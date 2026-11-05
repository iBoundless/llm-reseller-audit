<#
.SYNOPSIS
  Fingerprint probe for an OpenAI-compatible LLM endpoint.

.DESCRIPTION
  Sends a fixed battery of requests and records everything observable about the
  model behind the endpoint: response id / model / system_fingerprint, selected
  HTTP headers, tokenizer count on a fixed text, next-token logprobs, vision OCR,
  the thinking (reasoning_content) field, tool-call support and, optionally, a
  long-context needle test.

  Run it once against the official API and once against the reseller, then
  compare the two JSON cards.

.EXAMPLE
  $env:DSEEK_API_KEY = 'sk-...'

  .\probe.ps1 -Name official -BaseUrl https://api.deepseek.com -Model deepseek-flash
  .\probe.ps1 -Name proxy    -BaseUrl https://reseller.example/v1 -Model deepseek-v4.1-flash-1017

  .\probe.ps1 -Compare .\probe-official.json,.\probe-proxy.json

.EXAMPLE
  # long-context check at ~150K tokens
  .\probe.ps1 -Name proxy -BaseUrl https://reseller.example/v1 -Model deepseek-v4.1-flash-1017 -LongTokens 150000
#>
#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Probe')]
param(
    [Parameter(ParameterSetName = 'Probe', Position = 0)]
    [string]$Name = 'probe',

    [Parameter(ParameterSetName = 'Probe')]
    [string]$BaseUrl = $env:DSEEK_BASE_URL,

    [Parameter(ParameterSetName = 'Probe')]
    [string]$ApiKey = $env:DSEEK_API_KEY,

    [Parameter(ParameterSetName = 'Probe')]
    [string]$Model = $env:DSEEK_MODEL,

    # 0 = skip; 150000 = needle test at ~150K tokens
    [Parameter(ParameterSetName = 'Probe')]
    [int]$LongTokens = 0,

    [Parameter(ParameterSetName = 'Probe')]
    [switch]$SkipVision,

    # thinking models spend max_tokens on reasoning first; disable it for plain probes
    [Parameter(ParameterSetName = 'Probe')]
    [switch]$NoThinking,

    [Parameter(ParameterSetName = 'Compare', Mandatory)]
    [string[]]$Compare
)

$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 does not preload System.Net.Http.
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch { }

# ---------------------------------------------------------------- compare mode
if ($PSCmdlet.ParameterSetName -eq 'Compare') {
    # accepts .\a.json,.\b.json or a single ".\a.json,.\b.json" string (powershell.exe -File)
    $files = $Compare -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $cards = foreach ($f in $files) {
        if (-not (Test-Path $f)) { throw "File not found: $f" }
        Get-Content $f -Raw | ConvertFrom-Json
    }
    $keys = 'model_returned', 'id_prefix', 'system_fingerprint', 'prompt_tokens',
            'vision', 'thinking_flag', 'tools_flag', 'logprob_top1', 'server'
    $labelWidth = ($cards.name | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum
    if ($labelWidth -lt 10) { $labelWidth = 10 }

    foreach ($k in $keys) {
        Write-Host ''
        Write-Host $k -ForegroundColor Cyan
        foreach ($c in $cards) {
            $v = $c.$k
            if ($null -eq $v -or "$v" -eq '') { $v = '-' }
            Write-Host ("  {0,-$labelWidth}  {1}" -f $c.name, $v)
        }
    }

    Write-Host ''
    Write-Host 'long_context' -ForegroundColor Cyan
    foreach ($c in $cards) {
        $lc = $c.long_context
        $v = if ($lc -and $lc.prompt_tokens) {
            "prompt_tokens=$($lc.prompt_tokens)  found=$($lc.found)  in=$($lc.found_in)  ms=$($lc.ms)"
        } else { "$lc" }
        Write-Host ("  {0,-$labelWidth}  {1}" -f $c.name, $v)
    }

    Write-Host ''
    Write-Host 'thinking' -ForegroundColor Cyan
    foreach ($c in $cards) {
        $t = $c.thinking
        $v = if ($t -and $t.reasoning_chars -ne $null) { "reasoning=$($t.reasoning_chars) chars  answer=$($t.answer)" } else { "$t" }
        Write-Host ("  {0,-$labelWidth}  {1}" -f $c.name, $v)
    }
    return
}

if (-not $BaseUrl) { throw 'BaseUrl is required (or set $env:DSEEK_BASE_URL)' }
if (-not $Model)   { throw 'Model is required (or set $env:DSEEK_MODEL)' }

$BaseUrl = $BaseUrl.TrimEnd('/')
$script:ApiKey = $ApiKey
$script:NoThinking = [bool]$NoThinking

# ------------------------------------------------------------------- transport
function Invoke-ChatOnce {
    param(
        [hashtable]$Body,
        [int]$TimeoutSec = 600
    )
    $json = $Body | ConvertTo-Json -Depth 20 -Compress
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $client = $null
    try {
        $client = New-Object System.Net.Http.HttpClient
        $client.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
        if ($script:ApiKey) {
            $client.DefaultRequestHeaders.Authorization =
                New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $script:ApiKey)
        }
        $content = New-Object System.Net.Http.StringContent($json, [System.Text.Encoding]::UTF8, 'application/json')
        $resp = $client.PostAsync("$BaseUrl/chat/completions", $content).GetAwaiter().GetResult()
        $text = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $sw.Stop()
        $obj = $null
        try { $obj = $text | ConvertFrom-Json } catch { }
        $h = @{}
        foreach ($pair in $resp.Headers) { $h[$pair.Key] = ($pair.Value -join ', ') }
        return [pscustomobject]@{
            Status = [int]$resp.StatusCode; Body = $obj; Raw = $text
            Headers = $h; Ms = [int]$sw.ElapsedMilliseconds
        }
    }
    catch {
        $sw.Stop()
        return [pscustomobject]@{
            Status = -1; Body = $null; Raw = $_.Exception.Message
            Headers = @{}; Ms = [int]$sw.ElapsedMilliseconds
        }
    }
    finally {
        if ($client) { $client.Dispose() }
    }
}

# Gateways behind Cloudflare sometimes answer 524 after a ~2 min hang; retry once.
function Invoke-Chat {
    param([hashtable]$Body, [int]$TimeoutSec = 600)
    $r = Invoke-ChatOnce -Body $Body -TimeoutSec $TimeoutSec
    if ($r.Status -eq -1 -or $r.Status -ge 500) {
        Start-Sleep -Milliseconds 800
        $r2 = Invoke-ChatOnce -Body $Body -TimeoutSec $TimeoutSec
        return $r2
    }
    return $r
}

function New-Body {
    param([object]$Messages, [hashtable]$Extra = @{})
    $b = @{ model = $Model; messages = $Messages; temperature = 0; stream = $false }
    foreach ($k in $Extra.Keys) { $b[$k] = $Extra[$k] }
    if ($script:NoThinking) { $b['thinking'] = @{ type = 'disabled' } }
    return $b
}

function Get-ReplyText {
    param($Response)
    if ($Response.Body -and $Response.Body.choices) { return $Response.Body.choices[0].message.content }
    return $null
}

$card = [ordered]@{
    name    = $Name
    baseUrl = $BaseUrl
    model   = $Model
    ts      = (Get-Date).ToString('s')
}

Write-Host "== $Name ==" -ForegroundColor Cyan
Write-Host "endpoint: $BaseUrl  model: $Model"

# ---------------------------------------------- 1. echo: id / model / fingerprint
$echo = Invoke-Chat -Body (New-Body -Messages @(@{ role = 'user'; content = 'ping' }) -Extra @{ max_tokens = 8 })
$echoOut = [ordered]@{ status = $echo.Status; ms = $echo.Ms }
if ($echo.Body) {
    $echoOut.id                 = $echo.Body.id
    $echoOut.id_prefix          = if ($echo.Body.id) { ($echo.Body.id -split '[-_]')[0] } else { $null }
    $echoOut.model_returned     = $echo.Body.model
    $echoOut.system_fingerprint = $echo.Body.system_fingerprint
    $echoOut.created            = $echo.Body.created
    $echoOut.usage              = $echo.Body.usage
}
else {
    $echoOut.error = $echo.Raw
}
$card.echo = $echoOut

$server = $echo.Headers['Server']
if (-not $server) { $server = $echo.Headers['server'] }
$card.server = $server
$card.model_returned = $echoOut.model_returned
$card.id_prefix = $echoOut.id_prefix
$card.system_fingerprint = $echoOut.system_fingerprint
$card.headers = [ordered]@{
    server    = $server
    via       = $echo.Headers['Via']
    powerded  = $echo.Headers['X-Powered-By']
    reqid     = $echo.Headers['X-Request-Id']
    all       = ($echo.Headers.Keys | Sort-Object) -join ', '
}

# ------------------------------------------------------ 2. tokenizer fingerprint
# The same message must produce the same prompt_tokens on the same tokenizer.
$fixed = @'
Токенизатор — это скучный, но надёжный отпечаток модели. Съешь же ещё этих мягких
французских булок, да выпей чаю. 0123456789 -- __init__ / v1/chat/completions.
'@
$tok = Invoke-Chat -Body (New-Body -Messages @(@{ role = 'user'; content = $fixed }) -Extra @{ max_tokens = 4 })
$tokOut = [ordered]@{ status = $tok.Status }
if ($tok.Body -and $tok.Body.usage) {
    $tokOut.prompt_tokens = $tok.Body.usage.prompt_tokens
    $tokOut.completion_tokens = $tok.Body.usage.completion_tokens
}
else { $tokOut.error = $tok.Raw }
$card.tokens = $tokOut
$card.prompt_tokens = $tokOut.prompt_tokens

# ------------------------------------------------------ 3. next-token distribution
# Same model + same version => same top-1 and very close logprobs.
$lp = Invoke-Chat -Body (New-Body -Messages @(@{
                role = 'user'
                content = 'Continue the sentence with exactly one word, no punctuation: "The capital of France is"'
            }) -Extra @{ max_tokens = 1; logprobs = $true; top_logprobs = 5 })
$lpOut = [ordered]@{ status = $lp.Status; supported = $false }
if ($lp.Body -and $lp.Body.choices) {
    $choice = $lp.Body.choices[0]
    if ($choice.logprobs) {
        $lpOut.supported = $true
        $content = $choice.logprobs.content
        if (-not $content -and $choice.logprobs.tokens) {
            $lpOut.top1 = $choice.logprobs.tokens[0]
            $lpOut.logprob = $choice.logprobs.token_logprobs[0]
        }
        elseif ($content) {
            $lpOut.top1 = $content[0].token
            $lpOut.logprob = $content[0].logprob
            $lpOut.top = @($content[0].top_logprobs | ForEach-Object { "$($_.token):$([math]::Round($_.logprob,3))" })
        }
    }
    $lpOut.text = $choice.message.content
}
else { $lpOut.error = $lp.Raw }
$card.logprobs = $lpOut
$card.logprob_top1 = $lpOut.top1

# --------------------------------------------------------------- 4. vision (OCR)
$digits = '4821'
$imgB64 = $null
if (-not $SkipVision) {
    try {
        Add-Type -AssemblyName System.Drawing
        $bmp = New-Object System.Drawing.Bitmap(220, 80)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $font = New-Object System.Drawing.Font('Consolas', 38, [System.Drawing.FontStyle]::Bold)
        $g.DrawString($digits, $font, [System.Drawing.Brushes]::Black, 12, 10)
        $g.Dispose()
        $ms = New-Object System.IO.MemoryStream
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        $imgB64 = [Convert]::ToBase64String($ms.ToArray())
    }
    catch { $imgB64 = $null }
}
if ($imgB64) {
    $vis = Invoke-Chat -Body (New-Body -Messages @(@{
                    role = 'user'
                    content = @(
                        @{ type = 'text'; text = 'Read the digits in the image. Reply with digits only.' },
                        @{ type = 'image_url'; image_url = @{ url = "data:image/png;base64,$imgB64" } }
                    )
                }) -Extra @{ max_tokens = 400 })
    $visText = Get-ReplyText $vis
    $visReason = $null
    if ($vis.Body -and $vis.Body.choices) { $visReason = $vis.Body.choices[0].message.reasoning_content }
    if ($vis.Status -eq 200 -and -not $visText -and $visReason) {
        $card.vision = "200: empty content, reasoning only ($($visReason.Length) chars) - rerun with -NoThinking"
    }
    elseif ($vis.Status -eq 200) { $card.vision = "200: $($visText -replace '\s', '')" }
    else { $card.vision = "$($vis.Status): $($vis.Raw)" }
}
else {
    $card.vision = 'skipped (image generation failed or -SkipVision)'
}

# ------------------------------------------------------- 5. thinking (reasoning)
$th = Invoke-Chat -Body (New-Body -Messages @(@{ role = 'user'; content = 'How many r letters are in the word strawberry? One short sentence.' }) -Extra @{ max_tokens = 400 })
$thOut = [ordered]@{ status = $th.Status }
if ($th.Body -and $th.Body.choices) {
    $msg = $th.Body.choices[0].message
    $rc = $msg.reasoning_content
    $thOut.has_reasoning_content = [bool]$rc
    $thOut.reasoning_chars = if ($rc) { $rc.Length } else { 0 }
    $thOut.usage = $th.Body.usage
    $thOut.answer = ($msg.content -replace '\s+', ' ')
}
else { $thOut.error = $th.Raw }
$card.thinking = $thOut
$card.thinking_flag = if ($thOut.has_reasoning_content) { 'yes' } else { 'no' }

# ------------------------------------------------------------ 6. tool-call support
$tool = @{
    type     = 'function'
    function = @{
        name        = 'get_weather'
        description = 'Get current weather for a city'
        parameters  = @{
            type       = 'object'
            properties = @{ city = @{ type = 'string' } }
            required   = @('city')
        }
    }
}
$tc = Invoke-Chat -Body (New-Body -Messages @(@{ role = 'user'; content = 'What is the weather in Paris? Use the tool.' }) -Extra @{ tools = @($tool); tool_choice = 'auto'; max_tokens = 100 })
$tcOut = [ordered]@{ status = $tc.Status; supported = $false }
if ($tc.Body -and $tc.Body.choices) {
    $calls = $tc.Body.choices[0].message.tool_calls
    if ($calls) {
        $tcOut.supported = $true
        $tcOut.call = "$($calls[0].function.name)($($calls[0].function.arguments))"
    }
    else { $tcOut.text = $tc.Body.choices[0].message.content }
}
else { $tcOut.error = $tc.Raw }
$card.tools = $tcOut
$card.tools_flag = if ($tcOut.supported) { 'yes' } else { 'no' }

# ------------------------------------------------------------ 7. long context
if ($LongTokens -gt 0) {
    $unit = 'The quick brown fox jumps over the lazy dog while the river runs quietly. '
    $secret = '74-19-QK'
    $chars = [int]($LongTokens * 3.9)
    $n = [int][math]::Ceiling($chars / $unit.Length)
    $before = -join ($unit * [int]($n * 0.6))
    $after = -join ($unit * [int]($n * 0.4))
    $prompt = "$before`nNOTE: the vault code is $secret.`n$after`n`nQuestion: what is the vault code? Reply with the code only."
    $lc = Invoke-Chat -Body (New-Body -Messages @(@{ role = 'user'; content = $prompt }) -Extra @{ max_tokens = 800 }) -TimeoutSec 900
    $lcText = Get-ReplyText $lc
    $lcReason = $null
    if ($lc.Body -and $lc.Body.choices) { $lcReason = $lc.Body.choices[0].message.reasoning_content }
    $haystack = "$lcText`n$lcReason"
    $found = [bool]($haystack -match [regex]::Escape($secret))
    $where = if (-not $found) { 'nowhere' }
             elseif ($lcText -match [regex]::Escape($secret)) { 'content' }
             else { 'reasoning only' }
    $card.long_context = [ordered]@{
        requested_tokens = $LongTokens
        status           = $lc.Status
        prompt_tokens    = if ($lc.Body -and $lc.Body.usage) { $lc.Body.usage.prompt_tokens } else { $null }
        found            = $found
        found_in         = $where
        answer           = if ($lcText) { $lcText.Trim() } elseif ($lcReason) { "(no content) $($lcReason.Trim())" } else { $lc.Raw }
        ms               = $lc.Ms
    }
    $card.long_context_flag = if ($lc.Status -ne 200) { "request failed ($($lc.Status))" }
                              elseif ($found) { "needle found ($where)" }
                              else { 'NEEDLE LOST' }
}
else {
    $card.long_context = 'skipped (-LongTokens 150000 to enable)'
    $card.long_context_flag = 'skipped'
}

# ------------------------------------------------------------------------ output
$outFile = Join-Path (Get-Location) "probe-$Name.json"
$card | ConvertTo-Json -Depth 20 | Set-Content -Path $outFile -Encoding utf8

Write-Host ''
Write-Host "id_prefix          : $($card.id_prefix)"
Write-Host "model_returned     : $($card.model_returned)"
Write-Host "system_fingerprint : $($card.system_fingerprint)"
Write-Host "server             : $($card.server)"
Write-Host "prompt_tokens      : $($card.prompt_tokens)"
Write-Host "logprob top1       : $($card.logprob_top1)  (supported: $($lpOut.supported))"
Write-Host "vision             : $($card.vision)"
Write-Host "thinking           : $($card.thinking_flag)  ($($thOut.reasoning_chars) chars)"
Write-Host "tools              : $($card.tools_flag)"
Write-Host "long context       : $($card.long_context_flag)"
Write-Host ''
Write-Host "saved: $outFile" -ForegroundColor Green
