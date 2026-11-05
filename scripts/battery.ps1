<#
.SYNOPSIS
  Behavioural battery: 10 tasks with objectively checkable answers.

.DESCRIPTION
  Runs the same fixed set of tasks against an OpenAI-compatible endpoint at
  temperature 0 with thinking disabled (so temperature actually applies) and
  records the answers, the objective verdict and the time.

  Run it once per endpoint, then compare the two result files.

.EXAMPLE
  .\battery.ps1 -Name official -BaseUrl https://api.deepseek.com -Model deepseek-flash -ApiKey "sk-..."
  .\battery.ps1 -Name proxy    -BaseUrl https://api.zexkora.cc/v1 -Model deepseek-v4.1-flash-1017 -ApiKey "sk-..."
  .\battery.ps1 -Compare .\battery-official.json,.\battery-proxy.json
#>
#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run', Position = 0)]
    [string]$Name = 'run',

    [Parameter(ParameterSetName = 'Run')]
    [string]$BaseUrl = $env:DSEEK_BASE_URL,

    [Parameter(ParameterSetName = 'Run')]
    [string]$ApiKey = $env:DSEEK_API_KEY,

    [Parameter(ParameterSetName = 'Run')]
    [string]$Model = $env:DSEEK_MODEL,

    # keep thinking ON (default is OFF so that temperature 0 is honoured)
    [Parameter(ParameterSetName = 'Run')]
    [switch]$Thinking,

    [Parameter(ParameterSetName = 'Compare', Mandatory)]
    [string[]]$Compare
)

$ErrorActionPreference = 'Stop'
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch { }

# ------------------------------------------------------------------- tasks
$tasks = @(
    @{ id = 'arith';   mode = 'number';   expect = '379';
       prompt = 'Compute (17 * 23) - (144 / 12). Reply with the number only.' }

    @{ id = 'percent'; mode = 'number';   expect = '4963';
       prompt = 'A price is 4800 roubles. It is raised by 17.5 percent, then discounted by 12 percent. Give the final price rounded to the nearest whole rouble. Reply with the number only.' }

    @{ id = 'letters'; mode = 'number';   expect = '4';
       prompt = 'How many times does the Cyrillic letter o appear in the word "молокопереработка"? Reply with a number only.' }

    @{ id = 'json';    mode = 'json';     expect = 'id=42, 3 lowercase tags, ok=true';
       prompt = 'Return ONLY a JSON object, no prose and no code fences: {"id": 42, "tags": [three lowercase strings], "ok": true}.' }

    @{ id = 'date';    mode = 'contains'; expect = 'thursday';
       prompt = 'What day of the week is 2026-11-05? Answer with one English word.' }

    @{ id = 'code';    mode = 'number';   expect = '3';
       prompt = 'What does this Python code print: print(len(set("banana")))? Reply with a number only.' }

    @{ id = 'logic';   mode = 'firstword'; expect = 'no';
       prompt = 'All A are B. Some B are C. Does it necessarily follow that some A are C? Answer yes or no, one word.' }

    @{ id = 'reverse'; mode = 'contains'; expect = '321cba';
       prompt = 'Reverse the string abc123. Reply with the result only.' }

    @{ id = 'fivewords'; mode = 'words5'; expect = '5 words, 5 distinct first letters';
       prompt = 'Write exactly five words about winter, each word starting with a different letter. Reply with the words only, separated by spaces.' }

    @{ id = 'unit';    mode = 'number';   expect = '5.63';
       prompt = 'Convert 3.5 miles to kilometres using 1 mile = 1.609344 km. Round to 2 decimals. Reply with the number only.' }
)

# ------------------------------------------------------------------ helpers
function Invoke-ChatOnce {
    param([hashtable]$Body, [int]$TimeoutSec = 300)
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
        $resp = $client.PostAsync("$($script:BaseUrl)/chat/completions", $content).GetAwaiter().GetResult()
        $text = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
        $sw.Stop()
        $obj = $null
        try { $obj = $text | ConvertFrom-Json } catch { }
        return [pscustomobject]@{ Status = [int]$resp.StatusCode; Body = $obj; Raw = $text; Ms = [int]$sw.ElapsedMilliseconds }
    }
    catch {
        $sw.Stop()
        return [pscustomobject]@{ Status = -1; Body = $null; Raw = $_.Exception.Message; Ms = [int]$sw.ElapsedMilliseconds }
    }
    finally { if ($client) { $client.Dispose() } }
}

function Invoke-Chat {
    param([hashtable]$Body, [int]$TimeoutSec = 300)
    $r = Invoke-ChatOnce -Body $Body -TimeoutSec $TimeoutSec
    if ($r.Status -eq -1 -or $r.Status -ge 500) {
        Start-Sleep -Milliseconds 800
        $r = Invoke-ChatOnce -Body $Body -TimeoutSec $TimeoutSec
    }
    return $r
}

function Normalize-Text {
    param([string]$s)
    if (-not $s) { return '' }
    $t = $s.ToLowerInvariant()
    $t = $t -replace '[«»"“”„]', ''
    $t = $t -replace '\s+', ' '
    $t = $t.Trim()
    $t = $t -replace '[\.\!\?;:,]+$', ''
    return $t
}

function Get-FirstNumber {
    param([string]$s)
    if (-not $s) { return $null }
    $m = [regex]::Match($s, '-?\d[\d \u00a0]*(?:[.,]\d+)?')
    if (-not $m.Success) { return $null }
    return ($m.Value -replace '[\s\u00a0]', '') -replace ',', '.'
}

function Test-Answer {
    param([hashtable]$Task, [string]$Text)
    $norm = Normalize-Text $Text
    $ok = $false
    switch ($Task.mode) {
        'number' {
            $n = Get-FirstNumber $Text
            $ok = ($null -ne $n) -and ([math]::Abs([double]$n - [double]$Task.expect) -le 0.02)
        }
        'contains' { $ok = $norm.Contains($Task.expect) }
        'firstword' {
            $w = ($norm -split '\s+')[0]
            $ok = ($w -eq $Task.expect)
        }
        'json' {
            $i = $Text.IndexOf('{'); $j = $Text.LastIndexOf('}')
            if ($i -ge 0 -and $j -gt $i) {
                try {
                    $o = $Text.Substring($i, $j - $i + 1) | ConvertFrom-Json
                    $lower = @($o.tags | Where-Object { $_ -cmatch '^[a-z]+$' })
                    $ok = ($o.id -eq 42) -and ($o.ok -eq $true) -and (@($o.tags).Count -eq 3) -and ($lower.Count -eq 3)
                }
                catch { $ok = $false }
            }
        }
        'words5' {
            $w = @($norm -split '\s+' | Where-Object { $_ })
            $first = @($w | ForEach-Object { $_.Substring(0, 1) } | Sort-Object -Unique)
            $ok = ($w.Count -eq 5) -and ($first.Count -eq 5)
        }
    }
    return [pscustomobject]@{ ok = $ok; norm = $norm }
}

# ---------------------------------------------------------------- compare mode
if ($PSCmdlet.ParameterSetName -eq 'Compare') {
    $files = $Compare -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $runs = foreach ($f in $files) {
        if (-not (Test-Path $f)) { throw "File not found: $f" }
        Get-Content $f -Raw | ConvertFrom-Json
    }

    $ids = @($runs[0].results | ForEach-Object { $_.id })
    Write-Host ''
    Write-Host ('{0,-11} ' -f 'task') -NoNewline
    foreach ($r in $runs) { Write-Host ('{0,-9} ' -f $r.name) -NoNewline }
    Write-Host ''
    Write-Host ('-' * (12 + 10 * $runs.Count))

    $totals = @{}
    foreach ($r in $runs) { $totals[$r.name] = 0 }

    foreach ($id in $ids) {
        $line = '{0,-11} ' -f $id
        foreach ($r in $runs) {
            $res = $r.results | Where-Object { $_.id -eq $id }
            $mark = if ($res.ok) { 'OK' } else { 'FAIL' }
            if ($res.ok) { $totals[$r.name]++ }
            $line += ('{0,-9} ' -f $mark)
        }
        Write-Host $line
    }

    Write-Host ''
    foreach ($r in $runs) { Write-Host ("{0}: {1}/{2} correct" -f $r.name, $totals[$r.name], $ids.Count) }

    $same = 0
    foreach ($id in $ids) {
        $norms = @($runs | ForEach-Object { ($_.results | Where-Object { $_.id -eq $id }).norm })
        $filled = @($norms | Where-Object { $_ })
        if ($norms.Count -gt 1 -and $filled.Count -eq $norms.Count -and (@($norms | Sort-Object -Unique).Count -eq 1)) { $same++ }
    }
    Write-Host ("identical answers: {0} / {1}" -f $same, $ids.Count)
    Write-Host ''
    Write-Host 'answers (truncated):'
    foreach ($id in $ids) {
        Write-Host "  $id"
        foreach ($r in $runs) {
            $res = $r.results | Where-Object { $_.id -eq $id }
            $a = if ($res.norm.Length -gt 90) { $res.norm.Substring(0, 90) + '...' } else { $res.norm }
            Write-Host ("    {0,-9} {1}" -f $r.name, $a)
        }
    }
    return
}

if (-not $BaseUrl) { throw 'BaseUrl is required' }
if (-not $Model) { throw 'Model is required' }

$script:BaseUrl = $BaseUrl.TrimEnd('/')
$script:ApiKey = $ApiKey

Write-Host "== battery: $Name ==" -ForegroundColor Cyan
Write-Host "endpoint: $script:BaseUrl  model: $Model  thinking: $(if ($Thinking) { 'on' } else { 'off' })"

$results = @()
foreach ($t in $tasks) {
    $body = @{
        model       = $Model
        messages    = @(@{ role = 'user'; content = $t.prompt })
        temperature = 0
        max_tokens  = 700
        stream      = $false
    }
    if (-not $Thinking) { $body['thinking'] = @{ type = 'disabled' } }

    $r = Invoke-Chat -Body $body
    $text = $null
    $reasoning = $null
    if ($r.Body -and $r.Body.choices) {
        $text = $r.Body.choices[0].message.content
        $reasoning = $r.Body.choices[0].message.reasoning_content
    }
    $check = Test-Answer -Task $t -Text $text

    $results += [pscustomobject]@{
        id                = $t.id
        expect            = $t.expect
        mode              = $t.mode
        status            = $r.Status
        ms                = $r.Ms
        ok                = $check.ok
        norm              = $check.norm
        answer            = $text
        reasoning_present = [bool]$reasoning
    }
    $mark = if ($check.ok) { 'OK  ' } else { 'FAIL' }
    Write-Host ("  {0,-11} {1} {2,6} ms  {3}" -f $t.id, $mark, $r.Ms, ($check.norm -replace '\s+', ' '))
}

$out = [ordered]@{
    name     = $Name
    baseUrl  = $script:BaseUrl
    model    = $Model
    thinking = [bool]$Thinking
    ts       = (Get-Date).ToString('s')
    results  = $results
}
$file = Join-Path (Get-Location) "battery-$Name.json"
$out | ConvertTo-Json -Depth 10 | Set-Content -Path $file -Encoding utf8

$correct = @($results | Where-Object { $_.ok }).Count
Write-Host ''
Write-Host ("correct: $correct / $($results.Count)") -ForegroundColor Green
Write-Host "saved: $file"
