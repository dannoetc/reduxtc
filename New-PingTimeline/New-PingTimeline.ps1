<# 
.SYNOPSIS
  Ping a target over time and export an HTML timeline report (PingPlotter-style).

.EXAMPLE
  .\New-PingTimeline.ps1 -Target 8.8.8.8 -IntervalSeconds 1 -DurationMinutes 5 -OutFile C:\Temp\PingReport.html

.PARAMETER Target
  Hostname or IP to ping.

.PARAMETER IntervalSeconds
  Seconds between pings (default 1).

.PARAMETER DurationMinutes
  Total duration in minutes (alternative to -Count).

.PARAMETER Count
  Number of pings to send (alternative to -DurationMinutes).

.PARAMETER TimeoutMs
  Ping timeout in milliseconds (default 1000).

.PARAMETER OutFile
  HTML report path. Defaults to $PWD\PingReport_<target>_<timestamp>.html

.PARAMETER CsvFile
  Optional CSV output path.

.NOTES
  Compatible with Windows PowerShell 5.1+.
#>

param(
    [Parameter(Mandatory=$true)][string]$Target,
    [int]$IntervalSeconds = 1,
    [int]$DurationMinutes,
    [int]$Count,
    [int]$TimeoutMs = 1000,
    [string]$OutFile,
    [string]$CsvFile
)

function New-DefaultPath([string]$ext, [string]$namePrefix="PingReport") {
    $safeTarget = ($Target -replace '[^\w\.-]', '_')
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    Join-Path -Path (Get-Location) -ChildPath ("{0}_{1}_{2}.{3}" -f $namePrefix,$safeTarget,$stamp,$ext.TrimStart('.'))
}

# Validate duration vs count
if (-not $DurationMinutes -and -not $Count) { $DurationMinutes = 5 }
if ($DurationMinutes -and $Count) {
    Write-Warning "Both -DurationMinutes and -Count were provided. Using -Count and ignoring -DurationMinutes."
    $DurationMinutes = $null
}

if (-not $OutFile) { $OutFile = New-DefaultPath -ext 'html' }
$EmitCsv = $false
if ($CsvFile) { $EmitCsv = $true }

# Storage
$data = New-Object System.Collections.Generic.List[object]
[int]$sent = 0
[int]$recv = 0

# Timing
$startTime = Get-Date
$stopTime = $null
if ($DurationMinutes) { $stopTime = $startTime.AddMinutes($DurationMinutes) }

# Pinger
$ping = New-Object System.Net.NetworkInformation.Ping
Write-Host ("Pinging {0} every {1}s, timeout {2}ms..." -f $Target,$IntervalSeconds,$TimeoutMs)

# Main loop
$index = 0
while ($true) {
    if ($Count) {
        if ($index -ge $Count) { break }
    } elseif ($stopTime) {
        if ((Get-Date) -ge $stopTime) { break }
    }

    $tickStart = Get-Date
    $roundtrip = $null
    $status = "Unknown"
    $ttl = $null
    $ok = $false

    try {
        $reply = $ping.Send($Target, $TimeoutMs)
        $sent++
        if ($reply -ne $null) {
            $status = [string]$reply.Status
            if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $ok = $true
                $recv++
                $roundtrip = [int]$reply.RoundtripTime
                if ($reply.Options -ne $null) { $ttl = $reply.Options.Ttl }
            }
        } else {
            $status = "NoReply"
        }
    } catch {
        $status = "Exception: " + $_.Exception.Message
    }

    $now = Get-Date
    $elapsedVal = $null
    if ($ok) { $elapsedVal = [int]$roundtrip }

    $data.Add([pscustomobject]@{
        Index     = $index
        Time      = $now
        ElapsedMs = $elapsedVal
        Success   = $ok
        Status    = $status
        TTL       = $ttl
    })
    $index++

    # Pace to interval
    $elapsedMs = [int]((Get-Date) - $tickStart).TotalMilliseconds
    $sleepMs = ($IntervalSeconds * 1000) - $elapsedMs
    if ($sleepMs -gt 0) { Start-Sleep -Milliseconds $sleepMs }
}

$endTime = Get-Date

# Compute stats
$successful = $data | Where-Object { $_.Success }
$latencies = @()
foreach ($row in $successful) { $latencies += [int]$row.ElapsedMs }

$min = $null; $max = $null; $avg = $null; $stdev = $null; $jitterMad = $null
if ($latencies.Count -gt 0) {
    $min = ($latencies | Measure-Object -Minimum).Minimum
    $max = ($latencies | Measure-Object -Maximum).Maximum
    $avg = [math]::Round((($latencies | Measure-Object -Average).Average),2)

    # Standard deviation (population)
    $mean = $avg
    $variance = 0
    foreach ($v in $latencies) { $variance += [math]::Pow(($v - $mean),2) }
    $variance = $variance / $latencies.Count
    $stdev = [math]::Round([math]::Sqrt($variance),2)

    # Jitter (median absolute difference of successive RTTs)
    $diffs = @()
    for ($i=1; $i -lt $latencies.Count; $i++) { $diffs += [math]::Abs($latencies[$i]-$latencies[$i-1]) }
    if ($diffs.Count -gt 0) {
        $sorted = $diffs | Sort-Object
        $mid = [int][math]::Floor($sorted.Count/2)
        if ($sorted.Count % 2 -eq 0) { $jitterMad = [math]::Round((($sorted[$mid-1] + $sorted[$mid]) / 2),2) }
        else { $jitterMad = [math]::Round($sorted[$mid],2) }
    }
}

$loss = $sent - $recv
$lossPct = 0
if ($sent -gt 0) { $lossPct = [math]::Round(($loss / $sent) * 100, 2) }

# Optional CSV
if ($EmitCsv) {
    try {
        $dir = [System.IO.Path]::GetDirectoryName($CsvFile)
        if ($dir -and -not (Test-Path $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
        $data | Select-Object Time,Index,ElapsedMs,Success,Status,TTL | Export-Csv -NoTypeInformation -Path $CsvFile -Encoding UTF8
        Write-Host "Wrote CSV: $CsvFile"
    } catch {
        Write-Warning "Failed to write CSV: $($_.Exception.Message)"
    }
}

# Prepare arrays for HTML
$timesIso = @()
$rtts = @()
$flags = @()
$statuses = @()
$ttls = @()
foreach ($row in $data) {
    $timesIso += $row.Time.ToString("o")
    if ($row.ElapsedMs -ne $null) { $rtts += [int]$row.ElapsedMs } else { $rtts += $null }
    $flags += [bool]$row.Success
    $statuses += ($row.Status -replace '"','\"')
    if ($row.TTL -ne $null) { $ttls += [int]$row.TTL } else { $ttls += $null }
}

# Infer Y max for chart
$maxY = 0
foreach ($v in $rtts) { if ($v -ne $null -and $v -gt $maxY) { $maxY = $v } }
if ($maxY -lt 100) { $maxY = 100 }
$maxY = [int]([math]::Ceiling($maxY / 50.0) * 50)

# Precompute display strings
$minStr = "-"; if ($min -ne $null) { $minStr = "$min" }
$avgStr = "-"; if ($avg -ne $null) { $avgStr = "$avg" }
$maxStr = "-"; if ($max -ne $null) { $maxStr = "$max" }
$jitStr = "-"; if ($jitterMad -ne $null) { $jitStr = "$jitterMad" }
$csvTip = "Tip: Use -CsvFile to export raw data."
if ($EmitCsv -and $CsvFile) { $csvTip = "CSV written to <code>$CsvFile</code>." }

# Ensure output folder exists
$outDir = [System.IO.Path]::GetDirectoryName($OutFile)
if ($outDir -and -not (Test-Path $outDir)) { $null = New-Item -ItemType Directory -Path $outDir -Force }

# JSON blobs for HTML
$jsonTimes = ($timesIso | ConvertTo-Json -Compress)
$jsonRtts  = ($rtts     | ConvertTo-Json -Compress)
$jsonOk    = ($flags    | ConvertTo-Json -Compress)
$jsonSts   = ($statuses | ConvertTo-Json -Compress)
$jsonTtls  = ($ttls     | ConvertTo-Json -Compress)

# HTML
$html = @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Ping Timeline — $Target</title>
<style>
  :root { --bg:#0b1020; --card:#121936; --ink:#f5f7ff; --muted:#a8b0d9; --accent:#62c4ff; --loss:#ff6b6b; --grid:#2a355f; }
  body { background:var(--bg); color:var(--ink); font:14px/1.4 -apple-system,Segoe UI,Roboto,Inter,Helvetica,Arial,sans-serif; margin:0; padding:24px; }
  .wrap { max-width:1000px; margin:0 auto; }
  .card { background:var(--card); border-radius:16px; padding:20px; box-shadow:0 6px 30px rgba(0,0,0,.25); }
  h1 { font-size:20px; margin:0 0 4px 0; }
  .sub { color:var(--muted); margin-bottom:16px; }
  .grid { display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin:12px 0 4px 0; }
  .stat { background:#0e1530; border:1px solid #1c2750; border-radius:12px; padding:12px; }
  .stat .k { color:var(--muted); font-size:12px; }
  .stat .v { font-size:16px; font-weight:600; }
  canvas { width:100%; height:360px; display:block; }
  .legend { display:flex; gap:16px; color:var(--muted); font-size:12px; margin-top:6px; }
  .dot { width:10px; height:10px; border-radius:50%; display:inline-block; vertical-align:middle; margin-right:6px; }
  .lat { background:var(--accent); }
  .los { background:var(--loss); }
  table { width:100%; border-collapse:collapse; margin-top:12px; }
  th, td { text-align:left; padding:6px 8px; border-bottom:1px solid #1e264a; color:#cdd4ff; }
  th { color:#9aa6e0; font-weight:600; }
  .footer { color:var(--muted); margin-top:10px; font-size:12px; }
  code { color:#8fd1ff; }
  .btn { background:#0e1530; color:#cdd4ff; border:1px solid #1c2750; border-radius:8px; padding:6px 10px; cursor:pointer; }
  .btn:hover { filter:brightness(1.1); }
</style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h1>Ping Timeline &mdash; $Target</h1>
    <div class="sub">
      $( $startTime.ToString("yyyy-MM-dd HH:mm:ss") ) &rarr; $( $endTime.ToString("yyyy-MM-dd HH:mm:ss") )
      &bull; Interval: $IntervalSeconds s &bull; Timeout: $TimeoutMs ms &bull; Samples: $($data.Count)
    </div>
    <div class="grid">
      <div class="stat"><div class="k">Sent</div><div class="v">$sent</div></div>
      <div class="stat"><div class="k">Received</div><div class="v">$recv</div></div>
      <div class="stat"><div class="k">Loss</div><div class="v">$loss ($lossPct`%)</div></div>
      <div class="stat"><div class="k">RTT (min/avg/max)</div><div class="v">$minStr / $avgStr / $maxStr&nbsp;ms</div></div>
      <div class="stat"><div class="k">Jitter (MAD)</div><div class="v">$jitStr&nbsp;ms</div></div>
    </div>

    <canvas id="chart" width="1000" height="360"></canvas>
    <div class="legend">
      <span><span class="dot lat"></span>Latency (ms)</span>
      <span><span class="dot los"></span>Loss</span>
    </div>

    <div style="text-align:right; margin-top:8px; margin-bottom:6px;">
      <button id="toggleRows" class="btn">Show all</button>
    </div>

    <table>
      <thead><tr><th>Time</th><th>RTT (ms)</th><th>Status</th><th>TTL</th></tr></thead>
      <tbody id="rows"></tbody>
    </table>

    <div class="footer">Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss"). $csvTip</div>
  </div>
</div>

<script>
(function(){
  const times = $jsonTimes;
  const rtts  = $jsonRtts;
  const sts   = $jsonSts;
  const ttls  = $jsonTtls;
  const ymax  = $maxY;

  const canvas = document.getElementById('chart');
  const ctx = canvas.getContext('2d');
  const W = canvas.width, H = canvas.height;
  const padL = 50, padR = 12, padT = 10, padB = 45;

  function x(i){ return padL + (i * (W - padL - padR) / Math.max(1,(rtts.length-1))); }
  function y(ms){ var v = Math.min(ms, ymax); return H - padB - (v * (H - padT - padB) / ymax); }

  // Background
  ctx.fillStyle = "#0b1020"; ctx.fillRect(0,0,W,H);

  // Grid
  ctx.strokeStyle = "#2a355f"; ctx.lineWidth = 1;
  var gridStep = Math.max(25, Math.ceil(ymax/6/25)*25);
  for (var ms=0; ms<=ymax; ms+=gridStep){
    var gy = y(ms);
    ctx.beginPath(); ctx.moveTo(padL,gy); ctx.lineTo(W-padR,gy); ctx.stroke();
    ctx.fillStyle="#a8b0d9"; ctx.font="12px Segoe UI, Roboto, Arial";
    ctx.fillText(ms+" ms", 6, gy-2);
  }
  ctx.beginPath(); ctx.moveTo(padL, H-padB); ctx.lineTo(W-padR, H-padB); ctx.stroke();

  // Line
  ctx.lineWidth = 2; ctx.strokeStyle = "#62c4ff"; ctx.beginPath();
  var started = false;
  for (var i=0;i<rtts.length;i++){
    var v = rtts[i];
    if (v===null){ started=false; continue; }
    if (!started){ ctx.moveTo(x(i), y(v)); started=true; } else { ctx.lineTo(x(i), y(v)); }
  }
  ctx.stroke();

  // Loss dots
  for (var j=0;j<rtts.length;j++){
    if (rtts[j]===null){ var cx=x(j), cy=H-padB-4; ctx.fillStyle="#ff6b6b"; ctx.beginPath(); ctx.arc(cx,cy,3,0,Math.PI*2); ctx.fill(); }
  }

  // Time labels (secondary x-axis)
  ctx.fillStyle="#a8b0d9"; ctx.font="11px Segoe UI, Roboto, Arial";
  var tickCount = Math.min(10, Math.max(3, Math.floor(rtts.length / 10)));
  var step = Math.max(1, Math.floor(rtts.length / tickCount));
  for (var i=0; i<rtts.length; i+=step){
    var t = new Date(times[i]);
    var label = t.toLocaleTimeString([], {hour:'2-digit',minute:'2-digit',second:'2-digit'});
    var tx = x(i);
    var tw = ctx.measureText(label).width;
    ctx.fillText(label, tx - tw/2, H - 8);
    ctx.strokeStyle="#2a355f";
    ctx.beginPath(); ctx.moveTo(tx, H - padB); ctx.lineTo(tx, H - padB + 4); ctx.stroke();
  }

  // Table with expand/collapse
  var tbody = document.getElementById('rows');
  var toggle = document.getElementById('toggleRows');
  var INITIAL_ROWS = 20;

  function rowHtml(idx){
    var t = new Date(times[idx]).toLocaleString();
    var v = (rtts[idx]===null ? "-" : rtts[idx]);
    var s = sts[idx] || "";
    var ttl = (ttls[idx]===null || ttls[idx]===undefined) ? "-" : ttls[idx];
    return "<tr><td>"+t+"</td><td>"+v+"</td><td>"+s+"</td><td>"+ttl+"</td></tr>";
  }

  function renderTable(limit){
    var len = times.length;
    var end = limit ? Math.min(limit, len) : len;
    var html = "";
    for (var k=0; k<end; k++){ html += rowHtml(k); }
    tbody.innerHTML = html;
    if (toggle){
      if (end < len){ toggle.textContent = "Show all ("+len+")"; }
      else { toggle.textContent = "Collapse to first 20"; }
    }
  }

  tbody.setAttribute('data-expanded','0');
  renderTable(INITIAL_ROWS);

  if (toggle){
    toggle.addEventListener('click', function(){
      var expanded = tbody.getAttribute('data-expanded') === '1';
      if (expanded){ tbody.setAttribute('data-expanded','0'); renderTable(INITIAL_ROWS); }
      else { tbody.setAttribute('data-expanded','1'); renderTable(0); }
    });
  }
})();
</script>
</body>
</html>
"@

try {
    $null = New-Item -ItemType File -Path $OutFile -Force
    Set-Content -Path $OutFile -Value $html -Encoding UTF8
    Write-Host "Report written to: $OutFile"
} catch {
    Write-Error "Failed to write HTML: $($_.Exception.Message)"
}
