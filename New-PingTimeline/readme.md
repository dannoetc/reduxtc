
# New-PingTimeline.ps1

A PowerShell 5.1–compatible utility for visualizing network latency over time — a lightweight, self-contained alternative to PingPlotter. It collects ping results, generates statistics, and outputs a standalone HTML report with a live-style timeline graph.

## Features

- ✅ **Windows PowerShell 5.1 compatible** (no `?`, `??`, or inline expressions)
- 📈 **Ping timeline visualization** in an interactive HTML report
- 🕒 **Dual-axis chart** — latency (ms) with timestamps below
- 📊 **Key metrics**: sent, received, loss %, min/avg/max, jitter (MAD)
- 💾 **Optional CSV export** for raw data
- 🧾 **Self-contained HTML** — no internet/CDN dependencies
- ⚙️ **No admin required**, runs as a normal user

## Usage

```powershell
.\New-PingTimeline.ps1 -Target <host> [-IntervalSeconds 1] [-DurationMinutes 5] [-Count <n>] [-TimeoutMs 1000] [-OutFile <path>] [-CsvFile <path>]
```

### Examples

Ping Google DNS for 5 minutes and generate a report:

```powershell
.\New-PingTimeline.ps1 -Target 8.8.8.8 -IntervalSeconds 5 -DurationMinutes 5 -OutFile C:\ReduxTC\PingReport.html
```

Ping for a fixed number of attempts (instead of duration):

```powershell
.\New-PingTimeline.ps1 -Target 1.1.1.1 -Count 300 -IntervalSeconds 1
```

Save a CSV file in addition to HTML:

```powershell
.\New-PingTimeline.ps1 -Target example.com -DurationMinutes 2 -CsvFile C:\Temp\ping.csv
```

## Output

- **HTML Report**  
  Displays latency timeline (with red loss markers), summary stats, and sample table.

- **CSV File (optional)**  
  Contains timestamp, index, RTT, success, status, and TTL.

## Notes

- The report is fully self-contained — just open it in any browser.
- Long-duration tests automatically thin the table to avoid performance issues.
- Works great for diagnosing intermittent connectivity or jitter over time.

## Example Screenshot

![Ping Timeline Example](https://github.com/dannoetc/reduxtc/blob/main/New-PingTimeline/pingtimeline.png?raw=true)

## Author

**Redux Technology Consulting**  
Created by [Dan Nelson](https://github.com/dannoetc) for use in MSP automation and diagnostics.

---

© 2025  — MIT License