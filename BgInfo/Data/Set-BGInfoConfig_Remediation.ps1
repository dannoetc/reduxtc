# Download the BGInfo config file to the root of C: 

Invoke-WebRequest https://github.com/dannoetc/reduxtc/raw/refs/heads/main/BgInfo/Data/config.bgi -outfile C:\config.bgi

# Verify BGInfo Exists 
$TestBgInfo = Test-Path "C:\Program Files\WindowsApps\Microsoft.SysinternalsSuite_2025.5.0.0_x64__8wekyb3d8bbwe\Tools"

# Add to startup if BGInfo's installed 
if ($TestBgInfo) {
$Shell = New-Object –ComObject ("WScript.Shell") 
$ShortCut = $Shell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\BGInfo.lnk")
$ShortCut.TargetPath="C:\Program Files\WindowsApps\Microsoft.SysinternalsSuite_2025.5.0.0_x64__8wekyb3d8bbwe\Tools\bginfo.exe"
$ShortCut.Arguments="C:\config.bgi /timer:0 /silent /nolicprompt" 
$ShortCut.IconLocation = "C:\Program Files\WindowsApps\Microsoft.SysinternalsSuite_2025.5.0.0_x64__8wekyb3d8bbwe\Tools\bginfo.exe, 0"; $ShortCut.Save()
}