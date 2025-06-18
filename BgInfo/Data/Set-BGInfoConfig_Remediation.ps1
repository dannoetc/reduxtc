$TestBgInfo = Test-Path "C:\Program Files\WindowsApps\Microsoft.SysinternalsSuite_2025.5.0.0_x64__8wekyb3d8bbwe\Tools"

if ($TestBgInfo) {
$Shell = New-Object –ComObject ("WScript.Shell") 
$ShortCut = $Shell.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\BGInfo.lnk")
$ShortCut.TargetPath="C:\Program Files\WindowsApps\Microsoft.SysinternalsSuite_2025.5.0.0_x64__8wekyb3d8bbwe\Tools\bginfo.exe"
$ShortCut.Arguments="C:\config.bgi /timer:0 /silent /nolicprompt" 
$ShortCut.IconLocation = "C:\Program Files\WindowsApps\Microsoft.SysinternalsSuite_2025.5.0.0_x64__8wekyb3d8bbwe\Tools\bginfo.exe, 0"; $ShortCut.Save()
}