$BGShortcut = Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\BGInfo.lnk"

if ($BGShortcut) { 
	exit 1 
	}
	
else { 
	exit 0
	}
	