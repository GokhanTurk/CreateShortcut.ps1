$ComObj = New-Object -ComObject WScript.Shell
$ShortCut = $ComObj.CreateShortcut("C:\Users\Public\Desktop\quickassist.lnk")
$ShortCut.TargetPath = "%windir%\system32\quickassist.exe"
$ShortCut.FullName 
$ShortCut.Save()
$ShortCut = $ComObj.CreateShortcut("C:\Users\Public\Desktop\Outlook.lnk")
$ShortCut.TargetPath = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$ShortCut.FullName 
$ShortCut.Save()
$ShortCut = $ComObj.CreateShortcut("C:\Users\Public\Desktop\Microsoft Teams.lnk")
$ShortCut.TargetPath = "$env:UserProfile\AppData\Local\Microsoft\Teams\Update.exe"
$ShortCut.Arguments = "--processStart `"Teams.exe`""
$ShortCut.FullName 
$ShortCut.Save()
