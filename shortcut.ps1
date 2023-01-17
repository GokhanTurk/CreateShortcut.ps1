$ComObj = New-Object -ComObject WScript.Shell
$ShortCut = $ComObj.CreateShortcut("$Env:USERPROFILE\desktop\quickassist.lnk")
$ShortCut.TargetPath = "%windir%\system32\quickassist.exe"
$ShortCut.FullName 
$ShortCut.Save()
