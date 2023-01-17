$ComObj = New-Object -ComObject WScript.Shell
$ShortCut = $ComObj.CreateShortcut("C:\Users\Public\Desktop\quickassist.lnk")
$ShortCut.TargetPath = "%windir%\system32\quickassist.exe"
$ShortCut.FullName 
$ShortCut.Save()
