$WshShell = New-Object -comObject WScript.Shell
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$Shortcut = $WshShell.CreateShortcut("$DesktopPath\駐車場データダッシュボード起動.lnk")
$Shortcut.TargetPath = "$PWD\launch_dashboard.bat"
$Shortcut.WorkingDirectory = "$PWD"
$Shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,130"
$Shortcut.Save()

Write-Host "デスクトップにショートカットを作成しました。"
