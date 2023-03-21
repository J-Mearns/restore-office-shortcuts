#Script restores Start Menu shortcuts for Office Suite and certain misc. apps if detected.
#=========================================================================================


#Restore Office Suite.
======================

#Restore Microsoft Edge
if (!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")){
$WScript = New-Object -ComObject WScript.Shell
$SourcePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$shortcut = $WScript.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Edge.lnk")
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "Microsoft Edge"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Restore Outlook
if (!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")){
$WScript = New-Object -ComObject WScript.Shell
$SourcePath = "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
$shortcut = $WScript.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Outlook.lnk")
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "Outlook"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Restore Word
if (!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")){
$WScript = New-Object -ComObject WScript.Shell
$SourcePath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
$shortcut = $WScript.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk")
$shortcut.TargetPath = $SourcePath
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Restore Powerpoint
if (!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk")){
$WScript = New-Object -ComObject WScript.Shell
$SourcePath = "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
$shortcut = $WScript.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Powerpoint.lnk")
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "PowerPoint"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Restore Excel
if (!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")){
$WScript = New-Object -ComObject WScript.Shell
$SourcePath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$shortcut = $WScript.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Excel.lnk")
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "Excel"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Restore Publisher
if (!(Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher.lnk")){
$WScript = New-Object -ComObject WScript.Shell
$SourcePath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$shortcut = $WScript.CreateShortcut("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Publisher.lnk")
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "Publisher"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Checks for licensed applications, if detected restores shortcuts.
#=================================================================

#Checks for Visio
$Check = "C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE"
if(Test-Path $Check -PathType Leaf)
{
$SourcePath = $Check
$StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Visio.lnk"
$WScript = New-Object -ComObject WScript.Shell
$shortcut = $WScript.CreateShortcut($StartMenuPath)
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "Visio"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}


#Checks for Project
$Check = "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE"
if(Test-Path $Check -PathType Leaf)
{
$SourcePath = $Check
$StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Project.lnk"
$WScript = New-Object -ComObject WScript.Shell
$shortcut = $WScript.CreateShortcut($StartMenuPath)
$shortcut.TargetPath = $SourcePath
$shortcut.Description = "Project"
$shortcut.FullName
$shortcut.WindowStyle = 1
$shortcut.Save()
}

#Checks for other applications, if detected restores shortcuts.
#==============================================================

#Checks for Zoom
$Check = "C:\Program Files\Zoom\bin\Zoom.exe"
if (Test-Path $Check -PathType Leaf)
{
$SourcePath = $Check
$StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Zoom.lnk"
$WScript = New-Object -ComObject WScript.Shell
$shortcut = $WScript.CreateShortcut($StartMenuPath)
$shortcut.TargetPath = $SourcePath
$shortcut.Save()
}

#Checks for TeamViewer
$Check = "C:\Program Files\TeamViewer\TeamViewer.exe"
if(Test-Path $Check -PathType Leaf)
{
$SourcePath = $Check
$StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TeamViewer.lnk"
$WScript = New-Object -ComObject WScript.Shell
$shortcut = $WScript.CreateShortcut($StartMenuPath)
$shortcut.TargetPath = $SourcePath
$shortcut.Save()
}

#Checks for Chrome
$Check = "C:\Program Files\Google\Chrome\Application\chrome.exe"
if(Test-Path $Check -PathType Leaf)
{
$SourcePath = $Check
$StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Chrome.lnk"
$WScript = New-Object -ComObject WScript.Shell
$shortcut = $WScript.CreateShortcut($StartMenuPath)
$shortcut.TargetPath = $SourcePath
$shortcut.Save()
}

#Checks for Firefox
$Check = "C:\Program Files\Mozilla Firefox\firefox.exe"
if(Test-Path $Check -PathType Leaf)
{
$SourcePath = $Check
$StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Firefox.lnk"
$WScript = New-Object -ComObject WScript.Shell
$shortcut = $WScript.CreateShortcut($StartMenuPath)
$shortcut.TargetPath = $SourcePath
$shortcut.Save()
}