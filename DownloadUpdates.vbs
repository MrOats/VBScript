WScript.Echo "Ensuring .NET 3.5 is installed to ensure COM support in VBScript!"
Set objShell = wscript.CreateObject("wscript.shell")
objShell.Run "%SYSTEMROOT%\System32\dism.exe /Online /NoRestart /Enable-Feature /FeatureName:NetFx3", 0, true


' ========Search Section========
Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "Windows Update"

Set updateSearcher = updateSession.CreateUpdateSearcher()

WScript.Echo "Searching for updates..." & vbCRLF

Set searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

If searchResult.Updates.Count = 0 Then
  WScript.Echo "There are no applicable updates."
  WScript.Quit
End If

WScript.Echo "List of applicable items on the machine:"
currentCount = 0
For Each update in searchResult.Updates
  currentCount = currentCount + 1
  WScript.Echo currentCount & "> " & update.Title
Next
' ========END Search Section========
' ========Download Section========
Set updatesToDownload = CreateObject("Microsoft.Update.UpdateColl")

For Each update in searchResult.Updates
  update.AcceptEula()
  IF update.isDownloaded = false THEN
    updatesToDownload.Add(update)
  END IF
Next

WScript.Echo vbCRLF & "Downloading updates..."
currentCount = 0
For Each update in updatesToDownload
  Set downloader = updateSession.CreateUpdateDownloader()
  currentCount = currentCount + 1
  size = Round(CStr(update.MaxDownloadSize) * 0.000001, 2)
  WScript.Echo "Downloading " & currentCount & " of " & updatesToDownload.Count & ": " & update.Title & " (Size: " & size & " MBs)"
  Set toDownload = CreateObject("Microsoft.Update.UpdateColl")
  toDownload.Add(update)
  downloader.Updates = toDownload
  downloader.Download()
Next
' ========END Download Section========

'We're done here.
WScript.Quit
