Set oShell = CreateObject("Shell.Application")
Set oFolder = oShell.BrowseForFolder(0, "Please choose a folder", 0)
If Not oFolder Is Nothing Then
    WScript.Echo oFolder.Self.Path
End If
