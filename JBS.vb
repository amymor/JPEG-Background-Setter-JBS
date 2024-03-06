Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.IO
' For Security settings (For Check if the current user or Administrator is the owner)
Imports System.Security.Principal
' For Process (cmd /c icacls)
Imports System.Diagnostics

Module Module1
    Private Const SPI_SETDESKWALLPAPER As Integer = &H14
    Private Const SPIF_UPDATEINIFILE As Integer = &H1
    Private Const SPIF_SENDCHANGE As Integer = &H2
    Private Const SHCNE_ASSOCCHANGED As Integer = &H8000000

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Function SystemParametersInfo(ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Function SHChangeNotify(ByVal eventId As Integer, ByVal flags As Integer, ByVal item1 As IntPtr, ByVal item2 As IntPtr) As Integer
    End Function

    Sub Main(args As String())
        If args.Length <> 1 Or args.Contains("/?") Then
            Console.WriteLine("")
            Console.WriteLine("JBS.exe /? (To show this help)")
            Console.WriteLine("")
            Console.WriteLine("Usage:")
            Console.WriteLine("JBS.exe ""Wallpaper_Path""")
            Console.WriteLine("JBS.exe /u (To Uninstall)" )
            Console.WriteLine("")
            Console.WriteLine("Wallpaper position setting would read from JBS.ini:")
            Console.WriteLine("  Position_Setting_Number = N")
            Console.WriteLine("")
            Console.WriteLine("The Position_Setting_Number should be the following numbers:")
            Console.WriteLine("  0 - Center")
            Console.WriteLine("  1 - Tile")
            Console.WriteLine("  2 - Stretch")
            Console.WriteLine("  6 - Fit")
            Console.WriteLine(" 10 - Fill")
            Console.WriteLine(" 22 - Span")
            Return
        End If

        ' Set variable  (used for /u and also for changing security settings)
        Dim cachedFilesFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Themes\CachedFiles"

        ' Define the process (used for /u and also for changing security settings)
        Dim process As New Process()
        process.StartInfo.FileName = "cmd.exe"
        process.StartInfo.Verb = "runas" ' Run as administrator
        process.StartInfo.RedirectStandardInput = True
        process.StartInfo.RedirectStandardOutput = True
        process.StartInfo.CreateNoWindow = True
        process.StartInfo.UseShellExecute = False

    ' Check if the first argument is /u or /U
    If args(0).ToLower() = "/u" Then
        ' if the app run with  /u or /U as parameter then reset security settings and delete CachedFiles folder
        process.Start()
        ' process.StandardInput.WriteLine("takeown /f " & cachedFilesFolder & " /r /d")
        process.StandardInput.WriteLine("icacls " & cachedFilesFolder & " /reset")
        process.StandardInput.WriteLine("icacls " & cachedFilesFolder & " /inheritance:e")
        process.StandardInput.WriteLine("RD /S /Q " & cachedFilesFolder)
        process.StandardInput.Close()
        process.WaitForExit()
        ' Directory.Delete(cachedFilesFolder, True)
		Environment.Exit(0)
    Else
        ' msgbox ("error1", , "E")
        ' Treat the first argument as a file path
        Dim imagePath As String = args(0)
        ' Copy and Replace the "imagePath" variable file to "%AppData%\Microsoft\Windows\Themes" (rename as "TranscodedWallpaper" without extension)
        File.Copy(imagePath, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Themes\TranscodedWallpaper", True)
    End If

        ' Read the position setting from the JBS.ini file
        Dim positionSetting As Integer
        Using sr As StreamReader = New StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\" & "JBS.ini")
            While Not sr.EndOfStream
                Dim line As String = sr.ReadLine()
                If line.StartsWith("Position_Setting_Number = ") Then
                    positionSetting = Integer.Parse(line.Substring("Position_Setting_Number = ".Length))
                    Exit While
                End If
            End While
        End Using

        ' Set the wallpaper position
        Dim key As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop", True)
        If key IsNot Nothing Then
            key.SetValue("WallpaperStyle", positionSetting, RegistryValueKind.String)
            key.Close()
        End If

        ' Set the TranscodedWallpaper itself as wallpaper (it will delete CachedFiles if owner is not SYSTEM)
		Dim Transcoded_Wallpaper As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Microsoft\Windows\Themes\TranscodedWallpaper"
        SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, Transcoded_Wallpaper, SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)

        ' Check if "%AppData%\Microsoft\Windows\Themes\CachedFiles" folder is not exists
        If Not Directory.Exists(cachedFilesFolder) Then
            ' then create CachedFiles folder
            Directory.CreateDirectory(cachedFilesFolder)
        End If

        ' If have access to CachedFiles folder change the owner to SYSTEM and remove inheritance
        ' Start the process
        process.Start()
        ' Run the icacls command to get the owner
        process.StandardInput.WriteLine("icacls " & cachedFilesFolder)
        process.StandardInput.Close()
        ' Read the output
        Dim output As String = process.StandardOutput.ReadToEnd()
        process.WaitForExit()
        ' Check if the current user or Administrator is the owner
        Dim currentUser As String = Environment.UserName
        Dim isAdmin As Boolean = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator)
        If output.Contains(currentUser) OrElse isAdmin Then
            ' msgbox ("error2", , "E")
            ' If the current user or Administrator is the owner, run the icacls commands to change the owner to SYSTEM and remove inheritance
            process.Start()
            process.StandardInput.WriteLine("icacls " & cachedFilesFolder & " /setowner SYSTEM")
            process.StandardInput.WriteLine("icacls " & cachedFilesFolder & " /inheritance:r")
            process.StandardInput.Close()
            process.WaitForExit()
        End If


        ' Refresh the desktop wallpaper
        SHChangeNotify(SHCNE_ASSOCCHANGED, &H0, IntPtr.Zero, IntPtr.Zero)
    End Sub
End Module
