Imports System
Imports System.IO
Imports System.Runtime.InteropServices

Module ListFiles
    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Function StrCmpLogicalW(ByVal x As String, ByVal y As String) As Integer
    End Function

    Sub Main(args As String())
        If args.Length = 0 Then
            Console.WriteLine("Usage: ListFiles.exe <directory>")
            Return
        End If

        Dim directoryPath As String = args(0)
        If Not Directory.Exists(directoryPath) Then
            Console.WriteLine("Directory does not exist: " & directoryPath)
            Return
        End If

        Try
            Dim files As String() = Directory.GetFiles(directoryPath)
            ' Filter files by specific extensions
            files = Array.FindAll(files, Function(f) IsValidExtension(f))
            Array.Sort(files, New Comparison(Of String)(AddressOf NaturalSort))
            For Each file In files
                Console.WriteLine(Path.GetFileName(file))
            Next
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    Private Function NaturalSort(x As String, y As String) As Integer
        Return StrCmpLogicalW(x, y)
    End Function

    Private Function IsValidExtension(filePath As String) As Boolean
        Dim validExtensions As String() = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif"}
        Dim extension As String = Path.GetExtension(filePath).ToLower()
        Return validExtensions.Contains(extension)
    End Function
End Module
