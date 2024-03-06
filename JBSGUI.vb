Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics

Public Class SlideshowForm
    Inherits Form

    Private TextBoxFolderPath As New TextBox()
    Private ButtonBrowse As New Button()
    Private ButtonCreateList As New Button() ' New button for creating list
    Private WithEvents NumericUpDownCurrentBG As New NumericUpDown()
    Private ButtonPrevious As New Button()
    Private ButtonNext As New Button()
    Private ButtonBGName As New Button()
    Public Sub New()
        InitializeComponent()
        LoadInitialData()
    End Sub

    Private Sub InitializeComponent()
		Me.Size = New Size(560, 120)
		' Me.AllowDrop = True
		Me.Text = "JBS GUI - amymor OgomnamO"
		' Icon is not look good, but the exe is not depend on icon file near it anymore
		Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
		' The "VBCompilerUI.ico" should be near "VBCompilerUI.exe" to works
		' Me.Icon = New Icon("VBCompilerUI.ico")
		Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.Font = New Font("Segoe UI", 12) ' Set font size
		
        ' Set up form and controls
        TextBoxFolderPath.Location = New Point(10, 10)
        TextBoxFolderPath.Size = New Size(380, 30)

        NumericUpDownCurrentBG.Location = New Point(10, 50)
        NumericUpDownCurrentBG.Size = New Size(120, 20)
        NumericUpDownCurrentBG.Font = New Font("Segoe UI", 14) ' Set font size
        NumericUpDownCurrentBG.Value = 1
		
        ButtonBrowse.Location = New Point(390, 10)
        ButtonBrowse.Size = New Size(75, 30)
        ButtonBrowse.Text = "Browse"

        ButtonCreateList.Location = New Point(465, 10) ' Position the new button
        ButtonCreateList.Size = New Size(85, 30)
        ButtonCreateList.Font = New Font("Segoe UI", 10) ' Set font size
        ButtonCreateList.Text = "Create list"

        ButtonPrevious.Location = New Point(139, 50)
        ButtonPrevious.Size = New Size(60, 32)
        ' ButtonPrevious.Font = New Font("Segoe UI", 13) ' Set font size
        ButtonPrevious.Text = "◀️"
        ' ButtonPrevious.Text = "Previous"

        ButtonNext.Location = New Point(200, 50)
        ButtonNext.Size = New Size(60, 32)
        ' ButtonNext.Font = New Font("Segoe UI", 13) ' Set font size
        ' ButtonNext.Text = "Next"
        ButtonNext.Text = "▶️"
		
		ButtonBGName.Location = New Point(270, 50)
		ButtonBGName.Size = New Size(280, 32)
		ButtonBGName.TextAlign = ContentAlignment.MiddleLeft

        ' Dark theme =======================================================
        Me.BackColor = Color.FromArgb(40, 35, 40)
        Me.ForeColor = Color.FromArgb(240, 240, 240)

        TextBoxFolderPath.BackColor = Color.FromArgb(30, 30, 30)
        TextBoxFolderPath.ForeColor = Color.FromArgb(240, 240, 240)
        NumericUpDownCurrentBG.BackColor = Color.FromArgb(30, 30, 30)
        NumericUpDownCurrentBG.ForeColor = Color.FromArgb(240, 240, 240)

        ButtonBrowse.BackColor = Color.FromArgb(30, 30, 50)
        ButtonBrowse.ForeColor = Color.FromArgb(240, 240, 240)

        ButtonCreateList.BackColor = Color.FromArgb(30, 30, 50)
        ButtonCreateList.ForeColor = Color.FromArgb(240, 240, 160)

        ButtonPrevious.BackColor = Color.FromArgb(30, 30, 50)
        ButtonPrevious.ForeColor = Color.FromArgb(150, 220, 245)

        ButtonNext.BackColor = Color.FromArgb(30, 30, 50)
        ButtonNext.ForeColor = Color.FromArgb(150, 220, 245)

        ButtonBGName.BackColor = Color.FromArgb(30, 30, 50)
        ButtonBGName.ForeColor = Color.FromArgb(240, 240, 240)

        ' Flat style =======================================================
        TextBoxFolderPath.BorderStyle = BorderStyle.FixedSingle
        NumericUpDownCurrentBG.BorderStyle = BorderStyle.FixedSingle

        ButtonBrowse.FlatStyle = FlatStyle.Flat
        ButtonBrowse.FlatAppearance.BorderColor = Color.FromArgb(84, 70, 177)
        ButtonBrowse.FlatAppearance.BorderSize = 1

        ButtonCreateList.FlatStyle = FlatStyle.Flat
        ButtonCreateList.FlatAppearance.BorderColor = Color.FromArgb(84, 70, 177)
        ButtonCreateList.FlatAppearance.BorderSize = 1

        ButtonPrevious.FlatStyle = FlatStyle.Flat
        ' ButtonPrevious.FlatAppearance.BorderColor = Color.FromArgb(160,70,160)
        ButtonPrevious.FlatAppearance.BorderColor = Color.FromArgb(84, 70, 177)
        ButtonPrevious.FlatAppearance.BorderSize = 1

        ButtonNext.FlatStyle = FlatStyle.Flat
        ' ButtonNext.FlatAppearance.BorderColor = Color.FromArgb(160,70,160)
        ButtonNext.FlatAppearance.BorderColor = Color.FromArgb(84, 70, 177)
        ButtonNext.FlatAppearance.BorderSize = 1

		ButtonBGName.FlatStyle = FlatStyle.Flat
        ' ButtonBGName.FlatAppearance.BorderColor = Color.FromArgb(160,70,160)
		ButtonBGName.FlatAppearance.BorderColor = Color.FromArgb(84, 70, 177)
		ButtonBGName.FlatAppearance.BorderSize = 1

' AddHandler=====================================================
        AddHandler TextBoxFolderPath.TextChanged, AddressOf TextBoxFolderPath_TextChanged ' Handle text
        AddHandler ButtonBrowse.Click, AddressOf ButtonBrowse_Click
        AddHandler ButtonCreateList.Click, AddressOf ButtonCreateList_Click
        AddHandler ButtonPrevious.Click, AddressOf ButtonPrevious_Click
        AddHandler ButtonNext.Click, AddressOf ButtonNext_Click
		AddHandler ButtonBGName.Click, AddressOf ButtonBGName_Click

        ' AddHandler Me.DragEnter, AddressOf Form_DragEnter
        ' AddHandler Me.DragDrop, AddressOf Form_DragDrop
' Controls-add=====================================================
        Me.Controls.Add(TextBoxFolderPath)
        Me.Controls.Add(ButtonBrowse)
        Me.Controls.Add(ButtonCreateList)
        Me.Controls.Add(NumericUpDownCurrentBG)
        Me.Controls.Add(ButtonPrevious)
        Me.Controls.Add(ButtonNext)
		Me.Controls.Add(ButtonBGName)
    End Sub
' DragDrop=====================================================
    ' Private Sub Form_DragEnter(sender As Object, e As DragEventArgs) ' Check if the dragged item is a folder
        ' If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            ' If files.Length > 0 AndAlso Directory.Exists(files(0)) Then
                ' e.Effect = DragDropEffects.Copy
            ' End If
        ' End If
    ' End Sub

    ' Private Sub Form_DragDrop(sender As Object, e As DragEventArgs) ' Update the text box and FolderPath.ini with the folder path
        ' Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        ' If files.Length > 0 AndAlso Directory.Exists(files(0)) Then
            ' TextBoxFolderPath.Text = files(0)
            ' File.WriteAllText("FolderPath.ini", files(0).Trim())
        ' End If
    ' End Sub
' live Text-box and Numeric-field=====================================================

    Private Sub LoadInitialData()
        ' Load initial folder path from FolderPath.ini
        If File.Exists("FolderPath.ini") Then
            TextBoxFolderPath.Text = File.ReadAllText("FolderPath.ini").Trim()
        End If
        ' Load current background number from CurrentBG*.ini
        Dim currentBGFile = Directory.GetFiles(".").FirstOrDefault(Function(f) Path.GetFileName(f).StartsWith("CurrentBG"))
        If currentBGFile IsNot Nothing Then
            Dim currentBGNumber = Integer.Parse(Path.GetFileNameWithoutExtension(currentBGFile).Substring("CurrentBG".Length))
            NumericUpDownCurrentBG.Maximum = Decimal.MaxValue ' Remove the limit or set a specific limit
            NumericUpDownCurrentBG.Value = currentBGNumber
            NumericUpDownCurrentBG.Minimum = 1 ' Minimum value
            ButtonBrowse.Select() ' focus to the ButtonBrowse at startup
        End If	
    End Sub

    Private Sub TextBoxFolderPath_TextChanged(sender As Object, e As EventArgs)
        ' Update FolderPath.ini with the new text box value
        File.WriteAllText("FolderPath.ini", TextBoxFolderPath.Text.Trim())
    End Sub
' Browse=====================================================
Private Sub ButtonBrowse_Click(sender As Object, e As EventArgs)
    ' Initialize the OpenFileDialog
    Using openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Select a Folder"
        openFileDialog.CheckFileExists = False
        openFileDialog.CheckPathExists = True
        openFileDialog.ValidateNames = False
        openFileDialog.FileName = "Folder Selection"
        openFileDialog.Filter = "Folders|*.folder"
        openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer)
        openFileDialog.RestoreDirectory = True

        ' Show the OpenFileDialog
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Get the selected folder path
            Dim selectedFolderPath As String = Path.GetDirectoryName(openFileDialog.FileName)
            TextBoxFolderPath.Text = selectedFolderPath
            File.WriteAllText("FolderPath.ini", selectedFolderPath)
        End If
    End Using
End Sub
' CreateList=====================================================
Private Sub ButtonCreateList_Click(sender As Object, e As EventArgs)
    ' Execute ListFiles.exe and redirect output to List.ini
    Dim folderPath As String = TextBoxFolderPath.Text.Trim()
    If Not String.IsNullOrEmpty(folderPath) Then
        ' Ensure the folder path is properly quoted
        Dim quotedFolderPath As String = String.Format("""{0}""", folderPath)
        Dim startInfo As New ProcessStartInfo("ListFiles.exe", quotedFolderPath)
        startInfo.RedirectStandardOutput = True
        startInfo.UseShellExecute = False
        startInfo.CreateNoWindow = True
        Using process As New Process()
            process.StartInfo = startInfo
            process.Start()
            Dim output As String = process.StandardOutput.ReadToEnd()
            process.WaitForExit()
            File.WriteAllText("List.ini", output)
        End Using
    End If
End Sub

' UpdateBackgroundNumber=====================================================
Private Sub ButtonPrevious_Click(sender As Object, e As EventArgs)
    ' Check if the List.ini file exists
    If File.Exists("List.ini") Then
        ' Determine the total number of lines in List.ini
        Dim totalLines As Integer = File.ReadAllLines("List.ini").Length
        If totalLines > 0 Then
            ' Check if the current value is at the minimum
            If NumericUpDownCurrentBG.Value = 1 Then
                ' Set to the last line
                NumericUpDownCurrentBG.Value = totalLines
            Else
                ' Decrement the current background number
                NumericUpDownCurrentBG.Value -= 1
            End If
            UpdateBackgroundNumber()
        End If
    Else
        ' Optionally, handle the case where the List.ini file does not exist
        ' MessageBox.Show("List.ini does not exist.")
        If NumericUpDownCurrentBG.Value > 1 Then
        NumericUpDownCurrentBG.Value -= 1
    End If
        ButtonBGName.Text = "List.ini does not exist." ' Update the button's text
    End If
End Sub

' ---------------------------------------------------------------------------------------------------------------------------
Private Sub ButtonNext_Click(sender As Object, e As EventArgs)
    ' Check if the List.ini file exists
    If File.Exists("List.ini") Then
        ' Determine the total number of lines in List.ini
        Dim totalLines As Integer = File.ReadAllLines("List.ini").Length
        If totalLines > 0 Then
            ' Check if the current value is at the maximum
            If NumericUpDownCurrentBG.Value = totalLines Then
                ' Reset to the first line
                NumericUpDownCurrentBG.Value = 1
            Else
                ' Increment the current background number
                NumericUpDownCurrentBG.Value += 1
            End If
            UpdateBackgroundNumber()
        End If
    Else
        ' Optionally, handle the case where the List.ini file does not exist
        ' MessageBox.Show("List.ini does not exist.")
        NumericUpDownCurrentBG.Value += 1
        ButtonBGName.Text = "List.ini does not exist." ' Update the button's text
    End If
End Sub
' ---------------------------------------------------------------------------------------------------------------------------
Private Sub NumericUpDownCurrentBG_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDownCurrentBG.ValueChanged
    ' Check if the List.ini file exists
    If File.Exists("List.ini") Then
        ' Determine the total number of lines in List.ini
        Dim totalLines As Integer = File.ReadAllLines("List.ini").Length
        If totalLines > 0 Then
            ' Update the Maximum property to match the total number of lines
            NumericUpDownCurrentBG.Maximum = totalLines

            ' Implement wrapping behavior
            If NumericUpDownCurrentBG.Value > totalLines Then
                ' Wrap to the first line
                NumericUpDownCurrentBG.Value = 1
            ElseIf NumericUpDownCurrentBG.Value < 1 Then
                ' Wrap to the last line
                NumericUpDownCurrentBG.Value = totalLines
            End If
        End If
    Else
        ' Optionally, handle the case where the List.ini file does not exist
        ' MessageBox.Show("List.ini does not exist.")
        ButtonBGName.Text = "List.ini does not exist." ' Update the button's text
    End If
    ' Update the background number
    UpdateBackgroundNumber()
End Sub
' ---------------------------------------------------------------------------------------------------------------------------
    Private Sub UpdateBackgroundNumber()
        ' Update the current background number in CurrentBG*.ini
        Dim currentBGNumber = NumericUpDownCurrentBG.Value
        Dim currentBGFile = Directory.GetFiles(".").FirstOrDefault(Function(f) Path.GetFileName(f).StartsWith("CurrentBG"))
        If currentBGFile IsNot Nothing Then
            File.Delete(currentBGFile)
        End If
        File.WriteAllText("CurrentBG" & currentBGNumber & ".ini", "")

' Set Background=====================================================
    ' Set the desktop background using JBS.exe
    Dim imageName = GetImageNameFromList(currentBGNumber)
    If Not String.IsNullOrEmpty(imageName) Then
        ButtonBGName.Text = Path.GetFileName(imageName) ' Update the button's text
        ' Construct the full path to the image, ensuring it's properly quoted
        Dim fullPath As String = String.Format("""{0}\{1}""", TextBoxFolderPath.Text, imageName)
        Process.Start("JBS.exe", fullPath)
    ' Else
        ' ButtonBGName.Text = "No Image Selected" ' Update the button's text
    End If
End Sub

Private Function GetImageNameFromList(currentBGNumber As Integer) As String
    ' Check if the List.ini file exists
    If File.Exists("List.ini") Then
        Dim lines = File.ReadAllLines("List.ini")
        If lines.Length >= currentBGNumber Then
            Return lines(currentBGNumber - 1)
        End If
    End If
    ' Return an empty string or a default message if the file does not exist or the index is out of range
    Return String.Empty ' Or you can return a default message like "No Image Selected"
End Function
' ButtonBGName=====================================================
Private Sub ButtonBGName_Click(sender As Object, e As EventArgs)
    Dim folderPath As String = TextBoxFolderPath.Text.Trim()
    Dim currentBGNumber = NumericUpDownCurrentBG.Value
    Dim imageName = GetImageNameFromList(currentBGNumber)
    If Not String.IsNullOrEmpty(folderPath) AndAlso Not String.IsNullOrEmpty(imageName) Then
        ' Construct the full path to the image
        Dim fullPath As String = Path.Combine(folderPath, imageName)
        ' Construct the command line argument for Explorer using String.Format
        Dim explorerArgs As String = String.Format("/select, ""{0}""", fullPath)
        ' Start Explorer with the command line argument
        Process.Start("explorer.exe", explorerArgs)
    End If
End Sub

' =====================================================
    ' Other existing event handlers...
End Class


Public Module Program
    Public Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New SlideshowForm())
    End Sub
End Module
