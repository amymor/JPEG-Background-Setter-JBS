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
    Private userInitiatedChange As Boolean = False

    Public Sub New()
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
		NumericUpDownCurrentBG.Maximum = 10000000  ' Remove the limit
		NumericUpDownCurrentBG.Minimum = 1 ' Minimum value
		
        ButtonBrowse.Location = New Point(390, 10)
        ButtonBrowse.Size = New Size(75, 30)
        ButtonBrowse.Text = "Browse"
        ButtonBrowse.Font = New Font("Segoe UI", 11) ' Set font size

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
        ButtonBGName.Font = New Font("Segoe UI", 10) ' Set font size

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

' = startup TextBoxFolderPath as FolderPath.ini=====================================================
        If File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FolderPath.ini")) Then
            TextBoxFolderPath.Text = File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FolderPath.ini")).Trim()
        End If
' = startup Numeric as CurrentBG*.ini and ButtonBGName as image Name ============================
        ' Load current background number from CurrentBG*.ini
    Dim currentBGFile = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory).FirstOrDefault(Function(f) Path.GetFileName(f).StartsWith("CurrentBG"))
    If currentBGFile IsNot Nothing Then
        Dim currentBGNumber = Integer.Parse(Path.GetFileNameWithoutExtension(currentBGFile).Substring("CurrentBG".Length))
        NumericUpDownCurrentBG.Value = currentBGNumber

    ' Initialize ButtonBGName.Text with the name of the current background image
    Dim imageName = GetImageNameFromList(currentBGNumber)
    If Not String.IsNullOrEmpty(imageName) Then
        ButtonBGName.Text = Path.GetFileName(imageName) ' Update the button's text
	Else
        ButtonBGName.Text = "List.ini is empty or does not exist." ' Update the button's text
    End If
	Else
        Dim imageName = GetImageNameFromList("1")
        ButtonBGName.Text = Path.GetFileName(imageName) ' Update the button's text
        ' ButtonBGName.Text = "1. Browse > 2. Create List"
    End If
' AddHandler=====================================================
        AddHandler TextBoxFolderPath.TextChanged, AddressOf TextBoxFolderPath_TextChanged ' Handle text
        AddHandler ButtonBrowse.Click, AddressOf ButtonBrowse_Click
        AddHandler ButtonCreateList.Click, AddressOf ButtonCreateList_Click
        AddHandler ButtonPrevious.Click, AddressOf ButtonPrevious_Click
        AddHandler ButtonNext.Click, AddressOf ButtonNext_Click
        AddHandler NumericUpDownCurrentBG.ValueChanged, AddressOf NumericUpDownCurrentBG_ValueChanged
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
		
		ButtonBrowse.Select() ' focus to the ButtonBrowse at startup
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

' live Text-box =====================================================

Private Sub TextBoxFolderPath_TextChanged(sender As Object, e As EventArgs)
    File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FolderPath.ini"), TextBoxFolderPath.Text.Trim())
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
            ' File.WriteAllText("FolderPath.ini", selectedFolderPath)
			File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FolderPath.ini"), selectedFolderPath)
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
        Dim startInfo As New ProcessStartInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ListFiles.exe"), quotedFolderPath)
        startInfo.RedirectStandardOutput = True
        startInfo.UseShellExecute = False
        startInfo.CreateNoWindow = True
        Using process As New Process()
            process.StartInfo = startInfo
            process.Start()
            Dim output As String = process.StandardOutput.ReadToEnd()
            process.WaitForExit()

            ' Check if the output is not empty
            If Not String.IsNullOrWhiteSpace(output) Then
                ' Write the output to List.ini
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini"), output)

        
        ' Rename the existing CurrentBG file to the new number
                Dim currentBGFile = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory).FirstOrDefault(Function(f) Path.GetFileName(f).StartsWith("CurrentBG"))
                If currentBGFile IsNot Nothing Then
					File.Move(currentBGFile, "CurrentBG1")
                Else
					' If there's no existing CurrentBG file, create a new CurrentBG1
					File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CurrentBG1"), "")
                End If
	
                UpdateBackgroundNumber()
            Else
                ' If the output is empty, update the button's text to indicate no images were found
                ButtonBGName.Text = "The folder did not contain any images."
            End If
        End Using
    End If
End Sub


' UpdateBackgroundNumber=====================================================
Private Sub ButtonPrevious_Click(sender As Object, e As EventArgs)
userInitiatedChange = True
    ' Check if the List.ini file exists
    If File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")) Then
        ' Determine the total number of lines in List.ini
        Dim totalLines As Integer = File.ReadAllLines(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")).Length
        If totalLines > 0 Then
            ' Check if the current value is at the minimum
            If NumericUpDownCurrentBG.Value > 1 Then
                ' Decrement the current background number
                NumericUpDownCurrentBG.Value -= 1
            Else
                ' Set to the last line
                NumericUpDownCurrentBG.Value = totalLines
            End If
            UpdateBackgroundNumber()
        End If
    Else
        ' Handle the case where the List.ini file does not exist
        ButtonBGName.Text = "List.ini is empty or does not exist." ' Update the button's text
    End If
End Sub

' ---------------------------------------------------------------------------------------------------------------------------
Private Sub ButtonNext_Click(sender As Object, e As EventArgs)
userInitiatedChange = True
    ' Check if the List.ini file exists
    If File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")) Then
        ' Determine the total number of lines in List.ini
        Dim totalLines As Integer = File.ReadAllLines(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")).Length
        If totalLines > 0 Then
            ' Check if the current value is at the maximum
            If NumericUpDownCurrentBG.Value < totalLines Then
                ' Increment the current background number
                NumericUpDownCurrentBG.Value += 1
            Else
                ' Reset to the first line
                NumericUpDownCurrentBG.Value = 1
            End If
            UpdateBackgroundNumber()
        End If
    Else
        ' Handle the case where the List.ini file does not exist
        ButtonBGName.Text = "List.ini is empty or does not exist." ' Update the button's text
    End If
End Sub

' ---------------------------------------------------------------------------------------------------------------------------
 Private Sub NumericUpDownCurrentBG_ValueChanged(sender As Object, e As EventArgs)
    If userInitiatedChange Then
    ' Check if the List.ini file exists
    If File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")) Then
        ' Determine the total number of lines in List.ini
        Dim totalLines As Integer = File.ReadAllLines(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")).Length
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
        ' Handle the case where the List.ini file does not exist
        ButtonBGName.Text = "List.ini is empty or does not exist." ' Update the button's text
    End If

    ' Update the background number immediately
    UpdateBackgroundNumber()
    End If
    ' set the flag to True to make the NumericUpDown live
    userInitiatedChange = True
End Sub

' ---------------------------------------------------------------------------------------------------------------------------
Private Sub UpdateBackgroundNumber()
    Dim totalLines As Integer = File.ReadAllLines(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")).Length
    If totalLines > 0 Then
        ' Update the Maximum property to match the total number of lines
        NumericUpDownCurrentBG.Maximum = totalLines
    End If
    ' Update the current background number in CurrentBG*
    Dim currentBGNumber = NumericUpDownCurrentBG.Value
    Dim currentBGFile = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory).FirstOrDefault(Function(f) Path.GetFileName(f).StartsWith("CurrentBG"))
    
    If currentBGFile IsNot Nothing Then
        ' Construct the new file name based on the current background number
        Dim newBGFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CurrentBG" & currentBGNumber)
        
        ' Rename the existing CurrentBG file to the new number
        File.Move(currentBGFile, newBGFileName)
    Else
        ' If there's no existing CurrentBG file, create a new one with the current background number
        File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CurrentBG" & currentBGNumber), "")
    End If


    ' Set the desktop background using JBS.exe
    Dim imageName = GetImageNameFromList(currentBGNumber)
    If Not String.IsNullOrEmpty(imageName) Then
        ButtonBGName.Text = Path.GetFileName(imageName) ' Update the button's text
        ' Construct the full path to the image, ensuring it's properly quoted
        Dim fullPath As String = String.Format("""{0}\{1}""", TextBoxFolderPath.Text, imageName)
        Process.Start(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "JBS.exe"), fullPath)
    ' Else
        ' ButtonBGName.Text = "No Image Selected" ' Update the button's text
    End If
End Sub
' ---------------------------------------------------------------------------------------------------------------------------
Private Function GetImageNameFromList(currentBGNumber As Integer) As String
    Dim listFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "List.ini")
    ' Check if the List.ini file exists
    If File.Exists(listFilePath) Then
        Dim lines = File.ReadAllLines(listFilePath)
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

    Public Shared Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New SlideshowForm())
    End Sub
End Class