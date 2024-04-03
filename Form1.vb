Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form1

    Private WithEvents TrayIcon As New NotifyIcon()
    Private WithEvents fileSaveWorker As New System.ComponentModel.BackgroundWorker

    Public Sub New()
        InitializeComponent()

        ' Initialize BackgroundWorker
        fileSaveWorker.WorkerSupportsCancellation = True
        fileSaveWorker.WorkerReportsProgress = False
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Load the saved file path from settings
        If My.Settings.FilePath IsNot Nothing AndAlso File.Exists(My.Settings.FilePath) Then
            LoadFile(My.Settings.FilePath)
        Else
            ' If there's no saved file path, open the file browse dialog
            BrowseToolStripMenuItem_Click(Nothing, Nothing)
        End If

        ' Set the icon for the NotifyIcon
        TrayIcon.Icon = Me.Icon

        ' Add a context menu to the NotifyIcon
        Dim contextMenu As New ContextMenuStrip()
        contextMenu.Items.Add("Exit", Nothing, AddressOf ExitToolStripMenuItem_Click)
        TrayIcon.ContextMenuStrip = contextMenu

        Dim textBoxContextMenu As New ContextMenuStrip()

        ' Existing menu items
        textBoxContextMenu.Items.Add("Refresh", Nothing, AddressOf RefreshToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("-")
        textBoxContextMenu.Items.Add("Copy", Nothing, AddressOf CopyToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("Paste", Nothing, AddressOf PasteToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("-")
        textBoxContextMenu.Items.Add("Undo", Nothing, AddressOf UndoToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("Select All", Nothing, AddressOf SelectAllToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("Copy Line", Nothing, AddressOf CopyLineToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("-")
        textBoxContextMenu.Items.Add("Source File", Nothing, AddressOf BrowseToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("Change Font", Nothing, AddressOf FontToolStripMenuItem_Click)
        textBoxContextMenu.Items.Add("-")

        ' New menu item for DarkMode
        Dim darkModeMenuItem As New ToolStripMenuItem("Dark Mode")
        darkModeMenuItem.CheckOnClick = True
        AddHandler darkModeMenuItem.CheckedChanged, AddressOf DarkModeMenuItem_CheckedChanged
        textBoxContextMenu.Items.Add(darkModeMenuItem)

        ' Menu items for Always on Top and Run at Startup
        Dim alwaysOnTopMenuItem As New ToolStripMenuItem("Always on Top")
        alwaysOnTopMenuItem.CheckOnClick = True
        AddHandler alwaysOnTopMenuItem.CheckedChanged, AddressOf AlwaysOnTopMenuItem_CheckedChanged
        textBoxContextMenu.Items.Add(alwaysOnTopMenuItem)

        Dim autoRunMenuItem As New ToolStripMenuItem("Run at Startup")
        autoRunMenuItem.CheckOnClick = True
        autoRunMenuItem.Checked = GetIsApplicationSetToRunAtStartup()
        AddHandler autoRunMenuItem.CheckedChanged, AddressOf AutoRunMenuItem_CheckedChanged
        textBoxContextMenu.Items.Add(autoRunMenuItem)

        ' Set the context menu to TextBox1
        TextBox1.ContextMenuStrip = textBoxContextMenu

        ' Load AutoRun setting
        autoRunMenuItem.Checked = My.Settings.AutoRun

        ' Load Always on Top setting
        alwaysOnTopMenuItem.Checked = My.Settings.OnTop

    End Sub
    Private Sub DarkModeMenuItem_CheckedChanged(sender As Object, e As EventArgs)
        Dim menuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If menuItem.Checked Then
            ' Dark mode enabled
            TextBox1.ForeColor = Color.White
            TextBox1.BackColor = Color.FromArgb(31, 31, 31) ' Very dark gray
            My.Settings.DarkMode = True ' Save DarkMode setting
        Else
            ' Dark mode disabled
            TextBox1.ForeColor = Color.Black
            TextBox1.BackColor = Color.White
            My.Settings.DarkMode = False ' Save DarkMode setting
        End If
    End Sub

    Private Sub CopyLineToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Ensure there is text in the textbox
        If TextBox1.Text.Length > 0 Then
            ' Get the current selection start point
            Dim currentPos As Integer = TextBox1.SelectionStart

            ' Find the start of the current line
            Dim startOfLine As Integer = currentPos
            While startOfLine > 0 AndAlso TextBox1.Text(startOfLine - 1) <> vbCr AndAlso TextBox1.Text(startOfLine - 1) <> vbLf
                startOfLine -= 1
            End While

            ' Find the end of the current line
            Dim endOfLine As Integer = currentPos
            While endOfLine < TextBox1.TextLength AndAlso TextBox1.Text(endOfLine) <> vbCr AndAlso TextBox1.Text(endOfLine) <> vbLf
                endOfLine += 1
            End While

            ' Select and copy the current line
            TextBox1.Select(startOfLine, endOfLine - startOfLine)
            Clipboard.SetText(TextBox1.SelectedText)
        End If
    End Sub

    Private Function GetIsApplicationSetToRunAtStartup() As Boolean
        Using key = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", False)
            Dim value As Object = key.GetValue(Application.ProductName)
            Return value IsNot Nothing
        End Using
    End Function

    Private Sub AutoRunMenuItem_CheckedChanged(sender As Object, e As EventArgs)
        Dim menuItem = CType(sender, ToolStripMenuItem)
        My.Settings.AutoRun = menuItem.Checked
        My.Settings.Save()
        If menuItem.Checked Then
            SetApplicationToRunAtStartup(True)
        Else
            SetApplicationToRunAtStartup(False)
        End If

        ' Save the setting
    End Sub

    Private Sub SetApplicationToRunAtStartup(enable As Boolean)
        Using key = My.Computer.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Run", True)
            If enable Then
                key.SetValue(Application.ProductName, """" & Application.ExecutablePath & """")
            Else
                key.DeleteValue(Application.ProductName, False)
            End If
        End Using
    End Sub

    Private Sub AlwaysOnTopMenuItem_CheckedChanged(sender As Object, e As EventArgs)
        ' Update the application's TopMost property based on the menu item state
        Me.TopMost = DirectCast(sender, ToolStripMenuItem).Checked

        ' Save the Always on Top setting
        My.Settings.OnTop = Me.TopMost

        ' Save the setting
        My.Settings.Save()
    End Sub

    Private Sub LoadFile(filePath As String)
        Try
            TextBox1.Text = File.ReadAllText(filePath)
        Catch ex As Exception
            MessageBox.Show("Error loading file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private WithEvents textChangeTimer As New Timer()
    Private WithEvents saveTimer As New Timer()
    Private Const SaveIntervalSeconds As Integer = 2
    Private Const TimeoutSeconds As Integer = 60 ' 1 minutes
    Private timeoutCounter As Integer = 0

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ' Reset the timeout counter
        timeoutCounter = 0

        ' Start or reset the text change timer
        textChangeTimer.Interval = 2000 ' 2 seconds
        textChangeTimer.Stop()
        textChangeTimer.Start()
    End Sub

    Private Sub textChangeTimer_Tick(sender As Object, e As EventArgs) Handles textChangeTimer.Tick
        ' Stop the text change timer
        textChangeTimer.Stop()

        ' Start the save timer
        saveTimer.Interval = SaveIntervalSeconds * 1000 ' Convert seconds to milliseconds
        saveTimer.Start()
    End Sub

    Private Sub saveTimer_Tick(sender As Object, e As EventArgs) Handles saveTimer.Tick
        ' Start a new file saving operation
        If My.Settings.FilePath IsNot Nothing Then
            ' Do the saving operation here
            fileSaveWorker.RunWorkerAsync(TextBox1.Text)

            ' Increment timeout counter
            timeoutCounter += SaveIntervalSeconds

            ' Check if timeout has reached
            If timeoutCounter >= TimeoutSeconds Then
                ' Timeout reached, stop saving
                saveTimer.Stop()
            End If
        End If
    End Sub

    Private Sub fileSaveWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles fileSaveWorker.DoWork
        Dim text As String = DirectCast(e.Argument, String)

        Try
            File.WriteAllText(My.Settings.FilePath, text)
        Catch ex As Exception
            MessageBox.Show("Error saving file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ' Minimize to system tray when form is minimized
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
            TrayIcon.Visible = True
        End If
    End Sub

    Private Sub TrayIcon_MouseClick(sender As Object, e As MouseEventArgs) Handles TrayIcon.MouseClick
        ' Check if the left mouse button is clicked
        If e.Button = MouseButtons.Left Then
            If Me.WindowState = FormWindowState.Normal Then
                ' If the form is currently shown, minimize it
                Me.WindowState = FormWindowState.Minimized
            Else
                ' Otherwise, restore the form
                Me.Show()
                Me.WindowState = FormWindowState.Normal
            End If
        End If
    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Clean up resources and exit the application
        TrayIcon.Dispose()
        Application.Exit()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Reload the contents of the source file into the textbox
        LoadFile(My.Settings.FilePath)
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Copy selected text
        TextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Paste text from clipboard
        TextBox1.Paste()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Undo the operation
        TextBox1.Undo()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ' Assuming textBox is the name of your TextBox control
        TextBox1.SelectAll()
    End Sub

    Private Sub BrowseToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim openFileDialog As New OpenFileDialog()

        openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        openFileDialog.FilterIndex = 1
        openFileDialog.RestoreDirectory = True

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            My.Settings.FilePath = filePath
            LoadFile(filePath)
        End If
    End Sub

    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim fontDialog As New FontDialog()
        fontDialog.Font = TextBox1.Font

        If fontDialog.ShowDialog() = DialogResult.OK Then
            TextBox1.Font = fontDialog.Font
            My.Settings.TextBoxFont = fontDialog.Font
            My.Settings.Save()
        End If
    End Sub
    Private Sub TextBox1_MouseWheel(sender As Object, e As MouseEventArgs) Handles TextBox1.MouseWheel
        If e.Delta > 0 Then
            ' Scroll up
            SendKeys.Send("{UP}")
        Else
            ' Scroll down
            SendKeys.Send("{DOWN}")
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            ' Minimize to system tray instead of closing
            Me.WindowState = FormWindowState.Minimized
            e.Cancel = True
        End If
    End Sub

End Class
