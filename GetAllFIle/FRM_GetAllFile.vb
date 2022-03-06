Imports System.IO

Public Class FRM_GetAllFile
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
            Getall(TextBox1.Text)
        End If


        'Dim baseDirectory As String = "F:\Document Control\"
        'Dim allFiles As String() = Directory.GetFiles(baseDirectory, "*.*", SearchOption.AllDirectories)
        'For Each file In allFiles
        '    ListBox1.Items.Add(file)
        'Next
    End Sub
    Public Sub Getall(FolderName As String)
        ''''https://docs.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=net-6.0
        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo(FolderName)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.*", SearchOption.AllDirectories)
        Dim fi As IO.FileInfo
        TestOrder()
        ClearDataGrid()
        Dim id As Integer = 0
        For Each fi In aryFi
            id = id + 1
            strFileSize = (Math.Round(fi.Length / 1024)).ToString()
            FileDirectory.Rows.Add(id, fi.DirectoryName, fi.FullName, fi.Name, strFileSize, fi.Extension, fi.LastAccessTime, (fi.Attributes.ReadOnly = True).ToString)

            'Console.WriteLine("File Name: {0}", fi.Name)
            'Console.WriteLine("File Full Name: {0}", fi.FullName)
            'Console.WriteLine("File Size (KB): {0}", strFileSize)
            'Console.WriteLine("File Extension: {0}", fi.Extension)
            'Console.WriteLine("Last Accessed: {0}", fi.LastAccessTime)
            'Console.WriteLine("Read Only: {0}", (fi.Attributes.ReadOnly = True).ToString)
        Next
        DataGridView1.DataSource = FileDirectory
    End Sub
    Public Sub ClearDataGrid()
        DataGridView1.CancelEdit()
        DataGridView1.Columns.Clear()
        DataGridView1.DataSource = Nothing
        '///  Button2 Edit File

    End Sub
    Dim FileDirectory As New DataTable
    Public Sub TestOrder()
        FileDirectory = New DataTable
        FileDirectory.Columns.Add(New DataColumn("id", GetType(Integer)))
        FileDirectory.Columns.Add(New DataColumn("DirectoryName", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("Full Directory FileName", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("File Name", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("File Size (KB)", GetType(Decimal)))
        FileDirectory.Columns.Add(New DataColumn("File Extension", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("Last Accessed", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("Read Only", GetType(String)))

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount() > 0 Then
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                ExportDataSet(FileDirectory, FolderBrowserDialog1.SelectedPath)
            End If
        Else
            MessageBox.Show("No Data Export", "error !!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
End Class
