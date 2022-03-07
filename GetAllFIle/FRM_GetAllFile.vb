Imports System.IO

Public Class FRM_GetAllFile
  
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
     Try

            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                TextBox1.Text = FolderBrowserDialog1.SelectedPath
                Getall(TextBox1.Text)
            End If
            Catch

        End Try

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
        Dim GetAllFiles As IO.FileInfo() = di.GetFiles("*.*", SearchOption.AllDirectories)
        Dim xFileInfo As IO.FileInfo
        TestOrder()
        ClearDataGrid()
        Dim id As Integer = 0
        For Each xFileInfo In GetAllFiles
            id = id + 1
        strFileSize = FormatBytes(  xFileInfo.Length  )
            FileDirectory.Rows.Add(id, xFileInfo.DirectoryName, xFileInfo.FullName, xFileInfo.Name, strFileSize, xFileInfo.Extension, xFileInfo.LastAccessTime, (xFileInfo.Attributes.ReadOnly = True).ToString)

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

Dim DoubleBytes As Double
Public Function FormatBytes(ByVal BytesCaller As ULong) As String

    Try
        Select Case BytesCaller
            Case Is >= 1099511627776
                DoubleBytes = CDbl(BytesCaller / 1099511627776) 'TB
                Return FormatNumber(DoubleBytes, 2) & " TB"
            Case 1073741824 To 1099511627775
                DoubleBytes = CDbl(BytesCaller / 1073741824) 'GB
                Return FormatNumber(DoubleBytes, 2) & " GB"
            Case 1048576 To 1073741823
                DoubleBytes = CDbl(BytesCaller / 1048576) 'MB
                Return FormatNumber(DoubleBytes, 2) & " MB"
            Case 1024 To 1048575
                DoubleBytes = CDbl(BytesCaller / 1024) 'KB
                Return FormatNumber(DoubleBytes, 2) & " KB"
            Case 0 To 1023
                DoubleBytes = BytesCaller ' bytes
                Return FormatNumber(DoubleBytes, 2) & " bytes"
            Case Else
                Return ""
        End Select
    Catch
        Return ""
    End Try

End Function
    Dim FileDirectory As New DataTable
    Public Sub TestOrder()
        FileDirectory = New DataTable
        FileDirectory.Columns.Add(New DataColumn("id", GetType(Integer)))
        FileDirectory.Columns.Add(New DataColumn("DirectoryName", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("Full Directory FileName", GetType(String)))
        FileDirectory.Columns.Add(New DataColumn("File Name", GetType(String)))
    FileDirectory.Columns.Add(New DataColumn("File Size ", GetType(String)))
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
