Imports DocumentFormat.OpenXml.Packaging

Module ExportData
    'https://stackoverflow.com/questions/41454836/vb-net-datatable-to-excel
    Public Sub ExportDataSet(ByVal DataTable_In As DataTable, ByVal Destination As String, Optional FileName As String = "ExcelFileName.xlsx", Optional ds As DataSet = Nothing)

        Try
            Using workbook = SpreadsheetDocument.Create(Destination & "\" & FileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                Dim workbookPart = workbook.AddWorkbookPart()
                workbook.WorkbookPart.Workbook = New DocumentFormat.OpenXml.Spreadsheet.Workbook()
                workbook.WorkbookPart.Workbook.Sheets = New DocumentFormat.OpenXml.Spreadsheet.Sheets()

                If Not DataTable_In Is Nothing Then

                    Dim sheetPart = workbook.WorkbookPart.AddNewPart(Of WorksheetPart)()
                    Dim sheetData = New DocumentFormat.OpenXml.Spreadsheet.SheetData()
                    sheetPart.Worksheet = New DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData)
                    Dim sheets As DocumentFormat.OpenXml.Spreadsheet.Sheets = workbook.WorkbookPart.Workbook.GetFirstChild(Of DocumentFormat.OpenXml.Spreadsheet.Sheets)()
                    Dim relationshipId As String = workbook.WorkbookPart.GetIdOfPart(sheetPart)
                    Dim sheetId As UInteger = 1

                    If sheets.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Sheet)().Count() > 0 Then
                        sheetId = sheets.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Sheet)().[Select](Function(s) s.SheetId.Value).Max() + 1
                    End If

                    Dim sheet As DocumentFormat.OpenXml.Spreadsheet.Sheet = New DocumentFormat.OpenXml.Spreadsheet.Sheet() With {
                    .Id = relationshipId,
                    .SheetId = sheetId,
                    .Name = DataTable_In.TableName
                 }
                    sheets.Append(sheet)
                    Dim headerRow As DocumentFormat.OpenXml.Spreadsheet.Row = New DocumentFormat.OpenXml.Spreadsheet.Row()
                    Dim columns As List(Of String) = New List(Of String)()

                    For Each column As System.Data.DataColumn In DataTable_In.Columns
                        columns.Add(column.ColumnName)
                        Dim cell As DocumentFormat.OpenXml.Spreadsheet.Cell = New DocumentFormat.OpenXml.Spreadsheet.Cell()
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                        cell.CellValue = New DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                        headerRow.AppendChild(cell)
                    Next

                    sheetData.AppendChild(headerRow)

                    For Each dsrow As System.Data.DataRow In DataTable_In.Rows
                        Dim newRow As DocumentFormat.OpenXml.Spreadsheet.Row = New DocumentFormat.OpenXml.Spreadsheet.Row()

                        For Each col As String In columns
                            Dim cell As DocumentFormat.OpenXml.Spreadsheet.Cell = New DocumentFormat.OpenXml.Spreadsheet.Cell()
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            cell.CellValue = New DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow(col).ToString())
                            newRow.AppendChild(cell)
                        Next

                        sheetData.AppendChild(newRow)
                    Next
                Else
                    For Each table As System.Data.DataTable In ds.Tables
                        Dim sheetPart = workbook.WorkbookPart.AddNewPart(Of WorksheetPart)()
                        Dim sheetData = New DocumentFormat.OpenXml.Spreadsheet.SheetData()
                        sheetPart.Worksheet = New DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData)
                        Dim sheets As DocumentFormat.OpenXml.Spreadsheet.Sheets = workbook.WorkbookPart.Workbook.GetFirstChild(Of DocumentFormat.OpenXml.Spreadsheet.Sheets)()
                        Dim relationshipId As String = workbook.WorkbookPart.GetIdOfPart(sheetPart)
                        Dim sheetId As UInteger = 1

                        If sheets.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Sheet)().Count() > 0 Then
                            sheetId = sheets.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Sheet)().[Select](Function(s) s.SheetId.Value).Max() + 1
                        End If

                        Dim sheet As DocumentFormat.OpenXml.Spreadsheet.Sheet = New DocumentFormat.OpenXml.Spreadsheet.Sheet() With {
                        .Id = relationshipId,
                        .SheetId = sheetId,
                        .Name = table.TableName
                     }
                        sheets.Append(sheet)
                        Dim headerRow As DocumentFormat.OpenXml.Spreadsheet.Row = New DocumentFormat.OpenXml.Spreadsheet.Row()
                        Dim columns As List(Of String) = New List(Of String)()

                        For Each column As System.Data.DataColumn In table.Columns
                            columns.Add(column.ColumnName)
                            Dim cell As DocumentFormat.OpenXml.Spreadsheet.Cell = New DocumentFormat.OpenXml.Spreadsheet.Cell()
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                            cell.CellValue = New DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName)
                            headerRow.AppendChild(cell)
                        Next

                        sheetData.AppendChild(headerRow)

                        For Each dsrow As System.Data.DataRow In table.Rows
                            Dim newRow As DocumentFormat.OpenXml.Spreadsheet.Row = New DocumentFormat.OpenXml.Spreadsheet.Row()

                            For Each col As String In columns
                                Dim cell As DocumentFormat.OpenXml.Spreadsheet.Cell = New DocumentFormat.OpenXml.Spreadsheet.Cell()
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                                cell.CellValue = New DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow(col).ToString())
                                newRow.AppendChild(cell)
                            Next

                            sheetData.AppendChild(newRow)
                        Next
                    Next
                End If


            End Using
        Catch ex As Exception

        End Try

    End Sub
End Module
