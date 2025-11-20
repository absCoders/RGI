Imports System.Data
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ARCCUSTM

    ''' <summary>
    ''' Exports a DataTable to a new Excel workbook in a user-selected folder.
    ''' The worksheet is filled using a column map: DataTable column -> (Excel Header, NumberFormat).
    ''' </summary>
    ''' <param name="data">The DataTable to export.</param>
    ''' <param name="columnMap">
    ''' Dictionary(Of DataTableColumnName, KeyValuePair(Of ExcelHeaderTitle, ExcelNumberFormat))
    ''' Example format strings: "@", "mm/dd/yyyy", "#,##0.00", "0.000"
    ''' </param>
    ''' <param name="sheetName">Optional worksheet name (defaults to "Data").</param>
    ''' <param name="workbookBaseName">Optional base filename (defaults to "Export").</param>
    Public Shared Function ExportToNewWorkbook(ByVal data As DataTable,
                                          ByVal columnMap As Dictionary(Of String, KeyValuePair(Of String, String)),
                                          Optional ByVal sheetName As String = "Data",
                                          Optional ByVal workbookBaseName As String = "Export") As String

        Dim Retval As String = ""

        If data Is Nothing Then Throw New ArgumentNullException(NameOf(data))
        If columnMap Is Nothing OrElse columnMap.Count = 0 Then
            Throw New ArgumentException("columnMap must contain at least one mapping.", NameOf(columnMap))
        End If

        ' Pick folder
        Dim folderPath As String = PromptForFolder()
        If String.IsNullOrEmpty(folderPath) Then
            ' User cancelled
            Return Retval
        End If

        ' Build a deterministic, sensible write order:
        '   Use the DataTable column order, filtered to those present in the map.
        Dim orderedCols As New List(Of String)
        For Each col As DataColumn In data.Columns
            If columnMap.ContainsKey(col.ColumnName) Then
                orderedCols.Add(col.ColumnName)
            End If
        Next

        If orderedCols.Count = 0 Then
            Throw New InvalidOperationException("None of the DataTable columns are present in the columnMap.")
        End If

        Dim xlApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim ws As Excel.Worksheet = Nothing
        Dim headerRange As Excel.Range = Nothing
        Dim dataRange As Excel.Range = Nothing

        Try
            xlApp = New Excel.Application()
            xlApp.DisplayAlerts = False

            wb = xlApp.Workbooks.Add()
            ws = CType(wb.Sheets(1), Excel.Worksheet)
            ws.Name = CleanWorksheetName(sheetName)

            ' Write headers
            Dim colIndex As Integer = 1
            For Each dtColName As String In orderedCols
                Dim headerTitle As String = columnMap(dtColName).Key
                ws.Cells(1, colIndex).Value = headerTitle
                colIndex += 1
            Next

            ' Write data (row 2 onward)
            Dim rowCount As Integer = data.Rows.Count
            Dim colCount As Integer = orderedCols.Count

            If rowCount > 0 Then
                ' Build a 2D array for faster transfer
                Dim values(0 To rowCount - 1, 0 To colCount - 1) As Object
                For r As Integer = 0 To rowCount - 1
                    For c As Integer = 0 To colCount - 1
                        values(r, c) = If(IsDBNull(data.Rows(r)(orderedCols(c))), Nothing, data.Rows(r)(orderedCols(c)))
                    Next
                Next

                dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(1 + rowCount, colCount))
                dataRange.Value = values
            End If

            ' Header styling
            headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount))
            headerRange.Font.Bold = True
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            headerRange.WrapText = False
            headerRange.Interior.ColorIndex = 15 ' light gray fill

            ' Number formats per column
            For c As Integer = 1 To colCount
                Dim mapEntry = columnMap(orderedCols(c - 1))
                Dim fmt As String = If(mapEntry.Value, "").Trim()
                If fmt.Length > 0 Then
                    Dim colRng As Excel.Range = ws.Range(ws.Cells(2, c), ws.Cells(Math.Max(2, rowCount + 1), c))
                    colRng.NumberFormat = fmt
                    ReleaseComObjectSafe(colRng)
                End If
            Next

            ' Freeze top row, autofit, filter
            ws.Activate()
            ws.Range("A2").Select()
            xlApp.ActiveWindow.FreezePanes = True
            ws.UsedRange.Columns.AutoFit()
            headerRange.AutoFilter(1)

            ' Save workbook
            Dim fileName As String = String.Format("{0}_{1:yyyyMMdd_HHmmss}.xlsx", workbookBaseName, DateTime.Now)
            Dim savePath As String = Path.Combine(folderPath, fileName)
            Retval = savePath
            wb.SaveAs(Filename:=savePath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
            wb.Close(SaveChanges:=False)

            ' Optional: open it for the user
            ' xlApp.Workbooks.Open(savePath) ' uncomment if you want to open
            Return Retval
        Catch ex As COMException
            Throw New InvalidOperationException("Excel automation failed. Ensure Microsoft Excel is installed.", ex)
            Return Retval
        Finally
            ' Cleanup COM in reverse order
            ReleaseComObjectSafe(headerRange)
            ReleaseComObjectSafe(dataRange)
            ReleaseComObjectSafe(ws)
            If wb IsNot Nothing Then
                Try : wb.Close(False) : Catch : End Try
            End If
            ReleaseComObjectSafe(wb)

            If xlApp IsNot Nothing Then
                Try : xlApp.Quit() : Catch : End Try
            End If
            ReleaseComObjectSafe(xlApp)

            ' Encourage GC to collect COM wrappers
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    ' --- Helpers ---

    Private Shared Function PromptForFolder() As String
        Using dlg As New FolderBrowserDialog()
            dlg.Description = "Choose a folder to save the new Excel workbook"
            dlg.ShowNewFolderButton = True
            Dim result = dlg.ShowDialog()
            If result = DialogResult.OK Then
                Return dlg.SelectedPath
            End If
        End Using
        Return Nothing
    End Function

    Private Shared Function CleanWorksheetName(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then name = "Data"
        Dim invalidChars As Char() = New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
        For Each ch As Char In invalidChars
            name = name.Replace(ch, "-"c)
        Next
        If name.Length > 31 Then name = name.Substring(0, 31)
        If String.IsNullOrWhiteSpace(name) Then name = "Data"
        Return name
    End Function

    Private Shared Sub ReleaseComObjectSafe(ByVal comObj As Object)
        If comObj IsNot Nothing AndAlso Marshal.IsComObject(comObj) Then
            Try
                Marshal.FinalReleaseComObject(comObj)
            Catch
                ' swallow
            End Try
        End If
    End Sub

End Class
