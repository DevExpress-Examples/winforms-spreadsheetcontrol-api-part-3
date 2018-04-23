Imports System
Imports DevExpress.Spreadsheet

Namespace SpreadsheetControl_API_Part03
    Public NotInheritable Class RowAndColumnActions

        Private Sub New()
        End Sub

        Private Shared Sub DeleteRowsBasedOnCondition(ByVal workbook As IWorkbook)
'            #Region "#DeleteRowsBasedOnCondition"
            workbook.LoadDocument("Documents\Document.xlsx")
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create a function specifying the condition to remove worksheet rows.
            Dim rowRemovalCondition As Func(Of Integer, Boolean) = Function(x) worksheet.Cells(x, 0).Value.NumericValue > 3.0 AndAlso worksheet.Cells(x, 0).Value.NumericValue < 14.0

            ' Fill cells with data.
            For i As Integer = 0 To 14
                worksheet.Cells(i, 0).Value = i + 1
                worksheet.Cells(0, i).Value = i + 1
            Next i

            ' Delete all rows that meet the specified condition.
            'worksheet.Rows.Remove(rowRemovalCondition);

            ' Delete rows that meet the specified condition starting from the 7th row.
            worksheet.Rows.Remove(7, rowRemovalCondition)

            ' Delete rows that meet the specified condition starting from the 5th row to the 14th row.
            'worksheet.Rows.Remove(5, 14, rowRemovalCondition);
'            #End Region ' #DeleteRowsBasedOnCondition
        End Sub

        Private Shared Sub DeleteColumnssBasedOnCondition(ByVal workbook As IWorkbook)
'            #Region "#DeleteColumnsBasedOnCondition"
            workbook.LoadDocument("Documents\Document.xlsx")
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Create a function specifying the condition to remove worksheet columns.
            Dim columnRemovalCondition As Func(Of Integer, Boolean) = Function(x) worksheet.Cells(0, x).Value.NumericValue > 3.0 AndAlso worksheet.Cells(0, x).Value.NumericValue < 14.0

            ' Fill cells with data.
            For i As Integer = 0 To 14
                worksheet.Cells(i, 0).Value = i + 1
                worksheet.Cells(0, i).Value = i + 1
            Next i

            ' Delete all columns that meet the specified condition.
            'worksheet.Columns.Remove(columnRemovalCondition);

            ' Delete columns that meet the specified condition starting from the 7th column.
            worksheet.Columns.Remove(7, columnRemovalCondition)

            ' Delete columns that meet the specified condition starting from the 5th column to the 14th column.
            'worksheet.Columns.Remove(5, 14, columnRemovalCondition);

'            #End Region ' #DeleteColumnsBasedOnCondition
        End Sub
    End Class
End Namespace
