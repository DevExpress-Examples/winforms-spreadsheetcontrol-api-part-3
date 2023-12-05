Imports System
Imports DevExpress.Spreadsheet

Namespace SpreadsheetControl_API_Part03

    Public Module RowAndColumnActions

        Private Sub DeleteRowsBasedOnCondition(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#DeleteRowsBasedOnCondition"
            workbook.LoadDocument("Documents\Document.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Create a function specifying the condition to remove worksheet rows.
            Dim rowRemovalCondition As System.Func(Of Integer, Boolean) = Function(x) worksheet.Cells(CInt((x)), CInt((0))).Value.NumericValue > 3.0 AndAlso worksheet.Cells(CInt((x)), CInt((0))).Value.NumericValue < 14.0
            ' Fill cells with data.
            For i As Integer = 0 To 15 - 1
                worksheet.Cells(CInt((i)), CInt((0))).Value = i + 1
                worksheet.Cells(CInt((0)), CInt((i))).Value = i + 1
            Next

            ' Delete all rows that meet the specified condition.
            'worksheet.Rows.Remove(rowRemovalCondition);
            ' Delete rows that meet the specified condition starting from the 7th row.
            worksheet.Rows.Remove(7, rowRemovalCondition)
        ' Delete rows that meet the specified condition starting from the 5th row to the 14th row.
        'worksheet.Rows.Remove(5, 14, rowRemovalCondition);
#End Region  ' #DeleteRowsBasedOnCondition
        End Sub

        Private Sub DeleteColumnssBasedOnCondition(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#DeleteColumnsBasedOnCondition"
            workbook.LoadDocument("Documents\Document.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Create a function specifying the condition to remove worksheet columns.
            Dim columnRemovalCondition As System.Func(Of Integer, Boolean) = Function(x) worksheet.Cells(CInt((0)), CInt((x))).Value.NumericValue > 3.0 AndAlso worksheet.Cells(CInt((0)), CInt((x))).Value.NumericValue < 14.0
            ' Fill cells with data.
            For i As Integer = 0 To 15 - 1
                worksheet.Cells(CInt((i)), CInt((0))).Value = i + 1
                worksheet.Cells(CInt((0)), CInt((i))).Value = i + 1
            Next

            ' Delete all columns that meet the specified condition.
            'worksheet.Columns.Remove(columnRemovalCondition);
            ' Delete columns that meet the specified condition starting from the 7th column.
            worksheet.Columns.Remove(7, columnRemovalCondition)
        ' Delete columns that meet the specified condition starting from the 5th column to the 14th column.
        'worksheet.Columns.Remove(5, 14, columnRemovalCondition);
#End Region  ' #DeleteColumnsBasedOnCondition
        End Sub
    End Module
End Namespace
