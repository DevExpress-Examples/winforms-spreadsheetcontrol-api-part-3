using System;
using DevExpress.Spreadsheet;

namespace SpreadsheetControl_API_Part03
{
    public static class RowAndColumnActions
    {
        static void DeleteRowsBasedOnCondition(IWorkbook workbook)
        {
            #region #DeleteRowsBasedOnCondition
            workbook.LoadDocument("Documents\\Document.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            // Create a function specifying the condition to remove worksheet rows.
            Func<int, bool> rowRemovalCondition = x => worksheet.Cells[x, 0].Value.NumericValue > 3.0 && worksheet.Cells[x, 0].Value.NumericValue < 14.0;

            // Fill cells with data.
            for (int i = 0; i < 15; i++)
            {
                worksheet.Cells[i, 0].Value = i + 1;
                worksheet.Cells[0, i].Value = i + 1;
            }

            // Delete all rows that meet the specified condition.
            //worksheet.Rows.Remove(rowRemovalCondition);

            // Delete rows that meet the specified condition starting from the 7th row.
            worksheet.Rows.Remove(7, rowRemovalCondition);

            // Delete rows that meet the specified condition starting from the 5th row to the 14th row.
            //worksheet.Rows.Remove(5, 14, rowRemovalCondition);
            #endregion #DeleteRowsBasedOnCondition
        }

        static void DeleteColumnssBasedOnCondition(IWorkbook workbook)
        {
            #region #DeleteColumnsBasedOnCondition
            workbook.LoadDocument("Documents\\Document.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            // Create a function specifying the condition to remove worksheet columns.
            Func<int, bool> columnRemovalCondition = x => worksheet.Cells[0, x].Value.NumericValue > 3.0 && worksheet.Cells[0, x].Value.NumericValue < 14.0;

            // Fill cells with data.
            for (int i = 0; i < 15; i++)
            {
                worksheet.Cells[i, 0].Value = i + 1;
                worksheet.Cells[0, i].Value = i + 1;
            }

            // Delete all columns that meet the specified condition.
            //worksheet.Columns.Remove(columnRemovalCondition);

            // Delete columns that meet the specified condition starting from the 7th column.
            worksheet.Columns.Remove(7, columnRemovalCondition);

            // Delete columns that meet the specified condition starting from the 5th column to the 14th column.
            //worksheet.Columns.Remove(5, 14, columnRemovalCondition);

            #endregion #DeleteColumnsBasedOnCondition
        }
    }
}
