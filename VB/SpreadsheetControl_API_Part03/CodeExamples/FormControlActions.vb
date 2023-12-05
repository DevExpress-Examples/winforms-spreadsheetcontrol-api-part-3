Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace SpreadsheetControl_API_Part03.CodeExamples

    Friend Class FormControlActions

        Private Shared Sub CreateFormControls(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#CreateFormControls"
            Dim formControls = workbook.Worksheets(CInt((0))).FormControls
            ' Create a button form control:
            Dim buttonCellRange = workbook.Worksheets(CInt((0))).Range("B2:C2")
            Dim buttonFormControl = formControls.AddButton(buttonCellRange)
            buttonFormControl.PlainText = "Click Here"
            ' Create a list box form control:
            Dim comboCellRange = workbook.Worksheets(CInt((0))).Range("B4:C4")
            Dim comboBoxControl = formControls.AddComboBox(comboCellRange)
            comboBoxControl.DropDownLines = 3
            comboBoxControl.SourceRange = workbook.Worksheets(CInt((0))).Range("E2:E6")
            comboBoxControl.SelectedIndex = 1
            ' Create a check box form control:
            Dim checkRange = workbook.Worksheets(CInt((0))).Range("D5:E5")
            Dim checkBoxControl = formControls.AddCheckBox(checkRange)
            checkBoxControl.CheckState = DevExpress.Spreadsheet.FormControlCheckState.Checked
            checkBoxControl.PlainText = "Reviewed"
#End Region  ' #CreateFormControls
        End Sub

        Private Shared Sub EditFormControls(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#EditFormControls"
            workbook.LoadDocument("Documents\FormControls.xlsx")
            Dim formControls = workbook.Worksheets(CInt((0))).FormControls
            For Each formControl As DevExpress.Spreadsheet.FormControl In formControls
                formControl.PrintObject = False
            Next
#End Region  ' #EditFormControls
        End Sub
    End Class
End Namespace
