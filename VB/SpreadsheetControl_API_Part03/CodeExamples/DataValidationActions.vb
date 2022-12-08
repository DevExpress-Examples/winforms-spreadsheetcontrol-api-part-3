Imports System
Imports DevExpress.Spreadsheet
Imports System.Drawing
Imports System.Linq

Namespace SpreadsheetControl_API_Part03

    Public Module DataValidationActions

        Private Sub AddDataValidation(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#AddDataValidation"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            worksheet(CStr(("C1"))).SetValue(System.DateTime.Now)
            worksheet(CStr(("C1"))).NumberFormat = "mmm/d/yyyy h:mm"
            ' Restrict data entry to a whole number from 10 to 20.
            worksheet.DataValidations.Add(worksheet("B1"), DevExpress.Spreadsheet.DataValidationType.WholeNumber, DevExpress.Spreadsheet.DataValidationOperator.Between, 10, 20)
            ' Restrict data entry to a number within limits.
            Dim validation As DevExpress.Spreadsheet.DataValidation = worksheet.DataValidations.Add(worksheet("F4:F11"), DevExpress.Spreadsheet.DataValidationType.[Decimal], DevExpress.Spreadsheet.DataValidationOperator.Between, 10, 40)
            ' Restrict data entry using criteria calculated by a worksheet formula.
            worksheet.DataValidations.Add(worksheet("B4:B11"), DevExpress.Spreadsheet.DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)")
            ' Restrict data entry to 3 symbols.
            worksheet.DataValidations.Add(worksheet("D4:D11"), DevExpress.Spreadsheet.DataValidationType.TextLength, DevExpress.Spreadsheet.DataValidationOperator.Equal, 3)
            ' Restrict data entry to values in a drop-down list specified in code. 
            ' Note that the list in code should always use comma to separate entries, 
            ' but the list in UI is displayed using culture-specific list separator.
            worksheet.DataValidations.Add(worksheet("A4:A11"), DevExpress.Spreadsheet.DataValidationType.List, "PASS, FAIL")
            ' Restrict data entry to values in a drop-down list obtained from a worksheet.
            worksheet.DataValidations.Add(worksheet("E4:E11"), DevExpress.Spreadsheet.DataValidationType.List, DevExpress.Spreadsheet.ValueObject.FromRange(worksheet(CStr(("H4:H9"))).GetRangeWithAbsoluteReference()))
            ' Restrict data entry to a time before the specified time.
            worksheet.DataValidations.Add(worksheet("C1"), DevExpress.Spreadsheet.DataValidationType.Time, DevExpress.Spreadsheet.DataValidationOperator.LessThanOrEqual, System.DateTime.Now)
            ' Highlight data validation ranges.
            worksheet(CStr(("H4:H9"))).FillColor = System.Drawing.Color.LightGray
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3, &HFFDFC4, &HFFDAE9}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #AddDataValidation
        End Sub

        Private Sub ChangeCriteria(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#ChangeCriteria"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Restrict data entry to a number within limits.
            Dim validation As DevExpress.Spreadsheet.DataValidation = worksheet.DataValidations.Add(worksheet("F4:F11"), DevExpress.Spreadsheet.DataValidationType.[Decimal], DevExpress.Spreadsheet.DataValidationOperator.Between, 10, 40)
            ' Change the validation operator and criteria.
            ' Range F4:F11 should contain numbers greater than or equal 20.
            validation.[Operator] = DevExpress.Spreadsheet.DataValidationOperator.GreaterThanOrEqual
            validation.Criteria = 20
            validation.Criteria2 = DevExpress.Spreadsheet.ValueObject.Empty
            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #ChangeCriteria
        End Sub

        Private Sub UseUnionRange(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#UseUnionRange"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Create a union range.
            Dim range As DevExpress.Spreadsheet.CellRange = worksheet.Range.Union(worksheet("F4:F5"), worksheet("F6:F11"))
            ' Restrict data entry to a number within limits.
            worksheet.DataValidations.Add(range, DevExpress.Spreadsheet.DataValidationType.[Decimal], DevExpress.Spreadsheet.DataValidationOperator.Between, 10, 40)
            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #UseUnionRange
        End Sub

        Private Sub ShowInputMessage(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#ShowInputMessage"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Restrict data entry to a 5-digit number
            Dim validation As DevExpress.Spreadsheet.DataValidation = worksheet.DataValidations.Add(worksheet("B4:B11"), DevExpress.Spreadsheet.DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)")
            ' Show input message.
            validation.InputTitle = "Employee Id"
            validation.InputMessage = "Please enter 5-digit number"
            validation.ShowInputMessage = True
            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #ShowInputMessage
        End Sub

        Private Sub ShowErrorMessage(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#ShowErrorMessage"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Restrict data entry to a 5-digit number.
            Dim validation As DevExpress.Spreadsheet.DataValidation = worksheet.DataValidations.Add(worksheet("B4:B11"), DevExpress.Spreadsheet.DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)")
            ' Show error message.
            validation.ErrorTitle = "Wrong Employee Id"
            validation.ErrorMessage = "The value you entered is not valid. Use 5-digit number for the employee ID."
            validation.ErrorStyle = DevExpress.Spreadsheet.DataValidationErrorStyle.Information
            validation.ShowErrorMessage = True
            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #ShowErrorMessage
        End Sub

        Private Sub GetDataValidation(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#GetDataValidation"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Add data validations.
            worksheet.DataValidations.Add(worksheet("D4:D11"), DevExpress.Spreadsheet.DataValidationType.TextLength, DevExpress.Spreadsheet.DataValidationOperator.Equal, 3)
            worksheet.DataValidations.Add(worksheet("E4:E11"), DevExpress.Spreadsheet.DataValidationType.List, DevExpress.Spreadsheet.ValueObject.FromRange(worksheet(CStr(("H4:H9"))).GetRangeWithAbsoluteReference()))
            ' Get data validation entry associated with a particular cell.
            worksheet.DataValidations.GetDataValidation(CType((worksheet.Cells(CStr(("E4")))), DevExpress.Spreadsheet.Cell)).Criteria = DevExpress.Spreadsheet.ValueObject.FromRange(worksheet("H4:H5"))
            ' Get data validation entries for the specified range.
            Dim myValidation = worksheet.DataValidations.GetDataValidations(worksheet("D4:E11")).Where(Function(d) d.ValidationType = DevExpress.Spreadsheet.DataValidationType.TextLength).SingleOrDefault()
            If myValidation IsNot Nothing Then myValidation.Criteria = 4
            ' Get data validation entries that meet certain criteria.
            For Each d In worksheet.DataValidations.GetDataValidations(DevExpress.Spreadsheet.DataValidationType.TextLength, DevExpress.Spreadsheet.DataValidationOperator.Equal, 4, DevExpress.Spreadsheet.ValueObject.Empty)
                ' Change criteria operator.
                ' Range D4:D11 should contain text with more than 4 characters.
                d.[Operator] = DevExpress.Spreadsheet.DataValidationOperator.GreaterThan
            Next

            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #GetDataValidation
        End Sub

        Private Sub ValidateCellValue(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#ValidateCellValue"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Add data validations.
            worksheet.DataValidations.Add(worksheet("D4:D11"), DevExpress.Spreadsheet.DataValidationType.TextLength, DevExpress.Spreadsheet.DataValidationOperator.Equal, 3)
            'Check whether the cell value meets the validation criteria:
            Dim isValid As Boolean = worksheet.DataValidations.Validate(worksheet.Cells("D4"), worksheet.Cells(CStr(("J4"))).Value)
            If isValid Then
                worksheet(CStr(("D4"))).CopyFrom(worksheet("J4"))
            End If
'#End Region  ' #ValidateCellValue
        End Sub

        Private Sub RemoveDataValidation(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#RemoveDataValidation"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Add data validations.
            worksheet.DataValidations.Add(worksheet("D4:D11"), DevExpress.Spreadsheet.DataValidationType.TextLength, DevExpress.Spreadsheet.DataValidationOperator.Equal, 3)
            worksheet.DataValidations.Add(worksheet("E4:E11"), DevExpress.Spreadsheet.DataValidationType.List, DevExpress.Spreadsheet.ValueObject.FromRange(worksheet(CStr(("H4:H9"))).GetRangeWithAbsoluteReference()))
            ' Remove data validation by index.
            worksheet.DataValidations.RemoveAt(1)
            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #RemoveDataValidation
        End Sub

        Private Sub RemoveAllDataValidations(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
'#Region "#RemoveAllDataValidations"
            workbook.LoadDocument("Documents\DataValidation.xlsx")
            Dim worksheet As DevExpress.Spreadsheet.Worksheet = workbook.Worksheets(0)
            ' Add data validations.
            worksheet.DataValidations.Add(worksheet("D4:D11"), DevExpress.Spreadsheet.DataValidationType.TextLength, DevExpress.Spreadsheet.DataValidationOperator.Equal, 3)
            worksheet.DataValidations.Add(worksheet("E4:E11"), DevExpress.Spreadsheet.DataValidationType.List, DevExpress.Spreadsheet.ValueObject.FromRange(worksheet(CStr(("H4:H9"))).GetRangeWithAbsoluteReference()))
            ' Remove all data validations.
            worksheet.DataValidations.Clear()
            ' Highlight data validation ranges.
            Dim MyColorScheme As Integer() = New Integer() {&HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3}
            For i As Integer = 0 To worksheet.DataValidations.Count - 1
                worksheet.DataValidations(CInt((i))).Range.FillColor = System.Drawing.Color.FromArgb(MyColorScheme(i))
            Next
'#End Region  ' #RemoveAllDataValidations
        End Sub
    End Module
End Namespace
