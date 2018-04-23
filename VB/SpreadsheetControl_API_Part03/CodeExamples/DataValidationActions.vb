Imports System
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet
Imports System.Drawing
Imports System.Linq

Namespace SpreadsheetControl_API_Part03
	Public NotInheritable Class DataValidationActions

		Private Sub New()
		End Sub


		Private Shared Sub AddDataValidation(ByVal workbook As IWorkbook)
'			#Region "#AddDataValidation"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			worksheet("C1").SetValue(Date.Now)
			worksheet("C1").NumberFormat = "mmm/d/yyyy h:mm"

			' Restrict data entry to a whole number from 10 to 20.
			worksheet.DataValidations.Add(worksheet("B1"), DataValidationType.WholeNumber, DataValidationOperator.Between, 10, 20)

			' Restrict data entry to a number within limits.
			Dim validation As DataValidation = worksheet.DataValidations.Add(worksheet("F4:F11"), DataValidationType.Decimal, DataValidationOperator.Between, 10, 40)

			' Restrict data entry using criteria calculated by a worksheet formula.
			worksheet.DataValidations.Add(worksheet("B4:B11"), DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)")

			' Restrict data entry to 3 symbols.
			worksheet.DataValidations.Add(worksheet("D4:D11"), DataValidationType.TextLength, DataValidationOperator.Equal, 3)

			' Restrict data entry to values in a drop-down list specified in code. 
			' Note that the list in code should always use comma to separate entries, 
			' but the list in UI is displayed using culture-specific list separator.
			worksheet.DataValidations.Add(worksheet("A4:A11"), DataValidationType.List, "PASS, FAIL")

			' Restrict data entry to values in a drop-down list obtained from a worksheet.
			worksheet.DataValidations.Add(worksheet("E4:E11"), DataValidationType.List, ValueObject.FromRange(worksheet("H4:H9").GetRangeWithAbsoluteReference()))

			' Restrict data entry to a time before the specified time.
			worksheet.DataValidations.Add(worksheet("C1"), DataValidationType.Time, DataValidationOperator.LessThanOrEqual, Date.Now)

			' Highlight data validation ranges.
			worksheet("H4:H9").FillColor = Color.LightGray
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3, &HFFDFC4, &HFFDAE9}
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #AddDataValidation
		End Sub

		Private Shared Sub ChangeCriteria(ByVal workbook As IWorkbook)
'			#Region "#ChangeCriteria"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Restrict data entry to a number within limits.
			Dim validation As DataValidation = worksheet.DataValidations.Add(worksheet("F4:F11"), DataValidationType.Decimal, DataValidationOperator.Between, 10, 40)

			' Change the validation operator and criteria.
			' Range F4:F11 should contain numbers greater than or equal 20.
			validation.Operator = DataValidationOperator.GreaterThanOrEqual
			validation.Criteria = 20
			validation.Criteria2 = ValueObject.Empty

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #ChangeCriteria
		End Sub

		Private Shared Sub UseUnionRange(ByVal workbook As IWorkbook)
'			#Region "#UseUnionRange"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Create a union range.
			Dim range As Range = worksheet.Range.Union(worksheet("F4:F5"), worksheet("F6:F11"))
			' Restrict data entry to a number within limits.
			worksheet.DataValidations.Add(range, DataValidationType.Decimal, DataValidationOperator.Between, 10, 40)

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #UseUnionRange
		End Sub

		Private Shared Sub ShowInputMessage(ByVal workbook As IWorkbook)
'			#Region "#ShowInputMessage"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Restrict data entry to a 5-digit number
			Dim validation As DataValidation = worksheet.DataValidations.Add(worksheet("B4:B11"), DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)")

			' Show input message.
			validation.InputTitle = "Employee Id"
			validation.InputMessage = "Please enter 5-digit number"
			validation.ShowInputMessage = True

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #ShowInputMessage
		End Sub

		Private Shared Sub ShowErrorMessage(ByVal workbook As IWorkbook)
'			#Region "#ShowErrorMessage"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Restrict data entry to a 5-digit number.
			Dim validation As DataValidation = worksheet.DataValidations.Add(worksheet("B4:B11"), DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)")

			' Show error message.
			validation.ErrorTitle = "Wrong Employee Id"
			validation.ErrorMessage = "The value you entered is not valid. Use 5-digit number for the employee ID."
			validation.ErrorStyle = DataValidationErrorStyle.Information
			validation.ShowErrorMessage = True

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #ShowErrorMessage
		End Sub

		Private Shared Sub GetDataValidation(ByVal workbook As IWorkbook)
'			#Region "#GetDataValidation"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Add data validations.
			worksheet.DataValidations.Add(worksheet("D4:D11"), DataValidationType.TextLength, DataValidationOperator.Equal, 3)
			worksheet.DataValidations.Add(worksheet("E4:E11"), DataValidationType.List, ValueObject.FromRange(worksheet("H4:H9").GetRangeWithAbsoluteReference()))

			' Get data validation entry associated with a particular cell.
			worksheet.DataValidations.GetDataValidation(worksheet.Cells("E4")).Criteria = ValueObject.FromRange(worksheet("H4:H5"))

			' Get data validation entries for the specified range.
			Dim myValidation = worksheet.DataValidations.GetDataValidations(worksheet("D4:E11")).Where(Function(d) d.ValidationType = DataValidationType.TextLength).SingleOrDefault()
			If myValidation IsNot Nothing Then
				myValidation.Criteria = 4
			End If

			' Get data validation entries that meet certain criteria.
			For Each d In worksheet.DataValidations.GetDataValidations(DataValidationType.TextLength, DataValidationOperator.Equal, 4, ValueObject.Empty)
				' Change criteria operator.
				' Range D4:D11 should contain text with more than 4 characters.
				d.Operator = DataValidationOperator.GreaterThan
			Next d

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #GetDataValidation
		End Sub

		Private Shared Sub RemoveDataValidation(ByVal workbook As IWorkbook)
'			#Region "#RemoveDataValidation"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Add data validations.
			worksheet.DataValidations.Add(worksheet("D4:D11"), DataValidationType.TextLength, DataValidationOperator.Equal, 3)
			worksheet.DataValidations.Add(worksheet("E4:E11"), DataValidationType.List, ValueObject.FromRange(worksheet("H4:H9").GetRangeWithAbsoluteReference()))

			' Remove data validation by index.
			worksheet.DataValidations.RemoveAt(1)

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #RemoveDataValidation
		End Sub

		Private Shared Sub RemoveAllDataValidations(ByVal workbook As IWorkbook)
'			#Region "#RemoveAllDataValidations"
			workbook.LoadDocument("Documents\DataValidation.xlsx")
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Add data validations.
			worksheet.DataValidations.Add(worksheet("D4:D11"), DataValidationType.TextLength, DataValidationOperator.Equal, 3)
			worksheet.DataValidations.Add(worksheet("E4:E11"), DataValidationType.List, ValueObject.FromRange(worksheet("H4:H9").GetRangeWithAbsoluteReference()))

			' Remove all data validations.
			worksheet.DataValidations.Clear()

			' Highlight data validation ranges.
			Dim MyColorScheme() As Integer = { &HFFC4C4, &HFFD9D9, &HFFF6F6, &HFFECEC, &HE9D3D3 }
			For i As Integer = 0 To worksheet.DataValidations.Count - 1
				worksheet.DataValidations(i).Range.FillColor = Color.FromArgb(MyColorScheme(i))
			Next i
'			#End Region ' #RemoveAllDataValidations
		End Sub
	End Class
End Namespace
