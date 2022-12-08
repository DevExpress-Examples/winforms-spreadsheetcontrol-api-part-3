Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports DevExpress.XtraSpreadsheet

Namespace SpreadsheetControl_API_Part03

    Public Partial Class DisplayResultControl
        Inherits UserControl

        Public Sub New()
            InitializeComponent()
        End Sub

        Public ReadOnly Property Speadsheet As SpreadsheetControl
            Get
                Return spreadsheetControl1
            End Get
        End Property
    End Class
End Namespace
