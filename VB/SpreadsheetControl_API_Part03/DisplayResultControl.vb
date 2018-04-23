Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.Data
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DevExpress.XtraSpreadsheet

Namespace SpreadsheetControl_API_Part03
    Partial Public Class DisplayResultControl
        Inherits UserControl

        Public Sub New()
            InitializeComponent()
        End Sub

        Public ReadOnly Property Speadsheet() As SpreadsheetControl
            Get
                Return spreadsheetControl1
            End Get
        End Property

    End Class
End Namespace
