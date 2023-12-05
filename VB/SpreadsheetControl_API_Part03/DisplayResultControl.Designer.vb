Namespace SpreadsheetControl_API_Part03

    Partial Class DisplayResultControl

        ''' <summary> 
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary> 
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (Me.components IsNot Nothing) Then
                Me.components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

#Region "Component Designer generated code"
        ''' <summary> 
        ''' Required method for Designer support - do not modify 
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Me.dockManager1 = New DevExpress.XtraBars.Docking.DockManager(Me.components)
            Me.hideContainerLeft = New DevExpress.XtraBars.Docking.AutoHideContainer()
            Me.spreadsheetControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetControl()
            Me.spreadsheetFormulaBarControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetFormulaBarControl()
            Me.spreadsheetNameBoxControl1 = New DevExpress.XtraSpreadsheet.SpreadsheetNameBoxControl()
            Me.splitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.splitterControl1 = New DevExpress.XtraEditors.SplitterControl()
            CType((Me.dockManager1), System.ComponentModel.ISupportInitialize).BeginInit()
            CType((Me.spreadsheetNameBoxControl1.Properties), System.ComponentModel.ISupportInitialize).BeginInit()
            CType((Me.splitContainerControl1), System.ComponentModel.ISupportInitialize).BeginInit()
            Me.splitContainerControl1.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' dockManager1
            ' 
            Me.dockManager1.Form = Me
            Me.dockManager1.TopZIndexControls.AddRange(New String() {"DevExpress.XtraBars.BarDockControl", "DevExpress.XtraBars.StandaloneBarDockControl", "System.Windows.Forms.StatusBar", "System.Windows.Forms.MenuStrip", "System.Windows.Forms.StatusStrip", "DevExpress.XtraBars.Ribbon.RibbonStatusBar", "DevExpress.XtraBars.Ribbon.RibbonControl", "DevExpress.XtraBars.Navigation.OfficeNavigationBar", "DevExpress.XtraBars.Navigation.TileNavPane"})
            ' 
            ' hideContainerLeft
            ' 
            Me.hideContainerLeft.BackColor = System.Drawing.SystemColors.Control
            Me.hideContainerLeft.Dock = System.Windows.Forms.DockStyle.Left
            Me.hideContainerLeft.Location = New System.Drawing.Point(0, 0)
            Me.hideContainerLeft.Name = "hideContainerLeft"
            Me.hideContainerLeft.Size = New System.Drawing.Size(19, 300)
            ' 
            ' spreadsheetControl1
            ' 
            Me.spreadsheetControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.spreadsheetControl1.Location = New System.Drawing.Point(0, 25)
            Me.spreadsheetControl1.Name = "spreadsheetControl1"
            Me.spreadsheetControl1.Size = New System.Drawing.Size(800, 575)
            Me.spreadsheetControl1.TabIndex = 0
            Me.spreadsheetControl1.Text = "spreadsheetControl1"
            ' 
            ' spreadsheetFormulaBarControl1
            ' 
            Me.spreadsheetFormulaBarControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.spreadsheetFormulaBarControl1.Location = New System.Drawing.Point(0, 0)
            Me.spreadsheetFormulaBarControl1.MinimumSize = New System.Drawing.Size(0, 20)
            Me.spreadsheetFormulaBarControl1.Name = "spreadsheetFormulaBarControl1"
            Me.spreadsheetFormulaBarControl1.Size = New System.Drawing.Size(650, 20)
            Me.spreadsheetFormulaBarControl1.SpreadsheetControl = Me.spreadsheetControl1
            Me.spreadsheetFormulaBarControl1.TabIndex = 0
            ' 
            ' spreadsheetNameBoxControl1
            ' 
            Me.spreadsheetNameBoxControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.spreadsheetNameBoxControl1.EditValue = "A1"
            Me.spreadsheetNameBoxControl1.Location = New System.Drawing.Point(0, 0)
            Me.spreadsheetNameBoxControl1.MinimumSize = New System.Drawing.Size(0, 20)
            Me.spreadsheetNameBoxControl1.Name = "spreadsheetNameBoxControl1"
            Me.spreadsheetNameBoxControl1.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.spreadsheetNameBoxControl1.Size = New System.Drawing.Size(145, 20)
            Me.spreadsheetNameBoxControl1.SpreadsheetControl = Me.spreadsheetControl1
            Me.spreadsheetNameBoxControl1.TabIndex = 0
            ' 
            ' splitContainerControl1
            ' 
            Me.splitContainerControl1.Dock = System.Windows.Forms.DockStyle.Top
            Me.splitContainerControl1.Location = New System.Drawing.Point(0, 0)
            Me.splitContainerControl1.MinimumSize = New System.Drawing.Size(0, 20)
            Me.splitContainerControl1.Name = "splitContainerControl1"
            Me.splitContainerControl1.Panel1.Controls.Add(Me.spreadsheetNameBoxControl1)
            Me.splitContainerControl1.Panel2.Controls.Add(Me.spreadsheetFormulaBarControl1)
            Me.splitContainerControl1.Size = New System.Drawing.Size(800, 20)
            Me.splitContainerControl1.SplitterPosition = 145
            Me.splitContainerControl1.TabIndex = 2
            ' 
            ' splitterControl1
            ' 
            Me.splitterControl1.Dock = System.Windows.Forms.DockStyle.Top
            Me.splitterControl1.Location = New System.Drawing.Point(0, 20)
            Me.splitterControl1.MinSize = 20
            Me.splitterControl1.Name = "splitterControl1"
            Me.splitterControl1.Size = New System.Drawing.Size(800, 5)
            Me.splitterControl1.TabIndex = 1
            Me.splitterControl1.TabStop = False
            ' 
            ' DisplayResultControl
            ' 
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.Controls.Add(Me.spreadsheetControl1)
            Me.Controls.Add(Me.splitterControl1)
            Me.Controls.Add(Me.splitContainerControl1)
            Me.Name = "DisplayResultControl"
            Me.Size = New System.Drawing.Size(800, 600)
            CType((Me.dockManager1), System.ComponentModel.ISupportInitialize).EndInit()
            CType((Me.spreadsheetNameBoxControl1.Properties), System.ComponentModel.ISupportInitialize).EndInit()
            CType((Me.splitContainerControl1), System.ComponentModel.ISupportInitialize).EndInit()
            Me.splitContainerControl1.ResumeLayout(False)
            Me.ResumeLayout(False)
        End Sub

#End Region
        Private dockManager1 As DevExpress.XtraBars.Docking.DockManager

        Private hideContainerLeft As DevExpress.XtraBars.Docking.AutoHideContainer

        Private spreadsheetControl1 As DevExpress.XtraSpreadsheet.SpreadsheetControl

        Private splitterControl1 As DevExpress.XtraEditors.SplitterControl

        Private splitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl

        Private spreadsheetNameBoxControl1 As DevExpress.XtraSpreadsheet.SpreadsheetNameBoxControl

        Private spreadsheetFormulaBarControl1 As DevExpress.XtraSpreadsheet.SpreadsheetFormulaBarControl
    End Class
End Namespace
