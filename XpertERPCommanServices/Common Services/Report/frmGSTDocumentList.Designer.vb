<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmGSTDocumentList
    Inherits FrmMainTranScreen

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim RadListDataItem7 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem8 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem9 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem1 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem2 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem3 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim TableViewDefinition3 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition4 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.RadPageViewPage1 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.RadGroupBox5 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtToDate = New Telerik.WinControls.UI.RadDateTimePicker()
        Me.RadLabel2 = New common.Controls.MyLabel()
        Me.txtFromDate = New Telerik.WinControls.UI.RadDateTimePicker()
        Me.RadLabel3 = New common.Controls.MyLabel()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.ddlTransaction = New common.Controls.MyComboBox()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.ddlInvoiceType = New common.Controls.MyComboBox()
        Me.lbltype = New common.Controls.MyLabel()
        Me.ddlType = New common.Controls.MyComboBox()
        Me.RadPageViewPage2 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.gv = New common.UserControls.MyRadGridView()
        Me.RadPageViewPage3 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.gvDetails = New common.UserControls.MyRadGridView()
        Me.RadSplitButton1 = New Telerik.WinControls.UI.RadSplitButton()
        Me.QExpExcel = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnGO = New Telerik.WinControls.UI.RadButton()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.btnReset = New Telerik.WinControls.UI.RadButton()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.RadPageViewPage1.SuspendLayout()
        CType(Me.RadGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox5.SuspendLayout()
        CType(Me.txtToDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ddlTransaction, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ddlInvoiceType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lbltype, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ddlType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage2.SuspendLayout()
        CType(Me.gv, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage3.SuspendLayout()
        CType(Me.gvDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvDetails.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadPageView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.RadSplitButton1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnGO)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReset)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 405
        Me.SplitContainer1.TabIndex = 0
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage1)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage2)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage3)
        Me.RadPageView1.DefaultPage = Me.RadPageViewPage1
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.RadPageViewPage1
        Me.RadPageView1.Size = New System.Drawing.Size(800, 405)
        Me.RadPageView1.TabIndex = 0
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        '
        'RadPageViewPage1
        '
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox5)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel2)
        Me.RadPageViewPage1.Controls.Add(Me.ddlTransaction)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.ddlInvoiceType)
        Me.RadPageViewPage1.Controls.Add(Me.lbltype)
        Me.RadPageViewPage1.Controls.Add(Me.ddlType)
        Me.RadPageViewPage1.ItemSize = New System.Drawing.SizeF(41.0!, 28.0!)
        Me.RadPageViewPage1.Location = New System.Drawing.Point(10, 37)
        Me.RadPageViewPage1.Name = "RadPageViewPage1"
        Me.RadPageViewPage1.Size = New System.Drawing.Size(779, 357)
        Me.RadPageViewPage1.Text = "Filter"
        '
        'RadGroupBox5
        '
        Me.RadGroupBox5.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox5.Controls.Add(Me.txtToDate)
        Me.RadGroupBox5.Controls.Add(Me.RadLabel2)
        Me.RadGroupBox5.Controls.Add(Me.txtFromDate)
        Me.RadGroupBox5.Controls.Add(Me.RadLabel3)
        Me.RadGroupBox5.HeaderText = "Date Range"
        Me.RadGroupBox5.Location = New System.Drawing.Point(16, 6)
        Me.RadGroupBox5.Name = "RadGroupBox5"
        Me.RadGroupBox5.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox5.Size = New System.Drawing.Size(317, 52)
        Me.RadGroupBox5.TabIndex = 319
        Me.RadGroupBox5.Text = "Date Range"
        '
        'txtToDate
        '
        Me.txtToDate.CustomFormat = "dd/MM/yyyy"
        Me.txtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtToDate.Location = New System.Drawing.Point(200, 21)
        Me.txtToDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDate.Size = New System.Drawing.Size(88, 20)
        Me.txtToDate.TabIndex = 320
        Me.txtToDate.TabStop = False
        Me.txtToDate.Text = "31/08/2011"
        Me.txtToDate.Value = New Date(2011, 8, 31, 23, 50, 36, 937)
        '
        'RadLabel2
        '
        Me.RadLabel2.FieldName = Nothing
        Me.RadLabel2.Location = New System.Drawing.Point(13, 23)
        Me.RadLabel2.Name = "RadLabel2"
        Me.RadLabel2.Size = New System.Drawing.Size(34, 18)
        Me.RadLabel2.TabIndex = 6
        Me.RadLabel2.Text = "From:"
        '
        'txtFromDate
        '
        Me.txtFromDate.CustomFormat = "dd/MM/yyyy"
        Me.txtFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtFromDate.Location = New System.Drawing.Point(59, 21)
        Me.txtFromDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtFromDate.Name = "txtFromDate"
        Me.txtFromDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtFromDate.Size = New System.Drawing.Size(92, 20)
        Me.txtFromDate.TabIndex = 7
        Me.txtFromDate.TabStop = False
        Me.txtFromDate.Text = "31/08/2011"
        Me.txtFromDate.Value = New Date(2011, 8, 31, 23, 50, 36, 937)
        '
        'RadLabel3
        '
        Me.RadLabel3.FieldName = Nothing
        Me.RadLabel3.Location = New System.Drawing.Point(173, 23)
        Me.RadLabel3.Name = "RadLabel3"
        Me.RadLabel3.Size = New System.Drawing.Size(21, 18)
        Me.RadLabel3.TabIndex = 5
        Me.RadLabel3.Text = "To:"
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Location = New System.Drawing.Point(16, 121)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(64, 18)
        Me.MyLabel2.TabIndex = 315
        Me.MyLabel2.Text = "Transaction"
        '
        'ddlTransaction
        '
        Me.ddlTransaction.AutoCompleteDisplayMember = Nothing
        Me.ddlTransaction.AutoCompleteValueMember = Nothing
        Me.ddlTransaction.CalculationExpression = Nothing
        Me.ddlTransaction.DropDownAnimationEnabled = True
        Me.ddlTransaction.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.ddlTransaction.FieldCode = Nothing
        Me.ddlTransaction.FieldDesc = Nothing
        Me.ddlTransaction.FieldMaxLength = 0
        Me.ddlTransaction.FieldName = Nothing
        Me.ddlTransaction.isCalculatedField = False
        Me.ddlTransaction.IsSourceFromTable = False
        Me.ddlTransaction.IsSourceFromValueList = False
        Me.ddlTransaction.IsUnique = False
        Me.ddlTransaction.Location = New System.Drawing.Point(99, 120)
        Me.ddlTransaction.MendatroryField = False
        Me.ddlTransaction.MyLinkLable1 = Nothing
        Me.ddlTransaction.MyLinkLable2 = Nothing
        Me.ddlTransaction.Name = "ddlTransaction"
        Me.ddlTransaction.ReferenceFieldDesc = Nothing
        Me.ddlTransaction.ReferenceFieldName = Nothing
        Me.ddlTransaction.ReferenceTableName = Nothing
        Me.ddlTransaction.Size = New System.Drawing.Size(234, 20)
        Me.ddlTransaction.TabIndex = 316
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Location = New System.Drawing.Point(16, 97)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(69, 18)
        Me.MyLabel1.TabIndex = 317
        Me.MyLabel1.Text = "Invoice Type"
        '
        'ddlInvoiceType
        '
        Me.ddlInvoiceType.AutoCompleteDisplayMember = Nothing
        Me.ddlInvoiceType.AutoCompleteValueMember = Nothing
        Me.ddlInvoiceType.CalculationExpression = Nothing
        Me.ddlInvoiceType.DropDownAnimationEnabled = True
        Me.ddlInvoiceType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.ddlInvoiceType.FieldCode = Nothing
        Me.ddlInvoiceType.FieldDesc = Nothing
        Me.ddlInvoiceType.FieldMaxLength = 0
        Me.ddlInvoiceType.FieldName = Nothing
        Me.ddlInvoiceType.isCalculatedField = False
        Me.ddlInvoiceType.IsSourceFromTable = False
        Me.ddlInvoiceType.IsSourceFromValueList = False
        Me.ddlInvoiceType.IsUnique = False
        RadListDataItem7.Text = "B2B"
        RadListDataItem8.Text = "B2C"
        RadListDataItem9.Text = "Cancel"
        Me.ddlInvoiceType.Items.Add(RadListDataItem7)
        Me.ddlInvoiceType.Items.Add(RadListDataItem8)
        Me.ddlInvoiceType.Items.Add(RadListDataItem9)
        Me.ddlInvoiceType.Location = New System.Drawing.Point(99, 96)
        Me.ddlInvoiceType.MendatroryField = False
        Me.ddlInvoiceType.MyLinkLable1 = Nothing
        Me.ddlInvoiceType.MyLinkLable2 = Nothing
        Me.ddlInvoiceType.Name = "ddlInvoiceType"
        Me.ddlInvoiceType.ReferenceFieldDesc = Nothing
        Me.ddlInvoiceType.ReferenceFieldName = Nothing
        Me.ddlInvoiceType.ReferenceTableName = Nothing
        Me.ddlInvoiceType.Size = New System.Drawing.Size(234, 20)
        Me.ddlInvoiceType.TabIndex = 318
        '
        'lbltype
        '
        Me.lbltype.FieldName = Nothing
        Me.lbltype.Location = New System.Drawing.Point(16, 73)
        Me.lbltype.Name = "lbltype"
        Me.lbltype.Size = New System.Drawing.Size(30, 18)
        Me.lbltype.TabIndex = 315
        Me.lbltype.Text = "Type"
        '
        'ddlType
        '
        Me.ddlType.AutoCompleteDisplayMember = Nothing
        Me.ddlType.AutoCompleteValueMember = Nothing
        Me.ddlType.CalculationExpression = Nothing
        Me.ddlType.DropDownAnimationEnabled = True
        Me.ddlType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.ddlType.FieldCode = Nothing
        Me.ddlType.FieldDesc = Nothing
        Me.ddlType.FieldMaxLength = 0
        Me.ddlType.FieldName = Nothing
        Me.ddlType.isCalculatedField = False
        Me.ddlType.IsSourceFromTable = False
        Me.ddlType.IsSourceFromValueList = False
        Me.ddlType.IsUnique = False
        RadListDataItem1.Text = "Taxable"
        RadListDataItem2.Text = "Non-Taxable"
        RadListDataItem3.Text = "Both"
        Me.ddlType.Items.Add(RadListDataItem1)
        Me.ddlType.Items.Add(RadListDataItem2)
        Me.ddlType.Items.Add(RadListDataItem3)
        Me.ddlType.Location = New System.Drawing.Point(99, 72)
        Me.ddlType.MendatroryField = False
        Me.ddlType.MyLinkLable1 = Nothing
        Me.ddlType.MyLinkLable2 = Nothing
        Me.ddlType.Name = "ddlType"
        Me.ddlType.ReferenceFieldDesc = Nothing
        Me.ddlType.ReferenceFieldName = Nothing
        Me.ddlType.ReferenceTableName = Nothing
        Me.ddlType.Size = New System.Drawing.Size(234, 20)
        Me.ddlType.TabIndex = 316
        '
        'RadPageViewPage2
        '
        Me.RadPageViewPage2.Controls.Add(Me.gv)
        Me.RadPageViewPage2.ItemSize = New System.Drawing.SizeF(50.0!, 28.0!)
        Me.RadPageViewPage2.Location = New System.Drawing.Point(10, 37)
        Me.RadPageViewPage2.Name = "RadPageViewPage2"
        Me.RadPageViewPage2.Size = New System.Drawing.Size(779, 357)
        Me.RadPageViewPage2.Text = "Report"
        '
        'gv
        '
        Me.gv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gv.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.gv.MasterTemplate.AllowAddNewRow = False
        Me.gv.MasterTemplate.AllowEditRow = False
        Me.gv.MasterTemplate.EnableFiltering = True
        Me.gv.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gv.MasterTemplate.ShowFilteringRow = False
        Me.gv.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv.MasterTemplate.ViewDefinition = TableViewDefinition3
        Me.gv.MyExportAPI = False
        Me.gv.MyExportFilePath = ""
        Me.gv.MyStopExport = False
        Me.gv.Name = "gv"
        Me.gv.ShowGroupPanel = False
        Me.gv.ShowHeaderCellButtons = True
        Me.gv.Size = New System.Drawing.Size(779, 357)
        Me.gv.TabIndex = 1
        Me.gv.VarID = ""
        '
        'RadPageViewPage3
        '
        Me.RadPageViewPage3.Controls.Add(Me.gvDetails)
        Me.RadPageViewPage3.ItemSize = New System.Drawing.SizeF(87.0!, 28.0!)
        Me.RadPageViewPage3.Location = New System.Drawing.Point(10, 37)
        Me.RadPageViewPage3.Name = "RadPageViewPage3"
        Me.RadPageViewPage3.Size = New System.Drawing.Size(779, 357)
        Me.RadPageViewPage3.Text = "Report Details"
        '
        'gvDetails
        '
        Me.gvDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gvDetails.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.gvDetails.MasterTemplate.AllowAddNewRow = False
        Me.gvDetails.MasterTemplate.AllowEditRow = False
        Me.gvDetails.MasterTemplate.EnableFiltering = True
        Me.gvDetails.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gvDetails.MasterTemplate.ShowFilteringRow = False
        Me.gvDetails.MasterTemplate.ShowHeaderCellButtons = True
        Me.gvDetails.MasterTemplate.ViewDefinition = TableViewDefinition4
        Me.gvDetails.MyExportAPI = False
        Me.gvDetails.MyExportFilePath = ""
        Me.gvDetails.MyStopExport = False
        Me.gvDetails.Name = "gvDetails"
        Me.gvDetails.ShowGroupPanel = False
        Me.gvDetails.ShowHeaderCellButtons = True
        Me.gvDetails.Size = New System.Drawing.Size(779, 357)
        Me.gvDetails.TabIndex = 2
        Me.gvDetails.VarID = ""
        '
        'RadSplitButton1
        '
        Me.RadSplitButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadSplitButton1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.QExpExcel})
        Me.RadSplitButton1.Location = New System.Drawing.Point(170, 11)
        Me.RadSplitButton1.Name = "RadSplitButton1"
        Me.RadSplitButton1.Size = New System.Drawing.Size(77, 21)
        Me.RadSplitButton1.TabIndex = 330
        Me.RadSplitButton1.Text = "Export"
        '
        'QExpExcel
        '
        Me.QExpExcel.Name = "QExpExcel"
        Me.QExpExcel.Text = "Excel"
        Me.QExpExcel.UseCompatibleTextRendering = False
        '
        'btnGO
        '
        Me.btnGO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGO.Location = New System.Drawing.Point(12, 11)
        Me.btnGO.Name = "btnGO"
        Me.btnGO.Size = New System.Drawing.Size(77, 21)
        Me.btnGO.TabIndex = 331
        Me.btnGO.Text = ">>>"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(716, 11)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(77, 21)
        Me.btnClose.TabIndex = 328
        Me.btnClose.Text = "Close"
        '
        'btnReset
        '
        Me.btnReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReset.Location = New System.Drawing.Point(91, 11)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(77, 21)
        Me.btnReset.TabIndex = 329
        Me.btnReset.Text = "Reset"
        '
        'frmGSTDocumentList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "frmGSTDocumentList"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "frmGSTDocumentList"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageView1.ResumeLayout(False)
        Me.RadPageViewPage1.ResumeLayout(False)
        Me.RadPageViewPage1.PerformLayout()
        CType(Me.RadGroupBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox5.ResumeLayout(False)
        Me.RadGroupBox5.PerformLayout()
        CType(Me.txtToDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFromDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ddlTransaction, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ddlInvoiceType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lbltype, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ddlType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage2.ResumeLayout(False)
        CType(Me.gv.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage3.ResumeLayout(False)
        CType(Me.gvDetails.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvDetails, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents RadPageView1 As RadPageView
    Friend WithEvents RadPageViewPage1 As RadPageViewPage
    Friend WithEvents RadPageViewPage2 As RadPageViewPage
    Friend WithEvents RadSplitButton1 As RadSplitButton
    Friend WithEvents QExpExcel As RadMenuItem
    Friend WithEvents btnGO As RadButton
    Friend WithEvents btnClose As RadButton
    Friend WithEvents btnReset As RadButton
    Friend WithEvents RadLabel3 As common.Controls.MyLabel
    Friend WithEvents txtFromDate As RadDateTimePicker
    Friend WithEvents RadLabel2 As common.Controls.MyLabel
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents ddlTransaction As common.Controls.MyComboBox
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents ddlInvoiceType As common.Controls.MyComboBox
    Friend WithEvents lbltype As common.Controls.MyLabel
    Friend WithEvents ddlType As common.Controls.MyComboBox
    Friend WithEvents RadGroupBox5 As RadGroupBox
    Friend WithEvents txtToDate As RadDateTimePicker
    Friend WithEvents gv As common.UserControls.MyRadGridView
    Friend WithEvents RadPageViewPage3 As RadPageViewPage
    Friend WithEvents gvDetails As common.UserControls.MyRadGridView
End Class
