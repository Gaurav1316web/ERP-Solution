<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class rptProductCreditSaleReport
    'Inherits System.Windows.Forms.Form
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
        Dim TableViewDefinition2 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.RadPageViewPage1 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.RadGroupBox5 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rbtnNonTaxable = New common.Controls.MyRadioButton()
        Me.rbtnTaxable = New common.Controls.MyRadioButton()
        Me.RadGroupBox3 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rdbDelete = New common.Controls.MyRadioButton()
        Me.rdbCancel = New common.Controls.MyRadioButton()
        Me.chkShowAmt = New System.Windows.Forms.CheckBox()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rdbInvoiceCount = New common.Controls.MyRadioButton()
        Me.rdbProductSummary = New common.Controls.MyRadioButton()
        Me.rdbItemWiseCustomer = New common.Controls.MyRadioButton()
        Me.rdbSaleReturn = New common.Controls.MyRadioButton()
        Me.rdbGstInvoice = New common.Controls.MyRadioButton()
        Me.rdbCancelInvoice = New common.Controls.MyRadioButton()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.TxtCustomerType = New common.UserControls.txtMultiSelectFinder()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.TxtSubLocation = New common.UserControls.txtMultiSelectFinder()
        Me.MyLabel4 = New common.Controls.MyLabel()
        Me.txtItem = New common.UserControls.txtMultiSelectFinder()
        Me.TxtTransaction = New common.UserControls.txtMultiSelectFinder()
        Me.MyLabel10 = New common.Controls.MyLabel()
        Me.RadGroupBox2 = New Telerik.WinControls.UI.RadGroupBox()
        Me.rbtnDetail = New common.Controls.MyRadioButton()
        Me.rbtnSummary = New common.Controls.MyRadioButton()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.txtMultiCustomer = New common.UserControls.txtMultiSelectFinder()
        Me.RadGroupBox4 = New Telerik.WinControls.UI.RadGroupBox()
        Me.MyLabel17 = New common.Controls.MyLabel()
        Me.MyLabel18 = New common.Controls.MyLabel()
        Me.txtToDate = New Telerik.WinControls.UI.RadDateTimePicker()
        Me.txtfDate = New Telerik.WinControls.UI.RadDateTimePicker()
        Me.RadPageViewPage3 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.gvdata = New common.UserControls.MyRadGridView()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.RadSplitButton1 = New Telerik.WinControls.UI.RadSplitButton()
        Me.rmenuExport = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmenuPDF = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnReset = New Telerik.WinControls.UI.RadButton()
        Me.btnGo = New Telerik.WinControls.UI.RadButton()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.RadPageViewPage1.SuspendLayout()
        CType(Me.RadGroupBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox5.SuspendLayout()
        CType(Me.rbtnNonTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox3.SuspendLayout()
        CType(Me.rdbDelete, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.rdbInvoiceCount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbProductSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbItemWiseCustomer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbSaleReturn, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbGstInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rdbCancelInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox2.SuspendLayout()
        CType(Me.rbtnDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox4.SuspendLayout()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtToDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtfDate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage3.SuspendLayout()
        CType(Me.gvdata, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvdata.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RadSplitButton1)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnReset)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnGo)
        Me.SplitContainer1.Size = New System.Drawing.Size(800, 450)
        Me.SplitContainer1.SplitterDistance = 413
        Me.SplitContainer1.TabIndex = 0
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage1)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage3)
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.RadPageViewPage1
        Me.RadPageView1.Size = New System.Drawing.Size(800, 413)
        Me.RadPageView1.TabIndex = 77
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        '
        'RadPageViewPage1
        '
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox5)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox3)
        Me.RadPageViewPage1.Controls.Add(Me.chkShowAmt)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox1)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel2)
        Me.RadPageViewPage1.Controls.Add(Me.TxtCustomerType)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.TxtSubLocation)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel4)
        Me.RadPageViewPage1.Controls.Add(Me.txtItem)
        Me.RadPageViewPage1.Controls.Add(Me.TxtTransaction)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel10)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox2)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel13)
        Me.RadPageViewPage1.Controls.Add(Me.txtMultiCustomer)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox4)
        Me.RadPageViewPage1.ItemSize = New System.Drawing.SizeF(46.0!, 28.0!)
        Me.RadPageViewPage1.Location = New System.Drawing.Point(10, 37)
        Me.RadPageViewPage1.Name = "RadPageViewPage1"
        Me.RadPageViewPage1.Size = New System.Drawing.Size(779, 365)
        Me.RadPageViewPage1.Text = "Filters"
        '
        'RadGroupBox5
        '
        Me.RadGroupBox5.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox5.Controls.Add(Me.rbtnNonTaxable)
        Me.RadGroupBox5.Controls.Add(Me.rbtnTaxable)
        Me.RadGroupBox5.HeaderText = "Invoice Type"
        Me.RadGroupBox5.Location = New System.Drawing.Point(443, 5)
        Me.RadGroupBox5.Name = "RadGroupBox5"
        Me.RadGroupBox5.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox5.Size = New System.Drawing.Size(158, 40)
        Me.RadGroupBox5.TabIndex = 1508
        Me.RadGroupBox5.Text = "Invoice Type"
        Me.RadGroupBox5.Visible = False
        '
        'rbtnNonTaxable
        '
        Me.rbtnNonTaxable.Location = New System.Drawing.Point(6, 11)
        Me.rbtnNonTaxable.MyLinkLable1 = Nothing
        Me.rbtnNonTaxable.MyLinkLable2 = Nothing
        Me.rbtnNonTaxable.Name = "rbtnNonTaxable"
        Me.rbtnNonTaxable.Size = New System.Drawing.Size(84, 18)
        Me.rbtnNonTaxable.TabIndex = 396
        Me.rbtnNonTaxable.TabStop = False
        Me.rbtnNonTaxable.Text = "Non-Taxable"
        '
        'rbtnTaxable
        '
        Me.rbtnTaxable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rbtnTaxable.Location = New System.Drawing.Point(94, 11)
        Me.rbtnTaxable.MyLinkLable1 = Nothing
        Me.rbtnTaxable.MyLinkLable2 = Nothing
        Me.rbtnTaxable.Name = "rbtnTaxable"
        Me.rbtnTaxable.Size = New System.Drawing.Size(58, 18)
        Me.rbtnTaxable.TabIndex = 391
        Me.rbtnTaxable.Text = "Taxable"
        Me.rbtnTaxable.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'RadGroupBox3
        '
        Me.RadGroupBox3.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox3.Controls.Add(Me.rdbDelete)
        Me.RadGroupBox3.Controls.Add(Me.rdbCancel)
        Me.RadGroupBox3.HeaderText = ""
        Me.RadGroupBox3.Location = New System.Drawing.Point(611, 6)
        Me.RadGroupBox3.Name = "RadGroupBox3"
        Me.RadGroupBox3.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox3.Size = New System.Drawing.Size(158, 39)
        Me.RadGroupBox3.TabIndex = 1507
        Me.RadGroupBox3.Visible = False
        '
        'rdbDelete
        '
        Me.rdbDelete.Location = New System.Drawing.Point(15, 12)
        Me.rdbDelete.MyLinkLable1 = Nothing
        Me.rdbDelete.MyLinkLable2 = Nothing
        Me.rdbDelete.Name = "rdbDelete"
        Me.rdbDelete.Size = New System.Drawing.Size(52, 18)
        Me.rdbDelete.TabIndex = 396
        Me.rdbDelete.TabStop = False
        Me.rdbDelete.Text = "Delete"
        '
        'rdbCancel
        '
        Me.rdbCancel.Location = New System.Drawing.Point(79, 12)
        Me.rdbCancel.MyLinkLable1 = Nothing
        Me.rdbCancel.MyLinkLable2 = Nothing
        Me.rdbCancel.Name = "rdbCancel"
        Me.rdbCancel.Size = New System.Drawing.Size(53, 18)
        Me.rdbCancel.TabIndex = 391
        Me.rdbCancel.TabStop = False
        Me.rdbCancel.Text = "Cancel"
        '
        'chkShowAmt
        '
        Me.chkShowAmt.AutoSize = True
        Me.chkShowAmt.Location = New System.Drawing.Point(276, 19)
        Me.chkShowAmt.Name = "chkShowAmt"
        Me.chkShowAmt.Size = New System.Drawing.Size(99, 17)
        Me.chkShowAmt.TabIndex = 1506
        Me.chkShowAmt.Text = "Show Amount"
        Me.chkShowAmt.UseVisualStyleBackColor = True
        Me.chkShowAmt.Visible = False
        '
        'RadGroupBox1
        '
        Me.RadGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox1.Controls.Add(Me.rdbInvoiceCount)
        Me.RadGroupBox1.Controls.Add(Me.rdbProductSummary)
        Me.RadGroupBox1.Controls.Add(Me.rdbItemWiseCustomer)
        Me.RadGroupBox1.Controls.Add(Me.rdbSaleReturn)
        Me.RadGroupBox1.Controls.Add(Me.rdbGstInvoice)
        Me.RadGroupBox1.Controls.Add(Me.rdbCancelInvoice)
        Me.RadGroupBox1.HeaderText = ""
        Me.RadGroupBox1.Location = New System.Drawing.Point(419, 52)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox1.Size = New System.Drawing.Size(357, 65)
        Me.RadGroupBox1.TabIndex = 455
        Me.RadGroupBox1.Visible = False
        '
        'rdbInvoiceCount
        '
        Me.rdbInvoiceCount.Location = New System.Drawing.Point(254, 33)
        Me.rdbInvoiceCount.MyLinkLable1 = Nothing
        Me.rdbInvoiceCount.MyLinkLable2 = Nothing
        Me.rdbInvoiceCount.Name = "rdbInvoiceCount"
        Me.rdbInvoiceCount.Size = New System.Drawing.Size(89, 18)
        Me.rdbInvoiceCount.TabIndex = 400
        Me.rdbInvoiceCount.TabStop = False
        Me.rdbInvoiceCount.Text = "Invoice Count"
        '
        'rdbProductSummary
        '
        Me.rdbProductSummary.Location = New System.Drawing.Point(139, 34)
        Me.rdbProductSummary.MyLinkLable1 = Nothing
        Me.rdbProductSummary.MyLinkLable2 = Nothing
        Me.rdbProductSummary.Name = "rdbProductSummary"
        Me.rdbProductSummary.Size = New System.Drawing.Size(109, 18)
        Me.rdbProductSummary.TabIndex = 399
        Me.rdbProductSummary.TabStop = False
        Me.rdbProductSummary.Text = "Product Summary"
        '
        'rdbItemWiseCustomer
        '
        Me.rdbItemWiseCustomer.Location = New System.Drawing.Point(15, 34)
        Me.rdbItemWiseCustomer.MyLinkLable1 = Nothing
        Me.rdbItemWiseCustomer.MyLinkLable2 = Nothing
        Me.rdbItemWiseCustomer.Name = "rdbItemWiseCustomer"
        Me.rdbItemWiseCustomer.Size = New System.Drawing.Size(118, 18)
        Me.rdbItemWiseCustomer.TabIndex = 398
        Me.rdbItemWiseCustomer.Text = "ItemWise Customer"
        '
        'rdbSaleReturn
        '
        Me.rdbSaleReturn.Location = New System.Drawing.Point(212, 9)
        Me.rdbSaleReturn.MyLinkLable1 = Nothing
        Me.rdbSaleReturn.MyLinkLable2 = Nothing
        Me.rdbSaleReturn.Name = "rdbSaleReturn"
        Me.rdbSaleReturn.Size = New System.Drawing.Size(77, 18)
        Me.rdbSaleReturn.TabIndex = 397
        Me.rdbSaleReturn.TabStop = False
        Me.rdbSaleReturn.Text = "Sale Return"
        '
        'rdbGstInvoice
        '
        Me.rdbGstInvoice.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rdbGstInvoice.Location = New System.Drawing.Point(15, 9)
        Me.rdbGstInvoice.MyLinkLable1 = Nothing
        Me.rdbGstInvoice.MyLinkLable2 = Nothing
        Me.rdbGstInvoice.Name = "rdbGstInvoice"
        Me.rdbGstInvoice.Size = New System.Drawing.Size(78, 18)
        Me.rdbGstInvoice.TabIndex = 396
        Me.rdbGstInvoice.Text = "GST Invoice"
        Me.rdbGstInvoice.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'rdbCancelInvoice
        '
        Me.rdbCancelInvoice.Location = New System.Drawing.Point(99, 9)
        Me.rdbCancelInvoice.MyLinkLable1 = Nothing
        Me.rdbCancelInvoice.MyLinkLable2 = Nothing
        Me.rdbCancelInvoice.Name = "rdbCancelInvoice"
        Me.rdbCancelInvoice.Size = New System.Drawing.Size(107, 18)
        Me.rdbCancelInvoice.TabIndex = 391
        Me.rdbCancelInvoice.TabStop = False
        Me.rdbCancelInvoice.Text = "Cancelled Invoice"
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(13, 73)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(82, 18)
        Me.MyLabel2.TabIndex = 453
        Me.MyLabel2.Text = "Customer Type"
        Me.MyLabel2.Visible = False
        '
        'TxtCustomerType
        '
        Me.TxtCustomerType.arrDispalyMember = Nothing
        Me.TxtCustomerType.arrValueMember = Nothing
        Me.TxtCustomerType.Location = New System.Drawing.Point(97, 74)
        Me.TxtCustomerType.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCustomerType.MyLinkLable1 = Me.MyLabel2
        Me.TxtCustomerType.MyLinkLable2 = Nothing
        Me.TxtCustomerType.MyNullText = "All"
        Me.TxtCustomerType.Name = "TxtCustomerType"
        Me.TxtCustomerType.Size = New System.Drawing.Size(299, 19)
        Me.TxtCustomerType.TabIndex = 452
        Me.TxtCustomerType.Visible = False
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(13, 51)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(71, 18)
        Me.MyLabel1.TabIndex = 451
        Me.MyLabel1.Text = "Sub Location"
        Me.MyLabel1.Visible = False
        '
        'TxtSubLocation
        '
        Me.TxtSubLocation.arrDispalyMember = Nothing
        Me.TxtSubLocation.arrValueMember = Nothing
        Me.TxtSubLocation.Location = New System.Drawing.Point(97, 52)
        Me.TxtSubLocation.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtSubLocation.MyLinkLable1 = Me.MyLabel1
        Me.TxtSubLocation.MyLinkLable2 = Nothing
        Me.TxtSubLocation.MyNullText = "All"
        Me.TxtSubLocation.Name = "TxtSubLocation"
        Me.TxtSubLocation.Size = New System.Drawing.Size(299, 19)
        Me.TxtSubLocation.TabIndex = 450
        Me.TxtSubLocation.Visible = False
        '
        'MyLabel4
        '
        Me.MyLabel4.FieldName = Nothing
        Me.MyLabel4.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(12, 146)
        Me.MyLabel4.Margin = New System.Windows.Forms.Padding(4)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(29, 18)
        Me.MyLabel4.TabIndex = 449
        Me.MyLabel4.Text = "Item"
        Me.MyLabel4.Visible = False
        '
        'txtItem
        '
        Me.txtItem.arrDispalyMember = Nothing
        Me.txtItem.arrValueMember = Nothing
        Me.txtItem.Location = New System.Drawing.Point(98, 146)
        Me.txtItem.Margin = New System.Windows.Forms.Padding(4)
        Me.txtItem.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItem.MyLinkLable1 = Nothing
        Me.txtItem.MyLinkLable2 = Nothing
        Me.txtItem.MyNullText = "All"
        Me.txtItem.Name = "txtItem"
        Me.txtItem.Size = New System.Drawing.Size(299, 19)
        Me.txtItem.TabIndex = 448
        Me.txtItem.Visible = False
        '
        'TxtTransaction
        '
        Me.TxtTransaction.arrDispalyMember = Nothing
        Me.TxtTransaction.arrValueMember = Nothing
        Me.TxtTransaction.Location = New System.Drawing.Point(98, 122)
        Me.TxtTransaction.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTransaction.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtTransaction.MyLinkLable1 = Me.MyLabel10
        Me.TxtTransaction.MyLinkLable2 = Nothing
        Me.TxtTransaction.MyNullText = "All"
        Me.TxtTransaction.Name = "TxtTransaction"
        Me.TxtTransaction.Size = New System.Drawing.Size(299, 19)
        Me.TxtTransaction.TabIndex = 443
        Me.TxtTransaction.Visible = False
        '
        'MyLabel10
        '
        Me.MyLabel10.FieldName = Nothing
        Me.MyLabel10.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.Location = New System.Drawing.Point(13, 123)
        Me.MyLabel10.Margin = New System.Windows.Forms.Padding(4)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(64, 18)
        Me.MyLabel10.TabIndex = 442
        Me.MyLabel10.Text = "Transaction"
        Me.MyLabel10.Visible = False
        '
        'RadGroupBox2
        '
        Me.RadGroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox2.Controls.Add(Me.rbtnDetail)
        Me.RadGroupBox2.Controls.Add(Me.rbtnSummary)
        Me.RadGroupBox2.HeaderText = "Data View Preference"
        Me.RadGroupBox2.Location = New System.Drawing.Point(449, 184)
        Me.RadGroupBox2.Name = "RadGroupBox2"
        Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox2.Size = New System.Drawing.Size(158, 39)
        Me.RadGroupBox2.TabIndex = 441
        Me.RadGroupBox2.Text = "Data View Preference"
        Me.RadGroupBox2.Visible = False
        '
        'rbtnDetail
        '
        Me.rbtnDetail.Location = New System.Drawing.Point(15, 9)
        Me.rbtnDetail.MyLinkLable1 = Nothing
        Me.rbtnDetail.MyLinkLable2 = Nothing
        Me.rbtnDetail.Name = "rbtnDetail"
        Me.rbtnDetail.Size = New System.Drawing.Size(49, 18)
        Me.rbtnDetail.TabIndex = 396
        Me.rbtnDetail.TabStop = False
        Me.rbtnDetail.Text = "Detail"
        '
        'rbtnSummary
        '
        Me.rbtnSummary.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rbtnSummary.Location = New System.Drawing.Point(79, 9)
        Me.rbtnSummary.MyLinkLable1 = Nothing
        Me.rbtnSummary.MyLinkLable2 = Nothing
        Me.rbtnSummary.Name = "rbtnSummary"
        Me.rbtnSummary.Size = New System.Drawing.Size(67, 18)
        Me.rbtnSummary.TabIndex = 391
        Me.rbtnSummary.Text = "Summary"
        Me.rbtnSummary.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(12, 97)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(55, 18)
        Me.MyLabel13.TabIndex = 334
        Me.MyLabel13.Text = "Customer"
        Me.MyLabel13.Visible = False
        '
        'txtMultiCustomer
        '
        Me.txtMultiCustomer.arrDispalyMember = Nothing
        Me.txtMultiCustomer.arrValueMember = Nothing
        Me.txtMultiCustomer.Location = New System.Drawing.Point(98, 98)
        Me.txtMultiCustomer.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMultiCustomer.MyLinkLable1 = Me.MyLabel13
        Me.txtMultiCustomer.MyLinkLable2 = Nothing
        Me.txtMultiCustomer.MyNullText = "All"
        Me.txtMultiCustomer.Name = "txtMultiCustomer"
        Me.txtMultiCustomer.Size = New System.Drawing.Size(299, 19)
        Me.txtMultiCustomer.TabIndex = 333
        Me.txtMultiCustomer.Visible = False
        '
        'RadGroupBox4
        '
        Me.RadGroupBox4.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox4.Controls.Add(Me.MyLabel17)
        Me.RadGroupBox4.Controls.Add(Me.MyLabel18)
        Me.RadGroupBox4.Controls.Add(Me.txtToDate)
        Me.RadGroupBox4.Controls.Add(Me.txtfDate)
        Me.RadGroupBox4.HeaderText = "Date Range"
        Me.RadGroupBox4.Location = New System.Drawing.Point(16, 3)
        Me.RadGroupBox4.Name = "RadGroupBox4"
        Me.RadGroupBox4.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox4.Size = New System.Drawing.Size(254, 42)
        Me.RadGroupBox4.TabIndex = 53
        Me.RadGroupBox4.Text = "Date Range"
        '
        'MyLabel17
        '
        Me.MyLabel17.FieldName = Nothing
        Me.MyLabel17.Location = New System.Drawing.Point(135, 16)
        Me.MyLabel17.Name = "MyLabel17"
        Me.MyLabel17.Size = New System.Drawing.Size(19, 18)
        Me.MyLabel17.TabIndex = 3
        Me.MyLabel17.Text = "To"
        '
        'MyLabel18
        '
        Me.MyLabel18.FieldName = Nothing
        Me.MyLabel18.Location = New System.Drawing.Point(5, 16)
        Me.MyLabel18.Name = "MyLabel18"
        Me.MyLabel18.Size = New System.Drawing.Size(32, 18)
        Me.MyLabel18.TabIndex = 2
        Me.MyLabel18.Text = "From"
        '
        'txtToDate
        '
        Me.txtToDate.CustomFormat = "dd/MM/yyyy"
        Me.txtToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtToDate.Location = New System.Drawing.Point(159, 15)
        Me.txtToDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDate.Name = "txtToDate"
        Me.txtToDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDate.Size = New System.Drawing.Size(85, 20)
        Me.txtToDate.TabIndex = 1
        Me.txtToDate.TabStop = False
        Me.txtToDate.Text = "24/10/2011"
        Me.txtToDate.Value = New Date(2011, 10, 24, 2, 29, 0, 265)
        '
        'txtfDate
        '
        Me.txtfDate.CustomFormat = "dd/MM/yyyy"
        Me.txtfDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtfDate.Location = New System.Drawing.Point(44, 15)
        Me.txtfDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtfDate.Name = "txtfDate"
        Me.txtfDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtfDate.Size = New System.Drawing.Size(88, 20)
        Me.txtfDate.TabIndex = 0
        Me.txtfDate.TabStop = False
        Me.txtfDate.Text = "24/10/2011"
        Me.txtfDate.Value = New Date(2011, 10, 24, 2, 29, 0, 265)
        '
        'RadPageViewPage3
        '
        Me.RadPageViewPage3.Controls.Add(Me.gvdata)
        Me.RadPageViewPage3.ItemSize = New System.Drawing.SizeF(50.0!, 28.0!)
        Me.RadPageViewPage3.Location = New System.Drawing.Point(10, 37)
        Me.RadPageViewPage3.Name = "RadPageViewPage3"
        Me.RadPageViewPage3.Size = New System.Drawing.Size(779, 365)
        Me.RadPageViewPage3.Text = "Report"
        '
        'gvdata
        '
        Me.gvdata.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.gvdata.Cursor = System.Windows.Forms.Cursors.Default
        Me.gvdata.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gvdata.Font = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.gvdata.ForeColor = System.Drawing.Color.Black
        Me.gvdata.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.gvdata.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.gvdata.MasterTemplate.AllowAddNewRow = False
        Me.gvdata.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gvdata.MasterTemplate.ShowHeaderCellButtons = True
        Me.gvdata.MasterTemplate.ViewDefinition = TableViewDefinition2
        Me.gvdata.MyExportAPI = False
        Me.gvdata.MyExportFilePath = ""
        Me.gvdata.MyStopExport = False
        Me.gvdata.Name = "gvdata"
        Me.gvdata.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gvdata.ShowHeaderCellButtons = True
        Me.gvdata.Size = New System.Drawing.Size(779, 365)
        Me.gvdata.TabIndex = 1
        Me.gvdata.VarID = ""
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Location = New System.Drawing.Point(722, 7)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(68, 18)
        Me.btnClose.TabIndex = 170
        Me.btnClose.Text = "Close"
        '
        'RadSplitButton1
        '
        Me.RadSplitButton1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RadSplitButton1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.rmenuExport, Me.rmenuPDF})
        Me.RadSplitButton1.Location = New System.Drawing.Point(151, 7)
        Me.RadSplitButton1.Name = "RadSplitButton1"
        Me.RadSplitButton1.Size = New System.Drawing.Size(68, 18)
        Me.RadSplitButton1.TabIndex = 169
        Me.RadSplitButton1.Text = "Export"
        '
        'rmenuExport
        '
        Me.rmenuExport.Name = "rmenuExport"
        Me.rmenuExport.Text = "Export"
        Me.rmenuExport.UseCompatibleTextRendering = False
        '
        'rmenuPDF
        '
        Me.rmenuPDF.Name = "rmenuPDF"
        Me.rmenuPDF.Text = "PDF"
        Me.rmenuPDF.UseCompatibleTextRendering = False
        '
        'btnReset
        '
        Me.btnReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnReset.Location = New System.Drawing.Point(81, 7)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(68, 18)
        Me.btnReset.TabIndex = 168
        Me.btnReset.Text = "Reset"
        '
        'btnGo
        '
        Me.btnGo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGo.Location = New System.Drawing.Point(11, 7)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(68, 18)
        Me.btnGo.TabIndex = 167
        Me.btnGo.Text = ">>>"
        '
        'rptProductCreditSaleReport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "rptProductCreditSaleReport"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "rptProductCreditSaleReport"
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
        CType(Me.rbtnNonTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox3.ResumeLayout(False)
        Me.RadGroupBox3.PerformLayout()
        CType(Me.rdbDelete, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.rdbInvoiceCount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbProductSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbItemWiseCustomer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbSaleReturn, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbGstInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rdbCancelInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox2.ResumeLayout(False)
        Me.RadGroupBox2.PerformLayout()
        CType(Me.rbtnDetail, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnSummary, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox4.ResumeLayout(False)
        Me.RadGroupBox4.PerformLayout()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtToDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtfDate, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage3.ResumeLayout(False)
        CType(Me.gvdata.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvdata, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadSplitButton1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents RadPageView1 As RadPageView
    Friend WithEvents RadPageViewPage1 As RadPageViewPage
    Friend WithEvents RadGroupBox5 As RadGroupBox
    Friend WithEvents rbtnNonTaxable As common.Controls.MyRadioButton
    Friend WithEvents rbtnTaxable As common.Controls.MyRadioButton
    Friend WithEvents RadGroupBox3 As RadGroupBox
    Friend WithEvents rdbDelete As common.Controls.MyRadioButton
    Friend WithEvents rdbCancel As common.Controls.MyRadioButton
    Friend WithEvents chkShowAmt As CheckBox
    Friend WithEvents RadGroupBox1 As RadGroupBox
    Friend WithEvents rdbInvoiceCount As common.Controls.MyRadioButton
    Friend WithEvents rdbProductSummary As common.Controls.MyRadioButton
    Friend WithEvents rdbItemWiseCustomer As common.Controls.MyRadioButton
    Friend WithEvents rdbSaleReturn As common.Controls.MyRadioButton
    Friend WithEvents rdbGstInvoice As common.Controls.MyRadioButton
    Friend WithEvents rdbCancelInvoice As common.Controls.MyRadioButton
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents TxtCustomerType As common.UserControls.txtMultiSelectFinder
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents TxtSubLocation As common.UserControls.txtMultiSelectFinder
    Friend WithEvents MyLabel4 As common.Controls.MyLabel
    Friend WithEvents txtItem As common.UserControls.txtMultiSelectFinder
    Friend WithEvents TxtTransaction As common.UserControls.txtMultiSelectFinder
    Friend WithEvents MyLabel10 As common.Controls.MyLabel
    Friend WithEvents RadGroupBox2 As RadGroupBox
    Friend WithEvents rbtnDetail As common.Controls.MyRadioButton
    Friend WithEvents rbtnSummary As common.Controls.MyRadioButton
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents txtMultiCustomer As common.UserControls.txtMultiSelectFinder
    Friend WithEvents RadGroupBox4 As RadGroupBox
    Friend WithEvents MyLabel17 As common.Controls.MyLabel
    Friend WithEvents MyLabel18 As common.Controls.MyLabel
    Friend WithEvents txtToDate As RadDateTimePicker
    Friend WithEvents txtfDate As RadDateTimePicker
    Friend WithEvents RadPageViewPage3 As RadPageViewPage
    Friend WithEvents gvdata As common.UserControls.MyRadGridView
    Friend WithEvents btnClose As RadButton
    Friend WithEvents RadSplitButton1 As RadSplitButton
    Friend WithEvents rmenuExport As RadMenuItem
    Friend WithEvents rmenuPDF As RadMenuItem
    Friend WithEvents btnReset As RadButton
    Friend WithEvents btnGo As RadButton
End Class
