<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmARInvoiceEntry
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmARInvoiceEntry))
        Dim TableViewDefinition1 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition2 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition3 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.RadPageViewPage1 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.chkEinvoice = New Telerik.WinControls.UI.RadCheckBox()
        Me.pnlLocation = New System.Windows.Forms.Panel()
        Me.RadLabel7 = New Telerik.WinControls.UI.RadLabel()
        Me.txtLocationPrefixName = New common.Controls.MyLabel()
        Me.txtLocationPrefix = New common.UserControls.txtFinder()
        Me.MyLabel10 = New common.Controls.MyLabel()
        Me.txtDataAndTimeSelection = New common.Controls.MyDateTimePicker()
        Me.MyLabel9 = New common.Controls.MyLabel()
        Me.txtTapalNo = New common.Controls.MyTextBox()
        Me.chkAgainstServiceInvoice = New Telerik.WinControls.UI.RadCheckBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.txtSercurityReceiptNo = New common.UserControls.txtFinder()
        Me.MyLabel8 = New common.Controls.MyLabel()
        Me.chkAgainstSecurityReceiptNo = New Telerik.WinControls.UI.RadCheckBox()
        Me.txtServiceVisitCode = New common.Controls.MyTextBox()
        Me.txtVCGNo = New common.Controls.MyTextBox()
        Me.txtSaleReturnNo = New common.Controls.MyTextBox()
        Me.txtSaleInvoice = New common.Controls.MyTextBox()
        Me.MyLabel7 = New common.Controls.MyLabel()
        Me.ddlSecDepositType = New common.Controls.MyComboBox()
        Me.chkSecurityDposit = New Telerik.WinControls.UI.RadCheckBox()
        Me.MyLabel3 = New common.Controls.MyLabel()
        Me.LblLocDesp = New common.Controls.MyLabel()
        Me.MyLabel6 = New common.Controls.MyLabel()
        Me.MyLabel5 = New common.Controls.MyLabel()
        Me.lblReturnType = New common.Controls.MyLabel()
        Me.ddlReturnType = New common.Controls.MyComboBox()
        Me.pnlPCJ = New System.Windows.Forms.Panel()
        Me.MyLabel4 = New common.Controls.MyLabel()
        Me.fndProject = New common.UserControls.txtFinder()
        Me.lblProject = New common.Controls.MyLabel()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.txtlocation = New common.UserControls.txtFinder()
        Me.RadLabel3 = New Telerik.WinControls.UI.RadLabel()
        Me.btnDrillDown = New Telerik.WinControls.UI.RadButton()
        Me.txtACSet = New common.UserControls.txtFinder()
        Me.TxtCustomer = New common.UserControls.txtFinder()
        Me.txtRefDocNo = New common.UserControls.txtFinder()
        Me.UsLock1 = New common.usLock()
        Me.RadLabel15 = New common.Controls.MyLabel()
        Me.RadLabel12 = New common.Controls.MyLabel()
        Me.lblTotRAmt1 = New common.Controls.MyLabel()
        Me.RadGroupBox2 = New Telerik.WinControls.UI.RadGroupBox()
        Me.gv1 = New common.UserControls.MyRadGridView()
        Me.RadLabel8 = New common.Controls.MyLabel()
        Me.RadLabel6 = New common.Controls.MyLabel()
        Me.RadLabel5 = New common.Controls.MyLabel()
        Me.RadLabel4 = New common.Controls.MyLabel()
        Me.chkOnHold = New Telerik.WinControls.UI.RadCheckBox()
        Me.lblVendorName = New common.Controls.MyLabel()
        Me.RadLabel2 = New common.Controls.MyLabel()
        Me.RadLabel1 = New common.Controls.MyLabel()
        Me.btnAddNew = New Telerik.WinControls.UI.RadButton()
        Me.txtDocNo = New common.UserControls.txtNavigator()
        Me.txtOrderNo = New common.Controls.MyTextBox()
        Me.cboDocType = New common.Controls.MyComboBox()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.txtDesc = New common.Controls.MyTextBox()
        Me.RadPageViewPage2 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbtnTaxCalManual = New common.Controls.MyRadioButton()
        Me.rbtnTaxCalAutomatic = New common.Controls.MyRadioButton()
        Me.txtTaxGroup = New common.UserControls.txtFinder()
        Me.RadLabel10 = New common.Controls.MyLabel()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtTermCode = New common.UserControls.txtFinder()
        Me.RadLabel16 = New common.Controls.MyLabel()
        Me.txtDueDate = New common.Controls.MyDateTimePicker()
        Me.RadLabel17 = New common.Controls.MyLabel()
        Me.lblTermName = New common.Controls.MyLabel()
        Me.gv2 = New common.UserControls.MyRadGridView()
        Me.lblTaxGrpName = New common.Controls.MyLabel()
        Me.RadLabel11 = New common.Controls.MyLabel()
        Me.RadPageViewPage3 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.gvAC = New common.UserControls.MyRadGridView()
        Me.RadLabel31 = New common.Controls.MyLabel()
        Me.lblAddCharges = New common.Controls.MyLabel()
        Me.pvpCustomFields = New Telerik.WinControls.UI.RadPageViewPage()
        Me.UcCustomFields1 = New XpertERPEngine.ucCustomFields()
        Me.Attachments = New Telerik.WinControls.UI.RadPageViewPage()
        Me.UcAttachment1 = New XpertERPEngine.ucAttachment()
        Me.RadPageViewPage4 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.btnAdminCancel = New Telerik.WinControls.UI.RadButton()
        Me.MyLabel14 = New common.Controls.MyLabel()
        Me.btnCancel = New Telerik.WinControls.UI.RadButton()
        Me.lblTotalSubsidyAmt = New common.Controls.MyLabel()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.lblRoundoffAmount = New common.Controls.MyLabel()
        Me.pnlCurrConv = New System.Windows.Forms.Panel()
        Me.txtConversionRate = New common.MyNumBox()
        Me.txtCurrencyCode = New common.UserControls.txtFinder()
        Me.lblEffectiveFrom = New common.Controls.MyLabel()
        Me.txtApplicableFrom = New common.Controls.MyLabel()
        Me.lblCurrency = New common.Controls.MyLabel()
        Me.lblConvRate = New common.Controls.MyLabel()
        Me.btnReverse = New Telerik.WinControls.UI.RadButton()
        Me.RadLabel32 = New common.Controls.MyLabel()
        Me.lblAddCharges1 = New common.Controls.MyLabel()
        Me.RadLabel9 = New common.Controls.MyLabel()
        Me.RadLabel27 = New common.Controls.MyLabel()
        Me.lblTotRAmt = New common.Controls.MyLabel()
        Me.RadLabel25 = New common.Controls.MyLabel()
        Me.lblTaxAmt = New common.Controls.MyLabel()
        Me.lblAmtAfterDiscount = New common.Controls.MyLabel()
        Me.lblDiscountAmt = New common.Controls.MyLabel()
        Me.lblAmtWithDiscount = New common.Controls.MyLabel()
        Me.RadLabel22 = New common.Controls.MyLabel()
        Me.RadLabel19 = New common.Controls.MyLabel()
        Me.RadPageViewPage5 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.RadGroupBox3 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtAckDate = New common.Controls.MyDateTimePicker()
        Me.EinvoiceBtnUpdate = New Telerik.WinControls.UI.RadButton()
        Me.EInvoiceQrCode = New common.Controls.MyTextBox()
        Me.MyLabel11 = New common.Controls.MyLabel()
        Me.MyLabel17 = New common.Controls.MyLabel()
        Me.EinvoiceAckNo = New common.Controls.MyTextBox()
        Me.MyLabel12 = New common.Controls.MyLabel()
        Me.EInvoiceIRNNo = New common.Controls.MyTextBox()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.btnHistory = New Telerik.WinControls.UI.RadButton()
        Me.butCostCenterAndHirerachy_Update_AfterPost = New Telerik.WinControls.UI.RadButton()
        Me.btnPrintServiceInvoice = New Telerik.WinControls.UI.RadButton()
        Me.btnsetting = New Telerik.WinControls.UI.RadSplitButton()
        Me.btnsend = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnSendForApproval = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnPrint = New Telerik.WinControls.UI.RadButton()
        Me.btnDelete = New Telerik.WinControls.UI.RadButton()
        Me.btnPost = New Telerik.WinControls.UI.RadButton()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.btnSave = New Telerik.WinControls.UI.RadButton()
        Me.RadMenuItem1 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem2 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem4 = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuExportARInvoiceTrans = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuExportCN = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuExportDN = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuExportOpeningCredit = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuExortOpeningDebit = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmiExportOPInvoiceForTCS = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem3 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem5 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem6 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem7 = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuImportCR = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuImportDN = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuARInvoiceImport = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmiOPInvoiceForTCS = New Telerik.WinControls.UI.RadMenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.RadMenu1 = New Telerik.WinControls.UI.RadMenu()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.RadPageViewPage1.SuspendLayout()
        CType(Me.chkEinvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLocation.SuspendLayout()
        CType(Me.RadLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLocationPrefixName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDataAndTimeSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTapalNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAgainstServiceInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAgainstSecurityReceiptNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtServiceVisitCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVCGNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSaleReturnNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSaleInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ddlSecDepositType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSecurityDposit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LblLocDesp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReturnType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ddlReturnType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPCJ.SuspendLayout()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblProject, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDrillDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotRAmt1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox2.SuspendLayout()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkOnHold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVendorName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtOrderNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboDocType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.rbtnTaxCalManual, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnTaxCalAutomatic, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.RadLabel16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDueDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTermName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv2.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTaxGrpName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.gvAC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvAC.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAddCharges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pvpCustomFields.SuspendLayout()
        Me.Attachments.SuspendLayout()
        Me.RadPageViewPage4.SuspendLayout()
        CType(Me.btnAdminCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotalSubsidyAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRoundoffAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCurrConv.SuspendLayout()
        CType(Me.txtConversionRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblEffectiveFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtApplicableFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCurrency, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblConvRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAddCharges1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotRAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTaxAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAmtAfterDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDiscountAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAmtWithDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel19, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage5.SuspendLayout()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox3.SuspendLayout()
        CType(Me.txtAckDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EinvoiceBtnUpdate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EInvoiceQrCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EinvoiceAckNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.EInvoiceIRNNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.butCostCenterAndHirerachy_Update_AfterPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrintServiceInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnsetting, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnHistory)
        Me.SplitContainer1.Panel2.Controls.Add(Me.butCostCenterAndHirerachy_Update_AfterPost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrintServiceInvoice)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnsetting)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrint)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Size = New System.Drawing.Size(1102, 455)
        Me.SplitContainer1.SplitterDistance = 416
        Me.SplitContainer1.TabIndex = 1
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage1)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage2)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage3)
        Me.RadPageView1.Controls.Add(Me.pvpCustomFields)
        Me.RadPageView1.Controls.Add(Me.Attachments)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage4)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage5)
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.RadPageViewPage1
        Me.RadPageView1.Size = New System.Drawing.Size(1102, 416)
        Me.RadPageView1.TabIndex = 0
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).ItemAlignment = Telerik.WinControls.UI.StripViewItemAlignment.Near
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripAlignment = Telerik.WinControls.UI.StripViewAlignment.Top
        '
        'RadPageViewPage1
        '
        Me.RadPageViewPage1.Controls.Add(Me.chkEinvoice)
        Me.RadPageViewPage1.Controls.Add(Me.pnlLocation)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel10)
        Me.RadPageViewPage1.Controls.Add(Me.txtDataAndTimeSelection)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel9)
        Me.RadPageViewPage1.Controls.Add(Me.txtTapalNo)
        Me.RadPageViewPage1.Controls.Add(Me.chkAgainstServiceInvoice)
        Me.RadPageViewPage1.Controls.Add(Me.Panel2)
        Me.RadPageViewPage1.Controls.Add(Me.chkAgainstSecurityReceiptNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtServiceVisitCode)
        Me.RadPageViewPage1.Controls.Add(Me.txtVCGNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtSaleReturnNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtSaleInvoice)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel7)
        Me.RadPageViewPage1.Controls.Add(Me.ddlSecDepositType)
        Me.RadPageViewPage1.Controls.Add(Me.chkSecurityDposit)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel3)
        Me.RadPageViewPage1.Controls.Add(Me.LblLocDesp)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel6)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel5)
        Me.RadPageViewPage1.Controls.Add(Me.lblReturnType)
        Me.RadPageViewPage1.Controls.Add(Me.ddlReturnType)
        Me.RadPageViewPage1.Controls.Add(Me.pnlPCJ)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.txtlocation)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel3)
        Me.RadPageViewPage1.Controls.Add(Me.btnDrillDown)
        Me.RadPageViewPage1.Controls.Add(Me.txtACSet)
        Me.RadPageViewPage1.Controls.Add(Me.TxtCustomer)
        Me.RadPageViewPage1.Controls.Add(Me.txtRefDocNo)
        Me.RadPageViewPage1.Controls.Add(Me.UsLock1)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel15)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel12)
        Me.RadPageViewPage1.Controls.Add(Me.lblTotRAmt1)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox2)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel8)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel6)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel5)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel4)
        Me.RadPageViewPage1.Controls.Add(Me.chkOnHold)
        Me.RadPageViewPage1.Controls.Add(Me.lblVendorName)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel2)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.btnAddNew)
        Me.RadPageViewPage1.Controls.Add(Me.txtDocNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtOrderNo)
        Me.RadPageViewPage1.Controls.Add(Me.cboDocType)
        Me.RadPageViewPage1.Controls.Add(Me.txtDate)
        Me.RadPageViewPage1.Controls.Add(Me.txtDesc)
        Me.RadPageViewPage1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage1.ItemSize = New System.Drawing.SizeF(68.0!, 26.0!)
        Me.RadPageViewPage1.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage1.Name = "RadPageViewPage1"
        Me.RadPageViewPage1.Size = New System.Drawing.Size(1081, 370)
        Me.RadPageViewPage1.Text = "Document"
        '
        'chkEinvoice
        '
        Me.chkEinvoice.Location = New System.Drawing.Point(759, 129)
        Me.chkEinvoice.Name = "chkEinvoice"
        Me.chkEinvoice.Size = New System.Drawing.Size(101, 18)
        Me.chkEinvoice.TabIndex = 12148
        Me.chkEinvoice.Text = "Create E-Invoice"
        '
        'pnlLocation
        '
        Me.pnlLocation.Controls.Add(Me.RadLabel7)
        Me.pnlLocation.Controls.Add(Me.txtLocationPrefixName)
        Me.pnlLocation.Controls.Add(Me.txtLocationPrefix)
        Me.pnlLocation.Location = New System.Drawing.Point(543, 154)
        Me.pnlLocation.Name = "pnlLocation"
        Me.pnlLocation.Size = New System.Drawing.Size(410, 23)
        Me.pnlLocation.TabIndex = 12147
        '
        'RadLabel7
        '
        Me.RadLabel7.Location = New System.Drawing.Point(3, 3)
        Me.RadLabel7.Name = "RadLabel7"
        Me.RadLabel7.Size = New System.Drawing.Size(49, 18)
        Me.RadLabel7.TabIndex = 12145
        Me.RadLabel7.Text = "Location"
        '
        'txtLocationPrefixName
        '
        Me.txtLocationPrefixName.AutoSize = False
        Me.txtLocationPrefixName.BorderVisible = True
        Me.txtLocationPrefixName.FieldName = Nothing
        Me.txtLocationPrefixName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocationPrefixName.Location = New System.Drawing.Point(214, 3)
        Me.txtLocationPrefixName.Name = "txtLocationPrefixName"
        Me.txtLocationPrefixName.Size = New System.Drawing.Size(194, 18)
        Me.txtLocationPrefixName.TabIndex = 12146
        '
        'txtLocationPrefix
        '
        Me.txtLocationPrefix.CalculationExpression = Nothing
        Me.txtLocationPrefix.FieldCode = Nothing
        Me.txtLocationPrefix.FieldDesc = Nothing
        Me.txtLocationPrefix.FieldMaxLength = 0
        Me.txtLocationPrefix.FieldName = Nothing
        Me.txtLocationPrefix.isCalculatedField = False
        Me.txtLocationPrefix.IsSourceFromTable = False
        Me.txtLocationPrefix.IsSourceFromValueList = False
        Me.txtLocationPrefix.IsUnique = False
        Me.txtLocationPrefix.Location = New System.Drawing.Point(89, 2)
        Me.txtLocationPrefix.MendatroryField = True
        Me.txtLocationPrefix.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocationPrefix.MyLinkLable1 = Nothing
        Me.txtLocationPrefix.MyLinkLable2 = Nothing
        Me.txtLocationPrefix.MyReadOnly = False
        Me.txtLocationPrefix.MyShowMasterFormButton = False
        Me.txtLocationPrefix.Name = "txtLocationPrefix"
        Me.txtLocationPrefix.ReferenceFieldDesc = Nothing
        Me.txtLocationPrefix.ReferenceFieldName = Nothing
        Me.txtLocationPrefix.ReferenceTableName = Nothing
        Me.txtLocationPrefix.Size = New System.Drawing.Size(124, 18)
        Me.txtLocationPrefix.TabIndex = 12144
        Me.txtLocationPrefix.Value = ""
        '
        'MyLabel10
        '
        Me.MyLabel10.FieldName = Nothing
        Me.MyLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.Location = New System.Drawing.Point(311, 157)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(82, 16)
        Me.MyLabel10.TabIndex = 176
        Me.MyLabel10.Text = "Date And Time"
        '
        'txtDataAndTimeSelection
        '
        Me.txtDataAndTimeSelection.CalculationExpression = Nothing
        Me.txtDataAndTimeSelection.CustomFormat = "dd/MM/yyyy hh:mm tt"
        Me.txtDataAndTimeSelection.FieldCode = Nothing
        Me.txtDataAndTimeSelection.FieldDesc = Nothing
        Me.txtDataAndTimeSelection.FieldMaxLength = 0
        Me.txtDataAndTimeSelection.FieldName = Nothing
        Me.txtDataAndTimeSelection.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDataAndTimeSelection.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDataAndTimeSelection.isCalculatedField = False
        Me.txtDataAndTimeSelection.IsSourceFromTable = False
        Me.txtDataAndTimeSelection.IsSourceFromValueList = False
        Me.txtDataAndTimeSelection.IsUnique = False
        Me.txtDataAndTimeSelection.Location = New System.Drawing.Point(393, 156)
        Me.txtDataAndTimeSelection.MendatroryField = False
        Me.txtDataAndTimeSelection.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDataAndTimeSelection.MyLinkLable1 = Nothing
        Me.txtDataAndTimeSelection.MyLinkLable2 = Nothing
        Me.txtDataAndTimeSelection.Name = "txtDataAndTimeSelection"
        Me.txtDataAndTimeSelection.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDataAndTimeSelection.ReferenceFieldDesc = Nothing
        Me.txtDataAndTimeSelection.ReferenceFieldName = Nothing
        Me.txtDataAndTimeSelection.ReferenceTableName = Nothing
        Me.txtDataAndTimeSelection.ShowCheckBox = True
        Me.txtDataAndTimeSelection.Size = New System.Drawing.Size(143, 18)
        Me.txtDataAndTimeSelection.TabIndex = 175
        Me.txtDataAndTimeSelection.TabStop = False
        Me.txtDataAndTimeSelection.Text = "13/06/2011 11:29 AM"
        Me.txtDataAndTimeSelection.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'MyLabel9
        '
        Me.MyLabel9.FieldName = Nothing
        Me.MyLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel9.Location = New System.Drawing.Point(3, 157)
        Me.MyLabel9.Name = "MyLabel9"
        Me.MyLabel9.Size = New System.Drawing.Size(52, 16)
        Me.MyLabel9.TabIndex = 174
        Me.MyLabel9.Text = "Tapal No"
        '
        'txtTapalNo
        '
        Me.txtTapalNo.CalculationExpression = Nothing
        Me.txtTapalNo.FieldCode = Nothing
        Me.txtTapalNo.FieldDesc = Nothing
        Me.txtTapalNo.FieldMaxLength = 0
        Me.txtTapalNo.FieldName = Nothing
        Me.txtTapalNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTapalNo.isCalculatedField = False
        Me.txtTapalNo.IsSourceFromTable = False
        Me.txtTapalNo.IsSourceFromValueList = False
        Me.txtTapalNo.IsUnique = False
        Me.txtTapalNo.Location = New System.Drawing.Point(115, 156)
        Me.txtTapalNo.MendatroryField = False
        Me.txtTapalNo.MyLinkLable1 = Nothing
        Me.txtTapalNo.MyLinkLable2 = Nothing
        Me.txtTapalNo.Name = "txtTapalNo"
        Me.txtTapalNo.ReferenceFieldDesc = Nothing
        Me.txtTapalNo.ReferenceFieldName = Nothing
        Me.txtTapalNo.ReferenceTableName = Nothing
        Me.txtTapalNo.Size = New System.Drawing.Size(191, 18)
        Me.txtTapalNo.TabIndex = 173
        '
        'chkAgainstServiceInvoice
        '
        Me.chkAgainstServiceInvoice.Location = New System.Drawing.Point(869, 108)
        Me.chkAgainstServiceInvoice.Name = "chkAgainstServiceInvoice"
        Me.chkAgainstServiceInvoice.Size = New System.Drawing.Size(135, 18)
        Me.chkAgainstServiceInvoice.TabIndex = 172
        Me.chkAgainstServiceInvoice.Text = "Against Service Invoice"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.txtSercurityReceiptNo)
        Me.Panel2.Controls.Add(Me.MyLabel8)
        Me.Panel2.Location = New System.Drawing.Point(865, 126)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(317, 24)
        Me.Panel2.TabIndex = 171
        Me.Panel2.Visible = False
        '
        'txtSercurityReceiptNo
        '
        Me.txtSercurityReceiptNo.CalculationExpression = Nothing
        Me.txtSercurityReceiptNo.FieldCode = Nothing
        Me.txtSercurityReceiptNo.FieldDesc = Nothing
        Me.txtSercurityReceiptNo.FieldMaxLength = 0
        Me.txtSercurityReceiptNo.FieldName = Nothing
        Me.txtSercurityReceiptNo.isCalculatedField = False
        Me.txtSercurityReceiptNo.IsSourceFromTable = False
        Me.txtSercurityReceiptNo.IsSourceFromValueList = False
        Me.txtSercurityReceiptNo.IsUnique = False
        Me.txtSercurityReceiptNo.Location = New System.Drawing.Point(110, 2)
        Me.txtSercurityReceiptNo.MendatroryField = True
        Me.txtSercurityReceiptNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSercurityReceiptNo.MyLinkLable1 = Nothing
        Me.txtSercurityReceiptNo.MyLinkLable2 = Nothing
        Me.txtSercurityReceiptNo.MyReadOnly = False
        Me.txtSercurityReceiptNo.MyShowMasterFormButton = False
        Me.txtSercurityReceiptNo.Name = "txtSercurityReceiptNo"
        Me.txtSercurityReceiptNo.ReferenceFieldDesc = Nothing
        Me.txtSercurityReceiptNo.ReferenceFieldName = Nothing
        Me.txtSercurityReceiptNo.ReferenceTableName = Nothing
        Me.txtSercurityReceiptNo.Size = New System.Drawing.Size(204, 19)
        Me.txtSercurityReceiptNo.TabIndex = 11
        Me.txtSercurityReceiptNo.Value = ""
        '
        'MyLabel8
        '
        Me.MyLabel8.FieldName = Nothing
        Me.MyLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel8.Location = New System.Drawing.Point(3, 3)
        Me.MyLabel8.Name = "MyLabel8"
        Me.MyLabel8.Size = New System.Drawing.Size(106, 16)
        Me.MyLabel8.TabIndex = 12
        Me.MyLabel8.Text = "Security Receipt No"
        '
        'chkAgainstSecurityReceiptNo
        '
        Me.chkAgainstSecurityReceiptNo.Location = New System.Drawing.Point(759, 109)
        Me.chkAgainstSecurityReceiptNo.Name = "chkAgainstSecurityReceiptNo"
        Me.chkAgainstSecurityReceiptNo.Size = New System.Drawing.Size(100, 18)
        Me.chkAgainstSecurityReceiptNo.TabIndex = 169
        Me.chkAgainstSecurityReceiptNo.Text = "Against Security"
        '
        'txtServiceVisitCode
        '
        Me.txtServiceVisitCode.CalculationExpression = Nothing
        Me.txtServiceVisitCode.FieldCode = Nothing
        Me.txtServiceVisitCode.FieldDesc = Nothing
        Me.txtServiceVisitCode.FieldMaxLength = 0
        Me.txtServiceVisitCode.FieldName = Nothing
        Me.txtServiceVisitCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtServiceVisitCode.isCalculatedField = False
        Me.txtServiceVisitCode.IsSourceFromTable = False
        Me.txtServiceVisitCode.IsSourceFromValueList = False
        Me.txtServiceVisitCode.IsUnique = False
        Me.txtServiceVisitCode.Location = New System.Drawing.Point(869, 66)
        Me.txtServiceVisitCode.MaxLength = 30
        Me.txtServiceVisitCode.MendatroryField = False
        Me.txtServiceVisitCode.MyLinkLable1 = Nothing
        Me.txtServiceVisitCode.MyLinkLable2 = Nothing
        Me.txtServiceVisitCode.Name = "txtServiceVisitCode"
        Me.txtServiceVisitCode.ReferenceFieldDesc = Nothing
        Me.txtServiceVisitCode.ReferenceFieldName = Nothing
        Me.txtServiceVisitCode.ReferenceTableName = Nothing
        Me.txtServiceVisitCode.Size = New System.Drawing.Size(203, 18)
        Me.txtServiceVisitCode.TabIndex = 168
        '
        'txtVCGNo
        '
        Me.txtVCGNo.CalculationExpression = Nothing
        Me.txtVCGNo.FieldCode = Nothing
        Me.txtVCGNo.FieldDesc = Nothing
        Me.txtVCGNo.FieldMaxLength = 0
        Me.txtVCGNo.FieldName = Nothing
        Me.txtVCGNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVCGNo.isCalculatedField = False
        Me.txtVCGNo.IsSourceFromTable = False
        Me.txtVCGNo.IsSourceFromValueList = False
        Me.txtVCGNo.IsUnique = False
        Me.txtVCGNo.Location = New System.Drawing.Point(869, 43)
        Me.txtVCGNo.MaxLength = 30
        Me.txtVCGNo.MendatroryField = False
        Me.txtVCGNo.MyLinkLable1 = Nothing
        Me.txtVCGNo.MyLinkLable2 = Nothing
        Me.txtVCGNo.Name = "txtVCGNo"
        Me.txtVCGNo.ReferenceFieldDesc = Nothing
        Me.txtVCGNo.ReferenceFieldName = Nothing
        Me.txtVCGNo.ReferenceTableName = Nothing
        Me.txtVCGNo.Size = New System.Drawing.Size(203, 18)
        Me.txtVCGNo.TabIndex = 168
        '
        'txtSaleReturnNo
        '
        Me.txtSaleReturnNo.CalculationExpression = Nothing
        Me.txtSaleReturnNo.FieldCode = Nothing
        Me.txtSaleReturnNo.FieldDesc = Nothing
        Me.txtSaleReturnNo.FieldMaxLength = 0
        Me.txtSaleReturnNo.FieldName = Nothing
        Me.txtSaleReturnNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSaleReturnNo.isCalculatedField = False
        Me.txtSaleReturnNo.IsSourceFromTable = False
        Me.txtSaleReturnNo.IsSourceFromValueList = False
        Me.txtSaleReturnNo.IsUnique = False
        Me.txtSaleReturnNo.Location = New System.Drawing.Point(869, 23)
        Me.txtSaleReturnNo.MaxLength = 30
        Me.txtSaleReturnNo.MendatroryField = False
        Me.txtSaleReturnNo.MyLinkLable1 = Nothing
        Me.txtSaleReturnNo.MyLinkLable2 = Nothing
        Me.txtSaleReturnNo.Name = "txtSaleReturnNo"
        Me.txtSaleReturnNo.ReferenceFieldDesc = Nothing
        Me.txtSaleReturnNo.ReferenceFieldName = Nothing
        Me.txtSaleReturnNo.ReferenceTableName = Nothing
        Me.txtSaleReturnNo.Size = New System.Drawing.Size(203, 18)
        Me.txtSaleReturnNo.TabIndex = 168
        '
        'txtSaleInvoice
        '
        Me.txtSaleInvoice.CalculationExpression = Nothing
        Me.txtSaleInvoice.FieldCode = Nothing
        Me.txtSaleInvoice.FieldDesc = Nothing
        Me.txtSaleInvoice.FieldMaxLength = 0
        Me.txtSaleInvoice.FieldName = Nothing
        Me.txtSaleInvoice.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSaleInvoice.isCalculatedField = False
        Me.txtSaleInvoice.IsSourceFromTable = False
        Me.txtSaleInvoice.IsSourceFromValueList = False
        Me.txtSaleInvoice.IsUnique = False
        Me.txtSaleInvoice.Location = New System.Drawing.Point(869, 1)
        Me.txtSaleInvoice.MaxLength = 30
        Me.txtSaleInvoice.MendatroryField = False
        Me.txtSaleInvoice.MyLinkLable1 = Nothing
        Me.txtSaleInvoice.MyLinkLable2 = Nothing
        Me.txtSaleInvoice.Name = "txtSaleInvoice"
        Me.txtSaleInvoice.ReferenceFieldDesc = Nothing
        Me.txtSaleInvoice.ReferenceFieldName = Nothing
        Me.txtSaleInvoice.ReferenceTableName = Nothing
        Me.txtSaleInvoice.Size = New System.Drawing.Size(203, 18)
        Me.txtSaleInvoice.TabIndex = 167
        '
        'MyLabel7
        '
        Me.MyLabel7.FieldName = Nothing
        Me.MyLabel7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel7.Location = New System.Drawing.Point(759, 68)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(86, 16)
        Me.MyLabel7.TabIndex = 166
        Me.MyLabel7.Text = "Service Visit No"
        '
        'ddlSecDepositType
        '
        Me.ddlSecDepositType.AutoCompleteDisplayMember = Nothing
        Me.ddlSecDepositType.AutoCompleteValueMember = Nothing
        Me.ddlSecDepositType.CalculationExpression = Nothing
        Me.ddlSecDepositType.DropDownAnimationEnabled = True
        Me.ddlSecDepositType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.ddlSecDepositType.FieldCode = Nothing
        Me.ddlSecDepositType.FieldDesc = Nothing
        Me.ddlSecDepositType.FieldMaxLength = 0
        Me.ddlSecDepositType.FieldName = Nothing
        Me.ddlSecDepositType.isCalculatedField = False
        Me.ddlSecDepositType.IsSourceFromTable = False
        Me.ddlSecDepositType.IsSourceFromValueList = False
        Me.ddlSecDepositType.IsUnique = False
        Me.ddlSecDepositType.Location = New System.Drawing.Point(869, 87)
        Me.ddlSecDepositType.MendatroryField = False
        Me.ddlSecDepositType.MyLinkLable1 = Nothing
        Me.ddlSecDepositType.MyLinkLable2 = Nothing
        Me.ddlSecDepositType.Name = "ddlSecDepositType"
        Me.ddlSecDepositType.ReferenceFieldDesc = Nothing
        Me.ddlSecDepositType.ReferenceFieldName = Nothing
        Me.ddlSecDepositType.ReferenceTableName = Nothing
        Me.ddlSecDepositType.Size = New System.Drawing.Size(98, 20)
        Me.ddlSecDepositType.TabIndex = 164
        '
        'chkSecurityDposit
        '
        Me.chkSecurityDposit.Location = New System.Drawing.Point(759, 88)
        Me.chkSecurityDposit.Name = "chkSecurityDposit"
        Me.chkSecurityDposit.Size = New System.Drawing.Size(101, 18)
        Me.chkSecurityDposit.TabIndex = 163
        Me.chkSecurityDposit.Text = "Security Deposit"
        '
        'MyLabel3
        '
        Me.MyLabel3.FieldName = Nothing
        Me.MyLabel3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel3.Location = New System.Drawing.Point(759, 46)
        Me.MyLabel3.Name = "MyLabel3"
        Me.MyLabel3.Size = New System.Drawing.Size(55, 16)
        Me.MyLabel3.TabIndex = 162
        Me.MyLabel3.Text = "VCGL No"
        '
        'LblLocDesp
        '
        Me.LblLocDesp.AutoSize = False
        Me.LblLocDesp.BorderVisible = True
        Me.LblLocDesp.FieldName = Nothing
        Me.LblLocDesp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblLocDesp.Location = New System.Drawing.Point(264, 25)
        Me.LblLocDesp.Name = "LblLocDesp"
        Me.LblLocDesp.Size = New System.Drawing.Size(492, 18)
        Me.LblLocDesp.TabIndex = 160
        Me.LblLocDesp.TextWrap = False
        '
        'MyLabel6
        '
        Me.MyLabel6.FieldName = Nothing
        Me.MyLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel6.Location = New System.Drawing.Point(759, 23)
        Me.MyLabel6.Name = "MyLabel6"
        Me.MyLabel6.Size = New System.Drawing.Size(84, 16)
        Me.MyLabel6.TabIndex = 159
        Me.MyLabel6.Text = "Sale Return No"
        '
        'MyLabel5
        '
        Me.MyLabel5.FieldName = Nothing
        Me.MyLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(759, 1)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(86, 16)
        Me.MyLabel5.TabIndex = 158
        Me.MyLabel5.Text = "Sale Invoice No"
        '
        'lblReturnType
        '
        Me.lblReturnType.Enabled = False
        Me.lblReturnType.FieldName = Nothing
        Me.lblReturnType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReturnType.Location = New System.Drawing.Point(533, 133)
        Me.lblReturnType.Name = "lblReturnType"
        Me.lblReturnType.Size = New System.Drawing.Size(92, 16)
        Me.lblReturnType.TabIndex = 156
        Me.lblReturnType.Text = "Credit Note Type"
        Me.lblReturnType.Visible = False
        '
        'ddlReturnType
        '
        Me.ddlReturnType.AutoCompleteDisplayMember = Nothing
        Me.ddlReturnType.AutoCompleteValueMember = Nothing
        Me.ddlReturnType.CalculationExpression = Nothing
        Me.ddlReturnType.DropDownAnimationEnabled = True
        Me.ddlReturnType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.ddlReturnType.Enabled = False
        Me.ddlReturnType.FieldCode = Nothing
        Me.ddlReturnType.FieldDesc = Nothing
        Me.ddlReturnType.FieldMaxLength = 0
        Me.ddlReturnType.FieldName = Nothing
        Me.ddlReturnType.isCalculatedField = False
        Me.ddlReturnType.IsSourceFromTable = False
        Me.ddlReturnType.IsSourceFromValueList = False
        Me.ddlReturnType.IsUnique = False
        Me.ddlReturnType.Location = New System.Drawing.Point(633, 131)
        Me.ddlReturnType.MendatroryField = True
        Me.ddlReturnType.MyLinkLable1 = Me.lblReturnType
        Me.ddlReturnType.MyLinkLable2 = Nothing
        Me.ddlReturnType.Name = "ddlReturnType"
        Me.ddlReturnType.ReferenceFieldDesc = Nothing
        Me.ddlReturnType.ReferenceFieldName = Nothing
        Me.ddlReturnType.ReferenceTableName = Nothing
        Me.ddlReturnType.Size = New System.Drawing.Size(121, 20)
        Me.ddlReturnType.TabIndex = 155
        Me.ddlReturnType.Visible = False
        '
        'pnlPCJ
        '
        Me.pnlPCJ.Controls.Add(Me.MyLabel4)
        Me.pnlPCJ.Controls.Add(Me.fndProject)
        Me.pnlPCJ.Controls.Add(Me.lblProject)
        Me.pnlPCJ.Location = New System.Drawing.Point(3, 129)
        Me.pnlPCJ.Name = "pnlPCJ"
        Me.pnlPCJ.Size = New System.Drawing.Size(536, 23)
        Me.pnlPCJ.TabIndex = 11
        '
        'MyLabel4
        '
        Me.MyLabel4.FieldName = Nothing
        Me.MyLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(2, 1)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(41, 16)
        Me.MyLabel4.TabIndex = 34
        Me.MyLabel4.Text = "Project"
        '
        'fndProject
        '
        Me.fndProject.CalculationExpression = Nothing
        Me.fndProject.FieldCode = Nothing
        Me.fndProject.FieldDesc = Nothing
        Me.fndProject.FieldMaxLength = 0
        Me.fndProject.FieldName = Nothing
        Me.fndProject.isCalculatedField = False
        Me.fndProject.IsSourceFromTable = False
        Me.fndProject.IsSourceFromValueList = False
        Me.fndProject.IsUnique = False
        Me.fndProject.Location = New System.Drawing.Point(113, 3)
        Me.fndProject.MendatroryField = False
        Me.fndProject.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fndProject.MyLinkLable1 = Me.MyLabel4
        Me.fndProject.MyLinkLable2 = Me.lblProject
        Me.fndProject.MyReadOnly = False
        Me.fndProject.MyShowMasterFormButton = False
        Me.fndProject.Name = "fndProject"
        Me.fndProject.ReferenceFieldDesc = Nothing
        Me.fndProject.ReferenceFieldName = Nothing
        Me.fndProject.ReferenceTableName = Nothing
        Me.fndProject.Size = New System.Drawing.Size(143, 17)
        Me.fndProject.TabIndex = 0
        Me.fndProject.Value = ""
        '
        'lblProject
        '
        Me.lblProject.AutoSize = False
        Me.lblProject.BorderVisible = True
        Me.lblProject.FieldName = Nothing
        Me.lblProject.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProject.Location = New System.Drawing.Point(262, 3)
        Me.lblProject.Name = "lblProject"
        Me.lblProject.Size = New System.Drawing.Size(264, 18)
        Me.lblProject.TabIndex = 1
        Me.lblProject.TextWrap = False
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.Location = New System.Drawing.Point(3, 47)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(32, 16)
        Me.MyLabel1.TabIndex = 11
        Me.MyLabel1.Text = "Desc"
        '
        'txtlocation
        '
        Me.txtlocation.CalculationExpression = Nothing
        Me.txtlocation.FieldCode = Nothing
        Me.txtlocation.FieldDesc = Nothing
        Me.txtlocation.FieldMaxLength = 0
        Me.txtlocation.FieldName = Nothing
        Me.txtlocation.isCalculatedField = False
        Me.txtlocation.IsSourceFromTable = False
        Me.txtlocation.IsSourceFromValueList = False
        Me.txtlocation.IsUnique = False
        Me.txtlocation.Location = New System.Drawing.Point(117, 25)
        Me.txtlocation.MendatroryField = True
        Me.txtlocation.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtlocation.MyLinkLable1 = Nothing
        Me.txtlocation.MyLinkLable2 = Nothing
        Me.txtlocation.MyReadOnly = False
        Me.txtlocation.MyShowMasterFormButton = False
        Me.txtlocation.Name = "txtlocation"
        Me.txtlocation.ReferenceFieldDesc = Nothing
        Me.txtlocation.ReferenceFieldName = Nothing
        Me.txtlocation.ReferenceTableName = Nothing
        Me.txtlocation.Size = New System.Drawing.Size(143, 18)
        Me.txtlocation.TabIndex = 2
        Me.txtlocation.Value = ""
        '
        'RadLabel3
        '
        Me.RadLabel3.Location = New System.Drawing.Point(3, 25)
        Me.RadLabel3.Name = "RadLabel3"
        Me.RadLabel3.Size = New System.Drawing.Size(96, 18)
        Me.RadLabel3.TabIndex = 24
        Me.RadLabel3.Text = "Location Segment"
        '
        'btnDrillDown
        '
        Me.btnDrillDown.Image = CType(resources.GetObject("btnDrillDown.Image"), System.Drawing.Image)
        Me.btnDrillDown.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnDrillDown.Location = New System.Drawing.Point(403, 1)
        Me.btnDrillDown.Name = "btnDrillDown"
        Me.btnDrillDown.Size = New System.Drawing.Size(20, 20)
        Me.btnDrillDown.TabIndex = 2
        '
        'txtACSet
        '
        Me.txtACSet.CalculationExpression = Nothing
        Me.txtACSet.FieldCode = Nothing
        Me.txtACSet.FieldDesc = Nothing
        Me.txtACSet.FieldMaxLength = 0
        Me.txtACSet.FieldName = Nothing
        Me.txtACSet.isCalculatedField = False
        Me.txtACSet.IsSourceFromTable = False
        Me.txtACSet.IsSourceFromValueList = False
        Me.txtACSet.IsUnique = False
        Me.txtACSet.Location = New System.Drawing.Point(633, 88)
        Me.txtACSet.MendatroryField = False
        Me.txtACSet.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtACSet.MyLinkLable1 = Nothing
        Me.txtACSet.MyLinkLable2 = Nothing
        Me.txtACSet.MyReadOnly = False
        Me.txtACSet.MyShowMasterFormButton = False
        Me.txtACSet.Name = "txtACSet"
        Me.txtACSet.ReferenceFieldDesc = Nothing
        Me.txtACSet.ReferenceFieldName = Nothing
        Me.txtACSet.ReferenceTableName = Nothing
        Me.txtACSet.Size = New System.Drawing.Size(121, 18)
        Me.txtACSet.TabIndex = 8
        Me.txtACSet.Value = ""
        '
        'TxtCustomer
        '
        Me.TxtCustomer.CalculationExpression = Nothing
        Me.TxtCustomer.FieldCode = Nothing
        Me.TxtCustomer.FieldDesc = Nothing
        Me.TxtCustomer.FieldMaxLength = 0
        Me.TxtCustomer.FieldName = Nothing
        Me.TxtCustomer.isCalculatedField = False
        Me.TxtCustomer.IsSourceFromTable = False
        Me.TxtCustomer.IsSourceFromValueList = False
        Me.TxtCustomer.IsUnique = False
        Me.TxtCustomer.Location = New System.Drawing.Point(117, 67)
        Me.TxtCustomer.MendatroryField = True
        Me.TxtCustomer.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCustomer.MyLinkLable1 = Nothing
        Me.TxtCustomer.MyLinkLable2 = Nothing
        Me.TxtCustomer.MyReadOnly = False
        Me.TxtCustomer.MyShowMasterFormButton = False
        Me.TxtCustomer.Name = "TxtCustomer"
        Me.TxtCustomer.ReferenceFieldDesc = Nothing
        Me.TxtCustomer.ReferenceFieldName = Nothing
        Me.TxtCustomer.ReferenceTableName = Nothing
        Me.TxtCustomer.Size = New System.Drawing.Size(143, 18)
        Me.TxtCustomer.TabIndex = 5
        Me.TxtCustomer.Value = ""
        '
        'txtRefDocNo
        '
        Me.txtRefDocNo.CalculationExpression = Nothing
        Me.txtRefDocNo.FieldCode = Nothing
        Me.txtRefDocNo.FieldDesc = Nothing
        Me.txtRefDocNo.FieldMaxLength = 0
        Me.txtRefDocNo.FieldName = Nothing
        Me.txtRefDocNo.isCalculatedField = False
        Me.txtRefDocNo.IsSourceFromTable = False
        Me.txtRefDocNo.IsSourceFromValueList = False
        Me.txtRefDocNo.IsUnique = False
        Me.txtRefDocNo.Location = New System.Drawing.Point(117, 109)
        Me.txtRefDocNo.MendatroryField = False
        Me.txtRefDocNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRefDocNo.MyLinkLable1 = Nothing
        Me.txtRefDocNo.MyLinkLable2 = Nothing
        Me.txtRefDocNo.MyReadOnly = False
        Me.txtRefDocNo.MyShowMasterFormButton = False
        Me.txtRefDocNo.Name = "txtRefDocNo"
        Me.txtRefDocNo.ReferenceFieldDesc = Nothing
        Me.txtRefDocNo.ReferenceFieldName = Nothing
        Me.txtRefDocNo.ReferenceTableName = Nothing
        Me.txtRefDocNo.Size = New System.Drawing.Size(143, 18)
        Me.txtRefDocNo.TabIndex = 9
        Me.txtRefDocNo.Value = ""
        '
        'UsLock1
        '
        Me.UsLock1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.UsLock1.Location = New System.Drawing.Point(635, -1)
        Me.UsLock1.MyFont = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsLock1.Name = "UsLock1"
        Me.UsLock1.Size = New System.Drawing.Size(118, 20)
        Me.UsLock1.Status = common.ERPTransactionStatus.Pending
        Me.UsLock1.TabIndex = 5
        '
        'RadLabel15
        '
        Me.RadLabel15.FieldName = Nothing
        Me.RadLabel15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel15.Location = New System.Drawing.Point(3, 110)
        Me.RadLabel15.Name = "RadLabel15"
        Me.RadLabel15.Size = New System.Drawing.Size(60, 16)
        Me.RadLabel15.TabIndex = 8
        Me.RadLabel15.Text = "Invoice No"
        '
        'RadLabel12
        '
        Me.RadLabel12.FieldName = Nothing
        Me.RadLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel12.Location = New System.Drawing.Point(264, 110)
        Me.RadLabel12.Name = "RadLabel12"
        Me.RadLabel12.Size = New System.Drawing.Size(100, 16)
        Me.RadLabel12.TabIndex = 17
        Me.RadLabel12.Text = "Document Amount"
        '
        'lblTotRAmt1
        '
        Me.lblTotRAmt1.AutoSize = False
        Me.lblTotRAmt1.BackColor = System.Drawing.Color.Transparent
        Me.lblTotRAmt1.BorderVisible = True
        Me.lblTotRAmt1.FieldName = Nothing
        Me.lblTotRAmt1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRAmt1.Location = New System.Drawing.Point(368, 109)
        Me.lblTotRAmt1.Name = "lblTotRAmt1"
        Me.lblTotRAmt1.Size = New System.Drawing.Size(160, 18)
        Me.lblTotRAmt1.TabIndex = 14
        Me.lblTotRAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadGroupBox2
        '
        Me.RadGroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadGroupBox2.Controls.Add(Me.gv1)
        Me.RadGroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadGroupBox2.HeaderText = "Details"
        Me.RadGroupBox2.Location = New System.Drawing.Point(1, 183)
        Me.RadGroupBox2.Name = "RadGroupBox2"
        Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox2.Size = New System.Drawing.Size(1079, 186)
        Me.RadGroupBox2.TabIndex = 12
        Me.RadGroupBox2.Text = "Details"
        '
        'gv1
        '
        Me.gv1.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.gv1.Cursor = System.Windows.Forms.Cursors.Default
        Me.gv1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gv1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gv1.ForeColor = System.Drawing.Color.Black
        Me.gv1.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.gv1.Location = New System.Drawing.Point(10, 20)
        '
        '
        '
        Me.gv1.MasterTemplate.AllowDeleteRow = False
        Me.gv1.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gv1.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv1.MasterTemplate.ViewDefinition = TableViewDefinition1
        Me.gv1.MyExportAPI = False
        Me.gv1.MyExportFilePath = ""
        Me.gv1.MyStopExport = False
        Me.gv1.Name = "gv1"
        Me.gv1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gv1.ShowGroupPanel = False
        Me.gv1.ShowHeaderCellButtons = True
        Me.gv1.Size = New System.Drawing.Size(1059, 156)
        Me.gv1.TabIndex = 0
        Me.gv1.TabStop = False
        Me.gv1.VarID = ""
        '
        'RadLabel8
        '
        Me.RadLabel8.FieldName = Nothing
        Me.RadLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel8.Location = New System.Drawing.Point(533, 110)
        Me.RadLabel8.Name = "RadLabel8"
        Me.RadLabel8.Size = New System.Drawing.Size(53, 16)
        Me.RadLabel8.TabIndex = 19
        Me.RadLabel8.Text = "Order No"
        '
        'RadLabel6
        '
        Me.RadLabel6.FieldName = Nothing
        Me.RadLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel6.Location = New System.Drawing.Point(533, 89)
        Me.RadLabel6.Name = "RadLabel6"
        Me.RadLabel6.Size = New System.Drawing.Size(67, 16)
        Me.RadLabel6.TabIndex = 20
        Me.RadLabel6.Text = "Account Set"
        '
        'RadLabel5
        '
        Me.RadLabel5.FieldName = Nothing
        Me.RadLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel5.Location = New System.Drawing.Point(264, 89)
        Me.RadLabel5.Name = "RadLabel5"
        Me.RadLabel5.Size = New System.Drawing.Size(86, 16)
        Me.RadLabel5.TabIndex = 10
        Me.RadLabel5.Text = "Document Type"
        '
        'RadLabel4
        '
        Me.RadLabel4.FieldName = Nothing
        Me.RadLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel4.Location = New System.Drawing.Point(3, 89)
        Me.RadLabel4.Name = "RadLabel4"
        Me.RadLabel4.Size = New System.Drawing.Size(85, 16)
        Me.RadLabel4.TabIndex = 9
        Me.RadLabel4.Text = "Document Date"
        '
        'chkOnHold
        '
        Me.chkOnHold.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnHold.Location = New System.Drawing.Point(426, 3)
        Me.chkOnHold.Name = "chkOnHold"
        Me.chkOnHold.Size = New System.Drawing.Size(62, 16)
        Me.chkOnHold.TabIndex = 3
        Me.chkOnHold.Text = "On Hold"
        '
        'lblVendorName
        '
        Me.lblVendorName.AutoSize = False
        Me.lblVendorName.BorderVisible = True
        Me.lblVendorName.FieldName = Nothing
        Me.lblVendorName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorName.Location = New System.Drawing.Point(264, 67)
        Me.lblVendorName.Name = "lblVendorName"
        Me.lblVendorName.Size = New System.Drawing.Size(492, 18)
        Me.lblVendorName.TabIndex = 8
        '
        'RadLabel2
        '
        Me.RadLabel2.FieldName = Nothing
        Me.RadLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel2.Location = New System.Drawing.Point(3, 68)
        Me.RadLabel2.Name = "RadLabel2"
        Me.RadLabel2.Size = New System.Drawing.Size(55, 16)
        Me.RadLabel2.TabIndex = 10
        Me.RadLabel2.Text = "Customer"
        '
        'RadLabel1
        '
        Me.RadLabel1.FieldName = Nothing
        Me.RadLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel1.Location = New System.Drawing.Point(3, 3)
        Me.RadLabel1.Name = "RadLabel1"
        Me.RadLabel1.Size = New System.Drawing.Size(75, 16)
        Me.RadLabel1.TabIndex = 11
        Me.RadLabel1.Text = "Document No"
        '
        'btnAddNew
        '
        Me.btnAddNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Image = Global.ERP.My.Resources.Resources._new
        Me.btnAddNew.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnAddNew.Location = New System.Drawing.Point(383, 1)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(20, 20)
        Me.btnAddNew.TabIndex = 1
        '
        'txtDocNo
        '
        Me.txtDocNo.FieldName = Nothing
        Me.txtDocNo.Location = New System.Drawing.Point(117, 1)
        Me.txtDocNo.MendatroryField = False
        Me.txtDocNo.MyCharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.txtDocNo.MyFont = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.txtDocNo.MyLinkLable1 = Nothing
        Me.txtDocNo.MyLinkLable2 = Nothing
        Me.txtDocNo.MyMaxLength = 32767
        Me.txtDocNo.MyReadOnly = False
        Me.txtDocNo.Name = "txtDocNo"
        Me.txtDocNo.Size = New System.Drawing.Size(266, 20)
        Me.txtDocNo.TabIndex = 0
        Me.txtDocNo.TabStop = False
        Me.txtDocNo.Value = ""
        '
        'txtOrderNo
        '
        Me.txtOrderNo.CalculationExpression = Nothing
        Me.txtOrderNo.FieldCode = Nothing
        Me.txtOrderNo.FieldDesc = Nothing
        Me.txtOrderNo.FieldMaxLength = 0
        Me.txtOrderNo.FieldName = Nothing
        Me.txtOrderNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOrderNo.isCalculatedField = False
        Me.txtOrderNo.IsSourceFromTable = False
        Me.txtOrderNo.IsSourceFromValueList = False
        Me.txtOrderNo.IsUnique = False
        Me.txtOrderNo.Location = New System.Drawing.Point(633, 109)
        Me.txtOrderNo.MaxLength = 30
        Me.txtOrderNo.MendatroryField = False
        Me.txtOrderNo.MyLinkLable1 = Nothing
        Me.txtOrderNo.MyLinkLable2 = Nothing
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.ReferenceFieldDesc = Nothing
        Me.txtOrderNo.ReferenceFieldName = Nothing
        Me.txtOrderNo.ReferenceTableName = Nothing
        Me.txtOrderNo.Size = New System.Drawing.Size(121, 18)
        Me.txtOrderNo.TabIndex = 10
        '
        'cboDocType
        '
        Me.cboDocType.AllowDrop = True
        Me.cboDocType.AutoCompleteDisplayMember = Nothing
        Me.cboDocType.AutoCompleteValueMember = Nothing
        Me.cboDocType.CalculationExpression = Nothing
        Me.cboDocType.DropDownAnimationEnabled = True
        Me.cboDocType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cboDocType.FieldCode = Nothing
        Me.cboDocType.FieldDesc = Nothing
        Me.cboDocType.FieldMaxLength = 0
        Me.cboDocType.FieldName = Nothing
        Me.cboDocType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDocType.isCalculatedField = False
        Me.cboDocType.IsSourceFromTable = False
        Me.cboDocType.IsSourceFromValueList = False
        Me.cboDocType.IsUnique = False
        Me.cboDocType.Location = New System.Drawing.Point(368, 88)
        Me.cboDocType.MendatroryField = False
        Me.cboDocType.MyLinkLable1 = Nothing
        Me.cboDocType.MyLinkLable2 = Nothing
        Me.cboDocType.Name = "cboDocType"
        Me.cboDocType.ReferenceFieldDesc = Nothing
        Me.cboDocType.ReferenceFieldName = Nothing
        Me.cboDocType.ReferenceTableName = Nothing
        Me.cboDocType.Size = New System.Drawing.Size(160, 18)
        Me.cboDocType.TabIndex = 7
        Me.cboDocType.Text = "RadDropDownList1"
        '
        'txtDate
        '
        Me.txtDate.CalculationExpression = Nothing
        Me.txtDate.CustomFormat = "dd/MM/yyyy hh:mm tt"
        Me.txtDate.FieldCode = Nothing
        Me.txtDate.FieldDesc = Nothing
        Me.txtDate.FieldMaxLength = 0
        Me.txtDate.FieldName = Nothing
        Me.txtDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDate.isCalculatedField = False
        Me.txtDate.IsSourceFromTable = False
        Me.txtDate.IsSourceFromValueList = False
        Me.txtDate.IsUnique = False
        Me.txtDate.Location = New System.Drawing.Point(117, 88)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Nothing
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(143, 18)
        Me.txtDate.TabIndex = 6
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "13/06/2011 11:29 AM"
        Me.txtDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'txtDesc
        '
        Me.txtDesc.CalculationExpression = Nothing
        Me.txtDesc.FieldCode = Nothing
        Me.txtDesc.FieldDesc = Nothing
        Me.txtDesc.FieldMaxLength = 0
        Me.txtDesc.FieldName = Nothing
        Me.txtDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDesc.isCalculatedField = False
        Me.txtDesc.IsSourceFromTable = False
        Me.txtDesc.IsSourceFromValueList = False
        Me.txtDesc.IsUnique = False
        Me.txtDesc.Location = New System.Drawing.Point(117, 46)
        Me.txtDesc.MendatroryField = False
        Me.txtDesc.MyLinkLable1 = Nothing
        Me.txtDesc.MyLinkLable2 = Nothing
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReferenceFieldDesc = Nothing
        Me.txtDesc.ReferenceFieldName = Nothing
        Me.txtDesc.ReferenceTableName = Nothing
        Me.txtDesc.Size = New System.Drawing.Size(639, 18)
        Me.txtDesc.TabIndex = 4
        '
        'RadPageViewPage2
        '
        Me.RadPageViewPage2.Controls.Add(Me.GroupBox1)
        Me.RadPageViewPage2.Controls.Add(Me.txtTaxGroup)
        Me.RadPageViewPage2.Controls.Add(Me.RadLabel10)
        Me.RadPageViewPage2.Controls.Add(Me.RadGroupBox1)
        Me.RadPageViewPage2.Controls.Add(Me.gv2)
        Me.RadPageViewPage2.Controls.Add(Me.lblTaxGrpName)
        Me.RadPageViewPage2.Controls.Add(Me.RadLabel11)
        Me.RadPageViewPage2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage2.ItemSize = New System.Drawing.SizeF(47.0!, 26.0!)
        Me.RadPageViewPage2.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage2.Name = "RadPageViewPage2"
        Me.RadPageViewPage2.Size = New System.Drawing.Size(1081, 370)
        Me.RadPageViewPage2.Text = "Taxes"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbtnTaxCalManual)
        Me.GroupBox1.Controls.Add(Me.rbtnTaxCalAutomatic)
        Me.GroupBox1.Location = New System.Drawing.Point(553, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(163, 36)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tax Calculation Type"
        '
        'rbtnTaxCalManual
        '
        Me.rbtnTaxCalManual.Location = New System.Drawing.Point(89, 13)
        Me.rbtnTaxCalManual.MyLinkLable1 = Nothing
        Me.rbtnTaxCalManual.MyLinkLable2 = Nothing
        Me.rbtnTaxCalManual.Name = "rbtnTaxCalManual"
        Me.rbtnTaxCalManual.Size = New System.Drawing.Size(57, 18)
        Me.rbtnTaxCalManual.TabIndex = 1
        Me.rbtnTaxCalManual.Text = "Manual"
        '
        'rbtnTaxCalAutomatic
        '
        Me.rbtnTaxCalAutomatic.Location = New System.Drawing.Point(7, 13)
        Me.rbtnTaxCalAutomatic.MyLinkLable1 = Nothing
        Me.rbtnTaxCalAutomatic.MyLinkLable2 = Nothing
        Me.rbtnTaxCalAutomatic.Name = "rbtnTaxCalAutomatic"
        Me.rbtnTaxCalAutomatic.Size = New System.Drawing.Size(72, 18)
        Me.rbtnTaxCalAutomatic.TabIndex = 0
        Me.rbtnTaxCalAutomatic.Text = "Automatic"
        '
        'txtTaxGroup
        '
        Me.txtTaxGroup.CalculationExpression = Nothing
        Me.txtTaxGroup.FieldCode = Nothing
        Me.txtTaxGroup.FieldDesc = Nothing
        Me.txtTaxGroup.FieldMaxLength = 0
        Me.txtTaxGroup.FieldName = Nothing
        Me.txtTaxGroup.isCalculatedField = False
        Me.txtTaxGroup.IsSourceFromTable = False
        Me.txtTaxGroup.IsSourceFromValueList = False
        Me.txtTaxGroup.IsUnique = False
        Me.txtTaxGroup.Location = New System.Drawing.Point(69, 5)
        Me.txtTaxGroup.MendatroryField = True
        Me.txtTaxGroup.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTaxGroup.MyLinkLable1 = Nothing
        Me.txtTaxGroup.MyLinkLable2 = Nothing
        Me.txtTaxGroup.MyReadOnly = False
        Me.txtTaxGroup.MyShowMasterFormButton = False
        Me.txtTaxGroup.Name = "txtTaxGroup"
        Me.txtTaxGroup.ReferenceFieldDesc = Nothing
        Me.txtTaxGroup.ReferenceFieldName = Nothing
        Me.txtTaxGroup.ReferenceTableName = Nothing
        Me.txtTaxGroup.Size = New System.Drawing.Size(143, 19)
        Me.txtTaxGroup.TabIndex = 0
        Me.txtTaxGroup.Value = ""
        '
        'RadLabel10
        '
        Me.RadLabel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadLabel10.FieldName = Nothing
        Me.RadLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel10.ForeColor = System.Drawing.SystemColors.Highlight
        Me.RadLabel10.Location = New System.Drawing.Point(923, 266)
        Me.RadLabel10.Name = "RadLabel10"
        Me.RadLabel10.Size = New System.Drawing.Size(155, 16)
        Me.RadLabel10.TabIndex = 4
        Me.RadLabel10.Text = "Double click To Chage Rate"
        '
        'RadGroupBox1
        '
        Me.RadGroupBox1.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadGroupBox1.Controls.Add(Me.txtTermCode)
        Me.RadGroupBox1.Controls.Add(Me.RadLabel16)
        Me.RadGroupBox1.Controls.Add(Me.txtDueDate)
        Me.RadGroupBox1.Controls.Add(Me.RadLabel17)
        Me.RadGroupBox1.Controls.Add(Me.lblTermName)
        Me.RadGroupBox1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadGroupBox1.HeaderText = "Terms"
        Me.RadGroupBox1.Location = New System.Drawing.Point(1, 279)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox1.Size = New System.Drawing.Size(1081, 87)
        Me.RadGroupBox1.TabIndex = 3
        Me.RadGroupBox1.Text = "Terms"
        '
        'txtTermCode
        '
        Me.txtTermCode.CalculationExpression = Nothing
        Me.txtTermCode.FieldCode = Nothing
        Me.txtTermCode.FieldDesc = Nothing
        Me.txtTermCode.FieldMaxLength = 0
        Me.txtTermCode.FieldName = Nothing
        Me.txtTermCode.isCalculatedField = False
        Me.txtTermCode.IsSourceFromTable = False
        Me.txtTermCode.IsSourceFromValueList = False
        Me.txtTermCode.IsUnique = False
        Me.txtTermCode.Location = New System.Drawing.Point(70, 24)
        Me.txtTermCode.MendatroryField = False
        Me.txtTermCode.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTermCode.MyLinkLable1 = Nothing
        Me.txtTermCode.MyLinkLable2 = Nothing
        Me.txtTermCode.MyReadOnly = False
        Me.txtTermCode.MyShowMasterFormButton = False
        Me.txtTermCode.Name = "txtTermCode"
        Me.txtTermCode.ReferenceFieldDesc = Nothing
        Me.txtTermCode.ReferenceFieldName = Nothing
        Me.txtTermCode.ReferenceTableName = Nothing
        Me.txtTermCode.Size = New System.Drawing.Size(143, 18)
        Me.txtTermCode.TabIndex = 0
        Me.txtTermCode.Value = ""
        '
        'RadLabel16
        '
        Me.RadLabel16.FieldName = Nothing
        Me.RadLabel16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel16.Location = New System.Drawing.Point(6, 26)
        Me.RadLabel16.Name = "RadLabel16"
        Me.RadLabel16.Size = New System.Drawing.Size(63, 16)
        Me.RadLabel16.TabIndex = 13
        Me.RadLabel16.Text = "Term Code"
        '
        'txtDueDate
        '
        Me.txtDueDate.CalculationExpression = Nothing
        Me.txtDueDate.CustomFormat = "dd-MM-yyyy"
        Me.txtDueDate.FieldCode = Nothing
        Me.txtDueDate.FieldDesc = Nothing
        Me.txtDueDate.FieldMaxLength = 0
        Me.txtDueDate.FieldName = Nothing
        Me.txtDueDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtDueDate.isCalculatedField = False
        Me.txtDueDate.IsSourceFromTable = False
        Me.txtDueDate.IsSourceFromValueList = False
        Me.txtDueDate.IsUnique = False
        Me.txtDueDate.Location = New System.Drawing.Point(70, 57)
        Me.txtDueDate.MendatroryField = False
        Me.txtDueDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDueDate.MyLinkLable1 = Nothing
        Me.txtDueDate.MyLinkLable2 = Nothing
        Me.txtDueDate.Name = "txtDueDate"
        Me.txtDueDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDueDate.ReferenceFieldDesc = Nothing
        Me.txtDueDate.ReferenceFieldName = Nothing
        Me.txtDueDate.ReferenceTableName = Nothing
        Me.txtDueDate.Size = New System.Drawing.Size(116, 18)
        Me.txtDueDate.TabIndex = 2
        Me.txtDueDate.TabStop = False
        Me.txtDueDate.Text = "13-06-2011"
        Me.txtDueDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'RadLabel17
        '
        Me.RadLabel17.FieldName = Nothing
        Me.RadLabel17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel17.Location = New System.Drawing.Point(6, 58)
        Me.RadLabel17.Name = "RadLabel17"
        Me.RadLabel17.Size = New System.Drawing.Size(54, 16)
        Me.RadLabel17.TabIndex = 12
        Me.RadLabel17.Text = "Due Date"
        '
        'lblTermName
        '
        Me.lblTermName.AutoSize = False
        Me.lblTermName.BorderVisible = True
        Me.lblTermName.FieldName = Nothing
        Me.lblTermName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTermName.Location = New System.Drawing.Point(225, 24)
        Me.lblTermName.Name = "lblTermName"
        Me.lblTermName.Size = New System.Drawing.Size(321, 20)
        Me.lblTermName.TabIndex = 1
        '
        'gv2
        '
        Me.gv2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gv2.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.gv2.BeginEditMode = Telerik.WinControls.RadGridViewBeginEditMode.BeginEditOnEnter
        Me.gv2.Cursor = System.Windows.Forms.Cursors.Default
        Me.gv2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gv2.ForeColor = System.Drawing.Color.Black
        Me.gv2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.gv2.Location = New System.Drawing.Point(2, 39)
        '
        '
        '
        Me.gv2.MasterTemplate.AllowAddNewRow = False
        Me.gv2.MasterTemplate.AllowDeleteRow = False
        Me.gv2.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gv2.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv2.MasterTemplate.ViewDefinition = TableViewDefinition2
        Me.gv2.MyExportAPI = False
        Me.gv2.MyExportFilePath = ""
        Me.gv2.MyStopExport = False
        Me.gv2.Name = "gv2"
        Me.gv2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gv2.ShowHeaderCellButtons = True
        Me.gv2.Size = New System.Drawing.Size(1076, 224)
        Me.gv2.TabIndex = 1
        Me.gv2.TabStop = False
        Me.gv2.VarID = ""
        '
        'lblTaxGrpName
        '
        Me.lblTaxGrpName.AutoSize = False
        Me.lblTaxGrpName.BorderVisible = True
        Me.lblTaxGrpName.FieldName = Nothing
        Me.lblTaxGrpName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxGrpName.Location = New System.Drawing.Point(225, 5)
        Me.lblTaxGrpName.Name = "lblTaxGrpName"
        Me.lblTaxGrpName.Size = New System.Drawing.Size(321, 20)
        Me.lblTaxGrpName.TabIndex = 1
        '
        'RadLabel11
        '
        Me.RadLabel11.FieldName = Nothing
        Me.RadLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel11.Location = New System.Drawing.Point(3, 7)
        Me.RadLabel11.Name = "RadLabel11"
        Me.RadLabel11.Size = New System.Drawing.Size(60, 16)
        Me.RadLabel11.TabIndex = 2
        Me.RadLabel11.Text = "Tax Group"
        '
        'RadPageViewPage3
        '
        Me.RadPageViewPage3.Controls.Add(Me.SplitContainer2)
        Me.RadPageViewPage3.ItemSize = New System.Drawing.SizeF(116.0!, 26.0!)
        Me.RadPageViewPage3.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage3.Name = "RadPageViewPage3"
        Me.RadPageViewPage3.Size = New System.Drawing.Size(1081, 370)
        Me.RadPageViewPage3.Text = "Additional Chrarges"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.gvAC)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.RadLabel31)
        Me.SplitContainer2.Panel2.Controls.Add(Me.lblAddCharges)
        Me.SplitContainer2.Size = New System.Drawing.Size(1081, 370)
        Me.SplitContainer2.SplitterDistance = 333
        Me.SplitContainer2.TabIndex = 1
        '
        'gvAC
        '
        Me.gvAC.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.gvAC.Cursor = System.Windows.Forms.Cursors.Default
        Me.gvAC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gvAC.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gvAC.ForeColor = System.Drawing.Color.Black
        Me.gvAC.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.gvAC.Location = New System.Drawing.Point(0, 0)
        '
        '
        '
        Me.gvAC.MasterTemplate.AllowDeleteRow = False
        Me.gvAC.MasterTemplate.SelectionMode = Telerik.WinControls.UI.GridViewSelectionMode.CellSelect
        Me.gvAC.MasterTemplate.ShowHeaderCellButtons = True
        Me.gvAC.MasterTemplate.ViewDefinition = TableViewDefinition3
        Me.gvAC.MyExportAPI = False
        Me.gvAC.MyExportFilePath = ""
        Me.gvAC.MyStopExport = False
        Me.gvAC.Name = "gvAC"
        Me.gvAC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gvAC.ShowGroupPanel = False
        Me.gvAC.ShowHeaderCellButtons = True
        Me.gvAC.Size = New System.Drawing.Size(1081, 333)
        Me.gvAC.TabIndex = 1
        Me.gvAC.TabStop = False
        Me.gvAC.VarID = ""
        '
        'RadLabel31
        '
        Me.RadLabel31.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadLabel31.FieldName = Nothing
        Me.RadLabel31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel31.Location = New System.Drawing.Point(836, 5)
        Me.RadLabel31.Name = "RadLabel31"
        Me.RadLabel31.Size = New System.Drawing.Size(130, 16)
        Me.RadLabel31.TabIndex = 126
        Me.RadLabel31.Text = "Total Additional Charges"
        '
        'lblAddCharges
        '
        Me.lblAddCharges.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAddCharges.AutoSize = False
        Me.lblAddCharges.BorderVisible = True
        Me.lblAddCharges.FieldName = Nothing
        Me.lblAddCharges.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddCharges.Location = New System.Drawing.Point(968, 5)
        Me.lblAddCharges.Name = "lblAddCharges"
        Me.lblAddCharges.Size = New System.Drawing.Size(110, 18)
        Me.lblAddCharges.TabIndex = 127
        Me.lblAddCharges.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'pvpCustomFields
        '
        Me.pvpCustomFields.Controls.Add(Me.UcCustomFields1)
        Me.pvpCustomFields.ItemSize = New System.Drawing.SizeF(89.0!, 26.0!)
        Me.pvpCustomFields.Location = New System.Drawing.Point(10, 35)
        Me.pvpCustomFields.Name = "pvpCustomFields"
        Me.pvpCustomFields.Size = New System.Drawing.Size(1081, 370)
        Me.pvpCustomFields.Text = "Custom Fields"
        '
        'UcCustomFields1
        '
        Me.UcCustomFields1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UcCustomFields1.Location = New System.Drawing.Point(0, 0)
        Me.UcCustomFields1.Name = "UcCustomFields1"
        Me.UcCustomFields1.Size = New System.Drawing.Size(1081, 370)
        Me.UcCustomFields1.TabIndex = 1
        '
        'Attachments
        '
        Me.Attachments.Controls.Add(Me.UcAttachment1)
        Me.Attachments.ItemSize = New System.Drawing.SizeF(79.0!, 26.0!)
        Me.Attachments.Location = New System.Drawing.Point(10, 35)
        Me.Attachments.Name = "Attachments"
        Me.Attachments.Size = New System.Drawing.Size(1081, 370)
        Me.Attachments.Text = "Attachments"
        '
        'UcAttachment1
        '
        Me.UcAttachment1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UcAttachment1.Location = New System.Drawing.Point(0, 0)
        Me.UcAttachment1.Name = "UcAttachment1"
        Me.UcAttachment1.Size = New System.Drawing.Size(1081, 370)
        Me.UcAttachment1.TabIndex = 0
        '
        'RadPageViewPage4
        '
        Me.RadPageViewPage4.Controls.Add(Me.btnAdminCancel)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel14)
        Me.RadPageViewPage4.Controls.Add(Me.btnCancel)
        Me.RadPageViewPage4.Controls.Add(Me.lblTotalSubsidyAmt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel2)
        Me.RadPageViewPage4.Controls.Add(Me.lblRoundoffAmount)
        Me.RadPageViewPage4.Controls.Add(Me.pnlCurrConv)
        Me.RadPageViewPage4.Controls.Add(Me.btnReverse)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel32)
        Me.RadPageViewPage4.Controls.Add(Me.lblAddCharges1)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel9)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel27)
        Me.RadPageViewPage4.Controls.Add(Me.lblTotRAmt)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel25)
        Me.RadPageViewPage4.Controls.Add(Me.lblTaxAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblAmtAfterDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.lblDiscountAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblAmtWithDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel22)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel19)
        Me.RadPageViewPage4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage4.ItemSize = New System.Drawing.SizeF(41.0!, 26.0!)
        Me.RadPageViewPage4.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage4.Name = "RadPageViewPage4"
        Me.RadPageViewPage4.Size = New System.Drawing.Size(1081, 370)
        Me.RadPageViewPage4.Text = "Total"
        '
        'btnAdminCancel
        '
        Me.btnAdminCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdminCancel.Location = New System.Drawing.Point(779, 245)
        Me.btnAdminCancel.Name = "btnAdminCancel"
        Me.btnAdminCancel.Size = New System.Drawing.Size(267, 22)
        Me.btnAdminCancel.TabIndex = 139
        Me.btnAdminCancel.Text = "Admin Cancel"
        Me.btnAdminCancel.Visible = False
        '
        'MyLabel14
        '
        Me.MyLabel14.FieldName = Nothing
        Me.MyLabel14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel14.Location = New System.Drawing.Point(123, 219)
        Me.MyLabel14.Name = "MyLabel14"
        Me.MyLabel14.Size = New System.Drawing.Size(89, 16)
        Me.MyLabel14.TabIndex = 138
        Me.MyLabel14.Text = "Subsidy Amount"
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(779, 217)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(267, 22)
        Me.btnCancel.TabIndex = 136
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.Visible = False
        '
        'lblTotalSubsidyAmt
        '
        Me.lblTotalSubsidyAmt.AutoSize = False
        Me.lblTotalSubsidyAmt.BorderVisible = True
        Me.lblTotalSubsidyAmt.FieldName = Nothing
        Me.lblTotalSubsidyAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalSubsidyAmt.Location = New System.Drawing.Point(225, 218)
        Me.lblTotalSubsidyAmt.Name = "lblTotalSubsidyAmt"
        Me.lblTotalSubsidyAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTotalSubsidyAmt.TabIndex = 137
        Me.lblTotalSubsidyAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(124, 196)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(97, 16)
        Me.MyLabel2.TabIndex = 135
        Me.MyLabel2.Text = "Round off Amount"
        '
        'lblRoundoffAmount
        '
        Me.lblRoundoffAmount.AutoSize = False
        Me.lblRoundoffAmount.BorderVisible = True
        Me.lblRoundoffAmount.FieldName = Nothing
        Me.lblRoundoffAmount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRoundoffAmount.Location = New System.Drawing.Point(225, 194)
        Me.lblRoundoffAmount.Name = "lblRoundoffAmount"
        Me.lblRoundoffAmount.Size = New System.Drawing.Size(110, 18)
        Me.lblRoundoffAmount.TabIndex = 134
        Me.lblRoundoffAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCurrConv
        '
        Me.pnlCurrConv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCurrConv.Controls.Add(Me.txtConversionRate)
        Me.pnlCurrConv.Controls.Add(Me.txtCurrencyCode)
        Me.pnlCurrConv.Controls.Add(Me.lblEffectiveFrom)
        Me.pnlCurrConv.Controls.Add(Me.txtApplicableFrom)
        Me.pnlCurrConv.Controls.Add(Me.lblCurrency)
        Me.pnlCurrConv.Controls.Add(Me.lblConvRate)
        Me.pnlCurrConv.Location = New System.Drawing.Point(37, 11)
        Me.pnlCurrConv.Name = "pnlCurrConv"
        Me.pnlCurrConv.Size = New System.Drawing.Size(736, 38)
        Me.pnlCurrConv.TabIndex = 0
        '
        'txtConversionRate
        '
        Me.txtConversionRate.BackColor = System.Drawing.Color.White
        Me.txtConversionRate.CalculationExpression = Nothing
        Me.txtConversionRate.DecimalPlaces = 2
        Me.txtConversionRate.FieldCode = Nothing
        Me.txtConversionRate.FieldDesc = Nothing
        Me.txtConversionRate.FieldMaxLength = 0
        Me.txtConversionRate.FieldName = Nothing
        Me.txtConversionRate.isCalculatedField = False
        Me.txtConversionRate.IsSourceFromTable = False
        Me.txtConversionRate.IsSourceFromValueList = False
        Me.txtConversionRate.IsUnique = False
        Me.txtConversionRate.Location = New System.Drawing.Point(353, 8)
        Me.txtConversionRate.MendatroryField = False
        Me.txtConversionRate.MyLinkLable1 = Nothing
        Me.txtConversionRate.MyLinkLable2 = Nothing
        Me.txtConversionRate.Name = "txtConversionRate"
        Me.txtConversionRate.ReferenceFieldDesc = Nothing
        Me.txtConversionRate.ReferenceFieldName = Nothing
        Me.txtConversionRate.ReferenceTableName = Nothing
        Me.txtConversionRate.Size = New System.Drawing.Size(124, 20)
        Me.txtConversionRate.TabIndex = 1
        Me.txtConversionRate.Text = "1"
        Me.txtConversionRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtConversionRate.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'txtCurrencyCode
        '
        Me.txtCurrencyCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCurrencyCode.CalculationExpression = Nothing
        Me.txtCurrencyCode.FieldCode = Nothing
        Me.txtCurrencyCode.FieldDesc = Nothing
        Me.txtCurrencyCode.FieldMaxLength = 0
        Me.txtCurrencyCode.FieldName = Nothing
        Me.txtCurrencyCode.isCalculatedField = False
        Me.txtCurrencyCode.IsSourceFromTable = False
        Me.txtCurrencyCode.IsSourceFromValueList = False
        Me.txtCurrencyCode.IsUnique = False
        Me.txtCurrencyCode.Location = New System.Drawing.Point(80, 9)
        Me.txtCurrencyCode.MendatroryField = False
        Me.txtCurrencyCode.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurrencyCode.MyLinkLable1 = Nothing
        Me.txtCurrencyCode.MyLinkLable2 = Nothing
        Me.txtCurrencyCode.MyReadOnly = False
        Me.txtCurrencyCode.MyShowMasterFormButton = False
        Me.txtCurrencyCode.Name = "txtCurrencyCode"
        Me.txtCurrencyCode.ReferenceFieldDesc = Nothing
        Me.txtCurrencyCode.ReferenceFieldName = Nothing
        Me.txtCurrencyCode.ReferenceTableName = Nothing
        Me.txtCurrencyCode.Size = New System.Drawing.Size(170, 24)
        Me.txtCurrencyCode.TabIndex = 0
        Me.txtCurrencyCode.Value = ""
        '
        'lblEffectiveFrom
        '
        Me.lblEffectiveFrom.FieldName = Nothing
        Me.lblEffectiveFrom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEffectiveFrom.Location = New System.Drawing.Point(483, 12)
        Me.lblEffectiveFrom.Name = "lblEffectiveFrom"
        Me.lblEffectiveFrom.Size = New System.Drawing.Size(88, 16)
        Me.lblEffectiveFrom.TabIndex = 141
        Me.lblEffectiveFrom.Text = "Applicable From"
        '
        'txtApplicableFrom
        '
        Me.txtApplicableFrom.AutoSize = False
        Me.txtApplicableFrom.BorderVisible = True
        Me.txtApplicableFrom.FieldName = Nothing
        Me.txtApplicableFrom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtApplicableFrom.Location = New System.Drawing.Point(579, 10)
        Me.txtApplicableFrom.Name = "txtApplicableFrom"
        Me.txtApplicableFrom.Size = New System.Drawing.Size(152, 18)
        Me.txtApplicableFrom.TabIndex = 2
        Me.txtApplicableFrom.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrency
        '
        Me.lblCurrency.FieldName = Nothing
        Me.lblCurrency.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrency.Location = New System.Drawing.Point(22, 12)
        Me.lblCurrency.Name = "lblCurrency"
        Me.lblCurrency.Size = New System.Drawing.Size(52, 16)
        Me.lblCurrency.TabIndex = 137
        Me.lblCurrency.Text = "Currency"
        '
        'lblConvRate
        '
        Me.lblConvRate.FieldName = Nothing
        Me.lblConvRate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConvRate.Location = New System.Drawing.Point(256, 12)
        Me.lblConvRate.Name = "lblConvRate"
        Me.lblConvRate.Size = New System.Drawing.Size(91, 16)
        Me.lblConvRate.TabIndex = 139
        Me.lblConvRate.Text = "Conversion Rate"
        '
        'btnReverse
        '
        Me.btnReverse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReverse.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReverse.Location = New System.Drawing.Point(811, 58)
        Me.btnReverse.Name = "btnReverse"
        Me.btnReverse.Size = New System.Drawing.Size(267, 22)
        Me.btnReverse.TabIndex = 2
        Me.btnReverse.Text = "Revese and Unpost"
        Me.btnReverse.Visible = False
        '
        'RadLabel32
        '
        Me.RadLabel32.FieldName = Nothing
        Me.RadLabel32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel32.Location = New System.Drawing.Point(83, 171)
        Me.RadLabel32.Name = "RadLabel32"
        Me.RadLabel32.Size = New System.Drawing.Size(140, 16)
        Me.RadLabel32.TabIndex = 133
        Me.RadLabel32.Text = "+ Total Additional Charges"
        '
        'lblAddCharges1
        '
        Me.lblAddCharges1.AutoSize = False
        Me.lblAddCharges1.BorderVisible = True
        Me.lblAddCharges1.FieldName = Nothing
        Me.lblAddCharges1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddCharges1.Location = New System.Drawing.Point(225, 170)
        Me.lblAddCharges1.Name = "lblAddCharges1"
        Me.lblAddCharges1.Size = New System.Drawing.Size(110, 18)
        Me.lblAddCharges1.TabIndex = 6
        Me.lblAddCharges1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel9
        '
        Me.RadLabel9.FieldName = Nothing
        Me.RadLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel9.Location = New System.Drawing.Point(103, 115)
        Me.RadLabel9.Name = "RadLabel9"
        Me.RadLabel9.Size = New System.Drawing.Size(120, 16)
        Me.RadLabel9.TabIndex = 120
        Me.RadLabel9.Text = "Amount After Discount"
        '
        'RadLabel27
        '
        Me.RadLabel27.FieldName = Nothing
        Me.RadLabel27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel27.Location = New System.Drawing.Point(123, 243)
        Me.RadLabel27.Name = "RadLabel27"
        Me.RadLabel27.Size = New System.Drawing.Size(100, 16)
        Me.RadLabel27.TabIndex = 123
        Me.RadLabel27.Text = "Document Amount"
        '
        'lblTotRAmt
        '
        Me.lblTotRAmt.AutoSize = False
        Me.lblTotRAmt.BorderVisible = True
        Me.lblTotRAmt.FieldName = Nothing
        Me.lblTotRAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRAmt.Location = New System.Drawing.Point(225, 242)
        Me.lblTotRAmt.Name = "lblTotRAmt"
        Me.lblTotRAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTotRAmt.TabIndex = 7
        Me.lblTotRAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel25
        '
        Me.RadLabel25.FieldName = Nothing
        Me.RadLabel25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel25.Location = New System.Drawing.Point(146, 143)
        Me.RadLabel25.Name = "RadLabel25"
        Me.RadLabel25.Size = New System.Drawing.Size(77, 16)
        Me.RadLabel25.TabIndex = 122
        Me.RadLabel25.Text = "+ Tax Amount"
        '
        'lblTaxAmt
        '
        Me.lblTaxAmt.AutoSize = False
        Me.lblTaxAmt.BorderVisible = True
        Me.lblTaxAmt.FieldName = Nothing
        Me.lblTaxAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxAmt.Location = New System.Drawing.Point(225, 142)
        Me.lblTaxAmt.Name = "lblTaxAmt"
        Me.lblTaxAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTaxAmt.TabIndex = 5
        Me.lblTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtAfterDiscount
        '
        Me.lblAmtAfterDiscount.AutoSize = False
        Me.lblAmtAfterDiscount.BorderVisible = True
        Me.lblAmtAfterDiscount.FieldName = Nothing
        Me.lblAmtAfterDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtAfterDiscount.Location = New System.Drawing.Point(225, 114)
        Me.lblAmtAfterDiscount.Name = "lblAmtAfterDiscount"
        Me.lblAmtAfterDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtAfterDiscount.TabIndex = 4
        Me.lblAmtAfterDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDiscountAmt
        '
        Me.lblDiscountAmt.AutoSize = False
        Me.lblDiscountAmt.BorderVisible = True
        Me.lblDiscountAmt.FieldName = Nothing
        Me.lblDiscountAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiscountAmt.Location = New System.Drawing.Point(225, 86)
        Me.lblDiscountAmt.Name = "lblDiscountAmt"
        Me.lblDiscountAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblDiscountAmt.TabIndex = 3
        Me.lblDiscountAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtWithDiscount
        '
        Me.lblAmtWithDiscount.AutoSize = False
        Me.lblAmtWithDiscount.BorderVisible = True
        Me.lblAmtWithDiscount.FieldName = Nothing
        Me.lblAmtWithDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtWithDiscount.Location = New System.Drawing.Point(225, 58)
        Me.lblAmtWithDiscount.Name = "lblAmtWithDiscount"
        Me.lblAmtWithDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtWithDiscount.TabIndex = 1
        Me.lblAmtWithDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel22
        '
        Me.RadLabel22.FieldName = Nothing
        Me.RadLabel22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel22.Location = New System.Drawing.Point(124, 87)
        Me.RadLabel22.Name = "RadLabel22"
        Me.RadLabel22.Size = New System.Drawing.Size(99, 16)
        Me.RadLabel22.TabIndex = 121
        Me.RadLabel22.Text = "- Discount Amount"
        '
        'RadLabel19
        '
        Me.RadLabel19.FieldName = Nothing
        Me.RadLabel19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel19.Location = New System.Drawing.Point(37, 59)
        Me.RadLabel19.Name = "RadLabel19"
        Me.RadLabel19.Size = New System.Drawing.Size(186, 16)
        Me.RadLabel19.TabIndex = 119
        Me.RadLabel19.Text = "Document Amount without Discount"
        '
        'RadPageViewPage5
        '
        Me.RadPageViewPage5.Controls.Add(Me.RadGroupBox3)
        Me.RadPageViewPage5.ItemSize = New System.Drawing.SizeF(63.0!, 26.0!)
        Me.RadPageViewPage5.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage5.Name = "RadPageViewPage5"
        Me.RadPageViewPage5.Size = New System.Drawing.Size(1081, 370)
        Me.RadPageViewPage5.Text = "E-Invoice"
        '
        'RadGroupBox3
        '
        Me.RadGroupBox3.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox3.Controls.Add(Me.txtAckDate)
        Me.RadGroupBox3.Controls.Add(Me.EinvoiceBtnUpdate)
        Me.RadGroupBox3.Controls.Add(Me.EInvoiceQrCode)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel17)
        Me.RadGroupBox3.Controls.Add(Me.EinvoiceAckNo)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel11)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel12)
        Me.RadGroupBox3.Controls.Add(Me.EInvoiceIRNNo)
        Me.RadGroupBox3.Controls.Add(Me.MyLabel13)
        Me.RadGroupBox3.HeaderText = "E-Invoice Detail"
        Me.RadGroupBox3.Location = New System.Drawing.Point(17, 15)
        Me.RadGroupBox3.Name = "RadGroupBox3"
        Me.RadGroupBox3.Size = New System.Drawing.Size(680, 209)
        Me.RadGroupBox3.TabIndex = 1422
        Me.RadGroupBox3.Text = "E-Invoice Detail"
        '
        'txtAckDate
        '
        Me.txtAckDate.CalculationExpression = Nothing
        Me.txtAckDate.CustomFormat = "dd/MM/yyyy hh:mm tt"
        Me.txtAckDate.FieldCode = Nothing
        Me.txtAckDate.FieldDesc = Nothing
        Me.txtAckDate.FieldMaxLength = 0
        Me.txtAckDate.FieldName = Nothing
        Me.txtAckDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAckDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtAckDate.isCalculatedField = False
        Me.txtAckDate.IsSourceFromTable = False
        Me.txtAckDate.IsSourceFromValueList = False
        Me.txtAckDate.IsUnique = False
        Me.txtAckDate.Location = New System.Drawing.Point(84, 60)
        Me.txtAckDate.MendatroryField = False
        Me.txtAckDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtAckDate.MyLinkLable1 = Me.RadLabel4
        Me.txtAckDate.MyLinkLable2 = Nothing
        Me.txtAckDate.Name = "txtAckDate"
        Me.txtAckDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtAckDate.ReferenceFieldDesc = Nothing
        Me.txtAckDate.ReferenceFieldName = Nothing
        Me.txtAckDate.ReferenceTableName = Nothing
        Me.txtAckDate.Size = New System.Drawing.Size(126, 18)
        Me.txtAckDate.TabIndex = 1410
        Me.txtAckDate.TabStop = False
        Me.txtAckDate.Text = "13/06/2011 11:29 AM"
        Me.txtAckDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'EinvoiceBtnUpdate
        '
        Me.EinvoiceBtnUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.EinvoiceBtnUpdate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EinvoiceBtnUpdate.Location = New System.Drawing.Point(84, 176)
        Me.EinvoiceBtnUpdate.Name = "EinvoiceBtnUpdate"
        Me.EinvoiceBtnUpdate.Size = New System.Drawing.Size(141, 20)
        Me.EinvoiceBtnUpdate.TabIndex = 41
        Me.EinvoiceBtnUpdate.Text = "Update"
        '
        'EInvoiceQrCode
        '
        Me.EInvoiceQrCode.AutoSize = False
        Me.EInvoiceQrCode.CalculationExpression = Nothing
        Me.EInvoiceQrCode.FieldCode = Nothing
        Me.EInvoiceQrCode.FieldDesc = Nothing
        Me.EInvoiceQrCode.FieldMaxLength = 0
        Me.EInvoiceQrCode.FieldName = Nothing
        Me.EInvoiceQrCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EInvoiceQrCode.isCalculatedField = False
        Me.EInvoiceQrCode.IsSourceFromTable = False
        Me.EInvoiceQrCode.IsSourceFromValueList = False
        Me.EInvoiceQrCode.IsUnique = False
        Me.EInvoiceQrCode.Location = New System.Drawing.Point(84, 80)
        Me.EInvoiceQrCode.MaxLength = 500
        Me.EInvoiceQrCode.MendatroryField = False
        Me.EInvoiceQrCode.Multiline = True
        Me.EInvoiceQrCode.MyLinkLable1 = Me.MyLabel11
        Me.EInvoiceQrCode.MyLinkLable2 = Nothing
        Me.EInvoiceQrCode.Name = "EInvoiceQrCode"
        Me.EInvoiceQrCode.ReferenceFieldDesc = Nothing
        Me.EInvoiceQrCode.ReferenceFieldName = Nothing
        Me.EInvoiceQrCode.ReferenceTableName = Nothing
        Me.EInvoiceQrCode.Size = New System.Drawing.Size(588, 93)
        Me.EInvoiceQrCode.TabIndex = 1409
        '
        'MyLabel11
        '
        Me.MyLabel11.FieldName = Nothing
        Me.MyLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel11.Location = New System.Drawing.Point(11, 63)
        Me.MyLabel11.Name = "MyLabel11"
        Me.MyLabel11.Size = New System.Drawing.Size(52, 16)
        Me.MyLabel11.TabIndex = 1407
        Me.MyLabel11.Text = "Ack Date"
        '
        'MyLabel17
        '
        Me.MyLabel17.FieldName = Nothing
        Me.MyLabel17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel17.Location = New System.Drawing.Point(11, 85)
        Me.MyLabel17.Name = "MyLabel17"
        Me.MyLabel17.Size = New System.Drawing.Size(53, 16)
        Me.MyLabel17.TabIndex = 1408
        Me.MyLabel17.Text = "QR Code"
        '
        'EinvoiceAckNo
        '
        Me.EinvoiceAckNo.CalculationExpression = Nothing
        Me.EinvoiceAckNo.FieldCode = Nothing
        Me.EinvoiceAckNo.FieldDesc = Nothing
        Me.EinvoiceAckNo.FieldMaxLength = 0
        Me.EinvoiceAckNo.FieldName = Nothing
        Me.EinvoiceAckNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EinvoiceAckNo.isCalculatedField = False
        Me.EinvoiceAckNo.IsSourceFromTable = False
        Me.EinvoiceAckNo.IsSourceFromValueList = False
        Me.EinvoiceAckNo.IsUnique = False
        Me.EinvoiceAckNo.Location = New System.Drawing.Point(84, 41)
        Me.EinvoiceAckNo.MaxLength = 30
        Me.EinvoiceAckNo.MendatroryField = False
        Me.EinvoiceAckNo.MyLinkLable1 = Me.MyLabel11
        Me.EinvoiceAckNo.MyLinkLable2 = Nothing
        Me.EinvoiceAckNo.Name = "EinvoiceAckNo"
        Me.EinvoiceAckNo.ReferenceFieldDesc = Nothing
        Me.EinvoiceAckNo.ReferenceFieldName = Nothing
        Me.EinvoiceAckNo.ReferenceTableName = Nothing
        Me.EinvoiceAckNo.Size = New System.Drawing.Size(300, 18)
        Me.EinvoiceAckNo.TabIndex = 1406
        '
        'MyLabel12
        '
        Me.MyLabel12.FieldName = Nothing
        Me.MyLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel12.Location = New System.Drawing.Point(20, 42)
        Me.MyLabel12.Name = "MyLabel12"
        Me.MyLabel12.Size = New System.Drawing.Size(43, 16)
        Me.MyLabel12.TabIndex = 1405
        Me.MyLabel12.Text = "Ack No"
        '
        'EInvoiceIRNNo
        '
        Me.EInvoiceIRNNo.CalculationExpression = Nothing
        Me.EInvoiceIRNNo.FieldCode = Nothing
        Me.EInvoiceIRNNo.FieldDesc = Nothing
        Me.EInvoiceIRNNo.FieldMaxLength = 0
        Me.EInvoiceIRNNo.FieldName = Nothing
        Me.EInvoiceIRNNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.EInvoiceIRNNo.isCalculatedField = False
        Me.EInvoiceIRNNo.IsSourceFromTable = False
        Me.EInvoiceIRNNo.IsSourceFromValueList = False
        Me.EInvoiceIRNNo.IsUnique = False
        Me.EInvoiceIRNNo.Location = New System.Drawing.Point(84, 22)
        Me.EInvoiceIRNNo.MaxLength = 30
        Me.EInvoiceIRNNo.MendatroryField = False
        Me.EInvoiceIRNNo.MyLinkLable1 = Me.MyLabel13
        Me.EInvoiceIRNNo.MyLinkLable2 = Nothing
        Me.EInvoiceIRNNo.Name = "EInvoiceIRNNo"
        Me.EInvoiceIRNNo.ReferenceFieldDesc = Nothing
        Me.EInvoiceIRNNo.ReferenceFieldName = Nothing
        Me.EInvoiceIRNNo.ReferenceTableName = Nothing
        Me.EInvoiceIRNNo.Size = New System.Drawing.Size(300, 18)
        Me.EInvoiceIRNNo.TabIndex = 1402
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(18, 23)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(43, 16)
        Me.MyLabel13.TabIndex = 1403
        Me.MyLabel13.Text = "IRN No"
        '
        'btnHistory
        '
        Me.btnHistory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHistory.Location = New System.Drawing.Point(698, 6)
        Me.btnHistory.Name = "btnHistory"
        Me.btnHistory.Size = New System.Drawing.Size(69, 22)
        Me.btnHistory.TabIndex = 10
        Me.btnHistory.Text = "History"
        '
        'butCostCenterAndHirerachy_Update_AfterPost
        '
        Me.butCostCenterAndHirerachy_Update_AfterPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butCostCenterAndHirerachy_Update_AfterPost.Location = New System.Drawing.Point(506, 6)
        Me.butCostCenterAndHirerachy_Update_AfterPost.Name = "butCostCenterAndHirerachy_Update_AfterPost"
        Me.butCostCenterAndHirerachy_Update_AfterPost.Size = New System.Drawing.Size(186, 22)
        Me.butCostCenterAndHirerachy_Update_AfterPost.TabIndex = 9
        Me.butCostCenterAndHirerachy_Update_AfterPost.Text = "Update Cost Center And Hirerachy"
        '
        'btnPrintServiceInvoice
        '
        Me.btnPrintServiceInvoice.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrintServiceInvoice.Location = New System.Drawing.Point(374, 6)
        Me.btnPrintServiceInvoice.Name = "btnPrintServiceInvoice"
        Me.btnPrintServiceInvoice.Size = New System.Drawing.Size(126, 22)
        Me.btnPrintServiceInvoice.TabIndex = 8
        Me.btnPrintServiceInvoice.Text = "Print Service Invoice"
        '
        'btnsetting
        '
        Me.btnsetting.Items.AddRange(New Telerik.WinControls.RadItem() {Me.btnsend, Me.btnSendForApproval})
        Me.btnsetting.Location = New System.Drawing.Point(285, 6)
        Me.btnsetting.Name = "btnsetting"
        Me.btnsetting.Size = New System.Drawing.Size(88, 22)
        Me.btnsetting.TabIndex = 7
        Me.btnsetting.Text = "E-Mail/SMS"
        '
        'btnsend
        '
        Me.btnsend.AccessibleDescription = "Send"
        Me.btnsend.AccessibleName = "Send"
        Me.btnsend.Name = "btnsend"
        Me.btnsend.Text = "Send Email/SMS"
        '
        'btnSendForApproval
        '
        Me.btnSendForApproval.Name = "btnSendForApproval"
        Me.btnSendForApproval.Text = "Send For Approval"
        '
        'btnPrint
        '
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(215, 6)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(69, 22)
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.Text = "Print"
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(145, 6)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(69, 22)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        '
        'btnPost
        '
        Me.btnPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPost.Location = New System.Drawing.Point(75, 6)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(69, 22)
        Me.btnPost.TabIndex = 1
        Me.btnPost.Text = "Post"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(1028, 6)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(69, 22)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(5, 6)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(69, 22)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'RadMenuItem1
        '
        Me.RadMenuItem1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem2, Me.RadMenuItem3})
        Me.RadMenuItem1.Name = "RadMenuItem1"
        Me.RadMenuItem1.Text = "Setting"
        '
        'RadMenuItem2
        '
        Me.RadMenuItem2.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem4, Me.mnuExportARInvoiceTrans, Me.mnuExportCN, Me.mnuExportDN, Me.mnuExportOpeningCredit, Me.mnuExortOpeningDebit, Me.rmiExportOPInvoiceForTCS})
        Me.RadMenuItem2.Name = "RadMenuItem2"
        Me.RadMenuItem2.Text = "Export"
        '
        'RadMenuItem4
        '
        Me.RadMenuItem4.Name = "RadMenuItem4"
        Me.RadMenuItem4.Text = "Opening Balance ( Blank Sheet )"
        '
        'mnuExportARInvoiceTrans
        '
        Me.mnuExportARInvoiceTrans.Name = "mnuExportARInvoiceTrans"
        Me.mnuExportARInvoiceTrans.Text = "AR Invoice Transaction"
        '
        'mnuExportCN
        '
        Me.mnuExportCN.AccessibleDescription = "mnuExportCN"
        Me.mnuExportCN.AccessibleName = "mnuExportCN"
        Me.mnuExportCN.Name = "mnuExportCN"
        Me.mnuExportCN.Text = "Export[Credit Note]"
        '
        'mnuExportDN
        '
        Me.mnuExportDN.AccessibleDescription = "RadMenuItem8"
        Me.mnuExportDN.AccessibleName = "RadMenuItem8"
        Me.mnuExportDN.Name = "mnuExportDN"
        Me.mnuExportDN.Text = "Export[Debit Note]"
        '
        'mnuExportOpeningCredit
        '
        Me.mnuExportOpeningCredit.Name = "mnuExportOpeningCredit"
        Me.mnuExportOpeningCredit.Text = "Opening Export[Credit Note]"
        '
        'mnuExortOpeningDebit
        '
        Me.mnuExortOpeningDebit.AccessibleDescription = "RadMenuItem9"
        Me.mnuExortOpeningDebit.AccessibleName = "RadMenuItem9"
        Me.mnuExortOpeningDebit.Name = "mnuExortOpeningDebit"
        Me.mnuExortOpeningDebit.Text = "Opening Export [Debit Note]"
        '
        'rmiExportOPInvoiceForTCS
        '
        Me.rmiExportOPInvoiceForTCS.Name = "rmiExportOPInvoiceForTCS"
        Me.rmiExportOPInvoiceForTCS.Text = "Export OP Invoice For TCS"
        '
        'RadMenuItem3
        '
        Me.RadMenuItem3.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem5, Me.RadMenuItem6, Me.RadMenuItem7, Me.mnuImportCR, Me.mnuImportDN, Me.mnuARInvoiceImport, Me.rmiOPInvoiceForTCS})
        Me.RadMenuItem3.Name = "RadMenuItem3"
        Me.RadMenuItem3.Text = "Import"
        '
        'RadMenuItem5
        '
        Me.RadMenuItem5.Name = "RadMenuItem5"
        Me.RadMenuItem5.Text = "Opening Balance"
        '
        'RadMenuItem6
        '
        Me.RadMenuItem6.AccessibleDescription = "Opening Balance Credit"
        Me.RadMenuItem6.AccessibleName = "Opening Balance Credit"
        Me.RadMenuItem6.Name = "RadMenuItem6"
        Me.RadMenuItem6.Text = "Opening Credit Note"
        '
        'RadMenuItem7
        '
        Me.RadMenuItem7.AccessibleDescription = "Opening Balance Debit"
        Me.RadMenuItem7.AccessibleName = "Opening Balance Debit"
        Me.RadMenuItem7.Name = "RadMenuItem7"
        Me.RadMenuItem7.Text = "Opening Debit Note"
        '
        'mnuImportCR
        '
        Me.mnuImportCR.Name = "mnuImportCR"
        Me.mnuImportCR.Text = "Import[Credit Note]"
        '
        'mnuImportDN
        '
        Me.mnuImportDN.Name = "mnuImportDN"
        Me.mnuImportDN.Text = "Import[Debit Note]"
        '
        'mnuARInvoiceImport
        '
        Me.mnuARInvoiceImport.Name = "mnuARInvoiceImport"
        Me.mnuARInvoiceImport.Text = "Import AR Invoice"
        '
        'rmiOPInvoiceForTCS
        '
        Me.rmiOPInvoiceForTCS.Name = "rmiOPInvoiceForTCS"
        Me.rmiOPInvoiceForTCS.Text = "Import OP Invoice For TCS"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SplitContainer1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 20)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1102, 455)
        Me.Panel1.TabIndex = 3
        '
        'RadMenu1
        '
        Me.RadMenu1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem1})
        Me.RadMenu1.Location = New System.Drawing.Point(0, 0)
        Me.RadMenu1.Name = "RadMenu1"
        Me.RadMenu1.Size = New System.Drawing.Size(1102, 20)
        Me.RadMenu1.TabIndex = 1
        '
        'FrmARInvoiceEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1102, 475)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RadMenu1)
        Me.MinimumSize = New System.Drawing.Size(890, 467)
        Me.Name = "FrmARInvoiceEntry"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "AR Invoice Entry"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageView1.ResumeLayout(False)
        Me.RadPageViewPage1.ResumeLayout(False)
        Me.RadPageViewPage1.PerformLayout()
        CType(Me.chkEinvoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLocation.ResumeLayout(False)
        Me.pnlLocation.PerformLayout()
        CType(Me.RadLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLocationPrefixName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDataAndTimeSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTapalNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAgainstServiceInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAgainstSecurityReceiptNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtServiceVisitCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVCGNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSaleReturnNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSaleInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ddlSecDepositType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSecurityDposit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LblLocDesp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReturnType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ddlReturnType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPCJ.ResumeLayout(False)
        Me.pnlPCJ.PerformLayout()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblProject, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDrillDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotRAmt1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox2.ResumeLayout(False)
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkOnHold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVendorName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtOrderNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboDocType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDesc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage2.ResumeLayout(False)
        Me.RadPageViewPage2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.rbtnTaxCalManual, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnTaxCalAutomatic, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.RadLabel16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDueDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTermName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv2.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTaxGrpName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.gvAC.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvAC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAddCharges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pvpCustomFields.ResumeLayout(False)
        Me.Attachments.ResumeLayout(False)
        Me.RadPageViewPage4.ResumeLayout(False)
        Me.RadPageViewPage4.PerformLayout()
        CType(Me.btnAdminCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotalSubsidyAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRoundoffAmount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCurrConv.ResumeLayout(False)
        Me.pnlCurrConv.PerformLayout()
        CType(Me.txtConversionRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblEffectiveFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtApplicableFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCurrency, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblConvRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAddCharges1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotRAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTaxAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAmtAfterDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDiscountAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAmtWithDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel19, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage5.ResumeLayout(False)
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox3.ResumeLayout(False)
        Me.RadGroupBox3.PerformLayout()
        CType(Me.txtAckDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EinvoiceBtnUpdate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EInvoiceQrCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EinvoiceAckNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.EInvoiceIRNNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.butCostCenterAndHirerachy_Update_AfterPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrintServiceInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnsetting, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents RadPageView1 As Telerik.WinControls.UI.RadPageView
    Friend WithEvents RadPageViewPage1 As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents RadPageViewPage2 As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents RadPageViewPage4 As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents btnAddNew As Telerik.WinControls.UI.RadButton
    Friend WithEvents chkOnHold As Telerik.WinControls.UI.RadCheckBox
    Friend WithEvents txtDesc As common.Controls.MyTextBox
    Friend WithEvents txtDate As common.Controls.MyDateTimePicker
    Friend WithEvents txtOrderNo As common.Controls.MyTextBox
    Friend WithEvents cboDocType As common.Controls.MyComboBox
    Friend WithEvents gv1 As common.UserControls.MyRadGridView
    Friend WithEvents gv2 As common.UserControls.MyRadGridView
    Friend WithEvents btnClose As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnSave As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtDueDate As common.Controls.MyDateTimePicker
    Friend WithEvents RadGroupBox1 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents RadGroupBox2 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents btnPost As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnDelete As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtDocNo As common.UserControls.txtNavigator
    Friend WithEvents btnPrint As Telerik.WinControls.UI.RadButton
    Friend WithEvents UsLock1 As common.usLock
    Friend WithEvents txtRefDocNo As common.UserControls.txtFinder
    Friend WithEvents TxtCustomer As common.UserControls.txtFinder
    Friend WithEvents txtTaxGroup As common.UserControls.txtFinder
    Friend WithEvents txtTermCode As common.UserControls.txtFinder
    Friend WithEvents txtACSet As common.UserControls.txtFinder
    Friend WithEvents RadPageViewPage3 As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents gvAC As common.UserControls.MyRadGridView
    Friend WithEvents btnDrillDown As Telerik.WinControls.UI.RadButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbtnTaxCalManual As common.Controls.MyRadioButton
    Friend WithEvents rbtnTaxCalAutomatic As common.Controls.MyRadioButton
    Friend WithEvents RadLabel2 As common.Controls.MyLabel
    Friend WithEvents RadLabel1 As common.Controls.MyLabel
    Friend WithEvents lblVendorName As common.Controls.MyLabel
    Friend WithEvents RadLabel4 As common.Controls.MyLabel
    Friend WithEvents RadLabel5 As common.Controls.MyLabel
    Friend WithEvents RadLabel8 As common.Controls.MyLabel
    Friend WithEvents RadLabel6 As common.Controls.MyLabel
    Friend WithEvents lblTaxGrpName As common.Controls.MyLabel
    Friend WithEvents RadLabel11 As common.Controls.MyLabel
    Friend WithEvents RadLabel25 As common.Controls.MyLabel
    Friend WithEvents lblTaxAmt As common.Controls.MyLabel
    Friend WithEvents lblAmtAfterDiscount As common.Controls.MyLabel
    Friend WithEvents lblDiscountAmt As common.Controls.MyLabel
    Friend WithEvents lblAmtWithDiscount As common.Controls.MyLabel
    Friend WithEvents RadLabel22 As common.Controls.MyLabel
    Friend WithEvents RadLabel19 As common.Controls.MyLabel
    Friend WithEvents RadLabel27 As common.Controls.MyLabel
    Friend WithEvents lblTotRAmt As common.Controls.MyLabel
    Friend WithEvents RadLabel17 As common.Controls.MyLabel
    Friend WithEvents lblTermName As common.Controls.MyLabel
    Friend WithEvents RadLabel16 As common.Controls.MyLabel
    Friend WithEvents RadLabel9 As common.Controls.MyLabel
    Friend WithEvents RadLabel10 As common.Controls.MyLabel
    Friend WithEvents RadLabel12 As common.Controls.MyLabel
    Friend WithEvents lblTotRAmt1 As common.Controls.MyLabel
    Friend WithEvents RadLabel15 As common.Controls.MyLabel
    Friend WithEvents RadLabel31 As common.Controls.MyLabel
    Friend WithEvents lblAddCharges As common.Controls.MyLabel
    Friend WithEvents RadLabel32 As common.Controls.MyLabel
    Friend WithEvents lblAddCharges1 As common.Controls.MyLabel
    Friend WithEvents RadMenuItem1 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem2 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem4 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem3 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem5 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadMenuItem6 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem7 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadLabel3 As Telerik.WinControls.UI.RadLabel
    Friend WithEvents txtlocation As common.UserControls.txtFinder
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents btnReverse As Telerik.WinControls.UI.RadButton
    Friend WithEvents pnlCurrConv As System.Windows.Forms.Panel
    Friend WithEvents txtCurrencyCode As common.UserControls.txtFinder
    Friend WithEvents lblEffectiveFrom As common.Controls.MyLabel
    Friend WithEvents txtApplicableFrom As common.Controls.MyLabel
    Friend WithEvents lblCurrency As common.Controls.MyLabel
    Friend WithEvents lblConvRate As common.Controls.MyLabel
    Friend WithEvents pvpCustomFields As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents UcCustomFields1 As XpertERPEngine.ucCustomFields
    Friend WithEvents txtConversionRate As common.MyNumBox
    Friend WithEvents Attachments As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents UcAttachment1 As XpertERPEngine.ucAttachment
    Friend WithEvents pnlPCJ As System.Windows.Forms.Panel
    Friend WithEvents MyLabel4 As common.Controls.MyLabel
    Friend WithEvents fndProject As common.UserControls.txtFinder
    Friend WithEvents lblProject As common.Controls.MyLabel
    Friend WithEvents btnsetting As Telerik.WinControls.UI.RadSplitButton
    Friend WithEvents btnsend As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents btnSendForApproval As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents lblReturnType As common.Controls.MyLabel
    Friend WithEvents ddlReturnType As common.Controls.MyComboBox
    Friend WithEvents MyLabel6 As common.Controls.MyLabel
    Friend WithEvents MyLabel5 As common.Controls.MyLabel
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents lblRoundoffAmount As common.Controls.MyLabel
    Friend WithEvents LblLocDesp As common.Controls.MyLabel
    Friend WithEvents MyLabel3 As common.Controls.MyLabel
    Friend WithEvents ddlSecDepositType As common.Controls.MyComboBox
    Friend WithEvents chkSecurityDposit As Telerik.WinControls.UI.RadCheckBox
    Friend WithEvents MyLabel7 As common.Controls.MyLabel
    Friend WithEvents mnuExportARInvoiceTrans As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuExportCN As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuExportDN As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuImportCR As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuImportDN As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenu1 As Telerik.WinControls.UI.RadMenu
    Friend WithEvents mnuExportOpeningCredit As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuExortOpeningDebit As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuARInvoiceImport As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents txtServiceVisitCode As common.Controls.MyTextBox
    Friend WithEvents txtVCGNo As common.Controls.MyTextBox
    Friend WithEvents txtSaleReturnNo As common.Controls.MyTextBox
    Friend WithEvents txtSaleInvoice As common.Controls.MyTextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents txtSercurityReceiptNo As common.UserControls.txtFinder
    Friend WithEvents MyLabel8 As common.Controls.MyLabel
    Friend WithEvents chkAgainstSecurityReceiptNo As Telerik.WinControls.UI.RadCheckBox
    Friend WithEvents chkAgainstServiceInvoice As Telerik.WinControls.UI.RadCheckBox
    Friend WithEvents btnPrintServiceInvoice As Telerik.WinControls.UI.RadButton
    Friend WithEvents MyLabel10 As common.Controls.MyLabel
    Friend WithEvents txtDataAndTimeSelection As common.Controls.MyDateTimePicker
    Friend WithEvents MyLabel9 As common.Controls.MyLabel
    Friend WithEvents txtTapalNo As common.Controls.MyTextBox
    Friend WithEvents btnCancel As RadButton
    Friend WithEvents butCostCenterAndHirerachy_Update_AfterPost As RadButton
    Friend WithEvents rmiOPInvoiceForTCS As RadMenuItem
    Friend WithEvents rmiExportOPInvoiceForTCS As RadMenuItem
    Friend WithEvents btnHistory As RadButton
    Friend WithEvents txtLocationPrefixName As common.Controls.MyLabel
    Friend WithEvents txtLocationPrefix As common.UserControls.txtFinder
    Friend WithEvents RadLabel7 As RadLabel
    Friend WithEvents pnlLocation As Panel
    Friend WithEvents chkEinvoice As RadCheckBox
    Friend WithEvents RadPageViewPage5 As RadPageViewPage
    Friend WithEvents RadGroupBox3 As RadGroupBox
    Friend WithEvents txtAckDate As common.Controls.MyDateTimePicker
    Friend WithEvents EinvoiceBtnUpdate As RadButton
    Friend WithEvents EInvoiceQrCode As common.Controls.MyTextBox
    Friend WithEvents MyLabel11 As common.Controls.MyLabel
    Friend WithEvents MyLabel17 As common.Controls.MyLabel
    Friend WithEvents EinvoiceAckNo As common.Controls.MyTextBox
    Friend WithEvents MyLabel12 As common.Controls.MyLabel
    Friend WithEvents EInvoiceIRNNo As common.Controls.MyTextBox
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents MyLabel14 As common.Controls.MyLabel
    Friend WithEvents lblTotalSubsidyAmt As common.Controls.MyLabel
    Friend WithEvents btnAdminCancel As RadButton
End Class

