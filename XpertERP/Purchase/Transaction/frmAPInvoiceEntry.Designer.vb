<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmAPInvoiceEntry
    Inherits FrmMainTranScreen

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim RadListDataItem1 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem2 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem3 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem4 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim TableViewDefinition1 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAPInvoiceEntry))
        Dim RadListDataItem5 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem6 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim TableViewDefinition2 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition3 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.RadPageViewPage1 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.pnlLocation = New System.Windows.Forms.Panel()
        Me.RadLabel23 = New Telerik.WinControls.UI.RadLabel()
        Me.txtLocationPrefixName = New common.Controls.MyLabel()
        Me.txtLocationPrefix = New common.UserControls.txtFinder()
        Me.txtSecurityAdjusment = New common.UserControls.txtMultiSelectFinder()
        Me.MyLabel15 = New common.Controls.MyLabel()
        Me.ChkTrnsferToSvng = New common.Controls.MyCheckBox()
        Me.grpVendorBankDetails = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtVendor_Bank_ACNo = New common.Controls.MyTextBox()
        Me.MyLabel17 = New common.Controls.MyLabel()
        Me.txtBankCode = New common.Controls.MyTextBox()
        Me.MyLabel22 = New common.Controls.MyLabel()
        Me.TxtIFSCCode = New common.Controls.MyTextBox()
        Me.txtbranchname = New common.Controls.MyTextBox()
        Me.TxtBankName = New common.Controls.MyTextBox()
        Me.MyLabel23 = New common.Controls.MyLabel()
        Me.chkSecurityAdjusment = New common.Controls.MyCheckBox()
        Me.chkTDSProvision = New common.Controls.MyCheckBox()
        Me.chkRCM = New common.Controls.MyCheckBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.lblGstinNo = New common.Controls.MyLabel()
        Me.MyLabel49 = New common.Controls.MyLabel()
        Me.MyLabel45 = New common.Controls.MyLabel()
        Me.lblRegisterOrUnregister = New common.Controls.MyLabel()
        Me.RadGroupBox4 = New Telerik.WinControls.UI.RadGroupBox()
        Me.CboxITCCateogory = New common.Controls.MyComboBox()
        Me.MyLabel12 = New common.Controls.MyLabel()
        Me.CboxITCType = New common.Controls.MyComboBox()
        Me.chkITCEligible = New Telerik.WinControls.UI.RadCheckBox()
        Me.RadGroupBox2 = New Telerik.WinControls.UI.RadGroupBox()
        Me.gv1 = New common.UserControls.MyRadGridView()
        Me.pnlDeduction = New System.Windows.Forms.Panel()
        Me.pnlSecondaryTranporterPaymentCycle = New System.Windows.Forms.Panel()
        Me.MyLabel11 = New common.Controls.MyLabel()
        Me.MyLabel10 = New common.Controls.MyLabel()
        Me.txtToDateSecTrans = New common.Controls.MyDateTimePicker()
        Me.txtFromDateSecTrans = New common.Controls.MyDateTimePicker()
        Me.ChkSecurity = New common.Controls.MyCheckBox()
        Me.chkSecTrans = New common.Controls.MyCheckBox()
        Me.RadLabel21 = New Telerik.WinControls.UI.RadLabel()
        Me.LblCostCentre = New common.Controls.MyLabel()
        Me.RadLabel20 = New Telerik.WinControls.UI.RadLabel()
        Me.pnlPCJ = New System.Windows.Forms.Panel()
        Me.MyLabel4 = New common.Controls.MyLabel()
        Me.fndProject = New common.UserControls.txtFinder()
        Me.lblProject = New common.Controls.MyLabel()
        Me.RadLabel18 = New Telerik.WinControls.UI.RadLabel()
        Me.btnDrillDown = New Telerik.WinControls.UI.RadButton()
        Me.chkOnHold = New Telerik.WinControls.UI.RadCheckBox()
        Me.btnAddNew = New Telerik.WinControls.UI.RadButton()
        Me.txtMCC = New common.Controls.MyTextBox()
        Me.lblMCC = New common.Controls.MyTextBox()
        Me.lblMCC2 = New common.Controls.MyLabel()
        Me.txtDataAndTimeSelection = New common.Controls.MyDateTimePicker()
        Me.MyLabel16 = New common.Controls.MyLabel()
        Me.txtTapalNo = New common.Controls.MyTextBox()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.ddlEmployeeType = New common.Controls.MyComboBox()
        Me.txtTotalAmt = New common.MyNumBox()
        Me.lblcaption2 = New common.Controls.MyLabel()
        Me.dtBillDate = New common.Controls.MyDateTimePicker()
        Me.lblCaption1 = New common.Controls.MyLabel()
        Me.txtBillNo = New common.Controls.MyTextBox()
        Me.chkProRated = New common.Controls.MyCheckBox()
        Me.MyLabel9 = New common.Controls.MyLabel()
        Me.txtAdd_Doc_TYpe = New common.Controls.MyLabel()
        Me.LblVCGL = New common.Controls.MyTextBox()
        Me.TxtCostCentre = New common.UserControls.txtFinder()
        Me.LblHirerachy = New common.Controls.MyLabel()
        Me.TxtHirerachy = New common.UserControls.txtFinder()
        Me.MyLabel8 = New common.Controls.MyLabel()
        Me.chkDeduction = New common.Controls.MyCheckBox()
        Me.LblLocDesp = New common.Controls.MyLabel()
        Me.chkProvision = New common.Controls.MyCheckBox()
        Me.txtlocation = New common.UserControls.txtFinder()
        Me.chkQuickMode = New common.Controls.MyCheckBox()
        Me.txtACSet = New common.UserControls.txtFinder()
        Me.RadLabel6 = New common.Controls.MyLabel()
        Me.TxtVendorNo = New common.UserControls.txtFinder()
        Me.RadLabel2 = New common.Controls.MyLabel()
        Me.lblVendorName = New common.Controls.MyLabel()
        Me.txtRefDocNo = New common.UserControls.txtFinder()
        Me.RadLabel15 = New common.Controls.MyLabel()
        Me.UsLock1 = New common.usLock()
        Me.RadLabel14 = New common.Controls.MyLabel()
        Me.RadLabel13 = New common.Controls.MyLabel()
        Me.RadLabel12 = New common.Controls.MyLabel()
        Me.lblTotRAmt1 = New common.Controls.MyLabel()
        Me.RadLabel3 = New common.Controls.MyLabel()
        Me.RadLabel8 = New common.Controls.MyLabel()
        Me.RadLabel7 = New common.Controls.MyLabel()
        Me.RadLabel5 = New common.Controls.MyLabel()
        Me.RadLabel4 = New common.Controls.MyLabel()
        Me.RadLabel1 = New common.Controls.MyLabel()
        Me.cmbRefType = New common.Controls.MyComboBox()
        Me.txtDocNo = New common.UserControls.txtNavigator()
        Me.txtVendorInvDatre = New common.Controls.MyDateTimePicker()
        Me.txtVendorInvoiceNo = New common.Controls.MyTextBox()
        Me.txtOrderNo = New common.Controls.MyTextBox()
        Me.txtPONo = New common.Controls.MyTextBox()
        Me.cboDocType = New common.Controls.MyComboBox()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.txtDesc = New common.Controls.MyTextBox()
        Me.grpProvision = New System.Windows.Forms.GroupBox()
        Me.btlShowProvision = New System.Windows.Forms.LinkLabel()
        Me.MyLabel7 = New common.Controls.MyLabel()
        Me.txtProvAmt = New common.Controls.MyTextBox()
        Me.btnProvSelect = New System.Windows.Forms.LinkLabel()
        Me.MyLabel6 = New common.Controls.MyLabel()
        Me.dtpToProv = New common.Controls.MyDateTimePicker()
        Me.MyLabel5 = New common.Controls.MyLabel()
        Me.dtpFromProv = New common.Controls.MyDateTimePicker()
        Me.RadPageViewPage2 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbtnTaxCalManual = New common.Controls.MyRadioButton()
        Me.rbtnTaxCalAutomatic = New common.Controls.MyRadioButton()
        Me.txtTaxGroup = New common.UserControls.txtFinder()
        Me.RadLabel11 = New common.Controls.MyLabel()
        Me.lblTaxGrpName = New common.Controls.MyLabel()
        Me.RadLabel10 = New common.Controls.MyLabel()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtTermCode = New common.UserControls.txtFinder()
        Me.RadLabel16 = New common.Controls.MyLabel()
        Me.lblTermName = New common.Controls.MyLabel()
        Me.txtDueDate = New common.Controls.MyDateTimePicker()
        Me.RadLabel17 = New common.Controls.MyLabel()
        Me.gv2 = New common.UserControls.MyRadGridView()
        Me.RadPageViewPage3 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.gvAC = New common.UserControls.MyRadGridView()
        Me.RadLabel31 = New common.Controls.MyLabel()
        Me.lblAddCharges = New common.Controls.MyLabel()
        Me.pvpCustomFields = New Telerik.WinControls.UI.RadPageViewPage()
        Me.Attachments = New Telerik.WinControls.UI.RadPageViewPage()
        Me.UcAttachment1 = New XpertERPEngine.ucAttachment()
        Me.RadPageViewPage4 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.MyLabel14 = New common.Controls.MyLabel()
        Me.lblRoundOff = New common.Controls.MyLabel()
        Me.TxtCashDiscount = New common.MyNumBox()
        Me.txtAmtAfterCashDiscount = New common.Controls.MyLabel()
        Me.lblAmtAfterCashDiscount = New common.Controls.MyLabel()
        Me.lblCashDiscount = New common.Controls.MyLabel()
        Me.pnlCurrConv = New System.Windows.Forms.Panel()
        Me.txtConversionRate = New common.MyNumBox()
        Me.txtCurrencyCode = New common.UserControls.txtFinder()
        Me.lblEffectiveFrom = New common.Controls.MyLabel()
        Me.txtApplicableFrom = New common.Controls.MyLabel()
        Me.lblCurrency = New common.Controls.MyLabel()
        Me.lblConvRate = New common.Controls.MyLabel()
        Me.btnReverse = New Telerik.WinControls.UI.RadButton()
        Me.MyLabel3 = New common.Controls.MyLabel()
        Me.lblLandedAmt = New common.Controls.MyLabel()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.lblTotEmptyAmt = New common.Controls.MyLabel()
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
        Me.btnHistory = New Telerik.WinControls.UI.RadButton()
        Me.butCostCenterAndHirerachy_Update_AfterPost = New Telerik.WinControls.UI.RadButton()
        Me.btnsetting = New Telerik.WinControls.UI.RadSplitButton()
        Me.btnSend = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnSendForApproval = New Telerik.WinControls.UI.RadMenuItem()
        Me.btnPrint = New Telerik.WinControls.UI.RadButton()
        Me.btnDelete = New Telerik.WinControls.UI.RadButton()
        Me.btnPost = New Telerik.WinControls.UI.RadButton()
        Me.btnViewTDSDetails = New Telerik.WinControls.UI.RadButton()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.btnSave = New Telerik.WinControls.UI.RadButton()
        Me.RadMenu1 = New Telerik.WinControls.UI.RadMenu()
        Me.RadMenuItem1 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem2 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem4 = New Telerik.WinControls.UI.RadMenuItem()
        Me.mbtnExportApTransaction = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmiExportDrNote = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuCreditNoteExport = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem3 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem5 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem6 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem7 = New Telerik.WinControls.UI.RadMenuItem()
        Me.mbtnImportApTransaction = New Telerik.WinControls.UI.RadMenuItem()
        Me.rmiImportDrNote = New Telerik.WinControls.UI.RadMenuItem()
        Me.mnuImportCreditNote = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem10 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem8 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem9 = New Telerik.WinControls.UI.RadMenuItem()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.RadPageViewPage1.SuspendLayout()
        Me.pnlLocation.SuspendLayout()
        CType(Me.RadLabel23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLocationPrefixName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChkTrnsferToSvng, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpVendorBankDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpVendorBankDetails.SuspendLayout()
        CType(Me.txtVendor_Bank_ACNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBankCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtIFSCCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtbranchname, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtBankName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSecurityAdjusment, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkTDSProvision, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkRCM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        CType(Me.lblGstinNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel49, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel45, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRegisterOrUnregister, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox4.SuspendLayout()
        CType(Me.CboxITCCateogory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CboxITCType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkITCEligible, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox2.SuspendLayout()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlDeduction.SuspendLayout()
        Me.pnlSecondaryTranporterPaymentCycle.SuspendLayout()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtToDateSecTrans, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFromDateSecTrans, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ChkSecurity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSecTrans, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel21, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadLabel21.SuspendLayout()
        CType(Me.LblCostCentre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel20, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPCJ.SuspendLayout()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblProject, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDrillDown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkOnHold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtMCC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMCC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMCC2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDataAndTimeSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTapalNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ddlEmployeeType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTotalAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblcaption2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtBillDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCaption1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBillNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkProRated, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAdd_Doc_TYpe, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LblVCGL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LblHirerachy, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkDeduction, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LblLocDesp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkProvision, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkQuickMode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVendorName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotRAmt1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbRefType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVendorInvDatre, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVendorInvoiceNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtOrderNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPONo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboDocType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpProvision.SuspendLayout()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtProvAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpToProv, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpFromProv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.rbtnTaxCalManual, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnTaxCalAutomatic, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTaxGrpName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.RadLabel16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTermName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDueDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv2.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage3.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.gvAC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gvAC.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAddCharges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Attachments.SuspendLayout()
        Me.RadPageViewPage4.SuspendLayout()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRoundOff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtCashDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAmtAfterCashDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAmtAfterCashDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCashDiscount, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCurrConv.SuspendLayout()
        CType(Me.txtConversionRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblEffectiveFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtApplicableFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCurrency, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblConvRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLandedAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotEmptyAmt, System.ComponentModel.ISupportInitialize).BeginInit()
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
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.butCostCenterAndHirerachy_Update_AfterPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnsetting, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnViewTDSDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.SplitContainer1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SplitContainer1.IsSplitterFixed = True
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
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnsetting)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrint)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnViewTDSDetails)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Size = New System.Drawing.Size(1126, 530)
        Me.SplitContainer1.SplitterDistance = 492
        Me.SplitContainer1.TabIndex = 0
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage1)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage2)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage3)
        Me.RadPageView1.Controls.Add(Me.pvpCustomFields)
        Me.RadPageView1.Controls.Add(Me.Attachments)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage4)
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.RadPageViewPage4
        Me.RadPageView1.Size = New System.Drawing.Size(1126, 492)
        Me.RadPageView1.TabIndex = 0
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).ItemAlignment = Telerik.WinControls.UI.StripViewItemAlignment.Near
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripAlignment = Telerik.WinControls.UI.StripViewAlignment.Top
        '
        'RadPageViewPage1
        '
        Me.RadPageViewPage1.Controls.Add(Me.pnlLocation)
        Me.RadPageViewPage1.Controls.Add(Me.txtSecurityAdjusment)
        Me.RadPageViewPage1.Controls.Add(Me.ChkTrnsferToSvng)
        Me.RadPageViewPage1.Controls.Add(Me.grpVendorBankDetails)
        Me.RadPageViewPage1.Controls.Add(Me.chkSecurityAdjusment)
        Me.RadPageViewPage1.Controls.Add(Me.chkTDSProvision)
        Me.RadPageViewPage1.Controls.Add(Me.chkRCM)
        Me.RadPageViewPage1.Controls.Add(Me.Panel4)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox4)
        Me.RadPageViewPage1.Controls.Add(Me.chkITCEligible)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox2)
        Me.RadPageViewPage1.Controls.Add(Me.pnlDeduction)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel21)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel20)
        Me.RadPageViewPage1.Controls.Add(Me.pnlPCJ)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel18)
        Me.RadPageViewPage1.Controls.Add(Me.btnDrillDown)
        Me.RadPageViewPage1.Controls.Add(Me.chkOnHold)
        Me.RadPageViewPage1.Controls.Add(Me.btnAddNew)
        Me.RadPageViewPage1.Controls.Add(Me.txtMCC)
        Me.RadPageViewPage1.Controls.Add(Me.lblMCC)
        Me.RadPageViewPage1.Controls.Add(Me.lblMCC2)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel15)
        Me.RadPageViewPage1.Controls.Add(Me.txtDataAndTimeSelection)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel16)
        Me.RadPageViewPage1.Controls.Add(Me.txtTapalNo)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel13)
        Me.RadPageViewPage1.Controls.Add(Me.ddlEmployeeType)
        Me.RadPageViewPage1.Controls.Add(Me.txtTotalAmt)
        Me.RadPageViewPage1.Controls.Add(Me.lblcaption2)
        Me.RadPageViewPage1.Controls.Add(Me.dtBillDate)
        Me.RadPageViewPage1.Controls.Add(Me.lblCaption1)
        Me.RadPageViewPage1.Controls.Add(Me.txtBillNo)
        Me.RadPageViewPage1.Controls.Add(Me.chkProRated)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel9)
        Me.RadPageViewPage1.Controls.Add(Me.txtAdd_Doc_TYpe)
        Me.RadPageViewPage1.Controls.Add(Me.LblVCGL)
        Me.RadPageViewPage1.Controls.Add(Me.TxtCostCentre)
        Me.RadPageViewPage1.Controls.Add(Me.LblHirerachy)
        Me.RadPageViewPage1.Controls.Add(Me.TxtHirerachy)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel8)
        Me.RadPageViewPage1.Controls.Add(Me.chkDeduction)
        Me.RadPageViewPage1.Controls.Add(Me.LblLocDesp)
        Me.RadPageViewPage1.Controls.Add(Me.chkProvision)
        Me.RadPageViewPage1.Controls.Add(Me.txtlocation)
        Me.RadPageViewPage1.Controls.Add(Me.chkQuickMode)
        Me.RadPageViewPage1.Controls.Add(Me.txtACSet)
        Me.RadPageViewPage1.Controls.Add(Me.TxtVendorNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtRefDocNo)
        Me.RadPageViewPage1.Controls.Add(Me.UsLock1)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel15)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel14)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel13)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel12)
        Me.RadPageViewPage1.Controls.Add(Me.lblTotRAmt1)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel3)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel8)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel7)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel6)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel5)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel4)
        Me.RadPageViewPage1.Controls.Add(Me.lblVendorName)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel2)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.cmbRefType)
        Me.RadPageViewPage1.Controls.Add(Me.txtDocNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtVendorInvDatre)
        Me.RadPageViewPage1.Controls.Add(Me.txtVendorInvoiceNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtOrderNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtPONo)
        Me.RadPageViewPage1.Controls.Add(Me.cboDocType)
        Me.RadPageViewPage1.Controls.Add(Me.txtDate)
        Me.RadPageViewPage1.Controls.Add(Me.txtDesc)
        Me.RadPageViewPage1.Controls.Add(Me.grpProvision)
        Me.RadPageViewPage1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage1.ItemSize = New System.Drawing.SizeF(68.0!, 26.0!)
        Me.RadPageViewPage1.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage1.Name = "RadPageViewPage1"
        Me.RadPageViewPage1.Size = New System.Drawing.Size(1105, 446)
        Me.RadPageViewPage1.Text = "Document"
        '
        'pnlLocation
        '
        Me.pnlLocation.Controls.Add(Me.RadLabel23)
        Me.pnlLocation.Controls.Add(Me.txtLocationPrefixName)
        Me.pnlLocation.Controls.Add(Me.txtLocationPrefix)
        Me.pnlLocation.Location = New System.Drawing.Point(5, 258)
        Me.pnlLocation.Name = "pnlLocation"
        Me.pnlLocation.Size = New System.Drawing.Size(644, 23)
        Me.pnlLocation.TabIndex = 12148
        '
        'RadLabel23
        '
        Me.RadLabel23.Location = New System.Drawing.Point(3, 3)
        Me.RadLabel23.Name = "RadLabel23"
        Me.RadLabel23.Size = New System.Drawing.Size(49, 18)
        Me.RadLabel23.TabIndex = 12145
        Me.RadLabel23.Text = "Location"
        '
        'txtLocationPrefixName
        '
        Me.txtLocationPrefixName.AutoSize = False
        Me.txtLocationPrefixName.BorderVisible = True
        Me.txtLocationPrefixName.FieldName = Nothing
        Me.txtLocationPrefixName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocationPrefixName.Location = New System.Drawing.Point(236, 3)
        Me.txtLocationPrefixName.Name = "txtLocationPrefixName"
        Me.txtLocationPrefixName.Size = New System.Drawing.Size(373, 18)
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
        Me.txtLocationPrefix.Location = New System.Drawing.Point(106, 2)
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
        'txtSecurityAdjusment
        '
        Me.txtSecurityAdjusment.arrDispalyMember = Nothing
        Me.txtSecurityAdjusment.arrValueMember = Nothing
        Me.txtSecurityAdjusment.Location = New System.Drawing.Point(860, 149)
        Me.txtSecurityAdjusment.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSecurityAdjusment.MyLinkLable1 = Me.MyLabel15
        Me.txtSecurityAdjusment.MyLinkLable2 = Nothing
        Me.txtSecurityAdjusment.MyNullText = "Please select"
        Me.txtSecurityAdjusment.Name = "txtSecurityAdjusment"
        Me.txtSecurityAdjusment.Size = New System.Drawing.Size(231, 20)
        Me.txtSecurityAdjusment.TabIndex = 615
        Me.txtSecurityAdjusment.Visible = False
        '
        'MyLabel15
        '
        Me.MyLabel15.FieldName = Nothing
        Me.MyLabel15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel15.Location = New System.Drawing.Point(308, 215)
        Me.MyLabel15.Name = "MyLabel15"
        Me.MyLabel15.Size = New System.Drawing.Size(82, 16)
        Me.MyLabel15.TabIndex = 607
        Me.MyLabel15.Text = "Date And Time"
        '
        'ChkTrnsferToSvng
        '
        Me.ChkTrnsferToSvng.Location = New System.Drawing.Point(544, 214)
        Me.ChkTrnsferToSvng.MyLinkLable1 = Nothing
        Me.ChkTrnsferToSvng.MyLinkLable2 = Nothing
        Me.ChkTrnsferToSvng.Name = "ChkTrnsferToSvng"
        Me.ChkTrnsferToSvng.Size = New System.Drawing.Size(112, 18)
        Me.ChkTrnsferToSvng.TabIndex = 615
        Me.ChkTrnsferToSvng.Tag1 = Nothing
        Me.ChkTrnsferToSvng.Text = "Transfer To Saving"
        Me.ChkTrnsferToSvng.Visible = False
        '
        'grpVendorBankDetails
        '
        Me.grpVendorBankDetails.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.grpVendorBankDetails.Controls.Add(Me.txtVendor_Bank_ACNo)
        Me.grpVendorBankDetails.Controls.Add(Me.MyLabel17)
        Me.grpVendorBankDetails.Controls.Add(Me.txtBankCode)
        Me.grpVendorBankDetails.Controls.Add(Me.TxtIFSCCode)
        Me.grpVendorBankDetails.Controls.Add(Me.MyLabel22)
        Me.grpVendorBankDetails.Controls.Add(Me.txtbranchname)
        Me.grpVendorBankDetails.Controls.Add(Me.TxtBankName)
        Me.grpVendorBankDetails.Controls.Add(Me.MyLabel23)
        Me.grpVendorBankDetails.HeaderText = "Vendor Bank Details"
        Me.grpVendorBankDetails.Location = New System.Drawing.Point(704, 174)
        Me.grpVendorBankDetails.Name = "grpVendorBankDetails"
        Me.grpVendorBankDetails.Size = New System.Drawing.Size(385, 84)
        Me.grpVendorBankDetails.TabIndex = 611
        Me.grpVendorBankDetails.Text = "Vendor Bank Details"
        '
        'txtVendor_Bank_ACNo
        '
        Me.txtVendor_Bank_ACNo.CalculationExpression = Nothing
        Me.txtVendor_Bank_ACNo.FieldCode = Nothing
        Me.txtVendor_Bank_ACNo.FieldDesc = Nothing
        Me.txtVendor_Bank_ACNo.FieldMaxLength = 0
        Me.txtVendor_Bank_ACNo.FieldName = Nothing
        Me.txtVendor_Bank_ACNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendor_Bank_ACNo.isCalculatedField = False
        Me.txtVendor_Bank_ACNo.IsSourceFromTable = False
        Me.txtVendor_Bank_ACNo.IsSourceFromValueList = False
        Me.txtVendor_Bank_ACNo.IsUnique = False
        Me.txtVendor_Bank_ACNo.Location = New System.Drawing.Point(93, 63)
        Me.txtVendor_Bank_ACNo.MaxLength = 50
        Me.txtVendor_Bank_ACNo.MendatroryField = False
        Me.txtVendor_Bank_ACNo.MyLinkLable1 = Me.MyLabel17
        Me.txtVendor_Bank_ACNo.MyLinkLable2 = Nothing
        Me.txtVendor_Bank_ACNo.Name = "txtVendor_Bank_ACNo"
        Me.txtVendor_Bank_ACNo.ReferenceFieldDesc = Nothing
        Me.txtVendor_Bank_ACNo.ReferenceFieldName = Nothing
        Me.txtVendor_Bank_ACNo.ReferenceTableName = Nothing
        Me.txtVendor_Bank_ACNo.Size = New System.Drawing.Size(311, 18)
        Me.txtVendor_Bank_ACNo.TabIndex = 123
        '
        'MyLabel17
        '
        Me.MyLabel17.FieldName = Nothing
        Me.MyLabel17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel17.Location = New System.Drawing.Point(29, 65)
        Me.MyLabel17.Name = "MyLabel17"
        Me.MyLabel17.Size = New System.Drawing.Size(65, 16)
        Me.MyLabel17.TabIndex = 124
        Me.MyLabel17.Text = "Account No"
        '
        'txtBankCode
        '
        Me.txtBankCode.CalculationExpression = Nothing
        Me.txtBankCode.FieldCode = Nothing
        Me.txtBankCode.FieldDesc = Nothing
        Me.txtBankCode.FieldMaxLength = 0
        Me.txtBankCode.FieldName = Nothing
        Me.txtBankCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBankCode.isCalculatedField = False
        Me.txtBankCode.IsSourceFromTable = False
        Me.txtBankCode.IsSourceFromValueList = False
        Me.txtBankCode.IsUnique = False
        Me.txtBankCode.Location = New System.Drawing.Point(92, 19)
        Me.txtBankCode.MaxLength = 50
        Me.txtBankCode.MendatroryField = False
        Me.txtBankCode.MyLinkLable1 = Me.MyLabel22
        Me.txtBankCode.MyLinkLable2 = Nothing
        Me.txtBankCode.Name = "txtBankCode"
        Me.txtBankCode.ReferenceFieldDesc = Nothing
        Me.txtBankCode.ReferenceFieldName = Nothing
        Me.txtBankCode.ReferenceTableName = Nothing
        Me.txtBankCode.Size = New System.Drawing.Size(130, 18)
        Me.txtBankCode.TabIndex = 122
        '
        'MyLabel22
        '
        Me.MyLabel22.FieldName = Nothing
        Me.MyLabel22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel22.Location = New System.Drawing.Point(28, 43)
        Me.MyLabel22.Name = "MyLabel22"
        Me.MyLabel22.Size = New System.Drawing.Size(62, 16)
        Me.MyLabel22.TabIndex = 29
        Me.MyLabel22.Text = "IFSC Code"
        '
        'TxtIFSCCode
        '
        Me.TxtIFSCCode.CalculationExpression = Nothing
        Me.TxtIFSCCode.FieldCode = Nothing
        Me.TxtIFSCCode.FieldDesc = Nothing
        Me.TxtIFSCCode.FieldMaxLength = 0
        Me.TxtIFSCCode.FieldName = Nothing
        Me.TxtIFSCCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtIFSCCode.isCalculatedField = False
        Me.TxtIFSCCode.IsSourceFromTable = False
        Me.TxtIFSCCode.IsSourceFromValueList = False
        Me.TxtIFSCCode.IsUnique = False
        Me.TxtIFSCCode.Location = New System.Drawing.Point(92, 41)
        Me.TxtIFSCCode.MaxLength = 50
        Me.TxtIFSCCode.MendatroryField = False
        Me.TxtIFSCCode.MyLinkLable1 = Me.MyLabel22
        Me.TxtIFSCCode.MyLinkLable2 = Nothing
        Me.TxtIFSCCode.Name = "TxtIFSCCode"
        Me.TxtIFSCCode.ReferenceFieldDesc = Nothing
        Me.TxtIFSCCode.ReferenceFieldName = Nothing
        Me.TxtIFSCCode.ReferenceTableName = Nothing
        Me.TxtIFSCCode.Size = New System.Drawing.Size(130, 18)
        Me.TxtIFSCCode.TabIndex = 5
        '
        'txtbranchname
        '
        Me.txtbranchname.CalculationExpression = Nothing
        Me.txtbranchname.FieldCode = Nothing
        Me.txtbranchname.FieldDesc = Nothing
        Me.txtbranchname.FieldMaxLength = 0
        Me.txtbranchname.FieldName = Nothing
        Me.txtbranchname.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtbranchname.isCalculatedField = False
        Me.txtbranchname.IsSourceFromTable = False
        Me.txtbranchname.IsSourceFromValueList = False
        Me.txtbranchname.IsUnique = False
        Me.txtbranchname.Location = New System.Drawing.Point(225, 41)
        Me.txtbranchname.MaxLength = 150
        Me.txtbranchname.MendatroryField = False
        Me.txtbranchname.MyLinkLable1 = Nothing
        Me.txtbranchname.MyLinkLable2 = Nothing
        Me.txtbranchname.Name = "txtbranchname"
        Me.txtbranchname.ReferenceFieldDesc = Nothing
        Me.txtbranchname.ReferenceFieldName = Nothing
        Me.txtbranchname.ReferenceTableName = Nothing
        Me.txtbranchname.Size = New System.Drawing.Size(179, 18)
        Me.txtbranchname.TabIndex = 4
        '
        'TxtBankName
        '
        Me.TxtBankName.CalculationExpression = Nothing
        Me.TxtBankName.FieldCode = Nothing
        Me.TxtBankName.FieldDesc = Nothing
        Me.TxtBankName.FieldMaxLength = 0
        Me.TxtBankName.FieldName = Nothing
        Me.TxtBankName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtBankName.isCalculatedField = False
        Me.TxtBankName.IsSourceFromTable = False
        Me.TxtBankName.IsSourceFromValueList = False
        Me.TxtBankName.IsUnique = False
        Me.TxtBankName.Location = New System.Drawing.Point(225, 19)
        Me.TxtBankName.MaxLength = 50
        Me.TxtBankName.MendatroryField = False
        Me.TxtBankName.MyLinkLable1 = Me.MyLabel23
        Me.TxtBankName.MyLinkLable2 = Nothing
        Me.TxtBankName.Name = "TxtBankName"
        Me.TxtBankName.ReferenceFieldDesc = Nothing
        Me.TxtBankName.ReferenceFieldName = Nothing
        Me.TxtBankName.ReferenceTableName = Nothing
        Me.TxtBankName.Size = New System.Drawing.Size(178, 18)
        Me.TxtBankName.TabIndex = 1
        Me.TxtBankName.TabStop = False
        '
        'MyLabel23
        '
        Me.MyLabel23.FieldName = Nothing
        Me.MyLabel23.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel23.Location = New System.Drawing.Point(28, 19)
        Me.MyLabel23.Name = "MyLabel23"
        Me.MyLabel23.Size = New System.Drawing.Size(62, 16)
        Me.MyLabel23.TabIndex = 26
        Me.MyLabel23.Text = "Bank Code"
        '
        'chkSecurityAdjusment
        '
        Me.chkSecurityAdjusment.Location = New System.Drawing.Point(736, 150)
        Me.chkSecurityAdjusment.MyLinkLable1 = Nothing
        Me.chkSecurityAdjusment.MyLinkLable2 = Nothing
        Me.chkSecurityAdjusment.Name = "chkSecurityAdjusment"
        Me.chkSecurityAdjusment.Size = New System.Drawing.Size(121, 18)
        Me.chkSecurityAdjusment.TabIndex = 614
        Me.chkSecurityAdjusment.Tag1 = Nothing
        Me.chkSecurityAdjusment.Text = "Security Adjustment"
        '
        'chkTDSProvision
        '
        Me.chkTDSProvision.Location = New System.Drawing.Point(645, 150)
        Me.chkTDSProvision.MyLinkLable1 = Nothing
        Me.chkTDSProvision.MyLinkLable2 = Nothing
        Me.chkTDSProvision.Name = "chkTDSProvision"
        Me.chkTDSProvision.ReadOnly = True
        Me.chkTDSProvision.Size = New System.Drawing.Size(89, 18)
        Me.chkTDSProvision.TabIndex = 613
        Me.chkTDSProvision.Tag1 = Nothing
        Me.chkTDSProvision.Text = "TDS Provision"
        '
        'chkRCM
        '
        Me.chkRCM.Location = New System.Drawing.Point(598, 150)
        Me.chkRCM.MyLinkLable1 = Nothing
        Me.chkRCM.MyLinkLable2 = Nothing
        Me.chkRCM.Name = "chkRCM"
        Me.chkRCM.ReadOnly = True
        Me.chkRCM.Size = New System.Drawing.Size(44, 18)
        Me.chkRCM.TabIndex = 612
        Me.chkRCM.Tag1 = Nothing
        Me.chkRCM.Text = "RCM"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.lblGstinNo)
        Me.Panel4.Controls.Add(Me.MyLabel49)
        Me.Panel4.Controls.Add(Me.MyLabel45)
        Me.Panel4.Controls.Add(Me.lblRegisterOrUnregister)
        Me.Panel4.Location = New System.Drawing.Point(430, 169)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(272, 46)
        Me.Panel4.TabIndex = 603
        '
        'lblGstinNo
        '
        Me.lblGstinNo.AutoSize = False
        Me.lblGstinNo.BorderVisible = True
        Me.lblGstinNo.FieldName = Nothing
        Me.lblGstinNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGstinNo.Location = New System.Drawing.Point(101, 25)
        Me.lblGstinNo.Name = "lblGstinNo"
        Me.lblGstinNo.Size = New System.Drawing.Size(168, 19)
        Me.lblGstinNo.TabIndex = 356
        Me.lblGstinNo.TextWrap = False
        '
        'MyLabel49
        '
        Me.MyLabel49.FieldName = Nothing
        Me.MyLabel49.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel49.Location = New System.Drawing.Point(3, 28)
        Me.MyLabel49.Name = "MyLabel49"
        Me.MyLabel49.Size = New System.Drawing.Size(98, 16)
        Me.MyLabel49.TabIndex = 355
        Me.MyLabel49.Text = "Vendor GSTIN No"
        '
        'MyLabel45
        '
        Me.MyLabel45.FieldName = Nothing
        Me.MyLabel45.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel45.Location = New System.Drawing.Point(3, 4)
        Me.MyLabel45.Name = "MyLabel45"
        Me.MyLabel45.Size = New System.Drawing.Size(46, 16)
        Me.MyLabel45.TabIndex = 76
        Me.MyLabel45.Text = "Vendor "
        '
        'lblRegisterOrUnregister
        '
        Me.lblRegisterOrUnregister.AutoSize = False
        Me.lblRegisterOrUnregister.BorderVisible = True
        Me.lblRegisterOrUnregister.FieldName = Nothing
        Me.lblRegisterOrUnregister.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRegisterOrUnregister.Location = New System.Drawing.Point(101, 4)
        Me.lblRegisterOrUnregister.Name = "lblRegisterOrUnregister"
        Me.lblRegisterOrUnregister.Size = New System.Drawing.Size(168, 19)
        Me.lblRegisterOrUnregister.TabIndex = 354
        Me.lblRegisterOrUnregister.TextWrap = False
        '
        'RadGroupBox4
        '
        Me.RadGroupBox4.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox4.Controls.Add(Me.CboxITCCateogory)
        Me.RadGroupBox4.Controls.Add(Me.MyLabel12)
        Me.RadGroupBox4.Controls.Add(Me.CboxITCType)
        Me.RadGroupBox4.Enabled = False
        Me.RadGroupBox4.HeaderText = "ITC Eligible"
        Me.RadGroupBox4.Location = New System.Drawing.Point(111, 168)
        Me.RadGroupBox4.Name = "RadGroupBox4"
        Me.RadGroupBox4.Size = New System.Drawing.Size(314, 44)
        Me.RadGroupBox4.TabIndex = 152
        Me.RadGroupBox4.Text = "ITC Eligible"
        '
        'CboxITCCateogory
        '
        Me.CboxITCCateogory.AutoCompleteDisplayMember = Nothing
        Me.CboxITCCateogory.AutoCompleteValueMember = Nothing
        Me.CboxITCCateogory.CalculationExpression = Nothing
        Me.CboxITCCateogory.DropDownAnimationEnabled = True
        Me.CboxITCCateogory.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.CboxITCCateogory.FieldCode = Nothing
        Me.CboxITCCateogory.FieldDesc = Nothing
        Me.CboxITCCateogory.FieldMaxLength = 0
        Me.CboxITCCateogory.FieldName = Nothing
        Me.CboxITCCateogory.isCalculatedField = False
        Me.CboxITCCateogory.IsSourceFromTable = False
        Me.CboxITCCateogory.IsSourceFromValueList = False
        Me.CboxITCCateogory.IsUnique = False
        RadListDataItem1.Text = "Yes"
        RadListDataItem2.Text = "No"
        Me.CboxITCCateogory.Items.Add(RadListDataItem1)
        Me.CboxITCCateogory.Items.Add(RadListDataItem2)
        Me.CboxITCCateogory.Location = New System.Drawing.Point(90, 19)
        Me.CboxITCCateogory.MendatroryField = False
        Me.CboxITCCateogory.MyLinkLable1 = Me.MyLabel12
        Me.CboxITCCateogory.MyLinkLable2 = Nothing
        Me.CboxITCCateogory.Name = "CboxITCCateogory"
        Me.CboxITCCateogory.ReferenceFieldDesc = Nothing
        Me.CboxITCCateogory.ReferenceFieldName = Nothing
        Me.CboxITCCateogory.ReferenceTableName = Nothing
        Me.CboxITCCateogory.Size = New System.Drawing.Size(222, 20)
        Me.CboxITCCateogory.TabIndex = 45
        '
        'MyLabel12
        '
        Me.MyLabel12.FieldName = Nothing
        Me.MyLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel12.Location = New System.Drawing.Point(4, 21)
        Me.MyLabel12.Name = "MyLabel12"
        Me.MyLabel12.Size = New System.Drawing.Size(31, 16)
        Me.MyLabel12.TabIndex = 44
        Me.MyLabel12.Text = "Type"
        '
        'CboxITCType
        '
        Me.CboxITCType.AutoCompleteDisplayMember = Nothing
        Me.CboxITCType.AutoCompleteValueMember = Nothing
        Me.CboxITCType.CalculationExpression = Nothing
        Me.CboxITCType.DropDownAnimationEnabled = True
        Me.CboxITCType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.CboxITCType.FieldCode = Nothing
        Me.CboxITCType.FieldDesc = Nothing
        Me.CboxITCType.FieldMaxLength = 0
        Me.CboxITCType.FieldName = Nothing
        Me.CboxITCType.isCalculatedField = False
        Me.CboxITCType.IsSourceFromTable = False
        Me.CboxITCType.IsSourceFromValueList = False
        Me.CboxITCType.IsUnique = False
        RadListDataItem3.Text = "Yes"
        RadListDataItem4.Text = "No"
        Me.CboxITCType.Items.Add(RadListDataItem3)
        Me.CboxITCType.Items.Add(RadListDataItem4)
        Me.CboxITCType.Location = New System.Drawing.Point(37, 19)
        Me.CboxITCType.MendatroryField = True
        Me.CboxITCType.MyLinkLable1 = Me.MyLabel12
        Me.CboxITCType.MyLinkLable2 = Nothing
        Me.CboxITCType.Name = "CboxITCType"
        Me.CboxITCType.ReferenceFieldDesc = Nothing
        Me.CboxITCType.ReferenceFieldName = Nothing
        Me.CboxITCType.ReferenceTableName = Nothing
        Me.CboxITCType.Size = New System.Drawing.Size(52, 20)
        Me.CboxITCType.TabIndex = 43
        '
        'chkITCEligible
        '
        Me.chkITCEligible.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkITCEligible.Location = New System.Drawing.Point(5, 184)
        Me.chkITCEligible.Name = "chkITCEligible"
        '
        '
        '
        Me.chkITCEligible.RootElement.StretchHorizontally = True
        Me.chkITCEligible.RootElement.StretchVertically = True
        Me.chkITCEligible.Size = New System.Drawing.Size(93, 15)
        Me.chkITCEligible.TabIndex = 151
        Me.chkITCEligible.Text = "ITC Eligible"
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
        Me.RadGroupBox2.Location = New System.Drawing.Point(2, 279)
        Me.RadGroupBox2.Name = "RadGroupBox2"
        Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox2.Size = New System.Drawing.Size(1100, 153)
        Me.RadGroupBox2.TabIndex = 19
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
        Me.gv1.MasterTemplate.ShowFilteringRow = False
        Me.gv1.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv1.MasterTemplate.ViewDefinition = TableViewDefinition1
        Me.gv1.MyExportAPI = False
        Me.gv1.MyExportFilePath = ""
        Me.gv1.MyStopExport = False
        Me.gv1.Name = "gv1"
        Me.gv1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gv1.ShowGroupPanel = False
        Me.gv1.ShowHeaderCellButtons = True
        Me.gv1.Size = New System.Drawing.Size(1080, 123)
        Me.gv1.TabIndex = 0
        Me.gv1.TabStop = False
        Me.gv1.VarID = ""
        '
        'pnlDeduction
        '
        Me.pnlDeduction.Controls.Add(Me.pnlSecondaryTranporterPaymentCycle)
        Me.pnlDeduction.Controls.Add(Me.ChkSecurity)
        Me.pnlDeduction.Controls.Add(Me.chkSecTrans)
        Me.pnlDeduction.Location = New System.Drawing.Point(678, 107)
        Me.pnlDeduction.Name = "pnlDeduction"
        Me.pnlDeduction.Size = New System.Drawing.Size(241, 43)
        Me.pnlDeduction.TabIndex = 57
        Me.pnlDeduction.Visible = False
        '
        'pnlSecondaryTranporterPaymentCycle
        '
        Me.pnlSecondaryTranporterPaymentCycle.Controls.Add(Me.MyLabel11)
        Me.pnlSecondaryTranporterPaymentCycle.Controls.Add(Me.MyLabel10)
        Me.pnlSecondaryTranporterPaymentCycle.Controls.Add(Me.txtToDateSecTrans)
        Me.pnlSecondaryTranporterPaymentCycle.Controls.Add(Me.txtFromDateSecTrans)
        Me.pnlSecondaryTranporterPaymentCycle.Location = New System.Drawing.Point(1, 21)
        Me.pnlSecondaryTranporterPaymentCycle.Name = "pnlSecondaryTranporterPaymentCycle"
        Me.pnlSecondaryTranporterPaymentCycle.Size = New System.Drawing.Size(235, 20)
        Me.pnlSecondaryTranporterPaymentCycle.TabIndex = 58
        '
        'MyLabel11
        '
        Me.MyLabel11.FieldName = Nothing
        Me.MyLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel11.Location = New System.Drawing.Point(120, 3)
        Me.MyLabel11.Name = "MyLabel11"
        Me.MyLabel11.Size = New System.Drawing.Size(37, 16)
        Me.MyLabel11.TabIndex = 60
        Me.MyLabel11.Text = "TDate"
        '
        'MyLabel10
        '
        Me.MyLabel10.FieldName = Nothing
        Me.MyLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.Location = New System.Drawing.Point(5, 3)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(37, 16)
        Me.MyLabel10.TabIndex = 58
        Me.MyLabel10.Text = "FDate"
        '
        'txtToDateSecTrans
        '
        Me.txtToDateSecTrans.CalculationExpression = Nothing
        Me.txtToDateSecTrans.CustomFormat = "dd/MM/yyyy"
        Me.txtToDateSecTrans.FieldCode = Nothing
        Me.txtToDateSecTrans.FieldDesc = Nothing
        Me.txtToDateSecTrans.FieldMaxLength = 0
        Me.txtToDateSecTrans.FieldName = Nothing
        Me.txtToDateSecTrans.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtToDateSecTrans.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtToDateSecTrans.isCalculatedField = False
        Me.txtToDateSecTrans.IsSourceFromTable = False
        Me.txtToDateSecTrans.IsSourceFromValueList = False
        Me.txtToDateSecTrans.IsUnique = False
        Me.txtToDateSecTrans.Location = New System.Drawing.Point(157, 2)
        Me.txtToDateSecTrans.MendatroryField = False
        Me.txtToDateSecTrans.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDateSecTrans.MyLinkLable1 = Me.MyLabel11
        Me.txtToDateSecTrans.MyLinkLable2 = Nothing
        Me.txtToDateSecTrans.Name = "txtToDateSecTrans"
        Me.txtToDateSecTrans.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtToDateSecTrans.ReferenceFieldDesc = Nothing
        Me.txtToDateSecTrans.ReferenceFieldName = Nothing
        Me.txtToDateSecTrans.ReferenceTableName = Nothing
        Me.txtToDateSecTrans.Size = New System.Drawing.Size(76, 18)
        Me.txtToDateSecTrans.TabIndex = 59
        Me.txtToDateSecTrans.TabStop = False
        Me.txtToDateSecTrans.Text = "13/06/2011"
        Me.txtToDateSecTrans.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'txtFromDateSecTrans
        '
        Me.txtFromDateSecTrans.CalculationExpression = Nothing
        Me.txtFromDateSecTrans.CustomFormat = "dd/MM/yyyy"
        Me.txtFromDateSecTrans.FieldCode = Nothing
        Me.txtFromDateSecTrans.FieldDesc = Nothing
        Me.txtFromDateSecTrans.FieldMaxLength = 0
        Me.txtFromDateSecTrans.FieldName = Nothing
        Me.txtFromDateSecTrans.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFromDateSecTrans.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtFromDateSecTrans.isCalculatedField = False
        Me.txtFromDateSecTrans.IsSourceFromTable = False
        Me.txtFromDateSecTrans.IsSourceFromValueList = False
        Me.txtFromDateSecTrans.IsUnique = False
        Me.txtFromDateSecTrans.Location = New System.Drawing.Point(43, 2)
        Me.txtFromDateSecTrans.MendatroryField = False
        Me.txtFromDateSecTrans.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtFromDateSecTrans.MyLinkLable1 = Me.MyLabel10
        Me.txtFromDateSecTrans.MyLinkLable2 = Nothing
        Me.txtFromDateSecTrans.Name = "txtFromDateSecTrans"
        Me.txtFromDateSecTrans.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtFromDateSecTrans.ReferenceFieldDesc = Nothing
        Me.txtFromDateSecTrans.ReferenceFieldName = Nothing
        Me.txtFromDateSecTrans.ReferenceTableName = Nothing
        Me.txtFromDateSecTrans.Size = New System.Drawing.Size(76, 18)
        Me.txtFromDateSecTrans.TabIndex = 57
        Me.txtFromDateSecTrans.TabStop = False
        Me.txtFromDateSecTrans.Text = "13/06/2011"
        Me.txtFromDateSecTrans.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'ChkSecurity
        '
        Me.ChkSecurity.Location = New System.Drawing.Point(4, 1)
        Me.ChkSecurity.MyLinkLable1 = Nothing
        Me.ChkSecurity.MyLinkLable2 = Nothing
        Me.ChkSecurity.Name = "ChkSecurity"
        Me.ChkSecurity.Size = New System.Drawing.Size(60, 18)
        Me.ChkSecurity.TabIndex = 43
        Me.ChkSecurity.Tag1 = Nothing
        Me.ChkSecurity.Text = "Security"
        Me.ChkSecurity.Visible = False
        '
        'chkSecTrans
        '
        Me.chkSecTrans.Location = New System.Drawing.Point(63, 1)
        Me.chkSecTrans.MyLinkLable1 = Nothing
        Me.chkSecTrans.MyLinkLable2 = Nothing
        Me.chkSecTrans.Name = "chkSecTrans"
        Me.chkSecTrans.Size = New System.Drawing.Size(161, 18)
        Me.chkSecTrans.TabIndex = 56
        Me.chkSecTrans.Tag1 = Nothing
        Me.chkSecTrans.Text = "Secondary Transporter ( F5 )"
        '
        'RadLabel21
        '
        Me.RadLabel21.Controls.Add(Me.LblCostCentre)
        Me.RadLabel21.Location = New System.Drawing.Point(849, 66)
        Me.RadLabel21.Name = "RadLabel21"
        Me.RadLabel21.Size = New System.Drawing.Size(65, 18)
        Me.RadLabel21.TabIndex = 51
        Me.RadLabel21.Text = "Cost Centre"
        Me.RadLabel21.Visible = False
        '
        'LblCostCentre
        '
        Me.LblCostCentre.AutoSize = False
        Me.LblCostCentre.BorderVisible = True
        Me.LblCostCentre.FieldName = Nothing
        Me.LblCostCentre.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblCostCentre.Location = New System.Drawing.Point(65, 1)
        Me.LblCostCentre.Name = "LblCostCentre"
        Me.LblCostCentre.Size = New System.Drawing.Size(18, 18)
        Me.LblCostCentre.TabIndex = 52
        Me.LblCostCentre.TextWrap = False
        Me.LblCostCentre.Visible = False
        '
        'RadLabel20
        '
        Me.RadLabel20.Location = New System.Drawing.Point(808, 67)
        Me.RadLabel20.Name = "RadLabel20"
        Me.RadLabel20.Size = New System.Drawing.Size(82, 18)
        Me.RadLabel20.TabIndex = 48
        Me.RadLabel20.Text = "Hirerachy Level"
        Me.RadLabel20.Visible = False
        '
        'pnlPCJ
        '
        Me.pnlPCJ.Controls.Add(Me.MyLabel4)
        Me.pnlPCJ.Controls.Add(Me.fndProject)
        Me.pnlPCJ.Controls.Add(Me.lblProject)
        Me.pnlPCJ.Location = New System.Drawing.Point(3, 126)
        Me.pnlPCJ.Name = "pnlPCJ"
        Me.pnlPCJ.Size = New System.Drawing.Size(646, 22)
        Me.pnlPCJ.TabIndex = 18
        '
        'MyLabel4
        '
        Me.MyLabel4.FieldName = Nothing
        Me.MyLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(2, 1)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(41, 16)
        Me.MyLabel4.TabIndex = 0
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
        Me.fndProject.Location = New System.Drawing.Point(110, 2)
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
        Me.lblProject.Location = New System.Drawing.Point(261, 2)
        Me.lblProject.Name = "lblProject"
        Me.lblProject.Size = New System.Drawing.Size(407, 18)
        Me.lblProject.TabIndex = 1
        Me.lblProject.TextWrap = False
        '
        'RadLabel18
        '
        Me.RadLabel18.Location = New System.Drawing.Point(3, 25)
        Me.RadLabel18.Name = "RadLabel18"
        Me.RadLabel18.Size = New System.Drawing.Size(96, 18)
        Me.RadLabel18.TabIndex = 31
        Me.RadLabel18.Text = "Location Segment"
        '
        'btnDrillDown
        '
        Me.btnDrillDown.Image = CType(resources.GetObject("btnDrillDown.Image"), System.Drawing.Image)
        Me.btnDrillDown.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnDrillDown.Location = New System.Drawing.Point(376, 1)
        Me.btnDrillDown.Name = "btnDrillDown"
        Me.btnDrillDown.Size = New System.Drawing.Size(20, 20)
        Me.btnDrillDown.TabIndex = 2
        '
        'chkOnHold
        '
        Me.chkOnHold.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOnHold.Location = New System.Drawing.Point(678, 3)
        Me.chkOnHold.Name = "chkOnHold"
        Me.chkOnHold.Size = New System.Drawing.Size(62, 16)
        Me.chkOnHold.TabIndex = 3
        Me.chkOnHold.Text = "On Hold"
        '
        'btnAddNew
        '
        Me.btnAddNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Image = Global.ERP.My.Resources.Resources._new
        Me.btnAddNew.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnAddNew.Location = New System.Drawing.Point(357, 1)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(19, 20)
        Me.btnAddNew.TabIndex = 1
        '
        'txtMCC
        '
        Me.txtMCC.BackColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(187, Byte), Integer))
        Me.txtMCC.CalculationExpression = Nothing
        Me.txtMCC.Enabled = False
        Me.txtMCC.FieldCode = Nothing
        Me.txtMCC.FieldDesc = Nothing
        Me.txtMCC.FieldMaxLength = 0
        Me.txtMCC.FieldName = Nothing
        Me.txtMCC.isCalculatedField = False
        Me.txtMCC.IsSourceFromTable = False
        Me.txtMCC.IsSourceFromValueList = False
        Me.txtMCC.IsUnique = False
        Me.txtMCC.Location = New System.Drawing.Point(111, 235)
        Me.txtMCC.MendatroryField = False
        Me.txtMCC.MyLinkLable1 = Nothing
        Me.txtMCC.MyLinkLable2 = Nothing
        Me.txtMCC.Name = "txtMCC"
        Me.txtMCC.ReferenceFieldDesc = Nothing
        Me.txtMCC.ReferenceFieldName = Nothing
        Me.txtMCC.ReferenceTableName = Nothing
        Me.txtMCC.Size = New System.Drawing.Size(88, 20)
        Me.txtMCC.TabIndex = 610
        '
        'lblMCC
        '
        Me.lblMCC.BackColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(187, Byte), Integer), CType(CType(187, Byte), Integer))
        Me.lblMCC.CalculationExpression = Nothing
        Me.lblMCC.Enabled = False
        Me.lblMCC.FieldCode = Nothing
        Me.lblMCC.FieldDesc = Nothing
        Me.lblMCC.FieldMaxLength = 0
        Me.lblMCC.FieldName = Nothing
        Me.lblMCC.isCalculatedField = False
        Me.lblMCC.IsSourceFromTable = False
        Me.lblMCC.IsSourceFromValueList = False
        Me.lblMCC.IsUnique = False
        Me.lblMCC.Location = New System.Drawing.Point(200, 235)
        Me.lblMCC.MendatroryField = False
        Me.lblMCC.MyLinkLable1 = Nothing
        Me.lblMCC.MyLinkLable2 = Nothing
        Me.lblMCC.Name = "lblMCC"
        Me.lblMCC.ReferenceFieldDesc = Nothing
        Me.lblMCC.ReferenceFieldName = Nothing
        Me.lblMCC.ReferenceTableName = Nothing
        Me.lblMCC.Size = New System.Drawing.Size(140, 20)
        Me.lblMCC.TabIndex = 608
        '
        'lblMCC2
        '
        Me.lblMCC2.FieldName = Nothing
        Me.lblMCC2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMCC2.Location = New System.Drawing.Point(5, 237)
        Me.lblMCC2.Name = "lblMCC2"
        Me.lblMCC2.Size = New System.Drawing.Size(72, 16)
        Me.lblMCC2.TabIndex = 609
        Me.lblMCC2.Text = "MCC/PLANT"
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
        Me.txtDataAndTimeSelection.Location = New System.Drawing.Point(395, 214)
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
        Me.txtDataAndTimeSelection.TabIndex = 606
        Me.txtDataAndTimeSelection.TabStop = False
        Me.txtDataAndTimeSelection.Text = "13/06/2011 11:29 AM"
        Me.txtDataAndTimeSelection.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'MyLabel16
        '
        Me.MyLabel16.FieldName = Nothing
        Me.MyLabel16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel16.Location = New System.Drawing.Point(5, 215)
        Me.MyLabel16.Name = "MyLabel16"
        Me.MyLabel16.Size = New System.Drawing.Size(52, 16)
        Me.MyLabel16.TabIndex = 605
        Me.MyLabel16.Text = "Tapal No"
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
        Me.txtTapalNo.Location = New System.Drawing.Point(111, 214)
        Me.txtTapalNo.MendatroryField = False
        Me.txtTapalNo.MyLinkLable1 = Nothing
        Me.txtTapalNo.MyLinkLable2 = Nothing
        Me.txtTapalNo.Name = "txtTapalNo"
        Me.txtTapalNo.ReferenceFieldDesc = Nothing
        Me.txtTapalNo.ReferenceFieldName = Nothing
        Me.txtTapalNo.ReferenceTableName = Nothing
        Me.txtTapalNo.Size = New System.Drawing.Size(191, 18)
        Me.txtTapalNo.TabIndex = 604
        '
        'MyLabel1
        '
        Me.MyLabel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.MyLabel1.Location = New System.Drawing.Point(783, 433)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(318, 16)
        Me.MyLabel1.TabIndex = 23
        Me.MyLabel1.Text = "Double click on Tax Amount Column To Set Item wise Tax"
        Me.MyLabel1.Visible = False
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(343, 237)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(132, 16)
        Me.MyLabel13.TabIndex = 154
        Me.MyLabel13.Text = "Employee Expense Type"
        '
        'ddlEmployeeType
        '
        Me.ddlEmployeeType.AutoCompleteDisplayMember = Nothing
        Me.ddlEmployeeType.AutoCompleteValueMember = Nothing
        Me.ddlEmployeeType.CalculationExpression = Nothing
        Me.ddlEmployeeType.DropDownAnimationEnabled = True
        Me.ddlEmployeeType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.ddlEmployeeType.FieldCode = Nothing
        Me.ddlEmployeeType.FieldDesc = Nothing
        Me.ddlEmployeeType.FieldMaxLength = 0
        Me.ddlEmployeeType.FieldName = Nothing
        Me.ddlEmployeeType.isCalculatedField = False
        Me.ddlEmployeeType.IsSourceFromTable = False
        Me.ddlEmployeeType.IsSourceFromValueList = False
        Me.ddlEmployeeType.IsUnique = False
        RadListDataItem5.Text = "Yes"
        RadListDataItem6.Text = "No"
        Me.ddlEmployeeType.Items.Add(RadListDataItem5)
        Me.ddlEmployeeType.Items.Add(RadListDataItem6)
        Me.ddlEmployeeType.Location = New System.Drawing.Point(481, 235)
        Me.ddlEmployeeType.MendatroryField = True
        Me.ddlEmployeeType.MyLinkLable1 = Me.MyLabel13
        Me.ddlEmployeeType.MyLinkLable2 = Nothing
        Me.ddlEmployeeType.Name = "ddlEmployeeType"
        Me.ddlEmployeeType.ReferenceFieldDesc = Nothing
        Me.ddlEmployeeType.ReferenceFieldName = Nothing
        Me.ddlEmployeeType.ReferenceTableName = Nothing
        Me.ddlEmployeeType.Size = New System.Drawing.Size(132, 20)
        Me.ddlEmployeeType.TabIndex = 153
        '
        'txtTotalAmt
        '
        Me.txtTotalAmt.BackColor = System.Drawing.Color.White
        Me.txtTotalAmt.CalculationExpression = Nothing
        Me.txtTotalAmt.DecimalPlaces = 2
        Me.txtTotalAmt.FieldCode = Nothing
        Me.txtTotalAmt.FieldDesc = Nothing
        Me.txtTotalAmt.FieldMaxLength = 0
        Me.txtTotalAmt.FieldName = Nothing
        Me.txtTotalAmt.isCalculatedField = False
        Me.txtTotalAmt.IsSourceFromTable = False
        Me.txtTotalAmt.IsSourceFromValueList = False
        Me.txtTotalAmt.IsUnique = False
        Me.txtTotalAmt.Location = New System.Drawing.Point(679, 85)
        Me.txtTotalAmt.MendatroryField = False
        Me.txtTotalAmt.MyLinkLable1 = Nothing
        Me.txtTotalAmt.MyLinkLable2 = Nothing
        Me.txtTotalAmt.Name = "txtTotalAmt"
        Me.txtTotalAmt.ReferenceFieldDesc = Nothing
        Me.txtTotalAmt.ReferenceFieldName = Nothing
        Me.txtTotalAmt.ReferenceTableName = Nothing
        Me.txtTotalAmt.Size = New System.Drawing.Size(98, 20)
        Me.txtTotalAmt.TabIndex = 62
        Me.txtTotalAmt.Text = "0"
        Me.txtTotalAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTotalAmt.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'lblcaption2
        '
        Me.lblcaption2.FieldName = Nothing
        Me.lblcaption2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblcaption2.Location = New System.Drawing.Point(234, 151)
        Me.lblcaption2.Name = "lblcaption2"
        Me.lblcaption2.Size = New System.Drawing.Size(89, 16)
        Me.lblcaption2.TabIndex = 61
        Me.lblcaption2.Text = "Service Bill Date"
        Me.lblcaption2.Visible = False
        '
        'dtBillDate
        '
        Me.dtBillDate.CalculationExpression = Nothing
        Me.dtBillDate.CustomFormat = "dd/MM/yyyy"
        Me.dtBillDate.FieldCode = Nothing
        Me.dtBillDate.FieldDesc = Nothing
        Me.dtBillDate.FieldMaxLength = 0
        Me.dtBillDate.FieldName = Nothing
        Me.dtBillDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtBillDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtBillDate.isCalculatedField = False
        Me.dtBillDate.IsSourceFromTable = False
        Me.dtBillDate.IsSourceFromValueList = False
        Me.dtBillDate.IsUnique = False
        Me.dtBillDate.Location = New System.Drawing.Point(325, 150)
        Me.dtBillDate.MendatroryField = False
        Me.dtBillDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.dtBillDate.MyLinkLable1 = Me.lblcaption2
        Me.dtBillDate.MyLinkLable2 = Nothing
        Me.dtBillDate.Name = "dtBillDate"
        Me.dtBillDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.dtBillDate.ReferenceFieldDesc = Nothing
        Me.dtBillDate.ReferenceFieldName = Nothing
        Me.dtBillDate.ReferenceTableName = Nothing
        Me.dtBillDate.Size = New System.Drawing.Size(114, 18)
        Me.dtBillDate.TabIndex = 60
        Me.dtBillDate.TabStop = False
        Me.dtBillDate.Text = "13/06/2011"
        Me.dtBillDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        Me.dtBillDate.Visible = False
        '
        'lblCaption1
        '
        Me.lblCaption1.FieldName = Nothing
        Me.lblCaption1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCaption1.Location = New System.Drawing.Point(5, 151)
        Me.lblCaption1.Name = "lblCaption1"
        Me.lblCaption1.Size = New System.Drawing.Size(80, 16)
        Me.lblCaption1.TabIndex = 59
        Me.lblCaption1.Text = "Service Bill No"
        Me.lblCaption1.Visible = False
        '
        'txtBillNo
        '
        Me.txtBillNo.CalculationExpression = Nothing
        Me.txtBillNo.FieldCode = Nothing
        Me.txtBillNo.FieldDesc = Nothing
        Me.txtBillNo.FieldMaxLength = 0
        Me.txtBillNo.FieldName = Nothing
        Me.txtBillNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBillNo.isCalculatedField = False
        Me.txtBillNo.IsSourceFromTable = False
        Me.txtBillNo.IsSourceFromValueList = False
        Me.txtBillNo.IsUnique = False
        Me.txtBillNo.Location = New System.Drawing.Point(114, 150)
        Me.txtBillNo.MaxLength = 30
        Me.txtBillNo.MendatroryField = False
        Me.txtBillNo.MyLinkLable1 = Me.lblCaption1
        Me.txtBillNo.MyLinkLable2 = Nothing
        Me.txtBillNo.Name = "txtBillNo"
        Me.txtBillNo.ReferenceFieldDesc = Nothing
        Me.txtBillNo.ReferenceFieldName = Nothing
        Me.txtBillNo.ReferenceTableName = Nothing
        Me.txtBillNo.Size = New System.Drawing.Size(114, 18)
        Me.txtBillNo.TabIndex = 58
        Me.txtBillNo.Visible = False
        '
        'chkProRated
        '
        Me.chkProRated.Location = New System.Drawing.Point(523, 150)
        Me.chkProRated.MyLinkLable1 = Nothing
        Me.chkProRated.MyLinkLable2 = Nothing
        Me.chkProRated.Name = "chkProRated"
        Me.chkProRated.Size = New System.Drawing.Size(69, 18)
        Me.chkProRated.TabIndex = 17
        Me.chkProRated.Tag1 = Nothing
        Me.chkProRated.Text = "Pro Rated"
        '
        'MyLabel9
        '
        Me.MyLabel9.FieldName = Nothing
        Me.MyLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel9.Location = New System.Drawing.Point(871, 26)
        Me.MyLabel9.Name = "MyLabel9"
        Me.MyLabel9.Size = New System.Drawing.Size(31, 16)
        Me.MyLabel9.TabIndex = 55
        Me.MyLabel9.Text = "Type"
        '
        'txtAdd_Doc_TYpe
        '
        Me.txtAdd_Doc_TYpe.AutoSize = False
        Me.txtAdd_Doc_TYpe.BackColor = System.Drawing.Color.Transparent
        Me.txtAdd_Doc_TYpe.BorderVisible = True
        Me.txtAdd_Doc_TYpe.FieldName = Nothing
        Me.txtAdd_Doc_TYpe.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdd_Doc_TYpe.Location = New System.Drawing.Point(911, 23)
        Me.txtAdd_Doc_TYpe.Name = "txtAdd_Doc_TYpe"
        Me.txtAdd_Doc_TYpe.Size = New System.Drawing.Size(185, 20)
        Me.txtAdd_Doc_TYpe.TabIndex = 54
        Me.txtAdd_Doc_TYpe.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'LblVCGL
        '
        Me.LblVCGL.CalculationExpression = Nothing
        Me.LblVCGL.FieldCode = Nothing
        Me.LblVCGL.FieldDesc = Nothing
        Me.LblVCGL.FieldMaxLength = 0
        Me.LblVCGL.FieldName = Nothing
        Me.LblVCGL.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblVCGL.isCalculatedField = False
        Me.LblVCGL.IsSourceFromTable = False
        Me.LblVCGL.IsSourceFromValueList = False
        Me.LblVCGL.IsUnique = False
        Me.LblVCGL.Location = New System.Drawing.Point(911, 3)
        Me.LblVCGL.MendatroryField = False
        Me.LblVCGL.MyLinkLable1 = Nothing
        Me.LblVCGL.MyLinkLable2 = Nothing
        Me.LblVCGL.Name = "LblVCGL"
        Me.LblVCGL.ReadOnly = True
        Me.LblVCGL.ReferenceFieldDesc = Nothing
        Me.LblVCGL.ReferenceFieldName = Nothing
        Me.LblVCGL.ReferenceTableName = Nothing
        Me.LblVCGL.Size = New System.Drawing.Size(185, 18)
        Me.LblVCGL.TabIndex = 53
        '
        'TxtCostCentre
        '
        Me.TxtCostCentre.CalculationExpression = Nothing
        Me.TxtCostCentre.FieldCode = Nothing
        Me.TxtCostCentre.FieldDesc = Nothing
        Me.TxtCostCentre.FieldMaxLength = 0
        Me.TxtCostCentre.FieldName = Nothing
        Me.TxtCostCentre.isCalculatedField = False
        Me.TxtCostCentre.IsSourceFromTable = False
        Me.TxtCostCentre.IsSourceFromValueList = False
        Me.TxtCostCentre.IsUnique = False
        Me.TxtCostCentre.Location = New System.Drawing.Point(903, 48)
        Me.TxtCostCentre.MendatroryField = True
        Me.TxtCostCentre.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtCostCentre.MyLinkLable1 = Nothing
        Me.TxtCostCentre.MyLinkLable2 = Nothing
        Me.TxtCostCentre.MyReadOnly = False
        Me.TxtCostCentre.MyShowMasterFormButton = False
        Me.TxtCostCentre.Name = "TxtCostCentre"
        Me.TxtCostCentre.ReferenceFieldDesc = Nothing
        Me.TxtCostCentre.ReferenceFieldName = Nothing
        Me.TxtCostCentre.ReferenceTableName = Nothing
        Me.TxtCostCentre.Size = New System.Drawing.Size(31, 16)
        Me.TxtCostCentre.TabIndex = 50
        Me.TxtCostCentre.Value = ""
        Me.TxtCostCentre.Visible = False
        '
        'LblHirerachy
        '
        Me.LblHirerachy.AutoSize = False
        Me.LblHirerachy.BorderVisible = True
        Me.LblHirerachy.FieldName = Nothing
        Me.LblHirerachy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblHirerachy.Location = New System.Drawing.Point(918, 67)
        Me.LblHirerachy.Name = "LblHirerachy"
        Me.LblHirerachy.Size = New System.Drawing.Size(16, 18)
        Me.LblHirerachy.TabIndex = 49
        Me.LblHirerachy.TextWrap = False
        Me.LblHirerachy.Visible = False
        '
        'TxtHirerachy
        '
        Me.TxtHirerachy.CalculationExpression = Nothing
        Me.TxtHirerachy.FieldCode = Nothing
        Me.TxtHirerachy.FieldDesc = Nothing
        Me.TxtHirerachy.FieldMaxLength = 0
        Me.TxtHirerachy.FieldName = Nothing
        Me.TxtHirerachy.isCalculatedField = False
        Me.TxtHirerachy.IsSourceFromTable = False
        Me.TxtHirerachy.IsSourceFromValueList = False
        Me.TxtHirerachy.IsUnique = False
        Me.TxtHirerachy.Location = New System.Drawing.Point(868, 47)
        Me.TxtHirerachy.MendatroryField = True
        Me.TxtHirerachy.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtHirerachy.MyLinkLable1 = Nothing
        Me.TxtHirerachy.MyLinkLable2 = Nothing
        Me.TxtHirerachy.MyReadOnly = False
        Me.TxtHirerachy.MyShowMasterFormButton = False
        Me.TxtHirerachy.Name = "TxtHirerachy"
        Me.TxtHirerachy.ReferenceFieldDesc = Nothing
        Me.TxtHirerachy.ReferenceFieldName = Nothing
        Me.TxtHirerachy.ReferenceTableName = Nothing
        Me.TxtHirerachy.Size = New System.Drawing.Size(31, 17)
        Me.TxtHirerachy.TabIndex = 47
        Me.TxtHirerachy.Value = ""
        Me.TxtHirerachy.Visible = False
        '
        'MyLabel8
        '
        Me.MyLabel8.FieldName = Nothing
        Me.MyLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel8.Location = New System.Drawing.Point(871, 2)
        Me.MyLabel8.Name = "MyLabel8"
        Me.MyLabel8.Size = New System.Drawing.Size(37, 16)
        Me.MyLabel8.TabIndex = 46
        Me.MyLabel8.Text = "VCGL"
        '
        'chkDeduction
        '
        Me.chkDeduction.Location = New System.Drawing.Point(445, 150)
        Me.chkDeduction.MyLinkLable1 = Nothing
        Me.chkDeduction.MyLinkLable2 = Nothing
        Me.chkDeduction.Name = "chkDeduction"
        Me.chkDeduction.Size = New System.Drawing.Size(72, 18)
        Me.chkDeduction.TabIndex = 44
        Me.chkDeduction.Tag1 = Nothing
        Me.chkDeduction.Text = "Deduction"
        '
        'LblLocDesp
        '
        Me.LblLocDesp.AutoSize = False
        Me.LblLocDesp.BorderVisible = True
        Me.LblLocDesp.FieldName = Nothing
        Me.LblLocDesp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblLocDesp.Location = New System.Drawing.Point(263, 25)
        Me.LblLocDesp.Name = "LblLocDesp"
        Me.LblLocDesp.Size = New System.Drawing.Size(410, 18)
        Me.LblLocDesp.TabIndex = 41
        Me.LblLocDesp.TextWrap = False
        '
        'chkProvision
        '
        Me.chkProvision.Location = New System.Drawing.Point(761, 66)
        Me.chkProvision.MyLinkLable1 = Nothing
        Me.chkProvision.MyLinkLable2 = Nothing
        Me.chkProvision.Name = "chkProvision"
        Me.chkProvision.Size = New System.Drawing.Size(66, 18)
        Me.chkProvision.TabIndex = 16
        Me.chkProvision.Tag1 = Nothing
        Me.chkProvision.Text = "Provision"
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
        Me.txtlocation.Location = New System.Drawing.Point(114, 24)
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
        Me.txtlocation.Size = New System.Drawing.Size(142, 18)
        Me.txtlocation.TabIndex = 4
        Me.txtlocation.Value = ""
        '
        'chkQuickMode
        '
        Me.chkQuickMode.Location = New System.Drawing.Point(678, 66)
        Me.chkQuickMode.MyLinkLable1 = Nothing
        Me.chkQuickMode.MyLinkLable2 = Nothing
        Me.chkQuickMode.Name = "chkQuickMode"
        Me.chkQuickMode.Size = New System.Drawing.Size(81, 18)
        Me.chkQuickMode.TabIndex = 15
        Me.chkQuickMode.Tag1 = Nothing
        Me.chkQuickMode.Text = "Quick Mode"
        '
        'txtACSet
        '
        Me.txtACSet.CalculationExpression = Nothing
        Me.txtACSet.Enabled = False
        Me.txtACSet.FieldCode = Nothing
        Me.txtACSet.FieldDesc = Nothing
        Me.txtACSet.FieldMaxLength = 0
        Me.txtACSet.FieldName = Nothing
        Me.txtACSet.isCalculatedField = False
        Me.txtACSet.IsSourceFromTable = False
        Me.txtACSet.IsSourceFromValueList = False
        Me.txtACSet.IsUnique = False
        Me.txtACSet.Location = New System.Drawing.Point(751, 26)
        Me.txtACSet.MendatroryField = False
        Me.txtACSet.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtACSet.MyLinkLable1 = Me.RadLabel6
        Me.txtACSet.MyLinkLable2 = Nothing
        Me.txtACSet.MyReadOnly = False
        Me.txtACSet.MyShowMasterFormButton = False
        Me.txtACSet.Name = "txtACSet"
        Me.txtACSet.ReferenceFieldDesc = Nothing
        Me.txtACSet.ReferenceFieldName = Nothing
        Me.txtACSet.ReferenceTableName = Nothing
        Me.txtACSet.Size = New System.Drawing.Size(119, 20)
        Me.txtACSet.TabIndex = 5
        Me.txtACSet.Value = ""
        '
        'RadLabel6
        '
        Me.RadLabel6.FieldName = Nothing
        Me.RadLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel6.Location = New System.Drawing.Point(678, 26)
        Me.RadLabel6.Name = "RadLabel6"
        Me.RadLabel6.Size = New System.Drawing.Size(67, 16)
        Me.RadLabel6.TabIndex = 30
        Me.RadLabel6.Text = "Account Set"
        '
        'TxtVendorNo
        '
        Me.TxtVendorNo.CalculationExpression = Nothing
        Me.TxtVendorNo.FieldCode = Nothing
        Me.TxtVendorNo.FieldDesc = Nothing
        Me.TxtVendorNo.FieldMaxLength = 0
        Me.TxtVendorNo.FieldName = Nothing
        Me.TxtVendorNo.isCalculatedField = False
        Me.TxtVendorNo.IsSourceFromTable = False
        Me.TxtVendorNo.IsSourceFromValueList = False
        Me.TxtVendorNo.IsUnique = False
        Me.TxtVendorNo.Location = New System.Drawing.Point(114, 45)
        Me.TxtVendorNo.MendatroryField = True
        Me.TxtVendorNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtVendorNo.MyLinkLable1 = Me.RadLabel2
        Me.TxtVendorNo.MyLinkLable2 = Me.lblVendorName
        Me.TxtVendorNo.MyReadOnly = False
        Me.TxtVendorNo.MyShowMasterFormButton = False
        Me.TxtVendorNo.Name = "TxtVendorNo"
        Me.TxtVendorNo.ReferenceFieldDesc = Nothing
        Me.TxtVendorNo.ReferenceFieldName = Nothing
        Me.TxtVendorNo.ReferenceTableName = Nothing
        Me.TxtVendorNo.Size = New System.Drawing.Size(143, 18)
        Me.TxtVendorNo.TabIndex = 6
        Me.TxtVendorNo.Value = ""
        '
        'RadLabel2
        '
        Me.RadLabel2.FieldName = Nothing
        Me.RadLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel2.Location = New System.Drawing.Point(3, 46)
        Me.RadLabel2.Name = "RadLabel2"
        Me.RadLabel2.Size = New System.Drawing.Size(60, 16)
        Me.RadLabel2.TabIndex = 27
        Me.RadLabel2.Text = "Vendor No"
        '
        'lblVendorName
        '
        Me.lblVendorName.AutoSize = False
        Me.lblVendorName.BorderVisible = True
        Me.lblVendorName.FieldName = Nothing
        Me.lblVendorName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorName.Location = New System.Drawing.Point(263, 45)
        Me.lblVendorName.Name = "lblVendorName"
        Me.lblVendorName.Size = New System.Drawing.Size(410, 18)
        Me.lblVendorName.TabIndex = 7
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
        Me.txtRefDocNo.Location = New System.Drawing.Point(347, 107)
        Me.txtRefDocNo.MendatroryField = False
        Me.txtRefDocNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRefDocNo.MyLinkLable1 = Me.RadLabel15
        Me.txtRefDocNo.MyLinkLable2 = Nothing
        Me.txtRefDocNo.MyReadOnly = False
        Me.txtRefDocNo.MyShowMasterFormButton = False
        Me.txtRefDocNo.Name = "txtRefDocNo"
        Me.txtRefDocNo.ReferenceFieldDesc = Nothing
        Me.txtRefDocNo.ReferenceFieldName = Nothing
        Me.txtRefDocNo.ReferenceTableName = Nothing
        Me.txtRefDocNo.Size = New System.Drawing.Size(326, 18)
        Me.txtRefDocNo.TabIndex = 17
        Me.txtRefDocNo.Value = ""
        '
        'RadLabel15
        '
        Me.RadLabel15.FieldName = Nothing
        Me.RadLabel15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel15.Location = New System.Drawing.Point(274, 108)
        Me.RadLabel15.Name = "RadLabel15"
        Me.RadLabel15.Size = New System.Drawing.Size(65, 16)
        Me.RadLabel15.TabIndex = 21
        Me.RadLabel15.Text = "Ref Doc No"
        '
        'UsLock1
        '
        Me.UsLock1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.UsLock1.Location = New System.Drawing.Point(751, 1)
        Me.UsLock1.MyFont = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsLock1.Name = "UsLock1"
        Me.UsLock1.Size = New System.Drawing.Size(118, 20)
        Me.UsLock1.Status = common.ERPTransactionStatus.Pending
        Me.UsLock1.TabIndex = 34
        '
        'RadLabel14
        '
        Me.RadLabel14.FieldName = Nothing
        Me.RadLabel14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel14.Location = New System.Drawing.Point(3, 108)
        Me.RadLabel14.Name = "RadLabel14"
        Me.RadLabel14.Size = New System.Drawing.Size(72, 16)
        Me.RadLabel14.TabIndex = 22
        Me.RadLabel14.Text = "Ref DocType"
        '
        'RadLabel13
        '
        Me.RadLabel13.FieldName = Nothing
        Me.RadLabel13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel13.Location = New System.Drawing.Point(230, 87)
        Me.RadLabel13.Name = "RadLabel13"
        Me.RadLabel13.Size = New System.Drawing.Size(109, 16)
        Me.RadLabel13.TabIndex = 23
        Me.RadLabel13.Text = "Vendor Invoice Date"
        '
        'RadLabel12
        '
        Me.RadLabel12.FieldName = Nothing
        Me.RadLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel12.Location = New System.Drawing.Point(465, 87)
        Me.RadLabel12.Name = "RadLabel12"
        Me.RadLabel12.Size = New System.Drawing.Size(100, 16)
        Me.RadLabel12.TabIndex = 19
        Me.RadLabel12.Text = "Document Amount"
        '
        'lblTotRAmt1
        '
        Me.lblTotRAmt1.AutoSize = False
        Me.lblTotRAmt1.BackColor = System.Drawing.Color.Transparent
        Me.lblTotRAmt1.BorderVisible = True
        Me.lblTotRAmt1.FieldName = Nothing
        Me.lblTotRAmt1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRAmt1.Location = New System.Drawing.Point(565, 86)
        Me.lblTotRAmt1.Name = "lblTotRAmt1"
        Me.lblTotRAmt1.Size = New System.Drawing.Size(108, 18)
        Me.lblTotRAmt1.TabIndex = 14
        Me.lblTotRAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel3
        '
        Me.RadLabel3.FieldName = Nothing
        Me.RadLabel3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel3.Location = New System.Drawing.Point(230, 67)
        Me.RadLabel3.Name = "RadLabel3"
        Me.RadLabel3.Size = New System.Drawing.Size(99, 16)
        Me.RadLabel3.TabIndex = 26
        Me.RadLabel3.Text = "Vendor Invoice No"
        '
        'RadLabel8
        '
        Me.RadLabel8.FieldName = Nothing
        Me.RadLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel8.Location = New System.Drawing.Point(678, 47)
        Me.RadLabel8.Name = "RadLabel8"
        Me.RadLabel8.Size = New System.Drawing.Size(53, 16)
        Me.RadLabel8.TabIndex = 29
        Me.RadLabel8.Text = "Order No"
        '
        'RadLabel7
        '
        Me.RadLabel7.FieldName = Nothing
        Me.RadLabel7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel7.Location = New System.Drawing.Point(3, 87)
        Me.RadLabel7.Name = "RadLabel7"
        Me.RadLabel7.Size = New System.Drawing.Size(40, 16)
        Me.RadLabel7.TabIndex = 24
        Me.RadLabel7.Text = "PO No"
        '
        'RadLabel5
        '
        Me.RadLabel5.FieldName = Nothing
        Me.RadLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel5.Location = New System.Drawing.Point(465, 67)
        Me.RadLabel5.Name = "RadLabel5"
        Me.RadLabel5.Size = New System.Drawing.Size(86, 16)
        Me.RadLabel5.TabIndex = 20
        Me.RadLabel5.Text = "Document Type"
        '
        'RadLabel4
        '
        Me.RadLabel4.FieldName = Nothing
        Me.RadLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel4.Location = New System.Drawing.Point(3, 67)
        Me.RadLabel4.Name = "RadLabel4"
        Me.RadLabel4.Size = New System.Drawing.Size(85, 16)
        Me.RadLabel4.TabIndex = 25
        Me.RadLabel4.Text = "Document Date"
        '
        'RadLabel1
        '
        Me.RadLabel1.FieldName = Nothing
        Me.RadLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel1.Location = New System.Drawing.Point(3, 3)
        Me.RadLabel1.Name = "RadLabel1"
        Me.RadLabel1.Size = New System.Drawing.Size(75, 16)
        Me.RadLabel1.TabIndex = 32
        Me.RadLabel1.Text = "Document No"
        '
        'cmbRefType
        '
        Me.cmbRefType.AllowDrop = True
        Me.cmbRefType.AutoCompleteDisplayMember = Nothing
        Me.cmbRefType.AutoCompleteValueMember = Nothing
        Me.cmbRefType.CalculationExpression = Nothing
        Me.cmbRefType.DropDownAnimationEnabled = True
        Me.cmbRefType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbRefType.FieldCode = Nothing
        Me.cmbRefType.FieldDesc = Nothing
        Me.cmbRefType.FieldMaxLength = 0
        Me.cmbRefType.FieldName = Nothing
        Me.cmbRefType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRefType.isCalculatedField = False
        Me.cmbRefType.IsSourceFromTable = False
        Me.cmbRefType.IsSourceFromValueList = False
        Me.cmbRefType.IsUnique = False
        Me.cmbRefType.Location = New System.Drawing.Point(114, 107)
        Me.cmbRefType.MendatroryField = False
        Me.cmbRefType.MyLinkLable1 = Me.RadLabel14
        Me.cmbRefType.MyLinkLable2 = Nothing
        Me.cmbRefType.Name = "cmbRefType"
        Me.cmbRefType.ReferenceFieldDesc = Nothing
        Me.cmbRefType.ReferenceFieldName = Nothing
        Me.cmbRefType.ReferenceTableName = Nothing
        Me.cmbRefType.Size = New System.Drawing.Size(114, 18)
        Me.cmbRefType.TabIndex = 16
        Me.cmbRefType.Text = "RadDropDownList1"
        '
        'txtDocNo
        '
        Me.txtDocNo.FieldName = Nothing
        Me.txtDocNo.Location = New System.Drawing.Point(115, 1)
        Me.txtDocNo.MendatroryField = False
        Me.txtDocNo.MyCharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.txtDocNo.MyFont = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.txtDocNo.MyLinkLable1 = Me.RadLabel1
        Me.txtDocNo.MyLinkLable2 = Nothing
        Me.txtDocNo.MyMaxLength = 30
        Me.txtDocNo.MyReadOnly = False
        Me.txtDocNo.Name = "txtDocNo"
        Me.txtDocNo.Size = New System.Drawing.Size(242, 20)
        Me.txtDocNo.TabIndex = 0
        Me.txtDocNo.TabStop = False
        Me.txtDocNo.Value = ""
        '
        'txtVendorInvDatre
        '
        Me.txtVendorInvDatre.CalculationExpression = Nothing
        Me.txtVendorInvDatre.CustomFormat = "dd/MM/yyyy"
        Me.txtVendorInvDatre.FieldCode = Nothing
        Me.txtVendorInvDatre.FieldDesc = Nothing
        Me.txtVendorInvDatre.FieldMaxLength = 0
        Me.txtVendorInvDatre.FieldName = Nothing
        Me.txtVendorInvDatre.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorInvDatre.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtVendorInvDatre.isCalculatedField = False
        Me.txtVendorInvDatre.IsSourceFromTable = False
        Me.txtVendorInvDatre.IsSourceFromValueList = False
        Me.txtVendorInvDatre.IsUnique = False
        Me.txtVendorInvDatre.Location = New System.Drawing.Point(347, 86)
        Me.txtVendorInvDatre.MendatroryField = False
        Me.txtVendorInvDatre.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtVendorInvDatre.MyLinkLable1 = Me.RadLabel13
        Me.txtVendorInvDatre.MyLinkLable2 = Nothing
        Me.txtVendorInvDatre.Name = "txtVendorInvDatre"
        Me.txtVendorInvDatre.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtVendorInvDatre.ReferenceFieldDesc = Nothing
        Me.txtVendorInvDatre.ReferenceFieldName = Nothing
        Me.txtVendorInvDatre.ReferenceTableName = Nothing
        Me.txtVendorInvDatre.Size = New System.Drawing.Size(114, 18)
        Me.txtVendorInvDatre.TabIndex = 13
        Me.txtVendorInvDatre.TabStop = False
        Me.txtVendorInvDatre.Text = "13/06/2011"
        Me.txtVendorInvDatre.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'txtVendorInvoiceNo
        '
        Me.txtVendorInvoiceNo.CalculationExpression = Nothing
        Me.txtVendorInvoiceNo.FieldCode = Nothing
        Me.txtVendorInvoiceNo.FieldDesc = Nothing
        Me.txtVendorInvoiceNo.FieldMaxLength = 0
        Me.txtVendorInvoiceNo.FieldName = Nothing
        Me.txtVendorInvoiceNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorInvoiceNo.isCalculatedField = False
        Me.txtVendorInvoiceNo.IsSourceFromTable = False
        Me.txtVendorInvoiceNo.IsSourceFromValueList = False
        Me.txtVendorInvoiceNo.IsUnique = False
        Me.txtVendorInvoiceNo.Location = New System.Drawing.Point(347, 66)
        Me.txtVendorInvoiceNo.MaxLength = 30
        Me.txtVendorInvoiceNo.MendatroryField = True
        Me.txtVendorInvoiceNo.MyLinkLable1 = Me.RadLabel3
        Me.txtVendorInvoiceNo.MyLinkLable2 = Nothing
        Me.txtVendorInvoiceNo.Name = "txtVendorInvoiceNo"
        Me.txtVendorInvoiceNo.ReferenceFieldDesc = Nothing
        Me.txtVendorInvoiceNo.ReferenceFieldName = Nothing
        Me.txtVendorInvoiceNo.ReferenceTableName = Nothing
        Me.txtVendorInvoiceNo.Size = New System.Drawing.Size(114, 18)
        Me.txtVendorInvoiceNo.TabIndex = 10
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
        Me.txtOrderNo.Location = New System.Drawing.Point(751, 48)
        Me.txtOrderNo.MaxLength = 30
        Me.txtOrderNo.MendatroryField = False
        Me.txtOrderNo.MyLinkLable1 = Me.RadLabel8
        Me.txtOrderNo.MyLinkLable2 = Nothing
        Me.txtOrderNo.Name = "txtOrderNo"
        Me.txtOrderNo.ReferenceFieldDesc = Nothing
        Me.txtOrderNo.ReferenceFieldName = Nothing
        Me.txtOrderNo.ReferenceTableName = Nothing
        Me.txtOrderNo.Size = New System.Drawing.Size(118, 18)
        Me.txtOrderNo.TabIndex = 8
        '
        'txtPONo
        '
        Me.txtPONo.CalculationExpression = Nothing
        Me.txtPONo.FieldCode = Nothing
        Me.txtPONo.FieldDesc = Nothing
        Me.txtPONo.FieldMaxLength = 0
        Me.txtPONo.FieldName = Nothing
        Me.txtPONo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPONo.isCalculatedField = False
        Me.txtPONo.IsSourceFromTable = False
        Me.txtPONo.IsSourceFromValueList = False
        Me.txtPONo.IsUnique = False
        Me.txtPONo.Location = New System.Drawing.Point(114, 86)
        Me.txtPONo.MaxLength = 30
        Me.txtPONo.MendatroryField = False
        Me.txtPONo.MyLinkLable1 = Me.RadLabel7
        Me.txtPONo.MyLinkLable2 = Nothing
        Me.txtPONo.Name = "txtPONo"
        Me.txtPONo.ReferenceFieldDesc = Nothing
        Me.txtPONo.ReferenceFieldName = Nothing
        Me.txtPONo.ReferenceTableName = Nothing
        Me.txtPONo.Size = New System.Drawing.Size(114, 18)
        Me.txtPONo.TabIndex = 12
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
        Me.cboDocType.Location = New System.Drawing.Point(567, 66)
        Me.cboDocType.MendatroryField = False
        Me.cboDocType.MyLinkLable1 = Me.RadLabel5
        Me.cboDocType.MyLinkLable2 = Nothing
        Me.cboDocType.Name = "cboDocType"
        Me.cboDocType.ReferenceFieldDesc = Nothing
        Me.cboDocType.ReferenceFieldName = Nothing
        Me.cboDocType.ReferenceTableName = Nothing
        Me.cboDocType.Size = New System.Drawing.Size(106, 18)
        Me.cboDocType.TabIndex = 11
        Me.cboDocType.Text = "RadDropDownList1"
        '
        'txtDate
        '
        Me.txtDate.CalculationExpression = Nothing
        Me.txtDate.CustomFormat = "dd/MM/yyyy"
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
        Me.txtDate.Location = New System.Drawing.Point(114, 66)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Me.RadLabel4
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(114, 18)
        Me.txtDate.TabIndex = 9
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "13/06/2011"
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
        Me.txtDesc.Location = New System.Drawing.Point(400, 2)
        Me.txtDesc.MendatroryField = False
        Me.txtDesc.MyLinkLable1 = Nothing
        Me.txtDesc.MyLinkLable2 = Nothing
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.ReferenceFieldDesc = Nothing
        Me.txtDesc.ReferenceFieldName = Nothing
        Me.txtDesc.ReferenceTableName = Nothing
        Me.txtDesc.Size = New System.Drawing.Size(271, 18)
        Me.txtDesc.TabIndex = 2
        '
        'grpProvision
        '
        Me.grpProvision.Controls.Add(Me.btlShowProvision)
        Me.grpProvision.Controls.Add(Me.MyLabel7)
        Me.grpProvision.Controls.Add(Me.txtProvAmt)
        Me.grpProvision.Controls.Add(Me.btnProvSelect)
        Me.grpProvision.Controls.Add(Me.MyLabel6)
        Me.grpProvision.Controls.Add(Me.dtpToProv)
        Me.grpProvision.Controls.Add(Me.MyLabel5)
        Me.grpProvision.Controls.Add(Me.dtpFromProv)
        Me.grpProvision.Location = New System.Drawing.Point(940, 42)
        Me.grpProvision.Name = "grpProvision"
        Me.grpProvision.Size = New System.Drawing.Size(156, 95)
        Me.grpProvision.TabIndex = 42
        Me.grpProvision.TabStop = False
        '
        'btlShowProvision
        '
        Me.btlShowProvision.AutoSize = True
        Me.btlShowProvision.Location = New System.Drawing.Point(1, 79)
        Me.btlShowProvision.Name = "btlShowProvision"
        Me.btlShowProvision.Size = New System.Drawing.Size(116, 14)
        Me.btlShowProvision.TabIndex = 47
        Me.btlShowProvision.TabStop = True
        Me.btlShowProvision.Text = "List Selected Provision"
        '
        'MyLabel7
        '
        Me.MyLabel7.FieldName = Nothing
        Me.MyLabel7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel7.Location = New System.Drawing.Point(2, 62)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(56, 16)
        Me.MyLabel7.TabIndex = 46
        Me.MyLabel7.Text = "Prov Amt."
        '
        'txtProvAmt
        '
        Me.txtProvAmt.CalculationExpression = Nothing
        Me.txtProvAmt.FieldCode = Nothing
        Me.txtProvAmt.FieldDesc = Nothing
        Me.txtProvAmt.FieldMaxLength = 0
        Me.txtProvAmt.FieldName = Nothing
        Me.txtProvAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProvAmt.isCalculatedField = False
        Me.txtProvAmt.IsSourceFromTable = False
        Me.txtProvAmt.IsSourceFromValueList = False
        Me.txtProvAmt.IsUnique = False
        Me.txtProvAmt.Location = New System.Drawing.Point(71, 61)
        Me.txtProvAmt.MaxLength = 30
        Me.txtProvAmt.MendatroryField = False
        Me.txtProvAmt.MyLinkLable1 = Me.MyLabel7
        Me.txtProvAmt.MyLinkLable2 = Nothing
        Me.txtProvAmt.Name = "txtProvAmt"
        Me.txtProvAmt.ReadOnly = True
        Me.txtProvAmt.ReferenceFieldDesc = Nothing
        Me.txtProvAmt.ReferenceFieldName = Nothing
        Me.txtProvAmt.ReferenceTableName = Nothing
        Me.txtProvAmt.Size = New System.Drawing.Size(78, 18)
        Me.txtProvAmt.TabIndex = 45
        '
        'btnProvSelect
        '
        Me.btnProvSelect.AutoSize = True
        Me.btnProvSelect.Location = New System.Drawing.Point(-1, 7)
        Me.btnProvSelect.Name = "btnProvSelect"
        Me.btnProvSelect.Size = New System.Drawing.Size(123, 14)
        Me.btnProvSelect.TabIndex = 44
        Me.btnProvSelect.TabStop = True
        Me.btnProvSelect.Text = "Click To Select Provision"
        '
        'MyLabel6
        '
        Me.MyLabel6.FieldName = Nothing
        Me.MyLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel6.Location = New System.Drawing.Point(2, 43)
        Me.MyLabel6.Name = "MyLabel6"
        Me.MyLabel6.Size = New System.Drawing.Size(46, 16)
        Me.MyLabel6.TabIndex = 27
        Me.MyLabel6.Text = "To Date"
        '
        'dtpToProv
        '
        Me.dtpToProv.CalculationExpression = Nothing
        Me.dtpToProv.CustomFormat = "dd/MM/yyyy"
        Me.dtpToProv.FieldCode = Nothing
        Me.dtpToProv.FieldDesc = Nothing
        Me.dtpToProv.FieldMaxLength = 0
        Me.dtpToProv.FieldName = Nothing
        Me.dtpToProv.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpToProv.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToProv.isCalculatedField = False
        Me.dtpToProv.IsSourceFromTable = False
        Me.dtpToProv.IsSourceFromValueList = False
        Me.dtpToProv.IsUnique = False
        Me.dtpToProv.Location = New System.Drawing.Point(71, 42)
        Me.dtpToProv.MendatroryField = False
        Me.dtpToProv.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.dtpToProv.MyLinkLable1 = Me.MyLabel6
        Me.dtpToProv.MyLinkLable2 = Nothing
        Me.dtpToProv.Name = "dtpToProv"
        Me.dtpToProv.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.dtpToProv.ReferenceFieldDesc = Nothing
        Me.dtpToProv.ReferenceFieldName = Nothing
        Me.dtpToProv.ReferenceTableName = Nothing
        Me.dtpToProv.Size = New System.Drawing.Size(78, 18)
        Me.dtpToProv.TabIndex = 26
        Me.dtpToProv.TabStop = False
        Me.dtpToProv.Text = "13/06/2011"
        Me.dtpToProv.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'MyLabel5
        '
        Me.MyLabel5.FieldName = Nothing
        Me.MyLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(2, 24)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(60, 16)
        Me.MyLabel5.TabIndex = 25
        Me.MyLabel5.Text = "From Date"
        '
        'dtpFromProv
        '
        Me.dtpFromProv.CalculationExpression = Nothing
        Me.dtpFromProv.CustomFormat = "dd/MM/yyyy"
        Me.dtpFromProv.FieldCode = Nothing
        Me.dtpFromProv.FieldDesc = Nothing
        Me.dtpFromProv.FieldMaxLength = 0
        Me.dtpFromProv.FieldName = Nothing
        Me.dtpFromProv.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpFromProv.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromProv.isCalculatedField = False
        Me.dtpFromProv.IsSourceFromTable = False
        Me.dtpFromProv.IsSourceFromValueList = False
        Me.dtpFromProv.IsUnique = False
        Me.dtpFromProv.Location = New System.Drawing.Point(71, 23)
        Me.dtpFromProv.MendatroryField = False
        Me.dtpFromProv.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.dtpFromProv.MyLinkLable1 = Me.MyLabel5
        Me.dtpFromProv.MyLinkLable2 = Nothing
        Me.dtpFromProv.Name = "dtpFromProv"
        Me.dtpFromProv.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.dtpFromProv.ReferenceFieldDesc = Nothing
        Me.dtpFromProv.ReferenceFieldName = Nothing
        Me.dtpFromProv.ReferenceTableName = Nothing
        Me.dtpFromProv.Size = New System.Drawing.Size(78, 18)
        Me.dtpFromProv.TabIndex = 24
        Me.dtpFromProv.TabStop = False
        Me.dtpFromProv.Text = "13/06/2011"
        Me.dtpFromProv.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'RadPageViewPage2
        '
        Me.RadPageViewPage2.Controls.Add(Me.GroupBox1)
        Me.RadPageViewPage2.Controls.Add(Me.txtTaxGroup)
        Me.RadPageViewPage2.Controls.Add(Me.RadLabel10)
        Me.RadPageViewPage2.Controls.Add(Me.RadGroupBox1)
        Me.RadPageViewPage2.Controls.Add(Me.lblTaxGrpName)
        Me.RadPageViewPage2.Controls.Add(Me.gv2)
        Me.RadPageViewPage2.Controls.Add(Me.RadLabel11)
        Me.RadPageViewPage2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage2.ItemSize = New System.Drawing.SizeF(47.0!, 26.0!)
        Me.RadPageViewPage2.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage2.Name = "RadPageViewPage2"
        Me.RadPageViewPage2.Size = New System.Drawing.Size(1105, 446)
        Me.RadPageViewPage2.Text = "Taxes"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbtnTaxCalManual)
        Me.GroupBox1.Controls.Add(Me.rbtnTaxCalAutomatic)
        Me.GroupBox1.Location = New System.Drawing.Point(553, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(163, 36)
        Me.GroupBox1.TabIndex = 1
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
        Me.txtTaxGroup.MyLinkLable1 = Me.RadLabel11
        Me.txtTaxGroup.MyLinkLable2 = Me.lblTaxGrpName
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
        'RadLabel11
        '
        Me.RadLabel11.FieldName = Nothing
        Me.RadLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel11.Location = New System.Drawing.Point(3, 7)
        Me.RadLabel11.Name = "RadLabel11"
        Me.RadLabel11.Size = New System.Drawing.Size(60, 16)
        Me.RadLabel11.TabIndex = 3
        Me.RadLabel11.Text = "Tax Group"
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
        Me.lblTaxGrpName.TabIndex = 4
        '
        'RadLabel10
        '
        Me.RadLabel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadLabel10.FieldName = Nothing
        Me.RadLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel10.ForeColor = System.Drawing.SystemColors.Highlight
        Me.RadLabel10.Location = New System.Drawing.Point(936, 342)
        Me.RadLabel10.Name = "RadLabel10"
        Me.RadLabel10.Size = New System.Drawing.Size(162, 16)
        Me.RadLabel10.TabIndex = 4
        Me.RadLabel10.Text = "Double click To Change Rate"
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
        Me.RadGroupBox1.Location = New System.Drawing.Point(1, 355)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox1.Size = New System.Drawing.Size(1105, 87)
        Me.RadGroupBox1.TabIndex = 2
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
        Me.txtTermCode.MyLinkLable1 = Me.RadLabel16
        Me.txtTermCode.MyLinkLable2 = Me.lblTermName
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
        Me.RadLabel16.TabIndex = 3
        Me.RadLabel16.Text = "Term Code"
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
        Me.lblTermName.TabIndex = 4
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
        Me.txtDueDate.MyLinkLable1 = Me.RadLabel17
        Me.txtDueDate.MyLinkLable2 = Nothing
        Me.txtDueDate.Name = "txtDueDate"
        Me.txtDueDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDueDate.ReferenceFieldDesc = Nothing
        Me.txtDueDate.ReferenceFieldName = Nothing
        Me.txtDueDate.ReferenceTableName = Nothing
        Me.txtDueDate.Size = New System.Drawing.Size(116, 18)
        Me.txtDueDate.TabIndex = 1
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
        Me.RadLabel17.TabIndex = 2
        Me.RadLabel17.Text = "Due Date"
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
        Me.gv2.MasterTemplate.ShowFilteringRow = False
        Me.gv2.MasterTemplate.ShowHeaderCellButtons = True
        Me.gv2.MasterTemplate.ViewDefinition = TableViewDefinition2
        Me.gv2.MyExportAPI = False
        Me.gv2.MyExportFilePath = ""
        Me.gv2.MyStopExport = False
        Me.gv2.Name = "gv2"
        Me.gv2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gv2.ShowHeaderCellButtons = True
        Me.gv2.Size = New System.Drawing.Size(1100, 300)
        Me.gv2.TabIndex = 2
        Me.gv2.TabStop = False
        Me.gv2.VarID = ""
        '
        'RadPageViewPage3
        '
        Me.RadPageViewPage3.Controls.Add(Me.SplitContainer2)
        Me.RadPageViewPage3.ItemSize = New System.Drawing.SizeF(112.0!, 26.0!)
        Me.RadPageViewPage3.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage3.Name = "RadPageViewPage3"
        Me.RadPageViewPage3.Size = New System.Drawing.Size(1105, 446)
        Me.RadPageViewPage3.Text = "Additional Charges"
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
        Me.SplitContainer2.Size = New System.Drawing.Size(1105, 446)
        Me.SplitContainer2.SplitterDistance = 402
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
        Me.gvAC.MasterTemplate.ShowFilteringRow = False
        Me.gvAC.MasterTemplate.ShowHeaderCellButtons = True
        Me.gvAC.MasterTemplate.ViewDefinition = TableViewDefinition3
        Me.gvAC.MyExportAPI = False
        Me.gvAC.MyExportFilePath = ""
        Me.gvAC.MyStopExport = False
        Me.gvAC.Name = "gvAC"
        Me.gvAC.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gvAC.ShowGroupPanel = False
        Me.gvAC.ShowHeaderCellButtons = True
        Me.gvAC.Size = New System.Drawing.Size(1105, 402)
        Me.gvAC.TabIndex = 0
        Me.gvAC.TabStop = False
        Me.gvAC.VarID = ""
        '
        'RadLabel31
        '
        Me.RadLabel31.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadLabel31.FieldName = Nothing
        Me.RadLabel31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel31.Location = New System.Drawing.Point(860, 3)
        Me.RadLabel31.Name = "RadLabel31"
        Me.RadLabel31.Size = New System.Drawing.Size(130, 16)
        Me.RadLabel31.TabIndex = 1
        Me.RadLabel31.Text = "Total Additional Charges"
        '
        'lblAddCharges
        '
        Me.lblAddCharges.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblAddCharges.AutoSize = False
        Me.lblAddCharges.BorderVisible = True
        Me.lblAddCharges.FieldName = Nothing
        Me.lblAddCharges.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddCharges.Location = New System.Drawing.Point(992, 3)
        Me.lblAddCharges.Name = "lblAddCharges"
        Me.lblAddCharges.Size = New System.Drawing.Size(110, 18)
        Me.lblAddCharges.TabIndex = 0
        Me.lblAddCharges.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'pvpCustomFields
        '
        Me.pvpCustomFields.ItemSize = New System.Drawing.SizeF(89.0!, 26.0!)
        Me.pvpCustomFields.Location = New System.Drawing.Point(10, 35)
        Me.pvpCustomFields.Name = "pvpCustomFields"
        Me.pvpCustomFields.Size = New System.Drawing.Size(1105, 446)
        Me.pvpCustomFields.Text = "Custom Fields"
        '
        'Attachments
        '
        Me.Attachments.Controls.Add(Me.UcAttachment1)
        Me.Attachments.ItemSize = New System.Drawing.SizeF(79.0!, 26.0!)
        Me.Attachments.Location = New System.Drawing.Point(10, 35)
        Me.Attachments.Name = "Attachments"
        Me.Attachments.Size = New System.Drawing.Size(1105, 446)
        Me.Attachments.Text = "Attachments"
        '
        'UcAttachment1
        '
        Me.UcAttachment1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UcAttachment1.Location = New System.Drawing.Point(0, 0)
        Me.UcAttachment1.Name = "UcAttachment1"
        Me.UcAttachment1.Size = New System.Drawing.Size(1105, 446)
        Me.UcAttachment1.TabIndex = 0
        Me.UcAttachment1.TabStop = False
        '
        'RadPageViewPage4
        '
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel14)
        Me.RadPageViewPage4.Controls.Add(Me.lblRoundOff)
        Me.RadPageViewPage4.Controls.Add(Me.TxtCashDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.txtAmtAfterCashDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.lblAmtAfterCashDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.lblCashDiscount)
        Me.RadPageViewPage4.Controls.Add(Me.pnlCurrConv)
        Me.RadPageViewPage4.Controls.Add(Me.btnReverse)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel3)
        Me.RadPageViewPage4.Controls.Add(Me.lblLandedAmt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel2)
        Me.RadPageViewPage4.Controls.Add(Me.lblTotEmptyAmt)
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
        Me.RadPageViewPage4.Size = New System.Drawing.Size(1105, 446)
        Me.RadPageViewPage4.Text = "Total"
        '
        'MyLabel14
        '
        Me.MyLabel14.FieldName = Nothing
        Me.MyLabel14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel14.Location = New System.Drawing.Point(159, 212)
        Me.MyLabel14.Name = "MyLabel14"
        Me.MyLabel14.Size = New System.Drawing.Size(58, 16)
        Me.MyLabel14.TabIndex = 164
        Me.MyLabel14.Text = "Round Off"
        '
        'lblRoundOff
        '
        Me.lblRoundOff.AutoSize = False
        Me.lblRoundOff.BorderVisible = True
        Me.lblRoundOff.FieldName = Nothing
        Me.lblRoundOff.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRoundOff.Location = New System.Drawing.Point(220, 211)
        Me.lblRoundOff.Name = "lblRoundOff"
        Me.lblRoundOff.Size = New System.Drawing.Size(110, 18)
        Me.lblRoundOff.TabIndex = 165
        Me.lblRoundOff.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'TxtCashDiscount
        '
        Me.TxtCashDiscount.BackColor = System.Drawing.Color.White
        Me.TxtCashDiscount.CalculationExpression = Nothing
        Me.TxtCashDiscount.DecimalPlaces = 2
        Me.TxtCashDiscount.FieldCode = Nothing
        Me.TxtCashDiscount.FieldDesc = Nothing
        Me.TxtCashDiscount.FieldMaxLength = 0
        Me.TxtCashDiscount.FieldName = Nothing
        Me.TxtCashDiscount.isCalculatedField = False
        Me.TxtCashDiscount.IsSourceFromTable = False
        Me.TxtCashDiscount.IsSourceFromValueList = False
        Me.TxtCashDiscount.IsUnique = False
        Me.TxtCashDiscount.Location = New System.Drawing.Point(220, 257)
        Me.TxtCashDiscount.MendatroryField = False
        Me.TxtCashDiscount.MyLinkLable1 = Nothing
        Me.TxtCashDiscount.MyLinkLable2 = Nothing
        Me.TxtCashDiscount.Name = "TxtCashDiscount"
        Me.TxtCashDiscount.ReferenceFieldDesc = Nothing
        Me.TxtCashDiscount.ReferenceFieldName = Nothing
        Me.TxtCashDiscount.ReferenceTableName = Nothing
        Me.TxtCashDiscount.Size = New System.Drawing.Size(109, 20)
        Me.TxtCashDiscount.TabIndex = 163
        Me.TxtCashDiscount.Text = "0"
        Me.TxtCashDiscount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TxtCashDiscount.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'txtAmtAfterCashDiscount
        '
        Me.txtAmtAfterCashDiscount.AutoSize = False
        Me.txtAmtAfterCashDiscount.BorderVisible = True
        Me.txtAmtAfterCashDiscount.FieldName = Nothing
        Me.txtAmtAfterCashDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmtAfterCashDiscount.Location = New System.Drawing.Point(220, 282)
        Me.txtAmtAfterCashDiscount.Name = "txtAmtAfterCashDiscount"
        Me.txtAmtAfterCashDiscount.Size = New System.Drawing.Size(109, 18)
        Me.txtAmtAfterCashDiscount.TabIndex = 162
        Me.txtAmtAfterCashDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtAfterCashDiscount
        '
        Me.lblAmtAfterCashDiscount.FieldName = Nothing
        Me.lblAmtAfterCashDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtAfterCashDiscount.Location = New System.Drawing.Point(68, 281)
        Me.lblAmtAfterCashDiscount.Name = "lblAmtAfterCashDiscount"
        Me.lblAmtAfterCashDiscount.Size = New System.Drawing.Size(149, 16)
        Me.lblAmtAfterCashDiscount.TabIndex = 161
        Me.lblAmtAfterCashDiscount.Text = "Amount After Cash Discount"
        '
        'lblCashDiscount
        '
        Me.lblCashDiscount.FieldName = Nothing
        Me.lblCashDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCashDiscount.Location = New System.Drawing.Point(137, 258)
        Me.lblCashDiscount.Name = "lblCashDiscount"
        Me.lblCashDiscount.Size = New System.Drawing.Size(80, 16)
        Me.lblCashDiscount.TabIndex = 160
        Me.lblCashDiscount.Text = "Cash Discount"
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
        Me.pnlCurrConv.Location = New System.Drawing.Point(31, 6)
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
        Me.lblEffectiveFrom.TabIndex = 3
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
        Me.lblCurrency.TabIndex = 5
        Me.lblCurrency.Text = "Currency"
        '
        'lblConvRate
        '
        Me.lblConvRate.FieldName = Nothing
        Me.lblConvRate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConvRate.Location = New System.Drawing.Point(256, 12)
        Me.lblConvRate.Name = "lblConvRate"
        Me.lblConvRate.Size = New System.Drawing.Size(91, 16)
        Me.lblConvRate.TabIndex = 4
        Me.lblConvRate.Text = "Conversion Rate"
        '
        'btnReverse
        '
        Me.btnReverse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnReverse.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReverse.Location = New System.Drawing.Point(592, 51)
        Me.btnReverse.Name = "btnReverse"
        Me.btnReverse.Size = New System.Drawing.Size(267, 22)
        Me.btnReverse.TabIndex = 1
        Me.btnReverse.Text = "Revese and Unpost"
        Me.btnReverse.Visible = False
        '
        'MyLabel3
        '
        Me.MyLabel3.FieldName = Nothing
        Me.MyLabel3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel3.Location = New System.Drawing.Point(121, 189)
        Me.MyLabel3.Name = "MyLabel3"
        Me.MyLabel3.Size = New System.Drawing.Size(96, 16)
        Me.MyLabel3.TabIndex = 14
        Me.MyLabel3.Text = "+ Landed Amount"
        '
        'lblLandedAmt
        '
        Me.lblLandedAmt.AutoSize = False
        Me.lblLandedAmt.BorderVisible = True
        Me.lblLandedAmt.FieldName = Nothing
        Me.lblLandedAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLandedAmt.Location = New System.Drawing.Point(220, 188)
        Me.lblLandedAmt.Name = "lblLandedAmt"
        Me.lblLandedAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblLandedAmt.TabIndex = 15
        Me.lblLandedAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(127, 166)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(90, 16)
        Me.MyLabel2.TabIndex = 12
        Me.MyLabel2.Text = "+ Empty Amount"
        '
        'lblTotEmptyAmt
        '
        Me.lblTotEmptyAmt.AutoSize = False
        Me.lblTotEmptyAmt.BorderVisible = True
        Me.lblTotEmptyAmt.FieldName = Nothing
        Me.lblTotEmptyAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotEmptyAmt.Location = New System.Drawing.Point(220, 165)
        Me.lblTotEmptyAmt.Name = "lblTotEmptyAmt"
        Me.lblTotEmptyAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTotEmptyAmt.TabIndex = 13
        Me.lblTotEmptyAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel32
        '
        Me.RadLabel32.FieldName = Nothing
        Me.RadLabel32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel32.Location = New System.Drawing.Point(77, 143)
        Me.RadLabel32.Name = "RadLabel32"
        Me.RadLabel32.Size = New System.Drawing.Size(140, 16)
        Me.RadLabel32.TabIndex = 10
        Me.RadLabel32.Text = "+ Total Additional Charges"
        '
        'lblAddCharges1
        '
        Me.lblAddCharges1.AutoSize = False
        Me.lblAddCharges1.BorderVisible = True
        Me.lblAddCharges1.FieldName = Nothing
        Me.lblAddCharges1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddCharges1.Location = New System.Drawing.Point(220, 142)
        Me.lblAddCharges1.Name = "lblAddCharges1"
        Me.lblAddCharges1.Size = New System.Drawing.Size(110, 18)
        Me.lblAddCharges1.TabIndex = 11
        Me.lblAddCharges1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel9
        '
        Me.RadLabel9.FieldName = Nothing
        Me.RadLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel9.Location = New System.Drawing.Point(97, 97)
        Me.RadLabel9.Name = "RadLabel9"
        Me.RadLabel9.Size = New System.Drawing.Size(120, 16)
        Me.RadLabel9.TabIndex = 6
        Me.RadLabel9.Text = "Amount After Discount"
        '
        'RadLabel27
        '
        Me.RadLabel27.FieldName = Nothing
        Me.RadLabel27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel27.Location = New System.Drawing.Point(117, 235)
        Me.RadLabel27.Name = "RadLabel27"
        Me.RadLabel27.Size = New System.Drawing.Size(100, 16)
        Me.RadLabel27.TabIndex = 16
        Me.RadLabel27.Text = "Document Amount"
        '
        'lblTotRAmt
        '
        Me.lblTotRAmt.AutoSize = False
        Me.lblTotRAmt.BorderVisible = True
        Me.lblTotRAmt.FieldName = Nothing
        Me.lblTotRAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRAmt.Location = New System.Drawing.Point(220, 234)
        Me.lblTotRAmt.Name = "lblTotRAmt"
        Me.lblTotRAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTotRAmt.TabIndex = 17
        Me.lblTotRAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel25
        '
        Me.RadLabel25.FieldName = Nothing
        Me.RadLabel25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel25.Location = New System.Drawing.Point(140, 120)
        Me.RadLabel25.Name = "RadLabel25"
        Me.RadLabel25.Size = New System.Drawing.Size(77, 16)
        Me.RadLabel25.TabIndex = 8
        Me.RadLabel25.Text = "+ Tax Amount"
        '
        'lblTaxAmt
        '
        Me.lblTaxAmt.AutoSize = False
        Me.lblTaxAmt.BorderVisible = True
        Me.lblTaxAmt.FieldName = Nothing
        Me.lblTaxAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxAmt.Location = New System.Drawing.Point(220, 119)
        Me.lblTaxAmt.Name = "lblTaxAmt"
        Me.lblTaxAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTaxAmt.TabIndex = 9
        Me.lblTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtAfterDiscount
        '
        Me.lblAmtAfterDiscount.AutoSize = False
        Me.lblAmtAfterDiscount.BorderVisible = True
        Me.lblAmtAfterDiscount.FieldName = Nothing
        Me.lblAmtAfterDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtAfterDiscount.Location = New System.Drawing.Point(220, 96)
        Me.lblAmtAfterDiscount.Name = "lblAmtAfterDiscount"
        Me.lblAmtAfterDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtAfterDiscount.TabIndex = 7
        Me.lblAmtAfterDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDiscountAmt
        '
        Me.lblDiscountAmt.AutoSize = False
        Me.lblDiscountAmt.BorderVisible = True
        Me.lblDiscountAmt.FieldName = Nothing
        Me.lblDiscountAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiscountAmt.Location = New System.Drawing.Point(220, 73)
        Me.lblDiscountAmt.Name = "lblDiscountAmt"
        Me.lblDiscountAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblDiscountAmt.TabIndex = 5
        Me.lblDiscountAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtWithDiscount
        '
        Me.lblAmtWithDiscount.AutoSize = False
        Me.lblAmtWithDiscount.BorderVisible = True
        Me.lblAmtWithDiscount.FieldName = Nothing
        Me.lblAmtWithDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtWithDiscount.Location = New System.Drawing.Point(220, 50)
        Me.lblAmtWithDiscount.Name = "lblAmtWithDiscount"
        Me.lblAmtWithDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtWithDiscount.TabIndex = 3
        Me.lblAmtWithDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel22
        '
        Me.RadLabel22.FieldName = Nothing
        Me.RadLabel22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel22.Location = New System.Drawing.Point(118, 74)
        Me.RadLabel22.Name = "RadLabel22"
        Me.RadLabel22.Size = New System.Drawing.Size(99, 16)
        Me.RadLabel22.TabIndex = 4
        Me.RadLabel22.Text = "- Discount Amount"
        '
        'RadLabel19
        '
        Me.RadLabel19.FieldName = Nothing
        Me.RadLabel19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel19.Location = New System.Drawing.Point(31, 51)
        Me.RadLabel19.Name = "RadLabel19"
        Me.RadLabel19.Size = New System.Drawing.Size(186, 16)
        Me.RadLabel19.TabIndex = 2
        Me.RadLabel19.Text = "Document Amount without Discount"
        '
        'btnHistory
        '
        Me.btnHistory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHistory.Location = New System.Drawing.Point(662, 6)
        Me.btnHistory.Name = "btnHistory"
        Me.btnHistory.Size = New System.Drawing.Size(69, 22)
        Me.btnHistory.TabIndex = 9
        Me.btnHistory.Text = "History"
        '
        'butCostCenterAndHirerachy_Update_AfterPost
        '
        Me.butCostCenterAndHirerachy_Update_AfterPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butCostCenterAndHirerachy_Update_AfterPost.Location = New System.Drawing.Point(473, 6)
        Me.butCostCenterAndHirerachy_Update_AfterPost.Name = "butCostCenterAndHirerachy_Update_AfterPost"
        Me.butCostCenterAndHirerachy_Update_AfterPost.Size = New System.Drawing.Size(186, 22)
        Me.butCostCenterAndHirerachy_Update_AfterPost.TabIndex = 8
        Me.butCostCenterAndHirerachy_Update_AfterPost.Text = "Update Cost Center And Hirerachy"
        '
        'btnsetting
        '
        Me.btnsetting.Items.AddRange(New Telerik.WinControls.RadItem() {Me.btnSend, Me.btnSendForApproval})
        Me.btnsetting.Location = New System.Drawing.Point(380, 6)
        Me.btnsetting.Name = "btnsetting"
        Me.btnsetting.Size = New System.Drawing.Size(87, 22)
        Me.btnsetting.TabIndex = 7
        Me.btnsetting.Text = "E-Mail/SMS"
        '
        'btnSend
        '
        Me.btnSend.Name = "btnSend"
        Me.btnSend.Text = "Send Email/SMS"
        '
        'btnSendForApproval
        '
        Me.btnSendForApproval.Name = "btnSendForApproval"
        Me.btnSendForApproval.Text = "Send For Approval"
        '
        'btnPrint
        '
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(305, 6)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(69, 22)
        Me.btnPrint.TabIndex = 4
        Me.btnPrint.Text = "Print"
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(155, 6)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(69, 22)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        '
        'btnPost
        '
        Me.btnPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPost.Location = New System.Drawing.Point(80, 6)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(69, 22)
        Me.btnPost.TabIndex = 1
        Me.btnPost.Text = "Post"
        '
        'btnViewTDSDetails
        '
        Me.btnViewTDSDetails.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnViewTDSDetails.Location = New System.Drawing.Point(230, 6)
        Me.btnViewTDSDetails.Name = "btnViewTDSDetails"
        Me.btnViewTDSDetails.Size = New System.Drawing.Size(69, 22)
        Me.btnViewTDSDetails.TabIndex = 3
        Me.btnViewTDSDetails.Text = " TDS Detail"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(1052, 6)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(69, 22)
        Me.btnClose.TabIndex = 5
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
        'RadMenu1
        '
        Me.RadMenu1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem1})
        Me.RadMenu1.Location = New System.Drawing.Point(0, 0)
        Me.RadMenu1.Name = "RadMenu1"
        Me.RadMenu1.Size = New System.Drawing.Size(1126, 20)
        Me.RadMenu1.TabIndex = 0
        '
        'RadMenuItem1
        '
        Me.RadMenuItem1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem2, Me.RadMenuItem3, Me.RadMenuItem8, Me.RadMenuItem9})
        Me.RadMenuItem1.Name = "RadMenuItem1"
        Me.RadMenuItem1.Text = "Setting"
        '
        'RadMenuItem2
        '
        Me.RadMenuItem2.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem4, Me.mbtnExportApTransaction, Me.rmiExportDrNote, Me.mnuCreditNoteExport})
        Me.RadMenuItem2.Name = "RadMenuItem2"
        Me.RadMenuItem2.Text = "Export"
        '
        'RadMenuItem4
        '
        Me.RadMenuItem4.AccessibleDescription = "Opening Balance"
        Me.RadMenuItem4.AccessibleName = "Opening Balance"
        Me.RadMenuItem4.Name = "RadMenuItem4"
        Me.RadMenuItem4.Text = "Opening Balance (Blank Sheet)"
        '
        'mbtnExportApTransaction
        '
        Me.mbtnExportApTransaction.AccessibleDescription = "Export AP Transaction Blank"
        Me.mbtnExportApTransaction.AccessibleName = "Export AP Transaction Blank"
        Me.mbtnExportApTransaction.Name = "mbtnExportApTransaction"
        Me.mbtnExportApTransaction.Text = "AP Transactions Blank"
        '
        'rmiExportDrNote
        '
        Me.rmiExportDrNote.Name = "rmiExportDrNote"
        Me.rmiExportDrNote.Text = "Export [Debit Note]"
        '
        'mnuCreditNoteExport
        '
        Me.mnuCreditNoteExport.Name = "mnuCreditNoteExport"
        Me.mnuCreditNoteExport.Text = "Export [Credit Note]"
        '
        'RadMenuItem3
        '
        Me.RadMenuItem3.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem5, Me.RadMenuItem6, Me.RadMenuItem7, Me.mbtnImportApTransaction, Me.rmiImportDrNote, Me.mnuImportCreditNote, Me.RadMenuItem10})
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
        Me.RadMenuItem6.Name = "RadMenuItem6"
        Me.RadMenuItem6.Text = "Credit Note"
        '
        'RadMenuItem7
        '
        Me.RadMenuItem7.Name = "RadMenuItem7"
        Me.RadMenuItem7.Text = "Debit Note"
        '
        'mbtnImportApTransaction
        '
        Me.mbtnImportApTransaction.Name = "mbtnImportApTransaction"
        Me.mbtnImportApTransaction.Text = "AP Transactions"
        '
        'rmiImportDrNote
        '
        Me.rmiImportDrNote.Name = "rmiImportDrNote"
        Me.rmiImportDrNote.Text = "Import [Debit Note]"
        '
        'mnuImportCreditNote
        '
        Me.mnuImportCreditNote.Name = "mnuImportCreditNote"
        Me.mnuImportCreditNote.Text = "Import [Credit Note]"
        '
        'RadMenuItem10
        '
        Me.RadMenuItem10.Name = "RadMenuItem10"
        Me.RadMenuItem10.Text = "Venodr Invoice(Pivot of Date)"
        '
        'RadMenuItem8
        '
        Me.RadMenuItem8.Name = "RadMenuItem8"
        Me.RadMenuItem8.Text = "Save Layout"
        '
        'RadMenuItem9
        '
        Me.RadMenuItem9.Name = "RadMenuItem9"
        Me.RadMenuItem9.Text = "Delete Layout"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.SplitContainer1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 20)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1126, 530)
        Me.Panel1.TabIndex = 3
        '
        'FrmAPInvoiceEntry
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1126, 550)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.RadMenu1)
        Me.MinimumSize = New System.Drawing.Size(890, 467)
        Me.Name = "FrmAPInvoiceEntry"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "AP Invoice Entry"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageView1.ResumeLayout(False)
        Me.RadPageViewPage1.ResumeLayout(False)
        Me.RadPageViewPage1.PerformLayout()
        Me.pnlLocation.ResumeLayout(False)
        Me.pnlLocation.PerformLayout()
        CType(Me.RadLabel23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLocationPrefixName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChkTrnsferToSvng, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpVendorBankDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpVendorBankDetails.ResumeLayout(False)
        Me.grpVendorBankDetails.PerformLayout()
        CType(Me.txtVendor_Bank_ACNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBankCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtIFSCCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtbranchname, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtBankName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSecurityAdjusment, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkTDSProvision, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkRCM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.lblGstinNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel49, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel45, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRegisterOrUnregister, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox4.ResumeLayout(False)
        Me.RadGroupBox4.PerformLayout()
        CType(Me.CboxITCCateogory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CboxITCType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkITCEligible, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox2.ResumeLayout(False)
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlDeduction.ResumeLayout(False)
        Me.pnlDeduction.PerformLayout()
        Me.pnlSecondaryTranporterPaymentCycle.ResumeLayout(False)
        Me.pnlSecondaryTranporterPaymentCycle.PerformLayout()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtToDateSecTrans, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFromDateSecTrans, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ChkSecurity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSecTrans, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel21, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadLabel21.ResumeLayout(False)
        CType(Me.LblCostCentre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel20, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPCJ.ResumeLayout(False)
        Me.pnlPCJ.PerformLayout()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblProject, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDrillDown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkOnHold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtMCC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMCC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMCC2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDataAndTimeSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTapalNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ddlEmployeeType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTotalAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblcaption2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtBillDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCaption1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBillNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkProRated, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAdd_Doc_TYpe, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LblVCGL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LblHirerachy, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkDeduction, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LblLocDesp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkProvision, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkQuickMode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVendorName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotRAmt1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbRefType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVendorInvDatre, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVendorInvoiceNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtOrderNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPONo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboDocType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDesc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpProvision.ResumeLayout(False)
        Me.grpProvision.PerformLayout()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtProvAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpToProv, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpFromProv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage2.ResumeLayout(False)
        Me.RadPageViewPage2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.rbtnTaxCalManual, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnTaxCalAutomatic, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTaxGrpName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.RadLabel16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTermName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDueDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv2.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage3.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.gvAC.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gvAC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAddCharges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Attachments.ResumeLayout(False)
        Me.RadPageViewPage4.ResumeLayout(False)
        Me.RadPageViewPage4.PerformLayout()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRoundOff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtCashDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAmtAfterCashDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAmtAfterCashDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCashDiscount, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCurrConv.ResumeLayout(False)
        Me.pnlCurrConv.PerformLayout()
        CType(Me.txtConversionRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblEffectiveFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtApplicableFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCurrency, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblConvRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnReverse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLandedAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotEmptyAmt, System.ComponentModel.ISupportInitialize).EndInit()
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
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.butCostCenterAndHirerachy_Update_AfterPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnsetting, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnViewTDSDetails, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
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
    Friend WithEvents txtPONo As common.Controls.MyTextBox
    Friend WithEvents cboDocType As common.Controls.MyComboBox
    Friend WithEvents gv1 As common.UserControls.MyRadGridView
    Friend WithEvents gv2 As common.UserControls.MyRadGridView
    Friend WithEvents btnClose As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnSave As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtVendorInvoiceNo As common.Controls.MyTextBox
    Friend WithEvents txtDueDate As common.Controls.MyDateTimePicker
    Friend WithEvents RadGroupBox1 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents RadGroupBox2 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents btnViewTDSDetails As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnPost As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnDelete As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtVendorInvDatre As common.Controls.MyDateTimePicker
    Friend WithEvents txtDocNo As common.UserControls.txtNavigator
    Friend WithEvents btnPrint As Telerik.WinControls.UI.RadButton
    Friend WithEvents cmbRefType As common.Controls.MyComboBox
    Friend WithEvents UsLock1 As common.usLock
    Friend WithEvents txtRefDocNo As common.UserControls.txtFinder
    Friend WithEvents TxtVendorNo As common.UserControls.txtFinder
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
    Friend WithEvents RadLabel7 As common.Controls.MyLabel
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
    Friend WithEvents RadLabel3 As common.Controls.MyLabel
    Friend WithEvents RadLabel17 As common.Controls.MyLabel
    Friend WithEvents lblTermName As common.Controls.MyLabel
    Friend WithEvents RadLabel16 As common.Controls.MyLabel
    Friend WithEvents RadLabel9 As common.Controls.MyLabel
    Friend WithEvents RadLabel10 As common.Controls.MyLabel
    Friend WithEvents RadLabel12 As common.Controls.MyLabel
    Friend WithEvents lblTotRAmt1 As common.Controls.MyLabel
    Friend WithEvents RadLabel13 As common.Controls.MyLabel
    Friend WithEvents RadLabel14 As common.Controls.MyLabel
    Friend WithEvents RadLabel15 As common.Controls.MyLabel
    Friend WithEvents RadLabel31 As common.Controls.MyLabel
    Friend WithEvents lblAddCharges As common.Controls.MyLabel
    Friend WithEvents RadLabel32 As common.Controls.MyLabel
    Friend WithEvents lblAddCharges1 As common.Controls.MyLabel
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents chkQuickMode As common.Controls.MyCheckBox
    Friend WithEvents RadMenu1 As Telerik.WinControls.UI.RadMenu
    Friend WithEvents RadMenuItem1 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem2 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem4 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem3 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem5 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadMenuItem6 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem7 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents txtlocation As common.UserControls.txtFinder
    Friend WithEvents RadLabel18 As Telerik.WinControls.UI.RadLabel
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents lblTotEmptyAmt As common.Controls.MyLabel
    Friend WithEvents mbtnExportApTransaction As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mbtnImportApTransaction As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem8 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem9 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents MyLabel3 As common.Controls.MyLabel
    Friend WithEvents lblLandedAmt As common.Controls.MyLabel
    Friend WithEvents btnReverse As Telerik.WinControls.UI.RadButton
    Friend WithEvents pnlCurrConv As System.Windows.Forms.Panel
    Friend WithEvents txtCurrencyCode As common.UserControls.txtFinder
    Friend WithEvents lblEffectiveFrom As common.Controls.MyLabel
    Friend WithEvents txtApplicableFrom As common.Controls.MyLabel
    Friend WithEvents lblCurrency As common.Controls.MyLabel
    Friend WithEvents lblConvRate As common.Controls.MyLabel
    Friend WithEvents pvpCustomFields As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents txtConversionRate As common.MyNumBox
    Friend WithEvents Attachments As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents UcAttachment1 As XpertERPEngine.ucAttachment
    Friend WithEvents pnlPCJ As System.Windows.Forms.Panel
    Friend WithEvents MyLabel4 As common.Controls.MyLabel
    Friend WithEvents fndProject As common.UserControls.txtFinder
    Friend WithEvents lblProject As common.Controls.MyLabel
    Friend WithEvents btnsetting As Telerik.WinControls.UI.RadSplitButton
    Friend WithEvents btnSend As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents btnSendForApproval As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents LblLocDesp As common.Controls.MyLabel
    Friend WithEvents grpProvision As System.Windows.Forms.GroupBox
    Friend WithEvents chkProvision As common.Controls.MyCheckBox
    Friend WithEvents MyLabel6 As common.Controls.MyLabel
    Friend WithEvents dtpToProv As common.Controls.MyDateTimePicker
    Friend WithEvents MyLabel5 As common.Controls.MyLabel
    Friend WithEvents dtpFromProv As common.Controls.MyDateTimePicker
    Friend WithEvents btnProvSelect As System.Windows.Forms.LinkLabel
    Friend WithEvents MyLabel7 As common.Controls.MyLabel
    Friend WithEvents txtProvAmt As common.Controls.MyTextBox
    Friend WithEvents btlShowProvision As System.Windows.Forms.LinkLabel
    Friend WithEvents ChkSecurity As common.Controls.MyCheckBox
    Friend WithEvents chkDeduction As common.Controls.MyCheckBox
    Friend WithEvents chkProRated As common.Controls.MyCheckBox
    Friend WithEvents rmiExportDrNote As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents rmiImportDrNote As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents MyLabel8 As common.Controls.MyLabel
    Friend WithEvents LblHirerachy As common.Controls.MyLabel
    Friend WithEvents TxtHirerachy As common.UserControls.txtFinder
    Friend WithEvents RadLabel20 As Telerik.WinControls.UI.RadLabel
    Friend WithEvents LblCostCentre As common.Controls.MyLabel
    Friend WithEvents TxtCostCentre As common.UserControls.txtFinder
    Friend WithEvents RadLabel21 As Telerik.WinControls.UI.RadLabel
    Friend WithEvents mnuCreditNoteExport As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents mnuImportCreditNote As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents UcCustomFields1 As XpertERPEngine.ucCustomFields
    Friend WithEvents LblVCGL As common.Controls.MyTextBox
    Friend WithEvents txtAmtAfterCashDiscount As common.Controls.MyLabel
    Friend WithEvents lblAmtAfterCashDiscount As common.Controls.MyLabel
    Friend WithEvents lblCashDiscount As common.Controls.MyLabel
    Friend WithEvents TxtCashDiscount As common.MyNumBox
    Friend WithEvents RadMenuItem10 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents MyLabel9 As common.Controls.MyLabel
    Friend WithEvents txtAdd_Doc_TYpe As common.Controls.MyLabel
    Friend WithEvents pnlDeduction As System.Windows.Forms.Panel
    Friend WithEvents chkSecTrans As common.Controls.MyCheckBox
    Friend WithEvents MyLabel11 As common.Controls.MyLabel
    Friend WithEvents txtToDateSecTrans As common.Controls.MyDateTimePicker
    Friend WithEvents MyLabel10 As common.Controls.MyLabel
    Friend WithEvents txtFromDateSecTrans As common.Controls.MyDateTimePicker
    Friend WithEvents pnlSecondaryTranporterPaymentCycle As System.Windows.Forms.Panel
    Friend WithEvents lblCaption1 As common.Controls.MyLabel
    Friend WithEvents txtBillNo As common.Controls.MyTextBox
    Friend WithEvents lblcaption2 As common.Controls.MyLabel
    Friend WithEvents dtBillDate As common.Controls.MyDateTimePicker
    Friend WithEvents txtTotalAmt As common.MyNumBox
    Friend WithEvents RadGroupBox4 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents CboxITCCateogory As common.Controls.MyComboBox
    Friend WithEvents MyLabel12 As common.Controls.MyLabel
    Friend WithEvents CboxITCType As common.Controls.MyComboBox
    Friend WithEvents chkITCEligible As Telerik.WinControls.UI.RadCheckBox
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents ddlEmployeeType As common.Controls.MyComboBox
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents lblGstinNo As common.Controls.MyLabel
    Friend WithEvents MyLabel49 As common.Controls.MyLabel
    Friend WithEvents MyLabel45 As common.Controls.MyLabel
    Friend WithEvents lblRegisterOrUnregister As common.Controls.MyLabel
    Friend WithEvents MyLabel14 As common.Controls.MyLabel
    Friend WithEvents lblRoundOff As common.Controls.MyLabel
    Friend WithEvents MyLabel15 As common.Controls.MyLabel
    Friend WithEvents txtDataAndTimeSelection As common.Controls.MyDateTimePicker
    Friend WithEvents MyLabel16 As common.Controls.MyLabel
    Friend WithEvents txtTapalNo As common.Controls.MyTextBox
    Friend WithEvents txtMCC As common.Controls.MyTextBox
    Friend WithEvents lblMCC As common.Controls.MyTextBox
    Friend WithEvents lblMCC2 As common.Controls.MyLabel
    Friend WithEvents grpVendorBankDetails As RadGroupBox
    Friend WithEvents TxtIFSCCode As common.Controls.MyTextBox
    Friend WithEvents MyLabel22 As common.Controls.MyLabel
    Friend WithEvents txtbranchname As common.Controls.MyTextBox
    Friend WithEvents TxtBankName As common.Controls.MyTextBox
    Friend WithEvents MyLabel23 As common.Controls.MyLabel
    Friend WithEvents txtBankCode As common.Controls.MyTextBox
    Friend WithEvents txtVendor_Bank_ACNo As common.Controls.MyTextBox
    Friend WithEvents MyLabel17 As common.Controls.MyLabel
    Friend WithEvents butCostCenterAndHirerachy_Update_AfterPost As RadButton
    Friend WithEvents chkRCM As common.Controls.MyCheckBox
    Friend WithEvents chkTDSProvision As common.Controls.MyCheckBox
    Friend WithEvents chkSecurityAdjusment As common.Controls.MyCheckBox
    Friend WithEvents txtSecurityAdjusment As common.UserControls.txtMultiSelectFinder
    Friend WithEvents ChkTrnsferToSvng As common.Controls.MyCheckBox
    Friend WithEvents btnHistory As RadButton
    Friend WithEvents pnlLocation As Panel
    Friend WithEvents RadLabel23 As RadLabel
    Friend WithEvents txtLocationPrefixName As common.Controls.MyLabel
    Friend WithEvents txtLocationPrefix As common.UserControls.txtFinder
End Class

