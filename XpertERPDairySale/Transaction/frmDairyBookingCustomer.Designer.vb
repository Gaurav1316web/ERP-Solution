<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDairyBookingCustomer
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
        Dim RadListDataItem5 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim RadListDataItem6 As Telerik.WinControls.UI.RadListDataItem = New Telerik.WinControls.UI.RadListDataItem()
        Dim TableViewDefinition1 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Dim TableViewDefinition2 As Telerik.WinControls.UI.TableViewDefinition = New Telerik.WinControls.UI.TableViewDefinition()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.RadPageView1 = New Telerik.WinControls.UI.RadPageView()
        Me.RadPageViewPage1 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.txtMobileNo = New common.Controls.MyTextBox()
        Me.MyLabel11 = New common.Controls.MyLabel()
        Me.lblCategory = New common.Controls.MyLabel()
        Me.txtCategory = New common.UserControls.txtFinder()
        Me.lblRouteNo = New common.Controls.MyLabel()
        Me.MyLabel21 = New common.Controls.MyLabel()
        Me.txtAdharNo = New common.Controls.MyTextBox()
        Me.lblReceiptAmt = New common.Controls.MyLabel()
        Me.lblReceiptAmtDesc = New common.Controls.MyLabel()
        Me.lblReceipt = New common.Controls.MyLabel()
        Me.txtReceipt = New common.UserControls.txtFinder()
        Me.chkNoTranspoter = New Telerik.WinControls.UI.RadCheckBox()
        Me.chkAPS = New Telerik.WinControls.UI.RadCheckBox()
        Me.txtIndentDate = New common.Controls.MyDateTimePicker()
        Me.RadLabel4 = New common.Controls.MyLabel()
        Me.UsLock1 = New common.usLock()
        Me.lblIndentDate = New common.Controls.MyLabel()
        Me.chkExcludeKKFMND = New Telerik.WinControls.UI.RadCheckBox()
        Me.txtIndentNo = New common.Controls.MyTextBox()
        Me.txtSupplyDate = New common.Controls.MyDateTimePicker()
        Me.lblSupplyDate = New common.Controls.MyLabel()
        Me.lblIndentNo = New common.Controls.MyLabel()
        Me.chkIsEwayBill = New Telerik.WinControls.UI.RadCheckBox()
        Me.chkNoCrateIssue = New Telerik.WinControls.UI.RadCheckBox()
        Me.chkisTCS = New Telerik.WinControls.UI.RadCheckBox()
        Me.chkTPT = New Telerik.WinControls.UI.RadCheckBox()
        Me.lblShipToLocation = New common.Controls.MyLabel()
        Me.txtShipToLocation = New common.UserControls.txtFinder()
        Me.RadLabel18 = New common.Controls.MyLabel()
        Me.chkManualVehicle = New Telerik.WinControls.UI.RadCheckBox()
        Me.MyLabel31 = New common.Controls.MyLabel()
        Me.lblTransporter = New common.Controls.MyLabel()
        Me.lblVehicleNo = New common.Controls.MyLabel()
        Me.txtManualVehicle = New common.Controls.MyTextBox()
        Me.fndTransporter = New common.UserControls.txtFinder()
        Me.MyLabel20 = New common.Controls.MyLabel()
        Me.txtDescription = New common.Controls.MyTextBox()
        Me.lblReceiverName = New common.Controls.MyLabel()
        Me.txtReceiverName = New common.Controls.MyTextBox()
        Me.lblChequeNo = New common.Controls.MyLabel()
        Me.txtChequeNo = New common.Controls.MyTextBox()
        Me.cmbPaymentType = New common.Controls.MyComboBox()
        Me.MyLabel7 = New common.Controls.MyLabel()
        Me.lblPaymentType = New common.Controls.MyLabel()
        Me.chkcashsale = New Telerik.WinControls.UI.RadCheckBox()
        Me.chkGhee = New Telerik.WinControls.UI.RadCheckBox()
        Me.rgbTaxNonTax = New Telerik.WinControls.UI.RadGroupBox()
        Me.lblFATPER = New common.Controls.MyLabel()
        Me.txtMBRTHours = New common.Controls.MyTextBox()
        Me.txtFATPER = New common.Controls.MyTextBox()
        Me.lblMBRT = New common.Controls.MyLabel()
        Me.lblSNFPER = New common.Controls.MyLabel()
        Me.txtTemp = New common.Controls.MyTextBox()
        Me.txtSNFPER = New common.Controls.MyTextBox()
        Me.lblTEMP = New common.Controls.MyLabel()
        Me.lblAcidity = New common.Controls.MyLabel()
        Me.txtAcidity = New common.Controls.MyTextBox()
        Me.lblSubLocation = New common.Controls.MyLabel()
        Me.MyLabel60 = New common.Controls.MyLabel()
        Me.txtSubLocation = New common.UserControls.txtFinder()
        Me.chkDistributor = New Telerik.WinControls.UI.RadCheckBox()
        Me.cmbGatePassType = New common.Controls.MyComboBox()
        Me.lblGatePassType = New common.Controls.MyLabel()
        Me.lblShiftType = New Telerik.WinControls.UI.RadLabel()
        Me.txtCouponDate = New common.Controls.MyDateTimePicker()
        Me.lblCouponDate = New common.Controls.MyLabel()
        Me.txtBPLRemark = New common.Controls.MyTextBox()
        Me.txtBPLName = New common.Controls.MyTextBox()
        Me.txtCouponCode = New common.Controls.MyTextBox()
        Me.lblBPLRemark = New common.Controls.MyLabel()
        Me.lblBPLName = New common.Controls.MyLabel()
        Me.lblCouponCode = New common.Controls.MyLabel()
        Me.chkBPL = New Telerik.WinControls.UI.RadCheckBox()
        Me.rgbItemType = New Telerik.WinControls.UI.RadGroupBox()
        Me.rbtnNonTax = New Telerik.WinControls.UI.RadRadioButton()
        Me.rbtnTaxable = New Telerik.WinControls.UI.RadRadioButton()
        Me.txtVehicleCode = New common.UserControls.txtFinder()
        Me.txtVehicleName = New common.Controls.MyTextBox()
        Me.txtRouteName1 = New common.Controls.MyTextBox()
        Me.txtRouteCode1 = New common.Controls.MyTextBox()
        Me.btnCC = New Telerik.WinControls.UI.RadButton()
        Me.txtDate = New common.Controls.MyDateTimePicker()
        Me.cmbcashcredit = New common.Controls.MyComboBox()
        Me.lblCredit = New common.Controls.MyLabel()
        Me.lblOutstandingDesc = New common.Controls.MyLabel()
        Me.lblOutStanding = New common.Controls.MyLabel()
        Me.lblRoute1 = New common.Controls.MyLabel()
        Me.lblPriceCode1 = New common.Controls.MyLabel()
        Me.lblPriceCodeDesc = New common.Controls.MyLabel()
        Me.lblVehicle1 = New common.Controls.MyLabel()
        Me.chkSampling = New Telerik.WinControls.UI.RadCheckBox()
        Me.lblLocation = New common.Controls.MyLabel()
        Me.txtLocation = New common.UserControls.txtFinder()
        Me.RadLabel15 = New common.Controls.MyLabel()
        Me.lblRouteDesc = New common.Controls.MyLabel()
        Me.lblAbandonmentNo = New common.Controls.MyLabel()
        Me.RadGroupBox2 = New Telerik.WinControls.UI.RadGroupBox()
        Me.gv1 = New common.UserControls.MyRadGridView()
        Me.lblVendorName = New common.Controls.MyLabel()
        Me.txtRouteNo = New common.UserControls.txtFinder()
        Me.RadLabel2 = New common.Controls.MyLabel()
        Me.RadLabel1 = New common.Controls.MyLabel()
        Me.txtVendorNo = New common.UserControls.txtFinder()
        Me.txtDocNo = New common.UserControls.txtNavigator()
        Me.btnAddNew = New Telerik.WinControls.UI.RadButton()
        Me.RadPageViewPage6 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.MyLabel19 = New common.Controls.MyLabel()
        Me.RadGroupBox1 = New Telerik.WinControls.UI.RadGroupBox()
        Me.txtTermCode = New common.UserControls.txtFinder()
        Me.RadLabel16 = New common.Controls.MyLabel()
        Me.lblTermName = New common.Controls.MyLabel()
        Me.txtDueDate = New common.Controls.MyDateTimePicker()
        Me.RadLabel17 = New common.Controls.MyLabel()
        Me.gv2 = New common.UserControls.MyRadGridView()
        Me.txtSchemeTaxGroup = New common.UserControls.txtFinder()
        Me.MyLabel44 = New common.Controls.MyLabel()
        Me.lblTaxGroupScheme = New common.Controls.MyLabel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.rbtnTaxCalManual = New common.Controls.MyRadioButton()
        Me.rbtnTaxCalAutomatic = New common.Controls.MyRadioButton()
        Me.txtTaxGroup = New common.UserControls.txtFinder()
        Me.RadLabel11 = New common.Controls.MyLabel()
        Me.lblTaxGrpName = New common.Controls.MyLabel()
        Me.RadPageViewPage2 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.RadPageViewPage4 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.MyLabel71 = New common.Controls.MyLabel()
        Me.lblGrossAmount = New common.Controls.MyLabel()
        Me.lblTPTComAmt = New common.Controls.MyLabel()
        Me.txtTPTComAmt = New common.Controls.MyLabel()
        Me.lblRoundOff = New common.Controls.MyLabel()
        Me.btnUpdateVehicle = New Telerik.WinControls.UI.RadButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btn_QtyReset = New Telerik.WinControls.UI.RadButton()
        Me.btnHistory = New Telerik.WinControls.UI.RadButton()
        Me.BtnRecieptEntry = New Telerik.WinControls.UI.RadButton()
        Me.txtTCAmt = New common.Controls.MyLabel()
        Me.lblTCAmt = New common.Controls.MyLabel()
        Me.txtSecurity = New common.Controls.MyLabel()
        Me.lblSecuirty = New common.Controls.MyLabel()
        Me.txtDCAmt = New common.Controls.MyLabel()
        Me.lblDCAmt = New common.Controls.MyLabel()
        Me.TxtRoundoff = New common.Controls.MyLabel()
        Me.txtTCSTaxRate = New common.MyNumBox()
        Me.MyLabel57 = New common.Controls.MyLabel()
        Me.lblActualTCSTaxBaseAmt = New common.Controls.MyLabel()
        Me.MyLabel58 = New common.Controls.MyLabel()
        Me.txttcstaxbaseamount = New common.MyNumBox()
        Me.lblInvoiceDiscAmt = New common.Controls.MyLabel()
        Me.MyLabel6 = New common.Controls.MyLabel()
        Me.MyLabel3 = New common.Controls.MyLabel()
        Me.RadGroupBox3 = New Telerik.WinControls.UI.RadGroupBox()
        Me.chkDiscountOnAmt = New Telerik.WinControls.UI.RadRadioButton()
        Me.chkDiscountOnRate = New Telerik.WinControls.UI.RadRadioButton()
        Me.txtDiscAmt = New common.MyNumBox()
        Me.txtDiscPer = New common.MyNumBox()
        Me.MyLabel1 = New common.Controls.MyLabel()
        Me.pnlCurrConv = New System.Windows.Forms.Panel()
        Me.txtConversionRate = New common.MyNumBox()
        Me.txtCurrencyCode = New common.UserControls.txtFinder()
        Me.lblEffectiveFrom = New common.Controls.MyLabel()
        Me.txtApplicableFrom = New common.Controls.MyLabel()
        Me.lblCurrency = New common.Controls.MyLabel()
        Me.lblConvRate = New common.Controls.MyLabel()
        Me.RadLabel32 = New common.Controls.MyLabel()
        Me.lblAddCharges1 = New common.Controls.MyLabel()
        Me.txtComment = New System.Windows.Forms.TextBox()
        Me.RadLabel14 = New common.Controls.MyLabel()
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
        Me.RadPageViewPage3 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.lblLastCollectionDate = New common.Controls.MyLabel()
        Me.txtLastCollectionDate = New common.Controls.MyLabel()
        Me.MyLabel5 = New common.Controls.MyLabel()
        Me.lblTotRAmt1 = New common.Controls.MyLabel()
        Me.lblDCSDemand = New common.Controls.MyLabel()
        Me.txtDCSDemandNo = New common.Controls.MyLabel()
        Me.lblUnbilledMilk = New common.Controls.MyLabel()
        Me.lblUnbilledMilkAmt = New common.Controls.MyLabel()
        Me.pnlTCS = New Telerik.WinControls.UI.RadPanel()
        Me.lblTCSBaseAmt = New common.Controls.MyLabel()
        Me.txtTCSBaseAmt = New common.Controls.MyLabel()
        Me.MyLabel17 = New common.Controls.MyLabel()
        Me.lblTCSAmount = New common.Controls.MyLabel()
        Me.RadPanel2 = New Telerik.WinControls.UI.RadPanel()
        Me.LblBox = New common.Controls.MyLabel()
        Me.txtBox = New common.Controls.MyLabel()
        Me.lblCrate = New common.Controls.MyLabel()
        Me.txtCrate = New common.Controls.MyLabel()
        Me.lblCan = New common.Controls.MyLabel()
        Me.txtCan = New common.Controls.MyLabel()
        Me.chkDCS = New Telerik.WinControls.UI.RadCheckBox()
        Me.btnGatePassPrint = New Telerik.WinControls.UI.RadButton()
        Me.btnPrint = New Telerik.WinControls.UI.RadButton()
        Me.btnCreateDO = New Telerik.WinControls.UI.RadButton()
        Me.txtSalesman1 = New common.UserControls.txtFinder()
        Me.lblSalesmandesc1 = New common.Controls.MyLabel()
        Me.lblsalesman1 = New common.Controls.MyLabel()
        Me.txtSalesman = New common.UserControls.txtFinder()
        Me.lblSalesman = New common.Controls.MyLabel()
        Me.MyLabel2 = New common.Controls.MyLabel()
        Me.lblBoothStation = New common.Controls.MyLabel()
        Me.txtEx_Factory_Date = New common.Controls.MyDateTimePicker()
        Me.lbl_ExFactoryDate = New common.Controls.MyLabel()
        Me.MyLabel12 = New common.Controls.MyLabel()
        Me.txtPONo = New common.Controls.MyTextBox()
        Me.MyLabel4 = New common.Controls.MyLabel()
        Me.txtCustPODate = New common.Controls.MyDateTimePicker()
        Me.chkGatePass = New Telerik.WinControls.UI.RadCheckBox()
        Me.ItemTypePanel = New System.Windows.Forms.Panel()
        Me.rbtn_Ambient = New Telerik.WinControls.UI.RadRadioButton()
        Me.rbtn_Fresh = New Telerik.WinControls.UI.RadRadioButton()
        Me.MyLabel40 = New common.Controls.MyLabel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.MyLabel25 = New common.Controls.MyLabel()
        Me.lblARSecurity = New common.Controls.MyLabel()
        Me.MyLabel24 = New common.Controls.MyLabel()
        Me.lblReverseAdvanceSec = New common.Controls.MyLabel()
        Me.MyLabel23 = New common.Controls.MyLabel()
        Me.lblReverseRefund = New common.Controls.MyLabel()
        Me.MyLabel35 = New common.Controls.MyLabel()
        Me.lblTotalOutstansing = New common.Controls.MyLabel()
        Me.MyLabel33 = New common.Controls.MyLabel()
        Me.lblRefund = New common.Controls.MyLabel()
        Me.MyLabel30 = New common.Controls.MyLabel()
        Me.lblLedgerOutstanding = New common.Controls.MyLabel()
        Me.MyLabel28 = New common.Controls.MyLabel()
        Me.lblShortcloseDO = New common.Controls.MyLabel()
        Me.MyLabel26 = New common.Controls.MyLabel()
        Me.lblPendingDO = New common.Controls.MyLabel()
        Me.MyLabel22 = New common.Controls.MyLabel()
        Me.lblAdvanceSecurity = New common.Controls.MyLabel()
        Me.MyLabel15 = New common.Controls.MyLabel()
        Me.lblCreditLimit = New common.Controls.MyLabel()
        Me.PanelSearchItem = New System.Windows.Forms.Panel()
        Me.btnSerach = New Telerik.WinControls.UI.RadButton()
        Me.MyLabel18 = New common.Controls.MyLabel()
        Me.txtItemSearch = New common.Controls.MyTextBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.lbltotalOutstanding1 = New common.Controls.MyLabel()
        Me.lblTotalSecurity11 = New common.Controls.MyLabel()
        Me.MyLabel16 = New common.Controls.MyLabel()
        Me.MyLabel14 = New common.Controls.MyLabel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtBOstatus = New common.Controls.MyLabel()
        Me.txtDOStatus = New common.Controls.MyLabel()
        Me.lblLoginUserZone = New common.Controls.MyLabel()
        Me.MyLabel13 = New common.Controls.MyLabel()
        Me.MyLabel9 = New common.Controls.MyLabel()
        Me.lblroute = New common.Controls.MyLabel()
        Me.lblroutename = New common.Controls.MyLabel()
        Me.lblroutecode = New common.Controls.MyLabel()
        Me.lblPriceCode = New common.Controls.MyLabel()
        Me.txtPriceCode = New common.Controls.MyLabel()
        Me.MyLabel29 = New common.Controls.MyLabel()
        Me.MyLabel27 = New common.Controls.MyLabel()
        Me.lblUploadingDate = New common.Controls.MyLabel()
        Me.lblCreatedDateAndTime = New common.Controls.MyLabel()
        Me.lblvehicle = New common.Controls.MyLabel()
        Me.MyLabel10 = New common.Controls.MyLabel()
        Me.LblUpdatedVehicleDesc = New common.Controls.MyLabel()
        Me.lblvehiclecode = New common.Controls.MyLabel()
        Me.LblUpdatedVehicleCode = New common.Controls.MyLabel()
        Me.lblvehicleName = New common.Controls.MyLabel()
        Me.lblDO = New common.Controls.MyLabel()
        Me.lblDONumber = New common.Controls.MyLabel()
        Me.MyLabel8 = New common.Controls.MyLabel()
        Me.lblCancelStatus = New common.Controls.MyLabel()
        Me.RadPageViewPage5 = New Telerik.WinControls.UI.RadPageViewPage()
        Me.UcAttachment1 = New XpertERPEngine.ucAttachment()
        Me.btnCreateEWB = New Telerik.WinControls.UI.RadButton()
        Me.rpOpening = New Telerik.WinControls.UI.RadPanel()
        Me.lblDocAmt = New common.Controls.MyLabel()
        Me.lblTotalDocAmt = New common.Controls.MyLabel()
        Me.lblCrAmt = New common.Controls.MyLabel()
        Me.txtCrAmt = New common.Controls.MyLabel()
        Me.lblNetbalance = New common.Controls.MyLabel()
        Me.txtClosingBal = New common.Controls.MyLabel()
        Me.lblOpeningBal = New common.Controls.MyLabel()
        Me.txtOpeningbal = New common.Controls.MyLabel()
        Me.lblDrAmt = New common.Controls.MyLabel()
        Me.txtDrAmt = New common.Controls.MyLabel()
        Me.btnprinte_wayBill = New Telerik.WinControls.UI.RadButton()
        Me.RadLabel10 = New common.Controls.MyLabel()
        Me.btnPrintChallan = New Telerik.WinControls.UI.RadButton()
        Me.btnGatepass = New Telerik.WinControls.UI.RadButton()
        Me.lblCreatedByValue = New common.Controls.MyLabel()
        Me.lblCreatedBy = New common.Controls.MyLabel()
        Me.btnCreateAndPrintInvoice = New Telerik.WinControls.UI.RadButton()
        Me.btnreverse = New Telerik.WinControls.UI.RadButton()
        Me.btnCancel = New Telerik.WinControls.UI.RadButton()
        Me.btnDelete = New Telerik.WinControls.UI.RadButton()
        Me.btnPost = New Telerik.WinControls.UI.RadButton()
        Me.btnClose = New Telerik.WinControls.UI.RadButton()
        Me.btnSave = New Telerik.WinControls.UI.RadButton()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.RadMenu1 = New Telerik.WinControls.UI.RadMenu()
        Me.RadMenuItem2 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem3 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem1 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem4 = New Telerik.WinControls.UI.RadMenuItem()
        Me.RadMenuItem5 = New Telerik.WinControls.UI.RadMenuItem()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageView1.SuspendLayout()
        Me.RadPageViewPage1.SuspendLayout()
        CType(Me.txtMobileNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCategory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRouteNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAdharNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReceiptAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReceiptAmtDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReceipt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkNoTranspoter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkAPS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtIndentDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblIndentDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkExcludeKKFMND, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtIndentNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSupplyDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSupplyDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblIndentNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkIsEwayBill, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkNoCrateIssue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkisTCS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkTPT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblShipToLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkManualVehicle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTransporter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVehicleNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtManualVehicle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel20, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReceiverName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtReceiverName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblChequeNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtChequeNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbPaymentType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPaymentType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkcashsale, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkGhee, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rgbTaxNonTax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.rgbTaxNonTax.SuspendLayout()
        CType(Me.lblFATPER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtMBRTHours, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtFATPER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblMBRT, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSNFPER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSNFPER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTEMP, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAcidity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtAcidity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSubLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel60, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkDistributor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbGatePassType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblGatePassType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblShiftType, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCouponDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCouponDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBPLRemark, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBPLName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCouponCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblBPLRemark, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblBPLName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCouponCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkBPL, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rgbItemType, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.rgbItemType.SuspendLayout()
        CType(Me.rbtnNonTax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnTaxable, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtVehicleName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRouteName1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRouteCode1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbcashcredit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCredit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblOutstandingDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblOutStanding, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRoute1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPriceCode1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPriceCodeDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVehicle1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkSampling, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLocation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRouteDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAbandonmentNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox2.SuspendLayout()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblVendorName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage6.SuspendLayout()
        CType(Me.MyLabel19, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox1.SuspendLayout()
        CType(Me.RadLabel16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTermName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDueDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gv2.MasterTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel44, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTaxGroupScheme, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.rbtnTaxCalManual, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtnTaxCalAutomatic, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTaxGrpName, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage4.SuspendLayout()
        CType(Me.MyLabel71, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblGrossAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTPTComAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTPTComAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRoundOff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnUpdateVehicle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.btn_QtyReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BtnRecieptEntry, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTCAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTCAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtSecurity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSecuirty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDCAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDCAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TxtRoundoff, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTCSTaxRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel57, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblActualTCSTaxBaseAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel58, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txttcstaxbaseamount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblInvoiceDiscAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadGroupBox3.SuspendLayout()
        CType(Me.chkDiscountOnAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkDiscountOnRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDiscAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDiscPer, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCurrConv.SuspendLayout()
        CType(Me.txtConversionRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblEffectiveFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtApplicableFrom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCurrency, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblConvRate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAddCharges1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.RadPageViewPage3.SuspendLayout()
        CType(Me.lblLastCollectionDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtLastCollectionDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotRAmt1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDCSDemand, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDCSDemandNo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnbilledMilk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUnbilledMilkAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pnlTCS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTCS.SuspendLayout()
        CType(Me.lblTCSBaseAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTCSBaseAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTCSAmount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPanel2.SuspendLayout()
        CType(Me.LblBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCrate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCrate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCan, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCan, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkDCS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGatePassPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCreateDO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSalesmandesc1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblsalesman1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblSalesman, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblBoothStation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtEx_Factory_Date, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lbl_ExFactoryDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPONo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCustPODate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chkGatePass, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ItemTypePanel.SuspendLayout()
        CType(Me.rbtn_Ambient, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rbtn_Fresh, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel40, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.MyLabel25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblARSecurity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel24, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReverseAdvanceSec, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblReverseRefund, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel35, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotalOutstansing, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel33, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblRefund, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel30, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLedgerOutstanding, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel28, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblShortcloseDO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPendingDO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblAdvanceSecurity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCreditLimit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelSearchItem.SuspendLayout()
        CType(Me.btnSerach, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtItemSearch, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        CType(Me.lbltotalOutstanding1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotalSecurity11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.txtBOstatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDOStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblLoginUserZone, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblroute, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblroutename, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblroutecode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblPriceCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPriceCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel29, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblUploadingDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCreatedDateAndTime, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblvehicle, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LblUpdatedVehicleDesc, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblvehiclecode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LblUpdatedVehicleCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblvehicleName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDONumber, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCancelStatus, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.RadPageViewPage5.SuspendLayout()
        CType(Me.btnCreateEWB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rpOpening, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.rpOpening.SuspendLayout()
        CType(Me.lblDocAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblTotalDocAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCrAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtCrAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblNetbalance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtClosingBal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblOpeningBal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtOpeningbal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblDrAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDrAmt, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnprinte_wayBill, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RadLabel10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPrintChallan, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnGatepass, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCreatedByValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.lblCreatedBy, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCreateAndPrintInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnreverse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 25)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.RadPageView1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnCreateEWB)
        Me.SplitContainer1.Panel2.Controls.Add(Me.rpOpening)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnprinte_wayBill)
        Me.SplitContainer1.Panel2.Controls.Add(Me.RadLabel10)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPrintChallan)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnGatepass)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblCreatedByValue)
        Me.SplitContainer1.Panel2.Controls.Add(Me.lblCreatedBy)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnCreateAndPrintInvoice)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnreverse)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnCancel)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnDelete)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnPost)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnClose)
        Me.SplitContainer1.Panel2.Controls.Add(Me.btnSave)
        Me.SplitContainer1.Size = New System.Drawing.Size(1281, 523)
        Me.SplitContainer1.SplitterDistance = 455
        Me.SplitContainer1.TabIndex = 1
        '
        'RadPageView1
        '
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage1)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage6)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage2)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage4)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage3)
        Me.RadPageView1.Controls.Add(Me.RadPageViewPage5)
        Me.RadPageView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RadPageView1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageView1.Location = New System.Drawing.Point(0, 0)
        Me.RadPageView1.Name = "RadPageView1"
        Me.RadPageView1.SelectedPage = Me.RadPageViewPage1
        Me.RadPageView1.Size = New System.Drawing.Size(1281, 455)
        Me.RadPageView1.TabIndex = 0
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripButtons = Telerik.WinControls.UI.StripViewButtons.None
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).ItemAlignment = Telerik.WinControls.UI.StripViewItemAlignment.Near
        CType(Me.RadPageView1.GetChildAt(0), Telerik.WinControls.UI.RadPageViewStripElement).StripAlignment = Telerik.WinControls.UI.StripViewAlignment.Top
        '
        'RadPageViewPage1
        '
        Me.RadPageViewPage1.Controls.Add(Me.txtMobileNo)
        Me.RadPageViewPage1.Controls.Add(Me.lblCategory)
        Me.RadPageViewPage1.Controls.Add(Me.txtCategory)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel21)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel11)
        Me.RadPageViewPage1.Controls.Add(Me.txtAdharNo)
        Me.RadPageViewPage1.Controls.Add(Me.lblReceiptAmt)
        Me.RadPageViewPage1.Controls.Add(Me.lblReceiptAmtDesc)
        Me.RadPageViewPage1.Controls.Add(Me.lblReceipt)
        Me.RadPageViewPage1.Controls.Add(Me.txtReceipt)
        Me.RadPageViewPage1.Controls.Add(Me.chkNoTranspoter)
        Me.RadPageViewPage1.Controls.Add(Me.chkAPS)
        Me.RadPageViewPage1.Controls.Add(Me.txtIndentDate)
        Me.RadPageViewPage1.Controls.Add(Me.UsLock1)
        Me.RadPageViewPage1.Controls.Add(Me.lblIndentDate)
        Me.RadPageViewPage1.Controls.Add(Me.chkExcludeKKFMND)
        Me.RadPageViewPage1.Controls.Add(Me.txtIndentNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtSupplyDate)
        Me.RadPageViewPage1.Controls.Add(Me.lblIndentNo)
        Me.RadPageViewPage1.Controls.Add(Me.lblSupplyDate)
        Me.RadPageViewPage1.Controls.Add(Me.chkIsEwayBill)
        Me.RadPageViewPage1.Controls.Add(Me.chkNoCrateIssue)
        Me.RadPageViewPage1.Controls.Add(Me.chkisTCS)
        Me.RadPageViewPage1.Controls.Add(Me.chkTPT)
        Me.RadPageViewPage1.Controls.Add(Me.lblShipToLocation)
        Me.RadPageViewPage1.Controls.Add(Me.txtShipToLocation)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel18)
        Me.RadPageViewPage1.Controls.Add(Me.chkManualVehicle)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel31)
        Me.RadPageViewPage1.Controls.Add(Me.lblTransporter)
        Me.RadPageViewPage1.Controls.Add(Me.lblVehicleNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtManualVehicle)
        Me.RadPageViewPage1.Controls.Add(Me.fndTransporter)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel20)
        Me.RadPageViewPage1.Controls.Add(Me.txtDescription)
        Me.RadPageViewPage1.Controls.Add(Me.lblReceiverName)
        Me.RadPageViewPage1.Controls.Add(Me.txtReceiverName)
        Me.RadPageViewPage1.Controls.Add(Me.lblChequeNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtChequeNo)
        Me.RadPageViewPage1.Controls.Add(Me.cmbPaymentType)
        Me.RadPageViewPage1.Controls.Add(Me.lblPaymentType)
        Me.RadPageViewPage1.Controls.Add(Me.chkcashsale)
        Me.RadPageViewPage1.Controls.Add(Me.chkGhee)
        Me.RadPageViewPage1.Controls.Add(Me.rgbTaxNonTax)
        Me.RadPageViewPage1.Controls.Add(Me.lblSubLocation)
        Me.RadPageViewPage1.Controls.Add(Me.MyLabel60)
        Me.RadPageViewPage1.Controls.Add(Me.txtSubLocation)
        Me.RadPageViewPage1.Controls.Add(Me.chkDistributor)
        Me.RadPageViewPage1.Controls.Add(Me.cmbGatePassType)
        Me.RadPageViewPage1.Controls.Add(Me.lblGatePassType)
        Me.RadPageViewPage1.Controls.Add(Me.lblShiftType)
        Me.RadPageViewPage1.Controls.Add(Me.txtCouponDate)
        Me.RadPageViewPage1.Controls.Add(Me.lblCouponDate)
        Me.RadPageViewPage1.Controls.Add(Me.txtBPLRemark)
        Me.RadPageViewPage1.Controls.Add(Me.txtBPLName)
        Me.RadPageViewPage1.Controls.Add(Me.txtCouponCode)
        Me.RadPageViewPage1.Controls.Add(Me.lblBPLRemark)
        Me.RadPageViewPage1.Controls.Add(Me.lblBPLName)
        Me.RadPageViewPage1.Controls.Add(Me.lblCouponCode)
        Me.RadPageViewPage1.Controls.Add(Me.chkBPL)
        Me.RadPageViewPage1.Controls.Add(Me.rgbItemType)
        Me.RadPageViewPage1.Controls.Add(Me.txtVehicleCode)
        Me.RadPageViewPage1.Controls.Add(Me.txtVehicleName)
        Me.RadPageViewPage1.Controls.Add(Me.txtRouteName1)
        Me.RadPageViewPage1.Controls.Add(Me.txtRouteCode1)
        Me.RadPageViewPage1.Controls.Add(Me.btnCC)
        Me.RadPageViewPage1.Controls.Add(Me.txtDate)
        Me.RadPageViewPage1.Controls.Add(Me.cmbcashcredit)
        Me.RadPageViewPage1.Controls.Add(Me.lblCredit)
        Me.RadPageViewPage1.Controls.Add(Me.lblOutstandingDesc)
        Me.RadPageViewPage1.Controls.Add(Me.lblOutStanding)
        Me.RadPageViewPage1.Controls.Add(Me.lblRoute1)
        Me.RadPageViewPage1.Controls.Add(Me.lblPriceCode1)
        Me.RadPageViewPage1.Controls.Add(Me.lblPriceCodeDesc)
        Me.RadPageViewPage1.Controls.Add(Me.lblVehicle1)
        Me.RadPageViewPage1.Controls.Add(Me.chkSampling)
        Me.RadPageViewPage1.Controls.Add(Me.lblLocation)
        Me.RadPageViewPage1.Controls.Add(Me.txtLocation)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel15)
        Me.RadPageViewPage1.Controls.Add(Me.lblRouteDesc)
        Me.RadPageViewPage1.Controls.Add(Me.lblAbandonmentNo)
        Me.RadPageViewPage1.Controls.Add(Me.lblRouteNo)
        Me.RadPageViewPage1.Controls.Add(Me.RadGroupBox2)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel4)
        Me.RadPageViewPage1.Controls.Add(Me.lblVendorName)
        Me.RadPageViewPage1.Controls.Add(Me.txtRouteNo)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel2)
        Me.RadPageViewPage1.Controls.Add(Me.RadLabel1)
        Me.RadPageViewPage1.Controls.Add(Me.txtVendorNo)
        Me.RadPageViewPage1.Controls.Add(Me.txtDocNo)
        Me.RadPageViewPage1.Controls.Add(Me.btnAddNew)
        Me.RadPageViewPage1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadPageViewPage1.ItemSize = New System.Drawing.SizeF(89.0!, 26.0!)
        Me.RadPageViewPage1.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage1.Name = "RadPageViewPage1"
        Me.RadPageViewPage1.Size = New System.Drawing.Size(1260, 409)
        Me.RadPageViewPage1.Text = "Booking Order"
        '
        'txtMobileNo
        '
        Me.txtMobileNo.CalculationExpression = Nothing
        Me.txtMobileNo.FieldCode = Nothing
        Me.txtMobileNo.FieldDesc = Nothing
        Me.txtMobileNo.FieldMaxLength = 0
        Me.txtMobileNo.FieldName = Nothing
        Me.txtMobileNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMobileNo.isCalculatedField = False
        Me.txtMobileNo.IsSourceFromTable = False
        Me.txtMobileNo.IsSourceFromValueList = False
        Me.txtMobileNo.IsUnique = False
        Me.txtMobileNo.Location = New System.Drawing.Point(1103, 133)
        Me.txtMobileNo.MaxLength = 200
        Me.txtMobileNo.MendatroryField = False
        Me.txtMobileNo.MyLinkLable1 = Me.MyLabel11
        Me.txtMobileNo.MyLinkLable2 = Nothing
        Me.txtMobileNo.Name = "txtMobileNo"
        Me.txtMobileNo.ReferenceFieldDesc = Nothing
        Me.txtMobileNo.ReferenceFieldName = Nothing
        Me.txtMobileNo.ReferenceTableName = Nothing
        Me.txtMobileNo.Size = New System.Drawing.Size(144, 18)
        Me.txtMobileNo.TabIndex = 1607
        '
        'MyLabel11
        '
        Me.MyLabel11.FieldName = Nothing
        Me.MyLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel11.Location = New System.Drawing.Point(231, 181)
        Me.MyLabel11.Name = "MyLabel11"
        Me.MyLabel11.Size = New System.Drawing.Size(63, 16)
        Me.MyLabel11.TabIndex = 1605
        Me.MyLabel11.Text = "Adhaar No."
        '
        'lblCategory
        '
        Me.lblCategory.FieldName = Nothing
        Me.lblCategory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCategory.Location = New System.Drawing.Point(940, 112)
        Me.lblCategory.Name = "lblCategory"
        Me.lblCategory.Size = New System.Drawing.Size(52, 16)
        Me.lblCategory.TabIndex = 1539
        Me.lblCategory.Text = "Category"
        '
        'txtCategory
        '
        Me.txtCategory.CalculationExpression = Nothing
        Me.txtCategory.FieldCode = Nothing
        Me.txtCategory.FieldDesc = Nothing
        Me.txtCategory.FieldMaxLength = 0
        Me.txtCategory.FieldName = Nothing
        Me.txtCategory.isCalculatedField = False
        Me.txtCategory.IsSourceFromTable = False
        Me.txtCategory.IsSourceFromValueList = False
        Me.txtCategory.IsUnique = False
        Me.txtCategory.Location = New System.Drawing.Point(995, 110)
        Me.txtCategory.MendatroryField = False
        Me.txtCategory.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCategory.MyLinkLable1 = Me.lblRouteNo
        Me.txtCategory.MyLinkLable2 = Nothing
        Me.txtCategory.MyReadOnly = False
        Me.txtCategory.MyShowMasterFormButton = False
        Me.txtCategory.Name = "txtCategory"
        Me.txtCategory.ReferenceFieldDesc = Nothing
        Me.txtCategory.ReferenceFieldName = Nothing
        Me.txtCategory.ReferenceTableName = Nothing
        Me.txtCategory.Size = New System.Drawing.Size(113, 22)
        Me.txtCategory.TabIndex = 1540
        Me.txtCategory.Value = ""
        '
        'lblRouteNo
        '
        Me.lblRouteNo.FieldName = Nothing
        Me.lblRouteNo.Location = New System.Drawing.Point(-1, 25)
        Me.lblRouteNo.Name = "lblRouteNo"
        Me.lblRouteNo.Size = New System.Drawing.Size(54, 18)
        Me.lblRouteNo.TabIndex = 119
        Me.lblRouteNo.Text = "Route No"
        '
        'MyLabel21
        '
        Me.MyLabel21.FieldName = Nothing
        Me.MyLabel21.Location = New System.Drawing.Point(1035, 134)
        Me.MyLabel21.Name = "MyLabel21"
        Me.MyLabel21.Size = New System.Drawing.Size(62, 18)
        Me.MyLabel21.TabIndex = 1606
        Me.MyLabel21.Text = "Mobile No."
        '
        'txtAdharNo
        '
        Me.txtAdharNo.CalculationExpression = Nothing
        Me.txtAdharNo.FieldCode = Nothing
        Me.txtAdharNo.FieldDesc = Nothing
        Me.txtAdharNo.FieldMaxLength = 0
        Me.txtAdharNo.FieldName = Nothing
        Me.txtAdharNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAdharNo.isCalculatedField = False
        Me.txtAdharNo.IsSourceFromTable = False
        Me.txtAdharNo.IsSourceFromValueList = False
        Me.txtAdharNo.IsUnique = False
        Me.txtAdharNo.Location = New System.Drawing.Point(300, 180)
        Me.txtAdharNo.MaxLength = 200
        Me.txtAdharNo.MendatroryField = False
        Me.txtAdharNo.MyLinkLable1 = Me.MyLabel11
        Me.txtAdharNo.MyLinkLable2 = Nothing
        Me.txtAdharNo.Name = "txtAdharNo"
        Me.txtAdharNo.ReferenceFieldDesc = Nothing
        Me.txtAdharNo.ReferenceFieldName = Nothing
        Me.txtAdharNo.ReferenceTableName = Nothing
        Me.txtAdharNo.Size = New System.Drawing.Size(119, 18)
        Me.txtAdharNo.TabIndex = 1604
        '
        'lblReceiptAmt
        '
        Me.lblReceiptAmt.FieldName = Nothing
        Me.lblReceiptAmt.Location = New System.Drawing.Point(889, 89)
        Me.lblReceiptAmt.Name = "lblReceiptAmt"
        Me.lblReceiptAmt.Size = New System.Drawing.Size(67, 18)
        Me.lblReceiptAmt.TabIndex = 1602
        Me.lblReceiptAmt.Text = "Receipt Amt"
        '
        'lblReceiptAmtDesc
        '
        Me.lblReceiptAmtDesc.AutoSize = False
        Me.lblReceiptAmtDesc.BorderVisible = True
        Me.lblReceiptAmtDesc.FieldName = Nothing
        Me.lblReceiptAmtDesc.Location = New System.Drawing.Point(963, 89)
        Me.lblReceiptAmtDesc.Name = "lblReceiptAmtDesc"
        Me.lblReceiptAmtDesc.Size = New System.Drawing.Size(139, 19)
        Me.lblReceiptAmtDesc.TabIndex = 1603
        Me.lblReceiptAmtDesc.TextWrap = False
        '
        'lblReceipt
        '
        Me.lblReceipt.FieldName = Nothing
        Me.lblReceipt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReceipt.Location = New System.Drawing.Point(911, 66)
        Me.lblReceipt.Name = "lblReceipt"
        Me.lblReceipt.Size = New System.Drawing.Size(45, 16)
        Me.lblReceipt.TabIndex = 1601
        Me.lblReceipt.Text = "Receipt"
        '
        'txtReceipt
        '
        Me.txtReceipt.CalculationExpression = Nothing
        Me.txtReceipt.FieldCode = Nothing
        Me.txtReceipt.FieldDesc = Nothing
        Me.txtReceipt.FieldMaxLength = 0
        Me.txtReceipt.FieldName = Nothing
        Me.txtReceipt.isCalculatedField = False
        Me.txtReceipt.IsSourceFromTable = False
        Me.txtReceipt.IsSourceFromValueList = False
        Me.txtReceipt.IsUnique = False
        Me.txtReceipt.Location = New System.Drawing.Point(961, 62)
        Me.txtReceipt.MendatroryField = False
        Me.txtReceipt.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReceipt.MyLinkLable1 = Me.lblReceipt
        Me.txtReceipt.MyLinkLable2 = Nothing
        Me.txtReceipt.MyReadOnly = False
        Me.txtReceipt.MyShowMasterFormButton = False
        Me.txtReceipt.Name = "txtReceipt"
        Me.txtReceipt.ReferenceFieldDesc = Nothing
        Me.txtReceipt.ReferenceFieldName = Nothing
        Me.txtReceipt.ReferenceTableName = Nothing
        Me.txtReceipt.Size = New System.Drawing.Size(138, 22)
        Me.txtReceipt.TabIndex = 1600
        Me.txtReceipt.Value = ""
        '
        'chkNoTranspoter
        '
        Me.chkNoTranspoter.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNoTranspoter.Location = New System.Drawing.Point(1017, 46)
        Me.chkNoTranspoter.Name = "chkNoTranspoter"
        Me.chkNoTranspoter.Size = New System.Drawing.Size(99, 16)
        Me.chkNoTranspoter.TabIndex = 1599
        Me.chkNoTranspoter.Text = "NO Transporter"
        '
        'chkAPS
        '
        Me.chkAPS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAPS.Location = New System.Drawing.Point(377, 47)
        Me.chkAPS.Name = "chkAPS"
        Me.chkAPS.Size = New System.Drawing.Size(43, 16)
        Me.chkAPS.TabIndex = 1596
        Me.chkAPS.Text = "APS"
        '
        'txtIndentDate
        '
        Me.txtIndentDate.CalculationExpression = Nothing
        Me.txtIndentDate.CustomFormat = "dd/MM/yyyy"
        Me.txtIndentDate.FieldCode = Nothing
        Me.txtIndentDate.FieldDesc = Nothing
        Me.txtIndentDate.FieldMaxLength = 0
        Me.txtIndentDate.FieldName = Nothing
        Me.txtIndentDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIndentDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtIndentDate.isCalculatedField = False
        Me.txtIndentDate.IsSourceFromTable = False
        Me.txtIndentDate.IsSourceFromValueList = False
        Me.txtIndentDate.IsUnique = False
        Me.txtIndentDate.Location = New System.Drawing.Point(915, 193)
        Me.txtIndentDate.MendatroryField = False
        Me.txtIndentDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtIndentDate.MyLinkLable1 = Me.RadLabel4
        Me.txtIndentDate.MyLinkLable2 = Nothing
        Me.txtIndentDate.Name = "txtIndentDate"
        Me.txtIndentDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtIndentDate.ReferenceFieldDesc = Nothing
        Me.txtIndentDate.ReferenceFieldName = Nothing
        Me.txtIndentDate.ReferenceTableName = Nothing
        Me.txtIndentDate.Size = New System.Drawing.Size(106, 18)
        Me.txtIndentDate.TabIndex = 1594
        Me.txtIndentDate.TabStop = False
        Me.txtIndentDate.Text = "13/06/2011"
        Me.txtIndentDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'RadLabel4
        '
        Me.RadLabel4.FieldName = Nothing
        Me.RadLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel4.Location = New System.Drawing.Point(367, 3)
        Me.RadLabel4.Name = "RadLabel4"
        Me.RadLabel4.Size = New System.Drawing.Size(30, 16)
        Me.RadLabel4.TabIndex = 1
        Me.RadLabel4.Text = "Date"
        '
        'UsLock1
        '
        Me.UsLock1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.UsLock1.Location = New System.Drawing.Point(809, -2)
        Me.UsLock1.MyFont = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsLock1.Name = "UsLock1"
        Me.UsLock1.Size = New System.Drawing.Size(97, 16)
        Me.UsLock1.Status = common.ERPTransactionStatus.Pending
        Me.UsLock1.TabIndex = 1587
        '
        'lblIndentDate
        '
        Me.lblIndentDate.FieldName = Nothing
        Me.lblIndentDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIndentDate.Location = New System.Drawing.Point(830, 194)
        Me.lblIndentDate.Name = "lblIndentDate"
        Me.lblIndentDate.Size = New System.Drawing.Size(84, 16)
        Me.lblIndentDate.TabIndex = 1595
        Me.lblIndentDate.Text = "PO/Indent Date"
        '
        'chkExcludeKKFMND
        '
        Me.chkExcludeKKFMND.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkExcludeKKFMND.Location = New System.Drawing.Point(1017, 25)
        Me.chkExcludeKKFMND.Name = "chkExcludeKKFMND"
        Me.chkExcludeKKFMND.Size = New System.Drawing.Size(142, 16)
        Me.chkExcludeKKFMND.TabIndex = 1586
        Me.chkExcludeKKFMND.Text = "Exclude KKF Mandi Tax"
        '
        'txtIndentNo
        '
        Me.txtIndentNo.CalculationExpression = Nothing
        Me.txtIndentNo.FieldCode = Nothing
        Me.txtIndentNo.FieldDesc = Nothing
        Me.txtIndentNo.FieldMaxLength = 0
        Me.txtIndentNo.FieldName = Nothing
        Me.txtIndentNo.isCalculatedField = False
        Me.txtIndentNo.IsSourceFromTable = False
        Me.txtIndentNo.IsSourceFromValueList = False
        Me.txtIndentNo.IsUnique = False
        Me.txtIndentNo.Location = New System.Drawing.Point(710, 191)
        Me.txtIndentNo.MendatroryField = False
        Me.txtIndentNo.MyLinkLable1 = Nothing
        Me.txtIndentNo.MyLinkLable2 = Nothing
        Me.txtIndentNo.Name = "txtIndentNo"
        Me.txtIndentNo.ReferenceFieldDesc = Nothing
        Me.txtIndentNo.ReferenceFieldName = Nothing
        Me.txtIndentNo.ReferenceTableName = Nothing
        Me.txtIndentNo.Size = New System.Drawing.Size(111, 20)
        Me.txtIndentNo.TabIndex = 1593
        '
        'txtSupplyDate
        '
        Me.txtSupplyDate.CalculationExpression = Nothing
        Me.txtSupplyDate.CustomFormat = "dd/MM/yyyy hh:mm:tt"
        Me.txtSupplyDate.FieldCode = Nothing
        Me.txtSupplyDate.FieldDesc = Nothing
        Me.txtSupplyDate.FieldMaxLength = 0
        Me.txtSupplyDate.FieldName = Nothing
        Me.txtSupplyDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSupplyDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtSupplyDate.isCalculatedField = False
        Me.txtSupplyDate.IsSourceFromTable = False
        Me.txtSupplyDate.IsSourceFromValueList = False
        Me.txtSupplyDate.IsUnique = False
        Me.txtSupplyDate.Location = New System.Drawing.Point(603, 0)
        Me.txtSupplyDate.MendatroryField = False
        Me.txtSupplyDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtSupplyDate.MyLinkLable1 = Me.lblSupplyDate
        Me.txtSupplyDate.MyLinkLable2 = Nothing
        Me.txtSupplyDate.Name = "txtSupplyDate"
        Me.txtSupplyDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtSupplyDate.ReferenceFieldDesc = Nothing
        Me.txtSupplyDate.ReferenceFieldName = Nothing
        Me.txtSupplyDate.ReferenceTableName = Nothing
        Me.txtSupplyDate.Size = New System.Drawing.Size(125, 18)
        Me.txtSupplyDate.TabIndex = 1585
        Me.txtSupplyDate.TabStop = False
        Me.txtSupplyDate.Text = "13/06/2011 11:29:AM"
        Me.txtSupplyDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'lblSupplyDate
        '
        Me.lblSupplyDate.FieldName = Nothing
        Me.lblSupplyDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSupplyDate.Location = New System.Drawing.Point(533, 3)
        Me.lblSupplyDate.Name = "lblSupplyDate"
        Me.lblSupplyDate.Size = New System.Drawing.Size(68, 16)
        Me.lblSupplyDate.TabIndex = 1584
        Me.lblSupplyDate.Text = "Supply Date"
        '
        'lblIndentNo
        '
        Me.lblIndentNo.FieldName = Nothing
        Me.lblIndentNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIndentNo.Location = New System.Drawing.Point(633, 192)
        Me.lblIndentNo.Name = "lblIndentNo"
        Me.lblIndentNo.Size = New System.Drawing.Size(75, 16)
        Me.lblIndentNo.TabIndex = 1592
        Me.lblIndentNo.Text = "PO/Indent No"
        '
        'chkIsEwayBill
        '
        Me.chkIsEwayBill.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIsEwayBill.Location = New System.Drawing.Point(923, 134)
        Me.chkIsEwayBill.Name = "chkIsEwayBill"
        Me.chkIsEwayBill.Size = New System.Drawing.Size(110, 16)
        Me.chkIsEwayBill.TabIndex = 1583
        Me.chkIsEwayBill.Text = "Create E-Way Bill"
        '
        'chkNoCrateIssue
        '
        Me.chkNoCrateIssue.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNoCrateIssue.Location = New System.Drawing.Point(531, 180)
        Me.chkNoCrateIssue.Name = "chkNoCrateIssue"
        Me.chkNoCrateIssue.Size = New System.Drawing.Size(96, 16)
        Me.chkNoCrateIssue.TabIndex = 1582
        Me.chkNoCrateIssue.Text = "No Crate Issue"
        '
        'chkisTCS
        '
        Me.chkisTCS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkisTCS.Location = New System.Drawing.Point(425, 180)
        Me.chkisTCS.Name = "chkisTCS"
        Me.chkisTCS.Size = New System.Drawing.Size(43, 16)
        Me.chkisTCS.TabIndex = 1581
        Me.chkisTCS.Text = "TCS"
        '
        'chkTPT
        '
        Me.chkTPT.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTPT.Location = New System.Drawing.Point(481, 179)
        Me.chkTPT.Name = "chkTPT"
        Me.chkTPT.Size = New System.Drawing.Size(42, 16)
        Me.chkTPT.TabIndex = 1580
        Me.chkTPT.Text = "TPT"
        '
        'lblShipToLocation
        '
        Me.lblShipToLocation.AutoSize = False
        Me.lblShipToLocation.BorderVisible = True
        Me.lblShipToLocation.Enabled = False
        Me.lblShipToLocation.FieldName = Nothing
        Me.lblShipToLocation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblShipToLocation.Location = New System.Drawing.Point(199, 134)
        Me.lblShipToLocation.Name = "lblShipToLocation"
        Me.lblShipToLocation.Size = New System.Drawing.Size(135, 17)
        Me.lblShipToLocation.TabIndex = 1579
        Me.lblShipToLocation.TextWrap = False
        '
        'txtShipToLocation
        '
        Me.txtShipToLocation.CalculationExpression = Nothing
        Me.txtShipToLocation.FieldCode = Nothing
        Me.txtShipToLocation.FieldDesc = Nothing
        Me.txtShipToLocation.FieldMaxLength = 0
        Me.txtShipToLocation.FieldName = Nothing
        Me.txtShipToLocation.isCalculatedField = False
        Me.txtShipToLocation.IsSourceFromTable = False
        Me.txtShipToLocation.IsSourceFromValueList = False
        Me.txtShipToLocation.IsUnique = False
        Me.txtShipToLocation.Location = New System.Drawing.Point(75, 133)
        Me.txtShipToLocation.MendatroryField = False
        Me.txtShipToLocation.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtShipToLocation.MyLinkLable1 = Nothing
        Me.txtShipToLocation.MyLinkLable2 = Nothing
        Me.txtShipToLocation.MyReadOnly = False
        Me.txtShipToLocation.MyShowMasterFormButton = False
        Me.txtShipToLocation.Name = "txtShipToLocation"
        Me.txtShipToLocation.ReferenceFieldDesc = Nothing
        Me.txtShipToLocation.ReferenceFieldName = Nothing
        Me.txtShipToLocation.ReferenceTableName = Nothing
        Me.txtShipToLocation.Size = New System.Drawing.Size(115, 19)
        Me.txtShipToLocation.TabIndex = 1578
        Me.txtShipToLocation.Value = ""
        '
        'RadLabel18
        '
        Me.RadLabel18.FieldName = Nothing
        Me.RadLabel18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel18.Location = New System.Drawing.Point(1, 134)
        Me.RadLabel18.Name = "RadLabel18"
        Me.RadLabel18.Size = New System.Drawing.Size(67, 16)
        Me.RadLabel18.TabIndex = 1577
        Me.RadLabel18.Text = "Ship To Loc"
        '
        'chkManualVehicle
        '
        Me.chkManualVehicle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkManualVehicle.Location = New System.Drawing.Point(5, 159)
        Me.chkManualVehicle.Name = "chkManualVehicle"
        Me.chkManualVehicle.Size = New System.Drawing.Size(98, 16)
        Me.chkManualVehicle.TabIndex = 1576
        Me.chkManualVehicle.Text = "Manual Vehicle"
        '
        'MyLabel31
        '
        Me.MyLabel31.FieldName = Nothing
        Me.MyLabel31.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel31.Location = New System.Drawing.Point(-3, 89)
        Me.MyLabel31.Name = "MyLabel31"
        Me.MyLabel31.Size = New System.Drawing.Size(77, 16)
        Me.MyLabel31.TabIndex = 1575
        Me.MyLabel31.Text = "Transporter Id"
        '
        'lblTransporter
        '
        Me.lblTransporter.AutoSize = False
        Me.lblTransporter.BorderVisible = True
        Me.lblTransporter.FieldName = Nothing
        Me.lblTransporter.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTransporter.Location = New System.Drawing.Point(199, 91)
        Me.lblTransporter.Name = "lblTransporter"
        Me.lblTransporter.Size = New System.Drawing.Size(172, 17)
        Me.lblTransporter.TabIndex = 1574
        Me.lblTransporter.TextWrap = False
        '
        'lblVehicleNo
        '
        Me.lblVehicleNo.FieldName = Nothing
        Me.lblVehicleNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVehicleNo.Location = New System.Drawing.Point(124, 158)
        Me.lblVehicleNo.Name = "lblVehicleNo"
        Me.lblVehicleNo.Size = New System.Drawing.Size(61, 16)
        Me.lblVehicleNo.TabIndex = 1573
        Me.lblVehicleNo.Text = "Vehicle No"
        '
        'txtManualVehicle
        '
        Me.txtManualVehicle.CalculationExpression = Nothing
        Me.txtManualVehicle.FieldCode = Nothing
        Me.txtManualVehicle.FieldDesc = Nothing
        Me.txtManualVehicle.FieldMaxLength = 0
        Me.txtManualVehicle.FieldName = Nothing
        Me.txtManualVehicle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtManualVehicle.isCalculatedField = False
        Me.txtManualVehicle.IsSourceFromTable = False
        Me.txtManualVehicle.IsSourceFromValueList = False
        Me.txtManualVehicle.IsUnique = False
        Me.txtManualVehicle.Location = New System.Drawing.Point(191, 156)
        Me.txtManualVehicle.MaxLength = 200
        Me.txtManualVehicle.MendatroryField = False
        Me.txtManualVehicle.MyLinkLable1 = Me.lblVehicleNo
        Me.txtManualVehicle.MyLinkLable2 = Nothing
        Me.txtManualVehicle.Name = "txtManualVehicle"
        Me.txtManualVehicle.ReferenceFieldDesc = Nothing
        Me.txtManualVehicle.ReferenceFieldName = Nothing
        Me.txtManualVehicle.ReferenceTableName = Nothing
        Me.txtManualVehicle.Size = New System.Drawing.Size(136, 18)
        Me.txtManualVehicle.TabIndex = 1572
        '
        'fndTransporter
        '
        Me.fndTransporter.CalculationExpression = Nothing
        Me.fndTransporter.FieldCode = Nothing
        Me.fndTransporter.FieldDesc = Nothing
        Me.fndTransporter.FieldMaxLength = 0
        Me.fndTransporter.FieldName = Nothing
        Me.fndTransporter.isCalculatedField = False
        Me.fndTransporter.IsSourceFromTable = False
        Me.fndTransporter.IsSourceFromValueList = False
        Me.fndTransporter.IsUnique = False
        Me.fndTransporter.Location = New System.Drawing.Point(77, 89)
        Me.fndTransporter.MendatroryField = False
        Me.fndTransporter.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fndTransporter.MyLinkLable1 = Nothing
        Me.fndTransporter.MyLinkLable2 = Nothing
        Me.fndTransporter.MyReadOnly = False
        Me.fndTransporter.MyShowMasterFormButton = False
        Me.fndTransporter.Name = "fndTransporter"
        Me.fndTransporter.ReferenceFieldDesc = Nothing
        Me.fndTransporter.ReferenceFieldName = Nothing
        Me.fndTransporter.ReferenceTableName = Nothing
        Me.fndTransporter.Size = New System.Drawing.Size(115, 20)
        Me.fndTransporter.TabIndex = 1570
        Me.fndTransporter.Value = ""
        '
        'MyLabel20
        '
        Me.MyLabel20.FieldName = Nothing
        Me.MyLabel20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel20.Location = New System.Drawing.Point(626, 108)
        Me.MyLabel20.Name = "MyLabel20"
        Me.MyLabel20.Size = New System.Drawing.Size(63, 16)
        Me.MyLabel20.TabIndex = 1569
        Me.MyLabel20.Text = "Description"
        '
        'txtDescription
        '
        Me.txtDescription.CalculationExpression = Nothing
        Me.txtDescription.FieldCode = Nothing
        Me.txtDescription.FieldDesc = Nothing
        Me.txtDescription.FieldMaxLength = 0
        Me.txtDescription.FieldName = Nothing
        Me.txtDescription.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDescription.isCalculatedField = False
        Me.txtDescription.IsSourceFromTable = False
        Me.txtDescription.IsSourceFromValueList = False
        Me.txtDescription.IsUnique = False
        Me.txtDescription.Location = New System.Drawing.Point(700, 107)
        Me.txtDescription.MaxLength = 200
        Me.txtDescription.MendatroryField = False
        Me.txtDescription.MyLinkLable1 = Me.MyLabel20
        Me.txtDescription.MyLinkLable2 = Nothing
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.ReferenceFieldDesc = Nothing
        Me.txtDescription.ReferenceFieldName = Nothing
        Me.txtDescription.ReferenceTableName = Nothing
        Me.txtDescription.Size = New System.Drawing.Size(215, 18)
        Me.txtDescription.TabIndex = 1568
        '
        'lblReceiverName
        '
        Me.lblReceiverName.FieldName = Nothing
        Me.lblReceiverName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReceiverName.Location = New System.Drawing.Point(339, 154)
        Me.lblReceiverName.Name = "lblReceiverName"
        Me.lblReceiverName.Size = New System.Drawing.Size(84, 16)
        Me.lblReceiverName.TabIndex = 1567
        Me.lblReceiverName.Text = "Receiver Name"
        '
        'txtReceiverName
        '
        Me.txtReceiverName.CalculationExpression = Nothing
        Me.txtReceiverName.FieldCode = Nothing
        Me.txtReceiverName.FieldDesc = Nothing
        Me.txtReceiverName.FieldMaxLength = 0
        Me.txtReceiverName.FieldName = Nothing
        Me.txtReceiverName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReceiverName.isCalculatedField = False
        Me.txtReceiverName.IsSourceFromTable = False
        Me.txtReceiverName.IsSourceFromValueList = False
        Me.txtReceiverName.IsUnique = False
        Me.txtReceiverName.Location = New System.Drawing.Point(429, 153)
        Me.txtReceiverName.MaxLength = 200
        Me.txtReceiverName.MendatroryField = False
        Me.txtReceiverName.MyLinkLable1 = Me.lblReceiverName
        Me.txtReceiverName.MyLinkLable2 = Nothing
        Me.txtReceiverName.Name = "txtReceiverName"
        Me.txtReceiverName.ReferenceFieldDesc = Nothing
        Me.txtReceiverName.ReferenceFieldName = Nothing
        Me.txtReceiverName.ReferenceTableName = Nothing
        Me.txtReceiverName.Size = New System.Drawing.Size(188, 18)
        Me.txtReceiverName.TabIndex = 1566
        '
        'lblChequeNo
        '
        Me.lblChequeNo.FieldName = Nothing
        Me.lblChequeNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChequeNo.Location = New System.Drawing.Point(626, 134)
        Me.lblChequeNo.Name = "lblChequeNo"
        Me.lblChequeNo.Size = New System.Drawing.Size(128, 16)
        Me.lblChequeNo.TabIndex = 1563
        Me.lblChequeNo.Text = "Reference / Cheque No."
        '
        'txtChequeNo
        '
        Me.txtChequeNo.CalculationExpression = Nothing
        Me.txtChequeNo.FieldCode = Nothing
        Me.txtChequeNo.FieldDesc = Nothing
        Me.txtChequeNo.FieldMaxLength = 0
        Me.txtChequeNo.FieldName = Nothing
        Me.txtChequeNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeNo.isCalculatedField = False
        Me.txtChequeNo.IsSourceFromTable = False
        Me.txtChequeNo.IsSourceFromValueList = False
        Me.txtChequeNo.IsUnique = False
        Me.txtChequeNo.Location = New System.Drawing.Point(759, 133)
        Me.txtChequeNo.MaxLength = 200
        Me.txtChequeNo.MendatroryField = False
        Me.txtChequeNo.MyLinkLable1 = Me.lblChequeNo
        Me.txtChequeNo.MyLinkLable2 = Nothing
        Me.txtChequeNo.Name = "txtChequeNo"
        Me.txtChequeNo.ReferenceFieldDesc = Nothing
        Me.txtChequeNo.ReferenceFieldName = Nothing
        Me.txtChequeNo.ReferenceTableName = Nothing
        Me.txtChequeNo.Size = New System.Drawing.Size(154, 18)
        Me.txtChequeNo.TabIndex = 1562
        '
        'cmbPaymentType
        '
        Me.cmbPaymentType.AutoCompleteDisplayMember = Nothing
        Me.cmbPaymentType.AutoCompleteValueMember = Nothing
        Me.cmbPaymentType.CalculationExpression = Nothing
        Me.cmbPaymentType.DropDownAnimationEnabled = True
        Me.cmbPaymentType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbPaymentType.FieldCode = Nothing
        Me.cmbPaymentType.FieldDesc = Nothing
        Me.cmbPaymentType.FieldMaxLength = 0
        Me.cmbPaymentType.FieldName = Nothing
        Me.cmbPaymentType.isCalculatedField = False
        Me.cmbPaymentType.IsSourceFromTable = False
        Me.cmbPaymentType.IsSourceFromValueList = False
        Me.cmbPaymentType.IsUnique = False
        Me.cmbPaymentType.Location = New System.Drawing.Point(522, 132)
        Me.cmbPaymentType.MendatroryField = True
        Me.cmbPaymentType.MyLinkLable1 = Me.MyLabel7
        Me.cmbPaymentType.MyLinkLable2 = Nothing
        Me.cmbPaymentType.Name = "cmbPaymentType"
        Me.cmbPaymentType.ReferenceFieldDesc = Nothing
        Me.cmbPaymentType.ReferenceFieldName = Nothing
        Me.cmbPaymentType.ReferenceTableName = Nothing
        Me.cmbPaymentType.Size = New System.Drawing.Size(93, 20)
        Me.cmbPaymentType.TabIndex = 1561
        '
        'MyLabel7
        '
        Me.MyLabel7.FieldName = Nothing
        Me.MyLabel7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel7.Location = New System.Drawing.Point(147, 152)
        Me.MyLabel7.Name = "MyLabel7"
        Me.MyLabel7.Size = New System.Drawing.Size(58, 16)
        Me.MyLabel7.TabIndex = 1452
        Me.MyLabel7.Text = "BO Status"
        '
        'lblPaymentType
        '
        Me.lblPaymentType.FieldName = Nothing
        Me.lblPaymentType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPaymentType.Location = New System.Drawing.Point(439, 134)
        Me.lblPaymentType.Name = "lblPaymentType"
        Me.lblPaymentType.Size = New System.Drawing.Size(79, 16)
        Me.lblPaymentType.TabIndex = 1560
        Me.lblPaymentType.Text = "Payment Type"
        '
        'chkcashsale
        '
        Me.chkcashsale.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkcashsale.Location = New System.Drawing.Point(337, 134)
        Me.chkcashsale.Name = "chkcashsale"
        Me.chkcashsale.Size = New System.Drawing.Size(100, 16)
        Me.chkcashsale.TabIndex = 139
        Me.chkcashsale.Text = "Against Receipt"
        '
        'chkGhee
        '
        Me.chkGhee.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGhee.Location = New System.Drawing.Point(1017, 5)
        Me.chkGhee.Name = "chkGhee"
        Me.chkGhee.Size = New System.Drawing.Size(74, 16)
        Me.chkGhee.TabIndex = 1557
        Me.chkGhee.Text = "Only Ghee"
        '
        'rgbTaxNonTax
        '
        Me.rgbTaxNonTax.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.rgbTaxNonTax.Controls.Add(Me.lblFATPER)
        Me.rgbTaxNonTax.Controls.Add(Me.txtMBRTHours)
        Me.rgbTaxNonTax.Controls.Add(Me.txtFATPER)
        Me.rgbTaxNonTax.Controls.Add(Me.lblMBRT)
        Me.rgbTaxNonTax.Controls.Add(Me.lblSNFPER)
        Me.rgbTaxNonTax.Controls.Add(Me.txtTemp)
        Me.rgbTaxNonTax.Controls.Add(Me.txtSNFPER)
        Me.rgbTaxNonTax.Controls.Add(Me.lblTEMP)
        Me.rgbTaxNonTax.Controls.Add(Me.lblAcidity)
        Me.rgbTaxNonTax.Controls.Add(Me.txtAcidity)
        Me.rgbTaxNonTax.HeaderText = ""
        Me.rgbTaxNonTax.Location = New System.Drawing.Point(623, 153)
        Me.rgbTaxNonTax.Name = "rgbTaxNonTax"
        Me.rgbTaxNonTax.Size = New System.Drawing.Size(670, 33)
        Me.rgbTaxNonTax.TabIndex = 1556
        '
        'lblFATPER
        '
        Me.lblFATPER.FieldName = Nothing
        Me.lblFATPER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFATPER.Location = New System.Drawing.Point(20, 8)
        Me.lblFATPER.Name = "lblFATPER"
        Me.lblFATPER.Size = New System.Drawing.Size(41, 16)
        Me.lblFATPER.TabIndex = 1546
        Me.lblFATPER.Text = "FAT %"
        '
        'txtMBRTHours
        '
        Me.txtMBRTHours.CalculationExpression = Nothing
        Me.txtMBRTHours.FieldCode = Nothing
        Me.txtMBRTHours.FieldDesc = Nothing
        Me.txtMBRTHours.FieldMaxLength = 0
        Me.txtMBRTHours.FieldName = Nothing
        Me.txtMBRTHours.isCalculatedField = False
        Me.txtMBRTHours.IsSourceFromTable = False
        Me.txtMBRTHours.IsSourceFromValueList = False
        Me.txtMBRTHours.IsUnique = False
        Me.txtMBRTHours.Location = New System.Drawing.Point(585, 5)
        Me.txtMBRTHours.MendatroryField = False
        Me.txtMBRTHours.MyLinkLable1 = Nothing
        Me.txtMBRTHours.MyLinkLable2 = Nothing
        Me.txtMBRTHours.Name = "txtMBRTHours"
        Me.txtMBRTHours.ReferenceFieldDesc = Nothing
        Me.txtMBRTHours.ReferenceFieldName = Nothing
        Me.txtMBRTHours.ReferenceTableName = Nothing
        Me.txtMBRTHours.Size = New System.Drawing.Size(70, 20)
        Me.txtMBRTHours.TabIndex = 1555
        '
        'txtFATPER
        '
        Me.txtFATPER.CalculationExpression = Nothing
        Me.txtFATPER.FieldCode = Nothing
        Me.txtFATPER.FieldDesc = Nothing
        Me.txtFATPER.FieldMaxLength = 0
        Me.txtFATPER.FieldName = Nothing
        Me.txtFATPER.isCalculatedField = False
        Me.txtFATPER.IsSourceFromTable = False
        Me.txtFATPER.IsSourceFromValueList = False
        Me.txtFATPER.IsUnique = False
        Me.txtFATPER.Location = New System.Drawing.Point(67, 6)
        Me.txtFATPER.MendatroryField = False
        Me.txtFATPER.MyLinkLable1 = Nothing
        Me.txtFATPER.MyLinkLable2 = Nothing
        Me.txtFATPER.Name = "txtFATPER"
        Me.txtFATPER.ReferenceFieldDesc = Nothing
        Me.txtFATPER.ReferenceFieldName = Nothing
        Me.txtFATPER.ReferenceTableName = Nothing
        Me.txtFATPER.Size = New System.Drawing.Size(70, 20)
        Me.txtFATPER.TabIndex = 1547
        '
        'lblMBRT
        '
        Me.lblMBRT.FieldName = Nothing
        Me.lblMBRT.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMBRT.Location = New System.Drawing.Point(509, 8)
        Me.lblMBRT.Name = "lblMBRT"
        Me.lblMBRT.Size = New System.Drawing.Size(73, 16)
        Me.lblMBRT.TabIndex = 1554
        Me.lblMBRT.Text = "MBRT (HRS)"
        '
        'lblSNFPER
        '
        Me.lblSNFPER.FieldName = Nothing
        Me.lblSNFPER.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSNFPER.Location = New System.Drawing.Point(143, 8)
        Me.lblSNFPER.Name = "lblSNFPER"
        Me.lblSNFPER.Size = New System.Drawing.Size(42, 16)
        Me.lblSNFPER.TabIndex = 1548
        Me.lblSNFPER.Text = "SNF %"
        '
        'txtTemp
        '
        Me.txtTemp.CalculationExpression = Nothing
        Me.txtTemp.FieldCode = Nothing
        Me.txtTemp.FieldDesc = Nothing
        Me.txtTemp.FieldMaxLength = 0
        Me.txtTemp.FieldName = Nothing
        Me.txtTemp.isCalculatedField = False
        Me.txtTemp.IsSourceFromTable = False
        Me.txtTemp.IsSourceFromValueList = False
        Me.txtTemp.IsUnique = False
        Me.txtTemp.Location = New System.Drawing.Point(431, 5)
        Me.txtTemp.MendatroryField = False
        Me.txtTemp.MyLinkLable1 = Nothing
        Me.txtTemp.MyLinkLable2 = Nothing
        Me.txtTemp.Name = "txtTemp"
        Me.txtTemp.ReferenceFieldDesc = Nothing
        Me.txtTemp.ReferenceFieldName = Nothing
        Me.txtTemp.ReferenceTableName = Nothing
        Me.txtTemp.Size = New System.Drawing.Size(70, 20)
        Me.txtTemp.TabIndex = 1553
        '
        'txtSNFPER
        '
        Me.txtSNFPER.CalculationExpression = Nothing
        Me.txtSNFPER.FieldCode = Nothing
        Me.txtSNFPER.FieldDesc = Nothing
        Me.txtSNFPER.FieldMaxLength = 0
        Me.txtSNFPER.FieldName = Nothing
        Me.txtSNFPER.isCalculatedField = False
        Me.txtSNFPER.IsSourceFromTable = False
        Me.txtSNFPER.IsSourceFromValueList = False
        Me.txtSNFPER.IsUnique = False
        Me.txtSNFPER.Location = New System.Drawing.Point(187, 5)
        Me.txtSNFPER.MendatroryField = False
        Me.txtSNFPER.MyLinkLable1 = Nothing
        Me.txtSNFPER.MyLinkLable2 = Nothing
        Me.txtSNFPER.Name = "txtSNFPER"
        Me.txtSNFPER.ReferenceFieldDesc = Nothing
        Me.txtSNFPER.ReferenceFieldName = Nothing
        Me.txtSNFPER.ReferenceTableName = Nothing
        Me.txtSNFPER.Size = New System.Drawing.Size(65, 20)
        Me.txtSNFPER.TabIndex = 1549
        '
        'lblTEMP
        '
        Me.lblTEMP.FieldName = Nothing
        Me.lblTEMP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTEMP.Location = New System.Drawing.Point(392, 8)
        Me.lblTEMP.Name = "lblTEMP"
        Me.lblTEMP.Size = New System.Drawing.Size(41, 16)
        Me.lblTEMP.TabIndex = 1552
        Me.lblTEMP.Text = "Temp. "
        '
        'lblAcidity
        '
        Me.lblAcidity.FieldName = Nothing
        Me.lblAcidity.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAcidity.Location = New System.Drawing.Point(258, 8)
        Me.lblAcidity.Name = "lblAcidity"
        Me.lblAcidity.Size = New System.Drawing.Size(53, 16)
        Me.lblAcidity.TabIndex = 1550
        Me.lblAcidity.Text = "Acidity %"
        '
        'txtAcidity
        '
        Me.txtAcidity.CalculationExpression = Nothing
        Me.txtAcidity.FieldCode = Nothing
        Me.txtAcidity.FieldDesc = Nothing
        Me.txtAcidity.FieldMaxLength = 0
        Me.txtAcidity.FieldName = Nothing
        Me.txtAcidity.isCalculatedField = False
        Me.txtAcidity.IsSourceFromTable = False
        Me.txtAcidity.IsSourceFromValueList = False
        Me.txtAcidity.IsUnique = False
        Me.txtAcidity.Location = New System.Drawing.Point(317, 6)
        Me.txtAcidity.MendatroryField = False
        Me.txtAcidity.MyLinkLable1 = Nothing
        Me.txtAcidity.MyLinkLable2 = Nothing
        Me.txtAcidity.Name = "txtAcidity"
        Me.txtAcidity.ReferenceFieldDesc = Nothing
        Me.txtAcidity.ReferenceFieldName = Nothing
        Me.txtAcidity.ReferenceTableName = Nothing
        Me.txtAcidity.Size = New System.Drawing.Size(70, 20)
        Me.txtAcidity.TabIndex = 1551
        '
        'lblSubLocation
        '
        Me.lblSubLocation.AutoSize = False
        Me.lblSubLocation.BorderVisible = True
        Me.lblSubLocation.FieldName = Nothing
        Me.lblSubLocation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubLocation.Location = New System.Drawing.Point(199, 110)
        Me.lblSubLocation.Name = "lblSubLocation"
        Me.lblSubLocation.Size = New System.Drawing.Size(172, 17)
        Me.lblSubLocation.TabIndex = 1544
        Me.lblSubLocation.TextWrap = False
        '
        'MyLabel60
        '
        Me.MyLabel60.FieldName = Nothing
        Me.MyLabel60.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel60.Location = New System.Drawing.Point(-1, 113)
        Me.MyLabel60.Name = "MyLabel60"
        Me.MyLabel60.Size = New System.Drawing.Size(72, 16)
        Me.MyLabel60.TabIndex = 1545
        Me.MyLabel60.Text = "Sub Location"
        '
        'txtSubLocation
        '
        Me.txtSubLocation.CalculationExpression = Nothing
        Me.txtSubLocation.FieldCode = Nothing
        Me.txtSubLocation.FieldDesc = Nothing
        Me.txtSubLocation.FieldMaxLength = 0
        Me.txtSubLocation.FieldName = Nothing
        Me.txtSubLocation.isCalculatedField = False
        Me.txtSubLocation.IsSourceFromTable = False
        Me.txtSubLocation.IsSourceFromValueList = False
        Me.txtSubLocation.IsUnique = False
        Me.txtSubLocation.Location = New System.Drawing.Point(76, 111)
        Me.txtSubLocation.MendatroryField = False
        Me.txtSubLocation.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSubLocation.MyLinkLable1 = Me.MyLabel60
        Me.txtSubLocation.MyLinkLable2 = Me.lblSubLocation
        Me.txtSubLocation.MyReadOnly = False
        Me.txtSubLocation.MyShowMasterFormButton = False
        Me.txtSubLocation.Name = "txtSubLocation"
        Me.txtSubLocation.ReferenceFieldDesc = Nothing
        Me.txtSubLocation.ReferenceFieldName = Nothing
        Me.txtSubLocation.ReferenceTableName = Nothing
        Me.txtSubLocation.Size = New System.Drawing.Size(115, 19)
        Me.txtSubLocation.TabIndex = 1543
        Me.txtSubLocation.Value = ""
        '
        'chkDistributor
        '
        Me.chkDistributor.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDistributor.Location = New System.Drawing.Point(730, 1)
        Me.chkDistributor.Name = "chkDistributor"
        Me.chkDistributor.Size = New System.Drawing.Size(72, 16)
        Me.chkDistributor.TabIndex = 1485
        Me.chkDistributor.Text = "Distributor"
        '
        'cmbGatePassType
        '
        Me.cmbGatePassType.AutoCompleteDisplayMember = Nothing
        Me.cmbGatePassType.AutoCompleteValueMember = Nothing
        Me.cmbGatePassType.CalculationExpression = Nothing
        Me.cmbGatePassType.DropDownAnimationEnabled = True
        Me.cmbGatePassType.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbGatePassType.FieldCode = Nothing
        Me.cmbGatePassType.FieldDesc = Nothing
        Me.cmbGatePassType.FieldMaxLength = 0
        Me.cmbGatePassType.FieldName = Nothing
        Me.cmbGatePassType.isCalculatedField = False
        Me.cmbGatePassType.IsSourceFromTable = False
        Me.cmbGatePassType.IsSourceFromValueList = False
        Me.cmbGatePassType.IsUnique = False
        RadListDataItem1.Text = "Select"
        RadListDataItem2.Text = "AM"
        RadListDataItem3.Text = "PM"
        Me.cmbGatePassType.Items.Add(RadListDataItem1)
        Me.cmbGatePassType.Items.Add(RadListDataItem2)
        Me.cmbGatePassType.Items.Add(RadListDataItem3)
        Me.cmbGatePassType.Location = New System.Drawing.Point(171, 179)
        Me.cmbGatePassType.MendatroryField = False
        Me.cmbGatePassType.MyLinkLable1 = Nothing
        Me.cmbGatePassType.MyLinkLable2 = Nothing
        Me.cmbGatePassType.Name = "cmbGatePassType"
        Me.cmbGatePassType.ReferenceFieldDesc = Nothing
        Me.cmbGatePassType.ReferenceFieldName = Nothing
        Me.cmbGatePassType.ReferenceTableName = Nothing
        Me.cmbGatePassType.Size = New System.Drawing.Size(54, 20)
        Me.cmbGatePassType.TabIndex = 1538
        '
        'lblGatePassType
        '
        Me.lblGatePassType.FieldName = Nothing
        Me.lblGatePassType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGatePassType.Location = New System.Drawing.Point(78, 181)
        Me.lblGatePassType.Name = "lblGatePassType"
        Me.lblGatePassType.Size = New System.Drawing.Size(87, 16)
        Me.lblGatePassType.TabIndex = 1537
        Me.lblGatePassType.Text = "Gate Pass Type"
        '
        'lblShiftType
        '
        Me.lblShiftType.Location = New System.Drawing.Point(686, -3)
        Me.lblShiftType.Name = "lblShiftType"
        Me.lblShiftType.Size = New System.Drawing.Size(2, 2)
        Me.lblShiftType.TabIndex = 1536
        '
        'txtCouponDate
        '
        Me.txtCouponDate.CalculationExpression = Nothing
        Me.txtCouponDate.CustomFormat = "dd/MM/yyyy"
        Me.txtCouponDate.FieldCode = Nothing
        Me.txtCouponDate.FieldDesc = Nothing
        Me.txtCouponDate.FieldMaxLength = 0
        Me.txtCouponDate.FieldName = Nothing
        Me.txtCouponDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCouponDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtCouponDate.isCalculatedField = False
        Me.txtCouponDate.IsSourceFromTable = False
        Me.txtCouponDate.IsSourceFromValueList = False
        Me.txtCouponDate.IsUnique = False
        Me.txtCouponDate.Location = New System.Drawing.Point(509, 69)
        Me.txtCouponDate.MendatroryField = False
        Me.txtCouponDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtCouponDate.MyLinkLable1 = Me.RadLabel4
        Me.txtCouponDate.MyLinkLable2 = Nothing
        Me.txtCouponDate.Name = "txtCouponDate"
        Me.txtCouponDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtCouponDate.ReferenceFieldDesc = Nothing
        Me.txtCouponDate.ReferenceFieldName = Nothing
        Me.txtCouponDate.ReferenceTableName = Nothing
        Me.txtCouponDate.Size = New System.Drawing.Size(106, 18)
        Me.txtCouponDate.TabIndex = 1525
        Me.txtCouponDate.TabStop = False
        Me.txtCouponDate.Text = "13/06/2011"
        Me.txtCouponDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'lblCouponDate
        '
        Me.lblCouponDate.FieldName = Nothing
        Me.lblCouponDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCouponDate.Location = New System.Drawing.Point(429, 70)
        Me.lblCouponDate.Name = "lblCouponDate"
        Me.lblCouponDate.Size = New System.Drawing.Size(73, 16)
        Me.lblCouponDate.TabIndex = 1535
        Me.lblCouponDate.Text = "Coupon Date"
        '
        'txtBPLRemark
        '
        Me.txtBPLRemark.CalculationExpression = Nothing
        Me.txtBPLRemark.FieldCode = Nothing
        Me.txtBPLRemark.FieldDesc = Nothing
        Me.txtBPLRemark.FieldMaxLength = 0
        Me.txtBPLRemark.FieldName = Nothing
        Me.txtBPLRemark.isCalculatedField = False
        Me.txtBPLRemark.IsSourceFromTable = False
        Me.txtBPLRemark.IsSourceFromValueList = False
        Me.txtBPLRemark.IsUnique = False
        Me.txtBPLRemark.Location = New System.Drawing.Point(506, 113)
        Me.txtBPLRemark.MendatroryField = False
        Me.txtBPLRemark.MyLinkLable1 = Nothing
        Me.txtBPLRemark.MyLinkLable2 = Nothing
        Me.txtBPLRemark.Name = "txtBPLRemark"
        Me.txtBPLRemark.ReferenceFieldDesc = Nothing
        Me.txtBPLRemark.ReferenceFieldName = Nothing
        Me.txtBPLRemark.ReferenceTableName = Nothing
        Me.txtBPLRemark.Size = New System.Drawing.Size(111, 20)
        Me.txtBPLRemark.TabIndex = 1531
        '
        'txtBPLName
        '
        Me.txtBPLName.CalculationExpression = Nothing
        Me.txtBPLName.FieldCode = Nothing
        Me.txtBPLName.FieldDesc = Nothing
        Me.txtBPLName.FieldMaxLength = 0
        Me.txtBPLName.FieldName = Nothing
        Me.txtBPLName.isCalculatedField = False
        Me.txtBPLName.IsSourceFromTable = False
        Me.txtBPLName.IsSourceFromValueList = False
        Me.txtBPLName.IsUnique = False
        Me.txtBPLName.Location = New System.Drawing.Point(506, 91)
        Me.txtBPLName.MendatroryField = False
        Me.txtBPLName.MyLinkLable1 = Nothing
        Me.txtBPLName.MyLinkLable2 = Nothing
        Me.txtBPLName.Name = "txtBPLName"
        Me.txtBPLName.ReferenceFieldDesc = Nothing
        Me.txtBPLName.ReferenceFieldName = Nothing
        Me.txtBPLName.ReferenceTableName = Nothing
        Me.txtBPLName.Size = New System.Drawing.Size(111, 20)
        Me.txtBPLName.TabIndex = 1531
        '
        'txtCouponCode
        '
        Me.txtCouponCode.CalculationExpression = Nothing
        Me.txtCouponCode.FieldCode = Nothing
        Me.txtCouponCode.FieldDesc = Nothing
        Me.txtCouponCode.FieldMaxLength = 0
        Me.txtCouponCode.FieldName = Nothing
        Me.txtCouponCode.isCalculatedField = False
        Me.txtCouponCode.IsSourceFromTable = False
        Me.txtCouponCode.IsSourceFromValueList = False
        Me.txtCouponCode.IsUnique = False
        Me.txtCouponCode.Location = New System.Drawing.Point(506, 46)
        Me.txtCouponCode.MendatroryField = False
        Me.txtCouponCode.MyLinkLable1 = Nothing
        Me.txtCouponCode.MyLinkLable2 = Nothing
        Me.txtCouponCode.Name = "txtCouponCode"
        Me.txtCouponCode.ReferenceFieldDesc = Nothing
        Me.txtCouponCode.ReferenceFieldName = Nothing
        Me.txtCouponCode.ReferenceTableName = Nothing
        Me.txtCouponCode.Size = New System.Drawing.Size(111, 20)
        Me.txtCouponCode.TabIndex = 1530
        '
        'lblBPLRemark
        '
        Me.lblBPLRemark.FieldName = Nothing
        Me.lblBPLRemark.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBPLRemark.Location = New System.Drawing.Point(432, 113)
        Me.lblBPLRemark.Name = "lblBPLRemark"
        Me.lblBPLRemark.Size = New System.Drawing.Size(46, 16)
        Me.lblBPLRemark.TabIndex = 1517
        Me.lblBPLRemark.Text = "Remark"
        '
        'lblBPLName
        '
        Me.lblBPLName.FieldName = Nothing
        Me.lblBPLName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBPLName.Location = New System.Drawing.Point(432, 93)
        Me.lblBPLName.Name = "lblBPLName"
        Me.lblBPLName.Size = New System.Drawing.Size(36, 16)
        Me.lblBPLName.TabIndex = 1517
        Me.lblBPLName.Text = "Name"
        '
        'lblCouponCode
        '
        Me.lblCouponCode.FieldName = Nothing
        Me.lblCouponCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCouponCode.Location = New System.Drawing.Point(429, 47)
        Me.lblCouponCode.Name = "lblCouponCode"
        Me.lblCouponCode.Size = New System.Drawing.Size(76, 16)
        Me.lblCouponCode.TabIndex = 1517
        Me.lblCouponCode.Text = "Coupon Code"
        '
        'chkBPL
        '
        Me.chkBPL.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBPL.Location = New System.Drawing.Point(377, 27)
        Me.chkBPL.Name = "chkBPL"
        Me.chkBPL.Size = New System.Drawing.Size(42, 16)
        Me.chkBPL.TabIndex = 1485
        Me.chkBPL.Text = "BPL"
        '
        'rgbItemType
        '
        Me.rgbItemType.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.rgbItemType.Controls.Add(Me.rbtnNonTax)
        Me.rgbItemType.Controls.Add(Me.rbtnTaxable)
        Me.rgbItemType.HeaderText = "Item Type"
        Me.rgbItemType.Location = New System.Drawing.Point(911, -4)
        Me.rgbItemType.Name = "rgbItemType"
        Me.rgbItemType.Size = New System.Drawing.Size(100, 64)
        Me.rgbItemType.TabIndex = 1534
        Me.rgbItemType.Text = "Item Type"
        '
        'rbtnNonTax
        '
        Me.rbtnNonTax.Location = New System.Drawing.Point(5, 44)
        Me.rbtnNonTax.Name = "rbtnNonTax"
        Me.rbtnNonTax.Size = New System.Drawing.Size(84, 18)
        Me.rbtnNonTax.TabIndex = 1
        Me.rbtnNonTax.Text = "Non-Taxable"
        '
        'rbtnTaxable
        '
        Me.rbtnTaxable.Location = New System.Drawing.Point(6, 22)
        Me.rbtnTaxable.Name = "rbtnTaxable"
        Me.rbtnTaxable.Size = New System.Drawing.Size(58, 18)
        Me.rbtnTaxable.TabIndex = 0
        Me.rbtnTaxable.Text = "Taxable"
        '
        'txtVehicleCode
        '
        Me.txtVehicleCode.CalculationExpression = Nothing
        Me.txtVehicleCode.FieldCode = Nothing
        Me.txtVehicleCode.FieldDesc = Nothing
        Me.txtVehicleCode.FieldMaxLength = 0
        Me.txtVehicleCode.FieldName = Nothing
        Me.txtVehicleCode.isCalculatedField = False
        Me.txtVehicleCode.IsSourceFromTable = False
        Me.txtVehicleCode.IsSourceFromValueList = False
        Me.txtVehicleCode.IsUnique = False
        Me.txtVehicleCode.Location = New System.Drawing.Point(682, 62)
        Me.txtVehicleCode.MendatroryField = False
        Me.txtVehicleCode.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVehicleCode.MyLinkLable1 = Me.lblRouteNo
        Me.txtVehicleCode.MyLinkLable2 = Nothing
        Me.txtVehicleCode.MyReadOnly = False
        Me.txtVehicleCode.MyShowMasterFormButton = False
        Me.txtVehicleCode.Name = "txtVehicleCode"
        Me.txtVehicleCode.ReferenceFieldDesc = Nothing
        Me.txtVehicleCode.ReferenceFieldName = Nothing
        Me.txtVehicleCode.ReferenceTableName = Nothing
        Me.txtVehicleCode.Size = New System.Drawing.Size(90, 20)
        Me.txtVehicleCode.TabIndex = 1533
        Me.txtVehicleCode.Value = ""
        '
        'txtVehicleName
        '
        Me.txtVehicleName.CalculationExpression = Nothing
        Me.txtVehicleName.FieldCode = Nothing
        Me.txtVehicleName.FieldDesc = Nothing
        Me.txtVehicleName.FieldMaxLength = 0
        Me.txtVehicleName.FieldName = Nothing
        Me.txtVehicleName.isCalculatedField = False
        Me.txtVehicleName.IsSourceFromTable = False
        Me.txtVehicleName.IsSourceFromValueList = False
        Me.txtVehicleName.IsUnique = False
        Me.txtVehicleName.Location = New System.Drawing.Point(775, 62)
        Me.txtVehicleName.MendatroryField = False
        Me.txtVehicleName.MyLinkLable1 = Nothing
        Me.txtVehicleName.MyLinkLable2 = Nothing
        Me.txtVehicleName.Name = "txtVehicleName"
        Me.txtVehicleName.ReferenceFieldDesc = Nothing
        Me.txtVehicleName.ReferenceFieldName = Nothing
        Me.txtVehicleName.ReferenceTableName = Nothing
        Me.txtVehicleName.Size = New System.Drawing.Size(131, 20)
        Me.txtVehicleName.TabIndex = 1532
        '
        'txtRouteName1
        '
        Me.txtRouteName1.CalculationExpression = Nothing
        Me.txtRouteName1.FieldCode = Nothing
        Me.txtRouteName1.FieldDesc = Nothing
        Me.txtRouteName1.FieldMaxLength = 0
        Me.txtRouteName1.FieldName = Nothing
        Me.txtRouteName1.isCalculatedField = False
        Me.txtRouteName1.IsSourceFromTable = False
        Me.txtRouteName1.IsSourceFromValueList = False
        Me.txtRouteName1.IsUnique = False
        Me.txtRouteName1.Location = New System.Drawing.Point(749, 19)
        Me.txtRouteName1.MendatroryField = False
        Me.txtRouteName1.MyLinkLable1 = Nothing
        Me.txtRouteName1.MyLinkLable2 = Nothing
        Me.txtRouteName1.Name = "txtRouteName1"
        Me.txtRouteName1.ReferenceFieldDesc = Nothing
        Me.txtRouteName1.ReferenceFieldName = Nothing
        Me.txtRouteName1.ReferenceTableName = Nothing
        Me.txtRouteName1.Size = New System.Drawing.Size(154, 20)
        Me.txtRouteName1.TabIndex = 1530
        '
        'txtRouteCode1
        '
        Me.txtRouteCode1.CalculationExpression = Nothing
        Me.txtRouteCode1.FieldCode = Nothing
        Me.txtRouteCode1.FieldDesc = Nothing
        Me.txtRouteCode1.FieldMaxLength = 0
        Me.txtRouteCode1.FieldName = Nothing
        Me.txtRouteCode1.isCalculatedField = False
        Me.txtRouteCode1.IsSourceFromTable = False
        Me.txtRouteCode1.IsSourceFromValueList = False
        Me.txtRouteCode1.IsUnique = False
        Me.txtRouteCode1.Location = New System.Drawing.Point(682, 19)
        Me.txtRouteCode1.MendatroryField = False
        Me.txtRouteCode1.MyLinkLable1 = Nothing
        Me.txtRouteCode1.MyLinkLable2 = Nothing
        Me.txtRouteCode1.Name = "txtRouteCode1"
        Me.txtRouteCode1.ReferenceFieldDesc = Nothing
        Me.txtRouteCode1.ReferenceFieldName = Nothing
        Me.txtRouteCode1.ReferenceTableName = Nothing
        Me.txtRouteCode1.Size = New System.Drawing.Size(65, 20)
        Me.txtRouteCode1.TabIndex = 1529
        '
        'btnCC
        '
        Me.btnCC.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnCC.Location = New System.Drawing.Point(336, 1)
        Me.btnCC.Name = "btnCC"
        Me.btnCC.Size = New System.Drawing.Size(20, 21)
        Me.btnCC.TabIndex = 6
        Me.btnCC.Text = "CC"
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
        Me.txtDate.Location = New System.Drawing.Point(400, 2)
        Me.txtDate.MendatroryField = False
        Me.txtDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.MyLinkLable1 = Me.RadLabel4
        Me.txtDate.MyLinkLable2 = Nothing
        Me.txtDate.Name = "txtDate"
        Me.txtDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDate.ReferenceFieldDesc = Nothing
        Me.txtDate.ReferenceFieldName = Nothing
        Me.txtDate.ReferenceTableName = Nothing
        Me.txtDate.Size = New System.Drawing.Size(127, 18)
        Me.txtDate.TabIndex = 1524
        Me.txtDate.TabStop = False
        Me.txtDate.Text = "13/06/2011 11:29 AM"
        Me.txtDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'cmbcashcredit
        '
        Me.cmbcashcredit.AutoCompleteDisplayMember = Nothing
        Me.cmbcashcredit.AutoCompleteValueMember = Nothing
        Me.cmbcashcredit.CalculationExpression = Nothing
        Me.cmbcashcredit.DropDownAnimationEnabled = True
        Me.cmbcashcredit.DropDownStyle = Telerik.WinControls.RadDropDownStyle.DropDownList
        Me.cmbcashcredit.FieldCode = Nothing
        Me.cmbcashcredit.FieldDesc = Nothing
        Me.cmbcashcredit.FieldMaxLength = 0
        Me.cmbcashcredit.FieldName = Nothing
        Me.cmbcashcredit.isCalculatedField = False
        Me.cmbcashcredit.IsSourceFromTable = False
        Me.cmbcashcredit.IsSourceFromValueList = False
        Me.cmbcashcredit.IsUnique = False
        RadListDataItem4.Text = "Select"
        RadListDataItem5.Text = "CASH"
        RadListDataItem6.Text = "CREDIT"
        Me.cmbcashcredit.Items.Add(RadListDataItem4)
        Me.cmbcashcredit.Items.Add(RadListDataItem5)
        Me.cmbcashcredit.Items.Add(RadListDataItem6)
        Me.cmbcashcredit.Location = New System.Drawing.Point(506, 23)
        Me.cmbcashcredit.MendatroryField = False
        Me.cmbcashcredit.MyLinkLable1 = Nothing
        Me.cmbcashcredit.MyLinkLable2 = Nothing
        Me.cmbcashcredit.Name = "cmbcashcredit"
        Me.cmbcashcredit.ReferenceFieldDesc = Nothing
        Me.cmbcashcredit.ReferenceFieldName = Nothing
        Me.cmbcashcredit.ReferenceTableName = Nothing
        Me.cmbcashcredit.Size = New System.Drawing.Size(99, 20)
        Me.cmbcashcredit.TabIndex = 1515
        '
        'lblCredit
        '
        Me.lblCredit.FieldName = Nothing
        Me.lblCredit.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCredit.Location = New System.Drawing.Point(434, 25)
        Me.lblCredit.Name = "lblCredit"
        Me.lblCredit.Size = New System.Drawing.Size(66, 16)
        Me.lblCredit.TabIndex = 1516
        Me.lblCredit.Text = "Cash/Credit"
        '
        'lblOutstandingDesc
        '
        Me.lblOutstandingDesc.AutoSize = False
        Me.lblOutstandingDesc.BorderVisible = True
        Me.lblOutstandingDesc.FieldName = Nothing
        Me.lblOutstandingDesc.Location = New System.Drawing.Point(683, 84)
        Me.lblOutstandingDesc.Name = "lblOutstandingDesc"
        Me.lblOutstandingDesc.Size = New System.Drawing.Size(133, 19)
        Me.lblOutstandingDesc.TabIndex = 1494
        Me.lblOutstandingDesc.TextWrap = False
        '
        'lblOutStanding
        '
        Me.lblOutStanding.FieldName = Nothing
        Me.lblOutStanding.Location = New System.Drawing.Point(626, 84)
        Me.lblOutStanding.Name = "lblOutStanding"
        Me.lblOutStanding.Size = New System.Drawing.Size(53, 18)
        Me.lblOutStanding.TabIndex = 1493
        Me.lblOutStanding.Text = "Total O/S"
        '
        'lblRoute1
        '
        Me.lblRoute1.FieldName = Nothing
        Me.lblRoute1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRoute1.Location = New System.Drawing.Point(623, 23)
        Me.lblRoute1.Name = "lblRoute1"
        Me.lblRoute1.Size = New System.Drawing.Size(36, 16)
        Me.lblRoute1.TabIndex = 1489
        Me.lblRoute1.Text = "Route"
        '
        'lblPriceCode1
        '
        Me.lblPriceCode1.FieldName = Nothing
        Me.lblPriceCode1.Location = New System.Drawing.Point(623, 41)
        Me.lblPriceCode1.Name = "lblPriceCode1"
        Me.lblPriceCode1.Size = New System.Drawing.Size(60, 18)
        Me.lblPriceCode1.TabIndex = 1491
        Me.lblPriceCode1.Text = "Price Code"
        '
        'lblPriceCodeDesc
        '
        Me.lblPriceCodeDesc.AutoSize = False
        Me.lblPriceCodeDesc.BorderVisible = True
        Me.lblPriceCodeDesc.FieldName = Nothing
        Me.lblPriceCodeDesc.Location = New System.Drawing.Point(683, 41)
        Me.lblPriceCodeDesc.Name = "lblPriceCodeDesc"
        Me.lblPriceCodeDesc.Size = New System.Drawing.Size(221, 19)
        Me.lblPriceCodeDesc.TabIndex = 1492
        Me.lblPriceCodeDesc.TextWrap = False
        '
        'lblVehicle1
        '
        Me.lblVehicle1.FieldName = Nothing
        Me.lblVehicle1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVehicle1.Location = New System.Drawing.Point(623, 62)
        Me.lblVehicle1.Name = "lblVehicle1"
        Me.lblVehicle1.Size = New System.Drawing.Size(43, 16)
        Me.lblVehicle1.TabIndex = 1490
        Me.lblVehicle1.Text = "Vehicle"
        '
        'chkSampling
        '
        Me.chkSampling.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSampling.Location = New System.Drawing.Point(5, 181)
        Me.chkSampling.Name = "chkSampling"
        Me.chkSampling.Size = New System.Drawing.Size(67, 16)
        Me.chkSampling.TabIndex = 1455
        Me.chkSampling.Text = "Sampling"
        '
        'lblLocation
        '
        Me.lblLocation.AutoSize = False
        Me.lblLocation.BorderVisible = True
        Me.lblLocation.FieldName = Nothing
        Me.lblLocation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLocation.Location = New System.Drawing.Point(199, 71)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(172, 17)
        Me.lblLocation.TabIndex = 154
        Me.lblLocation.TextWrap = False
        '
        'txtLocation
        '
        Me.txtLocation.CalculationExpression = Nothing
        Me.txtLocation.FieldCode = Nothing
        Me.txtLocation.FieldDesc = Nothing
        Me.txtLocation.FieldMaxLength = 0
        Me.txtLocation.FieldName = Nothing
        Me.txtLocation.isCalculatedField = False
        Me.txtLocation.IsSourceFromTable = False
        Me.txtLocation.IsSourceFromValueList = False
        Me.txtLocation.IsUnique = False
        Me.txtLocation.Location = New System.Drawing.Point(78, 71)
        Me.txtLocation.MendatroryField = True
        Me.txtLocation.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocation.MyLinkLable1 = Me.RadLabel15
        Me.txtLocation.MyLinkLable2 = Nothing
        Me.txtLocation.MyReadOnly = False
        Me.txtLocation.MyShowMasterFormButton = False
        Me.txtLocation.Name = "txtLocation"
        Me.txtLocation.ReferenceFieldDesc = Nothing
        Me.txtLocation.ReferenceFieldName = Nothing
        Me.txtLocation.ReferenceTableName = Nothing
        Me.txtLocation.Size = New System.Drawing.Size(115, 17)
        Me.txtLocation.TabIndex = 153
        Me.txtLocation.Value = ""
        '
        'RadLabel15
        '
        Me.RadLabel15.FieldName = Nothing
        Me.RadLabel15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel15.Location = New System.Drawing.Point(-1, 67)
        Me.RadLabel15.Name = "RadLabel15"
        Me.RadLabel15.Size = New System.Drawing.Size(49, 16)
        Me.RadLabel15.TabIndex = 152
        Me.RadLabel15.Text = "Location"
        '
        'lblRouteDesc
        '
        Me.lblRouteDesc.AutoSize = False
        Me.lblRouteDesc.BorderVisible = True
        Me.lblRouteDesc.FieldName = Nothing
        Me.lblRouteDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRouteDesc.Location = New System.Drawing.Point(199, 27)
        Me.lblRouteDesc.Name = "lblRouteDesc"
        Me.lblRouteDesc.Size = New System.Drawing.Size(172, 19)
        Me.lblRouteDesc.TabIndex = 144
        Me.lblRouteDesc.TextWrap = False
        '
        'lblAbandonmentNo
        '
        Me.lblAbandonmentNo.FieldName = Nothing
        Me.lblAbandonmentNo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAbandonmentNo.Location = New System.Drawing.Point(764, 11)
        Me.lblAbandonmentNo.Name = "lblAbandonmentNo"
        Me.lblAbandonmentNo.Size = New System.Drawing.Size(2, 2)
        Me.lblAbandonmentNo.TabIndex = 27
        '
        'RadGroupBox2
        '
        Me.RadGroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadGroupBox2.Controls.Add(Me.gv1)
        Me.RadGroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadGroupBox2.HeaderText = "Item Details"
        Me.RadGroupBox2.Location = New System.Drawing.Point(2, 217)
        Me.RadGroupBox2.Name = "RadGroupBox2"
        Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox2.Size = New System.Drawing.Size(1255, 184)
        Me.RadGroupBox2.TabIndex = 28
        Me.RadGroupBox2.Text = "Item Details"
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
        Me.gv1.Size = New System.Drawing.Size(1235, 154)
        Me.gv1.TabIndex = 17
        Me.gv1.TabStop = False
        Me.gv1.VarID = ""
        '
        'lblVendorName
        '
        Me.lblVendorName.AutoSize = False
        Me.lblVendorName.BorderVisible = True
        Me.lblVendorName.FieldName = Nothing
        Me.lblVendorName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVendorName.Location = New System.Drawing.Point(199, 50)
        Me.lblVendorName.Name = "lblVendorName"
        Me.lblVendorName.Size = New System.Drawing.Size(172, 17)
        Me.lblVendorName.TabIndex = 10
        Me.lblVendorName.TextWrap = False
        '
        'txtRouteNo
        '
        Me.txtRouteNo.CalculationExpression = Nothing
        Me.txtRouteNo.FieldCode = Nothing
        Me.txtRouteNo.FieldDesc = Nothing
        Me.txtRouteNo.FieldMaxLength = 0
        Me.txtRouteNo.FieldName = Nothing
        Me.txtRouteNo.isCalculatedField = False
        Me.txtRouteNo.IsSourceFromTable = False
        Me.txtRouteNo.IsSourceFromValueList = False
        Me.txtRouteNo.IsUnique = False
        Me.txtRouteNo.Location = New System.Drawing.Point(78, 27)
        Me.txtRouteNo.MendatroryField = False
        Me.txtRouteNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRouteNo.MyLinkLable1 = Me.lblRouteNo
        Me.txtRouteNo.MyLinkLable2 = Nothing
        Me.txtRouteNo.MyReadOnly = False
        Me.txtRouteNo.MyShowMasterFormButton = False
        Me.txtRouteNo.Name = "txtRouteNo"
        Me.txtRouteNo.ReferenceFieldDesc = Nothing
        Me.txtRouteNo.ReferenceFieldName = Nothing
        Me.txtRouteNo.ReferenceTableName = Nothing
        Me.txtRouteNo.Size = New System.Drawing.Size(115, 19)
        Me.txtRouteNo.TabIndex = 22
        Me.txtRouteNo.Value = ""
        '
        'RadLabel2
        '
        Me.RadLabel2.FieldName = Nothing
        Me.RadLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel2.Location = New System.Drawing.Point(-1, 47)
        Me.RadLabel2.Name = "RadLabel2"
        Me.RadLabel2.Size = New System.Drawing.Size(55, 16)
        Me.RadLabel2.TabIndex = 21
        Me.RadLabel2.Text = "Customer"
        '
        'RadLabel1
        '
        Me.RadLabel1.FieldName = Nothing
        Me.RadLabel1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel1.Location = New System.Drawing.Point(-1, 3)
        Me.RadLabel1.Name = "RadLabel1"
        Me.RadLabel1.Size = New System.Drawing.Size(33, 16)
        Me.RadLabel1.TabIndex = 22
        Me.RadLabel1.Text = "Code"
        '
        'txtVendorNo
        '
        Me.txtVendorNo.CalculationExpression = Nothing
        Me.txtVendorNo.FieldCode = Nothing
        Me.txtVendorNo.FieldDesc = Nothing
        Me.txtVendorNo.FieldMaxLength = 0
        Me.txtVendorNo.FieldName = Nothing
        Me.txtVendorNo.isCalculatedField = False
        Me.txtVendorNo.IsSourceFromTable = False
        Me.txtVendorNo.IsSourceFromValueList = False
        Me.txtVendorNo.IsUnique = False
        Me.txtVendorNo.Location = New System.Drawing.Point(78, 50)
        Me.txtVendorNo.MendatroryField = True
        Me.txtVendorNo.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVendorNo.MyLinkLable1 = Me.RadLabel2
        Me.txtVendorNo.MyLinkLable2 = Me.lblVendorName
        Me.txtVendorNo.MyReadOnly = False
        Me.txtVendorNo.MyShowMasterFormButton = True
        Me.txtVendorNo.Name = "txtVendorNo"
        Me.txtVendorNo.ReferenceFieldDesc = Nothing
        Me.txtVendorNo.ReferenceFieldName = Nothing
        Me.txtVendorNo.ReferenceTableName = Nothing
        Me.txtVendorNo.Size = New System.Drawing.Size(116, 18)
        Me.txtVendorNo.TabIndex = 9
        Me.txtVendorNo.Value = ""
        '
        'txtDocNo
        '
        Me.txtDocNo.FieldName = Nothing
        Me.txtDocNo.Location = New System.Drawing.Point(46, 1)
        Me.txtDocNo.MendatroryField = False
        Me.txtDocNo.MyCharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.txtDocNo.MyFont = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.txtDocNo.MyLinkLable1 = Me.RadLabel1
        Me.txtDocNo.MyLinkLable2 = Nothing
        Me.txtDocNo.MyMaxLength = 32767
        Me.txtDocNo.MyReadOnly = False
        Me.txtDocNo.Name = "txtDocNo"
        Me.txtDocNo.Size = New System.Drawing.Size(269, 22)
        Me.txtDocNo.TabIndex = 0
        Me.txtDocNo.TabStop = False
        Me.txtDocNo.Value = ""
        '
        'btnAddNew
        '
        Me.btnAddNew.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddNew.Image = Global.XpertERPDairySale.My.Resources.Resources._new
        Me.btnAddNew.ImageAlignment = System.Drawing.ContentAlignment.MiddleCenter
        Me.btnAddNew.Location = New System.Drawing.Point(314, 1)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(20, 21)
        Me.btnAddNew.TabIndex = 1
        '
        'RadPageViewPage6
        '
        Me.RadPageViewPage6.Controls.Add(Me.MyLabel19)
        Me.RadPageViewPage6.Controls.Add(Me.RadGroupBox1)
        Me.RadPageViewPage6.Controls.Add(Me.gv2)
        Me.RadPageViewPage6.Controls.Add(Me.txtSchemeTaxGroup)
        Me.RadPageViewPage6.Controls.Add(Me.MyLabel44)
        Me.RadPageViewPage6.Controls.Add(Me.lblTaxGroupScheme)
        Me.RadPageViewPage6.Controls.Add(Me.GroupBox1)
        Me.RadPageViewPage6.Controls.Add(Me.txtTaxGroup)
        Me.RadPageViewPage6.Controls.Add(Me.lblTaxGrpName)
        Me.RadPageViewPage6.Controls.Add(Me.RadLabel11)
        Me.RadPageViewPage6.ItemSize = New System.Drawing.SizeF(47.0!, 26.0!)
        Me.RadPageViewPage6.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage6.Name = "RadPageViewPage6"
        Me.RadPageViewPage6.Size = New System.Drawing.Size(1260, 409)
        Me.RadPageViewPage6.Text = "Taxes"
        '
        'MyLabel19
        '
        Me.MyLabel19.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MyLabel19.FieldName = Nothing
        Me.MyLabel19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel19.ForeColor = System.Drawing.SystemColors.Highlight
        Me.MyLabel19.Location = New System.Drawing.Point(844, 321)
        Me.MyLabel19.Name = "MyLabel19"
        Me.MyLabel19.Size = New System.Drawing.Size(155, 16)
        Me.MyLabel19.TabIndex = 18
        Me.MyLabel19.Text = "Double click To Chage Rate"
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
        Me.RadGroupBox1.Location = New System.Drawing.Point(12, 335)
        Me.RadGroupBox1.Name = "RadGroupBox1"
        Me.RadGroupBox1.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox1.Size = New System.Drawing.Size(992, 73)
        Me.RadGroupBox1.TabIndex = 17
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
        Me.txtTermCode.Location = New System.Drawing.Point(68, 25)
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
        Me.txtTermCode.Size = New System.Drawing.Size(143, 20)
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
        Me.RadLabel16.TabIndex = 2
        Me.RadLabel16.Text = "Term Code"
        '
        'lblTermName
        '
        Me.lblTermName.AutoSize = False
        Me.lblTermName.BorderVisible = True
        Me.lblTermName.FieldName = Nothing
        Me.lblTermName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTermName.Location = New System.Drawing.Point(220, 23)
        Me.lblTermName.Name = "lblTermName"
        Me.lblTermName.Size = New System.Drawing.Size(321, 20)
        Me.lblTermName.TabIndex = 1
        Me.lblTermName.TextWrap = False
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
        Me.txtDueDate.Location = New System.Drawing.Point(70, 50)
        Me.txtDueDate.MendatroryField = False
        Me.txtDueDate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDueDate.MyLinkLable1 = Me.RadLabel17
        Me.txtDueDate.MyLinkLable2 = Nothing
        Me.txtDueDate.Name = "txtDueDate"
        Me.txtDueDate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtDueDate.ReferenceFieldDesc = Nothing
        Me.txtDueDate.ReferenceFieldName = Nothing
        Me.txtDueDate.ReferenceTableName = Nothing
        Me.txtDueDate.Size = New System.Drawing.Size(81, 18)
        Me.txtDueDate.TabIndex = 2
        Me.txtDueDate.TabStop = False
        Me.txtDueDate.Text = "13-06-2011"
        Me.txtDueDate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'RadLabel17
        '
        Me.RadLabel17.FieldName = Nothing
        Me.RadLabel17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel17.Location = New System.Drawing.Point(6, 51)
        Me.RadLabel17.Name = "RadLabel17"
        Me.RadLabel17.Size = New System.Drawing.Size(54, 16)
        Me.RadLabel17.TabIndex = 4
        Me.RadLabel17.Text = "Due Date"
        '
        'gv2
        '
        Me.gv2.BackColor = System.Drawing.Color.FromArgb(CType(CType(233, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.gv2.BeginEditMode = Telerik.WinControls.RadGridViewBeginEditMode.BeginEditOnEnter
        Me.gv2.Cursor = System.Windows.Forms.Cursors.Default
        Me.gv2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gv2.ForeColor = System.Drawing.Color.Black
        Me.gv2.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.gv2.Location = New System.Drawing.Point(12, 41)
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
        Me.gv2.Size = New System.Drawing.Size(1084, 274)
        Me.gv2.TabIndex = 16
        Me.gv2.TabStop = False
        Me.gv2.VarID = ""
        '
        'txtSchemeTaxGroup
        '
        Me.txtSchemeTaxGroup.CalculationExpression = Nothing
        Me.txtSchemeTaxGroup.FieldCode = Nothing
        Me.txtSchemeTaxGroup.FieldDesc = Nothing
        Me.txtSchemeTaxGroup.FieldMaxLength = 0
        Me.txtSchemeTaxGroup.FieldName = Nothing
        Me.txtSchemeTaxGroup.isCalculatedField = False
        Me.txtSchemeTaxGroup.IsSourceFromTable = False
        Me.txtSchemeTaxGroup.IsSourceFromValueList = False
        Me.txtSchemeTaxGroup.IsUnique = False
        Me.txtSchemeTaxGroup.Location = New System.Drawing.Point(843, 10)
        Me.txtSchemeTaxGroup.MendatroryField = True
        Me.txtSchemeTaxGroup.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSchemeTaxGroup.MyLinkLable1 = Me.MyLabel44
        Me.txtSchemeTaxGroup.MyLinkLable2 = Me.lblTaxGroupScheme
        Me.txtSchemeTaxGroup.MyReadOnly = False
        Me.txtSchemeTaxGroup.MyShowMasterFormButton = False
        Me.txtSchemeTaxGroup.Name = "txtSchemeTaxGroup"
        Me.txtSchemeTaxGroup.ReferenceFieldDesc = Nothing
        Me.txtSchemeTaxGroup.ReferenceFieldName = Nothing
        Me.txtSchemeTaxGroup.ReferenceTableName = Nothing
        Me.txtSchemeTaxGroup.Size = New System.Drawing.Size(132, 20)
        Me.txtSchemeTaxGroup.TabIndex = 13
        Me.txtSchemeTaxGroup.Value = ""
        '
        'MyLabel44
        '
        Me.MyLabel44.FieldName = Nothing
        Me.MyLabel44.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel44.Location = New System.Drawing.Point(725, 13)
        Me.MyLabel44.Name = "MyLabel44"
        Me.MyLabel44.Size = New System.Drawing.Size(104, 16)
        Me.MyLabel44.TabIndex = 15
        Me.MyLabel44.Text = "Scheme Tax Group"
        '
        'lblTaxGroupScheme
        '
        Me.lblTaxGroupScheme.AutoSize = False
        Me.lblTaxGroupScheme.BorderVisible = True
        Me.lblTaxGroupScheme.FieldName = Nothing
        Me.lblTaxGroupScheme.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxGroupScheme.Location = New System.Drawing.Point(983, 10)
        Me.lblTaxGroupScheme.Name = "lblTaxGroupScheme"
        Me.lblTaxGroupScheme.Size = New System.Drawing.Size(259, 20)
        Me.lblTaxGroupScheme.TabIndex = 14
        Me.lblTaxGroupScheme.TextWrap = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbtnTaxCalManual)
        Me.GroupBox1.Controls.Add(Me.rbtnTaxCalAutomatic)
        Me.GroupBox1.Location = New System.Drawing.Point(556, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(163, 36)
        Me.GroupBox1.TabIndex = 11
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
        Me.rbtnTaxCalManual.TabStop = False
        Me.rbtnTaxCalManual.Text = "Manual"
        '
        'rbtnTaxCalAutomatic
        '
        Me.rbtnTaxCalAutomatic.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rbtnTaxCalAutomatic.Location = New System.Drawing.Point(7, 13)
        Me.rbtnTaxCalAutomatic.MyLinkLable1 = Nothing
        Me.rbtnTaxCalAutomatic.MyLinkLable2 = Nothing
        Me.rbtnTaxCalAutomatic.Name = "rbtnTaxCalAutomatic"
        Me.rbtnTaxCalAutomatic.Size = New System.Drawing.Size(72, 18)
        Me.rbtnTaxCalAutomatic.TabIndex = 0
        Me.rbtnTaxCalAutomatic.Text = "Automatic"
        Me.rbtnTaxCalAutomatic.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
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
        Me.txtTaxGroup.Location = New System.Drawing.Point(78, 9)
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
        Me.txtTaxGroup.TabIndex = 9
        Me.txtTaxGroup.Value = ""
        '
        'RadLabel11
        '
        Me.RadLabel11.FieldName = Nothing
        Me.RadLabel11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel11.Location = New System.Drawing.Point(12, 12)
        Me.RadLabel11.Name = "RadLabel11"
        Me.RadLabel11.Size = New System.Drawing.Size(60, 16)
        Me.RadLabel11.TabIndex = 12
        Me.RadLabel11.Text = "Tax Group"
        '
        'lblTaxGrpName
        '
        Me.lblTaxGrpName.AutoSize = False
        Me.lblTaxGrpName.BorderVisible = True
        Me.lblTaxGrpName.FieldName = Nothing
        Me.lblTaxGrpName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaxGrpName.Location = New System.Drawing.Point(229, 9)
        Me.lblTaxGrpName.Name = "lblTaxGrpName"
        Me.lblTaxGrpName.Size = New System.Drawing.Size(321, 20)
        Me.lblTaxGrpName.TabIndex = 10
        Me.lblTaxGrpName.TextWrap = False
        '
        'RadPageViewPage2
        '
        Me.RadPageViewPage2.ItemSize = New System.Drawing.SizeF(76.0!, 26.0!)
        Me.RadPageViewPage2.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage2.Name = "RadPageViewPage2"
        Me.RadPageViewPage2.Size = New System.Drawing.Size(1260, 409)
        Me.RadPageViewPage2.Text = "Item Details"
        '
        'RadPageViewPage4
        '
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel71)
        Me.RadPageViewPage4.Controls.Add(Me.lblGrossAmount)
        Me.RadPageViewPage4.Controls.Add(Me.lblTPTComAmt)
        Me.RadPageViewPage4.Controls.Add(Me.txtTPTComAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblRoundOff)
        Me.RadPageViewPage4.Controls.Add(Me.btnUpdateVehicle)
        Me.RadPageViewPage4.Controls.Add(Me.GroupBox2)
        Me.RadPageViewPage4.Controls.Add(Me.txtTCAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblTCAmt)
        Me.RadPageViewPage4.Controls.Add(Me.txtSecurity)
        Me.RadPageViewPage4.Controls.Add(Me.lblSecuirty)
        Me.RadPageViewPage4.Controls.Add(Me.txtDCAmt)
        Me.RadPageViewPage4.Controls.Add(Me.lblDCAmt)
        Me.RadPageViewPage4.Controls.Add(Me.TxtRoundoff)
        Me.RadPageViewPage4.Controls.Add(Me.txtTCSTaxRate)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel57)
        Me.RadPageViewPage4.Controls.Add(Me.lblActualTCSTaxBaseAmt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel58)
        Me.RadPageViewPage4.Controls.Add(Me.txttcstaxbaseamount)
        Me.RadPageViewPage4.Controls.Add(Me.lblInvoiceDiscAmt)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel6)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel3)
        Me.RadPageViewPage4.Controls.Add(Me.RadGroupBox3)
        Me.RadPageViewPage4.Controls.Add(Me.txtDiscAmt)
        Me.RadPageViewPage4.Controls.Add(Me.txtDiscPer)
        Me.RadPageViewPage4.Controls.Add(Me.MyLabel1)
        Me.RadPageViewPage4.Controls.Add(Me.pnlCurrConv)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel32)
        Me.RadPageViewPage4.Controls.Add(Me.lblAddCharges1)
        Me.RadPageViewPage4.Controls.Add(Me.txtComment)
        Me.RadPageViewPage4.Controls.Add(Me.RadLabel14)
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
        Me.RadPageViewPage4.Size = New System.Drawing.Size(1260, 409)
        Me.RadPageViewPage4.Text = "Total"
        '
        'MyLabel71
        '
        Me.MyLabel71.FieldName = Nothing
        Me.MyLabel71.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel71.Location = New System.Drawing.Point(122, 329)
        Me.MyLabel71.Name = "MyLabel71"
        Me.MyLabel71.Size = New System.Drawing.Size(79, 16)
        Me.MyLabel71.TabIndex = 1580
        Me.MyLabel71.Text = "Gross Amount"
        '
        'lblGrossAmount
        '
        Me.lblGrossAmount.AutoSize = False
        Me.lblGrossAmount.BorderVisible = True
        Me.lblGrossAmount.FieldName = Nothing
        Me.lblGrossAmount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGrossAmount.Location = New System.Drawing.Point(224, 328)
        Me.lblGrossAmount.Name = "lblGrossAmount"
        Me.lblGrossAmount.Size = New System.Drawing.Size(110, 18)
        Me.lblGrossAmount.TabIndex = 1579
        Me.lblGrossAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTPTComAmt
        '
        Me.lblTPTComAmt.FieldName = Nothing
        Me.lblTPTComAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTPTComAmt.Location = New System.Drawing.Point(94, 307)
        Me.lblTPTComAmt.Name = "lblTPTComAmt"
        Me.lblTPTComAmt.Size = New System.Drawing.Size(123, 16)
        Me.lblTPTComAmt.TabIndex = 1578
        Me.lblTPTComAmt.Text = "- TPT Commission Amt"
        '
        'txtTPTComAmt
        '
        Me.txtTPTComAmt.AutoSize = False
        Me.txtTPTComAmt.BorderVisible = True
        Me.txtTPTComAmt.FieldName = Nothing
        Me.txtTPTComAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTPTComAmt.Location = New System.Drawing.Point(224, 306)
        Me.txtTPTComAmt.Name = "txtTPTComAmt"
        Me.txtTPTComAmt.Size = New System.Drawing.Size(110, 18)
        Me.txtTPTComAmt.TabIndex = 1577
        Me.txtTPTComAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRoundOff
        '
        Me.lblRoundOff.FieldName = Nothing
        Me.lblRoundOff.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRoundOff.Location = New System.Drawing.Point(429, 309)
        Me.lblRoundOff.Name = "lblRoundOff"
        Me.lblRoundOff.Size = New System.Drawing.Size(81, 16)
        Me.lblRoundOff.TabIndex = 1568
        Me.lblRoundOff.Text = "Round Off Amt"
        '
        'btnUpdateVehicle
        '
        Me.btnUpdateVehicle.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnUpdateVehicle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpdateVehicle.Location = New System.Drawing.Point(841, 340)
        Me.btnUpdateVehicle.Name = "btnUpdateVehicle"
        Me.btnUpdateVehicle.Size = New System.Drawing.Size(148, 22)
        Me.btnUpdateVehicle.TabIndex = 1567
        Me.btnUpdateVehicle.Text = "Update Details For EWB"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btn_QtyReset)
        Me.GroupBox2.Controls.Add(Me.btnHistory)
        Me.GroupBox2.Controls.Add(Me.BtnRecieptEntry)
        Me.GroupBox2.Location = New System.Drawing.Point(995, 173)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(262, 176)
        Me.GroupBox2.TabIndex = 1488
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "GroupBox2"
        '
        'btn_QtyReset
        '
        Me.btn_QtyReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_QtyReset.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_QtyReset.Location = New System.Drawing.Point(28, 24)
        Me.btn_QtyReset.Name = "btn_QtyReset"
        Me.btn_QtyReset.Size = New System.Drawing.Size(57, 22)
        Me.btn_QtyReset.TabIndex = 1487
        Me.btn_QtyReset.Text = "Qty Reset"
        Me.btn_QtyReset.Visible = False
        '
        'btnHistory
        '
        Me.btnHistory.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnHistory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHistory.Location = New System.Drawing.Point(28, 54)
        Me.btnHistory.Name = "btnHistory"
        Me.btnHistory.Size = New System.Drawing.Size(57, 22)
        Me.btnHistory.TabIndex = 1529
        Me.btnHistory.Text = "History"
        '
        'BtnRecieptEntry
        '
        Me.BtnRecieptEntry.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnRecieptEntry.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnRecieptEntry.Location = New System.Drawing.Point(28, 81)
        Me.BtnRecieptEntry.Name = "BtnRecieptEntry"
        Me.BtnRecieptEntry.Size = New System.Drawing.Size(76, 22)
        Me.BtnRecieptEntry.TabIndex = 1503
        Me.BtnRecieptEntry.Text = "Reciept Entry"
        '
        'txtTCAmt
        '
        Me.txtTCAmt.AutoSize = False
        Me.txtTCAmt.BorderVisible = True
        Me.txtTCAmt.FieldName = Nothing
        Me.txtTCAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTCAmt.Location = New System.Drawing.Point(584, 283)
        Me.txtTCAmt.Name = "txtTCAmt"
        Me.txtTCAmt.Size = New System.Drawing.Size(82, 18)
        Me.txtTCAmt.TabIndex = 1411
        Me.txtTCAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTCAmt
        '
        Me.lblTCAmt.FieldName = Nothing
        Me.lblTCAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTCAmt.Location = New System.Drawing.Point(427, 284)
        Me.lblTCAmt.Name = "lblTCAmt"
        Me.lblTCAmt.Size = New System.Drawing.Size(153, 16)
        Me.lblTCAmt.TabIndex = 1412
        Me.lblTCAmt.Text = "Transporter Commission Amt"
        '
        'txtSecurity
        '
        Me.txtSecurity.AutoSize = False
        Me.txtSecurity.BorderVisible = True
        Me.txtSecurity.FieldName = Nothing
        Me.txtSecurity.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSecurity.Location = New System.Drawing.Point(584, 261)
        Me.txtSecurity.Name = "txtSecurity"
        Me.txtSecurity.Size = New System.Drawing.Size(82, 18)
        Me.txtSecurity.TabIndex = 1409
        Me.txtSecurity.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSecuirty
        '
        Me.lblSecuirty.FieldName = Nothing
        Me.lblSecuirty.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSecuirty.Location = New System.Drawing.Point(435, 262)
        Me.lblSecuirty.Name = "lblSecuirty"
        Me.lblSecuirty.Size = New System.Drawing.Size(118, 16)
        Me.lblSecuirty.TabIndex = 1410
        Me.lblSecuirty.Text = "Total Secuirty Amount"
        '
        'txtDCAmt
        '
        Me.txtDCAmt.AutoSize = False
        Me.txtDCAmt.BorderVisible = True
        Me.txtDCAmt.FieldName = Nothing
        Me.txtDCAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDCAmt.Location = New System.Drawing.Point(584, 237)
        Me.txtDCAmt.Name = "txtDCAmt"
        Me.txtDCAmt.Size = New System.Drawing.Size(82, 18)
        Me.txtDCAmt.TabIndex = 1407
        Me.txtDCAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDCAmt
        '
        Me.lblDCAmt.FieldName = Nothing
        Me.lblDCAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDCAmt.Location = New System.Drawing.Point(434, 238)
        Me.lblDCAmt.Name = "lblDCAmt"
        Me.lblDCAmt.Size = New System.Drawing.Size(144, 16)
        Me.lblDCAmt.TabIndex = 1408
        Me.lblDCAmt.Text = "Distributor Commssion Amt"
        '
        'TxtRoundoff
        '
        Me.TxtRoundoff.AutoSize = False
        Me.TxtRoundoff.BorderVisible = True
        Me.TxtRoundoff.FieldName = Nothing
        Me.TxtRoundoff.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtRoundoff.Location = New System.Drawing.Point(584, 307)
        Me.TxtRoundoff.Name = "TxtRoundoff"
        Me.TxtRoundoff.Size = New System.Drawing.Size(82, 18)
        Me.TxtRoundoff.TabIndex = 1406
        Me.TxtRoundoff.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtTCSTaxRate
        '
        Me.txtTCSTaxRate.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.txtTCSTaxRate.CalculationExpression = Nothing
        Me.txtTCSTaxRate.DecimalPlaces = 0
        Me.txtTCSTaxRate.FieldCode = Nothing
        Me.txtTCSTaxRate.FieldDesc = Nothing
        Me.txtTCSTaxRate.FieldMaxLength = 0
        Me.txtTCSTaxRate.FieldName = Nothing
        Me.txtTCSTaxRate.isCalculatedField = False
        Me.txtTCSTaxRate.IsSourceFromTable = False
        Me.txtTCSTaxRate.IsSourceFromValueList = False
        Me.txtTCSTaxRate.IsUnique = False
        Me.txtTCSTaxRate.Location = New System.Drawing.Point(225, 393)
        Me.txtTCSTaxRate.MendatroryField = False
        Me.txtTCSTaxRate.MyLinkLable1 = Nothing
        Me.txtTCSTaxRate.MyLinkLable2 = Nothing
        Me.txtTCSTaxRate.Name = "txtTCSTaxRate"
        Me.txtTCSTaxRate.ReadOnly = True
        Me.txtTCSTaxRate.ReferenceFieldDesc = Nothing
        Me.txtTCSTaxRate.ReferenceFieldName = Nothing
        Me.txtTCSTaxRate.ReferenceTableName = Nothing
        Me.txtTCSTaxRate.Size = New System.Drawing.Size(115, 20)
        Me.txtTCSTaxRate.TabIndex = 1405
        Me.txtTCSTaxRate.Text = "0"
        Me.txtTCSTaxRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtTCSTaxRate.Value = New Decimal(New Integer() {0, 0, 0, 0})
        Me.txtTCSTaxRate.Visible = False
        '
        'MyLabel57
        '
        Me.MyLabel57.FieldName = Nothing
        Me.MyLabel57.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel57.Location = New System.Drawing.Point(82, 349)
        Me.MyLabel57.Name = "MyLabel57"
        Me.MyLabel57.Size = New System.Drawing.Size(138, 16)
        Me.MyLabel57.TabIndex = 1404
        Me.MyLabel57.Text = "Actual TCS Tax Base Amt"
        '
        'lblActualTCSTaxBaseAmt
        '
        Me.lblActualTCSTaxBaseAmt.AutoSize = False
        Me.lblActualTCSTaxBaseAmt.BorderVisible = True
        Me.lblActualTCSTaxBaseAmt.FieldName = Nothing
        Me.lblActualTCSTaxBaseAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActualTCSTaxBaseAmt.Location = New System.Drawing.Point(224, 348)
        Me.lblActualTCSTaxBaseAmt.Name = "lblActualTCSTaxBaseAmt"
        Me.lblActualTCSTaxBaseAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblActualTCSTaxBaseAmt.TabIndex = 1403
        Me.lblActualTCSTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel58
        '
        Me.MyLabel58.FieldName = Nothing
        Me.MyLabel58.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel58.Location = New System.Drawing.Point(94, 371)
        Me.MyLabel58.Name = "MyLabel58"
        Me.MyLabel58.Size = New System.Drawing.Size(122, 16)
        Me.MyLabel58.TabIndex = 1402
        Me.MyLabel58.Text = "TCS Tax Base Amount"
        '
        'txttcstaxbaseamount
        '
        Me.txttcstaxbaseamount.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.txttcstaxbaseamount.CalculationExpression = Nothing
        Me.txttcstaxbaseamount.DecimalPlaces = 0
        Me.txttcstaxbaseamount.FieldCode = Nothing
        Me.txttcstaxbaseamount.FieldDesc = Nothing
        Me.txttcstaxbaseamount.FieldMaxLength = 0
        Me.txttcstaxbaseamount.FieldName = Nothing
        Me.txttcstaxbaseamount.isCalculatedField = False
        Me.txttcstaxbaseamount.IsSourceFromTable = False
        Me.txttcstaxbaseamount.IsSourceFromValueList = False
        Me.txttcstaxbaseamount.IsUnique = False
        Me.txttcstaxbaseamount.Location = New System.Drawing.Point(225, 368)
        Me.txttcstaxbaseamount.MendatroryField = False
        Me.txttcstaxbaseamount.MyLinkLable1 = Nothing
        Me.txttcstaxbaseamount.MyLinkLable2 = Nothing
        Me.txttcstaxbaseamount.Name = "txttcstaxbaseamount"
        Me.txttcstaxbaseamount.ReadOnly = True
        Me.txttcstaxbaseamount.ReferenceFieldDesc = Nothing
        Me.txttcstaxbaseamount.ReferenceFieldName = Nothing
        Me.txttcstaxbaseamount.ReferenceTableName = Nothing
        Me.txttcstaxbaseamount.Size = New System.Drawing.Size(110, 20)
        Me.txttcstaxbaseamount.TabIndex = 1401
        Me.txttcstaxbaseamount.Text = "0"
        Me.txttcstaxbaseamount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txttcstaxbaseamount.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'lblInvoiceDiscAmt
        '
        Me.lblInvoiceDiscAmt.AutoSize = False
        Me.lblInvoiceDiscAmt.BorderVisible = True
        Me.lblInvoiceDiscAmt.FieldName = Nothing
        Me.lblInvoiceDiscAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceDiscAmt.Location = New System.Drawing.Point(223, 147)
        Me.lblInvoiceDiscAmt.Name = "lblInvoiceDiscAmt"
        Me.lblInvoiceDiscAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblInvoiceDiscAmt.TabIndex = 7
        Me.lblInvoiceDiscAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel6
        '
        Me.MyLabel6.FieldName = Nothing
        Me.MyLabel6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel6.Location = New System.Drawing.Point(78, 147)
        Me.MyLabel6.Name = "MyLabel6"
        Me.MyLabel6.Size = New System.Drawing.Size(138, 16)
        Me.MyLabel6.TabIndex = 150
        Me.MyLabel6.Text = "- Invoice Discount Amount"
        '
        'MyLabel3
        '
        Me.MyLabel3.FieldName = Nothing
        Me.MyLabel3.Location = New System.Drawing.Point(147, 99)
        Me.MyLabel3.Name = "MyLabel3"
        Me.MyLabel3.Size = New System.Drawing.Size(68, 18)
        Me.MyLabel3.TabIndex = 3
        Me.MyLabel3.Text = "Discount On"
        '
        'RadGroupBox3
        '
        Me.RadGroupBox3.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        Me.RadGroupBox3.Controls.Add(Me.chkDiscountOnAmt)
        Me.RadGroupBox3.Controls.Add(Me.chkDiscountOnRate)
        Me.RadGroupBox3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadGroupBox3.HeaderText = ""
        Me.RadGroupBox3.Location = New System.Drawing.Point(222, 99)
        Me.RadGroupBox3.Name = "RadGroupBox3"
        Me.RadGroupBox3.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        Me.RadGroupBox3.Size = New System.Drawing.Size(141, 21)
        Me.RadGroupBox3.TabIndex = 4
        '
        'chkDiscountOnAmt
        '
        Me.chkDiscountOnAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDiscountOnAmt.Location = New System.Drawing.Point(66, 2)
        Me.chkDiscountOnAmt.Name = "chkDiscountOnAmt"
        Me.chkDiscountOnAmt.Size = New System.Drawing.Size(59, 16)
        Me.chkDiscountOnAmt.TabIndex = 1
        Me.chkDiscountOnAmt.Text = "Amount"
        '
        'chkDiscountOnRate
        '
        Me.chkDiscountOnRate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDiscountOnRate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDiscountOnRate.Location = New System.Drawing.Point(15, 2)
        Me.chkDiscountOnRate.Name = "chkDiscountOnRate"
        Me.chkDiscountOnRate.Size = New System.Drawing.Size(44, 16)
        Me.chkDiscountOnRate.TabIndex = 0
        Me.chkDiscountOnRate.Text = "Rate"
        Me.chkDiscountOnRate.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'txtDiscAmt
        '
        Me.txtDiscAmt.BackColor = System.Drawing.Color.White
        Me.txtDiscAmt.CalculationExpression = Nothing
        Me.txtDiscAmt.DecimalPlaces = 5
        Me.txtDiscAmt.FieldCode = Nothing
        Me.txtDiscAmt.FieldDesc = Nothing
        Me.txtDiscAmt.FieldMaxLength = 0
        Me.txtDiscAmt.FieldName = Nothing
        Me.txtDiscAmt.isCalculatedField = False
        Me.txtDiscAmt.IsSourceFromTable = False
        Me.txtDiscAmt.IsSourceFromValueList = False
        Me.txtDiscAmt.IsUnique = False
        Me.txtDiscAmt.Location = New System.Drawing.Point(283, 123)
        Me.txtDiscAmt.MendatroryField = False
        Me.txtDiscAmt.MyLinkLable1 = Nothing
        Me.txtDiscAmt.MyLinkLable2 = Nothing
        Me.txtDiscAmt.Name = "txtDiscAmt"
        Me.txtDiscAmt.ReferenceFieldDesc = Nothing
        Me.txtDiscAmt.ReferenceFieldName = Nothing
        Me.txtDiscAmt.ReferenceTableName = Nothing
        Me.txtDiscAmt.Size = New System.Drawing.Size(80, 20)
        Me.txtDiscAmt.TabIndex = 6
        Me.txtDiscAmt.Text = "0"
        Me.txtDiscAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDiscAmt.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'txtDiscPer
        '
        Me.txtDiscPer.BackColor = System.Drawing.Color.White
        Me.txtDiscPer.CalculationExpression = Nothing
        Me.txtDiscPer.DecimalPlaces = 5
        Me.txtDiscPer.FieldCode = Nothing
        Me.txtDiscPer.FieldDesc = Nothing
        Me.txtDiscPer.FieldMaxLength = 0
        Me.txtDiscPer.FieldName = Nothing
        Me.txtDiscPer.isCalculatedField = False
        Me.txtDiscPer.IsSourceFromTable = False
        Me.txtDiscPer.IsSourceFromValueList = False
        Me.txtDiscPer.IsUnique = False
        Me.txtDiscPer.Location = New System.Drawing.Point(222, 123)
        Me.txtDiscPer.MendatroryField = False
        Me.txtDiscPer.MyLinkLable1 = Nothing
        Me.txtDiscPer.MyLinkLable2 = Nothing
        Me.txtDiscPer.Name = "txtDiscPer"
        Me.txtDiscPer.ReferenceFieldDesc = Nothing
        Me.txtDiscPer.ReferenceFieldName = Nothing
        Me.txtDiscPer.ReferenceTableName = Nothing
        Me.txtDiscPer.Size = New System.Drawing.Size(39, 20)
        Me.txtDiscPer.TabIndex = 5
        Me.txtDiscPer.Text = "0"
        Me.txtDiscPer.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDiscPer.Value = New Decimal(New Integer() {0, 0, 0, 0})
        '
        'MyLabel1
        '
        Me.MyLabel1.FieldName = Nothing
        Me.MyLabel1.Location = New System.Drawing.Point(264, 125)
        Me.MyLabel1.Name = "MyLabel1"
        Me.MyLabel1.Size = New System.Drawing.Size(15, 18)
        Me.MyLabel1.TabIndex = 148
        Me.MyLabel1.Text = "%"
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
        Me.pnlCurrConv.Location = New System.Drawing.Point(96, 28)
        Me.pnlCurrConv.Name = "pnlCurrConv"
        Me.pnlCurrConv.Size = New System.Drawing.Size(688, 38)
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
        Me.txtConversionRate.Location = New System.Drawing.Point(290, 8)
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
        Me.txtCurrencyCode.Enabled = False
        Me.txtCurrencyCode.FieldCode = Nothing
        Me.txtCurrencyCode.FieldDesc = Nothing
        Me.txtCurrencyCode.FieldMaxLength = 0
        Me.txtCurrencyCode.FieldName = Nothing
        Me.txtCurrencyCode.isCalculatedField = False
        Me.txtCurrencyCode.IsSourceFromTable = False
        Me.txtCurrencyCode.IsSourceFromValueList = False
        Me.txtCurrencyCode.IsUnique = False
        Me.txtCurrencyCode.Location = New System.Drawing.Point(59, 10)
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
        Me.txtCurrencyCode.Size = New System.Drawing.Size(136, 22)
        Me.txtCurrencyCode.TabIndex = 0
        Me.txtCurrencyCode.Value = ""
        '
        'lblEffectiveFrom
        '
        Me.lblEffectiveFrom.FieldName = Nothing
        Me.lblEffectiveFrom.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEffectiveFrom.Location = New System.Drawing.Point(420, 12)
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
        Me.txtApplicableFrom.Location = New System.Drawing.Point(516, 10)
        Me.txtApplicableFrom.Name = "txtApplicableFrom"
        Me.txtApplicableFrom.Size = New System.Drawing.Size(119, 18)
        Me.txtApplicableFrom.TabIndex = 2
        Me.txtApplicableFrom.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCurrency
        '
        Me.lblCurrency.FieldName = Nothing
        Me.lblCurrency.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrency.Location = New System.Drawing.Point(4, 12)
        Me.lblCurrency.Name = "lblCurrency"
        Me.lblCurrency.Size = New System.Drawing.Size(52, 16)
        Me.lblCurrency.TabIndex = 137
        Me.lblCurrency.Text = "Currency"
        '
        'lblConvRate
        '
        Me.lblConvRate.FieldName = Nothing
        Me.lblConvRate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConvRate.Location = New System.Drawing.Point(193, 12)
        Me.lblConvRate.Name = "lblConvRate"
        Me.lblConvRate.Size = New System.Drawing.Size(91, 16)
        Me.lblConvRate.TabIndex = 139
        Me.lblConvRate.Text = "Conversion Rate"
        '
        'RadLabel32
        '
        Me.RadLabel32.FieldName = Nothing
        Me.RadLabel32.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel32.Location = New System.Drawing.Point(80, 260)
        Me.RadLabel32.Name = "RadLabel32"
        Me.RadLabel32.Size = New System.Drawing.Size(140, 16)
        Me.RadLabel32.TabIndex = 135
        Me.RadLabel32.Text = "+ Total Additional Charges"
        '
        'lblAddCharges1
        '
        Me.lblAddCharges1.AutoSize = False
        Me.lblAddCharges1.BorderVisible = True
        Me.lblAddCharges1.FieldName = Nothing
        Me.lblAddCharges1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddCharges1.Location = New System.Drawing.Point(223, 260)
        Me.lblAddCharges1.Name = "lblAddCharges1"
        Me.lblAddCharges1.Size = New System.Drawing.Size(110, 18)
        Me.lblAddCharges1.TabIndex = 11
        Me.lblAddCharges1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtComment
        '
        Me.txtComment.Location = New System.Drawing.Point(429, 68)
        Me.txtComment.Multiline = True
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(355, 164)
        Me.txtComment.TabIndex = 2
        '
        'RadLabel14
        '
        Me.RadLabel14.FieldName = Nothing
        Me.RadLabel14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel14.Location = New System.Drawing.Point(369, 72)
        Me.RadLabel14.Name = "RadLabel14"
        Me.RadLabel14.Size = New System.Drawing.Size(55, 16)
        Me.RadLabel14.TabIndex = 127
        Me.RadLabel14.Text = "Comment"
        '
        'RadLabel9
        '
        Me.RadLabel9.FieldName = Nothing
        Me.RadLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel9.Location = New System.Drawing.Point(96, 203)
        Me.RadLabel9.Name = "RadLabel9"
        Me.RadLabel9.Size = New System.Drawing.Size(120, 16)
        Me.RadLabel9.TabIndex = 120
        Me.RadLabel9.Text = "Amount After Discount"
        '
        'RadLabel27
        '
        Me.RadLabel27.FieldName = Nothing
        Me.RadLabel27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel27.Location = New System.Drawing.Point(117, 287)
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
        Me.lblTotRAmt.Location = New System.Drawing.Point(223, 285)
        Me.lblTotRAmt.Name = "lblTotRAmt"
        Me.lblTotRAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTotRAmt.TabIndex = 12
        Me.lblTotRAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel25
        '
        Me.RadLabel25.FieldName = Nothing
        Me.RadLabel25.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel25.Location = New System.Drawing.Point(140, 233)
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
        Me.lblTaxAmt.Location = New System.Drawing.Point(223, 233)
        Me.lblTaxAmt.Name = "lblTaxAmt"
        Me.lblTaxAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblTaxAmt.TabIndex = 10
        Me.lblTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtAfterDiscount
        '
        Me.lblAmtAfterDiscount.AutoSize = False
        Me.lblAmtAfterDiscount.BorderVisible = True
        Me.lblAmtAfterDiscount.FieldName = Nothing
        Me.lblAmtAfterDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtAfterDiscount.Location = New System.Drawing.Point(223, 203)
        Me.lblAmtAfterDiscount.Name = "lblAmtAfterDiscount"
        Me.lblAmtAfterDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtAfterDiscount.TabIndex = 9
        Me.lblAmtAfterDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDiscountAmt
        '
        Me.lblDiscountAmt.AutoSize = False
        Me.lblDiscountAmt.BorderVisible = True
        Me.lblDiscountAmt.FieldName = Nothing
        Me.lblDiscountAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiscountAmt.Location = New System.Drawing.Point(223, 173)
        Me.lblDiscountAmt.Name = "lblDiscountAmt"
        Me.lblDiscountAmt.Size = New System.Drawing.Size(110, 18)
        Me.lblDiscountAmt.TabIndex = 8
        Me.lblDiscountAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAmtWithDiscount
        '
        Me.lblAmtWithDiscount.AutoSize = False
        Me.lblAmtWithDiscount.BorderVisible = True
        Me.lblAmtWithDiscount.FieldName = Nothing
        Me.lblAmtWithDiscount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAmtWithDiscount.Location = New System.Drawing.Point(223, 72)
        Me.lblAmtWithDiscount.Name = "lblAmtWithDiscount"
        Me.lblAmtWithDiscount.Size = New System.Drawing.Size(110, 18)
        Me.lblAmtWithDiscount.TabIndex = 1
        Me.lblAmtWithDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadLabel22
        '
        Me.RadLabel22.FieldName = Nothing
        Me.RadLabel22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel22.Location = New System.Drawing.Point(116, 173)
        Me.RadLabel22.Name = "RadLabel22"
        Me.RadLabel22.Size = New System.Drawing.Size(99, 16)
        Me.RadLabel22.TabIndex = 121
        Me.RadLabel22.Text = "- Discount Amount"
        '
        'RadLabel19
        '
        Me.RadLabel19.FieldName = Nothing
        Me.RadLabel19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel19.Location = New System.Drawing.Point(32, 72)
        Me.RadLabel19.Name = "RadLabel19"
        Me.RadLabel19.Size = New System.Drawing.Size(186, 16)
        Me.RadLabel19.TabIndex = 119
        Me.RadLabel19.Text = "Document Amount without Discount"
        '
        'RadPageViewPage3
        '
        Me.RadPageViewPage3.Controls.Add(Me.lblLastCollectionDate)
        Me.RadPageViewPage3.Controls.Add(Me.txtLastCollectionDate)
        Me.RadPageViewPage3.Controls.Add(Me.MyLabel5)
        Me.RadPageViewPage3.Controls.Add(Me.lblTotRAmt1)
        Me.RadPageViewPage3.Controls.Add(Me.lblDCSDemand)
        Me.RadPageViewPage3.Controls.Add(Me.txtDCSDemandNo)
        Me.RadPageViewPage3.Controls.Add(Me.lblUnbilledMilk)
        Me.RadPageViewPage3.Controls.Add(Me.lblUnbilledMilkAmt)
        Me.RadPageViewPage3.Controls.Add(Me.pnlTCS)
        Me.RadPageViewPage3.Controls.Add(Me.RadPanel2)
        Me.RadPageViewPage3.Controls.Add(Me.chkDCS)
        Me.RadPageViewPage3.Controls.Add(Me.btnGatePassPrint)
        Me.RadPageViewPage3.Controls.Add(Me.btnPrint)
        Me.RadPageViewPage3.Controls.Add(Me.btnCreateDO)
        Me.RadPageViewPage3.Controls.Add(Me.txtSalesman1)
        Me.RadPageViewPage3.Controls.Add(Me.lblSalesmandesc1)
        Me.RadPageViewPage3.Controls.Add(Me.lblsalesman1)
        Me.RadPageViewPage3.Controls.Add(Me.txtSalesman)
        Me.RadPageViewPage3.Controls.Add(Me.lblSalesman)
        Me.RadPageViewPage3.Controls.Add(Me.MyLabel2)
        Me.RadPageViewPage3.Controls.Add(Me.lblBoothStation)
        Me.RadPageViewPage3.Controls.Add(Me.txtEx_Factory_Date)
        Me.RadPageViewPage3.Controls.Add(Me.lbl_ExFactoryDate)
        Me.RadPageViewPage3.Controls.Add(Me.MyLabel12)
        Me.RadPageViewPage3.Controls.Add(Me.txtPONo)
        Me.RadPageViewPage3.Controls.Add(Me.MyLabel4)
        Me.RadPageViewPage3.Controls.Add(Me.txtCustPODate)
        Me.RadPageViewPage3.Controls.Add(Me.chkGatePass)
        Me.RadPageViewPage3.Controls.Add(Me.ItemTypePanel)
        Me.RadPageViewPage3.Controls.Add(Me.Panel3)
        Me.RadPageViewPage3.Controls.Add(Me.PanelSearchItem)
        Me.RadPageViewPage3.Controls.Add(Me.Panel4)
        Me.RadPageViewPage3.Controls.Add(Me.Panel1)
        Me.RadPageViewPage3.ItemSize = New System.Drawing.SizeF(50.0!, 26.0!)
        Me.RadPageViewPage3.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage3.Name = "RadPageViewPage3"
        Me.RadPageViewPage3.Size = New System.Drawing.Size(1260, 409)
        Me.RadPageViewPage3.Text = "Others"
        '
        'lblLastCollectionDate
        '
        Me.lblLastCollectionDate.FieldName = Nothing
        Me.lblLastCollectionDate.Location = New System.Drawing.Point(687, 380)
        Me.lblLastCollectionDate.Name = "lblLastCollectionDate"
        Me.lblLastCollectionDate.Size = New System.Drawing.Size(106, 18)
        Me.lblLastCollectionDate.TabIndex = 1570
        Me.lblLastCollectionDate.Text = "Last Collection Date"
        '
        'txtLastCollectionDate
        '
        Me.txtLastCollectionDate.AutoSize = False
        Me.txtLastCollectionDate.BorderVisible = True
        Me.txtLastCollectionDate.FieldName = Nothing
        Me.txtLastCollectionDate.Location = New System.Drawing.Point(797, 379)
        Me.txtLastCollectionDate.Name = "txtLastCollectionDate"
        Me.txtLastCollectionDate.Size = New System.Drawing.Size(86, 19)
        Me.txtLastCollectionDate.TabIndex = 1571
        Me.txtLastCollectionDate.TextWrap = False
        '
        'MyLabel5
        '
        Me.MyLabel5.FieldName = Nothing
        Me.MyLabel5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel5.Location = New System.Drawing.Point(771, 205)
        Me.MyLabel5.Name = "MyLabel5"
        Me.MyLabel5.Size = New System.Drawing.Size(108, 16)
        Me.MyLabel5.TabIndex = 1564
        Me.MyLabel5.Text = "Document Amount"
        '
        'lblTotRAmt1
        '
        Me.lblTotRAmt1.AutoSize = False
        Me.lblTotRAmt1.BorderVisible = True
        Me.lblTotRAmt1.FieldName = Nothing
        Me.lblTotRAmt1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotRAmt1.Location = New System.Drawing.Point(887, 203)
        Me.lblTotRAmt1.Name = "lblTotRAmt1"
        Me.lblTotRAmt1.Size = New System.Drawing.Size(104, 20)
        Me.lblTotRAmt1.TabIndex = 1565
        Me.lblTotRAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDCSDemand
        '
        Me.lblDCSDemand.FieldName = Nothing
        Me.lblDCSDemand.Location = New System.Drawing.Point(859, 236)
        Me.lblDCSDemand.Name = "lblDCSDemand"
        Me.lblDCSDemand.Size = New System.Drawing.Size(91, 18)
        Me.lblDCSDemand.TabIndex = 1562
        Me.lblDCSDemand.Text = "DCS Demand No"
        '
        'txtDCSDemandNo
        '
        Me.txtDCSDemandNo.AutoSize = False
        Me.txtDCSDemandNo.BorderVisible = True
        Me.txtDCSDemandNo.FieldName = Nothing
        Me.txtDCSDemandNo.Location = New System.Drawing.Point(956, 236)
        Me.txtDCSDemandNo.Name = "txtDCSDemandNo"
        Me.txtDCSDemandNo.Size = New System.Drawing.Size(133, 19)
        Me.txtDCSDemandNo.TabIndex = 1563
        Me.txtDCSDemandNo.TextWrap = False
        '
        'lblUnbilledMilk
        '
        Me.lblUnbilledMilk.FieldName = Nothing
        Me.lblUnbilledMilk.Location = New System.Drawing.Point(651, 236)
        Me.lblUnbilledMilk.Name = "lblUnbilledMilk"
        Me.lblUnbilledMilk.Size = New System.Drawing.Size(96, 18)
        Me.lblUnbilledMilk.TabIndex = 1560
        Me.lblUnbilledMilk.Text = "Unbilled Milk Amt"
        '
        'lblUnbilledMilkAmt
        '
        Me.lblUnbilledMilkAmt.AutoSize = False
        Me.lblUnbilledMilkAmt.BorderVisible = True
        Me.lblUnbilledMilkAmt.FieldName = Nothing
        Me.lblUnbilledMilkAmt.Location = New System.Drawing.Point(753, 236)
        Me.lblUnbilledMilkAmt.Name = "lblUnbilledMilkAmt"
        Me.lblUnbilledMilkAmt.Size = New System.Drawing.Size(101, 19)
        Me.lblUnbilledMilkAmt.TabIndex = 1561
        Me.lblUnbilledMilkAmt.TextWrap = False
        '
        'pnlTCS
        '
        Me.pnlTCS.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTCS.Controls.Add(Me.lblTCSBaseAmt)
        Me.pnlTCS.Controls.Add(Me.txtTCSBaseAmt)
        Me.pnlTCS.Controls.Add(Me.MyLabel17)
        Me.pnlTCS.Controls.Add(Me.lblTCSAmount)
        Me.pnlTCS.Location = New System.Drawing.Point(566, 345)
        Me.pnlTCS.Name = "pnlTCS"
        Me.pnlTCS.Size = New System.Drawing.Size(388, 32)
        Me.pnlTCS.TabIndex = 1532
        '
        'lblTCSBaseAmt
        '
        Me.lblTCSBaseAmt.FieldName = Nothing
        Me.lblTCSBaseAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTCSBaseAmt.Location = New System.Drawing.Point(5, 8)
        Me.lblTCSBaseAmt.Name = "lblTCSBaseAmt"
        Me.lblTCSBaseAmt.Size = New System.Drawing.Size(84, 16)
        Me.lblTCSBaseAmt.TabIndex = 137
        Me.lblTCSBaseAmt.Text = "TCS Base Amt"
        '
        'txtTCSBaseAmt
        '
        Me.txtTCSBaseAmt.AutoSize = False
        Me.txtTCSBaseAmt.BorderVisible = True
        Me.txtTCSBaseAmt.FieldName = Nothing
        Me.txtTCSBaseAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTCSBaseAmt.Location = New System.Drawing.Point(92, 6)
        Me.txtTCSBaseAmt.Name = "txtTCSBaseAmt"
        Me.txtTCSBaseAmt.Size = New System.Drawing.Size(104, 20)
        Me.txtTCSBaseAmt.TabIndex = 138
        Me.txtTCSBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'MyLabel17
        '
        Me.MyLabel17.FieldName = Nothing
        Me.MyLabel17.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel17.Location = New System.Drawing.Point(198, 8)
        Me.MyLabel17.Name = "MyLabel17"
        Me.MyLabel17.Size = New System.Drawing.Size(75, 16)
        Me.MyLabel17.TabIndex = 126
        Me.MyLabel17.Text = "TCS Amount"
        '
        'lblTCSAmount
        '
        Me.lblTCSAmount.AutoSize = False
        Me.lblTCSAmount.BorderVisible = True
        Me.lblTCSAmount.FieldName = Nothing
        Me.lblTCSAmount.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTCSAmount.Location = New System.Drawing.Point(278, 6)
        Me.lblTCSAmount.Name = "lblTCSAmount"
        Me.lblTCSAmount.Size = New System.Drawing.Size(104, 20)
        Me.lblTCSAmount.TabIndex = 136
        Me.lblTCSAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'RadPanel2
        '
        Me.RadPanel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadPanel2.Controls.Add(Me.LblBox)
        Me.RadPanel2.Controls.Add(Me.txtBox)
        Me.RadPanel2.Controls.Add(Me.lblCrate)
        Me.RadPanel2.Controls.Add(Me.txtCrate)
        Me.RadPanel2.Controls.Add(Me.lblCan)
        Me.RadPanel2.Controls.Add(Me.txtCan)
        Me.RadPanel2.Location = New System.Drawing.Point(580, 307)
        Me.RadPanel2.Name = "RadPanel2"
        Me.RadPanel2.Size = New System.Drawing.Size(294, 32)
        Me.RadPanel2.TabIndex = 1531
        '
        'LblBox
        '
        Me.LblBox.FieldName = Nothing
        Me.LblBox.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblBox.Location = New System.Drawing.Point(199, 8)
        Me.LblBox.Name = "LblBox"
        Me.LblBox.Size = New System.Drawing.Size(28, 16)
        Me.LblBox.TabIndex = 139
        Me.LblBox.Text = "Box"
        '
        'txtBox
        '
        Me.txtBox.AutoSize = False
        Me.txtBox.BorderVisible = True
        Me.txtBox.FieldName = Nothing
        Me.txtBox.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox.Location = New System.Drawing.Point(232, 6)
        Me.txtBox.Name = "txtBox"
        Me.txtBox.Size = New System.Drawing.Size(54, 20)
        Me.txtBox.TabIndex = 140
        Me.txtBox.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCrate
        '
        Me.lblCrate.FieldName = Nothing
        Me.lblCrate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCrate.Location = New System.Drawing.Point(99, 8)
        Me.lblCrate.Name = "lblCrate"
        Me.lblCrate.Size = New System.Drawing.Size(35, 16)
        Me.lblCrate.TabIndex = 137
        Me.lblCrate.Text = "Crate"
        '
        'txtCrate
        '
        Me.txtCrate.AutoSize = False
        Me.txtCrate.BorderVisible = True
        Me.txtCrate.FieldName = Nothing
        Me.txtCrate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCrate.Location = New System.Drawing.Point(139, 6)
        Me.txtCrate.Name = "txtCrate"
        Me.txtCrate.Size = New System.Drawing.Size(54, 20)
        Me.txtCrate.TabIndex = 138
        Me.txtCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCan
        '
        Me.lblCan.FieldName = Nothing
        Me.lblCan.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCan.Location = New System.Drawing.Point(6, 8)
        Me.lblCan.Name = "lblCan"
        Me.lblCan.Size = New System.Drawing.Size(31, 16)
        Me.lblCan.TabIndex = 126
        Me.lblCan.Text = "CAN"
        '
        'txtCan
        '
        Me.txtCan.AutoSize = False
        Me.txtCan.BorderVisible = True
        Me.txtCan.FieldName = Nothing
        Me.txtCan.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCan.Location = New System.Drawing.Point(40, 6)
        Me.txtCan.Name = "txtCan"
        Me.txtCan.Size = New System.Drawing.Size(54, 20)
        Me.txtCan.TabIndex = 136
        Me.txtCan.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkDCS
        '
        Me.chkDCS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDCS.Location = New System.Drawing.Point(580, 287)
        Me.chkDCS.Name = "chkDCS"
        Me.chkDCS.Size = New System.Drawing.Size(44, 16)
        Me.chkDCS.TabIndex = 1530
        Me.chkDCS.Text = "DCS"
        Me.chkDCS.Visible = False
        '
        'btnGatePassPrint
        '
        Me.btnGatePassPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGatePassPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGatePassPrint.Location = New System.Drawing.Point(892, 267)
        Me.btnGatePassPrint.Name = "btnGatePassPrint"
        Me.btnGatePassPrint.Size = New System.Drawing.Size(92, 22)
        Me.btnGatePassPrint.TabIndex = 1529
        Me.btnGatePassPrint.Text = "Gate Pass Print"
        '
        'btnPrint
        '
        Me.btnPrint.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrint.Location = New System.Drawing.Point(986, 267)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.Size = New System.Drawing.Size(47, 22)
        Me.btnPrint.TabIndex = 1528
        Me.btnPrint.Text = "Print"
        Me.btnPrint.Visible = False
        '
        'btnCreateDO
        '
        Me.btnCreateDO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCreateDO.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateDO.Location = New System.Drawing.Point(776, 267)
        Me.btnCreateDO.Name = "btnCreateDO"
        Me.btnCreateDO.Size = New System.Drawing.Size(110, 22)
        Me.btnCreateDO.TabIndex = 1527
        Me.btnCreateDO.Text = "Create and Post DO"
        '
        'txtSalesman1
        '
        Me.txtSalesman1.CalculationExpression = Nothing
        Me.txtSalesman1.FieldCode = Nothing
        Me.txtSalesman1.FieldDesc = Nothing
        Me.txtSalesman1.FieldMaxLength = 0
        Me.txtSalesman1.FieldName = Nothing
        Me.txtSalesman1.isCalculatedField = False
        Me.txtSalesman1.IsSourceFromTable = False
        Me.txtSalesman1.IsSourceFromValueList = False
        Me.txtSalesman1.IsUnique = False
        Me.txtSalesman1.Location = New System.Drawing.Point(291, 236)
        Me.txtSalesman1.MendatroryField = False
        Me.txtSalesman1.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSalesman1.MyLinkLable1 = Me.lblRouteNo
        Me.txtSalesman1.MyLinkLable2 = Nothing
        Me.txtSalesman1.MyReadOnly = False
        Me.txtSalesman1.MyShowMasterFormButton = False
        Me.txtSalesman1.Name = "txtSalesman1"
        Me.txtSalesman1.ReferenceFieldDesc = Nothing
        Me.txtSalesman1.ReferenceFieldName = Nothing
        Me.txtSalesman1.ReferenceTableName = Nothing
        Me.txtSalesman1.Size = New System.Drawing.Size(115, 22)
        Me.txtSalesman1.TabIndex = 1526
        Me.txtSalesman1.Value = ""
        '
        'lblSalesmandesc1
        '
        Me.lblSalesmandesc1.AutoSize = False
        Me.lblSalesmandesc1.BorderVisible = True
        Me.lblSalesmandesc1.FieldName = Nothing
        Me.lblSalesmandesc1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSalesmandesc1.Location = New System.Drawing.Point(412, 219)
        Me.lblSalesmandesc1.Name = "lblSalesmandesc1"
        Me.lblSalesmandesc1.Size = New System.Drawing.Size(227, 17)
        Me.lblSalesmandesc1.TabIndex = 1524
        Me.lblSalesmandesc1.TextWrap = False
        '
        'lblsalesman1
        '
        Me.lblsalesman1.FieldName = Nothing
        Me.lblsalesman1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblsalesman1.Location = New System.Drawing.Point(222, 219)
        Me.lblsalesman1.Name = "lblsalesman1"
        Me.lblsalesman1.Size = New System.Drawing.Size(57, 16)
        Me.lblsalesman1.TabIndex = 1525
        Me.lblsalesman1.Text = "Salesman"
        '
        'txtSalesman
        '
        Me.txtSalesman.CalculationExpression = Nothing
        Me.txtSalesman.FieldCode = Nothing
        Me.txtSalesman.FieldDesc = Nothing
        Me.txtSalesman.FieldMaxLength = 0
        Me.txtSalesman.FieldName = Nothing
        Me.txtSalesman.isCalculatedField = False
        Me.txtSalesman.IsSourceFromTable = False
        Me.txtSalesman.IsSourceFromValueList = False
        Me.txtSalesman.IsUnique = False
        Me.txtSalesman.Location = New System.Drawing.Point(138, 373)
        Me.txtSalesman.MendatroryField = False
        Me.txtSalesman.MyFont = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSalesman.MyLinkLable1 = Me.lblRouteNo
        Me.txtSalesman.MyLinkLable2 = Nothing
        Me.txtSalesman.MyReadOnly = False
        Me.txtSalesman.MyShowMasterFormButton = False
        Me.txtSalesman.Name = "txtSalesman"
        Me.txtSalesman.ReferenceFieldDesc = Nothing
        Me.txtSalesman.ReferenceFieldName = Nothing
        Me.txtSalesman.ReferenceTableName = Nothing
        Me.txtSalesman.Size = New System.Drawing.Size(115, 18)
        Me.txtSalesman.TabIndex = 1520
        Me.txtSalesman.Value = ""
        '
        'lblSalesman
        '
        Me.lblSalesman.AutoSize = False
        Me.lblSalesman.BorderVisible = True
        Me.lblSalesman.FieldName = Nothing
        Me.lblSalesman.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSalesman.Location = New System.Drawing.Point(258, 340)
        Me.lblSalesman.Name = "lblSalesman"
        Me.lblSalesman.Size = New System.Drawing.Size(227, 16)
        Me.lblSalesman.TabIndex = 1518
        Me.lblSalesman.TextWrap = False
        Me.lblSalesman.Visible = False
        '
        'MyLabel2
        '
        Me.MyLabel2.FieldName = Nothing
        Me.MyLabel2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel2.Location = New System.Drawing.Point(58, 339)
        Me.MyLabel2.Name = "MyLabel2"
        Me.MyLabel2.Size = New System.Drawing.Size(57, 16)
        Me.MyLabel2.TabIndex = 1519
        Me.MyLabel2.Text = "Salesman"
        '
        'lblBoothStation
        '
        Me.lblBoothStation.AutoSize = False
        Me.lblBoothStation.BorderVisible = True
        Me.lblBoothStation.FieldName = Nothing
        Me.lblBoothStation.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBoothStation.Location = New System.Drawing.Point(176, 321)
        Me.lblBoothStation.Name = "lblBoothStation"
        Me.lblBoothStation.Size = New System.Drawing.Size(347, 18)
        Me.lblBoothStation.TabIndex = 1515
        Me.lblBoothStation.TextWrap = False
        Me.lblBoothStation.Visible = False
        '
        'txtEx_Factory_Date
        '
        Me.txtEx_Factory_Date.CalculationExpression = Nothing
        Me.txtEx_Factory_Date.CustomFormat = "dd/MM/yyyy"
        Me.txtEx_Factory_Date.FieldCode = Nothing
        Me.txtEx_Factory_Date.FieldDesc = Nothing
        Me.txtEx_Factory_Date.FieldMaxLength = 0
        Me.txtEx_Factory_Date.FieldName = Nothing
        Me.txtEx_Factory_Date.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEx_Factory_Date.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtEx_Factory_Date.isCalculatedField = False
        Me.txtEx_Factory_Date.IsSourceFromTable = False
        Me.txtEx_Factory_Date.IsSourceFromValueList = False
        Me.txtEx_Factory_Date.IsUnique = False
        Me.txtEx_Factory_Date.Location = New System.Drawing.Point(349, 271)
        Me.txtEx_Factory_Date.MendatroryField = False
        Me.txtEx_Factory_Date.MinDate = New Date(1990, 1, 1, 0, 0, 0, 0)
        Me.txtEx_Factory_Date.MyLinkLable1 = Me.RadLabel4
        Me.txtEx_Factory_Date.MyLinkLable2 = Nothing
        Me.txtEx_Factory_Date.Name = "txtEx_Factory_Date"
        Me.txtEx_Factory_Date.NullDate = New Date(1990, 1, 1, 0, 0, 0, 0)
        Me.txtEx_Factory_Date.ReferenceFieldDesc = Nothing
        Me.txtEx_Factory_Date.ReferenceFieldName = Nothing
        Me.txtEx_Factory_Date.ReferenceTableName = Nothing
        Me.txtEx_Factory_Date.ShowCheckBox = True
        Me.txtEx_Factory_Date.Size = New System.Drawing.Size(105, 18)
        Me.txtEx_Factory_Date.TabIndex = 1511
        Me.txtEx_Factory_Date.TabStop = False
        Me.txtEx_Factory_Date.Value = New Date(1990, 1, 1, 0, 0, 0, 0)
        '
        'lbl_ExFactoryDate
        '
        Me.lbl_ExFactoryDate.FieldName = Nothing
        Me.lbl_ExFactoryDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_ExFactoryDate.Location = New System.Drawing.Point(255, 273)
        Me.lbl_ExFactoryDate.Name = "lbl_ExFactoryDate"
        Me.lbl_ExFactoryDate.Size = New System.Drawing.Size(88, 16)
        Me.lbl_ExFactoryDate.TabIndex = 1510
        Me.lbl_ExFactoryDate.Text = "Ex-Factory Date"
        '
        'MyLabel12
        '
        Me.MyLabel12.FieldName = Nothing
        Me.MyLabel12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel12.Location = New System.Drawing.Point(31, 249)
        Me.MyLabel12.Name = "MyLabel12"
        Me.MyLabel12.Size = New System.Drawing.Size(92, 16)
        Me.MyLabel12.TabIndex = 1506
        Me.MyLabel12.Text = "Customer PO No"
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
        Me.txtPONo.Location = New System.Drawing.Point(124, 247)
        Me.txtPONo.MaxLength = 200
        Me.txtPONo.MendatroryField = False
        Me.txtPONo.MyLinkLable1 = Me.MyLabel12
        Me.txtPONo.MyLinkLable2 = Nothing
        Me.txtPONo.Name = "txtPONo"
        Me.txtPONo.ReferenceFieldDesc = Nothing
        Me.txtPONo.ReferenceFieldName = Nothing
        Me.txtPONo.ReferenceTableName = Nothing
        Me.txtPONo.Size = New System.Drawing.Size(142, 18)
        Me.txtPONo.TabIndex = 1507
        '
        'MyLabel4
        '
        Me.MyLabel4.FieldName = Nothing
        Me.MyLabel4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel4.Location = New System.Drawing.Point(272, 250)
        Me.MyLabel4.Name = "MyLabel4"
        Me.MyLabel4.Size = New System.Drawing.Size(74, 16)
        Me.MyLabel4.TabIndex = 1508
        Me.MyLabel4.Text = "Cust Po Date"
        '
        'txtCustPODate
        '
        Me.txtCustPODate.CalculationExpression = Nothing
        Me.txtCustPODate.CustomFormat = "dd/MM/yyyy"
        Me.txtCustPODate.FieldCode = Nothing
        Me.txtCustPODate.FieldDesc = Nothing
        Me.txtCustPODate.FieldMaxLength = 0
        Me.txtCustPODate.FieldName = Nothing
        Me.txtCustPODate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCustPODate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.txtCustPODate.isCalculatedField = False
        Me.txtCustPODate.IsSourceFromTable = False
        Me.txtCustPODate.IsSourceFromValueList = False
        Me.txtCustPODate.IsUnique = False
        Me.txtCustPODate.Location = New System.Drawing.Point(350, 248)
        Me.txtCustPODate.MendatroryField = False
        Me.txtCustPODate.MinDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtCustPODate.MyLinkLable1 = Nothing
        Me.txtCustPODate.MyLinkLable2 = Nothing
        Me.txtCustPODate.Name = "txtCustPODate"
        Me.txtCustPODate.NullDate = New Date(1753, 1, 1, 0, 0, 0, 0)
        Me.txtCustPODate.ReferenceFieldDesc = Nothing
        Me.txtCustPODate.ReferenceFieldName = Nothing
        Me.txtCustPODate.ReferenceTableName = Nothing
        Me.txtCustPODate.ShowCheckBox = True
        Me.txtCustPODate.Size = New System.Drawing.Size(103, 18)
        Me.txtCustPODate.TabIndex = 1509
        Me.txtCustPODate.TabStop = False
        Me.txtCustPODate.Text = "13/06/2011"
        Me.txtCustPODate.Value = New Date(2011, 6, 13, 11, 29, 49, 421)
        '
        'chkGatePass
        '
        Me.chkGatePass.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkGatePass.Location = New System.Drawing.Point(35, 227)
        Me.chkGatePass.Name = "chkGatePass"
        Me.chkGatePass.Size = New System.Drawing.Size(73, 16)
        Me.chkGatePass.TabIndex = 1503
        Me.chkGatePass.Text = "Gate Pass"
        '
        'ItemTypePanel
        '
        Me.ItemTypePanel.Controls.Add(Me.rbtn_Ambient)
        Me.ItemTypePanel.Controls.Add(Me.rbtn_Fresh)
        Me.ItemTypePanel.Controls.Add(Me.MyLabel40)
        Me.ItemTypePanel.Location = New System.Drawing.Point(784, 8)
        Me.ItemTypePanel.Name = "ItemTypePanel"
        Me.ItemTypePanel.Size = New System.Drawing.Size(67, 57)
        Me.ItemTypePanel.TabIndex = 1502
        '
        'rbtn_Ambient
        '
        Me.rbtn_Ambient.Location = New System.Drawing.Point(5, 29)
        Me.rbtn_Ambient.Name = "rbtn_Ambient"
        Me.rbtn_Ambient.Size = New System.Drawing.Size(63, 18)
        Me.rbtn_Ambient.TabIndex = 39
        Me.rbtn_Ambient.TabStop = False
        Me.rbtn_Ambient.Text = "Ambient"
        '
        'rbtn_Fresh
        '
        Me.rbtn_Fresh.CheckState = System.Windows.Forms.CheckState.Checked
        Me.rbtn_Fresh.Location = New System.Drawing.Point(5, 11)
        Me.rbtn_Fresh.Name = "rbtn_Fresh"
        Me.rbtn_Fresh.Size = New System.Drawing.Size(47, 18)
        Me.rbtn_Fresh.TabIndex = 38
        Me.rbtn_Fresh.Text = "Fresh"
        Me.rbtn_Fresh.ToggleState = Telerik.WinControls.Enumerations.ToggleState.[On]
        '
        'MyLabel40
        '
        Me.MyLabel40.FieldName = Nothing
        Me.MyLabel40.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel40.Location = New System.Drawing.Point(3, -3)
        Me.MyLabel40.Name = "MyLabel40"
        Me.MyLabel40.Size = New System.Drawing.Size(57, 16)
        Me.MyLabel40.TabIndex = 36
        Me.MyLabel40.Text = "Item Type"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.MyLabel25)
        Me.Panel3.Controls.Add(Me.lblARSecurity)
        Me.Panel3.Controls.Add(Me.MyLabel24)
        Me.Panel3.Controls.Add(Me.lblReverseAdvanceSec)
        Me.Panel3.Controls.Add(Me.MyLabel23)
        Me.Panel3.Controls.Add(Me.lblReverseRefund)
        Me.Panel3.Controls.Add(Me.MyLabel35)
        Me.Panel3.Controls.Add(Me.lblTotalOutstansing)
        Me.Panel3.Controls.Add(Me.MyLabel33)
        Me.Panel3.Controls.Add(Me.lblRefund)
        Me.Panel3.Controls.Add(Me.MyLabel30)
        Me.Panel3.Controls.Add(Me.lblLedgerOutstanding)
        Me.Panel3.Controls.Add(Me.MyLabel28)
        Me.Panel3.Controls.Add(Me.lblShortcloseDO)
        Me.Panel3.Controls.Add(Me.MyLabel26)
        Me.Panel3.Controls.Add(Me.lblPendingDO)
        Me.Panel3.Controls.Add(Me.MyLabel22)
        Me.Panel3.Controls.Add(Me.lblAdvanceSecurity)
        Me.Panel3.Controls.Add(Me.MyLabel15)
        Me.Panel3.Controls.Add(Me.lblCreditLimit)
        Me.Panel3.Location = New System.Drawing.Point(529, 11)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(210, 194)
        Me.Panel3.TabIndex = 1501
        '
        'MyLabel25
        '
        Me.MyLabel25.FieldName = Nothing
        Me.MyLabel25.Location = New System.Drawing.Point(3, 59)
        Me.MyLabel25.Name = "MyLabel25"
        Me.MyLabel25.Size = New System.Drawing.Size(108, 18)
        Me.MyLabel25.TabIndex = 76
        Me.MyLabel25.Text = "AR Credit Security(-)"
        '
        'lblARSecurity
        '
        Me.lblARSecurity.AutoSize = False
        Me.lblARSecurity.BorderVisible = True
        Me.lblARSecurity.FieldName = Nothing
        Me.lblARSecurity.Location = New System.Drawing.Point(121, 58)
        Me.lblARSecurity.Name = "lblARSecurity"
        Me.lblARSecurity.Size = New System.Drawing.Size(87, 19)
        Me.lblARSecurity.TabIndex = 75
        Me.lblARSecurity.TextWrap = False
        '
        'MyLabel24
        '
        Me.MyLabel24.FieldName = Nothing
        Me.MyLabel24.Location = New System.Drawing.Point(3, 40)
        Me.MyLabel24.Name = "MyLabel24"
        Me.MyLabel24.Size = New System.Drawing.Size(118, 18)
        Me.MyLabel24.TabIndex = 74
        Me.MyLabel24.Text = "Revese Adv Security(-)"
        '
        'lblReverseAdvanceSec
        '
        Me.lblReverseAdvanceSec.AutoSize = False
        Me.lblReverseAdvanceSec.BorderVisible = True
        Me.lblReverseAdvanceSec.FieldName = Nothing
        Me.lblReverseAdvanceSec.Location = New System.Drawing.Point(121, 39)
        Me.lblReverseAdvanceSec.Name = "lblReverseAdvanceSec"
        Me.lblReverseAdvanceSec.Size = New System.Drawing.Size(87, 19)
        Me.lblReverseAdvanceSec.TabIndex = 73
        Me.lblReverseAdvanceSec.TextWrap = False
        '
        'MyLabel23
        '
        Me.MyLabel23.FieldName = Nothing
        Me.MyLabel23.Location = New System.Drawing.Point(3, 154)
        Me.MyLabel23.Name = "MyLabel23"
        Me.MyLabel23.Size = New System.Drawing.Size(119, 18)
        Me.MyLabel23.TabIndex = 72
        Me.MyLabel23.Text = "Reverse Refund Sec(+)"
        '
        'lblReverseRefund
        '
        Me.lblReverseRefund.AutoSize = False
        Me.lblReverseRefund.BorderVisible = True
        Me.lblReverseRefund.FieldName = Nothing
        Me.lblReverseRefund.Location = New System.Drawing.Point(121, 153)
        Me.lblReverseRefund.Name = "lblReverseRefund"
        Me.lblReverseRefund.Size = New System.Drawing.Size(87, 19)
        Me.lblReverseRefund.TabIndex = 71
        Me.lblReverseRefund.TextWrap = False
        '
        'MyLabel35
        '
        Me.MyLabel35.FieldName = Nothing
        Me.MyLabel35.Location = New System.Drawing.Point(4, 173)
        Me.MyLabel35.Name = "MyLabel35"
        Me.MyLabel35.Size = New System.Drawing.Size(96, 18)
        Me.MyLabel35.TabIndex = 70
        Me.MyLabel35.Text = "Total Outstanding"
        '
        'lblTotalOutstansing
        '
        Me.lblTotalOutstansing.AutoSize = False
        Me.lblTotalOutstansing.BorderVisible = True
        Me.lblTotalOutstansing.FieldName = Nothing
        Me.lblTotalOutstansing.Location = New System.Drawing.Point(122, 172)
        Me.lblTotalOutstansing.Name = "lblTotalOutstansing"
        Me.lblTotalOutstansing.Size = New System.Drawing.Size(87, 19)
        Me.lblTotalOutstansing.TabIndex = 69
        Me.lblTotalOutstansing.TextWrap = False
        '
        'MyLabel33
        '
        Me.MyLabel33.FieldName = Nothing
        Me.MyLabel33.Location = New System.Drawing.Point(3, 135)
        Me.MyLabel33.Name = "MyLabel33"
        Me.MyLabel33.Size = New System.Drawing.Size(99, 18)
        Me.MyLabel33.TabIndex = 68
        Me.MyLabel33.Text = "Refund Security (-)"
        '
        'lblRefund
        '
        Me.lblRefund.AutoSize = False
        Me.lblRefund.BorderVisible = True
        Me.lblRefund.FieldName = Nothing
        Me.lblRefund.Location = New System.Drawing.Point(121, 134)
        Me.lblRefund.Name = "lblRefund"
        Me.lblRefund.Size = New System.Drawing.Size(87, 19)
        Me.lblRefund.TabIndex = 67
        Me.lblRefund.TextWrap = False
        '
        'MyLabel30
        '
        Me.MyLabel30.FieldName = Nothing
        Me.MyLabel30.Location = New System.Drawing.Point(3, 116)
        Me.MyLabel30.Name = "MyLabel30"
        Me.MyLabel30.Size = New System.Drawing.Size(115, 18)
        Me.MyLabel30.TabIndex = 66
        Me.MyLabel30.Text = "Ledger Outstansing(-)"
        '
        'lblLedgerOutstanding
        '
        Me.lblLedgerOutstanding.AutoSize = False
        Me.lblLedgerOutstanding.BorderVisible = True
        Me.lblLedgerOutstanding.FieldName = Nothing
        Me.lblLedgerOutstanding.Location = New System.Drawing.Point(121, 115)
        Me.lblLedgerOutstanding.Name = "lblLedgerOutstanding"
        Me.lblLedgerOutstanding.Size = New System.Drawing.Size(87, 19)
        Me.lblLedgerOutstanding.TabIndex = 65
        Me.lblLedgerOutstanding.TextWrap = False
        '
        'MyLabel28
        '
        Me.MyLabel28.FieldName = Nothing
        Me.MyLabel28.Location = New System.Drawing.Point(4, 97)
        Me.MyLabel28.Name = "MyLabel28"
        Me.MyLabel28.Size = New System.Drawing.Size(109, 18)
        Me.MyLabel28.TabIndex = 64
        Me.MyLabel28.Text = "UnPostedDispatch(-)"
        '
        'lblShortcloseDO
        '
        Me.lblShortcloseDO.AutoSize = False
        Me.lblShortcloseDO.BorderVisible = True
        Me.lblShortcloseDO.FieldName = Nothing
        Me.lblShortcloseDO.Location = New System.Drawing.Point(121, 96)
        Me.lblShortcloseDO.Name = "lblShortcloseDO"
        Me.lblShortcloseDO.Size = New System.Drawing.Size(87, 19)
        Me.lblShortcloseDO.TabIndex = 63
        Me.lblShortcloseDO.TextWrap = False
        '
        'MyLabel26
        '
        Me.MyLabel26.FieldName = Nothing
        Me.MyLabel26.Location = New System.Drawing.Point(4, 78)
        Me.MyLabel26.Name = "MyLabel26"
        Me.MyLabel26.Size = New System.Drawing.Size(116, 18)
        Me.MyLabel26.TabIndex = 62
        Me.MyLabel26.Text = "Pending Posted DO(-)"
        '
        'lblPendingDO
        '
        Me.lblPendingDO.AutoSize = False
        Me.lblPendingDO.BorderVisible = True
        Me.lblPendingDO.FieldName = Nothing
        Me.lblPendingDO.Location = New System.Drawing.Point(121, 77)
        Me.lblPendingDO.Name = "lblPendingDO"
        Me.lblPendingDO.Size = New System.Drawing.Size(87, 19)
        Me.lblPendingDO.TabIndex = 61
        Me.lblPendingDO.TextWrap = False
        '
        'MyLabel22
        '
        Me.MyLabel22.FieldName = Nothing
        Me.MyLabel22.Location = New System.Drawing.Point(3, 21)
        Me.MyLabel22.Name = "MyLabel22"
        Me.MyLabel22.Size = New System.Drawing.Size(106, 18)
        Me.MyLabel22.TabIndex = 58
        Me.MyLabel22.Text = "Advance Security(+)"
        '
        'lblAdvanceSecurity
        '
        Me.lblAdvanceSecurity.AutoSize = False
        Me.lblAdvanceSecurity.BorderVisible = True
        Me.lblAdvanceSecurity.FieldName = Nothing
        Me.lblAdvanceSecurity.Location = New System.Drawing.Point(121, 20)
        Me.lblAdvanceSecurity.Name = "lblAdvanceSecurity"
        Me.lblAdvanceSecurity.Size = New System.Drawing.Size(87, 19)
        Me.lblAdvanceSecurity.TabIndex = 57
        Me.lblAdvanceSecurity.TextWrap = False
        '
        'MyLabel15
        '
        Me.MyLabel15.FieldName = Nothing
        Me.MyLabel15.Location = New System.Drawing.Point(3, 3)
        Me.MyLabel15.Name = "MyLabel15"
        Me.MyLabel15.Size = New System.Drawing.Size(78, 18)
        Me.MyLabel15.TabIndex = 56
        Me.MyLabel15.Text = "Credit Limit(+)"
        '
        'lblCreditLimit
        '
        Me.lblCreditLimit.AutoSize = False
        Me.lblCreditLimit.BorderVisible = True
        Me.lblCreditLimit.FieldName = Nothing
        Me.lblCreditLimit.Location = New System.Drawing.Point(121, 2)
        Me.lblCreditLimit.Name = "lblCreditLimit"
        Me.lblCreditLimit.Size = New System.Drawing.Size(87, 19)
        Me.lblCreditLimit.TabIndex = 25
        Me.lblCreditLimit.TextWrap = False
        '
        'PanelSearchItem
        '
        Me.PanelSearchItem.Controls.Add(Me.btnSerach)
        Me.PanelSearchItem.Controls.Add(Me.MyLabel18)
        Me.PanelSearchItem.Controls.Add(Me.txtItemSearch)
        Me.PanelSearchItem.Location = New System.Drawing.Point(29, 86)
        Me.PanelSearchItem.Name = "PanelSearchItem"
        Me.PanelSearchItem.Size = New System.Drawing.Size(200, 75)
        Me.PanelSearchItem.TabIndex = 1500
        Me.PanelSearchItem.Visible = False
        '
        'btnSerach
        '
        Me.btnSerach.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSerach.Location = New System.Drawing.Point(96, 41)
        Me.btnSerach.Name = "btnSerach"
        Me.btnSerach.Size = New System.Drawing.Size(97, 18)
        Me.btnSerach.TabIndex = 1447
        Me.btnSerach.Text = "Item Search"
        '
        'MyLabel18
        '
        Me.MyLabel18.FieldName = Nothing
        Me.MyLabel18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel18.Location = New System.Drawing.Point(6, 19)
        Me.MyLabel18.Name = "MyLabel18"
        Me.MyLabel18.Size = New System.Drawing.Size(28, 16)
        Me.MyLabel18.TabIndex = 1446
        Me.MyLabel18.Text = "Item"
        '
        'txtItemSearch
        '
        Me.txtItemSearch.CalculationExpression = Nothing
        Me.txtItemSearch.FieldCode = Nothing
        Me.txtItemSearch.FieldDesc = Nothing
        Me.txtItemSearch.FieldMaxLength = 0
        Me.txtItemSearch.FieldName = Nothing
        Me.txtItemSearch.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtItemSearch.isCalculatedField = False
        Me.txtItemSearch.IsSourceFromTable = False
        Me.txtItemSearch.IsSourceFromValueList = False
        Me.txtItemSearch.IsUnique = False
        Me.txtItemSearch.Location = New System.Drawing.Point(40, 18)
        Me.txtItemSearch.MaxLength = 200
        Me.txtItemSearch.MendatroryField = False
        Me.txtItemSearch.MyLinkLable1 = Nothing
        Me.txtItemSearch.MyLinkLable2 = Nothing
        Me.txtItemSearch.Name = "txtItemSearch"
        Me.txtItemSearch.ReferenceFieldDesc = Nothing
        Me.txtItemSearch.ReferenceFieldName = Nothing
        Me.txtItemSearch.ReferenceTableName = Nothing
        Me.txtItemSearch.Size = New System.Drawing.Size(153, 18)
        Me.txtItemSearch.TabIndex = 1445
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.lbltotalOutstanding1)
        Me.Panel4.Controls.Add(Me.lblTotalSecurity11)
        Me.Panel4.Controls.Add(Me.MyLabel16)
        Me.Panel4.Controls.Add(Me.MyLabel14)
        Me.Panel4.Location = New System.Drawing.Point(24, 11)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(200, 60)
        Me.Panel4.TabIndex = 1496
        '
        'lbltotalOutstanding1
        '
        Me.lbltotalOutstanding1.AutoSize = False
        Me.lbltotalOutstanding1.BorderVisible = True
        Me.lbltotalOutstanding1.FieldName = Nothing
        Me.lbltotalOutstanding1.Location = New System.Drawing.Point(101, 25)
        Me.lbltotalOutstanding1.Name = "lbltotalOutstanding1"
        Me.lbltotalOutstanding1.Size = New System.Drawing.Size(94, 19)
        Me.lbltotalOutstanding1.TabIndex = 74
        Me.lbltotalOutstanding1.TextWrap = False
        '
        'lblTotalSecurity11
        '
        Me.lblTotalSecurity11.AutoSize = False
        Me.lblTotalSecurity11.BorderVisible = True
        Me.lblTotalSecurity11.FieldName = Nothing
        Me.lblTotalSecurity11.Location = New System.Drawing.Point(60, 5)
        Me.lblTotalSecurity11.Name = "lblTotalSecurity11"
        Me.lblTotalSecurity11.Size = New System.Drawing.Size(135, 19)
        Me.lblTotalSecurity11.TabIndex = 73
        Me.lblTotalSecurity11.TextWrap = False
        '
        'MyLabel16
        '
        Me.MyLabel16.FieldName = Nothing
        Me.MyLabel16.Location = New System.Drawing.Point(6, 6)
        Me.MyLabel16.Name = "MyLabel16"
        Me.MyLabel16.Size = New System.Drawing.Size(46, 18)
        Me.MyLabel16.TabIndex = 72
        Me.MyLabel16.Text = "Security"
        '
        'MyLabel14
        '
        Me.MyLabel14.FieldName = Nothing
        Me.MyLabel14.Location = New System.Drawing.Point(3, 26)
        Me.MyLabel14.Name = "MyLabel14"
        Me.MyLabel14.Size = New System.Drawing.Size(96, 18)
        Me.MyLabel14.TabIndex = 71
        Me.MyLabel14.Text = "Total Outstanding"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.txtBOstatus)
        Me.Panel1.Controls.Add(Me.txtDOStatus)
        Me.Panel1.Controls.Add(Me.MyLabel7)
        Me.Panel1.Controls.Add(Me.lblLoginUserZone)
        Me.Panel1.Controls.Add(Me.MyLabel13)
        Me.Panel1.Controls.Add(Me.MyLabel9)
        Me.Panel1.Controls.Add(Me.lblroute)
        Me.Panel1.Controls.Add(Me.lblroutename)
        Me.Panel1.Controls.Add(Me.lblroutecode)
        Me.Panel1.Controls.Add(Me.lblPriceCode)
        Me.Panel1.Controls.Add(Me.txtPriceCode)
        Me.Panel1.Controls.Add(Me.MyLabel29)
        Me.Panel1.Controls.Add(Me.MyLabel27)
        Me.Panel1.Controls.Add(Me.lblUploadingDate)
        Me.Panel1.Controls.Add(Me.lblCreatedDateAndTime)
        Me.Panel1.Controls.Add(Me.lblvehicle)
        Me.Panel1.Controls.Add(Me.MyLabel10)
        Me.Panel1.Controls.Add(Me.LblUpdatedVehicleDesc)
        Me.Panel1.Controls.Add(Me.lblvehiclecode)
        Me.Panel1.Controls.Add(Me.LblUpdatedVehicleCode)
        Me.Panel1.Controls.Add(Me.lblvehicleName)
        Me.Panel1.Controls.Add(Me.lblDO)
        Me.Panel1.Controls.Add(Me.lblDONumber)
        Me.Panel1.Controls.Add(Me.MyLabel8)
        Me.Panel1.Controls.Add(Me.lblCancelStatus)
        Me.Panel1.Location = New System.Drawing.Point(230, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(293, 198)
        Me.Panel1.TabIndex = 1497
        '
        'txtBOstatus
        '
        Me.txtBOstatus.AutoSize = False
        Me.txtBOstatus.BorderVisible = True
        Me.txtBOstatus.FieldName = Nothing
        Me.txtBOstatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBOstatus.Location = New System.Drawing.Point(201, 151)
        Me.txtBOstatus.Name = "txtBOstatus"
        Me.txtBOstatus.Size = New System.Drawing.Size(86, 18)
        Me.txtBOstatus.TabIndex = 1451
        Me.txtBOstatus.TextWrap = False
        '
        'txtDOStatus
        '
        Me.txtDOStatus.AutoSize = False
        Me.txtDOStatus.BorderVisible = True
        Me.txtDOStatus.FieldName = Nothing
        Me.txtDOStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDOStatus.Location = New System.Drawing.Point(59, 152)
        Me.txtDOStatus.Name = "txtDOStatus"
        Me.txtDOStatus.Size = New System.Drawing.Size(82, 18)
        Me.txtDOStatus.TabIndex = 1453
        Me.txtDOStatus.TextWrap = False
        '
        'lblLoginUserZone
        '
        Me.lblLoginUserZone.AutoSize = False
        Me.lblLoginUserZone.BorderVisible = True
        Me.lblLoginUserZone.FieldName = Nothing
        Me.lblLoginUserZone.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLoginUserZone.ForeColor = System.Drawing.Color.Black
        Me.lblLoginUserZone.Location = New System.Drawing.Point(128, 77)
        Me.lblLoginUserZone.Name = "lblLoginUserZone"
        Me.lblLoginUserZone.Size = New System.Drawing.Size(156, 18)
        Me.lblLoginUserZone.TabIndex = 1495
        Me.lblLoginUserZone.TextWrap = False
        '
        'MyLabel13
        '
        Me.MyLabel13.FieldName = Nothing
        Me.MyLabel13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel13.Location = New System.Drawing.Point(3, 78)
        Me.MyLabel13.Name = "MyLabel13"
        Me.MyLabel13.Size = New System.Drawing.Size(103, 16)
        Me.MyLabel13.TabIndex = 1494
        Me.MyLabel13.Text = "Current User Zone "
        '
        'MyLabel9
        '
        Me.MyLabel9.FieldName = Nothing
        Me.MyLabel9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel9.Location = New System.Drawing.Point(5, 153)
        Me.MyLabel9.Name = "MyLabel9"
        Me.MyLabel9.Size = New System.Drawing.Size(58, 16)
        Me.MyLabel9.TabIndex = 1454
        Me.MyLabel9.Text = "DO Status"
        '
        'lblroute
        '
        Me.lblroute.FieldName = Nothing
        Me.lblroute.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblroute.Location = New System.Drawing.Point(3, 3)
        Me.lblroute.Name = "lblroute"
        Me.lblroute.Size = New System.Drawing.Size(36, 16)
        Me.lblroute.TabIndex = 150
        Me.lblroute.Text = "Route"
        '
        'lblroutename
        '
        Me.lblroutename.AutoSize = False
        Me.lblroutename.BorderVisible = True
        Me.lblroutename.FieldName = Nothing
        Me.lblroutename.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblroutename.Location = New System.Drawing.Point(128, 2)
        Me.lblroutename.Name = "lblroutename"
        Me.lblroutename.Size = New System.Drawing.Size(156, 18)
        Me.lblroutename.TabIndex = 147
        Me.lblroutename.TextWrap = False
        '
        'lblroutecode
        '
        Me.lblroutecode.AutoSize = False
        Me.lblroutecode.BorderVisible = True
        Me.lblroutecode.FieldName = Nothing
        Me.lblroutecode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblroutecode.Location = New System.Drawing.Point(63, 2)
        Me.lblroutecode.Name = "lblroutecode"
        Me.lblroutecode.Size = New System.Drawing.Size(63, 18)
        Me.lblroutecode.TabIndex = 149
        Me.lblroutecode.TextWrap = False
        '
        'lblPriceCode
        '
        Me.lblPriceCode.FieldName = Nothing
        Me.lblPriceCode.Location = New System.Drawing.Point(3, 21)
        Me.lblPriceCode.Name = "lblPriceCode"
        Me.lblPriceCode.Size = New System.Drawing.Size(60, 18)
        Me.lblPriceCode.TabIndex = 155
        Me.lblPriceCode.Text = "Price Code"
        '
        'txtPriceCode
        '
        Me.txtPriceCode.AutoSize = False
        Me.txtPriceCode.BorderVisible = True
        Me.txtPriceCode.FieldName = Nothing
        Me.txtPriceCode.Location = New System.Drawing.Point(63, 21)
        Me.txtPriceCode.Name = "txtPriceCode"
        Me.txtPriceCode.Size = New System.Drawing.Size(221, 19)
        Me.txtPriceCode.TabIndex = 156
        Me.txtPriceCode.TextWrap = False
        '
        'MyLabel29
        '
        Me.MyLabel29.FieldName = Nothing
        Me.MyLabel29.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel29.Location = New System.Drawing.Point(3, 115)
        Me.MyLabel29.Name = "MyLabel29"
        Me.MyLabel29.Size = New System.Drawing.Size(84, 16)
        Me.MyLabel29.TabIndex = 1490
        Me.MyLabel29.Text = "Uploading Date"
        '
        'MyLabel27
        '
        Me.MyLabel27.FieldName = Nothing
        Me.MyLabel27.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel27.Location = New System.Drawing.Point(3, 135)
        Me.MyLabel27.Name = "MyLabel27"
        Me.MyLabel27.Size = New System.Drawing.Size(74, 16)
        Me.MyLabel27.TabIndex = 1482
        Me.MyLabel27.Text = "Created Date"
        '
        'lblUploadingDate
        '
        Me.lblUploadingDate.AutoSize = False
        Me.lblUploadingDate.BorderVisible = True
        Me.lblUploadingDate.FieldName = Nothing
        Me.lblUploadingDate.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUploadingDate.Location = New System.Drawing.Point(90, 115)
        Me.lblUploadingDate.Name = "lblUploadingDate"
        Me.lblUploadingDate.Size = New System.Drawing.Size(148, 18)
        Me.lblUploadingDate.TabIndex = 1489
        Me.lblUploadingDate.TextWrap = False
        '
        'lblCreatedDateAndTime
        '
        Me.lblCreatedDateAndTime.AutoSize = False
        Me.lblCreatedDateAndTime.BorderVisible = True
        Me.lblCreatedDateAndTime.FieldName = Nothing
        Me.lblCreatedDateAndTime.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCreatedDateAndTime.Location = New System.Drawing.Point(90, 134)
        Me.lblCreatedDateAndTime.Name = "lblCreatedDateAndTime"
        Me.lblCreatedDateAndTime.Size = New System.Drawing.Size(148, 18)
        Me.lblCreatedDateAndTime.TabIndex = 1481
        Me.lblCreatedDateAndTime.TextWrap = False
        '
        'lblvehicle
        '
        Me.lblvehicle.FieldName = Nothing
        Me.lblvehicle.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblvehicle.Location = New System.Drawing.Point(3, 41)
        Me.lblvehicle.Name = "lblvehicle"
        Me.lblvehicle.Size = New System.Drawing.Size(43, 16)
        Me.lblvehicle.TabIndex = 151
        Me.lblvehicle.Text = "Vehicle"
        '
        'MyLabel10
        '
        Me.MyLabel10.FieldName = Nothing
        Me.MyLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel10.Location = New System.Drawing.Point(5, 176)
        Me.MyLabel10.Name = "MyLabel10"
        Me.MyLabel10.Size = New System.Drawing.Size(89, 16)
        Me.MyLabel10.TabIndex = 1460
        Me.MyLabel10.Text = "Updated Vehicle"
        '
        'LblUpdatedVehicleDesc
        '
        Me.LblUpdatedVehicleDesc.AutoSize = False
        Me.LblUpdatedVehicleDesc.BorderVisible = True
        Me.LblUpdatedVehicleDesc.FieldName = Nothing
        Me.LblUpdatedVehicleDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblUpdatedVehicleDesc.Location = New System.Drawing.Point(174, 175)
        Me.LblUpdatedVehicleDesc.Name = "LblUpdatedVehicleDesc"
        Me.LblUpdatedVehicleDesc.Size = New System.Drawing.Size(116, 18)
        Me.LblUpdatedVehicleDesc.TabIndex = 1459
        Me.LblUpdatedVehicleDesc.TextWrap = False
        '
        'lblvehiclecode
        '
        Me.lblvehiclecode.AutoSize = False
        Me.lblvehiclecode.BorderVisible = True
        Me.lblvehiclecode.FieldName = Nothing
        Me.lblvehiclecode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblvehiclecode.Location = New System.Drawing.Point(63, 40)
        Me.lblvehiclecode.Name = "lblvehiclecode"
        Me.lblvehiclecode.Size = New System.Drawing.Size(63, 18)
        Me.lblvehiclecode.TabIndex = 146
        Me.lblvehiclecode.TextWrap = False
        '
        'LblUpdatedVehicleCode
        '
        Me.LblUpdatedVehicleCode.AutoSize = False
        Me.LblUpdatedVehicleCode.BorderVisible = True
        Me.LblUpdatedVehicleCode.FieldName = Nothing
        Me.LblUpdatedVehicleCode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblUpdatedVehicleCode.Location = New System.Drawing.Point(98, 175)
        Me.LblUpdatedVehicleCode.Name = "LblUpdatedVehicleCode"
        Me.LblUpdatedVehicleCode.Size = New System.Drawing.Size(75, 18)
        Me.LblUpdatedVehicleCode.TabIndex = 1458
        Me.LblUpdatedVehicleCode.TextWrap = False
        '
        'lblvehicleName
        '
        Me.lblvehicleName.AutoSize = False
        Me.lblvehicleName.BorderVisible = True
        Me.lblvehicleName.FieldName = Nothing
        Me.lblvehicleName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblvehicleName.Location = New System.Drawing.Point(128, 40)
        Me.lblvehicleName.Name = "lblvehicleName"
        Me.lblvehicleName.Size = New System.Drawing.Size(156, 18)
        Me.lblvehicleName.TabIndex = 148
        Me.lblvehicleName.TextWrap = False
        '
        'lblDO
        '
        Me.lblDO.FieldName = Nothing
        Me.lblDO.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDO.Location = New System.Drawing.Point(3, 60)
        Me.lblDO.Name = "lblDO"
        Me.lblDO.Size = New System.Drawing.Size(123, 16)
        Me.lblDO.TabIndex = 158
        Me.lblDO.Text = "Delivery Order Number"
        '
        'lblDONumber
        '
        Me.lblDONumber.AutoSize = False
        Me.lblDONumber.BorderVisible = True
        Me.lblDONumber.FieldName = Nothing
        Me.lblDONumber.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDONumber.Location = New System.Drawing.Point(128, 59)
        Me.lblDONumber.Name = "lblDONumber"
        Me.lblDONumber.Size = New System.Drawing.Size(156, 18)
        Me.lblDONumber.TabIndex = 157
        Me.lblDONumber.TextWrap = False
        '
        'MyLabel8
        '
        Me.MyLabel8.FieldName = Nothing
        Me.MyLabel8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MyLabel8.Location = New System.Drawing.Point(3, 98)
        Me.MyLabel8.Name = "MyLabel8"
        Me.MyLabel8.Size = New System.Drawing.Size(77, 16)
        Me.MyLabel8.TabIndex = 1457
        Me.MyLabel8.Text = "Cancel Status"
        '
        'lblCancelStatus
        '
        Me.lblCancelStatus.AutoSize = False
        Me.lblCancelStatus.BorderVisible = True
        Me.lblCancelStatus.FieldName = Nothing
        Me.lblCancelStatus.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCancelStatus.Location = New System.Drawing.Point(90, 95)
        Me.lblCancelStatus.Name = "lblCancelStatus"
        Me.lblCancelStatus.Size = New System.Drawing.Size(148, 18)
        Me.lblCancelStatus.TabIndex = 1456
        Me.lblCancelStatus.TextWrap = False
        '
        'RadPageViewPage5
        '
        Me.RadPageViewPage5.Controls.Add(Me.UcAttachment1)
        Me.RadPageViewPage5.ItemSize = New System.Drawing.SizeF(73.0!, 26.0!)
        Me.RadPageViewPage5.Location = New System.Drawing.Point(10, 35)
        Me.RadPageViewPage5.Name = "RadPageViewPage5"
        Me.RadPageViewPage5.Size = New System.Drawing.Size(1260, 409)
        Me.RadPageViewPage5.Text = "Attachment"
        '
        'UcAttachment1
        '
        Me.UcAttachment1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UcAttachment1.Location = New System.Drawing.Point(0, 0)
        Me.UcAttachment1.Name = "UcAttachment1"
        Me.UcAttachment1.Size = New System.Drawing.Size(1260, 409)
        Me.UcAttachment1.TabIndex = 1
        Me.UcAttachment1.TabStop = False
        '
        'btnCreateEWB
        '
        Me.btnCreateEWB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCreateEWB.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateEWB.Location = New System.Drawing.Point(191, 39)
        Me.btnCreateEWB.Name = "btnCreateEWB"
        Me.btnCreateEWB.Size = New System.Drawing.Size(100, 22)
        Me.btnCreateEWB.TabIndex = 1413
        Me.btnCreateEWB.Text = "Create E-Way Bill"
        '
        'rpOpening
        '
        Me.rpOpening.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.rpOpening.Controls.Add(Me.lblDocAmt)
        Me.rpOpening.Controls.Add(Me.lblTotalDocAmt)
        Me.rpOpening.Controls.Add(Me.lblCrAmt)
        Me.rpOpening.Controls.Add(Me.txtCrAmt)
        Me.rpOpening.Controls.Add(Me.lblNetbalance)
        Me.rpOpening.Controls.Add(Me.txtClosingBal)
        Me.rpOpening.Controls.Add(Me.lblOpeningBal)
        Me.rpOpening.Controls.Add(Me.txtOpeningbal)
        Me.rpOpening.Controls.Add(Me.lblDrAmt)
        Me.rpOpening.Controls.Add(Me.txtDrAmt)
        Me.rpOpening.Location = New System.Drawing.Point(28, 2)
        Me.rpOpening.Name = "rpOpening"
        Me.rpOpening.Size = New System.Drawing.Size(1167, 32)
        Me.rpOpening.TabIndex = 1565
        '
        'lblDocAmt
        '
        Me.lblDocAmt.FieldName = Nothing
        Me.lblDocAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDocAmt.Location = New System.Drawing.Point(862, 5)
        Me.lblDocAmt.Name = "lblDocAmt"
        Me.lblDocAmt.Size = New System.Drawing.Size(108, 16)
        Me.lblDocAmt.TabIndex = 126
        Me.lblDocAmt.Text = "Document Amount"
        '
        'lblTotalDocAmt
        '
        Me.lblTotalDocAmt.AutoSize = False
        Me.lblTotalDocAmt.BorderVisible = True
        Me.lblTotalDocAmt.FieldName = Nothing
        Me.lblTotalDocAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalDocAmt.Location = New System.Drawing.Point(977, 3)
        Me.lblTotalDocAmt.Name = "lblTotalDocAmt"
        Me.lblTotalDocAmt.Size = New System.Drawing.Size(104, 20)
        Me.lblTotalDocAmt.TabIndex = 136
        Me.lblTotalDocAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCrAmt
        '
        Me.lblCrAmt.FieldName = Nothing
        Me.lblCrAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCrAmt.Location = New System.Drawing.Point(444, 5)
        Me.lblCrAmt.Name = "lblCrAmt"
        Me.lblCrAmt.Size = New System.Drawing.Size(64, 16)
        Me.lblCrAmt.TabIndex = 141
        Me.lblCrAmt.Text = "Credit Amt"
        '
        'txtCrAmt
        '
        Me.txtCrAmt.AutoSize = False
        Me.txtCrAmt.BorderVisible = True
        Me.txtCrAmt.FieldName = Nothing
        Me.txtCrAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCrAmt.Location = New System.Drawing.Point(512, 3)
        Me.txtCrAmt.Name = "txtCrAmt"
        Me.txtCrAmt.Size = New System.Drawing.Size(123, 20)
        Me.txtCrAmt.TabIndex = 142
        Me.txtCrAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNetbalance
        '
        Me.lblNetbalance.FieldName = Nothing
        Me.lblNetbalance.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNetbalance.Location = New System.Drawing.Point(636, 3)
        Me.lblNetbalance.Name = "lblNetbalance"
        Me.lblNetbalance.Size = New System.Drawing.Size(94, 16)
        Me.lblNetbalance.TabIndex = 139
        Me.lblNetbalance.Text = "Closing Balance"
        '
        'txtClosingBal
        '
        Me.txtClosingBal.AutoSize = False
        Me.txtClosingBal.BorderVisible = True
        Me.txtClosingBal.FieldName = Nothing
        Me.txtClosingBal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtClosingBal.Location = New System.Drawing.Point(733, 3)
        Me.txtClosingBal.Name = "txtClosingBal"
        Me.txtClosingBal.Size = New System.Drawing.Size(118, 20)
        Me.txtClosingBal.TabIndex = 140
        Me.txtClosingBal.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOpeningBal
        '
        Me.lblOpeningBal.FieldName = Nothing
        Me.lblOpeningBal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOpeningBal.Location = New System.Drawing.Point(5, 5)
        Me.lblOpeningBal.Name = "lblOpeningBal"
        Me.lblOpeningBal.Size = New System.Drawing.Size(99, 16)
        Me.lblOpeningBal.TabIndex = 137
        Me.lblOpeningBal.Text = "Opening Balance"
        '
        'txtOpeningbal
        '
        Me.txtOpeningbal.AutoSize = False
        Me.txtOpeningbal.BorderVisible = True
        Me.txtOpeningbal.FieldName = Nothing
        Me.txtOpeningbal.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOpeningbal.Location = New System.Drawing.Point(105, 3)
        Me.txtOpeningbal.Name = "txtOpeningbal"
        Me.txtOpeningbal.Size = New System.Drawing.Size(132, 20)
        Me.txtOpeningbal.TabIndex = 138
        Me.txtOpeningbal.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblDrAmt
        '
        Me.lblDrAmt.FieldName = Nothing
        Me.lblDrAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDrAmt.Location = New System.Drawing.Point(239, 5)
        Me.lblDrAmt.Name = "lblDrAmt"
        Me.lblDrAmt.Size = New System.Drawing.Size(60, 16)
        Me.lblDrAmt.TabIndex = 126
        Me.lblDrAmt.Text = "Debit Amt"
        '
        'txtDrAmt
        '
        Me.txtDrAmt.AutoSize = False
        Me.txtDrAmt.BorderVisible = True
        Me.txtDrAmt.FieldName = Nothing
        Me.txtDrAmt.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDrAmt.Location = New System.Drawing.Point(313, 3)
        Me.txtDrAmt.Name = "txtDrAmt"
        Me.txtDrAmt.Size = New System.Drawing.Size(119, 20)
        Me.txtDrAmt.TabIndex = 136
        Me.txtDrAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnprinte_wayBill
        '
        Me.btnprinte_wayBill.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnprinte_wayBill.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnprinte_wayBill.Location = New System.Drawing.Point(442, 40)
        Me.btnprinte_wayBill.Name = "btnprinte_wayBill"
        Me.btnprinte_wayBill.Size = New System.Drawing.Size(97, 21)
        Me.btnprinte_wayBill.TabIndex = 1528
        Me.btnprinte_wayBill.Text = "Print E-Way Bill"
        '
        'RadLabel10
        '
        Me.RadLabel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RadLabel10.FieldName = Nothing
        Me.RadLabel10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadLabel10.ForeColor = System.Drawing.SystemColors.Highlight
        Me.RadLabel10.Location = New System.Drawing.Point(1197, 7)
        Me.RadLabel10.Name = "RadLabel10"
        Me.RadLabel10.Size = New System.Drawing.Size(84, 16)
        Me.RadLabel10.TabIndex = 1485
        Me.RadLabel10.Text = "F5- Batch Item"
        Me.RadLabel10.Visible = False
        '
        'btnPrintChallan
        '
        Me.btnPrintChallan.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPrintChallan.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrintChallan.Location = New System.Drawing.Point(360, 40)
        Me.btnPrintChallan.Name = "btnPrintChallan"
        Me.btnPrintChallan.Size = New System.Drawing.Size(81, 22)
        Me.btnPrintChallan.TabIndex = 1504
        Me.btnPrintChallan.Text = "Print Challan"
        '
        'btnGatepass
        '
        Me.btnGatepass.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGatepass.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGatepass.Location = New System.Drawing.Point(292, 40)
        Me.btnGatepass.Name = "btnGatepass"
        Me.btnGatepass.Size = New System.Drawing.Size(67, 22)
        Me.btnGatepass.TabIndex = 4
        Me.btnGatepass.Text = "Gate Pass"
        '
        'lblCreatedByValue
        '
        Me.lblCreatedByValue.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCreatedByValue.FieldName = Nothing
        Me.lblCreatedByValue.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCreatedByValue.Location = New System.Drawing.Point(882, 39)
        Me.lblCreatedByValue.Name = "lblCreatedByValue"
        Me.lblCreatedByValue.Size = New System.Drawing.Size(2, 2)
        Me.lblCreatedByValue.TabIndex = 1501
        '
        'lblCreatedBy
        '
        Me.lblCreatedBy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblCreatedBy.FieldName = Nothing
        Me.lblCreatedBy.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCreatedBy.Location = New System.Drawing.Point(761, 41)
        Me.lblCreatedBy.Name = "lblCreatedBy"
        Me.lblCreatedBy.Size = New System.Drawing.Size(63, 16)
        Me.lblCreatedBy.TabIndex = 1500
        Me.lblCreatedBy.Text = "Created By"
        Me.lblCreatedBy.Visible = False
        '
        'btnCreateAndPrintInvoice
        '
        Me.btnCreateAndPrintInvoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCreateAndPrintInvoice.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreateAndPrintInvoice.Location = New System.Drawing.Point(540, 39)
        Me.btnCreateAndPrintInvoice.Name = "btnCreateAndPrintInvoice"
        Me.btnCreateAndPrintInvoice.Size = New System.Drawing.Size(122, 22)
        Me.btnCreateAndPrintInvoice.TabIndex = 22
        Me.btnCreateAndPrintInvoice.Text = "Create && Print Invoice"
        '
        'btnreverse
        '
        Me.btnreverse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnreverse.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnreverse.Location = New System.Drawing.Point(1059, 39)
        Me.btnreverse.Name = "btnreverse"
        Me.btnreverse.Size = New System.Drawing.Size(136, 22)
        Me.btnreverse.TabIndex = 21
        Me.btnreverse.Text = "Reverse/Recreate of BO"
        Me.btnreverse.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(1009, 39)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(46, 22)
        Me.btnCancel.TabIndex = 20
        Me.btnCancel.Text = "Cancel"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(128, 39)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(61, 22)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "Delete"
        '
        'btnPost
        '
        Me.btnPost.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnPost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPost.Location = New System.Drawing.Point(64, 39)
        Me.btnPost.Name = "btnPost"
        Me.btnPost.Size = New System.Drawing.Size(60, 22)
        Me.btnPost.TabIndex = 1
        Me.btnPost.Text = "Post"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(1201, 38)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(71, 22)
        Me.btnClose.TabIndex = 10
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(5, 39)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(57, 22)
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.RadMenu1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1281, 25)
        Me.Panel2.TabIndex = 5
        '
        'RadMenu1
        '
        Me.RadMenu1.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem2})
        Me.RadMenu1.Location = New System.Drawing.Point(0, 0)
        Me.RadMenu1.Name = "RadMenu1"
        Me.RadMenu1.Size = New System.Drawing.Size(1281, 20)
        Me.RadMenu1.TabIndex = 0
        '
        'RadMenuItem2
        '
        Me.RadMenuItem2.Items.AddRange(New Telerik.WinControls.RadItem() {Me.RadMenuItem3, Me.RadMenuItem1, Me.RadMenuItem4, Me.RadMenuItem5})
        Me.RadMenuItem2.Name = "RadMenuItem2"
        Me.RadMenuItem2.Text = "Setting"
        '
        'RadMenuItem3
        '
        Me.RadMenuItem3.Name = "RadMenuItem3"
        Me.RadMenuItem3.Text = "E-Mail/SMS Setting"
        '
        'RadMenuItem1
        '
        Me.RadMenuItem1.Name = "RadMenuItem1"
        Me.RadMenuItem1.Text = "Save Layout"
        '
        'RadMenuItem4
        '
        Me.RadMenuItem4.AccessibleDescription = "RadMenuItem4"
        Me.RadMenuItem4.AccessibleName = "RadMenuItem4"
        Me.RadMenuItem4.Name = "RadMenuItem4"
        Me.RadMenuItem4.Text = "Delete Layout"
        '
        'RadMenuItem5
        '
        Me.RadMenuItem5.AccessibleDescription = "RadMenuItem5"
        Me.RadMenuItem5.AccessibleName = "RadMenuItem5"
        Me.RadMenuItem5.Name = "RadMenuItem5"
        Me.RadMenuItem5.Text = "Footer Setting"
        '
        'frmDairyBookingCustomer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1281, 548)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.Panel2)
        Me.MinimumSize = New System.Drawing.Size(890, 467)
        Me.Name = "frmDairyBookingCustomer"
        '
        '
        '
        Me.RootElement.ApplyShapeToControl = True
        Me.Text = "Dairy Booking Customer"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.RadPageView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageView1.ResumeLayout(False)
        Me.RadPageViewPage1.ResumeLayout(False)
        Me.RadPageViewPage1.PerformLayout()
        CType(Me.txtMobileNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCategory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRouteNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAdharNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReceiptAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReceiptAmtDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReceipt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkNoTranspoter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkAPS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtIndentDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblIndentDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkExcludeKKFMND, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtIndentNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSupplyDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSupplyDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblIndentNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkIsEwayBill, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkNoCrateIssue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkisTCS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkTPT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblShipToLocation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkManualVehicle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTransporter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVehicleNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtManualVehicle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel20, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDescription, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReceiverName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtReceiverName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblChequeNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtChequeNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbPaymentType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPaymentType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkcashsale, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkGhee, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rgbTaxNonTax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.rgbTaxNonTax.ResumeLayout(False)
        Me.rgbTaxNonTax.PerformLayout()
        CType(Me.lblFATPER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtMBRTHours, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtFATPER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblMBRT, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSNFPER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSNFPER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTEMP, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAcidity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtAcidity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSubLocation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel60, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkDistributor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbGatePassType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblGatePassType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblShiftType, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCouponDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCouponDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBPLRemark, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBPLName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCouponCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblBPLRemark, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblBPLName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCouponCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkBPL, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rgbItemType, System.ComponentModel.ISupportInitialize).EndInit()
        Me.rgbItemType.ResumeLayout(False)
        Me.rgbItemType.PerformLayout()
        CType(Me.rbtnNonTax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnTaxable, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtVehicleName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRouteName1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRouteCode1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbcashcredit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCredit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblOutstandingDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblOutStanding, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRoute1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPriceCode1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPriceCodeDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVehicle1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkSampling, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLocation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRouteDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAbandonmentNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox2.ResumeLayout(False)
        CType(Me.gv1.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblVendorName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnAddNew, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage6.ResumeLayout(False)
        Me.RadPageViewPage6.PerformLayout()
        CType(Me.MyLabel19, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox1.ResumeLayout(False)
        Me.RadGroupBox1.PerformLayout()
        CType(Me.RadLabel16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTermName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDueDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv2.MasterTemplate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gv2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel44, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTaxGroupScheme, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.rbtnTaxCalManual, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtnTaxCalAutomatic, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTaxGrpName, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage4.ResumeLayout(False)
        Me.RadPageViewPage4.PerformLayout()
        CType(Me.MyLabel71, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblGrossAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTPTComAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTPTComAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRoundOff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnUpdateVehicle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.btn_QtyReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnHistory, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BtnRecieptEntry, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTCAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTCAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtSecurity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSecuirty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDCAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDCAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TxtRoundoff, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTCSTaxRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel57, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblActualTCSTaxBaseAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel58, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txttcstaxbaseamount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblInvoiceDiscAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadGroupBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadGroupBox3.ResumeLayout(False)
        Me.RadGroupBox3.PerformLayout()
        CType(Me.chkDiscountOnAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkDiscountOnRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDiscAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDiscPer, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCurrConv.ResumeLayout(False)
        Me.pnlCurrConv.PerformLayout()
        CType(Me.txtConversionRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblEffectiveFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtApplicableFrom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCurrency, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblConvRate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAddCharges1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel14, System.ComponentModel.ISupportInitialize).EndInit()
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
        Me.RadPageViewPage3.ResumeLayout(False)
        Me.RadPageViewPage3.PerformLayout()
        CType(Me.lblLastCollectionDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtLastCollectionDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotRAmt1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDCSDemand, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDCSDemandNo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnbilledMilk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUnbilledMilkAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pnlTCS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTCS.ResumeLayout(False)
        Me.pnlTCS.PerformLayout()
        CType(Me.lblTCSBaseAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTCSBaseAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTCSAmount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPanel2.ResumeLayout(False)
        Me.RadPanel2.PerformLayout()
        CType(Me.LblBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCrate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCrate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCan, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCan, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkDCS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGatePassPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrint, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCreateDO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSalesmandesc1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblsalesman1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblSalesman, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblBoothStation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtEx_Factory_Date, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lbl_ExFactoryDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPONo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCustPODate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chkGatePass, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ItemTypePanel.ResumeLayout(False)
        Me.ItemTypePanel.PerformLayout()
        CType(Me.rbtn_Ambient, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rbtn_Fresh, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel40, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.MyLabel25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblARSecurity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel24, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReverseAdvanceSec, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblReverseRefund, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel35, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotalOutstansing, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel33, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblRefund, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel30, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLedgerOutstanding, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel28, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblShortcloseDO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPendingDO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblAdvanceSecurity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCreditLimit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelSearchItem.ResumeLayout(False)
        Me.PanelSearchItem.PerformLayout()
        CType(Me.btnSerach, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtItemSearch, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        CType(Me.lbltotalOutstanding1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotalSecurity11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel14, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.txtBOstatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDOStatus, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblLoginUserZone, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblroute, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblroutename, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblroutecode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblPriceCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPriceCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel29, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblUploadingDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCreatedDateAndTime, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblvehicle, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LblUpdatedVehicleDesc, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblvehiclecode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LblUpdatedVehicleCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblvehicleName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDONumber, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MyLabel8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCancelStatus, System.ComponentModel.ISupportInitialize).EndInit()
        Me.RadPageViewPage5.ResumeLayout(False)
        CType(Me.btnCreateEWB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rpOpening, System.ComponentModel.ISupportInitialize).EndInit()
        Me.rpOpening.ResumeLayout(False)
        Me.rpOpening.PerformLayout()
        CType(Me.lblDocAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblTotalDocAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCrAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtCrAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblNetbalance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtClosingBal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblOpeningBal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtOpeningbal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblDrAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDrAmt, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnprinte_wayBill, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RadLabel10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPrintChallan, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnGatepass, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCreatedByValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.lblCreatedBy, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCreateAndPrintInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnreverse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnDelete, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnPost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.btnSave, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.RadMenu1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents RadPageView1 As Telerik.WinControls.UI.RadPageView
    Friend WithEvents RadPageViewPage4 As Telerik.WinControls.UI.RadPageViewPage
    Friend WithEvents btnClose As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnSave As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnPost As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnDelete As Telerik.WinControls.UI.RadButton
    Friend WithEvents txtComment As System.Windows.Forms.TextBox
    Friend WithEvents RadLabel25 As common.Controls.MyLabel
    Friend WithEvents lblTaxAmt As common.Controls.MyLabel
    Friend WithEvents lblAmtAfterDiscount As common.Controls.MyLabel
    Friend WithEvents lblDiscountAmt As common.Controls.MyLabel
    Friend WithEvents lblAmtWithDiscount As common.Controls.MyLabel
    Friend WithEvents RadLabel22 As common.Controls.MyLabel
    Friend WithEvents RadLabel19 As common.Controls.MyLabel
    Friend WithEvents RadLabel27 As common.Controls.MyLabel
    Friend WithEvents lblTotRAmt As common.Controls.MyLabel
    Friend WithEvents RadLabel9 As common.Controls.MyLabel
    Friend WithEvents RadLabel14 As common.Controls.MyLabel
    Friend WithEvents RadLabel32 As common.Controls.MyLabel
    Friend WithEvents lblAddCharges1 As common.Controls.MyLabel
    Friend WithEvents txtApplicableFrom As common.Controls.MyLabel
    Friend WithEvents lblEffectiveFrom As common.Controls.MyLabel
    Friend WithEvents lblConvRate As common.Controls.MyLabel
    Friend WithEvents lblCurrency As common.Controls.MyLabel
    Friend WithEvents pnlCurrConv As System.Windows.Forms.Panel
    Friend WithEvents txtCurrencyCode As common.UserControls.txtFinder
    Friend WithEvents txtConversionRate As common.MyNumBox
    Friend WithEvents MyLabel3 As common.Controls.MyLabel
    Friend WithEvents RadGroupBox3 As Telerik.WinControls.UI.RadGroupBox
    Friend WithEvents chkDiscountOnAmt As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents chkDiscountOnRate As Telerik.WinControls.UI.RadRadioButton
    Friend WithEvents txtDiscAmt As common.MyNumBox
    Friend WithEvents txtDiscPer As common.MyNumBox
    Friend WithEvents MyLabel1 As common.Controls.MyLabel
    Friend WithEvents lblInvoiceDiscAmt As common.Controls.MyLabel
    Friend WithEvents MyLabel6 As common.Controls.MyLabel
    Friend WithEvents RadMenuItem2 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem3 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenu1 As Telerik.WinControls.UI.RadMenu
    Friend WithEvents RadMenuItem1 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem4 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents RadMenuItem5 As Telerik.WinControls.UI.RadMenuItem
    Friend WithEvents btnCancel As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnreverse As Telerik.WinControls.UI.RadButton
    Friend WithEvents btnCreateAndPrintInvoice As Telerik.WinControls.UI.RadButton
    Friend WithEvents RadLabel10 As common.Controls.MyLabel
    Friend WithEvents RadPageViewPage2 As RadPageViewPage
    Friend WithEvents btn_QtyReset As RadButton
    Friend WithEvents lblCreatedBy As common.Controls.MyLabel
    Friend WithEvents lblCreatedByValue As common.Controls.MyLabel
    Friend WithEvents RadPageViewPage3 As RadPageViewPage
    Friend WithEvents ItemTypePanel As Panel
    Friend WithEvents rbtn_Ambient As RadRadioButton
    Friend WithEvents rbtn_Fresh As RadRadioButton
    Friend WithEvents MyLabel40 As common.Controls.MyLabel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents MyLabel25 As common.Controls.MyLabel
    Friend WithEvents lblARSecurity As common.Controls.MyLabel
    Friend WithEvents MyLabel24 As common.Controls.MyLabel
    Friend WithEvents lblReverseAdvanceSec As common.Controls.MyLabel
    Friend WithEvents MyLabel23 As common.Controls.MyLabel
    Friend WithEvents lblReverseRefund As common.Controls.MyLabel
    Friend WithEvents MyLabel35 As common.Controls.MyLabel
    Friend WithEvents lblTotalOutstansing As common.Controls.MyLabel
    Friend WithEvents MyLabel33 As common.Controls.MyLabel
    Friend WithEvents lblRefund As common.Controls.MyLabel
    Friend WithEvents MyLabel30 As common.Controls.MyLabel
    Friend WithEvents lblLedgerOutstanding As common.Controls.MyLabel
    Friend WithEvents MyLabel28 As common.Controls.MyLabel
    Friend WithEvents lblShortcloseDO As common.Controls.MyLabel
    Friend WithEvents MyLabel26 As common.Controls.MyLabel
    Friend WithEvents lblPendingDO As common.Controls.MyLabel
    Friend WithEvents MyLabel22 As common.Controls.MyLabel
    Friend WithEvents lblAdvanceSecurity As common.Controls.MyLabel
    Friend WithEvents MyLabel15 As common.Controls.MyLabel
    Friend WithEvents lblCreditLimit As common.Controls.MyLabel
    Friend WithEvents PanelSearchItem As Panel
    Friend WithEvents btnSerach As RadButton
    Friend WithEvents MyLabel18 As common.Controls.MyLabel
    Friend WithEvents txtItemSearch As common.Controls.MyTextBox
    Friend WithEvents Panel4 As Panel
    Friend WithEvents lbltotalOutstanding1 As common.Controls.MyLabel
    Friend WithEvents lblTotalSecurity11 As common.Controls.MyLabel
    Friend WithEvents MyLabel16 As common.Controls.MyLabel
    Friend WithEvents MyLabel14 As common.Controls.MyLabel
    Friend WithEvents txtBOstatus As common.Controls.MyLabel
    Friend WithEvents txtDOStatus As common.Controls.MyLabel
    Friend WithEvents MyLabel7 As common.Controls.MyLabel
    Friend WithEvents lblLoginUserZone As common.Controls.MyLabel
    Friend WithEvents MyLabel13 As common.Controls.MyLabel
    Friend WithEvents MyLabel9 As common.Controls.MyLabel
    Friend WithEvents lblroute As common.Controls.MyLabel
    Friend WithEvents lblroutename As common.Controls.MyLabel
    Friend WithEvents lblroutecode As common.Controls.MyLabel
    Friend WithEvents lblPriceCode As common.Controls.MyLabel
    Friend WithEvents txtPriceCode As common.Controls.MyLabel
    Friend WithEvents MyLabel29 As common.Controls.MyLabel
    Friend WithEvents MyLabel27 As common.Controls.MyLabel
    Friend WithEvents lblUploadingDate As common.Controls.MyLabel
    Friend WithEvents lblCreatedDateAndTime As common.Controls.MyLabel
    Friend WithEvents lblvehicle As common.Controls.MyLabel
    Friend WithEvents MyLabel10 As common.Controls.MyLabel
    Friend WithEvents LblUpdatedVehicleDesc As common.Controls.MyLabel
    Friend WithEvents lblvehiclecode As common.Controls.MyLabel
    Friend WithEvents LblUpdatedVehicleCode As common.Controls.MyLabel
    Friend WithEvents lblvehicleName As common.Controls.MyLabel
    Friend WithEvents lblDO As common.Controls.MyLabel
    Friend WithEvents lblDONumber As common.Controls.MyLabel
    Friend WithEvents MyLabel8 As common.Controls.MyLabel
    Friend WithEvents lblCancelStatus As common.Controls.MyLabel
    Friend WithEvents RadPageViewPage1 As RadPageViewPage
    Friend WithEvents chkSampling As RadCheckBox
    Friend WithEvents lblLocation As common.Controls.MyLabel
    Friend WithEvents txtLocation As common.UserControls.txtFinder
    Friend WithEvents RadLabel15 As common.Controls.MyLabel
    Friend WithEvents lblRouteDesc As common.Controls.MyLabel
    Friend WithEvents lblAbandonmentNo As common.Controls.MyLabel
    Friend WithEvents lblRouteNo As common.Controls.MyLabel
    Friend WithEvents RadGroupBox2 As RadGroupBox
    Friend WithEvents gv1 As common.UserControls.MyRadGridView
    Friend WithEvents RadLabel4 As common.Controls.MyLabel
    Friend WithEvents lblVendorName As common.Controls.MyLabel
    Friend WithEvents txtRouteNo As common.UserControls.txtFinder
    Friend WithEvents RadLabel2 As common.Controls.MyLabel
    Friend WithEvents RadLabel1 As common.Controls.MyLabel
    Friend WithEvents btnAddNew As RadButton
    Friend WithEvents txtVendorNo As common.UserControls.txtFinder
    Friend WithEvents txtDocNo As common.UserControls.txtNavigator
    Friend WithEvents txtEx_Factory_Date As common.Controls.MyDateTimePicker
    Friend WithEvents lbl_ExFactoryDate As common.Controls.MyLabel
    Friend WithEvents MyLabel12 As common.Controls.MyLabel
    Friend WithEvents txtPONo As common.Controls.MyTextBox
    Friend WithEvents MyLabel4 As common.Controls.MyLabel
    Friend WithEvents txtCustPODate As common.Controls.MyDateTimePicker
    Friend WithEvents chkGatePass As RadCheckBox
    Friend WithEvents lblOutstandingDesc As common.Controls.MyLabel
    Friend WithEvents lblOutStanding As common.Controls.MyLabel
    Friend WithEvents lblRoute1 As common.Controls.MyLabel
    Friend WithEvents lblPriceCode1 As common.Controls.MyLabel
    Friend WithEvents lblPriceCodeDesc As common.Controls.MyLabel
    Friend WithEvents lblVehicle1 As common.Controls.MyLabel
    Friend WithEvents cmbcashcredit As common.Controls.MyComboBox
    Friend WithEvents lblCredit As common.Controls.MyLabel
    Friend WithEvents lblBoothStation As common.Controls.MyLabel
    Friend WithEvents txtSalesman As common.UserControls.txtFinder
    Friend WithEvents lblSalesman As common.Controls.MyLabel
    Friend WithEvents MyLabel2 As common.Controls.MyLabel
    Friend WithEvents txtDate As common.Controls.MyDateTimePicker
    Friend WithEvents btnCC As RadButton
    Friend WithEvents txtVehicleName As common.Controls.MyTextBox
    Friend WithEvents txtRouteName1 As common.Controls.MyTextBox
    Friend WithEvents txtRouteCode1 As common.Controls.MyTextBox
    Friend WithEvents txtVehicleCode As common.UserControls.txtFinder
    Friend WithEvents rgbItemType As RadGroupBox
    Friend WithEvents rbtnNonTax As RadRadioButton
    Friend WithEvents rbtnTaxable As RadRadioButton
    Friend WithEvents chkBPL As RadCheckBox
    Friend WithEvents lblBPLRemark As common.Controls.MyLabel
    Friend WithEvents lblBPLName As common.Controls.MyLabel
    Friend WithEvents lblCouponCode As common.Controls.MyLabel
    Friend WithEvents txtBPLRemark As common.Controls.MyTextBox
    Friend WithEvents txtBPLName As common.Controls.MyTextBox
    Friend WithEvents txtCouponCode As common.Controls.MyTextBox
    Friend WithEvents RadPageViewPage5 As RadPageViewPage
    Friend WithEvents UcAttachment1 As ucAttachment
    Friend WithEvents txtCouponDate As common.Controls.MyDateTimePicker
    Friend WithEvents lblCouponDate As common.Controls.MyLabel
    Friend WithEvents lblShiftType As RadLabel
    Friend WithEvents cmbGatePassType As common.Controls.MyComboBox
    Friend WithEvents lblGatePassType As common.Controls.MyLabel
    Friend WithEvents chkDistributor As RadCheckBox
    Friend WithEvents txtCategory As common.UserControls.txtFinder
    Friend WithEvents lblCategory As common.Controls.MyLabel
    Friend WithEvents btnGatepass As RadButton
    Friend WithEvents BtnRecieptEntry As RadButton
    Friend WithEvents RadPageViewPage6 As RadPageViewPage
    Friend WithEvents txtSchemeTaxGroup As common.UserControls.txtFinder
    Friend WithEvents MyLabel44 As common.Controls.MyLabel
    Friend WithEvents lblTaxGroupScheme As common.Controls.MyLabel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rbtnTaxCalManual As common.Controls.MyRadioButton
    Friend WithEvents rbtnTaxCalAutomatic As common.Controls.MyRadioButton
    Friend WithEvents txtTaxGroup As common.UserControls.txtFinder
    Friend WithEvents RadLabel11 As common.Controls.MyLabel
    Friend WithEvents lblTaxGrpName As common.Controls.MyLabel
    Friend WithEvents gv2 As common.UserControls.MyRadGridView
    Friend WithEvents RadGroupBox1 As RadGroupBox
    Friend WithEvents txtTermCode As common.UserControls.txtFinder
    Friend WithEvents RadLabel16 As common.Controls.MyLabel
    Friend WithEvents lblTermName As common.Controls.MyLabel
    Friend WithEvents txtDueDate As common.Controls.MyDateTimePicker
    Friend WithEvents RadLabel17 As common.Controls.MyLabel
    Friend WithEvents MyLabel19 As common.Controls.MyLabel
    Friend WithEvents txtTCSTaxRate As common.MyNumBox
    Friend WithEvents MyLabel57 As common.Controls.MyLabel
    Friend WithEvents lblActualTCSTaxBaseAmt As common.Controls.MyLabel
    Friend WithEvents MyLabel58 As common.Controls.MyLabel
    Friend WithEvents txttcstaxbaseamount As common.MyNumBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents TxtRoundoff As common.Controls.MyLabel
    Friend WithEvents txtSecurity As common.Controls.MyLabel
    Friend WithEvents lblSecuirty As common.Controls.MyLabel
    Friend WithEvents txtDCAmt As common.Controls.MyLabel
    Friend WithEvents lblDCAmt As common.Controls.MyLabel
    Friend WithEvents lblSubLocation As common.Controls.MyLabel
    Friend WithEvents MyLabel60 As common.Controls.MyLabel
    Friend WithEvents txtSubLocation As common.UserControls.txtFinder
    Friend WithEvents btnPrintChallan As RadButton
    Friend WithEvents txtMBRTHours As common.Controls.MyTextBox
    Friend WithEvents lblMBRT As common.Controls.MyLabel
    Friend WithEvents txtTemp As common.Controls.MyTextBox
    Friend WithEvents lblTEMP As common.Controls.MyLabel
    Friend WithEvents txtAcidity As common.Controls.MyTextBox
    Friend WithEvents lblAcidity As common.Controls.MyLabel
    Friend WithEvents txtSNFPER As common.Controls.MyTextBox
    Friend WithEvents lblSNFPER As common.Controls.MyLabel
    Friend WithEvents txtFATPER As common.Controls.MyTextBox
    Friend WithEvents lblFATPER As common.Controls.MyLabel
    Friend WithEvents rgbTaxNonTax As RadGroupBox
    Friend WithEvents chkGhee As RadCheckBox
    Friend WithEvents txtTCAmt As common.Controls.MyLabel
    Friend WithEvents lblTCAmt As common.Controls.MyLabel
    Friend WithEvents chkcashsale As RadCheckBox
    Friend WithEvents cmbPaymentType As common.Controls.MyComboBox
    Friend WithEvents lblPaymentType As common.Controls.MyLabel
    Friend WithEvents lblChequeNo As common.Controls.MyLabel
    Friend WithEvents txtChequeNo As common.Controls.MyTextBox
    Friend WithEvents lblReceiverName As common.Controls.MyLabel
    Friend WithEvents txtReceiverName As common.Controls.MyTextBox
    Friend WithEvents txtSalesman1 As common.UserControls.txtFinder
    Friend WithEvents lblSalesmandesc1 As common.Controls.MyLabel
    Friend WithEvents lblsalesman1 As common.Controls.MyLabel
    Friend WithEvents MyLabel20 As common.Controls.MyLabel
    Friend WithEvents txtDescription As common.Controls.MyTextBox
    Friend WithEvents fndTransporter As common.UserControls.txtFinder
    Friend WithEvents lblVehicleNo As common.Controls.MyLabel
    Friend WithEvents txtManualVehicle As common.Controls.MyTextBox
    Friend WithEvents MyLabel31 As common.Controls.MyLabel
    Friend WithEvents lblTransporter As common.Controls.MyLabel
    Friend WithEvents chkManualVehicle As RadCheckBox
    Friend WithEvents lblShipToLocation As common.Controls.MyLabel
    Friend WithEvents txtShipToLocation As common.UserControls.txtFinder
    Friend WithEvents RadLabel18 As common.Controls.MyLabel
    Friend WithEvents chkTPT As RadCheckBox
    Friend WithEvents chkisTCS As RadCheckBox
    Friend WithEvents btnCreateDO As RadButton
    Friend WithEvents btnprinte_wayBill As RadButton
    Friend WithEvents btnHistory As RadButton
    Friend WithEvents chkNoCrateIssue As RadCheckBox
    Friend WithEvents chkIsEwayBill As RadCheckBox
    Friend WithEvents txtSupplyDate As common.Controls.MyDateTimePicker
    Friend WithEvents lblSupplyDate As common.Controls.MyLabel
    Friend WithEvents chkExcludeKKFMND As RadCheckBox
    Friend WithEvents btnGatePassPrint As RadButton
    Friend WithEvents btnPrint As RadButton
    Friend WithEvents chkDCS As RadCheckBox
    Friend WithEvents UsLock1 As common.usLock
    Friend WithEvents txtIndentDate As common.Controls.MyDateTimePicker
    Friend WithEvents lblIndentDate As common.Controls.MyLabel
    Friend WithEvents txtIndentNo As common.Controls.MyTextBox
    Friend WithEvents lblIndentNo As common.Controls.MyLabel
    Friend WithEvents lblDCSDemand As common.Controls.MyLabel
    Friend WithEvents txtDCSDemandNo As common.Controls.MyLabel
    Friend WithEvents lblUnbilledMilk As common.Controls.MyLabel
    Friend WithEvents lblUnbilledMilkAmt As common.Controls.MyLabel
    Friend WithEvents pnlTCS As RadPanel
    Friend WithEvents lblTCSBaseAmt As common.Controls.MyLabel
    Friend WithEvents txtTCSBaseAmt As common.Controls.MyLabel
    Friend WithEvents MyLabel17 As common.Controls.MyLabel
    Friend WithEvents lblTCSAmount As common.Controls.MyLabel
    Friend WithEvents RadPanel2 As RadPanel
    Friend WithEvents LblBox As common.Controls.MyLabel
    Friend WithEvents txtBox As common.Controls.MyLabel
    Friend WithEvents lblCrate As common.Controls.MyLabel
    Friend WithEvents txtCrate As common.Controls.MyLabel
    Friend WithEvents lblCan As common.Controls.MyLabel
    Friend WithEvents txtCan As common.Controls.MyLabel
    Friend WithEvents MyLabel5 As common.Controls.MyLabel
    Friend WithEvents lblTotRAmt1 As common.Controls.MyLabel
    Friend WithEvents rpOpening As RadPanel
    Friend WithEvents lblOpeningBal As common.Controls.MyLabel
    Friend WithEvents txtOpeningbal As common.Controls.MyLabel
    Friend WithEvents lblDrAmt As common.Controls.MyLabel
    Friend WithEvents txtDrAmt As common.Controls.MyLabel
    Friend WithEvents lblDocAmt As common.Controls.MyLabel
    Friend WithEvents lblTotalDocAmt As common.Controls.MyLabel
    Friend WithEvents lblNetbalance As common.Controls.MyLabel
    Friend WithEvents txtClosingBal As common.Controls.MyLabel
    Friend WithEvents lblLastCollectionDate As common.Controls.MyLabel
    Friend WithEvents txtLastCollectionDate As common.Controls.MyLabel
    Friend WithEvents lblCrAmt As common.Controls.MyLabel
    Friend WithEvents txtCrAmt As common.Controls.MyLabel
    Friend WithEvents chkAPS As RadCheckBox
    Friend WithEvents btnCreateEWB As RadButton
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnUpdateVehicle As RadButton
    Friend WithEvents chkNoTranspoter As RadCheckBox
    Friend WithEvents lblRoundOff As common.Controls.MyLabel
    Friend WithEvents lblReceiptAmt As common.Controls.MyLabel
    Friend WithEvents lblReceiptAmtDesc As common.Controls.MyLabel
    Friend WithEvents lblReceipt As common.Controls.MyLabel
    Friend WithEvents txtReceipt As common.UserControls.txtFinder
    Friend WithEvents MyLabel71 As common.Controls.MyLabel
    Friend WithEvents lblGrossAmount As common.Controls.MyLabel
    Friend WithEvents lblTPTComAmt As common.Controls.MyLabel
    Friend WithEvents txtTPTComAmt As common.Controls.MyLabel
    Friend WithEvents MyLabel21 As common.Controls.MyLabel
    Friend WithEvents MyLabel11 As common.Controls.MyLabel
    Friend WithEvents txtAdharNo As common.Controls.MyTextBox
    Friend WithEvents txtMobileNo As common.Controls.MyTextBox
End Class

