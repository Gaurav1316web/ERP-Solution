Imports System.Data.SqlClient
Imports System.IO
Imports common
Public Class frmMCCMaterialSale
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim ApplySaleRouteInDCSSale As Boolean = False
    Dim DefaultEnableNoTransporter As Boolean = False
    Dim ApplyEWBThresholdLimit As Boolean = False
    Dim EWBThresholdLimitForIntraCity As Integer = 0
    Dim EWBThresholdLimitForIntraState As Integer = 0
    Dim EWBThresholdLimitForInterState As Integer = 0
    Dim FlagDocumentIsTaxable As Integer = 0
    Dim ConvertIntoBillingUOM As Boolean = False
    Dim chkTaxGroup As String
    Dim DefaultEnableEWayBill As Boolean = False
    Dim arrMCCRights As ArrayList
    Dim CalculateTaxRatefromItemwsieTaxOnSale As Integer = 0
    Dim UseDescInsteadOFCodeOnMCCMAterialSale As Boolean = False
    Const colItemwiseTaxCode As String = "colItemwiseTaxCode"
    Dim GSTStatus As Boolean = False
    Dim CreateAutoMCCPriceChat As Boolean = False
    Dim EnableTCSRateValidityFrom01July2021 As Boolean = False
    Dim AmountToCheckCustomerOutstandingForTCSTax As Decimal = 0
    Dim TotalItemQty As Decimal = 0
    Private StrSql As String
    Private AutoScheme As Boolean = False
    Private MultiplySubsidyWithQuantity As Boolean = False
    'Private ExcludeKKFAndMandiForDCS As Boolean = False
    Private attachQry As String = ""
    Private blnBackCalculation As Boolean = False
    Private AllowChangeInvoiceType As Boolean = False
    Private IsBatchMFDEXDmandatory As Boolean = False
    Private PurchaseOneItemOneVendor As Boolean = False
    Private ItemRateEditable As Boolean = False
    Private ItemMRPEditable As Boolean = False
    Dim RunBatchFifowise As Integer = 0
    Dim RunBatchFifowisewithmodifyfunctionality As Boolean = False
    Public checkstockmrpwise As Boolean = False
    Public strExcise As Boolean
    Public intMRPwithabatement As Integer
    Private isPO_GRN_MRN_Editable As Boolean = False
    Public Const RowTypeItem As String = "Item"
    Public Const RowTypeMisc As String = "Misc"
    Const colOrgUnit As String = "COLORGUNIT"
    Const ReportID As String = "MCCMeterialSNItemGrid"
    Public strSRNno As String = Nothing
    Private isCellValueChangedOpen As Boolean = False
    Private isCellValueChangedTaxOpen As Boolean = False
    Dim iStxtTaxGroup_TxtChangedComplete As Boolean = True
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Const colLineNo As String = "COLLNO"
    Const colRowType As String = "COLTYPE"
    Const colComplete As String = "COMPLETE"
    Const colBalanceQty As String = "BALANCEQTY"
    Const colICode As String = "COLICODE"
    Const colShortDesc As String = "colShortDesc"
    Const colIName As String = "COLINAME"
    Const colDCSSaleZeroCost As String = "colDCSSaleZeroCost"
    Const colHSNNo As String = "COLHSNNo"
    Const colBarCode As String = "COLBARCODE"
    Const colPendingQty As String = "COLPENDINGQTY"
    Const colQty As String = "COLQTY"
    Const colBillingQty As String = "colBillingQty"
    Const colFreeQty As String = "COLFREEQTY"
    Const colUnit As String = "COLUNIT"
    Const colBillingUnit As String = "colBillingUnit"
    Const colUOMName As String = "COLUOMNAME"
    Const colRate As String = "COLRATE"
    Const colAmt As String = "COLAMT"
    Const colDisPer As String = "COLDISPER"
    Const colDisPerUnit As String = "COLDISPERUNIT"
    Const colDisUnitAmt As String = "COLDISUNITAMT"
    Const colDisAmt As String = "COLDISAMT"
    Const colAmtAfterDis As String = "COLAMTAFTERDIS"
    Const colTax1 As String = "COLTAX1"
    Const colTaxBaseAmt1 As String = "COLTAXBASEAMT1"
    Const colTaxRate1 As String = "COLTAXRATE1"
    Const colTaxAmt1 As String = "COLTAXAMT1"
    Const colIsTaxable1 As String = "ISTAXABLE1"
    Const colIsSurTax1 As String = "ISSURTAX1"
    Const colSurTaxCode1 As String = "SURTAXCODE1"
    Const colTaxRecoverable1 As String = "RECOVERTABLETAX1"
    Const colTax2 As String = "COLTAX2"
    Const colTaxBaseAmt2 As String = "COLTAXBASEAMT2"
    Const colTaxRate2 As String = "COLTAXRATE2"
    Const colTaxAmt2 As String = "COLTAXAMT2"
    Const colIsTaxable2 As String = "ISTAXABLE2"
    Const colIsSurTax2 As String = "ISSURTAX2"
    Const colSurTaxCode2 As String = "SURTAXCODE2"
    Const colTaxRecoverable2 As String = "RECOVERTABLETAX2"
    Const colTax3 As String = "COLTAX3"
    Const colTaxBaseAmt3 As String = "COLTAXBASEAMT3"
    Const colTaxRate3 As String = "COLTAXRATE3"
    Const colTaxAmt3 As String = "COLTAXAMT3"
    Const colIsTaxable3 As String = "ISTAXABLE3"
    Const colIsSurTax3 As String = "ISSURTAX3"
    Const colSurTaxCode3 As String = "SURTAXCODE3"
    Const colTaxRecoverable3 As String = "RECOVERTABLETAX3"
    Const colTax4 As String = "COLTAX4"
    Const colTaxBaseAmt4 As String = "COLTAXBASEAMT4"
    Const colTaxRate4 As String = "COLTAXRATE4"
    Const colTaxAmt4 As String = "COLTAXAMT4"
    Const colIsTaxable4 As String = "ISTAXABLE4"
    Const colIsSurTax4 As String = "ISSURTAX4"
    Const colSurTaxCode4 As String = "SURTAXCODE4"
    Const colTaxRecoverable4 As String = "RECOVERTABLETAX4"
    Const colTax5 As String = "COLTAX5"
    Const colTaxBaseAmt5 As String = "COLTAXBASEAMT5"
    Const colTaxRate5 As String = "COLTAXRATE5"
    Const colTaxAmt5 As String = "COLTAXAMT5"
    Const colIsTaxable5 As String = "ISTAXABLE5"
    Const colIsSurTax5 As String = "ISSURTAX5"
    Const colSurTaxCode5 As String = "SURTAXCODE5"
    Const colTaxRecoverable5 As String = "RECOVERTABLETAX5"
    Const colTax6 As String = "COLTAX6"
    Const colTaxBaseAmt6 As String = "COLTAXBASEAMT6"
    Const colTaxRate6 As String = "COLTAXRATE6"
    Const colTaxAmt6 As String = "COLTAXAMT6"
    Const colIsTaxable6 As String = "ISTAXABLE6"
    Const colIsSurTax6 As String = "ISSURTAX6"
    Const colSurTaxCode6 As String = "SURTAXCODE6"
    Const colTaxRecoverable6 As String = "RECOVERTABLETAX6"
    Const colTax7 As String = "COLTAX7"
    Const colTaxBaseAmt7 As String = "COLTAXBASEAMT7"
    Const colTaxRate7 As String = "COLTAXRATE7"
    Const colTaxAmt7 As String = "COLTAXAMT7"
    Const colIsTaxable7 As String = "ISTAXABLE7"
    Const colIsSurTax7 As String = "ISSURTAX7"
    Const colSurTaxCode7 As String = "SURTAXCODE7"
    Const colTaxRecoverable7 As String = "RECOVERTABLETAX7"
    Const colTax8 As String = "COLTAX8"
    Const colTaxBaseAmt8 As String = "COLTAXBASEAMT8"
    Const colTaxRate8 As String = "COLTAXRATE8"
    Const colTaxAmt8 As String = "COLTAXAMT8"
    Const colIsTaxable8 As String = "ISTAXABLE8"
    Const colIsSurTax8 As String = "ISSURTAX8"
    Const colSurTaxCode8 As String = "SURTAXCODE8"
    Const colTaxRecoverable8 As String = "RECOVERTABLETAX8"
    Const colTax9 As String = "COLTAX9"
    Const colTaxBaseAmt9 As String = "COLTAXBASEAMT9"
    Const colTaxRate9 As String = "COLTAXRATE9"
    Const colTaxAmt9 As String = "COLTAXAMT9"
    Const colIsTaxable9 As String = "ISTAXABLE9"
    Const colIsSurTax9 As String = "ISSURTAX9"
    Const colSurTaxCode9 As String = "SURTAXCODE9"
    Const colTaxRecoverable9 As String = "RECOVERTABLETAX9"
    Const colTax10 As String = "COLTAX10"
    Const colTaxBaseAmt10 As String = "COLTAXBASEAMT10"
    Const colTaxRate10 As String = "COLTAXRATE10"
    Const colTaxAmt10 As String = "COLTAXAMT10"
    Const colIsTaxable10 As String = "ISTAXABLE10"
    Const colIsTaxOnBaseAmount As String = "colIsTaxOnBaseAmount"
    Const colIsSurTax10 As String = "ISSURTAX10"
    Const colSurTaxCode10 As String = "SURTAXCODE10"
    Const colTaxRecoverable10 As String = "RECOVERTABLETAX10"
    Const colIsExcisable1 As String = "ISEXCISABLE1"
    Const colIsExcisable2 As String = "ISEXCISABLE2"
    Const colIsExcisable3 As String = "ISEXCISABLE3"
    Const colIsExcisable4 As String = "ISEXCISABLE4"
    Const colIsExcisable5 As String = "ISEXCISABLE5"
    Const colIsExcisable6 As String = "ISEXCISABLE6"
    Const colIsExcisable7 As String = "ISEXCISABLE7"
    Const colIsExcisable8 As String = "ISEXCISABLE8"
    Const colIsExcisable9 As String = "ISEXCISABLE9"
    Const colIsExcisable10 As String = "ISEXCISABLE10"
    Const colTotTaxAmt As String = "TAXAMT"
    Const colAmtAfterTax As String = "AMTAFTERTAX"
    Const colOrderNo As String = "ORDERNO"
    Const colHeaDDisPer As String = "colHeaDDisPer"
    Const colHeadDisPerAmt As String = "colHeadDisPerAmt"
    Const colLocationCode As String = "LOCATIONCODE"
    Const colLocationName As String = "LOCATIONNAME"
    Const colMRP As String = "MRP"
    Const colIsBatchItem As String = "colIsBatchItem"
    Const colBatchNo As String = "BATCHNO"
    Const colManualBatchNo As String = "ManualBATCHNO"
    Const colExpiry As String = "EXPIRYDATE"
    Const colManufactureDate As String = "MANUFACTUREDATE"
    Const colLandedRate As String = "LANDEDRATE"
    Const colLandedAmt As String = "LANDEDAMT"
    Const colSpecification As String = "SPECIFICATION"
    Const colRemarks As String = "REMARKS"
    Const colIsMannualAmt As String = "ISMANNUALAMT"
    Const colIsSerialseItem As String = "COLISSERIALSEITEM"
    Const colTTaxAutCode As String = "TAXAUTCODE"
    Const colTTaxAutName As String = "TAXAUTNAME"
    Const colTTaxRate As String = "TAXRATE"
    Const colTBaseAmt As String = "TAXBASEAMT"
    Const colTTaxAmt As String = "TAXAMT"
    Const colItmCost As String = "ItmCost"
    Const colACCode As String = "COLACCODE"
    Const colACName As String = "COLACNAME"
    Const colACAmount As String = "COLACAMOUNT"
    Const colIsEmptyValue As String = "ISEMPTYVALUE"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim repoBalQty As GridViewDecimalColumn
    Dim repoComplete As GridViewTextBoxColumn
    Public DocumentNo As String = Nothing
    Const colBinNo As String = "colBinNo"
    Const colPricipleCode As String = "colPricipleCode"
    Const colPricipleDesc As String = "colPricipleDesc"
    Const colvendorCode As String = "colvendorCode"
    Const colvendorDesc As String = "colvendorDesc"
    Const ColActualBalQty As String = "ColActualBalQty"
    Const colPriceDateColumn As String = "pricedate"
    Const colItemWeight As String = "colItemWeight"
    Const colConvF As String = "colConvF"
    Const colTotItemWt As String = "colTotItemWt"
    Const ColFOC As String = "ColFOC"
    Const colSchemeApplicable As String = "colSchemeApplicable"
    Const colDiscountAmount As String = "colDiscountAmount"
    Const colcustDiscount As String = "colcustDiscount"
    Const colActualCost As String = "colActualCost"
    Const colTotalMRP As String = "totalMRP"
    Const colTotalBasicAmount As String = "totalBasicAmount"
    Const colTotalDiscountAmount As String = "totalDiscountAmount"
    Const colTotalCustDiscount As String = "totalCustDiscount"
    Const colSchemeItem As String = "colSchemeItem"
    Const colFromSchemeCode As String = "colFromSchemeCode"
    Const ColCustDiscountQty As String = "ColCustDiscountQty"
    Const ColTPTRate As String = "ColTPTRate"
    Const ColTPTAmt As String = "ColTPTAmt"
    Const colAbatementPer As String = "colAbatementPer"
    Const colAbatementAmount As String = "colAbatementamount"
    Const colPriceCOde As String = "colPriceCOde"
    Const colMarkUpPercentage As String = "colMarkUpPercentage"
    Const colLandingCost As String = "colLandingCost"
    Const colMarkupOn As String = "colMarkupOn"
    Const colCustDiscPercentage As String = "colCustDiscPercentage"
    Const colCashDiscSchemeCode As String = "colCashDiscSchemeCode"
    Const colHeadDiscamt As String = "colHeadDiscamt"
    Const colPurCost As String = "colPurCost"
    Const colOrgCost As String = "colOrgCost"
    Const ColCommParty As String = "ColCommParty"
    Const ColCommPartyName As String = "ColCommPartyName"
    Const colCommRate As String = "colCommRate"
    Const ColCommAmt As String = "ColCommAmt"
    Const ColAmtAfterCOmm As String = "ColAmtAfterCOmm"
    Dim atchqry As String = ""
    Public IsDataImported As Boolean = False
    Public gvExcel As New RadGridView
    Public row_index As Integer
    Public DtExcel As DataTable
    Dim objsr As New clsSerialPort
    Dim isTaxExempted As Boolean = False
    Dim AllowRoundOff_onInvoice As Boolean = False
    Dim AllowRoundOffAmountOnDCSSale As Boolean = False
    Dim AllowPlandDeptMCCLocation As Boolean = False
    Dim AreaWiseBilling As Boolean = False
    Dim ShowAllCustomer As Boolean = False
    Dim EnableDynamicQRCodeForB2CInvoice As Boolean = False
    Dim arrLoc As String = ""

#End Region
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        If btnSave.Visible = True Then
            Export_details.Enabled = True
            Export_Head.Enabled = True
        Else
            Export_details.Enabled = False
            Export_Head.Enabled = False
        End If
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnPrint.Visible = MyBase.isPrintFlag
        rbtnCancel.Enabled = MyBase.isCancel_Flag
        If MyBase.isReverse Then
            btnReverseAndUnpost.Enabled = True
        Else
            btnReverseAndUnpost.Enabled = False
        End If
    End Sub
    Private Sub FrmAPInvoiceEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ApplySaleRouteInDCSSale = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ApplySaleRouteInDCSSale, clsFixedParameterCode.ApplySaleRouteInDCSSale, Nothing)) = "1", True, False))
        If ApplySaleRouteInDCSSale Then
            PnlSaleRoute.Visible = True
        End If
        arrMCCRights = clsMCCCodes.GetUserHavingMCCRights()
        CalculateTaxRatefromItemwsieTaxOnSale = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, Nothing))
        RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Collapsed
        SetUserMgmtNew()
        btnHistory.Enabled = False
        SetMailRight()
        ItemRateEditable = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.IsItemRateEditableOnSales & "'")) = 0, False, True)
        ItemMRPEditable = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.IsItemMRPEditableOnSales & "'")) = 0, False, True)
        RunBatchFifowise = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, Nothing))
        RunBatchFifowisewithmodifyfunctionality = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, Nothing)) = 1, True, False)
        checkstockmrpwise = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.checkstockMRPwise, clsFixedParameterCode.checkstockMRPwise, Nothing)) = 0, False, True)
        AutoScheme = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AutoSchemeOn & "'")) = 0, False, True)
        MultiplySubsidyWithQuantity = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MultiplySubsidyWithQuantity, clsFixedParameterCode.MultiplySubsidyWithQuantity, Nothing))
        'ExcludeKKFAndMandiForDCS = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ExcludeKKFAndMandiForDCS, clsFixedParameterCode.ExcludeKKFAndMandiForDCS, Nothing))
        blnBackCalculation = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select IsRateBackCalculation from TSPL_inv_parameters")) = 0, False, True)
        PurchaseOneItemOneVendor = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='PurchaseOneItemOneVendor'")) = 0, False, True)
        intMRPwithabatement = clsDBFuncationality.getSingleValue("select IsMRPwithAbatement from TSPL_INV_PARAMETERS")
        IsBatchMFDEXDmandatory = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select IsBatchNo_MFD_EXD_Mandatory from TSPL_inv_parameters")) = 0, False, True)
        AllowChangeInvoiceType = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Allow_Change_InvoiceType from TSPL_inv_parameters")) = 0, False, True)
        isPO_GRN_MRN_Editable = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select isMRNQtyEdiatableOnSRN from TSPL_inv_parameters")) = 0, False, True)
        UseDescInsteadOFCodeOnMCCMAterialSale = IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.UseDescInsteadOFCodeOnMCCMAterialSale, clsFixedParameterCode.UseDescInsteadOFCodeOnMCCMAterialSale, Nothing)) = "1", True, False)
        CreateAutoMCCPriceChat = IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateAutoMCCPriceChat, clsFixedParameterCode.CreateAutoMCCPriceChat, Nothing)) = "1", True, False)
        AllowRoundOff_onInvoice = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowRoundOff_OnCSASalePatti, clsFixedParameterCode.AllowRoundOff_OnCSASalePatti, Nothing)) = "1", True, False))
        AllowRoundOffAmountOnDCSSale = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowRoundOffAmountOnDCSSale, clsFixedParameterCode.AllowRoundOffAmountOnDCSSale, Nothing)) = "1", True, False))
        lblRound_Off.Visible = AllowRoundOffAmountOnDCSSale
        txtRoundOff.Visible = AllowRoundOffAmountOnDCSSale
        ShowAllCustomer = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ShowAllCustomerOnMccMaterialSale, clsFixedParameterCode.ShowAllCustomerOnMccMaterialSale, Nothing)) = "1", True, False))
        AreaWiseBilling = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, Nothing)) = "1", True, False))
        AllowPlandDeptMCCLocation = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.Allow_Plant_Depot_MCC_typeLocation, clsFixedParameterCode.Allow_Plant_Depot_MCC_typeLocation, Nothing)) = "1", True, False))
        EnableDynamicQRCodeForB2CInvoice = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.EnableDynamicQRCodeForB2CInvoice, clsFixedParameterCode.EnableDynamicQRCodeForB2CInvoice, Nothing)) = "1", True, False))
        AmountToCheckCustomerOutstandingForTCSTax = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.AmountToCheckCustomerOutstandingForTCSTax, clsFixedParameterCode.AmountToCheckCustomerOutstandingForTCSTax, Nothing))
        EnableTCSRateValidityFrom01July2021 = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.EnableTCSRateValidityFrom01July2021, clsFixedParameterCode.EnableTCSRateValidityFrom01July2021, Nothing)) = 0, False, True)
        ConvertIntoBillingUOM = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ConvertTOBillingUOM, clsFixedParameterCode.ConvertTOBillingUOM, Nothing)) = 1, True, False)
        DefaultEnableEWayBill = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DefaultEnableEWayBill, clsFixedParameterCode.DefaultEnableEWayBill, Nothing)) = 1, True, False)
        ApplyEWBThresholdLimit = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.ApplyEWBThresholdLimit, Nothing)) = 1, True, False)
        EWBThresholdLimitForIntraCity = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraCity, Nothing))
        EWBThresholdLimitForIntraState = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraState, Nothing))
        EWBThresholdLimitForInterState = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForInterState, Nothing))
        DefaultEnableNoTransporter = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DefaultEnableNoTransporter, clsFixedParameterCode.DefaultEnableNoTransporter, Nothing)) = 1, True, False)
        dtpChallan.Value = clsCommon.GETSERVERDATE
        dtpInvoice.Value = clsCommon.GETSERVERDATE
        chkVendorGrossReceipt.Visible = False
        txtVendorNo.MendatroryField = True
        ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D Delete Trasnaction")
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnAddNew, "Press Alt+N Adding New Trasnaction")
        ButtonToolTip.SetToolTip(btnAddNew, "Press Alt+A Create Additional Cost")
        RadPageView1.SelectedPage = RadPageViewPage1
        RadPageViewPage5.Item.Visibility = ElementVisibility.Collapsed
        LoadBlankGrid()
        LoadBlankGridTax()
        LoadItemType()
        LoadInvoiceType()
        LoadDeductionType()
        LoadBlankGridAC()
        AddNew()
        SetLength()
        If MultiplySubsidyWithQuantity Then
            txtRateAmt.Enabled = True
            txtDiscAmt.Enabled = True
            lblTotalSubsidy.Visible = True
            lblTotalDisSubsidy.Visible = True
        Else
            txtRateAmt.Enabled = False
            txtDiscAmt.Enabled = False
            lblTotalSubsidy.Visible = False
            lblTotalDisSubsidy.Visible = False
        End If
        If clsCommon.myLen(strSRNno) > 0 Then
            LoadData(strSRNno, NavigatorType.Current)
        End If
        chkRateDefaultSetting.ToggleState = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, Nothing)) = 1, ToggleState.On, ToggleState.Off)
        'For Custom Fields
        RadPageView1.Pages("pvpCustomFields").Item.Visibility = MyBase.customFieldTabProperty
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.Report_ID = MyBase.Form_ID
            UcCustomFields1.LoadCustomControls()
        End If
        'End of For Custom Fields
        SetMultiCurrencyVisibility()
        'If AllowPlandDeptMCCLocation = False Then


        '    Dim obj As New clsMCCCodes()
        '    obj = clsMCCCodes.GetData(True)
        '    If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Default_LocCode) > 1 Then
        '        txtBillToLocation.Value = obj.Default_LocCode
        '        lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "' "))
        '    Else
        '    End If
        'End If
        'lblBillToLocation.MyReadOnly = True

        If AreaWiseBilling Then
            txtBillToLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER inner join tspl_mcc_master on mcc_code=Default_Location  where User_Code='" + objCommonVar.CurrentUserCode + "' "))
            If clsCommon.myLen(txtBillToLocation.Value) > 0 Then
                lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_Location_Master where Location_Code='" + txtBillToLocation.Value + "' "))
            Else
                lblBillToLocation.Text = ""
            End If
            txtBillToLocation.Enabled = False
            lblBillToLocation.Enabled = False
        End If

        If objCommonVar.IsDemoERP Then
            UcAttachment1.Form_ID = MyBase.Form_ID
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Visible
        Else
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Collapsed
        End If
        If clsCommon.myLen(DocumentNo) > 0 Then
            LoadData(DocumentNo, NavigatorType.Current)
        End If
        If clsCommon.myLen(Me.Tag) > 0 Then
            LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
        End If
        txtBillToLocation.MyReadOnly = True
        If IsDataImported = True Then
            Me.CenterToParent()
            If LoadImportData(gvExcel, row_index, DtExcel) Then
                Me.Close()
            End If
        End If
        btnDeliveredTo.Enabled = False
        If objCommonVar.CreateAutoReceiptEntryDCSSale Then
            txtReceiptNo.Enabled = False
        End If
    End Sub
    Sub SetMultiCurrencyVisibility()
        Dim strq As String = ""
        If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
            pnlCurrConv.Visible = True
            If clsCommon.myLen(Me.txtVendorNo.Value) > 0 Then
                strq = "select currency_code from TSPL_CUSTOMER_MASTER where cust_code='" & clsCommon.myCstr(Me.txtVendorNo.Value) & "'"
                Me.txtCurrencyCode.Value = clsDBFuncationality.getSingleValue(strq).ToString
                ShowCurrencyDetail()
            End If
            ShowCurrencyDetail()
        Else
            pnlCurrConv.Visible = False
        End If
    End Sub
    Sub ShowCurrencyDetail()
        Dim dt As DataTable
        If clsCommon.myLen(clsCommon.myCstr(Me.txtVendorNo.Value)) = 0 Then
            Exit Sub
        End If
        If clsCommon.myLen(txtCurrencyCode.Value) > 0 Then
            dt = clsModuleCurrencyMapping.GetLatestCurConvRateDT(Me.txtDate.Value, txtCurrencyCode.Value)
            If dt.Rows.Count = 0 Then
                If Me.txtCurrencyCode.Value = objCommonVar.BaseCurrencyCode Then
                    Me.txtConversionRate.Text = 1
                    Me.txtApplicableFrom.Text = ""
                Else
                    clsCommon.MyMessageBoxShow("Conversion rate not entered for currency '" & Me.txtCurrencyCode.Value & "'")
                    Exit Sub
                End If
            Else
                Me.txtConversionRate.Text = dt.Rows(0).Item("Rate")
                Me.txtApplicableFrom.Text = clsCommon.GetPrintDate(dt.Rows(0).Item("FROM_DATE"), "dd/MMM/yyyy")
            End If
        Else
            Me.txtConversionRate.Text = 1
            Me.txtApplicableFrom.Text = ""
        End If
    End Sub
    Sub SetLength()
        txtDocNo.MyMaxLength = 30
        txtDesc.MaxLength = 200
        txtComment.MaxLength = 200
        txtRefNo.MaxLength = 50
        txtCarrier.MaxLength = 50
        txtGRNo.MaxLength = 50
        txtGENo.MaxLength = 50
    End Sub
    Sub LoadItemType()
        cboItemType.DataSource = clsItemMaster.GetItemType()
        cboItemType.ValueMember = "Code"
        cboItemType.DisplayMember = "Name"
        cboItemType.SelectedIndex = 2
        cboItemType.Visible = False
        RadLabel29.Visible = False
    End Sub
    Sub BlankAllControls()
        txtPaymentCycleNo.Text = ""
        txtFiscalYear.Text = ""
        txtEWayBillNo.Text = ""
        txtEWayBillDate.Checked = False
        txtTPTAmt.Text = 0
        txtEWayBillDate.Value = clsCommon.GETSERVERDATE()
        txtDiscAmt.Text = 0
        txtDiscPer.Text = 0
        lblDiscountAmt.Text = 0
        txtTransporterCode.Value = ""
        lblTransporterName.Text = ""
        lblInvoiceDiscAmt.Text = 0
        chkDiscountOnAmt.IsChecked = True
        chkRateDiffAmt.IsChecked = True
        chkNoTranspoter.IsChecked = False
        If chkRateDiffRate.IsChecked Then
            GroupBoxRateDiffAmtType.Visible = True
        ElseIf chkRateDiffAmt.IsChecked Then
            GroupBoxRateDiffAmtType.Visible = False
        End If
        txtNoOfInsallment.Value = 0
        chkExcludeKKFMandi.Checked = False
        LblVlc_Code.Text = ""
        LblVlc_Code.Tag = Nothing
        LblVlc_Name.Text = ""
        chkAutoTransfer.Checked = False
        txtFromLoc.Value = ""
        ddlInvoiceType.SelectedValue = ""
        txtDocNo.Value = ""
        txtDesc.Text = ""
        chkcashsale.Checked = False
        chkOther.Checked = False
        chkcashsale.Enabled = True
        chkAddTPT.Checked = True
        chkApplyTPT.Checked = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select count (1) from TSPL_DCS_TRANSPORTATION_CHARGES_HEAD ") > 0)
        txtTPTVendor.Value = ""
        chkOnHold.Checked = False
        txtVendorNo.Value = ""
        lblVendorName.Text = ""
        FndPriceCode.Value = ""
        txtSaleRoute.Value = ""
        FndPriceCode.Enabled = True
        FlagDocumentIsTaxable = 0
        setDCSBalance()
        txtDate.Value = clsCommon.GETSERVERDATE()
        dtpChallan.Value = clsCommon.GETSERVERDATE()
        dtpInvoice.Value = clsCommon.GETSERVERDATE()
        txtBillToLocation.Value = ""
        lblBillToLocation.Text = ""
        txtSubLocation.Value = ""
        lblSubLocation.Text = ""
        txtDesc.Text = ""
        txtInvNo.Text = ""
        txtComment.Text = ""
        txtRecommBy.Text = ""
        txtTaxGroup.Value = ""
        lblTaxGrpName.Text = ""
        txtTermCode.Value = ""
        lblTermName.Text = ""
        txtDueDate.Value = txtDate.Value
        txtRefNo.Text = ""
        lblAmtWithDiscount.Text = ""
        lblDiscountAmt.Text = ""
        lblAmtAfterDiscount.Text = ""
        lblTaxAmt.Text = ""
        lblTotRAmt.Text = ""
        lblTotRAmt1.Text = ""
        UsLock1.Status = ERPTransactionStatus.Pending
        txtCarrier.Text = ""
        txtGRNo.Text = ""
        txtGENo.Text = ""
        txtGEDate.Checked = False
        txtGEDate.Value = txtDate.Value
        txtDept.Value = ""
        lblDept.Text = ""
        cboItemType.SelectedIndex = 2
        cboItemType.Enabled = True
        txtBillToLocation.Enabled = True
        txtSubLocation.Enabled = True
        chkVendorGrossReceipt.Checked = False
        lblAddCharges1.Text = ""
        lblAddCharges1.Text = ""
        rbtnTaxCalAutomatic.IsChecked = True
        txtInvoiceNo.Text = ""
        txtMannaulInvoiceNo.Value = 0
        TxtInvoiceManualNoWithPrefix.Text = ""
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.BlankAllControls()
        End If
        UcAttachment1.BlankAllControls()
        fndProject.Value = ""
        lblProject.Text = ""
        txtPriceCode.Text = ""
        txtRoundOff.Text = ""
        chkTaxable.Checked = False
        cboDeductionType.SelectedValue = ""

        TxtEInvoiceUpdateIRNNo.Text = ""
        TxtEInvoiceUpdateAckNo.Text = ""
        TxtEInvoiceUpdateAckDate.Value = DateTime.Now
        TxtEInvoiceUpdateQCCode.Text = ""
        TxtEWayBillUpdateBillNo.Text = ""
        TxtEWayBillUpdateBillDate.Value = DateTime.Now
        TxtEWayBillUpdateValidDate.Value = DateTime.Now
        TxtEWayBillUpdateBillRemarks.Text = ""
        If DefaultEnableNoTransporter Then
            txtTransporterCode.Enabled = False
            lblTransporterName.Enabled = False
            chkNoTranspoter.Checked = True
        Else
            txtTransporterCode.Enabled = True
            lblTransporterName.Enabled = True
            chkNoTranspoter.Checked = False
        End If
        lblTPTVendor.Visible = False
        txtTPTVendor.Visible = False
    End Sub
    Public Shared Function GetItemType() As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("Code", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = RowTypeItem
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = RowTypeMisc
        dt.Rows.Add(dr)
        Return dt
    End Function
    Sub LoadBlankGrid()
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLineNo)
        repoComplete = New GridViewTextBoxColumn()
        repoComplete.FormatString = ""
        repoComplete.HeaderText = "Complete"
        repoComplete.Width = 70
        repoComplete.Name = colComplete
        repoComplete.ReadOnly = False
        repoComplete.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoComplete)
        Dim repoRowType As GridViewComboBoxColumn = New GridViewComboBoxColumn()
        repoRowType.FormatString = ""
        repoRowType.HeaderText = "Row Type"
        repoRowType.Name = colRowType
        repoRowType.Width = 0
        repoRowType.ReadOnly = False
        repoRowType.IsVisible = False
        repoRowType.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        repoRowType.DataSource = GetItemType()
        repoRowType.ValueMember = "Code"
        repoRowType.DisplayMember = "Code"
        gv1.MasterTemplate.Columns.Add(repoRowType)
        Dim repoShortDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoShortDesc.FormatString = ""
        repoShortDesc.HeaderText = "Short Desc"
        repoShortDesc.Name = colShortDesc
        repoShortDesc.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoShortDesc.TextImageRelation = TextImageRelation.TextBeforeImage
        repoShortDesc.Width = 100
        repoShortDesc.IsVisible = UseDescInsteadOFCodeOnMCCMAterialSale
        gv1.MasterTemplate.Columns.Add(repoShortDesc)
        Dim repoICode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoICode.FormatString = ""
        repoICode.HeaderText = "Item Code"
        repoICode.Name = colICode
        repoICode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoICode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoICode.Width = 100
        repoICode.ReadOnly = UseDescInsteadOFCodeOnMCCMAterialSale
        gv1.MasterTemplate.Columns.Add(repoICode)
        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIName.FormatString = ""
        repoIName.HeaderText = "Item Description"
        repoIName.Name = colIName
        repoIName.Width = 150
        repoIName.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIName)

        Dim repoManualBatchNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoManualBatchNo.FormatString = ""
        repoManualBatchNo.HeaderText = "Manual Batch No"
        repoManualBatchNo.Name = colManualBatchNo
        repoManualBatchNo.ReadOnly = True
        repoManualBatchNo.IsVisible = True
        repoManualBatchNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoManualBatchNo)

        Dim repoDCSSaleZeroCost As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDCSSaleZeroCost.FormatString = ""
        repoDCSSaleZeroCost.HeaderText = "Is DCS Sale Zero Cost"
        repoDCSSaleZeroCost.Name = colDCSSaleZeroCost
        repoDCSSaleZeroCost.Width = 150
        repoDCSSaleZeroCost.ReadOnly = True
        repoDCSSaleZeroCost.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoDCSSaleZeroCost)

        repoIName = New GridViewTextBoxColumn()
        repoIName.FormatString = ""
        repoIName.HeaderText = "HSN No"
        repoIName.Name = colHSNNo
        repoIName.Width = 150
        repoIName.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIName)
        Dim repoPriceDate As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoPriceDate.Format = DateTimePickerFormat.Custom
        repoPriceDate.CustomFormat = "dd-MM-yyyy"
        repoPriceDate.HeaderText = "Price Date"
        repoPriceDate.WrapText = True
        repoPriceDate.FormatString = "{0:d}"
        repoPriceDate.Name = colPriceDateColumn
        repoPriceDate.ReadOnly = True
        repoPriceDate.IsVisible = False
        repoPriceDate.Width = 0
        gv1.MasterTemplate.Columns.Add(repoPriceDate)
        Dim repoPriceCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPriceCode.FormatString = ""
        repoPriceCode.HeaderText = "Price Code"
        repoPriceCode.Name = colPriceCOde
        repoPriceCode.IsVisible = False
        repoPriceCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoPriceCode)
        Dim repoBarcode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBarcode.FormatString = ""
        repoBarcode.HeaderText = "BAR Code"
        repoBarcode.Name = colBarCode
        repoBarcode.IsVisible = False
        repoBarcode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoBarcode)
        Dim repoPendingQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPendingQty = New GridViewDecimalColumn()
        repoPendingQty.FormatString = ""
        repoPendingQty.HeaderText = "Pending Quantity"
        repoPendingQty.Name = colPendingQty
        repoPendingQty.IsVisible = False
        repoPendingQty.Minimum = 0
        repoPendingQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoPendingQty.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoPendingQty)
        repoBalQty = New GridViewDecimalColumn()
        repoBalQty.FormatString = ""
        repoBalQty.WrapText = True
        repoBalQty.HeaderText = "Balance Quantity"
        repoBalQty.Name = colBalanceQty
        repoBalQty.Width = 80
        repoBalQty.Minimum = 0
        repoBalQty.IsVisible = False
        repoBalQty.ReadOnly = True
        repoBalQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoBalQty)
        'Dim repoOrgSRNQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        'repoOrgSRNQty.FormatString = ""
        'repoOrgSRNQty.WrapText = True
        'repoOrgSRNQty.HeaderText = "Order Qty"
        'repoOrgSRNQty.Name = colOrgSOQty
        'repoOrgSRNQty.Width = 80
        'repoOrgSRNQty.Minimum = 0
        'repoOrgSRNQty.ReadOnly = Not isPO_GRN_MRN_Editable
        'repoOrgSRNQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        'gv1.MasterTemplate.Columns.Add(repoOrgSRNQty)
        Dim repoBillingUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBillingUnit.FormatString = ""
        repoBillingUnit.HeaderText = "Billing UOM"
        repoBillingUnit.Name = colBillingUnit
        repoBillingUnit.Width = 80
        repoBillingUnit.ReadOnly = True
        repoBillingUnit.IsVisible = False
        repoBillingUnit.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoBillingUnit.TextImageRelation = TextImageRelation.TextBeforeImage
        gv1.MasterTemplate.Columns.Add(repoBillingUnit)
        Dim repoBillingQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoBillingQty = New GridViewDecimalColumn()
        repoBillingQty.FormatString = ""
        repoBillingQty.HeaderText = "Billing Qty"
        repoBillingQty.Name = colBillingQty
        repoBillingQty.Width = 80
        repoBillingQty.Minimum = 0
        repoBillingQty.IsVisible = False
        repoBillingQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoBillingQty)
        Dim repoUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnit.FormatString = ""
        repoUnit.HeaderText = "UOM"
        repoUnit.Name = colUnit
        repoUnit.Width = 80
        repoUnit.ReadOnly = False
        repoUnit.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoUnit.TextImageRelation = TextImageRelation.TextBeforeImage
        gv1.MasterTemplate.Columns.Add(repoUnit)

        Dim repoUnitName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnitName.FormatString = ""
        repoUnitName.HeaderText = "UOM Description"
        repoUnitName.Name = colUOMName
        repoUnitName.Width = 150
        repoUnitName.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoUnitName)
        Dim repoOrgUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoOrgUnit.FormatString = ""
        repoOrgUnit.HeaderText = "ORG UOM"
        repoOrgUnit.Name = colOrgUnit
        repoOrgUnit.Width = 80
        repoOrgUnit.ReadOnly = False
        repoOrgUnit.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoOrgUnit)
        Dim repoQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoQty = New GridViewDecimalColumn()
        repoQty.FormatString = ""
        repoQty.HeaderText = "Quantity"
        repoQty.Name = colQty
        repoQty.Width = 80
        repoQty.Minimum = 0
        repoQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoQty)
        Dim repoActualBalQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoActualBalQty.FormatString = ""
        repoActualBalQty.HeaderText = "Actual Balance"
        repoActualBalQty.Name = ColActualBalQty
        repoActualBalQty.Width = 0
        repoActualBalQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoActualBalQty.ReadOnly = True
        repoActualBalQty.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoActualBalQty)
        Dim repoSchemeApp As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSchemeApp.AllowSort = False
        repoSchemeApp.FormatString = ""
        repoSchemeApp.HeaderText = "App. Qty Dis."
        repoSchemeApp.Name = colSchemeApplicable
        repoSchemeApp.ReadOnly = True
        repoSchemeApp.IsVisible = False
        repoSchemeApp.Width = 75
        gv1.MasterTemplate.Columns.Add(repoSchemeApp)
        Dim repoItemWt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoItemWt = New GridViewDecimalColumn()
        repoItemWt.FormatString = ""
        repoItemWt.HeaderText = "Item Weight"
        repoItemWt.Name = colItemWeight
        repoItemWt.Width = 80
        repoItemWt.Minimum = 0
        repoItemWt.IsVisible = False
        repoItemWt.ReadOnly = False
        repoItemWt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoItemWt)
        Dim repoConv As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoConv = New GridViewDecimalColumn()
        repoConv.FormatString = ""
        repoConv.HeaderText = "Conv. Factor"
        repoConv.Name = colConvF
        repoConv.Width = 80
        repoConv.Minimum = 0
        repoConv.ReadOnly = False
        repoConv.IsVisible = False
        repoConv.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoConv)
        Dim repoTotItemWt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotItemWt = New GridViewDecimalColumn()
        repoTotItemWt.FormatString = ""
        repoTotItemWt.HeaderText = "Tot Item Weight"
        repoTotItemWt.Name = colTotItemWt
        repoTotItemWt.Width = 80
        repoTotItemWt.Minimum = 0
        repoTotItemWt.ReadOnly = False
        repoTotItemWt.IsVisible = False
        repoTotItemWt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTotItemWt)
        Dim repoMRP As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMRP.FormatString = ""
        repoMRP.HeaderText = "MRP"
        repoMRP.Name = colMRP
        repoMRP.Width = 0
        repoMRP.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoMRP.ReadOnly = False
        repoMRP.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoMRP)
        Dim repoRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoRate = New GridViewDecimalColumn()
        repoRate.FormatString = ""
        repoRate.HeaderText = "Price Per Unit"
        repoRate.Name = colRate
        repoRate.Width = 100
        repoRate.Minimum = 0
        If CreateAutoMCCPriceChat Then
            repoRate.ReadOnly = False
        End If
        repoRate.DecimalPlaces = 6
        repoRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoRate)
        Dim repoOrgRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoOrgRate = New GridViewDecimalColumn()
        repoOrgRate.FormatString = ""
        repoOrgRate.HeaderText = "Org Basic Rate"
        repoOrgRate.Name = colOrgCost
        repoOrgRate.Width = 0
        repoOrgRate.Minimum = 0
        repoOrgRate.ReadOnly = True
        repoOrgRate.IsVisible = False
        repoOrgRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoOrgRate)
        Dim repoQtySchemeItem As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoQtySchemeItem.AllowSort = False
        repoQtySchemeItem.HeaderText = "Qty Scheme Item"
        repoQtySchemeItem.Name = colSchemeItem
        repoQtySchemeItem.ReadOnly = True
        repoQtySchemeItem.IsVisible = False
        repoQtySchemeItem.Width = 96
        gv1.MasterTemplate.Columns.Add(repoQtySchemeItem)
        Dim repoFromSchemeCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFromSchemeCode.HeaderText = "Scheme Code"
        repoFromSchemeCode.Name = colFromSchemeCode
        repoFromSchemeCode.Width = 80
        repoFromSchemeCode.ReadOnly = True
        repoFromSchemeCode.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoFromSchemeCode)
        Dim repoFreeQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoFreeQty.FormatString = ""
        repoFreeQty.HeaderText = "Free Quantity"
        repoFreeQty.Name = colFreeQty
        repoFreeQty.IsVisible = False
        repoFreeQty.Width = 80
        repoFreeQty.Minimum = 0
        repoFreeQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoFreeQty)
        Dim repoLocationCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoLocationCode.FormatString = ""
        repoLocationCode.HeaderText = "Location Code"
        repoLocationCode.Name = colLocationCode
        repoLocationCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoLocationCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoLocationCode.Width = 100
        repoLocationCode.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoLocationCode)
        Dim repoLocationName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoLocationName.FormatString = ""
        repoLocationName.HeaderText = "Location"
        repoLocationName.Name = colLocationName
        repoLocationName.ReadOnly = True
        repoLocationName.Width = 150
        repoLocationName.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoLocationName)
        Dim repoAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmt = New GridViewDecimalColumn()
        repoAmt.FormatString = ""
        repoAmt.HeaderText = "Basic Sale Amount"
        repoAmt.Name = colAmt
        repoAmt.Width = 80
        repoAmt.Minimum = 0
        repoAmt.ReadOnly = False
        repoAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmt.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoAmt)
        Dim repoAbatementRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAbatementRate = New GridViewDecimalColumn()
        repoAbatementRate.FormatString = ""
        repoAbatementRate.HeaderText = "Abatement %"
        repoAbatementRate.Name = colAbatementPer
        repoAbatementRate.Width = 80
        repoAbatementRate.Minimum = 0
        repoAbatementRate.ReadOnly = False
        repoAbatementRate.IsVisible = False
        repoAbatementRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAbatementRate)
        Dim repoAbatementamount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAbatementamount = New GridViewDecimalColumn()
        repoAbatementamount.FormatString = ""
        repoAbatementamount.HeaderText = "Abatement Amount"
        repoAbatementamount.Name = colAbatementAmount
        repoAbatementamount.Width = 80
        repoAbatementamount.Minimum = 0
        repoAbatementamount.ReadOnly = False
        repoAbatementamount.IsVisible = False
        repoAbatementamount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAbatementamount)
        Dim repoDisPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisPer = New GridViewDecimalColumn()
        repoDisPer.FormatString = ""
        repoDisPer.HeaderText = "Discount %"
        repoDisPer.Minimum = 0
        repoDisPer.Name = colDisPer
        repoDisPer.Width = 0
        repoDisPer.IsVisible = False
        repoDisPer.Maximum = 100
        repoDisPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDisPer)
        Dim repoDisAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisAmt.FormatString = ""
        repoDisAmt.HeaderText = "Discount Amount"
        repoDisAmt.WrapText = True
        repoDisAmt.Name = colDisAmt
        repoDisAmt.Width = 0
        repoDisAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisAmt.VisibleInColumnChooser = False
        repoDisAmt.ReadOnly = True
        repoDisAmt.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoDisAmt)
        Dim repoDisPerUnit As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisPerUnit = New GridViewDecimalColumn()
        repoDisPerUnit.FormatString = ""
        repoDisPerUnit.HeaderText = "Discount Per Unit"
        repoDisPerUnit.Minimum = 0
        repoDisPerUnit.Name = colDisPerUnit
        repoDisPerUnit.Width = 0
        repoDisPerUnit.IsVisible = True
        repoDisPerUnit.Maximum = 100
        repoDisPerUnit.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDisPerUnit)
        Dim repoDisUnitAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisUnitAmt.FormatString = ""
        repoDisUnitAmt.HeaderText = "Discount Unit Amount"
        repoDisUnitAmt.WrapText = True
        repoDisUnitAmt.Name = colDisUnitAmt
        repoDisUnitAmt.Width = 0
        repoDisUnitAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisUnitAmt.VisibleInColumnChooser = False
        repoDisUnitAmt.ReadOnly = True
        repoDisUnitAmt.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoDisUnitAmt)
        Dim repoHDisPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoHDisPer = New GridViewDecimalColumn()
        repoHDisPer.FormatString = ""
        repoHDisPer.HeaderText = "Head Discount %"
        repoHDisPer.Minimum = 0
        repoHDisPer.Name = colHeaDDisPer
        repoHDisPer.Width = 100
        repoHDisPer.ReadOnly = True
        repoHDisPer.IsVisible = False
        repoHDisPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoHDisPer)
        Dim repoHDisAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoHDisAmt = New GridViewDecimalColumn()
        repoHDisAmt.FormatString = ""
        repoHDisAmt.HeaderText = "Head Discount % Amount"
        repoHDisAmt.WrapText = True
        repoHDisAmt.Name = colHeadDisPerAmt
        repoHDisAmt.Width = 80
        repoHDisAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoHDisAmt.VisibleInColumnChooser = True
        repoHDisAmt.ReadOnly = True
        repoHDisAmt.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoHDisAmt)
        Dim repoHeadDisAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoHeadDisAmt = New GridViewDecimalColumn()
        repoHeadDisAmt.FormatString = ""
        repoHeadDisAmt.HeaderText = "Head Disc Amt"
        repoHeadDisAmt.WrapText = True
        repoHeadDisAmt.Name = colHeadDiscamt
        repoHeadDisAmt.Width = 80
        repoHeadDisAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoHeadDisAmt.VisibleInColumnChooser = False
        repoHeadDisAmt.ReadOnly = True
        repoHeadDisAmt.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoHeadDisAmt)

        Dim TransporterRate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        TransporterRate.FormatString = ""
        TransporterRate.HeaderText = "Transporter Commission Rate"
        TransporterRate.Name = ColTPTRate
        TransporterRate.Width = 100
        TransporterRate.ReadOnly = True
        TransporterRate.IsVisible = False
        TransporterRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(TransporterRate)

        Dim TransporterAmt As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        TransporterAmt.FormatString = ""
        TransporterAmt.HeaderText = "Transporter Commission Amt"
        TransporterAmt.Name = ColTPTAmt
        TransporterAmt.Width = 100
        TransporterAmt.ReadOnly = True
        TransporterAmt.IsVisible = False
        TransporterAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(TransporterAmt)

        Dim repoCustDiscountQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCustDiscountQty.HeaderText = "Cash Dis Qty."
        repoCustDiscountQty.MinWidth = 4
        repoCustDiscountQty.Name = ColCustDiscountQty
        repoCustDiscountQty.ReadOnly = True
        repoCustDiscountQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCustDiscountQty.Width = 0
        repoCustDiscountQty.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCustDiscountQty)
        Dim repoCustDiscountPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCustDiscountPer.HeaderText = "Cash Dis %."
        repoCustDiscountPer.MinWidth = 4
        repoCustDiscountPer.Name = colCustDiscPercentage
        repoCustDiscountPer.ReadOnly = True
        repoCustDiscountPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCustDiscountPer.Width = 0
        repoCustDiscountPer.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCustDiscountPer)
        Dim repoCustDiscount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCustDiscount.HeaderText = "Cash Dis Amt."
        repoCustDiscount.MinWidth = 4
        repoCustDiscount.Name = colcustDiscount
        repoCustDiscount.ReadOnly = True
        repoCustDiscount.IsVisible = False
        repoCustDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCustDiscount.Width = 0
        gv1.MasterTemplate.Columns.Add(repoCustDiscount)
        Dim repoCashSchemeCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCashSchemeCode.HeaderText = "Cash Scheme Code"
        repoCashSchemeCode.Name = colCashDiscSchemeCode
        repoCashSchemeCode.Width = 0
        repoCashSchemeCode.ReadOnly = True
        repoCashSchemeCode.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCashSchemeCode)
        Dim repoAcualCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAcualCost.AllowSort = False
        repoAcualCost.HeaderText = "Net Price"
        repoAcualCost.Name = colActualCost
        repoAcualCost.ReadOnly = True
        repoAcualCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAcualCost.Width = 0
        repoAcualCost.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAcualCost)
        Dim repoTotalMRP As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalMRP.AllowSort = False
        repoTotalMRP.HeaderText = "Total MRP"
        repoTotalMRP.Name = colTotalMRP
        repoTotalMRP.ReadOnly = True
        repoTotalMRP.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalMRP.Width = 0
        repoTotalMRP.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTotalMRP)
        Dim repoTotalBasicAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalBasicAmt.AllowSort = False
        repoTotalBasicAmt.HeaderText = "Total Basic Amount"
        repoTotalBasicAmt.Name = colTotalBasicAmount
        repoTotalBasicAmt.ReadOnly = True
        repoTotalBasicAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalBasicAmt.Width = 0
        repoTotalBasicAmt.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTotalBasicAmt)
        Dim repoTotalDiscount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalDiscount.AllowSort = False
        repoTotalDiscount.HeaderText = "Total Discount"
        repoTotalDiscount.Name = colTotalDiscountAmount
        repoTotalDiscount.ReadOnly = True
        repoTotalDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalDiscount.Width = 0
        repoTotalDiscount.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTotalDiscount)
        Dim repoTotalCustDiscount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalCustDiscount.HeaderText = "Total Cash Dis"
        repoTotalCustDiscount.Name = colTotalCustDiscount
        repoTotalCustDiscount.ReadOnly = True
        repoTotalCustDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalCustDiscount.Width = 0
        repoTotalCustDiscount.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTotalCustDiscount)
        Dim repoAmtAfterDis As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterDis = New GridViewDecimalColumn()
        repoAmtAfterDis.FormatString = ""
        repoAmtAfterDis.HeaderText = "Amount After Discount"
        repoAmtAfterDis.Name = colAmtAfterDis
        repoAmtAfterDis.WrapText = True
        repoAmtAfterDis.Width = 0
        repoAmtAfterDis.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterDis.VisibleInColumnChooser = True
        repoAmtAfterDis.ReadOnly = True
        repoAmtAfterDis.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterDis)
        Dim repoCommparty As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCommparty.FormatString = ""
        repoCommparty.HeaderText = "Comm. Party"
        repoCommparty.Name = ColCommParty
        repoCommparty.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoCommparty.TextImageRelation = TextImageRelation.TextBeforeImage
        repoCommparty.Width = 0
        repoCommparty.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCommparty)
        Dim repoCommPartyName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCommPartyName.FormatString = ""
        repoCommPartyName.HeaderText = "Comm. Party Name"
        repoCommPartyName.Name = ColCommPartyName
        repoCommPartyName.Width = 0
        repoCommPartyName.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCommPartyName)
        Dim repoCommPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCommPer = New GridViewDecimalColumn()
        repoCommPer.FormatString = ""
        repoCommPer.HeaderText = "Comm %"
        repoCommPer.Minimum = 0
        repoCommPer.Maximum = 100
        repoCommPer.Name = colCommRate
        repoCommPer.Width = 0
        repoCommPer.IsVisible = False
        repoCommPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoCommPer)
        Dim repoCommAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCommAmt = New GridViewDecimalColumn()
        repoCommAmt.FormatString = ""
        repoCommAmt.HeaderText = "Comm Amount"
        repoCommAmt.WrapText = True
        repoCommAmt.Name = ColCommAmt
        repoCommAmt.Width = 0
        repoCommAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCommAmt.VisibleInColumnChooser = False
        repoCommAmt.ReadOnly = True
        repoCommAmt.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCommAmt)
        Dim repoCommAmtAfterDis As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCommAmtAfterDis = New GridViewDecimalColumn()
        repoCommAmtAfterDis.FormatString = ""
        repoCommAmtAfterDis.HeaderText = "Amount After Comm."
        repoCommAmtAfterDis.Name = ColAmtAfterCOmm
        repoCommAmtAfterDis.WrapText = True
        repoCommAmtAfterDis.Width = 80
        repoCommAmtAfterDis.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCommAmtAfterDis.VisibleInColumnChooser = False
        repoCommAmtAfterDis.ReadOnly = True
        repoCommAmtAfterDis.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoCommAmtAfterDis)
        Dim repoPrincipleCOde As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPrincipleCOde.FormatString = ""
        repoPrincipleCOde.HeaderText = "Principle Code"
        repoPrincipleCOde.Name = colPricipleCode
        repoPrincipleCOde.Width = 150
        repoPrincipleCOde.ReadOnly = True
        repoPrincipleCOde.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoPrincipleCOde)
        Dim repoPrincipleDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPrincipleDesc.FormatString = ""
        repoPrincipleDesc.HeaderText = "Principle Desc"
        repoPrincipleDesc.Name = colPricipleDesc
        repoPrincipleDesc.Width = 150
        repoPrincipleDesc.ReadOnly = True
        repoPrincipleDesc.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoPrincipleDesc)
        Dim repoVCOde As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoVCOde.FormatString = ""
        repoVCOde.HeaderText = "Vendor Code"
        repoVCOde.Name = colvendorCode
        repoVCOde.Width = 150
        repoVCOde.ReadOnly = True
        repoVCOde.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoVCOde)
        Dim repoVDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoVDesc.FormatString = ""
        repoVDesc.HeaderText = "Vendor Desc"
        repoVDesc.Name = colvendorDesc
        repoVDesc.Width = 150
        repoVDesc.ReadOnly = True
        repoVDesc.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoVDesc)
        Dim repoMarkupPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMarkupPer = New GridViewDecimalColumn()
        repoMarkupPer.FormatString = ""
        repoMarkupPer.HeaderText = "Mark Up %"
        repoMarkupPer.Name = colMarkUpPercentage
        repoMarkupPer.WrapText = True
        repoMarkupPer.Width = 0
        repoMarkupPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoMarkupPer.VisibleInColumnChooser = False
        repoMarkupPer.ReadOnly = True
        repoMarkupPer.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoMarkupPer)
        Dim repoLandingCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLandingCost = New GridViewDecimalColumn()
        repoLandingCost.FormatString = ""
        repoLandingCost.HeaderText = "Landing Cost"
        repoLandingCost.Name = colLandingCost
        repoLandingCost.WrapText = True
        repoLandingCost.Width = 0
        repoLandingCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoLandingCost.VisibleInColumnChooser = False
        repoLandingCost.ReadOnly = True
        repoLandingCost.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoLandingCost)
        Dim repoPurCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPurCost = New GridViewDecimalColumn()
        repoPurCost.FormatString = ""
        repoPurCost.HeaderText = "purchase Cost"
        repoPurCost.Name = colPurCost
        repoPurCost.WrapText = True
        repoPurCost.Width = 0
        repoPurCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoPurCost.VisibleInColumnChooser = False
        repoPurCost.ReadOnly = True
        repoPurCost.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoPurCost)
        Dim repoMarkupOn As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoMarkupOn.FormatString = ""
        repoMarkupOn.HeaderText = "MarkUp On"
        repoMarkupOn.Name = colMarkupOn
        repoMarkupOn.ReadOnly = True
        repoMarkupOn.IsVisible = False
        repoMarkupOn.Width = 100
        gv1.MasterTemplate.Columns.Add(repoMarkupOn)
        Dim repoTax1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax1.FormatString = ""
        repoTax1.HeaderText = "Tax 1"
        repoTax1.Name = colTax1
        repoTax1.ReadOnly = True
        repoTax1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax1)
        Dim repoTaxBaseAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt1.FormatString = ""
        repoTaxBaseAmt1.HeaderText = "Tax Base Amount 1"
        repoTaxBaseAmt1.Name = colTaxBaseAmt1
        repoTaxBaseAmt1.ReadOnly = True
        repoTaxBaseAmt1.IsVisible = False
        repoTaxBaseAmt1.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt1)
        Dim repoTaxRate1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate1 = New GridViewDecimalColumn()
        repoTaxRate1.FormatString = ""
        repoTaxRate1.HeaderText = "Tax Rate 1"
        repoTaxRate1.Name = colTaxRate1
        repoTaxRate1.IsVisible = False
        repoTaxRate1.ReadOnly = True
        repoTaxRate1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate1)
        Dim repoTaxAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt1 = New GridViewDecimalColumn()
        repoTaxAmt1.FormatString = ""
        repoTaxAmt1.HeaderText = "Tax Amt 1"
        repoTaxAmt1.Name = colTaxAmt1
        repoTaxAmt1.IsVisible = False
        repoTaxAmt1.ReadOnly = True
        repoTaxAmt1.DecimalPlaces = 6
        repoTaxAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt1)
        Dim repoIsSurTax1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax1.HeaderText = "Is Surtax 1"
        repoIsSurTax1.Name = colIsSurTax1
        repoIsSurTax1.ReadOnly = True
        repoIsSurTax1.IsVisible = False
        repoIsSurTax1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax1)
        Dim repoSurTaxCode1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode1.FormatString = ""
        repoSurTaxCode1.HeaderText = "Surtax 1"
        repoSurTaxCode1.Name = colSurTaxCode1
        repoSurTaxCode1.ReadOnly = True
        repoSurTaxCode1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode1)
        Dim repoIsTaxable1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable1.HeaderText = "Is Taxable 1"
        repoIsTaxable1.Name = colIsTaxable1
        repoIsTaxable1.ReadOnly = True
        repoIsTaxable1.IsVisible = False
        repoIsTaxable1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable1)
        Dim repoTaxRecoverable1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable1.HeaderText = "Recoverable Tax 1"
        repoTaxRecoverable1.Name = colTaxRecoverable1
        repoTaxRecoverable1.ReadOnly = True
        repoTaxRecoverable1.IsVisible = False
        repoTaxRecoverable1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable1)
        Dim repoTax2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax2.FormatString = ""
        repoTax2.HeaderText = "Tax 2"
        repoTax2.Name = colTax2
        repoTax2.ReadOnly = True
        repoTax2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax2)
        Dim repoTaxBaseAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt2.FormatString = ""
        repoTaxBaseAmt2.HeaderText = "Tax Base Amount 2"
        repoTaxBaseAmt2.Name = colTaxBaseAmt2
        repoTaxBaseAmt2.ReadOnly = True
        repoTaxBaseAmt2.IsVisible = False
        repoTaxBaseAmt2.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt2)
        Dim repoTaxRate2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate2 = New GridViewDecimalColumn()
        repoTaxRate2.FormatString = ""
        repoTaxRate2.HeaderText = "Tax Rate 2"
        repoTaxRate2.Name = colTaxRate2
        repoTaxRate2.IsVisible = False
        repoTaxRate2.ReadOnly = True
        repoTaxRate2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate2)
        Dim repoTaxAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt2 = New GridViewDecimalColumn()
        repoTaxAmt2.FormatString = ""
        repoTaxAmt2.HeaderText = "Tax Amt 2"
        repoTaxAmt2.Name = colTaxAmt2
        repoTaxAmt2.IsVisible = False
        repoTaxAmt2.ReadOnly = True
        repoTaxAmt2.DecimalPlaces = 6
        repoTaxAmt2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt2)
        Dim repoIsSurTax2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax2.HeaderText = "Is Surtax 2"
        repoIsSurTax2.Name = colIsSurTax2
        repoIsSurTax2.ReadOnly = True
        repoIsSurTax2.IsVisible = False
        repoIsSurTax2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax2)
        Dim repoSurTaxCode2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode2.FormatString = ""
        repoSurTaxCode2.HeaderText = "Surtax 2"
        repoSurTaxCode2.Name = colSurTaxCode2
        repoSurTaxCode2.ReadOnly = True
        repoSurTaxCode2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode2)
        Dim repoIsTaxable2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable2.HeaderText = "Is Taxable 2"
        repoIsTaxable2.Name = colIsTaxable2
        repoIsTaxable2.ReadOnly = True
        repoIsTaxable2.IsVisible = False
        repoIsTaxable2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable2)
        Dim repoTaxRecoverable2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable2.HeaderText = "Recoverable Tax 2"
        repoTaxRecoverable2.Name = colTaxRecoverable2
        repoTaxRecoverable2.ReadOnly = True
        repoTaxRecoverable2.IsVisible = False
        repoTaxRecoverable2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable2)
        Dim repoTax3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax3.FormatString = ""
        repoTax3.HeaderText = "Tax 3"
        repoTax3.Name = colTax3
        repoTax3.ReadOnly = True
        repoTax3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax3)
        Dim repoTaxBaseAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt3.FormatString = ""
        repoTaxBaseAmt3.HeaderText = "Tax Base Amount 3"
        repoTaxBaseAmt3.Name = colTaxBaseAmt3
        repoTaxBaseAmt3.ReadOnly = True
        repoTaxBaseAmt3.IsVisible = False
        repoTaxBaseAmt3.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt3)
        Dim repoTaxRate3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate3 = New GridViewDecimalColumn()
        repoTaxRate3.FormatString = ""
        repoTaxRate3.HeaderText = "Tax Rate 3"
        repoTaxRate3.Name = colTaxRate3
        repoTaxRate3.IsVisible = False
        repoTaxRate3.ReadOnly = True
        repoTaxRate3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate3)
        Dim repoTaxAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt3 = New GridViewDecimalColumn()
        repoTaxAmt3.FormatString = ""
        repoTaxAmt3.HeaderText = "Tax Amt 3"
        repoTaxAmt3.Name = colTaxAmt3
        repoTaxAmt3.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt3.DecimalPlaces = 6
        repoTaxAmt3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt3)
        Dim repoIsSurTax3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax3.HeaderText = "Is Surtax 3"
        repoIsSurTax3.Name = colIsSurTax3
        repoIsSurTax3.ReadOnly = True
        repoIsSurTax3.IsVisible = False
        repoIsSurTax3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax3)
        Dim repoSurTaxCode3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode3.FormatString = ""
        repoSurTaxCode3.HeaderText = "Surtax 3"
        repoSurTaxCode3.Name = colSurTaxCode3
        repoSurTaxCode3.ReadOnly = True
        repoSurTaxCode3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode3)
        Dim repoIsTaxable3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable3.HeaderText = "Is Taxable 3"
        repoIsTaxable3.Name = colIsTaxable3
        repoIsTaxable3.ReadOnly = True
        repoIsTaxable3.IsVisible = False
        repoIsTaxable3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable3)
        Dim repoTaxRecoverable3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable3.HeaderText = "Recoverable Tax 3"
        repoTaxRecoverable3.Name = colTaxRecoverable3
        repoTaxRecoverable3.ReadOnly = True
        repoTaxRecoverable3.IsVisible = False
        repoTaxRecoverable3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable3)
        Dim repoTax4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax4.FormatString = ""
        repoTax4.HeaderText = "Tax 4"
        repoTax4.Name = colTax4
        repoTax4.ReadOnly = True
        repoTax4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax4)
        Dim repoTaxBaseAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt4.FormatString = ""
        repoTaxBaseAmt4.HeaderText = "Tax Base Amount 4"
        repoTaxBaseAmt4.Name = colTaxBaseAmt4
        repoTaxBaseAmt4.ReadOnly = True
        repoTaxBaseAmt4.IsVisible = False
        repoTaxBaseAmt4.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt4)
        Dim repoTaxRate4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate4 = New GridViewDecimalColumn()
        repoTaxRate4.FormatString = ""
        repoTaxRate4.HeaderText = "Tax Rate 4"
        repoTaxRate4.Name = colTaxRate4
        repoTaxRate4.IsVisible = False
        repoTaxRate4.ReadOnly = True
        repoTaxRate4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate4)
        Dim repoTaxAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt4 = New GridViewDecimalColumn()
        repoTaxAmt4.FormatString = ""
        repoTaxAmt4.HeaderText = "Tax Amt 4"
        repoTaxAmt4.Name = colTaxAmt4
        repoTaxAmt4.IsVisible = False
        repoTaxAmt4.ReadOnly = True
        repoTaxAmt4.DecimalPlaces = 6
        repoTaxAmt4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt4)
        Dim repoIsSurTax4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax4.HeaderText = "Is Surtax 4"
        repoIsSurTax4.Name = colIsSurTax4
        repoIsSurTax4.ReadOnly = True
        repoIsSurTax4.IsVisible = False
        repoIsSurTax4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax4)
        Dim repoSurTaxCode4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode4.FormatString = ""
        repoSurTaxCode4.HeaderText = "Surtax 4"
        repoSurTaxCode4.Name = colSurTaxCode4
        repoSurTaxCode4.ReadOnly = True
        repoSurTaxCode4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode4)
        Dim repoIsTaxable4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable4.HeaderText = "Is Taxable 4"
        repoIsTaxable4.Name = colIsTaxable4
        repoIsTaxable4.ReadOnly = True
        repoIsTaxable4.IsVisible = False
        repoIsTaxable4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable4)
        Dim repoTaxRecoverable4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable4.HeaderText = "Recoverable Tax 4"
        repoTaxRecoverable4.Name = colTaxRecoverable4
        repoTaxRecoverable4.ReadOnly = True
        repoTaxRecoverable4.IsVisible = False
        repoTaxRecoverable4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable4)
        Dim repoTax5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax5.FormatString = ""
        repoTax5.HeaderText = "Tax 5"
        repoTax5.Name = colTax5
        repoTax5.ReadOnly = True
        repoTax5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax5)
        Dim repoTaxBaseAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt5.FormatString = ""
        repoTaxBaseAmt5.HeaderText = "Tax Base Amount 5"
        repoTaxBaseAmt5.Name = colTaxBaseAmt5
        repoTaxBaseAmt5.DecimalPlaces = 6
        repoTaxBaseAmt5.ReadOnly = True
        repoTaxBaseAmt5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt5)
        Dim repoTaxRate5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate5 = New GridViewDecimalColumn()
        repoTaxRate5.FormatString = ""
        repoTaxRate5.HeaderText = "Tax Rate 5"
        repoTaxRate5.Name = colTaxRate5
        repoTaxRate5.IsVisible = False
        repoTaxRate5.ReadOnly = True
        repoTaxRate5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate5)
        Dim repoTaxAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt5 = New GridViewDecimalColumn()
        repoTaxAmt5.FormatString = ""
        repoTaxAmt5.HeaderText = "Tax Amt 5"
        repoTaxAmt5.Name = colTaxAmt5
        repoTaxAmt5.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt3.DecimalPlaces = 6
        repoTaxAmt5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt5)
        Dim repoIsSurTax5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax5.HeaderText = "Is Surtax 5"
        repoIsSurTax5.Name = colIsSurTax5
        repoIsSurTax5.ReadOnly = True
        repoIsSurTax5.IsVisible = False
        repoIsSurTax5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax5)
        Dim repoSurTaxCode5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode5.FormatString = ""
        repoSurTaxCode5.HeaderText = "Surtax 5"
        repoSurTaxCode5.Name = colSurTaxCode5
        repoSurTaxCode5.ReadOnly = True
        repoSurTaxCode5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode5)
        Dim repoIsTaxable5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable5.HeaderText = "Is Taxable 5"
        repoIsTaxable5.Name = colIsTaxable5
        repoIsTaxable5.ReadOnly = True
        repoIsTaxable5.IsVisible = False
        repoIsTaxable5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable5)
        Dim repoTaxRecoverable5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable5.HeaderText = "Recoverable Tax 5"
        repoTaxRecoverable5.Name = colTaxRecoverable5
        repoTaxRecoverable5.ReadOnly = True
        repoTaxRecoverable5.IsVisible = False
        repoTaxRecoverable5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable5)
        Dim repoTax6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax6.FormatString = ""
        repoTax6.HeaderText = "Tax 6"
        repoTax6.Name = colTax6
        repoTax6.ReadOnly = True
        repoTax6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax6)
        Dim repoTaxBaseAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt6.FormatString = ""
        repoTaxBaseAmt6.HeaderText = "Tax Base Amount 6"
        repoTaxBaseAmt6.Name = colTaxBaseAmt6
        repoTaxBaseAmt6.ReadOnly = True
        repoTaxBaseAmt6.IsVisible = False
        repoTaxBaseAmt6.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt6)
        Dim repoTaxRate6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate6 = New GridViewDecimalColumn()
        repoTaxRate6.FormatString = ""
        repoTaxRate6.HeaderText = "Tax Rate 6"
        repoTaxRate6.Name = colTaxRate6
        repoTaxRate6.IsVisible = False
        repoTaxRate6.ReadOnly = True
        repoTaxRate6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate6)
        Dim repoTaxAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt6 = New GridViewDecimalColumn()
        repoTaxAmt6.FormatString = ""
        repoTaxAmt6.HeaderText = "Tax Amt 6"
        repoTaxAmt6.Name = colTaxAmt6
        repoTaxAmt6.IsVisible = False
        repoTaxAmt6.ReadOnly = True
        repoTaxAmt6.DecimalPlaces = 6
        repoTaxAmt6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt6)
        Dim repoIsSurTax6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax6.HeaderText = "Is Surtax 6"
        repoIsSurTax6.Name = colIsSurTax6
        repoIsSurTax6.ReadOnly = True
        repoIsSurTax6.IsVisible = False
        repoIsSurTax6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax6)
        Dim repoSurTaxCode6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode6.FormatString = ""
        repoSurTaxCode6.HeaderText = "Surtax 6"
        repoSurTaxCode6.Name = colSurTaxCode6
        repoSurTaxCode6.ReadOnly = True
        repoSurTaxCode6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode6)
        Dim repoIsTaxable6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable6.HeaderText = "Is Taxable 6"
        repoIsTaxable6.Name = colIsTaxable6
        repoIsTaxable6.ReadOnly = True
        repoIsTaxable6.IsVisible = False
        repoIsTaxable6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable6)
        Dim repoTaxRecoverable6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable6.HeaderText = "Recoverable Tax 6"
        repoTaxRecoverable6.Name = colTaxRecoverable6
        repoTaxRecoverable6.ReadOnly = True
        repoTaxRecoverable6.IsVisible = False
        repoTaxRecoverable6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable6)
        Dim repoTax7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax7.FormatString = ""
        repoTax7.HeaderText = "Tax 7"
        repoTax7.Name = colTax7
        repoTax7.ReadOnly = True
        repoTax7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax7)
        Dim repoTaxBaseAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt7.FormatString = ""
        repoTaxBaseAmt7.HeaderText = "Tax Base Amount 7"
        repoTaxBaseAmt7.Name = colTaxBaseAmt7
        repoTaxBaseAmt7.ReadOnly = True
        repoTaxBaseAmt7.IsVisible = False
        repoTaxBaseAmt7.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt7)
        Dim repoTaxRate7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate7 = New GridViewDecimalColumn()
        repoTaxRate7.FormatString = ""
        repoTaxRate7.HeaderText = "Tax Rate 7"
        repoTaxRate7.Name = colTaxRate7
        repoTaxRate7.IsVisible = False
        repoTaxRate7.ReadOnly = True
        repoTaxRate7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate7)
        Dim repoTaxAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt7 = New GridViewDecimalColumn()
        repoTaxAmt7.FormatString = ""
        repoTaxAmt7.HeaderText = "Tax Amt 7"
        repoTaxAmt7.Name = colTaxAmt7
        repoTaxAmt7.DecimalPlaces = 6
        repoTaxAmt7.IsVisible = False
        repoTaxAmt7.ReadOnly = True
        repoTaxAmt7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt7)
        Dim repoIsSurTax7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax7.HeaderText = "Is Surtax 7"
        repoIsSurTax7.Name = colIsSurTax7
        repoIsSurTax7.ReadOnly = True
        repoIsSurTax7.IsVisible = False
        repoIsSurTax7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax7)
        Dim repoSurTaxCode7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode7.FormatString = ""
        repoSurTaxCode7.HeaderText = "Surtax 7"
        repoSurTaxCode7.Name = colSurTaxCode7
        repoSurTaxCode7.ReadOnly = True
        repoSurTaxCode7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode7)
        Dim repoIsTaxable7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable7.HeaderText = "Is Taxable 7"
        repoIsTaxable7.Name = colIsTaxable7
        repoIsTaxable7.ReadOnly = True
        repoIsTaxable7.IsVisible = False
        repoIsTaxable7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable7)
        Dim repoTaxRecoverable7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable7.HeaderText = "Recoverable Tax 7"
        repoTaxRecoverable7.Name = colTaxRecoverable7
        repoTaxRecoverable7.ReadOnly = True
        repoTaxRecoverable7.IsVisible = False
        repoTaxRecoverable7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable7)
        Dim repoTax8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax8.FormatString = ""
        repoTax8.HeaderText = "Tax 8"
        repoTax8.Name = colTax8
        repoTax8.ReadOnly = True
        repoTax8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax8)
        Dim repoTaxBaseAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt8.FormatString = ""
        repoTaxBaseAmt8.HeaderText = "Tax Base Amount 8"
        repoTaxBaseAmt8.Name = colTaxBaseAmt8
        repoTaxBaseAmt8.DecimalPlaces = 6
        repoTaxBaseAmt8.ReadOnly = True
        repoTaxBaseAmt8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt8)
        Dim repoTaxRate8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate8 = New GridViewDecimalColumn()
        repoTaxRate8.FormatString = ""
        repoTaxRate8.HeaderText = "Tax Rate 8"
        repoTaxRate8.Name = colTaxRate8
        repoTaxRate8.IsVisible = False
        repoTaxRate8.ReadOnly = True
        repoTaxRate8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate8)
        Dim repoTaxAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt8 = New GridViewDecimalColumn()
        repoTaxAmt8.FormatString = ""
        repoTaxAmt8.HeaderText = "Tax Amt 8"
        repoTaxAmt8.Name = colTaxAmt8
        repoTaxAmt8.IsVisible = False
        repoTaxAmt8.DecimalPlaces = 6
        repoTaxAmt8.ReadOnly = True
        repoTaxAmt8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt8)
        Dim repoIsSurTax8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax8.HeaderText = "Is Surtax 8"
        repoIsSurTax8.Name = colIsSurTax8
        repoIsSurTax8.ReadOnly = True
        repoIsSurTax8.IsVisible = False
        repoIsSurTax8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax8)
        Dim repoSurTaxCode8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode8.FormatString = ""
        repoSurTaxCode8.HeaderText = "Surtax 8"
        repoSurTaxCode8.Name = colSurTaxCode8
        repoSurTaxCode8.ReadOnly = True
        repoSurTaxCode8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode8)
        Dim repoIsTaxable8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable8.HeaderText = "Is Taxable 8"
        repoIsTaxable8.Name = colIsTaxable8
        repoIsTaxable8.ReadOnly = True
        repoIsTaxable8.IsVisible = False
        repoIsTaxable8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable8)
        Dim repoTaxRecoverable8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable8.HeaderText = "Recoverable Tax 8"
        repoTaxRecoverable8.Name = colTaxRecoverable8
        repoTaxRecoverable8.ReadOnly = True
        repoTaxRecoverable8.IsVisible = False
        repoTaxRecoverable8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable8)
        Dim repoTax9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax9.FormatString = ""
        repoTax9.HeaderText = "Tax 9"
        repoTax9.Name = colTax9
        repoTax9.ReadOnly = True
        repoTax9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax9)
        Dim repoTaxBaseAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt9.FormatString = ""
        repoTaxBaseAmt9.HeaderText = "Tax Base Amount 9"
        repoTaxBaseAmt9.Name = colTaxBaseAmt9
        repoTaxBaseAmt9.ReadOnly = True
        repoTaxBaseAmt9.DecimalPlaces = 6
        repoTaxBaseAmt9.IsVisible = False
        repoTaxBaseAmt9.DecimalPlaces = 6
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt9)
        Dim repoTaxRate9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate9 = New GridViewDecimalColumn()
        repoTaxRate9.FormatString = ""
        repoTaxRate9.HeaderText = "Tax Rate 9"
        repoTaxRate9.Name = colTaxRate9
        repoTaxRate9.IsVisible = False
        repoTaxRate9.ReadOnly = True
        repoTaxRate9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate9)
        Dim repoTaxAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt9 = New GridViewDecimalColumn()
        repoTaxAmt9.FormatString = ""
        repoTaxAmt9.HeaderText = "Tax Amt 9"
        repoTaxAmt9.Name = colTaxAmt9
        repoTaxAmt9.IsVisible = False
        repoTaxAmt9.ReadOnly = True
        repoTaxAmt9.DecimalPlaces = 6
        repoTaxAmt9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt9)
        Dim repoIsSurTax9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax9.HeaderText = "Is Surtax 9"
        repoIsSurTax9.Name = colIsSurTax9
        repoIsSurTax9.ReadOnly = True
        repoIsSurTax9.IsVisible = False
        repoIsSurTax9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax9)
        Dim repoSurTaxCode9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode9.FormatString = ""
        repoSurTaxCode9.HeaderText = "Surtax 9"
        repoSurTaxCode9.Name = colSurTaxCode9
        repoSurTaxCode9.ReadOnly = True
        repoSurTaxCode9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode9)
        Dim repoIsTaxable9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable9.HeaderText = "Is Taxable 9"
        repoIsTaxable9.Name = colIsTaxable9
        repoIsTaxable9.ReadOnly = True
        repoIsTaxable9.IsVisible = False
        repoIsTaxable9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable9)
        Dim repoTaxRecoverable9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable9.HeaderText = "Recoverable Tax 9"
        repoTaxRecoverable9.Name = colTaxRecoverable9
        repoTaxRecoverable9.ReadOnly = True
        repoTaxRecoverable9.IsVisible = False
        repoTaxRecoverable9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable9)
        Dim repoTax10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax10.FormatString = ""
        repoTax10.HeaderText = "Tax 10"
        repoTax10.Name = colTax10
        repoTax10.ReadOnly = True
        repoTax10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax10)
        Dim repoTaxBaseAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt10.FormatString = ""
        repoTaxBaseAmt10.HeaderText = "Tax Base Amount 10"
        repoTaxBaseAmt10.Name = colTaxBaseAmt10
        repoTaxBaseAmt10.ReadOnly = True
        repoTaxBaseAmt10.DecimalPlaces = 6
        repoTaxBaseAmt10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt10)
        Dim repoTaxRate10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate10 = New GridViewDecimalColumn()
        repoTaxRate10.FormatString = ""
        repoTaxRate10.HeaderText = "Tax Rate 10"
        repoTaxRate10.Name = colTaxRate10
        repoTaxRate10.IsVisible = False
        repoTaxRate10.ReadOnly = True
        repoTaxRate10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate10)
        Dim repoTaxAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt10 = New GridViewDecimalColumn()
        repoTaxAmt10.FormatString = ""
        repoTaxAmt10.HeaderText = "Tax Amt 10"
        repoTaxAmt10.Name = colTaxAmt10
        repoTaxAmt10.IsVisible = False
        repoTaxAmt10.ReadOnly = True
        repoTaxAmt10.DecimalPlaces = 6
        repoTaxAmt10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt10)
        Dim repoIsSurTax10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax10.HeaderText = "Is Surtax 10"
        repoIsSurTax10.Name = colIsSurTax10
        repoIsSurTax10.ReadOnly = True
        repoIsSurTax10.IsVisible = False
        repoIsSurTax10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSurTax10)
        Dim repoSurTaxCode10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode10.FormatString = ""
        repoSurTaxCode10.HeaderText = "Surtax 10"
        repoSurTaxCode10.Name = colSurTaxCode10
        repoSurTaxCode10.ReadOnly = True
        repoSurTaxCode10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSurTaxCode10)
        Dim repoIsTaxable10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable10.HeaderText = "Is Taxable 10"
        repoIsTaxable10.Name = colIsTaxable10
        repoIsTaxable10.ReadOnly = True
        repoIsTaxable10.IsVisible = False
        repoIsTaxable10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsTaxable10)
        For ii As Integer = 1 To 10
            Dim repoCheckBox As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
            repoCheckBox.HeaderText = "Is Tax On Base Amount " + clsCommon.myCstr(ii)
            repoCheckBox.Name = colIsTaxOnBaseAmount + clsCommon.myCstr(ii)
            repoCheckBox.ReadOnly = True
            repoCheckBox.IsVisible = False
            repoCheckBox.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
            repoCheckBox.WrapText = True
            gv1.MasterTemplate.Columns.Add(repoCheckBox)
        Next
        Dim repoTaxRecoverable10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoTaxRecoverable10.HeaderText = "Recoverable Tax 10"
        repoTaxRecoverable10.Name = colTaxRecoverable10
        repoTaxRecoverable10.ReadOnly = True
        repoTaxRecoverable10.IsVisible = False
        repoTaxRecoverable10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoTaxRecoverable10)
        Dim repoIsExcisable1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable1.HeaderText = "Is Excisable 1"
        repoIsExcisable1.Name = colIsExcisable1
        repoIsExcisable1.ReadOnly = True
        repoIsExcisable1.IsVisible = False
        repoIsExcisable1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable1)
        Dim repoIsExcisable2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable2.HeaderText = "Is Excisable 2"
        repoIsExcisable2.Name = colIsExcisable2
        repoIsExcisable2.ReadOnly = True
        repoIsExcisable2.IsVisible = False
        repoIsExcisable2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable2)
        Dim repoIsExcisable3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable3.HeaderText = "Is Excisable 3"
        repoIsExcisable3.Name = colIsExcisable3
        repoIsExcisable3.ReadOnly = True
        repoIsExcisable3.IsVisible = False
        repoIsExcisable3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable3)
        Dim repoIsExcisable4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable4.HeaderText = "Is Excisable 4"
        repoIsExcisable4.Name = colIsExcisable4
        repoIsExcisable4.ReadOnly = True
        repoIsExcisable4.IsVisible = False
        repoIsExcisable4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable4)
        Dim repoIsExcisable5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable5.HeaderText = "Is Excisable 5"
        repoIsExcisable5.Name = colIsExcisable5
        repoIsExcisable5.ReadOnly = True
        repoIsExcisable5.IsVisible = False
        repoIsExcisable5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable5)
        Dim repoIsExcisable6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable6.HeaderText = "Is Excisable 6"
        repoIsExcisable6.Name = colIsExcisable6
        repoIsExcisable6.ReadOnly = True
        repoIsExcisable6.IsVisible = False
        repoIsExcisable6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable6)
        Dim repoIsExcisable7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable7.HeaderText = "Is Excisable 7"
        repoIsExcisable7.Name = colIsExcisable7
        repoIsExcisable7.ReadOnly = True
        repoIsExcisable7.IsVisible = False
        repoIsExcisable7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable7)
        Dim repoIsExcisable8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable8.HeaderText = "Is Excisable 8"
        repoIsExcisable8.Name = colIsExcisable8
        repoIsExcisable8.ReadOnly = True
        repoIsExcisable8.IsVisible = False
        repoIsExcisable8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable8)
        Dim repoIsExcisable9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable9.HeaderText = "Is Excisable 9"
        repoIsExcisable9.Name = colIsExcisable9
        repoIsExcisable9.ReadOnly = True
        repoIsExcisable9.IsVisible = False
        repoIsExcisable9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable9)
        Dim repoIsExcisable10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsExcisable10.HeaderText = "Is Excisable 10"
        repoIsExcisable10.Name = colIsExcisable10
        repoIsExcisable10.ReadOnly = True
        repoIsExcisable10.IsVisible = False
        repoIsExcisable10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsExcisable10)
        Dim repoTotTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotTaxAmt.FormatString = ""
        repoTotTaxAmt.HeaderText = "Tax Amount"
        repoTotTaxAmt.Name = colTotTaxAmt
        repoTotTaxAmt.Width = 80
        repoTotTaxAmt.ReadOnly = True
        repoTotTaxAmt.DecimalPlaces = 6
        repoTotTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTotTaxAmt)
        Dim repoAmtAfterTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterTax.FormatString = ""
        repoAmtAfterTax.HeaderText = "Included Tax Amount"
        repoAmtAfterTax.Name = colAmtAfterTax
        repoAmtAfterTax.WrapText = True
        repoAmtAfterTax.DecimalPlaces = 6
        repoAmtAfterTax.Width = 80
        repoAmtAfterTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterTax.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterTax)
        Dim repoLandedRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLandedRate.FormatString = ""
        repoLandedRate.HeaderText = "Landed Rate"
        repoLandedRate.Name = colLandedRate
        repoLandedRate.WrapText = True
        repoLandedRate.Width = 80
        repoLandedRate.IsVisible = False
        repoLandedRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoLandedRate.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoLandedRate)
        Dim repoLandedAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLandedAmt.FormatString = ""
        repoLandedAmt.HeaderText = "Landed Amount"
        repoLandedAmt.Name = colLandedAmt
        repoLandedAmt.WrapText = True
        repoLandedAmt.Width = 80
        repoLandedAmt.IsVisible = False
        repoLandedAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoLandedAmt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoLandedAmt)
        Dim repoRequition As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRequition.FormatString = ""
        repoRequition.HeaderText = "Order No"
        repoRequition.Name = colOrderNo
        repoRequition.ReadOnly = True
        repoRequition.IsVisible = False
        repoRequition.Width = 100
        gv1.MasterTemplate.Columns.Add(repoRequition)
        Dim repoFOC As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoFOC.FormatString = ""
        repoFOC.HeaderText = "FOC"
        repoFOC.Name = ColFOC
        repoFOC.ReadOnly = True
        repoFOC.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoFOC)
        ''Dim repoAssessable As GridViewDecimalColumn = New GridViewDecimalColumn()
        ''repoAssessable.FormatString = ""
        ''repoAssessable.HeaderText = "Assessable"
        ''repoAssessable.Name = colAssessableRate
        ''repoAssessable.Width = 80
        ''repoAssessable.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        ''repoAssessable.ReadOnly = True
        ''gv1.MasterTemplate.Columns.Add(repoAssessable)
        ''Dim repoAssessableAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        ''repoAssessableAmt.WrapText = True
        ''repoAssessableAmt.ReadOnly = True
        ''repoAssessableAmt.FormatString = ""
        ''repoAssessableAmt.HeaderText = "Assessable Amount"
        ''repoAssessableAmt.Name = colAssessableAmount
        ''repoAssessableAmt.Width = 80
        ''repoAssessableAmt.Minimum = 0
        ''repoAssessableAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        ''gv1.MasterTemplate.Columns.Add(repoAssessableAmt)
        Dim repoBinNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBinNo.FormatString = ""
        repoBinNo.HeaderText = "Bin No"
        repoBinNo.Name = colBinNo
        repoBinNo.ReadOnly = False
        repoBinNo.IsVisible = False
        repoBinNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoBinNo)
        Dim repoBatchNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBatchNo.FormatString = ""
        repoBatchNo.HeaderText = "Batch No"
        repoBatchNo.Name = colBatchNo
        repoBatchNo.ReadOnly = False
        repoBatchNo.IsVisible = False
        repoBatchNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoBatchNo)
        Dim repoExpiry As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoExpiry.Format = DateTimePickerFormat.Custom
        repoExpiry.CustomFormat = "dd-MM-yyyy"
        repoExpiry.HeaderText = "Expiry Date"
        repoExpiry.FormatString = "{0:d}"
        repoExpiry.Name = colExpiry
        repoExpiry.WrapText = True
        repoExpiry.ReadOnly = False
        repoExpiry.IsVisible = False
        repoExpiry.Width = 80
        gv1.MasterTemplate.Columns.Add(repoExpiry)
        Dim repoManDate As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoManDate.Format = DateTimePickerFormat.Custom
        repoManDate.CustomFormat = "dd-MM-yyyy"
        repoManDate.HeaderText = "Manufacturer Date"
        repoManDate.WrapText = True
        repoManDate.FormatString = "{0:d}"
        repoManDate.Name = colManufactureDate
        repoManDate.ReadOnly = False
        repoManDate.IsVisible = False
        repoManDate.Width = 80
        gv1.MasterTemplate.Columns.Add(repoManDate)
        Dim repoSpecification As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSpecification.FormatString = ""
        repoSpecification.HeaderText = "Specification"
        repoSpecification.Name = colSpecification
        repoSpecification.Width = 100
        repoSpecification.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoSpecification)
        Dim repoItemTaxCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoItemTaxCode = New GridViewTextBoxColumn()
        repoItemTaxCode.FormatString = ""
        repoItemTaxCode.HeaderText = "Item Tax Code"
        repoItemTaxCode.Name = colItemwiseTaxCode
        repoItemTaxCode.Width = 100
        repoItemTaxCode.IsVisible = False
        repoItemTaxCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoItemTaxCode)
        Dim repoRemarks As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRemarks.FormatString = ""
        repoRemarks.HeaderText = "Remarks"
        repoRemarks.Name = colRemarks
        repoRemarks.Width = 100
        gv1.MasterTemplate.Columns.Add(repoRemarks)
        Dim repoMannulaAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMannulaAmt.FormatString = ""
        repoMannulaAmt.HeaderText = "Is Mannual amount"
        repoMannulaAmt.Name = colIsMannualAmt
        repoMannulaAmt.IsVisible = False
        repoMannulaAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoMannulaAmt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoMannulaAmt)
        Dim repoIsEmptyValue As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsEmptyValue.HeaderText = "Is Empty Value"
        repoIsEmptyValue.Name = colIsEmptyValue
        repoIsEmptyValue.ReadOnly = True
        repoIsEmptyValue.IsVisible = False
        repoIsEmptyValue.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsEmptyValue)
        Dim repoIsSerItem As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSerItem.HeaderText = "Is Serialize Item"
        repoIsSerItem.Name = colIsSerialseItem
        repoIsSerItem.ReadOnly = True
        repoIsSerItem.IsVisible = False
        repoIsSerItem.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsSerItem)

        Dim repoIsBatchItem As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsBatchItem.HeaderText = "Is Batch Item"
        repoIsBatchItem.Name = colIsBatchItem
        repoIsBatchItem.ReadOnly = True
        repoIsBatchItem.IsVisible = False
        repoIsBatchItem.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsBatchItem)

        clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = True
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40
        ' ReStoreGridLayout()
    End Sub
    Sub OpenSerialItem()
        If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsSerialseItem).Value) Then
            Dim frm As frmSerializeItemOut = New frmSerializeItemOut()
            frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
            frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
            frm.strLocationCode = txtBillToLocation.Value
            frm.strCurrDocNo = txtDocNo.Value
            frm.strCurrDocType = "SD-IN"
            frm.dblqty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
            frm.arr = TryCast(gv1.CurrentRow.Tag, List(Of clsSerializeInvenotry))
            frm.ShowDialog()
            If Not frm.isCencelButtonClicked Then
                gv1.CurrentRow.Tag = frm.arr
            End If
        End If
    End Sub
    Sub LoadBlankGridAC()
        gvAC.Rows.Clear()
        gvAC.Columns.Clear()
        Dim repoACCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoACCode.FormatString = ""
        repoACCode.HeaderText = "Addition Charges Code"
        repoACCode.Name = colACCode
        repoACCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoACCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoACCode.Width = 150
        repoACCode.ReadOnly = False
        gvAC.MasterTemplate.Columns.Add(repoACCode)
        Dim repoACName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoACName.FormatString = ""
        repoACName.HeaderText = "Addition Charges Description"
        repoACName.Name = colACName
        repoACName.Width = 300
        repoACName.ReadOnly = True
        gvAC.MasterTemplate.Columns.Add(repoACName)
        Dim repoACAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoACAmt.FormatString = ""
        repoACAmt.HeaderText = "Amount"
        repoACAmt.Name = colACAmount
        repoACAmt.Width = 100
        repoACAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoACAmt.ReadOnly = False
        gvAC.MasterTemplate.Columns.Add(repoACAmt)
        gvAC.AllowAddNewRow = False
        gvAC.ShowGroupPanel = False
        gvAC.AllowColumnReorder = True
        gvAC.AllowRowReorder = False
        gvAC.EnableSorting = False
        gvAC.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvAC.MasterTemplate.ShowRowHeaderColumn = False
        gvAC.TableElement.TableHeaderHeight = 40
    End Sub
    Sub LoadBlankGridTax()
        gv2.Rows.Clear()
        gv2.Columns.Clear()
        Dim repoTaxAuthCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthCode.FormatString = ""
        repoTaxAuthCode.HeaderText = "Tax Authority Code"
        repoTaxAuthCode.Name = colTTaxAutCode
        repoTaxAuthCode.Width = 150
        repoTaxAuthCode.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthCode)
        Dim repoTaxAuthName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthName.FormatString = ""
        repoTaxAuthName.HeaderText = "Tax Authority"
        repoTaxAuthName.Name = colTTaxAutName
        repoTaxAuthName.Width = 200
        repoTaxAuthName.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthName)
        Dim repoTaxBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt.FormatString = ""
        repoTaxBaseAmt.HeaderText = "Base Amount"
        repoTaxBaseAmt.Name = colTBaseAmt
        repoTaxBaseAmt.Width = 100
        repoTaxBaseAmt.ReadOnly = True
        repoTaxBaseAmt.DecimalPlaces = 6
        repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxBaseAmt)
        Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate.FormatString = ""
        repoTaxRate.HeaderText = "Tax Rate"
        repoTaxRate.Name = colTTaxRate
        repoTaxRate.Width = 100
        repoTaxRate.ReadOnly = True
        repoTaxRate.IsVisible = IIf(clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, Nothing)) = 1, True, False)
        repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxRate)
        Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt.FormatString = ""
        repoTaxAmt.HeaderText = "Tax Amount"
        repoTaxAmt.Name = colTTaxAmt
        repoTaxAmt.Width = 100
        repoTaxAmt.DecimalPlaces = 6
        repoTaxAmt.ReadOnly = False
        repoTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxAmt)
        gv1.AllowDeleteRow = True
        gv2.AllowAddNewRow = False
        gv2.ShowGroupPanel = False
        gv2.AllowColumnReorder = False
        gv2.AllowRowReorder = False
        gv2.EnableSorting = False
        gv2.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv2.MasterTemplate.ShowRowHeaderColumn = False
    End Sub
    Private Sub gv1_CellDoubleClick(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellDoubleClick
        Dim grow As GridViewRowInfo = TryCast(e.Row, GridViewRowInfo)
        If e.Column.Name = "complete" Then
            If grow.Cells(colComplete).Value = "No" Then
                grow.Cells(colComplete).Value = "Yes"
            ElseIf grow.Cells(colComplete).Value = "Yes" Then
                grow.Cells(colComplete).Value = "No"
            End If
        ElseIf e.Column.Name = colSchemeApplicable And grow.Cells(colSchemeItem).Value = "No" Then
            If grow.Cells(colSchemeApplicable).Value = "Yes" Then
                grow.Cells(colSchemeApplicable).Value = "No"
            ElseIf grow.Cells(colSchemeApplicable).Value = "No" Then
                grow.Cells(colSchemeApplicable).Value = "Yes"
            End If
        End If
    End Sub
    Private Sub gv1_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column.FieldName.StartsWith("_CFLD_") Then
                        clsCustomFieldGrid.getFinderForCustomFieldGrid(gv1, e.Column.Name.ToString, MyBase.Form_ID)
                    End If
                    If e.Column Is gv1.Columns(colTotTaxAmt) Then
                        UpdateCurrentRow(gv1.CurrentRow.Index) ''-1 is for current row
                        UpdateAllTotals()
                    End If
                    If e.Column Is gv1.Columns(colQty) Then
                        If RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                            OpenBatchItem(False)
                        End If
                    End If
                    If e.Column Is gv1.Columns(colICode) OrElse e.Column Is gv1.Columns(colDisPerUnit) OrElse e.Column Is gv1.Columns(colMRP) OrElse e.Column Is gv1.Columns(colRate) OrElse e.Column Is gv1.Columns(colHeaDDisPer) OrElse (e.Column Is gv1.Columns(colHeadDiscamt)) OrElse e.Column Is gv1.Columns(colSchemeApplicable) OrElse e.Column Is gv1.Columns(colAmt) OrElse e.Column Is gv1.Columns(colQty) OrElse e.Column Is gv1.Columns(colRate) OrElse e.Column Is gv1.Columns(colSpecification) OrElse e.Column Is gv1.Columns(colRemarks) OrElse e.Column Is gv1.Columns(colDisPer) OrElse e.Column Is gv1.Columns(colMRP) OrElse e.Column Is gv1.Columns(colBatchNo) OrElse e.Column Is gv1.Columns(colExpiry) OrElse e.Column Is gv1.Columns(colManufactureDate) OrElse e.Column Is gv1.Columns(colShortDesc) OrElse e.Column Is gv1.Columns(colUnit) OrElse (e.Column Is gv1.Columns(colAmt) AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal) Then
                        If (e.Column Is gv1.Columns(colQty) OrElse (e.Column Is gv1.Columns(colHeadDiscamt)) OrElse e.Column Is gv1.Columns(colHeaDDisPer) OrElse e.Column Is gv1.Columns(colDisPer) OrElse (e.Column Is gv1.Columns(colAmt) AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal)) Then
                            If ((e.Column Is gv1.Columns(colQty))) Then
                                Dim dblPendingQty As Decimal = 0
                                If (clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) > 0) Then
                                    dblPendingQty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colPendingQty).Value)
                                    Dim dblEnteredQty As Decimal = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
                                    Dim dblDamageQty As Decimal = 0 'clsCommon.myCDecimal(gv1.CurrentRow.Cells(colLeakQty).Value) + clsCommon.myCDecimal(gv1.CurrentRow.Cells(colBurstQty).Value) + clsCommon.myCDecimal(gv1.CurrentRow.Cells(colShortQty).Value)
                                    If (dblEnteredQty + dblDamageQty) > dblPendingQty Then
                                        common.clsCommon.MyMessageBoxShow("Entered Quantity Can't be more than Pending Quantity." + Environment.NewLine + "Entered Quantity : " + clsCommon.myCstr(dblEnteredQty) + ". Pending Quantity : " + clsCommon.myCstr(dblPendingQty))
                                        gv1.CurrentCell.Value = dblPendingQty
                                    End If
                                ElseIf clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) = 0 Then
                                    If AutoScheme Then
                                        gv1.CurrentRow.Cells(colSchemeApplicable).Value = "Yes"
                                    End If
                                End If
                                findQtyandPromoSchemeCode(False)
                                OpenSerialItem()
                                setBalance()
                            End If
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                            If rbtnTaxCalManual.IsChecked Then
                                For ii As Integer = 0 To gv1.Rows.Count - 1
                                    UpdateCurrentRow(ii)
                                Next
                            End If
                            UpdateAllTotals()
                        ElseIf e.Column Is gv1.Columns(colICode) Then
                            If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                                OpenICodeListCurrentCalaculation(False)
                                setBalance()
                            Else
                                Throw New Exception("Please select Customer first")
                            End If
                        ElseIf e.Column Is gv1.Columns(colShortDesc) Then
                            If UseDescInsteadOFCodeOnMCCMAterialSale Then
                                OpenICodeListShortDesc(False, clsCommon.myCstr(gv1.CurrentRow.Cells(colShortDesc).Value))
                                setBalance()
                            End If
                        ElseIf e.Column Is gv1.Columns(colUnit) Then
                            OpenUOMList(False)
                            UpdateCurrentRow(gv1.CurrentRow.Index) ''-1 is for current row
                            UpdateAllTotals()
                            If AutoScheme Then
                                gv1.CurrentRow.Cells(colSchemeApplicable).Value = "Yes"
                            End If
                            findQtyandPromoSchemeCode(False)
                            setBalance()
                        ElseIf e.Column Is gv1.Columns(colSchemeApplicable) Then
                            findQtyandPromoSchemeCode(False)
                        ElseIf e.Column Is gv1.Columns(colMRP) Then
                            OpenGetbalance(False)
                        ElseIf e.Column Is gv1.Columns(colRate) Then
                            gv1.CurrentRow.Cells(colOrgCost).Value = gv1.CurrentRow.Cells(colRate).Value
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                            UpdateAllTotals()
                        ElseIf e.Column Is gv1.Columns(colDisPerUnit) Then
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                            UpdateAllTotals()
                        ElseIf e.Column Is gv1.Columns(colAmt) Then
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                            If rbtnTaxCalManual.IsChecked Then
                                For ii As Integer = 0 To gv1.Rows.Count - 1
                                    UpdateCurrentRow(ii)
                                Next
                            End If
                            UpdateAllTotals()
                        End If
                    End If

                    If e.Column Is gv1.Columns(colRate) Then
                        '  SetDecimalPlaces()
                    End If
                    isCellValueChangedOpen = False
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            isCellValueChangedOpen = False
        End Try
    End Sub
    Sub OpenGetbalance(ByVal isButtonClick As Boolean)
        gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value))
    End Sub

    Private Sub findQtyandPromoSchemeCode(ByVal isButtonClick As Boolean)
        Dim dr1 As DataTable
        Dim schemeCodeCol As String
        Dim intRow As Integer
        Dim strOrderCode As String = ""
        Try
            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeApplicable).Value), "No") = CompairStringResult.Equal OrElse clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value) = 0 Then
                For schemeRow As Integer = gv1.Rows.Count - 1 To 0 Step -1
                    If clsCommon.myLen(gv1.CurrentRow.Cells(colFromSchemeCode).Value) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colFromSchemeCode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colFromSchemeCode).Value)) = CompairStringResult.Equal Then
                            If clsCommon.myCDecimal(gv1.Rows(schemeRow).Cells(colQty).Value) >= 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(ColFOC).Value), 1) = CompairStringResult.Equal Then
                                gv1.Rows.RemoveAt(schemeRow)
                            End If
                        End If
                    End If
                Next
                gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                gv1.CurrentRow.Cells(colFromSchemeCode).Value = Nothing
            ElseIf clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeApplicable).Value), "Yes") = CompairStringResult.Equal Then
                For schemeRow As Integer = gv1.Rows.Count - 1 To 0 Step -1
                    If clsCommon.myLen(gv1.CurrentRow.Cells(colFromSchemeCode).Value) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colFromSchemeCode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colFromSchemeCode).Value)) = CompairStringResult.Equal Then
                            If clsCommon.myCDecimal(gv1.Rows(schemeRow).Cells(colQty).Value) >= 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(ColFOC).Value), 1) = CompairStringResult.Equal Then
                                gv1.Rows.RemoveAt(schemeRow)
                            End If
                        End If
                    End If
                Next
                StrSql = "SELECT top 1 TSPL_SCHEME_MASTER_NEW.Scheme_Code,TSPL_SCHEME_MASTER_NEW.Item_Qty,TSPL_SCHEME_MASTER_NEW.Unit_Code " &
             "FROM TSPL_SCHEME_MASTER_NEW  left outer JOIN TSPL_SCHEME_DETAIL_NEW  ON  " &
             "TSPL_SCHEME_MASTER_NEW.Scheme_Code = TSPL_SCHEME_DETAIL_NEW.Scheme_Code  left outer join " &
             "TSPL_SCHEME_BENEFICIARY on TSPL_SCHEME_MASTER_NEW.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code " &
             "WHERE  Scheme_Type='Quantitive' and  (TSPL_SCHEME_MASTER_NEW.Item_Code = '" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "') AND TSPL_SCHEME_MASTER_NEW.Start_Date <='" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy") + "' " &
                  " AND (TSPL_SCHEME_MASTER_NEW.End_Date >='" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy") + "' OR TSPL_SCHEME_MASTER_NEW.End_Date is NULL )   and  " &
             "TSPL_SCHEME_MASTER_NEW.Item_Qty <= '" & clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value) & "'  AND " &
             "TSPL_SCHEME_MASTER_NEW.Unit_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) & "' AND " &
             "TSPL_SCHEME_MASTER_NEW.MRP='" & clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value) & "' AND TSPL_SCHEME_BENEFICIARY.Cust_Code  = '" & txtVendorNo.Value & "' order by Start_Date desc"
                dr1 = clsDBFuncationality.GetDataTable(StrSql)
                Dim discountRatio As Integer = 0
                If dr1 IsNot Nothing AndAlso dr1.Rows.Count > 0 Then
                    Dim intOldRow As Integer = gv1.CurrentRow.Index
                    schemeCodeCol = dr1.Rows(0)(0).ToString
                    Dim strPriceCode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colPriceCOde).Value)
                    Dim mainItemQty As Decimal = clsCommon.myCDecimal(dr1.Rows(0)(1).ToString())
                    Dim mainItemCode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    Dim mode As Decimal = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value) Mod clsCommon.myCDecimal(dr1.Rows(0)(1).ToString())
                    discountRatio = (clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value) - mode) / clsCommon.myCDecimal(dr1.Rows(0)(1).ToString())
                    If discountRatio > 0 Then
                        Dim dblSchemeItemActualQTy
                        StrSql = "SELECT TSPL_ITEM_MASTER.Weight_Value,TSPL_ITEM_PRICE_MASTER.Start_Date,TSPL_ITEM_PRICE_MASTER.Price_Code,TSPL_SCHEME_DETAIL_NEW.Item_Code,TSPL_ITEM_MASTER.Item_Desc, " &
                           "TSPL_SCHEME_DETAIL_NEW.Qty,TSPL_SCHEME_DETAIL_NEW.Unit_Code, TSPL_SCHEME_DETAIL_NEW.MRP,TSPL_SCHEME_DETAIL_NEW.Price_Date, " &
                           "TSPL_ITEM_PRICE_MASTER.Item_Basic_Price, TSPL_ITEM_PRICE_MASTER.Tax_group  , TSPL_ITEM_MASTER.Item_Type, TSPL_ITEM_MASTER.show,  " &
                           "TSPL_ITEM_MASTER.Sku_Seq,TAX1_Rate as Tax1Rate,TAX2_Rate as  Tax2Rate,TAX3_Rate as Tax3Rate,TAX4_Rate as Tax4Rate, " &
                           "TAX5_Rate as Tax5Rate,TAX6_Rate as Tax6Rate,TAX7_Rate as Tax7Rate,TAX8_Rate as Tax8Rate, TAX9_Rate as Tax9Rate, " &
                           "TAX10_Rate as Tax10Rate,TSPL_ITEM_PRICE_MASTER.TAX1 ,  " &
                           "TSPL_ITEM_PRICE_MASTER.TAX2,TSPL_ITEM_PRICE_MASTER.TAX3, TSPL_ITEM_PRICE_MASTER.TAX4,TSPL_ITEM_PRICE_MASTER.TAX5,  " &
                           "TSPL_ITEM_PRICE_MASTER.TAX6,TSPL_ITEM_PRICE_MASTER.TAX7, TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,  " &
                           "TSPL_ITEM_PRICE_MASTER.TAX10,abatement_rate FROM TSPL_SCHEME_MASTER_NEW left outer JOIN TSPL_SCHEME_DETAIL_NEW  " &
                           "ON  TSPL_SCHEME_MASTER_NEW.Scheme_Code = TSPL_SCHEME_DETAIL_NEW.Scheme_Code left outer join  " &
                           "TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SCHEME_DETAIL_NEW.Item_Code left outer  join tspl_item_price_master on " &
                           "tspl_item_price_master.Item_Code=TSPL_SCHEME_DETAIL_NEW.Item_Code and TSPL_ITEM_PRICE_MASTER.UOM=TSPL_SCHEME_DETAIL_NEW.Unit_Code and " &
                           "TSPL_ITEM_PRICE_MASTER.Start_Date=TSPL_SCHEME_DETAIL_NEW.Price_Date and " &
                           "TSPL_ITEM_PRICE_MASTER.Item_Basic_Net=TSPL_SCHEME_DETAIL_NEW.MRP " &
                           "WHERE Scheme_Type='Quantitive' and (TSPL_SCHEME_MASTER_NEW.Scheme_Code =  '" & schemeCodeCol & "') and TSPL_ITEM_PRICE_MASTER.Price_Code='" & strPriceCode & "'"
                        dr1 = clsDBFuncationality.GetDataTable(StrSql)
                        If dr1 IsNot Nothing AndAlso dr1.Rows.Count > 0 Then
                            Dim intParentRowIndexNo As Integer = gv1.CurrentRow.Index
                            gv1.Rows(intParentRowIndexNo).Cells(colFromSchemeCode).Value = schemeCodeCol
                            For Each tdr As DataRow In dr1.Rows
                                dblSchemeItemActualQTy = clsItemLocationDetails.getBalance(clsCommon.myCstr(tdr("Item_Code")), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(tdr("Unit_Code")), clsCommon.myCstr(tdr("MRP")))
                                intOldRow += 1
                                If clsCommon.myLen(gv1.Rows(intParentRowIndexNo).Cells(colOrderNo).Value) = 0 Then
                                    gv1.Rows.AddNew()
                                    intRow = gv1.Rows.Count - 1
                                Else
                                    strOrderCode = gv1.Rows(intParentRowIndexNo).Cells(colOrderNo).Value
                                    If clsCommon.CompairString(gv1.Rows(intParentRowIndexNo).Cells(colSchemeApplicable).Value, "Yes") = CompairStringResult.Equal Then
                                        gv1.Rows.AddNew()
                                        intRow = gv1.Rows.Count - 1
                                    End If
                                End If
                                gv1.Rows(intRow).Cells(colRowType).Value = RowTypeItem
                                gv1.Rows(intRow).Cells(colLineNo).Value = intRow + 1
                                gv1.Rows(intRow).Cells(colOrderNo).Value = strOrderCode
                                gv1.Rows(intRow).Cells(colICode).Value = clsCommon.myCstr(tdr("Item_Code"))
                                gv1.Rows(intRow).Cells(colIName).Value = clsCommon.myCstr(tdr("Item_Desc"))
                                gv1.Rows(intRow).Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(tdr("Item_Code")), Nothing)
                                gv1.Rows(intRow).Cells(colPriceCOde).Value = strPriceCode
                                gv1.Rows(intRow).Cells(colPriceDateColumn).Value = clsCommon.myCstr(tdr("Start_Date"))
                                gv1.Rows(intRow).Cells(colUnit).Value = clsCommon.myCstr(tdr("Unit_Code"))
                                gv1.Rows(intRow).Cells(colUnit).Value = clsDBFuncationality.getSingleValue("select Unit_Desc from tspl_unit_master where Unit_code= '" + clsCommon.myCstr(tdr("Unit_Code")) + "'")
                                gv1.Rows(intRow).Cells(colMRP).Value = clsCommon.myCDecimal(tdr("MRP"))
                                gv1.Rows(intRow).Cells(colRate).Value = clsCommon.myCDecimal(tdr("Item_Basic_Price"))
                                gv1.Rows(intRow).Cells(colQty).Value = clsCommon.myCDecimal(tdr("Qty")) * discountRatio
                                gv1.Rows(intRow).Cells(colItemWeight).Value = clsCommon.myCDecimal(tdr("Weight_Value"))
                                gv1.Rows(intRow).Cells(colSchemeApplicable).Value = "Yes"
                                gv1.Rows(intRow).Cells(colSchemeItem).Value = "Yes"
                                gv1.Rows(intRow).Cells(colFromSchemeCode).Value = schemeCodeCol
                                gv1.Rows(intRow).Cells(ColFOC).Value = 1
                                gv1.Rows(intRow).Cells(colAbatementPer).Value = clsCommon.myCDecimal(tdr("abatement_rate"))
                                gv1.Rows(intRow).Cells(colActualCost).Value = clsCommon.myCDecimal(tdr("Item_Basic_Price"))
                                gv1.Rows(intRow).Cells(colTaxRate1).Value = clsCommon.myCDecimal(tdr("Tax1Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate2).Value = clsCommon.myCDecimal(tdr("Tax2Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate3).Value = clsCommon.myCDecimal(tdr("Tax3Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate4).Value = clsCommon.myCDecimal(tdr("Tax4Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate5).Value = clsCommon.myCDecimal(tdr("Tax5Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate6).Value = clsCommon.myCDecimal(tdr("Tax6Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate7).Value = clsCommon.myCDecimal(tdr("Tax7Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate8).Value = clsCommon.myCDecimal(tdr("Tax8Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate9).Value = clsCommon.myCDecimal(tdr("Tax9Rate"))
                                gv1.Rows(intRow).Cells(colTaxRate10).Value = clsCommon.myCDecimal(tdr("Tax10Rate"))
                                gv1.Rows(intRow).Cells(colTax1).Value = clsCommon.myCstr(tdr("Tax1"))
                                gv1.Rows(intRow).Cells(colTax2).Value = clsCommon.myCstr(tdr("Tax2"))
                                gv1.Rows(intRow).Cells(colTax3).Value = clsCommon.myCstr(tdr("Tax3"))
                                gv1.Rows(intRow).Cells(colTax4).Value = clsCommon.myCstr(tdr("Tax4"))
                                gv1.Rows(intRow).Cells(colTax5).Value = clsCommon.myCstr(tdr("Tax5"))
                                gv1.Rows(intRow).Cells(colTax6).Value = clsCommon.myCstr(tdr("Tax6"))
                                gv1.Rows(intRow).Cells(colTax7).Value = clsCommon.myCstr(tdr("Tax7"))
                                gv1.Rows(intRow).Cells(colTax8).Value = clsCommon.myCstr(tdr("Tax8"))
                                gv1.Rows(intRow).Cells(colTax9).Value = clsCommon.myCstr(tdr("Tax9"))
                                gv1.Rows(intRow).Cells(colTax10).Value = clsCommon.myCstr(tdr("Tax10"))
                                Dim dblConvF As Decimal
                                dblConvF = GetConvFactor(gv1.Rows(intRow).Cells(colUnit).Value, gv1.Rows(intRow).Cells(colICode).Value)
                                gv1.Rows(intRow).Cells(colConvF).Value = dblConvF
                                gv1.Rows(intRow).Cells(colTotItemWt).Value = dblConvF * clsCommon.myCDecimal(gv1.Rows(intRow).Cells(colItemWeight).Value) * clsCommon.myCDecimal(gv1.Rows(intRow).Cells(colQty).Value)
                                gv1.Rows(intRow).Cells(ColActualBalQty).Value = dblSchemeItemActualQTy
                                If intRow <> intOldRow Then
                                    gv1.Rows.Move(intRow, intOldRow)
                                End If
                                UpdateCurrentRow(gv1.Rows(intRow).Index)
                                UpdateAllTotals()
                            Next
                        End If
                        If dr1.Rows.Count = 0 Then
                            gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                            gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                            gv1.CurrentRow.Cells(colFromSchemeCode).Value = Nothing
                        End If
                    End If
                Else
                    For schemeRow As Integer = gv1.Rows.Count - 1 To 0 Step -1
                        If clsCommon.myLen(gv1.CurrentRow.Cells(colFromSchemeCode).Value) > 0 Then
                            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colFromSchemeCode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colFromSchemeCode).Value)) = CompairStringResult.Equal Then
                                If clsCommon.myCDecimal(gv1.Rows(schemeRow).Cells(colQty).Value) >= 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(ColFOC).Value), 1) = CompairStringResult.Equal Then
                                    gv1.Rows.RemoveAt(schemeRow)
                                End If
                            End If
                        End If
                    Next
                    gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                    gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                    gv1.CurrentRow.Cells(colFromSchemeCode).Value = Nothing
                End If
            End If
            RefreshSerialNo()
        Catch ex As Exception
            myMessages.myExceptions(ex)
            gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
            gv1.CurrentRow.Cells(colFromSchemeCode).Value = String.Empty
            Exit Sub
        End Try
    End Sub
    Sub RefreshSerialNo()
        Dim intSerialNo As Integer
        For intCount As Integer = 0 To gv1.Rows.Count - 1
            intSerialNo += 1
            gv1.Rows(intCount).Cells(colLineNo).Value = intSerialNo
        Next
    End Sub
    Sub OpenUOMList(ByVal isButtonClick As Boolean)
        Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        If clsCommon.myLen(strICode) > 0 Then
            Dim qry As String = "select UOM_Code as Code,UOM_Description as [Description] from TSPL_ITEM_UOM_DETAIL"
            Dim whrCls As String = "Item_Code='" + strICode + "'"
            gv1.CurrentRow.Cells(colUnit).Value = clsCommon.ShowSelectForm("ShipmentItefndnder", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), "Code", isButtonClick)
            If ApplySaleRouteInDCSSale Then
                gv1.CurrentRow.Cells(colRate).Value = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), txtDate.Value, Nothing, FndPriceCode.Value)
            Else
                gv1.CurrentRow.Cells(colRate).Value = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), txtDate.Value) 'clsCommon.myCdbl(dr("BasicRate"))
            End If
        End If
    End Sub
    Private Sub setGridFocus()
        Try
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCDecimal(intCurrRow + 1)
            If intCurrRow = gv1.Rows.Count - 1 Then
                gv1.Rows.AddNew()
                gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                gv1.CurrentRow = gv1.Rows(intCurrRow)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub OpenICodeList(ByVal isButtonClick As Boolean)
        gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
        Dim strItemType As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value)
        If clsCommon.myLen(strItemType) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Please select Row Type", Me.Text)
            Exit Sub
        End If
        '------------------07/08/2014---------------n-level pivot cat.---------------------
        Dim pivotheader As String = ""
        pivotheader = clsItemMaster.GetNLevelPivotHeader()
        '------------------------------------------------------------------------
        Dim strPriceGrp As String = ""
        Dim strPriceCondition As String = ""
        Dim strPriceGrpStatus As String = ""
        strPriceGrp = "''"
        strPriceGrpStatus = "''"
        If clsCommon.myLen(txtPriceCode.Text) > 0 Then
            strPriceGrp = "''"
            strPriceGrpStatus = "''"
            strPriceCondition = " and Price_Code='" & txtPriceCode.Text & "'"
        End If
        If clsCommon.CompairString(strItemType, RowTypeItem) = CompairStringResult.Equal Then
            Dim repID As String = ""
            Dim qry As String
            Dim strTaxRate As String = ""
            Dim dr As DataRow = Nothing
            Dim COUNT As Integer = 0
            For Each grow As GridViewRowInfo In gv2.Rows
                COUNT += 1
                Dim Strii As String = clsCommon.myCstr(COUNT)
                If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colTTaxAutCode).Value)) > 0 Then
                    Dim intTaxRate As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("SELECT 1 FROM TSPL_LOCATION_SALE_TAX_MASTER WHERE Location_Code='" & txtBillToLocation.Value & "' AND Tax_Group_Code='" & txtTaxGroup.Value & "' AND Tax_Code='" & grow.Cells(colTTaxAutCode).Value & "'"))
                    If intTaxRate = 1 Then
                        If clsCommon.myLen(strTaxRate) > 0 Then
                            strTaxRate += " Or "
                        End If
                        strTaxRate += " exists (select 1 from TSPL_LOCATION_SALE_TAX_MASTER where Location_Code='" & txtBillToLocation.Value & "' and Tax_group=xxx.TAX" & Strii & " and TAX_Rate=xxx.TAX" & Strii & "_Rate  ) "
                    End If
                End If
            Next
            If clsCommon.myLen(strTaxRate) > 0 Then
                strTaxRate = " and ( " & strTaxRate & " )"
            End If
            If clsCommon.myLen(pivotheader) <= 0 Then
                If PurchaseOneItemOneVendor = True Then
                    repID = "ShipPurOneItm1"
                    qry = "SELECT Item_Code as Item,Item_Desc as ItemDesc,Customer_item_no as [Customer Item Code],PrincipleCode,PrincipleDesc,vendor_code as [Vendor code],vendor_desc as [Vendor Desc],Price_Code,Start_Date AS Start_Date,UOM as Unit,Item_Basic_Net as MRP,abatement_rate, " &
                    "Item_Basic_Price as BasicRate,ITF_CODE as [ITF CODE], " &
             "Weight_Value,markup_on,markup_percent,landing_cost,Purchase_Cost,TAX1_Rate as Tax1Rate,TAX2_Rate as  Tax2Rate,TAX3_Rate as Tax3Rate,TAX4_Rate as Tax4Rate, " &
             "TAX5_Rate as Tax5Rate,TAX6_Rate as Tax6Rate,TAX7_Rate as Tax7Rate,TAX8_Rate as Tax8Rate, " &
             "TAX9_Rate as Tax9Rate,TAX10_Rate as Tax10Rate,TAX1 ,TAX2,TAX3,TAX4,TAX5,TAX6,TAX7, TAX8,TAX9,TAX10  " &
             "FROM ( SELECT distinct TSPL_VENDOR_ITEM_DETAIL.vendor_code,TSPL_VENDOR_ITEM_DETAIL.vendor_desc,Customer_item_no,TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE as PrincipleCode,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as PrincipleDesc," & strPriceGrpStatus & " as PriceGroupStatus," & strPriceGrp & " as priceGroup,TSPL_ITEM_PRICE_MASTER.Purchase_Cost,TSPL_ITEM_MASTER.Weight_Value,TSPL_ITEM_PRICE_MASTER.abatement_rate,TSPL_ITEM_MASTER.Item_Code, TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.ITF_CODE,  " &
             "CONVERT(varchar(10), TSPL_ITEM_PRICE_MASTER.Start_Date, 103) AS Start_Date, TSPL_ITEM_PRICE_MASTER.UOM, " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code, TSPL_ITEM_PRICE_MASTER.Item_Basic_Net,TSPL_ITEM_MASTER.Batch_No , TSPL_ITEM_PRICE_MASTER.Tax_group  , " &
             "TSPL_ITEM_PRICE_MASTER.Item_Basic_Price, markup_on,markup_percent,landing_cost, " &
             "TSPL_ITEM_MASTER.Item_Type, TSPL_ITEM_MASTER.show, TSPL_ITEM_MASTER.Sku_Seq, TSPL_ITEM_PRICE_MASTER.TAX1_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX6_Rate,TSPL_ITEM_PRICE_MASTER.TAX7_Rate,TSPL_ITEM_PRICE_MASTER.TAX8_Rate,TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX10_Rate,TSPL_ITEM_PRICE_MASTER.TAX1 ,TSPL_ITEM_PRICE_MASTER.TAX2,TSPL_ITEM_PRICE_MASTER.TAX3, " &
             "TSPL_ITEM_PRICE_MASTER.TAX4,TSPL_ITEM_PRICE_MASTER.TAX5,TSPL_ITEM_PRICE_MASTER.TAX6,TSPL_ITEM_PRICE_MASTER.TAX7, " &
             "TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10   FROM TSPL_ITEM_PRICE_MASTER " &
             "INNER Join  (SELECT     Item_Code, UOM, MAX(Start_Date) AS MaxDateTime, Item_Basic_Net,  Price_Code, Tax_group  FROM  " &
             "TSPL_ITEM_PRICE_MASTER  where Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND (End_Date  >='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' OR End_Date IS NULL) GROUP BY Item_Code, UOM, Item_Basic_Net,  Price_Code, Tax_group   " &
             ")  AS groupedP  ON TSPL_ITEM_PRICE_MASTER.Item_Code = groupedP.Item_Code AND  TSPL_ITEM_PRICE_MASTER.Start_Date = groupedP.MaxDateTime AND " &
             "TSPL_ITEM_PRICE_MASTER.UOM = groupedP.UOM AND TSPL_ITEM_PRICE_MASTER.Item_Basic_Net = groupedP.Item_Basic_Net  AND  " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code = groupedP.Price_Code and TSPL_ITEM_PRICE_MASTER.Tax_group = groupedP.Tax_group  left outer join TSPL_PRICE_GROUP_MAPPING on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_PRICE_GROUP_MAPPING.price_code " &
             "INNER JOIN TSPL_ITEM_MASTER AS TSPL_ITEM_MASTER ON TSPL_ITEM_PRICE_MASTER.Item_Code = TSPL_ITEM_MASTER.Item_Code " &
             "left outer join TSPL_VENDOR_ITEM_DETAIL on TSPL_VENDOR_ITEM_DETAIL.item_no=TSPL_ITEM_MASTER.Item_Code left  outer join " &
             "TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_ITEM_MASTER_CATEGORY.SNO=1 left outer join " &
             "TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " &
             "left outer join TSPL_CUSTOMER_ITEM_MAPPING on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_MAPPING.item_no and TSPL_CUSTOMER_ITEM_MAPPING.Customer_Code='" & txtVendorNo.Value & "' " &
             " where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI') )xxx Where  Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             " " & strTaxRate & " " &
             "Order By Item_Code,Start_Date,UOM desc"
                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                Else
                    repID = "ShipPurOneItm2"
                    qry = "SELECT Item_Code as Item,Item_Desc as ItemDesc,Customer_item_no as [Customer Item Code],PrincipleCode,PrincipleDesc,Price_Code,Start_Date AS Start_Date,UOM as Unit,Item_Basic_Net as MRP,abatement_rate,Item_Basic_Price as BasicRate,ITF_CODE as [ITF CODE], " &
             "Weight_Value,markup_on,markup_percent,landing_cost,Purchase_Cost,TAX1_Rate as Tax1Rate,TAX2_Rate as  Tax2Rate,TAX3_Rate as Tax3Rate,TAX4_Rate as Tax4Rate, " &
             "TAX5_Rate as Tax5Rate,TAX6_Rate as Tax6Rate,TAX7_Rate as Tax7Rate,TAX8_Rate as Tax8Rate, " &
             "TAX9_Rate as Tax9Rate,TAX10_Rate as Tax10Rate,TAX1 ,TAX2,TAX3,TAX4,TAX5,TAX6,TAX7, TAX8,TAX9,TAX10  " &
             "FROM ( SELECT distinct Customer_item_no,TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE as PrincipleCode,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as PrincipleDesc," & strPriceGrpStatus & " as PriceGroupStatus," & strPriceGrp & " as priceGroup,TSPL_ITEM_PRICE_MASTER.Purchase_Cost,TSPL_ITEM_MASTER.Weight_Value,TSPL_ITEM_PRICE_MASTER.abatement_rate,TSPL_ITEM_MASTER.Item_Code, TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.ITF_CODE,  " &
             "CONVERT(varchar(10), TSPL_ITEM_PRICE_MASTER.Start_Date, 103) AS Start_Date, TSPL_ITEM_PRICE_MASTER.UOM, " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code, TSPL_ITEM_PRICE_MASTER.Item_Basic_Net,TSPL_ITEM_MASTER.Batch_No , TSPL_ITEM_PRICE_MASTER.Tax_group  , " &
             "TSPL_ITEM_PRICE_MASTER.Item_Basic_Price, markup_on,markup_percent,landing_cost, " &
             "TSPL_ITEM_MASTER.Item_Type, TSPL_ITEM_MASTER.show, TSPL_ITEM_MASTER.Sku_Seq, TSPL_ITEM_PRICE_MASTER.TAX1_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX6_Rate,TSPL_ITEM_PRICE_MASTER.TAX7_Rate,TSPL_ITEM_PRICE_MASTER.TAX8_Rate,TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX10_Rate,TSPL_ITEM_PRICE_MASTER.TAX1 ,TSPL_ITEM_PRICE_MASTER.TAX2,TSPL_ITEM_PRICE_MASTER.TAX3, " &
             "TSPL_ITEM_PRICE_MASTER.TAX4,TSPL_ITEM_PRICE_MASTER.TAX5,TSPL_ITEM_PRICE_MASTER.TAX6,TSPL_ITEM_PRICE_MASTER.TAX7, " &
             "TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10   FROM TSPL_ITEM_PRICE_MASTER " &
             "INNER Join  (SELECT     Item_Code, UOM, MAX(Start_Date) AS MaxDateTime, Item_Basic_Net,  Price_Code, Tax_group  FROM  " &
             "TSPL_ITEM_PRICE_MASTER  where Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND (End_Date  >='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' OR End_Date IS NULL) GROUP BY Item_Code, UOM, Item_Basic_Net,  Price_Code, Tax_group   " &
             ")  AS groupedP  ON TSPL_ITEM_PRICE_MASTER.Item_Code = groupedP.Item_Code AND  TSPL_ITEM_PRICE_MASTER.Start_Date = groupedP.MaxDateTime AND " &
             "TSPL_ITEM_PRICE_MASTER.UOM = groupedP.UOM AND TSPL_ITEM_PRICE_MASTER.Item_Basic_Net = groupedP.Item_Basic_Net  AND  " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code = groupedP.Price_Code and TSPL_ITEM_PRICE_MASTER.Tax_group = groupedP.Tax_group  left outer join TSPL_PRICE_GROUP_MAPPING on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_PRICE_GROUP_MAPPING.price_code " &
             "INNER JOIN TSPL_ITEM_MASTER AS TSPL_ITEM_MASTER ON TSPL_ITEM_PRICE_MASTER.Item_Code = TSPL_ITEM_MASTER.Item_Code " &
             "left  outer join " &
             "TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_ITEM_MASTER_CATEGORY.SNO=1 left outer join " &
             "TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " &
             "left outer join TSPL_CUSTOMER_ITEM_MAPPING on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_MAPPING.item_no and TSPL_CUSTOMER_ITEM_MAPPING.Customer_Code='" & txtVendorNo.Value & "' " &
             " where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI'))xxx Where  Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             " " & strTaxRate & " " &
             "Order By Item_Code,Start_Date,UOM desc"
                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                End If
            End If
            If clsCommon.myLen(pivotheader) > 0 Then
                dr = Nothing
                If PurchaseOneItemOneVendor = True Then
                    repID = "ShipPurOneItm3"
                    qry = "select * from (select a.DESCRIPTION,a.cat_value,b.* from (SELECT Item_Code as Item,Item_Desc as ItemDesc,Customer_item_no as [Customer Item Code],PrincipleCode,PrincipleDesc,vendor_code as [Vendor code],vendor_desc as [Vendor Desc],Price_Code,Start_Date AS Start_Date,UOM as Unit,Item_Basic_Net as MRP,abatement_rate, " &
                    "Item_Basic_Price as BasicRate,ITF_CODE as [ITF CODE], " &
             "Weight_Value,markup_on,markup_percent,landing_cost,Purchase_Cost,TAX1_Rate as Tax1Rate,TAX2_Rate as  Tax2Rate,TAX3_Rate as Tax3Rate,TAX4_Rate as Tax4Rate, " &
             "TAX5_Rate as Tax5Rate,TAX6_Rate as Tax6Rate,TAX7_Rate as Tax7Rate,TAX8_Rate as Tax8Rate, " &
             "TAX9_Rate as Tax9Rate,TAX10_Rate as Tax10Rate,TAX1 ,TAX2,TAX3,TAX4,TAX5,TAX6,TAX7, TAX8,TAX9,TAX10  " &
             "FROM ( SELECT distinct TSPL_VENDOR_ITEM_DETAIL.vendor_code,TSPL_VENDOR_ITEM_DETAIL.vendor_desc,Customer_item_no,TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE as PrincipleCode,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as PrincipleDesc," & strPriceGrpStatus & " as PriceGroupStatus," & strPriceGrp & " as priceGroup,TSPL_ITEM_PRICE_MASTER.Purchase_Cost,TSPL_ITEM_MASTER.Weight_Value,TSPL_ITEM_PRICE_MASTER.abatement_rate,TSPL_ITEM_MASTER.Item_Code, TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.ITF_CODE,  " &
             "CONVERT(varchar(10), TSPL_ITEM_PRICE_MASTER.Start_Date, 103) AS Start_Date, TSPL_ITEM_PRICE_MASTER.UOM, " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code, TSPL_ITEM_PRICE_MASTER.Item_Basic_Net,TSPL_ITEM_MASTER.Batch_No , TSPL_ITEM_PRICE_MASTER.Tax_group  , " &
             "TSPL_ITEM_PRICE_MASTER.Item_Basic_Price, markup_on,markup_percent,landing_cost, " &
             "TSPL_ITEM_MASTER.Item_Type, TSPL_ITEM_MASTER.show, TSPL_ITEM_MASTER.Sku_Seq, TSPL_ITEM_PRICE_MASTER.TAX1_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX6_Rate,TSPL_ITEM_PRICE_MASTER.TAX7_Rate,TSPL_ITEM_PRICE_MASTER.TAX8_Rate,TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX10_Rate,TSPL_ITEM_PRICE_MASTER.TAX1 ,TSPL_ITEM_PRICE_MASTER.TAX2,TSPL_ITEM_PRICE_MASTER.TAX3, " &
             "TSPL_ITEM_PRICE_MASTER.TAX4,TSPL_ITEM_PRICE_MASTER.TAX5,TSPL_ITEM_PRICE_MASTER.TAX6,TSPL_ITEM_PRICE_MASTER.TAX7, " &
             "TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10   FROM TSPL_ITEM_PRICE_MASTER " &
             "INNER Join  (SELECT     Item_Code, UOM, MAX(Start_Date) AS MaxDateTime, Item_Basic_Net,  Price_Code, Tax_group  FROM  " &
             "TSPL_ITEM_PRICE_MASTER  where Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND (End_Date  >='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' OR End_Date IS NULL) GROUP BY Item_Code, UOM, Item_Basic_Net,  Price_Code, Tax_group   " &
             ")  AS groupedP  ON TSPL_ITEM_PRICE_MASTER.Item_Code = groupedP.Item_Code AND  TSPL_ITEM_PRICE_MASTER.Start_Date = groupedP.MaxDateTime AND " &
             "TSPL_ITEM_PRICE_MASTER.UOM = groupedP.UOM AND TSPL_ITEM_PRICE_MASTER.Item_Basic_Net = groupedP.Item_Basic_Net  AND  " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code = groupedP.Price_Code and TSPL_ITEM_PRICE_MASTER.Tax_group = groupedP.Tax_group  left outer join TSPL_PRICE_GROUP_MAPPING on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_PRICE_GROUP_MAPPING.price_code " &
             "INNER JOIN TSPL_ITEM_MASTER AS TSPL_ITEM_MASTER ON TSPL_ITEM_PRICE_MASTER.Item_Code = TSPL_ITEM_MASTER.Item_Code " &
             "left outer join TSPL_VENDOR_ITEM_DETAIL on TSPL_VENDOR_ITEM_DETAIL.item_no=TSPL_ITEM_MASTER.Item_Code left  outer join " &
             "TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_ITEM_MASTER_CATEGORY.SNO=1 left outer join " &
             "TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " &
             "left outer join TSPL_CUSTOMER_ITEM_MAPPING on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_MAPPING.item_no and TSPL_CUSTOMER_ITEM_MAPPING.Customer_Code='" & txtVendorNo.Value & "' " &
             " where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI') )xxx Where  Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             " " & strTaxRate & " " &
             ")b left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code,TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a on a.Item_code=b.Item) as s pivot(max(cat_value) for description in (" + pivotheader + "))t"
                    '"Order By Item_Code,Start_Date,UOM desc"
                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                Else
                    repID = "ShipPurOneItm4"
                    qry = "select * from (select a.DESCRIPTION,a.cat_value,b.* from (SELECT Item_Code as Item,Item_Desc as ItemDesc,Customer_item_no as [Customer Item Code],PrincipleCode,PrincipleDesc,Price_Code,Start_Date AS Start_Date,UOM as Unit,Item_Basic_Net as MRP,abatement_rate,Item_Basic_Price as BasicRate,ITF_CODE as [ITF CODE], " &
             "Weight_Value,markup_on,markup_percent,landing_cost,Purchase_Cost,TAX1_Rate as Tax1Rate,TAX2_Rate as  Tax2Rate,TAX3_Rate as Tax3Rate,TAX4_Rate as Tax4Rate, " &
             "TAX5_Rate as Tax5Rate,TAX6_Rate as Tax6Rate,TAX7_Rate as Tax7Rate,TAX8_Rate as Tax8Rate, " &
             "TAX9_Rate as Tax9Rate,TAX10_Rate as Tax10Rate,TAX1 ,TAX2,TAX3,TAX4,TAX5,TAX6,TAX7, TAX8,TAX9,TAX10  " &
             "FROM ( SELECT distinct Customer_item_no,TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE as PrincipleCode,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as PrincipleDesc," & strPriceGrpStatus & " as PriceGroupStatus," & strPriceGrp & " as priceGroup,TSPL_ITEM_PRICE_MASTER.Purchase_Cost,TSPL_ITEM_MASTER.Weight_Value,TSPL_ITEM_PRICE_MASTER.abatement_rate,TSPL_ITEM_MASTER.Item_Code, TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.ITF_CODE,  " &
             "CONVERT(varchar(10), TSPL_ITEM_PRICE_MASTER.Start_Date, 103) AS Start_Date, TSPL_ITEM_PRICE_MASTER.UOM, " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code, TSPL_ITEM_PRICE_MASTER.Item_Basic_Net,TSPL_ITEM_MASTER.Batch_No , TSPL_ITEM_PRICE_MASTER.Tax_group  , " &
             "TSPL_ITEM_PRICE_MASTER.Item_Basic_Price, markup_on,markup_percent,landing_cost, " &
             "TSPL_ITEM_MASTER.Item_Type, TSPL_ITEM_MASTER.show, TSPL_ITEM_MASTER.Sku_Seq, TSPL_ITEM_PRICE_MASTER.TAX1_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX6_Rate,TSPL_ITEM_PRICE_MASTER.TAX7_Rate,TSPL_ITEM_PRICE_MASTER.TAX8_Rate,TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
             "TSPL_ITEM_PRICE_MASTER.TAX10_Rate,TSPL_ITEM_PRICE_MASTER.TAX1 ,TSPL_ITEM_PRICE_MASTER.TAX2,TSPL_ITEM_PRICE_MASTER.TAX3, " &
             "TSPL_ITEM_PRICE_MASTER.TAX4,TSPL_ITEM_PRICE_MASTER.TAX5,TSPL_ITEM_PRICE_MASTER.TAX6,TSPL_ITEM_PRICE_MASTER.TAX7, " &
             "TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10   FROM TSPL_ITEM_PRICE_MASTER " &
             "INNER Join  (SELECT     Item_Code, UOM, MAX(Start_Date) AS MaxDateTime, Item_Basic_Net,  Price_Code, Tax_group  FROM  " &
             "TSPL_ITEM_PRICE_MASTER  where Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND (End_Date  >='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' OR End_Date IS NULL) GROUP BY Item_Code, UOM, Item_Basic_Net,  Price_Code, Tax_group   " &
             ")  AS groupedP  ON TSPL_ITEM_PRICE_MASTER.Item_Code = groupedP.Item_Code AND  TSPL_ITEM_PRICE_MASTER.Start_Date = groupedP.MaxDateTime AND " &
             "TSPL_ITEM_PRICE_MASTER.UOM = groupedP.UOM AND TSPL_ITEM_PRICE_MASTER.Item_Basic_Net = groupedP.Item_Basic_Net  AND  " &
             "TSPL_ITEM_PRICE_MASTER.Price_Code = groupedP.Price_Code and TSPL_ITEM_PRICE_MASTER.Tax_group = groupedP.Tax_group  left outer join TSPL_PRICE_GROUP_MAPPING on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_PRICE_GROUP_MAPPING.price_code " &
             "INNER JOIN TSPL_ITEM_MASTER AS TSPL_ITEM_MASTER ON TSPL_ITEM_PRICE_MASTER.Item_Code = TSPL_ITEM_MASTER.Item_Code " &
             "left  outer join " &
             "TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_ITEM_MASTER_CATEGORY.SNO=1 left outer join " &
             "TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " &
             "left outer join TSPL_CUSTOMER_ITEM_MAPPING on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_MAPPING.item_no and TSPL_CUSTOMER_ITEM_MAPPING.Customer_Code='" & txtVendorNo.Value & "' " &
             " where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI'))xxx Where Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             " " & strTaxRate & " " &
             ")b left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code,TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a on a.Item_code=b.Item) as s pivot(max(cat_value) for description in (" + pivotheader + "))t"
                    '"Order By Item_Code,Start_Date,UOM desc"
                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                End If
            End If
            If Not dr Is Nothing Then
                If clsCommon.myLen(clsCommon.myCstr(dr("Item"))) > 0 Then
                    SetitemWiseTaxSetting(True, True)
                    gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
                    gv1.CurrentRow.Cells(colICode).Value = clsCommon.myCstr(dr("Item"))
                    gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dr("Unit")), clsCommon.myCDecimal(dr("MRP")))
                    gv1.CurrentRow.Cells(colIName).Value = clsCommon.myCstr(dr("ItemDesc"))
                    gv1.CurrentRow.Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dr("Item")), Nothing)
                    gv1.CurrentRow.Cells(colPriceCOde).Value = clsCommon.myCstr(dr("Price_Code"))
                    gv1.CurrentRow.Cells(colPriceDateColumn).Value = clsCommon.myCstr(dr("Start_Date"))
                    gv1.CurrentRow.Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                    gv1.CurrentRow.Cells(colUOMName).Value = clsDBFuncationality.getSingleValue("select Unit_Desc from tspl_unit_master where Unit_code= '" + clsCommon.myCstr(dr("Unit")) + "'")
                    gv1.CurrentRow.Cells(colOrgUnit).Value = clsCommon.myCstr(dr("Unit"))
                    gv1.CurrentRow.Cells(colMRP).Value = clsCommon.myCDecimal(dr("MRP"))
                    gv1.CurrentRow.Cells(colRate).Value = clsCommon.myCDecimal(dr("BasicRate"))
                    gv1.CurrentRow.Cells(colAbatementPer).Value = clsCommon.myCDecimal(dr("abatement_rate"))
                    gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                    gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                    gv1.CurrentRow.Cells(colActualCost).Value = clsCommon.myCDecimal(dr("BasicRate"))
                    gv1.CurrentRow.Cells(colTaxRate1).Value = clsCommon.myCDecimal(dr("Tax1Rate"))
                    gv1.CurrentRow.Cells(colTaxRate2).Value = clsCommon.myCDecimal(dr("Tax2Rate"))
                    gv1.CurrentRow.Cells(colTaxRate3).Value = clsCommon.myCDecimal(dr("Tax3Rate"))
                    gv1.CurrentRow.Cells(colTaxRate4).Value = clsCommon.myCDecimal(dr("Tax4Rate"))
                    gv1.CurrentRow.Cells(colTaxRate5).Value = clsCommon.myCDecimal(dr("Tax5Rate"))
                    gv1.CurrentRow.Cells(colTaxRate6).Value = clsCommon.myCDecimal(dr("Tax6Rate"))
                    gv1.CurrentRow.Cells(colTaxRate7).Value = clsCommon.myCDecimal(dr("Tax7Rate"))
                    gv1.CurrentRow.Cells(colTaxRate8).Value = clsCommon.myCDecimal(dr("Tax8Rate"))
                    gv1.CurrentRow.Cells(colTaxRate9).Value = clsCommon.myCDecimal(dr("Tax9Rate"))
                    gv1.CurrentRow.Cells(colTaxRate10).Value = clsCommon.myCDecimal(dr("Tax10Rate"))
                    gv1.CurrentRow.Cells(ColFOC).Value = 0
                    gv1.CurrentRow.Cells(colItemWeight).Value = clsCommon.myCDecimal(dr("Weight_Value"))
                    Dim dblConvF As Decimal
                    dblConvF = GetConvFactor(gv1.CurrentRow.Cells(colUnit).Value, gv1.CurrentRow.Cells(colICode).Value)
                    gv1.CurrentRow.Cells(colConvF).Value = dblConvF
                    gv1.CurrentRow.Cells(colMarkupOn).Value = clsCommon.myCstr(dr("markup_on"))
                    gv1.CurrentRow.Cells(colMarkUpPercentage).Value = clsCommon.myCDecimal(dr("markup_percent"))
                    gv1.CurrentRow.Cells(colLandingCost).Value = clsCommon.myCDecimal(dr("landing_cost"))
                    gv1.CurrentRow.Cells(colPurCost).Value = clsCommon.myCDecimal(dr("Purchase_Cost"))
                    gv1.CurrentRow.Cells(colOrgCost).Value = clsCommon.myCDecimal(dr("BasicRate"))
                End If
            Else
                SetBlankOfItemColumns()
            End If
        End If
        setBalance()
    End Sub
    Sub OpenICodeListShortDesc(ByVal isButtonClick As Boolean, ByVal strShortDesc As String)
        gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
        Dim strItemType As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value)
        Dim pivotheader As String = ""
        pivotheader = clsItemMaster.GetNLevelPivotHeader()
        '----------------------------------------------------------------
        If clsCommon.CompairString(strItemType, RowTypeItem) = CompairStringResult.Equal Then
            Dim qry As String
            qry = "select * from (select a.DESCRIPTION,a.cat_value, TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.item_code as Item,TSPL_ITEM_MASTER.item_desc as [ItemDesc]," _
            & " TSPL_ITEM_MASTER.Unit_Code as Unit , TSPL_ITEM_MASTER.Rate as BasicRate,TSPL_ITEM_MASTER.rate as MRP, Weight_Value as [Weight Value] from " _
            & " TSPL_ITEM_MASTER    left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code," _
            & " TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value " _
            & " from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE=" _
            & " TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' left outer join " _
            & " TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code" _
            & " and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a" _
            & " on a.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_ITEM_MASTER.item_code=a.item_code " _
             & "  where  TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and coalesce(Product_Type,'') not in ('MI') and Item_Type not in ('A') and coalesce(Item_used_as,'')='S' "
            If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
                qry += " and TSPL_ITEM_MASTER.IsTaxable='" + IIf(chkTaxable.Checked, "1", "0") + "'"
            End If
            If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)) > 0 Then
                qry += " and TSPL_ITEM_MASTER.Item_code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' "
            End If
            qry += " and TSPL_ITEM_MASTER.Short_Description='" & strShortDesc & "') as s pivot(max(cat_value) for description in ([BRAND],[ITEM TYPE1],[MAIN GROUP],[PACKING],[PRODUCT CATEGORY],[SKU],[SUB GROUP]))t  "
            gv1.CurrentRow.Cells(colShortDesc).Value = clsCommon.ShowSelectForm("MSA@Item", qry, "Short_Description", "", strShortDesc, "Short_Description", isButtonClick)
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing And dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If Not dr Is Nothing Then
                        If clsCommon.myLen(clsCommon.myCstr(dr("Item"))) > 0 Then
                            gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
                            gv1.CurrentRow.Cells(colICode).Value = clsCommon.myCstr(dr("Item"))
                            gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dr("Unit")), clsCommon.myCDecimal(dr("MRP")))
                            gv1.CurrentRow.Cells(colIName).Value = clsCommon.myCstr(dr("ItemDesc"))
                            gv1.CurrentRow.Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dr("Item")), Nothing)
                            gv1.CurrentRow.Cells(colPriceCOde).Value = ""
                            gv1.CurrentRow.Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                            gv1.CurrentRow.Cells(colUOMName).Value = clsDBFuncationality.getSingleValue("select Unit_Desc from tspl_unit_master where Unit_code= '" + clsCommon.myCstr(dr("Unit")) + "'")
                            gv1.CurrentRow.Cells(colOrgUnit).Value = clsCommon.myCstr(dr("Unit"))
                            gv1.CurrentRow.Cells(colMRP).Value = clsCommon.myCDecimal(dr("MRP"))
                            If ApplySaleRouteInDCSSale Then
                                gv1.CurrentRow.Cells(colRate).Value = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(dr("Item")), clsCommon.myCstr(dr("Unit")), txtDate.Value, Nothing, FndPriceCode.Value)
                            Else
                                gv1.CurrentRow.Cells(colRate).Value = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(dr("Item")), clsCommon.myCstr(dr("Unit")), txtDate.Value)
                            End If
                            gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                            gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                            gv1.CurrentRow.Cells(colActualCost).Value = clsCommon.myCDecimal(dr("BasicRate"))
                            gv1.CurrentRow.Cells(ColFOC).Value = 0
                            gv1.CurrentRow.Cells(colItemWeight).Value = clsCommon.myCDecimal(dr("Weight Value"))
                            gv1.CurrentRow.Cells(colOrgCost).Value = clsCommon.myCDecimal(dr("BasicRate"))
                            gv1.CurrentRow.Cells(colShortDesc).Value = clsCommon.myCstr(dr("Short_Description"))
                            Dim dblConvF As Decimal
                            dblConvF = GetConvFactor(gv1.CurrentRow.Cells(colUnit).Value, gv1.CurrentRow.Cells(colICode).Value)
                            gv1.CurrentRow.Cells(colConvF).Value = dblConvF
                            SetitemWiseTaxSetting(True, True)
                        End If
                    Else
                        SetBlankOfItemColumns()
                    End If
                Next
            Else
                SetBlankOfItemColumns()
            End If
        End If
        setBalance()
    End Sub
    Sub OpenICodeListCurrentCalaculation(ByVal isButtonClick As Boolean)
        If ApplySaleRouteInDCSSale Then
            If clsCommon.myLen(FndPriceCode.Value) <= 0 Then
                gv1.CurrentRow.Cells(colICode).Value = ""
                Throw New Exception("Please select Price Code first.")
            End If
        End If
        gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
        Dim strItemType As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value)
        Dim pivotheader As String = ""
        pivotheader = clsItemMaster.GetNLevelPivotHeader()
        '----------------------------------------------------------------
        If clsCommon.CompairString(strItemType, RowTypeItem) = CompairStringResult.Equal Then
            Dim qry As String
            qry = "select * from (
select a.DESCRIPTION,a.cat_value, TSPL_ITEM_MASTER.item_code as Item,TSPL_ITEM_MASTER.item_desc as [ItemDesc],TSPL_ITEM_MASTER.Short_Description, 
TSPL_ITEM_MASTER.Unit_Code as Unit , TSPL_ITEM_MASTER.Rate as BasicRate,TSPL_ITEM_MASTER.rate as MRP, Weight_Value as [Weight Value],isnull(DCS_Sale_Zero_Cost,0)as DCS_Sale_Zero_Cost
from TSPL_ITEM_MASTER    
left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code,
TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value  from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE= TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' 
left outer join  TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a on a.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_ITEM_MASTER.item_code=a.item_code "
            Dim whrcls As String = " "
            whrcls = " TSPL_ITEM_MASTER.Active = 1 and TSPL_ITEM_MASTER.Deduction ='" + clsCommon.myCstr(cboDeductionType.SelectedValue) + "'"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                whrcls += "  And coalesce(Product_Type,'') not in ('MI') and Item_Type not in ('A') and coalesce(Item_used_as,'')='S' "
            Else
                whrcls += "  And Is_FreshItem = 0 And coalesce(Product_Type,'') not in ('MI') and Item_Type not in ('A') and coalesce(Item_used_as,'')='S' "
            End If

            If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
                whrcls += " and TSPL_ITEM_MASTER.IsTaxable='" + IIf(chkTaxable.Checked, "1", "0") + "'"
            End If
            whrcls += " ) as s pivot(max(cat_value) for description in ([BRAND],[ITEM TYPE1],[MAIN GROUP],[PACKING],[PRODUCT CATEGORY],[SKU],[SUB GROUP]))t  "
            ' Dim dr As DataRow = clsCommon.ShowSelectFormForRow("MSA@Items", qry)
            gv1.CurrentRow.Cells(colICode).Value = clsCommon.ShowSelectForm("ITEMSTFND", qry, "Item", whrcls, gv1.CurrentRow.Cells(colICode).Value, "Item", False)

            qry += " where " & whrcls & ""
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            Dim dr As DataRow() = clsCommon.MyDTSelect(dt, "Item='" + gv1.CurrentRow.Cells(colICode).Value + "'")
            If dr.Length > 0 Then
                If clsCommon.myLen(clsCommon.myCstr(dr(0).Item("Item"))) > 0 Then
                    gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
                    gv1.CurrentRow.Cells(colICode).Value = clsCommon.myCstr(dr(0).Item("Item"))
                    gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dr(0).Item("Unit")), clsCommon.myCDecimal(dr(0).Item("MRP")))
                    gv1.CurrentRow.Cells(colIName).Value = clsCommon.myCstr(dr(0).Item("ItemDesc"))
                    gv1.CurrentRow.Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dr(0).Item("Item")), Nothing)
                    gv1.CurrentRow.Cells(colPriceCOde).Value = ""
                    '    gv1.CurrentRow.Cells(colUnit).Value = clsCommon.myCstr(dr(0).Item("Unit"))
                    gv1.CurrentRow.Cells(colUnit).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select UOM_Code from TSPL_ITEM_UOM_DETAIL where Item_Code = '" & clsCommon.myCstr(dr(0).Item("Item")) & "' and default_uom = 1"))
                    gv1.CurrentRow.Cells(colUOMName).Value = clsDBFuncationality.getSingleValue("select Unit_Desc from tspl_unit_master where Unit_code= '" + clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) + "'")
                    gv1.CurrentRow.Cells(colOrgUnit).Value = clsCommon.myCstr(dr(0).Item("Unit"))
                    gv1.CurrentRow.Cells(colMRP).Value = clsCommon.myCDecimal(dr(0).Item("MRP"))
                    gv1.CurrentRow.Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(gv1.CurrentRow.Cells(colICode).Value)
                    gv1.CurrentRow.Cells(colDCSSaleZeroCost).Value = clsCommon.myCBool(dr(0).Item("DCS_Sale_Zero_Cost") = 1)
                    If gv1.CurrentRow.Cells(colDCSSaleZeroCost).Value Then
                        gv1.CurrentRow.Cells(colRate).Value = 0
                    Else
                        If ApplySaleRouteInDCSSale Then
                            If clsCommon.myLen(FndPriceCode.Value) > 0 Then
                                gv1.CurrentRow.Cells(colRate).Value = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(dr(0).Item("Item")), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), txtDate.Value, Nothing, FndPriceCode.Value)
                            End If
                        Else
                            gv1.CurrentRow.Cells(colRate).Value = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(dr(0).Item("Item")), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), txtDate.Value) 'clsCommon.myCdbl(dr("BasicRate"))
                        End If
                    End If
                    gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                    gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                    gv1.CurrentRow.Cells(colActualCost).Value = clsCommon.myCDecimal(dr(0).Item("BasicRate"))
                    gv1.CurrentRow.Cells(ColFOC).Value = 0
                    gv1.CurrentRow.Cells(colItemWeight).Value = clsCommon.myCDecimal(dr(0).Item("Weight Value"))
                    gv1.CurrentRow.Cells(colOrgCost).Value = clsCommon.myCDecimal(dr(0).Item("BasicRate"))
                    gv1.CurrentRow.Cells(colShortDesc).Value = clsCommon.myCstr(dr(0).Item("Short_Description"))
                    Dim dblConvF As Decimal
                    dblConvF = GetConvFactor(gv1.CurrentRow.Cells(colUnit).Value, gv1.CurrentRow.Cells(colICode).Value)
                    gv1.CurrentRow.Cells(colConvF).Value = dblConvF
                    If chkTaxable.Checked Then
                        Dim qry1 As String = " select Top 1 TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code from TSPL_ITEM_WISE_TAX Inner Join TSPL_ITEM_WISE_TAX_AUTHORITY On TSPL_ITEM_WISE_TAX_AUTHORITY.HCODE=TSPL_ITEM_WISE_TAX.HCODE
Inner Join TSPL_ITEM_WISE_TAX_GROUP On TSPL_ITEM_WISE_TAX_GROUP.HCODE=TSPL_ITEM_WISE_TAX_AUTHORITY.HCODE
where TSPL_ITEM_WISE_TAX.Type='S' And TSPL_ITEM_WISE_TAX.Status=1 And TSPL_ITEM_WISE_TAX_GROUP.Item_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) + "'  "
                        If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                            qry1 += "  and TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code ='" & txtTaxGroup.Value & "'"
                        End If
                        qry1 +="Group By CONVERT(date,TSPL_ITEM_WISE_TAX.DOC_DATE,103),TSPL_ITEM_WISE_TAX_GROUP.Tax_Group_Code "
                        qry1 += " Order By CONVERT(date,TSPL_ITEM_WISE_TAX.DOC_DATE,103) Desc"
                        txtTaxGroup.Value = clsDBFuncationality.getSingleValue(qry1)
                        SetTaxDetails()
                        If clsCommon.myLen(chkTaxGroup) > 0 Then
                            If clsCommon.CompairString(txtTaxGroup.Value, chkTaxGroup) <> CompairStringResult.Equal Then
                                Throw New Exception("Item (" + clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) + ") does not belong to tax group (" + txtTaxGroup.Value + ") in Item Wise Tax.")
                            End If
                        End If
                        chkTaxGroup = txtTaxGroup.Value
                    End If
                    If chkApplyTPT.Checked Then
                        Dim dtTransportationCharges As DataTable = clsDCSTransportationCharges.PickTransportationRate(LblVlc_Code.Tag, gv1.CurrentRow.Cells(colICode).Value, txtDate.Value)
                        If dtTransportationCharges.Rows.Count > 0 AndAlso dtTransportationCharges IsNot Nothing Then
                            gv1.CurrentRow.Cells(ColTPTRate).Value = clsCommon.myCDecimal(dtTransportationCharges.Rows(0)("Transportation_Rate"))
                            gv1.CurrentRow.Cells(ColTPTRate).Tag = clsCommon.myCDecimal(dtTransportationCharges.Rows(0)("PK_ID"))
                        End If
                    End If
                    SetitemWiseTaxSetting(True, True)
                End If
            Else
                SetBlankOfItemColumns()
            End If
        End If
        setBalance()
    End Sub
    Private Function GetConvFactor(ByVal strUnit As String, ByVal strItem As String) As Decimal
        Dim dblConvF As Decimal = 0
        dblConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strUnit & "'"))
        Return dblConvF
    End Function
    Private Sub SetBlankOfItemColumns()
        gv1.CurrentRow.Cells(colICode).Value = ""
        gv1.CurrentRow.Cells(colIName).Value = ""
        gv1.CurrentRow.Cells(colShortDesc).Value = ""
        gv1.CurrentRow.Cells(colHSNNo).Value = ""
        gv1.CurrentRow.Cells(colUnit).Value = ""
        gv1.CurrentRow.Cells(colUOMName).Value = ""
        gv1.CurrentRow.Cells(colMRP).Value = 0
        gv1.CurrentRow.Cells(ColTPTRate).Value = 0
        gv1.CurrentRow.Cells(ColTPTRate).Tag = Nothing
    End Sub
    Private Sub gv1_UserAddedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserAddedRow
        For i As Integer = 0 To gv1.Rows.Count - 1
            gv1.Rows(0).Cells(0).Value = 1
            If i <> 0 Then
                gv1.Rows(i).Cells(colLineNo).Value = i + 1
            End If
        Next
    End Sub
    Private Function GetCurrentRowTotalTaxAmt(ByVal IntRowNo As Integer) As Decimal
        Dim dblTotTax As Decimal = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                dblTotTax = dblTotTax + clsCommon.myCDecimal(gv1.CurrentRow.Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value)
            Else
                dblTotTax = dblTotTax + clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value)
            End If
        Next
        Return dblTotTax
    End Function
    Private Function GetCurrentRowSurTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Decimal
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCDecimal(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            Else
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            End If
        Next
        Return 0
    End Function
    Private Function GetCurrentRowOtherTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Decimal
        Dim dblRet As Decimal = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If IntRowNo < 0 Then
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCDecimal(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                Else
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                End If
            Next
        Next
        Return dblRet
    End Function
    Private Sub BlankTaxDetails(ByVal intRowNo As Integer)
        BlankTaxDetails(intRowNo, True)
    End Sub
    Private Sub BlankTaxDetails(ByVal intRowNo As Integer, ByVal isBlankRate As Boolean)
        For ii As Integer = 1 To 10
            Dim strII As String = clsCommon.myCstr(ii)
            If intRowNo < 0 Then
                gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("COLTAXBASEAMT" + strII)).Value = Nothing
                If isBlankRate Then
                    gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                    gv1.CurrentRow.Cells(colItemwiseTaxCode).Value = Nothing
                End If
                gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            Else
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + strII)).Value = Nothing
                If isBlankRate Then
                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                    gv1.Rows(intRowNo).Cells(colItemwiseTaxCode).Value = Nothing
                End If
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            End If
        Next
    End Sub
    Private Sub UpdateAllTotals()
        Dim dblCashDisAmt As Decimal = 0
        Dim dblHeadDisPerAmt As Decimal = 0
        Dim dblCommAmt As Decimal = 0
        Dim dblTotAmt As Decimal = 0
        Dim dblTotDisAmt As Decimal = 0
        Dim dblAmtAfterDis As Decimal = 0
        Dim dblTotLandedCost As Decimal = 0
        Dim dblHeadDisAmt As Decimal = 0
        Dim dblTaxBaseAmt1 As Decimal = 0
        Dim dblTaxBaseAmt2 As Decimal = 0
        Dim dblTaxBaseAmt3 As Decimal = 0
        Dim dblTaxBaseAmt4 As Decimal = 0
        Dim dblTaxBaseAmt5 As Decimal = 0
        Dim dblTaxBaseAmt6 As Decimal = 0
        Dim dblTaxBaseAmt7 As Decimal = 0
        Dim dblTaxBaseAmt8 As Decimal = 0
        Dim dblTaxBaseAmt9 As Decimal = 0
        Dim dblTaxBaseAmt10 As Decimal = 0
        Dim dblTaxAmt1 As Decimal = 0
        Dim dblTaxAmt2 As Decimal = 0
        Dim dblTaxAmt3 As Decimal = 0
        Dim dblTaxAmt4 As Decimal = 0
        Dim dblTaxAmt5 As Decimal = 0
        Dim dblTaxAmt6 As Decimal = 0
        Dim dblTaxAmt7 As Decimal = 0
        Dim dblTaxAmt8 As Decimal = 0
        Dim dblTaxAmt9 As Decimal = 0
        Dim dblTaxAmt10 As Decimal = 0
        TotalItemQty = 0
        Dim dblTaxTotAmt As Decimal = 0
        Dim dblNetAmt As Decimal = 0
        Dim dblTotalTPTAmt As Double = 0

        For ii As Integer = 0 To gv1.Rows.Count - 1
            If (clsCommon.myLen(gv1.Rows(ii).Cells(colICode).Value) > 0) And clsCommon.myCDecimal(gv1.Rows(ii).Cells(ColFOC).Value) = 0 Then
                dblTotAmt = dblTotAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colAmt).Value)
                dblTotDisAmt = dblTotDisAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colDisAmt).Value)
                dblCashDisAmt = dblCashDisAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTotalCustDiscount).Value)
                dblHeadDisAmt = dblHeadDisAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colHeadDiscamt).Value)
                dblAmtAfterDis = dblAmtAfterDis + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colAmtAfterDis).Value)
                dblHeadDisPerAmt = dblHeadDisPerAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colHeadDisPerAmt).Value)
                dblCommAmt = dblCommAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(ColCommAmt).Value)
                dblTaxAmt1 = dblTaxAmt1 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt1).Value)
                dblTaxAmt2 = dblTaxAmt2 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt2).Value)
                dblTaxAmt3 = dblTaxAmt3 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt3).Value)
                dblTaxAmt4 = dblTaxAmt4 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt4).Value)
                dblTaxAmt5 = dblTaxAmt5 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt5).Value)
                dblTaxAmt6 = dblTaxAmt6 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt6).Value)
                dblTaxAmt7 = dblTaxAmt7 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt7).Value)
                dblTaxAmt8 = dblTaxAmt8 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt8).Value)
                dblTaxAmt9 = dblTaxAmt9 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt9).Value)
                dblTaxAmt10 = dblTaxAmt10 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxAmt10).Value)
                dblTaxBaseAmt1 = dblTaxBaseAmt1 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt1).Value)
                dblTaxBaseAmt2 = dblTaxBaseAmt2 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt2).Value)
                dblTaxBaseAmt3 = dblTaxBaseAmt3 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt3).Value)
                dblTaxBaseAmt4 = dblTaxBaseAmt4 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt4).Value)
                dblTaxBaseAmt5 = dblTaxBaseAmt5 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt5).Value)
                dblTaxBaseAmt6 = dblTaxBaseAmt6 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt6).Value)
                dblTaxBaseAmt7 = dblTaxBaseAmt7 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt7).Value)
                dblTaxBaseAmt8 = dblTaxBaseAmt8 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt8).Value)
                dblTaxBaseAmt9 = dblTaxBaseAmt9 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt9).Value)
                dblTaxBaseAmt10 = dblTaxBaseAmt10 + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTaxBaseAmt10).Value)
                TotalItemQty = TotalItemQty + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colQty).Value)
                dblTaxTotAmt = dblTaxTotAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colTotTaxAmt).Value)
                dblNetAmt = dblNetAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colAmtAfterTax).Value)
                dblTotLandedCost = dblTotLandedCost + clsCommon.myCDecimal(gv1.Rows(ii).Cells(colLandedAmt).Value)
                dblTotalTPTAmt = dblTotalTPTAmt + clsCommon.myCDecimal(gv1.Rows(ii).Cells(ColTPTAmt).Value)
            End If
        Next
        If rbtnTaxCalAutomatic.IsChecked Then
            For ii As Integer = 1 To gv2.Rows.Count
                Select Case (ii)
                    Case 1
                        If chkisTCS.Checked Then
                            If (clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblTaxBaseAmt1)
                                If clsCommon.myCDecimal(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt1 = Math.Round(txttcstaxbaseamount.Value, 6)
                                Else
                                    dblTaxBaseAmt1 = lblActualTCSTaxBaseAmt.Text
                                End If
                                dblTaxAmt1 = (dblTaxBaseAmt1 * txtTCSTaxRate.Value) / 100
                            End If
                        End If
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt1, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt1, 6)
                        If dblTaxBaseAmt1 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt1 * 100) / dblTaxBaseAmt1, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 2
                        If chkisTCS.Checked Then
                            If (clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = (dblTaxBaseAmt1 + dblTaxAmt1)
                                If clsCommon.myCDecimal(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt2 = Math.Round(txttcstaxbaseamount.Value, 6)
                                Else
                                    dblTaxBaseAmt2 = lblActualTCSTaxBaseAmt.Text
                                End If
                                dblTaxAmt2 = (dblTaxBaseAmt2 * txtTCSTaxRate.Value) / 100
                            End If
                        End If
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt2, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt2, 6)
                        If dblTaxBaseAmt2 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt2 * 100) / dblTaxBaseAmt2, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 3
                        If chkisTCS.Checked Then
                            If (clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = (dblTaxBaseAmt1 + dblTaxAmt1 + dblTaxAmt2)
                                If clsCommon.myCDecimal(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt3 = clsCommon.myCDecimal(txttcstaxbaseamount.Value)
                                Else
                                    dblTaxBaseAmt3 = clsCommon.myCDecimal(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt3 = (dblTaxBaseAmt3 * txtTCSTaxRate.Value) / 100
                            End If
                        End If
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt3, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt3, 6)
                        If dblTaxBaseAmt3 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt3 * 100) / dblTaxBaseAmt3, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 4
                        If chkisTCS.Checked Then
                            If (clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = (dblTaxBaseAmt1 + dblTaxAmt1 + dblTaxAmt2 + dblTaxAmt3)
                                If clsCommon.myCDecimal(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt4 = clsCommon.myCDecimal(txttcstaxbaseamount.Value)
                                Else
                                    dblTaxBaseAmt4 = clsCommon.myCDecimal(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt4 = (dblTaxBaseAmt4 * txtTCSTaxRate.Value) / 100
                            End If
                        End If
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt4, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt4, 6)
                        If dblTaxBaseAmt4 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt4 * 100) / dblTaxBaseAmt4, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 5
                        If chkisTCS.Checked Then
                            If (clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = (dblTaxBaseAmt1 + dblTaxAmt1 + dblTaxAmt2 + dblTaxAmt3 + dblTaxAmt4)
                                If clsCommon.myCDecimal(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt5 = clsCommon.myCDecimal(lblActualTCSTaxBaseAmt.Text)
                                Else
                                    dblTaxBaseAmt5 = clsCommon.myCDecimal(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt5 = (dblTaxBaseAmt5 * txtTCSTaxRate.Value) / 100
                            End If
                        End If
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt5, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt5, 6)
                        If dblTaxBaseAmt5 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt5 * 100) / dblTaxBaseAmt5, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 6
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt6, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt6, 6)
                        If dblTaxBaseAmt6 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt6 * 100) / dblTaxBaseAmt6, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 7
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt7, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt7, 6)
                        If dblTaxBaseAmt7 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt7 * 100) / dblTaxBaseAmt7, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 8
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt8, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt8, 6)
                        If dblTaxBaseAmt8 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt8 * 100) / dblTaxBaseAmt8, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 9
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt9, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt9, 6)
                        If dblTaxBaseAmt9 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt9 * 100) / dblTaxBaseAmt9, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 10
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt10, 6)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt10, 6)
                        If dblTaxBaseAmt10 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt10 * 100) / dblTaxBaseAmt10, 2)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                End Select
            Next
        End If
        Dim dblACAmount As Decimal = 0
        For ii As Integer = 0 To gvAC.Rows.Count - 1
            If (clsCommon.myLen(gvAC.Rows(ii).Cells(colACCode).Value) > 0) Then
                dblACAmount = dblACAmount + clsCommon.myCDecimal(gvAC.Rows(ii).Cells(colACAmount).Value)
            End If
        Next
        lblAmtWithDiscount.Text = clsCommon.myFormat(dblTotAmt)
        lblDiscountAmt.Text = clsCommon.myFormat(dblTotDisAmt + dblCashDisAmt)
        If MultiplySubsidyWithQuantity Then
            Dim AmtAfterDiscount As Decimal = lblAmtWithDiscount.Text
            Dim InvAmt As Decimal = clsCommon.myCDecimal(lblInvoiceDiscAmt.Text)
            AmtAfterDiscount = AmtAfterDiscount - InvAmt
            lblAmtAfterDiscount.Text = clsCommon.myFormat(AmtAfterDiscount)
            lblInvoiceDiscAmt.Text = lblTotalDisSubsidy.Text
        Else
            lblAmtAfterDiscount.Text = clsCommon.myFormat(dblAmtAfterDis)
            lblInvoiceDiscAmt.Text = dblHeadDisAmt + dblHeadDisPerAmt
        End If
        lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt)
        lblAddCharges.Text = clsCommon.myFormat(dblACAmount)
        lblAddCharges1.Text = clsCommon.myFormat(dblACAmount)
        dblNetAmt = dblNetAmt + dblACAmount
        If MultiplySubsidyWithQuantity Then
            Dim TotDocAmt As Decimal = lblAmtAfterDiscount.Text
            Dim TaxAmt As Decimal = lblTaxAmt.Text
            TotDocAmt = TotDocAmt + TaxAmt
            lblTotRAmt.Text = clsCommon.myFormat(TotDocAmt)
        Else
            lblTotRAmt.Text = clsCommon.myFormat(dblNetAmt)
        End If
        lblCommAmt.Text = clsCommon.myFormat(dblCommAmt)
        If AllowRoundOffAmountOnDCSSale Then
            Dim lstDecml As New List(Of Decimal)
            lstDecml = ClsScrapSaleHead.Calculate_RoundOffAmt(clsCommon.myCDecimal(lblTotRAmt.Text), Nothing)
            If lstDecml IsNot Nothing AndAlso lstDecml.Count > 0 Then
                lblTotRAmt.Text = clsCommon.myCDecimal(lstDecml(0))
                txtRoundOff.Text = clsCommon.myCDecimal(lstDecml(1))
            End If
        Else
            txtRoundOff.Text = 0
        End If
        lblTotRAmt1.Text = lblTotRAmt.Text
        txtTPTAmt.Text = dblTotalTPTAmt
        'lblGrossAmount.Text =
    End Sub
    Private Function GetBaseOtherTaxableAmount(ByVal intEndCol As Integer) As Decimal
        ''Dim dblRetVal As Decimal = 0
        ''For ii As Integer = 0 To intEndCol - 1
        ''    If clsCommon.myCBool(gv2.Rows(ii).Cells(colTIsTaxable).Value) Then
        ''        dblRetVal = dblRetVal + clsCommon.myCDecimal(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
        ''    End If
        ''Next
        ''Return dblRetVal
    End Function
    Private Function GetBaseTaxAmount(ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Decimal
        For ii As Integer = 0 To intEndCol
            If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAutCode).Value)) = CompairStringResult.Equal Then
                Return clsCommon.myCDecimal(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
            End If
        Next
        Return 0
    End Function
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        AddNew()
    End Sub
    Sub AddNew()
        rbtnCancel.Visible = False
        btnInvoiceJE.Visible = False
        BlankAllControls()
        fndProject.Enabled = True
        lblProject.Enabled = True
        LoadBlankGrid()
        LoadBlankGridAC()
        LoadBlankGridTax()
        isNewEntry = True
        btnSave.Text = "Save"
        chkisTCS.Checked = True
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnDelete.Enabled = True
        txtDate.Focus()
        gv1.Rows.AddNew()
        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
        ''gv1.Rows.AddNew()
        chkInternal.Checked = False
        gvAC.Rows.AddNew()
        gvAC.Rows.AddNew()
        txtDate.Enabled = True
        txtVendorNo.Enabled = True
        chkTaxable.Enabled = True
        btnHistory.Enabled = False
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.SetDefaultValues()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()
        If AllowChangeInvoiceType Then
            lblInvoiceType.Visible = True
            ddlInvoiceType.Visible = True
        Else
            lblInvoiceType.Visible = False
            ddlInvoiceType.Visible = False
        End If
        lblActualTCSTaxBaseAmt.Text = ""
        txttcstaxbaseamount.Text = ""
        txtElectronicRefNo.Text = ""
        txtTaxGroup.Enabled = True
        txtReceiptNo.Value = ""
        lblReceiptAmt.Text = ""
        txtBankCode.Value = ""
        txtBankCode.Enabled = True
        txtVehicleNo.Text = ""
        txtReceiverName.Text = ""
        lblTotalSubsidy.Text = ""
        If MultiplySubsidyWithQuantity Then
            txtRateAmt.Enabled = True
            txtDiscAmt.Enabled = True
        Else
            txtRateAmt.Enabled = False
            txtDiscAmt.Enabled = False
        End If
        txtRateAmt.Text = 0
        txtRatePer.Text = 0
        lblGrossAmount.Text = ""
        cmbPaymentType.SelectedIndex = 0
        'If DefaultEnableEWayBill Then
        '    chkIsEwayBill.Checked = True
        'Else
        '    chkIsEwayBill.Checked = False
        'End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
            If Not chkcashsale.Checked Then
                txtReceiptNo.Visible = False
                lblReceiptAmt.Visible = False
                lblReceiptNo.Visible = False
                lblPaymentType.Visible = False
            End If
        Else
            txtReceiptNo.Visible = False
            lblReceiptAmt.Visible = False
            lblReceiptNo.Visible = False
            'lblPaymentType.Visible = True
            'cmbPaymentType.Visible = True
            lblPaymentType.Visible = False
            cmbPaymentType.Visible = False
        End If

        If MultiplySubsidyWithQuantity Then
            lblTotalSubsidy.Visible = True
            lblTotalDisSubsidy.Visible = True
        Else
            lblTotalSubsidy.Visible = False
            lblTotalDisSubsidy.Visible = False

        End If
        lblTotalDisSubsidy.Text = 0
        lblTotalSubsidy.Text = 0
        SETGSTControl()
        'If AllowPlandDeptMCCLocation = True Then
        '    Dim obj As New clsMCCCodes()
        '    obj = clsMCCCodes.GetData(True)
        '    If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Default_LocCode) > 1 Then
        '        txtBillToLocation.Value = obj.Default_LocCode
        '        lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "' "))
        '    Else
        '    End If
        'End If

        txtBillToLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
        If clsCommon.myLen(txtBillToLocation.Value) > 0 Then
            lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "'"))
        End If
        If clsCommon.myLen(clsCommon.myCstr(txtBillToLocation.Value)) > 0 Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                txtSubLocation.Enabled = True
                txtSubLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Sub_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "'"))
                lblSubLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtSubLocation.Value + "'"))
            Else
                txtSubLocation.Enabled = False
                txtSubLocation.Value = ""
                lblSubLocation.Text = ""
            End If

        End If
        txtRecommBy.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select top 1 Recommended_By from TSPL_DCS_SALE_ENTRY where Bill_To_Location='" + txtBillToLocation.Value + "' order by Document_Date desc,TSPL_DCS_SALE_ENTRY.Document_Code "))
    End Sub
    Function AllowToSave(ByVal ChekPostBtn As Boolean) As Boolean
        Try
            Xtra.TransactionValidity(txtDate.Value)
            If AllowFutureDateTransaction(txtDate.Value, Nothing) = False Then
                txtDate.Focus()
                Return False
            End If
            If CalculateTaxRatefromItemwsieTaxOnSale = 1 Then
                SetitemWiseTaxSetting(True, False)
            End If
            If objCommonVar.CreateAutoReceiptEntryDCSSale Then
                If chkcashsale.Checked Then
                    If clsCommon.myLen(txtBankCode.Value) <= 0 Then
                        clsCommon.MyMessageBoxShow(Me, "Please select Bank Code", Me.Text)
                        txtBankCode.Focus()
                        Return False
                    End If
                End If
            End If
            If ApplySaleRouteInDCSSale Then
                If clsCommon.myLen(FndPriceCode.Value) <= 0 Then
                    FndPriceCode.Focus()
                    Throw New Exception("Price Code cannot be blank.")
                End If
                If clsCommon.myLen(txtSaleRoute.Value) <= 0 Then
                    txtSaleRoute.Focus()
                    Throw New Exception("Sale Route cannot be blank.")
                End If
            End If
            If chkApplyTPT.Checked Then
                If clsCommon.myLen(txtTPTVendor.Value) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Please select TPT Vendor", Me.Text)
                    txtTPTVendor.Focus()
                    Return False
                End If
            End If
            If clsCommon.myCDecimal(lblTotalSubsidy.Text) > clsCommon.myCDecimal(lblTotRAmt.Text) Then
                Throw New Exception("Subsidy amount cannot be greater than Document amount")
            End If
            If clsCommon.myCDecimal(txtRateAmt.Text) < 0 Then
                Throw New Exception("Rate amount cannot be negative")
            End If
            If clsCommon.myCDecimal(lblTotalSubsidy.Text) < 0 Then
                Throw New Exception("Subsidy amount cannot be negative")
            End If
            If clsCommon.myLen(txtTransporterCode.Value) < 0 AndAlso Not chkNoTranspoter.Checked Then
                Throw New Exception("Please Select Transpoter")
            End If
            For ii As Integer = 0 To gv1.RowCount - 1
                If chkApplyTPT.Checked Then
                    Dim dtTransportationCharges As DataTable = clsDCSTransportationCharges.PickTransportationRate(LblVlc_Code.Tag, clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value), txtDate.Value)
                    If dtTransportationCharges.Rows.Count > 0 AndAlso dtTransportationCharges IsNot Nothing Then
                        gv1.Rows(ii).Cells(ColTPTRate).Value = clsCommon.myCDecimal(dtTransportationCharges.Rows(0)("Transportation_Rate"))
                        gv1.Rows(ii).Cells(ColTPTRate).Tag = clsCommon.myCDecimal(dtTransportationCharges.Rows(0)("PK_ID"))
                    End If
                Else
                    gv1.Rows(ii).Cells(ColTPTRate).Value = 0
                    gv1.Rows(ii).Cells(ColTPTRate).Tag = Nothing
                End If
                If clsCommon.myLen(gv1.Rows(ii).Cells(colICode).Value) > 0 Then
                    If (gv1.Rows(ii).Cells(colQty).Value) <= 0 Then
                        clsCommon.MyMessageBoxShow(Me, "Please enter quantity of Item [" & gv1.Rows(ii).Cells(colICode).Value & "] ", Me.Text)
                        Return False
                    End If
                End If
                UpdateCurrentRow(ii)
            Next
            UpdateAllTotals()
            CalculateDiscountAmount()
            CalculateRateDiffAmount()
            If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer", Me.Text)
                txtVendorNo.Focus()
                Return False
            End If
            GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
            If GSTStatus = False OrElse (chkTaxable.Checked AndAlso GSTStatus = True) Then
                If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please select Tax Group", Me.Text)
                    txtTaxGroup.Focus()
                    Return False
                End If
            End If
            If clsCommon.myLen(txtBillToLocation.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Bill to Location", Me.Text)
                txtBillToLocation.Focus()
                Return False
            End If
            If clsCommon.myLen(clsCommon.myCstr(txtBillToLocation.Value)) > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                    If clsCommon.myLen(txtSubLocation.Value) <= 0 Then
                        Throw New Exception("Please select Sub Location")
                    End If
                End If
            End If
            '' Check EWB Threshold Limit 28-Jan-2026
            If ApplyEWBThresholdLimit AndAlso chkTaxable.Checked Then

                Dim EWBThresholdLimtCat As String = ""
                Dim EWBThresholdLimtCatForCust As String = ""
                chkIsEwayBill.Checked = False
                Dim EwbNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EWayBillNo from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & txtInvoiceNo.Text & "'"))
                If clsCommon.myLen(EwbNo) = 0 Then
                    If clsCommon.myLen(clsCommon.myCstr(txtBillToLocation.Value)) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "") = CompairStringResult.Equal Then
                                If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                                    EWBThresholdLimtCat = "select City_Code,State from TSPL_LOCATION_MASTER where Location_Code='" & txtSubLocation.Value & "'"

                                End If
                            End If
                        Else
                            EWBThresholdLimtCat = "select City_Code,State from TSPL_LOCATION_MASTER where Location_Code='" & txtBillToLocation.Value & "'"
                        End If
                        EWBThresholdLimtCatForCust = "select City_Code,State from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"
                        Dim dt_FromLocation As DataTable = clsDBFuncationality.GetDataTable(EWBThresholdLimtCat)
                        Dim dt_CustLocation As DataTable = clsDBFuncationality.GetDataTable(EWBThresholdLimtCatForCust)
                        If clsCommon.CompairString(clsCommon.myCstr(dt_FromLocation(0)("City_Code")), clsCommon.myCstr(dt_CustLocation(0)("City_Code"))) = CompairStringResult.Equal Then
                            If clsCommon.myCDecimal(lblTotRAmt1.Text) > EWBThresholdLimitForIntraCity Then
                                chkIsEwayBill.Checked = True
                            End If
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dt_FromLocation(0)("State")), clsCommon.myCstr(dt_CustLocation(0)("State"))) = CompairStringResult.Equal Then
                            If clsCommon.myCDecimal(lblTotRAmt1.Text) > EWBThresholdLimitForIntraState Then
                                chkIsEwayBill.Checked = True
                            End If
                        Else
                            If clsCommon.myCDecimal(lblTotRAmt1.Text) > EWBThresholdLimitForInterState Then
                                chkIsEwayBill.Checked = True
                            End If
                        End If
                    End If
                Else
                    chkIsEwayBill.Checked = True
                End If

            End If

            '' End  of EWB Threshold Limit
            If Not isNewEntry AndAlso clsCommon.myLen(txtDocNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Shipment No Not found to save", Me.Text)
                txtDocNo.Focus()
                Return False
            End If
            If AllowChangeInvoiceType Then
                If clsCommon.myLen(ddlInvoiceType.SelectedValue) <= 0 Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please select invoice  Type for creating auto invoice", Me.Text)
                    cboItemType.Focus()
                    Return False
                Else
                    If InvoiceType() = False Then
                        Return False
                    End If
                End If
            Else
                InvoiceType()
            End If
            Dim arrReqNo As New List(Of String)
            Dim arrICode As New List(Of String)()
            Dim arrProjNo As New List(Of String)
            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strReqNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colOrderNo).Value)
                Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                Dim strIName As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                Dim dblPendingQty As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colPendingQty).Value)
                Dim dblQty As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colQty).Value)
                Dim dblMRP As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colMRP).Value)
                Dim strUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                Dim strSchemeItem As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeItem).Value)
                Dim strProject As String = ""
                If clsCommon.myLen(strICode) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
                    If clsCommon.myLen(strUOM) <= 0 Then
                        common.clsCommon.MyMessageBoxShow("Please enter UOM for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                    Dim objCustItem As clsCustomeritemDetails = clsCustomeritemDetails.GetItemRateAndDiscount(txtVendorNo.Value, strICode, clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), txtDate.Value)
                    If objCustItem IsNot Nothing Then
                        If objCustItem.Min_Rate > (clsCommon.myCDecimal(gv1.Rows(ii).Cells(colAmtAfterDis).Value) / dblQty) Then
                            common.clsCommon.MyMessageBoxShow("Minimum Rate Can't be Less Then " + clsCommon.myCstr(objCustItem.Min_Rate) + " of Item " + strICode.Trim() + "( " + strIName.Trim() + " ) At Line No: " + clsCommon.myCstr(clsCommon.myCDecimal(ii + 1)))
                            Return False
                        End If
                    End If
                    If gv1.Rows(ii).Cells(colDCSSaleZeroCost).Value Then
                        gv1.Rows(ii).Cells(colRate).Value = 0
                    Else
                        If clsCommon.myCDecimal(gv1.Rows(ii).Cells(colRate).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow("Rate Can't be 0 of Item " + strICode.Trim() + "( " + strIName.Trim() + " ) At Line No: " + clsCommon.myCstr(clsCommon.myCDecimal(ii + 1)))
                            Return False
                        End If
                    End If

                    If clsCommon.myLen(strReqNo) > 0 Then
                        If Not (arrReqNo.Contains(strReqNo)) Then
                            arrReqNo.Add(strReqNo)
                        End If
                        If dblQty > dblPendingQty AndAlso clsCommon.CompairString(strSchemeItem, "No") = CompairStringResult.Equal Then
                            common.clsCommon.MyMessageBoxShow("Item " + strICode + "( " + strIName.Trim() + " ) Entered Quantity(" + clsCommon.myCstr(dblQty) + ") Can't be more Pending Quantity(" + clsCommon.myCstr(dblPendingQty) + ").At Line No: " + clsCommon.myCstr(clsCommon.myCDecimal(ii + 1)) + " ")
                            Return False
                        End If
                        If objCommonVar.IsDemoERP Then
                            If Not strProject Is Nothing Then
                                strProject = clsSNSalesOrderHead.IsValidProjectForSO(strReqNo, "")
                                If arrProjNo.Contains(strProject) And arrProjNo.Count > 0 Then
                                    common.clsCommon.MyMessageBoxShow("SO No:" + strReqNo + " and Item : " + strIName + " is not for PJC or not related to PJC At Line No: " + clsCommon.myCstr(clsCommon.myCDecimal(ii + 1)) + " ")
                                    Return False
                                Else
                                    arrProjNo.Add(strProject)
                                End If
                            End If
                        End If
                    End If
                    If IsBatchMFDEXDmandatory Then
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colBatchNo).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow("Please enter Batch No for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colManufactureDate).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow("Please enter Manufacture Date for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colExpiry).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow("Please enter Expiry Date for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                        If clsCommon.GetDateWithStartTime(clsCommon.myCDate(gv1.Rows(ii).Cells(colManufactureDate).Value)) > clsCommon.GetDateWithStartTime(clsCommon.myCDate(gv1.Rows(ii).Cells(colExpiry).Value)) Then
                            common.clsCommon.MyMessageBoxShow("Please enter Expiry Date greater than Manufacturing Date in " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                    End If
                    If Not arrICode.Contains(strICode) Then
                        arrICode.Add(strICode)
                    End If
                    If clsCommon.myLen(strICode) > 0 AndAlso clsCommon.CompairString(strSchemeItem, "No") = CompairStringResult.Equal Then
                        For jj As Integer = ii + 1 To gv1.Rows.Count - 1
                            Dim strInICode As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                            Dim strInUOM As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                            If clsCommon.CompairString(strICode, strInICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strUOM, strInUOM) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colSchemeItem).Value), "No") = CompairStringResult.Equal Then
                                common.clsCommon.MyMessageBoxShow("Item Code " + strICode + " and Unit " + strUOM + " is repeated at Row No" + clsCommon.myCstr(ii + 1) + " and " + clsCommon.myCstr(jj + 1), Me.Text)
                                Return False
                            End If
                        Next
                    End If
                    Dim dblOuterConvFac As Decimal = clsItemMaster.GetConvertionFactor(strICode, strUOM, Nothing)
                    Dim Location As String = Nothing
                    If clsCommon.myLen(clsCommon.myCstr(txtBillToLocation.Value)) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                            If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                                Location = txtSubLocation.Value
                            End If
                        Else
                            Location = txtBillToLocation.Value
                        End If
                    End If
                    Dim dblBalQty As Decimal = clsItemLocationDetails.getBalance(strICode, Location, txtDocNo.Value, txtDate.Value, Nothing, strUOM, dblMRP)
                    Dim dblEnteredQty As Decimal = dblQty
                    For jj As Integer = 0 To gv1.Rows.Count - 1
                        If ii = jj Then
                            Continue For
                        End If
                        Dim strICodeInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                        Dim strUOMInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                        Dim dblQtyInner As Decimal = clsCommon.myCDecimal(gv1.Rows(jj).Cells(colQty).Value)
                        Dim dblInnerConvFac As Decimal = clsItemMaster.GetConvertionFactor(strICodeInner, strUOMInner, Nothing)
                        If dblQtyInner > 0 AndAlso clsCommon.CompairString(strICodeInner, strICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strUOMInner, strUOM) = CompairStringResult.Equal Then
                            dblEnteredQty += dblQtyInner
                        End If
                    Next
                    dblEnteredQty = Math.Round(dblEnteredQty, 2, MidpointRounding.ToEven)
                    If dblEnteredQty > dblBalQty Then ' AndAlso clsCommon.CompairString(strSchemeItem, "No") = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow("Item - " + strICode + Environment.NewLine + "Entered Quantity - " + clsCommon.myCstr(dblEnteredQty) + " and Balance Quantity - " + clsCommon.myCstr(dblBalQty), Me.Text)
                        Return False
                    End If
                End If
                If clsCommon.myCBool(gv1.Rows(ii).Cells(colIsSerialseItem).Value) Then
                    Dim arrSerailNo As List(Of clsSerializeInvenotry) = TryCast(gv1.Rows(ii).Tag, List(Of clsSerializeInvenotry))
                    If arrSerailNo Is Nothing OrElse dblQty <> arrSerailNo.Count Then
                        common.clsCommon.MyMessageBoxShow("Please provice serial no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                End If
                Dim qry As String = " select Is_Batch_Item from TSPL_ITEM_MASTER where Item_Code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value) & "'"
                Dim isBatchAvtive As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                If dblQty > 0 AndAlso isBatchAvtive = "1" Then
                    If RunBatchFifowise = 1 Then
                        gv1.CurrentRow = gv1.Rows(ii)
                        OpenBatchItem(False)
                    End If
                    Dim arrBatchNo As List(Of clsBatchInventory) = TryCast(gv1.Rows(ii).Cells(colICode).Tag, List(Of clsBatchInventory))
                    If arrBatchNo Is Nothing Then
                        Throw New Exception("Please provide Batch no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                    Else
                        Dim tQty As Decimal = 0
                        For Each objBatch As clsBatchInventory In arrBatchNo
                            tQty += objBatch.Qty
                        Next
                        If tQty <> dblQty Then
                            Throw New Exception("Item : " + strICode + " Entered Qty " + clsCommon.myCstr(dblQty) + Environment.NewLine + "And Batchwise Qty " + clsCommon.myCstr(tQty) + " . At Line No" + clsCommon.myCstr(ii + 1))
                        End If
                    End If
                End If
            Next
            For ii As Integer = 0 To gvAC.Rows.Count - 1
                If clsCommon.myLen(gvAC.Rows(ii).Cells(colACCode).Value) > 0 Then
                    For jj As Integer = 0 To gvAC.Rows.Count - 1
                        If ii = jj Then
                            Continue For
                        End If
                        If clsCommon.CompairString(clsCommon.myCstr(gvAC.Rows(ii).Cells(colACCode).Value), clsCommon.myCstr(gvAC.Rows(jj).Cells(colACCode).Value)) = CompairStringResult.Equal Then
                            common.clsCommon.MyMessageBoxShow("Additional Charges: " + clsCommon.myCstr(gvAC.Rows(ii).Cells(colACCode).Value) + "Repeated at Row No " + clsCommon.myCstr(ii + 1) + " and " + clsCommon.myCstr(jj + 1) + "")
                            Return False
                        End If
                    Next
                End If
            Next
            clsSNSalesOrderHead.IsValidCustomer(arrReqNo, txtVendorNo.Value)
            If clsERPFuncationality.GetGSTStatus(txtDate.Value) Then
                clsItemMaster.isItemTaxableOrNonTaxable(arrICode, chkTaxable.Checked)
            Else
                If clsLocation.isLocatinExcisable(txtBillToLocation.Value) Then
                    Dim qry As String = "select Type  from TSPL_TAX_MASTER where Tax_Code ='" + clsCommon.myCstr(gv2.Rows(0).Cells(colTTaxAutCode).Value) + "' "
                    If Not clsCommon.CompairString("E", clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))) = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow("Tax should be excisable for excisable location", Me.Text)
                        Return False
                    End If
                End If
                Dim intx As Integer = clsItemMaster.isItemOfSameExempted(arrICode)
                If Not (intx = arrICode.Count OrElse intx = 0) Then
                    Throw New Exception("All item should be of Tax Exempted or not")
                Else
                    If intx > 0 AndAlso clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                        Dim qry As String = "select 1 from TSPL_Tax_Group_Master where Tax_Group_Code='" + txtTaxGroup.Value + "' and Is_Tax_Exempted=1"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Please select of Tax Exempted")
                        End If
                    End If
                End If
                isTaxExempted = False
                If intx > 0 Then
                    isTaxExempted = True
                End If
            End If
            'Bhole Baba- Credit Limit check
            Dim CreditLimit As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("Select TSPL_CUSTOMER_MASTER.Credit_Limit from TSPL_CUSTOMER_VENDOR_MAPPING left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.cust_code=TSPL_CUSTOMER_VENDOR_MAPPING.cust_code left outer join TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code WHERE (TSPL_VENDOR_MASTER.Is_Chilling_vendor=1 OR TSPL_VENDOR_MASTER.IsEmployee=1 OR TSPL_VENDOR_MASTER.Form_Type='PTM') AND TSPL_CUSTOMER_MASTER.Cust_Code='" + txtVendorNo.Value + "'"))
            If CreditLimit > 0 Then
                Dim SaleAmt As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select sum(Total_Amt) As Total_Amt  from TSPL_SD_SHIPMENT_HEAD where TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' and Customer_Code='" + txtVendorNo.Value + "' and convert(date,TSPL_SD_SHIPMENT_HEAD.document_date,103)>=convert(date,DATEADD(mm, DATEDIFF(mm, 0, convert(date,'" + txtDate.Value + "',103)), 0),103) and convert(date,TSPL_SD_SHIPMENT_HEAD.document_date,103)<=convert(date,DATEADD (dd, -1, DATEADD(mm, DATEDIFF(mm, 0, convert(date,'" + txtDate.Value + "',103)) + 1, 0)),103) and TSPL_SD_SHIPMENT_HEAD.document_code <> '" + txtDocNo.Value + "' "))
                If (SaleAmt + clsCommon.myCDecimal(lblTotRAmt.Text)) > CreditLimit Then
                    Throw New Exception("Sale limit exceed for this month.")
                End If
            End If
            UcCustomFields1.AllowToSave()
            UcAttachment1.AllowToSave()
            If GSTStatus = True AndAlso chkTaxable.Checked Then
                clsLocationWiseTax.IsValidTaxGroup(txtTaxGroup.Value, txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value, Nothing)
            End If
            For ii As Integer = 0 To gv1.Rows.Count - 1
                If gv1.Rows(ii).Cells(colDCSSaleZeroCost).Value Then
                    gv1.Rows(ii).Cells(colRate).Value = 0
                End If
            Next

            Dim cgstAmount As Decimal = 0
            Dim sgstAmount As Decimal = 0
            For Each grow As GridViewRowInfo In gv2.Rows
                Dim taxtype As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select type from tspl_Tax_master where Tax_Code='" + clsCommon.myCstr(grow.Cells(colTTaxAutCode).Value) + "' "))
                Dim amt As Decimal = clsCommon.myCDecimal(grow.Cells(colTTaxAmt).Value)
                If clsCommon.CompairString(taxtype, "CGST") = CompairStringResult.Equal Then
                    cgstAmount = amt
                ElseIf clsCommon.CompairString(taxtype, "SGST") = CompairStringResult.Equal Then
                    sgstAmount = amt
                End If
            Next
            If cgstAmount <> sgstAmount Then
                Throw New Exception("CGST " & cgstAmount & "  and SGST " & sgstAmount & " Amount must be same")
            End If

            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveData(False)
    End Sub
    Function SaveData(ByVal ChekPostBtn As Boolean) As Boolean
        Try
            If (AllowToSave(ChekPostBtn)) Then
                Dim obj As New clsMCCMaterialSale()
                obj.Sale_Invoice_No = txtInvoiceNo.Text
                obj.DCS_Price_Code = FndPriceCode.Value
                obj.Sale_Route_Code = txtSaleRoute.Value
                'obj.WayBillNo = txtWayBillno.Text
                ' obj.WayBillDate = txtWaybillDate.Value
                obj.Road_Permit_No = txtRoadPermitNo.Text
                obj.Total_Comm_Amt = lblCommAmt.Text
                obj.Invoice_Type = ddlInvoiceType.SelectedValue
                obj.Price_Code = txtPriceCode.Text
                obj.HeadDisc_Per = clsCommon.myCDecimal(txtDiscPer.Text)
                obj.IS_TCS = IIf(chkisTCS.Checked, 1, 0)
                obj.Deduction = clsCommon.myCstr(cboDeductionType.SelectedValue)
                obj.IsEwayBill = IIf(chkIsEwayBill.Checked, 1, 0)
                obj.No_Transporter = IIf(chkNoTranspoter.Checked, 1, 0)
                obj.Transport_Id = txtTransporterCode.Value
                obj.Transporter_Name = lblTransporterName.Text
                If obj.HeadDisc_Per > 0 Then
                    If MultiplySubsidyWithQuantity Then
                        obj.HeadDisc_PerAmt = obj.TotalSubsidyDisAmt
                    Else
                        obj.HeadDisc_PerAmt = lblInvoiceDiscAmt.Text
                    End If
                    obj.HeadDisc_Amt = 0
                Else
                    If MultiplySubsidyWithQuantity Then
                        obj.HeadDisc_Amt = txtDiscAmt.Text
                    Else
                        obj.HeadDisc_Amt = lblInvoiceDiscAmt.Text
                    End If
                    obj.HeadDisc_PerAmt = 0
                End If
                obj.RateDiff_Per = txtRatePer.Text
                If obj.RateDiff_Per > 0 Then
                    If rbtnTotalAmt.IsChecked Then
                        obj.Rate_Diff_Amount_Type = 1
                    ElseIf rbtnBasicAmt.IsChecked Then
                        obj.Rate_Diff_Amount_Type = 2
                    End If
                End If
                If chkRateDiffRate.IsChecked Then
                    obj.Is_Rate_Diff_Per_Amt = 1
                ElseIf chkRateDiffAmt.IsChecked Then
                    obj.Is_Rate_Diff_Per_Amt = 2
                End If
                obj.RateDiff_Amt = clsCommon.myCDecimal(txtRateAmt.Text)
                obj.Gross_Amount = lblGrossAmount.Text
                '-----------------richa 27/06/2014 Ticket No .BM00000002982-------  
                If clsCommon.myCDecimal(txtMannaulInvoiceNo.Value) > 0 Then
                    obj.Mannual_Invoice_No = txtMannaulInvoiceNo.Value
                Else
                    obj.InvoiceManualNowithPrefix = TxtInvoiceManualNoWithPrefix.Text
                End If
                obj.Transporter_Commission_TotalAmt = clsCommon.myCDecimal(txtTPTAmt.Text)
                '----------------------------------------------------------------
                obj.Document_Code = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Customer_Code = txtVendorNo.Value
                obj.Customer_Name = lblVendorName.Text
                obj.Ref_No = txtRefNo.Text
                obj.Inv_Date = clsCommon.GetPrintDate(dtpInvoice.Value, "dd/MMM/yyyy")
                obj.Challan_Date = clsCommon.GetPrintDate(dtpChallan.Value, "dd/MMM/yyyy")
                obj.Total_Tax_Amt = clsCommon.myCDecimal(lblTaxAmt.Text)
                obj.Inv_No = txtInvNo.Text
                obj.Bill_To_Location = txtBillToLocation.Value
                obj.Sub_Location_code = txtSubLocation.Value
                'obj.Ship_To_Location = txtShipToLocation.Value
                obj.Comments = txtComment.Text
                obj.On_Hold = chkOnHold.Checked
                obj.Description = txtDesc.Text
                obj.Tax_Group = txtTaxGroup.Value
                obj.Is_CashSale = IIf(chkcashsale.Checked, "Y", "N")
                obj.Is_Apply_TPT = IIf(chkApplyTPT.Checked, "1", "0")
                If chkApplyTPT.Checked Then
                    obj.TPT_Vendor = txtTPTVendor.Value
                End If
                obj.Is_Add_TPT = chkAddTPT.Checked
                obj.Recommended_By = txtRecommBy.Text
                obj.Is_Internal = chkInternal.Checked
                obj.PROJECT_ID = fndProject.Value
                obj.No_Of_Instalment = txtNoOfInsallment.Value
                obj.Exclude_KKF_And_Mandi = chkExcludeKKFMandi.Checked
                If (gv2.Rows.Count > 0) Then
                    obj.TAX1 = clsCommon.myCstr(gv2.Rows(0).Cells(colTTaxAutCode).Value)
                    obj.TAX1_Rate = clsCommon.myCDecimal(gv2.Rows(0).Cells(colTTaxRate).Value)
                    obj.TAX1_Base_Amt = clsCommon.myCDecimal(gv2.Rows(0).Cells(colTBaseAmt).Value)
                    obj.TAX1_Amt = clsCommon.myCDecimal(gv2.Rows(0).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 1) Then
                    obj.TAX2 = clsCommon.myCstr(gv2.Rows(1).Cells(colTTaxAutCode).Value)
                    obj.TAX2_Rate = clsCommon.myCDecimal(gv2.Rows(1).Cells(colTTaxRate).Value)
                    obj.TAX2_Base_Amt = clsCommon.myCDecimal(gv2.Rows(1).Cells(colTBaseAmt).Value)
                    obj.TAX2_Amt = clsCommon.myCDecimal(gv2.Rows(1).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 2) Then
                    obj.TAX3 = clsCommon.myCstr(gv2.Rows(2).Cells(colTTaxAutCode).Value)
                    obj.TAX3_Rate = clsCommon.myCDecimal(gv2.Rows(2).Cells(colTTaxRate).Value)
                    obj.TAX3_Base_Amt = clsCommon.myCDecimal(gv2.Rows(2).Cells(colTBaseAmt).Value)
                    obj.TAX3_Amt = clsCommon.myCDecimal(gv2.Rows(2).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 3) Then
                    obj.TAX4 = clsCommon.myCstr(gv2.Rows(3).Cells(colTTaxAutCode).Value)
                    obj.TAX4_Rate = clsCommon.myCDecimal(gv2.Rows(3).Cells(colTTaxRate).Value)
                    obj.TAX4_Base_Amt = clsCommon.myCDecimal(gv2.Rows(3).Cells(colTBaseAmt).Value)
                    obj.TAX4_Amt = clsCommon.myCDecimal(gv2.Rows(3).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 4) Then
                    obj.TAX5 = clsCommon.myCstr(gv2.Rows(4).Cells(colTTaxAutCode).Value)
                    obj.TAX5_Rate = clsCommon.myCDecimal(gv2.Rows(4).Cells(colTTaxRate).Value)
                    obj.TAX5_Base_Amt = clsCommon.myCDecimal(gv2.Rows(4).Cells(colTBaseAmt).Value)
                    obj.TAX5_Amt = clsCommon.myCDecimal(gv2.Rows(4).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 5) Then
                    obj.TAX6 = clsCommon.myCstr(gv2.Rows(5).Cells(colTTaxAutCode).Value)
                    obj.TAX6_Rate = clsCommon.myCDecimal(gv2.Rows(5).Cells(colTTaxRate).Value)
                    obj.TAX6_Base_Amt = clsCommon.myCDecimal(gv2.Rows(5).Cells(colTBaseAmt).Value)
                    obj.TAX6_Amt = clsCommon.myCDecimal(gv2.Rows(5).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 6) Then
                    obj.TAX7 = clsCommon.myCstr(gv2.Rows(6).Cells(colTTaxAutCode).Value)
                    obj.TAX7_Rate = clsCommon.myCDecimal(gv2.Rows(6).Cells(colTTaxRate).Value)
                    obj.TAX7_Base_Amt = clsCommon.myCDecimal(gv2.Rows(6).Cells(colTBaseAmt).Value)
                    obj.TAX7_Amt = clsCommon.myCDecimal(gv2.Rows(6).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 7) Then
                    obj.TAX8 = clsCommon.myCstr(gv2.Rows(7).Cells(colTTaxAutCode).Value)
                    obj.TAX8_Rate = clsCommon.myCDecimal(gv2.Rows(7).Cells(colTTaxRate).Value)
                    obj.TAX8_Base_Amt = clsCommon.myCDecimal(gv2.Rows(7).Cells(colTBaseAmt).Value)
                    obj.TAX8_Amt = clsCommon.myCDecimal(gv2.Rows(7).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 8) Then
                    obj.TAX9 = clsCommon.myCstr(gv2.Rows(8).Cells(colTTaxAutCode).Value)
                    obj.TAX9_Rate = clsCommon.myCDecimal(gv2.Rows(8).Cells(colTTaxRate).Value)
                    obj.TAX9_Base_Amt = clsCommon.myCDecimal(gv2.Rows(8).Cells(colTBaseAmt).Value)
                    obj.TAX9_Amt = clsCommon.myCDecimal(gv2.Rows(8).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 9) Then
                    obj.TAX10 = clsCommon.myCstr(gv2.Rows(9).Cells(colTTaxAutCode).Value)
                    obj.TAX10_Rate = clsCommon.myCDecimal(gv2.Rows(9).Cells(colTTaxRate).Value)
                    obj.TAX10_Base_Amt = clsCommon.myCDecimal(gv2.Rows(9).Cells(colTBaseAmt).Value)
                    obj.TAX10_Amt = clsCommon.myCDecimal(gv2.Rows(9).Cells(colTTaxAmt).Value)
                End If
                obj.Terms_Code = txtTermCode.Value
                obj.Due_Date = txtDueDate.Value
                obj.Discount_Base = clsCommon.myCDecimal(lblAmtWithDiscount.Text)
                obj.Discount_Amt = clsCommon.myCDecimal(lblDiscountAmt.Text)
                obj.Amount_Less_Discount = clsCommon.myCDecimal(lblAmtAfterDiscount.Text)
                obj.Total_Amt = clsCommon.myCDecimal(lblTotRAmt.Text)
                obj.RoundOffAmount = clsCommon.myCDecimal(txtRoundOff.Text)
                obj.Carrier = txtCarrier.Text
                obj.GRNo = txtGRNo.Text
                obj.GENo = txtGENo.Text
                If txtGEDate.Checked Then
                    obj.GEDate = txtGEDate.Value
                End If
                obj.Item_Type = clsCommon.myCstr(cboItemType.SelectedValue)
                obj.Dept = txtDept.Value
                obj.Dept_Desc = lblDept.Text
                If (gvAC.Rows.Count > 0) Then
                    If clsCommon.myLen(gvAC.Rows(0).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code1 = clsCommon.myCstr(gvAC.Rows(0).Cells(colACCode).Value)
                        obj.Add_Charge_Name1 = clsCommon.myCstr(gvAC.Rows(0).Cells(colACName).Value)
                        obj.Add_Charge_Amt1 = clsCommon.myCDecimal(gvAC.Rows(0).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 1) Then
                    If clsCommon.myLen(gvAC.Rows(1).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code2 = clsCommon.myCstr(gvAC.Rows(1).Cells(colACCode).Value)
                        obj.Add_Charge_Name2 = clsCommon.myCstr(gvAC.Rows(1).Cells(colACName).Value)
                        obj.Add_Charge_Amt2 = clsCommon.myCDecimal(gvAC.Rows(1).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 2) Then
                    If clsCommon.myLen(gvAC.Rows(2).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code3 = clsCommon.myCstr(gvAC.Rows(2).Cells(colACCode).Value)
                        obj.Add_Charge_Name3 = clsCommon.myCstr(gvAC.Rows(2).Cells(colACName).Value)
                        obj.Add_Charge_Amt3 = clsCommon.myCDecimal(gvAC.Rows(2).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 3) Then
                    If clsCommon.myLen(gvAC.Rows(3).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code4 = clsCommon.myCstr(gvAC.Rows(3).Cells(colACCode).Value)
                        obj.Add_Charge_Name4 = clsCommon.myCstr(gvAC.Rows(3).Cells(colACName).Value)
                        obj.Add_Charge_Amt4 = clsCommon.myCDecimal(gvAC.Rows(3).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 4) Then
                    If clsCommon.myLen(gvAC.Rows(4).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code5 = clsCommon.myCstr(gvAC.Rows(4).Cells(colACCode).Value)
                        obj.Add_Charge_Name5 = clsCommon.myCstr(gvAC.Rows(4).Cells(colACName).Value)
                        obj.Add_Charge_Amt5 = clsCommon.myCDecimal(gvAC.Rows(4).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 5) Then
                    If clsCommon.myLen(gvAC.Rows(5).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code6 = clsCommon.myCstr(gvAC.Rows(5).Cells(colACCode).Value)
                        obj.Add_Charge_Name6 = clsCommon.myCstr(gvAC.Rows(5).Cells(colACName).Value)
                        obj.Add_Charge_Amt6 = clsCommon.myCDecimal(gvAC.Rows(5).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 6) Then
                    If clsCommon.myLen(gvAC.Rows(6).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code7 = clsCommon.myCstr(gvAC.Rows(6).Cells(colACCode).Value)
                        obj.Add_Charge_Name7 = clsCommon.myCstr(gvAC.Rows(6).Cells(colACName).Value)
                        obj.Add_Charge_Amt7 = clsCommon.myCDecimal(gvAC.Rows(6).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 7) Then
                    If clsCommon.myLen(gvAC.Rows(7).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code8 = clsCommon.myCstr(gvAC.Rows(7).Cells(colACCode).Value)
                        obj.Add_Charge_Name8 = clsCommon.myCstr(gvAC.Rows(7).Cells(colACName).Value)
                        obj.Add_Charge_Amt8 = clsCommon.myCDecimal(gvAC.Rows(7).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 8) Then
                    If clsCommon.myLen(gvAC.Rows(8).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code9 = clsCommon.myCstr(gvAC.Rows(8).Cells(colACCode).Value)
                        obj.Add_Charge_Name9 = clsCommon.myCstr(gvAC.Rows(8).Cells(colACName).Value)
                        obj.Add_Charge_Amt9 = clsCommon.myCDecimal(gvAC.Rows(8).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 9) Then
                    If clsCommon.myLen(gvAC.Rows(9).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code10 = clsCommon.myCstr(gvAC.Rows(9).Cells(colACCode).Value)
                        obj.Add_Charge_Name10 = clsCommon.myCstr(gvAC.Rows(9).Cells(colACName).Value)
                        obj.Add_Charge_Amt10 = clsCommon.myCDecimal(gvAC.Rows(9).Cells(colACAmount).Value)
                    End If
                End If
                obj.Total_Add_Charge = clsCommon.myCDecimal(lblAddCharges.Text)
                If rbtnTaxCalAutomatic.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                ElseIf rbtnTaxCalManual.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                End If
                obj.Is_Create_Auto_Invoice = 1
                obj.Is_Create_Auto_Receipt = chkCreateAutoInvoice.Checked AndAlso chkCreateAutoReceipt.Checked
                obj.PROJECT_ID = fndProject.Value
                obj.Is_Taxable = chkTaxable.Checked
                If chkcashsale.Checked Then
                    obj.Receipt_No = txtReceiptNo.Value
                    obj.ReceiptAmt = clsCommon.myCDecimal(lblReceiptAmt.Text)
                    obj.Payment_Terms = cmbPaymentType.Text
                    obj.Bank_Code = txtBankCode.Value
                End If
                obj.VehicleNo = txtVehicleNo.Text
                obj.ReceiverName = txtReceiverName.Text
                obj.TotalSubsidyAmt = clsCommon.myCDecimal(lblTotalSubsidy.Text)
                obj.TotalSubsidyDisAmt = clsCommon.myCDecimal(lblTotalDisSubsidy.Text)
                obj.Arr = New List(Of clsMCCMaterialSaleDetail)
                If CreateAutoMCCPriceChat Then

                    Dim qry As String = Nothing
                    Dim count As Integer = 0
                    qry = "select count(*) from TSPL_MCC_RATE_UPLOADER_master   left join TSPL_MCC_RATE_UPLOADER_Detail on TSPL_MCC_RATE_UPLOADER_Detail.Code=TSPL_MCC_RATE_UPLOADER_MASTER.code where Item_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) + "'  and TSPL_MCC_RATE_UPLOADER_Detail.RATE_UOM='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) + "' and convert(date,TSPL_MCC_RATE_UPLOADER_Master.Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtDate.Value) + "',103) and convert(date,TSPL_MCC_RATE_UPLOADER_Master.Effective_date,103) >=convert(date,'" + clsCommon.GetPrintDate(txtDate.Value) + "',103) "
                    count = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry))
                    If count <= 0 Then
                        Dim objPrice As New clsMCCMaterailSalePriceChat
                        objPrice.DOCDate = clsCommon.GetPrintDate(txtDate.Value)
                        objPrice.Effective_Date = clsCommon.GetPrintDate(txtDate.Value.AddDays(365))
                        'Dim qry As String = Nothing
                        If AllowPlandDeptMCCLocation Then
                            qry = "select Location_Code AS Code, Location_Desc as Name  from TSPL_LOCATION_MASTER where Is_Sub_Location = 'N' AND Location_Category <> 'MCC' and GIT_Type  <> 'Y' order by Code "
                        Else
                            qry = "select MCC_Code as Code ,MCC_NAME as Name from TSPL_MCC_MASTER "
                        End If
                        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
                        Dim MccRate As New ArrayList()
                        If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                            For Each dr As DataRow In dt1.Rows
                                MccRate.Add(dr("Code"))
                            Next
                        End If
                        objPrice.ArrMCCRate = MccRate
                        objPrice.Arr = New List(Of clsMCCMaterailSalePriceChatDetail)
                        For Each gorw As GridViewRowInfo In gv1.Rows
                            If clsCommon.myCDecimal(gorw.Cells(colRate).Value) > 0 Then
                                Dim objTrprice As New clsMCCMaterailSalePriceChatDetail()
                                objTrprice.Item_Code = clsCommon.myCstr(gorw.Cells(colICode).Value)
                                objTrprice.Rate_UOM = clsCommon.myCstr(gorw.Cells(colUnit).Value)
                                objTrprice.Price = clsCommon.myCDecimal(gorw.Cells(colRate).Value)
                                objPrice.Arr.Add(objTrprice)
                            End If

                        Next
                        objPrice.SaveData(objPrice, True)

                    End If


                End If
                Dim isSubLocation As Boolean = False
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsSubLocationWise from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "'", Nothing)), "Y") = CompairStringResult.Equal Then
                    isSubLocation = True
                Else
                    isSubLocation = False
                End If
                For Each grow As GridViewRowInfo In gv1.Rows
                    Dim Rate_Mcc_Item As Decimal
                    If ApplySaleRouteInDCSSale Then
                        Rate_Mcc_Item = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(grow.Cells(colICode).Value), clsCommon.myCstr(grow.Cells(colUnit).Value), txtDate.Value, Nothing, FndPriceCode.Value)
                    Else
                        Rate_Mcc_Item = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(grow.Cells(colICode).Value), clsCommon.myCstr(grow.Cells(colUnit).Value), txtDate.Value)
                    End If
                    Dim objTr As New clsMCCMaterialSaleDetail()
                    If Rate_Mcc_Item <> clsCommon.myCDecimal(grow.Cells(colRate).Value) Then
                        obj.Rate_Status = 2
                    End If
                    objTr.Line_No = clsCommon.myCDecimal(grow.Cells(colLineNo).Value)
                    objTr.Row_Type = clsCommon.myCstr(grow.Cells(colRowType).Value)
                    objTr.Item_Code = clsCommon.myCstr(grow.Cells(colICode).Value)
                    objTr.Item_Desc = clsCommon.myCstr(grow.Cells(colIName).Value)
                    objTr.Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                    objTr.Free_Qty = clsCommon.myCDecimal(grow.Cells(colFreeQty).Value)
                    objTr.Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                    If ConvertIntoBillingUOM AndAlso clsCommon.myLen(objTr.Item_Code) > 0 Then
                        Dim Billing_UOM As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select UOM_Code from TSPL_ITEM_UOM_DETAIL where Item_Code='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "' and Billing_UOM=1"))
                        If clsCommon.myLen(Billing_UOM) > 0 Then
                            objTr.Billing_Unit_code = Billing_UOM
                        Else

                            Throw New Exception("Please Map Billing UOM for item [" & clsCommon.myCstr(grow.Cells(colIName).Value) & "]")
                        End If
                        Dim BillingUOMConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.Billing_UOM=1 "))
                        Dim BillingItemConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(grow.Cells(colUnit).Value) & "' "))
                        If BillingUOMConvFactor > 0 AndAlso BillingItemConvFactor > 0 Then
                            Dim DispatchQty As Decimal = clsCommon.myCDecimal(grow.Cells(colQty).Value) * BillingItemConvFactor
                            objTr.Billing_Qty = Math.Ceiling(DispatchQty / BillingUOMConvFactor)
                        End If
                    Else
                        objTr.Billing_Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                        objTr.Billing_Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                    End If
                    objTr.OrgUnit_code = clsCommon.myCstr(grow.Cells(colOrgUnit).Value)
                    objTr.Order_Code = clsCommon.myCstr(grow.Cells(colOrderNo).Value)
                    objTr.Item_Cost = clsCommon.myCDecimal(grow.Cells(colRate).Value)
                    objTr.Amount = clsCommon.myCDecimal(grow.Cells(colAmt).Value)
                    objTr.Disc_Per = clsCommon.myCDecimal(grow.Cells(colDisPer).Value)
                    objTr.Disc_Amt = clsCommon.myCDecimal(grow.Cells(colDisAmt).Value)
                    objTr.Disc_Per_Unit = clsCommon.myCDecimal(grow.Cells(colDisPerUnit).Value)
                    objTr.Disc_Unit_Amt = clsCommon.myCDecimal(grow.Cells(colDisUnitAmt).Value)
                    objTr.Amt_Less_Discount = clsCommon.myCDecimal(grow.Cells(colAmtAfterDis).Value)
                    objTr.TAX1 = clsCommon.myCstr(grow.Cells(colTax1).Value)
                    objTr.TAX1_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt1).Value)
                    objTr.TAX1_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate1).Value)
                    objTr.TAX1_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt1).Value)
                    objTr.TAX2 = clsCommon.myCstr(grow.Cells(colTax2).Value)
                    objTr.TAX2_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt2).Value)
                    objTr.TAX2_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate2).Value)
                    objTr.TAX2_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt2).Value)
                    objTr.TAX3 = clsCommon.myCstr(grow.Cells(colTax3).Value)
                    objTr.TAX3_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt3).Value)
                    objTr.TAX3_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate3).Value)
                    objTr.TAX3_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt3).Value)
                    objTr.TAX4 = clsCommon.myCstr(grow.Cells(colTax4).Value)
                    objTr.TAX4_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt4).Value)
                    objTr.TAX4_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate4).Value)
                    objTr.TAX4_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt4).Value)
                    objTr.TAX5 = clsCommon.myCstr(grow.Cells(colTax5).Value)
                    objTr.TAX5_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt5).Value)
                    objTr.TAX5_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate5).Value)
                    objTr.TAX5_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt5).Value)
                    objTr.TAX6 = clsCommon.myCstr(grow.Cells(colTax6).Value)
                    objTr.TAX6_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt6).Value)
                    objTr.TAX6_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate6).Value)
                    objTr.TAX6_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt6).Value)
                    objTr.TAX7 = clsCommon.myCstr(grow.Cells(colTax7).Value)
                    objTr.TAX7_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt7).Value)
                    objTr.TAX7_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate7).Value)
                    objTr.TAX7_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt7).Value)
                    objTr.TAX8 = clsCommon.myCstr(grow.Cells(colTax8).Value)
                    objTr.TAX8_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt8).Value)
                    objTr.TAX8_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate8).Value)
                    objTr.TAX8_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt8).Value)
                    objTr.TAX9 = clsCommon.myCstr(grow.Cells(colTax9).Value)
                    objTr.TAX9_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt9).Value)
                    objTr.TAX9_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate9).Value)
                    objTr.TAX9_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt9).Value)
                    objTr.TAX10 = clsCommon.myCstr(grow.Cells(colTax10).Value)
                    objTr.TAX10_Base_Amt = clsCommon.myCDecimal(grow.Cells(colTaxBaseAmt10).Value)
                    objTr.TAX10_Rate = clsCommon.myCDecimal(grow.Cells(colTaxRate10).Value)
                    objTr.TAX10_Amt = clsCommon.myCDecimal(grow.Cells(colTaxAmt10).Value)
                    objTr.Total_Tax_Amt = clsCommon.myCDecimal(grow.Cells(colTotTaxAmt).Value)
                    objTr.Item_Net_Amt = clsCommon.myCDecimal(grow.Cells(colAmtAfterTax).Value)
                    objTr.Location = IIf(isSubLocation, txtSubLocation.Value, txtBillToLocation.Value) 'clsCommon.myCstr(grow.Cells(colLocationCode).Value)
                    'If objTr.Location Is Nothing OrElse clsCommon.myLen(objTr.Location) = 0 Then
                    '    objTr.Location = txtBillToLocation.Value
                    'End If
                    objTr.MRP = clsCommon.myCDecimal(grow.Cells(colMRP).Value)
                    objTr.Scheme_Applicable = clsCommon.myCstr(grow.Cells(colSchemeApplicable).Value)
                    objTr.Scheme_Code = clsCommon.myCstr(grow.Cells(colFromSchemeCode).Value)
                    objTr.Scheme_Item = clsCommon.myCstr(grow.Cells(colSchemeItem).Value)
                    objTr.Item_Tax = clsCommon.myCDecimal(grow.Cells(colTotTaxAmt).Value)
                    objTr.Total_MRP_Amt = clsCommon.myCDecimal(grow.Cells(colTotalMRP).Value)
                    objTr.Total_Basic_Amt = clsCommon.myCDecimal(grow.Cells(colTotalBasicAmount).Value)
                    objTr.Total_Disc_Amt = clsCommon.myCDecimal(grow.Cells(colTotalDiscountAmount).Value)
                    objTr.Cust_Discount = clsCommon.myCDecimal(grow.Cells(colcustDiscount).Value)
                    objTr.Total_Cust_Discount = clsCommon.myCDecimal(grow.Cells(colTotalCustDiscount).Value)
                    objTr.ActualRate = clsCommon.myCDecimal(grow.Cells(colActualCost).Value)
                    objTr.Cust_DiscountQty = clsCommon.myCDecimal(grow.Cells(ColCustDiscountQty).Value)
                    objTr.Price_Date = clsCommon.myCDate(grow.Cells(colPriceDateColumn).Value)
                    objTr.Price_code = clsCommon.myCstr(grow.Cells(colPriceCOde).Value)
                    objTr.Abatement_Per = clsCommon.myCDecimal(grow.Cells(colAbatementPer).Value)
                    objTr.Abatement_Amt = clsCommon.myCDecimal(grow.Cells(colAbatementAmount).Value)
                    objTr.FOC_Item = clsCommon.myCDecimal(grow.Cells(ColFOC).Value)
                    objTr.Item_Weight = clsCommon.myCDecimal(grow.Cells(colItemWeight).Value)
                    objTr.Conv_Factor = clsCommon.myCDecimal(grow.Cells(colConvF).Value)
                    objTr.TotalItem_Weight = clsCommon.myCDecimal(grow.Cells(colTotItemWt).Value)
                    objTr.Markup_On = clsCommon.myCstr(grow.Cells(colMarkupOn).Value)
                    objTr.Markup_Percent = clsCommon.myCDecimal(grow.Cells(colMarkUpPercentage).Value)
                    objTr.Landing_Cost = clsCommon.myCDecimal(grow.Cells(colLandingCost).Value)
                    objTr.CustDiscPer = clsCommon.myCDecimal(grow.Cells(colCustDiscPercentage).Value)
                    objTr.HeadDiscAmt = clsCommon.myCDecimal(grow.Cells(colHeadDiscamt).Value)
                    objTr.CasdDiscScheme_Code = clsCommon.myCstr(grow.Cells(colCashDiscSchemeCode).Value)
                    objTr.Purchase_Cost = clsCommon.myCDecimal(grow.Cells(colPurCost).Value)
                    objTr.OrgRate = clsCommon.myCDecimal(grow.Cells(colOrgCost).Value)
                    objTr.PrincipleCode = clsCommon.myCstr(grow.Cells(colPricipleCode).Value)
                    objTr.PrincipleDesc = clsCommon.myCstr(grow.Cells(colPricipleDesc).Value)
                    objTr.vendor_code = clsCommon.myCstr(grow.Cells(colvendorCode).Value)
                    objTr.vendor_desc = clsCommon.myCstr(grow.Cells(colvendorDesc).Value)
                    objTr.HeadDiscPer = clsCommon.myCDecimal(grow.Cells(colHeaDDisPer).Value)
                    objTr.HeadDiscPerAmt = clsCommon.myCDecimal(grow.Cells(colHeadDisPerAmt).Value)
                    objTr.Batch_No = clsCommon.myCstr(grow.Cells(colBatchNo).Value)
                    objTr.Bin_No = clsCommon.myCstr(grow.Cells(colBinNo).Value)
                    If clsCommon.myLen(grow.Cells(colExpiry).Value) > 0 Then
                        objTr.Expiry_Date = clsCommon.myCDate(grow.Cells(colExpiry).Value, "dd-MM-yyyy")
                    End If
                    If clsCommon.myLen(grow.Cells(colManufactureDate).Value) > 0 Then
                        objTr.MFG_Date = clsCommon.myCDate(grow.Cells(colManufactureDate).Value)
                    End If
                    objTr.Specification = clsCommon.myCstr(grow.Cells(colSpecification).Value)
                    objTr.Remarks = clsCommon.myCstr(grow.Cells(colRemarks).Value)
                    objTr.Is_Mannual_Amt = clsCommon.myCDecimal(grow.Cells(colIsMannualAmt).Value)
                    objTr.Balance_Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                    objTr.Commission_Rate = clsCommon.myCDecimal(grow.Cells(colCommRate).Value)
                    objTr.Commission_Party = clsCommon.myCstr(grow.Cells(ColCommParty).Value)
                    objTr.Commission_Amt = clsCommon.myCDecimal(grow.Cells(ColCommAmt).Value)
                    objTr.Amt_Less_Commission = clsCommon.myCDecimal(grow.Cells(ColAmtAfterCOmm).Value)
                    objTr.ItemwiseTaxCode = clsCommon.myCstr(grow.Cells(colItemwiseTaxCode).Value)
                    objTr.Transporter_Commission_Rate = clsCommon.myCDecimal(grow.Cells(ColTPTRate).Value)
                    objTr.REF_TPT_PK_ID = clsCommon.myCDecimal(grow.Cells(ColTPTRate).Tag)
                    objTr.Transporter_Commission_Amt = clsCommon.myCDecimal(grow.Cells(ColTPTAmt).Value)
                    objTr.arrSrItem = TryCast(grow.Tag, List(Of clsSerializeInvenotry))
                    objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))  ' Change By prabhakar
                    If (clsCommon.myLen(objTr.Item_Code) > 0) Then
                        obj.Arr.Add(objTr)
                    End If
                Next
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Fill at list one Item", Me.Text)
                    Return False
                End If
                ''For Custom Fields
                obj.Form_ID = MyBase.Form_ID
                obj.arrCustomFields = New List(Of clsCustomFieldValues)
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.GetData(obj.arrCustomFields)
                End If
                If MyBase.ArrDetailFields IsNot Nothing AndAlso MyBase.ArrDetailFields.Count > 0 Then
                    clsCustomFieldGrid.GetData(obj.arrCustomFields, gv1, MyBase.ArrDetailFields, colICode)
                End If
                ''End of For Custom Fields
                '' CurrencConversion
                If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
                    obj.CURRENCY_CODE = Me.txtCurrencyCode.Value
                    ''''obj.ConvRate = clsCommon.myCDecimal(Me.txtConversionRate.Text)
                    If clsCommon.myLen(txtApplicableFrom.Text) > 0 Then
                        obj.ApplicableFrom = Me.txtApplicableFrom.Text
                    Else
                        obj.ApplicableFrom = Nothing
                    End If
                    obj.ConvRate = clsCommon.myCDecimal(Me.txtConversionRate.Text)
                Else
                    obj.CURRENCY_CODE = Nothing
                    obj.ConvRate = 1
                    obj.ApplicableFrom = Nothing
                End If
                obj.isTaxExempted = isTaxExempted
                '' end CurrencyConversion
                If clsMCCMaterialSale.checkSaveNotification(obj, Nothing) Then
                    If (obj.SaveData(obj, isNewEntry)) Then
                        UcAttachment1.SaveData(obj.Document_Code)
                        If ChekPostBtn = False And IsDataImported = False Then
                            common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                        End If
                        LoadData(obj.Document_Code, NavigatorType.Current)
                    End If
                End If
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            If ChekPostBtn = False Then
                common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
    End Function
    Function InvoiceType() As Boolean
        Dim dt As DataTable
        Dim strloc As String
        Dim qry As String
        strloc = txtBillToLocation.Value
        qry = "SELECT TSPL_LOCATION_MASTER.Excisable,TSPL_LOCATION_MASTER.State, " &
          "TSPL_LOCATION_MASTER.Sales_Tax_Group as LocalTaxGroup,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as Local_Tax_GroupName, " &
          "TSPL_LOCATION_MASTER.Sales_Tax_GroupIS as InterstateTaxGroup,TSPL_TAX_GROUP_MASTERIS.Tax_Group_Desc as Interstate_Tax_GroupName " &
          "FROM TSPL_LOCATION_MASTER left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' left outer join TSPL_TAX_GROUP_MASTER as TSPL_TAX_GROUP_MASTERIS on TSPL_TAX_GROUP_MASTERIS.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_GroupIS and TSPL_TAX_GROUP_MASTERIS.Tax_Group_Type='S' " &
          "WHERE TSPL_LOCATION_MASTER.Location_Code = '" + strloc + "'"
        dt = clsDBFuncationality.GetDataTable(qry)
        Dim strLocState As String = clsCommon.myCstr(dt.Rows(0)("State"))
        qry = "select Price_Code,price_CodeNon,State,Tin_No from TSPL_CUSTOMER_MASTER where Cust_Code='" + txtVendorNo.Value + "'"
        dt = clsDBFuncationality.GetDataTable(qry)
        Dim strCustState As String = clsCommon.myCstr(dt.Rows(0)("State"))
        Dim strTinNo As String = clsCommon.myCstr(dt.Rows(0)("Tin_No"))
        If AllowChangeInvoiceType = False Then
            If clsCommon.myLen(strTinNo) > 0 AndAlso clsCommon.CompairString(strLocState, strCustState) = CompairStringResult.Equal Then
                ddlInvoiceType.SelectedValue = "T"
            Else
                ddlInvoiceType.SelectedValue = "R"
            End If
        Else
            If clsCommon.myLen(ddlInvoiceType.SelectedValue) > 0 Then
                Dim strInvoiceType As String = String.Empty
                Dim strInvoiceTypeDesc As String = String.Empty
                If clsCommon.myLen(strTinNo) > 0 AndAlso clsCommon.CompairString(strLocState, strCustState) = CompairStringResult.Equal Then
                    strInvoiceType = "T"
                Else
                    strInvoiceType = "R"
                End If
                If Not clsCommon.CompairString(strInvoiceType, ddlInvoiceType.SelectedValue) = CompairStringResult.Equal Then
                    If strInvoiceType = "T" Then
                        strInvoiceTypeDesc = "Tax"
                    ElseIf strInvoiceType = "R" Then
                        strInvoiceTypeDesc = "Retail"
                    End If
                    If (common.clsCommon.MyMessageBoxShow("System is generating " & strInvoiceTypeDesc & "  Invoice Type.Do you still want to continue ?  ", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                        Return False
                    Else
                        ddlInvoiceType.SelectedValue = strInvoiceType
                        Return True
                    End If
                End If
            End If
        End If
        If GSTStatus Then
            If chkTaxable.Checked Then
                ddlInvoiceType.SelectedValue = "T"
            Else
                ddlInvoiceType.SelectedValue = "N"
            End If
        End If
        Return True
    End Function
    Private Sub fndProject__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndProject._MYValidating
        Dim qry As String = "select PROJECT_CODE as Code,SPECIFICATION,PROJECT_STATUS as Status from TSPL_PJC_PROJECT"
        fndProject.Value = clsCommon.ShowSelectForm("Project Code", qry, "Code", "", fndProject.Value, "", isButtonClicked)
        lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")
    End Sub
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try
            Dim obj As New clsMCCMaterialSale()
            obj = clsMCCMaterialSale.GetData(strCode, NavTyep)
            '  txtBillToLocation.Enabled = True
            txtSubLocation.Enabled = True
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
                btnSave.Enabled = True
                btnPost.Enabled = True
                btnDelete.Enabled = True
                isInsideLoadData = True
                isNewEntry = False
                btnSave.Text = "Update"
                BlankAllControls()
                LoadBlankGrid()
                LoadBlankGridTax()
                LoadBlankGridAC()
                cboItemType.Enabled = False
                txtBillToLocation.Enabled = False
                txtSubLocation.Enabled = False
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                    repoComplete.IsVisible = True
                    repoBalQty.IsVisible = True
                    chkIsEwayBill.Enabled = False
                    If obj.Is_Delivered = 1 Then
                        btnDeliveredTo.Enabled = False
                    Else
                        btnDeliveredTo.Enabled = True
                    End If
                    btnCreateEWB.Enabled = True
                Else
                    chkIsEwayBill.Enabled = True
                    btnCreateEWB.Enabled = False
                End If
                'txtWayBillno.Text = obj.WayBillNo
                'If clsCommon.myLen(obj.WayBillDate) > 0 Then
                '    txtWaybillDate.Value = obj.WayBillDate
                'End If
                chkIsEwayBill.Checked = IIf(obj.IsEwayBill = 1, True, False)
                chkisTCS.Checked = IIf(obj.IS_TCS = 1, True, False)
                txtEWayBillNo.Text = obj.EWayBillNo
                If obj.EWayBillDate IsNot Nothing Then
                    txtEWayBillDate.Value = obj.EWayBillDate
                    txtEWayBillDate.Checked = True
                End If
                chkVendorGrossReceipt.Checked = clsVendorMaster.isGrossReceipt(obj.Customer_Code)
                UsLock1.Status = obj.Status
                txtDocNo.Value = obj.Document_Code
                txtDate.Value = obj.Document_Date
                txtVendorNo.Value = obj.Customer_Code
                txtDate.Enabled = False
                txtVendorNo.Enabled = False
                chkTaxable.Enabled = False
                chkRateUserCustomer.ToggleState = ClsUserCustomerSettings.GetUserCustomerRateSetting(txtVendorNo.Value)
                txtRoadPermitNo.Text = obj.Road_Permit_No
                FndPriceCode.Value = obj.DCS_Price_Code
                txtSaleRoute.Value = obj.Sale_Route_Code
                lblVendorName.Text = obj.Customer_Name
                txtRefNo.Text = obj.Ref_No
                If clsCommon.myLen(obj.Challan_Date) > 0 Then
                    dtpChallan.Value = obj.Challan_Date
                    dtpChallan.Checked = True
                Else
                    dtpChallan.Checked = False
                End If
                If clsCommon.myLen(obj.Inv_Date) > 0 Then
                    dtpInvoice.Value = obj.Inv_Date
                    dtpInvoice.Checked = True
                Else
                    dtpInvoice.Checked = False
                End If
                txtTransporterCode.Value = obj.Transport_Id
                lblTransporterName.Text = obj.Transporter_Name
                chkOnHold.Checked = obj.On_Hold
                txtDesc.Text = obj.Description
                txtTaxGroup.Value = obj.Tax_Group
                chkcashsale.Checked = IIf(obj.Is_CashSale = "Y", True, False)
                chkApplyTPT.Checked = IIf(obj.Is_Apply_TPT = "1", True, False)
                chkNoTranspoter.Checked = IIf(obj.No_Transporter = 1, True, False)
                chkAddTPT.Checked = obj.Is_Add_TPT
                txtRecommBy.Text = obj.Recommended_By
                If chkApplyTPT.Checked Then
                    txtTPTVendor.Value = obj.TPT_Vendor
                End If
                chkcashsale.Enabled = False
                txtComment.Text = obj.Comments
                'txtShipToLocation.Value = obj.Ship_To_Location
                txtBillToLocation.Value = obj.Bill_To_Location
                txtSubLocation.Value = obj.Sub_Location_code
                If chkcashsale.Checked Then
                    txtReceiptNo.Value = obj.Receipt_No
                    lblReceiptAmt.Text = obj.ReceiptAmt
                    cmbPaymentType.SelectedValue = obj.Payment_Terms
                    txtBankCode.Value = obj.Bank_Code
                End If

                lblTotalSubsidy.Text = obj.TotalSubsidyAmt
                lblTotalDisSubsidy.Text = obj.TotalSubsidyDisAmt

                txtVehicleNo.Text = obj.VehicleNo
                txtReceiverName.Text = obj.ReceiverName
                If clsCommon.myLen(clsCommon.myCstr(txtBillToLocation.Value)) > 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                        txtSubLocation.Enabled = True
                    Else
                        txtSubLocation.Enabled = False
                    End If
                End If
                If (clsCommon.myLen(txtSubLocation.Value) > 0) Then
                    lblSubLocation.Text = clsLocation.GetName(txtSubLocation.Value, Nothing)
                End If
                txtInvNo.Text = obj.Inv_No
                Dim objTaxGrpMaster As New clsTaxGroupMaster()
                objTaxGrpMaster = objTaxGrpMaster.GetDataForSale(obj.Tax_Group)
                If (objTaxGrpMaster IsNot Nothing) Then
                    lblTaxGrpName.Text = objTaxGrpMaster.Tax_Group_Desc
                End If
                chkInternal.Checked = obj.Is_Internal
                cboItemType.SelectedValue = obj.Item_Type
                txtDept.Value = obj.Dept
                lblDept.Text = obj.Dept_Desc
                txtElectronicRefNo.Text = obj.Electronic_Ref_No
                txtTermCode.Value = obj.Terms_Code
                'lblTermName.Text = obj.Terms_Description
                If obj.Due_Date IsNot Nothing Then
                    txtDueDate.Value = obj.Due_Date
                End If
                txtNoOfInsallment.Value = obj.No_Of_Instalment
                chkExcludeKKFMandi.Checked = obj.Exclude_KKF_And_Mandi
                lblAmtWithDiscount.Text = clsCommon.myFormat(obj.Discount_Base)
                lblDiscountAmt.Text = clsCommon.myFormat(obj.Discount_Amt)
                lblAmtAfterDiscount.Text = clsCommon.myFormat(obj.Amount_Less_Discount)
                lblTaxAmt.Text = clsCommon.myFormat(obj.Total_Tax_Amt)
                lblTotRAmt.Text = clsCommon.myFormat(obj.Total_Amt)
                lblTotRAmt1.Text = lblTotRAmt.Text
                txtTPTAmt.Text = clsCommon.myFormat(obj.Transporter_Commission_TotalAmt)
                lblBillToLocation.Text = obj.BillToLocationName
                'lblShipToLocation.Text = obj.ShipToLocationName
                lblTaxGrpName.Text = obj.TaxGroupName
                lblTermName.Text = obj.TermsName
                txtRoundOff.Text = clsCommon.myFormat(obj.RoundOffAmount)
                txtCarrier.Text = obj.Carrier
                txtGRNo.Text = obj.GRNo
                txtGENo.Text = obj.GENo
                If obj.GEDate.HasValue Then
                    txtGEDate.Value = obj.GEDate
                    txtGEDate.Checked = True
                End If
                txtInvoiceNo.Text = obj.Invoice_No
                'richa Ticket No.BM00000002982
                txtMannaulInvoiceNo.Value = obj.Mannual_Invoice_No
                TxtInvoiceManualNoWithPrefix.Text = obj.InvoiceManualNowithPrefix
                '
                'txtReqNo.Value = obj.Against_Sales_Order
                fndProject.Value = obj.PROJECT_ID
                lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")
                'If clsCommon.myLen(txtReqNo.Value) > 0 Then
                '    lblProject.Enabled = False
                '    fndProject.Enabled = False
                'End If
                txtPriceCode.Text = obj.Price_Code
                txtDiscPer.Text = obj.HeadDisc_Per
                txtDiscAmt.Text = obj.HeadDisc_Amt

                txtRatePer.Text = obj.RateDiff_Per
                txtRateAmt.Text = obj.RateDiff_Amt

                If obj.RateDiff_Per > 0 Then
                    GroupBoxRateDiffAmtType.Visible = True
                    If obj.Rate_Diff_Amount_Type = 1 Then
                        rbtnTotalAmt.IsChecked = True
                    ElseIf obj.Rate_Diff_Amount_Type = 2 Then
                        rbtnBasicAmt.IsChecked = True
                    End If
                Else
                    GroupBoxRateDiffAmtType.Visible = False
                End If
                If clsCommon.myLen(txtDiscAmt.Text) <= 0 OrElse clsCommon.myLen(txtDiscPer.Text) <= 0 OrElse clsCommon.myCDecimal(txtDiscAmt.Text) = 0 OrElse clsCommon.myCDecimal(txtDiscPer.Text) = 0 Then
                    txtDiscPer.Text = obj.HeadDisc_Per
                    If clsCommon.myCDecimal(txtDiscPer.Text) = 0 Then
                        txtDiscAmt.Text = obj.HeadDisc_Amt
                        chkDiscountOnAmt.IsChecked = True
                        If MultiplySubsidyWithQuantity Then
                            lblInvoiceDiscAmt.Text = obj.TotalSubsidyDisAmt
                        Else
                            lblInvoiceDiscAmt.Text = obj.HeadDisc_Amt
                        End If
                    Else
                        chkDiscountOnRate.IsChecked = True
                        If MultiplySubsidyWithQuantity Then
                            lblInvoiceDiscAmt.Text = obj.TotalSubsidyDisAmt
                        Else
                            lblInvoiceDiscAmt.Text = obj.HeadDisc_PerAmt
                        End If
                    End If
                End If
                'If clsCommon.myLen(txtRateAmt.Text) <= 0 OrElse clsCommon.myLen(txtRatePer.Text) <= 0 OrElse clsCommon.myCDecimal(txtRateAmt.Text) = 0 OrElse clsCommon.myCDecimal(txtRatePer.Text) = 0 Then
                If True Then
                    txtRatePer.Text = obj.RateDiff_Per
                    If clsCommon.myCDecimal(txtRatePer.Text) = 0 Then
                        txtRateAmt.Text = obj.RateDiff_Amt
                        'chkRateDiffAmt.IsChecked = True
                    Else
                        ' chkRateDiffRate.IsChecked = True
                    End If
                End If
                If obj.Is_Rate_Diff_Per_Amt = 1 Then
                    chkRateDiffRate.IsChecked = True
                ElseIf obj.Is_Rate_Diff_Per_Amt = 2 Then
                    chkRateDiffAmt.IsChecked = True
                End If
                lblGrossAmount.Text = clsCommon.myCDecimal(obj.Gross_Amount)
                ddlInvoiceType.SelectedValue = obj.Invoice_Type
                lblCommAmt.Text = obj.Total_Comm_Amt
                Dim dt_vlc As DataTable = clsDBFuncationality.GetDataTable("SELECT TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code,TSPL_VLC_MASTER_HEAD.* FROM TSPL_CUSTOMER_VENDOR_MAPPING left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code = TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code WHERE TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "'")
                If dt_vlc.Rows.Count > 0 Then
                    LblVlc_Code.Text = dt_vlc.Rows(0).Item("vlc_code_vlc_Uploader")
                    LblVlc_Code.Tag = dt_vlc.Rows(0).Item("vlc_code")
                    LblVlc_Name.Text = dt_vlc.Rows(0).Item("vlc_name")
                Else
                    LblVlc_Code.Text = ""
                    LblVlc_Code.Tag = Nothing
                    LblVlc_Name.Text = ""
                End If
                setDCSBalance()
                If (clsCommon.myLen(obj.TAX1) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX1
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX1_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX1_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX1_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX1) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX2) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX2
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX2_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX2_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX2_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX2) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX3) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX3
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX3_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX3_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX3_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX3) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX4) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX4
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX4_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX4_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX4_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX4) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX5) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX5
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX5_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX5_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX5_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX5) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX6) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX6
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX6_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX6_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX6_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX6) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX7) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX7
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX7_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX7_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX7_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX7) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX8) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX8
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX8_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX8_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX8_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX8) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX9) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX9
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX9_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX9_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX9_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX9) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX10) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX10
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX10_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX10_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX10_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX10) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code1) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code1
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name1
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt1
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code2) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code2
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name2
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt2
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code3) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code3
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name3
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt3
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code4) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code4
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name4
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt4
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code5) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code5
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name5
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt5
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code6) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code6
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name6
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt6
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code7) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code7
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name7
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt7
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code8) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code8
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name8
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt8
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code9) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code9
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name9
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt9
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code10) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code10
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name10
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt10
                End If
                lblAddCharges.Text = clsCommon.myFormat(obj.Total_Add_Charge)
                lblAddCharges1.Text = clsCommon.myFormat(obj.Total_Add_Charge)
                If obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic Then
                    rbtnTaxCalAutomatic.IsChecked = True
                ElseIf obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual Then
                    rbtnTaxCalManual.IsChecked = True
                End If
                chkCreateAutoInvoice.Checked = obj.Is_Create_Auto_Invoice
                'chkCreateAutoReceipt.Visible = chkCreateAutoInvoice.Checked
                chkCreateAutoReceipt.Checked = obj.Is_Create_Auto_Receipt
                chkTaxable.Checked = obj.Is_Taxable
                cboDeductionType.SelectedValue = obj.Deduction
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsMCCMaterialSaleDetail In obj.Arr
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Tag = objTr.arrSrItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = objTr.Line_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = objTr.Row_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colComplete).Value = IIf(objTr.Status = 0, "No", "Yes")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = objTr.arrBatchItem ' change by prabhakar
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colShortDesc).Value = clsItemMaster.GetItemShortDescription(objTr.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objTr.Item_Desc
                        If objTr.arrBatchItem IsNot Nothing Then
                            Dim batchNoList As String = String.Join(","c, objTr.arrBatchItem.Where(Function(x) Not String.IsNullOrEmpty(x.Batch_No)).Select(Function(x) x.Batch_No.Replace("'", "''")).ToArray())
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colManualBatchNo).Value = batchNoList
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDCSSaleZeroCost).Value = clsCommon.myCBool(objTr.DCS_Sale_Zero_Cost = 1)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(objTr.Item_Code, Nothing)
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colBarCode).Value = objTr.Bar_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(objTr.Item_Code)
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgSOQty).Value = objTr.so_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = objTr.Balance_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = objTr.Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBillingQty).Value = objTr.Billing_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = objTr.Balance_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFreeQty).Value = objTr.Free_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgUnit).Value = objTr.OrgUnit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objTr.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBillingUnit).Value = objTr.Billing_Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUOMName).Value = clsDBFuncationality.getSingleValue("select Unit_Desc from tspl_unit_master where Unit_code= '" + objTr.Unit_code + "'")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = objTr.Order_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objTr.Item_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = objTr.Location
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = objTr.LocationName
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPer).Value = objTr.Disc_Per
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisAmt).Value = objTr.Disc_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPerUnit).Value = objTr.Disc_Per_Unit
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisUnitAmt).Value = objTr.Disc_Unit_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterDis).Value = objTr.Amt_Less_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax1).Value = objTr.TAX1
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt1).Value = objTr.TAX1_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = objTr.TAX1_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt1).Value = objTr.TAX1_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax2).Value = objTr.TAX2
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt2).Value = objTr.TAX2_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = objTr.TAX2_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt2).Value = objTr.TAX2_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax3).Value = objTr.TAX3
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt3).Value = objTr.TAX3_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = objTr.TAX3_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt3).Value = objTr.TAX3_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax4).Value = objTr.TAX4
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt4).Value = objTr.TAX4_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = objTr.TAX4_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt4).Value = objTr.TAX4_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax5).Value = objTr.TAX5
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt5).Value = objTr.TAX5_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = objTr.TAX5_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt5).Value = objTr.TAX5_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax6).Value = objTr.TAX6
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt6).Value = objTr.TAX6_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = objTr.TAX6_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt6).Value = objTr.TAX6_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax7).Value = objTr.TAX7
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt7).Value = objTr.TAX7_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = objTr.TAX7_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt7).Value = objTr.TAX7_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax8).Value = objTr.TAX8
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt8).Value = objTr.TAX8_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = objTr.TAX8_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt8).Value = objTr.TAX8_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax9).Value = objTr.TAX9
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt9).Value = objTr.TAX9_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = objTr.TAX9_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt9).Value = objTr.TAX9_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax10).Value = objTr.TAX10
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt10).Value = objTr.TAX10_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = objTr.TAX10_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt10).Value = objTr.TAX10_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = objTr.Total_Tax_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterTax).Value = objTr.Item_Net_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = objTr.MRP
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableRate).Value = objTr.Assessable
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableAmount).Value = objTr.AssessableAmt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = objTr.Batch_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = objTr.Bin_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHeaDDisPer).Value = objTr.HeadDiscPer
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDisPerAmt).Value = objTr.HeadDiscPerAmt
                        If objTr.Expiry_Date.HasValue Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colExpiry).Value = objTr.Expiry_Date
                        End If
                        If objTr.MFG_Date.HasValue Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colManufactureDate).Value = objTr.MFG_Date
                        End If
                        'If clsCommon.myLen(objTr.Order_Code) > 0 Then
                        '    gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = clsSNSalesOrderDetail.GetBalancePOQty(objTr.Order_Code, objTr.Item_Code, obj.Document_Code, objTr.Unit_code, objTr.MRP, objTr.Assessable)
                        'End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSpecification).Value = objTr.Specification
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRemarks).Value = objTr.Remarks
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsMannualAmt).Value = objTr.Is_Mannual_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = IIf(objTr.Scheme_Applicable = "Y", "Yes", "No")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFromSchemeCode).Value = objTr.Scheme_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = IIf(objTr.Scheme_Item = "Y", "Yes", "No")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = objTr.Item_Tax
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalMRP).Value = objTr.Total_MRP_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalBasicAmount).Value = objTr.Total_Basic_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalDiscountAmount).Value = objTr.Total_Disc_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colcustDiscount).Value = objTr.Cust_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalCustDiscount).Value = objTr.Total_Cust_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colActualCost).Value = objTr.ActualRate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCustDiscountQty).Value = objTr.Cust_DiscountQty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceDateColumn).Value = objTr.Price_Date
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceCOde).Value = objTr.Price_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementPer).Value = objTr.Abatement_Per
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementAmount).Value = objTr.Abatement_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColFOC).Value = objTr.FOC_Item
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value), clsCommon.myCDecimal(gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = objTr.Batch_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemWeight).Value = objTr.Item_Weight
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colConvF).Value = objTr.Conv_Factor
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotItemWt).Value = objTr.TotalItem_Weight
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkupOn).Value = objTr.Markup_On
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkUpPercentage).Value = objTr.Markup_Percent
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLandingCost).Value = objTr.Landing_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCustDiscPercentage).Value = objTr.CustDiscPer
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = objTr.HeadDiscAmt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCashDiscSchemeCode).Value = objTr.CasdDiscScheme_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPurCost).Value = objTr.Purchase_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgCost).Value = objTr.OrgRate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleCode).Value = objTr.PrincipleCode
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleDesc).Value = objTr.PrincipleDesc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorCode).Value = objTr.vendor_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorDesc).Value = objTr.vendor_desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCommRate).Value = objTr.Commission_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommParty).Value = objTr.Commission_Party
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommPartyName).Value = clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommParty).Value & "'")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommAmt).Value = objTr.Commission_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColAmtAfterCOmm).Value = objTr.Amt_Less_Commission
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemwiseTaxCode).Value = objTr.ItemwiseTaxCode
                        If clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(ColTPTRate)) > 0 Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColTPTRate).Value = objTr.Transporter_Commission_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColTPTRate).Tag = objTr.REF_TPT_PK_ID
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColTPTAmt).Value = objTr.Transporter_Commission_Amt
                        End If
                        If obj.Status = ERPTransactionStatus.Pending Then
                            If clsCommon.myLen(obj.TAX1) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable1).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX1)
                            End If
                            If clsCommon.myLen(obj.TAX2) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable2).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX2)
                            End If
                            If clsCommon.myLen(obj.TAX3) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable3).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX3)
                            End If
                            If clsCommon.myLen(obj.TAX4) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable4).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX4)
                            End If
                            If clsCommon.myLen(obj.TAX5) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable5).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX5)
                            End If
                            If clsCommon.myLen(obj.TAX6) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable6).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX6)
                            End If
                            If clsCommon.myLen(obj.TAX7) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable7).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX7)
                            End If
                            If clsCommon.myLen(obj.TAX8) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable8).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX8)
                            End If
                            If clsCommon.myLen(obj.TAX9) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable9).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX9)
                            End If
                            If clsCommon.myLen(obj.TAX10) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable10).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX10)
                            End If
                        End If
                    Next
                    If obj.Status = ERPTransactionStatus.Pending Then
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                        gvAC.Rows.AddNew()
                    End If
                    ''For Custom Fields
                    If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                        UcCustomFields1.LoadData(obj.Document_Code)
                    End If
                    clsCustomFieldGrid.FillDataInGrid(obj.Document_Code, MyBase.Form_ID, gv1)
                    ''End of For Custom Fields
                    '' MULTICURRENCY
                    Me.txtCurrencyCode.Value = obj.CURRENCY_CODE
                    Me.txtConversionRate.Text = obj.ConvRate
                    Me.txtApplicableFrom.Text = obj.ApplicableFrom.ToString
                    '' end  MULTICURRENCY
                    'LoadParentSHipCode()
                End If
                SetitemWiseTaxOnlySetting()
                ' ''RefreshReqNo()
                ' ''RefreshGRPNo()
                UcAttachment1.LoadData(obj.Document_Code)
                Dim qry As String = "select IRN_no,Ack_No,Ack_Date,QR_Code,EWayBillNo,EwayBillDate,EwayBillValidDate,EWayBillRemarks from TSPL_SD_SALE_INVOICE_HEAD where Document_Code = '" + obj.Invoice_No + "'"
                Dim dtInv As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dtInv IsNot Nothing AndAlso dtInv.Rows.Count > 0 Then
                    TxtEInvoiceUpdateIRNNo.Text = clsCommon.myCstr(dtInv.Rows(0)("IRN_no"))
                    TxtEInvoiceUpdateAckNo.Text = clsCommon.myCstr(dtInv.Rows(0)("Ack_No"))
                    If dtInv.Rows(0)("Ack_Date") IsNot DBNull.Value Then
                        TxtEInvoiceUpdateAckDate.Value = clsCommon.myCDate(dtInv.Rows(0)("Ack_Date"))
                    End If
                    TxtEInvoiceUpdateQCCode.Text = clsCommon.myCstr(dtInv.Rows(0)("QR_Code"))
                    TxtEWayBillUpdateBillNo.Text = clsCommon.myCstr(dtInv.Rows(0)("EWayBillNo"))
                    If dtInv.Rows(0)("EwayBillDate") IsNot DBNull.Value Then
                        TxtEWayBillUpdateBillDate.Value = clsCommon.myCDate(dtInv.Rows(0)("EwayBillDate"))
                    End If
                    If dtInv.Rows(0)("EwayBillValidDate") IsNot DBNull.Value Then
                        TxtEWayBillUpdateValidDate.Value = clsCommon.myCDate(dtInv.Rows(0)("EwayBillValidDate"))
                    End If
                    TxtEWayBillUpdateBillRemarks.Text = clsCommon.myCstr(dtInv.Rows(0)("EWayBillRemarks"))
                End If
            End If
            If clsCommon.myLen(txtInvoiceNo.Text) > 0 Then
                btnInvoiceJE.Visible = True
            Else
                btnInvoiceJE.Visible = False
            End If
            txtPaymentCycleNo.Text = clsGenratePaymentCycles.GetPaymentCycleNo(txtBillToLocation.Value, txtDate.Value)
            txtFiscalYear.Text = clsGenratePaymentCycles.GetPaymentFiscalCode(txtBillToLocation.Value, txtDate.Value)

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
        End Try
    End Sub
    Public Shared Function IsBatchDetailMandatory(ByVal strUOMCode As String) As Boolean
        Dim qry As String = "select 1 from TSPL_UNIT_MASTER where Unit_Code='" + strUOMCode + "' and Empty='Y'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return False
        End If
        Return True
    End Function
    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        ''Dim arrVN As New List(Of String)()
        ''Dim strCode As String = ""
        ''Try
        ''    Dim qry As String = "select Voucher_No from TSPL_JOURNAL_MASTER order by Voucher_No"
        ''    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        ''    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        ''        For Each dr As DataRow In dt.Rows
        ''            Try
        ''                strCode = clsCommon.myCstr(dr("Voucher_No"))
        ''                qry = "select convert(date,TSPL_JOURNAL_MASTER.Voucher_Date ,103) from TSPL_JOURNAL_MASTER  where Voucher_No='" + strCode + "'"
        ''                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
        ''            Catch ex As Exception
        ''                arrVN.Add(strCode)
        ''            End Try
        ''        Next
        ''    End If
        ''Catch ex As Exception
        ''End Try
        'Dim dt As DataTable = clsDBFuncationality.GetDataTable("select top 10 * from TSPL_JOURNAL_MASTER")
        'For ii As Integer = 0 To dt.Columns.Count - 1
        '    Dim oldColumnName As String = dt.Columns(ii).ColumnName
        '    Dim NewColumnName As String = ""
        '    If oldColumnName.Contains("_") Then
        '        NewColumnName = oldColumnName.Replace("_", Environment.NewLine)
        '    End If
        '    If clsCommon.myLen(NewColumnName) > 0 Then
        '        dt.Columns(ii).ColumnName = NewColumnName
        '    End If
        'Next
        CloseForm()
    End Sub
    Sub CloseForm()
        Me.Close()
    End Sub
    Private Sub gv2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv2.DoubleClick
        Try
            If CalculateTaxRatefromItemwsieTaxOnSale = 0 Then
                If rbtnTaxCalAutomatic.IsChecked Then
                    Dim qry As String = "select Tax_Rate_Code as [Rate Code],Tax_Rate_Desc as [Rate Description],Tax_Rate as [Rate] from TSPL_TAX_RATES "
                    Dim dblNewRate As Decimal = clsLocationWiseTax.FinderForTaxRate(txtBillToLocation.Value, txtTaxGroup.Value, clsCommon.myCstr(gv2.CurrentRow.Cells(colTTaxAutCode).Value), txtVendorNo.Value, "S")
                    Dim intRowNo As Integer = gv2.CurrentRow.Index
                    If gv1.RowCount > 0 AndAlso intRowNo >= 0 Then
                        Dim strII As String = clsCommon.myCstr(intRowNo + 1)
                        For ii As Integer = 0 To gv1.Rows.Count - 1
                            gv1.Rows(ii).Cells("COLTAXRATE" + strII).Value = dblNewRate
                        Next
                    End If
                    For ii As Integer = 0 To gv1.Rows.Count - 1
                        UpdateCurrentRow(ii)
                    Next
                    UpdateAllTotals()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        PostData()
    End Sub
    Private Function CreateSTN()
        Dim strTransferOut As String
        Dim strGITLoc As String
        If chkAutoTransfer.Checked = True Then
            If clsCommon.myLen(txtFromLoc.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select From Location", Me.Text)
                txtFromLoc.Focus()
                Return False
            End If
            If common.clsCommon.MyMessageBoxShow("Are you sure to Create Auto Transfer from location " & txtFromLoc.Value & " To " & txtBillToLocation.Value & "  ?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                Try
                    strGITLoc = clsDBFuncationality.getSingleValue("select GIT_Location from TSPL_LOCATION_MASTER where Location_Code='" & txtBillToLocation.Value & "'", trans)
                    If clsCommon.myLen(strGITLoc) <= 0 Then
                        clsCommon.MyMessageBoxShow("Pls create GIT location for " & txtBillToLocation.Value & " ")
                        trans.Rollback()
                        Return False
                    End If
                    Dim obj As New clsTransferDCC
                    obj.Form_ID = MyBase.Form_ID
                    obj.Document_Date = txtDate.Value
                    obj.Delivery_date = txtDate.Value
                    obj.Delivery_Duration = 0
                    obj.AutoTransfer = 1
                    obj.Ref_No = txtRefNo.Text
                    obj.Total_Tax_Amt = clsCommon.myCDecimal(lblTaxAmt.Text)
                    obj.Remarks = "Auto Transfer Out"
                    obj.From_Location = txtFromLoc.Value
                    obj.To_Location = strGITLoc
                    obj.Comments = txtComment.Text
                    obj.On_Hold = chkOnHold.Checked
                    obj.Mode_Of_Transport = "By Road"
                    obj.Tax_Group = ""
                    obj.Transfer_Type = "O"
                    obj.Terms_Code = txtTermCode.Value
                    obj.Terms_Remark = txtComment.Text
                    obj.Due_Date = txtDueDate.Value
                    obj.Discount_Base = clsCommon.myCDecimal(lblAmtWithDiscount.Text)
                    obj.Total_Amt_Less_Tax = clsCommon.myCDecimal(lblAmtWithDiscount.Text)
                    obj.Discount_Amt = 0
                    obj.Amount_Less_Discount = clsCommon.myCDecimal(lblAmtWithDiscount.Text)
                    obj.DOC_Total_Amt = clsCommon.myCDecimal(lblAmtWithDiscount.Text)
                    obj.TransferOutNo = ""
                    obj.Km_Reading = 1
                    obj.Is_AgainstFormF = 0
                    obj.Type = ""
                    If rbtnTaxCalAutomatic.IsChecked Then
                        obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                    ElseIf rbtnTaxCalManual.IsChecked Then
                        obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                    End If
                    obj.Arr = New List(Of clsTransferDCCDetail)
                    Dim arr As New List(Of String)
                    Dim strCode As String
                    For ii As Integer = 0 To gv1.Rows.Count - 1
                        'If (clsCommon.myCDecimal(gv1.Rows(ii).Cells(ColActualBalQty).Value) - clsCommon.myCDecimal(gv1.Rows(ii).Cells(colQty).Value)) < 0 Then
                        Dim objTr As New clsTransferDCCDetail()
                        Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                        Dim strMRPOuter As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colMRP).Value)
                        Dim strUnitOuter As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                        Dim dblOutQty As Decimal = 0
                        objTr.Line_No = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colLineNo).Value)
                        'objTr.Row_Type = clsCommon.myCstr(grow.Cells(colRowType).Value)
                        objTr.Item_Code = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                        objTr.Item_Desc = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                        objTr.Out_Qty = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colQty).Value)
                        objTr.Unit_code = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                        objTr.MRP = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colMRP).Value)
                        ''Check for Balance With Unapproved Qty
                        Dim arrlist As New ArrayList()
                        Dim strCurrMRP As Decimal = clsCommon.myCDecimal(objTr.MRP)
                        Dim dblOuterConvFac As Decimal = 0
                        Dim dblActualbalonashipment = 0
                        Dim dblBalQty As Decimal = clsItemLocationDetails.getBalance(objTr.Item_Code, txtFromLoc.Value, txtDocNo.Value, txtDate.Value, trans, objTr.Unit_code, objTr.MRP)
                        Dim dblEnteredQty As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colQty).Value)
                        Dim dblActualBalOuter As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(ColActualBalQty).Value)
                        For jj As Integer = 0 To gv1.Rows.Count - 1
                            If ii = jj Then
                                Continue For
                            End If
                            Dim strICodeInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                            Dim strUOMInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                            Dim dblMRPInner As Decimal = clsCommon.myCDecimal(gv1.Rows(jj).Cells(colMRP).Value)
                            Dim dblQtyInner As Decimal = clsCommon.myCDecimal(gv1.Rows(jj).Cells(colQty).Value)
                            Dim dblConvFInner As Decimal = clsCommon.myCDecimal(gv1.Rows(jj).Cells(colConvF).Value)
                            Dim dblActualBalInner As Decimal = clsCommon.myCDecimal(gv1.Rows(jj).Cells(ColActualBalQty).Value)
                            If dblQtyInner > 0 AndAlso dblMRPInner = (strMRPOuter) AndAlso clsCommon.CompairString(strICodeInner, strICode) = CompairStringResult.Equal Then
                                '''''' started
                                strCode = strICodeInner
                                If Not arr.Contains(strCode) Then
                                    arr.Add(strCode)
                                    dblOutQty = dblActualBalOuter - objTr.Out_Qty
                                    If dblOutQty < 0 Then
                                        dblOutQty = dblOutQty * -1
                                    Else
                                        dblOutQty = 0
                                    End If
                                Else
                                    dblOutQty = dblActualBalInner - dblQtyInner
                                    If dblOutQty < 0 Then
                                        dblOutQty = dblOutQty * -1
                                    Else
                                        dblOutQty = dblEnteredQty - dblOutQty
                                    End If
                                End If
                                '''''' ended
                                If clsCommon.CompairString(strUnitOuter, strUOMInner) = CompairStringResult.Equal Then
                                    dblEnteredQty += dblQtyInner
                                Else
                                    dblEnteredQty += (dblQtyInner * dblConvFInner)
                                End If
                            End If
                        Next
                        dblEnteredQty = Math.Round(dblEnteredQty, 2, MidpointRounding.ToEven)
                        If dblEnteredQty > dblBalQty Then
                            Throw New Exception("Item - " + strICode + " , MRP - " + clsCommon.myCstr(strMRPOuter) + Environment.NewLine + "Entered Quantity - " + clsCommon.myCstr(dblEnteredQty) + " and Balance Quantity - " + clsCommon.myCstr(dblBalQty) + " at Location " + txtFromLoc.Value)
                        End If
                        Dim dblBalQtyInLoc As Decimal = clsItemLocationDetails.getBalance(objTr.Item_Code, txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, trans, objTr.Unit_code, objTr.MRP)
                        If dblEnteredQty < dblBalQtyInLoc Then
                            common.clsCommon.MyMessageBoxShow("You have already balance for Item - " + txtBillToLocation.Value + Environment.NewLine + " ", Me.Text)
                            objTr.Item_Code = ""
                        End If
                        objTr.In_Qty = 0
                        objTr.Breakage = 0
                        objTr.Leak = 0
                        objTr.Shortage = 0
                        objTr.Out_Qty = dblOutQty
                        objTr.TransferOutNo = ""
                        'objTr.Location = clsCommon.myCstr(grow.Cells(colloc).Value)
                        objTr.Item_Cost = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colRate).Value)
                        objTr.Amount = Math.Round(dblOutQty * clsCommon.myCDecimal(gv1.Rows(ii).Cells(colRate).Value), 2)
                        objTr.Disc_Per = 0
                        objTr.Disc_Amt = 0
                        objTr.Amt_Less_Discount = objTr.Amount
                        objTr.Item_Net_Amt = objTr.Amount
                        objTr.Specification = ""
                        objTr.Remarks = ""
                        objTr.Location = txtFromLoc.Value 'clsCommon.myCstr(grow.Cells(colLocationCode).Value)
                        ''objTr.Assessable = clsCommon.myCDecimal(grow.Cells(colAssessableRate).Value)
                        If (clsCommon.myLen(objTr.Item_Code) > 0) And dblOutQty > 0 Then
                            obj.Arr.Add(objTr)
                        End If
                        'End If
                    Next
                    If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                        'common.clsCommon.MyMessageBoxShow("Please Fill at list one Item")
                        trans.Rollback()
                        Return True
                        Exit Function
                    End If
                    obj.SaveData(obj, True, False, trans)
                    obj.Document_No = obj.Document_No
                    clsTransferDCC.postTransfer(obj.Document_No, trans)
                    strTransferOut = obj.Document_No
                    '' for load in 
                    Dim objIn As New clsTransferDCC
                    Dim objTransfer As New clsTransferDCC
                    objTransfer = clsTransferDCC.GetData(strTransferOut, NavigatorType.Current, trans)
                    If (objTransfer IsNot Nothing AndAlso clsCommon.myLen(objTransfer.Document_No) > 0) Then
                        objIn.Document_Date = objTransfer.Document_Date
                        objIn.Delivery_date = objTransfer.Delivery_date
                        objIn.Delivery_Duration = 0
                        objIn.Ref_No = objTransfer.Ref_No
                        objIn.Total_Tax_Amt = objTransfer.Total_Tax_Amt
                        objIn.Remarks = "Auto Transfer In"
                        objIn.From_Location = strGITLoc
                        objIn.To_Location = txtBillToLocation.Value
                        objIn.Comments = objTransfer.Comments
                        objIn.On_Hold = chkOnHold.Checked
                        objIn.Mode_Of_Transport = "By Road"
                        objIn.AutoTransfer = 1
                        objIn.Tax_Group = ""
                        objIn.Transfer_Type = "I"
                        objIn.Terms_Code = objTransfer.Terms_Code
                        objIn.Terms_Remark = txtComment.Text
                        objIn.Due_Date = txtDueDate.Value
                        objIn.Discount_Base = objTransfer.Discount_Base
                        objIn.Total_Amt_Less_Tax = objTransfer.Total_Amt_Less_Tax
                        objIn.Discount_Amt = 0
                        objIn.Amount_Less_Discount = objTransfer.Amount_Less_Discount
                        objIn.DOC_Total_Amt = objTransfer.DOC_Total_Amt
                        objIn.TransferOutNo = strTransferOut
                        objIn.Vehicle_Code = objTransfer.Vehicle_Code
                        objIn.Vehicle_No = objTransfer.Vehicle_No
                        objIn.Km_Reading = 1
                        objIn.Is_AgainstFormF = 0
                        objIn.Type = ""
                        If rbtnTaxCalAutomatic.IsChecked Then
                            objIn.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                        ElseIf rbtnTaxCalManual.IsChecked Then
                            objIn.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                        End If
                        objIn.Arr = New List(Of clsTransferDCCDetail)
                        For Each objInDetail As clsTransferDCCDetail In objTransfer.Arr
                            'If (clsCommon.myCDecimal(grow.Cells(ColActualBalQty).Value) - clsCommon.myCDecimal(grow.Cells(colQty).Value)) < 0 Then
                            Dim objTrIn As New clsTransferDCCDetail()
                            objTrIn.Line_No = objInDetail.Line_No
                            'objTr.Row_Type = clsCommon.myCstr(grow.Cells(colRowType).Value)
                            objTrIn.Item_Code = objInDetail.Item_Code
                            objTrIn.Item_Desc = objInDetail.Item_Desc
                            objTrIn.In_Qty = objInDetail.Out_Qty
                            objTrIn.Unit_code = objInDetail.Unit_code
                            objTrIn.MRP = objInDetail.MRP
                            objTrIn.Breakage = 0
                            objTrIn.Leak = 0
                            objTrIn.Shortage = 0
                            objTrIn.TransferOutNo = strTransferOut
                            'objTr.Location = clsCommon.myCstr(grow.Cells(colloc).Value)
                            objTrIn.Item_Cost = objInDetail.Item_Cost
                            objTrIn.Amount = objInDetail.Amount
                            objTrIn.Disc_Per = objInDetail.Disc_Per
                            objTrIn.Disc_Amt = objInDetail.Disc_Amt
                            objTrIn.Amt_Less_Discount = objInDetail.Amt_Less_Discount
                            objTrIn.Item_Net_Amt = objInDetail.Item_Net_Amt
                            objTrIn.Specification = ""
                            objTrIn.Remarks = ""
                            objTrIn.Location = txtFromLoc.Value 'clsCommon.myCstr(grow.Cells(colLocationCode).Value)
                            'Dim dblBalQtyIn As Decimal = clsItemLocationDetails.getBalance(objTrIn.Item_Code, txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, trans, objTrIn.Unit_code, objTrIn.MRP)
                            'Dim dblInQty As Decimal = objTrIn.In_Qty
                            'If dblInQty < dblBalQtyIn Then
                            '    common.clsCommon.MyMessageBoxShow("You have already balance for Item - " + txtBillToLocation.Value + Environment.NewLine + " ", Me.Text)
                            '    objTrIn.Item_Code = ""
                            'End If
                            'objTrIn.In_Qty = dblInQty - dblBalQtyIn
                            ''objTr.Assessable = clsCommon.myCDecimal(grow.Cells(colAssessableRate).Value)
                            If (clsCommon.myLen(objTrIn.Item_Code) > 0) Then
                                objIn.Arr.Add(objTrIn)
                            End If
                            'End If
                        Next
                        If (objIn.Arr Is Nothing OrElse objIn.Arr.Count <= 0) Then
                            common.clsCommon.MyMessageBoxShow(Me, "Please Fill at list one Item", Me.Text)
                            trans.Rollback()
                        End If
                    End If
                    objIn.SaveData(objIn, True, False, trans)
                    objIn.Document_No = objIn.Document_No
                    clsTransferDCC.postTransfer(objIn.Document_No, trans)
                    trans.Commit()
                    Return True
                Catch ex As Exception
                    trans.Rollback()
                    common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                End Try
            End If
        End If
        Return True
    End Function
    Sub PostData()
        Try
            Dim msg As String = ""
            Dim qry As String = ""
            Dim dt As DataTable = Nothing
            If (myMessages.postConfirm()) Then
                'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                '    If chkcashsale.Checked Then
                '        If clsCommon.myLen(txtReceiptNo.Value) <= 0 Then
                '            Throw New Exception("Please select Receipt No.")
                '            txtReceiptNo.Focus()
                '        End If
                '    End If
                'End If
                If Not clsMCCMaterialSale.checkApprovalDocument(Form_ID, txtDocNo.Value) Then
                    btnPost.Enabled = False
                    Throw New Exception("Approval required for this Document")
                End If
                If SaveData(True) Then
                    If (clsMCCMaterialSale.PostData(MyBase.Form_ID, txtDocNo.Value)) Then
                        msg = "Successfully Posted"
                    Else
                        qry = "select No_Of_Level, LEVEL from TSPL_APPROVAL_LEVEL_SCREEN where User_Code='" + objCommonVar.CurrentUserCode + "' and Trans_Code='" + MyBase.Form_ID + "' "
                        dt = clsDBFuncationality.GetDataTable(qry)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            Dim level As String = dt.Rows(0)("LEVEL").ToString()
                            Dim NoOflevel As Integer = clsCommon.myCDecimal(dt.Rows(0)("No_Of_Level"))
                            If clsCommon.CompairString(level, "Level1") = CompairStringResult.Equal Then
                                msg = "Level 1 Approval done. "
                                If NoOflevel = 1 Then
                                    msg += "Successfully Posted. "
                                Else
                                    msg += "Level 2 Approval Required."
                                End If
                            ElseIf clsCommon.CompairString(level, "Level2") = CompairStringResult.Equal Then
                                msg = "Level 2 Approval done. "
                                If NoOflevel = 2 Then
                                    msg += "Successfully Posted "
                                Else
                                    msg += "Level 3 Approval Required."
                                End If
                            Else
                                msg = "Level 3 Approval done. Successfully Posted"
                            End If
                        End If
                    End If
                    common.clsCommon.MyMessageBoxShow(msg)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                    '===============if setting on then sms send
                    'If clsSMSAtPost_Sales.SMSATPOST_SALE() Then
                    '    SMSSENDONLY(True)
                    'End If
                    '=============================================
                    If (common.clsCommon.MyMessageBoxShow("Do you want to print", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes) Then
                        clsMCCMaterialSale.funPrint(MyBase.Form_ID, False, txtDate.Value, txtDocNo.Value, False)
                        'funPrint()
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        DeleteData()
    End Sub
    Sub DeleteData()
        Try
            Dim Reason As String = ""
            If (myMessages.deleteConfirm()) Then
                If clsCancelLog.CheckForReasonOnDelete() Then
                    '' REASON FOR DELETE 
                    Dim frm As New FrmFreeTxtBox1
                    frm.Text = "Remarks for Delete"
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 Then
                        Exit Sub
                    Else
                        Reason = frm.strRmks
                    End If
                End If
                If (clsMCCMaterialSale.DeleteData(txtDocNo.Value)) Then
                    saveCancelLog(Reason, "Delete", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Function saveCancelLog(ByVal Reason As String, ByVal Activity_Type As String, Optional ByVal trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim obj As New clsCancelLog
        obj.Program_Code = Form_ID
        obj.DOCUMENT_NO = clsCommon.myCstr(Me.txtDocNo.Value)
        obj.REASON = Reason
        obj.ACTIVITY_TYPE = Activity_Type
        Return clsCancelLog.SaveData(obj, True, trans)
    End Function
    Private Sub gv1_CellEditorInitialized(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellEditorInitialized
        If TypeOf Me.gv1.CurrentColumn Is GridViewMultiComboBoxColumn Then
            Dim editor As RadMultiColumnComboBoxElement = DirectCast(Me.gv1.ActiveEditor, RadMultiColumnComboBoxElement)
            editor.AutoSizeDropDownToBestFit = True
            editor.EditorControl.MasterTemplate.BestFitColumns()
            editor.DropDownStyle = RadDropDownStyle.DropDown
            editor.AutoFilter = True
            If editor.EditorControl.MasterTemplate.FilterDescriptors.Count = 0 Then
                Dim autoFilter As FilterDescriptor = New FilterDescriptor("Description", FilterOperator.StartsWith, "")
                autoFilter.IsFilterEditor = True
                editor.EditorControl.FilterDescriptors.Add(autoFilter)
            End If
        End If
    End Sub
    Private Sub txtDocNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocNo._MYNavigator
        Try
            '-------richa 30/07/2014 Ticket No. BM00000003242---------
            Dim strwherecls As String = ""
            Dim qst As String = ""
            strwherecls = Xtra.CustomerPermission()
            If clsCommon.myLen(strwherecls) = 0 Then
                qst = "select count(*) from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + txtDocNo.Value + "' and  TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC'"
            Else
                qst = "select count(*) from TSPL_SD_SHIPMENT_HEAD where Document_Code='" + txtDocNo.Value + "' and  TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' and TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + strwherecls + ")"
            End If
            '-----------------------------------------------------
            Dim count As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qst))
            If count = 0 Then
                txtDocNo.MyReadOnly = False
            Else
                txtDocNo.MyReadOnly = True
            End If
            LoadData(txtDocNo.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDocNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtDocNo._MYValidating
        ''-------richa 30/07/2014 Ticket No. BM00000003242---------
        '        Dim arrZone As New List(Of String)
        '        Dim ZoneQry As String = " SELECT DISTINCT v.Zone_Code FROM tspl_user_master u JOIN (
        '    SELECT Zone_Code, Vendor_Code FROM tspl_vendor_master  WHERE Vendor_code = (
        '        SELECT Vendor_Code   FROM TSPL_CUSTOMER_VENDOR_MAPPING  WHERE Cust_Code IN (
        '            SELECT Customer_Code FROM TSPL_SD_SHIPMENT_HEAD  ))
        ') v ON u.Vendor_Code = v.Vendor_Code
        'LEFT OUTER JOIN tspl_vendor_master vm ON v.Zone_Code = vm.Zone_Code"
        '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(ZoneQry)
        '        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '            For Each row As DataRow In dt.Rows
        '                arrZone.Add(row("Zone_Code"))
        '            Next
        '        End If



        Dim strwherecls As String = ""
        strwherecls = Xtra.CustomerPermission()
        '-----------------------------------------------------
        Dim qry As String = "select Document_Code as Code,SALE_INVOICE_NO AS [Invoice No],CONVERT(varchar(10), Document_Date,103)+' '+ CONVERT(varchar(5), Document_Date,114) as Date,Customer_Code as [Customer Code], Customer_Name as Customer,TSPL_SD_SHIPMENT_HEAD.Comments,Total_Amt as Amount,
case when TSPL_SD_SHIPMENT_HEAD.Status=0 then 'Pending' else 'Approved' end as [Status],Direct_Dispatch as [Direct Dispatch] ,
TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as[DCS Uploader],
TSPL_SD_SHIPMENT_HEAD.Created_By as [Created By],TSPL_SD_SHIPMENT_HEAD.Created_Date as [Created Date],TSPL_SD_SHIPMENT_HEAD.Sub_Location_code AS [Sub Location],
case when TSPL_SD_SHIPMENT_HEAD.Is_CashSale='Y' then 'Cash' else'Credit' end as SaleType
from TSPL_SD_SHIPMENT_HEAD 
left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code= TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
"
        Dim whrClas As String = ""
        '-------richa 30/07/2014 Ticket No. BM00000003242---------
        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        'End If
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 And clsCommon.myLen(strwherecls) > 0 Then
            whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and  TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' and TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + strwherecls + ")  "
        ElseIf clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and  TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC'  "
        ElseIf clsCommon.myLen(strwherecls) > 0 Then
            whrClas = " TSPL_SD_SHIPMENT_HEAD.Customer_Code in (" + strwherecls + ") and  TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC'   "
        Else
            whrClas = " TSPL_SD_SHIPMENT_HEAD.Trans_Type='MCC' "
        End If

        'If arrZone IsNot Nothing AndAlso arrZone.Count > 0 Then
        '    whrClas += " TSPL_VENDOR_MASTER.Zone_Code In (" + clsCommon.GetMulcallString(arrZone) + ") "
        'End If
        '' 


        '-----------------------------------------------------
        LoadData(clsCommon.ShowSelectForm("ShipmentCofnd", qry, "Code", whrClas, txtDocNo.Value, "Code", isButtonClicked, "Document_Date"), NavigatorType.Current)
    End Sub
    Private Sub FrmAPInvoiceEntry_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing AndAlso gv1.CurrentColumn Is gv1.Columns(colUnit) Then
            isCellValueChangedOpen = True
            gv1.CurrentColumn = gv1.Columns(colIName)
            OpenUOMList(True)
            gv1.CurrentColumn = gv1.Columns(colUnit)
            setGridFocus()
            isCellValueChangedOpen = False
        End If
        If e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing AndAlso clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) <= 0 Then
            isCellValueChangedOpen = True
            If gv1.CurrentColumn Is gv1.Columns(colICode) Then
                gv1.CurrentColumn = gv1.Columns(colIName)
                If blnBackCalculation = True Then
                    OpenICodeList(True)
                Else
                    OpenICodeListCurrentCalaculation(True)
                End If
                gv1.CurrentColumn = gv1.Columns(colICode)
            End If
            setGridFocus()
            isCellValueChangedOpen = False
        ElseIf e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
            AddNew()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso btnSave.Enabled AndAlso MyBase.isModifyFlag Then
            SaveData(False)
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnPost.Enabled AndAlso MyBase.isPostFlag Then
            PostData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso btnDelete.Enabled AndAlso MyBase.isDeleteFlag Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
            CloseForm()
        ElseIf e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.T Then
            chkRateDefaultSetting.Visible = Not chkRateDefaultSetting.Visible
            chkRateUserCustomer.Visible = Not chkRateUserCustomer.Visible
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.SIRC
            frm.strCode = clsFixedParameterCode.SIReversAndCreate
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                '-----------------richa 27/06/2014 Ticket No .BM00000002982------------
                pnlMannualInvoiceNo.Visible = True
                Dim desc As String = ""
                'Dim trans As SqlTransaction
                'trans = clsDBFuncationality.GetTransactin()
                'desc = IIf(clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.InvoiceManualNoWithPrefix & "'")) = 0, False, True)
                desc = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.InvoiceManualNoWithPrefix, clsFixedParameterCode.InvoiceManualNoWithPrefix, Nothing))
                If clsCommon.CompairString(desc, "0") = CompairStringResult.Equal Then
                    txtMannaulInvoiceNo.Visible = True
                    TxtInvoiceManualNoWithPrefix.Visible = False
                Else
                    txtMannaulInvoiceNo.Visible = False
                    TxtInvoiceManualNoWithPrefix.Visible = True
                End If
                '-----------------------------------------------------------
                btnReverseAndUnpost.Visible = True
                rbtnCancel.Visible = True
            End If
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F11 Then
            If RadPageViewPage5.Item.Visibility = ElementVisibility.Visible Then
                RadPageViewPage5.Item.Visibility = ElementVisibility.Collapsed
            Else
                Dim pwd As New FrmPWD(Nothing)
                pwd.strCode = clsFixedParameterCode.MilkSetting
                pwd.strType = clsFixedParameterType.MCCMilkSRNRepost
                pwd.ShowDialog()
                If pwd.isPasswordCorrect Then
                    RadPageViewPage5.Item.Visibility = ElementVisibility.Visible
                End If
            End If
        End If
    End Sub
    Private Sub txtTermCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTermCode._MYValidating
        Dim qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_TERMS_MASTER"
        txtTermCode.Value = clsCommon.ShowSelectForm("ShipmentTermCodefndd", qry, "Code", "", txtTermCode.Value, "Code", isButtonClicked)
        SetTermDetails()
    End Sub
    Sub SetTermDetails()
        Dim qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_TERMS_MASTER where Terms_Code='" + txtTermCode.Value + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Description"))
            txtDueDate.Value = txtDate.Value.AddDays(clsCommon.myCDecimal(dt.Rows(0)("No Of Days")))
        Else
            lblTermName.Text = ""
        End If
    End Sub
    Private Sub txtTaxGroup__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTaxGroup._MYValidating
        'Dim qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
        'txtTaxGroup.Value = clsCommon.ShowSelectForm("Shipmentfndid", qry, "Code", "Tax_Group_Type='S'", txtTaxGroup.Value, "Code", isButtonClicked)
        'SetTaxDetails()
        Try
            If clsCommon.myLen(txtBillToLocation.Value) > 0 Then
                Dim qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
                txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtTaxGroup.Value, isButtonClicked)
                SetTaxDetails()
            Else
                Throw New Exception("Please select Location First")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub SetTax()
    '    GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
    '    If GSTStatus = False OrElse (chkTaxable.Checked AndAlso GSTStatus) Then
    '        txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
    '        lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, Nothing)
    '    Else
    '        If chkTaxable.Checked = False Then
    '            txtTaxGroup.Value = clsLocationWiseTax.GetExempedDefaultTaxGroup(True, txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
    '            lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, Nothing)
    '        End If
    '    End If
    '    SetTaxDetails()
    'End Sub
    Private Sub SetTax()
        Try
            If clsCommon.myLen(txtBillToLocation.Value) <= 0 Then
                Throw New Exception("Please select Location")
            ElseIf clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                Throw New Exception("Please select Customer")
            End If
            txtTaxGroup.Value = clsTaxGroupMaster.GetTaxGroupFromTaxGroupMaster(txtBillToLocation.Value, txtVendorNo.Value, "S", chkTaxable.Checked, txtDate.Value)
            lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, Nothing)
            SetTaxDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub SetTaxDetails()
        LoadBlankGridTax()
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsOfSale(txtTaxGroup.Value, "S")
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If (dt.Rows.Count > 10) Then
                MessageBox.Show("Can't Handle More than 10 Tax Types in a Group")
                Return
            End If
            lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
            For Each dr As DataRow In dt.Rows
                gv2.Rows.AddNew()
                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dr("Tax_Code"))
                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = clsCommon.myCstr(dr("Tax_Code_Desc"))
                If rbtnTaxCalAutomatic.IsChecked Then

                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_TCS  from tspl_tax_master where tax_code ='" & clsCommon.myCstr(dr("Tax_Code")) & "' ")), "Y") = CompairStringResult.Equal Then
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsTCSnotApplicable ,0) from TSPL_CUSTOMER_MASTER where Cust_Code ='" & txtVendorNo.Value & "'")), "0") = CompairStringResult.Equal Then
                            If chkisTCS.Checked Then
                                If AmountToCheckCustomerOutstandingForTCSTax > 0 Then
                                    Dim dblOutstandingAmount As Decimal = clsCommon.myCDecimal(clsCustomerMaster.GetCustomerOutstandingForTCSTaxApplicableOnFY(txtVendorNo.Value, txtDate.Value))
                                    If dblOutstandingAmount < AmountToCheckCustomerOutstandingForTCSTax Then
                                        dblOutstandingAmount = dblOutstandingAmount + clsCommon.myCDecimal(clsCommon.myFormat(lblActualTCSTaxBaseAmt.Text))
                                        If dblOutstandingAmount > AmountToCheckCustomerOutstandingForTCSTax Then
                                            If clsCommon.myCDecimal(clsCommon.myFormat(lblActualTCSTaxBaseAmt.Text)) > 0 Then
                                                txttcstaxbaseamount.Value = Math.Round(clsCommon.myCDecimal(dblOutstandingAmount - AmountToCheckCustomerOutstandingForTCSTax), 2)
                                            End If
                                        End If
                                    End If
                                    If dblOutstandingAmount > AmountToCheckCustomerOutstandingForTCSTax Then
                                        If EnableTCSRateValidityFrom01July2021 Then
                                            Dim Is_ITR_Filled_And_TCSAmountGreater50K As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue(" SELECT CASE WHEN ISNULL(IsTCSGreaterthan50K,0)=1 AND ISNULL(IsITRfilledinLast2Years,0)=1 THEN 1 ELSE 0 END FROM TSPL_CUSTOMER_MASTER WHERE Cust_Code='" & txtVendorNo.Value & "'")) = 1, True, False)
                                            If Is_ITR_Filled_And_TCSAmountGreater50K = True Then
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                                            Else
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                                            End If
                                        Else
                                            Dim panno As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(pan,'')+isnull(Additional3 ,'') as PanNoAdhar from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'"))
                                            If clsCommon.myLen(panno) > 0 Then
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                                            Else
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                                            End If
                                        End If
                                    Else
                                        gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = 0
                                    End If
                                Else
                                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = dr("TaxRate")
                                End If
                            Else
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = 0
                            End If
                            txtTCSTaxRate.Value = clsCommon.myCDecimal(gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value)
                        End If
                    Else
                        gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = dr("TaxRate")
                    End If

                End If
            Next
            SetitemWiseTaxSetting(True, False)
        Else
            lblTaxGrpName.Text = ""
            For ii As Integer = 0 To gv1.Rows.Count - 1
                BlankTaxDetails(ii)
            Next
        End If
        For ii As Integer = 0 To gv1.Rows.Count - 1
            UpdateCurrentRow(ii)
        Next
        UpdateAllTotals()
    End Sub
    Sub SetitemWiseTaxSetting(ByVal isChangeRate As Boolean, ByVal isForCurrentRow As Boolean)
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsOfSale(txtTaxGroup.Value, "S")
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If isForCurrentRow Then
                BlankTaxDetails(gv1.CurrentRow.Index, isChangeRate)
                If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                        If (CalculateTaxRatefromItemwsieTaxOnSale = 0 OrElse clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal) Then
                            If isChangeRate Then
                                gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCDecimal(dr("TaxRate"))
                            End If
                        Else
                            If isChangeRate Then
                                Dim objTM As clsItemWiseTaxAuthority
                                objTM = clsItemWiseTaxAuthority.GetAutoItemwiseTaxRate(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(txtTaxGroup.Value), clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value), txtDate.Value, "S")
                                If objTM IsNot Nothing Then
                                    gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = objTM.TAX_Rate
                                    gv1.CurrentRow.Cells(colItemwiseTaxCode).Value = objTM.HCODE
                                End If
                            End If
                        End If
                        gv1.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        gv1.CurrentRow.Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("RECOVERTABLETAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_Recoverable")), "Y") = CompairStringResult.Equal)
                        gv1.CurrentRow.Cells(clsCommon.myCstr(colIsTaxOnBaseAmount + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_On_Base_Amount")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                End If
            Else
                For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                    BlankTaxDetails(intRowNo, isChangeRate)
                    If clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode).Value) > 0 Then
                        Dim ii As Integer = 1
                        For Each dr As DataRow In dt.Rows
                            Dim strII As String = clsCommon.myCstr(ii)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                            If (CalculateTaxRatefromItemwsieTaxOnSale = 0 OrElse clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal) Then
                                If isChangeRate Then
                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCDecimal(dr("TaxRate"))
                                End If
                            Else
                                If isChangeRate Then
                                    Dim objTM As clsItemWiseTaxAuthority
                                    objTM = clsItemWiseTaxAuthority.GetAutoItemwiseTaxRate(clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colICode).Value), clsCommon.myCstr(txtTaxGroup.Value), clsCommon.myCstr(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value), txtDate.Value, "S")
                                    If objTM IsNot Nothing Then
                                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = objTM.TAX_Rate
                                        gv1.Rows(intRowNo).Cells(colItemwiseTaxCode).Value = objTM.HCODE
                                    End If
                                End If
                            End If
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("RECOVERTABLETAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_Recoverable")), "Y") = CompairStringResult.Equal)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr(colIsTaxOnBaseAmount + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_On_Base_Amount")), "Y") = CompairStringResult.Equal)
                            ii = ii + 1
                        Next
                    End If
                Next
            End If
        End If
    End Sub
    Sub SetitemWiseTaxOnlySetting()
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsscrap(txtTaxGroup.Value)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode)) > 0 Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr(colIsTaxOnBaseAmount + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_On_Base_Amount")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                End If
            Next
        End If
    End Sub
    Private Sub txtVendorNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtVendorNo._MYValidating
        btnHistory.Enabled = True
        Dim qry As String = ""
        Try
            If chkOther.Checked Then
                qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,Customer_Name as Name,TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER.City_Name)>0 then ', '+TSPL_CITY_MASTER.City_Name else ' ' end + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Address
           ,TSPL_CUSTOMER_MASTER.Terms_Code as [Term Code] , TSPL_TERMS_MASTER.Terms_Desc as [Term Description] ,TSPL_CUSTOMER_MASTER.Tax_Group as [Tax Group],TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as [Tax Group Description],Salesman_Code as [Salesman Code],Salesman_Desc as Salesman  ,TSPL_CUSTOMER_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_ROUTE_MASTER.vehicle_code,TSPL_VEHICLE_MASTER.Number ,'' as [VSP Code],''  as [VSP Name],''  as Vlc_Code,'' as [VLc Name]
            from TSPL_CUSTOMER_MASTER  left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code left outer join TSPL_TDS_STATE_MASTER on TSPL_TDS_STATE_MASTER.State_Code=TSPL_CUSTOMER_MASTER.State left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S'
            left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  
            left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id "
                Dim WhrCls As String = " 2=2 and TSPL_CUSTOMER_MASTER.CUSTOMER_FORM_TYPE = 'ALL'  "

                txtVendorNo.Value = clsCommon.ShowSelectForm("AllCustomerInsteadVSP", qry, "Code", WhrCls, txtVendorNo.Value, "Code", isButtonClicked)
                qry += " where 2=2 and TSPL_CUSTOMER_MASTER.Cust_Code  ='" + txtVendorNo.Value + "'"
            Else
                If Not AreaWiseBilling Then
                    If AllowPlandDeptMCCLocation Then
                        If clsCommon.myLen(txtBillToLocation.Value) = 0 Then
                            If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                                txtBillToLocation.Value = clsDBFuncationality.getSingleValue(" select MCC from TSPL_VLC_MASTER_HEAD where (VLC_Code_VLC_Uploader = '" + clsCommon.myCstr(txtVendorNo.Value) + "' or VSP_Code = '" + clsCommon.myCstr(txtVendorNo.Value) + "') ")
                                lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc  from TSPL_LOCATION_MASTER where Location_Code ='" + txtBillToLocation.Value + "'"))
                            End If
                            If clsCommon.myLen(txtBillToLocation.Value) = 0 Then
                                If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                                    clsCommon.MyMessageBoxShow("Please Enter valid customer Code/VLC Uploader Code", Me.Text)
                                    txtVendorNo.Value = ""
                                    txtVendorNo.Focus()
                                Else
                                    clsCommon.MyMessageBoxShow("Please select Location first", Me.Text)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If AllowPlandDeptMCCLocation Then
                    qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,Customer_Name as Name,TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER.City_Name)>0 then ', '+TSPL_CITY_MASTER.City_Name else ' ' end + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Address
           ,TSPL_CUSTOMER_MASTER.Terms_Code as [Term Code] , TSPL_TERMS_MASTER.Terms_Desc as [Term Description] ,TSPL_CUSTOMER_MASTER.Tax_Group as [Tax Group],TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as [Tax Group Description],Salesman_Code as [Salesman Code],Salesman_Desc as Salesman  ,TSPL_CUSTOMER_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_ROUTE_MASTER.vehicle_code,TSPL_VEHICLE_MASTER.Number ,TSPL_VENDOR_MASTER.Vendor_Code as [VSP Code],Vendor_Name as [VSP Name],VLC_Code_VLC_Uploader as Vlc_Code,VLC_Name as [VLc Name] 
            from TSPL_CUSTOMER_MASTER  left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code left outer join TSPL_TDS_STATE_MASTER on TSPL_TDS_STATE_MASTER.State_Code=TSPL_CUSTOMER_MASTER.State left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S'
            left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  
            left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id 
            left join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code 
            inner join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code 
            left join TSPL_VLC_MASTER_HEAD on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code"
                    'left join TSPL_USER_CUSTOMER_ZONE on TSPL_USER_CUSTOMER_ZONE.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code"
                    Dim WhrCls As String = " 2=2 and TSPL_VENDOR_MASTER.Is_Inactive_In_Milk_Procurement=0
            and TSPL_VENDOR_MASTER.Form_Type='VSP' "
                    If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
                        WhrCls += " and TSPL_VENDOR_MASTER.zone_code in (" + objCommonVar.strCurrUserZones + ")"
                    End If
                    'If clsCommon.myLen(objCommonVar.CurrentUserCode) > 0 Then
                    '    WhrCls += " And TSPL_USER_CUSTOMER_ZONE.USER_Code= ('" + objCommonVar.CurrentUserCode + "')"
                    'End If
                    'and TSPL_USER_CUSTOMER_ZONE.USER_Code= ('" + objCommonVar.CurrentUserCode + "')"
                    txtVendorNo.Value = clsCommon.ShowSelectForm("AllCustomerInsteadMcc", qry, "Vlc_Code", WhrCls, txtVendorNo.Value, "Code", isButtonClicked)
                    qry += " where 2=2 and TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader  ='" + txtVendorNo.Value + "'"
                Else
                    If ShowAllCustomer = True AndAlso chkcashsale.Checked = True AndAlso chkOther.Checked = True Then
                        qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,Customer_Name as Name,TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER.City_Name)>0 then ', '+TSPL_CITY_MASTER.City_Name else ' ' end + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Address,TSPL_CUSTOMER_MASTER.Terms_Code as [Term Code] , TSPL_TERMS_MASTER.Terms_Desc as [Term Description] ,TSPL_CUSTOMER_MASTER.Tax_Group as [Tax Group],TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as [Tax Group Description],Salesman_Code as [Salesman Code],Salesman_Desc as Salesman  " &
                       ",TSPL_CUSTOMER_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_ROUTE_MASTER.vehicle_code,TSPL_VEHICLE_MASTER.Number ,TSPL_VENDOR_MASTER.Vendor_Code as [VSP Code],Vendor_Name as [VSP Name],VLC_Code_VLC_Uploader as Vlc_Code,VLC_Name as [VLc Name] "
                        qry += " from TSPL_CUSTOMER_MASTER "
                        qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code"
                        qry += " left outer join TSPL_TDS_STATE_MASTER on TSPL_TDS_STATE_MASTER.State_Code=TSPL_CUSTOMER_MASTER.State"
                        qry += " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code"
                        qry += " left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S'" &
                    "left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  " &
                    "left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id  " &
                    " left join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code " &
                    " left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code " &
                    " left join TSPL_VLC_MASTER_HEAD on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code "
                        '' " and  TSPL_VLC_MASTER_HEAD.mcc='" & txtBillToLocation.Value & "'"
                        txtVendorNo.Value = clsCommon.ShowSelectForm("MCC Customer Lists", qry, "Vlc_Code", " TSPL_CUSTOMER_MASTER.Status = 'N' ", txtVendorNo.Value, "Code", isButtonClicked)
                        qry += " where 2=2 and TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader  ='" + txtVendorNo.Value + "'"
                    Else
                        ''-------richa 30/07/2014 Ticket No. BM00000003242---------
                        Dim strwherecls As String = ""
                        strwherecls = MyBase.Cust_CustomerVendorMapping()
                        If clsCommon.myLen(strwherecls) <= 0 Then
                            clsCommon.MyMessageBoxShow("No Customer Found", Me.Text)
                            Exit Sub
                        End If
                        '-----------------------------------------------------
                        If chkOther.Checked Then
                            qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,Customer_Name as Name,TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER.City_Name)>0 then ', '+TSPL_CITY_MASTER.City_Name else ' ' end + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Address,TSPL_CUSTOMER_MASTER.Terms_Code as [Term Code] , TSPL_TERMS_MASTER.Terms_Desc as [Term Description] ,TSPL_CUSTOMER_MASTER.Tax_Group as [Tax Group],TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as [Tax Group Description],Salesman_Code as [Salesman Code],Salesman_Desc as Salesman  " &
                             ",TSPL_CUSTOMER_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_ROUTE_MASTER.vehicle_code,TSPL_VEHICLE_MASTER.Number ,TSPL_VENDOR_MASTER.Vendor_Code as [VSP Code],Vendor_Name as [VSP Name],VLC_Code_VLC_Uploader as Vlc_Code,VLC_Name as [VLc Name] "
                            qry += " from TSPL_CUSTOMER_MASTER "
                            qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code"
                            qry += " left outer join TSPL_TDS_STATE_MASTER on TSPL_TDS_STATE_MASTER.State_Code=TSPL_CUSTOMER_MASTER.State"
                            qry += " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code"
                            qry += " left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S'" &
                        "left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  " &
                        "left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id  " &
                        " left join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code " &
                        " left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code " &
                        " left join TSPL_VLC_MASTER_HEAD on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code "
                            'qry = "select Vendor_Code AS [Code],Vendor_Name As [Name],Add1 + ' ' + Add2 + ' ' + Add3 AS [Vendor Address],Closing_Date As [Closing Date] " & _
                            '                " ,Vendor_Group_Code As [Vendor Group Code],Vendor_Group_Code_Desc  AS [Vendor Group Code Desc] ,City_Code As [City Code],City_Code_Desc AS [City Code Desc], " & _
                            '                " State,Country ,Phone1 ,Phone2 ,Fax,Email ,WebSite,Contact_Person_Name As [Contact Person Name],Contact_Person_Email As [Contact Person Email] " & _
                            '                " ,Contact_Person_Phone As [Contact Person Phone],Contact_Person_Fax As [Contact Person Fax],Contact_Person_Website AS [Contact Person Website] " & _
                            '                " ,Terms_Code AS [Terms Code],Terms_Code_Desc As [Terms Code Desc],Vendor_Account As [Vendor Account],Vendor_Account_Desc As [Vendor Account Desc],Payment_Code As [Payment Code],Bank_Code As [Bank Code],Tax_Group As [Tax Group],Ven_Type_Code As [Ven Type Code],Status ,OnHold ,CURRENCY_CODE As [Currency Code],Form_Type As [Form Type],State_Code As [State Code],Country_Code AS [Country Code],IFSC_Code As [IFSC Code],Account_Type As [Account Type] from TSPL_VENDOR_MASTER "
                            txtVendorNo.Value = clsCommon.ShowSelectForm("MCC Customer Lists", qry, "Vlc_Code", " TSPL_CUSTOMER_MASTER.Status <> 'Y' ", txtVendorNo.Value, "Code", isButtonClicked)
                            ''txtVendorNo.Value = clsCommon.ShowSelectForm("MCC Customer Lists", qry, "Code", "  TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code is null and TSPL_VENDOR_MASTER.Is_Inactive_In_Milk_Procurement=0", txtVendorNo.Value, "Code", isButtonClicked)
                            qry += " where 2=2 and TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader  ='" + txtVendorNo.Value + "'"
                        Else
                            qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,Customer_Name as Name,TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER.City_Name)>0 then ', '+TSPL_CITY_MASTER.City_Name else ' ' end + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Address,TSPL_CUSTOMER_MASTER.Terms_Code as [Term Code] , TSPL_TERMS_MASTER.Terms_Desc as [Term Description] ,TSPL_CUSTOMER_MASTER.Tax_Group as [Tax Group],TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as [Tax Group Description],Salesman_Code as [Salesman Code],Salesman_Desc as Salesman  " &
                             ",TSPL_CUSTOMER_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_ROUTE_MASTER.vehicle_code,TSPL_VEHICLE_MASTER.Number ,TSPL_VENDOR_MASTER.Vendor_Code as [VSP Code],Vendor_Name as [VSP Name],VLC_Code_VLC_Uploader as Vlc_Code,VLC_Name as [VLc Name] "
                            qry += " from TSPL_CUSTOMER_MASTER "
                            qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code"
                            qry += " left outer join TSPL_TDS_STATE_MASTER on TSPL_TDS_STATE_MASTER.State_Code=TSPL_CUSTOMER_MASTER.State"
                            qry += " left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code"
                            qry += " left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S'" &
                        "left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  " &
                        " left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " &
                        " left join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code " &
                        " left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code " &
                        " left join TSPL_VLC_MASTER_HEAD on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code  "
                            '-------richa 30/07/2014 Ticket No. BM00000003242---------
                            '=============and TSPL_VLC_MASTER_HEAD.VSP_Code is null update by Preeti GUpta Against Ticket No[8543]=======
                            If UseDescInsteadOFCodeOnMCCMAterialSale = True Then
                                LblVlc_Code.Text = clsCommon.ShowSelectForm("MCCCustomerList1", qry, "Vlc_Code", " TSPL_CUSTOMER_MASTER.Cust_Code in (" + strwherecls + ") and TSPL_VENDOR_MASTER.Is_Inactive_In_Milk_Procurement=0 and mcc='" & txtBillToLocation.Value & "'", txtVendorNo.Value, "Vlc_Code", isButtonClicked)
                                LblVlc_Code.Tag = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vlc_Code from TSPL_VLC_MASTER_HEAD  where VLC_Code_VLC_Uploader = '" + clsCommon.myCstr(LblVlc_Code.Text) + "' "))
                                qry += " where 2=2 and VLC_Code_VLC_Uploader ='" + LblVlc_Code.Text + "' and mcc='" & txtBillToLocation.Value & "' "
                            Else
                                'If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                                '    txtVendorNo.Value = clsDBFuncationality.getSingleValue(" select  VSP_Code from TSPL_VLC_MASTER_HEAD  where (VSP_Code = '" + txtVendorNo.Value + "' or VLC_Code_VLC_Uploader = '" + txtVendorNo.Value + "' )")
                                'End If
                                Dim whr As String = " TSPL_CUSTOMER_MASTER.Cust_Code in (" + strwherecls + ") and TSPL_VENDOR_MASTER.Is_Inactive_In_Milk_Procurement=0  "
                                If AllowPlandDeptMCCLocation Then
                                    whr += "and mcc='" & txtBillToLocation.Value & "'"
                                End If
                                txtVendorNo.Value = clsCommon.ShowSelectForm("MCC Customer List", qry, "Vlc_Code", whr, txtVendorNo.Value, "Code", isButtonClicked)
                                qry += " where 2=2 and TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader  ='" + txtVendorNo.Value + "'"
                                If Not AreaWiseBilling Then
                                    txtBillToLocation.Value = clsDBFuncationality.getSingleValue("select mcc from TSPL_VLC_MASTER_HEAD  where VLC_Code_VLC_Uploader = '" + clsCommon.myCstr(txtVendorNo.Value) + "' ")
                                    lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc  from TSPL_LOCATION_MASTER where Location_Code ='" + txtBillToLocation.Value + "'"))
                                End If
                            End If
                        End If
                    End If
                End If


            End If
            '-----------------------------------------------------
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                If UseDescInsteadOFCodeOnMCCMAterialSale = True Then
                    txtVendorNo.Value = clsCommon.myCstr(dt.Rows(0)("Code"))
                End If
                txtVendorNo.Enabled = False
                txtVendorNo.Value = clsCommon.myCstr(dt.Rows(0)("Code"))
                lblVendorName.Text = clsCommon.myCstr(dt.Rows(0)("Name"))
                txtTermCode.Value = clsCommon.myCstr(dt.Rows(0)("Term Code"))
                lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Term Description"))
                txtTaxGroup.Value = clsCommon.myCstr(dt.Rows(0)("Tax Group"))
                lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax Group Description"))
                LblVlc_Code.Text = clsCommon.myCstr(dt.Rows(0)("Vlc_Code"))
                LblVlc_Code.Tag = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Vlc_Code from TSPL_VLC_MASTER_HEAD  where VLC_Code_VLC_Uploader = '" + clsCommon.myCstr(LblVlc_Code.Text) + "' "))
                LblVlc_Name.Text = clsCommon.myCstr(dt.Rows(0)("Vlc Name"))
                'txtDate.Enabled = False
                '' code commented by Panch Raj against ticket no: KDI/17/05/18-000318 after discussion with Ranjana mam
                'txtVendorNo.Enabled = False
                chkTaxable.Enabled = False
                chkRateUserCustomer.ToggleState = ClsUserCustomerSettings.GetUserCustomerRateSetting(txtVendorNo.Value)
                SetMultiCurrencyVisibility()
            Else
                lblVendorName.Text = ""
                txtTermCode.Value = ""
                lblTermName.Text = ""
                txtTaxGroup.Value = ""
                lblTaxGrpName.Text = ""
                LblVlc_Code.Text = ""
                LblVlc_Code.Tag = Nothing
                LblVlc_Name.Text = ""
                Me.txtCurrencyCode.Value = ""
                Me.txtConversionRate.Text = 1
                Me.txtApplicableFrom.Text = ""
            End If
            '''' priti change start here
            qry = "SELECT TSPL_LOCATION_MASTER.Excisable,TSPL_LOCATION_MASTER.State, " &
            "TSPL_LOCATION_MASTER.Sales_Tax_Group as LocalTaxGroup,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc as Local_Tax_GroupName, " &
            "TSPL_LOCATION_MASTER.Sales_Tax_GroupIS as InterstateTaxGroup,TSPL_TAX_GROUP_MASTERIS.Tax_Group_Desc as Interstate_Tax_GroupName " &
            "FROM TSPL_LOCATION_MASTER left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' left outer join TSPL_TAX_GROUP_MASTER as TSPL_TAX_GROUP_MASTERIS on TSPL_TAX_GROUP_MASTERIS.Tax_Group_Code=TSPL_LOCATION_MASTER.Sales_Tax_GroupIS and TSPL_TAX_GROUP_MASTERIS.Tax_Group_Type='S' WHERE TSPL_LOCATION_MASTER.Location_Code = '" + Convert.ToString(txtBillToLocation.Value) + "'"
            Dim dtLocation As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dtLocation IsNot Nothing AndAlso dtLocation.Rows.Count > 0 Then
                Dim loc As String = clsCommon.myCstr(dtLocation.Rows(0)("Excisable"))
                Dim strLocState As String = clsCommon.myCstr(dtLocation.Rows(0)("State"))
                If clsCommon.CompairString(loc, "T") = CompairStringResult.Equal Then
                    strExcise = True
                Else
                    strExcise = False
                End If
                If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                    qry = "select Price_Code,price_CodeNon,State,price_group_code from TSPL_CUSTOMER_MASTER where Cust_Code='" + txtVendorNo.Value + "'"
                    dt = clsDBFuncationality.GetDataTable(qry)
                    If clsCommon.CompairString(loc, "T") = CompairStringResult.Equal OrElse clsCommon.CompairString(loc, "Y") = CompairStringResult.Equal Then
                        txtPriceCode.Text = clsCommon.myCstr(dt.Rows(0)("Price_Code"))
                    Else
                        txtPriceCode.Text = clsCommon.myCstr(dt.Rows(0)("price_CodeNon"))
                    End If
                    '' code commented by Panch Raj against ticket no: KDI/17/05/18-000318 after discussion with Ranjana mam
                    'txtVendorNo.Enabled = False
                    chkTaxable.Enabled = False
                    'If clsCommon.CompairString(clsCommon.myCstr(dtLocation.Rows(0)("State")), clsCommon.myCstr(dt.Rows(0)("State"))) = CompairStringResult.Equal Then
                    '    txtTaxGroup.Value = clsCommon.myCstr(dtLocation.Rows(0)("LocalTaxGroup"))
                    '    lblTaxGrpName.Text = clsCommon.myCstr(dtLocation.Rows(0)("Local_Tax_GroupName"))
                    'Else
                    '    txtTaxGroup.Value = clsCommon.myCstr(dtLocation.Rows(0)("InterstateTaxGroup"))
                    '    lblTaxGrpName.Text = clsCommon.myCstr(dtLocation.Rows(0)("Interstate_Tax_GroupName"))
                    'End If
                End If

                '''' priti change ends here
                SetTax()
                If chkTaxable.Checked = False AndAlso clsCommon.myLen(txtTaxGroup.Value) = 0 Then
                    txtVendorNo.Value = ""
                    clsCommon.MyMessageBoxShow("Please Map exempted Tax Group on Location " & txtBillToLocation.Value)
                    Exit Sub
                End If
            End If
            SetTermDetails()
            setDCSBalance()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtBillToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtBillToLocation._MYValidating
        'Dim obj As clsLocation = clsLocation.FinderForPhysicalLoaction(txtBillToLocation.Value, isButtonClicked)
        'If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Code) > 0 Then0
        '    txtBillToLocation.Value = obj.Code
        '    lblBillToLocation.Text = obj.Name
        'Else
        '    txtBillToLocation.Value = ""
        '    lblBillToLocation.Text = ""
        'End If
        Dim qry As String = Nothing
        Dim WhrCls As String = Nothing
        Dim WhrClsDefLoc As String = Nothing
        Dim Loc As String = clsDBFuncationality.getSingleValue("Select Default_Location from TSPL_USER_MASTER Where User_Code='" + objCommonVar.CurrentUserCode + "'")

        Dim isMainPlant As Boolean = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select IsMainPlant from TSPL_LOCATION_MASTER where Location_Code = '" & Loc & "'") = 1)



        WhrClsDefLoc = " (tspl_location_master.loc_segment_Code = '" & txtBillToLocation.Value & "' or tspl_mcc_master.mcc_Code = '" & txtBillToLocation.Value & "')"
        If AllowPlandDeptMCCLocation Then
            qry = "select Location_Code AS Code,Location_Desc as Name  from TSPL_LOCATION_MASTER"
            WhrCls = " Is_Sub_Location = 'N' AND Location_Category <> 'MCC' and GIT_Type  <> 'Y' "
        Else
            qry = "select distinct Location_Code as Code,Location_Desc as Name from TSPL_LOCATION_MASTER "
            If (Not isMainPlant) Then
                qry += " Left outer join tspl_user_master on tspl_user_master.default_location=tspl_location_master.location_code "
            End If

            WhrCls = " Location_Type='Physical' and CSA_Type='N' and Is_Section='N' " ' and Is_Sub_Location='N' and location_category='MCC' "
            WhrCls += "   and  Location_Code in (" + clsCommon.GetMulcallString(arrMCCRights) + ")"
            If (Not isMainPlant) Then
                'Dim Loc As String = clsDBFuncationality.getSingleValue("Select Default_Location from TSPL_USER_MASTER Where User_Code='" + objCommonVar.CurrentUserCode + "'")
                'If clsCommon.myLen(Loc) > 0 AndAlso clsCommon.CompairString(Loc, "GNG") <> CompairStringResult.Equal Then
                WhrCls += " AND TSPL_USER_MASTER.User_Code ='" + objCommonVar.CurrentUserCode + "'"
                'End If
            End If

        End If
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            WhrCls += "  and  Location_Code in (" + objCommonVar.strCurrUserLocations + ") "
        End If
        If UseDescInsteadOFCodeOnMCCMAterialSale = True Then
            lblBillToLocation.Text = clsCommon.ShowSelectForm("ShipmentMasteidfndr", qry, "Name", WhrCls, txtBillToLocation.Value, "Name", isButtonClicked)
            txtBillToLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_LOCATION_MASTER where Location_Desc ='" + lblBillToLocation.Text + "'"))
        Else


            txtBillToLocation.Value = clsCommon.ShowSelectForm("ShipmentMasteidfndr", qry, "Code", WhrCls, txtBillToLocation.Value, "Code", isButtonClicked)
            lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "'"))
        End If
        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
            txtSubLocation.Enabled = True
        Else
            txtSubLocation.Enabled = False
        End If
        txtPaymentCycleNo.Text = clsGenratePaymentCycles.GetPaymentCycleNo(txtBillToLocation.Value, txtDate.Value)
        txtFiscalYear.Text = clsGenratePaymentCycles.GetPaymentFiscalCode(txtBillToLocation.Value, txtDate.Value)
        Try
            SetTax()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub txtShipToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean)
    '    If clsCommon.myLen(txtBillToLocation.Value) = 0 Then
    '        clsCommon.MyMessageBoxShow("Please select Location first")
    '        txtBillToLocation.Focus()
    '        Exit Sub
    '    End If
    '    If clsCommon.myLen(txtVendorNo.Value) = 0 Then
    '        clsCommon.MyMessageBoxShow("Please select Customer first")
    '        txtVendorNo.Focus()
    '        Exit Sub
    '    End If
    '    Dim qry As String = " select TSPL_SHIP_TO_LOCATION.Ship_To_Code as [Code],TSPL_SHIP_TO_LOCATION.Ship_To_Desc as [Description],TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code as[Customer Code] ,TSPL_SHIP_TO_LOCATION.Ship_To_Type_Desc  as [Customer Discription], replace(case when ISNULL (TSPL_CUSTOMER_MASTER.Add1,'')='' then '' else TSPL_CUSTOMER_MASTER.add1 +',' end + case when ISNULL (TSPL_CUSTOMER_MASTER.Add2,'')='' then '' else TSPL_CUSTOMER_MASTER.add2 +',' end + case when ISNULL (TSPL_CUSTOMER_MASTER.Add3,'')='' then '' else TSPL_CUSTOMER_MASTER.add3 +',' end  ,',,',',') as [Customer Address] ,replace(case when ISNULL (TSPL_SHIP_TO_LOCATION.Add1,'')='' then '' else TSPL_SHIP_TO_LOCATION.add1 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add2,'')='' then '' else TSPL_SHIP_TO_LOCATION.add2 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add3,'')='' then '' else TSPL_SHIP_TO_LOCATION.add3 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add4,'')='' then '' else TSPL_SHIP_TO_LOCATION.add4 +',' end ,',,',',') as [Ship to Address],TSPL_SHIP_TO_LOCATION.CST_No as [CST NO],TSPL_SHIP_TO_LOCATION.Tin_No as [TIN No]  from TSPL_SHIP_TO_LOCATION left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code "
    '    txtShipToLocation.Value = clsCommon.ShowSelectForm("ShipmentShindrlter", qry, "Code", "Ship_To_Type_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' and loc_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'", txtShipToLocation.Value, "Code", isButtonClicked)
    '    'txtShipToLocation.Value = clsShipToLocation.getFinder("Ship_To_Type_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' and loc_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'", txtShipToLocation.Value, isButtonClicked)
    '    lblShipToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Ship_To_Desc from TSPL_SHIP_TO_LOCATION where Ship_To_Code='" + txtShipToLocation.Value + "'"))
    'End Sub
    Private Sub btnRequistionItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectMRNItems()
    End Sub
    Sub SelectMRNItems()
        isInsideLoadData = True
        Dim frm As New frmPendingSaleOrderPS()
        frm.VendorCode = txtVendorNo.Value
        frm.strCurrCode = txtDocNo.Value
        frm.ShowDialog()
        LoadBlankGrid()
        Dim objOrderHead As clsPSSalesOrder = Nothing
        If frm.ArrReturn IsNot Nothing AndAlso frm.ArrReturn.Count > 0 Then
            If clsCommon.myLen(frm.ArrReturn(0).Document_Code) > 0 Then
                objOrderHead = clsPSSalesOrder.GetData(frm.ArrReturn(0).Document_Code, NavigatorType.Current)
                If objOrderHead IsNot Nothing AndAlso clsCommon.myLen(objOrderHead.Document_Code) > 0 Then
                    txtBillToLocation.Value = ""
                    lblBillToLocation.Text = ""
                    chkcashsale.Checked = False
                    '' currency details
                    txtCurrencyCode.Value = objOrderHead.CURRENCY_CODE
                    Me.txtConversionRate.Text = objOrderHead.ConvRate
                    If objOrderHead.ApplicableFrom IsNot Nothing Then
                        Me.txtApplicableFrom.Text = objOrderHead.ApplicableFrom
                    End If
                    If clsCommon.myLen(txtRefNo.Text) <= 0 Then
                        txtRefNo.Text = objOrderHead.Ref_No
                    End If
                    If clsCommon.myLen(txtDesc.Text) <= 0 Then
                        txtDesc.Text = objOrderHead.Description
                    End If
                    If clsCommon.myLen(txtInvNo.Text) <= 0 Then
                        'txtRemarks.Text = objOrderHead.Remarks
                    End If
                    If clsCommon.myLen(txtRoadPermitNo.Text) <= 0 Then
                        txtRoadPermitNo.Text = objOrderHead.Road_Permit_No
                    End If
                    'If (clsCommon.myLen(txtShipToLocation.Value)) <= 0 Then
                    '    txtShipToLocation.Value = objOrderHead.Ship_To_Location
                    'End If
                    If (clsCommon.myLen(cboItemType.SelectedValue) <= 0) Then
                        cboItemType.SelectedValue = objOrderHead.Item_Type
                    End If
                    If (clsCommon.myLen(txtDept.Value) <= 0) Then
                        txtDept.Value = objOrderHead.Dept
                        lblDept.Text = objOrderHead.Dept_Desc
                    End If
                    If (clsCommon.myLen(txtBillToLocation.Value) <= 0) Then
                        txtBillToLocation.Value = objOrderHead.Bill_To_Location
                        lblBillToLocation.Text = objOrderHead.BillToLocationName
                    End If
                    If (clsCommon.myLen(txtVendorNo.Value) <= 0) Then
                        txtVendorNo.Value = frm.VendorCode
                        lblVendorName.Text = frm.VendorName
                        chkVendorGrossReceipt.Checked = clsVendorMaster.isGrossReceipt(frm.VendorCode)
                    End If
                    If (clsCommon.myLen(txtTermCode.Value) <= 0) Then
                        txtTermCode.Value = objOrderHead.Terms_Code
                        lblTermName.Text = objOrderHead.TermsName
                        txtDueDate.Value = objOrderHead.Due_Date
                    End If
                    If (clsCommon.myLen(txtPriceCode.Text)) <= 0 Then
                        txtPriceCode.Text = objOrderHead.Price_Code
                    End If
                    If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                        txtTaxGroup.Value = objOrderHead.Tax_Group
                        SetTaxDetails()
                    End If
                    If (clsCommon.myLen(fndProject.Value) <= 0) Then
                        fndProject.Value = objOrderHead.PROJECT_ID
                        lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")
                        fndProject.Enabled = False
                        lblProject.Enabled = False
                    End If
                    If clsCommon.myLen(txtDiscAmt.Text) <= 0 OrElse clsCommon.myLen(txtDiscPer.Text) <= 0 OrElse clsCommon.myCDecimal(txtDiscAmt.Text) = 0 OrElse clsCommon.myCDecimal(txtDiscPer.Text) = 0 Then
                        If clsCommon.myLen(txtDiscAmt.Text) <= 0 OrElse clsCommon.myLen(txtDiscPer.Text) <= 0 OrElse clsCommon.myCDecimal(txtDiscAmt.Text) = 0 OrElse clsCommon.myCDecimal(txtDiscPer.Text) = 0 Then
                            txtDiscPer.Text = objOrderHead.HeadDisc_Per
                            If clsCommon.myCDecimal(txtDiscPer.Text) = 0 Then
                                txtDiscAmt.Text = objOrderHead.HeadDisc_Amt
                                chkDiscountOnAmt.IsChecked = True
                                lblInvoiceDiscAmt.Text = objOrderHead.HeadDisc_Amt
                            Else
                                chkDiscountOnRate.IsChecked = True
                                lblInvoiceDiscAmt.Text = objOrderHead.HeadDisc_PerAmt
                            End If
                        End If
                    End If
                    LoadBlankGridAC()
                    txtTaxGroup.Enabled = False
                    SETGSTControl()
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code1) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code1
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name1
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt1
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code2) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code2
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name2
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt2
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code3) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code3
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name3
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt3
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code4) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code4
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name4
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt4
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code5) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code5
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name5
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt5
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code6) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code6
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name6
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt6
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code7) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code7
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name7
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt7
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code8) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code8
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name8
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt8
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code9) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code9
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name9
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt9
                    End If
                    If (clsCommon.myLen(objOrderHead.Add_Charge_Code10) > 0) Then
                        gvAC.Rows.AddNew()
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = objOrderHead.Add_Charge_Code10
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = objOrderHead.Add_Charge_Name10
                        gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = objOrderHead.Add_Charge_Amt10
                    End If
                    gvAC.Rows.AddNew()
                End If
            End If
            If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
                gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
            End If
            Dim mrnno As String = ""
            Dim arr As New List(Of String)
            For ii As Integer = 0 To frm.ArrReturn.Count - 1
                If clsCommon.myLen(frm.ArrReturn(ii).Document_Code) > 0 Then
                    Dim strCode As String = frm.ArrReturn(ii).Document_Code
                    'If Not arr.Contains(strCode) Then
                    '    arr.Add(strCode)
                    objOrderHead = clsPSSalesOrder.GetData(frm.ArrReturn(ii).Document_Code, NavigatorType.Current)
                    For Each obj As clsPSSalesOrderDetail In objOrderHead.Arr
                        If (obj.Item_Code = frm.ArrReturn(ii).Item_Code AndAlso obj.Unit_code = frm.ArrReturn(ii).Unit_code AndAlso obj.Scheme_Item = "N") Then ''OrElse (obj.Scheme_Code = frm.ArrReturn(ii).Scheme_Code AndAlso clsCommon.myLen(obj.Scheme_Code) > 0) Then
                            If IsValidItem(obj) Then
                                gv1.Rows.AddNew()
                                txtBillToLocation.Value = obj.Location
                                lblBillToLocation.Text = obj.LocationName
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = obj.Document_Code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = obj.Row_Type
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(obj.Item_Code)
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = obj.Location
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = obj.LocationName
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = frm.ArrReturn(ii).Balance_Qty
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = frm.ArrReturn(ii).Balance_Qty
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgUnit).Value = obj.OrgUnit_code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = frm.ArrReturn(ii).Balance_Qty
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = obj.TAX1_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = obj.TAX2_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = obj.TAX3_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = obj.TAX4_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = obj.TAX5_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = obj.TAX6_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = obj.TAX7_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = obj.TAX8_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = obj.TAX9_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = obj.TAX10_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPer).Value = obj.Disc_Per
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
                                ''gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableRate).Value = obj.Assessable
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = obj.Batch_No
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = obj.HeadDiscAmt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colHeaDDisPer).Value = obj.HeadDiscPer
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDisPerAmt).Value = obj.HeadDiscPerAmt
                                'Dim dt As DataTable = clsMCCMaterialSale.GetOriginalQty(obj.Document_Code, obj.Item_Code, obj.Unit_code, obj.Assessable, obj.MRP, Nothing)
                                'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                '    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgSOQty).Value = clsCommon.myCDecimal(dt.Rows(0)("MRN_Qty"))
                                'End If
                                If obj.MFG_Date.HasValue Then
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colManufactureDate).Value = obj.MFG_Date
                                End If
                                If obj.Expiry_Date.HasValue Then
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colExpiry).Value = obj.Expiry_Date
                                End If
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = IIf(obj.Scheme_Applicable = "Y", "Yes", "No")
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colFromSchemeCode).Value = obj.Scheme_Code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = IIf(obj.Scheme_Item = "Y", "Yes", "No")
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = obj.Item_Tax
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalMRP).Value = obj.Total_MRP_Amt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalBasicAmount).Value = obj.Total_Basic_Amt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalDiscountAmount).Value = obj.Total_Disc_Amt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colcustDiscount).Value = obj.Cust_Discount
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalCustDiscount).Value = obj.Total_Cust_Discount
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colActualCost).Value = obj.ActualRate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColCustDiscountQty).Value = obj.Cust_DiscountQty
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceDateColumn).Value = obj.Price_Date
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceCOde).Value = obj.Price_code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementPer).Value = obj.Abatement_Per
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementAmount).Value = obj.Abatement_Amt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColFOC).Value = obj.FOC_Item
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value), clsCommon.myCDecimal(gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = obj.Batch_No
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colItemWeight).Value = obj.Item_Weight
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colConvF).Value = obj.Conv_Factor
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTotItemWt).Value = obj.TotalItem_Weight
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkupOn).Value = obj.Markup_On
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkUpPercentage).Value = obj.Markup_Percent
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLandingCost).Value = obj.Landing_Cost
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCustDiscPercentage).Value = obj.CustDiscPer
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = obj.HeadDiscAmt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCashDiscSchemeCode).Value = obj.CasdDiscScheme_Code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPurCost).Value = obj.Purchase_Cost
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgCost).Value = obj.OrgRate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleCode).Value = obj.PrincipleCode
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleDesc).Value = obj.PrincipleDesc
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorCode).Value = obj.vendor_code
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorDesc).Value = obj.vendor_desc
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCommRate).Value = obj.Commission_Rate
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommParty).Value = obj.Commission_Party
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommPartyName).Value = clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommParty).Value & "'")
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColCommAmt).Value = obj.Commission_Amt
                                gv1.Rows(gv1.Rows.Count - 1).Cells(ColAmtAfterCOmm).Value = obj.Amt_Less_Commission
                                findQtyandPromoSchemeCode(False)
                            End If
                        End If
                    Next
                    'End If
                End If
                'mrnno = obj.Document_Code
            Next
            If objOrderHead.Arr IsNot Nothing AndAlso objOrderHead.Arr.Count > 0 Then
                For Each objTr As clsPSSalesOrderDetail In objOrderHead.Arr
                    If objTr.Row_Type = "Misc" Then
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = mrnno
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = objTr.Row_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objTr.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objTr.Item_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objTr.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = 0
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = objTr.TAX1_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = objTr.TAX2_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = objTr.TAX3_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = objTr.TAX4_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = objTr.TAX5_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = objTr.TAX6_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = objTr.TAX7_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = objTr.TAX8_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = objTr.TAX9_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = objTr.TAX10_Rate
                    End If
                Next
            End If
            'gv1.Rows.AddNew()
            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
            SetitemWiseTaxSetting(False, False)
            For ii As Integer = 0 To gv1.RowCount - 1
                UpdateCurrentRow(ii)
            Next
        End If
        isInsideLoadData = False
        UpdateAllTotals()
        'RefreshReqNo()
        'LoadParentSHipCode()
    End Sub
    Function IsValidItem(ByVal obj As clsPSSalesOrderDetail)
        If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
            txtTaxGroup.Value = obj.SOTax_Group
            SetTaxDetails()
        End If
        ''If Not clsCommon.CompairString(txtTaxGroup.Value, obj.MRNTax_Group) = CompairStringResult.Equal Then
        ''    common.clsCommon.MyMessageBoxShow("Item : " + obj.Item_Desc + " not Added Current Tax Group : " + txtTaxGroup.Value + " MRN No: " + obj.MRN_No + "  contain Tax Group :" + obj.MRNTax_Group + Environment.NewLine)
        ''    Return False
        ''End If
        For ii As Integer = 0 To gv1.RowCount - 1
            Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
            Dim strReqCode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colOrderNo).Value)
            Dim strUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
            Dim dblMRP As Decimal = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colMRP).Value)
            Dim strSchemeItem As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeItem).Value)
            If clsCommon.myLen(strReqCode) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value), "N") = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Scheme_Item, "N") = CompairStringResult.Equal AndAlso clsCommon.CompairString(strReqCode, obj.Document_Code) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strICode, obj.Item_Code) = CompairStringResult.Equal Then
                Dim strMsg As String = "Already exist at row no:" + clsCommon.myCstr(ii + 1) + Environment.NewLine + "Order No : " + obj.Document_Code + "  Item : " + obj.Item_Desc + Environment.NewLine + ""
                If dblMRP > 0 Then
                    strMsg = strMsg + Environment.NewLine + "MRP : " + clsCommon.myCstr(dblMRP)
                End If
                common.clsCommon.MyMessageBoxShow(strMsg)
                Return False
            End If
        Next
        Return True
    End Function
    Private Sub gv1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv1.DoubleClick
        Try
            If gv1.CurrentColumn Is gv1.Columns(colTotTaxAmt) AndAlso rbtnTaxCalAutomatic.IsChecked Then
                If CalculateTaxRatefromItemwsieTaxOnSale = 0 Then
                    Dim frm As New FrmPOItemTaxDetails()
                    frm.strLineNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colLineNo).Value)
                    frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    frm.dblTotTax = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colTotTaxAmt).Value)
                    frm.dblAmtAfterDis = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colAmtAfterDis).Value)
                    ''New Column for location wise
                    frm.strTaxGroup = txtTaxGroup.Value
                    frm.strTransLocation = txtBillToLocation.Value
                    frm.strTaxType = "S"
                    frm.strVendorCustomerCode = txtVendorNo.Value
                    ''End of New Column for location wise
                    If clsCommon.myLen(frm.strItemCode) > 0 Then
                        frm.ArrIn = New List(Of clsTempItemTaxDetails)
                        For ii As Integer = 1 To 10
                            Dim strii As String = clsCommon.myCstr(ii)
                            Dim obj As New clsTempItemTaxDetails()
                            obj.AuthorityCode = clsCommon.myCstr(gv1.CurrentRow.Cells("COLTAX" + strii).Value)
                            If clsCommon.myLen(obj.AuthorityCode) > 0 Then
                                obj.AuthorityName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Code_Desc from TSPL_TAX_MASTER WHERE Tax_Code='" + obj.AuthorityCode + "'"))
                                obj.Rate = clsCommon.myCDecimal(gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value)
                                obj.BaseAmt = clsCommon.myCDecimal(gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value)
                                obj.TaxAmt = clsCommon.myCDecimal(gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value)
                                obj.isSurTax = clsCommon.myCBool(gv1.CurrentRow.Cells("ISSURTAX" + strii).Value)
                                obj.SurTax = clsCommon.myCstr(gv1.CurrentRow.Cells("SURTAXCODE" + strii).Value)
                                obj.IsTaxable = clsCommon.myCBool(gv1.CurrentRow.Cells("ISTAXABLE" + strii).Value)
                                frm.ArrIn.Add(obj)
                            End If
                        Next
                        frm.ShowDialog()
                        If frm.ArrOut IsNot Nothing AndAlso frm.ArrOut.Count > 0 Then
                            BlankTaxDetails(gv1.CurrentRow.Index)
                            For ii As Integer = 0 To frm.ArrOut.Count - 1
                                Dim strii As String = clsCommon.myCstr(ii + 1)
                                gv1.CurrentRow.Cells("COLTAX" + strii).Value = frm.ArrOut(ii).AuthorityCode
                                gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value = frm.ArrOut(ii).Rate
                                gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value = frm.ArrOut(ii).BaseAmt
                                gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value = frm.ArrOut(ii).TaxAmt
                                gv1.CurrentRow.Cells("ISSURTAX" + strii).Value = frm.ArrOut(ii).isSurTax
                                gv1.CurrentRow.Cells("SURTAXCODE" + strii).Value = frm.ArrOut(ii).SurTax
                                gv1.CurrentRow.Cells("ISTAXABLE" + strii).Value = frm.ArrOut(ii).IsTaxable
                            Next
                            gv1.CurrentRow.Cells(colTotTaxAmt).Value = frm.dblTotTax
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                            UpdateAllTotals()
                        End If
                    End If
                End If
            ElseIf gv1.CurrentColumn Is gv1.Columns(colComplete) AndAlso UsLock1.Status = ERPTransactionStatus.Approved Then
                Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                Dim intSNo As Integer = Convert.ToInt32((clsCommon.myCDecimal(gv1.CurrentRow.Cells(colLineNo).Value)))
                Dim strStatus As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colComplete).Value)
                If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso clsCommon.myLen(strICode) > 0 AndAlso intSNo > 0 AndAlso clsCommon.CompairString(strStatus, "No") = CompairStringResult.Equal Then
                    If common.clsCommon.MyMessageBoxShow("Do you want to complete the item " + clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value), Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                        If clsMCCMaterialSaleDetail.CompleteSRN(txtDocNo.Value, strICode, intSNo) Then
                            common.clsCommon.MyMessageBoxShow("Successfully Completed")
                            LoadData(txtDocNo.Value, NavigatorType.Current)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                myMessages.blankValue(Me, "Invoice not found to Print", Me.Text)
            Else
                clsMCCMaterialSale.funPrint(MyBase.Form_ID, False, txtDate.Value, txtDocNo.Value, False)
                'funPrint()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Function GetAttachQry() As String
    '    Dim QryShowStatus As String = ""
    '    Dim ShowStatusForSale As Decimal = clsDBFuncationality.getSingleValue("SELECT Description  FROM TSPL_FIXED_PARAMETER  WHERE Code ='ShowStatusForSales' And Type ='ShowStatusForSales'")
    '    If clsCommon.CompairString(clsCommon.myCstr(ShowStatusForSale), "1") = CompairStringResult.Equal Then
    '        QryShowStatus = " ,(case when TSPL_SD_SHIPMENT_HEAD.status =1 then 'AUTHORIZED' else 'NOT AUTHORIZED' end) as Status "
    '    Else
    '        QryShowStatus = ""
    '    End If
    '    Dim SerialNo As String = ""
    '    Dim SerialNoColumn As String = ""
    '    Dim ShowSerialNoForSales As Decimal = clsDBFuncationality.getSingleValue("SELECT Description  FROM TSPL_FIXED_PARAMETER  WHERE Code ='ShowSerialNoForSales' And Type ='ShowSerialNoForSales'")
    '    If clsCommon.CompairString(clsCommon.myCstr(ShowSerialNoForSales), "1") = CompairStringResult.Equal Then
    '        SerialNoColumn = " ,1 As SerialNoText , aa.Serial_No As [SerialNo]  "
    '        'SerialNo = "  left outer join TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL  on TSPL_SD_SHIPMENT_DETAIL.Item_Code  =TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.Main_Item_Code And TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.IS_Principle=1 ANd TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL.Location_Code =TSPL_SD_SHIPMENT_DETAIL.Location "
    '        SerialNo = " left outer join (select distinct Doc_No,Serial_No,Main_Item_Code,Location_Code from TSPL_MF_PRINCIPLE_RECEIPT_SERIAL_DETAIL WHERE Is_principle='1' AND ISNULL(Serial_No,'')<>'' and Doc_No in (select Doc_No from TSPL_MF_PRINCIPLE_RECEIPT_HEAD where Status='1'))aa  on TSPL_SD_SHIPMENT_DETAIL.Item_Code  =AA.Main_Item_Code  ANd aa.Location_Code =TSPL_SD_SHIPMENT_DETAIL.Location  "
    '    Else
    '        SerialNoColumn = " ,0 As SerialNoText "
    '        SerialNo = ""
    '    End If
    '    Dim isShowQRcode As Integer = 0
    '    If clsERPFuncationality.GetQRCodeStatus(txtDate.Value) = True AndAlso EnableDynamicQRCodeForB2CInvoice = True AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EInvoice_Type from  TSPL_SD_SALE_INVOICE_HEAD where Against_Shipment_No = '" + txtDocNo.Value + "'")), "BC") = CompairStringResult.Equal Then
    '        isShowQRcode = 1
    '    End If
    '    Dim Qry As String = "  select " + clsCommon.myCstr(isShowQRcode) + " as isShowQRcode, TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type ,  cast(TSPL_SD_SALE_INVOICE_HEAD.BarCode_Img as image) As BarCode_Img, TSPL_SD_SHIPMENT_HEAD.Bill_To_Location, TSPL_LOCATION_MASTER.Location_Desc, TSPL_SD_SHIPMENT_HEAD.Is_Taxable,TSPL_SD_SHIPMENT_HEAD.Electronic_Ref_No,TSPL_SD_SHIPMENT_HEAD.EWayBillNo, convert(varchar,TSPL_SD_SHIPMENT_HEAD.EWayBillDate,103) as EWayBillDate,  StateMasterForLocation.GST_STATE_CODE AS From_GstStateCode," &
    '      " TSPL_LOCATION_MASTER.GSTNO as From_Loc_GstinNo, TSPL_STATE_MASTER.GST_STATE_Code AS Cust_GstStateCode,  TSPL_CUSTOMER_MASTER.GSTNO as Cust_GstInNo,    case when coalesce(p_cust.GST_STATE_CODE,'')='' then TSPL_state_Master.GST_STATE_CODE       when coalesce(p_cust.GST_STATE_CODE,'')<>'' then p_cust .GST_STATE_CODE    end as P_GST_STATE_CODE," &
    '      " case when coalesce(p_cust.P_GSTIN_NO,'')='' then TSPL_CUSTOMER_MASTER .GSTNO  when coalesce(p_cust.P_GSTIN_NO,'')<>'' then p_cust .P_GSTIN_NO end as P_GSTIN_NO,    TSPL_ITEM_MASTER.HSN_Code,  " &
    '      " dtax1.Type as tax1Type,dtax2.Type as tax2Type,dtax3.Type as tax3Type,dtax4.Type as tax4Type,dtax5.Type as tax5Type,dtax6.Type as tax6Type,dtax7.Type as tax7Type,dtax8.Type as tax8Type,dtax9.Type as tax9Type,dtax10.Type as tax10Type, " &
    '    " ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt ,0)as DTax1_Amt,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt ,0)as DTax2_Amt," &
    '   " ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt ,0)as DTax3_Amt,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt ,0)as DTax4_Amt," &
    '    "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt ,0)as DTax5_Amt,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt ,0)as DTax6_Amt," &
    '     " ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt ,0)as DTax7_Amt,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt ,0)as DTax8_Amt," &
    '      "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt ,0)as DTax9_Amt,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt ,0)as DTax10_Amt," &
    '      "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX1_Rate,0)as DTax1_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX2_Rate,0)as DTax2_Rate," &
    '      "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX3_Rate,0)as DTax3_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX4_Rate,0)as DTax4_Rate," &
    '      "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX5_Rate,0)as DTax5_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX6_Rate,0)as DTax6_Rate," &
    '      "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX7_Rate,0)as DTax7_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX8_Rate,0)as DTax8_Rate," &
    '      "ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX9_Rate,0)as DTax9_Rate,ISNULL(TSPL_SD_SHIPMENT_DETAIL.TAX10_Rate,0)as DTax10_Rate," &
    '     " TSPL_LOCATION_MASTER.State as loc_state_code,tspl_location_master.HOAdd1 ,tspl_location_master.HOAdd2,TSPL_SD_SHIPMENT_HEAD.HeadDisc_PerAmt,'" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MM/yyyy") + "' as RunDate,"
    '    Qry += " TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_fOR_Comp.City_Name)>0 then ', '+TSPL_CITY_MASTER_fOR_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER_For_Comp.STATE_NAME  )>0 then ', '+ TSPL_STATE_MASTER_For_Comp.STATE_NAME else ' ' end"
    '    Qry += "  + case when len(TSPL_COMPANY_MASTER.Pincode    )>0 then ', Pin Code - '+ cast(TSPL_COMPANY_MASTER.Pincode as varchar)  else ' ' end"
    '    Qry += "  + case when len(TSPL_COMPANY_MASTER.Tin_No     )>0 then ', Tin No - '+ cast(TSPL_COMPANY_MASTER.Tin_No as varchar)  else ' ' end"
    '    Qry += "  + case when len(TSPL_COMPANY_MASTER.CINNo      )>0 then ', CIN No - '+ cast(TSPL_COMPANY_MASTER.CINNo as varchar)  else ' ' end"
    '    Qry += "  + case when len(TSPL_COMPANY_MASTER.Fax     )>0 then ',Fax '+ TSPL_COMPANY_MASTER.Fax else '' end"
    '    Qry += "+ Case when len(ISNULL(TSPL_COMPANY_MASTER.Phone1,''))>0 and TSPL_COMPANY_MASTER.Phone1='(+__)__________' then '' else ',Phone'+TSPL_COMPANY_MASTER.Phone1 end "
    '    Qry += "+  Case When   ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ',  '+ TSPL_COMPANY_MASTER.Phone2 Else'' End "
    '    Qry += "  + case when len(TSPL_COMPANY_MASTER.Email    )>0 then ',Email - '+ TSPL_COMPANY_MASTER.Email else '' end "
    '    Qry += " as Comp_Address,TSPL_SD_SHIPMENT_HEAD.RoundOffAmount , "
    '    '' Anubhooti 20-Feb-2015 (Show VLC,VSP Name)
    '    Qry += " tspl_vlc_master_head.vlc_code_vlc_uploader As VLC_Code,tspl_vlc_master_head.vlc_name,tspl_vlc_master_head.VSP_Code, "
    '    ''
    '    Qry += "   case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add1  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add1 end as P_Add1, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add2  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add2 end as P_Add2, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Add3  when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_add3 end as P_Add3, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .PIN_Code   when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Pin_No  end as P_PinNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .CST    when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_CST_No   end as P_CstNo,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Tin_No     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .p_Tin_No   end as P_TinNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Email    when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Email  end as P_Email,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Fax     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Fax   end as P_Fax, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .PAN      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Cust_PAN    end as P_Cust_PAN, case when coalesce(p_cust.P_cust_code,'')='' then  TSPL_CUSTOMER_MASTER  .CST      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_CST_No    end as P_CSTNo,case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Lst_No     when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_LST_No    end as P_LstNo, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Cust_Code      when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_code   end as P_CustCode, case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CUSTOMER_MASTER  .Customer_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_cust_name    end as P_Cust_Name,"
    '    Qry += " case when coalesce(p_cust.P_cust_code,'')='' then TSPL_CITY_MASTER   .City_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_City_Name    end as P_City_Name,"
    '    Qry += " case when coalesce(p_cust.P_cust_code,'')='' then TSPL_state_Master.state_Name       when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_state_Name    end as P_State_Name,"
    '    Qry += " case when coalesce(p_cust.P_cust_code,'')='' then     case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When   ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End   when coalesce(p_cust.P_cust_code,'')<>'' then p_cust .P_Phn    end as P_Cust_Phn,TSPL_CUSTOMER_MASTER.Cust_Code ,  TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as Cust_Add1,TSPL_CUSTOMER_MASTER.Add2 as Cust_add2,TSPL_CUSTOMER_MASTER.Add3 as cust_add3,  case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When   ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End  as Cust_Phn,TSPL_CUSTOMER_MASTER.Tin_No  as Cust_TinNo ,TSPL_CUSTOMER_MASTER.CST as Cust_CSTNo,TSPL_CUSTOMER_MASTER.Lst_No as Cust_LSTNo,TSPL_CUSTOMER_MASTER.Email as Cust_Email ,TSPL_CUSTOMER_MASTER.PAN as Customer_PAN,TSPL_CUSTOMER_MASTER.PIN_Code as Cust_PinCode,TSPL_CITY_MASTER.City_Name as Loctn_City_Name_Desc,TSPL_CUSTOMER_MASTER.Fax as Cust_Fax,TSPL_STATE_MASTER .STATE_NAME  as Cust_State_Name, "
    '    Qry += " TSPL_STATE_MASTER.STATE_NAME  as Cust_State_Name ,TSPL_CITY_MASTER_ForCustomer.City_Name  as Cust_City_Name,case when (TSPL_SD_SHIPMENT_HEAD.Dispatch_Terms )='FE' then 'Freight Extra' else  TSPL_SD_SHIPMENT_HEAD.Dispatch_Terms  end  as Dispatch_Terms,TSPL_LOCATION_MASTER .Add1 as Loc_Add1,TSPL_LOCATION_MASTER.Add2 as Loc_ADd2,TSPL_LOCATION_MASTER.Add3  as Loc_Add3, StateMasterForLocation.State_Name as LocationState, TSPL_LOCATION_MASTER.Pin_Code as Loc_Pin_Code,TSPL_LOCATION_MASTER.TIN_No as Loc_TinNo,Case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as  Loc_Phn,TSPL_LOCATION_MASTER.Email as Loc_Email,TSPL_SD_SHIPMENT_HEAD.Total_Add_Charge,convert(varchar,TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Date ,103) as Delivery_Date, TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS,( case when TSPL_SD_SHIPMENT_HEAD.Invoice_Type='R' then 'Retail Invoice' else 'Tax Invoice' end) as Invoice_Type,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Cash_Sale_InvNo,convert(varchar(15),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Cash_Sale_InvDate,TSPL_SD_SHIPMENT_HEAD.Invoice_Type as Inv_Type,TSPL_SD_SHIPMENT_HEAD.Is_CashSale,TSPL_LOCATION_MASTER.Loc_Short_Name,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Inv_Date,103) as Inv_Date,case when (TSPL_SD_SHIPMENT_HEAD.Payment_Terms)='A' then 'Advance' else TSPL_SD_SHIPMENT_HEAD.Payment_Terms end  as Payment_Terms ,TSPL_SD_SHIPMENT_HEAD.Transporter_Name ,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No   ,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,TSPL_COMPANY_MASTER.Tin_No as Comp_Tin_No,TSPL_COMPANY_MASTER.Pan_No as Comp_PANNO,TSPL_COMPANY_MASTER.CST_LST as Comp_CST_No,TSPL_COMPANY_MASTER.CINNo as Comp_CinNo,TSPL_COMPANY_MASTER.Pincode  as Comp_Pin_Code, TSPL_SHIP_TO_LOCATION.Ship_To_Desc as shipName, TSPL_SHIP_TO_LOCATION.add1 as ship_Add1, TSPL_SHIP_TO_LOCATION.Add2 as ship_add2 ,TSPL_SHIP_TO_LOCATION.Add3 as ship_add3  ,TSPL_SHIP_TO_LOCATION.Pin_Code,TSPL_CITY_MASTER.STATE_CODE  ,Tspl_City_master.City_Name,TSPL_EMPLOYEE_MASTER.Emp_Name as SalesManName,TSPL_SD_SHIPMENT_HEAD.Inv_No, TSPL_SD_SHIPMENT_HEAD.Dept_Desc , TSPL_SD_SHIPMENT_HEAD.Remarks ,  TSPL_SD_SHIPMENT_HEAD.Terms_Code,TSPL_SD_SHIPMENT_HEAD.VehicleNo ,TSPL_SD_SHIPMENT_HEAD.Challan_No,TSPL_SD_SHIPMENT_HEAD.Ref_No,TSPL_SD_SHIPMENT_HEAD.RateDiff_Amt,TSPL_SD_SHIPMENT_HEAD.Gross_Amount,TSPL_VENDOR_MASTER.Zone_Code, "
    '    Qry += " TSPL_SD_SHIPMENT_DETAIL .Specification as  specification,   TSPL_SD_SHIPMENT_HEAD.sale_invoice_no as DocNo , TSPL_SD_SHIPMENT_HEAD.Description, "
    '    Qry += "  convert(varchar ,TSPL_SD_SHIPMENT_HEAD .Document_Date,103)as Document_Date , TSPL_SD_SHIPMENT_HEAD.Against_Sales_Order, TSPL_SD_SHIPMENT_HEAD.Item_Type ,  TSPL_SD_SHIPMENT_HEAD.Customer_Code, "
    '    Qry += " TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 as Customer_Add1,TSPL_CUSTOMER_MASTER.add2 as customer_Add2,TSPL_CUSTOMER_MASTER.Add3 as customer_Add3 ,TSPL_CUSTOMER_MASTER.State as customer_city_State,TSPL_CUSTOMER_MASTER.Tin_No  as Cust_Tin_No ,TSPL_CUSTOMER_MASTER.PIN_Code as Customer_Pin_Code , TSPL_SD_SHIPMENT_HEAD .Terms_Code as termscode ,TSPL_SD_SHIPMENT_HEAD .Ref_No as ref_no ,"
    '    Qry += " TSPL_SD_SHIPMENT_HEAD .Comments as comments ,  TSPL_SD_SHIPMENT_HEAD .Discount_Amt as dis_amt,"
    '    If MultiplySubsidyWithQuantity Then
    '        Qry += "  TSPL_SD_SHIPMENT_HEAD.TotalSubsidyDisAmt  as dis_amt1,"
    '    Else
    '        Qry += "  TSPL_SD_SHIPMENT_DETAIL.Total_Disc_Amt  as dis_amt1,"
    '    End If
    '    Qry += " TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount  as aftrdiscount ,TSPL_SD_SHIPMENT_HEAD .Total_Amt as Total_amount,"
    '    Qry += " TSPL_SD_SHIPMENT_HEAD.Discount_Base as bfrdisc_amount, TSPL_COMPANY_MASTER.Access_Officer as FSSAI,TSPL_COMPANY_MASTER.Email,TSPL_COMPANY_MASTER.Tcan_No AS WebSite ,TSPL_COMPANY_MASTER.Phone1 AS COMP_PHONE,  "
    '    Qry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SHIPMENT_HEAD.tax1_amt,0) as txt1amt, TSPL_SD_SHIPMENT_HEAD.TAX1_Rate,  tax2.Tax_Code_Desc as tax2name,isnull (TSPL_SD_SHIPMENT_HEAD.tax2_amt,0) as txt2amt,TSPL_SD_SHIPMENT_HEAD.TAX2_Rate, tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SHIPMENT_HEAD.tax3_amt,0) as txt3amt,TSPL_SD_SHIPMENT_HEAD.TAX3_Rate,   tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SHIPMENT_HEAD.tax4_amt,0) as txt4amt,TSPL_SD_SHIPMENT_HEAD.TAX4_Rate,   tax5.Tax_Code_Desc as tax5name,isnull (TSPL_SD_SHIPMENT_HEAD.tax5_amt,0) as txt5amt,TSPL_SD_SHIPMENT_HEAD.TAX5_Rate,   tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SHIPMENT_HEAD.tax6_amt,0) as txt6amt,TSPL_SD_SHIPMENT_HEAD.TAX6_Rate,   tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SHIPMENT_HEAD.tax7_amt,0) as txt7amt,TSPL_SD_SHIPMENT_HEAD.TAX7_Rate,   tax8.Tax_Code_Desc as tax8name,isnull (TSPL_SD_SHIPMENT_HEAD.tax8_amt,0) as txt8amt,TSPL_SD_SHIPMENT_HEAD.TAX8_Rate,    tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SHIPMENT_HEAD.tax9_amt,0) as txt9amt,TSPL_SD_SHIPMENT_HEAD.TAX9_Rate,   tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SHIPMENT_HEAD.tax10_amt,0) as txt10amt, TSPL_SD_SHIPMENT_HEAD.TAX10_Rate, "
    '    'Qry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SHIPMENT_HEAD.tax1_amt,0) as txt1amt,  "
    '    'Qry += " tax2.Tax_Code_Desc as tax2name,isnull (TSPL_SD_SHIPMENT_HEAD.tax2_amt,0) as txt2amt,  "
    '    'Qry += " tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SHIPMENT_HEAD.tax3_amt,0) as txt3amt,  "
    '    'Qry += " tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SHIPMENT_HEAD.tax4_amt,0) as txt4amt,  "
    '    'Qry += " tax5.Tax_Code_Desc as tax5name,isnull (TSPL_SD_SHIPMENT_HEAD.tax5_amt,0) as txt5amt,  "
    '    'Qry += " tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SHIPMENT_HEAD.tax6_amt,0) as txt6amt,  "
    '    'Qry += " tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SHIPMENT_HEAD.tax7_amt,0) as txt7amt,  "
    '    'Qry += " tax8.Tax_Code_Desc as tax8name,isnull (TSPL_SD_SHIPMENT_HEAD.tax8_amt,0) as txt8amt,   "
    '    'Qry += " tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SHIPMENT_HEAD.tax9_amt,0) as txt9amt,  "
    '    'Qry += " tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SHIPMENT_HEAD.tax10_amt,0) as txt10amt,  "
    '    Qry += " isnull(TSPL_SD_SHIPMENT_HEAD .Total_Tax_Amt,0) as total_tax_amt, TSPL_SD_SHIPMENT_HEAD.Total_Amt as DocAmt,  TSPL_COMPANY_MASTER.Comp_Name as compname,ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+ Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as Phone,TSPL_COMPANY_MASTER.Fax as Comp_Fax,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.GSTReg_No As SellerGST,TSPL_COMPANY_MASTER.Pan_No,ISNULL(tspl_company_Master.ADD1,'') as Comp_add,"
    '    Qry += " TSPL_SD_SHIPMENT_DETAIL.item_code as item_code, TSPL_ITEM_MASTER.Item_Desc + case when Scheme_Item ='Y' then ' (Free Scheme)' else '' end    as itemdesc, TSPL_SD_SHIPMENT_DETAIL.Row_Type,TSPL_SD_SHIPMENT_DETAIL.Qty as qty,TSPL_SD_SHIPMENT_DETAIL.unit_code as uom,TSPL_SD_SHIPMENT_DETAIL.item_cost as itemcost,TSPL_SD_SHIPMENT_DETAIL.amount as amount,TSPL_SD_SHIPMENT_HEAD.TAX1,TSPL_SD_SHIPMENT_HEAD.TAX2,TSPL_SD_SHIPMENT_HEAD.TAX3,TSPL_SD_SHIPMENT_HEAD.TAX4,TSPL_SD_SHIPMENT_HEAD.TAX5,TSPL_SD_SHIPMENT_HEAD.Total_Add_Charge,TSPL_SD_SHIPMENT_HEAD.ReceiverName, TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt "
    '    Qry += " " & QryShowStatus & " "
    '    Qry += " " & SerialNoColumn & "  "
    '    Qry += " from TSPL_SD_SHIPMENT_DETAIL   "
    '    Qry += " " & SerialNo & " "
    '    Qry += " left outer join TSPL_SD_SHIPMENT_HEAD  on TSPL_SD_SHIPMENT_HEAD.Document_Code  =TSPL_SD_SHIPMENT_DETAIL.Document_Code   "
    '    Qry += " left outer join  TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code =TSPL_SD_SHIPMENT_HEAD .Ship_To_Location "
    '    Qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_SHIP_TO_LOCATION.City_Code "
    '    Qry += " left outer join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE =TSPL_SD_SHIPMENT_HEAD.Salesman_Code "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SHIPMENT_HEAD.tax1  "
    '    Qry += " left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SHIPMENT_HEAD.tax2  "
    '    Qry += " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SHIPMENT_HEAD .TAX3  "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SHIPMENT_HEAD .tax4  "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SHIPMENT_HEAD .tax5  "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX6  "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX7  "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX8  "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX9 "
    '    Qry += " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX10     "
    '    Qry += "left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_SD_SHIPMENT_DETAIL.tax1  " &
    '        " left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_SD_SHIPMENT_DETAIL.tax2   " &
    '        " left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_SD_SHIPMENT_DETAIL.TAX3  " &
    '        " left outer join TSPL_TAX_MASTER as dtax4 on dtax4.Tax_Code= TSPL_SD_SHIPMENT_DETAIL.tax4  " &
    '         "  left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_SD_SHIPMENT_DETAIL.tax5  " &
    '           " left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX6  " &
    '            " left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX7  " &
    '             "  left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX8   " &
    '            " left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX9 " &
    '        "left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_SD_SHIPMENT_DETAIL.TAX10  "
    '    Qry += " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_SD_SHIPMENT_HEAD.comp_code  "
    '    Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SHIPMENT_HEAD.Customer_Code  left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = TSPL_CUSTOMER_MASTER.Cust_Code   LEFT join (select TSPL_CUSTOMER_MASTER.GSTNO AS P_GSTIN_NO,GST_STATE_CODE, pan as P_Cust_PAN, Cust_Code as P_cust_code,Customer_Name as P_cust_name,Add1 as P_cust_add1 ,Add2  as P_cust_add2,add3  as P_cust_add3,PIN_Code as P_Pin_No,Tin_No as p_Tin_No ,State as P_state,Email as P_Email,fax as P_Fax,TSPL_CITY_MASTER.City_Name as  P_City_Name,TSPL_STATE_MASTER.STATE_NAME as P_State_Name  ,case when ISNULL(Phone1,'')='(+__)__________' then '' else Phone1 end +  Case When   ISNULL(Phone2,'')<>'(+__)__________' Then ', '+ Phone2 Else'' End as  P_Phn,CST as P_CST_No,Terms_Code as P_Terms,Lst_No as P_LST_No  from TSPL_CUSTOMER_MASTER   left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_CUSTOMER_MASTER.City_Code "
    '    Qry += " left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State =TSPL_STATE_MASTER.STATE_CODE  "
    '    Qry += " ) p_cust on p_cust.P_cust_code=TSPL_CUSTOMER_MASTER.Parent_Customer_No and TSPL_CUSTOMER_MASTER.Parent_Customer_YN='N'"
    '    '' Anubhooti 20-Feb-2015 (Show VLC,VSP Name)
    '    Qry += "  LEFT OUTER JOIN tspl_customer_vendor_mapping ON tspl_customer_vendor_mapping.Cust_Code =TSPL_CUSTOMER_MASTER.Cust_Code   "
    '    Qry += "LEFT OUTER JOIN tspl_vlc_master_head ON tspl_customer_vendor_mapping.Vendor_Code =tspl_vlc_master_head.vsp_code"
    '    Qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_SD_SHIPMENT_HEAD.Bill_To_Location "
    '    Qry += " left outer join TSPL_CITY_MASTER  as TSPL_CITY_MASTER_ForCustomer on TSPL_CITY_MASTER_ForCustomer.City_Code =TSPL_CUSTOMER_MASTER.City_Code"
    '    Qry += " left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE = TSPL_CUSTOMER_MASTER.State"
    '    Qry += " Left Outer Join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code "
    '    Qry += " left outer join TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE on TSPL_DELIVERY_ORDER_HEAD_PRODUCTSALE.Document_Code =TSPL_SD_SHIPMENT_HEAD.Delivery_Code_PS"
    '    Qry += " LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code "
    '    Qry += " LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State " &
    '    " LEFT OUTER JOIN TSPL_STATE_MASTER StateMasterForLocation ON StateMasterForLocation.State_Code=TSPL_LOCATION_MASTER.State " &
    '    " left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code "
    '    Qry += "  where 2=2 "
    '    Qry += "  and  TSPL_SD_SHIPMENT_HEAD.Document_Code = '" + txtDocNo.Value + "'"
    '    Return Qry
    'End Function
    'Private Sub funPrint()
    '    Try
    '        '====================================================
    '        'Dim strUrl As String = "upi://pay?pa=1234566@ICIC12345.ifsc.npci&pn=TEST&mc=0000&tr=API123456&am=10000&mam=10000&cu=INR&mode=01&b2cSellerGstin=29AABCT1332L000&b2cUPIID=null&b2cBankAcNo=1234566&b2cIFSCCode=ICIC12345&b2cInvNo=API123456&b2cInvDate=10-06-2021&b2cCGSTAmount=0&b2cSGSTAmount=0&b2cIGSTAmount=0&b2cCESS=0&size=150"
    '        'Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(strUrl)
    '        'clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_SD_SALE_INVOICE_HEAD", "TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No='" & txtDocNo.Value & "'")
    '        '===================================================
    '        Dim IsMandiTax As Decimal = 0
    '        IsMandiTax = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code='" & txtTaxGroup.Value & "' and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y')"))
    '        Dim IsExampted As Decimal = 0
    '        IsExampted = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select COUNT(Is_Tax_Exempted)  from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" & txtTaxGroup.Value & "' AND Tax_Group_Type='S' AND Is_Tax_Exempted=1"))
    '        atchqry = GetAttachQry()
    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)
    '        If dt.Rows.Count > 0 Then
    '            ' ==============================Ticket No  ERO/27/11/18-000421 By Prabhakar for  Customer Dashboard on Print==============================================
    '            Dim dtCustomerOutstanding As DataTable = Nothing
    '            dtCustomerOutstanding = clsCustomerMaster.getCustomerOutstandingOfAmt_Can_Crate("'" & clsCommon.myCstr(dt.Rows(0)("Customer_Code")) & "'", clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")).AddDays(-1), "dd/MMM/yyyy"), clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")), "dd/MMM/yyyy"))
    '            '=============================================================================
    '            Dim frmCRV As New frmCrystalReportViewer()
    '            If clsERPFuncationality.GetGSTStatus(clsCommon.myCDate(dt.Rows(0)("Document_Date"))) Then
    '                If clsCommon.myCDecimal(dt.Rows(0)("Is_Taxable")) = 1 Then
    '                    If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("LocationState")), clsCommon.myCstr(dt.Rows(0)("Cust_State_Name"))) = CompairStringResult.Equal Then
    '                        If IsExampted > 0 Then
    '                            frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
    '                        ElseIf IsMandiTax > 0 Then
    '                            frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxables", "MCC Material Sale Local With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
    '                            'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local_WithMandi", "MCC Material Sale Local With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
    '                        Else
    '                            frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
    '                            'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
    '                        End If
    '                    Else
    '                        If IsExampted > 0 Then
    '                            frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
    '                        ElseIf IsMandiTax > 0 Then
    '                            frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxables", "MCC Material Sale Interstate With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
    '                            'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState_WithMandi", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
    '                        Else
    '                            frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
    '                        End If
    '                    End If
    '                Else
    '                    frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
    '                End If
    '            Else
    '                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Is_CashSale")), "Y") = CompairStringResult.Equal Then
    '                    frmCRV.funreport(CrystalReportFolder.MilkProcurement, dt, "rptMccMaterialSaleCashMemo", "Cash Memo")
    '                Else
    '                    frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale", "MCC Material Sale", "rptCompanyAddress.rpt")
    '                End If
    '            End If
    '            'frmCrystalReportViewer.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale", "MCC Material Sale", "rptCompanyAddress.rpt")
    '            frmCRV = Nothing
    '        End If
    '    Catch ex As Exception
    '        common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Private Function SetItemWiseTax(ByVal dtAfterModify As DataTable, ByVal strShipFrm As String) As DataTable
        dtAfterModify.Columns.Add("TAX1_Rate1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX1_Rate2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX1_Rate3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX1_Amt1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX1_Amt2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX1_Amt3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX2_Rate1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX2_Rate2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX2_Rate3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX2_Amt1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX2_Amt2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX2_Amt3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX3_Rate1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX3_Rate2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX3_Rate3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX3_Amt1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX3_Amt2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX3_Amt3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX4_Rate1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX4_Rate2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX4_Rate3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX4_Amt1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX4_Amt2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX4_Amt3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX5_Rate1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX5_Rate2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX5_Rate3", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX5_Amt1", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX5_Amt2", GetType(Decimal))
        dtAfterModify.Columns.Add("TAX5_Amt3", GetType(Decimal))
        Dim qry As String = "select Tax,Rate,SUM(Amt) as TaxAmt"
        qry += " from ("
        qry += " select TAX1 as Tax,TAX1_Rate as Rate,TAX1_Amt as Amt"
        qry += " from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strShipFrm + "' "
        qry += " union all "
        qry += " select TAX2 as Tax,TAX2_Rate as Rate,TAX2_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX3 as Tax,TAX3_Rate as Rate,TAX3_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX4 as Tax,TAX4_Rate as Rate,TAX4_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX5 as Tax,TAX5_Rate as Rate,TAX5_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strShipFrm + "'   "
        qry += " union all "
        qry += " select TAX6 as Tax,TAX6_Rate as Rate,TAX6_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_DETAIL where Document_Code='" + strShipFrm + "'   "
        qry += " )xxx "
        qry += " group by Tax,Rate   having SUM(Amt)>0   "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                For ii As Integer = 1 To 5
                    Dim strCol As String = "TAX" + clsCommon.myCstr(ii) + ""
                    If clsCommon.CompairString(clsCommon.myCstr(dtAfterModify.Rows(0)(strCol)), clsCommon.myCstr(dr("Tax"))) = CompairStringResult.Equal Then
                        If clsCommon.myCDecimal(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt1")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate1") = clsCommon.myCDecimal(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt1") = clsCommon.myCDecimal(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCDecimal(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt2")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate2") = clsCommon.myCDecimal(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt2") = clsCommon.myCDecimal(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCDecimal(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt3")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate3") = clsCommon.myCDecimal(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt3") = clsCommon.myCDecimal(dr("TaxAmt"))
                            Next
                        Else
                            Throw New Exception("Printing Support only 3 Diffent Rates")
                        End If
                    End If
                Next
            Next
        End If
        Return dtAfterModify
    End Function
    Private Sub gv1_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv1.CellFormatting
        Try
            If e.Column.Index >= 0 Then
                If (e.Column Is gv1.Columns(colExpiry)) OrElse (e.Column Is gv1.Columns(colManufactureDate)) Then
                    gv1.Columns(colExpiry).FormatString = "{0:d}"
                ElseIf e.Column Is gv1.Columns(colICode) OrElse e.Column Is gv1.Columns(colMRP) Then
                    If clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) > 0 Then
                        gv1.CurrentRow.Cells(colICode).ReadOnly = True
                        gv1.CurrentRow.Cells(colMRP).ReadOnly = True
                        'gv1.CurrentRow.Cells(colOrgSOQty).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colICode).ReadOnly = False
                        gv1.CurrentRow.Cells(colMRP).ReadOnly = False
                        'gv1.CurrentRow.Cells(colOrgSOQty).ReadOnly = False
                    End If
                ElseIf e.Column Is gv1.Columns(colQty) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
                        gv1.CurrentRow.Cells(colQty).ReadOnly = False
                        gv1.CurrentRow.Cells(colFreeQty).ReadOnly = False
                    Else
                        gv1.CurrentRow.Cells(colQty).ReadOnly = True
                        gv1.CurrentRow.Cells(colFreeQty).ReadOnly = True
                    End If
                ElseIf e.Column Is gv1.Columns(colUnit) Then
                    'If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) <= 0 Then
                    '    gv1.CurrentRow.Cells(colUnit).ReadOnly = False
                    'Else
                    '    gv1.CurrentRow.Cells(colUnit).ReadOnly = True
                    'End If
                ElseIf e.Column Is gv1.Columns(colRate) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
                        If chkRateUserCustomer.ToggleState = ToggleState.Indeterminate Then
                            gv1.CurrentRow.Cells(colRate).ReadOnly = Not chkRateDefaultSetting.Checked
                        Else
                            gv1.CurrentRow.Cells(colRate).ReadOnly = Not chkRateUserCustomer.Checked
                        End If
                        If ItemRateEditable Then
                            gv1.CurrentRow.Cells(colRate).ReadOnly = False
                        Else
                            If CreateAutoMCCPriceChat Then
                                gv1.CurrentRow.Cells(colRate).ReadOnly = False
                            Else
                                gv1.CurrentRow.Cells(colRate).ReadOnly = True
                            End If
                        End If
                    Else
                        If CreateAutoMCCPriceChat Then
                            gv1.CurrentRow.Cells(colRate).ReadOnly = False
                        Else
                            gv1.CurrentRow.Cells(colRate).ReadOnly = True
                        End If
                    End If
                ElseIf e.Column Is gv1.Columns(ColCommParty) Then
                    'If chkCommApply.Checked And clsCommon.myLen(txtReqNo.Value) = 0 Then
                    '    gv1.CurrentRow.Cells(ColCommParty).ReadOnly = False
                    '    gv1.CurrentRow.Cells(colCommRate).ReadOnly = False
                    'Else
                    gv1.CurrentRow.Cells(ColCommParty).ReadOnly = True
                    gv1.CurrentRow.Cells(colCommRate).ReadOnly = True
                    ' End If
                ElseIf e.Column Is gv1.Columns(colDisPer) Then
                    If ItemMRPEditable Then
                        gv1.CurrentRow.Cells(colMRP).ReadOnly = False
                    Else
                        gv1.CurrentRow.Cells(colMRP).ReadOnly = True
                    End If
                ElseIf e.Column Is gv1.Columns(colAmt) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal OrElse clsCommon.myCDecimal(gv1.CurrentRow.Cells(colIsMannualAmt).Value) = 1 Then
                        gv1.CurrentRow.Cells(colAmt).ReadOnly = False
                    Else
                        gv1.CurrentRow.Cells(colAmt).ReadOnly = True
                    End If
                    ''ElseIf (e.Column Is gv1.Columns(colOrgPOQty) OrElse e.Column Is gv1.Columns(colOrgSOQty) OrElse e.Column Is gv1.Columns(colOrgGRNQty)) Then
                    ''    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso isPO_GRN_MRN_Editable Then
                    ''        gv1.CurrentRow.Cells(colOrgPOQty).ReadOnly = False
                    ''        gv1.CurrentRow.Cells(colOrgGRNQty).ReadOnly = False
                    ''        gv1.CurrentRow.Cells(colOrgSOQty).ReadOnly = False
                    ''    Else
                    ''        gv1.CurrentRow.Cells(colQty).ReadOnly = True
                    ''    End If
                End If
                'Dim cell As GridDataCellElement = TryCast(e.CellElement, GridDataCellElement)
                'cell.GradientStyle = GradientStyles.Solid
                'cell.BackColor = Color.FromArgb(243, 181, 51)
            End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        End Try
    End Sub
    Private Sub gv1_UserDeletingRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub
    Private Sub gv1_UserDeletedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserDeletedRow
        UpdateAllTotals()
        For ii As Integer = 1 To gv1.Rows.Count
            gv1.Rows(ii - 1).Cells(colLineNo).Value = ii
        Next
        'RefreshReqNo()
    End Sub
    Private Sub txtDept__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtDept._MYValidating
        Try
            Dim obj As clsDepartment = clsDepartment.Finder(txtDept.Value, isButtonClicked)
            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Code) > 0 Then
                txtDept.Value = obj.Code
                lblDept.Text = obj.Name
            Else
                txtDept.Value = ""
                lblDept.Text = ""
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtReqNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean)
        SelectMRNItems()
    End Sub
    'Sub RefreshReqNo()
    '    txtReqNo.Value = ""
    '    If gv1.Rows.Count > 0 Then
    '        For ii As Integer = 0 To gv1.Rows.Count - 1
    '            Dim strReqNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colOrderNo).Value)
    '            If clsCommon.myLen(strReqNo) > 0 Then
    '                txtReqNo.Value = clsCommon.myCstr(strReqNo)
    '                Exit Sub
    '            End If
    '        Next
    '    End If
    'End Sub
    Public Sub funrejprint()
        Try
            Dim qry As String = "SELECT     TSPL_SD_SHIPMENT_HEAD.Document_Code, TSPL_SD_SHIPMENT_HEAD.Document_Date, TSPL_SD_SHIPMENT_HEAD.Customer_Name, TSPL_SD_SHIPMENT_HEAD.Ship_To_Location, TSPL_SD_SHIPMENT_HEAD.Bill_To_Location, TSPL_SD_SHIPMENT_HEAD.RMDA_No, TSPL_SD_SHIPMENT_HEAD.RMDA_Date,TSPL_SD_SHIPMENT_HEAD.Remarks,TSPL_SD_SHIPMENT_HEAD.Description, TSPL_SRN_DETAIL.Item_Code,   TSPL_SRN_DETAIL.Item_Desc, TSPL_SRN_DETAIL.Rejected_Qty, TSPL_SRN_DETAIL.Item_Cost,TSPL_SRN_DETAIL.Unit_code,TSPL_SRN_DETAIL.Rejected_Qty*TSPL_SRN_DETAIL.Item_Cost as Amount,TSPL_COMPANY_MASTER.comp_name,TSPL_COMPANY_MASTER.Logo_Img, TSPL_COMPANY_MASTER.Logo_Img2  FROM         TSPL_SD_SHIPMENT_HEAD INNER JOIN    TSPL_SRN_DETAIL ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SRN_DETAIL.Document_Code LEFT OUTER JOIN     TSPL_COMPANY_MASTER ON TSPL_SD_SHIPMENT_HEAD.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code  where TSPL_SD_SHIPMENT_HEAD.Document_Code='" + txtDocNo.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            Dim frmCRV As New frmCrystalReportViewer()
            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptMRDA", "MRDA Report")
            frmCRV = Nothing
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Function GetItemType(ByVal strItmType As String) As String
        Dim qry As String = "select distinct Item_Type  from TSPL_ITEM_MASTER where Item_Code ='" + strItmType + "'"
        strItmType = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
        If strItmType = "F" Then
            strItmType = 0
        Else
            strItmType = 1
        End If
        Return strItmType
    End Function
    Public Function GetTaxGrp(ByVal strItmType As String) As String
        Dim qry As String = "select Tax_Group  from TSPL_VENDOR_MASTER where Customer_Code ='" + strItmType + "'"
        strItmType = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
        Return strItmType
    End Function
    Private Function GetConvRate(ByVal IntRowNo As Integer) As Decimal
        Dim strItem As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
        Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colOrgUnit).Value)
        Dim strCurrentUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
        Dim dblCurrentConvF As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colConvF).Value)
        Dim dblRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colRate).Value)
        Dim dblConvRate As Decimal = 0
        If clsCommon.myLen(strOrgUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strOrgUnit & "'"))
            Dim dblStockingUnitConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and Stocking_Unit='Y' "))
            dblConvRate = (dblRate / dblOrgConvF) * dblStockingUnitConvF * dblCurrentConvF
        End If
        Return dblConvRate
    End Function
    Private Function GetConvMRP(ByVal IntRowNo As Integer) As Decimal
        Dim strItem As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
        Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colOrgUnit).Value)
        Dim strCurrentUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
        Dim dblCurrentConvF As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colConvF).Value)
        Dim dblRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colRate).Value)
        Dim dblMRP As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colMRP).Value)
        Dim dblConvRate As Decimal = 0
        Dim dblConvMRP As Decimal = 0
        If clsCommon.myLen(strOrgUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strOrgUnit & "'"))
            Dim dblStockingUnitConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and Stocking_Unit='Y' "))
            dblConvMRP = (dblMRP / dblOrgConvF) * dblStockingUnitConvF * dblCurrentConvF
        End If
        Return dblConvMRP
    End Function
    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer)
        Try
            Dim arrTaxableAuth As New List(Of String)
            Dim arrTaxableAuth1 As New List(Of String)
            Dim dblFAmt As Decimal = 0
            'If chkVendorGrossReceipt.Checked OrElse clsCommon.CompairString(cboItemType.SelectedValue, "F") = CompairStringResult.Equal Then
            '    dblQty = dblQty
            'End If
            'If clsCommon.myLen(strFCode) > 0 Then
            '    dblFAmt = dblFRate * dblLeak
            'End If
            Dim dblQty As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colQty).Value)
            Dim dblRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblMRP As Decimal = 0
            Dim dblBasicRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblConvF As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colConvF).Value)
            Dim dblItemWeight As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colItemWeight).Value)
            Dim dblheadDiscamt As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colHeadDiscamt).Value)
            'Dim dblOrgBasicRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colOrgCost).Value)
            'Dim dblConvBasicRate As Decimal
            'If dblConvF = 1 And dblOrgBasicRate Then
            '    dblConvBasicRate = dblOrgBasicRate
            'Else
            '    dblConvBasicRate = dblOrgBasicRate * dblConvF
            'End If
            'Dim dblBasicAmt As Decimal = dblQty * dblConvBasicRate
            'dblBasicRate = dblConvBasicRate
            'dblRate = dblBasicRate
            'dblRate = clsEkoPro.GetRateMccSale(txtBillToLocation.Value, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value), dtpChallan.Value) 'GetConvRate(IntRowNo)
            dblMRP = GetConvMRP(IntRowNo)
            Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblBasicAmt As Decimal = dblQty * dblRate
            Dim dblMRPAmt As Decimal = dblQty * dblMRP
            Dim dblAmt As Decimal = (dblQty * dblRate) ''+ dblFAmt
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colIsMannualAmt).Value) = 0 Then
                'gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 6)
                gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2)
            Else
                dblAmt = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colAmt).Value)
            End If
            ''''' to calculate customer disc
            Dim dt As New DataTable
            Dim dblOrderQty As Decimal = 0
            Dim dblCustDiscQty As Decimal = 0
            Dim dblCustDiscAmt As Decimal = 0
            Dim dblCustDiscPercentage As Decimal = 0
            Dim dblApplyCustDisc As Decimal = 0
            Dim dblTotCustDisc As Decimal = 0
            Dim strICode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
            If clsCommon.myLen(strICode) > 0 And clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0 Then
                StrSql = "select top 1 Item_Qty,Amount,percentage,TSPL_SCHEME_MASTER_NEW.scheme_code from TSPL_SCHEME_MASTER_NEW left outer join TSPL_SCHEME_BENEFICIARY on  " &
            "TSPL_SCHEME_MASTER_NEW.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code where Scheme_Type='Cash' and Cust_Code='" & txtVendorNo.Value & "' and " &
            "Start_Date < = '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' and Item_Code='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and " &
            "Unit_Code='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "' and " &
            "MRP='" & clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colMRP).Value) & "' " &
            "and Basic_Price='" & clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colRate).Value) & "'  order by TSPL_SCHEME_MASTER_NEW.scheme_code desc "
                dt = clsDBFuncationality.GetDataTable(StrSql)
                If dt IsNot Nothing And dt.Rows.Count > 0 Then
                    dblOrderQty = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colQty).Value)
                    dblCustDiscQty = clsCommon.myCDecimal(dt.Rows(0)("Item_Qty"))
                    dblCustDiscAmt = clsCommon.myCDecimal(dt.Rows(0)("Amount"))
                    dblCustDiscPercentage = clsCommon.myCDecimal(dt.Rows(0)("percentage"))
                    If dblCustDiscPercentage = 0 Then
                        gv1.Rows(IntRowNo).Cells(ColCustDiscountQty).Value = dblCustDiscQty
                        gv1.Rows(IntRowNo).Cells(colcustDiscount).Value = dblCustDiscAmt
                        dblApplyCustDisc = Math.Truncate(dblOrderQty / dblCustDiscQty)
                        dblTotCustDisc = dblApplyCustDisc * dblCustDiscAmt
                    Else
                        gv1.Rows(IntRowNo).Cells(ColCustDiscountQty).Value = dblCustDiscQty
                        gv1.Rows(IntRowNo).Cells(colCustDiscPercentage).Value = dblCustDiscPercentage
                        dblApplyCustDisc = (dblAmt * dblCustDiscPercentage) / 100
                        dblTotCustDisc = dblApplyCustDisc
                    End If
                    If dblTotCustDisc > 0 Then
                        gv1.Rows(IntRowNo).Cells(colCashDiscSchemeCode).Value = clsCommon.myCstr(dt.Rows(0)("scheme_code"))
                    End If
                    gv1.Rows(IntRowNo).Cells(colTotalCustDiscount).Value = dblTotCustDisc
                End If
            End If
            ''''' end 
            Dim dblDisPer As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            Dim dblDisAmt As Decimal = (dblAmt * dblDisPer) / 100
            Dim dblDisPerUnit As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colDisPerUnit).Value)
            Dim dblDisUnitAmt As Decimal = dblDisPerUnit * dblQty
            Dim dblHeadDisPer As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colHeaDDisPer).Value)
            Dim dblHeadPerDisAmt As Decimal = (dblAmt * dblHeadDisPer) / 100
            Dim dblTotDiscAmt = dblheadDiscamt + dblHeadPerDisAmt + dblDisAmt + dblDisUnitAmt
            Dim dblAmtAfterDis As Decimal = dblAmt - dblDisAmt - dblTotCustDisc - dblheadDiscamt - dblHeadPerDisAmt
            Dim dblAbatementRate As Decimal = gv1.CurrentRow.Cells(colAbatementPer).Value
            Dim dblAbatementAmt As Decimal = ((dblMRP * dblAbatementRate) / 100) * dblQty
            Dim dblTPTRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColTPTRate).Value)
            Dim dblTPTAmt As Decimal = dblTPTRate * dblQty
            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)
                If rbtnTaxCalAutomatic.IsChecked Then
                    Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value)
                    Dim strTaxType As String = clsDBFuncationality.getSingleValue(" select Type from TSPL_TAX_MASTER where Tax_Code = '" + strTaxCode + "'")
                    If clsCommon.myLen(strTaxCode) > 0 Then
                        Dim dblTaxRate As Decimal = 0
                        If clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") = CompairStringResult.Equal Then
                            If chkisTCS.Checked = True Then
                                dblTaxRate = clsCommon.myCDecimal(txtTCSTaxRate.Value)
                            Else
                                dblTaxRate = 0
                            End If
                        Else
                            Dim FormType As String = clsDBFuncationality.getSingleValue("select CUSTOMER_FORM_TYPE from TSPL_CUSTOMER_MASTER where  cust_code = '" + txtVendorNo.Value + "' ")
                            If chkExcludeKKFMandi.Checked AndAlso clsCommon.CompairString(FormType, "VSP") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(strTaxType, "K") = CompairStringResult.Equal OrElse clsCommon.CompairString(strTaxType, "M") = CompairStringResult.Equal) Then
                                dblTaxRate = 0
                            Else
                                dblTaxRate = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                            End If

                        End If
                        Dim IsSurTax As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value)
                        Dim strSurTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value)
                        Dim IsTaxable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                        Dim IsExcisable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value)
                        Dim IsTaxonBaseAmount As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(colIsTaxOnBaseAmount + clsCommon.myCstr(ii)).Value)
                        Dim dblBaseAmt As Decimal = 0
                        Dim dblTaxAmt As Decimal = 0
                        If IsSurTax Then
                            Dim dblSurTaxAmt As Decimal = GetCurrentRowSurTaxAmt(IntRowNo, ii, strSurTaxCode)
                            dblBaseAmt = dblSurTaxAmt
                        Else
                            Dim dblOtherTaxAmt As Decimal = 0
                            ''richa 25   Sep 2020 changes according to tax
                            If Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") <> CompairStringResult.Equal Then
                                dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)
                            ElseIf Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") = CompairStringResult.Equal Then
                                dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth1)
                            End If
                            If strExcise = True AndAlso intMRPwithabatement = 1 AndAlso IsExcisable = True Then
                                dblBaseAmt = (dblAbatementAmt + dblOtherTaxAmt)
                            Else
                                dblBaseAmt = (dblAmtAfterDis + dblOtherTaxAmt)
                            End If
                        End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Math.Round(dblBaseAmt, 6)
                        dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Math.Round(dblTaxAmt, 6)
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Math.Round(dblTaxAmt, 2)
                        If (IsTaxable AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper())) AndAlso (clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "CGST") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "SGST") <> CompairStringResult.Equal) Then
                            arrTaxableAuth.Add(strTaxCode.ToUpper())
                        End If
                        If (IsTaxable AndAlso Not arrTaxableAuth1.Contains(strTaxCode.ToUpper())) Then
                            arrTaxableAuth1.Add(strTaxCode.ToUpper())
                        End If
                    Else
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxRate" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value = Nothing
                    End If
                ElseIf rbtnTaxCalManual.IsChecked Then
                    If gv2.Rows.Count >= ii Then
                        Dim dblTaxAmt As Decimal = clsCommon.myCDecimal(gv2.Rows(ii - 1).Cells(colTTaxAmt).Value)
                        Dim dblCurrRowAmt As Decimal = clsCommon.myCDecimal(gv1.Rows(clsCommon.myCDecimal(IntRowNo)).Cells(colAmt).Value)
                        Dim dblTotAmt As Decimal = 0
                        For jj As Integer = 0 To gv1.Rows.Count - 1
                            dblTotAmt += clsCommon.myCDecimal(gv1.Rows(jj).Cells(colAmt).Value)
                        Next
                        Dim dblCurrCalTax As Decimal = 0
                        If dblTotAmt <> 0 Then
                            dblCurrCalTax = Math.Round(clsCommon.myCDecimal(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 6, MidpointRounding.ToEven)
                        End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = dblCurrCalTax
                    End If
                End If
            Next
            If dblQty > 0 Then
                Dim dblNetPrice As Decimal = dblAmtAfterDis / dblQty
                gv1.Rows(IntRowNo).Cells(colActualCost).Value = dblNetPrice
            End If
            Dim dblTotTaxAmt As Decimal = GetCurrentRowTotalTaxAmt(IntRowNo)
            Dim dblAmtAfterTax As Decimal = dblAmtAfterDis + dblTotTaxAmt
            gv1.Rows(IntRowNo).Cells(colDisAmt).Value = Math.Round(dblDisAmt, 2)
            gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = Math.Round(dblAmtAfterDis, 6)
            gv1.Rows(IntRowNo).Cells(colTotTaxAmt).Value = Math.Round(dblTotTaxAmt, 6)
            'gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = Math.Round(dblAmtAfterTax, 6)
            gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = Math.Round(dblAmtAfterTax, 2)
            gv1.Rows(IntRowNo).Cells(colAbatementAmount).Value = Math.Round(dblAbatementAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalMRP).Value = Math.Round(dblMRPAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalBasicAmount).Value = Math.Round(dblBasicAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotItemWt).Value = Math.Round(dblConvF * dblItemWeight * dblQty, 2)
            gv1.Rows(IntRowNo).Cells(colTotalCustDiscount).Value = Math.Round(dblTotCustDisc, 2)
            gv1.Rows(IntRowNo).Cells(colRate).Value = dblRate
            gv1.Rows(IntRowNo).Cells(colDisUnitAmt).Value = Math.Round(dblDisUnitAmt, 2)
            gv1.Rows(IntRowNo).Cells(colHeadDisPerAmt).Value = Math.Round(dblHeadPerDisAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalDiscountAmount).Value = Math.Round(dblTotDiscAmt, 2)
            gv1.Rows(IntRowNo).Cells(colOrgUnit).Value = strOrgUnit
            gv1.Rows(IntRowNo).Cells(colMRP).Value = Math.Round(dblMRP, 2)
            gv1.Rows(IntRowNo).Cells(ColTPTAmt).Value = Math.Round(dblTPTAmt, 2)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub RadMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem1.Click
        If clsCommon.myLen(ReportID) > 0 Then
            gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = ReportID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv1.SaveLayout(obj.GridLayout)
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            obj.GridColumns = gv1.ColumnCount
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow("Layout saved successfully", "Information")
            End If
            ''stuti regarding memory leakage
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub
    Private Sub gvAC_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gvAC.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column Is gvAC.Columns(colACAmount) Then
                        UpdateAllTotals()
                    ElseIf e.Column Is gvAC.Columns(colACCode) Then
                        Dim obj As clsAdditionalCharge = clsAdditionalCharge.GetFinder(clsCommon.myCstr(gvAC.CurrentRow.Cells(colACCode).Value), False)
                        If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Code) > 0 Then
                            gvAC.CurrentRow.Cells(colACCode).Value = obj.Code
                            gvAC.CurrentRow.Cells(colACName).Value = obj.desc
                        Else
                            gvAC.CurrentRow.Cells(colACCode).Value = ""
                            gvAC.CurrentRow.Cells(colACName).Value = ""
                            gvAC.CurrentRow.Cells(colACAmount).Value = 0
                        End If
                    End If
                End If
                setGridFocusAC()
                isCellValueChangedOpen = False
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub setGridFocusAC()
        Try
            Dim intCurrRow As Integer = gvAC.CurrentRow.Index
            If intCurrRow = gvAC.Rows.Count - 1 AndAlso gvAC.Rows.Count <= 10 Then
                gvAC.Rows.AddNew()
                gvAC.CurrentRow = gvAC.Rows(intCurrRow)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub RadMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem4.Click
        clsGridLayout.DeleteData(ReportID, objCommonVar.CurrentUserCode)
    End Sub
    Private Sub rbtnTaxCalAutomatic_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtnTaxCalAutomatic.ToggleStateChanged, rbtnTaxCalManual.ToggleStateChanged
        If Not isInsideLoadData Then
            If rbtnTaxCalAutomatic.IsChecked Then
                SetTaxDetails()
            ElseIf rbtnTaxCalManual.IsChecked Then
                For intRowNo As Integer = 0 To gv2.Rows.Count - 1
                    gv2.Rows(intRowNo).Cells(colTTaxRate).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTTaxAmt).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTBaseAmt).Value = Nothing
                Next
                For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                    For ii As Integer = 1 To 10
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                    Next
                Next
            End If
        End If
    End Sub
    Private Sub gv2_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv2.CellFormatting
        Try
            If e.Column.Index >= 0 Then
                If (e.Column Is gv2.Columns(colTTaxAmt)) Then
                    gv2.CurrentRow.Cells(colTTaxAmt).ReadOnly = rbtnTaxCalAutomatic.IsChecked
                End If
                Dim cell As GridDataCellElement = TryCast(e.CellElement, GridDataCellElement)
                cell.GradientStyle = GradientStyles.Solid
                cell.BackColor = Color.FromArgb(243, 181, 51)
            End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        End Try
    End Sub
    Private Sub gv2_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv2.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedTaxOpen Then
                    isCellValueChangedTaxOpen = True
                    If (e.Column Is (gv2.Columns(colTTaxAmt)) AndAlso rbtnTaxCalManual.IsChecked) Then
                        For ii As Integer = 0 To gv1.Rows.Count - 1
                            UpdateCurrentRow(ii)
                        Next
                        UpdateAllTotals()
                    End If
                    isCellValueChangedTaxOpen = False
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles gv1.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.F7 Then
            If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
                gv1.CurrentRow.Cells(colIsMannualAmt).Value = IIf(clsCommon.myCDecimal(gv1.CurrentRow.Cells(colIsMannualAmt).Value) = 1, 0, 1)
            End If
            If clsCommon.myCDecimal(gv1.CurrentRow.Cells(colIsMannualAmt).Value) = 0 Then
                UpdateCurrentRow(gv1.CurrentRow.Index)
            End If
        ElseIf e.KeyCode = Keys.F4 Then
            OpenSerialItem()
        End If
        If e.KeyCode = Keys.F5 Then
            '======update by preeti gupta 17/10/2018
            If RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                OpenBatchItem(True)
            Else
                OpenBatchItemIfFIFIOSettingON()
            End If
        End If
    End Sub
    '============created by preeti gupta===============
    Public Sub OpenBatchItemIfFIFIOSettingON()
        Dim arr As List(Of clsBatchInventory) = Nothing
        Dim strBatchunion As String = ""
        If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
            arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
        End If
        If Not arr Is Nothing Then
            If arr.Count > 0 Then
                For Each obj As clsBatchInventory In arr
                    strBatchunion += " Batch No - " & clsCommon.myCstr(obj.Batch_No) & "         Qty - " & clsCommon.myCstr(obj.Qty) + Environment.NewLine
                Next
                clsCommon.MyMessageBoxShow(strBatchunion, Me.Text)
            End If
        End If
    End Sub
    'Sub OpenBatchItem()
    '    Try
    '        If (clsCommon.myLen(clsCommon.myCDecimal(gv1.CurrentRow.Cells(colICode).Value)) > 0) Then
    '            Dim qry As String = " select Is_Batch_Item from TSPL_ITEM_MASTER where Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "'"
    '            Dim isBatchAvtive As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
    '            Dim isSubLocation As Boolean = False
    '            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsSubLocationWise from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "'", Nothing)), "Y") = CompairStringResult.Equal Then
    '                isSubLocation = True
    '            Else
    '                isSubLocation = False
    '            End If
    '            If isBatchAvtive = "1" Then
    '                Dim frm As frmBatchItemOut = New frmBatchItemOut()
    '                frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
    '                frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
    '                If isSubLocation Then
    '                    frm.strLocationCode = txtSubLocation.Value
    '                Else
    '                    frm.strLocationCode = txtBillToLocation.Value
    '                End If

    '                frm.strCurrDocNo = txtDocNo.Value
    '                frm.strCurrDocType = "MCC-MSALE"
    '                frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
    '                frm.dblMRP = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value)
    '                frm.dblqty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
    '                frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
    '                If RunBatchFifowise = 0 Then
    '                    frm.ShowDialog()
    '                    If Not frm.isCencelButtonClicked Then
    '                        gv1.CurrentRow.Cells(colICode).Tag = frm.arr
    '                    End If
    '                Else
    '                    frm.OpenSerialList(0, "")
    '                    gv1.CurrentRow.Cells(colICode).Tag = frm.arr
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        isCellValueChangedOpen = False
    '        common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub

    Sub OpenBatchItem(ByVal isFromF5 As Boolean)
        Dim idx As Integer = gv1.CurrentRow.Index
        Dim blnBatchqty As Boolean = False
        Dim isNewDocumentorExistingdoc As Boolean = True
        Dim isSubLocation As Boolean = False
        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsSubLocationWise from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "'", Nothing)), "Y") = CompairStringResult.Equal Then
            isSubLocation = True
        Else
            isSubLocation = False
        End If
        If clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
            If RunBatchFifowisewithmodifyfunctionality = True Then
                If Not isFromF5 Then

                    For ii As Integer = 0 To gv1.Rows.Count - 1
                        If gv1.Rows(ii).Cells(colIsBatchItem).Value = True Then
                            Dim strBatchunion As String = ""
                            'If RunBatchFifowise = 1 Then
                            If ii > 0 Then
                                Dim strICodeOuter As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                                For jj As Integer = 0 To ii - 1
                                    Dim strICodeInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                                    If clsCommon.CompairString(strICodeOuter, strICodeInner) = CompairStringResult.Equal Then
                                        Dim arr As List(Of clsBatchInventory) = Nothing
                                        arr = TryCast(gv1.Rows(jj).Cells(colICode).Tag, List(Of clsBatchInventory))
                                        If arr IsNot Nothing Then
                                            For Each obj As clsBatchInventory In arr
                                                Dim dblqty As Double = obj.Qty
                                                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value), clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)) <> CompairStringResult.Equal Then
                                                    dblqty = GetConvQuantity(strICodeInner, clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value), obj.Qty)
                                                End If
                                                strBatchunion += " union all select '" & clsCommon.myCstr(obj.Batch_No) & "' as Batch_No, " &
                                                        "'" & clsCommon.myCstr(obj.Manual_BatchNo) & "' as Manual_BatchNo,'O' as In_Out_Type, " &
                                                        "'" & clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value) & "' as OrgUOM," & dblqty & " as OrgQty,0 as OrgMRP, " &
                                                        "'" & clsCommon.GetPrintDate(obj.Expiry_Date, "dd/MMM/yyyy") & "' as Expiry_Date, " &
                                                        "'" & clsCommon.GetPrintDate(obj.Manufacture_Date, "dd/MMM/yyyy") & "' as Manufacture_Date, " &
                                                        "" & dblqty & " as Qty, 0 as MRP "
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                            gv1.CurrentRow = gv1.Rows(ii)
                            Dim frm As frmBatchItemOut = New frmBatchItemOut()
                            frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                            If isSubLocation Then
                                frm.strLocationCode = txtSubLocation.Value
                            Else
                                frm.strLocationCode = txtBillToLocation.Value
                            End If
                            frm.strCurrDocNo = txtDocNo.Value
                            frm.strCurrDocType = "MCC-MSALE"
                            '"PS-SH"
                            frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            frm.dblqty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
                            frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                            Dim isMilkItem As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_Milk_Pouch  from TSPL_ITEM_MASTER where item_code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "'")) = 0, False, True)
                            If frm.OpenSerialList(0, "", strBatchunion, False, IIf(isMilkItem = True, True, False)) Then
                                gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                                blnBatchqty = True
                            Else
                                Dim batchQty As Double = 0
                                For Each obj As clsBatchInventory In frm.arr
                                    batchQty += obj.Qty
                                Next
                                clsCommon.MyMessageBoxShow(Me, "Please increase stock Item Code - " & frm.strItemCode & " , Entered Qty - " & clsCommon.myCstr(frm.dblqty) & " Batch Qty - " & clsCommon.myCstr(batchQty), Me.Text)
                                blnBatchqty = False
                                Exit Sub
                            End If
                            'End If
                        End If
                    Next
                Else
                    If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsBatchItem).Value) Then
                        Dim frm As frmBatchItemOut = New frmBatchItemOut()
                        frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                        frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                        If isSubLocation Then
                            frm.strLocationCode = txtSubLocation.Value
                        Else
                            frm.strLocationCode = txtBillToLocation.Value
                        End If
                        frm.strCurrDocNo = txtDocNo.Value
                        frm.strCurrDocType = "MCC-MSALE"
                        '"PS-SH"
                        frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                        'frm.dblMRP = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value) comment by balwinder On UDL On 02/12/2016
                        frm.dblqty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
                        If checkstockmrpwise Then
                            frm.dblMRP = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value)
                        End If
                        frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                        frm.ShowDialog()
                        If Not frm.isCencelButtonClicked Then
                            gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                        End If
                    End If
                End If
            ElseIf RunBatchFifowise = 0 Then
                ' w/o fifo start here
                If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsBatchItem).Value) Then
                    Dim frm As frmBatchItemOut = New frmBatchItemOut()
                    frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    If isSubLocation Then
                        frm.strLocationCode = txtSubLocation.Value
                    Else
                        frm.strLocationCode = txtBillToLocation.Value
                    End If
                    frm.strCurrDocNo = txtDocNo.Value
                    frm.strCurrDocType = "MCC-MSALE"
                    frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                    frm.dblqty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
                    If checkstockmrpwise Then
                        frm.dblMRP = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value)
                    End If
                    frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    frm.ShowDialog()
                    If Not frm.isCencelButtonClicked Then
                        gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                    End If
                End If
                ' w/o fifo ends here
            Else
                ' fifo start
                For ii As Integer = 0 To gv1.Rows.Count - 1
                    If gv1.Rows(ii).Cells(colIsBatchItem).Value = True Then
                        Dim strBatchunion As String = ""
                        If RunBatchFifowise = 1 Then
                            If ii > 0 Then
                                Dim strICodeOuter As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                                For jj As Integer = 0 To ii - 1
                                    Dim strICodeInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                                    If clsCommon.CompairString(strICodeOuter, strICodeInner) = CompairStringResult.Equal Then
                                        Dim arr As List(Of clsBatchInventory) = Nothing
                                        arr = TryCast(gv1.Rows(jj).Cells(colICode).Tag, List(Of clsBatchInventory))
                                        For Each obj As clsBatchInventory In arr
                                            Dim dblqty As Double = obj.Qty
                                            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value), clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)) <> CompairStringResult.Equal Then
                                                dblqty = GetConvQuantity(strICodeInner, clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value), obj.Qty)
                                            End If
                                            strBatchunion += " union all select '" & clsCommon.myCstr(obj.Batch_No) & "' as Batch_No, " &
                                                "'" & clsCommon.myCstr(obj.Manual_BatchNo) & "' as Manual_BatchNo,'O' as In_Out_Type, " &
                                                "'" & clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value) & "' as OrgUOM," & dblqty & " as OrgQty,0 as OrgMRP, " &
                                                "'" & clsCommon.GetPrintDate(obj.Expiry_Date, "dd/MMM/yyyy") & "' as Expiry_Date, " &
                                                "'" & clsCommon.GetPrintDate(obj.Manufacture_Date, "dd/MMM/yyyy") & "' as Manufacture_Date, " &
                                                "" & dblqty & " as Qty, 0 as MRP "
                                        Next
                                    End If
                                Next
                            End If
                            gv1.CurrentRow = gv1.Rows(ii)
                            Dim frm As frmBatchItemOut = New frmBatchItemOut()
                            frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                            If isSubLocation Then
                                frm.strLocationCode = txtSubLocation.Value
                            Else
                                frm.strLocationCode = txtBillToLocation.Value
                            End If
                            frm.strCurrDocNo = txtDocNo.Value
                            frm.strCurrDocType = "MCC-MSALE"
                            '"PS-SH"
                            frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            frm.dblqty = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colQty).Value)
                            frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                            If frm.OpenSerialList(0, "", strBatchunion) Then
                                gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                                blnBatchqty = True
                            Else
                                Dim batchQty As Double = 0
                                For Each obj As clsBatchInventory In frm.arr
                                    batchQty += obj.Qty
                                Next
                                clsCommon.MyMessageBoxShow(Me, "Please increase stock Item Code - " & frm.strItemCode & " , Entered Qty - " & clsCommon.myCstr(frm.dblqty) & " Batch Qty - " & clsCommon.myCstr(batchQty), Me.Text)
                                blnBatchqty = False
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                ' fifo ends here
            End If
        End If
        Try
            gv1.CurrentRow = gv1.Rows(idx)
        Catch ex As Exception
        End Try

    End Sub

    Public Shared Function GetConvQuantity(ByVal strItem As String, ByVal strCurrentUnit As String, ByVal strConvertedUnit As String, ByVal dblQty As Double) As Double
        Dim dblCurrentConvF As Double = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strCurrentUnit & "'"))
        Dim dblConvQty As Double = 0
        If clsCommon.myLen(strConvertedUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strConvertedUnit & "'"))
            If dblCurrentConvF > 0 Then
                dblConvQty = Math.Round(Math.Round((dblOrgConvF / dblCurrentConvF), 2) * dblQty, 2)
            End If
        End If
        Return dblConvQty
    End Function
    Public Shared Function GetConvQuantity(ByVal strItem As String, ByVal strCurrentUnit As String, ByVal strConvertedUnit As String, ByVal dblQty As Double, ByVal trans As SqlTransaction) As Double
        Dim dblCurrentConvF As Double = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strCurrentUnit & "'", trans))
        Dim dblConvQty As Double = 0
        If clsCommon.myLen(strConvertedUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strConvertedUnit & "'", trans))
            If dblCurrentConvF > 0 Then
                dblConvQty = Math.Round(Math.Round((dblOrgConvF / dblCurrentConvF), 2) * dblQty, 2)
            End If
        End If
        Return dblConvQty
    End Function
    Private Sub gv1_RowFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.RowFormattingEventArgs) Handles gv1.RowFormatting
        Try
            If clsCommon.CompairString(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myCDecimal(e.RowElement.RowInfo.Cells(colIsMannualAmt).Value) > 0 Then
                e.RowElement.ForeColor = Color.Blue
            Else
                e.RowElement.ForeColor = Color.Black
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv1_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gv1.CurrentColumnChanged
        If gv1.RowCount > 0 Then
            'If Not clsCommon.myLen(txtReqNo.Value) > 0 Then
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCDecimal(intCurrRow + 1)
            If intCurrRow = gv1.Rows.Count - 1 Then
                gv1.Rows.AddNew()
                gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                gv1.CurrentRow = gv1.Rows(intCurrRow)
            End If
        End If
        'End If
    End Sub
    Public Shared Function GetInvoiceType() As DataTable
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = ""
        dr("Name") = "Select"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "R"
        dr("Name") = "Retail"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "T"
        dr("Name") = "Tax"
        dt.Rows.Add(dr)
        dr = dt.NewRow()
        dr("Code") = "N"
        dr("Name") = "Non Taxable"
        dt.Rows.Add(dr)
        Return dt
    End Function
    Sub LoadDeductionType()
        ' Dim qry As String = "select '' as Code, '<--Select-->' as Name union all select Code, Description as Name from TSPL_DEDUCTION_MASTER"
        Dim qry As String = "select '' as Code, '<--Select-->' as Name union all 
                             Select  Distinct Deduction,max(Description) as Name  from TSPL_ITEM_MASTER 
                             left outer join TSPL_DEDUCTION_MASTER On TSPL_DEDUCTION_MASTER.Code=TSPL_ITEM_MASTER.Deduction
                             where 2=2 and Deduction is not null group by Deduction "
        cboDeductionType.DataSource = clsDBFuncationality.GetDataTable(qry)
        cboDeductionType.ValueMember = "Code"
        cboDeductionType.DisplayMember = "Name"
    End Sub
    Sub LoadInvoiceType()
        ddlInvoiceType.DataSource = GetInvoiceType()
        ddlInvoiceType.ValueMember = "Code"
        ddlInvoiceType.DisplayMember = "Name"
    End Sub
    Private Sub chkCreateAutoInvoice_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkCreateAutoInvoice.ToggleStateChanged
        'chkCreateAutoReceipt.Visible = chkCreateAutoInvoice.Checked
        If AllowChangeInvoiceType Then
            lblInvoiceType.Visible = True
            ddlInvoiceType.Visible = True
        Else
            lblInvoiceType.Visible = False
            ddlInvoiceType.Visible = False
        End If
        If chkCreateAutoInvoice.Checked Then
            'btnPrintInvoice.Visible = True
            lblInvoiceNo.Visible = True
            txtInvoiceNo.Visible = True
            '-----------------richa 27/06/2014 Ticket No .BM00000002982------------
            ddlInvoiceType.Visible = True
            lblInvoiceType.Visible = True
            '---------------------------------------------------------------------
        Else
            'btnPrintInvoice.Visible = True
            lblInvoiceNo.Visible = False
            txtInvoiceNo.Visible = False
            '-----------------richa 27/06/2014 Ticket No .BM00000002982------------
            ddlInvoiceType.Visible = False
            lblInvoiceType.Visible = False
            '---------------------------------------------------------------------
        End If
    End Sub
    Private Sub txtCurrencyCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtCurrencyCode._MYValidating
        Dim qry As String = "select CURRENCY_CODE AS Code, CURRENCY_NAME as Name from TSPL_CURRENCY_MASTER"
        txtCurrencyCode.Value = clsCommon.ShowSelectForm("CURRENCY_MASTER", qry, "Code", "", txtCurrencyCode.Value, "CURRENCY_CODE", isButtonClicked)
        ShowCurrencyDetail()
    End Sub
    Sub setBalance()
        UcItemBalance1.ItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        UcItemBalance1.ItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
        UcItemBalance1.ItemMRP = clsCommon.myCDecimal(gv1.CurrentRow.Cells(colMRP).Value)
        If AllowPlandDeptMCCLocation Then
            If clsCommon.myLen(clsCommon.myCstr(txtBillToLocation.Value)) > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                    UcItemBalance1.LocationCode = txtSubLocation.Value
                    UcItemBalance1.LocationName = lblSubLocation.Text
                Else
                    UcItemBalance1.LocationCode = txtBillToLocation.Value
                    UcItemBalance1.LocationName = lblBillToLocation.Text
                End If
            End If
        Else
            UcItemBalance1.LocationCode = txtBillToLocation.Value
            UcItemBalance1.LocationName = lblBillToLocation.Text
        End If
        UcItemBalance1.UOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
        UcItemBalance1.TransNo = txtDocNo.Value
        UcItemBalance1.TransDate = txtDate.Value
        UcItemBalance1.ShowSOQty = True
        UcItemBalance1.CommitedQty = True
        UcItemBalance1.CommitedQtyLbl = True
        UcItemBalance1.RefreshData()
    End Sub
    Private Sub gv1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles gv1.Click
        If gv1.CurrentRow IsNot Nothing Then
            setBalance()
        End If
    End Sub
    Private Sub gv1_CurrentRowChanged(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.CurrentRowChangedEventArgs) Handles gv1.CurrentRowChanged
        If gv1.CurrentRow IsNot Nothing AndAlso Not e.CurrentRow.Index < 0 Then
            setBalance()
        End If
    End Sub
    Private Sub RadGroupBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadGroupBox2.Click
    End Sub
    Public Shared Function GetCost(ByVal CostMethod As EnumCostingMethod, ByVal strICode As String, ByVal strLocation As String, ByVal dblqty As Decimal, ByVal dtDocumentDate As DateTime, ByVal dtPostingDate As DateTime, ByVal isApplyCostOnPostDate As Boolean, ByVal trans As SqlTransaction) As Decimal
        Dim dblRetCost As Decimal = 0
        If Not CostMethod = EnumCostingMethod.NA AndAlso dblqty > 0 Then
            Dim strSymbolCost As String = " >= "
            If CostMethod = EnumCostingMethod.LIFO Then
                strSymbolCost = " <= "
            End If
            Dim strDateColumn As String = " Punching_Date "
            Dim strDateForCheck As String = clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtDocumentDate), "dd/MMM/yyyy hh:mm tt")
            If isApplyCostOnPostDate Then
                strDateColumn = " Posting_Date "
                strDateForCheck = clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtPostingDate), "dd/MMM/yyyy hh:mm tt")
            End If
            Dim qry As String
            If CostMethod = EnumCostingMethod.Averege Then
                qry = "select case when Qty=0 then 0 else abs(Amt/Qty)*" + clsCommon.myCstr(dblqty) + "  end as AvgCost from( select  sum(Amt * RI) as Amt,sum(Qty * RI) as Qty from(" + Environment.NewLine
                qry += " select Stock_Qty as Qty,( Avg_Cost) as Amt,case when InOut='O' then -1 else 1 end as RI  from TSPL_INVENTORY_MOVEMENT where Item_Code='" + strICode + "' and Location_Code='" + strLocation + "' and  " + strDateColumn + " <= '" + strDateForCheck + "' " + Environment.NewLine
                qry += " )xxx )xxxx" + Environment.NewLine
                dblRetCost = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry, trans))
            Else
                qry = ";WITH cteStockSum AS ( " + Environment.NewLine
                qry += " SELECT   Item_Code ,SUM(Stock_Qty * CASE WHEN  InOut = 'O' THEN -1 ELSE 1 END) AS TotalStock FROM  TSPL_INVENTORY_MOVEMENT where Item_Code='" + strICode + "' and Location_Code='" + strLocation + "' and " + strDateColumn + " <= '" + strDateForCheck + "'  GROUP BY Item_Code)," + Environment.NewLine
                qry += " cteReverseInSum AS (" + Environment.NewLine
                qry += " SELECT  s.Item_Code ,s." + strDateColumn + " as TranDate ,(SELECT SUM(i.Stock_Qty) FROM TSPL_INVENTORY_MOVEMENT AS i  WHERE i.Item_Code = s.Item_Code AND i.InOut IN ( 'I' ) and i." + strDateColumn + " <= '" + strDateForCheck + "' AND i." + strDateColumn + " " + strSymbolCost + " s." + strDateColumn + " --for FIFO  >= " + Environment.NewLine
                qry += " ) AS RollingStock ,s.Stock_Qty AS ThisStock FROM TSPL_INVENTORY_MOVEMENT AS s WHERE  s.Item_Code='" + strICode + "' and s.Location_Code='" + strLocation + "' and s." + strDateColumn + " <= '" + strDateForCheck + "'  and s.InOut IN ( 'I' ))," + Environment.NewLine
                qry += " cteWithLastTranDate  AS ( " + Environment.NewLine
                qry += " SELECT   w.Item_Code ,w.TotalStock ,LastPartialStock. TranDate ,LastPartialStock.StockToUse ,LastPartialStock.RunningTotal ,w.TotalStock -LastPartialStock.RunningTotal+ LastPartialStock.StockToUse AS UseThisStock FROM cteStockSum AS w" + Environment.NewLine
                qry += " CROSS APPLY ( SELECT TOP ( 1 )z. TranDate ,z.ThisStock AS StockToUse ,z.RollingStock AS RunningTotal FROM  cteReverseInSum AS z WHERE z.Item_Code = w.Item_Code AND z.RollingStock >= w.TotalStock ORDER BY  z.TranDate " + IIf(CostMethod = EnumCostingMethod.FIFO, "DESC", "") + " --for FIFO DESC" + Environment.NewLine
                qry += " ) AS LastPartialStock" + Environment.NewLine
                qry += " )" + Environment.NewLine
                qry += " select *  from (" + Environment.NewLine
                qry += " SELECT  y.Item_Code ,y.TotalStock AS CurrentItems ,e.Basic_Cost,e." + strDateColumn + " as TranDate,(CASE WHEN e." + strDateColumn + " = y.TranDate THEN y.UseThisStock" + Environment.NewLine
                qry += " ELSE e.Stock_Qty END * Price.Basic_Cost) AS CurrentValue,(CASE WHEN e. " + strDateColumn + "  = y.TranDate THEN y.UseThisStock  ELSE e.Stock_Qty END  ) as BalanceQty FROM cteWithLastTranDate AS y INNER JOIN TSPL_INVENTORY_MOVEMENT AS e ON e.Item_Code = y.Item_Code and e." + strDateColumn + " <= '" + strDateForCheck + "' AND e." + strDateColumn + " " + strSymbolCost + " y.TranDate -- for Fifo >=" + Environment.NewLine
                qry += " AND e.InOut IN ('I')" + Environment.NewLine
                qry += " CROSS APPLY ( SELECT TOP ( 1 ) case when Stock_Qty =0 then 0 else  (p.Basic_Cost*p.Qty)/p.Stock_Qty end as Basic_Cost FROM TSPL_INVENTORY_MOVEMENT AS p  WHERE p.Item_Code = e.Item_Code " + Environment.NewLine
                qry += " AND p." + strDateColumn + " <= e." + strDateColumn + "  " + Environment.NewLine
                qry += " AND p.InOut = 'I'  ORDER BY p." + strDateColumn + " DESC ) AS Price" + Environment.NewLine
                qry += ")xxx   " + IIf(CostMethod = EnumCostingMethod.FIFO, " order by TranDate ", IIf(CostMethod = EnumCostingMethod.LIFO, "order by TranDate DESC", "")) + " --For Fifo not Desc order" + Environment.NewLine
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim dblbalanceQty As Decimal = dblqty
                    For Each dr As DataRow In dt.Rows
                        Dim dblCurrQty As Decimal = clsCommon.myCDecimal(dr("BalanceQty"))
                        If dblbalanceQty >= dblCurrQty Then
                            dblRetCost += clsCommon.myCDecimal(dr("CurrentValue"))
                        Else
                            dblRetCost += (clsCommon.myCDecimal(dr("CurrentValue")) * dblbalanceQty) / dblCurrQty
                        End If
                        dblbalanceQty -= dblCurrQty
                        If dblbalanceQty <= 0 Then
                            dblRetCost = dblRetCost
                            Exit For
                        End If
                    Next
                    If dblbalanceQty > 0 Then
                        Throw New Exception("Quantity Not available for " + strICode)
                    End If
                End If
            End If
        End If
        'Throw New Exception("cost " + clsCommon.myCstr(dblRetCost))
        Return dblRetCost
    End Function
    'Private Sub txtBarCode_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
    '    If clsCommon.myLen(txtBarCode.Text) > 0 Then
    '        Dim obj As clsBarCodeGenerator = clsBarCodeGenerator.GetData(txtBarCode.Text)
    '        If obj Is Nothing Then
    '            clsCommon.MyMessageBoxShow("Not a Valid Barcode", Me.Text)
    '            txtBarCode.Text = ""
    '            Exit Sub
    '        End If
    '        Dim isFound As Boolean = False
    '        Dim CurrentRow As Integer = 1
    '        For ii As Integer = 0 To gv1.Rows.Count - 1
    '            If clsCommon.CompairString(txtBarCode.Text, clsCommon.myCstr(gv1.Rows(ii).Cells(colBarCode).Value)) = CompairStringResult.Equal Then
    '                gv1.Rows(ii).Cells(colQty).Value = clsCommon.myCDecimal(gv1.Rows(ii).Cells(colQty).Value) + 1
    '                CurrentRow = ii
    '                isFound = True
    '                Exit For
    '            End If
    '        Next
    '        If Not isFound Then
    '            gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
    '            gv1.Rows(gv1.Rows.Count - 1).Cells(colBarCode).Value = obj.Bar_Code
    '            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
    '            If blnBackCalculation = True Then
    '                OpenICodeList(False)
    '            Else
    '                OpenICodeListCurrentCalaculation(False)
    '            End If
    '            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = 1
    '            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Selling_Price
    '            gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.Item_MRP
    '            CurrentRow = gv1.Rows.Count - 1
    '            For ii As Integer = 1 To gv1.Rows.Count
    '                gv1.Rows(ii - 1).Cells(colLineNo).Value = ii
    '            Next
    '            gv1.Rows.AddNew()
    '        End If
    '        UpdateCurrentRow(CurrentRow)
    '        UpdateAllTotals()
    '        txtBarCode.Text = ""
    '        txtBarCode.Focus()
    '    End If
    'End Sub
    Private Sub chkDiscountOnAmt_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkDiscountOnAmt.ToggleStateChanged, chkDiscountOnRate.ToggleStateChanged
        If chkDiscountOnAmt.IsChecked Then
            txtDiscAmt.Enabled = True
            txtDiscPer.Enabled = False
            txtDiscPer.Text = 0
            For Each gro As GridViewRowInfo In gv1.Rows
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCDecimal(gro.Cells(ColFOC).Value) = 0 Then
                    gro.Cells(colHeaDDisPer).Value = 0
                Else
                End If
            Next
            lblInvoiceDiscAmt.Text = 0
        Else
            txtDiscAmt.Enabled = False
            txtDiscPer.Enabled = True
            txtDiscAmt.Text = 0
            For Each gro As GridViewRowInfo In gv1.Rows
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCDecimal(gro.Cells(ColFOC).Value) = 0 Then
                    gro.Cells(colHeadDiscamt).Value = 0
                Else
                End If
            Next
            lblInvoiceDiscAmt.Text = 0
        End If
    End Sub
    Private Sub txtDiscAmt_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscAmt.Leave
        Try
            CalculateDiscountAmount()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDiscAmt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscAmt.Click
    End Sub
    Private Sub txtDiscPer_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscPer.Leave
        CalculateDiscountAmount()
    End Sub
    Private Sub txtDiscPer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscPer.Click
    End Sub
    'Private Sub CalculateDiscountAmount()
    '    If clsCommon.myCDecimal(txtDiscAmt.Text) > clsCommon.myCDecimal(lblAmtWithDiscount.Text) Then
    '        isCellValueChangedOpen = False
    '        Throw New Exception("Discount amount cannot be greater than Doc amount")
    '    End If
    '    Dim discountrate As Decimal = Decimal.Parse(txtDiscPer.Text)
    '    'If chkDiscountOnAmt.IsChecked AndAlso clsCommon.myCDecimal(lblAmtWithDiscount.Text) > 0 Then
    '    '    discountrate = clsCommon.myCstr(Math.Round(((txtDiscAmt.Value * 100) / clsCommon.myCDecimal(lblAmtWithDiscount.Text)), 5, MidpointRounding.ToEven))
    '    '    txtDiscPer.Value = discountrate
    '    'End If
    '    Dim dblDiscountAmtPerUnit As Decimal = 0
    '    Dim dblDiscountAmt As Decimal = 0
    '    Dim dblCustDiscountNoTax As Decimal = 0
    '    If String.IsNullOrEmpty(lblAmtWithDiscount.Text) Then
    '        lblAmtWithDiscount.Text = 0
    '    End If
    '    If chkDiscountOnAmt.IsChecked Then
    '        For Each gro As GridViewRowInfo In gv1.Rows
    '            gv1.CurrentRow = gro.Cells(colHeadDiscamt).RowInfo
    '            If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCDecimal(gro.Cells(ColFOC).Value) = 0 Then
    '                dblDiscountAmt = Math.Round((clsCommon.myCDecimal(gro.Cells(colAmt).Value) * txtDiscAmt.Value) / clsCommon.myCDecimal(lblAmtWithDiscount.Text), 2)
    '                gro.Cells(colHeadDiscamt).Value = Math.Round((dblDiscountAmt), 2)
    '            Else
    '                gro.Cells(colHeadDiscamt).Value = 0
    '            End If
    '        Next
    '    Else
    '        For Each gro As GridViewRowInfo In gv1.Rows
    '            gv1.CurrentRow = gro.Cells(colHeadDiscamt).RowInfo
    '            If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCDecimal(gro.Cells(ColFOC).Value) = 0 Then
    '                gro.Cells(colDisPer).Value = Math.Round((discountrate), 2)
    '            Else
    '            End If
    '        Next
    '    End If
    'End Sub
    Private Sub CalculateDiscountAmount()
        If clsCommon.myCDecimal(txtDiscAmt.Text) > clsCommon.myCDecimal(lblAmtWithDiscount.Text) Then
            isCellValueChangedOpen = False
            Throw New Exception("Discount amount cannot be greater than Doc amount")
        End If
        Dim discountrate As Decimal
        'Dim discountrate As Decimal = Decimal.Parse(txtDiscPer.Text)
        'If chkDiscountOnAmt.IsChecked AndAlso clsCommon.myCDecimal(lblAmtWithDiscount.Text) > 0 Then
        '    discountrate = clsCommon.myCstr(Math.Round(((txtDiscAmt.Value * 100) / clsCommon.myCDecimal(lblAmtWithDiscount.Text)), 5, MidpointRounding.ToEven))
        '    txtDiscPer.Value = discountrate
        'End If
        If clsCommon.myCDecimal(txtDiscPer.Text) > 0 Then
            discountrate = Decimal.Parse(txtDiscPer.Text)
            txtDiscAmt.Text = 0
        ElseIf clsCommon.myCDecimal(txtDiscAmt.Text) > 0 Then
            txtDiscPer.Text = 0
        End If
        Dim dblDiscountAmt As Decimal = 0
        Dim dblCustDiscountNoTax As Decimal = 0
        If String.IsNullOrEmpty(lblAmtWithDiscount.Text) Then
            lblAmtWithDiscount.Text = 0
        End If
        If chkDiscountOnAmt.IsChecked Then
            For Each gro As GridViewRowInfo In gv1.Rows
                gv1.CurrentRow = gro.Cells(colHeadDiscamt).RowInfo
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCDecimal(gro.Cells(ColFOC).Value) = 0 Then
                    'dblDiscountAmt = Math.Round((clsCommon.myCDecimal(gro.Cells(colAmt).Value) * txtDiscAmt.Value) / clsCommon.myCDecimal(lblAmtWithDiscount.Text), 2)
                    dblDiscountAmt = Math.Round(txtDiscAmt.Value * TotalItemQty, 2)
                    gro.Cells(colHeadDiscamt).Value = Math.Round((dblDiscountAmt), 2)
                Else
                    gro.Cells(colHeadDiscamt).Value = 0
                End If
            Next
        Else
            For Each gro As GridViewRowInfo In gv1.Rows
                gv1.CurrentRow = gro.Cells(colHeadDiscamt).RowInfo
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCDecimal(gro.Cells(ColFOC).Value) = 0 Then
                    gro.Cells(colHeaDDisPer).Value = Math.Round((discountrate), 2)
                Else
                End If
            Next
        End If
    End Sub
    Private Sub lblVendorName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblVendorName.Click
    End Sub
    Private Sub RadLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadLabel2.Click
    End Sub
    Private Sub txtFromLoc__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtFromLoc._MYValidating
        Dim qry As String = "select Location_Code as Code,Location_Desc as Name from TSPL_LOCATION_MASTER "
        Dim WhrCls As String = " Location_Type='Physical' and location_category='MCC' "
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            WhrCls += "  and  Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
        End If
        txtFromLoc.Value = clsCommon.ShowSelectForm("Transferfndr", qry, "Code", WhrCls, txtFromLoc.Value, "Code", isButtonClicked)
    End Sub
    Private Sub btnPrintInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrintInvoice.Click
        If clsCommon.myLen(txtInvoiceNo.Text) <= 0 Then
            myMessages.blankValue(Me, "Invoice not found to Print", Me.Text)
        Else
            'Dim objInvoice As New frmSNSaleInvoice
            'objInvoice.funPrintForMCC(txtInvoiceNo.Text)
            Dim ObjclsSNInvoiceHead As New clsSNInvoiceHead()
            ObjclsSNInvoiceHead.funPrintForMCC(MyBase.Form_ID, txtInvoiceNo.Text)
        End If
    End Sub
    Private Sub btnReverseAndUnpost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverseAndUnpost.Click
        Try
            If common.clsCommon.MyMessageBoxShow("Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                If clsMCCMaterialSale.ReverseAndUnpost(txtDocNo.Value, False) Then
                    common.clsCommon.MyMessageBoxShow("Successfully Reversed and Recreated", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnDrillDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If clsCommon.myLen(txtReqNo.Value) > 0 Then
        '    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSNSalesOrder, txtReqNo.Value)
        'Else
        '    common.clsCommon.MyMessageBoxShow("No data found")
        'End If
    End Sub
    '-------------------------Done By Monika For Mailing System--------------------------
#Region "Mail SMS Setting"
    Private Function GetPdfPath() As String
        Dim StrRptPath As String = ""
        Dim IsMandiTax As Decimal = 0
        IsMandiTax = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code='" & txtTaxGroup.Value & "' and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y')"))
        Dim IsExampted As Decimal = 0
        IsExampted = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select COUNT(Is_Tax_Exempted)  from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='" & txtTaxGroup.Value & "' AND Tax_Group_Type='S' AND Is_Tax_Exempted=1"))
        'atchqry = GetAttachQry()
        atchqry = clsMCCMaterialSale.GetAttachQry(False, txtDate.Value, txtDocNo.Value)
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)
        If dt.Rows.Count > 0 Then
            ' ==============================Ticket No  ERO/27/11/18-000421 By Prabhakar for  Customer Dashboard on Print==============================================
            Dim dtCustomerOutstanding As DataTable = Nothing
            dtCustomerOutstanding = clsCustomerMaster.getCustomerOutstandingOfAmt_Can_Crate("'" & clsCommon.myCstr(dt.Rows(0)("Customer_Code")) & "'", clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")).AddDays(-1), "dd/MMM/yyyy"), clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_Date")), "dd/MMM/yyyy"))
            '=============================================================================
            Dim frmCRV As New frmCrystalReportViewer()
            If clsERPFuncationality.GetGSTStatus(clsCommon.myCDate(dt.Rows(0)("Document_Date"))) Then
                If clsCommon.myCDecimal(dt.Rows(0)("Is_Taxable")) = 1 Then
                    If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("LocationState")), clsCommon.myCstr(dt.Rows(0)("Cust_State_Name"))) = CompairStringResult.Equal Then
                        If IsExampted > 0 Then
                            StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptcustomerOutstandingErod.rpt", dtCustomerOutstanding)
                        ElseIf IsMandiTax > 0 Then
                            StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local_WithMandi", "MCC Material Sale Local With Mandi", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                        Else
                            StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                            ''frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_Local", "MCC Material Sale Local", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt")
                        End If
                    Else
                        If IsExampted > 0 Then
                            StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                        ElseIf IsMandiTax > 0 Then
                            StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState_WithMandi", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                        Else
                            StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_InterState", "MCC Material Sale Interstate", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                        End If
                    End If
                Else
                    StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale_NonTaxable", "MCC Material Sale Non Taxable", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptCompanyAddress.rpt", "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
                End If
            Else
                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Is_CashSale")), "Y") = CompairStringResult.Equal Then
                    StrRptPath = frmCRV.funreport(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, "rptMccMaterialSaleCashMemo", "Cash Memo")
                Else
                    StrRptPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, True, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptMCCMaterialSale", "MCC Material Sale", "rptCompanyAddress.rpt")
                End If
            End If
            frmCRV = Nothing
        End If
        Return StrRptPath
    End Function
    Private Sub SendSMSandEmail(ByVal lstUsers As List(Of String), ByVal isSendForApproval As Boolean)
        Try
            Dim strContactperson As String = ""
            Dim strStartupPath As String = ""
            Dim strPdfPath As String = ""
            'sanjay
            Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + Form_ID + "'", Nothing)
            If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
                Dim objEmailH As New clsEMailHead()
                objEmailH.arrEMail = New List(Of String)()
                objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.SaleOrderNo, txtDocNo.Value)
                objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
                objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SaleOrderNo, txtDocNo.Value)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerNo, txtVendorNo.Value)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerName, lblVendorName.Text)
                'objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.ContactPerson, strContactperson)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.TotalAmount, lblTotRAmt.Text)
                objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Form_Code, MyBase.Form_ID)
                'atchqry = GetAttachQry()
                'Dim dtFinal As DataTable = clsDBFuncationality.GetDataTable(atchqry)
                strPdfPath = GetPdfPath()
                'strStartupPath = clsERPFuncationality.GetReportPathWithoutRptName(CrystalReportFolder.MilkProcurement, clsCommon.myCDate(dtFinal.Rows(0)("Document_Date")))
                'If clsERPFuncationality.exportCrystalToPDF(dtFinal, strStartupPath, strRptName, "Attachment", strStartupPath) Then
                objEmailH.Attachment_1_Path = strPdfPath
                For Each strUser As String In lstUsers
                    Dim lstReceiptents As New List(Of String)
                    Dim qry As String = ""
                    Dim emailId As String = ""
                    If isSendForApproval Then
                        strContactperson = strUser
                        qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" + strUser + "') "
                        emailId = clsDBFuncationality.getSingleValue(qry)
                    Else
                        strContactperson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
                        emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
                    End If
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.ContactPerson, strContactperson)
                    If clsCommon.myLen(emailId) > 0 Then
                        objEmailH.arrEMail.Add(emailId)
                    End If
                    'lstReceiptents.Add(emailId)
                Next
                objEmailH.SaveData(MyBase.Form_ID, objEmailH, Nothing)
                objEmailH = Nothing
                clsCommon.MyMessageBoxShow(Me, "E-Mail Send Successfully", Me.Text)
                'SMSSENDONLY(False)
                Dim obj As clsMCCMaterialSale = clsMCCMaterialSale.GetData(txtDocNo.Value, NavigatorType.Current)
                clsMCCMaterialSale.SMSSENDONLY(obj, Nothing, False)
                'End If
            Else
                clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
            End If
            'sanjay
            'Dim obj As clsEmailSMSSettingNew = clsEmailSMSSettingNew.GetData(clsUserMgtCode.frmSNShipment)
            'If obj Is Nothing Then
            '    clsCommon.MyMessageBoxShow("First do setting of email/sms", Me.Text)
            '    Return
            'End If
            'If clsCommon.myLen(obj.mailsubjct) <= 0 Then
            '    clsCommon.MyMessageBoxShow("First do setting of email/sms", Me.Text)
            '    Return
            'End If
            'Dim strSubject As String = obj.mailsubjct.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'strSubject = strSubject.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
            'Dim strbody As String = obj.mailbody.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'strbody = strbody.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
            'strbody = strbody.Replace(clsEmailSMSConstants.CustomerNo, txtVendorNo.Value)
            'strbody = strbody.Replace(clsEmailSMSConstants.CustomerName, lblVendorName.Text)
            'strbody = strbody.Replace(clsEmailSMSConstants.ContactPerson, strContactperson)
            'strbody = strbody.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
            'strbody = strbody.Replace(clsEmailSMSConstants.Form_Code, MyBase.Form_ID)
            ''------------------------code for attchament-------------------------------------
            'Dim strRptPath As String = ""
            'If obj.atchmnt = "Y" Then
            '    attachQry = GetAttachQry()
            '    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(attachQry)
            '    If dt1.Rows.Count > 0 Then
            '        SetItemWiseTax(dt1, txtDocNo.Value)
            '        Dim frmCRV As New frmCrystalReportViewer()
            '        strRptPath = frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt1, "crptShipment", "Shipment Detail")
            '        frmCRV = Nothing
            '    End If
            'End If
            ''---------------------------------------------------------------------------
            'Dim qry As String = "Select * from TSPL_APPROVAL_LEVEL_SCREEN where 1=1 and Trans_Code='" + MyBase.Form_ID + "' "
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            'Dim lstUsers As New List(Of String)
            'Dim lstReceiptents As New List(Of String)
            'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '    For Each dr As DataRow In dt.Rows
            '        qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" + dr("User_Code").ToString() + "') "
            '        Dim emailId As String = clsDBFuncationality.getSingleValue(qry)
            '        lstReceiptents.Add(emailId)
            '        Dim body As String = strbody.Replace(clsEmailSMSConstants.UserCode, dr("User_Code").ToString())
            '        clsMailViaOutlook.SendEmail(strSubject, body, lstReceiptents, Nothing, strRptPath)
            '    Next
            'End If
            'clsCommon.MyMessageBoxShow("E-Mail Send Successfully", Me.Text)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        'Try
        '    Dim client As New System.Net.WebClient()
        '    'strMes = "Dear " & strContactperson & " (" & strCustomer & ")" & "   your Bill/Invoice No " & txtDocNo.Value & "  dated  " & clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy") & "  has been dispatched with amount  " & lblTotRAmt.Text
        '    If clsCommon.myLen(obj.smsbody) <= 0 Then
        '        Return
        '    End If
        '    strMes = obj.smsbody
        '    If strMes.Contains(clsEmailSMSConstants.SaleOrderNo) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.SaleOrderDate) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.VendorNo) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.VendorNo, txtVendorNo.Value)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.VendorName) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.VendorName, lblVendorName.Text)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.TotalAmount) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
        '    End If
        '    strphone = clsDBFuncationality.getSingleValue("select Phone1 from TSPL_customer_MASTER where cust_code ='" & txtVendorNo.Value & "' ")
        '    Dim baseurl As String = "http://bulksms.mysmsmantra.com:8080/WebSMS/SMSAPI.jsp?username=tecxpert&password=1818948263&sendername=vipin&mobileno=91" + strphone + "&message=" + strMes
        '    Dim data As Stream = client.OpenRead(baseurl)
        '    Dim reader As StreamReader = New StreamReader(data)
        '    Dim s As String = reader.ReadToEnd()
        '    data.Close()
        '    reader.Close()
        '    clsCommon.MyMessageBoxShow("SMS Send Successfully", Me.Text)
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'End Try
    End Sub
    Private Sub SetMailRight()
        'Dim obj As clsCheckMailSetting = clsCheckMailSetting.CheckMailRight()
        If objCommonVar.IsMailSend Then
            btnsend.Enabled = True
        Else
            btnsend.Enabled = False
        End If
    End Sub
    Public Function GetMailPrint(ByVal strCode As String)
        atchqry = "  select '" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MM/yyyy") + "' as RunDate, TSPL_SHIP_TO_LOCATION.Ship_To_Desc as shipName, TSPL_SHIP_TO_LOCATION.add1 as ship_Add1, TSPL_SHIP_TO_LOCATION.Add2 as ship_add2 ,TSPL_SHIP_TO_LOCATION.Add3 as ship_add3  ,TSPL_SHIP_TO_LOCATION.Pin_Code,TSPL_CITY_MASTER.STATE_CODE  ,City_Name,TSPL_EMPLOYEE_MASTER.Emp_Name as SalesManName,TSPL_SD_SHIPMENT_HEAD.Inv_No, TSPL_SD_SHIPMENT_HEAD.Dept_Desc , TSPL_SD_SHIPMENT_HEAD.Remarks ,  TSPL_SD_SHIPMENT_HEAD.Terms_Code,TSPL_SD_SHIPMENT_HEAD.VehicleNo , "
        atchqry += " TSPL_SD_SHIPMENT_DETAIL .Specification as  specification,   TSPL_SD_SHIPMENT_HEAD.Document_Code as DocNo , TSPL_SD_SHIPMENT_HEAD.Description, "
        atchqry += "  convert(varchar ,TSPL_SD_SHIPMENT_HEAD .Document_Date,103)as Document_Date , TSPL_SD_SHIPMENT_HEAD.Against_Sales_Order, TSPL_SD_SHIPMENT_HEAD.Item_Type ,  TSPL_SD_SHIPMENT_HEAD.Customer_Code, "
        atchqry += " TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 as Customer_Add1,TSPL_CUSTOMER_MASTER.add2 as customer_Add2,TSPL_CUSTOMER_MASTER.Add3 as customer_Add3 ,TSPL_CUSTOMER_MASTER.State as customer_city_State ,TSPL_CUSTOMER_MASTER.PIN_Code as Customer_Pin_Code , TSPL_SD_SHIPMENT_HEAD .Terms_Code as termscode ,TSPL_SD_SHIPMENT_HEAD .Ref_No as ref_no ,"
        atchqry += " TSPL_SD_SHIPMENT_HEAD .Comments as comments ,  TSPL_SD_SHIPMENT_HEAD .Discount_Amt as dis_amt,TSPL_SD_SHIPMENT_DETAIL .Disc_Amt  as dis_amt1,"
        atchqry += " TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount  as aftrdiscount ,TSPL_SD_SHIPMENT_HEAD .Total_Amt as Total_amount,"
        atchqry += " TSPL_SD_SHIPMENT_HEAD.Discount_Base as bfrdisc_amount,  "
        atchqry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SHIPMENT_HEAD.tax1_amt,0) as txt1amt,  "
        atchqry += " tax2.Tax_Code_Desc as tax2name,isnull (TSPL_SD_SHIPMENT_HEAD.tax2_amt,0) as txt2amt,  "
        atchqry += " tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SHIPMENT_HEAD.tax3_amt,0) as txt3amt,  "
        atchqry += " tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SHIPMENT_HEAD.tax4_amt,0) as txt4amt,  "
        atchqry += " tax5.Tax_Code_Desc as tax5name,isnull (TSPL_SD_SHIPMENT_HEAD.tax5_amt,0) as txt5amt,  "
        atchqry += " tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SHIPMENT_HEAD.tax6_amt,0) as txt6amt,  "
        atchqry += " tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SHIPMENT_HEAD.tax7_amt,0) as txt7amt,  "
        atchqry += " tax8.Tax_Code_Desc as tax8name,isnull (TSPL_SD_SHIPMENT_HEAD.tax8_amt,0) as txt8amt,   "
        atchqry += " tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SHIPMENT_HEAD.tax9_amt,0) as txt9amt,  "
        atchqry += " tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SHIPMENT_HEAD.tax10_amt,0) as txt10amt,  "
        atchqry += " isnull(TSPL_SD_SHIPMENT_HEAD .Total_Tax_Amt,0) as total_tax_amt, TSPL_SD_SHIPMENT_HEAD.Total_Amt as DocAmt,  TSPL_COMPANY_MASTER.Comp_Name as compname,ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+ Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as Phone,TSPL_COMPANY_MASTER.Fax ,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,ISNULL(tspl_company_Master.ADD1,'') as address1,"
        atchqry += " TSPL_SD_SHIPMENT_DETAIL.item_code as item_code, TSPL_ITEM_MASTER.Item_Desc   as itemdesc, TSPL_SD_SHIPMENT_DETAIL.Row_Type,TSPL_SD_SHIPMENT_DETAIL.Qty as qty,TSPL_SD_SHIPMENT_DETAIL.unit_code as uom,TSPL_SD_SHIPMENT_DETAIL.item_cost as itemcost,TSPL_SD_SHIPMENT_DETAIL.amount as amount,TSPL_SD_SHIPMENT_HEAD.TAX1,TSPL_SD_SHIPMENT_HEAD.TAX2,TSPL_SD_SHIPMENT_HEAD.TAX3,TSPL_SD_SHIPMENT_HEAD.TAX4,TSPL_SD_SHIPMENT_HEAD.TAX5,TSPL_SD_SHIPMENT_HEAD.Total_Add_Charge from TSPL_SD_SHIPMENT_DETAIL  "
        atchqry += " left outer join TSPL_SD_SHIPMENT_HEAD  on TSPL_SD_SHIPMENT_HEAD.Document_Code  =TSPL_SD_SHIPMENT_DETAIL.Document_Code   "
        atchqry += " left outer join  TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code =TSPL_SD_SHIPMENT_HEAD .Ship_To_Location "
        atchqry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_SHIP_TO_LOCATION.City_Code "
        atchqry += " left outer join TSPL_EMPLOYEE_MASTER on TSPL_EMPLOYEE_MASTER.EMP_CODE =TSPL_SD_SHIPMENT_HEAD.Salesman_Code "
        atchqry += " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SHIPMENT_HEAD.tax1  "
        atchqry += " left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SHIPMENT_HEAD.tax2  "
        atchqry += " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SHIPMENT_HEAD .TAX3  "
        atchqry += " left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SHIPMENT_HEAD .tax4  "
        atchqry += " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SHIPMENT_HEAD .tax5  "
        atchqry += " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX6  "
        atchqry += " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX7  "
        atchqry += " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX8  "
        atchqry += " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX9 "
        atchqry += " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SHIPMENT_HEAD .TAX10     "
        atchqry += " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_SD_SHIPMENT_HEAD.comp_code  "
        atchqry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SHIPMENT_HEAD.Customer_Code   "
        atchqry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_SD_SHIPMENT_HEAD.Bill_To_Location "
        atchqry += " Left Outer Join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  where 2=2 "
        atchqry += "  and  TSPL_SD_SHIPMENT_HEAD.Document_Code = '" + strCode + "'"
        SetItemWiseTax(clsDBFuncationality.GetDataTable(atchqry), txtDocNo.Value)
        Return atchqry
    End Function
    Private Sub RadMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem5.Click
        Dim frm As New FrmMailSMSSettingNew2()
        frm.FormId = clsUserMgtCode.frmShipmentProductSale
        frm.ShowDialog()
    End Sub
    Private Sub btnpreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpreview.Click
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            clsCommon.MyMessageBoxShow("First Select Shipment Document No.", Me.Text)
            txtDocNo.Focus()
            txtDocNo.Focus()
            Return
        End If
        atchqry = GetMailPrint(txtDocNo.Value)
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            System.Diagnostics.Process.Start(frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt, "crptShipment", "Shipment Detail"))
            frmCRV = Nothing
        End If
    End Sub
    Private Sub btnsend_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsend.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Please Select Document No. First", Me.Text)
                'txtReqNo.Focus()
                'txtReqNo.Select()
                Return
            End If
            If Not (common.clsCommon.MyMessageBoxShow("Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
                Return
            End If
            LoadData(txtDocNo.Value, NavigatorType.Current)
            Dim lstUsers As New List(Of String)
            lstUsers.Add(txtVendorNo.Value)
            'SendEmail(lstUsers, False)
            SendSMSandEmail(lstUsers, False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
#End Region
    Private Sub btnSendForApproval_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendForApproval.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Please Select Document No. First", Me.Text)
                'txtReqNo.Focus()
                'txtReqNo.Select()
                Return
            End If
            If Not (common.clsCommon.MyMessageBoxShow("Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
                Return
            End If
            LoadData(txtDocNo.Value, NavigatorType.Current)
            Dim qry As String = "Select * from TSPL_APPROVAL_LEVEL_SCREEN where 1=1 and Trans_Code='" + MyBase.Form_ID + "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            Dim lstUsers As New List(Of String)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    lstUsers.Add(dr("User_Code").ToString())
                Next
            End If
            If lstUsers.Count = 0 Then
                Throw New Exception("No Receiptent Found")
            End If
            SendSMSandEmail(lstUsers, True)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub SMSSENDONLY(ByVal isPost As Boolean)
    '    Try
    '        Dim strContactPerson As String = ""
    '        Dim strotherno As String = clsDBFuncationality.getSingleValue("select REPLACE( REPLACE( REPLACE(Phone1,'_',''),'(+)',''),'(+91)','') as Phone1 from TSPL_customer_MASTER where cust_code ='" & txtVendorNo.Value & "' ")
    '        If clsCommon.myLen(strotherno) > 0 Then
    '            If clsCommon.myLen(strotherno) = 10 Then
    '                strContactPerson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_customer_MASTER where cust_code ='" & txtVendorNo.Value & "' ")
    '                Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmMCCMaterial + "'", Nothing)
    '                Dim objSMSH As New clsSMSHead()
    '                objSMSH.arrMobilNo = New List(Of String)()
    '                objSMSH.arrMobilNo.Add(clsCommon.myCstr(strotherno))
    '                If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
    '                    If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
    '                        Dim strItem As String = ""
    '                        Dim strQty As String = ""
    '                        For ii As Integer = 0 To gv1.Rows.Count - 1
    '                            If clsCommon.myLen(gv1.Rows(ii).Cells(colICode).Value) > 0 Then
    '                                If clsCommon.myLen(strItem) > 0 Then
    '                                    strItem += ", "
    '                                    strQty += ", "
    '                                End If
    '                                strItem += clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
    '                                strQty += clsCommon.myCstr(gv1.Rows(ii).Cells(colQty).Value)
    '                            End If
    '                        Next
    '                        objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.InvoiceNo, txtInvoiceNo.Text)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SaleOrderNo, txtDocNo.Value)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerNo, txtVendorNo.Value)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerName, lblVendorName.Text)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(lblTotRAmt.Text))
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Form_Code, clsUserMgtCode.frmMCCMaterial)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ContactPerson, strContactPerson)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.ItemCode, strItem)
    '                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Qty, strQty)
    '                    End If
    '                    If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
    '                        objSMSH.SaveData(clsUserMgtCode.frmMCCMaterial, objSMSH, Nothing)
    '                        objSMSH = Nothing
    '                        If Not isPost Then
    '                            clsCommon.MyMessageBoxShow("SMS Send Successfully", Me.Text)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Sub
    '----------------------Done By Preeti Gupta 29/05/2014-------BM00000002659----------
    Private Sub btnHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHistory.Click
        Dim frm As New FrmSaleHistory
        frm.strFormId = MyBase.Form_ID
        frm.strCustId = txtVendorNo.Value
        frm.strCustName = lblVendorName.Text
        Dim strvendor As String = txtVendorNo.Value
        frm.ShowDialog()
        frm.WindowState = FormWindowState.Maximized
    End Sub
    'Public Function ExporttoExcel_transaction(ByVal sql As String, ByVal whrClaus As String, ByVal OrderByClaus As String, ByVal frm As RadForm) As Boolean
    '    Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
    '    Dim wBook As Microsoft.Office.Interop.Excel.Workbook
    '    'Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
    '    ''************* Filter Block Start
    '    Dim sfd As SaveFileDialog = New SaveFileDialog()
    '    Dim path As String
    '    sfd.FileName = frm.Text
    '    sfd.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
    '    If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
    '        path = sfd.FileName
    '    Else
    '        Return False
    '    End If
    '    Dim frmFilter As New frmFilterToExport()
    '    Dim gv As New UserControls.MyRadGridView
    '    wBook = excel.Workbooks.Add()
    '    Dim iQueryIndex As Integer = 0
    '    For Each row_sql As String In Regex.Split(sql, "####")
    '        Try
    '            iQueryIndex += 1
    '            Dim filename As String = row_sql.Substring(row_sql.IndexOf("From ") + 5, row_sql.Length - row_sql.IndexOf("From ") - 5)
    '            frmFilter.qry = row_sql
    '            frmFilter.whrCls = " Where 2=2 " + whrClaus
    '            If clsCommon.myLen(OrderByClaus) > 0 Then
    '                frmFilter.orderByClause = " Order by " + OrderByClaus
    '            End If
    '            frmFilter.ShowDialog()
    '            If frmFilter.isCancel Then
    '                Return False
    '            End If
    '            row_sql = frmFilter.qry
    '            ''************* Filter Block End
    '            If Not path.Equals(String.Empty) Then
    '                Try
    '                    Dim exporter As New RadGridViewExcelExporter()
    '                    gv.Name = "gTax"
    '                    frm.Controls.Add(gv)
    '                    FillGridView(row_sql, gv)
    '                    If gv.Rows.Count = 0 Then
    '                        common.clsCommon.MyMessageBoxShow("There is no data to transfer.")
    '                        Return False
    '                    End If
    '                    Dim i As Integer = 0
    '                    For i = 0 To gv.ColumnCount - 1
    '                        Dim grow As GridViewRowInfo = TryCast(gv.Rows(0), GridViewRowInfo)
    '                        If TypeOf grow.Cells(i).Value Is DateTime Then
    '                            Dim datecol As GridViewDateTimeColumn = TryCast(gv.Columns(i), GridViewDateTimeColumn)
    '                            datecol.ExcelExportType = DisplayFormatType.ShortDate
    '                        End If
    '                    Next i
    '                    'AddHandler exporter.Progress, AddressOf PB.exporter_Progress
    '                    'ShowPb(gv.Rows.Count)
    '                    'exporter.Export(gv, path, frm.Text)
    '                    'HidePb()
    '                    clsCommon.ProgressBarShow()
    '                    '=================Export================
    '                    exportdata_transaction(gv, path, filename, wBook, excel, iQueryIndex) 'path.Substring(path.LastIndexOf("\") + 1, path.Length - path.LastIndexOf("\") - 1), wBook) 'frm.Text)
    '                    '===============================================================
    '                    'MessageBox.Show(wBook.Worksheets.Count)
    '                Catch ex As Exception
    '                    frm.Controls.Remove(gv)
    '                    clsCommon.ProgressBarHide()
    '                    'HidePb()
    '                    Throw New Exception(ex.Message)
    '                    Return False
    '                End Try
    '            End If
    '        Catch ex As Exception
    '            common.clsCommon.MyMessageBoxShow("No data transfered." + Environment.NewLine + ex.Message, "Export Error", MessageBoxButtons.OK)
    '        End Try
    '    Next
    '    Dim strFileName As String = path
    '    Dim blnFileOpen As Boolean = False
    '    Try
    '        Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
    '        fileTemp.Close()
    '    Catch ex As Exception
    '        blnFileOpen = False
    '    End Try
    '    If System.IO.File.Exists(strFileName) Then
    '        System.IO.File.Delete(strFileName)
    '    End If
    '    wBook.SaveAs(strFileName)
    '    Dim dp As Object = wBook.BuiltinDocumentProperties
    '    dp("Keywords").Value = clsUserMgtCode.frmComplaintMaster
    '    excel.DisplayAlerts = False
    '    wBook.Close(True)
    '    'Dim exporter As New ExportToExcelML(gv)
    '    'AddHandler exporter.ExcelCellFormatting, AddressOf exporter_ExcelCellFormatting
    '    'exporter.ExportHierarchy = True
    '    '' exporter.ExportVisualSettings = True
    '    'exporter.SheetMaxRows = ExcelMaxRows._65536
    '    'exporter.SheetName = frm.Text
    '    'exporter.RunExport(path)
    '    frm.Controls.Remove(gv)
    '    clsCommon.ProgressBarHide()
    '    common.clsCommon.MyMessageBoxShow("Data transfer Completed!", "Export", MessageBoxButtons.OK)
    '    'Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
    '    excel.Workbooks.Open(path)
    '    excel.Visible = True
    '    Return True
    'End Function
    'Public Function exportdata_transaction(ByVal gv As RadGridView, ByVal flname As String, ByVal sname As String, ByVal wBook As Microsoft.Office.Interop.Excel.Workbook, ByVal excel As Microsoft.Office.Interop.Excel.ApplicationClass, ByVal Queryindex As Integer)
    '    If ((gv.Columns.Count = 0) Or (gv.Rows.Count = 0)) Then
    '        Exit Function
    '    End If
    '    'Creating dataset to export
    '    Dim dset As New DataSet
    '    'add table to dataset
    '    dset.Tables.Add()
    '    'add column to that table
    '    For i As Integer = 0 To gv.ColumnCount - 1
    '        dset.Tables(0).Columns.Add(gv.Columns(i).HeaderText)
    '    Next
    '    'add rows to the table
    '    Dim dr1 As DataRow
    '    For i As Integer = 0 To gv.RowCount - 1
    '        dr1 = dset.Tables(0).NewRow
    '        For j As Integer = 0 To gv.Columns.Count - 1
    '            dr1(j) = gv.Rows(i).Cells(j).Value.ToString
    '        Next
    '        dset.Tables(0).Rows.Add(dr1)
    '    Next
    '    'Dim excel As New Microsoft.Office.Interop.Excel.ApplicationClass
    '    'Dim wBook As Microsoft.Office.Interop.Excel.Workbook
    '    Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
    '    ' wBook = excel.Workbooks.Add()
    '    wSheet = wBook.Sheets("Sheet" & Queryindex & "") 'wBook.ActiveSheet()
    '    ' wSheet = wBook.ActiveSheet
    '    If clsCommon.myLen(sname) > 31 Then
    '        sname = sname.Substring(0, 31)
    '    End If
    '    'Dim dp As Object
    '    ' dp = wBook.BuiltinDocumentProperties
    '    'dp("Tags").Value = clsUserMgtCode.frmComplaintMaster
    '    wSheet.Name = sname
    '    wSheet.Cells.NumberFormat = "@"
    '    Try
    '        '  CType(wBook.Sheets("Sheet2"), Microsoft.Office.Interop.Excel.Worksheet).Delete()
    '        ' CType(wBook.Sheets("Sheet3"), Microsoft.Office.Interop.Excel.Worksheet).Delete()
    '    Catch ex As Exception
    '    End Try
    '    wSheet.Activate()
    '    Dim dt As System.Data.DataTable = dset.Tables(0)
    '    Dim dc As System.Data.DataColumn
    '    Dim dr As System.Data.DataRow
    '    Dim colIndex As Integer = 0fsetuser
    '    Dim rowIndex As Integer = 0
    '    For Each dc In dt.Columns
    '        colIndex = colIndex + 1
    '        wSheet.Cells(1, colIndex) = dc.ColumnName 'excel.
    '    Next
    '    For Each dr In dt.Rows
    '        rowIndex = rowIndex + 1
    '        colIndex = 0
    '        For Each dc In dt.Columns
    '            colIndex = colIndex + 1
    '            wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName).ToString 'excel.
    '        Next
    '    Next
    '    wSheet.Columns.AutoFit()
    '    'Dim strFileName As String = flname
    '    'Dim blnFileOpen As Boolean = False
    '    'Try
    '    '    Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
    '    '    fileTemp.Close()
    '    'Catch ex As Exception
    '    '    blnFileOpen = False
    '    'End Try
    '    'If System.IO.File.Exists(strFileName) Then
    '    '    System.IO.File.Delete(strFileName)
    '    'End If
    '    'wBook.SaveAs(strFileName)
    '    'Dim dp As Object = wBook.BuiltinDocumentProperties
    '    'dp("Keywords").Value = clsUserMgtCode.frmComplaintMaster
    '    'excel.DisplayAlerts = False
    '    'wBook.Close(True)
    'End Function
    Private Sub Export_Head_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export_Head.Click
        Dim sQuery As String = "select Document_Code as [Document Code],RANK() over(order by Document_Code desc) as [S No],Document_Date AS [Document Date],Customer_Code as [Customer Code],Description,Remarks,Bill_To_Location as [Bill To Location]," _
        & " Ship_To_Location as [Ship To Location],Tax_Group as [Tax Group],PROJECT_ID as [Project Id],Carrier,VehicleNo,Vehicle_Code" _
        & " as [Vehicle Code],GRNo,GENo,GEDate,Dept,Salesman_Code as [Salesman Code],Salesman_Name as [Salesman Name],Price_code as [Price Code]," _
        & " Route_No as [Route No],Add_Charge_Code1,Add_Charge_Amt1,Add_Charge_Code2,Add_Charge_Amt2,Add_Charge_Code3,Add_Charge_Amt3" _
        & " ,Add_Charge_Code4,Add_Charge_Amt4,Add_Charge_Code5,Add_Charge_Amt5,Add_Charge_Code6,Add_Charge_Amt6,Add_Charge_Code7,Add_Charge_Amt7" _
        & " ,Add_Charge_Code8,Add_Charge_Amt8,Add_Charge_Code9,Add_Charge_Amt9,Add_Charge_Code10,Add_Charge_Amt10 from TSPL_SD_SHIPMENT_HEAD"
        transportSql.ExporttoExcel(sQuery, Me)
    End Sub
    Private Sub Export_details_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export_details.Click
        Dim sQuery As String = "select DOCUMENT_CODE as [Document Code],RANK() over(order by Document_Code desc) as [S No],Item_Code as [Item Code],Price_Date as [Price Date]," _
        & " Unit_code as [UOM],Qty as [Quantity],MRP,Item_Cost  as [Item Cost],Location,Remarks,Specification from tspl_SD_Shipment_detail"
        transportSql.ExporttoExcel(sQuery, Me)
    End Sub
    Function LoadImportData(ByVal gv As RadGridView, ByVal rowindex As Integer, ByVal DtExcel As DataTable)
        Dim iswithouterror As Boolean = False
        Try
            btnSave.Enabled = True
            btnPost.Enabled = True
            btnDelete.Enabled = True
            isInsideLoadData = True
            isNewEntry = True
            btnSave.Text = "Update"
            BlankAllControls()
            LoadBlankGrid()
            LoadBlankGridTax()
            LoadBlankGridAC()
            cboItemType.Enabled = False
            txtBillToLocation.Enabled = False
            Dim obj As New frmShipmentImportExport()
            obj.SaveData(gv, rowindex, DtExcel)
            If (obj IsNot Nothing) Then 'AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                    repoComplete.IsVisible = True
                    repoBalQty.IsVisible = True
                End If
                chkVendorGrossReceipt.Checked = clsVendorMaster.isGrossReceipt(obj.Customer_Code)
                UsLock1.Status = obj.Status
                txtDocNo.Value = obj.Document_Code
                txtDate.Value = obj.Document_Date
                txtVendorNo.Value = obj.Customer_Code
                txtDate.Enabled = False
                txtVendorNo.Enabled = False
                chkTaxable.Enabled = False
                chkRateUserCustomer.ToggleState = ClsUserCustomerSettings.GetUserCustomerRateSetting(txtVendorNo.Value)
                lblVendorName.Text = obj.Customer_Name
                Dim dt_vlc As DataTable = clsDBFuncationality.GetDataTable("select * from tspl_vlc_master_Head where vsp_code='" & clsCommon.myCstr(txtVendorNo.Value) & "'")
                If dt_vlc.Rows.Count > 0 Then
                    LblVlc_Code.Text = dt_vlc.Rows(0).Item("vlc_code_vlc_Uploader")
                    LblVlc_Code.Tag = dt_vlc.Rows(0).Item("vlc_code")
                    LblVlc_Name.Text = dt_vlc.Rows(0).Item("vlc_name")
                Else
                    LblVlc_Code.Text = ""
                    LblVlc_Code.Tag = Nothing
                    LblVlc_Name.Text = ""
                End If
                txtRefNo.Text = obj.Ref_No
                If clsCommon.myLen(obj.Challan_Date) > 0 Then
                    dtpChallan.Value = obj.Challan_Date
                    dtpChallan.Checked = True
                Else
                    dtpChallan.Checked = False
                End If
                If clsCommon.myLen(obj.Inv_Date) > 0 Then
                    dtpInvoice.Value = obj.Inv_Date
                    dtpInvoice.Checked = True
                Else
                    dtpInvoice.Checked = False
                End If
                chkOnHold.Checked = obj.On_Hold
                txtDesc.Text = obj.Description
                txtTaxGroup.Value = obj.Tax_Group
                txtComment.Text = obj.Comments
                'txtShipToLocation.Value = obj.Ship_To_Location
                txtBillToLocation.Value = obj.Bill_To_Location
                txtInvNo.Text = obj.Inv_No
                Dim objTaxGrpMaster As New clsTaxGroupMaster()
                objTaxGrpMaster = objTaxGrpMaster.GetDataForSale(obj.Tax_Group)
                If (objTaxGrpMaster IsNot Nothing) Then
                    lblTaxGrpName.Text = objTaxGrpMaster.Tax_Group_Desc
                End If
                chkInternal.Checked = obj.Is_Internal
                cboItemType.SelectedValue = obj.Item_Type
                txtDept.Value = obj.Dept
                lblDept.Text = obj.Dept_Desc
                txtTermCode.Value = obj.Terms_Code
                'lblTermName.Text = obj.Terms_Description
                If obj.Due_Date IsNot Nothing Then
                    txtDueDate.Value = obj.Due_Date
                End If
                lblAmtWithDiscount.Text = clsCommon.myFormat(obj.Discount_Base)
                lblDiscountAmt.Text = clsCommon.myFormat(obj.Discount_Amt)
                lblAmtAfterDiscount.Text = clsCommon.myFormat(obj.Amount_Less_Discount)
                lblTaxAmt.Text = clsCommon.myFormat(obj.Total_Tax_Amt)
                lblTotRAmt.Text = clsCommon.myFormat(obj.Total_Amt)
                lblTotRAmt1.Text = lblTotRAmt.Text
                lblBillToLocation.Text = obj.BillToLocationName
                'lblShipToLocation.Text = obj.ShipToLocationName
                lblTaxGrpName.Text = obj.TaxGroupName
                lblTermName.Text = obj.TermsName
                txtCarrier.Text = obj.Carrier
                txtGRNo.Text = obj.GRNo
                txtGENo.Text = obj.GENo
                If obj.GEDate.HasValue Then
                    txtGEDate.Value = obj.GEDate
                    txtGEDate.Checked = True
                End If
                txtInvoiceNo.Text = obj.Invoice_No
                'txtReqNo.Value = obj.Against_Sales_Order
                fndProject.Value = obj.PROJECT_ID
                lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")
                'If clsCommon.myLen(txtReqNo.Value) > 0 Then
                '    lblProject.Enabled = False
                '    fndProject.Enabled = False
                'End If
                txtPriceCode.Text = obj.Price_Code
                txtDiscPer.Text = obj.HeadDisc_Per
                txtDiscAmt.Text = obj.HeadDisc_Amt
                ddlInvoiceType.SelectedValue = obj.Invoice_Type
                If (clsCommon.myLen(obj.TAX1) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX1
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX1_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX1_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX1_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX1) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX2) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX2
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX2_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX2_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX2_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX2) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX3) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX3
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX3_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX3_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX3_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX3) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX4) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX4
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX4_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX4_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX4_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX4) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX5) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX5
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX5_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX5_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX5_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX5) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX6) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX6
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX6_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX6_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX6_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX6) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX7) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX7
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX7_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX7_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX7_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX7) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX8) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX8
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX8_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX8_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX8_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX8) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX9) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX9
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX9_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX9_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX9_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX9) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX10) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX10
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX10_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.TAX10_Base_Amt
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX10_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX10) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                ''gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code1) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code1
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name1
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt1
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code2) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code2
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name2
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt2
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code3) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code3
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name3
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt3
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code4) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code4
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name4
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt4
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code5) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code5
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name5
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt5
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code6) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code6
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name6
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt6
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code7) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code7
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name7
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt7
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code8) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code8
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name8
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt8
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code9) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code9
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name9
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt9
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code10) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACCode).Value = obj.Add_Charge_Code10
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACName).Value = obj.Add_Charge_Name10
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colACAmount).Value = obj.Add_Charge_Amt10
                End If
                lblAddCharges.Text = clsCommon.myFormat(obj.Total_Add_Charge)
                lblAddCharges1.Text = clsCommon.myFormat(obj.Total_Add_Charge)
                If obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic Then
                    rbtnTaxCalAutomatic.IsChecked = True
                ElseIf obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual Then
                    rbtnTaxCalManual.IsChecked = True
                End If
                chkCreateAutoInvoice.Checked = obj.Is_Create_Auto_Invoice
                'chkCreateAutoReceipt.Visible = chkCreateAutoInvoice.Checked
                chkCreateAutoReceipt.Checked = obj.Is_Create_Auto_Receipt
                If obj.arrList IsNot Nothing AndAlso obj.arrList.Count > 0 Then
                    For Each objTr As frmShipmentImportExport In obj.arrList
                        gv1.Rows.AddNew()
                        ' gv1.Rows(gv1.Rows.Count - 1).Tag = objTr.arrSrItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = objTr.Line_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = objTr.Row_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colComplete).Value = IIf(objTr.Status = 0, "No", "Yes")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objTr.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHSNNo).Value = clsItemMaster.GetItemHSNCode(objTr.Item_Code, Nothing)
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colBarCode).Value = objTr.Bar_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(objTr.Item_Code)
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgSOQty).Value = objTr.so_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(objTr.Unit_code), 0)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(objTr.Unit_code), 0)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = objTr.Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFreeQty).Value = objTr.Free_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objTr.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = objTr.Order_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objTr.Item_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = objTr.strLocation
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = objTr.LocationName
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPer).Value = objTr.Disc_Per
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisAmt).Value = objTr.Disc_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterDis).Value = objTr.Amt_Less_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax1).Value = objTr.TAX1
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt1).Value = objTr.TAX1_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = objTr.TAX1_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt1).Value = objTr.TAX1_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax2).Value = objTr.TAX2
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt2).Value = objTr.TAX2_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = objTr.TAX2_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt2).Value = objTr.TAX2_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax3).Value = objTr.TAX3
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt3).Value = objTr.TAX3_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = objTr.TAX3_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt3).Value = objTr.TAX3_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax4).Value = objTr.TAX4
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt4).Value = objTr.TAX4_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = objTr.TAX4_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt4).Value = objTr.TAX4_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax5).Value = objTr.TAX5
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt5).Value = objTr.TAX5_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = objTr.TAX5_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt5).Value = objTr.TAX5_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax6).Value = objTr.TAX6
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt6).Value = objTr.TAX6_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = objTr.TAX6_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt6).Value = objTr.TAX6_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax7).Value = objTr.TAX7
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt7).Value = objTr.TAX7_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = objTr.TAX7_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt7).Value = objTr.TAX7_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax8).Value = objTr.TAX8
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt8).Value = objTr.TAX8_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = objTr.TAX8_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt8).Value = objTr.TAX8_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax9).Value = objTr.TAX9
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt9).Value = objTr.TAX9_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = objTr.TAX9_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt9).Value = objTr.TAX9_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax10).Value = objTr.TAX10
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt10).Value = objTr.TAX10_Base_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = objTr.TAX10_Rate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt10).Value = objTr.TAX10_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = objTr.Total_Tax_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterTax).Value = objTr.Item_Net_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = objTr.MRP
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableRate).Value = objTr.Assessable
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colAssessableAmount).Value = objTr.AssessableAmt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = objTr.Batch_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = objTr.Bin_No
                        If objTr.Expiry_Date.HasValue Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colExpiry).Value = objTr.Expiry_Date
                        End If
                        If objTr.MFG_Date.HasValue Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colManufactureDate).Value = objTr.MFG_Date
                        End If
                        If clsCommon.myLen(objTr.Order_Code) > 0 Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = clsSNSalesOrderDetail.GetBalancePOQty(objTr.Order_Code, objTr.Item_Code, obj.Document_Code, objTr.Unit_code, objTr.MRP, objTr.Assessable)
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSpecification).Value = objTr.Specification
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRemarks).Value = objTr.Remarks
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsMannualAmt).Value = objTr.Is_Mannual_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = IIf(objTr.Scheme_Applicable = "Y", "Yes", "No")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFromSchemeCode).Value = objTr.Scheme_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = IIf(objTr.Scheme_Item = "Y", "Yes", "No")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = objTr.Item_Tax
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalMRP).Value = objTr.Total_MRP_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalBasicAmount).Value = objTr.Total_Basic_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalDiscountAmount).Value = objTr.Total_Disc_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colcustDiscount).Value = objTr.Cust_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalCustDiscount).Value = objTr.Total_Cust_Discount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colActualCost).Value = objTr.ActualRate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCustDiscountQty).Value = objTr.Cust_DiscountQty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceDateColumn).Value = objTr.Price_Date
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceCOde).Value = objTr.Price_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementPer).Value = objTr.Abatement_Per
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementAmount).Value = objTr.Abatement_Amt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColFOC).Value = objTr.FOC_Item
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value), 0)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = objTr.Batch_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemWeight).Value = objTr.Item_Weight
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colConvF).Value = objTr.Conv_Factor
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTotItemWt).Value = objTr.TotalItem_Weight
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkupOn).Value = objTr.Markup_On
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkUpPercentage).Value = objTr.Markup_Percent
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLandingCost).Value = objTr.Landing_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCustDiscPercentage).Value = objTr.CustDiscPer
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = objTr.HeadDiscAmt
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCashDiscSchemeCode).Value = objTr.CasdDiscScheme_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPurCost).Value = objTr.Purchase_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgCost).Value = objTr.OrgRate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleCode).Value = objTr.PrincipleCode
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleDesc).Value = objTr.PrincipleDesc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorCode).Value = objTr.vendor_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorDesc).Value = objTr.vendor_desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = "Yes"
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = "Yes"
                        If obj.Status = ERPTransactionStatus.Pending Then
                            If clsCommon.myLen(obj.TAX1) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable1).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX1)
                            End If
                            If clsCommon.myLen(obj.TAX2) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable2).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX2)
                            End If
                            If clsCommon.myLen(obj.TAX3) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable3).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX3)
                            End If
                            If clsCommon.myLen(obj.TAX4) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable4).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX4)
                            End If
                            If clsCommon.myLen(obj.TAX5) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable5).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX5)
                            End If
                            If clsCommon.myLen(obj.TAX6) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable6).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX6)
                            End If
                            If clsCommon.myLen(obj.TAX7) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable7).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX7)
                            End If
                            If clsCommon.myLen(obj.TAX8) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable8).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX8)
                            End If
                            If clsCommon.myLen(obj.TAX9) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable9).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX9)
                            End If
                            If clsCommon.myLen(obj.TAX10) > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRecoverable10).Value = clsTaxMaster.IsTaxRecoverableAC(obj.TAX10)
                            End If
                        End If
                    Next
                    If obj.Status = ERPTransactionStatus.Pending Then
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                        gvAC.Rows.AddNew()
                    End If
                    ''For Custom Fields
                    If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                        UcCustomFields1.LoadData(obj.Document_Code)
                    End If
                    clsCustomFieldGrid.FillDataInGrid(obj.Document_Code, MyBase.Form_ID, gv1)
                    ''End of For Custom Fields
                    '' MULTICURRENCY
                    Me.txtCurrencyCode.Value = obj.CURRENCY_CODE
                    Me.txtConversionRate.Text = obj.ConvRate
                    Me.txtApplicableFrom.Text = obj.ApplicableFrom.ToString
                    '' end  MULTICURRENCY
                End If
                SetitemWiseTaxOnlySetting()
                For Each row As GridViewRowInfo In gv1.Rows
                    UpdateCurrentRow(row.Index)
                Next
                UpdateAllTotals()
                Dim ee As System.EventArgs = Nothing
                If AllowToSave(False) = False Then
                    iswithouterror = False
                    GoTo a
                End If
                btnSave_Click("", ee)
                DocumentNo = txtDocNo.Value
                ' ''RefreshReqNo()
                ' ''RefreshGRPNo()
                '  UcAttachment1.LoadData(obj.Document_Code)
                iswithouterror = True
a:          End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            iswithouterror = False
        Finally
            isInsideLoadData = False
        End Try
        Return iswithouterror
    End Function
    Private Sub btnDeliveredTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeliveredTo.Click
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        'Try
        '    Dim strDatabaseName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select DataBase_Name from TSPL_COMPANY_MASTER  where Cust_Code='" & txtVendorNo.Value & "'", trans))
        '    If clsCommon.myLen(strDatabaseName) = 0 Then
        '        Throw New Exception("Please set main database name for company " + strDatabaseName + " ")
        '    End If
        '    If common.clsCommon.MyMessageBoxShow("Are you sure to deliver Shipment ?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
        '        Dim qry = "Update TSPL_SD_SHIPMENT_HEAD set Is_Delivered=1  where Document_Code='" + txtDocNo.Value + "'"
        '        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        '        qry = "Update " + strDatabaseName + ".dbo.TSPL_PURCHASE_ORDER_HEAD set SaleInvoiceNo='" & txtInvoiceNo.Text & "' where PurchaseOrder_No='" & txtPONo.Text & "'"
        '        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        '        clsCommon.MyMessageBoxShow("Data Delivered successfully")
        '    End If
        '    trans.Commit()
        '    LoadData(txtDocNo.Value, NavigatorType.Current)
        'Catch ex As Exception
        '    trans.Rollback()
        '    Throw New Exception(ex.Message)
        'End Try
    End Sub
    Private Sub BtnSaveLayout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSaveLayout.Click
        '    If clsCommon.myLen(GetReportID()) > 0 Then
        gv1.MasterTemplate.FilterDescriptors.Clear()
        Dim obj As New clsGridLayout()
        obj.ReportID = "MilkMaterialGrid"
        obj.UserID = objCommonVar.CurrentUserCode
        obj.GridLayout = New MemoryStream()
        gv1.SaveLayout(obj.GridLayout)
        obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
        obj.GridColumns = gv1.ColumnCount
        If obj.SaveData() Then
            common.clsCommon.MyMessageBoxShow("Layout saved successfully", "Information")
        End If
        ''stuti regarding memory leakage
        obj.GridLayout.Close()
        obj.GridLayout.Dispose()
        'End If
    End Sub
    Private Sub BtnDeleteLayout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnDeleteLayout.Click
        clsGridLayout.DeleteData("MilkMaterialGrid", objCommonVar.CurrentUserCode)
        clsCommon.MyMessageBoxShow("Layout Deleted Successfully.", Me.Text)
        LoadData(txtDocNo.Value, NavigatorType.Current)
    End Sub
    Private Sub ReStoreGridLayout()
        Try
            'If clsCommon.myLen("LoadinMainGrid") > 0 Then
            Dim obj As clsGridLayout = New clsGridLayout()
            obj = CType(obj.GetData("MilkMaterialGrid", "", objCommonVar.CurrentUserCode), clsGridLayout)
            If Not obj Is Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
                Dim ii As Integer
                For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
                    gv1.Columns(ii).IsVisible = False
                    gv1.Columns(ii).VisibleInColumnChooser = True
                Next
                gv1.LoadLayout(obj.GridLayout)
                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            End If
            'End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs)
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                Dim obj As New clsMCCMaterialSale
                obj.EWayBillNo = txtEWayBillNo.Text
                obj.Electronic_Ref_No = txtElectronicRefNo.Text
                If txtEWayBillDate.Checked Then
                    obj.EWayBillDate = txtEWayBillDate.Value
                End If
                If clsMCCMaterialSale.UpdateAfterPosting(obj, txtDocNo.Value, Nothing) Then
                    clsCommon.MyMessageBoxShow("Information updated successfully.", Me.Text)
                End If
            Else
                Throw New Exception("Document no not found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtDate.Validating
        SETGSTControl()
    End Sub
    Sub SETGSTControl()
        GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        If (GSTStatus AndAlso chkTaxable.Checked = True) OrElse (GSTStatus = False) Then
            txtTaxGroup.Enabled = True
        Else
            txtTaxGroup.Enabled = False
        End If
    End Sub
    Private Sub chkTaxable_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkTaxable.ToggleStateChanged
        SETGSTControl()
        If clsCommon.myLen(txtDocNo.Value) = 0 AndAlso clsCommon.myLen(txtBillToLocation.Value) > 0 AndAlso clsCommon.myLen(txtVendorNo.Value) > 0 Then
            SetTax()
        End If
    End Sub
    Private Sub btnInvoiceJE_Click(sender As Object, e As EventArgs) Handles btnInvoiceJE.Click
        clsOpenJEAgainstInvoice.ShowInvoiceJE(txtInvoiceNo.Text)
    End Sub
    ' Ticket No : TEC/29/10/18-000352 By Prabhakar
    Private Sub btnShowInventory_Click(sender As Object, e As EventArgs) Handles btnShowInventory.Click
        clsOpenInventory.ShowInventoryDatails(txtDocNo.Value)
    End Sub
    Private Sub txtDate_ValueChanged(sender As Object, e As EventArgs) Handles txtDate.ValueChanged
        Try
            txtPaymentCycleNo.Text = clsGenratePaymentCycles.GetPaymentCycleNo(txtBillToLocation.Value, txtDate.Value)
            txtFiscalYear.Text = clsGenratePaymentCycles.GetPaymentFiscalCode(txtBillToLocation.Value, txtDate.Value)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub lblBillToLocation_Click(sender As Object, e As EventArgs) Handles lblBillToLocation.Click
    End Sub
    Private Sub txtSubLocation__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtSubLocation._MYValidating
        If clsCommon.myLen(txtBillToLocation.Value) = 0 Then
            clsCommon.MyMessageBoxShow("Please select Bill To location code before sub location", Me.Text)
            Exit Sub
        End If
        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
            txtSubLocation.Value = clsLocation.getFinder(" (isnull(is_sub_location,'N')='Y' or isnull(Is_Section,'N')='Y') and Main_Location_Code='" + txtBillToLocation.Value + "'", txtSubLocation.Value, isButtonClicked)
        End If
        If clsCommon.myLen(txtSubLocation.Value) > 0 Then
            lblSubLocation.Text = clsLocation.GetName(txtSubLocation.Value, Nothing)
        Else
            lblSubLocation.Text = ""
        End If
    End Sub
    Private Sub chkRateDiffAmt_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkRateDiffAmt.ToggleStateChanged, chkRateDiffRate.ToggleStateChanged
        If Not isInsideLoadData Then
            If chkRateDiffAmt.IsChecked Then
                txtRateAmt.Enabled = True
                txtRatePer.Enabled = False
                txtRatePer.Text = 0
                GroupBoxRateDiffAmtType.Visible = False
            Else
                txtRateAmt.Enabled = False
                txtRatePer.Enabled = True
                txtRatePer.Text = 0
                GroupBoxRateDiffAmtType.Visible = True
            End If
        End If
    End Sub
    Private Sub txtRateAmt_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRateAmt.Leave
        CalculateRateDiffAmount()
    End Sub
    Private Sub txtRatePer_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtRatePer.Leave
        CalculateRateDiffAmount()
    End Sub
    Private Sub CalculateRateDiffAmount()
        Try
            If clsCommon.myCDecimal(txtRateAmt.Text) > (clsCommon.myCDecimal(lblAmtAfterDiscount.Text) + clsCommon.myCDecimal(lblTaxAmt.Text)) Then
                txtRateAmt.Text = 0
                lblTotalSubsidy.Text = 0
                Throw New Exception("Rate Difference amount cannot be greater than sum of Discount after amount and Tax amount")
            End If
            If clsCommon.myCDecimal(txtRatePer.Text) > 0 Then
                If rbtnBasicAmt.IsChecked Then
                    txtRateAmt.Text = clsCommon.myCDecimal(lblAmtWithDiscount.Text) * clsCommon.myCDecimal(txtRatePer.Text) / 100
                ElseIf rbtnTotalAmt.IsChecked Then
                    txtRateAmt.Text = clsCommon.myCDecimal(lblTotRAmt.Text) * clsCommon.myCDecimal(txtRatePer.Text) / 100
                End If
            ElseIf clsCommon.myCDecimal(txtRatePer.Text) = 0 Then
                If chkRateDiffRate.IsChecked Then
                    txtRateAmt.Text = 0
                End If
            ElseIf clsCommon.myCDecimal(txtRateAmt.Text) > 0 Then
                txtRatePer.Text = 0
            End If
            If chkRateDiffAmt.IsChecked Then
                If MultiplySubsidyWithQuantity Then
                    lblTotalSubsidy.Text = clsCommon.myCDecimal(txtRateAmt.Text) * TotalItemQty
                    lblGrossAmount.Text = clsCommon.myCDecimal(lblTotRAmt.Text) - clsCommon.myCDecimal(lblTotalSubsidy.Text)
                    If chkAddTPT.Checked Then
                        lblGrossAmount.Text = clsCommon.myCDecimal(lblGrossAmount.Text) + clsCommon.myCDecimal(txtTPTAmt.Text)
                    Else
                        lblGrossAmount.Text = clsCommon.myCDecimal(lblGrossAmount.Text) - clsCommon.myCDecimal(txtTPTAmt.Text)
                    End If

                Else
                    lblGrossAmount.Text = clsCommon.myCDecimal(lblTotRAmt.Text - txtRateAmt.Text) '

                    If chkAddTPT.Checked Then
                        lblGrossAmount.Text = clsCommon.myCDecimal(lblGrossAmount.Text) + clsCommon.myCDecimal(txtTPTAmt.Text)
                    Else
                        lblGrossAmount.Text = clsCommon.myCDecimal(lblGrossAmount.Text) - clsCommon.myCDecimal(txtTPTAmt.Text)
                    End If
                End If
            Else
                If rbtnTotalAmt.IsChecked Then
                    lblTotalSubsidy.Text = clsCommon.myCDecimal(lblTotRAmt.Text * txtRatePer.Text) / 100
                ElseIf rbtnBasicAmt.IsChecked Then
                    lblTotalSubsidy.Text = clsCommon.myCDecimal(lblAmtWithDiscount.Text * txtRatePer.Text) / 100
                End If

                lblGrossAmount.Text = clsCommon.myCDecimal(lblTotRAmt.Text - lblTotalSubsidy.Text) '
                If chkAddTPT.Checked Then
                    lblGrossAmount.Text = clsCommon.myCDecimal(lblGrossAmount.Text) + clsCommon.myCDecimal(txtTPTAmt.Text)
                Else
                    lblGrossAmount.Text = clsCommon.myCDecimal(lblGrossAmount.Text) - clsCommon.myCDecimal(txtTPTAmt.Text)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub chkcashsale_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkcashsale.ToggleStateChanged
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
            If chkcashsale.Checked Then
                lblReceiptNo.Visible = True
                txtReceiptNo.Visible = True
                lblReceiptAmt.Visible = True
                lblBankCode.Visible = True
                If objCommonVar.CreateAutoReceiptEntryDCSSale Then
                    lblBankCode.Visible = True
                    txtBankCode.Visible = True
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetBankcodeData(False))
                    If dt.Rows.Count > 0 Then
                        If dt.Rows.Count = 1 Then
                            txtBankCode.Value = dt.Rows(0)("Code")
                        End If
                    End If
                End If
                'lblTotalSubsidy.Visible = True
                lblPaymentType.Visible = True
                cmbPaymentType.Visible = True
                lblPaymentType.Location = New System.Drawing.Point(794, 1)
                cmbPaymentType.Location = New System.Drawing.Point(873, 1)
                MyLabel14.Text = "Subsidy"
                'lblPaymentType
            Else
                lblBankCode.Visible = False
                txtBankCode.Visible = False
                lblReceiptNo.Visible = False
                txtReceiptNo.Visible = False
                lblReceiptAmt.Visible = False
                '  lblTotalSubsidy.Visible = False
                lblPaymentType.Visible = False
                cmbPaymentType.Visible = False
                MyLabel14.Text = "Rate Difference"
            End If
        Else
            If chkcashsale.Checked Then
                lblReceiptNo.Visible = True
                If objCommonVar.CreateAutoReceiptEntryDCSSale Then
                    lblBankCode.Visible = True
                    txtBankCode.Visible = True
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetBankcodeData(False))
                    If dt.Rows.Count > 0 Then
                        If dt.Rows.Count = 1 Then
                            txtBankCode.Value = dt.Rows(0)("Code")
                        End If
                    End If
                End If
                txtReceiptNo.Visible = True
                lblReceiptAmt.Visible = True
                'lblTotalSubsidy.Visible = True
                lblPaymentType.Visible = True
                cmbPaymentType.Visible = True
                MyLabel14.Text = "Subsidy"
            Else
                lblBankCode.Visible = False
                txtBankCode.Visible = False
                lblReceiptNo.Visible = False
                txtReceiptNo.Visible = False
                lblReceiptAmt.Visible = False
                ' lblTotalSubsidy.Visible = False
                lblPaymentType.Visible = False
                cmbPaymentType.Visible = False
                MyLabel14.Text = "Rate Difference"
            End If

        End If
    End Sub
    Private Sub txtReceiptNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtReceiptNo._MYValidating
        If clsCommon.myLen(txtVendorNo.Value) > 0 Then
            Dim qry As String = "select TSPL_RECEIPT_HEADER.Receipt_No as Code,TSPL_RECEIPT_HEADER.Receipt_Amount as Amount,TSPL_RECEIPT_HEADER.Receipt_Date,TSPL_RECEIPT_HEADER.Receipt_Post_Date,TSPL_RECEIPT_HEADER.Receipt_Type,TSPL_RECEIPT_HEADER.Balance_Amt from TSPL_RECEIPT_HEADER"
            Dim whrcls As String = " TSPL_RECEIPT_HEADER.Cust_Code='" + clsCommon.myCstr(txtVendorNo.Value) + "'"
            whrcls += " and not exists (select TSPL_SD_SHIPMENT_HEAD.Receipt_No from TSPL_SD_SHIPMENT_HEAD where TSPL_RECEIPT_HEADER.Receipt_No =TSPL_SD_SHIPMENT_HEAD.Receipt_No ) and TSPL_RECEIPT_HEADER.Posted='Y'"
            txtReceiptNo.Value = clsCommon.ShowSelectForm("DBC-Receipt", qry, "Code", whrcls, txtReceiptNo.Value, "Code", isButtonClicked)
            lblReceiptAmt.Text = clsDBFuncationality.getSingleValue("select Receipt_Amount from TSPL_RECEIPT_HEADER where Receipt_No='" + clsCommon.myCstr(txtReceiptNo.Value) + "'")
        Else
            clsCommon.MyMessageBoxShow(Me, "Please Select Customer", Me.Text)
        End If
    End Sub
    Private Sub txtRateAmt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtRateAmt.KeyPress
        'If MultiplySubsidyWithQuantity Then
        lblTotalSubsidy.Text = clsCommon.myCDecimal(txtRateAmt.Text) * TotalItemQty
        'End If
    End Sub

    Private Sub txtDiscAmt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDiscAmt.KeyPress
        'If MultiplySubsidyWithQuantity Then
        lblTotalDisSubsidy.Text = clsCommon.myCDecimal(txtDiscAmt.Text) * TotalItemQty
        ' End If
    End Sub
    Private Sub chkisTCS_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkisTCS.ToggleStateChanged
        Try
            For ii As Integer = 0 To gv1.Rows.Count - 1
                UpdateCurrentRow(ii)
            Next
            UpdateAllTotals()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrintEwayBill_Click(sender As Object, e As EventArgs) Handles btnPrintEwayBill.Click
        Try
            If clsCommon.myLen(TxtEWayBillUpdateBillNo.Text) > 0 Then

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(clsPSInvoiceHead.PrintEWayBill(TxtEWayBillUpdateBillNo.Text, txtBillToLocation.Value))
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rpte-waybill", "E-WayBill", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    frmCRV = Nothing
                Else
                    Throw New Exception("No Data Found ")
                End If
            Else
                Throw New Exception("e-Way Bill not found!")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnCancel_Click(sender As Object, e As EventArgs) Handles rbtnCancel.Click
        Try
            Dim EInvoiceType As String = ""
            If clsCommon.myLen(txtInvoiceNo.Text) <= 0 Then
                Throw New Exception("Code is empty")
            End If
            If clsCommon.MyMessageBoxShow(Me, "Are you sure to Cancel the Record?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            Dim strSaleReturnNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Document_Code from TSPL_SD_SALE_RETURN_HEAD where Against_Invoice_No='" & txtInvoiceNo.Text & "' "))
            If clsCommon.myLen(strSaleReturnNo) > 0 Then
                Throw New Exception("You cannot cancelled this document because its Sale Return (" + clsCommon.myCstr(strSaleReturnNo) + ") has been created.")
            End If
            Dim strReceiptCount As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select receipt_no from TSPL_RECEIPT_DETAIL where Document_No in (Select Document_No from TSPL_Customer_Invoice_Head  where against_Sale_no='" & txtInvoiceNo.Text & "') "))
            If clsCommon.myLen(strReceiptCount) > 0 Then
                Throw New Exception("You cannot cancelled this document because receiving (" + clsCommon.myCstr(strReceiptCount) + ") has been done against its AR Invoice.")
            End If
            '' richa ERO/10/11/21-001547
            Dim strDairyGAtePassCount As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select gpcode from TSPL_SD_SHIPMENT_HEAD where document_code='" & txtDocNo.Value & "' "))
            If clsCommon.myLen(strDairyGAtePassCount) > 0 Then
                Throw New Exception("You cannot cancelled this document because Dairy GAte Pass (" + clsCommon.myCstr(strDairyGAtePassCount) + ") has been created.")
            End If
            If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                Dim EInvoiceCancelTimeValid As Int64 = 0
                EInvoiceCancelTimeValid = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,EInvoice_Posting_Date,GETDATE()),0) as PostedHours from tspl_sd_sale_invoice_head where  document_code = '" + txtInvoiceNo.Text + "'"))
                If EInvoiceCancelTimeValid >= 24 Then
                    Throw New Exception("Invoice can not be cancelled.It has been more than 24 hours.")
                End If
            End If
            clsPSShipmentHead.CancelData(Me.Form_ID, txtDocNo.Value, txtInvoiceNo.Text, NavigatorType.Current, False)
            clsCommon.MyMessageBoxShow(Me, "Successfully Cancelled", Me.Text)
            AddNew()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnHistory1_Click(sender As Object, e As EventArgs)
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocNo.Value, "Document_Code", "TSPL_SD_SHIPMENT_HEAD", "TSPL_SD_SHIPMENT_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub txtDiscPer_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDiscPer.KeyPress
        lblTotalDisSubsidy.Text = clsCommon.myRoundOFF(clsCommon.myCDecimal(lblAmtWithDiscount.Text) * (clsCommon.myCDecimal(txtDiscPer.Text) / 100), 2, 4)
    End Sub

    Private Sub txtTPTVendor__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTPTVendor._MYValidating
        Try
            Dim qry As String = "select Vendor_Code AS Code,Vendor_Name as Name from tspl_vendor_master "
            txtTPTVendor.Value = clsCommon.ShowSelectForm("DCSTPTVendor", qry, "Code", "  ISNULL(Transporter,'')='Y' ", txtTPTVendor.Value, "Code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkApplyTPT_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkApplyTPT.ToggleStateChanged
        If chkApplyTPT.Checked Then
            lblTPTVendor.Visible = True
            txtTPTVendor.Visible = True
        Else
            lblTPTVendor.Visible = False
            txtTPTVendor.Visible = False
        End If
    End Sub
    Private Sub txtBankCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtBankCode._MYValidating
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetBankcodeData(False))
            If dt.Rows.Count > 1 Then
                txtBankCode.Value = clsCommon.ShowSelectForm("ShipmentBankCodefnd", GetBankcodeData(True), "Code", " TSPL_USER_BANK_MASTER.User_Code='" + objCommonVar.CurrentUserCode + "'", txtBankCode.Value, "Code", isButtonClicked)
            Else
                txtBankCode.Value = clsCommon.ShowSelectForm("ShipmentBankCodefnd", GetBankcodeData(True), "Code", " TSPL_USER_BANK_MASTER.User_Code='" + objCommonVar.CurrentUserCode + "'", txtBankCode.Value, "Code", isButtonClicked)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Function GetBankcodeData(ByVal ShowAllbank As Boolean) As String
        Dim qry As String = "SELECT TSPL_BANK_MASTER.BANK_CODE AS Code,DESCRIPTION as Name FROM TSPL_BANK_MASTER left outer join TSPL_USER_BANK_MASTER on TSPL_USER_BANK_MASTER.BANK_CODE= TSPL_BANK_MASTER.Bank_Code  "
        If Not ShowAllbank Then
            qry += " where TSPL_USER_BANK_MASTER.User_Code='" + objCommonVar.CurrentUserCode + "'"
        End If
        Return qry
    End Function

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        Dim Tran As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(txtInvoiceNo.Text) <= 0 Then
                Throw New Exception("Please provice Invoice No")
            End If
            If clsCommon.myLen(TxtEInvoiceUpdateIRNNo.Text) <= 0 Then
                Throw New Exception("Please provice " + TxtEInvoiceUpdateIRNNo.MyLinkLable1.Text)
            End If
            If clsCommon.myLen(TxtEInvoiceUpdateAckNo.Text) <= 0 Then
                Throw New Exception("Please provice " + TxtEInvoiceUpdateAckNo.MyLinkLable1.Text)
            End If
            If clsCommon.myLen(TxtEInvoiceUpdateQCCode.Text) <= 0 Then
                Throw New Exception("Please provice " + TxtEInvoiceUpdateQCCode.MyLinkLable1.Text)
            End If
            If Not UsLock1.Status = ERPTransactionStatus.Pending Then
                Throw New Exception("Transaction should be Pending ")
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, txtInvoiceNo.Text, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Tran)
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "IRN_no", TxtEInvoiceUpdateIRNNo.Text)
            clsCommon.AddColumnsForChange(coll, "Ack_No", TxtEInvoiceUpdateAckNo.Text)
            clsCommon.AddColumnsForChange(coll, "Ack_Date", clsCommon.GetPrintDate(TxtEInvoiceUpdateAckDate.Value, "dd/MMM/yyyy hh:mm:ss tt"))
            clsCommon.AddColumnsForChange(coll, "QR_Code", TxtEInvoiceUpdateQCCode.Text)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_INVOICE_HEAD", OMInsertOrUpdate.Update, "TSPL_SD_SALE_INVOICE_HEAD.Document_Code='" + txtInvoiceNo.Text + "'", Tran)
            Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(TxtEInvoiceUpdateQCCode.Text)
            clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_SD_SALE_INVOICE_HEAD", "TSPL_SD_SALE_INVOICE_HEAD.Document_Code='" & txtInvoiceNo.Text & "'", Tran)
            Tran.Commit()
            clsCommon.MyMessageBoxShow(Me, "E-Invoice Details Updated successfully", Me.Text)
        Catch ex As Exception
            Tran.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadButton2_Click_1(sender As Object, e As EventArgs) Handles RadButton2.Click
        Dim Tran As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(txtInvoiceNo.Text) <= 0 Then
                Throw New Exception("Please provice Invoice No")
            End If
            If clsCommon.myLen(TxtEWayBillUpdateBillNo.Text) <= 0 Then
                Throw New Exception("Please provice " + TxtEWayBillUpdateBillNo.MyLinkLable1.Text)
            End If
            If Not UsLock1.Status = ERPTransactionStatus.Pending Then
                Throw New Exception("Transaction should be Pending")
            End If
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, txtInvoiceNo.Text, "TSPL_SD_SALE_INVOICE_HEAD", "Document_Code", "TSPL_SD_SALE_INVOICE_DETAIL", "Document_Code", Tran)
            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "EWayBillNo", TxtEWayBillUpdateBillNo.Text)
            clsCommon.AddColumnsForChange(coll, "EwayBillDate", clsCommon.GetPrintDate(TxtEWayBillUpdateBillDate.Value, "dd/MMM/yyyy hh:mm:ss tt"))
            clsCommon.AddColumnsForChange(coll, "EwayBillValidDate", clsCommon.GetPrintDate(TxtEWayBillUpdateValidDate.Value, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "EWayBillRemarks", TxtEWayBillUpdateBillRemarks.Text)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SD_SALE_INVOICE_HEAD", OMInsertOrUpdate.Update, "TSPL_SD_SALE_INVOICE_HEAD.Document_Code='" + txtInvoiceNo.Text + "'", Tran)
            Tran.Commit()
            clsCommon.MyMessageBoxShow(Me, "E-Waybill Details Updated successfully", Me.Text)
        Catch ex As Exception
            Tran.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub setDCSBalance()
        UcDCSBalance1.DCSCode = clsCommon.myCstr(LblVlc_Code.Tag)
        UcDCSBalance1.DCSName = clsCommon.myCstr(LblVlc_Name.Text)
        UcDCSBalance1.VendorCode = txtVendorNo.Value
        UcDCSBalance1.DCSUploaderCode = LblVlc_Code.Text
        UcDCSBalance1.TransDate = txtDate.Value
        UcDCSBalance1.RefreshData()
    End Sub

    Private Sub chkOther_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkOther.ToggleStateChanged
        If chkOther.Checked Then
            chkcashsale.Checked = True
            chkcashsale.Enabled = False
        Else
            chkcashsale.Checked = False
            chkcashsale.Enabled = True
        End If
    End Sub

    Private Sub btnCreateEWB_Click(sender As Object, e As EventArgs) Handles btnCreateEWB.Click
        Dim tran As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Create_Ewb(tran)
            tran.Commit()
            clsCommon.MyMessageBoxShow(Me, "EWB Created Successfully", Me.Text)
            LoadData(txtDocNo.Value, NavigatorType.Current)
        Catch ex As Exception
            tran.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Function Create_Ewb(ByVal trans As SqlTransaction) As Boolean

        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso clsCommon.myLen(txtInvoiceNo.Text) > 0 Then
                Dim EwbNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EWayBillNo from TSPL_SD_SALE_INVOICE_HEAD where  document_code = '" & txtInvoiceNo.Text & "' ", trans))
                If clsCommon.myLen(EwbNo) = 0 Then
                    If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value, trans) Then
                        Dim EInvoiceCancelTimeValid As Int64 = 0
                        EInvoiceCancelTimeValid = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,EInvoice_Posting_Date,GETDATE()),0) as PostedHours from tspl_sd_sale_invoice_head where  document_code = '" & txtInvoiceNo.Text & "'", trans))
                        If EInvoiceCancelTimeValid >= 24 Then
                            Throw New Exception("E-Way Bill can not be Generated.It has been more than 24 hours.")
                        End If
                    End If
                    If clsCommon.myLen(GetEWayBillNo(txtInvoiceNo.Text, trans)) <= 0 Then
                        clsPSInvoiceHead.EWayBill_Implementation(txtInvoiceNo.Text, txtBillToLocation.Value, trans, True)
                        If clsCommon.myLen(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" & txtInvoiceNo.Text & "'", trans)) <= 0 Then
                            Throw New Exception("E-Way Bill For Sales Invoice No [" + txtInvoiceNo.Text + "] is not generated")
                        End If
                    End If
                Else
                    Throw New Exception("Ewb no [" & clsCommon.myCstr(EwbNo) & "] is already exists")
                End If

            Else
                Throw New Exception("Please Select Document/Invoice ")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetEWayBillNo(strDocNo As String, trans As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" + strDocNo + "'", trans))
    End Function

    Private Sub btnUpdateVehicle_Click(sender As Object, e As EventArgs) Handles btnUpdateVehicle.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso UsLock1.Status = ERPTransactionStatus.Approved Then
                Dim frm As New FrmCommonUpdatesForEWB
                frm.strDocNo = txtDocNo.Value
                frm.strCustCode = txtVendorNo.Value
                frm.strTransId = ""
                frm.strTransName = ""
                frm.strVehicleNo = txtVehicleNo.Text
                frm.strScreenType = "DCSSale"
                frm.ShowDialog()
            Else
                Throw New Exception("Unable to proceed. You must select a Document Number that has already been approved.")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnTotalAmt_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnTotalAmt.ToggleStateChanged, rbtnBasicAmt.ToggleStateChanged
        If chkRateDiffAmt.IsChecked OrElse chkRateDiffRate.IsChecked Then
            CalculateRateDiffAmount()
        End If
    End Sub

    Private Sub txtTransporterCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTransporterCode._MYValidating
        Try
            Dim qry As String = "select Transport_Id,Transporter_Name from TSPL_TRANSPORT_MASTER"
            txtTransporterCode.Value = clsCommon.ShowSelectForm("DCSTransport No", qry, "Transport_Id", "", txtTransporterCode.Value, "Transport_Id", isButtonClicked)
            lblTransporterName.Text = connectSql.RunScalar("Select Transporter_Name  from TSPL_TRANSPORT_MASTER where Transport_Id = '" + Convert.ToString(txtTransporterCode.Value) + "'")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub chkNoTranspoter_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkNoTranspoter.ToggleStateChanged
        If chkNoTranspoter.Checked Then
            txtTransporterCode.Enabled = False
            lblTransporterName.Enabled = False
            txtTransporterCode.Value = ""
            lblTransporterName.Text = ""
        Else
            txtTransporterCode.Enabled = True
            lblTransporterName.Enabled = True
        End If
    End Sub
    Private Sub txtSaleRoute__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtSaleRoute._MYValidating
        Try
            Dim qry As String = " select Route_No as Code,Route_Desc as Name from TSPL_Route_Master "
            txtSaleRoute.Value = clsCommon.ShowSelectForm("SaleRoute", qry, "Code", "", txtSaleRoute.Value, "", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub FndPriceCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles FndPriceCode._MYValidating
        Try
            Dim qry As String = "Select Distinct Price_Code as Code, Price_Code_Desc as [Description] ,isnull(Transfer,0) as Transfer from TSPL_PRICE_COMPONENT_MAPPING"
            FndPriceCode.Value = clsCommon.ShowSelectForm("PriceDCSSaleEn", qry, "Code", "", FndPriceCode.Value, "", isButtonClicked)
            If clsCommon.myLen(FndPriceCode.Value) > 0 Then
                FndPriceCode.Enabled = False
            Else
                FndPriceCode.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class
