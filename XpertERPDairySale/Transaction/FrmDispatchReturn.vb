Imports common
Imports System.Globalization
Imports System.Data.SqlClient
Imports System.IO
'Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Net.Mail
Imports System.Net.WebClient
Imports System.Net
Imports Telerik.WinControls.Data

Public Class FrmDispatchReturn
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim SaleInvoiceDate As DateTime
    Dim ApplyCommission As Boolean = False
    Dim ApplyBoothWiseReturn As Boolean = False
    Dim ApplyCommissionRateWithTax As Boolean = False
    Dim CalculateTaxRatefromItemwsieTaxOnSale As Integer = 0
    Dim ApplyRoundOffZero As Boolean = False
    Public IncreaseCrateQtyOnFiftyPercent As Boolean = False
    Dim CreateCommonDairyDispatchforFreshAmbient As Boolean = False
    Public checkstockmrpwise As Boolean = False
    Dim AllowCrateCanPhysicalStock As Integer = 0
    Dim AutoCalculateCrate As Integer = 0
    Dim AutoCalculateCAN As Integer = 0
    Dim DispatchCommissionDecimalPlaces As Decimal = 4
    Dim ShowMulMRPOfSameItemOnDairyBookingCustomer As Boolean = False
    Const colCrate As String = "colCrate"
    Const colCan As String = "colCan"
    Public AllowModifcationByApprovalUser As Boolean = False
    Dim loading As Integer = 0
    Private StrSql As String
    Private blnBackCalculation As Boolean = False
    Private AllowChangeInvoiceType As Boolean = False
    Private IsBatchMFDEXDmandatory As Boolean = False
    Private PurchaseOneItemOneVendor As Boolean = False
    Private strExcise As Boolean
    Private intMRPwithabatement As Integer
    Private isPO_GRN_MRN_Editable As Boolean = False
    Public Const RowTypeItem As String = "Item"
    Public Const RowTypeMisc As String = "Misc"
    Const colOrgUnit As String = "COLORGUNIT"
    Const ReportID As String = "DSSaleReturnItemGrid"
    Public strSRNno As String = Nothing
    Private isCellValueChangedOpen As Boolean = False
    Private isCellValueChangedTaxOpen As Boolean = False
    Dim iStxtTaxGroup_TxtChangedComplete As Boolean = True
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Private isInvoiceLoadData As Boolean = False
    Const colRateUnitQty As String = "colRateUnitQty"
    Const colAlterUnitQty As String = "colAlterUnitQty"
    Const colLineNo As String = "COLLNO"
    Const colRowType As String = "COLTYPE"
    Const colComplete As String = "COMPLETE"
    Const colBalanceQty As String = "BALANCEQTY"
    Const colICode As String = "COLICODE"
    Const colIName As String = "COLINAME"
    Const colIHSN As String = "colIHSN"
    Const colPendingQty As String = "COLPENDINGQTY"

    'Const colOrgSOQty As String = "COLORGSOQTY"
    Const colQty As String = "COLQTY"

    Const colFreeQty As String = "COLFREEQTY"
    Const colUnit As String = "COLUNIT"
    Const colActualQty As String = "COLActQty"
    Const colConvAmt As String = "colConvAmt"
    Const colActualUOM As String = "COLActUom"
    Const colActualRetQty As String = "colActualRetQty"
    Const colRate As String = "COLRATE"
    Const colAmt As String = "COLAMT"
    Const colDisPer As String = "COLDISPER"
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


    ''Const colLocationCode As String = "LOCATIONCODE"
    ''Const colLocationName As String = "LOCATIONNAME"


    Const colMRP As String = "MRP"
    Const colItemBasicPrice As String = "colItemBasicPrice"
    '' ''Const colAssessableRate As String = "ASSESSABLERATE"
    '' ''Const colAssessableAmount As String = "ASSESSABLEAMT"

    Const colBatchNo As String = "BATCHNO"
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
    'Const colTIsTaxable As String = "ISTAXABLE"
    Const colTTaxRate As String = "TAXRATE"
    'Const colTIsSurTax As String = "ISSURTAX"
    'Const colTSurTaxCode As String = "SURTAXCODE"
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
    Const colIsTaxOnBaseAmount As String = "colIsTaxOnBaseAmount"
    Const colBinNo As String = "colBinNo"
    Const colPricipleCode As String = "colPricipleCode"
    Const colPricipleDesc As String = "colPricipleDesc"
    Const colvendorCode As String = "colvendorCode"
    Const colvendorDesc As String = "colvendorDesc"
    Const colReturnAmt As String = "colReturnAmt"
    Const colDamageAmt As String = "colDamageAmt"
    Const colDamageQty As String = "colDamageQty"
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
    Const colInvoiceSchemeCode As String = "colInvoiceSchemeCode"
    Const ColCustDiscountQty As String = "ColCustDiscountQty"
    Const colAbatementPer As String = "colAbatementPer"
    Const colAbatementAmount As String = "colAbatementamount"
    Const colPriceCOde As String = "colPriceCOde"
    Const colMarkUpPercentage As String = "colMarkUpPercentage"
    Const colLandingCost As String = "colLandingCost"
    Const colMarkupOn As String = "colMarkupOn"

    Const colCustDiscPercentage As String = "colCustDiscPercentage"
    Const colCashDiscSchemeCode As String = "colCashDiscSchemeCode"
    Const colHeadDiscamt As String = "colHeadDiscamt"
    Const colCashSchemeCode As String = "CashSchemeCode"
    Const colCashSchemeType As String = "CashSchType"
    Const colCash_Pers As String = "CashScPers"
    Const colCash_Amt As String = "CashSc_Amt"
    Const colSchmCodeType As String = "SchmCodeType"
    Const colMainIcode As String = "MainIcode"
    Const colMainIQty As String = "MainIQty"
    Const colMainIUOM As String = "MainIUOM"
    Const colICodeGrp As String = "colICodeGrp"
    Const colSchemeType As String = "colSchemeType"
    Const colVolumeScheme As String = "colVolumeScheme"
    Const colUnitALter As String = "colUnitALter"
    Const colUnitRate As String = "colUnitRate"
    Const colTAX_PAID As String = "colTAX_PAID"
    Const colPurCost As String = "colPurCost"
    Const colOrgCost As String = "colOrgCost"
    Public DocumentNo As String = Nothing
    Const colHeaDDisPer As String = "colHeaDDisPer"
    Const colHeadDisPerAmt As String = "colHeadDisPerAmt"

    Const ColDCApplicableDate As String = "ColDCApplicableDate"
    Const ColDCUOM As String = "ColDCUOM"
    Const ColDCRate As String = "ColDCRate"
    Const ColDCRateWithTax As String = "ColDCRateWithTax"
    Const ColDCUnitCF As String = "ColDCUnitCF"
    Const ColDCQtyinSU As String = "ColDCQtyinSU"
    Const ColDCCFUOM As String = "ColDCCFUOM"
    Const ColDCAmt As String = "ColDCAmt"
    Const ColTCRate As String = "ColTCRate"
    Const ColTCAmt As String = "ColTCAmt"
    Const ColSCRate As String = "ColSCRate"
    Const ColSCAmt As String = "ColSCAmt"
    Const ColDCPKID As String = "ColDCPKID"
    Const colCF As String = "colCF"
    Const colCFKG As String = "colCFKG"
    Const colQtyinKG As String = "colQtyinKG"
    Const colTransporter As String = "colTransporter"

    Dim atchmntqry As String = ""
    Dim UOMAtDiarySaleReturn As Boolean = Nothing
    Dim IsFormLoad As Boolean = False
    Const colIsBatchItem As String = "colIsBatchItem"
    Dim StrPDFPath As String = Nothing
    Const colDelivery_Code As String = "colDelivery_Code"
    Const colDocno As String = "colDocno"
    Const colBoothCode As String = "colBoothCode"
    Const colBoothName As String = "colBoothName"
    Const colItem_Code As String = "colItem_Code"
    Const colItem_Name As String = "colItem_Name"
    Const colUnit_Code As String = "colUnit_Code"
    Const colBQty As String = "colBQty"
    Const colBRQty As String = "colBRQty"




    Dim FlagDocumentIsTaxable As Integer = 0
    Dim EInvoiceType As String = ""
#End Region
    Public Sub SetUserMgmtNew()
        ''Richa Check in 19/06/2020
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")

        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnPrint.Visible = MyBase.isPrintFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnCancel.Visible = MyBase.isCancel_Flag_After_Posting
        btnReverseAndUnpost.Visible = False

        'If MyBase.isReverse Then
        '    btnReverseAndUnpost.Enabled = True
        'Else
        '    btnReverseAndUnpost.Enabled = False
        'End If
    End Sub

    Private Sub FrmDispatchReturn_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btnHistory.Enabled = False
        SetUserMgmtNew()
        SetMailRight()
        IncreaseCrateQtyOnFiftyPercent = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.IncreaseCrateQtyOnFiftyPercent, clsFixedParameterCode.IncreaseCrateQtyOnFiftyPercent, Nothing)) = 1, True, False)
        AllowCrateCanPhysicalStock = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowCratePhysicalStock, clsFixedParameterCode.AllowCratePhysicalStock, Nothing))
        AutoCalculateCrate = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AutoCalculateCrateOnDairyDispatch & "'")) = 0, 0, 1)
        blnBackCalculation = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsRateBackCalculation from TSPL_inv_parameters")) = 0, False, True)
        PurchaseOneItemOneVendor = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='PurchaseOneItemOneVendor'")) = 0, False, True)
        isPO_GRN_MRN_Editable = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select isMRNQtyEdiatableOnSRN from TSPL_inv_parameters")) = 0, False, True)
        AutoCalculateCAN = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AutoCalculateCANOnDairyDispatch & "'")) = 0, 0, 1)
        ShowMulMRPOfSameItemOnDairyBookingCustomer = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowMulMRPOfSameItemOnDairyBookingCustomer, clsFixedParameterCode.ShowMulMRPOfSameItemOnDairyBookingCustomer, Nothing)) = 1, True, False)
        CreateCommonDairyDispatchforFreshAmbient = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateCommonDairyDispatchforFreshAmbient, clsFixedParameterCode.CreateCommonDairyDispatchforFreshAmbient, Nothing)) = 1, True, False)
        checkstockmrpwise = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.checkstockMRPwise, clsFixedParameterCode.checkstockMRPwise, Nothing)) = 0, False, True)
        ApplyCommission = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyCommission, clsFixedParameterCode.ApplyCommission, Nothing)) = 1, True, False)
        ApplyCommissionRateWithTax = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyCommissionRateWithTax, clsFixedParameterCode.ApplyCommissionRateWithTax, Nothing)) = 1, True, False)
        ApplyRoundOffZero = If(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyRoundOffZero, clsFixedParameterCode.ApplyRoundOffZero, Nothing)) = 1, True, False)
        'ApplyBoothWiseReturn = If(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBoothWiseReturn, clsFixedParameterCode.ApplyBoothWiseReturn, Nothing)) = 1, True, False)
        CalculateTaxRatefromItemwsieTaxOnSale = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, Nothing))
        DispatchCommissionDecimalPlaces = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DispatchCommissionDecimalPlaces, clsFixedParameterCode.DispatchCommissionDecimalPlaces, Nothing))
        'crateTable()
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
        LoadBlankGrid()
        LoadBlankGridTax()
        LoadItemType()
        LoadBlankGridAC()
        'LoadBlankGridBooth()
        AddNew()
        SetLength()
        LoadReturnType()
        If clsCommon.myLen(strSRNno) > 0 Then
            LoadData(strSRNno, NavigatorType.Current)
        Else
            strSRNno = Me.Tag
            If clsCommon.myLen(strSRNno) > 0 Then
                LoadData(strSRNno, NavigatorType.Current)
            End If
        End If
        chkRateDefaultSetting.ToggleState = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SalesRateEditable, clsFixedParameterCode.SalesRateEditable, Nothing)) = 1, ToggleState.On, ToggleState.Off)

        ''For Custom Fields
        RadPageView1.Pages("pvpCustomFields").Item.Visibility = MyBase.customFieldTabProperty
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.Report_ID = MyBase.Form_ID
            UcCustomFields1.LoadCustomControls()
        End If
        ''End of For Custom Fields

        '' MultiCurrency
        '   RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
        'If ApplyBoothWiseReturn Then
        '    RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Visible
        'Else
        '    RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
        'End If
        SetMultiCurrencyVisibility()
        '' End of MultiCurrency
        txtBillToLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
        If clsCommon.myLen(txtBillToLocation.Value) > 0 Then
            lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_Location_Master where Location_Code='" + txtBillToLocation.Value + "' "))
        End If
        ''For Attachment
        If objCommonVar.IsDemoERP Then
            UcAttachment1.Form_ID = MyBase.Form_ID
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Visible
        Else
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Collapsed
        End If

        ''End of For Attachment
        If clsCommon.myLen(DocumentNo) > 0 Then
            LoadData(DocumentNo, NavigatorType.Current)
        End If
        fndProject.Enabled = False
        lblProject.Enabled = False

        UOMAtDiarySaleReturn = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.UOMAtDiarySaleReturn, clsFixedParameterCode.UOMAtDiarySaleReturn, Nothing)) = 1, True, False)
        If UOMAtDiarySaleReturn = True Then
            gv1.Columns(colQty).ReadOnly = True
            gv1.Columns(colUnit).ReadOnly = True
            gv1.Columns(colActualQty).IsVisible = True
            gv1.Columns(colActualUOM).IsVisible = True
            gv1.Columns(colConvAmt).IsVisible = True
        Else
            gv1.Columns(colActualQty).IsVisible = True
            gv1.Columns(colActualUOM).IsVisible = True
            'gv1.Columns(colActualRetQty).IsVisible = False
            gv1.Columns(colConvAmt).IsVisible = False
        End If
        If AllowCrateCanPhysicalStock = 1 And AutoCalculateCrate = 0 Then
            clsDBFuncationality.ExecuteNonQuery("update tspl_fixed_parameter set Description=1 where code='Auto Calculate Crate on Dairy Dispatch'")
            AutoCalculateCrate = 1
        End If

        If AutoCalculateCAN = 1 Then
            Panel2.Visible = True
        End If
        RadMenuItem5.Visibility = ElementVisibility.Collapsed
        IsFormLoad = True
    End Sub
    Sub SetMultiCurrencyVisibility()
        Dim strq As String = ""
        '=======shivani
        Dim Currency As Integer = clsModuleCurrencyMapping.GetmulticurrencyDecimalPlaces()
        txtConversionRate.DecimalPlaces = clsCommon.myCdbl(Currency)
        '================
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
        'Dim strq As String
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
    Public Shared Function GetReturnType() As DataTable
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))

        Dim dr As DataRow = dt.NewRow()
        dr("Code") = "D"
        dr("Name") = "Damaged Goods"
        dt.Rows.Add(dr)


        dr = dt.NewRow()
        dr("Code") = "P"
        dr("Name") = "Price Only"
        dt.Rows.Add(dr)


        dr = dt.NewRow()
        dr("Code") = "I"
        dr("Name") = "Inventory Type"
        dt.Rows.Add(dr)

        Return dt
    End Function
    Sub LoadReturnType()
        ddlReturnType.DataSource = Nothing
        loading = 1
        ddlReturnType.DataSource = GetReturnType()
        loading = 0
        ddlReturnType.ValueMember = "Code"
        ddlReturnType.DisplayMember = "Name"
        ddlReturnType.SelectedValue = "I"
    End Sub
    Sub SetLength()
        txtDocNo.MyMaxLength = 30
        txtDesc.MaxLength = 200
        txtComment.MaxLength = 200
        txtRefNo.MaxLength = 50
        txtCarrier.MaxLength = 50

        'txtGRNo.MaxLength = 50
        txtGENo.MaxLength = 50

        txtPONo.MaxLength = 200

    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
                        gv1.Columns(ii).IsVisible = False
                        gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Sub LoadItemType()
        cboItemType.DataSource = clsItemMaster.GetItemType()
        cboItemType.ValueMember = "Code"
        cboItemType.DisplayMember = "Name"
        'cboItemType.SelectedIndex = 2
        cboItemType.Visible = False
        RadLabel29.Visible = False
    End Sub


    Sub BlankAllControls()
        rbtnMorning.IsChecked = True
        isInvoiceLoadData = False
        TxtTotalCAN.Value = 0
        txtShippedCan.Value = 0
        txtCrateQty.Value = 0
        txtCrate.Value = 0
        txtJaali.Value = 0
        txtBox.Value = 0
        ddlReturnType.Enabled = True
        ddlReturnType.SelectedValue = "I"
        txtPONo.Text = ""
        txtPriceCode.Text = ""
        txtRouteNo.Value = ""
        lblRouteDesc.Text = ""
        txtPriceGroupCode.Text = ""
        txtInvoiceType.Text = ""
        txtVehcileCode.Value = ""
        txtVehicleNo.Text = ""
        txtSalesman.Value = ""
        lblSalesman.Text = ""
        txtDocNo.Value = ""
        txtDesc.Text = ""
        chkOnHold.Checked = False
        txtVendorNo.Value = ""
        lblVendorName.Text = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        dtpChallan.Value = clsCommon.GETSERVERDATE()
        dtpInvoice.Value = clsCommon.GETSERVERDATE()
        txtBillToLocation.Value = ""
        lblBillToLocation.Text = ""
        txtShipToLocation.Value = ""
        lblShipToLocation.Text = ""
        txtSubLocation.Value = ""
        lblSubLocation.Text = ""
        txtDesc.Text = ""
        txtInvNo.Text = ""
        txtComment.Text = ""
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
        TxtRoundoff.Text = ""
        UsLock1.Status = ERPTransactionStatus.Pending
        txtCarrier.Text = ""
        txtVehicleNo.Text = ""
        'txtGRNo.Text = ""
        txtGENo.Text = ""
        txtGEDate.Checked = False
        txtGEDate.Value = txtDate.Value

        txtDept.Value = ""
        lblDept.Text = ""
        cboItemType.SelectedIndex = 2
        cboItemType.Enabled = True
        txtBillToLocation.Enabled = True
        txtReqNo.Value = ""
        If ApplyBoothWiseReturn Then
            txtReqNo.Enabled = False
        Else
            txtReqNo.Enabled = True
        End If
        chkVendorGrossReceipt.Checked = False
        lblAddCharges1.Text = ""
        lblAddCharges1.Text = ""
        rbtnTaxCalAutomatic.IsChecked = True
        EInvoiceType = ""
        FlagDocumentIsTaxable = 0
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.BlankAllControls()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()
        fndProject.Text = ""
        lblProject.Text = ""

        If UOMAtDiarySaleReturn = True Then
            gv1.Columns(colQty).ReadOnly = True
            gv1.Columns(colUnit).ReadOnly = True
            gv1.Columns(colActualQty).IsVisible = True
            gv1.Columns(colActualUOM).IsVisible = True
            gv1.Columns(colConvAmt).IsVisible = True
        Else
            gv1.Columns(colActualQty).IsVisible = False
            gv1.Columns(colActualUOM).IsVisible = False
            'gv1.Columns(colActualRetQty).IsVisible = False
            gv1.Columns(colConvAmt).IsVisible = False
        End If
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
        repoRowType.Width = 50
        repoRowType.ReadOnly = False
        repoRowType.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        repoRowType.DataSource = GetItemType()
        repoRowType.ValueMember = "Code"
        repoRowType.DisplayMember = "Code"
        gv1.MasterTemplate.Columns.Add(repoRowType)


        Dim repoICode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoICode.FormatString = ""
        repoICode.HeaderText = "Item Code"
        repoICode.Name = colICode
        repoICode.HeaderImage = My.Resources.search4
        repoICode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoICode.Width = 100
        gv1.MasterTemplate.Columns.Add(repoICode)

        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIName.FormatString = ""
        repoIName.HeaderText = "Item Description"
        repoIName.Name = colIName
        repoIName.Width = 150
        repoIName.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIName)

        Dim repoIHSN As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIHSN.FormatString = ""
        repoIHSN.HeaderText = "HSN Code"
        repoIHSN.Name = colIHSN
        repoIHSN.Width = 150
        repoIHSN.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIHSN)

        Dim repoICodeGrp As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoICodeGrp.FormatString = ""
        repoICodeGrp.HeaderText = "Item Group"
        repoICodeGrp.Name = colICodeGrp
        repoICodeGrp.HeaderImage = My.Resources.search4
        repoICodeGrp.TextImageRelation = TextImageRelation.TextBeforeImage
        repoICodeGrp.Width = 100
        repoICodeGrp.ReadOnly = True
        repoICodeGrp.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoICodeGrp)


        Dim repoPriceDate As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoPriceDate.Format = DateTimePickerFormat.Custom
        repoPriceDate.CustomFormat = "dd-MM-yyyy"
        repoPriceDate.HeaderText = "Price Date"
        repoPriceDate.WrapText = True
        repoPriceDate.FormatString = "{0:d}"
        repoPriceDate.Name = colPriceDateColumn
        repoPriceDate.ReadOnly = True
        repoPriceDate.IsVisible = False
        repoPriceDate.Width = 80
        gv1.MasterTemplate.Columns.Add(repoPriceDate)


        Dim repoPriceCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPriceCode.FormatString = ""
        repoPriceCode.HeaderText = "Price Code"
        repoPriceCode.Name = colPriceCOde
        repoPriceCode.IsVisible = False
        repoPriceCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoPriceCode)

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

        Dim repoUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnit.FormatString = ""
        repoUnit.HeaderText = "UOM"
        repoUnit.Name = colUnit
        repoUnit.Width = 80
        repoUnit.IsVisible = False
        repoUnit.HeaderImage = My.Resources.search4
        repoUnit.TextImageRelation = TextImageRelation.TextBeforeImage
        gv1.MasterTemplate.Columns.Add(repoUnit)



        Dim repoAlterUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAlterUnit.FormatString = ""
        repoAlterUnit.HeaderText = "Alter UOM"
        repoAlterUnit.Name = colUnitALter
        repoAlterUnit.Width = 80
        repoAlterUnit.ReadOnly = True
        repoAlterUnit.IsVisible = False
        repoAlterUnit.VisibleInColumnChooser = False
        gv1.MasterTemplate.Columns.Add(repoAlterUnit)

        Dim reporateUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        reporateUnit.FormatString = ""
        reporateUnit.HeaderText = "Rate UOM"
        reporateUnit.Name = colUnitRate
        reporateUnit.Width = 80
        reporateUnit.IsVisible = False
        reporateUnit.VisibleInColumnChooser = False
        gv1.MasterTemplate.Columns.Add(reporateUnit)

        Dim repoOrgUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoOrgUnit.FormatString = ""
        repoOrgUnit.HeaderText = "ORG UOM"
        repoOrgUnit.Name = colOrgUnit
        repoOrgUnit.Width = 80
        repoOrgUnit.ReadOnly = False
        repoOrgUnit.IsVisible = False
        repoOrgUnit.VisibleInColumnChooser = False
        gv1.MasterTemplate.Columns.Add(repoOrgUnit)


        Dim repoQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoQty = New GridViewDecimalColumn()
        repoQty.FormatString = "{0:n3}"
        repoQty.HeaderText = "Return Quantity"
        repoQty.Name = colQty
        repoQty.Width = 80
        repoQty.Minimum = 0
        repoQty.IsVisible = False
        repoQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoQty)

        Dim repoCrate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCrate.FormatString = ""
        repoCrate.HeaderText = "Crate"
        repoCrate.Name = colCrate
        If AutoCalculateCrate = 1 Then
            repoCrate.IsVisible = True
        Else
            repoCrate.IsVisible = False
        End If
        repoCrate.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoCrate)

        Dim repoCan As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCan.FormatString = ""
        repoCan.HeaderText = "CAN"
        repoCan.Name = colCan
        If AutoCalculateCAN = 1 Then
            repoCan.IsVisible = True
        Else
            repoCan.IsVisible = False
        End If
        repoCan.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoCan)

        'Dim repoActualRetQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        'repoActualRetQty.FormatString = "{0:n3}"
        'repoActualRetQty.HeaderText = "Actual Return Qty"
        'repoActualRetQty.Name = colActualRetQty
        'repoActualRetQty.Width = 80
        'repoActualRetQty.Minimum = 0
        'repoActualRetQty.IsVisible = False
        'repoActualRetQty.TextImageRelation = TextImageRelation.TextBeforeImage
        'gv1.MasterTemplate.Columns.Add(repoActualRetQty)


        Dim repoActualUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoActualUnit.FormatString = ""
        repoActualUnit.HeaderText = "Return UOM"
        repoActualUnit.Name = colActualUOM
        repoActualUnit.Width = 80
        repoActualUnit.IsVisible = True
        repoActualUnit.HeaderImage = My.Resources.search4
        repoActualUnit.TextImageRelation = TextImageRelation.TextBeforeImage
        gv1.MasterTemplate.Columns.Add(repoActualUnit)

        Dim repoActualQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoActualQty.FormatString = "{0:n3}"
        repoActualQty.WrapText = True
        repoActualQty.HeaderText = "Return Qty"
        repoActualQty.Name = colActualQty
        repoActualQty.Width = 80
        repoActualQty.Minimum = 0
        repoActualQty.IsVisible = True
        repoActualQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoActualQty.DecimalPlaces = 3
        gv1.MasterTemplate.Columns.Add(repoActualQty)

        Dim repoActuConvAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoActuConvAmt.FormatString = ""
        repoActuConvAmt.WrapText = True
        repoActuConvAmt.HeaderText = "Conv.Amt"
        repoActuConvAmt.Name = colConvAmt
        repoActuConvAmt.Width = 80
        repoActuConvAmt.Minimum = 0
        repoActuConvAmt.ReadOnly = True
        repoActuConvAmt.IsVisible = False
        repoActuConvAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoActuConvAmt)

        Dim repoDamageQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDamageQty = New GridViewDecimalColumn()
        repoDamageQty.FormatString = "{0:n3}"
        repoDamageQty.HeaderText = "Damage Quantity"
        repoDamageQty.Name = colDamageQty
        repoDamageQty.Width = 80
        repoDamageQty.Minimum = 0
        repoDamageQty.IsVisible = True
        repoDamageQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDamageQty)


        Dim repoMainUnitQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMainUnitQty = New GridViewDecimalColumn()
        repoMainUnitQty.FormatString = ""
        repoMainUnitQty.HeaderText = "Rate Unit Qty"
        repoMainUnitQty.Name = colRateUnitQty
        repoMainUnitQty.Width = 80
        repoMainUnitQty.Minimum = 0
        repoMainUnitQty.ShowUpDownButtons = False
        repoMainUnitQty.Step = 0
        repoMainUnitQty.IsVisible = False
        repoMainUnitQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoMainUnitQty)

        Dim repoAlterUnitQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAlterUnitQty = New GridViewDecimalColumn()
        repoAlterUnitQty.FormatString = ""
        repoAlterUnitQty.HeaderText = "Alter Unit Qty"
        repoAlterUnitQty.Name = colAlterUnitQty
        repoAlterUnitQty.Width = 80
        repoAlterUnitQty.Minimum = 0
        repoAlterUnitQty.ShowUpDownButtons = False
        repoAlterUnitQty.Step = 0
        repoAlterUnitQty.IsVisible = False
        repoAlterUnitQty.VisibleInColumnChooser = False
        repoAlterUnitQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAlterUnitQty)

        Dim repoActualBalQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoActualBalQty.FormatString = ""
        repoActualBalQty.HeaderText = "Actual Balance"
        repoActualBalQty.Name = ColActualBalQty
        repoActualBalQty.Width = 80
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
        repoSchemeApp.Width = 75
        gv1.MasterTemplate.Columns.Add(repoSchemeApp)

        Dim repoItemWt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoItemWt = New GridViewDecimalColumn()
        repoItemWt.FormatString = ""
        repoItemWt.HeaderText = "Item Weight"
        repoItemWt.Name = colItemWeight
        repoItemWt.Width = 80
        repoItemWt.Minimum = 0
        repoItemWt.ReadOnly = False
        repoItemWt.IsVisible = False
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

        Dim TAX_PAID As New GridViewComboBoxColumn
        TAX_PAID.FormatString = ""
        TAX_PAID.HeaderText = "Tax Paid"
        TAX_PAID.Name = colTAX_PAID
        TAX_PAID.Width = 100
        TAX_PAID.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        TAX_PAID.DataSource = clsCSABooking.GetTaxPaid()
        TAX_PAID.ValueMember = "Code"
        TAX_PAID.DisplayMember = "Name"
        TAX_PAID.ReadOnly = True
        gv1.Columns.Add(TAX_PAID)


        Dim repoMRP As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMRP.FormatString = ""
        repoMRP.HeaderText = "MRP"
        repoMRP.Name = colMRP
        repoMRP.Width = 80
        repoMRP.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoMRP.ReadOnly = False
        gv1.MasterTemplate.Columns.Add(repoMRP)

        Dim repoItemBasicPrice As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoItemBasicPrice = New GridViewDecimalColumn()
        repoItemBasicPrice.FormatString = ""
        repoItemBasicPrice.HeaderText = "Basic Price"
        repoItemBasicPrice.Name = colItemBasicPrice
        repoItemBasicPrice.Width = 80
        repoItemBasicPrice.Minimum = 0
        repoItemBasicPrice.ReadOnly = True
        repoItemBasicPrice.IsVisible = True
        repoItemBasicPrice.VisibleInColumnChooser = False
        repoItemBasicPrice.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoItemBasicPrice)


        Dim repoFreeQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoFreeQty.FormatString = ""
        repoFreeQty.HeaderText = "Free Quantity"
        repoFreeQty.Name = colFreeQty
        repoFreeQty.Width = 80
        repoFreeQty.Minimum = 0
        repoFreeQty.IsVisible = False
        repoFreeQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoFreeQty)

        ''Dim repoLocationCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        ''repoLocationCode.FormatString = ""
        ''repoLocationCode.HeaderText = "Location Code"
        ''repoLocationCode.Name = colLocationCode
        ''repoLocationCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        ''repoLocationCode.TextImageRelation = TextImageRelation.TextBeforeImage
        ''repoLocationCode.Width = 100
        ''gv1.MasterTemplate.Columns.Add(repoLocationCode)

        ''Dim repoLocationName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        ''repoLocationName.FormatString = ""
        ''repoLocationName.HeaderText = "Location"
        ''repoLocationName.Name = colLocationName
        ''repoLocationName.ReadOnly = True
        ''repoLocationName.Width = 150
        ''gv1.MasterTemplate.Columns.Add(repoLocationName)



        Dim repoRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoRate = New GridViewDecimalColumn()
        repoRate.FormatString = ""
        repoRate.HeaderText = "Basic Rate"
        repoRate.Name = colRate
        repoRate.Width = 80
        repoRate.Minimum = 0
        repoRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoRate)

        Dim repoOrgRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoOrgRate = New GridViewDecimalColumn()
        repoOrgRate.FormatString = ""
        repoOrgRate.HeaderText = "Org Basic Rate"
        repoOrgRate.Name = colOrgCost
        repoOrgRate.Width = 80
        repoOrgRate.Minimum = 0
        repoOrgRate.IsVisible = False
        repoOrgRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoOrgRate)


        Dim repoIsSchmItem13 As GridViewComboBoxColumn = New GridViewComboBoxColumn()
        repoIsSchmItem13.FormatString = ""
        repoIsSchmItem13.HeaderText = "Scheme Type"
        repoIsSchmItem13.Name = colSchmCodeType
        repoIsSchmItem13.Width = 50
        repoIsSchmItem13.ReadOnly = True
        repoIsSchmItem13.DataSource = clsSchemeApplyOnDairy.GetSchemeTypes()
        repoIsSchmItem13.DisplayMember = "Code"
        repoIsSchmItem13.ValueMember = "Name"
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem13)

        Dim repoQtySchemeItem As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoQtySchemeItem.AllowSort = False
        repoQtySchemeItem.HeaderText = "Qty Scheme Item"
        repoQtySchemeItem.Name = colSchemeItem
        repoQtySchemeItem.ReadOnly = True
        repoQtySchemeItem.Width = 96
        gv1.MasterTemplate.Columns.Add(repoQtySchemeItem)

        Dim repoFromSchemeCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFromSchemeCode.HeaderText = "Scheme Code"
        repoFromSchemeCode.Name = colFromSchemeCode
        repoFromSchemeCode.Width = 80
        repoFromSchemeCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoFromSchemeCode)

        Dim repoInvoiceSchemeCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoInvoiceSchemeCode.HeaderText = "Invoice Scheme Code"
        repoInvoiceSchemeCode.Name = colInvoiceSchemeCode
        repoInvoiceSchemeCode.Width = 80
        repoInvoiceSchemeCode.ReadOnly = True
        repoInvoiceSchemeCode.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoInvoiceSchemeCode)

        Dim repoIsSchmItem3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIsSchmItem3.FormatString = ""
        repoIsSchmItem3.HeaderText = "Main Item Code"
        repoIsSchmItem3.Name = colMainIcode
        repoIsSchmItem3.Width = 50
        repoIsSchmItem3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem3)

        Dim repoIsSchmItem4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIsSchmItem4.FormatString = ""
        repoIsSchmItem4.HeaderText = "Main Item UOM"
        repoIsSchmItem4.Name = colMainIUOM
        repoIsSchmItem4.Width = 50
        repoIsSchmItem4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem4)

        Dim repoIsSchmItem5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIsSchmItem5.FormatString = ""
        repoIsSchmItem5.HeaderText = "Main Item Qty"
        repoIsSchmItem5.Name = colMainIQty
        repoIsSchmItem5.Width = 50
        repoIsSchmItem5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem5)

        Dim repoIsSchmItem1c As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIsSchmItem1c.FormatString = ""
        repoIsSchmItem1c.HeaderText = "Cash Scheme Code"
        repoIsSchmItem1c.Name = colCashSchemeCode
        repoIsSchmItem1c.Width = 50
        repoIsSchmItem1c.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem1c)

        Dim repoIsSchmItem1c1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIsSchmItem1c1.FormatString = ""
        repoIsSchmItem1c1.HeaderText = "Cash Scheme Type"
        repoIsSchmItem1c1.Name = colCashSchemeType
        repoIsSchmItem1c1.Width = 50
        repoIsSchmItem1c1.ReadOnly = True
        repoIsSchmItem1c1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem1c1)

        Dim repoIsSchmItem1c2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoIsSchmItem1c2.FormatString = ""
        repoIsSchmItem1c2.HeaderText = "Cash %"
        repoIsSchmItem1c2.Name = colCash_Pers
        repoIsSchmItem1c2.Width = 50
        repoIsSchmItem1c2.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem1c2)

        Dim repoIsSchmItem1c3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoIsSchmItem1c3.FormatString = ""
        repoIsSchmItem1c3.HeaderText = "Cash Amount"
        repoIsSchmItem1c3.Name = colCash_Amt
        repoIsSchmItem1c3.Width = 50
        repoIsSchmItem1c3.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIsSchmItem1c3)


        Dim repoAbatementRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAbatementRate = New GridViewDecimalColumn()
        repoAbatementRate.FormatString = ""
        repoAbatementRate.HeaderText = "Abatement %"
        repoAbatementRate.Name = colAbatementPer
        repoAbatementRate.Width = 80
        repoAbatementRate.Minimum = 0
        repoAbatementRate.ReadOnly = False
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
        repoAbatementamount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAbatementamount)

        Dim repoReturnAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoReturnAmt = New GridViewDecimalColumn()
        repoReturnAmt.FormatString = ""
        repoReturnAmt.HeaderText = "Return Amt"
        repoReturnAmt.Name = colReturnAmt
        repoReturnAmt.Width = 80
        repoReturnAmt.Minimum = 0
        repoReturnAmt.ReadOnly = True
        repoReturnAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoReturnAmt)

        Dim repoDamageAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDamageAmt = New GridViewDecimalColumn()
        repoDamageAmt.FormatString = ""
        repoDamageAmt.HeaderText = "Damage Amt"
        repoDamageAmt.Name = colDamageAmt
        repoDamageAmt.Width = 80
        repoDamageAmt.Minimum = 0
        repoDamageAmt.ReadOnly = True
        repoDamageAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDamageAmt)


        Dim repoAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmt = New GridViewDecimalColumn()
        repoAmt.FormatString = ""
        repoAmt.HeaderText = "Total Amount"
        repoAmt.Name = colAmt
        repoAmt.Width = 80
        repoAmt.Minimum = 0
        repoAmt.ReadOnly = False
        repoAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAmt)



        Dim repoDisPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisPer = New GridViewDecimalColumn()
        repoDisPer.FormatString = ""
        repoDisPer.HeaderText = "Discount %"
        repoDisPer.Minimum = 0
        repoDisPer.Name = colDisPer
        repoDisPer.Width = 100
        repoDisPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDisPer)

        Dim repoDisAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisAmt = New GridViewDecimalColumn()
        repoDisAmt.FormatString = ""
        repoDisAmt.HeaderText = "Discount Amount"
        repoDisAmt.WrapText = True
        repoDisAmt.Name = colDisAmt
        repoDisAmt.Width = 80
        repoDisAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisAmt.VisibleInColumnChooser = False
        repoDisAmt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoDisAmt)


        Dim repoHDisPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoHDisPer = New GridViewDecimalColumn()
        repoHDisPer.FormatString = ""
        repoHDisPer.HeaderText = "Head Discount %"
        repoHDisPer.Minimum = 0
        repoHDisPer.Name = colHeaDDisPer
        repoHDisPer.Width = 100
        repoHDisPer.ReadOnly = True
        repoHDisPer.IsVisible = True
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
        repoHDisAmt.VisibleInColumnChooser = False
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
        gv1.MasterTemplate.Columns.Add(repoHeadDisAmt)

        Dim repoCustDiscountQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCustDiscountQty.HeaderText = "Cash Dis Qty."
        repoCustDiscountQty.MinWidth = 4
        repoCustDiscountQty.Name = ColCustDiscountQty
        repoCustDiscountQty.ReadOnly = True
        repoCustDiscountQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCustDiscountQty.Width = 54
        gv1.MasterTemplate.Columns.Add(repoCustDiscountQty)


        Dim repoCustDiscountPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCustDiscountPer.HeaderText = "Cash Dis %."
        repoCustDiscountPer.MinWidth = 4
        repoCustDiscountPer.Name = colCustDiscPercentage
        repoCustDiscountPer.ReadOnly = True
        repoCustDiscountPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCustDiscountPer.Width = 54
        gv1.MasterTemplate.Columns.Add(repoCustDiscountPer)

        Dim repoCustDiscount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCustDiscount.HeaderText = "Cash Dis Amt."
        repoCustDiscount.MinWidth = 4
        repoCustDiscount.Name = colcustDiscount
        repoCustDiscount.ReadOnly = True
        repoCustDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoCustDiscount.Width = 54
        gv1.MasterTemplate.Columns.Add(repoCustDiscount)

        Dim repoCashSchemeCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCashSchemeCode.HeaderText = "Cash Scheme Code"
        repoCashSchemeCode.Name = colCashDiscSchemeCode
        repoCashSchemeCode.Width = 80
        repoCashSchemeCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoCashSchemeCode)


        Dim repoAcualCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAcualCost.AllowSort = False
        repoAcualCost.HeaderText = "Net Price"
        repoAcualCost.Name = colActualCost
        repoAcualCost.ReadOnly = True
        repoAcualCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAcualCost.Width = 55
        gv1.MasterTemplate.Columns.Add(repoAcualCost)


        Dim repoTotalMRP As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalMRP.AllowSort = False
        repoTotalMRP.HeaderText = "Total MRP"
        repoTotalMRP.Name = colTotalMRP
        repoTotalMRP.ReadOnly = True
        repoTotalMRP.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalMRP.Width = 62
        gv1.MasterTemplate.Columns.Add(repoTotalMRP)

        Dim repoTotalBasicAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalBasicAmt.AllowSort = False
        repoTotalBasicAmt.HeaderText = "Total Basic Amount"
        repoTotalBasicAmt.Name = colTotalBasicAmount
        repoTotalBasicAmt.ReadOnly = True
        repoTotalBasicAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalBasicAmt.Width = 106
        gv1.MasterTemplate.Columns.Add(repoTotalBasicAmt)

        Dim repoTotalDiscount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalDiscount.AllowSort = False
        repoTotalDiscount.HeaderText = "Total Discount"
        repoTotalDiscount.Name = colTotalDiscountAmount
        repoTotalDiscount.ReadOnly = True
        repoTotalDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalDiscount.Width = 81
        gv1.MasterTemplate.Columns.Add(repoTotalDiscount)

        Dim repoTotalCustDiscount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotalCustDiscount.HeaderText = "Total Cash Dis"
        repoTotalCustDiscount.Name = colTotalCustDiscount
        repoTotalCustDiscount.ReadOnly = True
        repoTotalCustDiscount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTotalCustDiscount.Width = 79
        gv1.MasterTemplate.Columns.Add(repoTotalCustDiscount)

        Dim repoAmtAfterDis As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterDis = New GridViewDecimalColumn()
        repoAmtAfterDis.FormatString = ""
        repoAmtAfterDis.HeaderText = "Amount After Discount"
        repoAmtAfterDis.Name = colAmtAfterDis
        repoAmtAfterDis.WrapText = True
        repoAmtAfterDis.Width = 80
        repoAmtAfterDis.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterDis.VisibleInColumnChooser = False
        repoAmtAfterDis.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterDis)

        Dim repoPrincipleCOde As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPrincipleCOde.FormatString = ""
        repoPrincipleCOde.HeaderText = "Principle Code"
        repoPrincipleCOde.Name = colPricipleCode
        repoPrincipleCOde.Width = 150
        repoPrincipleCOde.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoPrincipleCOde)

        Dim repoPrincipleDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPrincipleDesc.FormatString = ""
        repoPrincipleDesc.HeaderText = "Principle Desc"
        repoPrincipleDesc.Name = colPricipleDesc
        repoPrincipleDesc.Width = 150
        repoPrincipleDesc.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoPrincipleDesc)

        Dim repoVCOde As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoVCOde.FormatString = ""
        repoVCOde.HeaderText = "Vendor Code"
        repoVCOde.Name = colvendorCode
        repoVCOde.Width = 150
        repoVCOde.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoVCOde)

        Dim repoVDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoVDesc.FormatString = ""
        repoVDesc.HeaderText = "Vendor Desc"
        repoVDesc.Name = colvendorDesc
        repoVDesc.Width = 150
        repoVDesc.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoVDesc)

        Dim repoMarkupPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMarkupPer = New GridViewDecimalColumn()
        repoMarkupPer.FormatString = ""
        repoMarkupPer.HeaderText = "Mark Up %"
        repoMarkupPer.Name = colMarkUpPercentage
        repoMarkupPer.WrapText = True
        repoMarkupPer.Width = 80
        repoMarkupPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoMarkupPer.VisibleInColumnChooser = False
        repoMarkupPer.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoMarkupPer)



        Dim repoLandingCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLandingCost = New GridViewDecimalColumn()
        repoLandingCost.FormatString = ""
        repoLandingCost.HeaderText = "Landing Cost"
        repoLandingCost.Name = colLandingCost
        repoLandingCost.WrapText = True
        repoLandingCost.Width = 80
        repoLandingCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoLandingCost.VisibleInColumnChooser = False
        repoLandingCost.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoLandingCost)

        Dim repoPurCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPurCost = New GridViewDecimalColumn()
        repoPurCost.FormatString = ""
        repoPurCost.HeaderText = "purchase Cost"
        repoPurCost.Name = colPurCost
        repoPurCost.WrapText = True
        repoPurCost.Width = 80
        repoPurCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoPurCost.VisibleInColumnChooser = False
        repoPurCost.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoPurCost)

        Dim repoMarkupOn As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoMarkupOn.FormatString = ""
        repoMarkupOn.HeaderText = "MarkUp On"
        repoMarkupOn.Name = colMarkupOn
        repoMarkupOn.ReadOnly = True
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
        repoTaxBaseAmt9.IsVisible = False
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
        repoTotTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTotTaxAmt)

        Dim repoAmtAfterTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterTax.FormatString = ""
        repoAmtAfterTax.HeaderText = "Included Tax Amount"
        repoAmtAfterTax.Name = colAmtAfterTax
        repoAmtAfterTax.WrapText = True
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
        repoRequition.HeaderText = "Shipment No"
        repoRequition.Name = colOrderNo
        repoRequition.ReadOnly = True
        repoRequition.Width = 100
        gv1.MasterTemplate.Columns.Add(repoRequition)


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
        repoBinNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoBinNo)

        Dim repoBatchNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBatchNo.FormatString = ""
        repoBatchNo.HeaderText = "Batch No"
        repoBatchNo.Name = colBatchNo
        repoBatchNo.ReadOnly = False
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
        repoManDate.Width = 80
        gv1.MasterTemplate.Columns.Add(repoManDate)

        Dim repoSpecification As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSpecification.FormatString = ""
        repoSpecification.HeaderText = "Specification"
        repoSpecification.Name = colSpecification
        repoSpecification.Width = 100
        gv1.MasterTemplate.Columns.Add(repoSpecification)


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



        Dim repoFOC As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoFOC.FormatString = ""
        repoFOC.HeaderText = "FOC"
        repoFOC.Name = ColFOC
        repoFOC.ReadOnly = True
        repoFOC.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoFOC)

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

        Dim repoDeliveryCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDeliveryCode.FormatString = ""
        repoDeliveryCode.HeaderText = "Delivery No"
        repoDeliveryCode.Name = colDelivery_Code
        repoDeliveryCode.ReadOnly = True
        repoDeliveryCode.Width = 100
        gv1.MasterTemplate.Columns.Add(repoDeliveryCode)

        Dim DC_PKID As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_PKID.FormatString = ""
        DC_PKID.HeaderText = "Distributor PKID"
        DC_PKID.Name = ColDCPKID
        DC_PKID.Width = 100
        DC_PKID.ReadOnly = True
        DC_PKID.IsVisible = True
        DC_PKID.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_PKID)
        Dim DC_ApplicableDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_ApplicableDate.FormatString = ""
        DC_ApplicableDate.HeaderText = "Applicable Date"
        DC_ApplicableDate.Name = ColDCApplicableDate
        DC_ApplicableDate.Width = 100
        DC_ApplicableDate.ReadOnly = True
        DC_ApplicableDate.IsVisible = True
        DC_ApplicableDate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_ApplicableDate)
        Dim DC_UOM As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_UOM.FormatString = ""
        DC_UOM.HeaderText = "Commission UOM"
        DC_UOM.Name = ColDCUOM
        DC_UOM.Width = 100
        DC_UOM.ReadOnly = True
        DC_UOM.IsVisible = True
        DC_UOM.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_UOM)
        Dim DC_Rate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_Rate.FormatString = ""
        DC_Rate.HeaderText = "Commission Rate"
        DC_Rate.Name = ColDCRate
        DC_Rate.Width = 100
        DC_Rate.ReadOnly = True
        DC_Rate.IsVisible = True
        DC_Rate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_Rate)
        Dim DC_RateWithTax As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_RateWithTax.FormatString = ""
        DC_RateWithTax.HeaderText = "Commission Rate Wtih TAx"
        DC_RateWithTax.Name = ColDCRateWithTax
        DC_RateWithTax.Width = 100
        DC_RateWithTax.ReadOnly = True
        DC_RateWithTax.IsVisible = True
        DC_RateWithTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_RateWithTax)
        Dim DC_UnitCF As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_UnitCF.FormatString = ""
        DC_UnitCF.HeaderText = "Conversion Factor"
        DC_UnitCF.Name = ColDCUnitCF
        DC_UnitCF.Width = 100
        DC_UnitCF.ReadOnly = True
        DC_UnitCF.IsVisible = True
        DC_UnitCF.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_UnitCF)
        Dim DC_QtyInSU As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_QtyInSU.FormatString = ""
        DC_QtyInSU.HeaderText = "Qty in Stocking Unit"
        DC_QtyInSU.Name = ColDCQtyinSU
        DC_QtyInSU.Width = 100
        DC_QtyInSU.ReadOnly = True
        DC_QtyInSU.IsVisible = True
        DC_QtyInSU.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_QtyInSU)
        Dim DC_CFUOM As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_CFUOM.FormatString = ""
        DC_CFUOM.HeaderText = "Commission Conversion Factor"
        DC_CFUOM.Name = ColDCCFUOM
        DC_CFUOM.Width = 100
        DC_CFUOM.ReadOnly = True
        DC_CFUOM.IsVisible = True
        DC_CFUOM.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_CFUOM)
        Dim DC_Amt As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_Amt.FormatString = ""
        DC_Amt.HeaderText = "Distributor Commission Amt"
        DC_Amt.Name = ColDCAmt
        DC_Amt.Width = 100
        DC_Amt.ReadOnly = True
        DC_Amt.IsVisible = True
        DC_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_Amt)
        Dim TransporterRate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        TransporterRate.FormatString = ""
        TransporterRate.HeaderText = "Transporter Commission Rate"
        TransporterRate.Name = ColTCRate
        TransporterRate.Width = 100
        TransporterRate.ReadOnly = True
        TransporterRate.IsVisible = False
        TransporterRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(TransporterRate)
        Dim TransporterAmt As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        TransporterAmt.FormatString = ""
        TransporterAmt.HeaderText = "Transporter Commission Amt"
        TransporterAmt.Name = ColTCAmt
        TransporterAmt.Width = 100
        TransporterAmt.ReadOnly = True
        TransporterAmt.IsVisible = False
        TransporterAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(TransporterAmt)
        Dim SC_Rate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        SC_Rate.FormatString = ""
        SC_Rate.HeaderText = "Security Rate"
        SC_Rate.Name = ColSCRate
        SC_Rate.Width = 100
        SC_Rate.ReadOnly = True
        SC_Rate.IsVisible = True
        SC_Rate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(SC_Rate)
        Dim SC_Amt As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        SC_Amt.FormatString = ""
        SC_Amt.HeaderText = "Security Amt"
        SC_Amt.Name = ColSCAmt
        SC_Amt.Width = 100
        SC_Amt.ReadOnly = True
        SC_Amt.IsVisible = True
        SC_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(SC_Amt)
        Dim repoCF As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCF.FormatString = ""
        repoCF.HeaderText = "Convertion Factor"
        repoCF.Name = colCF
        repoCF.Width = 100
        repoCF.ReadOnly = True
        repoCF.IsVisible = False
        repoCF.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoCF)
        Dim repoCFinKg As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCFinKg.FormatString = ""
        repoCFinKg.HeaderText = "Convertion Factor in KG"
        repoCFinKg.Name = colCFKG
        repoCFinKg.Width = 100
        repoCFinKg.ReadOnly = True
        repoCFinKg.IsVisible = False
        repoCFinKg.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoCFinKg)
        Dim repoQtyinKg As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoQtyinKg.FormatString = ""
        repoQtyinKg.HeaderText = "Total Qty in KG"
        repoQtyinKg.Name = colQtyinKG
        repoQtyinKg.Width = 100
        repoQtyinKg.ReadOnly = True
        repoQtyinKg.IsVisible = False
        repoQtyinKg.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoQtyinKg)
        'clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields, trans)
        Dim repoTransporter As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTransporter.FormatString = ""
        repoTransporter.HeaderText = "Transporter"
        repoTransporter.Name = colTransporter
        repoTransporter.Width = 100
        repoTransporter.ReadOnly = True
        repoTransporter.IsVisible = False
        repoTransporter.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTransporter)

        clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)

        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = True
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40
        gv1.BestFitColumns()
        ReStoreGridLayout()
    End Sub
    Public Sub GetDCDetails(ByVal intRow As Integer)
        Dim docDate As DateTime = Nothing
        If clsCommon.myLen(txtReqNo.Value) > 0 Then
            docDate = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select Document_Date from TSPL_SD_SHIPMENT_HEAD where Document_Code='" & txtReqNo.Value & "'"))
        End If
        Dim DCQry As String = "select top 1 TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Commision_UOM,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Rate,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Security_Rate from TSPL_DISTRIBUTOR_COMMISSION_HEAD
left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date<='" & IIf(clsCommon.myLen(txtReqNo.Value) > 0, clsCommon.GetPrintDate(docDate), clsCommon.GetPrintDate(txtDate.Value)) & "' and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code='" + clsCommon.myCstr(txtVendorNo.Value) + "' and TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code='" + clsCommon.myCstr(gv1.Rows(intRow).Cells(colICode).Value) + "' and TSPL_DISTRIBUTOR_COMMISSION_HEAD.IsPosted=1 and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code='" + clsCommon.myCstr(txtRouteNo.Value) + "' 
 order by TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date desc,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No desc"
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(DCQry)
        If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
            gv1.Rows(intRow).Cells(ColDCPKID).Value = clsCommon.myCstr(dt1.Rows(0)("PK_ID"))
            gv1.Rows(intRow).Cells(ColDCApplicableDate).Value = clsCommon.myCstr(dt1.Rows(0)("Applicable_Date"))
            gv1.Rows(intRow).Cells(ColDCUOM).Value = clsCommon.myCstr(dt1.Rows(0)("Commision_UOM"))
            gv1.Rows(intRow).Cells(ColDCRate).Value = clsCommon.myCstr(dt1.Rows(0)("Rate"))
            gv1.Rows(intRow).Cells(ColTCRate).Value = clsCommon.myCstr(dt1.Rows(0)("Transporter_Rate"))
            gv1.Rows(intRow).Cells(ColSCRate).Value = clsCommon.myCstr(dt1.Rows(0)("Security_Rate"))
            'gv1.CurrentRow.Cells(ColDCRateWithTax).Value = Math.Round(gv1.Rows(IntRowNo).Cells(ColDCRate).Value * 100 / (100 + dblTotTaxRate), 2)
            gv1.Rows(intRow).Cells(ColDCUnitCF).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" + clsCommon.myCstr(gv1.Rows(intRow).Cells(colActualUOM).Value) + "' and Item_Code='" + clsCommon.myCstr(gv1.Rows(intRow).Cells(colICode).Value) + "'")
            gv1.Rows(intRow).Cells(ColDCCFUOM).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" + clsCommon.myCstr(gv1.Rows(intRow).Cells(ColDCUOM).Value) + "' and Item_Code='" + clsCommon.myCstr(gv1.Rows(intRow).Cells(colICode).Value) + "'")
            'gv1.CurrentRow.Cells(ColDCQtyinSU).Value = (gv1.Rows(IntRowNo).Cells(colQty).Value * gv1.Rows(IntRowNo).Cells(ColDCUnitCF).Value) / gv1.Rows(IntRowNo).Cells(ColDCCFUOM).Value
            'gv1.CurrentRow.Cells(ColDCAmt).Value = gv1.CurrentRow.Cells(ColDCQtyinSU).Value * gv1.CurrentRow.Cells(ColDCRateWithTax).Value
            'dblTotalDCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCAmt).Value)
        End If
    End Sub
    'Sub LoadBlankGridAC()
    '    gvAC.Rows.Clear()
    '    gvAC.Columns.Clear()

    '    Dim repoACCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
    '    repoACCode.FormatString = ""
    '    repoACCode.HeaderText = "Addition Charges Code"
    '    repoACCode.Name = colACCode
    '    repoACCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
    '    repoACCode.TextImageRelation = TextImageRelation.TextBeforeImage
    '    repoACCode.Width = 150
    '    repoACCode.ReadOnly = False
    '    gvAC.MasterTemplate.Columns.Add(repoACCode)

    '    Dim repoACName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
    '    repoACName.FormatString = ""
    '    repoACName.HeaderText = "Addition Charges Description"
    '    repoACName.Name = colACName
    '    repoACName.Width = 300
    '    repoACName.ReadOnly = True
    '    gvAC.MasterTemplate.Columns.Add(repoACName)

    '    Dim repoACAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
    '    repoACAmt.FormatString = ""
    '    repoACAmt.HeaderText = "Amount"
    '    repoACAmt.Name = colACAmount
    '    repoACAmt.Width = 100
    '    repoACAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
    '    repoACAmt.ReadOnly = False
    '    gvAC.MasterTemplate.Columns.Add(repoACAmt)

    '    gvAC.AllowAddNewRow = False
    '    gvAC.ShowGroupPanel = False
    '    gvAC.AllowColumnReorder = True
    '    gvAC.AllowRowReorder = False
    '    gvAC.EnableSorting = False
    '    gvAC.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
    '    gvAC.MasterTemplate.ShowRowHeaderColumn = False
    '    gvAC.TableElement.TableHeaderHeight = 40
    'End Sub

    'Sub LoadBlankGridTax()
    '    gv2.Rows.Clear()
    '    gv2.Columns.Clear()

    '    Dim repoTaxAuthCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
    '    repoTaxAuthCode.FormatString = ""
    '    repoTaxAuthCode.HeaderText = "Tax Authority Code"
    '    repoTaxAuthCode.Name = colTTaxAutCode
    '    repoTaxAuthCode.Width = 150
    '    repoTaxAuthCode.ReadOnly = True
    '    gv2.MasterTemplate.Columns.Add(repoTaxAuthCode)

    '    Dim repoTaxAuthName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
    '    repoTaxAuthName.FormatString = ""
    '    repoTaxAuthName.HeaderText = "Tax Authority"
    '    repoTaxAuthName.Name = colTTaxAutName
    '    repoTaxAuthName.Width = 200
    '    repoTaxAuthName.ReadOnly = True
    '    gv2.MasterTemplate.Columns.Add(repoTaxAuthName)

    '    Dim repoTaxBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
    '    repoTaxBaseAmt.FormatString = ""
    '    repoTaxBaseAmt.HeaderText = "Base Amount"
    '    repoTaxBaseAmt.Name = colTBaseAmt
    '    repoTaxBaseAmt.Width = 100
    '    repoTaxBaseAmt.ReadOnly = True
    '    repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
    '    gv2.MasterTemplate.Columns.Add(repoTaxBaseAmt)

    '    Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
    '    repoTaxRate.FormatString = ""
    '    repoTaxRate.HeaderText = "Tax Rate"
    '    repoTaxRate.Name = colTTaxRate
    '    repoTaxRate.Width = 100
    '    repoTaxRate.ReadOnly = True
    '    repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
    '    gv2.MasterTemplate.Columns.Add(repoTaxRate)

    '    Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
    '    repoTaxAmt.FormatString = ""
    '    repoTaxAmt.HeaderText = "Tax Amount"
    '    repoTaxAmt.Name = colTTaxAmt
    '    repoTaxAmt.Width = 100
    '    repoTaxAmt.ReadOnly = False
    '    repoTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
    '    gv2.MasterTemplate.Columns.Add(repoTaxAmt)

    '    clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)

    '    gv1.AllowAddNewRow = False
    '    gv1.ShowGroupPanel = False
    '    gv1.AllowColumnReorder = True
    '    gv1.AllowRowReorder = False
    '    gv1.EnableSorting = False
    '    gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
    '    gv1.MasterTemplate.ShowRowHeaderColumn = False
    '    gv1.TableElement.TableHeaderHeight = 40

    '    ReStoreGridLayout()
    'End Sub

    Sub OpenSerialItem()
        If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsSerialseItem).Value) Then
            Dim frm As frmSerializeItemOut = New frmSerializeItemOut()
            frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
            frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
            frm.strLocationCode = txtBillToLocation.Value
            frm.strCurrDocNo = txtDocNo.Value
            frm.strCurrDocType = "Sale Return"
            frm.strAgaintsDocNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colOrderNo).Value)
            frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
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
        repoACCode.HeaderImage = My.Resources.search4
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
        repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxBaseAmt)

        Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate.FormatString = ""
        repoTaxRate.HeaderText = "Tax Rate"
        repoTaxRate.Name = colTTaxRate
        repoTaxRate.Width = 100
        repoTaxRate.ReadOnly = True
        repoTaxRate.IsVisible = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, Nothing)) = 1, True, False)
        repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxRate)

        Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt.FormatString = ""
        repoTaxAmt.HeaderText = "Tax Amount"
        repoTaxAmt.Name = colTTaxAmt
        repoTaxAmt.Width = 100
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
    Public Sub ItemPrice(ByVal strItem As String, ByVal strUnit As String, ByVal intQty As Decimal, ByVal introw As Integer, ByVal isFORPrice As Boolean)
        Try
            Dim dt As New DataTable()
            Dim dblRate As Double = 0
            Dim dblTotal As Double = 0
            Dim dblItemBasicPrice As Double = 0
            Dim whrcls As String = ""
            Dim qry As String = ""
            Dim Price_code As String = ""

            Price_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT price_CodeNon FROM tspl_customer_master where cust_code='" + txtVendorNo.Value + "'"))
            whrcls = " and TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "'"


            qry = " Select Is_With_Tax, RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price,XXXE.TAX1_Rate, " &
            " XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
            "  XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, " &
            " XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, " &
            " XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7, " &
            " XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.TAX1_Amt, " &
            " XXXE.TAX2_Amt,XXXE.TAX3_Amt,XXXE.TAX4_Amt,XXXE.Against_Plan_TR_Code  from ( " &
            "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
            "Start_Date Desc) as RowNo,Is_With_Tax, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
            "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,  " &
            "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
            " TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
            " TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
            " TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7, " &
            " TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.TAX1_Amt , TSPL_ITEM_PRICE_MASTER.TAX2_Amt ,TSPL_ITEM_PRICE_MASTER.TAX3_Amt ,TSPL_ITEM_PRICE_MASTER.TAX4_Amt,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
            "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
            "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  2=( case when CONVERT(date,Start_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,Start_Date,103)<='" + IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "' then 2 else 3 end)  end)  and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null)  " & whrcls & "  " &
            " and UOM='" & strUnit & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & strItem & "'  AND Location_Code='" & clsCommon.myCstr(txtBillToLocation.Value) & "'  " &
            ") XXXE WHERE RowNo=1  "
            dt = clsDBFuncationality.GetDataTable(qry)
            If dt.Rows.Count > 0 Then
                gv1.Rows(introw).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                gv1.Rows(introw).Cells(colMRP).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Net"))
                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0).Item("Is_With_Tax")), "N") = CompairStringResult.Equal Then
                    dblItemBasicPrice = Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price")) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Amt")), 2), 2)
                Else
                    dblItemBasicPrice = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price"))
                End If
            Else
                Throw New Exception("Please Create Price Chart For  " + Environment.NewLine +
                                                     "Location    --  " & clsCommon.myCstr(txtBillToLocation.Value) & " " + Environment.NewLine +
                                                     "Price Code --   " & clsCommon.myCstr(txtPriceCode.Text) & " " + Environment.NewLine +
                                                     "State Date < =  " & clsCommon.myCstr(txtDate.Value) & " " + Environment.NewLine +
                                                     "Item         --   " & clsCommon.myCstr(gv1.Rows(introw).Cells(colICode).Value) & " " + Environment.NewLine +
                                                     "Unit         --   " & clsCommon.myCstr(strUnit) & ".")
            End If
            gv1.Rows(introw).Cells(colItemBasicPrice).Value = dblItemBasicPrice
            'gv1.Rows(introw).Cells().Value = dblItemBasicPrice * intQty
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                UpdateCurrentRow1(introw)

            Else
                UpdateCurrentRow(introw)

            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

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
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                            UpdateCurrentRow1(gv1.CurrentRow.Index)
                        Else
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                        End If
                        'UpdateCurrentRow(gv1.CurrentRow.Index) ''-1 is for current row
                        UpdateAllTotals()
                    End If
                    If e.Column Is gv1.Columns(colICode) OrElse e.Column Is gv1.Columns(colHeadDiscamt) OrElse e.Column Is gv1.Columns(colHeaDDisPer) OrElse (e.Column Is gv1.Columns(colHeadDiscamt)) OrElse e.Column Is gv1.Columns(colAmt) OrElse e.Column Is gv1.Columns(colQty) OrElse e.Column Is gv1.Columns(colDamageQty) OrElse e.Column Is gv1.Columns(colRate) OrElse e.Column Is gv1.Columns(colSpecification) OrElse e.Column Is gv1.Columns(colRemarks) OrElse e.Column Is gv1.Columns(colDisPer) OrElse e.Column Is gv1.Columns(colMRP) OrElse e.Column Is gv1.Columns(colBatchNo) OrElse e.Column Is gv1.Columns(colExpiry) OrElse e.Column Is gv1.Columns(colManufactureDate) OrElse e.Column Is gv1.Columns(colUnit) OrElse e.Column Is gv1.Columns(colActualQty) OrElse e.Column Is gv1.Columns(colActualUOM) OrElse (e.Column Is gv1.Columns(colAmt) AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal) Then
                        If (e.Column Is gv1.Columns(colQty) OrElse e.Column Is gv1.Columns(colHeadDiscamt) OrElse e.Column Is gv1.Columns(colHeaDDisPer) OrElse (e.Column Is gv1.Columns(colHeadDiscamt)) OrElse e.Column Is gv1.Columns(colDamageQty) OrElse e.Column Is gv1.Columns(colRate) OrElse e.Column Is gv1.Columns(colDisPer) OrElse (e.Column Is gv1.Columns(colAmt) OrElse e.Column Is gv1.Columns(colActualQty) AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal)) Then
                            If ((e.Column Is gv1.Columns(colQty))) Then
                                Dim dblPendingQty As Double = 0
                                If (clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) > 0) Then
                                    dblPendingQty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colPendingQty).Value)

                                End If

                                If (clsCommon.myLen(gv1.CurrentRow.Cells(colOrderNo).Value) > 0) Then
                                    Dim dblEnteredQty As Double = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
                                    Dim dblDamageQty As Double = 0 'clsCommon.myCdbl(gv1.CurrentRow.Cells(colLeakQty).Value) + clsCommon.myCdbl(gv1.CurrentRow.Cells(colBurstQty).Value) + clsCommon.myCdbl(gv1.CurrentRow.Cells(colShortQty).Value)
                                    If (dblEnteredQty + dblDamageQty) > dblPendingQty Then
                                        'common.clsCommon.MyMessageBoxShow("Entered Quantity Can't be more than Pending Quantity." + Environment.NewLine + "Entered Quantity : " + clsCommon.myCstr(dblEnteredQty) + ". Pending Quantity : " + clsCommon.myCstr(dblPendingQty) + ". Damage Quantity : " + clsCommon.myCstr(dblDamageQty))
                                        'gv1.CurrentCell.Value = 0
                                    End If
                                End If
                                OpenSerialItem()
                                OpenBatchItem()
                            End If
                            'If clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value) > 0 Then
                            findQtyandPromoSchemeCode(False, gv1.CurrentRow.Cells(colInvoiceSchemeCode).Value)
                            findVolumeStructureSchemeCode(False, "")
                            'End If
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                UpdateCurrentRow1(gv1.CurrentRow.Index)
                            Else
                                UpdateCurrentRow(gv1.CurrentRow.Index)
                            End If
                            'UpdateCurrentRow(gv1.CurrentRow.Index)
                            If rbtnTaxCalManual.IsChecked Then
                                For ii As Integer = 0 To gv1.Rows.Count - 1
                                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                        UpdateCurrentRow1(ii)
                                    Else
                                        UpdateCurrentRow(ii)
                                    End If
                                    'UpdateCurrentRow(ii)
                                Next
                            End If
                            UpdateAllTotals()


                        ElseIf e.Column Is gv1.Columns(colICode) Then
                            OpenItemList(False)
                        ElseIf e.Column Is gv1.Columns(colActualUOM) Then
                            Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            If clsCommon.myLen(strICode) > 0 Then
                                OpenUOMList(False)
                            Else
                                Throw New Exception("Please fill item first.")
                                isCellValueChangedOpen = False
                                Exit Sub
                            End If
                            Dim strActualUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colActualUOM).Value)
                            ItemPrice(strICode, strActualUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value), gv1.CurrentRow.Index, False)
                        ElseIf e.Column Is gv1.Columns(colActualQty) Then
                            Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            Dim strActualUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colActualUOM).Value)
                            '' Dim dblActualqty As Decimal = clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value)
                            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colActualUOM).Value)) = CompairStringResult.Equal Then
                                gv1.CurrentRow.Cells(colQty).Value = clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value)
                            Else
                                gv1.CurrentRow.Cells(colQty).Value = ActualConversionInReturn(strICode, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), strActualUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value))
                            End If
                            ItemPrice(strICode, strActualUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value), gv1.CurrentRow.Index, False)
                            UpdateAllTotals()
                        ElseIf e.Column Is gv1.Columns(colUnit) Then
                            OpenUOMList(False)

                        ElseIf UOMAtDiarySaleReturn = True Then
                            ''richa agarwal GKD/28/06/18-000153 calculate functions not called at actual uom so now it is called thats why problem occurs in calculation BHA/17/08/18-000446
                            If e.Column Is gv1.Columns(colActualUOM) Then
                                OpenUOMList(False)
                                CalAltQty(gv1.CurrentRow.Index)
                                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                    UpdateCurrentRow1(gv1.CurrentRow.Index)
                                Else
                                    UpdateCurrentRow(gv1.CurrentRow.Index)
                                End If
                                'UpdateCurrentRow(gv1.CurrentRow.Index)
                                UpdateAllTotals()
                            ElseIf e.Column Is gv1.Columns(colActualQty) Then
                                CalAltQty(gv1.CurrentRow.Index)
                                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                    UpdateCurrentRow1(gv1.CurrentRow.Index)
                                Else
                                    UpdateCurrentRow(gv1.CurrentRow.Index)
                                End If
                                'UpdateCurrentRow(gv1.CurrentRow.Index)
                                UpdateAllTotals()
                                OpenBatchItem()
                                ''richa agarwal 8 June,2020
                                'If Not (clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchmCodeType).Value), "VolumeSlab") = CompairStringResult.Equal) Then
                                If Not ((clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchmCodeType).Value), "VolumeSlab") = CompairStringResult.Equal) Or (clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchmCodeType).Value), "Quantitive") = CompairStringResult.Equal)) Then
                                    Dim Inv_Date As Date = clsCommon.myCDate(clsDBFuncationality.getSingleValue("Select CONVERT(date, Document_Date,103) from TSPL_SD_SALE_INVOICE_HEAD where Document_Code ='" + txtReqNo.Value + "'"))
                                    findQtyandPromoSchemeCode(False, gv1.CurrentRow.Cells(colInvoiceSchemeCode).Value, Inv_Date)
                                End If
                                If Not clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal Then
                                    findVolumeStructureSchemeCode(False, "")
                                End If

                            End If
                            'GetDCDetails()

                            'ElseIf e.Column Is gv1.Columns(colMRP) Then
                            '    Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            '    Dim Rateqry As String = "select Item_Rate,MRP from TSPL_VENDOR_ITEM_DETAIL where Customer_Code='" + txtVendorNo.Value + "' and item_no='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) + "' and UOM='FC' and MRP='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colMRP).Value) + "' "
                            '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(Rateqry)
                            '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            '        Dim VendrItemRate As Double = clsCommon.myCdbl(dt.Rows(0)("Item_Rate"))
                            '        Dim conversionFact As Double = clsCommon.myCdbl(clsItemMaster.GetConvertionFactor(strICode, gv1.CurrentRow.Cells(colUnit).Value, Nothing))
                            '        If VendrItemRate <> 0 Then
                            '            Dim Itemrate As Double = clsCommon.myCdbl(clsCommon.myCdbl(VendrItemRate) / clsCommon.myCdbl(conversionFact))
                            '            gv1.CurrentRow.Cells(colRate).Value = Math.Round(Itemrate, 2)
                            '        End If
                            '        UpdateCurrentRow(gv1.CurrentRow.Index)
                            '        If rbtnTaxCalManual.IsChecked Then
                            '            For ii As Integer = 0 To gv1.Rows.Count - 1
                            '                UpdateCurrentRow(ii)
                            '            Next
                            '        End If
                            '        UpdateAllTotals()
                            '    End If
                        ElseIf e.Column Is gv1.Columns(colAmt) Then

                            ' ''If clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value) > 0 Then
                            ' ''    gv1.CurrentRow.Cells(colRate).Value = Math.Round(clsCommon.myCdbl(gv1.CurrentRow.Cells(colAmt).Value) / clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), 2, MidpointRounding.ToEven)
                            ' ''Else
                            ' ''    gv1.CurrentRow.Cells(colAmt).Value = 0
                            ' ''End If
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                UpdateCurrentRow1(gv1.CurrentRow.Index)
                            Else
                                UpdateCurrentRow(gv1.CurrentRow.Index)
                            End If
                            'UpdateCurrentRow(gv1.CurrentRow.Index)
                            If rbtnTaxCalManual.IsChecked Then
                                For ii As Integer = 0 To gv1.Rows.Count - 1
                                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                        UpdateCurrentRow1(ii)
                                    Else
                                        UpdateCurrentRow(ii)
                                    End If
                                    'UpdateCurrentRow(ii)
                                Next
                            End If
                            UpdateAllTotals()
                        End If
                        'setGridFocus()
                    End If
                    isCellValueChangedOpen = False
                End If
            End If
        Catch ex As Exception
            isCellValueChangedOpen = False
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function ActualConversionInReturn(ByVal itemCode As String, ByVal ReturnUOM As String, ByVal ActualUOM As String, ByVal ActualQty As Double) As Double
        Dim ActInReturn As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Convert(decimal(18,3),(" + clsCommon.myCstr(ActualQty) + "*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/CInFactor.Conversion_Factor)
from  TSPL_ITEM_MASTER 
Left Outer Join  TSPL_ITEM_UOM_DETAIL On TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
Left Outer Join (Select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL Where Item_Code='" + itemCode + "' And UOM_Code='" + ReturnUOM + "') CInFactor On CInFactor.Item_Code=TSPL_ITEM_MASTER.Item_Code
Where TSPL_ITEM_MASTER.Item_Code='" + itemCode + "' And TSPL_ITEM_UOM_DETAIL.UOM_Code='" + ActualUOM + "' "))
        Return ActInReturn
    End Function

    Sub OpenUOMList(ByVal isButtonClick As Boolean)
        Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        If clsCommon.myLen(strICode) > 0 Then
            Dim qry As String = "select UOM_Code as Code,UOM_Description as [Description] from TSPL_ITEM_UOM_DETAIL"
            Dim whrCls As String = "Item_Code='" + strICode + "'"
            gv1.CurrentRow.Cells(colActualUOM).Value = clsCommon.ShowSelectForm("PSSaleReturnActUnit", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colActualUOM).Value), "Code", isButtonClick)

            Dim dblConvF As Double
            dblConvF = GetConvFactor(gv1.CurrentRow.Cells(colActualUOM).Value, gv1.CurrentRow.Cells(colICode).Value)
            gv1.CurrentRow.Cells(colConvF).Value = dblConvF

            UpdateCurrentRow(gv1.CurrentRow.Index) ''-1 is for current row
            UpdateAllTotals()
        End If
    End Sub



    Private Sub setGridFocus()
        Try
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
            If intCurrRow = gv1.Rows.Count - 1 Then
                gv1.Rows.AddNew()
                gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                gv1.CurrentRow = gv1.Rows(intCurrRow)
            End If


            'If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
            '    If gv1.CurrentColumn Is gv1.Columns(colICode) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow)
            '        gv1.CurrentColumn = gv1.Columns(colQty)
            '    ElseIf gv1.CurrentColumn Is gv1.Columns(colQty) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow)
            '        gv1.CurrentColumn = gv1.Columns(colLeakQty)
            '    ElseIf gv1.CurrentColumn Is gv1.Columns(colLeakQty) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow)
            '        gv1.CurrentColumn = gv1.Columns(colBurstQty)
            '    ElseIf gv1.CurrentColumn Is gv1.Columns(colBurstQty) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow)
            '        gv1.CurrentColumn = gv1.Columns(colShortQty)
            '    ElseIf gv1.CurrentColumn Is gv1.Columns(colShortQty) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow)
            '        gv1.CurrentColumn = gv1.Columns(colRate)
            '    ElseIf gv1.CurrentColumn Is gv1.Columns(colRate) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow)
            '        gv1.CurrentColumn = gv1.Columns(colDisPer)
            '    ElseIf gv1.CurrentColumn Is gv1.Columns(colDisPer) Then
            '        gv1.CurrentRow = gv1.Rows(intCurrRow + 1)
            '        gv1.CurrentColumn = gv1.Columns(colICode)
            '    End If
            'End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub setBoothGridFocus()
    '    Try
    '        Dim intCurrRow As Integer = gvBooth.CurrentRow.Index
    '        'gvBooth.CurrentRow.Cells(colBoothCode).Value = clsCommon.myCdbl(intCurrRow + 1)
    '        If intCurrRow = gvBooth.Rows.Count - 1 Then
    '            gvBooth.Rows.AddNew()
    '            'gvBooth.CurrentRow = gvBooth.Rows(intCurrRow +)
    '        End If

    '    Catch ex As Exception
    '        common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Sub OpenICodeListCurrentCalaculation(ByVal isButtonClick As Boolean)
        gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
        Dim strItemType As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value)
        If clsCommon.myLen(strItemType) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Please select Row Type", Me.Text)
            Exit Sub
        End If

        '------------07/08/2014--------pivot category---------------------------
        Dim pivotheader As String = ""
        pivotheader = clsItemMaster.GetNLevelPivotHeader()
        '---------------------------------------------------------------------

        If clsCommon.CompairString(strItemType, RowTypeItem) = CompairStringResult.Equal Then
            Dim qry As String = ""
            If clsCommon.myLen(pivotheader) <= 0 Then
                If PurchaseOneItemOneVendor = True Then
                    qry = "select TSPL_CUSTOMER_ITEM_DETAIL.item_no as Item,TSPL_CUSTOMER_ITEM_DETAIL.item_desc as [ItemDesc],TSPL_CUSTOMER_ITEM_DETAIL.uom as Unit , " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.item_rate as BasicRate,TSPL_CUSTOMER_ITEM_DETAIL.Item_MRP as MRP, " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.Start_Date as [Start Date],vendor_code as [Vendor code],vendor_desc as [Vendor Desc]  , " &
                    "TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE as PrincipleCode,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as PrincipleDesc, " &
                    "Weight_Value as [Weight Value] from TSPL_CUSTOMER_ITEM_DETAIL left outer join TSPL_ITEM_MASTER on " &
                    "TSPL_ITEM_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_DETAIL.item_no left outer join " &
                    "TSPL_VENDOR_ITEM_DETAIL on TSPL_CUSTOMER_ITEM_DETAIL.item_no=TSPL_VENDOR_ITEM_DETAIL.item_no  left  outer join " &
                    "TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code=TSPL_ITEM_MASTER.Item_Code and " &
                    "TSPL_ITEM_MASTER_CATEGORY.SNO=1 left outer join " &
                    "TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " &
                    "where Customer_Code='" & txtVendorNo.Value & "' and TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI') "
                Else

                    qry = "select TSPL_CUSTOMER_ITEM_DETAIL.item_no as Item,TSPL_CUSTOMER_ITEM_DETAIL.item_desc as [ItemDesc],TSPL_CUSTOMER_ITEM_DETAIL.uom as Unit , " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.item_rate as BasicRate,TSPL_CUSTOMER_ITEM_DETAIL.Item_MRP as MRP, " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.Start_Date as [Start Date], " &
                    "Weight_Value as [Weight Value] from TSPL_CUSTOMER_ITEM_DETAIL left outer join TSPL_ITEM_MASTER on " &
                    "TSPL_ITEM_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_DETAIL.item_no  " &
                    "where Customer_Code='" & txtVendorNo.Value & "' and TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI') "
                End If
            End If

            '-------------------------cate. in columns---------------------------------
            If clsCommon.myLen(pivotheader) > 0 Then
                If PurchaseOneItemOneVendor = True Then
                    qry = "select * from (select a.DESCRIPTION,a.cat_value, TSPL_CUSTOMER_ITEM_DETAIL.item_no as Item,TSPL_CUSTOMER_ITEM_DETAIL.item_desc as [ItemDesc],TSPL_CUSTOMER_ITEM_DETAIL.uom as Unit , " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.item_rate as BasicRate,TSPL_CUSTOMER_ITEM_DETAIL.Item_MRP as MRP, " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.Start_Date as [Start Date],vendor_code as [Vendor code],vendor_desc as [Vendor Desc]  , " &
                    "TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE as PrincipleCode,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as PrincipleDesc, " &
                    "Weight_Value as [Weight Value] from TSPL_CUSTOMER_ITEM_DETAIL left outer join TSPL_ITEM_MASTER on " &
                    "TSPL_ITEM_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_DETAIL.item_no left outer join " &
                    "TSPL_VENDOR_ITEM_DETAIL on TSPL_CUSTOMER_ITEM_DETAIL.item_no=TSPL_VENDOR_ITEM_DETAIL.item_no  left  outer join " &
                    "TSPL_ITEM_MASTER_CATEGORY on TSPL_ITEM_MASTER_CATEGORY.Item_code=TSPL_ITEM_MASTER.Item_Code and " &
                    "TSPL_ITEM_MASTER_CATEGORY.SNO=1 left outer join " &
                    "TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values " &
                    " left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code,TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a on a.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_CUSTOMER_ITEM_DETAIL.item_no=a.item_code " &
                    "where Customer_Code='" & txtVendorNo.Value & "' and TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI')) as s pivot(max(cat_value) for description in (" + pivotheader + "))t  "
                Else
                    qry = "select * from (select a.DESCRIPTION,a.cat_value,TSPL_CUSTOMER_ITEM_DETAIL.item_no as Item,TSPL_CUSTOMER_ITEM_DETAIL.item_desc as [ItemDesc],TSPL_CUSTOMER_ITEM_DETAIL.uom as Unit , " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.item_rate as BasicRate,TSPL_CUSTOMER_ITEM_DETAIL.Item_MRP as MRP, " &
                    "TSPL_CUSTOMER_ITEM_DETAIL.Start_Date as [Start Date], " &
                    "Weight_Value as [Weight Value] from TSPL_CUSTOMER_ITEM_DETAIL left outer join TSPL_ITEM_MASTER on " &
                    "TSPL_ITEM_MASTER.Item_Code=TSPL_CUSTOMER_ITEM_DETAIL.item_no  " &
                    " left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code,TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a on a.Item_code=TSPL_ITEM_MASTER.Item_Code and TSPL_CUSTOMER_ITEM_DETAIL.item_no=a.item_code " &
                    "where Customer_Code='" & txtVendorNo.Value & "' and TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI')) as s pivot(max(cat_value) for description in (" + pivotheader + "))t  "
                End If
            End If
            '-----------------------------------------------------------------

            Dim dr As DataRow = clsCommon.ShowSelectFormForRow("SRD@uyit5", qry)
            If Not dr Is Nothing Then
                If clsCommon.myLen(clsCommon.myCstr(dr("Item"))) > 0 Then
                    SetitemWiseTaxSetting(True, True)
                    gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
                    gv1.CurrentRow.Cells(colICode).Value = clsCommon.myCstr(dr("Item"))
                    gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), IIf(clsCommon.myLen(clsCommon.myCstr(txtSubLocation.Value)) > 0, txtSubLocation.Value, txtBillToLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dr("Unit")), clsCommon.myCdbl(dr("MRP")))

                    If PurchaseOneItemOneVendor = True Then
                        gv1.CurrentRow.Cells(colPricipleCode).Value = clsCommon.myCstr(dr("PrincipleCode"))
                        gv1.CurrentRow.Cells(colPricipleDesc).Value = clsCommon.myCstr(dr("PrincipleDesc"))
                        gv1.CurrentRow.Cells(colvendorCode).Value = clsCommon.myCstr(dr("Vendor code"))
                        gv1.CurrentRow.Cells(colvendorDesc).Value = clsCommon.myCstr(dr("Vendor Desc"))
                    End If
                    gv1.CurrentRow.Cells(colIName).Value = clsCommon.myCstr(dr("ItemDesc"))
                    gv1.CurrentRow.Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), Nothing)
                    gv1.CurrentRow.Cells(colPriceCOde).Value = ""
                    gv1.CurrentRow.Cells(colPriceDateColumn).Value = clsCommon.myCstr(dr("Start Date"))
                    gv1.CurrentRow.Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                    gv1.CurrentRow.Cells(colOrgUnit).Value = clsCommon.myCstr(dr("Unit"))

                    gv1.CurrentRow.Cells(colMRP).Value = clsCommon.myCdbl(dr("MRP"))
                    gv1.CurrentRow.Cells(colRate).Value = clsCommon.myCdbl(dr("BasicRate"))
                    'gv1.CurrentRow.Cells(colAbatementPer).Value = clsCommon.myCdbl(dr("abatement_rate"))
                    gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                    gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                    gv1.CurrentRow.Cells(colActualCost).Value = clsCommon.myCdbl(dr("BasicRate"))
                    'gv1.CurrentRow.Cells(colTaxRate1).Value = 
                    'gv1.CurrentRow.Cells(colTaxRate2).Value = clsCommon.myCdbl(dr("Tax2Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate3).Value = clsCommon.myCdbl(dr("Tax3Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate4).Value = clsCommon.myCdbl(dr("Tax4Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate5).Value = clsCommon.myCdbl(dr("Tax5Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate6).Value = clsCommon.myCdbl(dr("Tax6Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate7).Value = clsCommon.myCdbl(dr("Tax7Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate8).Value = clsCommon.myCdbl(dr("Tax8Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate9).Value = clsCommon.myCdbl(dr("Tax9Rate"))
                    'gv1.CurrentRow.Cells(colTaxRate10).Value = clsCommon.myCdbl(dr("Tax10Rate"))

                    'gv1.CurrentRow.Cells(colMarkupOn).Value = clsCommon.myCstr(dr("markup_on"))
                    'gv1.CurrentRow.Cells(colMarkUpPercentage).Value = clsCommon.myCdbl(dr("markup_percent"))
                    'gv1.CurrentRow.Cells(colLandingCost).Value = clsCommon.myCdbl(dr("landing_cost"))

                    gv1.CurrentRow.Cells(ColFOC).Value = 0
                    gv1.CurrentRow.Cells(colItemWeight).Value = clsCommon.myCdbl(dr("Weight Value"))
                    'gv1.CurrentRow.Cells(colPurCost).Value = clsCommon.myCdbl(dr("Purchase_Cost"))
                    gv1.CurrentRow.Cells(colOrgCost).Value = clsCommon.myCdbl(dr("BasicRate"))
                    Dim dblConvF As Double
                    dblConvF = GetConvFactor(gv1.CurrentRow.Cells(colUnit).Value, gv1.CurrentRow.Cells(colICode).Value)
                    gv1.CurrentRow.Cells(colConvF).Value = dblConvF
                End If
            Else
                SetBlankOfItemColumns()
            End If


        End If


        'setBalance()

    End Sub
    Sub RefreshSerialNo()
        Dim intSerialNo As Integer
        For intCount As Integer = 0 To gv1.Rows.Count - 1
            intSerialNo += 1
            gv1.Rows(intCount).Cells(colLineNo).Value = intSerialNo
        Next


    End Sub
    Private Sub findVolumeStructureSchemeCode(ByVal isButtonClick As Boolean, Optional ByVal SchemeCode As String = Nothing)
        Dim isDefault As Boolean = False
        isDefault = clsCommon.myCBool(IIf(clsFixedParameter.GetData(clsFixedParameterType.IsVolumeSchemeBydefault, clsFixedParameterCode.IsVolumeSchemeBydefault, Nothing) = "1", True, False))
        If isDefault Then
            For ii As Integer = gv1.Rows.Count - 1 To 0 Step -1
                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colSchmCodeType).Value), "VolumeSlab") = CompairStringResult.Equal And clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeApplicable).Value) = "No" Then
                    gv1.Rows.RemoveAt(ii)
                End If
            Next

            Dim arrStructure As List(Of String) = New List(Of String)
            For ii As Integer = 0 To gv1.Rows.Count - 1
                If Not clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal Then
                    Dim dblTotalQty As Double = 0
                    Dim MainUOM As String = ""
                    Dim MainQty As Decimal = 0
                    Dim Schm_Qty As Double = 0
                    Dim Schm_Icode As String = ""
                    Dim Schm_Item_Uom As String = ""
                    Dim qty As Decimal = 0
                    Dim mode As Decimal = 0
                    Dim free_qty As Double = 0
                    Dim strOuterItem As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                    '' Dim strOuterstructure As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIStruct).Value)
                    Dim strOuterstructure As String = clsItemMaster.GetItemStructureCode(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value), Nothing)
                    If clsCommon.myLen(strOuterstructure) > 0 Then
                        If Not arrStructure.Contains(strOuterstructure) Then
                            For jj As Integer = 0 To gv1.Rows.Count - 1

                                If clsCommon.myCdbl(gv1.Rows(jj).Cells(colQty).Value) > 0 AndAlso (clsCommon.myCstr(gv1.Rows(jj).Cells(colSchemeItem).Value) = "" OrElse clsCommon.myCstr(gv1.Rows(jj).Cells(colSchemeItem).Value) = "No") AndAlso Not clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colSchmCodeType).Value), "VolumeSlab") = CompairStringResult.Equal Then
                                    Dim wt_qty As Double = 0
                                    Dim wt_unit As String = Nothing
                                    'Dim strInnerstructure As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colIStruct).Value)

                                    Dim strInnerstructure As String = clsItemMaster.GetItemStructureCode(clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value), Nothing)
                                    Dim Main_Item_Code As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                                    Dim Main_Item_Qty As Double = clsCommon.myCdbl(gv1.Rows(jj).Cells(colQty).Value)
                                    Dim Main_Item_Unit As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                                    Dim Product_Type As String = clsCommon.myCstr(clsItemMaster.GetItemProductType(Main_Item_Code, Nothing))
                                    If clsCommon.CompairString(strOuterstructure, strInnerstructure) = CompairStringResult.Equal Then
                                        wt_qty = clsItemMaster.GetItemWeightValue(Main_Item_Code, Nothing)
                                        wt_unit = clsItemMaster.GetItemWeightUnit(Main_Item_Code, Nothing)
                                        Main_Item_Qty = Main_Item_Qty * wt_qty * clsCommon.myCdbl(clsItemMaster.GetConvertionFactor(Main_Item_Code, Main_Item_Unit, Nothing))

                                        Dim Cnvrsn_qty As Double = 1
                                        qty += (Main_Item_Qty / Cnvrsn_qty)

                                    End If
                                End If
                            Next
                            ''richa agarwal 19 June,2019 add equal to with greater than and less than slab range ERO/19/06/19-000647
                            Dim qry = "select top 1 TSPL_SCHEME_MASTER_NEW.Scheme_Code,TSPL_SCHEME_MASTER_NEW.Scheme_Type,TSPL_SCHEME_MASTER_VOLUME_SLAB.Item_Code as schemeItemCode, " &
                                   "TSPL_SCHEME_MASTER_VOLUME_SLAB.qty as schemeQty,TSPL_SCHEME_MASTER_VOLUME_SLAB.Unit_Code as SchemeUnit, " &
                                   "TSPL_SCHEME_MASTER_VOLUME_SLAB.Min_Range,Max_Range from TSPL_SCHEME_MASTER_NEW left outer join " &
                                   "TSPL_SCHEME_MASTER_VOLUME_SLAB on TSPL_SCHEME_MASTER_VOLUME_SLAB.Scheme_Code=TSPL_SCHEME_MASTER_NEW.Scheme_Code  left outer join " &
                                   "tspl_item_master on tspl_item_master.item_code=TSPL_SCHEME_MASTER_VOLUME_SLAB.Item_Code  where " &
                                   "TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select scheme_code from (select ROW_NUMBER() over (partition by scheme_type order by MaxlimitStart_Date desc) as sno, " &
                                   "TSPL_SCHEME_MASTER_NEW.Scheme_Code from  TSPL_SCHEME_MASTER_NEW where TSPL_SCHEME_MASTER_NEW.Scheme_Type='VolumeSlab' and " &
                                   "TSPL_SCHEME_MASTER_NEW.Structure_Code='" & strOuterstructure & "' and TSPL_SCHEME_MASTER_NEW.Status='Active' and " &
                                   "TSPL_SCHEME_MASTER_NEW.MaxlimitStart_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' and (TSPL_SCHEME_MASTER_NEW.MaxlimitEnd_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or " &
                                   "TSPL_SCHEME_MASTER_NEW.MaxlimitEnd_Date is null  ) " &
                                   "and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select TSPL_SCHEME_MASTER_VOLUME_SLAB.Scheme_Code from  TSPL_SCHEME_MASTER_VOLUME_SLAB left outer join TSPL_SCHEME_BENEFICIARY on " &
                                   "TSPL_SCHEME_MASTER_VOLUME_SLAB.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code " &
                                   "where TSPL_SCHEME_MASTER_NEW.Structure_Code='" & strOuterstructure & "' and Cust_Code='" & txtVendorNo.Value & "'))a where a.sno=1) " &
                                   "and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select Scheme_Code from TSPL_SCHEME_BENEFICIARY where Cust_Code='" & txtVendorNo.Value & "') " &
                                   "and TSPL_SCHEME_MASTER_NEW.Status='Active' and TSPL_SCHEME_MASTER_NEW.Structure_Code='" & strOuterstructure & "'  and " &
                                   "TSPL_SCHEME_MASTER_NEW.Scheme_Type='VolumeSlab' and " & qty & " >= TSPL_SCHEME_MASTER_VOLUME_SLAB.Min_Range and " &
                                   " " & qty & " <= TSPL_SCHEME_MASTER_VOLUME_SLAB.Max_Range  order by TSPL_SCHEME_MASTER_NEW.Scheme_Code ,Min_Range "
                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                            For Each dr As DataRow In dt.Rows
                                SchemeCode = clsCommon.myCstr(dr("Scheme_Code"))
                                Schm_Icode = clsCommon.myCstr(dr("schemeItemCode"))
                                Schm_Item_Uom = clsCommon.myCstr(dr("SchemeUnit"))
                                Schm_Qty = clsCommon.myCdbl(dr("schemeQty"))
                            Next
                            If Schm_Qty > 0 Then
                                If clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <> 0 Then
                                    gv1.Rows.AddNew()
                                End If
                                Dim Count As Integer = 0
                                For kk As Integer = 0 To gv1.Rows.Count - 1
                                    If clsCommon.myLen(gv1.Rows(kk).Cells(colICode).Value) <> 0 Then
                                        Count += 1
                                    End If
                                Next
                                gv1.Rows(Count).Cells(colRowType).Value = RowTypeItem
                                gv1.Rows(Count).Cells(colLineNo).Value = Count + 1
                                gv1.Rows(Count).Cells(colICode).Value = Schm_Icode
                                gv1.Rows(Count).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(Schm_Icode, Nothing)
                                gv1.Rows(Count).Cells(colIName).Value = clsItemMaster.GetItemName(Schm_Icode, Nothing)
                                ' gv1.Rows(Count).Cells(colIStruct).Value = clsItemMaster.GetItemStructureCode(Schm_Icode, Nothing)
                                'gv1.Rows(Count).Cells(colLocationCode).Value = txtBillToLocation.Value
                                'gv1.Rows(Count).Cells(colLocationName).Value = lblBillToLocation.Text
                                gv1.Rows(Count).Cells(colPriceCOde).Value = ""
                                'gv1.Rows(Count).Cells(colPriceDateColumn).Value = ""
                                gv1.Rows(Count).Cells(colUnit).Value = Schm_Item_Uom
                                gv1.Rows(Count).Cells(colUnitRate).Value = Schm_Item_Uom
                                gv1.Rows(Count).Cells(colMRP).Value = 0
                                gv1.Rows(Count).Cells(colRate).Value = 0
                                gv1.Rows(Count).Cells(colQty).Value = Schm_Qty
                                gv1.Rows(Count).Cells(colItemWeight).Value = 0
                                gv1.Rows(Count).Cells(colSchemeApplicable).Value = "No"
                                gv1.Rows(Count).Cells(colSchemeItem).Value = "Yes"
                                gv1.Rows(Count).Cells(colFromSchemeCode).Value = SchemeCode
                                gv1.Rows(Count).Cells(ColFOC).Value = 1
                                gv1.Rows(Count).Cells(colAbatementPer).Value = 0
                                gv1.Rows(Count).Cells(colActualCost).Value = 0
                                'If clsCommon.CompairString(clsCommon.myCstr(cmbDisItemType.SelectedValue), "NT") = CompairStringResult.Equal Then
                                '    If ShowSchemeItemRate Then
                                '        gv1.Rows(Count).Cells(colRate).Value = ItemPrice(gv1.Rows(Count).Cells(colICode).Value, gv1.Rows(Count).Cells(colUnit).Value, gv1.Rows(Count).Index)
                                '    End If
                                'ElseIf clsCommon.CompairString(clsCommon.myCstr(cmbDisItemType.SelectedValue), "T") = CompairStringResult.Equal Then
                                '    If ShowSchemeItemRateTaxable Then
                                '        gv1.Rows(Count).Cells(colRate).Value = ItemPrice(gv1.Rows(Count).Cells(colICode).Value, gv1.Rows(Count).Cells(colUnit).Value, gv1.Rows(Count).Index)
                                '    End If
                                'End If

                                gv1.Rows(Count).Cells(colRate).Value = 0
                                gv1.Rows(Count).Cells(colSchmCodeType).Value = "VolumeSlab"
                                gv1.Rows(Count).Cells(colMainIcode).Value = strOuterstructure
                                gv1.Rows(Count).Cells(colMainIQty).Value = ""
                                gv1.Rows(Count).Cells(colMainIUOM).Value = ""
                                If clsCommon.myLen(txtReqNo.Value) > 0 Then
                                    gv1.Rows(Count).Cells(colOrderNo).Value = txtReqNo.Value
                                Else
                                    gv1.Rows(Count).Cells(colOrderNo).Value = ""
                                End If

                                gv1.Rows(Count).Cells(ColFOC).Value = 1
                                gv1.Rows(Count).Cells(colICodeGrp).Value = clsDBFuncationality.getSingleValue("select CSA_TYPE from TSPL_ITEM_MASTER where Item_Code='" & gv1.Rows(Count).Cells(colICode).Value & "' ")
                                gv1.Rows(Count).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(Schm_Icode)
                                'gv1.Rows(Count).Cells(colSchemeApplicable).ReadOnly = True
                                gv1.Rows(Count).Cells(colQty).ReadOnly = True
                                gv1.Rows(Count).Cells(colSchmCodeType).ReadOnly = True
                                'gv1.Rows(Count).Cells(colActualUOM).ReadOnly = True
                                'gv1.Rows(Count).Cells(colActualQty).ReadOnly = True
                                arrStructure.Add(strOuterstructure)
                            End If
                        End If
                    End If
                End If
            Next

            For ii As Integer = 0 To gv1.Rows.Count - 1
                gv1.Rows(ii).Cells(colLineNo).Value = ii + 1
            Next
        End If
        ' findQtyandPromoSchemeCode(False, "")
    End Sub


    Private Sub findQtyandPromoSchemeCode(ByVal isButtonClick As Boolean, ByVal strSchemeCode As String, Optional ByVal OrderDate? As DateTime = Nothing)

        'Dim dr1 As DataTable
        ' Dim schemeCodeCol As String
        Dim LocCodeCol As String = String.Empty
        Dim LocNameCol As String = String.Empty
        Dim intRow As Integer = 0
        Dim strOrderCode As String = ""
        'LocCodeCol = clsCommon.myCstr(gv1.CurrentRow.Cells(colLocationCode).Value)
        'LocNameCol = clsCommon.myCstr(gv1.CurrentRow.Cells(colLocationName).Value)
        Try
            Dim Index As Integer = gv1.CurrentRow.Index

            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeApplicable).Value), "No") = CompairStringResult.Equal OrElse clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value) = 0 Then
                For schemeRow As Integer = gv1.Rows.Count - 1 To 0 Step -1
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colSchemeApplicable).Value), "No") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(ColFOC).Value), 1) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colMainIcode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)) = CompairStringResult.Equal Then
                        gv1.Rows.RemoveAt(schemeRow)
                    End If
                Next
                gv1.Rows(Index).Cells(colCash_Amt).Value = Nothing
                gv1.Rows(Index).Cells(colCash_Pers).Value = Nothing
                gv1.Rows(Index).Cells(colCashSchemeCode).Value = Nothing
                gv1.Rows(Index).Cells(colCashSchemeType).Value = Nothing
                gv1.Rows(Index).Cells(colFromSchemeCode).Value = Nothing
                gv1.Rows(Index).Cells(colSchmCodeType).Value = Nothing
                gv1.Rows(Index).Cells(colMainIcode).Value = Nothing
                gv1.Rows(Index).Cells(colMainIQty).Value = Nothing
                gv1.Rows(Index).Cells(colMainIUOM).Value = Nothing
                gv1.Rows(Index).Cells(colSchemeApplicable).Value = "No"
                'gv1.Rows(Index).Cells(colSchemeItem).Value = "No"

                gv1.Rows(Index).Cells(ColFOC).Value = 0

                RefreshSerialNo()
            End If
            gv1.CurrentRow.Cells(colSchemeApplicable).Value = "Yes"
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(Index).Cells(colSchmCodeType).Value), "None") <> CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(Index).Cells(colSchmCodeType).Value), "") <> CompairStringResult.Equal Then

                If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeApplicable).Value), "Yes") = CompairStringResult.Equal Then
                    For schemeRow As Integer = gv1.Rows.Count - 1 To 0 Step -1
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colSchemeApplicable).Value), "No") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(ColFOC).Value), 1) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(schemeRow).Cells(colMainIcode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)) = CompairStringResult.Equal Then
                            gv1.Rows.RemoveAt(schemeRow)
                        End If
                    Next
                    '-------------fill cash scheme---------------------------------------
                    'Dim obj_Cash As clsSchemeApplyOnDairy = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(Index).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(Index).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(Index).Cells(colQty).Value), txtVendorNo.Value)
                    'If obj_Cash IsNot Nothing Then
                    '    gv1.Rows(Index).Cells(colCash_Amt).Value = obj_Cash.Cash_Amt
                    '    gv1.Rows(Index).Cells(colCash_Pers).Value = obj_Cash.Cash_Pers
                    '    gv1.Rows(Index).Cells(colCashSchemeCode).Value = obj_Cash.Schm_Code
                    '    If clsCommon.myCdbl(obj_Cash.Cash_Pers) > 0 Then
                    '        gv1.Rows(Index).Cells(colCashSchemeType).Value = "P"
                    '        gv1.CurrentRow.Cells(colCash_Amt).Value = System.Math.Round((clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value) * clsCommon.myCdbl(gv1.CurrentRow.Cells(colRate).Value) * clsCommon.myCdbl(gv1.CurrentRow.Cells(colCash_Pers).Value)) / 100, 2)
                    '    ElseIf clsCommon.myCdbl(obj_Cash.Cash_Amt) > 0 Then
                    '        gv1.Rows(Index).Cells(colCashSchemeType).Value = "A"
                    '    End If
                    '    'gv1.Rows(Index).Cells(colSchemeItem).Value = "Yes"
                    'Else
                    '    gv1.Rows(Index).Cells(colCash_Amt).Value = Nothing
                    '    gv1.Rows(Index).Cells(colCash_Pers).Value = Nothing
                    '    gv1.Rows(Index).Cells(colCashSchemeCode).Value = Nothing
                    '    gv1.Rows(Index).Cells(colCashSchemeType).Value = Nothing
                    '    gv1.Rows(Index).Cells(colSchemeItem).Value = "No"

                    '    '==========if cash scheme is there but no quantitive or volumn then also scheme item set to Y
                    '    'If clsCommon.myLen(gv1.Rows(Index).Cells(colMainIcode).Value) > 0 Then
                    '    '    gv1.Rows(Index).Cells(colIsSchmItem).Value = "Y"
                    '    'End If
                    'End If
                    '------------------------------------------------------------------

                    'Dim objD As clsSchemeApplyOnDairy = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnitRate).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colQty).Value), txtVendorNo.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colSchmCodeType).Value))
                    Dim objD As clsSchemeApplyOnDairy = Nothing
                    ''richa BHA/25/02/19-000823 '' remove schemecode from below function
                    objD = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), txtVendorNo.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colSchmCodeType).Value), strSchemeCode, OrderDate)
                    If objD.Arr.Count = 0 AndAlso clsCommon.myLen(gv1.CurrentRow.Cells(colUnitALter).Value) > 0 Then
                        objD = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnitALter).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colQty).Value), txtVendorNo.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colSchmCodeType).Value), strSchemeCode, OrderDate)
                    End If
                    If objD IsNot Nothing AndAlso objD.Arr.Count > 0 Then
                        For Each objtr As clsSchemeApplyOnDairy In objD.Arr
                            If clsCommon.myLen(LocCodeCol) = 0 Then
                                LocCodeCol = txtBillToLocation.Value
                                LocNameCol = lblBillToLocation.Text
                            End If
                            'Dim dblSchemeItemActualQTy As Double = clsItemLocationDetails.getBalance(objtr.Schm_Icode, LocCodeCol, txtDocNo.Value, txtDate.Value, Nothing, objtr.Schm_IUnit_Rate, Nothing)
                            '--------------update free itemcode in main item row------------------
                            gv1.Rows(Index).Cells(colFromSchemeCode).Value = objtr.Schm_Code
                            gv1.Rows(Index).Cells(colSchmCodeType).Value = objtr.schm_Type
                            gv1.Rows(Index).Cells(colMainIcode).Value = objtr.Schm_Icode
                            gv1.Rows(Index).Cells(colMainIQty).Value = objtr.Schm_Qty
                            gv1.Rows(Index).Cells(colMainIUOM).Value = objtr.Schm_Item_Uom
                            gv1.Rows(Index).Cells(ColFOC).Value = 0
                            gv1.Rows(Index).Cells(colSchemeApplicable).Value = "Yes"
                            Dim MainItemCode = gv1.Rows(Index).Cells(colICode).Value
                            Dim MainItemUnit = gv1.Rows(Index).Cells(colUnitRate).Value
                            Dim MainItemQty As Double = gv1.Rows(Index).Cells(colQty).Value
                            Dim MainSaleOrderCode As String = gv1.Rows(Index).Cells(colOrderNo).Value
                            Dim MainTaxPaid As String = gv1.Rows(Index).Cells(colTAX_PAID).Value
                            '-------------------------------------------------------------

                            '-------------------------------------------------------------

                            gv1.Rows.AddNew()

                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = intRow + 1
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objtr.Schm_Icode
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objtr.Schm_Iname
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(objtr.Schm_Icode, Nothing)
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = LocCodeCol
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = LocNameCol
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceCOde).Value = ""
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceDateColumn).Value = ""
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objtr.Schm_Item_Uom
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitRate).Value = objtr.Schm_Item_Uom
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitALter).Value = objtr.Schm_Item_Uom
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = 0
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTAX_PAID).Value = MainTaxPaid
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = 0
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = objtr.Schm_Qty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colItemWeight).Value = 0
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = "No"
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = "Yes"
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colFromSchemeCode).Value = objtr.Schm_Code
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColFOC).Value = 1
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementPer).Value = 0
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colActualCost).Value = 0
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = Math.Round(clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), IIf(clsCommon.myLen(clsCommon.myCstr(txtSubLocation.Value)) > 0, txtSubLocation.Value, txtBillToLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnitRate).Value), 0), 2)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(objtr.Schm_Icode)
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colSchmCodeType).Value = objtr.schm_Type
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colMainIcode).Value = MainItemCode
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colMainIQty).Value = MainItemQty
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colMainIUOM).Value = MainItemUnit
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = MainSaleOrderCode
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColFOC).Value = 1
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colICodeGrp).Value = clsDBFuncationality.getSingleValue("select CSA_TYPE from TSPL_ITEM_MASTER where Item_Code='" & gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value & "' ")
                            ''richa agarwal to add batch details for scheme item BHA/25/04/19-000867
                            Dim dtscheme As DataTable = clsDBFuncationality.GetDataTable(" select TSPL_SD_SHIPMENT_HEAD.Document_Code ,TSPL_SD_SHIPMENT_HEAD.Screen_Type,TSPL_SD_SHIPMENT_DETAIL.Line_No,TSPL_SD_SHIPMENT_DETAIL.Scheme_Applicable ,TSPL_SD_SHIPMENT_HEAD.Trans_Type    from TSPL_SD_SHIPMENT_HEAD left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.Document_Code where TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No ='" & clsCommon.myCstr(txtReqNo.Value) & "' AND TSPL_SD_SHIPMENT_DETAIL .Item_Code ='" & clsCommon.myCstr(objtr.Schm_Icode) & "' AND Scheme_Applicable ='N'")
                            If dtscheme IsNot Nothing AndAlso dtscheme.Rows.Count > 0 Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = clsBatchInventory.GetData(clsCommon.myCstr(dtscheme.Rows(0)("Trans_Type")), clsCommon.myCstr(dtscheme.Rows(0)("Document_Code")), objtr.Schm_Icode, clsCommon.myCstr(dtscheme.Rows(0)("Line_No")), Nothing, clsCommon.myCstr(dtscheme.Rows(0)("Screen_Type")))
                            End If

                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualUOM).ReadOnly = True
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualQty).ReadOnly = True
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualRetQty).ReadOnly = True
                            'gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).ReadOnly = True
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).ReadOnly = True
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colSchmCodeType).ReadOnly = True

                            Dim dblConvF As Double = clsItemMaster.GetConvertionFactor(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value, gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitRate).Value, Nothing)
                            gv1.Rows(intRow).Cells(colConvF).Value = dblConvF
                            gv1.Rows(intRow).Cells(colTotItemWt).Value = dblConvF * clsCommon.myCdbl(gv1.Rows(intRow).Cells(colItemWeight).Value) * clsCommon.myCdbl(gv1.Rows(intRow).Cells(colQty).Value)
                            Dim qry As String = ""
                            Dim Weight_UOM As String = ""
                            Dim SKU_VALUE As Decimal = 0

                            qry = "select Item_Code,Weight_UOM,Weight_Value from TSPL_ITEM_MASTER where Item_Code='" & gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value & "'"
                            Dim dtWt As New DataTable()
                            dtWt = clsDBFuncationality.GetDataTable(qry)
                            If dtWt.Rows.Count > 0 Then
                                SKU_VALUE = clsCommon.myCdbl(dtWt.Rows(0).Item("Weight_Value"))
                                Weight_UOM = clsCommon.myCstr(dtWt.Rows(0).Item("Weight_UOM"))
                            End If

                            'If clsCommon.myLen(Weight_UOM) <= 0 Then
                            '    clsCommon.MyMessageBoxShow("Weight UOM not defined for Item  " & gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value & "")
                            '    Exit Sub
                            'End If

                            gv1.Rows.Move(gv1.Rows.Count - 1, Index + 1)


                            'End If
                        Next
                    Else
                        gv1.Rows(Index).Cells(colFromSchemeCode).Value = Nothing
                        gv1.Rows(Index).Cells(colSchmCodeType).Value = Nothing
                        gv1.Rows(Index).Cells(colMainIcode).Value = Nothing
                        gv1.Rows(Index).Cells(colMainIQty).Value = Nothing
                        gv1.Rows(Index).Cells(colMainIUOM).Value = Nothing
                        gv1.Rows(Index).Cells(ColFOC).Value = 0
                        gv1.Rows(Index).Cells(colSchemeApplicable).Value = "No"
                    End If
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
    Sub OpenItemList(ByVal isButtonClick As Boolean)
        Dim whrCls As String = String.Empty
        If clsCommon.myLen(txtReqNo.Value) = 0 Then
            Dim qry As String = String.Empty
            qry = "select distinct TSPL_INVENTORY_MOVEMENT.Item_Code as [Code] ,tspl_item_master.Item_Desc as [Item Description],Sku_Seq as [Seq. No.],ISNULL(Short_Description,'') AS Short_Description,ISNULL(CSA_TYPE,'') AS [Item Group Type] from  TSPL_INVENTORY_MOVEMENT left outer join tspl_item_master on  TSPL_INVENTORY_MOVEMENT.Item_Code=tspl_item_master.item_code"
            whrCls = " inout='I' and (tspl_item_master.Is_FreshItem=1 or tspl_item_master.Is_Ambient=1 )  and Item_Type='F' and Is_Serial_Item=0"
            gv1.CurrentRow.Cells(colICode).Value = clsCommon.ShowSelectForm("PSSaleReturnItemfnd", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), "Code", isButtonClick)
        End If
        'SetBlankOfItemColumns()
        gv1.CurrentRow.Cells(colIName).Value = clsDBFuncationality.getSingleValue("select Item_Desc from TSPL_ITEM_MASTER where Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        gv1.CurrentRow.Cells(colICodeGrp).Value = clsDBFuncationality.getSingleValue("select CSA_TYPE from TSPL_ITEM_MASTER where Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        gv1.CurrentRow.Cells(colUnit).Value = clsDBFuncationality.getSingleValue("select UOM_Code from TSPL_ITEM_UOM_DETAIL where Default_UOM=1 and Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        gv1.CurrentRow.Cells(colOrgUnit).Value = gv1.CurrentRow.Cells(colUnit).Value
        If clsCommon.CompairString(ddlReturnType.SelectedValue, "P") = CompairStringResult.Equal Then
            gv1.CurrentRow.Cells(colQty).Value = 1
        End If

        'If clsCommon.CompairString(gv1.CurrentRow.Cells(colICodeGrp).Value, "CPD-DESI GHEE") = CompairStringResult.Equal Then
        '    CalDiffrate(CInt(gv1.CurrentRow.Index), True, colUnitRate)
        'End If
        SetTaxDetails()
    End Sub
    Sub OpenICodeList(ByVal isButtonClick As Boolean)
        gv1.CurrentRow.Cells(colRowType).Value = RowTypeItem
        Dim strItemType As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value)
        If clsCommon.myLen(strItemType) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Please select Row Type", Me.Text)
            Exit Sub
        End If

        '--------------07/08/2014------------------------n-level cat in query------------
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
        ElseIf clsCommon.myLen(txtPriceGroupCode.Text) > 0 Then
            strPriceGrp = "TSPL_PRICE_GROUP_MAPPING.price_group_code "
            strPriceCondition = " and (priceGroup ='" & txtPriceGroupCode.Text & "' and PriceGroupStatus='Y')"
            strPriceGrpStatus = "TSPL_PRICE_GROUP_MAPPING.status"
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
                    Dim intTaxRate As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT 1 FROM TSPL_LOCATION_SALE_TAX_MASTER WHERE Location_Code='" & txtBillToLocation.Value & "' AND Tax_Group_Code='" & txtTaxGroup.Value & "' AND Tax_Code='" & grow.Cells(colTTaxAutCode).Value & "'"))
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
                    repID = "RetPurOneItm1"
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
             "where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI'))xxx Where  Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             "" & strTaxRate & " " &
             "Order By Item_Code,Start_Date,UOM desc"
                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                Else
                    repID = "RetPurOneItm2"
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
             "where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI'))xxx Where Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             "" & strTaxRate & " " &
             "Order By Item_Code,Start_Date,UOM desc"

                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                End If
            End If

            '--------------cate. in columns-------------------------------------------
            If clsCommon.myLen(pivotheader) > 0 Then
                If PurchaseOneItemOneVendor = True Then
                    repID = "RetPurOneItm3"
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
             "where TSPL_ITEM_MASTER.Active=1  and Is_FreshItem=0 and Product_Type not in ('MI'))xxx Where  Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             "" & strTaxRate & " " &
             ")b left outer join (select TSPL_ITEM_MASTER_CATEGORY.Item_code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code,TSPL_ITEM_CATEGORY_LEVEL.DESCRIPTION,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as cat_value from TSPL_ITEM_MASTER_CATEGORY left outer join TSPL_ITEM_CATEGORY_LEVEL on TSPL_ITEM_CATEGORY_LEVEL.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and ISNULL(TSPL_ITEM_CATEGORY_LEVEL.Form_Type,'item')='item' left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  and ISNULL(TSPL_ITEM_CATEGORY_LEVEL_VALUES.Form_Type,'item')='item')a on a.Item_code=b.Item) as s pivot(max(cat_value) for description in (" + pivotheader + "))t"
                    '"Order By Item_Code,Start_Date,UOM desc"
                    dr = clsCommon.ShowSelectFormForRow(repID, qry)
                Else
                    repID = "RetPurOneItm4"
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
             "where TSPL_ITEM_MASTER.Active=1 and Is_FreshItem=0 and Product_Type not in ('MI') )xxx Where Tax_group='" & txtTaxGroup.Value & "' " & strPriceCondition & " " &
             "" & strTaxRate & " " &
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
                    'gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dr("Unit")), clsCommon.myCdbl(dr("MRP")))
                    'gv1.CurrentRow.Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalanceWithUnapprove(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtBillToLocation.Value, clsCommon.myCdbl(dr("MRP")), clsCommon.myCstr(dr("Unit")), txtDocNo.Value, txtDate.Value)
                    'If gv1.CurrentRow.Cells(ColActualBalQty).Value = 0 Then
                    '    isCellValueChangedOpen = False

                    '    Throw New Exception("Qty is not avaliable for item " & gv1.CurrentRow.Cells(colICode).Value & " at location " & txtBillToLocation.Value & " ")
                    'End If
                    If PurchaseOneItemOneVendor = True Then
                        gv1.CurrentRow.Cells(colPricipleCode).Value = clsCommon.myCstr(dr("PrincipleCode"))
                        gv1.CurrentRow.Cells(colPricipleDesc).Value = clsCommon.myCstr(dr("PrincipleDesc"))
                        gv1.CurrentRow.Cells(colvendorCode).Value = clsCommon.myCstr(dr("Vendor code"))
                        gv1.CurrentRow.Cells(colvendorDesc).Value = clsCommon.myCstr(dr("Vendor Desc"))
                    End If
                    gv1.CurrentRow.Cells(colIName).Value = clsCommon.myCstr(dr("ItemDesc"))
                    gv1.CurrentRow.Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), Nothing)
                    gv1.CurrentRow.Cells(colPriceCOde).Value = clsCommon.myCstr(dr("Price_Code"))
                    gv1.CurrentRow.Cells(colPriceDateColumn).Value = clsCommon.myCstr(dr("Start_Date"))
                    gv1.CurrentRow.Cells(colUnit).Value = clsCommon.myCstr(dr("Unit"))
                    gv1.CurrentRow.Cells(colOrgUnit).Value = clsCommon.myCstr(dr("Unit"))
                    gv1.CurrentRow.Cells(colMRP).Value = clsCommon.myCdbl(dr("MRP"))
                    gv1.CurrentRow.Cells(colRate).Value = clsCommon.myCstr(dr("BasicRate"))
                    gv1.CurrentRow.Cells(colAbatementPer).Value = clsCommon.myCdbl(dr("abatement_rate"))
                    gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
                    gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
                    gv1.CurrentRow.Cells(colActualCost).Value = clsCommon.myCdbl(dr("BasicRate"))
                    gv1.CurrentRow.Cells(colTaxRate1).Value = clsCommon.myCdbl(dr("Tax1Rate"))
                    gv1.CurrentRow.Cells(colTaxRate2).Value = clsCommon.myCdbl(dr("Tax2Rate"))
                    gv1.CurrentRow.Cells(colTaxRate3).Value = clsCommon.myCdbl(dr("Tax3Rate"))
                    gv1.CurrentRow.Cells(colTaxRate4).Value = clsCommon.myCdbl(dr("Tax4Rate"))
                    gv1.CurrentRow.Cells(colTaxRate5).Value = clsCommon.myCdbl(dr("Tax5Rate"))
                    gv1.CurrentRow.Cells(colTaxRate6).Value = clsCommon.myCdbl(dr("Tax6Rate"))
                    gv1.CurrentRow.Cells(colTaxRate7).Value = clsCommon.myCdbl(dr("Tax7Rate"))
                    gv1.CurrentRow.Cells(colTaxRate8).Value = clsCommon.myCdbl(dr("Tax8Rate"))
                    gv1.CurrentRow.Cells(colTaxRate9).Value = clsCommon.myCdbl(dr("Tax9Rate"))
                    gv1.CurrentRow.Cells(colTaxRate10).Value = clsCommon.myCdbl(dr("Tax10Rate"))

                    gv1.CurrentRow.Cells(ColFOC).Value = 0
                    gv1.CurrentRow.Cells(colItemWeight).Value = clsCommon.myCdbl(dr("Weight_Value"))
                    Dim dblConvF As Double
                    dblConvF = GetConvFactor(gv1.CurrentRow.Cells(colUnit).Value, gv1.CurrentRow.Cells(colICode).Value)
                    gv1.CurrentRow.Cells(colConvF).Value = dblConvF

                    gv1.CurrentRow.Cells(colMarkupOn).Value = clsCommon.myCstr(dr("markup_on"))
                    gv1.CurrentRow.Cells(colMarkUpPercentage).Value = clsCommon.myCdbl(dr("markup_percent"))
                    gv1.CurrentRow.Cells(colLandingCost).Value = clsCommon.myCdbl(dr("landing_cost"))
                    gv1.CurrentRow.Cells(colPurCost).Value = clsCommon.myCdbl(dr("Purchase_Cost"))
                    gv1.CurrentRow.Cells(colOrgCost).Value = clsCommon.myCdbl(dr("BasicRate"))


                End If
            Else
                SetBlankOfItemColumns()
            End If


        End If


        'setBalance()
        '    Dim obj As clsItemMaster = clsItemMaster.FinderForItem(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(cboItemType.SelectedValue), True, isButtonClick, "", txtVendorNo.Value)
        '    If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Item_Code) > 0 Then
        '        gv1.CurrentRow.Cells(colICode).Value = obj.Item_Code
        '        gv1.CurrentRow.Cells(colIName).Value = obj.Item_Desc
        '        gv1.CurrentRow.Cells(colUnit).Value = obj.Unit_Code
        '        gv1.CurrentRow.Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value))
        '        gv1.CurrentRow.Cells(colIsSerialseItem).Value = obj.Is_Serial_Item
        '        Dim objCustItem As clsCustomeritemDetails = clsCustomeritemDetails.GetItemRateAndDiscount(txtVendorNo.Value, obj.Item_Code, obj.Unit_Code, txtDate.Value)
        '        If objCustItem IsNot Nothing Then
        '            gv1.CurrentRow.Cells(colRate).Value = objCustItem.Approval_Item_Rate
        '            gv1.CurrentRow.Cells(colDisPer).Value = objCustItem.Discount_Per
        '        End If
        '    Else
        '        SetBlankOfItemColumns()
        '    End If
        '    ''End If
        '    Dim objVItem As clsVendorItemDetail = clsVendorItemDetail.GetItemRateAndMRP(txtVendorNo.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value))
        '    If objVItem IsNot Nothing Then
        '        gv1.CurrentRow.Cells(colRate).Value = objVItem.item_rate
        '        gv1.CurrentRow.Cells(colMRP).Value = objVItem.MRP
        '    Else

        '    End If
        'Else
        '    ''For Open Misc Charges 
        '    Dim obj As clsAdditionalCharge = clsAdditionalCharge.GetFinder(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), isButtonClick)
        '    If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Code) > 0 Then
        '        gv1.CurrentRow.Cells(colICode).Value = obj.Code
        '        gv1.CurrentRow.Cells(colIName).Value = obj.desc
        '        gv1.CurrentRow.Cells(colUnit).Value = Nothing
        '        gv1.CurrentRow.Cells(colQty).Value = Nothing
        '        gv1.CurrentRow.Cells(colRate).Value = Nothing
        '    Else
        '        SetBlankOfItemColumns()
        '    End If
        '    ''End of Misc Charges 
        'End If
        'SetitemWiseTaxSetting(True, True)
        'setBalance()
    End Sub



    Private Function GetConvFactor(ByVal strUnit As String, ByVal strItem As String) As Double
        Dim dblConvF As Double = 0
        dblConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strUnit & "'"))
        Return dblConvF
    End Function
    Private Sub SetBlankOfItemColumns()
        gv1.CurrentRow.Cells(colICode).Value = ""
        gv1.CurrentRow.Cells(colIName).Value = ""
        gv1.CurrentRow.Cells(colUnit).Value = ""
        gv1.CurrentRow.Cells(colMRP).Value = 0


    End Sub
    Private Sub gv1_UserAddedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserAddedRow
        For i As Integer = 0 To gv1.Rows.Count - 1
            gv1.Rows(0).Cells(0).Value = 1
            If i <> 0 Then
                gv1.Rows(i).Cells(colLineNo).Value = i + 1
            End If
        Next
    End Sub
    Private Function GetCurrentRowTotalTCSTaxAmt(ByVal IntRowNo As Integer) As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value), "TCS") = CompairStringResult.Equal Then
                If IntRowNo < 0 Then
                    dblTotTax = dblTotTax + Math.Round(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value), 2)
                Else
                    dblTotTax = dblTotTax + Math.Round(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value), 2)
                End If
            End If

        Next
        Return dblTotTax
    End Function
    Private Function GetCurrentRowTotalTaxAmt(ByVal IntRowNo As Integer) As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If Not clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value), "TCS") = CompairStringResult.Equal Then
                If IntRowNo < 0 Then
                    dblTotTax = dblTotTax + Math.Round(clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value), 2)
                Else
                    dblTotTax = dblTotTax + Math.Round(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMT" + strii)).Value), 2)
                End If
            End If
        Next
        Return dblTotTax
    End Function

    Private Function GetCurrentRowSurTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            Else
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            End If
        Next
        Return 0
    End Function

    Private Function GetCurrentRowOtherTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Double
        Dim dblRet As Double = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If IntRowNo < 0 Then
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                Else
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
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
                End If
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            End If
        Next
    End Sub

    Private Sub UpdateAllTotals()
        Dim dblCashDisAmt As Double = 0
        Dim dblTotAmt As Double = 0
        Dim dblTotDisAmt As Double = 0
        Dim dblAmtAfterDis As Double = 0
        Dim dblTotLandedCost As Double = 0


        Dim dblTaxBaseAmt1 As Double = 0
        Dim dblTaxBaseAmt2 As Double = 0
        Dim dblTaxBaseAmt3 As Double = 0
        Dim dblTaxBaseAmt4 As Double = 0
        Dim dblTaxBaseAmt5 As Double = 0
        Dim dblTaxBaseAmt6 As Double = 0
        Dim dblTaxBaseAmt7 As Double = 0
        Dim dblTaxBaseAmt8 As Double = 0
        Dim dblTaxBaseAmt9 As Double = 0
        Dim dblTaxBaseAmt10 As Double = 0

        Dim dblTaxAmt1 As Double = 0
        Dim dblTaxAmt2 As Double = 0
        Dim dblTaxAmt3 As Double = 0
        Dim dblTaxAmt4 As Double = 0
        Dim dblTaxAmt5 As Double = 0
        Dim dblTaxAmt6 As Double = 0
        Dim dblTaxAmt7 As Double = 0
        Dim dblTaxAmt8 As Double = 0
        Dim dblTaxAmt9 As Double = 0
        Dim dblTaxAmt10 As Double = 0
        Dim dblHeadDisPerAmt As Double = 0
        Dim dblHeadDisAmt As Double = 0
        Dim dblCrateQty As Double = 0
        Dim dblCanQty As Double = 0
        Dim dblTaxTotAmt As Double = 0
        Dim dblNetAmt As Double = 0
        For ii As Integer = 0 To gv1.Rows.Count - 1
            If (clsCommon.myLen(gv1.Rows(ii).Cells(colICode).Value) > 0) AndAlso clsCommon.myCdbl(gv1.Rows(ii).Cells(ColFOC).Value) = 0 Then
                dblCrateQty = dblCrateQty + clsCommon.myCdbl(gv1.Rows(ii).Cells(colCrate).Value)
                dblCanQty = dblCanQty + clsCommon.myCdbl(gv1.Rows(ii).Cells(colCan).Value)
                dblTotAmt = dblTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmt).Value)
                dblTotDisAmt = dblTotDisAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colDisAmt).Value)
                dblCashDisAmt = dblCashDisAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colCash_Amt).Value)
                'dblTotDisAmt = dblTotDisAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colDisAmt).Value)
                dblAmtAfterDis = dblAmtAfterDis + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmtAfterDis).Value)
                dblHeadDisAmt = dblHeadDisAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colHeadDiscamt).Value)
                dblHeadDisPerAmt = dblHeadDisPerAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colHeadDisPerAmt).Value)

                dblTaxAmt1 = dblTaxAmt1 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt1).Value), 2)
                dblTaxAmt2 = dblTaxAmt2 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt2).Value), 2)
                dblTaxAmt3 = dblTaxAmt3 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt3).Value), 2)
                dblTaxAmt4 = dblTaxAmt4 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt4).Value), 2)
                dblTaxAmt5 = dblTaxAmt5 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt5).Value), 2)
                dblTaxAmt6 = dblTaxAmt6 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt6).Value), 2)
                dblTaxAmt7 = dblTaxAmt7 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt7).Value), 2)
                dblTaxAmt8 = dblTaxAmt8 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt8).Value), 2)
                dblTaxAmt9 = dblTaxAmt9 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt9).Value), 2)
                dblTaxAmt10 = dblTaxAmt10 + Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt10).Value), 2)

                dblTaxBaseAmt1 = dblTaxBaseAmt1 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt1).Value)
                dblTaxBaseAmt2 = dblTaxBaseAmt2 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt2).Value)
                dblTaxBaseAmt3 = dblTaxBaseAmt3 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt3).Value)
                dblTaxBaseAmt4 = dblTaxBaseAmt4 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt4).Value)
                dblTaxBaseAmt5 = dblTaxBaseAmt5 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt5).Value)
                dblTaxBaseAmt6 = dblTaxBaseAmt6 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt6).Value)
                dblTaxBaseAmt7 = dblTaxBaseAmt7 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt7).Value)
                dblTaxBaseAmt8 = dblTaxBaseAmt8 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt8).Value)
                dblTaxBaseAmt9 = dblTaxBaseAmt9 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt9).Value)
                dblTaxBaseAmt10 = dblTaxBaseAmt10 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt10).Value)

                dblTaxTotAmt = dblTaxTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTotTaxAmt).Value)
                dblNetAmt = dblNetAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmtAfterTax).Value)

                dblTotLandedCost = dblTotLandedCost + clsCommon.myCdbl(gv1.Rows(ii).Cells(colLandedAmt).Value)
            End If
        Next

        If rbtnTaxCalAutomatic.IsChecked Then

            For ii As Integer = 1 To gv2.Rows.Count
                Select Case (ii)
                    Case 1
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt1, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt1, 2)
                        If dblTaxBaseAmt1 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt1 * 100) / dblTaxBaseAmt1, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 2
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt2, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt2, 2)
                        If dblTaxBaseAmt2 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt2 * 100) / dblTaxBaseAmt2, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 3
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt3, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt3, 2)
                        If dblTaxBaseAmt3 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt3 * 100) / dblTaxBaseAmt3, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 4
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt4, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt4, 2)
                        If dblTaxBaseAmt4 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt4 * 100) / dblTaxBaseAmt4, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 5
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt5, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt5, 2)
                        If dblTaxBaseAmt5 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt5 * 100) / dblTaxBaseAmt5, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 6
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt6, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt6, 2)
                        If dblTaxBaseAmt6 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt6 * 100) / dblTaxBaseAmt6, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 7
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt7, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt7, 2)
                        If dblTaxBaseAmt7 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt7 * 100) / dblTaxBaseAmt7, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 8
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt8, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt8, 2)
                        If dblTaxBaseAmt8 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt8 * 100) / dblTaxBaseAmt8, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 9
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt9, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt9, 2)
                        If dblTaxBaseAmt9 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt9 * 100) / dblTaxBaseAmt9, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 10
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt10, 2)
                        gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt10, 2)
                        If dblTaxBaseAmt10 <> 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt10 * 100) / dblTaxBaseAmt10, 3)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                End Select
            Next
        End If
        Dim dblACAmount As Double = 0
        For ii As Integer = 0 To gvAC.Rows.Count - 1
            If (clsCommon.myLen(gvAC.Rows(ii).Cells(colACCode).Value) > 0) Then
                dblACAmount = dblACAmount + clsCommon.myCdbl(gvAC.Rows(ii).Cells(colACAmount).Value)
            End If
        Next



        lblAmtWithDiscount.Text = clsCommon.myFormat(dblTotAmt)
        'lblDiscountAmt.Text = clsCommon.myFormat(dblTotDisAmt)
        lblDiscountAmt.Text = clsCommon.myFormat(dblTotDisAmt + dblCashDisAmt)
        lblAmtAfterDiscount.Text = clsCommon.myFormat(dblAmtAfterDis)
        lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt)

        lblAddCharges.Text = clsCommon.myFormat(dblACAmount)
        lblAddCharges1.Text = clsCommon.myFormat(dblACAmount)

        dblNetAmt = dblNetAmt + dblACAmount
        lblInvoiceDiscAmt.Text = dblHeadDisAmt + dblHeadDisPerAmt
        If UOMAtDiarySaleReturn = True Then
            lblTotRAmt.Text = System.Math.Round(dblNetAmt, 2)
        Else
            lblTotRAmt.Text = clsCommon.myFormat(dblNetAmt)
        End If

        'lblTotRAmt1.Text = lblTotRAmt.Text

        If UOMAtDiarySaleReturn = True Then

        End If
        txtCrateQty.Value = dblCrateQty
        TxtTotalCAN.Value = dblCanQty
        txtShippedCan.Value = TxtTotalCAN.Value

        'done by richa ERO/10/01/20-001173
        If ApplyRoundOffZero Then
            If Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0, MidpointRounding.AwayFromZero) > clsCommon.myCdbl(lblTotRAmt.Text) Then
                TxtRoundoff.Text = Math.Round(Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0, MidpointRounding.AwayFromZero) - clsCommon.myCdbl(lblTotRAmt.Text), 2)
                lblTotRAmt.Text = Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0, MidpointRounding.AwayFromZero)
            Else
                TxtRoundoff.Text = Math.Round(Math.Round(clsCommon.myCdbl(lblTotRAmt.Text)) - clsCommon.myCdbl(lblTotRAmt.Text), 2)
                lblTotRAmt.Text = Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0)
            End If
        End If

        lblTotRAmt1.Text = lblTotRAmt.Text
    End Sub

    'Private Function GetBaseOtherTaxableAmount(ByVal intEndCol As Integer) As Double
    ''Dim dblRetVal As Double = 0
    ''For ii As Integer = 0 To intEndCol - 1
    ''    If clsCommon.myCBool(gv2.Rows(ii).Cells(colTIsTaxable).Value) Then
    ''        dblRetVal = dblRetVal + clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
    ''    End If
    ''Next
    ''Return dblRetVal
    'End Function

    Private Function GetBaseTaxAmount(ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 0 To intEndCol
            If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAutCode).Value)) = CompairStringResult.Equal Then
                Return clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
            End If
        Next
        Return 0
    End Function

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        AddNew()

    End Sub

    Sub AddNew()
        chkBoothWiseReturn.Checked = False
        'If ApplyBoothWiseReturn Then
        '    RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Visible
        'Else
        '    RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
        'End If
        btnInvoiceJE.Visible = False

        BlankAllControls()
        LoadBlankGrid()
        LoadBlankGridAC()
        LoadBlankGridTax()
        'LoadBlankGridBooth()
        isNewEntry = True
        btnSave.Text = "Save"
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnDelete.Enabled = True
        txtDate.Focus()
        gv1.Rows.AddNew()
        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
        ''gv1.Rows.AddNew()
        chkInternal.Checked = False
        gvAC.Rows.AddNew()
        'gvAC.Rows.AddNew()
        txtDate.Enabled = True
        txtVendorNo.Enabled = True
        btnHistory.Enabled = False
        chkCncelDSR.Enabled = True
        chkCncelDSR.Checked = False
        SaleInvoiceDate = Nothing
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.SetDefaultValues()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()

    End Sub

    Function AllowToSave() As Boolean
        Try
            Xtra.TransactionValidity(txtDate.Value)
            'If ApplyBoothWiseReturn Then
            '    If gvBooth Is Nothing OrElse gvBooth.Rows.Count <= 0 Then
            '        Throw New Exception("Please Fill Booth Detail")
            '    End If
            '    MergeBooth()
            'End If
            RefreshReqNo()
            Dim strTaxgrp = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Group from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & txtReqNo.Value & "'"))
            UpdateAllTotals()
            If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer", Me.Text)
                txtVendorNo.Focus()
                Return False
            End If

            If clsCommon.myLen(strTaxgrp) > 0 Then
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
            If Not isNewEntry AndAlso clsCommon.myLen(txtDocNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Return No Not found to save", Me.Text)
                txtDocNo.Focus()
                Return False
            End If
            'If clsCommon.myLen(cboItemType.SelectedValue) <= 0 Then
            '    common.clsCommon.MyMessageBoxShow("Please select Item Type")
            '    cboItemType.Focus()
            '    Return False
            'End If
            Dim strReceiptNo = clsDBFuncationality.getSingleValue("select isnull((Select distinct '['+TSPL_RECEIPT_DETAIL.Receipt_No+']    ' from tspl_sd_sale_invoice_head  " &
           "left outer join TSPL_Customer_Invoice_Head on tspl_sd_sale_invoice_head.Document_Code=TSPL_Customer_Invoice_Head.Against_Sale_No  " &
           "left outer join TSPL_RECEIPT_DETAIL on TSPL_Customer_Invoice_Head.Document_No=TSPL_RECEIPT_DETAIL.Document_No " &
           "left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_DETAIL.Receipt_No=TSPL_RECEIPT_HEADER.Receipt_No  " &
           "where TSPL_RECEIPT_HEADER.Posted ='N' and tspl_sd_sale_invoice_head.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "' and tspl_sd_sale_invoice_head.Screen_Type='DS' and TSPL_SD_SALE_INVOICE_HEAD.Document_Code='" & txtReqNo.Value & "' for xml path('')),'')  as DocNo ")
            If clsCommon.myLen(strReceiptNo) > 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please Post these Receipt " & strReceiptNo & " before creating sale return")
                Return False
            End If
            If clsCommon.myCDate(clsDBFuncationality.getSingleValue("Select CONVERT(date, Document_Date,103) from TSPL_SD_SALE_INVOICE_HEAD where Document_Code ='" + txtReqNo.Value + "'")) > clsCommon.myCDate(txtDate.Value) Then
                txtDate.Focus()
                Throw New Exception("Date cannot be less than from Dispatch Date")
            End If
            If AllowCrateCanPhysicalStock = 1 Then
                Dim dblCrateQty As Double = clsCommon.myCdbl(txtCrate.Value)
                Dim dblTotalCan As Double = clsCommon.myCdbl(TxtTotalCAN.Value)
                Dim dblShippedCan As Double = clsCommon.myCdbl(txtShippedCan.Value)
                If dblShippedCan > dblTotalCan Then
                    Throw New Exception("Shipped Can cannot be greater than total can.")
                End If
            End If
            Dim arrReqNo As New List(Of String)
            Dim arrICode As New List(Of String)()

            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strReqNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colOrderNo).Value)
                Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                Dim strIName As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                Dim dblPendingQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colPendingQty).Value)
                Dim dblQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colQty).Value)
                Dim dblDamageQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colDamageQty).Value)
                Dim dblActualQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value)
                Dim dblBalSaleQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colBalanceQty).Value)
                Dim strUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                Dim strActualUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value)
                If clsCommon.myLen(strICode) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
                    If clsCommon.myLen(strUOM) <= 0 Then
                        common.clsCommon.MyMessageBoxShow(Me, "Please enter UOM for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                    If clsCommon.myLen(strReqNo) > 0 Then
                        If Not (arrReqNo.Contains(strReqNo)) Then
                            arrReqNo.Add(strReqNo)
                        End If
                        If (dblQty + dblDamageQty) > dblPendingQty Then
                            'common.clsCommon.MyMessageBoxShow("Item " + strICode + "( " + strIName.Trim() + " ) Entered Quantity(" + clsCommon.myCstr(dblQty) + ") Can't be more Pending Quantity(" + clsCommon.myCstr(dblPendingQty) + ").At Line No: " + clsCommon.myCstr(clsCommon.myCdbl(ii + 1)) + " ")
                            'Return False
                        End If
                        If clsCommon.CompairString(strUOM, clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value)) = CompairStringResult.Equal Then
                            If (dblActualQty + dblDamageQty) > dblQty Then
                                common.clsCommon.MyMessageBoxShow("Actual Quantity and Damaged Quantity can't exceed Return Quantity")
                                Return False
                            End If
                        Else
                            Dim ActualQty As Decimal = 0
                            Dim DamageQty As Decimal = 0
                            If clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value) > 0 Then
                                Dim Qry As String = "select CONVERT(DECIMAL(18, 2), (" & clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colActualQty).Value)) & " * ISNULL(TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 1)) / ISNULL(ConvertDiv.Conversion_Factor, 1) ) AS Actual_QTY
                                  from TSPL_ITEM_UOM_DETAIL
                                  INNER JOIN TSPL_ITEM_UOM_DETAIL AS ConvertDiv ON ConvertDiv.Item_Code ='" & strICode & "' AND ConvertDiv.UOM_Code ='" & strUOM & "'
                                  where TSPL_ITEM_UOM_DETAIL.Item_Code='" & strICode & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='" & strActualUOM & "'"

                                ActualQty = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
                            End If
                            If clsCommon.myCdbl(gv1.Rows(ii).Cells(colDamageQty).Value) > 0 Then
                                Dim query As String = "select CONVERT(DECIMAL(18, 2), (" & clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colDamageQty).Value)) & " * ISNULL(TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 1)) / ISNULL(ConvertDiv.Conversion_Factor, 1) ) AS Actual_QTY
                                  from TSPL_ITEM_UOM_DETAIL
                                  INNER JOIN TSPL_ITEM_UOM_DETAIL AS ConvertDiv ON ConvertDiv.Item_Code ='" & strICode & "' AND ConvertDiv.UOM_Code ='" & strUOM & "'
                                  where TSPL_ITEM_UOM_DETAIL.Item_Code='" & strICode & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='" & strActualUOM & "'"

                                DamageQty = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(query))
                            End If
                            If (ActualQty + DamageQty) > dblQty Then
                                common.clsCommon.MyMessageBoxShow("Actual Quantity and Damaged Quantity can't exceed Return Quantity")
                                Return False
                            End If
                        End If

                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeItem).Value), "No") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(strUOM, clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value)) = CompairStringResult.Equal Then
                                If dblQty > dblBalSaleQty Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Item " + strICode + "( " + strIName.Trim() + " ) Entered Quantity(" + clsCommon.myCstr(dblQty) + ") Can't be more Pending Quantity(" + clsCommon.myCstr(dblBalSaleQty) + ").At Line No: " + clsCommon.myCstr(clsCommon.myCdbl(ii + 1)) + " ")
                                    Return False
                                End If
                            End If


                            ''richa agarwal CHECK QTY with balance qty through conversion ERO/30/06/20-001270
                            If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value)) > 0 Then
                                Dim Qry As String = "select convert(decimal(18,3),(" & clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value) & "/LtrUnit.conversion_factor)*StockUnit.conversion_factor*CurrentUnit.conversion_factor) as Ltr_Qty FROM  tspl_item_uom_detail LtrUnit " & Environment.NewLine &
" left join tspl_item_uom_detail StockUnit on StockUnit.item_code='" & strICode & "'    and StockUnit.Stocking_Unit ='Y' " & Environment.NewLine &
" left join tspl_item_uom_detail CurrentUnit on CurrentUnit.item_code='" & strICode & "' WHERE  CurrentUnit.uom_code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value) & "' AND  LtrUnit.item_code='" & strICode & "' and LtrUnit.UOM_Code='" & strUOM & "'"

                                Dim dblConvertedQty = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
                                If dblConvertedQty > dblBalSaleQty Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Item " + strICode + "( " + strIName.Trim() + " ) Entered Quantity(" + clsCommon.myCstr(dblConvertedQty) + ") Can't be more Pending Quantity(" + clsCommon.myCstr(dblBalSaleQty) + ").At Line No: " + clsCommon.myCstr(clsCommon.myCdbl(ii + 1)) + " ")
                                    Return False
                                End If
                            End If

                        End If
                        ''richa agarwal CHECK QTY for scheme item 
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(strUOM, clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value)) = CompairStringResult.Equal Then
                                If clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value) > dblQty Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Item " + strICode + "( " + strIName.Trim() + " ) Entered Quantity(" + clsCommon.myCstr(dblQty) + ") Can't be more Pending Quantity(" + clsCommon.myCstr(dblBalSaleQty) + ").At Line No: " + clsCommon.myCstr(clsCommon.myCdbl(ii + 1)) + " ")
                                    Return False
                                End If
                            End If

                            If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value)) > 0 Then
                                Dim Qry As String = "select convert(decimal(18,3),(" & clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value) & "/LtrUnit.conversion_factor)*StockUnit.conversion_factor*CurrentUnit.conversion_factor) as Ltr_Qty FROM  tspl_item_uom_detail LtrUnit " & Environment.NewLine &
" left join tspl_item_uom_detail StockUnit on StockUnit.item_code='" & strICode & "'    and StockUnit.Stocking_Unit ='Y' " & Environment.NewLine &
" left join tspl_item_uom_detail CurrentUnit on CurrentUnit.item_code='" & strICode & "' WHERE  CurrentUnit.uom_code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colActualUOM).Value) & "' AND  LtrUnit.item_code='" & strICode & "' and LtrUnit.UOM_Code='" & strUOM & "'"

                                Dim dblConvertedQty = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
                                If dblConvertedQty > dblQty Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Item " + strICode + "( " + strIName.Trim() + " ) Entered Quantity(" + clsCommon.myCstr(dblConvertedQty) + ") Can't be more Pending Quantity(" + clsCommon.myCstr(dblBalSaleQty) + ").At Line No: " + clsCommon.myCstr(clsCommon.myCdbl(ii + 1)) + " ")
                                    Return False
                                End If
                            End If

                        End If

                        ''end of scheme item


                    End If
                    If clsCommon.CompairString(clsCommon.myCstr(cboItemType.SelectedValue), "F") = CompairStringResult.Equal AndAlso IsBatchDetailMandatory(gv1.Rows(ii).Cells(colUnit).Value) Then
                        If clsCommon.myCdbl(gv1.Rows(ii).Cells(colMRP).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow(Me, "Please enter MRP No for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colBatchNo).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow(Me, "Please enter Batch No for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colExpiry).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow(Me, "Please enter Expiry Date for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                        If clsCommon.myLen(gv1.Rows(ii).Cells(colManufactureDate).Value) <= 0 Then
                            common.clsCommon.MyMessageBoxShow(Me, "Please enter Manufacture Date for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            Return False
                        End If
                    End If
                    If Not arrICode.Contains(strICode) Then
                        arrICode.Add(strICode)
                    End If

                    If clsCommon.CompairString(clsCommon.myCstr(cboItemType.SelectedValue), "O") = CompairStringResult.Equal Then
                        For jj As Integer = ii + 1 To gv1.Rows.Count - 1
                            Dim strInICode As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                            Dim strInUOM As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                            If clsCommon.CompairString(strICode, strInICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colSchemeItem).Value), "No") = CompairStringResult.Equal Then
                                common.clsCommon.MyMessageBoxShow(Me, "Item Code " + strICode + " and Unit " + strUOM + " is repeted at Row No" + clsCommon.myCstr(ii + 1) + " and " + clsCommon.myCstr(jj + 1), Me.Text)
                                Return False
                            End If
                        Next
                    End If
                End If
                If clsCommon.myCBool(gv1.Rows(ii).Cells(colIsSerialseItem).Value) Then
                    Dim arrSerailNo As List(Of clsSerializeInvenotry) = TryCast(gv1.Rows(ii).Tag, List(Of clsSerializeInvenotry))
                    If arrSerailNo Is Nothing OrElse dblQty <> arrSerailNo.Count Then
                        common.clsCommon.MyMessageBoxShow(Me, "Please provice serial no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                End If

                'Batch wise item return Ticket No- ALF/22/05/18-000066
                If dblQty > 0 AndAlso clsCommon.myCBool(gv1.Rows(ii).Cells(colIsBatchItem).Value) Then
                    Dim arrBatchNo As List(Of clsBatchInventory) = TryCast(gv1.Rows(ii).Cells(colICode).Tag, List(Of clsBatchInventory))
                    If arrBatchNo Is Nothing Then
                        Throw New Exception("Please provide Batch no for item : " + strICode + " . At Line No" + clsCommon.myCstr(ii + 1))
                    Else
                        Dim tQty As Decimal = 0
                        For Each objBatch As clsBatchInventory In arrBatchNo
                            tQty += objBatch.Qty
                        Next
                        If Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value), 2, MidpointRounding.AwayFromZero) > 0 Then
                            dblQty = Math.Round(clsCommon.myCdbl(gv1.Rows(ii).Cells(colActualQty).Value), 2, MidpointRounding.AwayFromZero)
                        End If
                        If tQty.ToString("N3") <> dblQty.ToString("N3") Then
                            Throw New Exception("Item : " + strICode + " Entered Qty " + clsCommon.myCstr(dblQty) + Environment.NewLine + "And Batchwise Qty " + clsCommon.myCstr(tQty) + " . At Line No" + clsCommon.myCstr(ii + 1))
                        End If
                        '' BHA/17/08/18-000446
                        'If tQty < dblQty Then
                        '    Throw New Exception("Item : " + strICode + " Entered Qty " + clsCommon.myCstr(dblQty) + Environment.NewLine + "And Batchwise Qty " + clsCommon.myCstr(tQty) + " . At Line No" + clsCommon.myCstr(ii + 1))
                        'End If
                    End If
                End If
                'Batch wise item return Ticket No- ALF/22/05/18-000066
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
            'clsItemMaster.isItemOfSameType(clsCommon.myCstr(cboItemType.SelectedValue), cboItemType.Text, arrICode)
            clsPSInvoiceHead.IsValidCustomer(arrReqNo, txtVendorNo.Value, txtDate.Value)
            UcCustomFields1.AllowToSave()
            UcAttachment1.AllowToSave()

            If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True And clsCommon.myLen(strTaxgrp) > 0 Then
                If clsLocationWiseTax.IsValidTaxGroup(txtTaxGroup.Value, txtBillToLocation.Value, txtVendorNo.Value, "S", clsCommon.myCDate(txtDate.Value), Nothing) = False Then
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try

    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveData(False)
    End Sub

    Sub SaveData(ByVal ChekPostBtn As Boolean)
        Try
            '' Anubhooti 13-Sep-2014 BM00000003735
            If ChekPostBtn = False Then
                If FrmMainTranScreen.ValidateTransactionAccToFinYear("Sale Return", txtDate.Value) = False Then
                    Exit Sub
                End If
            End If
            ''
            Dim totalqty As Decimal = 0
            If Not AllowModifcationByApprovalUser Then
                clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value))
            End If

            If (AllowToSave()) Then


                Dim obj As New clsShipmentReturnHead()
                'Save Cancel Ticket No- ALF/16/05/18-000064
                obj.Is_Cancelled = IIf(chkCncelDSR.Checked, 1, 0)
                'Save Cancel Ticket No- ALF/16/05/18-000064
                obj.CrateMan = txtCrate.Value
                obj.jaali = txtJaali.Value
                obj.Box = txtBox.Value
                obj.ShippedCAN = txtShippedCan.Value
                obj.TotalCAN = TxtTotalCAN.Value
                obj.CrateQty = txtCrateQty.Value
                obj.Return_Type = ddlReturnType.SelectedValue
                obj.Shift_Type = IIf(rbtnMorning.IsChecked, "Morning", "Evening")
                If clsCommon.CompairString(obj.Return_Type, "D") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Return_Type, "I") = CompairStringResult.Equal Then
                    obj.Damage_Type = IIf(rbtn_leak.Checked, "0", "1")
                End If
                obj.Cust_PO_No = txtPONo.Text
                obj.Route_No = txtRouteNo.Value
                obj.Route_Desc = lblRouteDesc.Text
                obj.Price_Code = txtPriceCode.Text
                obj.Price_Group_Code = txtPriceGroupCode.Text
                'obj.HeadDisc_Per = txtDiscPer.Text
                'obj.HeadDisc_Amt = txtDiscAmt.Text
                obj.HeadDisc_Per = txtDiscPer.Text
                If obj.HeadDisc_Per > 0 Then
                    obj.HeadDisc_PerAmt = lblInvoiceDiscAmt.Text
                    obj.HeadDisc_Amt = 0
                Else
                    obj.HeadDisc_Amt = lblInvoiceDiscAmt.Text
                    obj.HeadDisc_PerAmt = 0
                End If
                obj.Invoice_Type = txtInvoiceType.Text

                If clsERPFuncationality.GetGSTStatus(txtDate.Value) = True Then
                    If rbtn_Ambient.IsChecked = True Then
                        obj.Is_Taxable = 1
                    Else
                        obj.Is_Taxable = 0
                    End If
                End If

                obj.Document_Code = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Customer_Code = txtVendorNo.Value
                obj.Customer_Name = lblVendorName.Text
                obj.Ref_No = txtRefNo.Text
                obj.Inv_Date = clsCommon.GetPrintDate(dtpInvoice.Value, "dd/MMM/yyyy")
                obj.Challan_Date = clsCommon.GetPrintDate(dtpChallan.Value, "dd/MMM/yyyy")
                obj.Total_Tax_Amt = clsCommon.myCdbl(lblTaxAmt.Text)
                obj.Inv_No = txtInvNo.Text
                obj.Bill_To_Location = txtBillToLocation.Value
                obj.Ship_To_Location = txtShipToLocation.Value
                obj.Sub_Location_code = txtSubLocation.Value
                obj.Comments = txtComment.Text
                obj.On_Hold = chkOnHold.Checked
                obj.Description = txtDesc.Text
                obj.Tax_Group = txtTaxGroup.Value
                obj.PROJECT_ID = fndProject.Text
                obj.Is_Internal = chkInternal.Checked
                If (gv2.Rows.Count > 0) Then
                    obj.TAX1 = clsCommon.myCstr(gv2.Rows(0).Cells(colTTaxAutCode).Value)
                    obj.TAX1_Rate = clsCommon.myCdbl(gv2.Rows(0).Cells(colTTaxRate).Value)
                    obj.TAX1_Base_Amt = clsCommon.myCdbl(gv2.Rows(0).Cells(colTBaseAmt).Value)
                    obj.TAX1_Amt = clsCommon.myCdbl(gv2.Rows(0).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 1) Then
                    obj.TAX2 = clsCommon.myCstr(gv2.Rows(1).Cells(colTTaxAutCode).Value)
                    obj.TAX2_Rate = clsCommon.myCdbl(gv2.Rows(1).Cells(colTTaxRate).Value)
                    obj.TAX2_Base_Amt = clsCommon.myCdbl(gv2.Rows(1).Cells(colTBaseAmt).Value)
                    obj.TAX2_Amt = clsCommon.myCdbl(gv2.Rows(1).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 2) Then
                    obj.TAX3 = clsCommon.myCstr(gv2.Rows(2).Cells(colTTaxAutCode).Value)
                    obj.TAX3_Rate = clsCommon.myCdbl(gv2.Rows(2).Cells(colTTaxRate).Value)
                    obj.TAX3_Base_Amt = clsCommon.myCdbl(gv2.Rows(2).Cells(colTBaseAmt).Value)
                    obj.TAX3_Amt = clsCommon.myCdbl(gv2.Rows(2).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 3) Then
                    obj.TAX4 = clsCommon.myCstr(gv2.Rows(3).Cells(colTTaxAutCode).Value)
                    obj.TAX4_Rate = clsCommon.myCdbl(gv2.Rows(3).Cells(colTTaxRate).Value)
                    obj.TAX4_Base_Amt = clsCommon.myCdbl(gv2.Rows(3).Cells(colTBaseAmt).Value)
                    obj.TAX4_Amt = clsCommon.myCdbl(gv2.Rows(3).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 4) Then
                    obj.TAX5 = clsCommon.myCstr(gv2.Rows(4).Cells(colTTaxAutCode).Value)
                    obj.TAX5_Rate = clsCommon.myCdbl(gv2.Rows(4).Cells(colTTaxRate).Value)
                    obj.TAX5_Base_Amt = clsCommon.myCdbl(gv2.Rows(4).Cells(colTBaseAmt).Value)
                    obj.TAX5_Amt = clsCommon.myCdbl(gv2.Rows(4).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 5) Then
                    obj.TAX6 = clsCommon.myCstr(gv2.Rows(5).Cells(colTTaxAutCode).Value)
                    obj.TAX6_Rate = clsCommon.myCdbl(gv2.Rows(5).Cells(colTTaxRate).Value)
                    obj.TAX6_Base_Amt = clsCommon.myCdbl(gv2.Rows(5).Cells(colTBaseAmt).Value)
                    obj.TAX6_Amt = clsCommon.myCdbl(gv2.Rows(5).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 6) Then
                    obj.TAX7 = clsCommon.myCstr(gv2.Rows(6).Cells(colTTaxAutCode).Value)
                    obj.TAX7_Rate = clsCommon.myCdbl(gv2.Rows(6).Cells(colTTaxRate).Value)
                    obj.TAX7_Base_Amt = clsCommon.myCdbl(gv2.Rows(6).Cells(colTBaseAmt).Value)
                    obj.TAX7_Amt = clsCommon.myCdbl(gv2.Rows(6).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 7) Then
                    obj.TAX8 = clsCommon.myCstr(gv2.Rows(7).Cells(colTTaxAutCode).Value)
                    obj.TAX8_Rate = clsCommon.myCdbl(gv2.Rows(7).Cells(colTTaxRate).Value)
                    obj.TAX8_Base_Amt = clsCommon.myCdbl(gv2.Rows(7).Cells(colTBaseAmt).Value)
                    obj.TAX8_Amt = clsCommon.myCdbl(gv2.Rows(7).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 8) Then
                    obj.TAX9 = clsCommon.myCstr(gv2.Rows(8).Cells(colTTaxAutCode).Value)
                    obj.TAX9_Rate = clsCommon.myCdbl(gv2.Rows(8).Cells(colTTaxRate).Value)
                    obj.TAX9_Base_Amt = clsCommon.myCdbl(gv2.Rows(8).Cells(colTBaseAmt).Value)
                    obj.TAX9_Amt = clsCommon.myCdbl(gv2.Rows(8).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 9) Then
                    obj.TAX10 = clsCommon.myCstr(gv2.Rows(9).Cells(colTTaxAutCode).Value)
                    obj.TAX10_Rate = clsCommon.myCdbl(gv2.Rows(9).Cells(colTTaxRate).Value)
                    obj.TAX10_Base_Amt = clsCommon.myCdbl(gv2.Rows(9).Cells(colTBaseAmt).Value)
                    obj.TAX10_Amt = clsCommon.myCdbl(gv2.Rows(9).Cells(colTTaxAmt).Value)
                End If

                obj.Terms_Code = txtTermCode.Value
                obj.Due_Date = txtDueDate.Value
                obj.Discount_Base = clsCommon.myCdbl(lblAmtWithDiscount.Text)
                obj.Discount_Amt = clsCommon.myCdbl(lblDiscountAmt.Text)
                obj.Amount_Less_Discount = clsCommon.myCdbl(lblAmtAfterDiscount.Text)
                obj.Total_Amt = clsCommon.myCdbl(lblTotRAmt.Text)
                obj.RoundOffAmount = clsCommon.myCdbl(TxtRoundoff.Text)
                obj.Carrier = txtCarrier.Text
                obj.VehicleNo = txtVehicleNo.Text
                obj.Vehicle_Code = txtVehcileCode.Value
                'obj.GRNo = txtGRNo.Text
                obj.GENo = txtGENo.Text

                If txtGEDate.Checked Then
                    obj.GEDate = txtGEDate.Value
                End If
                obj.Item_Type = clsCommon.myCstr(cboItemType.SelectedValue)
                obj.Dept = txtDept.Value
                obj.Dept_Desc = lblDept.Text

                obj.Against_Invoice_No = txtReqNo.Value


                If (gvAC.Rows.Count > 0) Then
                    If clsCommon.myLen(gvAC.Rows(0).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code1 = clsCommon.myCstr(gvAC.Rows(0).Cells(colACCode).Value)
                        obj.Add_Charge_Name1 = clsCommon.myCstr(gvAC.Rows(0).Cells(colACName).Value)
                        obj.Add_Charge_Amt1 = clsCommon.myCdbl(gvAC.Rows(0).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 1) Then
                    If clsCommon.myLen(gvAC.Rows(1).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code2 = clsCommon.myCstr(gvAC.Rows(1).Cells(colACCode).Value)
                        obj.Add_Charge_Name2 = clsCommon.myCstr(gvAC.Rows(1).Cells(colACName).Value)
                        obj.Add_Charge_Amt2 = clsCommon.myCdbl(gvAC.Rows(1).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 2) Then
                    If clsCommon.myLen(gvAC.Rows(2).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code3 = clsCommon.myCstr(gvAC.Rows(2).Cells(colACCode).Value)
                        obj.Add_Charge_Name3 = clsCommon.myCstr(gvAC.Rows(2).Cells(colACName).Value)
                        obj.Add_Charge_Amt3 = clsCommon.myCdbl(gvAC.Rows(2).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 3) Then
                    If clsCommon.myLen(gvAC.Rows(3).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code4 = clsCommon.myCstr(gvAC.Rows(3).Cells(colACCode).Value)
                        obj.Add_Charge_Name4 = clsCommon.myCstr(gvAC.Rows(3).Cells(colACName).Value)
                        obj.Add_Charge_Amt4 = clsCommon.myCdbl(gvAC.Rows(3).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 4) Then
                    If clsCommon.myLen(gvAC.Rows(4).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code5 = clsCommon.myCstr(gvAC.Rows(4).Cells(colACCode).Value)
                        obj.Add_Charge_Name5 = clsCommon.myCstr(gvAC.Rows(4).Cells(colACName).Value)
                        obj.Add_Charge_Amt5 = clsCommon.myCdbl(gvAC.Rows(4).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 5) Then
                    If clsCommon.myLen(gvAC.Rows(5).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code6 = clsCommon.myCstr(gvAC.Rows(5).Cells(colACCode).Value)
                        obj.Add_Charge_Name6 = clsCommon.myCstr(gvAC.Rows(5).Cells(colACName).Value)
                        obj.Add_Charge_Amt6 = clsCommon.myCdbl(gvAC.Rows(5).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 6) Then
                    If clsCommon.myLen(gvAC.Rows(6).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code7 = clsCommon.myCstr(gvAC.Rows(6).Cells(colACCode).Value)
                        obj.Add_Charge_Name7 = clsCommon.myCstr(gvAC.Rows(6).Cells(colACName).Value)
                        obj.Add_Charge_Amt7 = clsCommon.myCdbl(gvAC.Rows(6).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 7) Then
                    If clsCommon.myLen(gvAC.Rows(7).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code8 = clsCommon.myCstr(gvAC.Rows(7).Cells(colACCode).Value)
                        obj.Add_Charge_Name8 = clsCommon.myCstr(gvAC.Rows(7).Cells(colACName).Value)
                        obj.Add_Charge_Amt8 = clsCommon.myCdbl(gvAC.Rows(7).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 8) Then
                    If clsCommon.myLen(gvAC.Rows(8).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code9 = clsCommon.myCstr(gvAC.Rows(8).Cells(colACCode).Value)
                        obj.Add_Charge_Name9 = clsCommon.myCstr(gvAC.Rows(8).Cells(colACName).Value)
                        obj.Add_Charge_Amt9 = clsCommon.myCdbl(gvAC.Rows(8).Cells(colACAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 9) Then
                    If clsCommon.myLen(gvAC.Rows(9).Cells(colACCode).Value) > 0 Then
                        obj.Add_Charge_Code10 = clsCommon.myCstr(gvAC.Rows(9).Cells(colACCode).Value)
                        obj.Add_Charge_Name10 = clsCommon.myCstr(gvAC.Rows(9).Cells(colACName).Value)
                        obj.Add_Charge_Amt10 = clsCommon.myCdbl(gvAC.Rows(9).Cells(colACAmount).Value)
                    End If
                End If
                obj.Total_Add_Charge = clsCommon.myCdbl(lblAddCharges.Text)
                If rbtnTaxCalAutomatic.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                ElseIf rbtnTaxCalManual.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                End If
                obj.IsSampling = clsDBFuncationality.getSingleValue("select isnull(IsSampling,0) as IsSampling from TSPL_SD_SALE_INVOICE_HEAD where DOCUMENT_CODE='" & obj.Against_Invoice_No & "'")
                obj.Salesman_Code = txtSalesman.Value
                obj.Salesman_Name = lblSalesman.Text
                obj.Trans_type = IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
                obj.Screen_Type = "DS"
                totalqty = 0
                obj.Arr = New List(Of clsShipmentReturnDetail)
                Dim DCTotalAmt As Decimal = 0
                Dim TCTotalAmt As Decimal = 0
                Dim SCTotalAmt As Decimal = 0
                Dim intLineNo As Integer = 1
                For Each grow As GridViewRowInfo In gv1.Rows
                    Dim objTr As New clsShipmentReturnDetail()
                    objTr.Sampling = obj.IsSampling
                    objTr.Delivery_Code = clsCommon.myCstr(grow.Cells(colDelivery_Code).Value)
                    objTr.Crate = clsCommon.myCdbl(grow.Cells(colCrate).Value)
                    objTr.CAN = clsCommon.myCdbl(grow.Cells(colCan).Value)
                    objTr.Alter_UnitQty = clsCommon.myCdbl(grow.Cells(colAlterUnitQty).Value)
                    objTr.Rate_UnitQty = clsCommon.myCdbl(grow.Cells(colRateUnitQty).Value)
                    objTr.InvoiceScheme_Code = clsCommon.myCstr(grow.Cells(colInvoiceSchemeCode).Value)
                    objTr.Scheme_Item_Code = clsCommon.myCstr(grow.Cells(colMainIcode).Value)
                    objTr.Scheme_Item_UOM = clsCommon.myCstr(grow.Cells(colMainIUOM).Value)
                    objTr.Scheme_Qty = clsCommon.myCdbl(grow.Cells(colMainIQty).Value)
                    objTr.Scheme_Type = clsCommon.myCstr(grow.Cells(colSchmCodeType).Value)
                    objTr.Cash_Scheme_Code = clsCommon.myCstr(grow.Cells(colCashSchemeCode).Value)
                    objTr.Cash_Scheme_Type = clsCommon.myCstr(grow.Cells(colCashSchemeType).Value)
                    objTr.Cash_Scheme_Pers = clsCommon.myCdbl(grow.Cells(colCash_Pers).Value)
                    objTr.Cash_Scheme_Amount = clsCommon.myCdbl(grow.Cells(colCash_Amt).Value)
                    objTr.RATE_UOM = clsCommon.myCstr(grow.Cells(colUnitRate).Value)
                    objTr.Alternate_UOM = clsCommon.myCstr(grow.Cells(colUnitALter).Value)
                    objTr.Item_Group = clsCommon.myCstr(grow.Cells(colICodeGrp).Value)
                    objTr.TAX_PAID = clsCommon.myCstr(grow.Cells(colTAX_PAID).Value)

                    objTr.Line_No = intLineNo
                    objTr.Row_Type = clsCommon.myCstr(grow.Cells(colRowType).Value)
                    objTr.Item_Code = clsCommon.myCstr(grow.Cells(colICode).Value)
                    objTr.Item_Desc = clsCommon.myCstr(grow.Cells(colIName).Value)
                    objTr.Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
                    totalqty += objTr.Qty
                    objTr.DamageQty = clsCommon.myCdbl(grow.Cells(colDamageQty).Value)
                    objTr.Free_Qty = clsCommon.myCdbl(grow.Cells(colFreeQty).Value)

                    objTr.Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                    objTr.OrgUnit_code = clsCommon.myCstr(grow.Cells(colOrgUnit).Value)
                    objTr.Shipment_Code = clsCommon.myCstr(grow.Cells(colOrderNo).Value)
                    'objTr.Location = clsCommon.myCstr(grow.Cells(colloc).Value)
                    objTr.Item_Cost = clsCommon.myCdbl(grow.Cells(colRate).Value)
                    objTr.Amount = clsCommon.myCdbl(grow.Cells(colAmt).Value)
                    objTr.Damage_Amount = clsCommon.myCdbl(grow.Cells(colDamageAmt).Value)
                    objTr.Return_Amount = clsCommon.myCdbl(grow.Cells(colReturnAmt).Value)
                    objTr.Disc_Per = clsCommon.myCdbl(grow.Cells(colDisPer).Value)
                    objTr.Disc_Amt = clsCommon.myCdbl(grow.Cells(colDisAmt).Value)
                    objTr.Amt_Less_Discount = clsCommon.myCdbl(grow.Cells(colAmtAfterDis).Value)
                    objTr.TAX1 = clsCommon.myCstr(grow.Cells(colTax1).Value)
                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt1).Value)
                    objTr.TAX1_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate1).Value)
                    objTr.TAX1_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt1).Value)
                    objTr.TAX2 = clsCommon.myCstr(grow.Cells(colTax2).Value)
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt2).Value)
                    objTr.TAX2_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate2).Value)
                    objTr.TAX2_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt2).Value)
                    objTr.TAX3 = clsCommon.myCstr(grow.Cells(colTax3).Value)
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt3).Value)
                    objTr.TAX3_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate3).Value)
                    objTr.TAX3_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt3).Value)
                    objTr.TAX4 = clsCommon.myCstr(grow.Cells(colTax4).Value)
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt4).Value)
                    objTr.TAX4_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate4).Value)
                    objTr.TAX4_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt4).Value)
                    objTr.TAX5 = clsCommon.myCstr(grow.Cells(colTax5).Value)
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt5).Value)
                    objTr.TAX5_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate5).Value)
                    objTr.TAX5_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt5).Value)
                    objTr.TAX6 = clsCommon.myCstr(grow.Cells(colTax6).Value)
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt6).Value)
                    objTr.TAX6_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate6).Value)
                    objTr.TAX6_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt6).Value)
                    objTr.TAX7 = clsCommon.myCstr(grow.Cells(colTax7).Value)
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt7).Value)
                    objTr.TAX7_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate7).Value)
                    objTr.TAX7_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt7).Value)
                    objTr.TAX8 = clsCommon.myCstr(grow.Cells(colTax8).Value)
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt8).Value)
                    objTr.TAX8_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate8).Value)
                    objTr.TAX8_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt8).Value)
                    objTr.TAX9 = clsCommon.myCstr(grow.Cells(colTax9).Value)
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt9).Value)
                    objTr.TAX9_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate9).Value)
                    objTr.TAX9_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt9).Value)
                    objTr.TAX10 = clsCommon.myCstr(grow.Cells(colTax10).Value)
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt10).Value)
                    objTr.TAX10_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate10).Value)
                    objTr.TAX10_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt10).Value)
                    objTr.Total_Tax_Amt = clsCommon.myCdbl(grow.Cells(colTotTaxAmt).Value)
                    objTr.Item_Net_Amt = clsCommon.myCdbl(grow.Cells(colAmtAfterTax).Value)
                    objTr.Location = txtBillToLocation.Value 'clsCommon.myCstr(grow.Cells(colLocationCode).Value)


                    objTr.MRP = clsCommon.myCdbl(grow.Cells(colMRP).Value)
                    ''objTr.Assessable = clsCommon.myCdbl(grow.Cells(colAssessableRate).Value)
                    ''objTr.AssessableAmt = clsCommon.myCdbl(grow.Cells(colAssessableAmount).Value)
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
                    objTr.Is_Mannual_Amt = clsCommon.myCdbl(grow.Cells(colIsMannualAmt).Value)
                    ''richa agarwal 08 june,2020
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colSchemeItem).Value), "No") = CompairStringResult.Equal Then
                        objTr.Balance_Qty = clsCommon.myCdbl(grow.Cells(colBalanceQty).Value)
                    Else
                        objTr.Balance_Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
                    End If


                    objTr.Scheme_Applicable = clsCommon.myCstr(grow.Cells(colSchemeApplicable).Value)
                    objTr.Scheme_Code = clsCommon.myCstr(grow.Cells(colFromSchemeCode).Value)
                    objTr.Scheme_Item = clsCommon.myCstr(grow.Cells(colSchemeItem).Value)
                    objTr.Item_Tax = clsCommon.myCdbl(grow.Cells(colTotTaxAmt).Value)
                    objTr.Total_MRP_Amt = clsCommon.myCdbl(grow.Cells(colTotalMRP).Value)
                    objTr.Total_Basic_Amt = clsCommon.myCdbl(grow.Cells(colTotalBasicAmount).Value)
                    objTr.Total_Disc_Amt = clsCommon.myCdbl(grow.Cells(colTotalDiscountAmount).Value)
                    objTr.Price_with_tax = clsCommon.myCdbl(grow.Cells(colItemBasicPrice).Value)
                    objTr.Cust_Discount = clsCommon.myCdbl(grow.Cells(colcustDiscount).Value)
                    objTr.Total_Cust_Discount = clsCommon.myCdbl(grow.Cells(colTotalCustDiscount).Value)
                    objTr.ActualRate = clsCommon.myCdbl(grow.Cells(colActualCost).Value)
                    objTr.Cust_DiscountQty = clsCommon.myCdbl(grow.Cells(ColCustDiscountQty).Value)
                    objTr.Price_Date = clsCommon.myCDate(grow.Cells(colPriceDateColumn).Value)
                    objTr.Price_code = clsCommon.myCstr(grow.Cells(colPriceCOde).Value)
                    objTr.Abatement_Per = clsCommon.myCdbl(grow.Cells(colAbatementPer).Value)
                    objTr.Abatement_Amt = clsCommon.myCdbl(grow.Cells(colAbatementAmount).Value)
                    objTr.FOC_Item = clsCommon.myCdbl(grow.Cells(ColFOC).Value)
                    objTr.Item_Weight = clsCommon.myCdbl(grow.Cells(colItemWeight).Value)
                    objTr.Conv_Factor = clsCommon.myCdbl(grow.Cells(colConvF).Value)
                    objTr.TotalItem_Weight = clsCommon.myCdbl(grow.Cells(colTotItemWt).Value)

                    objTr.Markup_On = clsCommon.myCstr(grow.Cells(colMarkupOn).Value)
                    objTr.Markup_Percent = clsCommon.myCdbl(grow.Cells(colMarkUpPercentage).Value)
                    objTr.Landing_Cost = clsCommon.myCdbl(grow.Cells(colLandingCost).Value)
                    objTr.CustDiscPer = clsCommon.myCdbl(grow.Cells(colCustDiscPercentage).Value)
                    objTr.HeadDiscAmt = clsCommon.myCdbl(grow.Cells(colHeadDiscamt).Value)
                    objTr.CasdDiscScheme_Code = clsCommon.myCstr(grow.Cells(colCashDiscSchemeCode).Value)

                    objTr.Purchase_Cost = clsCommon.myCdbl(grow.Cells(colPurCost).Value)
                    objTr.OrgRate = clsCommon.myCdbl(grow.Cells(colOrgCost).Value)
                    objTr.PrincipleCode = clsCommon.myCstr(grow.Cells(colPricipleCode).Value)
                    objTr.PrincipleDesc = clsCommon.myCstr(grow.Cells(colPricipleDesc).Value)
                    objTr.vendor_code = clsCommon.myCstr(grow.Cells(colvendorCode).Value)
                    objTr.vendor_desc = clsCommon.myCstr(grow.Cells(colvendorDesc).Value)

                    objTr.HeadDiscPer = clsCommon.myCdbl(grow.Cells(colHeaDDisPer).Value)
                    objTr.HeadDiscPerAmt = clsCommon.myCdbl(grow.Cells(colHeadDisPerAmt).Value)

                    objTr.ActualUOM = clsCommon.myCstr(grow.Cells(colActualUOM).Value)
                    objTr.ActuaQty = clsCommon.myCdbl(grow.Cells(colActualQty).Value)
                    objTr.ActualReturnQty = clsCommon.myCdbl(grow.Cells(colConvAmt).Value)

                    objTr.arrSrItem = TryCast(grow.Tag, List(Of clsSerializeInvenotry))
                    ' done by priti BHA/28/06/18-000113
                    objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    objTr.Distributor_Commission_PKID = clsCommon.myCstr(grow.Cells(ColDCPKID).Value)
                    objTr.Distributor_Commission_Rate = clsCommon.myCdbl(grow.Cells(ColDCRate).Value)
                    objTr.Distributor_Commission_Amt = clsCommon.myCdbl(grow.Cells(ColDCAmt).Value)
                    objTr.Transporter_Commission_Rate = clsCommon.myCdbl(grow.Cells(ColTCRate).Value)
                    objTr.Transporter_Commission_Amt = clsCommon.myCdbl(grow.Cells(ColTCAmt).Value)
                    objTr.Security_Amt = clsCommon.myCdbl(grow.Cells(ColSCAmt).Value)
                    objTr.Security_Rate = clsCommon.myCdbl(grow.Cells(ColSCRate).Value)
                    objTr.Security_Amt = clsCommon.myCdbl(grow.Cells(ColSCAmt).Value)
                    objTr.Transporter = clsCommon.myCstr(grow.Cells(colTransporter).Value)
                    DCTotalAmt += objTr.Distributor_Commission_Amt
                    TCTotalAmt += objTr.Transporter_Commission_Amt
                    SCTotalAmt += objTr.Security_Amt
                    objTr.Distributor_Commission_RateWithTax = clsCommon.myCdbl(grow.Cells(ColDCRateWithTax).Value)
                    'sanjay Ticket -ALF/07/06/19-000107,Allow zero qty in case of Misc Item
                    If (clsCommon.myLen(objTr.Item_Code) > 0 AndAlso (clsCommon.myCdbl(objTr.ActuaQty) > 0 Or clsCommon.myCdbl(objTr.DamageQty) > 0) And (clsCommon.myCdbl(objTr.Qty) > 0 Or clsCommon.CompairString(clsCommon.myCstr(objTr.Row_Type), "Misc") = CompairStringResult.Equal)) Then
                        obj.Arr.Add(objTr)
                        intLineNo += 1
                    End If
                Next
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Fill at list one Item", Me.Text)
                    Return
                End If
                obj.Distributor_Commission_TotalAmt = DCTotalAmt
                obj.Transporter_Commission_TotalAmt = TCTotalAmt
                obj.Security_TotalAmt = SCTotalAmt
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
                    obj.ConvRate = clsCommon.myCdbl(Me.txtConversionRate.Text)
                    If clsCommon.myLen(txtApplicableFrom.Text) > 0 Then
                        obj.ApplicableFrom = Me.txtApplicableFrom.Text
                    Else
                        obj.ApplicableFrom = Nothing
                    End If
                Else
                    obj.CURRENCY_CODE = Nothing
                    obj.ConvRate = 1
                    obj.ApplicableFrom = Nothing
                End If
                '' end CurrencyConversion

                If obj.SaveData(obj, isNewEntry) Then
                    UcAttachment1.SaveData(obj.Document_Code)
                    If ChekPostBtn = False Then
                        common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                    End If
                    clsApply_Approval.CheckApprovalRequired(MyBase.Form_ID, clsCommon.myCstr(obj.Document_Code), txtDate.Text, clsCommon.myCstr(txtDesc.Text), clsCommon.myCstr(txtComment.Text), clsCommon.myCdbl(lblTotRAmt.Text), clsCommon.myCdbl(totalqty), "")
                    LoadData(obj.Document_Code, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try

            Dim obj As New clsShipmentReturnHead()
            obj = clsShipmentReturnHead.GetData(strCode, NavTyep)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
                ''richa ERO/06/07/21-001438
                If clsCommon.CompairString(obj.Is_Taxable, "1") = CompairStringResult.Equal Then
                    rbtn_Ambient.IsChecked = True
                Else
                    rbtn_Fresh.IsChecked = True
                End If
                btnSave.Enabled = True
                btnPost.Enabled = True
                btnDelete.Enabled = True
                isInsideLoadData = True
                isNewEntry = False
                btnSave.Text = "Update"
                BlankAllControls()
                LoadBlankGrid()
                'LoadBlankGridBooth()
                LoadBlankGridTax()
                LoadBlankGridAC()
                cboItemType.Enabled = False
                txtBillToLocation.Enabled = False

                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                    repoComplete.IsVisible = True
                    repoBalQty.IsVisible = True
                    btnCancel.Enabled = True

                Else
                    btnCancel.Enabled = False
                End If
                'If obj.Booth_Arr IsNot Nothing AndAlso obj.Booth_Arr.Count > 0 Then
                '    chkBoothWiseReturn.Checked = True
                'End If
                txtShippedCan.Value = obj.ShippedCAN
                TxtTotalCAN.Value = obj.TotalCAN
                txtCrateQty.Value = obj.CrateQty
                chkCncelDSR.Enabled = False
                chkCncelDSR.Checked = IIf(obj.Is_Cancelled = 1, True, False)
                txtCrate.Value = obj.CrateMan
                txtJaali.Value = obj.jaali
                txtBox.Value = obj.Box
                chkVendorGrossReceipt.Checked = clsVendorMaster.isGrossReceipt(obj.Customer_Code)
                UsLock1.Status = obj.Status
                txtInvoiceType.Text = obj.Invoice_Type
                txtDocNo.Value = obj.Document_Code
                txtDate.Value = obj.Document_Date
                txtVendorNo.Value = obj.Customer_Code
                txtPONo.Text = obj.Cust_PO_No
                'txtDate.Enabled = False
                txtVendorNo.Enabled = False
                chkRateUserCustomer.ToggleState = ClsUserCustomerSettings.GetUserCustomerRateSetting(txtVendorNo.Value)
                If clsCommon.CompairString(obj.Shift_Type, "Morning") = CompairStringResult.Equal Then
                    rbtnMorning.IsChecked = True
                ElseIf clsCommon.CompairString(obj.Shift_Type, "Evening") = CompairStringResult.Equal Then
                    rbtnEvening.IsChecked = True
                End If

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
                chkOnHold.Checked = obj.On_Hold
                txtDesc.Text = obj.Description
                txtTaxGroup.Value = obj.Tax_Group

                txtComment.Text = obj.Comments
                txtShipToLocation.Value = obj.Ship_To_Location
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
                txtDueDate.Value = obj.Due_Date
                lblAmtWithDiscount.Text = clsCommon.myFormat(obj.Discount_Base)
                lblDiscountAmt.Text = clsCommon.myFormat(obj.Discount_Amt)
                lblAmtAfterDiscount.Text = clsCommon.myFormat(obj.Amount_Less_Discount)
                lblTaxAmt.Text = clsCommon.myFormat(obj.Total_Tax_Amt)
                lblTotRAmt.Text = clsCommon.myFormat(obj.Total_Amt)
                lblTotRAmt1.Text = lblTotRAmt.Text
                TxtRoundoff.Text = clsCommon.myFormat(obj.RoundOffAmount)
                lblBillToLocation.Text = obj.BillToLocationName
                lblShipToLocation.Text = obj.ShipToLocationName
                lblTaxGrpName.Text = obj.TaxGroupName
                lblTermName.Text = obj.TermsName

                txtCarrier.Text = obj.Carrier
                txtVehicleNo.Text = obj.VehicleNo
                'txtGRNo.Text = obj.GRNo
                txtGENo.Text = obj.GENo
                If obj.GEDate.HasValue Then
                    txtGEDate.Value = obj.GEDate
                    txtGEDate.Checked = True
                End If
                txtSalesman.Value = obj.Salesman_Code
                lblSalesman.Text = obj.Salesman_Name
                txtVehicleNo.Text = obj.VehicleNo
                txtVehcileCode.Value = obj.Vehicle_Code
                txtReqNo.Value = obj.Against_Invoice_No
                txtRouteNo.Value = obj.Route_No
                lblRouteDesc.Text = obj.Route_Desc
                txtPriceCode.Text = obj.Price_Code
                txtPriceGroupCode.Text = obj.Price_Group_Code
                fndProject.Text = obj.PROJECT_ID
                lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Text + "'")
                txtSubLocation.Value = obj.Sub_Location_code
                If clsCommon.myLen(txtReqNo.Value) > 0 Then
                    SaleInvoiceDate = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select document_date from TSPL_SD_SALE_INVOICE_HEAD where TSPL_SD_SALE_INVOICE_HEAD.Document_Code='" + txtReqNo.Value + "'"))
                    txtTaxGroup.Enabled = False
                End If
                If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                    lblSubLocation.Text = clsLocation.GetName(txtSubLocation.Value, Nothing)
                Else
                    lblSubLocation.Text = ""
                End If
                txtDiscPer.Text = obj.HeadDisc_Per
                txtDiscAmt.Text = obj.HeadDisc_Amt
                ddlReturnType.SelectedValue = obj.Return_Type
                If clsCommon.CompairString(obj.Damage_Type, "0") = CompairStringResult.Equal Then
                    rbtn_leak.Checked = True
                Else
                    rbtn_expire.Checked = True
                End If
                If clsCommon.myLen(txtDiscAmt.Text) <= 0 OrElse clsCommon.myLen(txtDiscPer.Text) <= 0 OrElse clsCommon.myCdbl(txtDiscAmt.Text) = 0 OrElse clsCommon.myCdbl(txtDiscPer.Text) = 0 Then
                    txtDiscPer.Text = obj.HeadDisc_Per
                    If clsCommon.myCdbl(txtDiscPer.Text) = 0 Then
                        txtDiscAmt.Text = obj.HeadDisc_Amt
                        chkDiscountOnAmt.IsChecked = True
                        lblInvoiceDiscAmt.Text = obj.HeadDisc_Amt
                    Else
                        chkDiscountOnRate.IsChecked = True
                        lblInvoiceDiscAmt.Text = obj.HeadDisc_PerAmt
                    End If
                End If
                FlagDocumentIsTaxable = obj.Is_Taxable
                EInvoiceType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_SD_SALE_RETURN_HEAD", "Document_Code", obj.Document_Code, Nothing)
                If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                    btnReverseAndUnpost.Enabled = False
                    If obj.Status = ERPTransactionStatus.Approved Then
                        btnCancel.Enabled = True
                    ElseIf obj.Status = ERPTransactionStatus.Pending Then
                        btnCancel.Enabled = False
                    End If
                End If
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

                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsShipmentReturnDetail In obj.Arr
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDelivery_Code).Value = objTr.Delivery_Code
                        gv1.Rows(gv1.Rows.Count - 1).Tag = objTr.arrSrItem
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCrate).Value = objTr.Crate
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCan).Value = objTr.CAN
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAlterUnitQty).Value = objTr.Alter_UnitQty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRateUnitQty).Value = objTr.Rate_UnitQty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colInvoiceSchemeCode).Value = objTr.InvoiceScheme_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCash_Amt).Value = objTr.Cash_Scheme_Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCash_Pers).Value = objTr.Cash_Scheme_Pers
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCashSchemeCode).Value = objTr.Cash_Scheme_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colCashSchemeType).Value = objTr.Cash_Scheme_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchmCodeType).Value = objTr.Scheme_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMainIcode).Value = objTr.Scheme_Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMainIQty).Value = objTr.Scheme_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colMainIUOM).Value = objTr.Scheme_Item_UOM

                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitRate).Value = objTr.RATE_UOM
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitALter).Value = objTr.Alternate_UOM
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTAX_PAID).Value = objTr.TAX_PAID
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICodeGrp).Value = objTr.Item_Group
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemBasicPrice).Value = objTr.Price_with_tax

                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = objTr.Line_No
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = objTr.Row_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colComplete).Value = IIf(objTr.Status = 0, "No", "Yes")
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objTr.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(objTr.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(objTr.Item_Code)

                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(objTr.Item_Code)
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgSOQty).Value = objTr.so_Qty

                        gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = objTr.Balance_Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = objTr.Qty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDamageQty).Value = objTr.DamageQty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFreeQty).Value = objTr.Free_Qty

                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objTr.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgUnit).Value = objTr.OrgUnit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = objTr.Shipment_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objTr.Item_Cost
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationCode).Value = objTr.Location
                        ''gv1.Rows(gv1.Rows.Count - 1).Cells(colLocationName).Value = objTr.LocationName
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDamageAmt).Value = objTr.Damage_Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colReturnAmt).Value = objTr.Return_Amount
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
                        'If clsCommon.myLen(objTr.Invoice_Code) > 0 Then
                        '    'gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = clsPSInvoiceHeadDetail.GetBalanceSRNQty(objTr.Invoice_Code, objTr.Item_Code, obj.Document_Code, objTr.Unit_code, objTr.MRP, objTr.Assessable)
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
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value))
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
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHeaDDisPer).Value = objTr.HeadDiscPer
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDisPerAmt).Value = objTr.HeadDiscPerAmt
                        'Batch wise item return Ticket No- ALF/22/05/18-000066
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(objTr.Item_Code)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = objTr.arrBatchItem
                        'Batch wise item return Ticket No- ALF/22/05/18-000066
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colActualUOM).Value = objTr.ActualUOM
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colActualQty).Value = objTr.ActuaQty
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colConvAmt).Value = objTr.ActualReturnQty

                        'If clsCommon.CompairString(gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value, "No") = CompairStringResult.Equal Then
                        '    'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualUOM).ReadOnly = False
                        '    'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualQty).ReadOnly = False
                        '    gv1.Rows(gv1.Rows.Count - 1).Cells(colActualRetQty).ReadOnly = False
                        'Else

                        '    'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualUOM).ReadOnly = True
                        '    'gv1.Rows(gv1.Rows.Count - 1).Cells(colActualQty).ReadOnly = True
                        '    gv1.Rows(gv1.Rows.Count - 1).Cells(colActualRetQty).ReadOnly = True
                        'End If

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
                        '''''''''''''' Distributor Commission Detail '''''''''''''''''''''''''''''''''
                        If clsCommon.myLen(objTr.Distributor_Commission_PKID) > 0 Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCPKID).Value = objTr.Distributor_Commission_PKID
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCRate).Value = objTr.Distributor_Commission_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCRateWithTax).Value = objTr.Distributor_Commission_RateWithTax
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCAmt).Value = objTr.Distributor_Commission_Amt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCApplicableDate).Value = clsDBFuncationality.getSingleValue("select TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date from TSPL_DISTRIBUTOR_COMMISSION_HEAD where Doc_No=(select Doc_No from TSPL_DISTRIBUTOR_COMMISSION_DETAIL where PK_ID='" + clsCommon.myCstr(objTr.Distributor_Commission_PKID) + "')")
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCUOM).Value = clsDBFuncationality.getSingleValue("select TSPL_DISTRIBUTOR_COMMISSION_HEAD.Commision_UOM from TSPL_DISTRIBUTOR_COMMISSION_HEAD where Doc_No=(select Doc_No from TSPL_DISTRIBUTOR_COMMISSION_DETAIL where PK_ID='" + clsCommon.myCstr(objTr.Distributor_Commission_PKID) + "')")
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCAmt).Value = objTr.Distributor_Commission_Amt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColTCRate).Value = objTr.Transporter_Commission_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColTCAmt).Value = objTr.Transporter_Commission_Amt
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCUnitCF).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" + clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value) + "' and Item_Code='" + clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) + "'")
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCCFUOM).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" + clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCUOM).Value) + "' and Item_Code='" + clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) + "'")
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCQtyinSU).Value = (gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value * gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCUnitCF).Value) / gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCCFUOM).Value
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColSCRate).Value = objTr.Security_Rate
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColSCAmt).Value = objTr.Security_Amt
                        End If
                        '''''''''''''' End of Distributor Commission Detail '''''''''''''''''''''''''''''''''
                    Next
                    'If obj.Status = ERPTransactionStatus.Pending Then
                    '    gv1.Rows.AddNew()
                    '    gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                    '    gvAC.Rows.AddNew()
                    'End If
                End If

                SetitemWiseTaxOnlySetting()
                ' ''RefreshReqNo()
                ' ''RefreshGRPNo()

                ''For Custom Fields
                ''richa agarwal BM00000008667
                If clsCommon.CompairString(obj.Status, "1") = CompairStringResult.Equal Then
                    txtDate.Enabled = False
                Else
                    txtDate.Enabled = True
                End If


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
                UcAttachment1.LoadData(obj.Document_Code)

                btnPost.Visible = MyBase.isPostFlag
                If Not clsApply_Approval.Visibility_PostButtonForApproval(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value), clsCommon.myCdbl(lblTotRAmt.Text), 0, "") Then
                    btnPost.Visible = False
                    If UsLock1.Status = ERPTransactionStatus.Pending Then
                        UsLock1.Status = clsApply_Approval.ApprovalCondCheck_Doc(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value), Nothing)
                    End If
                End If
            End If
            If UOMAtDiarySaleReturn = True Then
                gv1.Columns(colActualQty).IsVisible = True
                gv1.Columns(colActualUOM).IsVisible = True
                gv1.Columns(colConvAmt).IsVisible = True
                gv1.Columns(colQty).ReadOnly = True
                gv1.Columns(colUnit).ReadOnly = True
                'Else
                '    gv1.Columns(colActualQty).IsVisible = False
                '    gv1.Columns(colActualUOM).IsVisible = False
                '    gv1.Columns(colConvAmt).IsVisible = False

            End If
            If clsCommon.myLen(txtReqNo.Value) > 0 Then
                btnInvoiceJE.Visible = True
            Else
                btnInvoiceJE.Visible = False
            End If
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

    'Private Sub gv2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv2.DoubleClick
    '    Try
    '        If rbtnTaxCalAutomatic.IsChecked Then
    '            Dim qry As String = "select Tax_Rate_Code as [Rate Code],Tax_Rate_Desc as [Rate Description],Tax_Rate as [Rate] from TSPL_TAX_RATES "
    '            Dim dblNewRate As Double = clsCommon.myCdbl(clsCommon.ShowSelectForm("PSSaleReturnfnddxRate", qry, "Rate", "Tax_Code='" + clsCommon.myCstr(gv2.CurrentRow.Cells(colTTaxAutCode).Value) + "' and Tax_Type='S'", "", "", True))
    '            Dim intRowNo As Integer = gv2.CurrentRow.Index
    '            If gv1.RowCount > 0 AndAlso intRowNo >= 0 Then
    '                Dim strII As String = clsCommon.myCstr(intRowNo + 1)
    '                For ii As Integer = 0 To gv1.Rows.Count - 1
    '                    gv1.Rows(ii).Cells("COLTAXRATE" + strII).Value = dblNewRate
    '                Next
    '            End If
    '            For ii As Integer = 0 To gv1.Rows.Count - 1
    '                UpdateCurrentRow(ii)
    '            Next
    '            UpdateAllTotals()
    '        End If
    '    Catch ex As Exception
    '        common.clsCommon.MyMessageBoxShow(ex.Message)
    '    End Try
    'End Sub

    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        PostData()
    End Sub

    Sub PostData()
        Try
            Dim qry As String = ""
            Dim msg As String = ""
            Dim dt As DataTable = Nothing
            If (myMessages.postConfirm()) Then
                SaveData(True)
                '' Anubhooti 13-Sep-2014 BM00000003735
                If FrmMainTranScreen.ValidateTransactionAccToFinYear("Sale Return", txtDate.Value) = False Then
                    Exit Sub
                End If
                ''
                If (clsShipmentReturnHead.PostData(MyBase.Form_ID, txtDocNo.Value)) Then

                    msg = "Successfully Posted"
                Else
                    qry = "select No_Of_Level, LEVEL from TSPL_APPROVAL_LEVEL_SCREEN where User_Code='" + objCommonVar.CurrentUserCode + "' and Trans_Code='" + MyBase.Form_ID + "' "
                    dt = clsDBFuncationality.GetDataTable(qry)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        Dim level As String = dt.Rows(0)("LEVEL").ToString()
                        Dim NoOflevel As Integer = clsCommon.myCdbl(dt.Rows(0)("No_Of_Level"))
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
                            msg = "Level 3 Approval done. Successfully Posted. "
                        End If
                    End If
                End If

                common.clsCommon.MyMessageBoxShow(msg)
                LoadData(txtDocNo.Value, NavigatorType.Current)

                If clsSMSAtPost_Sales.SMSATPOST_SALE() Then
                    SMSSENDONLY(True)
                End If

                'If (common.clsCommon.MyMessageBoxShow(Me, "Do you want to print", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes) Then
                '    funPrint(txtDocNo.Value)
                'End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub CalAltQty(ByVal CurrentRow As Integer)
        Try
            ''richa BHA/30/10/18-000649
            If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(CurrentRow).Cells(colActualUOM).Value)) <= 0 Then
                gv1.Rows(CurrentRow).Cells(colActualQty).Value = 0
                Throw New Exception("Please select Actual UOM first.")
            End If
            ''-------------------
            Dim qty As Decimal = clsCommon.myCdbl(gv1.Rows(CurrentRow).Cells(colQty).Value)
            Dim ActualQty As Double = clsCommon.myCdbl(gv1.Rows(CurrentRow).Cells(colActualQty).Value)
            Dim Actualuom As String = clsCommon.myCstr(gv1.Rows(CurrentRow).Cells(colActualUOM).Value)
            Dim uom As String = clsCommon.myCstr(gv1.Rows(CurrentRow).Cells(colUnit).Value)

            Dim UnitAmt As Double = clsCommon.myCdbl(gv1.Rows(CurrentRow).Cells(colActualCost).Value)

            Dim Actconversion As Decimal = 0
            Actconversion = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select conversion_factor from TSPL_ITEM_UOM_DETAIL where item_code='" + clsCommon.myCstr(gv1.Rows(CurrentRow).Cells(colICode).Value) + "' and uom_code='" + Actualuom + "'"))

            'If altconversion <= 0 Then
            '    altconversion = 1
            'End If

            Dim conversion As Decimal = 0
            conversion = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select conversion_factor from TSPL_ITEM_UOM_DETAIL where item_code='" + clsCommon.myCstr(gv1.Rows(CurrentRow).Cells(colICode).Value) + "' and uom_code='" + uom + "'"))
            'If conversion <= 0 Then
            '    conversion = 1
            'End If

            Dim altqty As Decimal = 0
            Dim finalqty As Decimal = 0

            If Actconversion > 0 Then
                'altqty = System.Math.Round((ActualQty * Actconversion) / conversion, 3)
                altqty = ((ActualQty * Actconversion) / conversion)
                finalqty = System.Math.Round((UnitAmt * Actconversion) / conversion, 3)
            Else
                altqty = 0
            End If

            If Not clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeItem).Value), "Yes") = CompairStringResult.Equal Then
                gv1.Rows(CurrentRow).Cells(colQty).Value = System.Math.Round(altqty, 5)
                gv1.Rows(CurrentRow).Cells(colConvAmt).Value = finalqty
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        DeleteData()
    End Sub
    Sub DeleteData()
        Try
            Dim Reason As String = ""
            If (myMessages.deleteConfirm()) Then
                clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value))
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
                If (clsShipmentReturnHead.DeleteData(txtDocNo.Value)) Then
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
                Dim autoFilter As FilterDescriptor = New FilterDescriptor("Name", FilterOperator.StartsWith, "")
                autoFilter.IsFilterEditor = True
                editor.EditorControl.FilterDescriptors.Add(autoFilter)
            End If
        End If
    End Sub

    'Private Function funSetUserAccess() As Boolean
    '    Try
    '        btnSave.Visible = True
    '        btnDelete.Visible = True
    '        btnPost.Visible = True

    '        Dim strRights As String
    '        Dim strTemp() As String
    '        Dim strProgCode = "PO-SRN"
    '        strRights = enuUserRights.enuRead & "," & enuUserRights.enuModify & "," & enuUserRights.enuDelete & "," & enuUserRights.enuAuthorised
    '        strRights = modUserMgt.funGetPermissions(strRights, strProgCode)
    '        strTemp = Split(strRights, ",")
    '        If strTemp(0) = "0" Then
    '            MsgBox("Permission Denied", MsgBoxStyle.Critical, Me.Text)
    '            funSetUserAccess = False
    '            blnRead = False
    '            Me.Close()
    '            Exit Function
    '        Else
    '            blnRead = True
    '        End If
    '        If strTemp(1) = "0" Then 'Grant modify access
    '            btnSave.Visible = False
    '        End If
    '        If strTemp(2) = "0" Then 'Grant modify access
    '            btnDelete.Visible = False
    '        End If
    '        If strTemp(3) = "0" Then 'Grant Authorize access
    '            btnPost.Visible = False
    '        End If
    '        funSetUserAccess = True
    '    Catch er As Exception

    '    End Try
    'End Function

    Private Sub txtDocNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocNo._MYNavigator
        Try
            ''-------richa 30/08/2014 Ticket No. BM00000003242---------
            Dim strwherecls As String = ""
            Dim strcondition As String = ""
            strwherecls = Xtra.CustomerPermission()
            If clsCommon.myLen(strwherecls) > 0 Then
                'strcondition = "and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + strwherecls + ") and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "'"
                strcondition = "and TSPL_SD_Shipment_RETURN_HEAD.Customer_Code in (" + strwherecls + ") "
            End If

            '-----------------------------------------------------
            ' Dim qst As String = "select count(*) from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" + txtDocNo.Value + "' " + strcondition + " and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type='DS' "
            Dim qst As String = "select count(*) from TSPL_SD_Shipment_RETURN_HEAD where Document_Code='" + txtDocNo.Value + "' " + strcondition + "  and TSPL_SD_Shipment_RETURN_HEAD.Screen_Type='DS' "
            Dim count As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qst))
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
        '-------richa 12/08/2014 Ticket No. BM00000003242---------
        Dim strwherecls As String = ""
        strwherecls = Xtra.CustomerPermission()
        '-----------------------------------------------------

        Dim qry As String = "select Document_Code as Code,CONVERT(varchar(10), Document_Date,103)+' '+ CONVERT(varchar(5), Document_Date,114) as Date,isnull(Against_Invoice_No,'') as [Against Invoice No],Customer_Code as [Customer Code], Customer_Name as Customer,ISNULL(TSPL_CUSTOMER_MASTER.Alies_Name,'') As [Alies Name],Total_Amt as Amount,case when TSPL_SD_Shipment_RETURN_HEAD.Status=0 then 'Pending' else 'Approved' end as [Status] "
        'Sanjay Ticket No- ALF/03/07/18-000075 Add column Is cancelled in the finder
        qry += " ,case when isnull(TSPL_SD_Shipment_RETURN_HEAD.is_cancelled,0) = 0 then 'No' else 'Yes' end as [Is Cancelled] "
        'Sanjay Ticket No- ALF/03/07/18-000075 Add column Is cancelled in the finder
        qry += " from TSPL_SD_Shipment_RETURN_HEAD left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_Shipment_RETURN_HEAD.Customer_Code "


        Dim whrClas As String = ""

        '-------richa 12/08/2014 Ticket No. BM00000003242---------
        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        'End If

        'If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 And clsCommon.myLen(strwherecls) > 0 Then
        '    whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type='DS' and TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + strwherecls + ") "
        'ElseIf clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
        '    whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ") and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type='DS' "
        'ElseIf clsCommon.myLen(strwherecls) > 0 Then
        '    whrClas = " TSPL_SD_SALE_RETURN_HEAD.Customer_Code in (" + strwherecls + ") and TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type='DS' "
        'Else
        '    whrClas = " TSPL_SD_SALE_RETURN_HEAD.Trans_Type='" + IIf(rbtn_Fresh.IsChecked = True, "FS", "PS") + "' and TSPL_SD_SALE_RETURN_HEAD.Screen_Type='DS' "
        'End If

        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 AndAlso clsCommon.myLen(strwherecls) > 0 Then
            whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")  and TSPL_SD_Shipment_RETURN_HEAD.Screen_Type='DS' and TSPL_SD_Shipment_RETURN_HEAD.Customer_Code in (" + strwherecls + ") "
        ElseIf clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrClas = " Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")  and TSPL_SD_Shipment_RETURN_HEAD.Screen_Type='DS' "
        ElseIf clsCommon.myLen(strwherecls) > 0 Then
            whrClas = " TSPL_SD_Shipment_RETURN_HEAD.Customer_Code in (" + strwherecls + ")  and TSPL_SD_Shipment_RETURN_HEAD.Screen_Type='DS' "
        Else
            whrClas = "  TSPL_SD_Shipment_RETURN_HEAD.Screen_Type='DS' "
        End If

        '-----------------------------------------------------

        LoadData(clsCommon.ShowSelectForm("PSSaleReturnDocfnd", qry, "Code", whrClas, txtDocNo.Value, "Code", isButtonClicked, " TSPL_SD_Shipment_RETURN_HEAD.Document_date "), NavigatorType.Current)
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
        'If e.KeyCode = Keys.Enter AndAlso gvBooth.CurrentCell IsNot Nothing AndAlso gvBooth.CurrentColumn Is gvBooth.Columns(colBoothCode) Then
        '    isCellValueChangedOpen = True
        '    gvBooth.CurrentColumn = gvBooth.Columns(colBoothCode)
        '    setBoothGridFocus()
        '    isCellValueChangedOpen = False
        'End If
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
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control AndAlso e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then

                'Add Tool tip Task No- TEC/18/05/18-000237
                ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                                             "TSPL_SD_SHIPMENT_RETURN_HEAD " + Environment.NewLine +
                                             "TSPL_SD_SALE_RETURN_DETAIL " + Environment.NewLine +
                                             "TSPL_BATCH_ITEM ( If Item is batch type) " + Environment.NewLine +
                                             "TSPL_SERIAL_ITEM ( If Item is Serial type)" + Environment.NewLine +
                                             "TSPL_Customer_Invoice_Head ( For AR Credit Note Entry - After Posting)  " + Environment.NewLine +
                                             "TSPL_Customer_Invoice_Detail( After Posting)  " + Environment.NewLine +
                                             "TSPL_JOURNAL_MASTER (Journal Voucher Entry - For dispatch and invoice  - After Posting )  " + Environment.NewLine +
                                             "TSPL_JOURNAL_DETAILS ( After Posting) " + Environment.NewLine +
                                             "TSPL_INVENTORY_MOVEMENT  ( After Posting) ")
                'Add Tool tip Task No- TEC/18/05/18-000237

                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIRC
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverseAndUnpost.Visible = True
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
            End If
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control AndAlso e.KeyCode = Keys.F10 Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.SIRC
            frm.strCode = clsFixedParameterCode.AllowBoothWiseReturn
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                ApplyBoothWiseReturn = True
                clsCommon.MyMessageBoxShow(Me, "Booth Wise Sale Return activated")
                AddNew()
            End If
        End If
    End Sub

    Private Sub txtTermCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTermCode._MYValidating
        txtTermCode.Value = ClsReceivablePaymentTerms.getFinderWithSaleType(txtTermCode.Value, "P", isButtonClicked)
        lblTermName.Text = ClsReceivablePaymentTerms.GetName(txtTermCode.Value)
        SetTermDetails()
    End Sub

    Sub SetTermDetails()
        Dim qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_RECEIVABLE_PAYMENT_TERMS_MASTER where Terms_Code='" + txtTermCode.Value + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Description"))
            txtDueDate.Value = txtDate.Value.AddDays(clsCommon.myCdbl(dt.Rows(0)("No Of Days")))
        Else
            lblTermName.Text = ""
        End If
    End Sub


    Private Sub txtTaxGroup__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTaxGroup._MYValidating


        If clsCommon.myLen(txtBillToLocation.Value) > 0 Then

            If clsERPFuncationality.GetGSTStatus(txtDate.Value) AndAlso objCommonVar.GSTActiveTaxGroup Then
                Dim qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER"
                txtTaxGroup.Value = clsCommon.ShowSelectForm("SaleReturnfndid", qry, "Code", "Tax_Group_Type='S' AND Active=1", txtTaxGroup.Value, "Code", isButtonClicked)
                SetTaxDetails()
            Else

                Dim qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
                txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtTaxGroup.Value, isButtonClicked)
                'For ii As Integer = 0 To gv1.Rows.Count - 1
                '    If clsCommon.CompairString(gv1.Rows(ii).Cells(colICodeGrp).Value, "CPD-DESI GHEE") = CompairStringResult.Equal AndAlso clsCommon.CompairString(gv1.Rows(ii).Cells(colTAX_PAID).Value, "Yes") = CompairStringResult.Equal Then
                '        gv1.Rows(ii).Cells(colRate).Value = gv1.Rows(ii).Cells(colManualRate).Value
                '    End If
                'Next
                SetTaxDetails()
            End If
        Else
            common.clsCommon.MyMessageBoxShow(Me, "Please select Location First", Me.Text)
        End If

    End Sub

    Sub SetTaxDetails()
        LoadBlankGridTax()
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsByLocation(txtTaxGroup.Value, "S", txtVendorNo.Value, txtBillToLocation.Value)
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
                    If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Is_TCS from tspl_tax_master where Tax_Code='" & clsCommon.myCstr(dr("Tax_Code")) & "'")), "Y") = CompairStringResult.Equal) AndAlso Not (SaleInvoiceDate.Month() = txtDate.Value.Month() And SaleInvoiceDate.Year() = txtDate.Value.Year()) Then
                        gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = 0
                    Else
                        gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dr("TaxRate"))
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
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsByLocation(txtTaxGroup.Value, "S", txtVendorNo.Value, txtBillToLocation.Value)

        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If isForCurrentRow Then
                BlankTaxDetails(gv1.CurrentRow.Index, isChangeRate)
                If clsCommon.myLen(gv1.CurrentRow.Cells(colICode)) > 0 Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                        If isChangeRate Then
                            If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Is_TCS from tspl_tax_master where Tax_Code='" & clsCommon.myCstr(dr("Tax_Code")) & "'")), "Y") = CompairStringResult.Equal) AndAlso Not (SaleInvoiceDate.Month() = txtDate.Value.Month() And SaleInvoiceDate.Year() = txtDate.Value.Year()) Then
                                gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = 0
                            Else
                                gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))
                            End If
                        End If
                        If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Is_TCS from tspl_tax_master where Tax_Code='" & clsCommon.myCstr(dr("Tax_Code")) & "'")), "Y") = CompairStringResult.Equal) AndAlso Not (SaleInvoiceDate.Month() = txtDate.Value.Month() And SaleInvoiceDate.Year() = txtDate.Value.Year()) Then
                            gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = 0
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
                    If clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode)) > 0 Then
                        Dim ii As Integer = 1
                        For Each dr As DataRow In dt.Rows
                            Dim strII As String = clsCommon.myCstr(ii)
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                            If isChangeRate Then
                                If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Is_TCS from tspl_tax_master where Tax_Code='" & clsCommon.myCstr(dr("Tax_Code")) & "'")), "Y") = CompairStringResult.Equal) AndAlso Not (SaleInvoiceDate.Month() = txtDate.Value.Month() And SaleInvoiceDate.Year() = txtDate.Value.Year()) Then
                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = 0
                                Else
                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))
                                End If
                            End If
                            If (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Is_TCS from tspl_tax_master where Tax_Code='" & clsCommon.myCstr(dr("Tax_Code")) & "'")), "Y") = CompairStringResult.Equal) AndAlso Not (SaleInvoiceDate.Month() = txtDate.Value.Month() And SaleInvoiceDate.Year() = txtDate.Value.Year()) Then
                                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = 0
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
        ''Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsscrap(txtTaxGroup.Value)
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsByLocation(txtTaxGroup.Value, "S", txtVendorNo.Value, txtBillToLocation.Value)
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
        If clsCommon.myLen(txtBillToLocation.Value) = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select Location first", Me.Text)
            Exit Sub
        End If
        ''-------richa 12/08/2014 Ticket No. BM00000003242---------
        Dim strwherecls As String = ""
        Dim strwhrcondition As String = ""
        strwherecls = Xtra.CustomerPermission()
        '-----------------------------------------------------
        Dim qry As String = "select Cust_Code as Code,Customer_Name as Name,ISNULL(TSPL_CUSTOMER_MASTER.Alies_Name,'') As [Alies Name],TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER.City_Name)>0 then ', '+TSPL_CITY_MASTER.City_Name else ' ' end + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Address,TSPL_CUSTOMER_MASTER.Terms_Code as [Term Code] , TSPL_RECEIVABLE_PAYMENT_TERMS_MASTER.Terms_Desc as [Term Description] ,Tax_Group as [Tax Group],Tax_Group_Desc as [Tax Group Description],Salesman_Code as [Salesman Code],Salesman_Desc as Salesman   "
        qry += " from TSPL_CUSTOMER_MASTER "
        qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code"
        qry += " left outer join TSPL_TDS_STATE_MASTER on TSPL_TDS_STATE_MASTER.State_Code=TSPL_CUSTOMER_MASTER.State"
        qry += " left outer join TSPL_RECEIVABLE_PAYMENT_TERMS_MASTER on TSPL_RECEIVABLE_PAYMENT_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code"
        qry += " left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S'"

        '-------richa 12/08/2014 Ticket No. BM00000003242---------
        If clsCommon.myLen(strwherecls) > 0 Then
            strwhrcondition = " TSPL_CUSTOMER_MASTER.Cust_Code in (" + strwherecls + ")"
        End If
        '-----------------------------------------------------'
        txtVendorNo.Value = clsCommon.ShowSelectForm("PSSaleRetCustfnd", qry, "Code", strwhrcondition, txtVendorNo.Value, "Code", isButtonClicked)
        qry += " where 2=2 and TSPL_CUSTOMER_MASTER.Cust_Code ='" + txtVendorNo.Value + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            lblVendorName.Text = clsCommon.myCstr(dt.Rows(0)("Name"))
            txtTermCode.Value = clsCommon.myCstr(dt.Rows(0)("Term Code"))
            lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Term Description"))
            txtTaxGroup.Value = clsCommon.myCstr(dt.Rows(0)("Tax Group"))
            lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax Group Description"))
            txtSalesman.Value = clsCommon.myCstr(dt.Rows(0)("Salesman Code"))
            lblSalesman.Text = clsCommon.myCstr(dt.Rows(0)("Salesman"))

            '  txtDate.Enabled = False
            txtVendorNo.Enabled = False
            chkRateUserCustomer.ToggleState = ClsUserCustomerSettings.GetUserCustomerRateSetting(txtVendorNo.Value)
            SetMultiCurrencyVisibility()
        Else
            lblVendorName.Text = ""
            txtTermCode.Value = ""
            lblTermName.Text = ""
            txtTaxGroup.Value = ""
            lblTaxGrpName.Text = ""

            txtSalesman.Value = ""
            lblSalesman.Text = ""
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
            If clsCommon.myLen(txtPriceCode.Text) = 0 Then
                txtPriceGroupCode.Text = clsCommon.myCstr(dt.Rows(0)("price_group_code"))
            End If
            txtVendorNo.Enabled = False

            If clsCommon.CompairString(clsCommon.myCstr(dtLocation.Rows(0)("State")), clsCommon.myCstr(dt.Rows(0)("State"))) = CompairStringResult.Equal Then
                txtTaxGroup.Value = clsCommon.myCstr(dtLocation.Rows(0)("LocalTaxGroup"))
                lblTaxGrpName.Text = clsCommon.myCstr(dtLocation.Rows(0)("Local_Tax_GroupName"))
            Else
                txtTaxGroup.Value = clsCommon.myCstr(dtLocation.Rows(0)("InterstateTaxGroup"))
                lblTaxGrpName.Text = clsCommon.myCstr(dtLocation.Rows(0)("Interstate_Tax_GroupName"))
            End If

        End If


        '''' priti change ends here

        SetTaxDetails()
        SetTermDetails()
    End Sub
    Private Sub txtVehcileCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtVehcileCode._MYValidating
        Try
            Dim qry As String = "Select distinct  vehicle_id ,Description from TSPL_VEHICLE_MASTER"
            txtVehcileCode.Value = clsCommon.ShowSelectForm("Vehicle No", qry, "vehicle_id", "", txtVehcileCode.Value, "vehicle_id", isButtonClicked)
            txtVehicleNo.Text = connectSql.RunScalar("Select Description  from TSPL_VEHICLE_MASTER where Vehicle_Id = '" + Convert.ToString(txtVehcileCode.Value) + "'")
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtBillToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtBillToLocation._MYValidating


        Dim qry As String = "select Location_Code as Code,Location_Desc as Name,Loc_Short_Name as [Short Name] from TSPL_LOCATION_MASTER "
        Dim WhrCls As String = "  Location_Type='Physical' and Location_Category not in('MCC') and CSA_Type='N' and Is_Section='N' and Is_Sub_Location='N'  "
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            WhrCls += "  and  Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
        End If


        txtBillToLocation.Value = clsCommon.ShowSelectForm("PSSaleRetBillLoca", qry, "Code", WhrCls, txtBillToLocation.Value, "Code", isButtonClicked)
        lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtBillToLocation.Value + "'"))

    End Sub
    Private Sub fndRouteNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtRouteNo._MYValidating

        Dim qry As String = "Select Route_No as Code,Route_Desc as Description,Type,Employee_Code as 'Employee Code',Off_Day as 'Off Day' from TSPL_ROUTE_MASTER"
        txtRouteNo.Value = clsCommon.ShowSelectForm("PSSaleRetRoute", qry, "Code", "", txtRouteNo.Value, "", isButtonClicked)
        fndRouteNo_TextChanged()
    End Sub
    Private Sub fndRouteNo_TextChanged()
        Dim sql As String = "Select Route_Desc,Employee_Code from TSPL_ROUTE_MASTER where Route_No='" + txtRouteNo.Value + "'"
        Dim dr1 As DataTable = clsDBFuncationality.GetDataTable(sql)
        If dr1 IsNot Nothing AndAlso dr1.Rows.Count > 0 Then

            lblRouteDesc.Text = dr1.Rows(0)(0).ToString()
            txtSalesman.Value = dr1.Rows(0)(1).ToString()



        Else
            lblRouteDesc.Text = String.Empty
        End If
        If ApplyBoothWiseReturn Then
            Dim qry As String = "select top 1 TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code  from TSPL_DISTRIBUTOR_ROUTE_CUSTOMER
left join TSPL_DISTRIBUTOR_ROUTE on TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code=TSPL_DISTRIBUTOR_ROUTE.Code
                where TSPL_DISTRIBUTOR_ROUTE.Status=1 and TSPL_DISTRIBUTOR_ROUTE.ItemType='M'
 and TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" + txtRouteNo.Value + "' and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" + clsCommon.GetPrintDate(txtDate.Value) + "' then 2 else 3 end) end) order by Start_Date desc "
            txtVendorNo.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
            lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
            txtBillToLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
            If clsCommon.myLen(txtBillToLocation.Value) > 0 Then
                lblBillToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_Location_Master where Location_Code='" + txtBillToLocation.Value + "' "))
            End If
            SetTaxX()
        End If
    End Sub
    Private Sub txtShipToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtShipToLocation._MYValidating
        If clsCommon.myLen(txtBillToLocation.Value) = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select Location first", Me.Text)
            txtBillToLocation.Focus()
            Exit Sub
        End If
        If clsCommon.myLen(txtVendorNo.Value) = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select Customer first", Me.Text)
            txtVendorNo.Focus()
            Exit Sub
        End If

        'Dim qry As String = "select Ship_To_Code as Code,Ship_To_Desc as Description, Tin_No as [Tin No], CST_No as [CST No] from TSPL_SHIP_TO_LOCATION"
        'txtShipToLocation.Value = clsCommon.ShowSelectForm("PSSaleRetShipLoc", qry, "Code", "", txtShipToLocation.Value, "Code", isButtonClicked)
        'lblShipToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Ship_To_Desc from TSPL_SHIP_TO_LOCATION where Ship_To_Code='" + txtShipToLocation.Value + "'"))

        Dim qry As String = " select distinct TSPL_SHIP_TO_LOCATION.Ship_To_Code as [Code],TSPL_SHIP_TO_LOCATION.Ship_To_Desc as [Description],TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code as[Customer Code] ,TSPL_SHIP_TO_LOCATION.Ship_To_Type_Desc  as [Customer Discription], replace(case when ISNULL (TSPL_CUSTOMER_MASTER.Add1,'')='' then '' else TSPL_CUSTOMER_MASTER.add1 +',' end + case when ISNULL (TSPL_CUSTOMER_MASTER.Add2,'')='' then '' else TSPL_CUSTOMER_MASTER.add2 +',' end + case when ISNULL (TSPL_CUSTOMER_MASTER.Add3,'')='' then '' else TSPL_CUSTOMER_MASTER.add3 +',' end  ,',,',',') as [Customer Address] ,replace(case when ISNULL (TSPL_SHIP_TO_LOCATION.Add1,'')='' then '' else TSPL_SHIP_TO_LOCATION.add1 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add2,'')='' then '' else TSPL_SHIP_TO_LOCATION.add2 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add3,'')='' then '' else TSPL_SHIP_TO_LOCATION.add3 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add4,'')='' then '' else TSPL_SHIP_TO_LOCATION.add4 +',' end ,',,',',') as [Ship to Address],TSPL_SHIP_TO_LOCATION.CST_No as [CST NO],TSPL_SHIP_TO_LOCATION.Tin_No as [TIN No]  from TSPL_SHIP_TO_LOCATION left outer join TSPL_SHIP_TO_LOCATION_LOCATIONS on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SHIP_TO_LOCATION_LOCATIONS.Ship_To_Code left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code "
        txtShipToLocation.Value = clsCommon.ShowSelectForm("PSSaleRetShipLoc", qry, "Code", "Ship_To_Type_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' and (TSPL_SHIP_TO_LOCATION_LOCATIONS.Loc_Code ='" & txtBillToLocation.Value & "' or  TSPL_SHIP_TO_LOCATION.Loc_Code='" & txtShipToLocation.Value & "' )", txtShipToLocation.Value, "Code", isButtonClicked)
        lblShipToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Ship_To_Desc from TSPL_SHIP_TO_LOCATION where Ship_To_Code='" + txtShipToLocation.Value + "'"))

    End Sub

    Private Sub btnRequistionItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        SelectMRNItems()
    End Sub

    Sub SelectMRNItems()
        isInsideLoadData = True

        Dim frm As New frmPendingDispatch()
        frm.VendorCode = txtVendorNo.Value
        frm.strCurrCode = txtDocNo.Value
        If rbtn_Fresh.IsChecked Then
            frm.Is_Taxable = 0
        Else
            frm.Is_Taxable = 1
        End If

        frm.ShowDialog()
        LoadBlankGrid()
        Dim objOrderHead As clsPSShipmentHead = Nothing
        If frm.ArrReturn IsNot Nothing AndAlso frm.ArrReturn.Count > 0 Then
            isInvoiceLoadData = True
            If clsCommon.myLen(frm.ArrReturn(0).Document_Code) > 0 Then
                objOrderHead = clsPSShipmentHead.GetData(frm.ArrReturn(0).Document_Code, NavigatorType.Current)
                If objOrderHead IsNot Nothing AndAlso clsCommon.myLen(objOrderHead.Document_Code) > 0 Then
                    txtBillToLocation.Value = ""
                    SaleInvoiceDate = objOrderHead.Document_Date
                    '  txtShippedCan.Value = objOrderHead.ShippedCAN
                    If clsCommon.myLen(txtRefNo.Text) <= 0 Then
                        txtRefNo.Text = objOrderHead.Ref_No
                    End If
                    If clsCommon.myLen(txtDesc.Text) <= 0 Then
                        txtDesc.Text = objOrderHead.Description
                    End If
                    If clsCommon.myLen(txtPONo.Text) <= 0 Then
                        txtPONo.Text = objOrderHead.Cust_PO_No
                    End If
                    If clsCommon.myLen(txtInvNo.Text) <= 0 Then
                        'txtRemarks.Text = objOrderHead.Remarks
                    End If
                    If (clsCommon.myLen(txtShipToLocation.Value)) <= 0 Then
                        txtShipToLocation.Value = objOrderHead.Ship_To_Location
                    End If
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
                        txtInvoiceType.Text = objOrderHead.Invoice_Type
                    End If
                    txtSubLocation.Value = objOrderHead.Sub_Location_code
                    If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                        lblSubLocation.Text = clsLocation.GetName(txtSubLocation.Value, Nothing)
                    Else
                        lblSubLocation.Text = ""
                    End If
                    If (clsCommon.myLen(lblProject.Text) <= 0) Then
                        fndProject.Text = objOrderHead.PROJECT_ID
                        lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Text + "'")
                    End If
                    If (clsCommon.myLen(txtVehcileCode.Value) <= 0) Then
                        txtVehcileCode.Value = objOrderHead.Vehicle_Code
                        txtVehicleNo.Text = objOrderHead.VehicleNo
                    End If
                    If (clsCommon.myLen(txtBillToLocation.Value) <= 0) Then
                        txtBillToLocation.Value = objOrderHead.Bill_To_Location
                        lblBillToLocation.Text = objOrderHead.BillToLocationName
                    End If
                    If (clsCommon.myLen(txtRouteNo.Value)) <= 0 Then
                        txtRouteNo.Value = objOrderHead.Route_No
                        lblRouteDesc.Text = objOrderHead.Route_Desc
                    End If
                    If (clsCommon.myLen(txtPriceCode.Text)) <= 0 Then
                        txtPriceCode.Text = objOrderHead.Price_Code
                    End If
                    If (clsCommon.myLen(txtPriceGroupCode.Text)) <= 0 Then
                        txtPriceGroupCode.Text = objOrderHead.Price_Group_Code
                    End If
                    If (clsCommon.myLen(txtVendorNo.Value) <= 0) Then
                        txtVendorNo.Value = frm.VendorCode
                        lblVendorName.Text = frm.VendorName
                        chkVendorGrossReceipt.Checked = clsVendorMaster.isGrossReceipt(frm.VendorCode)

                        ' txtDate.Enabled = False
                        txtVendorNo.Enabled = False
                        chkRateUserCustomer.ToggleState = ClsUserCustomerSettings.GetUserCustomerRateSetting(txtVendorNo.Value)
                    End If
                    txtDate.Enabled = True
                    If (clsCommon.myLen(txtTermCode.Value) <= 0) Then
                        txtTermCode.Value = objOrderHead.Terms_Code
                        lblTermName.Text = objOrderHead.TermsName
                        If objOrderHead.Due_Date IsNot Nothing Then
                            txtDueDate.Value = objOrderHead.Due_Date

                        End If
                    End If
                    'ddlReturnType.Enabled = False
                    ddlReturnType.SelectedValue = "I"
                    If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                        txtTaxGroup.Value = objOrderHead.Tax_Group
                        SetTaxDetails()
                    End If

                    '' If clsCommon.myLen(txtSalesman.Value) <= 0 Then
                    txtSalesman.Value = objOrderHead.Salesman_Code
                    lblSalesman.Text = objOrderHead.Salesman_Name
                    ''End If
                    If clsCommon.myCdbl(txtCrate.Value) = 0 Then
                        txtCrate.Value = objOrderHead.Crate
                    Else
                        txtCrate.Value += objOrderHead.Crate
                    End If
                    If clsCommon.myCdbl(txtJaali.Value) = 0 Then
                        txtJaali.Value = objOrderHead.jaali
                    Else
                        txtJaali.Value += objOrderHead.jaali
                    End If
                    If clsCommon.myCdbl(txtBox.Value) = 0 Then
                        txtBox.Value = objOrderHead.Box
                    Else
                        txtBox.Value += objOrderHead.Box
                    End If

                    If clsCommon.myLen(txtDiscAmt.Text) <= 0 OrElse clsCommon.myLen(txtDiscPer.Text) <= 0 OrElse clsCommon.myCdbl(txtDiscAmt.Text) = 0 OrElse clsCommon.myCdbl(txtDiscPer.Text) = 0 Then
                        If clsCommon.myLen(txtDiscAmt.Text) <= 0 OrElse clsCommon.myLen(txtDiscPer.Text) <= 0 OrElse clsCommon.myCdbl(txtDiscAmt.Text) = 0 OrElse clsCommon.myCdbl(txtDiscPer.Text) = 0 Then
                            txtDiscPer.Text = objOrderHead.HeadDisc_Per
                            If clsCommon.myCdbl(txtDiscPer.Text) = 0 Then
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

            For Each obj As clsPSShipmentHeadDetail In frm.ArrReturn

                If IsValidItem(obj) Then
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDelivery_Code).Value = obj.Delivery_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = obj.Document_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = obj.Row_Type
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(obj.Item_Code)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code

                    gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = obj.Balance_Qty
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = obj.Balance_Qty
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = obj.Balance_Qty


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
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgUnit).Value = obj.Unit_code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = obj.Batch_No
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
                    If obj.MFG_Date.HasValue Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colManufactureDate).Value = obj.MFG_Date
                    End If
                    If obj.Expiry_Date.HasValue Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colExpiry).Value = obj.Expiry_Date
                    End If

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
                    ' gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value))
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
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = obj.HeadDiscAmt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeaDDisPer).Value = obj.HeadDiscPer
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDisPerAmt).Value = obj.HeadDiscPerAmt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = IIf(obj.Scheme_Applicable = "Y", "Yes", "No")
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colFromSchemeCode).Value = obj.Scheme_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colInvoiceSchemeCode).Value = obj.Scheme_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = IIf(obj.Scheme_Item = "Y", "Yes", "No")
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitALter).Value = obj.Alternate_UOM
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitRate).Value = obj.RATE_UOM
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)

                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = clsBatchInventory.GetData(objOrderHead.Trans_Type, objOrderHead.Document_Code, obj.Item_Code, obj.Line_No, Nothing, objOrderHead.Screen_Type, obj.Unit_code)
                    txtReqNo.Value = obj.Document_Code
                    findQtyandPromoSchemeCode(False, obj.Scheme_Code, objOrderHead.Document_Date)
                End If
            Next
            findVolumeStructureSchemeCode(False, "")

            'For ii As Integer = 0 To frm.ArrReturn.Count - 1
            '    If clsCommon.myLen(frm.ArrReturn(ii).Document_Code) > 0 Then
            '        Dim strCode As String = frm.ArrReturn(ii).Document_Code
            '        'If Not arr.Contains(strCode) Then
            '        '    arr.Add(strCode)
            '        objOrderHead = clsPSInvoiceHead.GetData(frm.ArrReturn(ii).Document_Code, "", NavigatorType.Current)
            '        For Each obj As clsPSInvoiceHeadDetail In objOrderHead.Arr
            '            'If (obj.Item_Code = frm.ArrReturn(ii).Item_Code AndAlso obj.Scheme_Item = "N") OrElse (obj.Scheme_Code = frm.ArrReturn(ii).Scheme_Code AndAlso clsCommon.myLen(obj.Scheme_Code) > 0) Then
            '            If (obj.Item_Code = frm.ArrReturn(ii).Item_Code AndAlso obj.Unit_code = frm.ArrReturn(ii).Unit_code AndAlso obj.Scheme_Item = "N") Then ''OrElse (obj.Scheme_Code = frm.ArrReturn(ii).Scheme_Code AndAlso clsCommon.myLen(obj.Scheme_Code) > 0) Then


            '            End If
            '        Next

            '    End If

            'Next


            'For ii As Integer = 0 To frm.ArrReturn.Count - 1
            '    If clsCommon.myLen(frm.ArrReturn(ii).Document_Code) > 0 Then
            '        Dim strCode As String = frm.ArrReturn(ii).Document_Code
            '        'If Not arr.Contains(strCode) Then
            '        '    arr.Add(strCode)
            '        objOrderHead = clsPSInvoiceHead.GetData(frm.ArrReturn(ii).Document_Code, "", NavigatorType.Current)
            '        For Each obj As clsPSInvoiceHeadDetail In objOrderHead.Arr
            '            'If (obj.Item_Code = frm.ArrReturn(ii).Item_Code AndAlso obj.Scheme_Item = "N") OrElse (obj.Scheme_Code = frm.ArrReturn(ii).Scheme_Code AndAlso clsCommon.myLen(obj.Scheme_Code) > 0) Then
            '            If (obj.Item_Code = frm.ArrReturn(ii).Item_Code AndAlso obj.Unit_code = frm.ArrReturn(ii).Unit_code AndAlso obj.Scheme_Item = "N") Then ''OrElse (obj.Scheme_Code = frm.ArrReturn(ii).Scheme_Code AndAlso clsCommon.myLen(obj.Scheme_Code) > 0) Then

            '                If IsValidItem(obj) Then
            '                    gv1.Rows.AddNew()
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDelivery_Code).Value = obj.Delivery_Code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = obj.Document_Code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = obj.Row_Type
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = obj.Item_Code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = obj.Item_Desc
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(obj.Item_Code, Nothing)
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(obj.Item_Code)
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(obj.Item_Code)
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = obj.Item_Cost
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = obj.Unit_code

            '                    'Special case (Shipment which contain two booking with same item,UOM)
            '                    Dim RecordCount As Integer = 0
            '                    Dim strqry As String = ""
            '                    If ShowMulMRPOfSameItemOnDairyBookingCustomer AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", Nothing)), "Others") = CompairStringResult.Equal Then
            '                        strqry = "select count(distinct TSPL_SD_SALE_INVOICE_DETAIL.Delivery_Code) as cc FROM TSPL_SD_SALE_INVOICE_DETAIL  " & _
            '                    " where TSPL_SD_SALE_INVOICE_DETAIL.Document_Code='" + obj.Document_Code + "'" & _
            '                    " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code='" + obj.Item_Code + "'" & _
            '                    " and TSPL_SD_SALE_INVOICE_DETAIL.Unit_code='" + obj.Unit_code + "'" & _
            '                    " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost  =" & obj.Item_Cost & " " & _
            '                    " group by TSPL_SD_SALE_INVOICE_DETAIL.Shipment_Code,TSPL_SD_SALE_INVOICE_DETAIL.Document_Code,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code,TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost,TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item "
            '                    Else
            '                        strqry = "select count(distinct TSPL_SD_SALE_INVOICE_DETAIL.Delivery_Code) as cc FROM TSPL_SD_SALE_INVOICE_DETAIL  " & _
            '                    " where TSPL_SD_SALE_INVOICE_DETAIL.Document_Code='" + obj.Document_Code + "'" & _
            '                    " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code='" + obj.Item_Code + "'" & _
            '                    " and TSPL_SD_SALE_INVOICE_DETAIL.Unit_code='" + obj.Unit_code + "'" & _
            '                    " group by TSPL_SD_SALE_INVOICE_DETAIL.Shipment_Code,TSPL_SD_SALE_INVOICE_DETAIL.Document_Code,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code,TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item "
            '                    End If

            '                    RecordCount = clsDBFuncationality.getSingleValue(strqry)
            '                    If ShowMulMRPOfSameItemOnDairyBookingCustomer AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", Nothing)), "Others") = CompairStringResult.Equal Then
            '                        If frm.ArrReturn.Count > 1 AndAlso gv1.Rows.Count > 1 AndAlso RecordCount > 1 AndAlso frm.ArrReturn(ii + 1).Document_Code = obj.Document_Code AndAlso frm.ArrReturn(ii + 1).Item_Code = obj.Item_Code AndAlso frm.ArrReturn(ii + 1).Unit_code = obj.Unit_code AndAlso frm.ArrReturn(ii + 1).Item_Cost = obj.Item_Cost Then
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                        ElseIf frm.ArrReturn.Count > 1 AndAlso gv1.Rows.Count > 1 AndAlso RecordCount = 1 AndAlso frm.ArrReturn(ii + 1).Document_Code = obj.Document_Code AndAlso frm.ArrReturn(ii + 1).Item_Code = obj.Item_Code AndAlso frm.ArrReturn(ii + 1).Unit_code = obj.Unit_code AndAlso frm.ArrReturn(ii + 1).Item_Cost = obj.Item_Cost Then
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                        Else
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = frm.ArrReturn(ii).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = frm.ArrReturn(ii).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = frm.ArrReturn(ii).Balance_Qty
            '                        End If
            '                    Else
            '                        If frm.ArrReturn.Count > 1 AndAlso gv1.Rows.Count > 1 AndAlso RecordCount > 1 AndAlso frm.ArrReturn(ii + 1).Document_Code = obj.Document_Code AndAlso frm.ArrReturn(ii + 1).Item_Code = obj.Item_Code AndAlso frm.ArrReturn(ii + 1).Unit_code = obj.Unit_code Then
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = frm.ArrReturn(ii + 1).Balance_Qty
            '                        Else
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = frm.ArrReturn(ii).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = frm.ArrReturn(ii).Balance_Qty
            '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = frm.ArrReturn(ii).Balance_Qty
            '                        End If
            '                    End If


            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = obj.TAX1_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = obj.TAX2_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = obj.TAX3_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = obj.TAX4_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = obj.TAX5_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = obj.TAX6_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = obj.TAX7_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = obj.TAX8_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = obj.TAX9_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = obj.TAX10_Rate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPer).Value = obj.Disc_Per
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value = obj.MRP
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgUnit).Value = obj.Unit_code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = obj.Batch_No
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBinNo).Value = obj.Bin_No
            '                    If obj.MFG_Date.HasValue Then
            '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colManufactureDate).Value = obj.MFG_Date
            '                    End If
            '                    If obj.Expiry_Date.HasValue Then
            '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colExpiry).Value = obj.Expiry_Date
            '                    End If

            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = obj.Item_Tax
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalMRP).Value = obj.Total_MRP_Amt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalBasicAmount).Value = obj.Total_Basic_Amt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalDiscountAmount).Value = obj.Total_Disc_Amt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colcustDiscount).Value = obj.Cust_Discount
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotalCustDiscount).Value = obj.Total_Cust_Discount
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colActualCost).Value = obj.ActualRate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColCustDiscountQty).Value = obj.Cust_DiscountQty
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceDateColumn).Value = obj.Price_Date
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceCOde).Value = obj.Price_code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementPer).Value = obj.Abatement_Per
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAbatementAmount).Value = obj.Abatement_Amt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColFOC).Value = obj.FOC_Item
            '                    ' gv1.Rows(gv1.Rows.Count - 1).Cells(ColActualBalQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value), txtBillToLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(gv1.Rows.Count - 1).Cells(colMRP).Value))
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = obj.Batch_No
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colItemWeight).Value = obj.Item_Weight
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colConvF).Value = obj.Conv_Factor
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotItemWt).Value = obj.TotalItem_Weight
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkupOn).Value = obj.Markup_On
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colMarkUpPercentage).Value = obj.Markup_Percent
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLandingCost).Value = obj.Landing_Cost
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCustDiscPercentage).Value = obj.CustDiscPer
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = obj.HeadDiscAmt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCashDiscSchemeCode).Value = obj.CasdDiscScheme_Code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPurCost).Value = obj.Purchase_Cost
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgCost).Value = obj.OrgRate
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleCode).Value = obj.PrincipleCode
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPricipleDesc).Value = obj.PrincipleDesc
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorCode).Value = obj.vendor_code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colvendorDesc).Value = obj.vendor_desc
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDiscamt).Value = obj.HeadDiscAmt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeaDDisPer).Value = obj.HeadDiscPer
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHeadDisPerAmt).Value = obj.HeadDiscPerAmt
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeApplicable).Value = IIf(obj.Scheme_Applicable = "Y", "Yes", "No")
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colFromSchemeCode).Value = obj.Scheme_Code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colInvoiceSchemeCode).Value = obj.Scheme_Code
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = IIf(obj.Scheme_Item = "Y", "Yes", "No")
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitALter).Value = obj.Alternate_UOM
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnitRate).Value = obj.RATE_UOM
            '                    'gv1.Rows(gv1.Rows.Count - 1).Cells(colICodeGrp).Value = o
            '                    'Batch wise item return Ticket No- ALF/22/05/18-000066
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(obj.Item_Code)

            '                    ' done by priti BHA/28/06/18-000113 added by richa agarwal 3 May,2019
            '                    ' gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = obj.arrBatchItem
            '                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Tag = clsBatchInventory.GetData(objOrderHead.Trans_type, objOrderHead.Against_Shipment_No, obj.Item_Code, obj.Line_No, Nothing, objOrderHead.Screen_Type)
            '                    'Batch wise item return Ticket No- ALF/22/05/18-000066
            '                    findQtyandPromoSchemeCode(False, obj.Scheme_Code)
            '                End If
            '                End If
            '        Next

            '    End If

            'Next
            If objOrderHead.Arr IsNot Nothing AndAlso objOrderHead.Arr.Count > 0 Then
                For Each objTr As clsPSShipmentHeadDetail In objOrderHead.Arr
                    If objTr.Row_Type = "Misc" Then
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = mrnno
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = objTr.Row_Type
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = objTr.Item_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = objTr.Item_Desc
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(objTr.Item_Code, Nothing)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = objTr.Item_Cost
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = objTr.Unit_code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = 0
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colReturnAmt).Value = objTr.Amount
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
            If UOMAtDiarySaleReturn = True Then
                gv1.Columns(colActualQty).IsVisible = True
                gv1.Columns(colActualUOM).IsVisible = True
                gv1.Columns(colConvAmt).IsVisible = True
                gv1.Columns(colQty).ReadOnly = True
                gv1.Columns(colUnit).ReadOnly = True
            Else
                gv1.Columns(colActualQty).IsVisible = False
                gv1.Columns(colActualUOM).IsVisible = False
                gv1.Columns(colConvAmt).IsVisible = False

            End If
            'gv1.Rows.AddNew()
            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
            SetitemWiseTaxSetting(False, False)
            RefreshReqNo()
            For ii As Integer = 0 To gv1.RowCount - 1
                UpdateCurrentRow(ii)
            Next
        End If
        isInsideLoadData = False
        UpdateAllTotals()
        isInvoiceLoadData = False
    End Sub

    Function IsValidItem(ByVal obj As clsPSShipmentHeadDetail)
        If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
            txtTaxGroup.Value = obj.SRNTax_Group
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
            Dim dblMRP As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colMRP).Value)
            Dim strSchemeItem As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colSchemeItem).Value)
            Dim strDelivery_Code As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colDelivery_Code).Value)
            Dim dblBasicRate As Double = clsCommon.myCstr(gv1.Rows(ii).Cells(colRate).Value)
            If clsCommon.myLen(strReqCode) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(strICode), obj.Item_Code) = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Scheme_Item, "N") = CompairStringResult.Equal AndAlso clsCommon.CompairString(strReqCode, obj.Document_Code) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strUOM, obj.Unit_code) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strDelivery_Code, obj.Delivery_Code) = CompairStringResult.Equal AndAlso (ShowMulMRPOfSameItemOnDairyBookingCustomer = True And clsCommon.CompairString(clsCommon.myCstr(dblBasicRate), obj.Item_Cost) = CompairStringResult.Equal) Then
                Dim strMsg As String = "Already exist at row no:" + clsCommon.myCstr(ii + 1) + Environment.NewLine + "Order No : " + obj.Document_Code + "  Item : " + obj.Item_Desc + Environment.NewLine + ""
                If dblMRP > 0 Then
                    strMsg = strMsg + Environment.NewLine + "MRP : " + clsCommon.myCstr(dblMRP)
                End If
                If clsCommon.CompairString(clsCommon.myCstr(dblBasicRate), obj.Item_Cost) = CompairStringResult.Equal Then
                    strMsg = strMsg + Environment.NewLine + "Rate : " + clsCommon.myCstr(dblBasicRate)
                End If
                Return False
            End If


        Next
        Return True
    End Function

    Private Sub gv1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv1.DoubleClick
        Try
            If gv1.CurrentColumn Is gv1.Columns(colTotTaxAmt) AndAlso rbtnTaxCalAutomatic.IsChecked Then
                Dim frm As New FrmPOItemTaxDetails()
                frm.strLineNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colLineNo).Value)
                frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                frm.dblTotTax = clsCommon.myCdbl(gv1.CurrentRow.Cells(colTotTaxAmt).Value)
                frm.dblAmtAfterDis = clsCommon.myCdbl(gv1.CurrentRow.Cells(colAmtAfterDis).Value)
                If clsCommon.myLen(frm.strItemCode) > 0 Then
                    frm.ArrIn = New List(Of clsTempItemTaxDetails)
                    For ii As Integer = 1 To 10
                        Dim strii As String = clsCommon.myCstr(ii)
                        Dim obj As New clsTempItemTaxDetails()
                        obj.AuthorityCode = clsCommon.myCstr(gv1.CurrentRow.Cells("COLTAX" + strii).Value)
                        If clsCommon.myLen(obj.AuthorityCode) > 0 Then
                            obj.AuthorityName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Code_Desc from TSPL_TAX_MASTER WHERE Tax_Code='" + obj.AuthorityCode + "'"))
                            obj.Rate = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value)
                            obj.BaseAmt = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value)
                            obj.TaxAmt = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value)
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
            ElseIf gv1.CurrentColumn Is gv1.Columns(colComplete) AndAlso UsLock1.Status = ERPTransactionStatus.Approved Then
                Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                Dim intSNo As Integer = Convert.ToInt32((clsCommon.myCdbl(gv1.CurrentRow.Cells(colLineNo).Value)))
                Dim strStatus As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colComplete).Value)
                If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso clsCommon.myLen(strICode) > 0 AndAlso intSNo > 0 AndAlso clsCommon.CompairString(strStatus, "No") = CompairStringResult.Equal Then
                    If common.clsCommon.MyMessageBoxShow(Me, "Do you want to complete the item " + clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value), Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                        If clsPSSalesReturnDetail.CompleteSRN(txtDocNo.Value, strICode, intSNo) Then
                            common.clsCommon.MyMessageBoxShow(Me, "Successfully Completed", Me.Text)
                            LoadData(txtDocNo.Value, NavigatorType.Current)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Function GetColumnsForTaxRateType(ByVal dt As DataTable)
        Dim cols As String = ""
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                'If (dt.Rows.IndexOf(dr) + 1) = dt.Rows.Count Then
                '    cols = cols & "'" & dr.Item("Tax_RateType_Name") & "'" & " as Tax_RateType_Name" & (dt.Rows.IndexOf(dr) + 1) & "," & dr.Item("Tax_RateType_Rate") & " as Tax_RateType_Rate" & (dt.Rows.IndexOf(dr) + 1) & "," & dr.Item("Tax_RateType_Amount") & " as Tax_RateType_Amount" & (dt.Rows.IndexOf(dr) + 1)
                'Else
                '    cols = cols & "'" & dr.Item("Tax_RateType_Name") & "'" & " as Tax_RateType_Name" & (dt.Rows.IndexOf(dr) + 1) & "," & dr.Item("Tax_RateType_Rate") & " as Tax_RateType_Rate" & (dt.Rows.IndexOf(dr) + 1) & "," & dr.Item("Tax_RateType_Amount") & " as Tax_RateType_Amount" & (dt.Rows.IndexOf(dr) + 1) & ","
                'End If
                cols = cols & "'" & dr.Item("Tax_RateType_Name") & "'" & " as Tax_RateType_Name" & (dt.Rows.IndexOf(dr) + 1) & "," & dr.Item("Tax_RateType_Rate") & " as Tax_RateType_Rate" & (dt.Rows.IndexOf(dr) + 1) & "," & dr.Item("Tax_RateType_Amount") & " as Tax_RateType_Amount" & (dt.Rows.IndexOf(dr) + 1) & ","
            Next
        End If
        For i As Integer = (dt.Rows.Count + 1) To 7
            If i = 7 Then
                cols = cols & "''" & " as Tax_RateType_Name" & (i) & "," & "null" & " as Tax_RateType_Rate" & (i) & "," & "null" & " as Tax_RateType_Amount" & (i)
            Else
                cols = cols & "''" & " as Tax_RateType_Name" & (i) & "," & "null" & " as Tax_RateType_Rate" & (i) & "," & "null" & " as Tax_RateType_Amount" & (i) & ","
            End If
        Next

        If clsCommon.myLen(cols) > 0 Then
            Return "," & cols
        Else
            Return ""
        End If
    End Function
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                myMessages.blankValue(Me, "Document not found to Print", Me.Text)
            Else
                funPrint(txtDocNo.Value)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Function GetTaxRateTypeDT(ByVal DocNo As String, Optional ByVal trans As SqlTransaction = Nothing) As DataTable
        Dim qry As String = ""
        qry = " select distinct * from (" &
              " select distinct TAX1 as Tax_RateType_Name,TAX1_Rate as Tax_RateType_Rate,sum(TAX1_Amt) as Tax_RateType_Amount  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX1,TAX1_Rate " &
              " union all " &
              " select distinct TAX2,TAX2_Rate,sum(TAX2_Amt) as TAX2_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX2,TAX2_Rate " &
              " union all " &
              " select distinct TAX3,TAX3_Rate,sum(TAX3_Amt) as TAX3_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX3,TAX3_Rate " &
              " union all " &
              " select distinct TAX4,TAX4_Rate,sum(TAX4_Amt) as TAX4_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX4,TAX4_Rate " &
              " union all " &
              " select distinct TAX5,TAX5_Rate,sum(TAX5_Amt) as TAX5_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX5,TAX5_Rate " &
              " union all " &
              " select distinct TAX6,TAX6_Rate,sum(TAX6_Amt) as TAX6_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX6,TAX6_Rate " &
              " union all " &
              " select distinct TAX7,TAX7_Rate,sum(TAX7_Amt) as TAX7_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX7,TAX7_Rate " &
              " union all " &
              " select distinct TAX8,TAX8_Rate,sum(TAX8_Amt) as TAX8_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX8,TAX8_Rate " &
              " union all " &
              " select distinct TAX9,TAX9_Rate,sum(TAX9_Amt) as TAX9_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX9,TAX9_Rate " &
              " union all " &
              " select distinct TAX10,TAX10_Rate,sum(TAX10_Amt) as TAX1_Amt  from TSPL_SD_SHIPMENT_RETURN_DETAIL where DOCUMENT_CODE='" & DocNo & "' group by TAX10,TAX10_Rate " &
              " ) as tax where Tax_RateType_Name is not null and Tax_RateType_Amount>0"

        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Return dt
    End Function

    Public Sub funPrint(ByVal StrCode As String, Optional ByVal IsPDF As Boolean = False)
        Try
            clsDSSalesReturnHead.funsaleReturnDairyPrint(MyBase.Form_ID, "Dispatch_Return", txtDate.Value, StrCode, IsPDF)


        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Function SetItemWiseTax(ByVal dtAfterModify As DataTable, ByVal strShipFrm As String) As DataTable
        dtAfterModify.Columns.Add("TAX1_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX1_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX2_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX2_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX3_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX3_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX4_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX4_Amt3", GetType(Double))

        dtAfterModify.Columns.Add("TAX5_Rate1", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate2", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Rate3", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt1", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt2", GetType(Double))
        dtAfterModify.Columns.Add("TAX5_Amt3", GetType(Double))

        Dim qry As String = "select Tax,Rate,SUM(Amt) as TaxAmt"
        qry += " from ("
        qry += " select TAX1 as Tax,TAX1_Rate as Rate,TAX1_Amt as Amt"
        qry += " from TSPL_SD_SHIPMENT_RETURN_DETAIL where Document_Code='" + strShipFrm + "' "
        qry += " union all "
        qry += " select TAX2 as Tax,TAX2_Rate as Rate,TAX2_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_RETURN_DETAIL where Document_Code='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX3 as Tax,TAX3_Rate as Rate,TAX3_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_RETURN_DETAIL where Document_Code='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX4 as Tax,TAX4_Rate as Rate,TAX4_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_RETURN_DETAIL where Document_Code='" + strShipFrm + "'  "
        qry += " union all "
        qry += " select TAX5 as Tax,TAX5_Rate as Rate,TAX5_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_RETURN_DETAIL where Document_Code='" + strShipFrm + "'   "
        qry += " union all "
        qry += " select TAX6 as Tax,TAX6_Rate as Rate,TAX6_Amt as Amt "
        qry += " from TSPL_SD_SHIPMENT_RETURN_DETAIL where Document_Code='" + strShipFrm + "'   "
        qry += " )xxx "
        qry += " group by Tax,Rate   having SUM(Amt)>0   "


        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                For ii As Integer = 1 To 5
                    Dim strCol As String = "TAX" + clsCommon.myCstr(ii) + ""
                    If clsCommon.CompairString(clsCommon.myCstr(dtAfterModify.Rows(0)(strCol)), clsCommon.myCstr(dr("Tax"))) = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt1")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate1") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt1") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt2")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate2") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt2") = clsCommon.myCdbl(dr("TaxAmt"))
                            Next
                        ElseIf clsCommon.myCdbl(dtAfterModify.Rows(0)("TAX" + clsCommon.myCstr(ii) + "_Amt3")) <= 0 Then
                            For jj As Integer = 0 To dtAfterModify.Rows.Count - 1
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Rate3") = clsCommon.myCdbl(dr("Rate"))
                                dtAfterModify.Rows(jj)("TAX" + clsCommon.myCstr(ii) + "_Amt3") = clsCommon.myCdbl(dr("TaxAmt"))
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

                    Else
                        gv1.CurrentRow.Cells(colICode).ReadOnly = False
                        gv1.CurrentRow.Cells(colMRP).ReadOnly = False

                    End If

                ElseIf e.Column Is gv1.Columns(colQty) Then
                    If clsCommon.CompairString(gv1.CurrentRow.Cells(colSchemeItem).Value, "Yes") = CompairStringResult.Equal Then
                        gv1.CurrentRow.Cells(colQty).ReadOnly = True
                    Else
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal Then
                            If clsCommon.myLen(txtReqNo.Value) = 0 AndAlso clsCommon.CompairString(ddlReturnType.SelectedValue, "D") = CompairStringResult.Equal Then
                                gv1.CurrentRow.Cells(colQty).ReadOnly = True
                                gv1.CurrentRow.Cells(colFreeQty).ReadOnly = True
                                gv1.CurrentRow.Cells(colDamageQty).ReadOnly = False
                            ElseIf clsCommon.myLen(txtReqNo.Value) = 0 AndAlso clsCommon.CompairString(ddlReturnType.SelectedValue, "P") = CompairStringResult.Equal Then
                                gv1.CurrentRow.Cells(colQty).ReadOnly = True
                                gv1.CurrentRow.Cells(colDamageQty).ReadOnly = True
                                gv1.CurrentRow.Cells(colFreeQty).ReadOnly = True
                            Else
                                gv1.CurrentRow.Cells(colQty).ReadOnly = False
                                gv1.CurrentRow.Cells(colFreeQty).ReadOnly = False
                                gv1.CurrentRow.Cells(colDamageQty).ReadOnly = False
                            End If
                        Else
                            gv1.CurrentRow.Cells(colQty).ReadOnly = True
                            gv1.CurrentRow.Cells(colFreeQty).ReadOnly = True
                        End If
                    End If


                ElseIf e.Column Is gv1.Columns(colRate) Then
                    If clsCommon.myLen(txtReqNo.Value) > 0 Then
                        gv1.CurrentRow.Cells(colRate).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colRate).ReadOnly = False
                    End If

                ElseIf e.Column Is gv1.Columns(colDisPer) Then
                    If chkRateUserCustomer.ToggleState = ToggleState.Indeterminate Then
                        gv1.CurrentRow.Cells(colDisPer).ReadOnly = Not chkRateDefaultSetting.Checked
                    Else
                        gv1.CurrentRow.Cells(colDisPer).ReadOnly = Not chkRateUserCustomer.Checked
                    End If
                ElseIf e.Column Is gv1.Columns(colAmt) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colRowType).Value), RowTypeMisc) = CompairStringResult.Equal OrElse clsCommon.myCdbl(gv1.CurrentRow.Cells(colIsMannualAmt).Value) = 1 Then
                        gv1.CurrentRow.Cells(colAmt).ReadOnly = False
                    Else
                        gv1.CurrentRow.Cells(colAmt).ReadOnly = True
                    End If

                ElseIf e.Column Is gv1.Columns(colUnit) Then
                    If clsCommon.myLen(txtReqNo.Value) > 0 Then
                        gv1.CurrentRow.Cells(colUnit).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colUnit).ReadOnly = False

                    End If

                End If
                'Dim cell As GridDataCellElement = TryCast(e.CellElement, GridDataCellElement)
                'cell.GradientStyle = GradientStyles.Solid
                'cell.BackColor = Color.FromArgb(243, 181, 51)
            End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(ex.Message)
        End Try
    End Sub

    Private Sub gv1_UserDeletingRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        common.clsCommon.MyMessageBoxShow(Me, "Can't Delete Row", Me.Text)
        e.Cancel = True
        'If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
        '    e.Cancel = True
        'ElseIf clsCommon.myCdbl(gv1.CurrentRow.Cells(ColFOC).Value) = 1 Then
        '    common.clsCommon.MyMessageBoxShow("Can't Delete The Current Row.This Item is Scheme Item")
        '    e.Cancel = True
        'End If
        'RefreshSerialNo()
    End Sub

    Private Sub gv1_UserDeletedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserDeletedRow
        UpdateAllTotals()
        For ii As Integer = 1 To gv1.Rows.Count
            gv1.Rows(ii - 1).Cells(colLineNo).Value = ii
        Next
        RefreshReqNo()
        RefreshSerialNo()
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

    Private Sub txtReqNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtReqNo._MYValidating
        SelectMRNItems()
    End Sub

    Sub RefreshReqNo()
        txtReqNo.Value = ""
        If gv1.Rows.Count > 0 Then
            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strReqNo As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colOrderNo).Value)
                If clsCommon.myLen(strReqNo) > 0 Then
                    txtReqNo.Value = clsCommon.myCstr(strReqNo)
                    Exit Sub
                End If
            Next
        End If
    End Sub

    'Public Sub funrejprint()
    '    Try
    '        Dim qry As String = "SELECT TSPL_SD_Shipment_RETURN_HEAD.Document_Code, TSPL_SD_Shipment_RETURN_HEAD.Document_Date, TSPL_SD_Shipment_RETURN_HEAD.Customer_Name, TSPL_SD_Shipment_RETURN_HEAD.Ship_To_Location, TSPL_SD_Shipment_RETURN_HEAD.Bill_To_Location, TSPL_SD_Shipment_RETURN_HEAD.RMDA_No, TSPL_SD_Shipment_RETURN_HEAD.RMDA_Date,TSPL_SD_Shipment_RETURN_HEAD.Remarks,TSPL_SD_SALE_RETURN_HEAD.Description, TSPL_SRN_DETAIL.Item_Code,   TSPL_SRN_DETAIL.Item_Desc, TSPL_SRN_DETAIL.Rejected_Qty, TSPL_SRN_DETAIL.Item_Cost,TSPL_SRN_DETAIL.Unit_code,TSPL_SRN_DETAIL.Rejected_Qty*TSPL_SRN_DETAIL.Item_Cost as Amount,TSPL_COMPANY_MASTER.comp_name,TSPL_COMPANY_MASTER.Logo_Img, TSPL_COMPANY_MASTER.Logo_Img2  FROM         TSPL_SD_SALE_RETURN_HEAD INNER JOIN    TSPL_SRN_DETAIL ON TSPL_SD_SALE_RETURN_HEAD.Document_Code = TSPL_SRN_DETAIL.Document_Code LEFT OUTER JOIN     TSPL_COMPANY_MASTER ON TSPL_SD_SALE_RETURN_HEAD.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code  where TSPL_SD_SALE_RETURN_HEAD.Document_Code='" + txtDocNo.Value + "'"

    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '        Dim frmCRV As New frmCrystalReportViewer()
    '        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.PurchaseOrder, dt, "rptMRDA", "MRDA Report")
    '        frmCRV = Nothing
    '    Catch ex As Exception
    '        common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub

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
    Private Sub chkDiscountOnAmt_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkDiscountOnAmt.ToggleStateChanged, chkDiscountOnRate.ToggleStateChanged
        If chkDiscountOnAmt.IsChecked Then
            txtDiscAmt.Enabled = True
            txtDiscPer.Enabled = False
            txtDiscPer.Text = 0
            For Each gro As GridViewRowInfo In gv1.Rows
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCdbl(gro.Cells(ColFOC).Value) = 0 Then
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
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCdbl(gro.Cells(ColFOC).Value) = 0 Then
                    gro.Cells(colHeadDiscamt).Value = 0
                Else

                End If
            Next
            lblInvoiceDiscAmt.Text = 0

        End If
    End Sub
    Private Sub txtDiscAmt_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscAmt.Leave
        CalculateDiscountAmount()
    End Sub

    Private Sub txtDiscPer_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscPer.Leave
        CalculateDiscountAmount()
    End Sub
    Private Sub CalculateDiscountAmount()
        If clsCommon.myCdbl(txtDiscAmt.Text) > clsCommon.myCdbl(lblAmtWithDiscount.Text) Then
            isCellValueChangedOpen = False
            Throw New Exception("Discount amount cannot be greater than Doc amount")

        End If
        Dim discountrate As Decimal
        'Dim discountrate As Decimal = Decimal.Parse(txtDiscPer.Text)
        'If chkDiscountOnAmt.IsChecked AndAlso clsCommon.myCdbl(lblAmtWithDiscount.Text) > 0 Then
        '    discountrate = clsCommon.myCstr(Math.Round(((txtDiscAmt.Value * 100) / clsCommon.myCdbl(lblAmtWithDiscount.Text)), 5, MidpointRounding.ToEven))
        '    txtDiscPer.Value = discountrate
        'End If
        If clsCommon.myCdbl(txtDiscPer.Text) > 0 Then
            discountrate = Decimal.Parse(txtDiscPer.Text)
            txtDiscAmt.Text = 0

        ElseIf clsCommon.myCdbl(txtDiscAmt.Text) > 0 Then
            txtDiscPer.Text = 0
        End If
        Dim dblDiscountAmt As Decimal = 0
        Dim dblCustDiscountNoTax As Double = 0
        If String.IsNullOrEmpty(lblAmtWithDiscount.Text) Then
            lblAmtWithDiscount.Text = 0
        End If

        If chkDiscountOnAmt.IsChecked Then
            For Each gro As GridViewRowInfo In gv1.Rows
                gv1.CurrentRow = gro.Cells(colHeadDiscamt).RowInfo
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCdbl(gro.Cells(ColFOC).Value) = 0 Then

                    dblDiscountAmt = Math.Round((clsCommon.myCdbl(gro.Cells(colAmt).Value) * txtDiscAmt.Value) / clsCommon.myCdbl(lblAmtWithDiscount.Text), 2)
                    gro.Cells(colHeadDiscamt).Value = Math.Round((dblDiscountAmt), 2)
                Else
                    gro.Cells(colHeadDiscamt).Value = 0

                End If

            Next
        Else
            For Each gro As GridViewRowInfo In gv1.Rows
                gv1.CurrentRow = gro.Cells(colHeadDiscamt).RowInfo
                If clsCommon.myLen(gro.Cells(colICode).Value) > 0 And clsCommon.myCdbl(gro.Cells(ColFOC).Value) = 0 Then
                    gro.Cells(colHeaDDisPer).Value = Math.Round((discountrate), 2)
                Else

                End If
            Next
        End If

    End Sub
    Private Function GetConvQty(ByVal IntRowNo As Integer) As Double
        Dim strItem As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
        Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
        Dim strCurrentUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnitRate).Value)
        Dim dblQTy As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value)
        Dim dblCurrentConvF As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strCurrentUnit & "'"))
        Dim dblConvQty As Double = 0
        If clsCommon.myLen(strOrgUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strOrgUnit & "'"))
            Dim dblStockingUnitConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and Stocking_Unit='Y' "))
            If dblCurrentConvF > 0 Then
                dblConvQty = Math.Round((dblQTy / dblCurrentConvF) * dblStockingUnitConvF * dblOrgConvF, 6)
            End If
        End If
        'gv1.Rows(IntRowNo).Cells(colOrgUnit).Value = strCurrentUnit
        'gv1.Rows(IntRowNo).Cells(colOrgUnitRate).Value = strCurrentUnit
        Return dblConvQty
    End Function
    Private Function GetConvAlterQty(ByVal IntRowNo As Integer) As Double
        Dim strItem As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
        Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
        Dim strCurrentUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnitALter).Value)
        Dim dblQTy As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value)
        Dim dblCurrentConvF As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strCurrentUnit & "'"))
        Dim dblConvQty As Double = 0
        If clsCommon.myLen(strOrgUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strOrgUnit & "'"))
            Dim dblStockingUnitConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and Stocking_Unit='Y' "))
            If dblCurrentConvF > 0 Then
                dblConvQty = Math.Round((dblQTy / dblCurrentConvF) * dblStockingUnitConvF * dblOrgConvF, 2)
            End If
        End If
        'gv1.Rows(IntRowNo).Cells(colOrgUnit).Value = strCurrentUnit
        'gv1.Rows(IntRowNo).Cells(colOrgUnitRate).Value = strCurrentUnit
        Return dblConvQty
    End Function
    Private Function GetConvRate(ByVal IntRowNo As Integer) As Double
        Dim strItem As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
        Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colOrgUnit).Value)
        Dim strCurrentUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
        Dim dblCurrentConvF As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colConvF).Value)
        Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
        Dim dblConvRate As Double = 0
        If clsCommon.myLen(strOrgUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strOrgUnit & "'"))
            Dim dblStockingUnitConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and Stocking_Unit='Y' "))
            dblConvRate = (dblRate / dblOrgConvF) * dblStockingUnitConvF * dblCurrentConvF
        End If
        Return dblConvRate
    End Function
    Private Function GetCurrentRowTotalTaxRate(ByVal IntRowNo As Integer) As Double
        Dim dblTotRate As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAX" + strii)).Value), "TCS") <> CompairStringResult.Equal Then
                If IntRowNo < 0 Then
                    dblTotRate = dblTotRate + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("COLTAXRate" + strii)).Value)
                Else
                    dblTotRate = dblTotRate + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRate" + strii)).Value)
                End If
            End If
        Next
        Return dblTotRate
    End Function
    Function TruncateToDecimalPlaces(value As Double, decimalPlaces As Integer) As Double
        Dim factor As Double = Math.Pow(10, decimalPlaces)
        Return Math.Truncate(value * factor) / factor
    End Function
    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer)
        Try
            Dim arrTaxableAuth As New List(Of String)
            Dim arrTaxableAuth1 As New List(Of String)
            Dim dblFAmt As Double = 0


            'If chkVendorGrossReceipt.Checked OrElse clsCommon.CompairString(cboItemType.SelectedValue, "F") = CompairStringResult.Equal Then
            '    dblQty = dblQty
            'End If

            Dim dblAlterQty As Double = 0
            Dim dblQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value)
            Dim dblActualQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colActualQty).Value)
            Dim dblDamageQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDamageQty).Value)
            Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblMRP As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colMRP).Value)
            Dim dblBasicRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblConvF As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colConvF).Value)
            Dim dblItemWeight As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colItemWeight).Value)
            Dim dblheadDiscamt As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colHeadDiscamt).Value)
            Dim dblOrgBasicRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colOrgCost).Value)
            'Dim dblConvBasicRate As Double = dblOrgBasicRate * dblConvF
            Dim dblMRPAmt As Double = dblQty * dblMRP


            Dim strICode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
            Dim strSchemeCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colFromSchemeCode).Value)
            If clsCommon.myLen(strICode) > 0 And clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0 Then

                'Dim obj_Cash As clsSchemeApplyOnDairy = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnitRate).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value)
                Dim obj_Cash As clsSchemeApplyOnDairy = Nothing
                'obj_Cash = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value, strSchemeCode)
                obj_Cash = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value, strSchemeCode)
                If clsCommon.myLen(obj_Cash.Schm_Code) = 0 AndAlso clsCommon.myLen(gv1.Rows(IntRowNo).Cells(colUnitALter).Value) > 0 Then
                    obj_Cash = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnitALter).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value, strSchemeCode)
                End If
                If obj_Cash IsNot Nothing Then
                    gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = obj_Cash.Cash_Amt
                    gv1.Rows(IntRowNo).Cells(colCash_Pers).Value = obj_Cash.Cash_Pers
                    gv1.Rows(IntRowNo).Cells(colCashSchemeCode).Value = obj_Cash.Schm_Code
                    If clsCommon.myCdbl(obj_Cash.Cash_Pers) > 0 Then
                        gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = "P"
                        gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = System.Math.Round((clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colCash_Pers).Value)) / 100, 2)
                    ElseIf clsCommon.myCdbl(obj_Cash.Cash_Amt) > 0 Then
                        gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = "A"
                    End If
                Else
                    gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = Nothing
                    gv1.Rows(IntRowNo).Cells(colCash_Pers).Value = Nothing
                    gv1.Rows(IntRowNo).Cells(colCashSchemeCode).Value = Nothing
                    gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = Nothing
                End If
            End If
            Dim dblCashAmt As Double = gv1.Rows(IntRowNo).Cells(colCash_Amt).Value
            Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblBasicAmt As Double = dblQty * dblRate
            Dim dblReturnAmt As Double = (dblActualQty * dblRate)
            Dim dblDamageAmt As Double = (dblDamageQty * dblRate)
            Dim dblAmt As Double = dblReturnAmt + dblDamageAmt ''+ dblFAmt
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colIsMannualAmt).Value) = 0 Then
                gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2) 'dblAmt 
                gv1.Rows(IntRowNo).Cells(colReturnAmt).Value = Math.Round(dblReturnAmt, 2) 'dblReturnAmt 
                gv1.Rows(IntRowNo).Cells(colDamageAmt).Value = Math.Round(dblDamageAmt, 2) 'dblDamageAmt 
            Else
                dblAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAmt).Value)
            End If

            If clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colCash_Amt).Value) > dblAmt Then
                gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = 0
                gv1.Rows(IntRowNo).Cells(colCash_Pers).Value = 0
                gv1.Rows(IntRowNo).Cells(colCashSchemeCode).Value = Nothing
                gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = Nothing
            End If

            GetDCDetails(IntRowNo)
            Dim dblTotalDCAmt As Double = 0
            Dim dblTotalTCAmt As Double = 0
            Dim dblTotTaxRate As Double = GetCurrentRowTotalTaxRate(IntRowNo)
            Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblSCRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColSCRate).Value)
            Dim dblDisAmt As Double = (dblAmt * dblDisPer) / 100
            'Dim dblSCAmt As Double = dblAmt * (dblSCRate / 100)
            'gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = dblSCAmt
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(ColDCRate).Value), "") <> CompairStringResult.Equal AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCRate).Value) >= 0 Then
                If Not ApplyCommissionRateWithTax Then
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = gv1.Rows(IntRowNo).Cells(ColDCRate).Value
                Else
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = clsCommon.myRoundOFF(gv1.Rows(IntRowNo).Cells(ColDCRate).Value * 100 / (100 + dblTotTaxRate), DispatchCommissionDecimalPlaces, 4)
                End If
                gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value = (clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colActualQty).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCUnitCF).Value)) / gv1.Rows(IntRowNo).Cells(ColDCCFUOM).Value
                gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value)
                gv1.Rows(IntRowNo).Cells(ColTCAmt).Value = clsCommon.myCstr(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColTCRate).Value))
                gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = clsCommon.myCstr(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColSCRate).Value))
                dblTotalDCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCAmt).Value)
                dblTotalTCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColTCAmt).Value)
                If dblTotalDCAmt > 0 Then
                    ' GetBoothWiseDCDetails(IntRowNo, trans)
                    If ApplyCommission Then
                        dblDisAmt = dblDisAmt + dblTotalDCAmt
                    End If
                End If
            End If

            'Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblDisAmt As Double = (dblAmt * dblDisPer) / 100
            Dim dblHeadDisPer As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colHeaDDisPer).Value)
            Dim dblHeadPerDisAmt As Double = (dblAmt * dblHeadDisPer) / 100

            Dim dblTotDiscAmt = dblheadDiscamt + dblHeadPerDisAmt + dblDisAmt + dblCashAmt
            Dim dblAmtAfterDis As Double = dblAmt - dblDisAmt - dblheadDiscamt - dblHeadPerDisAmt - dblCashAmt
            Dim dblAbatementRate As Double = gv1.CurrentRow.Cells(colAbatementPer).Value
            Dim dblAbatementAmt As Double = ((dblMRP * dblAbatementRate) / 100) * dblQty


            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)
                If rbtnTaxCalAutomatic.IsChecked Then
                    Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                    If clsCommon.myLen(strTaxCode) > 0 Then
                        Dim dblTaxRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                        Dim IsSurTax As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value)
                        Dim strSurTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value)
                        Dim IsTaxable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                        Dim IsExcisable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value)
                        Dim IsTaxonBaseAmount As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(colIsTaxOnBaseAmount + clsCommon.myCstr(ii)).Value)
                        Dim dblBaseAmt As Double = 0
                        Dim dblTaxAmt As Double = 0
                        If IsSurTax Then
                            Dim dblSurTaxAmt As Double = GetCurrentRowSurTaxAmt(IntRowNo, ii, strSurTaxCode)
                            dblBaseAmt = dblSurTaxAmt
                        Else
                            Dim dblOtherTaxAmt As Double = 0
                            ''richa 21 Sep 2020 changes according to tax
                            ' commented 
                            'If Not IsTaxonBaseAmount Then
                            '    dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)
                            'End If
                            If Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") <> CompairStringResult.Equal Then
                                dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)
                                'dblBaseAmt += dblOtherTaxAmt
                            ElseIf Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") = CompairStringResult.Equal Then
                                dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth1)
                                'dblBaseAmt += dblOtherTaxAmt
                            End If
                            ''If IsExcisable Then
                            ''    dblBaseAmt = (dblAssessableAmt + dblOtherTaxAmt)
                            ''Else
                            dblBaseAmt = (dblAmtAfterDis + dblOtherTaxAmt)
                            ''End If
                        End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Math.Round(dblBaseAmt, 3)
                        dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                        ' gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Math.Round(dblTaxAmt, 2)
                        'If IsTaxable AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper()) Then
                        '    arrTaxableAuth.Add(strTaxCode.ToUpper())
                        'End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Math.Round(dblTaxAmt, 6)
                        If (IsTaxable AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper())) AndAlso (clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "CGST") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "SGST") <> CompairStringResult.Equal) Then
                            arrTaxableAuth.Add(strTaxCode.ToUpper())
                            'ElseIf clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") = CompairStringResult.Equal Then
                            '    arrTaxableAuth.Add(strTaxCode.ToUpper())
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
                        Dim dblTaxAmt As Double = clsCommon.myCdbl(gv2.Rows(ii - 1).Cells(colTTaxAmt).Value)
                        Dim dblCurrRowAmt As Double = clsCommon.myCdbl(gv1.Rows(clsCommon.myCdbl(IntRowNo)).Cells(colAmt).Value)
                        Dim dblTotAmt As Double = 0
                        For jj As Integer = 0 To gv1.Rows.Count - 1
                            dblTotAmt += clsCommon.myCdbl(gv1.Rows(jj).Cells(colAmt).Value)
                        Next
                        Dim dblCurrCalTax As Double = 0
                        If dblTotAmt <> 0 Then
                            dblCurrCalTax = Math.Round(clsCommon.myCdbl(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 2, MidpointRounding.ToEven)
                        End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = dblCurrCalTax
                    End If
                End If
            Next
            Dim isSampling As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(isSampling,0) from tspl_sd_shipment_head where Sale_invoice_No='" & txtReqNo.Value & "'"))
            If clsCommon.CompairString(isSampling, "1") = CompairStringResult.Equal Then
                dblAmtAfterDis = 0
                dblTotDiscAmt = 0
            End If
            Dim dblTotTaxAmt As Double = GetCurrentRowTotalTaxAmt(IntRowNo)
            Dim dblTotTCSTaxAmt As Double = GetCurrentRowTotalTCSTaxAmt(IntRowNo)
            Dim dblAmtAfterTax As Double = dblAmtAfterDis + dblTotTaxAmt + dblTotTCSTaxAmt
            gv1.Rows(IntRowNo).Cells(colAlterUnitQty).Value = Math.Round(dblAlterQty, 2)
            gv1.Rows(IntRowNo).Cells(colRateUnitQty).Value = Math.Round(dblQty, 2)

            gv1.Rows(IntRowNo).Cells(colDisAmt).Value = Math.Round(dblDisAmt, 2)
            gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = Math.Round(dblAmtAfterDis, 2)
            gv1.Rows(IntRowNo).Cells(colTotTaxAmt).Value = Math.Round(dblTotTaxAmt, 2)
            gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = Math.Round(dblAmtAfterTax, 2)
            gv1.Rows(IntRowNo).Cells(colAbatementAmount).Value = Math.Round(dblAbatementAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalMRP).Value = Math.Round(dblMRPAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalBasicAmount).Value = Math.Round(dblBasicAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotItemWt).Value = Math.Round(dblConvF * dblItemWeight * dblQty, 2)
            'gv1.Rows(IntRowNo).Cells(colTotalCustDiscount).Value = Math.Round(dblTotCustDisc, 2)
            'gv1.Rows(IntRowNo).Cells(colRate).Value = dblRate
            gv1.Rows(IntRowNo).Cells(colHeadDisPerAmt).Value = Math.Round(dblHeadPerDisAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalDiscountAmount).Value = Math.Round(dblTotDiscAmt, 2)
            gv1.Rows(IntRowNo).Cells(colOrgUnit).Value = strOrgUnit

            'If AutoCalculateCrate = 1 And isInvoiceLoadData = False Then
            If AutoCalculateCrate = 1 Then
                If clsCommon.myLen(strICode) > 0 Then 'AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0
                    '' Anubhooti 11-Sep-2014 BM00000003847
                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "'"))
                    If ItemCrateType = 1 Then
                        Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "'"))
                        'If IsStockingUnit = "Y" Then
                        Dim CrateConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and tspl_unit_master.Crate_Type ='Y' "))
                        Dim ItemConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "' "))

                        If CrateConvFactor > 0 And ItemConvFactor > 0 Then
                            Dim DispatchQty As Double = gv1.Rows(IntRowNo).Cells(colQty).Value * ItemConvFactor
                            If DispatchQty >= CrateConvFactor Then
                                If IncreaseCrateQtyOnFiftyPercent = True Then
                                    Dim IntegerPart As Integer = Math.Floor(DispatchQty / CrateConvFactor)
                                    Dim fractionPart As Integer = ((DispatchQty / CrateConvFactor) - IntegerPart) * 100
                                    If fractionPart >= 50 Then
                                        gv1.Rows(IntRowNo).Cells(colCrate).Value = Math.Ceiling(DispatchQty / CrateConvFactor)
                                    Else
                                        gv1.Rows(IntRowNo).Cells(colCrate).Value = Math.Floor(DispatchQty / CrateConvFactor)
                                    End If
                                Else
                                    gv1.Rows(IntRowNo).Cells(colCrate).Value = Math.Floor(DispatchQty / CrateConvFactor)
                                End If
                            Else
                                gv1.Rows(IntRowNo).Cells(colCrate).Value = 0
                            End If
                        Else
                            clsCommon.MyMessageBoxShow(Me, "Please fill conversion factor for this unit at line no." & IntRowNo + 1 & "")
                        End If
                    End If
                End If
                Dim TotalCrate As Integer = 0
                For i As Integer = 0 To gv1.Rows.Count - 1
                    TotalCrate = TotalCrate + gv1.Rows(i).Cells(colCrate).Value
                Next
                If clsCommon.myCdbl(TotalCrate) > 0 Then
                    txtCrate.Value = TotalCrate
                Else
                    txtCrate.Value = 0
                End If

            End If

            If AutoCalculateCAN = 1 Then
                If clsCommon.myLen(strICode) > 0 Then 'AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0
                    '' Anubhooti 11-Sep-2014 BM00000003847
                    Dim ItemCanType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Is_CAN_Type  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "'"))
                    If ItemCanType = 1 Then
                        Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "'"))
                        'If IsStockingUnit = "Y" Then
                        Dim CanConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and tspl_unit_master.Can_Type ='Y' "))
                        Dim ItemConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "' "))

                        If CanConvFactor > 0 And ItemConvFactor > 0 Then
                            Dim DispatchQty As Double = gv1.Rows(IntRowNo).Cells(colQty).Value * ItemConvFactor
                            If DispatchQty >= CanConvFactor Then
                                If IncreaseCrateQtyOnFiftyPercent = True Then
                                    Dim IntegerPart As Integer = Math.Floor(DispatchQty / CanConvFactor)
                                    Dim fractionPart As Integer = ((DispatchQty / CanConvFactor) - IntegerPart) * 100
                                    If fractionPart >= 50 Then
                                        gv1.Rows(IntRowNo).Cells(colCan).Value = Math.Ceiling(DispatchQty / CanConvFactor)
                                    Else
                                        gv1.Rows(IntRowNo).Cells(colCan).Value = Math.Floor(DispatchQty / CanConvFactor)
                                    End If
                                Else
                                    gv1.Rows(IntRowNo).Cells(colCan).Value = Math.Floor(DispatchQty / CanConvFactor)
                                End If
                            Else
                                gv1.Rows(IntRowNo).Cells(colCan).Value = 0
                            End If
                        Else
                            clsCommon.MyMessageBoxShow(Me, "Please fill conversion factor for this unit at line no." & IntRowNo + 1 & "")
                        End If
                    End If
                End If

                ' for Total Can
                Dim TotalCan As Integer = 0
                For i As Integer = 0 To gv1.Rows.Count - 1
                    TotalCan = TotalCan + gv1.Rows(i).Cells(colCan).Value
                Next
                If clsCommon.myCdbl(TotalCan) > 0 Then
                    TxtTotalCAN.Value = TotalCan
                Else
                    TxtTotalCAN.Value = 0
                End If
            End If


        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub UpdateCurrentRow1(ByVal IntRowNo As Integer)
        Try
            Dim arrTaxableAuth As New List(Of String)
            Dim arrTaxableAuth1 As New List(Of String)
            Dim dblFAmt As Double = 0


            'If chkVendorGrossReceipt.Checked OrElse clsCommon.CompairString(cboItemType.SelectedValue, "F") = CompairStringResult.Equal Then
            '    dblQty = dblQty
            'End If

            Dim dblAlterQty As Double = 0
            Dim dblQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value)
            Dim dblActualQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colActualQty).Value)
            Dim dblDamageQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDamageQty).Value)
            Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblMRP As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colMRP).Value)
            Dim dblBasicRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colItemBasicPrice).Value)
            Dim dblConvF As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colConvF).Value)
            Dim dblItemWeight As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colItemWeight).Value)
            Dim dblheadDiscamt As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colHeadDiscamt).Value)
            Dim dblOrgBasicRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colOrgCost).Value)
            'Dim dblConvBasicRate As Double = dblOrgBasicRate * dblConvF
            Dim dblMRPAmt As Double = dblQty * dblBasicRate


            Dim strICode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
            Dim strSchemeCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colFromSchemeCode).Value)
            If clsCommon.myLen(strICode) > 0 And clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0 Then

                'Dim obj_Cash As clsSchemeApplyOnDairy = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnitRate).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value)
                Dim obj_Cash As clsSchemeApplyOnDairy = Nothing
                'obj_Cash = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value, strSchemeCode)
                obj_Cash = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value, strSchemeCode)
                If clsCommon.myLen(obj_Cash.Schm_Code) = 0 AndAlso clsCommon.myLen(gv1.Rows(IntRowNo).Cells(colUnitALter).Value) > 0 Then
                    obj_Cash = clsSchemeApplyOnDairy.GetPriceSchemeData(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnitALter).Value), clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value), txtVendorNo.Value, strSchemeCode)
                End If
                If obj_Cash IsNot Nothing Then
                    gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = obj_Cash.Cash_Amt
                    gv1.Rows(IntRowNo).Cells(colCash_Pers).Value = obj_Cash.Cash_Pers
                    gv1.Rows(IntRowNo).Cells(colCashSchemeCode).Value = obj_Cash.Schm_Code
                    If clsCommon.myCdbl(obj_Cash.Cash_Pers) > 0 Then
                        gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = "P"
                        gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = System.Math.Round((clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colCash_Pers).Value)) / 100, 2)
                    ElseIf clsCommon.myCdbl(obj_Cash.Cash_Amt) > 0 Then
                        gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = "A"
                    End If
                Else
                    gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = Nothing
                    gv1.Rows(IntRowNo).Cells(colCash_Pers).Value = Nothing
                    gv1.Rows(IntRowNo).Cells(colCashSchemeCode).Value = Nothing
                    gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = Nothing
                End If
            End If
            Dim dblCashAmt As Decimal = gv1.Rows(IntRowNo).Cells(colCash_Amt).Value
            Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblBasicAmt As Decimal = dblQty * dblBasicRate
            Dim dblReturnAmt As Decimal = (dblActualQty * dblBasicRate)
            Dim dblDamageAmt As Decimal = (dblDamageQty * dblBasicRate)
            Dim dblAmt As Decimal = dblReturnAmt + dblDamageAmt ''+ dblFAmt
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colIsMannualAmt).Value) = 0 Then
                gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2) 'dblAmt 
                gv1.Rows(IntRowNo).Cells(colReturnAmt).Value = Math.Round(dblReturnAmt, 2) 'dblReturnAmt 
                gv1.Rows(IntRowNo).Cells(colDamageAmt).Value = Math.Round(dblDamageAmt, 2) 'dblDamageAmt 
            Else
                dblAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAmt).Value)
            End If

            If clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colCash_Amt).Value) > dblAmt Then
                gv1.Rows(IntRowNo).Cells(colCash_Amt).Value = 0
                gv1.Rows(IntRowNo).Cells(colCash_Pers).Value = 0
                gv1.Rows(IntRowNo).Cells(colCashSchemeCode).Value = Nothing
                gv1.Rows(IntRowNo).Cells(colCashSchemeType).Value = Nothing
            End If

            GetDCDetails(IntRowNo)
            Dim dblTotalDCAmt As Decimal = 0
            Dim dblTotalTCAmt As Decimal = 0
            Dim dblTotTaxRate As Decimal = GetCurrentRowTotalTaxRate(IntRowNo)
            Dim dblDisPer As Decimal = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblSCRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColSCRate).Value)
            Dim dblDisAmt As Decimal = (dblAmt * dblDisPer) / 100
            'Dim dblSCAmt As Double = dblAmt * (dblSCRate / 100)
            'gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = dblSCAmt
            If Not gv1.Rows(IntRowNo).Cells(ColDCRate).Value = Nothing AndAlso clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColDCRate).Value) >= 0 Then
                If Not ApplyCommissionRateWithTax Then
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = gv1.Rows(IntRowNo).Cells(ColDCRate).Value
                Else
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = Math.Round(gv1.Rows(IntRowNo).Cells(ColDCRate).Value * 100 / (100 + dblTotTaxRate), 4)
                End If
                gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value = (gv1.Rows(IntRowNo).Cells(colQty).Value * gv1.Rows(IntRowNo).Cells(ColDCUnitCF).Value) / gv1.Rows(IntRowNo).Cells(ColDCCFUOM).Value
                gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value
                gv1.Rows(IntRowNo).Cells(ColTCAmt).Value = clsCommon.myCstr(clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColTCRate).Value))
                gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = clsCommon.myCstr(clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColSCRate).Value))
                dblTotalDCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCAmt).Value)
                dblTotalTCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColTCAmt).Value)
                If dblTotalDCAmt > 0 Then
                    ' GetBoothWiseDCDetails(IntRowNo, trans)
                    If ApplyCommission Then
                        dblDisAmt = dblDisAmt + dblTotalDCAmt
                    End If
                End If
            End If

            'Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblDisAmt As Double = (dblAmt * dblDisPer) / 100
            Dim dblHeadDisPer As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colHeaDDisPer).Value)
            Dim dblHeadPerDisAmt As Decimal = (dblAmt * dblHeadDisPer) / 100

            Dim dblTotDiscAmt = dblheadDiscamt + dblHeadPerDisAmt + dblDisAmt + dblCashAmt
            Dim dblAmtAfterDis As Decimal = dblAmt - dblDisAmt - dblheadDiscamt - dblHeadPerDisAmt - dblCashAmt
            Dim dblAbatementRate As Decimal = gv1.CurrentRow.Cells(colAbatementPer).Value
            Dim dblAbatementAmt As Decimal = ((dblMRP * dblAbatementRate) / 100) * dblQty

            Dim dblTotalNonTabxableRate As Decimal = 0
            Dim dblTotalNonTabxableAmount As Decimal = 0
            Dim dblKKFTaxRate As Decimal = 0
            Dim dblMNDTaxRate As Decimal = 0
            Dim dblGSTTaxRate As New List(Of Decimal)
            Dim dblGSTTaxValue1 As Decimal = 0
            Dim dblGSTTaxValue2 As Decimal = 0
            Dim dblKKFTaxValue As Decimal = 0
            Dim dblMNDTaxValue As Decimal = 0
            Dim dblTotalTaxValue As Decimal = 0
            Dim dblKKFMNDBaseAmt As Decimal = 0
            Dim dblTaxableValue As Decimal = 0
            Dim dblProductValue As Decimal = 0

            For ii As Integer = 1 To 10
                Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells("COLTAX" + clsCommon.myCstr(ii)).Value)
                If clsCommon.myLen(strTaxCode) > 0 Then
                    If clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "K") = CompairStringResult.Equal Then
                        dblKKFTaxRate = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("COLTAXRATE" + clsCommon.myCstr(ii)).Value)
                    ElseIf clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "M") = CompairStringResult.Equal Then
                        dblMNDTaxRate = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("COLTAXRATE" + clsCommon.myCstr(ii)).Value)
                    ElseIf clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "GST") = CompairStringResult.Equal Then
                        dblGSTTaxRate.Add(clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("COLTAXRATE" + clsCommon.myCstr(ii)).Value))
                    End If
                    'If Not clsCommon.myCBool(gv1.CurrentRow.Cells(colIsTaxable + clsCommon.myCstr(ii)).Value) OrElse (clsCommon.CompairString(strTaxCode, "CGST") = CompairStringResult.Equal Or clsCommon.CompairString(strTaxCode, "SGST") = CompairStringResult.Equal) Then
                    '        dblTotalNonTabxableRate = dblTotalNonTabxableRate + clsCommon.myCDecimal(gv1.CurrentRow.Cells(colTaxRate + clsCommon.myCstr(ii)).Value)
                    '    End If
                End If
            Next
            If dblGSTTaxRate.Count = 1 Then
                dblGSTTaxValue1 = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblBasicAmt / (100 + dblGSTTaxRate(0)) * dblGSTTaxRate(0), 3), 2, 4)
            ElseIf dblGSTTaxRate.Count = 2 Then
                dblGSTTaxValue1 = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblBasicAmt / (100 + dblGSTTaxRate(0) + dblGSTTaxRate(1)) * dblGSTTaxRate(0), 3), 2, 4)
                dblGSTTaxValue2 = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblBasicAmt / (100 + dblGSTTaxRate(0) + dblGSTTaxRate(1)) * dblGSTTaxRate(1), 3), 2, 4)
            End If
            dblKKFMNDBaseAmt = clsCommon.myRoundOFF(dblBasicAmt - (dblGSTTaxValue1 + dblGSTTaxValue2), 2, 4)
            dblKKFTaxValue = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblKKFMNDBaseAmt / (100 + dblKKFTaxRate + dblMNDTaxRate) * dblKKFTaxRate, 3), 2, 4)
            dblMNDTaxValue = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblKKFMNDBaseAmt / (100 + dblKKFTaxRate + dblMNDTaxRate) * dblMNDTaxRate, 3), 2, 4)
            dblTotalTaxValue = dblGSTTaxValue1 + dblGSTTaxValue2 + dblKKFTaxValue + dblMNDTaxValue
            dblTaxableValue = dblBasicAmt - (dblGSTTaxValue1 + dblGSTTaxValue2)
            dblProductValue = dblTaxableValue - (dblKKFTaxValue + dblMNDTaxValue)
            gv1.Rows(IntRowNo).Cells(colAmt).Value = dblBasicAmt - dblTotalTaxValue
            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)
                If rbtnTaxCalAutomatic.IsChecked Then
                    Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                    If clsCommon.myLen(strTaxCode) > 0 Then
                        Dim dblTaxRate As Double = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                        Dim IsSurTax As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value)
                        Dim strSurTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value)
                        Dim IsTaxable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                        Dim IsExcisable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value)
                        Dim IsTaxonBaseAmount As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(colIsTaxOnBaseAmount + clsCommon.myCstr(ii)).Value)
                        Dim dblBaseAmt As Decimal = 0
                        Dim dblTaxAmt As Decimal = 0

                        If clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "K") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTaxAmt" + clsCommon.myCstr(ii)).Value = dblKKFTaxValue
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = clsCommon.myRoundOFF(dblProductValue, 2, 4)
                        ElseIf clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "M") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTaxAmt" + clsCommon.myCstr(ii)).Value = dblMNDTaxValue
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = clsCommon.myRoundOFF(dblProductValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "CGST") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTaxAmt" + clsCommon.myCstr(ii)).Value = dblGSTTaxValue1
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "SGST") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTaxAmt" + clsCommon.myCstr(ii)).Value = dblGSTTaxValue2
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "IGST") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTaxAmt" + clsCommon.myCstr(ii)).Value = dblGSTTaxValue1
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "EXEMPTED") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTaxAmt" + clsCommon.myCstr(ii)).Value = 0
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 3, 4)

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
                        Dim dblTaxAmt As Double = clsCommon.myCdbl(gv2.Rows(ii - 1).Cells(colTTaxAmt).Value)
                        Dim dblCurrRowAmt As Double = clsCommon.myCdbl(gv1.Rows(clsCommon.myCdbl(IntRowNo)).Cells(colAmt).Value)
                        Dim dblTotAmt As Double = 0
                        For jj As Integer = 0 To gv1.Rows.Count - 1
                            dblTotAmt += clsCommon.myCdbl(gv1.Rows(jj).Cells(colAmt).Value)
                        Next
                        Dim dblCurrCalTax As Double = 0
                        If dblTotAmt <> 0 Then
                            dblCurrCalTax = Math.Round(clsCommon.myCdbl(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 2, MidpointRounding.ToEven)
                        End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = dblCurrCalTax
                    End If
                End If
            Next
            Dim isSampling As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(isSampling,0) from tspl_sd_shipment_head where Sale_invoice_No='" & txtReqNo.Value & "'"))
            If clsCommon.CompairString(isSampling, "1") = CompairStringResult.Equal Then
                dblAmtAfterDis = 0
                dblTotDiscAmt = 0
            End If
            Dim dblTotTaxAmt As Double = GetCurrentRowTotalTaxAmt(IntRowNo)
            Dim dblAmtAfterTax As Double = dblAmtAfterDis + dblTotTaxAmt
            gv1.Rows(IntRowNo).Cells(colAlterUnitQty).Value = clsCommon.myRoundOFF(dblAlterQty, 2, 4)
            gv1.Rows(IntRowNo).Cells(colRateUnitQty).Value = clsCommon.myRoundOFF(dblQty, 2, 4)

            gv1.Rows(IntRowNo).Cells(colDisAmt).Value = clsCommon.myRoundOFF(dblDisAmt, 2, 4)
            gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblAmtAfterDis - (dblTotTaxAmt), 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colTotTaxAmt).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblTotTaxAmt, 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colAmtAfterTax).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblAmtAfterTax - dblTotTaxAmt, 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colAbatementAmount).Value = Math.Round(dblAbatementAmt, 2)
            gv1.Rows(IntRowNo).Cells(colTotalMRP).Value = clsCommon.myRoundOFF(dblMRPAmt, 2, 4)
            gv1.Rows(IntRowNo).Cells(colTotalBasicAmount).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblBasicAmt, 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colTotItemWt).Value = clsCommon.myRoundOFF(dblConvF * dblItemWeight * dblQty, 2, 4)
            'gv1.Rows(IntRowNo).Cells(colTotalCustDiscount).Value = Math.Round(dblTotCustDisc, 2)
            'gv1.Rows(IntRowNo).Cells(colRate).Value = dblRate
            gv1.Rows(IntRowNo).Cells(colHeadDisPerAmt).Value = clsCommon.myRoundOFF(dblHeadPerDisAmt, 2, 4)
            gv1.Rows(IntRowNo).Cells(colTotalDiscountAmount).Value = clsCommon.myRoundOFF(dblTotDiscAmt, 2, 4)
            gv1.Rows(IntRowNo).Cells(colOrgUnit).Value = strOrgUnit

            'If AutoCalculateCrate = 1 And isInvoiceLoadData = False Then
            If AutoCalculateCrate = 1 Then
                If clsCommon.myLen(strICode) > 0 Then 'AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0
                    '' Anubhooti 11-Sep-2014 BM00000003847
                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "'"))
                    If ItemCrateType = 1 Then
                        Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "'"))
                        'If IsStockingUnit = "Y" Then
                        Dim CrateConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and tspl_unit_master.Crate_Type ='Y' "))
                        Dim ItemConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "' "))

                        If CrateConvFactor > 0 And ItemConvFactor > 0 Then
                            Dim DispatchQty As Double = gv1.Rows(IntRowNo).Cells(colQty).Value * ItemConvFactor
                            If DispatchQty >= CrateConvFactor Then
                                If IncreaseCrateQtyOnFiftyPercent = True Then
                                    Dim IntegerPart As Integer = Math.Floor(DispatchQty / CrateConvFactor)
                                    Dim fractionPart As Integer = ((DispatchQty / CrateConvFactor) - IntegerPart) * 100
                                    If fractionPart >= 50 Then
                                        gv1.Rows(IntRowNo).Cells(colCrate).Value = Math.Ceiling(DispatchQty / CrateConvFactor)
                                    Else
                                        gv1.Rows(IntRowNo).Cells(colCrate).Value = Math.Floor(DispatchQty / CrateConvFactor)
                                    End If
                                Else
                                    gv1.Rows(IntRowNo).Cells(colCrate).Value = Math.Floor(DispatchQty / CrateConvFactor)
                                End If
                            Else
                                gv1.Rows(IntRowNo).Cells(colCrate).Value = 0
                            End If
                        Else
                            clsCommon.MyMessageBoxShow(Me, "Please fill conversion factor for this unit at line no." & IntRowNo + 1 & "")
                        End If
                    End If
                End If
                Dim TotalCrate As Integer = 0
                For i As Integer = 0 To gv1.Rows.Count - 1
                    TotalCrate = TotalCrate + gv1.Rows(i).Cells(colCrate).Value
                Next
                If clsCommon.myCDecimal(TotalCrate) > 0 Then
                    txtCrate.Value = TotalCrate
                Else
                    txtCrate.Value = 0
                End If

            End If

            If AutoCalculateCAN = 1 Then
                If clsCommon.myLen(strICode) > 0 Then 'AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColFOC).Value) = 0
                    '' Anubhooti 11-Sep-2014 BM00000003847
                    Dim ItemCanType As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Is_CAN_Type  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "'"))
                    If ItemCanType = 1 Then
                        Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "'"))
                        'If IsStockingUnit = "Y" Then
                        Dim CanConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and tspl_unit_master.Can_Type ='Y' "))
                        Dim ItemConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "' "))

                        If CanConvFactor > 0 And ItemConvFactor > 0 Then
                            Dim DispatchQty As Double = gv1.Rows(IntRowNo).Cells(colQty).Value * ItemConvFactor
                            If DispatchQty >= CanConvFactor Then
                                If IncreaseCrateQtyOnFiftyPercent = True Then
                                    Dim IntegerPart As Integer = Math.Floor(DispatchQty / CanConvFactor)
                                    Dim fractionPart As Integer = ((DispatchQty / CanConvFactor) - IntegerPart) * 100
                                    If fractionPart >= 50 Then
                                        gv1.Rows(IntRowNo).Cells(colCan).Value = Math.Ceiling(DispatchQty / CanConvFactor)
                                    Else
                                        gv1.Rows(IntRowNo).Cells(colCan).Value = Math.Floor(DispatchQty / CanConvFactor)
                                    End If
                                Else
                                    gv1.Rows(IntRowNo).Cells(colCan).Value = Math.Floor(DispatchQty / CanConvFactor)
                                End If
                            Else
                                gv1.Rows(IntRowNo).Cells(colCan).Value = 0
                            End If
                        Else
                            clsCommon.MyMessageBoxShow(Me, "Please fill conversion factor for this unit at line no." & IntRowNo + 1 & "")
                        End If
                    End If
                End If

                ' for Total Can
                Dim TotalCan As Integer = 0
                For i As Integer = 0 To gv1.Rows.Count - 1
                    TotalCan = TotalCan + gv1.Rows(i).Cells(colCan).Value
                Next
                If clsCommon.myCDecimal(TotalCan) > 0 Then
                    TxtTotalCAN.Value = TotalCan
                Else
                    TxtTotalCAN.Value = 0
                End If
            End If


        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub RadMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem1.Click
        If clsCommon.myLen(MyBase.Form_ID) > 0 Then
            gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = MyBase.Form_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv1.SaveLayout(obj.GridLayout)
            obj.GridColumns = gv1.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
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
        'clsGridLayout.DeleteData(ReportID, objCommonVar.CurrentUserCode)
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
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
    Private Sub gv2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv2.DoubleClick
        Try
            If True Then
                Dim dblNewRate As Double = clsLocationWiseTax.FinderForTaxRate(txtBillToLocation.Value, txtTaxGroup.Value, clsCommon.myCstr(gv2.CurrentRow.Cells(colTTaxAutCode).Value), txtVendorNo.Value, "S")
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
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
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
            'common.clsCommon.MyMessageBoxShow(ex.Message)
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
                gv1.CurrentRow.Cells(colIsMannualAmt).Value = IIf(clsCommon.myCdbl(gv1.CurrentRow.Cells(colIsMannualAmt).Value) = 1, 0, 1)
            End If

            If clsCommon.myCdbl(gv1.CurrentRow.Cells(colIsMannualAmt).Value) = 0 Then
                UpdateCurrentRow(gv1.CurrentRow.Index)
            End If
        ElseIf e.KeyCode = Keys.F4 Then
            OpenSerialItem()
        ElseIf e.KeyCode = Keys.F5 Then
            OpenBatchItem()
        End If
    End Sub

    Private Sub gv1_RowFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.RowFormattingEventArgs) Handles gv1.RowFormatting
        Try
            If clsCommon.CompairString(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colRowType).Value), RowTypeItem) = CompairStringResult.Equal AndAlso clsCommon.myCdbl(e.RowElement.RowInfo.Cells(colIsMannualAmt).Value) > 0 Then
                e.RowElement.ForeColor = Color.Blue
            Else
                e.RowElement.ForeColor = Color.Black
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub


    Private Sub gv1_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gv1.CurrentColumnChanged
        If Not clsCommon.myLen(txtReqNo.Value) > 0 Then
            If gv1.RowCount > 0 Then
                Dim intCurrRow As Integer = gv1.CurrentRow.Index
                gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
                If intCurrRow = gv1.Rows.Count - 1 Then
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                    gv1.CurrentRow = gv1.Rows(intCurrRow)
                End If
            End If
        End If
    End Sub


    Private Sub txtSalesman__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtSalesman._MYValidating
        Dim qry As String = "select EMP_CODE as Code,Emp_Name as Name from TSPL_EMPLOYEE_MASTER"
        Dim whrcls As String = "Emp_type='Salesman'"
        txtSalesman.Value = clsCommon.ShowSelectForm("PSSaleRetSalesman", qry, "Code", whrcls, txtSalesman.Value, "Code", isButtonClicked)
        lblSalesman.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Emp_Name as Name from TSPL_EMPLOYEE_MASTER where EMP_CODE ='" + txtSalesman.Value + "' and Emp_type='Salesman'"))
    End Sub

    Private Sub txtCurrencyCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtCurrencyCode._MYValidating
        Dim qry As String = "select CURRENCY_CODE AS Code, CURRENCY_NAME as Name from TSPL_CURRENCY_MASTER"
        txtCurrencyCode.Value = clsCommon.ShowSelectForm("CURRENCY_MASTER", qry, "Code", "", txtCurrencyCode.Value, "CURRENCY_CODE", isButtonClicked)
        ShowCurrencyDetail()
    End Sub

    Private Sub txtVendorNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVendorNo.Load

    End Sub

    '----------------Done By Monika 30/04/2014----------------------------------------------------
#Region "Email SMS Setting"
    Private Sub RadMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem5.Click
        Dim frm As New FrmMailSMSSettingNew2()
        frm.FormId = clsUserMgtCode.frmSaleReturnProductSale
        frm.ShowDialog()
    End Sub

    Public Function GetAttachmentQuery(ByVal strcode As String)
        atchmntqry = "  select '" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MM/yyyy") + "' as RunDate, TSPL_SD_SALE_RETURN_HEAD.Inv_No, TSPL_SD_SALE_RETURN_HEAD.Dept_Desc , TSPL_SD_SALE_RETURN_HEAD.Remarks ,  TSPL_SD_SALE_RETURN_HEAD.Terms_Code,TSPL_SD_SALE_RETURN_HEAD.VehicleNo , "
        atchmntqry += " TSPL_SD_SALE_RETURN_DETAIL .Specification as  specification,   TSPL_SD_SALE_RETURN_HEAD.Document_Code as DocNo , TSPL_SD_SALE_RETURN_HEAD.Description, "
        atchmntqry += " convert(varchar(15),TSPL_SD_SALE_RETURN_HEAD.Document_Date,106) as Document_Date, TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No, TSPL_SD_SALE_INVOICE_HEAD.Total_Amt as SaleInvoiceAmt , TSPL_SD_SALE_RETURN_HEAD.Item_Type ,  TSPL_SD_SALE_RETURN_HEAD.Customer_Code, "
        atchmntqry += " TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_SD_SALE_RETURN_HEAD .Terms_Code as termscode ,TSPL_SD_SALE_RETURN_HEAD .Ref_No as ref_no ,"
        atchmntqry += " TSPL_SD_SALE_RETURN_HEAD .Comments as comments ,  TSPL_SD_SALE_RETURN_HEAD .Discount_Amt as dis_amt,TSPL_SD_SALE_RETURN_DETAIL .Disc_Amt  as dis_amt1,"
        atchmntqry += " TSPL_SD_SALE_RETURN_HEAD.Amount_Less_Discount  as aftrdiscount ,TSPL_SD_SALE_RETURN_HEAD .Total_Amt as Total_amount,"
        atchmntqry += " TSPL_SD_SALE_RETURN_HEAD.Discount_Base as bfrdisc_amount,  "
        atchmntqry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax1_amt,0) as txt1amt,  "
        atchmntqry += " tax2.Tax_Code_Desc as tax2name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax2_amt,0) as txt2amt,  "
        atchmntqry += " tax3.Tax_Code_Desc as tax3name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax3_amt,0) as txt3amt,  "
        atchmntqry += " tax4.Tax_Code_Desc as tax4name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax4_amt,0) as txt4amt,  "
        atchmntqry += " tax5.Tax_Code_Desc as tax5name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax5_amt,0) as txt5amt,  "
        atchmntqry += " tax6.Tax_Code_Desc as tax6name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax6_amt,0) as txt6amt,  "
        atchmntqry += " tax7.Tax_Code_Desc as tax7name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax7_amt,0) as txt7amt,  "
        atchmntqry += " tax8.Tax_Code_Desc as tax8name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax8_amt,0) as txt8amt,   "
        atchmntqry += " tax9.Tax_Code_Desc as tax9name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax9_amt,0) as txt9amt,  "
        atchmntqry += " tax10.Tax_Code_Desc as tax10name,isnull (TSPL_SD_SALE_RETURN_HEAD.tax10_amt,0) as txt10amt,  "
        atchmntqry += " isnull(TSPL_SD_SALE_RETURN_HEAD .Total_Tax_Amt,0) as total_tax_amt, TSPL_SD_SALE_RETURN_HEAD.Total_Amt as DocAmt,  TSPL_COMPANY_MASTER.Comp_Name as compname,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,ISNULL(tspl_company_Master.ADD1,'') as address1,"
        atchmntqry += " TSPL_SD_SALE_RETURN_DETAIL.item_code as item_code, TSPL_ITEM_MASTER.Item_Desc   as itemdesc, TSPL_SD_SALE_RETURN_DETAIL.Row_Type,TSPL_SD_SALE_RETURN_DETAIL.Qty as qty,TSPL_SD_SALE_RETURN_DETAIL.unit_code as uom,TSPL_SD_SALE_RETURN_DETAIL.item_cost as itemcost,TSPL_SD_SALE_RETURN_DETAIL.amount as amount,TSPL_SD_SALE_RETURN_HEAD.TAX1,TSPL_SD_SALE_RETURN_HEAD.TAX2,TSPL_SD_SALE_RETURN_HEAD.TAX3,TSPL_SD_SALE_RETURN_HEAD.TAX4,TSPL_SD_SALE_RETURN_HEAD.TAX5,TSPL_SD_SALE_RETURN_HEAD.Total_Add_Charge, "
        atchmntqry += " TSPL_SD_SALES_ORDER_HEAD.Document_Code as SalesOrderNo, TSPL_SD_SALES_ORDER_HEAD.Remarks as SO_Head_Remarks, TSPL_SD_SHIPMENT_HEAD.Document_Code as ShippmentNo, convert(varchar(15),TSPL_SD_SHIPMENT_HEAD.Document_Date,106) as ShippmentDate, TSPL_Customer_Invoice_Head.Document_No as CustomerInvoiceNo, " &
               " convert(varchar(15),TSPL_Customer_Invoice_Head.Document_Date,106) as CustomerInvoiceDate,TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No as SaleInvoiceNo, " &
               " convert(varchar(15),TSPL_SD_SALE_INVOICE_HEAD.Inv_Date,106) as SaleInvoiceDate, TSPL_CUSTOMER_MASTER.Customer_Name, (case when ISNULL(TSPL_CUSTOMER_MASTER.ADD1,'')<> '' then TSPL_CUSTOMER_MASTER.ADD1 else '' end + case when ISNULL(TSPL_CUSTOMER_MASTER.ADD2,'')<> '' then ', ' +TSPL_CUSTOMER_MASTER.ADD2 else '' end + case when ISNULL(TSPL_CUSTOMER_MASTER.ADD3,'')<> '' then ', ' +TSPL_CUSTOMER_MASTER.ADD3 else '' end ) as CustAddress,	" &
               " TSPL_COMPANY_MASTER.Tin_No as TinNo, TSPL_COMPANY_MASTER.CST_LST as CstNo, (case when ISNULL(TSPL_COMPANY_MASTER.ADD1,'')<> '' then TSPL_COMPANY_MASTER.ADD1 else '' end + case when ISNULL(TSPL_COMPANY_MASTER.ADD2,'')<> '' then ', ' +TSPL_COMPANY_MASTER.ADD2 else '' end + " &
               " case when ISNULL(TSPL_COMPANY_MASTER.ADD3,'')<> '' then ', ' +TSPL_COMPANY_MASTER.ADD3 else '' end ) as CompAddress, " &
               " COALESCE(TSPL_SD_SALES_ORDER_DETAIL.Qty,0) as OrdQty, COALESCE(TSPL_SD_SALE_RETURN_DETAIL.MRP,0) AS MRP, TSPL_SD_SALE_RETURN_DETAIL.unit_code as UOM, " &
               " COALESCE(TSPL_SD_SALE_RETURN_DETAIL.Qty,0) as ReturnQty, COALESCE(TSPL_SD_SALE_INVOICE_DETAIL.Qty,0) as CustInvQty, (COALESCE(TSPL_SD_SALE_INVOICE_DETAIL.Qty,0) - COALESCE(TSPL_SD_SALE_RETURN_DETAIL.Qty,0)) as Difference, TSPL_SD_SALE_RETURN_DETAIL.Remarks, TSPL_ITEM_BARCODE.Bar_Code "
        atchmntqry += " from TSPL_SD_SALE_RETURN_DETAIL  "
        atchmntqry += " LEFT outer join TSPL_SD_SALE_RETURN_HEAD  on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code   "
        atchmntqry += " LEFT outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_RETURN_HEAD.tax1  "
        atchmntqry += " left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_RETURN_HEAD.tax2  "
        atchmntqry += " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD .TAX3  "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_RETURN_HEAD .tax4  "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD .tax5  "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX6  "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX7  "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX8  "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX9 "
        atchmntqry += " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX10     "
        atchmntqry += " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_SD_SALE_RETURN_HEAD.comp_code  "
        atchmntqry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_RETURN_HEAD.Customer_Code   "
        atchmntqry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location "
        atchmntqry += " Left Outer Join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_DETAIL.Item_Code "
        atchmntqry += " LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No "
        atchmntqry += " LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code and TSPL_SD_SALE_INVOICE_DETAIL.LINE_NO = TSPL_SD_SALE_RETURN_DETAIL.LINE_NO  and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code= TSPL_SD_SALE_RETURN_DETAIL.Item_Code " &
               " LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No = TSPL_SD_SALE_INVOICE_HEAD.Document_Code " &
               " LEFT OUTER JOIN TSPL_SD_SALES_ORDER_HEAD ON TSPL_SD_SALES_ORDER_HEAD.Document_Code= TSPL_SD_SHIPMENT_HEAD.Against_Sales_Order " &
               " LEFT OUTER JOIN TSPL_SD_SALES_ORDER_DETAIL on TSPL_SD_SALES_ORDER_DETAIL.Document_Code = TSPL_SD_SALES_ORDER_HEAD.Document_Code and TSPL_SD_SALES_ORDER_DETAIL.LINE_NO = TSPL_SD_SALE_RETURN_DETAIL.LINE_NO and TSPL_SD_SALES_ORDER_DETAIL.Item_Code = TSPL_SD_SALE_RETURN_DETAIL.Item_Code " &
               " LEFT OUTER JOIN TSPL_Customer_Invoice_Head ON TSPL_Customer_Invoice_Head.Against_Sale_Return_No = TSPL_SD_SALE_RETURN_HEAD.Document_Code " &
               " LEFT OUTER JOIN TSPL_ITEM_BARCODE ON TSPL_ITEM_BARCODE.Item_Code = TSPL_SD_SALE_RETURN_DETAIL.Item_Code AND " &
               " TSPL_ITEM_BARCODE.Item_MRP = TSPL_SD_SALE_RETURN_DETAIL.MRP "
        atchmntqry += " where 2=2  and  TSPL_SD_SALE_RETURN_HEAD.Document_Code = '" + strcode + "'"
        SetItemWiseTax(clsDBFuncationality.GetDataTable(atchmntqry), txtDocNo.Value)
        Return atchmntqry
    End Function

    Private Sub SendSMSandEmail(ByVal lstUsers As List(Of String), ByVal isSendForApproval As Boolean)
        Try
            'Dim obj As clsEmailSMSSettingNew = clsEmailSMSSettingNew.GetData(clsUserMgtCode.frmSNSaleReturn)

            'If obj Is Nothing Then
            '    clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
            '    Return
            'End If
            'If clsCommon.myLen(obj.mailsubjct) <= 0 Then
            '    clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
            '    Return
            'End If

            'Dim strContactPerson As String = ""
            'Dim strSubject As String = obj.mailsubjct.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'strSubject = strSubject.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))

            'Dim strbody As String = obj.mailbody.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'strbody = strbody.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
            'strbody = strbody.Replace(clsEmailSMSConstants.CustomerNo, txtVendorNo.Value)
            'strbody = strbody.Replace(clsEmailSMSConstants.CustomerName, lblVendorName.Text)
            'strbody = strbody.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
            'strbody = strbody.Replace(clsEmailSMSConstants.Form_Code, MyBase.Form_ID)

            ''------------------------code for attchament-------------------------------------
            'Dim strRptPath As String = ""
            'If obj.atchmnt = "Y" Then
            '    atchmntqry = GetAttachmentQuery(txtDocNo.Value)
            '    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(atchmntqry)
            '    If dt1.Rows.Count > 0 Then
            '        SetItemWiseTax(dt1, txtDocNo.Value)
            '        Dim frmCRV As New frmCrystalReportViewer()
            '        strRptPath = frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt1, "crptShipment", "Shipment Detail")
            '        frmCRV = Nothing
            '    End If
            'End If
            ''---------------------------------------------------------------------------


            'For Each strUser As String In lstUsers
            '    'lstUsers.Add(dr("User_Code").ToString())
            '    Dim lstReceiptents As New List(Of String)
            '    Dim qry As String = ""
            '    Dim emailId As String = ""
            '    If isSendForApproval Then
            '        strContactPerson = strUser
            '        qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" + strUser + "') "
            '        emailId = clsDBFuncationality.getSingleValue(qry)
            '    Else
            '        strContactPerson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
            '        emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
            '    End If

            '    strbody = strbody.Replace(clsEmailSMSConstants.ContactPerson, strContactPerson)
            '    lstReceiptents.Add(emailId)

            '    Dim body As String = strbody.Replace(clsEmailSMSConstants.UserCode, strUser)

            '    clsMailViaOutlook.SendEmail(strSubject, body, lstReceiptents, Nothing, strRptPath)
            'Next
            Dim strContactPerson As String = ""
            Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmSaleReturndairy + "'", Nothing)
            Dim objEmailH As New clsEMailHead()
            objEmailH.arrEMail = New List(Of String)()
            If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then

                If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 Then
                    objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                    objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                    objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))

                    objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Cust_Code, txtVendorNo.Value)
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Cust_Name, lblVendorName.Text)
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.DocAmount, clsCommon.myCstr(lblTotRAmt.Text))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SalePerson_Code, clsCommon.myCstr(txtSalesman.Value))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SalePerson_Name, clsCommon.myCstr(lblSalesman.Text))


                    funPrint(txtDocNo.Value, True)
                    objEmailH.Attachment_1_Path = StrPDFPath


                    For Each strUser As String In lstUsers
                        Dim lstReceiptents As New List(Of String)
                        Dim qry As String = ""
                        Dim emailId As String = ""
                        If isSendForApproval Then
                            strContactPerson = strUser
                            qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" + strUser + "') "
                            emailId = clsDBFuncationality.getSingleValue(qry)
                        Else
                            strContactPerson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
                            emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
                        End If

                        lstReceiptents.Add(emailId)

                        objEmailH.arrEMail.Add(clsCommon.myCstr(emailId))

                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.ContactPerson, strContactPerson)
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.UserCode, strUser)

                    Next


                    objEmailH.SaveData(clsUserMgtCode.frmSaleReturndairy, objEmailH, Nothing)
                    objEmailH = Nothing
                    clsCommon.MyMessageBoxShow(Me, "E-Mail Send Successfully", Me.Text)
                    If Not clsSMSAtPost_Sales.SMSATPOST_SALE() Then
                        SMSSENDONLY(False)
                    End If
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "First do email and sms setting", Me.Text)
                Return
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try


    End Sub
    'Ticket No-TEC/09/08/19-000985
    Private Sub SMSSENDONLY(ByVal isPost As Boolean)
        Try
            'Dim obj As clsEmailSMSSettingNew = clsEmailSMSSettingNew.GetData(clsUserMgtCode.frmSNSaleReturn)

            'If obj Is Nothing Then
            '    clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
            '    Return
            'End If

            'If clsCommon.myLen(obj.smsbody) <= 0 Then
            '    Return
            'End If

            'Dim strMes As String = obj.smsbody
            'strMes = strMes.Replace("'", " ").Replace("`", "/")

            'If strMes.Contains(clsEmailSMSConstants.SaleOrderNo) Then
            '    strMes = strMes.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'End If
            'If strMes.Contains(clsEmailSMSConstants.SaleOrderDate) Then
            '    strMes = strMes.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
            'End If
            'If strMes.Contains(clsEmailSMSConstants.CustomerNo) Then
            '    strMes = strMes.Replace(clsEmailSMSConstants.CustomerNo, txtVendorNo.Value)
            'End If
            'If strMes.Contains(clsEmailSMSConstants.CustomerName) Then
            '    strMes = strMes.Replace(clsEmailSMSConstants.CustomerName, lblVendorName.Text)
            'End If
            'If strMes.Contains(clsEmailSMSConstants.ContactPerson) Then
            '    strMes = strMes.Replace(clsEmailSMSConstants.ContactPerson, "")
            'End If
            'If strMes.Contains(clsEmailSMSConstants.TotalAmount) Then
            '    strMes = strMes.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
            'End If


            'Dim strphone As String = clsDBFuncationality.getSingleValue("select Phone1 from TSPL_customer_MASTER where cust_code ='" & txtVendorNo.Value & "' ")

            'If clsSMSSend.SendSMS(clsUserMgtCode.frmSNSaleReturn, strMes, strphone) Then
            '    If Not isPost Then
            '        clsCommon.MyMessageBoxShow("SMS Send Successfully", Me.Text)
            '    End If
            'End If


            'sanjay
            Dim strContactPerson As String = ""
            Dim strotherno As String = Nothing
            strotherno = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Phone1 from TSPL_CUSTOMER_MASTER where cust_code='" + txtVendorNo.Value + "'"))

            Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmSaleReturndairy + "'", Nothing)

            Dim objSMSH As New clsSMSHead()
            objSMSH.arrMobilNo = New List(Of String)()
            objSMSH.arrMobilNo.Add(clsCommon.myCstr(strotherno))
            If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then

                If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then

                    objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Cust_Code, txtVendorNo.Value)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Cust_Name, lblVendorName.Text)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.DocAmount, clsCommon.myCstr(lblTotRAmt.Text))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SalePerson_Code, clsCommon.myCstr(txtSalesman.Value))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SalePerson_Name, clsCommon.myCstr(lblSalesman.Text))

                    objSMSH.SaveData(clsUserMgtCode.frmSaleReturndairy, objSMSH, Nothing)
                    objSMSH = Nothing
                    If Not isPost Then
                        clsCommon.MyMessageBoxShow(Me, "SMS Send Successfully", Me.Text)
                    End If
                End If
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub SetMailRight()
        'Dim obj As clsCheckMailSetting = clsCheckMailSetting.CheckMailRight()
        If objCommonVar.IsMailSend Then
            btnsend.Enabled = True
        Else
            btnsend.Enabled = False
        End If
    End Sub

    Private Sub btnpreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpreview.Click
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please Select First Sale Return No.", Me.Text)
            txtDocNo.Focus()
            txtDocNo.Select()
            Return
        End If

        atchmntqry = GetAttachmentQuery(txtDocNo.Value)
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(atchmntqry)
        If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            System.Diagnostics.Process.Start(frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt1, "crptSaleReturn", "Sales Return"))
            frmCRV = Nothing
        End If

    End Sub

    Private Sub btnsend_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsend.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select Document No. First", Me.Text)
                txtReqNo.Focus()
                txtReqNo.Select()
                Return
            End If

            If Not (common.clsCommon.MyMessageBoxShow(Me, "Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
                Return
            End If
            LoadData(txtDocNo.Value, NavigatorType.Current)
            Dim lstUsers As New List(Of String)
            lstUsers.Add(txtVendorNo.Value)
            SendSMSandEmail(lstUsers, False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
#End Region
    '--------------------------------------------------------

    Private Sub btnSendForApproval_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendForApproval.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select Document No. First", Me.Text)
                txtReqNo.Focus()
                txtReqNo.Select()
                Return
            End If

            If Not (common.clsCommon.MyMessageBoxShow(Me, "Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
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
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
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

    'Private Sub ddlReturnType_SelectedValueChanged(sender As Object, e As EventArgs) Handles ddlReturnType.SelectedValueChanged
    '    If ddlReturnType.Items.Count > 0 AndAlso ddlReturnType.SelectedIndex >= 0 AndAlso loading = 0 Then
    '        If clsCommon.CompairString(ddlReturnType.SelectedValue, "D") = CompairStringResult.Equal Or clsCommon.CompairString(ddlReturnType.SelectedValue, "I") = CompairStringResult.Equal Then
    '            pnl_damage.Visible = True
    '        Else
    '            pnl_damage.Visible = False
    '        End If
    '    End If
    'End Sub

    Private Sub rbtn_Fresh_CheckStateChanged(sender As Object, e As EventArgs) Handles rbtn_Fresh.CheckStateChanged
        If IsFormLoad = True Then
            AddNew()
        End If
    End Sub
    'Batch wise item return Ticket No- ALF/22/05/18-000066,[ERO/13/12/18-000441]
    Sub OpenBatchItem()
        Dim TransType_Str As String = ""
        If CreateCommonDairyDispatchforFreshAmbient Then
            TransType_Str = "FS-SH"
        Else
            TransType_Str = "FS" 'IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
        End If

        TransType_Str = TransType_Str & "-SR"
        If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsBatchItem).Value) Then
            Dim frm As frmBatchItemOut = New frmBatchItemOut()
            frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
            frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
            frm.strLocationCode = txtBillToLocation.Value
            frm.strCurrDocNo = txtDocNo.Value
            frm.strCurrDocType = TransType_Str
            frm.strUOM = IIf(clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colActualUOM).Value)) > 0, clsCommon.myCstr(gv1.CurrentRow.Cells(colActualUOM).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value))
            'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) done by Panch Raj on 28-06-2017 because product dispatch is not mrp wise (discussed with Ranjana mam and Balwinder sir)
            ''richa BHA/21/08/18-000471
            frm.dblqty = IIf(Math.Round(clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value), 2, MidpointRounding.AwayFromZero) > 0, Math.Round(clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value), 2, MidpointRounding.AwayFromZero), Math.Round(clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), 2, MidpointRounding.AwayFromZero))
            If checkstockmrpwise Then
                frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
            End If
            frm.strSplTransaction = "DSSaleReturn"
            frm.strShipmentNo = clsDBFuncationality.getSingleValue("select Against_Shipment_No from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & txtReqNo.Value & "'")

            frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
            frm.ShowDialog()
            If Not frm.isCencelButtonClicked Then
                gv1.CurrentRow.Cells(colICode).Tag = frm.arr
            End If
        End If
    End Sub
    'Batch wise item return Ticket No- ALF/22/05/18-000066

    Private Sub btnReverseAndUnpost_Click(sender As Object, e As EventArgs) Handles btnReverseAndUnpost.Click
        Try
            If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                If clsDSSalesReturnHead.ReverseAndUnpost(txtDocNo.Value) Then
                    common.clsCommon.MyMessageBoxShow("Successfully Reversed and Recreated", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnInvoiceJE_Click(sender As Object, e As EventArgs) Handles btnInvoiceJE.Click
        clsOpenJEAgainstInvoice.ShowInvoiceJEForReturn(txtDocNo.Value)
    End Sub
    ' Ticket : TEC/29/10/18-000348 By Prabhakar
    Private Sub btnShowInventory_Click(sender As Object, e As EventArgs) Handles btnShowInventory.Click
        clsOpenInventory.ShowInventoryDatails(txtDocNo.Value)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        CancelData()
    End Sub

    Function CancelData() As Boolean
        Try
            Dim isEinvoiceCancelled As Boolean = False
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Code is empty")
            End If
            If clsCommon.MyMessageBoxShow("Are you sure to Cancel the Record?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                Return False
            End If
            If clsCommon.MyMessageBoxShow(Me, "Is E-invoice Cancelled on GST Portal?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                isEinvoiceCancelled = True
            End If

            Dim strReceiptCount As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select receipt_no from TSPL_RECEIPT_DETAIL where Document_No in (Select Document_No from TSPL_Customer_Invoice_Head  where Against_Sale_Return_No='" & txtDocNo.Value & "') "))
            If clsCommon.myLen(strReceiptCount) > 0 Then
                Throw New Exception("You cannot cancelled this document because receiving (" + clsCommon.myCstr(strReceiptCount) + ") has been done against its AR Invoice.")
            End If

            If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal AndAlso Not isEinvoiceCancelled Then
                Dim EInvoiceCancelTimeValid As Int64 = 0
                EInvoiceCancelTimeValid = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,EInvoice_Posting_Date,GETDATE()),0) as PostedHours from TSPL_SD_SALE_RETURN_HEAD where  document_code = '" + txtDocNo.Value + "'"))
                If EInvoiceCancelTimeValid >= 24 Then
                    Throw New Exception("Invoice can not be cancelled.It has been more than 24 hours.")
                End If
            End If

            clsDSSalesReturnHead.CancelData(Me.Form_ID, txtDocNo.Value, NavigatorType.Current, False)
            clsCommon.MyMessageBoxShow(Me, "Successfully Cancelled", Me.Text)
            AddNew()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return True
    End Function

    'Private Sub gvBooth_CellValueChanged(sender As Object, e As GridViewCellEventArgs)
    '    Try
    '        If (Not isInsideLoadData) AndAlso Not isCellValueChangedOpen Then
    '            isCellValueChangedOpen = True
    '            If e.Column Is gvBooth.Columns(colBoothCode) Then
    '                OpenBoothCode(False)
    '            ElseIf e.Column Is gvBooth.Columns(colItem_Code) Then
    '                OpenBoothItem(False)
    '            ElseIf e.Column Is gvBooth.Columns(colUnit_Code) Then
    '                OpenBoothUnit(False)
    '            ElseIf e.Column Is gvBooth.Columns(colBRQty) Then
    '                If clsCommon.myCdbl(gvBooth.CurrentRow.Cells(colBRQty).Value) > 0 Then
    '                    MergeBooth()
    '                Else
    '                    gvBooth.CurrentRow.Cells(colBRQty).Value = 0
    '                    Throw New Exception("Please Fill Qty")
    '                End If
    '            End If
    '            isCellValueChangedOpen = False
    '        End If
    '    Catch ex As Exception
    '        isCellValueChangedOpen = False
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    'Private Sub OpenBoothCode(ByVal isButtonClick As Boolean)
    '    Try
    '        Dim whrCls As String = String.Empty
    '        If clsCommon.myLen(txtRouteNo.Value) > 0 Then
    '            Dim strQry As String = "select Cust_Code as Code,Customer_Name,Route_No,Credit_Customer as [Is Credit Customer] from TSPL_CUSTOMER_MASTER "
    '            whrCls = "  Route_No='" & txtRouteNo.Value & "' and Status='N' "
    '            gvBooth.CurrentRow.Cells(colBoothCode).Value = clsCommon.ShowSelectForm("PSSaleReturnBoothCode", strQry, "Code", whrCls, clsCommon.myCstr(gvBooth.CurrentRow.Cells(colBoothCode).Value), "Code", isButtonClick)
    '            gvBooth.CurrentRow.Cells(colBoothName).Value = clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where cust_code='" & gvBooth.CurrentRow.Cells(colBoothCode).Value & "' ")

    '        Else
    '            gvBooth.CurrentRow.Cells(colBoothCode).Value = ""
    '            Throw New Exception("Please Select Route No.")
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try

    'End Sub
    'Private Sub OpenBoothItem(ByVal isButtonClick As Boolean)
    '    Try
    '        Dim whrCls As String = String.Empty
    '        If clsCommon.myLen(txtRouteNo.Value) > 0 Then
    '            Dim strQry As String = "select distinct TSPL_INVENTORY_MOVEMENT.Item_Code as [Code] ,tspl_item_master.Item_Desc as [Item Description],Sku_Seq as [Seq. No.],ISNULL(Short_Description,'') AS Short_Description,ISNULL(CSA_TYPE,'') AS [Item Group Type] from  TSPL_INVENTORY_MOVEMENT left outer join tspl_item_master on  TSPL_INVENTORY_MOVEMENT.Item_Code=tspl_item_master.item_code "
    '            whrCls = "  inout='I' and Item_Type='F' and Is_Serial_Item=0 and TSPL_ITEM_MASTER.Active=1  "
    '            If rbtn_Fresh.IsChecked Then
    '                whrCls += " and tspl_item_master.Is_FreshItem=1 "
    '            Else
    '                whrCls += " and tspl_item_master.Is_Ambient=1 "
    '            End If
    '            gvBooth.CurrentRow.Cells(colItem_Code).Value = clsCommon.ShowSelectForm("PSSaleReturnItemCode", strQry, "Code", whrCls, clsCommon.myCstr(gvBooth.CurrentRow.Cells(colItem_Code).Value), "Code", isButtonClick)
    '            gvBooth.CurrentRow.Cells(colItem_Name).Value = clsDBFuncationality.getSingleValue("select Short_Description from TSPL_ITEM_MASTER where Item_Code='" & gvBooth.CurrentRow.Cells(colItem_Code).Value & "' ")
    '        Else
    '            gvBooth.CurrentRow.Cells(colItem_Code).Value = ""
    '            Throw New Exception("Please Select Route No.")
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)

    '    End Try

    'End Sub
    'Private Sub OpenBoothUnit(ByVal isButtonClick As Boolean)
    '    Try
    '        Dim strICode As String = clsCommon.myCstr(gvBooth.CurrentRow.Cells(colItem_Code).Value)
    '        If clsCommon.myLen(strICode) > 0 Then
    '            Dim qry As String = "select UOM_Code as Code,UOM_Description as [Description] from TSPL_ITEM_UOM_DETAIL"
    '            Dim whrCls As String = "Item_Code='" & strICode & "'"
    '            gvBooth.CurrentRow.Cells(colUnit_Code).Value = clsCommon.ShowSelectForm("PSSaleReturnActUnit", qry, "Code", whrCls, clsCommon.myCstr(gvBooth.CurrentRow.Cells(colUnit_Code).Value), "Code", isButtonClick)
    '        Else
    '            Throw New Exception("Please Select Route No.")
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)

    '    End Try

    'End Sub
    'Public Sub MergeBooth()
    '    Try
    '        LoadBlankGrid()
    '        LoadBlankGridAC()
    '        LoadBlankGridTax()
    '        Dim myDictionary As New Dictionary(Of String, clsSNShipmentDCSItemDetail)
    '        For ii As Integer = 0 To gvBooth.Rows.Count - 1
    '            If clsCommon.myLen(gvBooth.Rows(ii).Cells(colItem_Code).Value) > 0 AndAlso clsCommon.myLen(gvBooth.Rows(ii).Cells(colUnit_Code).Value) > 0 AndAlso clsCommon.myCDecimal(gvBooth.Rows(ii).Cells(colBRQty).Value) > 0 Then
    '                Dim strKey As String = clsCommon.myCstr(gvBooth.Rows(ii).Cells(colItem_Code).Value) & clsCommon.myCstr(gvBooth.Rows(ii).Cells(colUnit_Code).Value)
    '                If myDictionary.ContainsKey(strKey) Then
    '                    myDictionary(strKey).Qty += clsCommon.myCDecimal(gvBooth.Rows(ii).Cells(colBRQty).Value)
    '                Else
    '                    Dim obj As New clsSNShipmentDCSItemDetail
    '                    obj.ICode = clsCommon.myCstr(gvBooth.Rows(ii).Cells(colItem_Code).Value)
    '                    obj.UOM = clsCommon.myCstr(gvBooth.Rows(ii).Cells(colUnit_Code).Value)
    '                    obj.Qty = clsCommon.myCDecimal(gvBooth.Rows(ii).Cells(colBRQty).Value)

    '                    myDictionary.Add(strKey, obj)

    '                End If

    '            End If
    '        Next
    '        If myDictionary.Count > 0 Then
    '            If gv1.Rows.Count > 0 AndAlso clsCommon.myLen(gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value) <= 0 Then
    '                gv1.Rows.RemoveAt(gv1.Rows.Count - 1)
    '            End If

    '            Dim Lineno As Integer = 1
    '            For Each strKey As String In myDictionary.Keys
    '                gv1.Rows.AddNew()
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = Lineno
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colDelivery_Code).Value = ""
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = "Item"
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colOrderNo).Value = ""
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = myDictionary(strKey).ICode
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsItemMaster.GetItemName(myDictionary(strKey).ICode, Nothing)
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colICodeGrp).Value = clsDBFuncationality.getSingleValue("select CSA_TYPE from TSPL_ITEM_MASTER where Item_Code='" & myDictionary(strKey).ICode & "' ")
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(myDictionary(strKey).ICode, Nothing)
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsEmptyValue).Value = clsItemMaster.IsItemHaveEmptyValue(myDictionary(strKey).ICode, Nothing)
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIsSerialseItem).Value = clsItemMaster.IsSerializeItem(myDictionary(strKey).ICode)
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = myDictionary(strKey).UOM
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colActualUOM).Value = myDictionary(strKey).UOM

    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = myDictionary(strKey).Qty
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colPendingQty).Value = myDictionary(strKey).Qty
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colBalanceQty).Value = myDictionary(strKey).Qty
    '                gv1.Rows(gv1.Rows.Count - 1).Cells(colActualQty).Value = myDictionary(strKey).Qty
    '                'SetTaxX(myDictionary(strKey).ICode, Nothing)
    '                ItemPrice(myDictionary(strKey).ICode, myDictionary(strKey).UOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colActualQty).Value), gv1.CurrentRow.Index, False)
    '                'SetitemWiseTaxSetting(True, False)
    '                SetTaxDetails()

    '                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
    '                    UpdateCurrentRow1(gv1.CurrentRow.Index)

    '                Else
    '                    UpdateCurrentRow(gv1.CurrentRow.Index)

    '                End If
    '                UpdateAllTotals()
    '                Lineno += 1

    '            Next
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Sub
    Private Sub SetTaxX()
        Dim GSTStatus As Boolean = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        Dim trans As SqlTransaction = Nothing
        If GSTStatus = False OrElse (rbtn_Ambient.IsChecked AndAlso GSTStatus) Then
            'If CalculateTaxRatefromItemwsieTaxOnSale Then
            '    If clsCommon.myLen(txtBillToLocation.Value) > 0 Then
            '        Dim strTaxType As String = clsLocationWiseTax.TaxType(txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value, trans)
            '        If GSTStatus = True AndAlso clsCommon.CompairString(strTaxType, "L") = CompairStringResult.Equal Then
            '            ''pick tax from item wise tax master instead of Location master done by richa 25 Nov,2020
            '            ' txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
            '            txtTaxGroup.Value = clsItemWiseTaxAuthority.GetTaxGroupItemWise("L", "S", txtDate.Value, Item_Code, trans)
            '        Else
            '            'txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
            '            txtTaxGroup.Value = clsItemWiseTaxAuthority.GetTaxGroupItemWise("I", "S", txtDate.Value, Item_Code, trans)
            '        End If
            '    Else
            '        txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value, trans)
            '    End If
            'Else

            txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value, trans)
            lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, trans)
            'End If
        Else
            If rbtn_Fresh.IsChecked Then
                txtTaxGroup.Value = clsLocationWiseTax.GetExempedDefaultTaxGroup(True, txtBillToLocation.Value, txtVendorNo.Value, "S", txtDate.Value, trans)
                lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, trans)
            End If
        End If
    End Sub
    Private Sub crateTable()
        Try
            Dim coll As Dictionary(Of String, String)
            coll = New Dictionary(Of String, String)
            coll.Add("Document_Code", "varchar(30) NOT NULL Primary Key")
            coll.Add("Document_Date", "DateTime not NULL")
            coll.Add("Customer_Code", "varchar(12) NOT NULL")
            coll.Add("Status", "integer not null default 0")
            coll.Add("On_Hold", "integer not null default 0")
            coll.Add("Ref_No", "varchar(50) NULL")
            coll.Add("Description", "varchar(200) NULL")
            coll.Add("Remarks", "varchar(200) NULL")
            coll.Add("Tax_Group", "varchar(12) NOT NULL")
            coll.Add("Bill_To_Location", "varchar(12) NULL")
            coll.Add("Ship_To_Location", "varchar(12) NULL")
            coll.Add("TAX1", "varchar(12) NULL")
            coll.Add("TAX1_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX1_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX1_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX2", "varchar(12) NULL")
            coll.Add("TAX2_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX2_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX2_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX3", "varchar(12) NULL")
            coll.Add("TAX3_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX3_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX3_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX4", "varchar(12) NULL")
            coll.Add("TAX4_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX4_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX4_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX5", "varchar(12) NULL")
            coll.Add("TAX5_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX5_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX5_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX6", "varchar(12) NULL")
            coll.Add("TAX6_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX6_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX6_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX7", "varchar(12) NULL")
            coll.Add("TAX7_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX7_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX7_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX8", "varchar(12) NULL")
            coll.Add("TAX8_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX8_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX8_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX9", "varchar(12) NULL")
            coll.Add("TAX9_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX9_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX9_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX10", "varchar(12) NULL")
            coll.Add("TAX10_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX10_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX10_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("Discount_Base", "decimal(18, 6) NULL")
            coll.Add("Discount_Amt", "decimal(18, 6) NULL")
            coll.Add("Amount_Less_Discount", "decimal(18, 6) NULL")
            coll.Add("Total_Tax_Amt", "decimal(18, 6) NULL")
            coll.Add("Total_Amt", "decimal(18, 6) NULL")
            coll.Add("Comments", "varchar(max) NOT NULL")
            coll.Add("Comp_Code", "varchar(8) NOT NULL")
            coll.Add("Terms_Code", "varchar(12) NULL")
            coll.Add("Due_Date", "Datetime NULL")
            coll.Add("Created_By", "varchar(12) NOT NULL")
            coll.Add("Created_Date", "Datetime NOT NULL")
            coll.Add("Modify_By", "varchar(12) NOT NULL")
            coll.Add("Modify_Date", "Datetime NOT NULL")
            coll.Add("Posting_Date", "Datetime NULL")
            coll.Add("Carrier", "Varchar(50) NULL")
            coll.Add("Vehicle_Code", "varchar(12) NULL")
            coll.Add("VehicleNo", "Varchar(50) NULL")
            coll.Add("GRNo", "Varchar(50) NULL")
            coll.Add("GENo", "Varchar(50) NULL")
            coll.Add("GEDate", "Datetime NULL")
            coll.Add("Add_Charge_Code1", "varchar(30) NULL")
            coll.Add("Add_Charge_Name1", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt1", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code2", "varchar(30) NULL")
            coll.Add("Add_Charge_Name2", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt2", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code3", "varchar(30) NULL")
            coll.Add("Add_Charge_Name3", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt3", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code4", "varchar(30) NULL")
            coll.Add("Add_Charge_Name4", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt4", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code5", "varchar(30) NULL")
            coll.Add("Add_Charge_Name5", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt5", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code6", "varchar(30) NULL")
            coll.Add("Add_Charge_Name6", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt6", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code7", "varchar(30) NULL")
            coll.Add("Add_Charge_Name7", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt7", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code8", "varchar(30) NULL")
            coll.Add("Add_Charge_Name8", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt8", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code9", "varchar(30) NULL")
            coll.Add("Add_Charge_Name9", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt9", "decimal(18, 6) NULL")
            coll.Add("Add_Charge_Code10", "varchar(30) NULL")
            coll.Add("Add_Charge_Name10", "varchar(100) NULL")
            coll.Add("Add_Charge_Amt10", "decimal(18, 6) NULL")
            coll.Add("Total_Add_Charge", "decimal(18, 6) NULL")
            coll.Add("Dept", "Varchar(12) null")
            coll.Add("Dept_Desc", "Varchar(50) NULL")
            coll.Add("Item_Type", "char(1) NULL")
            coll.Add("Challan_No", "Varchar(50) NULL")
            coll.Add("Challan_Date", "Date NULL")
            coll.Add("Inv_No", "Varchar(50) NULL")
            coll.Add("Inv_Date", "Date NULL")
            coll.Add("Against_Invoice_No", "Varchar(30) null References TSPL_SD_SHIPMENT_HEAD(Document_Code)")
            coll.Add("Tax_Calculation_Type", "integer not null default 0")
            coll.Add("Is_Internal", "integer not null default 0")
            coll.Add("Salesman_Code", "Varchar(12) null References TSPL_EMPLOYEE_MASTER(EMP_CODE)")
            coll.Add("Salesman_Name", "varchar(30) NULL")
            coll.Add("CURRENCY_CODE", "VARCHAR(30)  NULL REFERENCES TSPL_CURRENCY_MASTER(CURRENCY_CODE) ")

            coll.Add("ConvRate", "Decimal(18,6) NOT NULL DEFAULT 1")
            coll.Add("ApplicableFrom", "DATE NULL")
            coll.Add("PROJECT_ID", "Varchar(30) NULL References TSPL_PJC_PROJECT(PROJECT_CODE)")

            coll.Add("Price_code", "varchar(12) NULL")
            coll.Add("Price_Group_Code", "varchar(20)")
            coll.Add("Route_No", "varchar(12) NULL")
            coll.Add("Route_Desc", "varchar(50) NULL")

            coll.Add("HeadDisc_Per", "decimal(18, 6) NULL")
            coll.Add("HeadDisc_PerAmt", "decimal(18, 6) NULL")
            coll.Add("HeadDisc_Amt", "decimal(18, 6) NULL")
            coll.Add("TotCashDiscAmt", "decimal(18, 6) NULL")
            coll.Add("Invoice_Type", "Char(1) NULL")
            coll.Add("Return_Type", "Char(1) NULL")
            coll.Add("Cust_PO_No", "varchar(200) NULL")
            coll.Add("Approval_Level", "char(6) null")
            coll.Add("Level1_User", "varchar(12) null references tspl_user_master(user_code)")
            coll.Add("Level2_User", "varchar(12) null references tspl_user_master(user_code)")
            coll.Add("Level3_User", "varchar(12) null references tspl_user_master(user_code)")
            coll.Add("Is_Approved", "int not null default 0")

            coll.Add("CrateQty", "float not null default 0")
            coll.Add("Trans_Type", "varchar(3) NOT NULL Default 'ALL'") 'used for define bulk,fresh,csa sale types.

            coll.Add("SO_Validity", "integer not null default 0")
            coll.Add("Commission_Apply", "integer not null default 0")
            coll.Add("Vehicle_Capacity", "float  null default 0")
            coll.Add("Dispatch_Terms", "varchar(20) NULL")
            coll.Add("Payment_Terms", "varchar(20) NULL")
            coll.Add("Dispatch_Period", "integer not null default 0")
            coll.Add("RoundOffAmount", "decimal(18,6) not null default 0")
            coll.Add("CSA_Loc_Code", "varchar(12) NULL")
            '====================notify party detail for Sale Return==================================
            coll.Add("Document_Type", "varchar(2) null") 'ex for export and mt for merchant trade
            coll.Add("Incentive_Declaration", "varchar(max)null")
            coll.Add("Cust_PO_Date", "datetime")
            coll.Add("BANK_CODE", "VARCHAR(12) NULL REFERENCES TSPL_BANK_MASTER(BANK_CODE)")
            coll.Add("BRANCH_CODE", "varchar(30) null")
            coll.Add("Import_Export_No", "Varchar(50) NULL")
            coll.Add("Exporter_Ref_No", "varchar(50) null")
            coll.Add("Pre_Carriage_By", "varchar(100) null")
            coll.Add("Final_Destination", "varchar(100) null")
            coll.Add("Final_Destination_Country", "varchar(30) null")
            coll.Add("Origin_Country", "varchar(30) null")
            coll.Add("Discharge_Port", "varchar(100) null")
            coll.Add("Commission_Paid", "varchar(3) NOT NULL default 'No'")
            coll.Add("Commission_Amount", "float null")
            coll.Add("Comm_Amt_Type", "varchar(1) null")
            coll.Add("Commission_Payee_Code", "varchar(12) null")
            coll.Add("Commission_Instruction", "varchar(1000) null")
            coll.Add("EX_Term_Code", "varchar(3) NULL")
            coll.Add("Advance_Percentage", "float null")
            coll.Add("Loading_Port", "varchar(100) null")
            coll.Add("Vessel_Flight_No", "varchar(100) null")
            coll.Add("CHA_Code", "varchar(12) NULL REFERENCES TSPL_VENDOR_MASTER (Vendor_Code)")
            coll.Add("CHA_Charge_Amount", "float null")
            coll.Add("CHA_FOB_Amount", "float null")
            coll.Add("CHA_Frieght_Kg_Amount", "float null")
            coll.Add("CHA_Basic_Freight_Amount", "float null")
            coll.Add("MT_Against_PO", "Varchar(30) null References TSPL_PURCHASE_ORDER_HEAD(PurchaseOrder_No)")
            coll.Add("MT_Against_PO_Date", "datetime null")
            coll.Add("Bin_No", "varchar(50) NULL")
            coll.Add("CrateMan", "float  null default 0")
            coll.Add("jaali", "float  null default 0")
            coll.Add("Box", "float  null default 0")
            coll.Add("Gate_ReturnNo", "varchar(30) NULL")
            coll.Add("Damage_Type", "Char(1) not NULL default '0'")
            coll.Add("Is_Taxable", "Integer not null default 0")
            coll.Add("Is_Cancelled", "Integer not null default 0")
            coll.Add("Screen_Type", "varchar(10) NOT NULL default('')")
            coll.Add("Gate_Entry_No", "Varchar(30) NULL ")
            coll.Add("Crate_Item", "Varchar(50) NULL References TSPL_ITEM_MASTER(Item_Code)")
            coll.Add("Crate_ItemUnit", "Varchar(12) NULL")
            coll.Add("Crate_ItemRate", "Float not null default 0")
            coll.Add("Can_Item", "Varchar(50) NULL References TSPL_ITEM_MASTER(Item_Code)")
            coll.Add("Can_ItemUnit", "Varchar(12) NULL")
            coll.Add("Can_ItemRate", "Float not null default 0")
            coll.Add("TotalCAN", "float  null")
            coll.Add("ShippedCAN", "float  null")
            coll.Add("LeakageAmount", "decimal(18,6) NULL")
            coll.Add("Sub_Location_code", "varchar(12) NULL")
            coll.Add("EInvoice_Type", "VARCHAR(2) NULL")
            coll.Add("IRN_No", "VARCHAR(100) NULL")
            coll.Add("QR_Code", "VARCHAR(MAX) NULL")
            coll.Add("Ack_No", "VARCHAR(20) NULL")
            coll.Add("Ack_Date", "Datetime NULL")
            coll.Add("BarCode_Img", "image null")
            coll.Add("EInvoice_Posting_Date", "Datetime NULL")
            coll.Add("IsSampling", "Integer not null default 0")
            coll.Add("Distributor_Commission_TotalAmt", "decimal(18,6) null")
            coll.Add("Transporter_Commission_TotalAmt", "decimal(18,6) null")
            coll.Add("Security_TotalAmt", "decimal(18,6) null")
            coll.Add("Is_ManualTCS", "Integer Default 0")
            coll.Add("Gross_Amount", "decimal(18, 6) NULL")
            coll.Add("RateDiff_Amt", "decimal(18, 6) NULL")
            coll.Add("RateDiff_Per", "decimal(18, 6) NULL")
            coll.Add("TotalSubsidyAmt", "Decimal(18,6) NULL")
            coll.Add("TotalSubsidyDisAmt", "Decimal(18,6) NULL")
            coll.Add("Shift_Type", "Varchar(12) NULL")

            clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_SD_SHIPMENT_RETURN_HEAD", coll, Nothing, True, True, "", "Document_Code", "Document_Date", True)

            coll = New Dictionary(Of String, String)
            coll.Add("DOCUMENT_CODE", "Varchar(30) not null References TSPL_SD_SHIPMENT_RETURN_HEAD(DOCUMENT_CODE)")
            coll.Add("Line_No", "integer not null default 0")
            coll.Add("Row_Type", "varchar(10) NULL")
            coll.Add("Item_Code", "varchar(50) NOT NULL")
            coll.Add("Qty", "decimal(18, 6) NULL")
            coll.Add("DamageQty", "decimal(18, 6) NULL")
            coll.Add("Balance_Qty", "decimal(18, 6) NULL")
            coll.Add("Shipment_Code", "Varchar(30) null References TSPL_SD_SHIPMENT_HEAD(DOCUMENT_CODE)")
            coll.Add("Unit_code", "varchar(12) NULL")
            coll.Add("Location", "varchar(12) NOT NULL")
            coll.Add("Item_Cost", "decimal(18, 6) NULL")
            coll.Add("TAX1", "varchar(12) NULL")
            coll.Add("TAX1_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX1_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX1_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX2", "varchar(12) NULL")
            coll.Add("TAX2_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX2_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX2_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX3", "varchar(12) NULL")
            coll.Add("TAX3_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX3_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX3_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX4", "varchar(12) NULL")
            coll.Add("TAX4_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX4_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX4_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX5", "varchar(12) NULL")
            coll.Add("TAX5_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX5_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX5_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX6", "varchar(12) NULL")
            coll.Add("TAX6_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX6_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX6_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX7", "varchar(12) NULL")
            coll.Add("TAX7_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX7_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX7_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX8", "varchar(12) NULL")
            coll.Add("TAX8_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX8_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX8_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX9", "varchar(12) NULL")
            coll.Add("TAX9_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX9_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX9_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX10", "varchar(12) NULL")
            coll.Add("TAX10_Base_Amt", "decimal(18, 6) NULL")
            coll.Add("TAX10_Rate", "decimal(18, 6) NULL")
            coll.Add("TAX10_Amt", "decimal(18, 6) NULL")
            coll.Add("Return_Amount", "decimal(18, 6) NULL")
            coll.Add("Damage_Amount", "decimal(18, 6) NULL")
            coll.Add("Amount", "decimal(18, 6) NULL")
            coll.Add("Disc_Per", "decimal(18, 6) NULL")
            coll.Add("Disc_Amt", "decimal(18, 6) NULL")
            coll.Add("Amt_Less_Discount", "decimal(18, 6) NULL")
            coll.Add("Total_Tax_Amt", "decimal(18, 6) NULL")
            coll.Add("Item_Net_Amt", "decimal(18, 6) NULL")
            coll.Add("Status", "integer not null default 0")
            coll.Add("MRP", "decimal(18, 2) NULL")
            coll.Add("Batch_No", "varchar(30) NULL")
            coll.Add("MFG_Date", "Date null")
            coll.Add("Expiry_Date", "Date null")
            coll.Add("Free_Qty", "decimal(18, 6) NULL")
            coll.Add("Specification", "varchar(200) NULL")
            coll.Add("Remarks", "varchar(200) NULL")
            coll.Add("Assessable", "decimal(18, 6) NULL")
            coll.Add("AssessableAmt", "decimal(18, 6) NULL")
            coll.Add("Is_Mannual_Amt", "integer not null default 0")
            'coll.Add("Unit_Cogs", "decimal(18, 2) NULL")


            coll.Add("Price_code", "varchar(12) NULL")
            coll.Add("Price_Date", "Date NULL")
            coll.Add("Abatement_Per", "decimal(18, 6) NULL")
            coll.Add("Abatement_Amt", "decimal(18, 6) NULL")
            coll.Add("Scheme_Code", "varchar(12) NULL")
            coll.Add("Scheme_Applicable", "char(1) NULL")
            coll.Add("Scheme_Item", "char(1) NULL")
            coll.Add("FOC_Item", "integer not null default 0")
            coll.Add("Item_Tax", "decimal(18, 6) NULL")
            coll.Add("Total_MRP_Amt", "decimal(18, 6) NULL")
            coll.Add("Total_Basic_Amt", "decimal(18, 6) NULL")
            coll.Add("Total_Disc_Amt", "decimal(18, 6) NULL")
            coll.Add("Cust_DiscountQty", "decimal(18, 6) NULL")
            coll.Add("Cust_Discount", "decimal(18, 6) NULL")
            coll.Add("Total_Cust_Discount", "decimal(18, 6) NULL")
            coll.Add("Price_Amount1", "decimal(18, 6) NULL")
            coll.Add("Price_Amount2", "decimal(18, 6) NULL")
            coll.Add("Price_Amount3", "decimal(18, 6) NULL")
            coll.Add("Price_Amount4", "decimal(18, 6) NULL")
            coll.Add("Price_Amount5", "decimal(18, 6) NULL")
            coll.Add("Price_Amount6", "decimal(18, 6) NULL")
            coll.Add("Price_Amount7", "decimal(18, 6) NULL")
            coll.Add("Price_Amount8", "decimal(18, 6) NULL")
            coll.Add("Price_Amount9", "decimal(18, 6) NULL")
            coll.Add("Price_Amount10", "decimal(18, 6) NULL")
            coll.Add("ActualRate", "decimal(18, 6) NULL")
            coll.Add("Item_Weight", "decimal(18, 6) NULL")
            coll.Add("Conv_Factor", "decimal(18, 6) NULL")
            coll.Add("TotalItem_Weight", "decimal(18, 6) NULL")

            coll.Add("Markup_On", "varchar(15) Null")
            coll.Add("Markup_Percent", "Decimal(18,6) NULL")
            coll.Add("Landing_Cost", "Decimal(18,6) NULL")
            coll.Add("HeadDiscAmt", "Decimal(18,6) NULL")
            coll.Add("HeadDiscPer", "Decimal(18,6) NULL")
            coll.Add("HeadDiscPerAmt", "Decimal(18,6) NULL")
            coll.Add("CustDiscPer", "decimal(18, 6) NULL ")
            coll.Add("CasdDiscScheme_Code", "varchar(12) NULL")
            coll.Add("Purchase_Cost", "Decimal(18,6) Null")
            coll.Add("OrgRate", "decimal(18, 6)  not null default 0")
            coll.Add("PrincipleCode", "varchar(30) Null")
            coll.Add("PrincipleDesc", "varchar(200) Null")
            coll.Add("vendor_code", "varchar(12) Null")
            coll.Add("vendor_desc", "varchar(200) Null")
            coll.Add("Bin_No", "varchar(50) NULL")
            coll.Add("Weight_UOM", "varchar(12) NULL ")
            coll.Add("Commission_Rate", "decimal(18, 6) NULL default 0")
            coll.Add("Commission_Party", "varchar(12) NULL")
            coll.Add("Commission_Amt", "decimal(18, 6) NULL default 0")
            coll.Add("Amt_Less_Commission", "decimal(18, 6) NULL default 0")
            coll.Add("OrgUnit_code", "Varchar(12) null")

            coll.Add("TAX_PAID", "varchar(10) NULL ") '' YES , NO
            coll.Add("Alternate_UOM", "varchar(12)  NULL REFERENCES TSPL_UNIT_MASTER(UNIT_CODE)")
            coll.Add("RATE_UOM", "varchar(12)  NULL REFERENCES TSPL_UNIT_MASTER(UNIT_CODE)")


            coll.Add("Scheme_Type", "varchar(20) NULL ") '' YES , NO
            coll.Add("Scheme_Item_Code", "varchar(50) NOT NULL default ''")
            coll.Add("Scheme_Qty", "float null")
            coll.Add("Scheme_Item_UOM", "varchar(12) null")
            coll.Add("Cash_Scheme_Code", "varchar(12) null")
            coll.Add("Cash_Scheme_Type", "char(1) null") 'P-% or A-Amount types
            coll.Add("Cash_Scheme_Pers", "float null")
            coll.Add("Cash_Scheme_Amount", "float null")
            coll.Add("Sampling", "integer not null default 0")
            coll.Add("InvoiceScheme_Code", "varchar(12) NULL")
            coll.Add("InvoiceCashScheme_Code", "varchar(12) NULL")
            coll.Add("Item_Group", "varchar(50)  NULL")
            coll.Add("Transfer_No", "varchar(30) null")
            coll.Add("Org_Transfer_Qty", "float null")
            coll.Add("Org_Transfer_UOM", "varchar(12) null")
            coll.Add("Alt_Transfer_Qty", "float null")
            coll.Add("Rate_UnitQty", "decimal(18, 6) NULL default 0")
            coll.Add("Alter_UnitQty", "decimal(18, 6) NULL default 0")
            coll.Add("Crate", "FLOAT NULL default 0")
            coll.Add("Delivery_Code", "Varchar(30) null ")
            coll.Add("Shipping_Mark", "varchar(100) null")
            coll.Add("Packing_Instruction", "varchar(200) null")
            coll.Add("No_Kind_Package", "varchar(100) null")
            coll.Add("Container_No", "varchar(100) null")
            coll.Add("Adjustment_No", "varchar(30) NULL")
            coll.Add("CSA_SalePatti_Return_No", "varchar(30) null")
            coll.Add("ActualQty", "decimal(18, 6) NULL")
            coll.Add("ActualUOM", "varchar(30) null")
            coll.Add("ActualReturnQty", "decimal(18, 6) NULL")
            coll.Add("ActualConvAmt", "decimal(18, 6) NULL")
            coll.Add("ItemwiseTaxCode", "varchar(30) null ")
            coll.Add("CAN", "float  null default 0")
            coll.Add("ItemLeakageAmount", "decimal(18,6) NULL")
            coll.Add("Distributor_Commission_PKID", "int null References TSPL_DISTRIBUTOR_COMMISSION_DETAIL(PK_ID)")
            coll.Add("Distributor_Commission_Rate", "decimal(18,6) NULL")
            coll.Add("Distributor_Commission_RateWithTax", "decimal(18,6) NULL")
            coll.Add("Distributor_Commission_Amt", "decimal(18,6) NULL")
            coll.Add("Transporter_Commission_Rate", "decimal(18,6) NULL")
            coll.Add("Transporter_Commission_Amt", "decimal(18,6) NULL")
            coll.Add("Security_Rate", "decimal(18,6) NULL")
            coll.Add("Security_Amt", "decimal(18,6) NULL")
            coll.Add("PK_ID", "integer NOT NULL identity NOT FOR REPLICATION primary key")
            coll.Add("Transporter", "varchar(12) NULL")
            coll.Add("price_with_tax", "decimal(18,6) NULL")
            clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_SD_SHIPMENT_RETURN_DETAIL", coll, Nothing, True, True, "TSPL_SD_SALE_RETURN_HEAD", "DOCUMENT_CODE", "", True)

        Catch ex As Exception

        End Try

    End Sub

    'Private Sub chkBoothWiseReturn_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkBoothWiseReturn.ToggleStateChanged
    '    If chkBoothWiseReturn.Checked Then
    '        ApplyBoothWiseReturn = True
    '        RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Visible
    '    Else
    '        ApplyBoothWiseReturn = False
    '        RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
    '    End If
    'End Sub
End Class