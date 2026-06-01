''richa VIJ/22/11/19-000079,VIJ/05/12/19-000097
'Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.Data.SqlClient
Imports System.IO
Imports common
Imports System.Globalization
Imports System.Text.RegularExpressions
''Checkin sanjay 22/06/2020
Public Class frmDairyBookingCustomer
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim DefaultEnableNoTransporter As Boolean = False
    Dim ManualBatchOnCustomerBooking As Boolean = False
    Dim GenerateCustomerWiseGatePass As Boolean = False
    Dim ServerDateTimeForTaxableInvoice As Boolean = False
    Dim CheckCustomeroutStandingAmt As Boolean = False
    Dim DeductTPTFromDocAmt As Boolean = False
    Dim ApplyEWBThresholdLimit As Boolean = False
    Dim EWBThresholdLimitForIntraCity As Integer = 0
    Dim EWBThresholdLimitForIntraState As Integer = 0
    Dim EWBThresholdLimitForInterState As Integer = 0
    Dim isRCDFRateControl As Boolean = False
    Dim OneTimeCheck As Boolean = False
    Dim ApplyDefaultTCSIsChecked As Boolean = False
    Dim ConvertIntoBillingUOM As Boolean = False
    Dim ApplyPricePlanOnDocumentDate As Boolean = False
    Dim DefaultEnableEWayBill As Boolean = False
    Dim HideOutstanding As Boolean = False
    Dim ApplyManualScheme As Boolean = False
    Dim ApplyItemCapacityLimit As Boolean = False
    Dim isloadBookingTypeValues As Boolean = True
    Dim EnableLocation As Boolean = True
    Dim ApplyCommission As Boolean = True
    Dim ApplyTPT As Boolean = True
    Dim ApplyCommissionRateWithTax As Boolean = True
    Dim FORPRICE As Double = 0
    Dim EnableTCSRateValidityFrom01July2021 As Boolean = False
    Dim ShowDemandDoc As Boolean = False
    Dim AmountToCheckCustomerOutstandingForTCSTax As Double = 0
    Dim CalculateTaxRatefromItemwsieTaxOnSale As Integer = 0
    Public checkstockmrpwise As Boolean = False
    Dim RunBatchFifowise As Integer = 0
    Dim RunBatchFifowisewithmodifyfunctionality As Boolean = False
    Private isCellValueChangedTaxOpen As Boolean = False
    Dim ApplyRoundOffZero As Boolean = False
    Dim FlagFirstRecord As Boolean = False
    Dim IsTotalQtyinKG As Boolean = False
    Dim IsExistsForPrice As Boolean = False
    Dim RecordCount As Integer = 0
    Dim FlagCreateDo As Boolean = False
    Dim TotalQuantity As Integer = 0
    Dim TotalCan As Integer = 0
    Dim TotalBox As Integer = 0
    Dim TotalCrate As Integer = 0
    Dim DOmsg As String = ""
    Private DOCreated As Boolean = False
    Dim blnSaveTotalQTy As Boolean = False
    Dim CheckOutstandingOnbooking As Integer = 0
    Dim dblNetAmt As Decimal = 0
    Dim GSTStatus As Boolean = False
    Dim AllowFreshPriceChartonProductSale As Integer = 0
    Dim AllowFreshPriceChartonBookingProductSale As Integer = 0
    Public PrintTruckSheetAfterGenerate As Boolean = False
    Dim vaddnew As String = "Y"
    Dim attachqry As String = ""
    Private StrSql As String
    Public StrDocNo As String
    Public strExcise As Boolean
    Dim blnPageLoad As Boolean = False
    Public Shared valueEntry As Boolean = False
    Dim CreateCommonDairyDispatchforFreshAmbient As Integer = 0
    Private PurchaseOneItemOneVendor As Boolean = False
    Private blnBackCalculation As Boolean = False
    Private IsBatchMFDEXDmandatory As Boolean = False
    Private AutoScheme As Boolean = False
    Private ItemRateEditable As Boolean = False
    Private ItemMRPEditable As Boolean = False
    Const colOrgUnit As String = "COLORGUNIT"
    Public intMRPwithabatement As Integer
    Private intApprovel_Required As Integer = 0
    Public Const RowTypeItem As String = "Item"
    Public Const RowTypeMisc As String = "Misc"
    Private isCellValueChangedOpen As Boolean = False
    Dim iStxtTaxGroup_TxtChangedComplete As Boolean = True
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Private isPageLoadData As Boolean = False
    Dim AllowWo_Outstanding As Boolean
    Dim EnableCustomerPODetailonDairyBooking As Integer = 0
    Dim CheckAvgQtyOnDairyBooking As Boolean = False
    Dim ShowAvailableQtyOnDairyBooking As Boolean = False
    Const colLineNo As String = "COLLNO"
    Const colICode As String = "COLICODE"
    Const colIType As String = "colIType"
    Const colIName As String = "COLINAME"
    Const colIShortName As String = "COLISHORTNAME"
    Const colIHSN As String = "colIHSN"
    Const ColAvailableQty As String = "ColAvailableQty"
    Const colPreviousQty As String = "colPreviousQty"
    Const ColAvgQty As String = "ColAvgQty"
    Const ColBillingQty As String = "ColBillingQty"
    Const colQty As String = "COLQTY"
    Const colUnit As String = "COLUNIT"
    Const colBillingUOM As String = "colBillingUOM"
    Const colRate As String = "COLRATE"
    Const colMRP As String = "colMRP"
    Const colAmt As String = "COLAMT"
    Const colIsBatchItem As String = "colIsBatchItem"
    Const colBatchNo As String = "BATCHNO"
    Const colPriceIDAppDate As String = "colPriceIDAppDate"
    Const colPricePlanNo As String = "colPricePlanNo"
    Const colPriceId As String = "colPriceId"
    Const colDisc_Scheme_Type As String = "colDisc_Scheme_Type"
    Const colDisc_Scheme_Code As String = "colDisc_Scheme_Code"
    Const colDisc_Scheme_Pers As String = "colDisc_Scheme_Pers"
    Const colDisc_Scheme_Amount As String = "colDisc_Scheme_Amount"
    Const colSellingRate As String = "colSellingRate"
    Const colOrgRate As String = "colOrgRate"
    Const colTTaxAutCode As String = "colTTaxAutCode"
    Const colTTaxAutName As String = "colTTaxAutName"
    Const colTTaxRate As String = "colTTaxRate"
    Const colTBaseAmt As String = "TAXBASEAMT"
    Const colTTaxAmt As String = "TAXAMT"
    Const colACCode As String = "colACCode"
    Const colACName As String = "colACName"
    Const colPKID As String = "colPKID"
    Const colACAmount As String = "colACAmount"
    Const colSchemeType As String = "colSchemeType"
    Const colSchemeItem As String = "colSchemeItem"
    Const colTax_NonTax As String = "colTax_NonTax"
    Const colFreshAmbient As String = "colFreshAmbient"
    Const colRemarks As String = "colRemarks"
    Const colItemBasicPrice As String = "colItemBasicPrice"
    Const colAmountWithTax As String = "colAmountWithTax"
    Const colDisAmt As String = "COLDISAMT"
    Const colAmtAfterDis As String = "COLAMTAFTERDIS"
    Const colTaxGroup As String = "colTaxGroup"
    'Const colIsKKF As String = "colIsKKF"
    'Const colIsMNDTax As String = "colIsMNDTax"
    Const colTax As String = "colTax"
    Const colTax_Base_Amt As String = "colTax_Base_Amt"
    Const colTax_Rate As String = "colTax_Rate"
    Const colTax_Amt As String = "colTax_Amt"
    Const colIsTaxOnBaseAmount As String = "colIsTaxOnBaseAmount"
    Const colCF As String = "colCF"
    Const colCFKG As String = "colCFKG"
    Const colQtyinKG As String = "colQtyinKG"
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
    Const ColMainItem As String = "ColMainItem"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim repoBalQty As GridViewDecimalColumn
    Dim repoComplete As GridViewTextBoxColumn
    Dim strVehicleCode As String = ""
    Dim strVehicleDesc As String = ""
    Dim strRoutecode As String = ""
    Dim strRouteDesc As String = ""
    Dim Price_code As String = ""
    Dim dblCustOutstandingAmt As Double = 0
    Dim DOStatus As Integer = 0
    Dim BookingStatus As String = 0
    Dim Is_Cancelled As Integer = 0
    Dim DoNotConsiderCustomerCreditLimit As Integer = 0
    Dim SettTagMultipleRouteWithCustomer As Boolean = False
    Dim SettDairyBookingTolleranceQty As Decimal = 0.0
    Private DonotAllowtoChangeUOMinDairyBookingCustomer As Boolean = False
    Private ShowBookingTypeDropDownonDairyBookingCustomer As Boolean = False
    Dim ShowMulMRPOfSameItemOnDairyBookingCustomer As Boolean = False
    Dim AllowZeroQtyOnDairyBooking As Boolean = False
    Dim DonotIncludeSecurityInCustomerOutstanding As Boolean = False
    Dim IsLoadBookingType As Boolean = False
    Dim EnableItemShortDescriptionInBooking As Boolean = False
    Dim AllowToCreateNoOfBookingPerDay As Integer = 0
    Dim settTCSRateforCustomerWithPanNo As Decimal = 0
    Dim settTCSRateforCustomerWithoutPanNo As Decimal = 0
    Dim ApplyIncludeTCSAmountInRouteTotalOnTruckSheet As Boolean = False
    Dim AutoCalculateCrate As Integer = 0
    Public Property custname As String
    Public Property txtloc As String
    Dim dblOutstandingAmount As Double = 0
    'Dim Tax1_code As String = ""
    'Dim Tax1_Rate As Decimal = 0
    'Dim Tax1_Base_Amt As Decimal = 0
    'Dim Tax1_Amt As Decimal = 0
    'Dim Tax2_code As String = ""
    'Dim Tax2_Rate As Decimal = 0
    'Dim Tax2_Base_Amt As Decimal = 0
    'Dim Tax2_Amt As Decimal = 0
    'Dim Tax3_code As String = ""
    'Dim Tax3_Rate As Decimal = 0
    'Dim Tax3_Base_Amt As Decimal = 0
    'Dim Tax3_Amt As Decimal = 0
    'Dim Tax4_code As String = ""
    'Dim Tax4_Rate As Decimal = 0
    'Dim Tax4_Base_Amt As Decimal = 0
    'Dim Tax4_Amt As Decimal = 0
    'Dim Tax5_code As String = ""
    'Dim Tax5_Rate As Decimal = 0
    'Dim Tax5_Base_Amt As Decimal = 0
    'Dim Tax5_Amt As Decimal = 0
    'Dim Tax6_code As String = ""
    'Dim Tax6_Rate As Decimal = 0
    'Dim Tax6_Base_Amt As Decimal = 0
    'Dim Tax6_Amt As Decimal = 0
#End Region
    '================Update by Preeti Gupta Against Ticket No[ERO/08/05/18-000300]
    '--Sanjay 20201606
    Public Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.frmBookingProductSale)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnCreateAndPrintInvoice.Visible = MyBase.isPrintFlag
        btnPrint.Visible = MyBase.isPrintFlag
        btnGatePassPrint.Visible = MyBase.isPrintFlag
        'btnCreateAndPrintInvoice.Visible = MyBase.isQuickExportFlag
        'If MyBase.isReverse Then
        '    btnreverse.Enabled = True
        'Else
        '    btnreverse.Enabled = False
        'End If
        btnreverse.Visible = False
    End Sub
    Private Sub fndRouteNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtRouteNo._MYValidating
        Dim qry As String = "Select TSPL_ROUTE_MASTER.Route_No as Code,Route_Desc as Description,Type,Employee_Code as 'Employee Code',Off_Day as 'Off Day' from TSPL_ROUTE_MASTER"
        Dim strWhere As String = ""
        If SettTagMultipleRouteWithCustomer AndAlso clsCommon.myLen(txtVendorNo.Value) > 0 Then
            qry += " inner join TSPL_Customer_Route_Master on TSPL_Customer_Route_Master.Route_No=TSPL_ROUTE_MASTER.Route_No "
            strWhere = " TSPL_Customer_Route_Master.cust_Code='" & txtVendorNo.Value & "'"
        End If
        txtRouteNo.Value = clsCommon.ShowSelectForm("DSRouteFinder", qry, "Code", strWhere, txtRouteNo.Value, "", isButtonClicked)
        fndRouteNo_TextChanged()
        lblroutecode.Text = txtRouteNo.Value
        lblroutename.Text = lblRouteDesc.Text
        txtRouteCode1.Text = txtRouteNo.Value
        txtRouteName1.Text = lblRouteDesc.Text
        qry = "select TSPL_VEHICLE_MASTER.Vehicle_Id from TSPL_ROUTE_MASTER left join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where TSPL_ROUTE_MASTER.Route_No='" & clsCommon.myCstr(txtRouteNo.Value) & "'"
        txtVehicleCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Nothing))
        txtVehicleName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Number from TSPL_VEHICLE_MASTER where Vehicle_Id='" & clsCommon.myCstr(txtVehicleCode.Value) & "'"))
        fndTransporter.Value = clsDBFuncationality.getSingleValue("select Transport_Id from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicleCode.Value & "'")
        lblTransporter.Text = clsDBFuncationality.getSingleValue("  select Transporter_Name from TSPL_TRANSPORT_MASTER where Transport_Id='" & fndTransporter.Value & "'")
        If SettTagMultipleRouteWithCustomer Then
            ''richa ERO/15/11/19-001104
            If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code " & Environment.NewLine &
                " from TSPL_CUSTOMER_MASTER " & Environment.NewLine &
                " inner join TSPL_Customer_Route_Master on TSPL_Customer_Route_Master.cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code" & Environment.NewLine &
                " where TSPL_Customer_Route_Master.Route_No='" & txtRouteNo.Value & "' and TSPL_Customer_Route_Master.cust_Code='" & txtVendorNo.Value & "' "
                txtVendorNo.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, Nothing))
                lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
                'setRouteDetail(txtVendorNo.Value, txtRouteNo.Value)
            Else
                txtVendorNo.Value = ""
                lblVendorName.Text = ""
            End If
        End If
    End Sub
    Private Sub fndRouteNo_TextChanged()
        Dim sql As String = "Select Route_Desc,Employee_Code from TSPL_ROUTE_MASTER where Route_No='" & txtRouteNo.Value & "'"
        Dim dr1 As DataTable = clsDBFuncationality.GetDataTable(sql)
        If dr1 IsNot Nothing AndAlso dr1.Rows.Count > 0 Then
            lblRouteDesc.Text = dr1.Rows(0)(0).ToString()
        Else
            lblRouteDesc.Text = String.Empty
        End If
    End Sub

    Private Sub FrmAPInvoiceEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        UcAttachment1.Form_ID = MyBase.Form_ID
        chkisTCS.Checked = True
        ' UcAttachment1.isDeleteTheAttachment = False
        'UcAttachment1.settAutoAttachment = True
        UcAttachment1.MandatoryPDFFileAny = True
        isPageLoadData = True
        isRCDFRateControl = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RCDFRateControl, clsFixedParameterCode.RCDFRateControl, Nothing)) = 1, True, False))
        CreateCommonDairyDispatchforFreshAmbient = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateCommonDairyDispatchforFreshAmbient, clsFixedParameterCode.CreateCommonDairyDispatchforFreshAmbient, Nothing))
        CheckOutstandingOnbooking = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CheckOutstandingCreditLimitOnBooking, clsFixedParameterCode.CheckOutstandingCreditLimitOnBooking, Nothing))
        AllowWo_Outstanding = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AllowDispatchOutstandingFS & "'")) = 0, False, True)
        EnableCustomerPODetailonDairyBooking = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableCustomerPODetailonDairyBooking, clsFixedParameterCode.EnableCustomerPODetailonDairyBooking, Nothing))
        CheckAvgQtyOnDairyBooking = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.DonotCheckAvgQtyOnDairyBooking & "'")) = 0, False, True)
        ShowAvailableQtyOnDairyBooking = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.ShowAvailableQtyOnDairyBooking & "'")) = 0, False, True)
        DoNotConsiderCustomerCreditLimit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DoNotConsiderCustomerCreditLimit, clsFixedParameterCode.DoNotConsiderCustomerCreditLimit, Nothing))
        SettTagMultipleRouteWithCustomer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TagMultipleRouteWithCustomer, clsFixedParameterCode.TagMultipleRouteWithCustomer, Nothing))
        SettDairyBookingTolleranceQty = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DairyBookingTolleranceQty, clsFixedParameterCode.DairyBookingTolleranceQty, Nothing))
        DonotAllowtoChangeUOMinDairyBookingCustomer = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DonotAllowtoChangeUOMinDairyBookingCustomer, clsFixedParameterCode.DonotAllowtoChangeUOMinDairyBookingCustomer, Nothing)) = 1, True, False)
        ShowBookingTypeDropDownonDairyBookingCustomer = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowBookingTypeDropDownonDairyBookingCustomer, clsFixedParameterCode.ShowBookingTypeDropDownonDairyBookingCustomer, Nothing)) = 1, True, False)
        ShowMulMRPOfSameItemOnDairyBookingCustomer = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowMulMRPOfSameItemOnDairyBookingCustomer, clsFixedParameterCode.ShowMulMRPOfSameItemOnDairyBookingCustomer, Nothing)) = 1, True, False)
        AllowZeroQtyOnDairyBooking = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowZeroQtyOnDairyBooking, clsFixedParameterCode.AllowZeroQtyOnDairyBooking, Nothing)) = 1, True, False)
        PrintTruckSheetAfterGenerate = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PrintTruckSheetAfterGenerate, clsFixedParameterCode.PrintTruckSheetAfterGenerate, Nothing)) = 1, True, False)
        DonotIncludeSecurityInCustomerOutstanding = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DonotIncludeSecurityInCustomerOutstanding, clsFixedParameterCode.DonotIncludeSecurityInCustomerOutstanding, Nothing)) = 1, True, False)
        EnableItemShortDescriptionInBooking = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableItemShortDescriptionInBooking, clsFixedParameterCode.EnableItemShortDescriptionInBooking, Nothing)) = 1, True, False)
        AllowToCreateNoOfBookingPerDay = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowToCreateNoOfBookingPerDay, clsFixedParameterCode.AllowToCreateNoOfBookingPerDay, Nothing))
        RunBatchFifowisewithmodifyfunctionality = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.RunBatchFifowisewithModifyfunctionality, clsFixedParameterCode.RunBatchFifowisewithModifyfunctionality, Nothing)) = 1, True, False)
        RunBatchFifowise = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.RunBatchFifowise, clsFixedParameterCode.RunBatchFifowise, Nothing))
        settTCSRateforCustomerWithPanNo = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
        settTCSRateforCustomerWithoutPanNo = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
        'ApplyIncludeTCSAmountInRouteTotalOnTruckSheet = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ApplyIncludeTCSAmountInRouteTotalOnTruckSheet, clsFixedParameterCode.ApplyIncludeTCSAmountInRouteTotalOnTruckSheet, Nothing)) = "1", True, False))
        'Dim ApplyIncludeTCSAmountInRouteTotalOnTruckSheet As Boolean = False
        AutoCalculateCrate = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AutoCalculateCrateOnDairyDispatch & "'")) = 0, 0, 1)
        EnableLocation = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableLocation, clsFixedParameterCode.EnableLocation, Nothing)) = 1, True, False)
        AmountToCheckCustomerOutstandingForTCSTax = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AmountToCheckCustomerOutstandingForTCSTax, clsFixedParameterCode.AmountToCheckCustomerOutstandingForTCSTax, Nothing))
        EnableTCSRateValidityFrom01July2021 = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableTCSRateValidityFrom01July2021, clsFixedParameterCode.EnableTCSRateValidityFrom01July2021, Nothing)) = 0, False, True)
        CalculateTaxRatefromItemwsieTaxOnSale = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, Nothing))
        ApplyRoundOffZero = If(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyRoundOffZero, clsFixedParameterCode.ApplyRoundOffZero, Nothing)) = 1, True, False)
        ShowDemandDoc = If(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowDemandDoc, clsFixedParameterCode.ShowDemandDoc, Nothing)) = 1, True, False)
        FORPRICE = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.FORPRICE, clsFixedParameterCode.FORPRICE, Nothing))
        ApplyCommission = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyCommission, clsFixedParameterCode.ApplyCommission, Nothing)) = 1, True, False)
        ApplyTPT = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyTPT, clsFixedParameterCode.ApplyTPT, Nothing)) = 1, True, False)
        ApplyCommissionRateWithTax = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyCommissionRateWithTax, clsFixedParameterCode.ApplyCommissionRateWithTax, Nothing)) = 1, True, False)
        checkstockmrpwise = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.checkstockMRPwise, clsFixedParameterCode.checkstockMRPwise, Nothing)) = 0, False, True)
        ApplyManualScheme = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyManualScheme, clsFixedParameterCode.ApplyManualScheme, Nothing)) = 0, False, True)
        ApplyDefaultTCSIsChecked = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyDefaultTCSIsChecked, clsFixedParameterCode.ApplyDefaultTCSIsChecked, Nothing)) = 0, False, True)
        ApplyItemCapacityLimit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyItemCapacityLimit, clsFixedParameterCode.ApplyItemCapacityLimit, Nothing)) = 1, True, False)
        ConvertIntoBillingUOM = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ConvertTOBillingUOM, clsFixedParameterCode.ConvertTOBillingUOM, Nothing)) = 1, True, False)
        ApplyPricePlanOnDocumentDate = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyPricePlanOnDocumentDate, clsFixedParameterCode.ApplyPricePlanOnDocumentDate, Nothing)) = 1, True, False)
        DefaultEnableEWayBill = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DefaultEnableEWayBill, clsFixedParameterCode.DefaultEnableEWayBill, Nothing)) = 1, True, False)
        ApplyEWBThresholdLimit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.ApplyEWBThresholdLimit, Nothing)) = 1, True, False)
        EWBThresholdLimitForIntraCity = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraCity, Nothing))
        EWBThresholdLimitForIntraState = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForIntraState, Nothing))
        EWBThresholdLimitForInterState = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyEWBThresholdLimit, clsFixedParameterCode.EWBThresholdLimitForInterState, Nothing))
        DefaultEnableNoTransporter = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DefaultEnableNoTransporter, clsFixedParameterCode.DefaultEnableNoTransporter, Nothing)) = 1, True, False)
        ManualBatchOnCustomerBooking = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ManualBatchOnCustomerBooking, clsFixedParameterCode.ManualBatchOnCustomerBooking, Nothing)) = 1, True, False)
        GenerateCustomerWiseGatePass = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.GenerateCustomerWiseGatePass, clsFixedParameterCode.GenerateCustomerWiseGatePass, Nothing)) = 1, True, False)
        CheckCustomeroutStandingAmt = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CheckCustomeroutStandingAmt, clsFixedParameterCode.CheckCustomeroutStandingAmt, Nothing)) = 1, True, False)
        DeductTPTFromDocAmt = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DeductTPTFromDocAmt, clsFixedParameterCode.DeductTPTFromDocAmt, Nothing)) = 1, True, False)
        ServerDateTimeForTaxableInvoice = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ServerDateTimeForTaxableInvoice, clsFixedParameterCode.ServerDateTimeForTaxableInvoice, Nothing)) = 1, True, False)

        'SetMailRight()
        SetUserMgmtNew()
        'btnCopy.Visible = False
        btnCreateDO.Visible = False
        btnCreateAndPrintInvoice.Visible = True
        If EnableCustomerPODetailonDairyBooking = 1 Then
            'MyLabel2.Visible = True
            txtSalesman.Visible = True
            lblSalesman.Visible = True
            MyLabel12.Visible = True
            txtPONo.Visible = True
            MyLabel4.Visible = True
            txtCustPODate.Visible = True
        Else
            'MyLabel2.Visible = False
            txtSalesman.Visible = False
            lblSalesman.Visible = False
            MyLabel12.Visible = False
            txtPONo.Visible = False
            MyLabel4.Visible = False
            txtCustPODate.Visible = False
        End If
        txtVendorNo.MendatroryField = True
        'LoadModeOfTrasport()
        ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D Delete Trasnaction")
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnAddNew, "Press Alt+N Adding New Trasnaction")
        RadPageView1.SelectedPage = RadPageViewPage1
        LoadBlankGrid()
        LoadBlankGridTax()
        AddNew()
        cmbGatePassType.Text = "Select"
        txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "' "))
        If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                txtSubLocation.Enabled = True
                txtSubLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Sub_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "'"))
                lblSubLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtSubLocation.Value & "'"))
            Else
                txtSubLocation.Enabled = False
                txtSubLocation.Value = ""
                lblSubLocation.Text = ""
            End If

        End If
        lblCredit.Visible = False
        cmbcashcredit.Visible = False
        intMRPwithabatement = clsDBFuncationality.getSingleValue("select IsMRPwithAbatement from TSPL_INV_PARAMETERS")
        blnBackCalculation = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsRateBackCalculation from TSPL_inv_parameters")) = 0, False, True)
        'If ShowBookingTypeDropDownonDairyBookingCustomer = True Then
        '    Me.RadPageViewPage2.Controls.Add(Me.RadGroupBox2)
        '    Me.RadPageViewPage2.ItemSize = New System.Drawing.SizeF(76.0!, 26.0!)
        '    Me.RadPageViewPage2.Location = New System.Drawing.Point(10, 35)
        '    Me.RadPageViewPage2.Name = "RadPageViewPage2"
        '    Me.RadPageViewPage2.Size = New System.Drawing.Size(1024, 409)
        '    Me.RadPageViewPage2.Text = "Item Details"
        '    Me.RadGroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Grouping
        '    Me.RadGroupBox2.Controls.Add(Me.gv1)
        '    Me.RadGroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        '    Me.RadGroupBox2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '    Me.RadGroupBox2.HeaderText = "Item Details"
        '    Me.RadGroupBox2.Location = New System.Drawing.Point(0, 0)
        '    Me.RadGroupBox2.Name = "RadGroupBox2"
        '    Me.RadGroupBox2.Padding = New System.Windows.Forms.Padding(10, 20, 10, 10)
        '    Me.RadGroupBox2.Size = New System.Drawing.Size(1024, 409)
        '    Me.RadGroupBox2.TabIndex = 28
        '    Me.RadGroupBox2.Text = "Item Details"
        'Else
        '    RadPageView1.Pages("RadPageViewPage2").Item.Visibility = ElementVisibility.Collapsed
        'End If
        'RadPageView1.Pages("RadPageViewPage4").Item.Visibility = ElementVisibility.Collapsed
        RadPageView1.Pages("RadPageViewPage2").Item.Visibility = ElementVisibility.Collapsed
        RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
        If clsCommon.myLen(StrDocNo) > 0 Then
            LoadData(StrDocNo, NavigatorType.Current)
        End If
        If CreateCommonDairyDispatchforFreshAmbient = 1 Then
            ItemTypePanel.Visible = False
        Else
            ItemTypePanel.Visible = True
        End If
        'txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "' "))
        'If clsCommon.myLen(txtLocation.Value) > 0 Then
        '    lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
        'End If
        RadMenuItem3.Visibility = ElementVisibility.Collapsed
        isPageLoadData = False
        Try
            If clsCommon.myLen(Me.Tag) > 0 Then
                LoadData(Me.Tag.ToString(), NavigatorType.Current)
            End If
        Catch ex As Exception
        End Try
        If ShowBookingTypeDropDownonDairyBookingCustomer Then
            txtVendorNo.TabIndex = 0
            txtVendorNo.Focus()
        End If
        rbtnTaxable.IsChecked = True
        lblLoginUserZone.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Zone_Code from TSPL_USER_MASTER where User_Code = '" & objCommonVar.CurrentUserCode & "'"))
        pnlTCS.Visible = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowTCSAmountOnBookingForOtherCustomer, clsFixedParameterCode.ShowTCSAmountOnBookingForOtherCustomer, Nothing)) = 1, True, False)
        BPLController(False)
        lblShiftType.Visible = False
        custname = txtVendorNo.Value
        'txtLocation.Value = txtloc
        btnGatepass.Enabled = False
        'CreateTable()
        ChkTaxNonTax()
        btnCreateEWB.Enabled = False
        btnprinte_wayBill.Visible = False
        If ManualBatchOnCustomerBooking Then
            RunBatchFifowise = 0
            RunBatchFifowisewithmodifyfunctionality = False
        End If
    End Sub
    'Sub CreateTable()
    '    Dim coll As Dictionary(Of String, String)
    '    coll = New Dictionary(Of String, String)()
    '    coll.Add("FAT_Per", "decimal(18,2) null")
    '    coll.Add("SNF_Per", "decimal(18,2) null")
    '    coll.Add("Acidity", "decimal(18,2) null")
    '    coll.Add("Temperature", "decimal(18,2) null")
    '    coll.Add("MBRT_Hours", "decimal(18,2) null")
    '    clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_BOOKING_MATSER", coll, Nothing, True, True, "", "Document_No", "Document_Date")
    'End Sub
    Sub BlankAllControls()
        'VendorCodeForChangeIndent = ""
        txtBOstatus.Text = ""
        txtDOStatus.Text = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        lblTotRAmt1.Text = ""
        lblTotalDocAmt.Text = ""
        lblTCSAmount.Text = ""
        txtTCSBaseAmt.Text = ""
        txtDocNo.Value = ""
        lblDONumber.Text = ""
        txtRouteNo.Value = ""
        lblRouteDesc.Text = ""
        txtVendorNo.Value = ""
        lblVendorName.Text = ""
        lblroutecode.Text = ""
        lblroutename.Text = ""
        lblvehiclecode.Text = ""
        lblvehicleName.Text = ""
        'txtLocation.Value = ""
        'lblLocation.Text = ""
        txtPriceCode.Text = ""
        txtPONo.Text = ""
        txtIndentNo.Text = ""
        txtIndentDate.Value = txtDate.Value
        lblCreatedByValue.Text = ""
        txtCustPODate.Checked = False
        txtOpeningbal.Text = ""
        txtDrAmt.Text = ""
        txtCrAmt.Text = ""
        txtClosingBal.Text = ""
        txtCustPODate.Value = clsCommon.GETSERVERDATE()
        UsLock1.Status = ERPTransactionStatus.Pending
        chkManualVehicle.Checked = True
        ChkTaxNonTax()
    End Sub
    Sub LoadBlankGrid()
        'Dim qry As String = String.Empty
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        'Dim TAX_PAID As New GridViewComboBoxColumn
        'Dim BOOK_RATE_UOM As New GridViewTextBoxColumn
        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLineNo)
        'Dim repoisKKF As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        ''repoisKKF.FormatString = ""
        'repoisKKF.HeaderText = " KKF "
        'repoisKKF.Name = colIsKKF
        'repoisKKF.Width = 60
        'repoisKKF.ReadOnly = False
        ''repostatus.ReadOnly = True
        'gv1.MasterTemplate.Columns.Add(repoisKKF)
        'Dim repoisMNDTAX As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        'repoisMNDTAX.Name = colIsMNDTax
        'repoisMNDTAX.Width = 60
        'repoisMNDTAX.HeaderText = "MANDI TAX"
        'repoisMNDTAX.ReadOnly = False
        'gv1.MasterTemplate.Columns.Add(repoisMNDTAX)
        Dim repoICode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoICode.FormatString = ""
        repoICode.HeaderText = "Item Code"
        repoICode.Name = colICode
        repoICode.HeaderImage = My.Resources.search4
        repoICode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoICode.Width = 100
        repoICode.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoICode)
        Dim repoIType As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIType.FormatString = ""
        repoIType.HeaderText = "Item Type"
        repoIType.Name = colIType
        'repoIType.HeaderImage = My.Resources.search4
        'repoIType.TextImageRelation = TextImageRelation.TextBeforeImage
        repoIType.Width = 100
        repoIType.IsVisible = False
        repoIType.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoIType)
        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIName.FormatString = ""
        repoIName.HeaderText = "Item Description"
        repoIName.Name = colIName
        repoIName.Width = 150
        repoIName.ReadOnly = True
        repoIName.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoIName)
        Dim repoIShortName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIShortName.FormatString = ""
        repoIShortName.HeaderText = "Item Short Description"
        repoIShortName.Name = colIShortName
        repoIShortName.Width = 150
        repoIShortName.ReadOnly = IIf(EnableItemShortDescriptionInBooking = True, False, True)
        repoIShortName.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoIShortName)
        'SKG
        'Dim repoIShortName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        'repoIShortName.FormatString = ""
        'repoIShortName.HeaderText = "Item Short Description"
        'repoIShortName.Name = colIShortName
        'repoIShortName.Width = 150
        'repoIShortName.IsVisible = True
        'gv1.MasterTemplate.Columns.Add(repoIShortName)
        'SKG
        Dim repoIHSN As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIHSN.FormatString = ""
        repoIHSN.HeaderText = "HSN Code"
        repoIHSN.Name = colIHSN
        repoIHSN.Width = 150
        repoIHSN.ReadOnly = True
        repoIHSN.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoIHSN)
        'If ApplyManualScheme Then
        Dim repoMainItem As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoMainItem.FormatString = ""
        repoMainItem.HeaderText = "Main Item Code"
        repoMainItem.Name = ColMainItem
        repoMainItem.HeaderImage = My.Resources.search4
        repoMainItem.TextImageRelation = TextImageRelation.TextBeforeImage
        repoMainItem.Width = 100
        repoMainItem.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoMainItem)
        'End If
        Dim repoPriceId As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPriceId.FormatString = ""
        repoPriceId.HeaderText = "Price Id"
        repoPriceId.Name = colPriceId
        repoPriceId.Width = 150
        repoPriceId.ReadOnly = True
        repoPriceId.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoPriceId)
        Dim repoPriceIDAppDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPriceIDAppDate.FormatString = ""
        repoPriceIDAppDate.HeaderText = "Price Id Start Date"
        repoPriceIDAppDate.Name = colPriceIDAppDate
        repoPriceIDAppDate.Width = 150
        repoPriceIDAppDate.ReadOnly = True
        repoPriceIDAppDate.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoPriceIDAppDate)
        Dim repoPricePlanNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoPricePlanNo.FormatString = ""
        repoPricePlanNo.HeaderText = "Price Plan No"
        repoPricePlanNo.Name = colPricePlanNo
        repoPricePlanNo.Width = 150
        repoPricePlanNo.ReadOnly = True
        repoPricePlanNo.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoPricePlanNo)
        Dim repoUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnit.FormatString = ""
        repoUnit.HeaderText = "UOM"
        repoUnit.Name = colUnit
        repoUnit.Width = 80
        gv1.MasterTemplate.Columns.Add(repoUnit)
        Dim repoBillingUnit As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBillingUnit.FormatString = ""
        repoBillingUnit.HeaderText = "Billing UOM"
        repoBillingUnit.Name = colBillingUOM
        repoBillingUnit.Width = 80
        repoBillingUnit.IsVisible = False
        repoBillingUnit.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoBillingUnit)
        Dim repoBillingQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoBillingQty.FormatString = ""
        repoBillingQty.HeaderText = "Billing Qty"
        repoBillingQty.Name = ColBillingQty
        repoBillingQty.Width = 100
        repoBillingQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoBillingQty.ReadOnly = True
        repoBillingQty.IsVisible = False
        repoBillingQty.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoBillingQty)
        Dim repoAvgQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAvgQty.FormatString = ""
        repoAvgQty.HeaderText = "Average Qty"
        repoAvgQty.Name = ColAvgQty
        repoAvgQty.Width = 100
        repoAvgQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAvgQty.ReadOnly = True
        repoAvgQty.IsVisible = False
        repoAvgQty.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoAvgQty)
        Dim repoActualBalQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoActualBalQty.FormatString = ""
        repoActualBalQty.HeaderText = "Available Qty"
        repoActualBalQty.Name = ColAvailableQty
        repoActualBalQty.Width = 100
        repoActualBalQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoActualBalQty.ReadOnly = True
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            repoActualBalQty.IsVisible = True
        Else
            repoActualBalQty.IsVisible = False
        End If
        'repoActualBalQty.IsVisible = False
        repoActualBalQty.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoActualBalQty)
        Dim repopREVIOUSQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repopREVIOUSQty = New GridViewDecimalColumn()
        repopREVIOUSQty.FormatString = ""
        repopREVIOUSQty.HeaderText = "Previous Quantity"
        repopREVIOUSQty.Name = colPreviousQty
        repopREVIOUSQty.Width = 80
        repopREVIOUSQty.Minimum = 0
        repopREVIOUSQty.ReadOnly = True
        repopREVIOUSQty.Step = 0
        repopREVIOUSQty.IsVisible = False
        repopREVIOUSQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repopREVIOUSQty)
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
        Dim repoAmtAfterDis As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterDis = New GridViewDecimalColumn()
        repoAmtAfterDis.FormatString = ""
        repoAmtAfterDis.HeaderText = "Amount After Discount"
        repoAmtAfterDis.Name = colAmtAfterDis
        repoAmtAfterDis.WrapText = True
        repoAmtAfterDis.Width = 80
        repoAmtAfterDis.DecimalPlaces = 2
        repoAmtAfterDis.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterDis.ReadOnly = True
        repoAmtAfterDis.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAmtAfterDis)
        Dim repoQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoQty = New GridViewDecimalColumn()
        repoQty.FormatString = ""
        repoQty.HeaderText = "Quantity"
        repoQty.Name = colQty
        repoQty.Width = 80
        repoQty.Minimum = 0
        repoQty.ShowUpDownButtons = False
        repoQty.Step = 0
        repoQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoQty)
        Dim repoRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoRate = New GridViewDecimalColumn()
        repoRate.FormatString = ""
        repoRate.HeaderText = "Rate"
        repoRate.Name = colRate
        repoRate.Width = 80
        repoRate.Minimum = 0
        If ShowMulMRPOfSameItemOnDairyBookingCustomer = True AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") = CompairStringResult.Equal Then
            repoRate.ReadOnly = False
        Else
            repoRate.ReadOnly = True
        End If
        repoRate.IsVisible = True
        repoRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoRate)
        Dim repoMRP As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMRP = New GridViewDecimalColumn()
        repoMRP.FormatString = ""
        repoMRP.HeaderText = "MRP"
        repoMRP.Name = colMRP
        repoMRP.Width = 80
        repoMRP.Minimum = 0
        repoMRP.ReadOnly = True
        repoMRP.IsVisible = True
        repoMRP.VisibleInColumnChooser = True
        repoMRP.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoMRP)
        Dim repoIsBatchItem As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsBatchItem.HeaderText = "Is Batch Item"
        repoIsBatchItem.Name = colIsBatchItem
        repoIsBatchItem.ReadOnly = True
        repoIsBatchItem.IsVisible = False
        repoIsBatchItem.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(repoIsBatchItem)
        clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)
        Dim repoBatchNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBatchNo.FormatString = ""
        repoBatchNo.HeaderText = "Batch No"
        repoBatchNo.Name = colBatchNo
        repoBatchNo.ReadOnly = False
        repoBatchNo.Width = 100
        gv1.MasterTemplate.Columns.Add(repoBatchNo)
        Dim repoAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmt = New GridViewDecimalColumn()
        repoAmt.FormatString = ""
        repoAmt.HeaderText = "Amount"
        repoAmt.Name = colAmt
        repoAmt.Width = 80
        repoAmt.Minimum = 0
        repoAmt.ReadOnly = True
        repoAmt.IsVisible = True
        repoAmt.VisibleInColumnChooser = True
        repoAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAmt)
        Dim repoTBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTBaseAmt = New GridViewDecimalColumn()
        repoTBaseAmt.FormatString = ""
        repoTBaseAmt.HeaderText = "Tax Base Amt"
        repoTBaseAmt.Name = colTBaseAmt
        repoTBaseAmt.Width = 80
        repoTBaseAmt.Minimum = 0
        repoTBaseAmt.ReadOnly = True
        repoTBaseAmt.IsVisible = False
        repoTBaseAmt.VisibleInColumnChooser = True
        repoTBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTBaseAmt)
        Dim repoTTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTTaxAmt = New GridViewDecimalColumn()
        repoTTaxAmt.FormatString = ""
        repoTTaxAmt.HeaderText = "Tax Amt"
        repoTTaxAmt.Name = colTTaxAmt
        repoTTaxAmt.Width = 80
        repoTTaxAmt.Minimum = 0
        repoTTaxAmt.ReadOnly = True
        repoTTaxAmt.IsVisible = False
        repoTTaxAmt.VisibleInColumnChooser = True
        repoTTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTTaxAmt)
        Dim repoOrgRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoOrgRate = New GridViewDecimalColumn()
        repoOrgRate.FormatString = ""
        repoOrgRate.HeaderText = "Org Basic Rate"
        repoOrgRate.Name = colOrgRate
        repoOrgRate.Width = 80
        repoOrgRate.Minimum = 0
        repoOrgRate.ReadOnly = True
        repoOrgRate.IsVisible = False
        repoOrgRate.VisibleInColumnChooser = True
        repoOrgRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoOrgRate)
        Dim repoSellingRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoSellingRate = New GridViewDecimalColumn()
        repoSellingRate.FormatString = ""
        repoSellingRate.HeaderText = "Selling Rate"
        repoSellingRate.Name = colSellingRate
        repoSellingRate.Width = 80
        repoSellingRate.Minimum = 0
        repoSellingRate.ReadOnly = True
        repoSellingRate.IsVisible = False
        repoSellingRate.VisibleInColumnChooser = False
        repoSellingRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoSellingRate)
        'colItemBasicPrice
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
        'colAmountWithTax
        Dim repoAmountWithTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmountWithTax = New GridViewDecimalColumn()
        repoAmountWithTax.FormatString = ""
        repoAmountWithTax.HeaderText = "Amount with Tax"
        repoAmountWithTax.Name = colAmountWithTax
        repoAmountWithTax.Width = 80
        repoAmountWithTax.Minimum = 0
        repoAmountWithTax.ReadOnly = True
        repoAmountWithTax.IsVisible = False
        repoAmountWithTax.VisibleInColumnChooser = False
        repoAmountWithTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAmountWithTax)
        Dim repoSchemeType As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSchemeType.AllowSort = False
        repoSchemeType.HeaderText = "Scheme Type"
        repoSchemeType.Name = colSchemeType
        repoSchemeType.ReadOnly = True
        repoSchemeType.Width = 96
        repoSchemeType.IsVisible = False
        repoSchemeType.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoSchemeType)
        Dim repoSchemeItem As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSchemeItem.AllowSort = False
        repoSchemeItem.HeaderText = "Scheme Item"
        repoSchemeItem.Name = colSchemeItem
        repoSchemeItem.ReadOnly = True
        repoSchemeItem.Width = 96
        repoSchemeItem.IsVisible = False
        repoSchemeItem.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoSchemeItem)
        Dim repoDisc_Scheme_Type As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDisc_Scheme_Type.HeaderText = "Cash Scheme Type"
        repoDisc_Scheme_Type.Name = colDisc_Scheme_Type
        repoDisc_Scheme_Type.Width = 80
        repoDisc_Scheme_Type.ReadOnly = True
        repoDisc_Scheme_Type.IsVisible = False
        repoDisc_Scheme_Type.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoDisc_Scheme_Type)
        Dim repoDisc_Scheme_Pers As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisc_Scheme_Pers.HeaderText = "Cash Scheme %."
        repoDisc_Scheme_Pers.MinWidth = 4
        repoDisc_Scheme_Pers.Name = colDisc_Scheme_Pers
        repoDisc_Scheme_Pers.ReadOnly = True
        repoDisc_Scheme_Pers.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisc_Scheme_Pers.Width = 54
        repoDisc_Scheme_Pers.IsVisible = False
        repoDisc_Scheme_Pers.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoDisc_Scheme_Pers)
        Dim repoDisc_Scheme_Amount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisc_Scheme_Amount.HeaderText = "Cash Scheme Amt."
        repoDisc_Scheme_Amount.MinWidth = 4
        repoDisc_Scheme_Amount.Name = colDisc_Scheme_Amount
        repoDisc_Scheme_Amount.ReadOnly = True
        repoDisc_Scheme_Amount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisc_Scheme_Amount.Width = 54
        repoDisc_Scheme_Amount.IsVisible = False
        repoDisc_Scheme_Amount.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoDisc_Scheme_Amount)
        Dim repoDisc_Scheme_Code As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDisc_Scheme_Code.HeaderText = "Cash Scheme Code"
        repoDisc_Scheme_Code.Name = colDisc_Scheme_Code
        repoDisc_Scheme_Code.Width = 80
        repoDisc_Scheme_Code.ReadOnly = True
        repoDisc_Scheme_Code.IsVisible = False
        repoDisc_Scheme_Code.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoDisc_Scheme_Code)
        'sanjay
        Dim repoTax_NonTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTax_NonTax.HeaderText = "Tax-NonTax"
        repoTax_NonTax.MinWidth = 4
        repoTax_NonTax.Name = colTax_NonTax
        repoTax_NonTax.ReadOnly = True
        repoTax_NonTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoTax_NonTax.Width = 54
        repoTax_NonTax.IsVisible = False
        repoTax_NonTax.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoTax_NonTax)
        Dim repoFreshAmbient As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFreshAmbient.HeaderText = "Fresh/Ambient"
        repoFreshAmbient.Name = colFreshAmbient
        repoFreshAmbient.Width = 80
        repoFreshAmbient.ReadOnly = True
        repoFreshAmbient.IsVisible = False
        repoFreshAmbient.VisibleInColumnChooser = True
        gv1.MasterTemplate.Columns.Add(repoFreshAmbient)
        Dim repoRemarks As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoRemarks.FormatString = ""
        repoRemarks.HeaderText = "Remarks"
        repoRemarks.Name = colRemarks
        repoRemarks.Width = 200
        gv1.MasterTemplate.Columns.Add(repoRemarks)
        Dim repoTaxGroup As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxGroup.FormatString = ""
        repoTaxGroup.HeaderText = "Tax Group"
        repoTaxGroup.Name = colTaxGroup
        repoTaxGroup.Width = 200
        repoTaxGroup.IsVisible = False
        repoTaxGroup.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoTaxGroup)
        Dim repoTax As GridViewTextBoxColumn
        Dim repoTax_Base_Amt As GridViewDecimalColumn
        Dim repoTax_Rate As GridViewDecimalColumn
        Dim repoTax_Amt As GridViewDecimalColumn
        For i As Integer = 1 To 10
            repoTax = New GridViewTextBoxColumn()
            repoTax.FormatString = ""
            repoTax.HeaderText = "Tax Code" + clsCommon.myCstr(i)
            repoTax.Name = colTax + clsCommon.myCstr(i)
            repoTax.Width = 80
            repoTax.ReadOnly = True
            repoTax.IsVisible = False
            gv1.MasterTemplate.Columns.Add(repoTax)
            repoTax_Base_Amt = New GridViewDecimalColumn()
            repoTax_Base_Amt.FormatString = ""
            repoTax_Base_Amt.HeaderText = "Tax Base Amt" + clsCommon.myCstr(i)
            repoTax_Base_Amt.Name = colTax_Base_Amt + clsCommon.myCstr(i)
            repoTax_Base_Amt.Width = 80
            repoTax_Base_Amt.Minimum = 0
            repoTax_Base_Amt.ReadOnly = True
            repoTax_Base_Amt.IsVisible = False
            repoTax_Base_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gv1.MasterTemplate.Columns.Add(repoTax_Base_Amt)
            'Dim repoTax1_Rate As GridViewDecimalColumn = New GridViewDecimalColumn()
            repoTax_Rate = New GridViewDecimalColumn()
            repoTax_Rate.FormatString = ""
            repoTax_Rate.HeaderText = "Tax Rate" + clsCommon.myCstr(i)
            repoTax_Rate.Name = colTax_Rate + clsCommon.myCstr(i)
            repoTax_Rate.Width = 80
            repoTax_Rate.Minimum = 0
            repoTax_Rate.ReadOnly = True
            repoTax_Rate.IsVisible = False
            repoTax_Rate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gv1.MasterTemplate.Columns.Add(repoTax_Rate)
            repoTax_Amt = New GridViewDecimalColumn()
            repoTax_Amt.FormatString = ""
            repoTax_Amt.HeaderText = "Tax Amt" + clsCommon.myCstr(i)
            repoTax_Amt.Name = colTax_Amt + clsCommon.myCstr(i)
            repoTax_Amt.Width = 80
            repoTax_Amt.Minimum = 0
            repoTax_Amt.ReadOnly = True
            repoTax_Amt.IsVisible = False
            repoTax_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gv1.MasterTemplate.Columns.Add(repoTax_Amt)
        Next
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
        Dim DC_PKID As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_PKID.FormatString = ""
        DC_PKID.HeaderText = "Distributor PKID"
        DC_PKID.Name = ColDCPKID
        DC_PKID.Width = 100
        DC_PKID.ReadOnly = True
        DC_PKID.IsVisible = False
        DC_PKID.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_PKID)
        Dim DC_ApplicableDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_ApplicableDate.FormatString = ""
        DC_ApplicableDate.HeaderText = "Applicable Date"
        DC_ApplicableDate.Name = ColDCApplicableDate
        DC_ApplicableDate.Width = 100
        DC_ApplicableDate.ReadOnly = True
        DC_ApplicableDate.IsVisible = False
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
        DC_Rate.IsVisible = False
        DC_Rate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_Rate)
        Dim DC_RateWithTax As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_RateWithTax.FormatString = ""
        DC_RateWithTax.HeaderText = "Commission Rate Wtih TAx"
        DC_RateWithTax.Name = ColDCRateWithTax
        DC_RateWithTax.Width = 100
        DC_RateWithTax.ReadOnly = True
        DC_RateWithTax.IsVisible = False
        DC_RateWithTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_RateWithTax)
        Dim DC_UnitCF As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_UnitCF.FormatString = ""
        DC_UnitCF.HeaderText = "Conversion Factor"
        DC_UnitCF.Name = ColDCUnitCF
        DC_UnitCF.Width = 100
        DC_UnitCF.ReadOnly = True
        DC_UnitCF.IsVisible = False
        DC_UnitCF.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_UnitCF)
        Dim DC_QtyInSU As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_QtyInSU.FormatString = ""
        DC_QtyInSU.HeaderText = "Qty in Stocking Unit"
        DC_QtyInSU.Name = ColDCQtyinSU
        DC_QtyInSU.Width = 100
        DC_QtyInSU.ReadOnly = True
        DC_QtyInSU.IsVisible = False
        DC_QtyInSU.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_QtyInSU)
        Dim DC_CFUOM As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_CFUOM.FormatString = ""
        DC_CFUOM.HeaderText = "Commission Conversion Factor"
        DC_CFUOM.Name = ColDCCFUOM
        DC_CFUOM.Width = 100
        DC_CFUOM.ReadOnly = True
        DC_CFUOM.IsVisible = False
        DC_CFUOM.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(DC_CFUOM)
        Dim DC_Amt As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        DC_Amt.FormatString = ""
        DC_Amt.HeaderText = "Distributor Commission Amt"
        DC_Amt.Name = ColDCAmt
        DC_Amt.Width = 100
        DC_Amt.ReadOnly = True
        DC_Amt.IsVisible = False
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
        SC_Rate.IsVisible = False
        SC_Rate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(SC_Rate)
        Dim SC_Amt As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        SC_Amt.FormatString = ""
        SC_Amt.HeaderText = "Security Amt"
        SC_Amt.Name = ColSCAmt
        SC_Amt.Width = 100
        SC_Amt.ReadOnly = True
        SC_Amt.IsVisible = False
        SC_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(SC_Amt)
        Dim PKID As GridViewDecimalColumn = New GridViewDecimalColumn()
        PKID = New GridViewDecimalColumn()
        PKID.FormatString = ""
        PKID.HeaderText = "PK ID"
        PKID.Name = colPKID
        PKID.Width = 80
        PKID.Minimum = 0
        PKID.ReadOnly = True
        PKID.IsVisible = False
        PKID.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(PKID)
        ' colItemBasicPrice
        clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)
        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40
        'ReStoreGridLayout()
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
    Function LoadCombobox() As DataTable
        Dim qry As String = "select * from (select 'YES' as Code,'YES' as Name union all select 'NO' as Code,'NO' as Name)a"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        Return dt
    End Function
    Public Sub TCSTaxRatio(ByVal intRow As Integer)
        'Dim tcstaxxbaseamt As Decimal = clsCommon.myCdbl(txttcstaxbaseamount.Value)
        Dim dbltoalbaseamt As Decimal = 0
        Dim tcsbaseamt As Decimal = 0 'gv1.Rows(intRow).Cells(colTax_Base_Amt)
        Dim Actualtcsbaseamt As Decimal = 0
        Dim tcsTaxRate As Decimal = 0
        If clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
            For i As Integer = 0 To gv1.Rows.Count - 1
                For j As Integer = 1 To 10
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(i).Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(j))).Value), "TCS") = CompairStringResult.Equal Then
                        dbltoalbaseamt += clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(j))).Value)
                    End If
                Next

            Next
            If chkisTCS.Checked Then
                For j As Integer = 1 To 10
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(intRow).Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(j))).Value), "TCS") = CompairStringResult.Equal Then
                        tcsbaseamt = clsCommon.myCdbl(gv1.Rows(intRow).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(j))).Value)
                        tcsTaxRate = clsCommon.myCdbl(gv1.Rows(intRow).Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(j))).Value)
                        Actualtcsbaseamt = Math.Round(clsCommon.myCdbl(txttcstaxbaseamount.Value) * (tcsbaseamt / dbltoalbaseamt))
                        gv1.Rows(intRow).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(j))).Value = Actualtcsbaseamt
                        gv1.Rows(intRow).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(j))).Value = (Actualtcsbaseamt * tcsTaxRate) / 100
                    End If
                Next
            End If

        End If
    End Sub
    'Private Sub ReStoreGridLayout()
    '    Try
    '        If clsCommon.myLen(MyBase.Form_ID) > 0 Then
    '            Dim obj As clsGridLayout = New clsGridLayout()
    '            obj = CType(obj.GetData(Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
    '            If Not obj Is Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
    '                Dim ii As Integer
    '                For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
    '                    gv1.Columns(ii).IsVisible = False
    '                    gv1.Columns(ii).VisibleInColumnChooser = True
    '                Next
    '                gv1.LoadLayout(obj.GridLayout)
    '                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
    '            End If
    '        End If
    '    Catch err As Exception
    '        MessageBox.Show(err.Message)
    '    End Try
    'End Sub
    Private Sub gv1_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            'gv1.CurrentRow.Cells(colDisPer).Value
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column.FieldName.StartsWith("_CFLD_") Then
                        clsCustomFieldGrid.getFinderForCustomFieldGrid(gv1, e.Column.Name.ToString, MyBase.Form_ID)
                    End If
                    'If e.Column Is gv1.Columns(colIShortName) OrElse e.Column Is gv1.Columns(colUnit) OrElse e.Column Is gv1.Columns(colICode) OrElse e.Column Is gv1.Columns(colQty) OrElse e.Column Is gv1.Columns(colRate) OrElse e.Column Is gv1.Columns(colIsKKF) OrElse e.Column Is gv1.Columns(colIsMNDTax) Then
                    If e.Column Is gv1.Columns(colIShortName) OrElse e.Column Is gv1.Columns(colUnit) OrElse e.Column Is gv1.Columns(colICode) OrElse e.Column Is gv1.Columns(colQty) OrElse e.Column Is gv1.Columns(colRate) Then
                        If e.Column Is gv1.Columns(colICode) Then
                            If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer First", Me.Text)
                                isCellValueChangedOpen = False
                                Exit Sub
                            End If
                            If clsCommon.myLen(txtLocation.Value) <= 0 Then
                                common.clsCommon.MyMessageBoxShow(Me, "Please select Location First", Me.Text)
                                isCellValueChangedOpen = False
                                Exit Sub
                            End If
                            OpenItemList(False)
                            Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            Dim strIUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            ItemPrice(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index, False)
                            SetTax(strICode, gv1.CurrentRow.Index)
                            rgbItemType.Enabled = False
                        ElseIf e.Column Is gv1.Columns(colIShortName) Then
                            If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer First", Me.Text)
                                isCellValueChangedOpen = False
                                Exit Sub
                            End If
                            If clsCommon.myLen(txtLocation.Value) <= 0 Then
                                common.clsCommon.MyMessageBoxShow(Me, "Please select Location First", Me.Text)
                                isCellValueChangedOpen = False
                                Exit Sub
                            End If
                            gv1.CurrentRow.Cells(colICode).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct Item_Code from tspl_item_master where Short_Description='" & gv1.CurrentRow.Cells(colIShortName).Value & "'"))
                            OpenItemList(False)
                            Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            Dim strIUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            ItemPrice(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index, False)
                            SetTax(strICode, gv1.CurrentRow.Index)
                        ElseIf e.Column Is gv1.Columns(colUnit) Then
                            Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            If clsCommon.myLen(strICode) > 0 Then
                                OpenUOMList(False)
                            Else
                                Throw New Exception("Please fill item first.")
                                isCellValueChangedOpen = False
                                Exit Sub
                            End If
                            Dim strIUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            ItemPrice(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index, False)
                            SetTax(strICode, gv1.CurrentRow.Index)
                            UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
                        ElseIf e.Column Is gv1.Columns(colQty) Then
                            Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            Dim strIUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            If ApplyItemCapacityLimit Then
                                CheckItemCapacityLimit(strICode, strIUOM, clsCommon.myCstr(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index, gv1.CurrentColumn.Index)

                            End If
                            'gv1.CurrentRow.Cells(colOrgCost).Value = gv1.CurrentRow.Cells(colRate).Value
                            SetTax(strICode, gv1.CurrentRow.Index)
                            SetTaxDetails(strICode, gv1.CurrentRow.Index)

                            ItemPrice(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index, False)
                            SetTaxDetails(strICode, gv1.CurrentRow.Index)
                            If ShowBookingTypeDropDownonDairyBookingCustomer = True Then
                                'Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                                'Dim strIUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                                'ItemPrice(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index)
                            End If
                            'If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Is_Batch_Item from TSPL_ITEM_MASTER where Item_Code='" & strICode & "'")) = 1 Then
                            '    OpenBatchItem()
                            'End If
                            If RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                                If gv1.CurrentRow.Cells(colIsBatchItem).Value = True Then
                                    OpenBatchItem()

                                End If

                            End If
                            'If RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                            '        OpenBatchItem()
                            '    End If
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                UpdateCurrentRow1(gv1.CurrentRow.Index)
                            Else
                                UpdateCurrentRow(gv1.CurrentRow.Index)
                            End If
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                IsTotalQtyinKG = False
                                gv1.CurrentRow.Cells(colCF).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) & "' and Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "'")
                                gv1.CurrentRow.Cells(colCFKG).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='KG' and Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "'")
                                If gv1.CurrentRow.Cells(colCFKG).Value IsNot Nothing AndAlso clsCommon.myCdbl(gv1.CurrentRow.Cells(colCFKG).Value) > 0 Then
                                    gv1.CurrentRow.Cells(colQtyinKG).Value = (clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value) * clsCommon.myCdbl(gv1.CurrentRow.Cells(colCF).Value)) / clsCommon.myCdbl(gv1.CurrentRow.Cells(colCFKG).Value)
                                    If clsCommon.myCdbl(gv1.CurrentRow.Cells(colQtyinKG).Value) > FORPRICE Then
                                        ItemPrice(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index, True)
                                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                            UpdateCurrentRow1(gv1.CurrentRow.Index)
                                        Else
                                            UpdateCurrentRow(gv1.CurrentRow.Index)
                                        End If
                                        IsTotalQtyinKG = True
                                    End If
                                End If
                            End If
                            UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
                            TCSTaxRatio(gv1.CurrentRow.Index)

                            UpdateAllTotals()
                        ElseIf e.Column Is gv1.Columns(colRate) Then
                            If ShowMulMRPOfSameItemOnDairyBookingCustomer Then
                                If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)) <= 0 Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Please select Item First", Me.Text)
                                    isCellValueChangedOpen = False
                                    Exit Sub
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)) <= 0 Then
                                    common.clsCommon.MyMessageBoxShow(Me, "Please select Item Uom First", Me.Text)
                                    isCellValueChangedOpen = False
                                    Exit Sub
                                End If
                                Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                                Dim strIUOM As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                                ItemPriceFinderForRateSelection(strICode, strIUOM, clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), gv1.CurrentRow.Index)
                                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                    UpdateCurrentRow1(gv1.CurrentRow.Index)
                                Else
                                    UpdateCurrentRow(gv1.CurrentRow.Index)
                                End If
                                UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
                                UpdateAllTotals()
                            End If
                            'ElseIf e.Column Is gv1.Columns(colIsKKF) Then
                            '    UpdateCurrentRow(gv1.CurrentRow.Index)
                            'ElseIf e.Column Is gv1.Columns(colIsMNDTax) Then
                            '    UpdateCurrentRow(gv1.CurrentRow.Index)
                        End If
                    ElseIf e.Column Is gv1.Columns(ColMainItem) Then
                        OpenSchemeItemList(gv1.CurrentRow.Index)
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                            UpdateCurrentRow1(gv1.CurrentRow.Index)
                        Else
                            UpdateCurrentRow(gv1.CurrentRow.Index)
                        End If
                    End If
                End If
                isCellValueChangedOpen = False
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub OpenSchemeItemList(ByVal IntRowNo As Integer)
        Try
            Dim lstItems As List(Of String) = New List(Of String)
            For i As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.myLen(gv1.Rows(i).Cells(colICode).Value) > 0 AndAlso clsCommon.myLen(gv1.Rows(i).Cells(ColMainItem).Value) = 0 Then
                    If Not lstItems.Contains(clsCommon.myCstr(gv1.Rows(i).Cells(colICode).Value)) Then
                        lstItems.Add(clsCommon.myCstr(gv1.Rows(i).Cells(colICode).Value))
                    End If
                End If
            Next
            If lstItems.Count > 0 Then
                Dim strqry As String = "SELECT item AS Code FROM (VALUES"
                For i As Integer = 0 To lstItems.Count - 2
                    strqry += "('" & lstItems(i) & "'),"
                Next
                strqry += "('" & lstItems(lstItems.Count - 1) & "')) AS Dummytbl(item)"
                gv1.Rows(IntRowNo).Cells(ColMainItem).Value = clsCommon.ShowSelectForm("Dummytbl", strqry, "Code", "", Nothing, "", True)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub RefreshSerialNo()
        Dim intSerialNo As Integer
        For intCount As Integer = 0 To gv1.Rows.Count - 1
            intSerialNo += 1
            gv1.Rows(intCount).Cells(colLineNo).Value = intSerialNo
        Next
    End Sub
    Sub AddCustomerDisc(ByVal isButtonClick As Boolean)
    End Sub
    Sub OpenItemList(ByVal isButtonClick As Boolean)
        'Dim strTax As String = Nothing
        chkSampling.Enabled = False
        GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        Dim whrCls As String = ""
        If CreateCommonDairyDispatchforFreshAmbient = 0 Then
            whrCls = IIf(rbtn_Fresh.IsChecked, " isnull(Is_FreshItem,0)=1 ", " isnull(Is_Ambient,0)=1 ")
        End If
        ' done by priti BHA/14/06/18-000053,ERO/18/06/18-000347
        If clsCommon.myLen(whrCls) > 0 Then
            whrCls += "  and isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0 "
        Else
            whrCls += "  isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0 "
        End If
        If rbtnNonTax.IsChecked Then
            whrCls += " and TSPL_ITEM_MASTER_TAXABLE.Is_Taxable=0"
        ElseIf rbtnTaxable.IsChecked Then
            whrCls += " and TSPL_ITEM_MASTER_TAXABLE.Is_Taxable=1"
            If chkGhee.Checked Then
                whrCls += " and TypeOfItm='G' "
            End If
        End If
        'Sanjay BHA/09/07/18-000142
        whrCls += " and isnull(TSPL_ITEM_MASTER.item_type,'')='F' "
        'Sanjay BHA/09/07/18-000142
        'Ticket No-ALF/05/03/19-000093 show only active item
        whrCls += " and tspl_item_master.Active=1 "
        '"Is_FreshItem=0 and Product_Type not in ('MI')  and Item_Type in ('F','T','S') and Is_Serial_Item=0 " & strTax
        gv1.CurrentRow.Cells(colICode).Value = clsItemMaster.getFinder(whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtDate.Value, False)
        If GSTStatus = False Then
            If CheckItemtaxType() = False Then
                Exit Sub
            End If
        End If
        gv1.CurrentRow.Cells(colIName).Value = clsDBFuncationality.getSingleValue("select Item_Desc from TSPL_ITEM_MASTER where Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        gv1.CurrentRow.Cells(colIType).Value = clsDBFuncationality.getSingleValue("select TypeOfItm from TSPL_ITEM_MASTER where Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        gv1.CurrentRow.Cells(colIShortName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct Short_Description from tspl_item_master where item_code='" & gv1.CurrentRow.Cells(colICode).Value & "'"))
        gv1.CurrentRow.Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(gv1.CurrentRow.Cells(colICode).Value, Nothing)
        gv1.CurrentRow.Cells(colUnit).Value = clsDBFuncationality.getSingleValue("select UOM_Code from TSPL_ITEM_UOM_DETAIL where Default_UOM=1 and Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        If ShowAvailableQtyOnDairyBooking Then
            gv1.CurrentRow.Cells(ColAvailableQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), IIf(clsCommon.myLen(txtSubLocation.Value) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), 0)
        End If
        gv1.CurrentRow.Cells(colTax_NonTax).Value = clsDBFuncationality.getSingleValue("select top 1 Is_Taxable from TSPL_ITEM_MASTER_TAXABLE where Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' and TSPL_ITEM_MASTER_TAXABLE.EFFECTIVE_DATE <= '" & clsCommon.GetPrintDate(txtDate.Value) & "' ORDER BY TSPL_ITEM_MASTER_TAXABLE.EFFECTIVE_DATE DESC ")
        gv1.CurrentRow.Cells(colFreshAmbient).Value = clsDBFuncationality.getSingleValue("select case when Is_Ambient=1 then 'PS' WHEN Is_FreshItem=1 THEN 'FS' ELSE '' END from TSPL_ITEM_MASTER where Item_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' ")
        gv1.CurrentRow.Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value))
        For i As Integer = 0 To gv1.Rows.Count - 1
            If Not chkSampling.Checked AndAlso DonotAllowtoChangeUOMinDairyBookingCustomer Then
                gv1.Rows(i).Cells(colUnit).ReadOnly = True
            Else
                gv1.Rows(i).Cells(colUnit).ReadOnly = False
            End If
        Next
        'SetTax(gv1.CurrentRow.Cells(colICode).Value)
    End Sub
    Public Sub ItemPrice(ByVal strItem As String, ByVal strUnit As String, ByVal intQty As Decimal, ByVal introw As Integer, ByVal isFORPrice As Boolean)
        'Pick Rate
        'Dim Price_code As String = clsDBFuncationality.getSingleValue("select price_CodeNon from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'")
        Dim dt As New DataTable()
        Dim dblRate As Double = 0
        'Dim dblTotal As Double = 0
        Dim dblItemBasicPrice As Double = 0
        Dim whrcls As String = ""
        Dim qry As String = ""
        If chkDCS.Checked Then
            If clsCommon.CompairString(cmbcashcredit.Text, "Cash") = CompairStringResult.Equal Then
                Price_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select VSP_Price_Code_Cash  from tspl_customer_group_master where Cust_Group_Code in(select Cust_Group_Code from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "')"))
                whrcls = " and TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "' and TSPL_ITEM_PRICE_MASTER.Is_For_Price=0"
            Else
                Price_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select VSP_Price_Code_Credit  from tspl_customer_group_master where Cust_Group_Code in(select Cust_Group_Code from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "')"))
                whrcls = " and TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "' and TSPL_ITEM_PRICE_MASTER.Is_For_Price=0"
            End If
        Else
            If isFORPrice Then
                Price_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select price_CodeFOR from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'"))
                whrcls = " and TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "' and TSPL_ITEM_PRICE_MASTER.Is_For_Price=1"
            Else
                Price_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select price_CodeNon from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'"))
                whrcls = " and TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "' and TSPL_ITEM_PRICE_MASTER.Is_For_Price=0"
            End If
        End If
        txtPriceCode.Text = Price_code
        lblPriceCodeDesc.Text = Price_code
        Dim PriceDate As DateTime = txtSupplyDate.Value
        If ApplyPricePlanOnDocumentDate Then
            PriceDate = txtDate.Value
        End If
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
        "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  2=( case when CONVERT(date,Start_Date,103)='" + clsCommon.GetPrintDate(PriceDate) + "' and Shift_Type='" + IIf(clsCommon.CompairString(clsCommon.myCstr(cmbGatePassType.Text), "AM") = CompairStringResult.Equal, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,Start_Date,103)<='" + IIf(clsCommon.CompairString(clsCommon.myCstr(cmbGatePassType.Text), "AM") = CompairStringResult.Equal, clsCommon.GetPrintDate(PriceDate.AddDays(-1)), clsCommon.GetPrintDate(PriceDate)) + "' then 2 else 3 end)  end) and (End_Date >= '" & clsCommon.GetPrintDate(PriceDate, "dd/MMM/yyyy") & "'  or End_date is null) " & whrcls & "  " &
        " and UOM='" & strUnit & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & strItem & "'  AND Location_Code='" & clsCommon.myCstr(txtLocation.Value) & "'  " &
        ") XXXE WHERE RowNo=1  "
        dt = clsDBFuncationality.GetDataTable(qry)
        If dt.Rows.Count > 0 Then
            If chkSampling.Checked Then
                dblRate = clsCommon.myCdbl(0)
                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0).Item("Is_With_Tax")), "N") = CompairStringResult.Equal Then
                    dblItemBasicPrice = Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price")) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Amt")), 2), 2)
                Else
                    dblItemBasicPrice = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price"))
                End If
                'If dblRate = 0 Then
                '    Throw New Exception("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
                '    Exit Sub
                'End If
                'End IfF
            Else
                dblRate = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0).Item("Is_With_Tax")), "N") = CompairStringResult.Equal Then
                    dblItemBasicPrice = Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price")) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Amt")), 2), 2)
                Else
                    dblItemBasicPrice = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price"))
                End If
                If dblRate = 0 Then
                    Throw New Exception("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
                    Exit Sub
                End If
                'End IfF
            End If

            'dblRate = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
            '    If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0).Item("Is_With_Tax")), "N") = CompairStringResult.Equal Then
            '        dblItemBasicPrice = Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price")) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Amt")), 2) + Math.Round(clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Amt")), 2), 2)
            '    Else
            '        dblItemBasicPrice = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price"))
            '    End If
            '    If dblRate = 0 Then
            '        Throw New Exception("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
            '        Exit Sub
            '    End If
            'End IfF
        Else
            If Not isFORPrice Then
                Throw New Exception("Please create Price chart for customer " & txtVendorNo.Value & " for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
                blnSaveTotalQTy = False
                Exit Sub
            End If
        End If
        '''''''''''scheme
        Dim obj_Cash As clsSchemeApplyOnDairy = Nothing
        Dim tax As Double = 0
        Dim tax_on_amt As Decimal = 0
        obj_Cash = clsSchemeApplyOnDairy.GetDiscountSchemeData(strItem, strUnit, intQty, txtVendorNo.Value, Nothing, Nothing)
        If obj_Cash IsNot Nothing Then
            gv1.Rows(introw).Cells(colDisc_Scheme_Amount).Value = obj_Cash.Cash_Amt 'objBookingitem.Disc_Scheme_Amount = obj_Cash.Cash_Amt
            gv1.Rows(introw).Cells(colDisc_Scheme_Pers).Value = obj_Cash.Cash_Pers  'objBookingitem.Disc_Scheme_Pers = obj_Cash.Cash_Pers
            gv1.Rows(introw).Cells(colDisc_Scheme_Code).Value = obj_Cash.Schm_Code        ' objBookingitem.Disc_Scheme_Code = obj_Cash.Schm_Code
            If clsCommon.myCdbl(obj_Cash.Cash_Pers) <> 0 Then
                gv1.Rows(introw).Cells(colDisc_Scheme_Type).Value = "P"
                gv1.Rows(introw).Cells(colDisc_Scheme_Amount).Value = System.Math.Round((dblRate * obj_Cash.Cash_Pers) / 100, 2)
            ElseIf clsCommon.myCdbl(obj_Cash.Cash_Amt) <> 0 Then
                gv1.Rows(introw).Cells(colDisc_Scheme_Type).Value = "A"
            End If
            dblRate = dblRate - gv1.Rows(introw).Cells(colDisc_Scheme_Amount).Value
            tax_on_amt = dblRate
            'Dim Alltax As Double = System.Math.Round((clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX5_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX6_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX7_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX8_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX9_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX10_Rate"))), 3)
            'tax = System.Math.Round((dblRate * Alltax), 3)
            'dblRate = dblRate + tax
        End If
        ''''''''''''scheme
        dt = clsDBFuncationality.GetDataTable(qry)
        If dt.Rows.Count > 0 Then
            If chkSampling.Checked Then
                gv1.Rows(introw).Cells(colSellingRate).Value = clsCommon.myCdbl(0)
                gv1.Rows(introw).Cells(colOrgRate).Value = clsCommon.myCdbl(0)
                gv1.Rows(introw).Cells(colMRP).Value = clsCommon.myCdbl(0)
                gv1.Rows(introw).Cells(colPriceId).Value = clsCommon.myCstr(dt.Rows(0).Item("Item_Price_ID"))
                gv1.Rows(introw).Cells(colPriceIDAppDate).Value = clsCommon.myCstr(dt.Rows(0).Item("Start_date"))
                gv1.Rows(introw).Cells(colPricePlanNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Plan_Code  from TSPL_ITEM_PRICE_PLAN_DETAIL WHERE PLAN_TR_CODE='" & clsCommon.myCstr(dt.Rows(0).Item("Against_Plan_TR_Code")) & "'"))
            Else
                gv1.Rows(introw).Cells(colSellingRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                gv1.Rows(introw).Cells(colOrgRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                gv1.Rows(introw).Cells(colMRP).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                gv1.Rows(introw).Cells(colPriceId).Value = clsCommon.myCstr(dt.Rows(0).Item("Item_Price_ID"))
                gv1.Rows(introw).Cells(colPriceIDAppDate).Value = clsCommon.myCstr(dt.Rows(0).Item("Start_date"))
                gv1.Rows(introw).Cells(colPricePlanNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Plan_Code  from TSPL_ITEM_PRICE_PLAN_DETAIL WHERE PLAN_TR_CODE='" & clsCommon.myCstr(dt.Rows(0).Item("Against_Plan_TR_Code")) & "'"))

            End If
            If ShowMulMRPOfSameItemOnDairyBookingCustomer Then
                isCellValueChangedOpen = True
            End If
            If chkSampling.Checked Then
                gv1.Rows(introw).Cells(colRate).Value = clsCommon.myCdbl(0)
            Else
                gv1.Rows(introw).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
            End If
            'gv1.Rows(introw).Cells(colRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
            If ShowMulMRPOfSameItemOnDairyBookingCustomer Then
                isCellValueChangedOpen = False
            End If
            gv1.Rows(introw).Cells(colTBaseAmt).Value = tax_on_amt
            gv1.Rows(introw).Cells(colTTaxAmt).Value = tax
            gv1.Rows(introw).Cells(colItemBasicPrice).Value = dblItemBasicPrice
            gv1.Rows(introw).Cells(colAmountWithTax).Value = dblItemBasicPrice * intQty
            gv1.Rows(introw).Cells(colTaxGroup).Value = txtTaxGroup.Value
            Dim DOCdateCurrent As Date? = Nothing
            DOCdateCurrent = clsCommon.GETSERVERDATE()
            ' Query to get scheme type of Item
            Dim qryScheme As String = "Select TSPL_SCHEME_MASTER_NEW.Scheme_Type from TSPL_SCHEME_MASTER_NEW left outer join TSPL_SCHEME_DETAIL_NEW on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_MASTER_NEW.Scheme_Code "
            qryScheme += " left outer join tspl_item_master on tspl_item_master.item_code=TSPL_SCHEME_DETAIL_NEW.Item_Code "
            qryScheme += " where TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select scheme_code from (select ROW_NUMBER() over (partition by scheme_type order by start_date desc) as sno, TSPL_SCHEME_MASTER_NEW.Scheme_Code from "
            qryScheme += " TSPL_SCHEME_MASTER_NEW where TSPL_SCHEME_MASTER_NEW.Scheme_Type='Quantitive' and TSPL_SCHEME_MASTER_NEW.Status='Active' and  TSPL_SCHEME_MASTER_NEW.Start_Date<='" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  and (TSPL_SCHEME_MASTER_NEW.End_Date >= '" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  or TSPL_SCHEME_MASTER_NEW.End_date is null) and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select TSPL_SCHEME_DETAIL_NEW.Scheme_Code from "
            qryScheme += " TSPL_SCHEME_DETAIL_NEW left outer join TSPL_SCHEME_BENEFICIARY on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code where MainItem_Code='" & strItem & "' and Cust_Code='" & txtVendorNo.Value & "'))a where a.sno=1)"
            qryScheme += " and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select Scheme_Code from TSPL_SCHEME_BENEFICIARY where Cust_Code='" & txtVendorNo.Value & "') and TSPL_SCHEME_MASTER_NEW.Status='Active'"
            qryScheme += " and TSPL_SCHEME_DETAIL_NEW.MainItem_Code='" & strItem & "' "
            qryScheme += " order by TSPL_SCHEME_MASTER_NEW.Scheme_Code"
            Dim SchemeType As String = clsDBFuncationality.getSingleValue(qryScheme)
            If clsCommon.myLen(clsCommon.myCstr(SchemeType)) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(SchemeType), "Quantitive") = CompairStringResult.Equal Then
                Dim objD As clsSchemeApplyOnDairy = Nothing
                objD = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(strItem), clsCommon.myCstr(strUnit), clsCommon.myCstr(intQty), txtVendorNo.Value, clsCommon.myCstr(SchemeType), Nothing, Nothing, Nothing)
                If objD IsNot Nothing AndAlso objD.Arr.Count > 0 Then
                    For Each objtrScheme As clsSchemeApplyOnDairy In objD.Arr
                        gv1.Rows(introw).Cells(colSchemeType).Value = objtrScheme.schm_Type
                    Next
                End If
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                UpdateCurrentRow1(introw)
            Else
                UpdateCurrentRow(introw)
            End If
            UpdateAllTotals()
        Else
            If Not isFORPrice Then
                clsCommon.MyMessageBoxShow("Please create Price chart for customer " & clsCommon.myCstr(txtVendorNo.Value) & " for Location " & clsCommon.myCstr(txtLocation.Value) & "  for item " & gv1.Rows(introw).Cells(colICode).Value & ".", Me.Text)
                gv1.Rows(introw).Cells(colICode).Value = Nothing
                gv1.Rows(introw).Cells(colUnit).Value = Nothing
                gv1.Rows(introw).Cells(colIName).Value = Nothing
                gv1.Rows(introw).Cells(colIHSN).Value = Nothing
                gv1.Rows(introw).Cells(colRate).Value = 0
                gv1.Rows(introw).Cells(colOrgRate).Value = 0
                gv1.Rows(introw).Cells(colTBaseAmt).Value = 0
                gv1.Rows(introw).Cells(colTTaxAmt).Value = 0
                Exit Sub
            End If
        End If
        'Return dblRate
    End Sub
    Public Sub GetDCDetails(ByVal intRow As Integer)
        Dim DCQry As String = "select top 1 TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Commision_UOM,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.PK_ID,TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Rate,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Transporter_Rate,TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Security_Rate from TSPL_DISTRIBUTOR_COMMISSION_HEAD
left join TSPL_DISTRIBUTOR_COMMISSION_DETAIL on TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
left join TSPL_DISTRIBUTOR_COMMISSION_ITEMS on TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Doc_No=TSPL_DISTRIBUTOR_COMMISSION_HEAD.Doc_No
where TSPL_DISTRIBUTOR_COMMISSION_HEAD.Applicable_Date<='" & clsCommon.GetPrintDate(txtDate.Value) & "' and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' and TSPL_DISTRIBUTOR_COMMISSION_ITEMS.Item_Code='" & clsCommon.myCstr(gv1.Rows(intRow).Cells(colICode).Value) & "' and TSPL_DISTRIBUTOR_COMMISSION_HEAD.IsPosted=1 and TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code='" & clsCommon.myCstr(txtRouteNo.Value) & "'
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
            gv1.Rows(intRow).Cells(ColDCUnitCF).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" & clsCommon.myCstr(gv1.Rows(intRow).Cells(colUnit).Value) & "' and Item_Code='" & clsCommon.myCstr(gv1.Rows(intRow).Cells(colICode).Value) & "'")
            gv1.Rows(intRow).Cells(ColDCCFUOM).Value = clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where UOM_Code='" & clsCommon.myCstr(gv1.Rows(intRow).Cells(ColDCUOM).Value) & "' and Item_Code='" & clsCommon.myCstr(gv1.Rows(intRow).Cells(colICode).Value) & "'")
            'gv1.CurrentRow.Cells(ColDCQtyinSU).Value = (gv1.Rows(IntRowNo).Cells(colQty).Value * gv1.Rows(IntRowNo).Cells(ColDCUnitCF).Value) / gv1.Rows(IntRowNo).Cells(ColDCCFUOM).Value
            'gv1.CurrentRow.Cells(ColDCAmt).Value = gv1.CurrentRow.Cells(ColDCQtyinSU).Value * gv1.CurrentRow.Cells(ColDCRateWithTax).Value
            'dblTotalDCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCAmt).Value)
        End If
    End Sub
    Private Sub ItemPriceFinderForRateSelection(ByVal strItem As String, ByVal strUnit As String, ByVal intQty As Decimal, ByVal introw As Integer)
        Dim dt As New DataTable()
        Dim dblRate As Double = 0
        'Dim dblTotal As Double = 0
        Dim dblItemBasicPrice As Double = 0
        Dim qry As String = ""
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
        "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null)  and  " &
        "TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "' and UOM='" & strUnit & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & strItem & "' AND Location_Code='" & clsCommon.myCstr(txtLocation.Value) & "'  " &
        ") XXXE "
        Dim dr As DataRow = Nothing
        dr = clsCommon.ShowSelectFormForRow("ItemrateFinder", qry)
        If dr IsNot Nothing Then
            dblRate = clsCommon.myCdbl(dr("Item_Selling_Price"))
            If clsCommon.CompairString(clsCommon.myCstr(dr("Is_With_Tax")), "N") = CompairStringResult.Equal Then
                dblItemBasicPrice = Math.Round(clsCommon.myCdbl(dr("Item_Basic_Price")) + Math.Round(clsCommon.myCdbl(dr("TAX1_Amt")), 2) + Math.Round(clsCommon.myCdbl(dr("TAX2_Amt")), 2) + Math.Round(clsCommon.myCdbl(dr("TAX3_Amt")), 2) + Math.Round(clsCommon.myCdbl(dr("TAX4_Amt")), 2), 2)
            Else
                dblItemBasicPrice = clsCommon.myCdbl(dr("Item_Basic_Price"))
            End If
            If dblRate = 0 Then
                Throw New Exception("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
                Exit Sub
            End If
        Else
            Throw New Exception("Please create Price chart for customer " & txtVendorNo.Value & " for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
            blnSaveTotalQTy = False
            Exit Sub
        End If
        '''''''''''scheme
        Dim obj_Cash As clsSchemeApplyOnDairy = Nothing
        Dim tax As Double = 0
        Dim tax_on_amt As Decimal = 0
        obj_Cash = clsSchemeApplyOnDairy.GetDiscountSchemeData(strItem, strUnit, intQty, txtVendorNo.Value, Nothing, Nothing)
        If obj_Cash IsNot Nothing Then
            gv1.CurrentRow.Cells(colDisc_Scheme_Amount).Value = obj_Cash.Cash_Amt 'objBookingitem.Disc_Scheme_Amount = obj_Cash.Cash_Amt
            gv1.CurrentRow.Cells(colDisc_Scheme_Pers).Value = obj_Cash.Cash_Pers  'objBookingitem.Disc_Scheme_Pers = obj_Cash.Cash_Pers
            gv1.CurrentRow.Cells(colDisc_Scheme_Code).Value = obj_Cash.Schm_Code        ' objBookingitem.Disc_Scheme_Code = obj_Cash.Schm_Code
            If clsCommon.myCdbl(obj_Cash.Cash_Pers) <> 0 Then
                gv1.CurrentRow.Cells(colDisc_Scheme_Type).Value = "P"
                gv1.CurrentRow.Cells(colDisc_Scheme_Amount).Value = System.Math.Round((dblRate * obj_Cash.Cash_Pers) / 100, 2)
            ElseIf clsCommon.myCdbl(obj_Cash.Cash_Amt) <> 0 Then
                gv1.CurrentRow.Cells(colDisc_Scheme_Type).Value = "A"
            End If
            dblRate = dblRate - gv1.CurrentRow.Cells(colDisc_Scheme_Amount).Value
            tax_on_amt = dblRate
        End If
        ''''''''''''scheme
        qry += " where RowNo='" & clsCommon.myCstr(dr("RowNo")) & "'"
        dt = clsDBFuncationality.GetDataTable(qry)
        If dt.Rows.Count > 0 Then
            gv1.CurrentRow.Cells(colSellingRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
            gv1.CurrentRow.Cells(colOrgRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
            gv1.CurrentRow.Cells(colPriceId).Value = clsCommon.myCstr(dt.Rows(0).Item("Item_Price_ID"))
            gv1.CurrentRow.Cells(colPriceIDAppDate).Value = clsCommon.myCstr(dt.Rows(0).Item("Start_date"))
            gv1.CurrentRow.Cells(colPricePlanNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Plan_Code  from TSPL_ITEM_PRICE_PLAN_DETAIL WHERE PLAN_TR_CODE='" & clsCommon.myCstr(dt.Rows(0).Item("Against_Plan_TR_Code")) & "'"))
            isCellValueChangedOpen = True
            gv1.CurrentRow.Cells(colRate).Value = dblRate
            isCellValueChangedOpen = False
            gv1.CurrentRow.Cells(colTBaseAmt).Value = tax_on_amt
            gv1.CurrentRow.Cells(colTTaxAmt).Value = tax
            gv1.CurrentRow.Cells(colItemBasicPrice).Value = dblItemBasicPrice
            gv1.CurrentRow.Cells(colAmountWithTax).Value = dblItemBasicPrice * intQty
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                UpdateCurrentRow1(gv1.CurrentRow.Index)
            Else
                UpdateCurrentRow(gv1.CurrentRow.Index)
            End If
            Dim DOCdateCurrent As Date? = Nothing
            DOCdateCurrent = clsCommon.GETSERVERDATE()
            ' Query to get scheme type of Item
            Dim qryScheme As String = "Select TSPL_SCHEME_MASTER_NEW.Scheme_Type from TSPL_SCHEME_MASTER_NEW left outer join TSPL_SCHEME_DETAIL_NEW on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_MASTER_NEW.Scheme_Code "
            qryScheme += " left outer join tspl_item_master on tspl_item_master.item_code=TSPL_SCHEME_DETAIL_NEW.Item_Code "
            qryScheme += " where TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select scheme_code from (select ROW_NUMBER() over (partition by scheme_type order by start_date desc) as sno, TSPL_SCHEME_MASTER_NEW.Scheme_Code from "
            qryScheme += " TSPL_SCHEME_MASTER_NEW where TSPL_SCHEME_MASTER_NEW.Scheme_Type='Quantitive' and TSPL_SCHEME_MASTER_NEW.Status='Active' and  TSPL_SCHEME_MASTER_NEW.Start_Date<='" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  and (TSPL_SCHEME_MASTER_NEW.End_Date >= '" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  or TSPL_SCHEME_MASTER_NEW.End_date is null) and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select TSPL_SCHEME_DETAIL_NEW.Scheme_Code from "
            qryScheme += " TSPL_SCHEME_DETAIL_NEW left outer join TSPL_SCHEME_BENEFICIARY on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code where MainItem_Code='" & strItem & "' and Cust_Code='" & txtVendorNo.Value & "'))a where a.sno=1)"
            qryScheme += " and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select Scheme_Code from TSPL_SCHEME_BENEFICIARY where Cust_Code='" & txtVendorNo.Value & "') and TSPL_SCHEME_MASTER_NEW.Status='Active'"
            qryScheme += " and TSPL_SCHEME_DETAIL_NEW.MainItem_Code='" & strItem & "' "
            qryScheme += " order by TSPL_SCHEME_MASTER_NEW.Scheme_Code"
            Dim SchemeType As String = clsDBFuncationality.getSingleValue(qryScheme)
            If clsCommon.myLen(clsCommon.myCstr(SchemeType)) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(SchemeType), "Quantitive") = CompairStringResult.Equal Then
                Dim objD As clsSchemeApplyOnDairy = Nothing
                objD = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(strItem), clsCommon.myCstr(strUnit), clsCommon.myCstr(intQty), txtVendorNo.Value, clsCommon.myCstr(SchemeType), Nothing, Nothing, Nothing)
                If objD IsNot Nothing AndAlso objD.Arr.Count > 0 Then
                    For Each objtrScheme As clsSchemeApplyOnDairy In objD.Arr
                        gv1.CurrentRow.Cells(colSchemeType).Value = objtrScheme.schm_Type
                    Next
                End If
            End If
        Else
            clsCommon.MyMessageBoxShow(Me, "Please create Price chart for customer " & clsCommon.myCstr(txtVendorNo.Value) & " for Location " & clsCommon.myCstr(txtLocation.Value) & "  for item " & gv1.Rows(introw).Cells(colICode).Value & ".", Me.Text)
            gv1.CurrentRow.Cells(colICode).Value = Nothing
            gv1.CurrentRow.Cells(colUnit).Value = Nothing
            gv1.CurrentRow.Cells(colIName).Value = Nothing
            gv1.CurrentRow.Cells(colIHSN).Value = Nothing
            gv1.CurrentRow.Cells(colRate).Value = 0
            gv1.CurrentRow.Cells(colOrgRate).Value = 0
            gv1.CurrentRow.Cells(colTBaseAmt).Value = 0
            gv1.CurrentRow.Cells(colTTaxAmt).Value = 0
            Exit Sub
        End If
    End Sub
    Sub OpenItemUOMList(ByVal isButtonClick As Boolean)
        Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        Dim qry As String = "select UOM_Code as Code,UOM_Description as [Description] from TSPL_ITEM_UOM_DETAIL"
        Dim whrCls As String = "Item_Code='" & strICode & "'"
        gv1.CurrentRow.Cells(colUnit).Value = clsCommon.ShowSelectForm("PS-BOItemUOMFndr", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), "Code", isButtonClick)
    End Sub
    Sub OpenUOMList(ByVal isButtonClick As Boolean)
        Dim strICode As String = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
        If clsCommon.myLen(strICode) > 0 Then
            Dim qry As String = "select UOM_Code as Code,UOM_Description as [Description] from TSPL_ITEM_UOM_DETAIL"
            Dim whrCls As String = "Item_Code='" & strICode & "'"
            gv1.CurrentRow.Cells(colUnit).Value = clsCommon.ShowSelectForm("PS-BOUOMFndr", qry, "Code", whrCls, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), "Code", isButtonClick)
            If ShowAvailableQtyOnDairyBooking Then
                gv1.CurrentRow.Cells(ColAvailableQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), IIf(clsCommon.myLen(txtSubLocation.Value) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), 0)
            End If
            'ItemPrice(gv1.CurrentRow.Cells(colICode).Value, gv1.CurrentRow.Cells(colUnit).Value, gv1.CurrentRow.Cells(colQty).Value, gv1.CurrentRow.Index)
            'End If
            'End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                UpdateCurrentRow1(gv1.CurrentRow.Index)
            Else
                UpdateCurrentRow(gv1.CurrentRow.Index)
            End If ''-1 is for current row
            UpdateAllTotals()
        End If
    End Sub
    Private Function GetConvFactor(ByVal strUnit As String, ByVal strItem As String) As Double
        Dim dblConvF As Double = 0
        dblConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strUnit & "'"))
        Return dblConvF
    End Function
    Private Sub SetBlankOfItemColumns()
        gv1.CurrentRow.Cells(colICode).Value = ""
        gv1.CurrentRow.Cells(colIName).Value = ""
        gv1.CurrentRow.Cells(colIHSN).Value = ""
        gv1.CurrentRow.Cells(colUnit).Value = ""
        gv1.CurrentRow.Cells(colOrgUnit).Value = ""
        'gv1.CurrentRow.Cells(colMRP).Value = 0
        'gv1.CurrentRow.Cells(colPriceDateColumn).Value = 0
        'gv1.CurrentRow.Cells(colPriceCOde).Value = 0
        'gv1.CurrentRow.Cells(colSchemeApplicable).Value = "No"
        'gv1.CurrentRow.Cells(colSchemeItem).Value = "No"
        'gv1.CurrentRow.Cells(colRate).Value = 0
        'gv1.CurrentRow.Cells(colMRP).Value = 0
        ''gv1.CurrentRow.Cells(colAssessableRate).Value = 0
    End Sub
    Private Sub gv1_UserAddedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserAddedRow
        For i As Integer = 0 To gv1.Rows.Count - 1
            gv1.Rows(0).Cells(0).Value = 1
            If i <> 0 Then
                gv1.Rows(i).Cells(colLineNo).Value = i + 1
            End If
        Next
    End Sub
    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer)
        Try
            'SetTax(gv1.Rows(IntRowNo).Cells(colICode).Value)
            Dim arrTaxableAuth As New List(Of String)
            Dim arrTaxableAuth1 As New List(Of String)
            'Dim dblFAmt As Double = 0
            GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
            If Not GSTStatus Then
                strExcise = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Excisable from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'")) = "T", True, False)
            End If
            'Dim dblAlterQty As Double = 0
            Dim strICode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
            Dim strUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblQty As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value)
            Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colRate).Value)
            If ApplyManualScheme AndAlso clsCommon.myLen(gv1.Rows(IntRowNo).Cells(ColMainItem).Value) > 0 Then
                'If clsCommon.myLen(gv1.Rows(IntRowNo).Cells(ColMainItem).Value) > 0 Then
                dblRate = 0
                'End If
            End If
            'Dim dblMainItem As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColMainItem).Value)
            Dim wt_unit As String = 0
            'Dim wt_qty As Double = 0
            'Dim Item_Weight As Double = 0
            Dim TotalItem_Weight As Double = 0
            'Dim TotalItem_WeightMetric As Double = 0
            If clsCommon.myLen(strICode) > 0 Then
                wt_unit = clsItemMaster.GetItemWeightUnit(strICode, Nothing)
                TotalItem_Weight = clsItemMaster.getTotalItemWeight(strICode, strUnit, dblQty, Nothing)
            End If
            'Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            'Dim dblBasicAmt As Double = dblQty * dblRate
            Dim dblAmt As Double = (dblQty * dblRate) ''+ dblFAmt
            gv1.Rows(IntRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2)
            ''''' to calculate customer disc
            'Dim dt As New DataTable
            'Dim dblOrderQty As Double = 0
            'Dim dblCustDiscQty As Double = 0
            'Dim dblCustDiscAmt As Double = 0
            'Dim dblCustDiscPercentage As Double = 0
            'Dim dblApplyCustDisc As Double = 0
            'Dim dblTotCustDisc As Double = 0

            If ApplyTPT Then
                If chkTPT.Checked Then
                    GetDCDetails(IntRowNo)
                End If
            ElseIf ApplyCommission Then
                GetDCDetails(IntRowNo)
            End If


            Dim dblTotalDCAmt As Double = 0
            Dim dblTotalTCAmt As Double = 0
            Dim dblTotTaxRate As Double = GetCurrentRowTotalTaxRate(IntRowNo)
            ' Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblSCRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColSCRate).Value)
            Dim dblDisAmt As Double = 0
            'Dim dblTCAmt As Double = 0
            'Dim dblSCAmt As Double = dblAmt * (dblSCRate / 100)
            'gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = dblSCAmt
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(ColDCRate).Value), "") <> CompairStringResult.Equal AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCRate).Value) >= 0 Then
                If Not ApplyCommissionRateWithTax Then
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = gv1.Rows(IntRowNo).Cells(ColDCRate).Value
                Else
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = Math.Round(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCRate).Value) * 100 / (100 + dblTotTaxRate), 4)

                End If
                gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value = (clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colQty).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCUnitCF).Value)) / clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCCFUOM).Value)
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRate).Value
                Else
                    gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value
                End If
                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(ColTCRate).Value), "") <> CompairStringResult.Equal AndAlso clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColTCRate).Value) >= 0 Then
                    gv1.Rows(IntRowNo).Cells(ColTCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColTCRate).Value
                End If
                ' gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value
                gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = clsCommon.myCdbl(clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColSCRate).Value))
                dblTotalDCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColDCAmt).Value)
                dblTotalTCAmt = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColTCAmt).Value)
                If dblTotalDCAmt > 0 AndAlso ApplyCommission AndAlso Not chkSampling.Checked Then
                    'If ApplyCommission Then
                    '    If Not chkSampling.Checked Then
                    dblDisAmt = dblDisAmt + dblTotalDCAmt

                    '    End If
                    'End If
                End If
                If dblTotalTCAmt > 0 AndAlso ApplyTPT AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") <> CompairStringResult.Equal Then
                    'If ApplyTPT AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") <> CompairStringResult.Equal Then
                    dblDisAmt = dblDisAmt + dblTotalTCAmt
                    'End If
                End If
            End If
            Dim dblTotDiscAmt As Double = 0
            Dim dblAmtAfterDis As Double = 0
            ''richa agarwal 06 Aug,2019 calculate discoiunt amount in case of scheme item TEC/06/08/19-000981
            dblTotDiscAmt = dblDisAmt
            dblAmtAfterDis = dblAmt - dblDisAmt
            gv1.Rows(IntRowNo).Cells(colDisAmt).Value = dblTotDiscAmt
            gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = dblAmtAfterDis
            SetitemWiseTaxSetting(False, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value), IntRowNo)
            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)

                If rbtnTaxCalAutomatic.IsChecked Then
                    Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" & Strii)).Value)
                    If clsCommon.myLen(strTaxCode) > 0 Then
                        Dim dblTaxRate As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Rate" & Strii)).Value)
                        If Not chkisTCS.Checked AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" & Strii)).Value, "TCS") = CompairStringResult.Equal Then
                            dblTaxRate = 0
                        End If

                        'Dim dblTaxBaseAmt As Double = gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value
                        Dim IsTaxonBaseAmount As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(colIsTaxOnBaseAmount + clsCommon.myCstr(ii)).Value)
                        Dim dblBaseAmt As Double = dblAmtAfterDis
                        Dim dblTaxAmt As Double = 0
                        Dim dblOtherTaxAmt As Double = 0
                        If Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") <> CompairStringResult.Equal Then
                            dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)
                            dblBaseAmt += dblOtherTaxAmt
                        ElseIf Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") = CompairStringResult.Equal Then
                            dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth1)
                            ' If chkisTCS.Checked = False Then
                            dblBaseAmt += dblOtherTaxAmt
                        End If
                        If Not IsTaxonBaseAmount AndAlso clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
                            ' lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblAmtAfterDis + dblOtherTaxAmt)
                            Dim dblTotalBasicPrice As Double = 0
                            For n As Integer = 0 To gv1.Rows.Count - 1
                                If clsCommon.myLen(gv1.Rows(n).Cells(colICode).Value) > 0 Then
                                    dblTotalBasicPrice = dblTotalBasicPrice + clsCommon.myCdbl(gv1.Rows(n).Cells(colAmtAfterDis).Value)
                                End If
                            Next
                            If dblTotalBasicPrice > 0 Then
                                dblBaseAmt = (clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value) * clsCommon.myCdbl(txttcstaxbaseamount.Value)) / dblTotalBasicPrice
                            End If
                        End If
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = Math.Round(dblBaseAmt, 2)
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblBaseAmt, 3), 2, 4)
                        dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" & Strii)).Value = Math.Round(dblTaxAmt, 2)
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" & Strii)).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblTaxAmt, 3), 2, 4)
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" + Strii)).Value = Math.Round(dblTaxBaseAmt, 2)
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + Strii)).Value = Math.Round((dblTaxBaseAmt * dblTaxRate) / 100, 2)
                        If ((rbtnTaxable.IsChecked AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper()))) AndAlso (clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "CGST") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "SGST") <> CompairStringResult.Equal) Then
                            arrTaxableAuth.Add(strTaxCode.ToUpper())
                        End If
                        If (rbtnTaxable.IsChecked AndAlso Not arrTaxableAuth1.Contains(strTaxCode.ToUpper())) Then
                            arrTaxableAuth1.Add(strTaxCode.ToUpper())
                        End If
                    Else
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" & Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Rate" & Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" & Strii)).Value = Nothing
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
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" & Strii)).Value = dblCurrCalTax
                    End If
                End If
            Next
            'If dblQty > 0 Then
            '    Dim dblNetPrice As Double = dblAmtAfterDis / dblQty
            'End If
            TCSTaxRatio(IntRowNo)
            Dim dblTotTaxAmt As Double = GetCurrentRowTotalTaxAmt(IntRowNo)
            gv1.Rows(IntRowNo).Cells(colTBaseAmt).Value = Math.Round(dblAmtAfterDis, 2)
            gv1.Rows(IntRowNo).Cells(colTTaxAmt).Value = Math.Round(dblTotTaxAmt, 2)
            gv1.Rows(IntRowNo).Cells(colAmountWithTax).Value = Math.Round(dblTotTaxAmt + dblAmtAfterDis, 2)
            If chkSampling.Checked Then

                gv1.Rows(IntRowNo).Cells(colDisAmt).Value = Math.Round(dblDisAmt, 2) + Math.Round(dblAmtAfterDis, 2)

            Else
                If ApplyManualScheme Then
                    If clsCommon.myLen(gv1.Rows(IntRowNo).Cells(ColMainItem).Value) > 0 Then
                        gv1.Rows(IntRowNo).Cells(colTBaseAmt).Value = 0
                        gv1.Rows(IntRowNo).Cells(colTTaxAmt).Value = 0
                        gv1.Rows(IntRowNo).Cells(colDisAmt).Value = 0
                    Else
                        gv1.Rows(IntRowNo).Cells(colDisAmt).Value = Math.Round(dblDisAmt, 2)

                    End If
                Else
                    gv1.Rows(IntRowNo).Cells(colDisAmt).Value = Math.Round(dblDisAmt, 2)

                End If

            End If
            If ApplyManualScheme Then
                If clsCommon.myLen(gv1.Rows(IntRowNo).Cells(ColMainItem).Value) > 0 Then
                    gv1.Rows(IntRowNo).Cells(colAmountWithTax).Value = 0
                    gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = 0
                Else
                    gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = Math.Round(dblAmtAfterDis, 2)
                End If


            Else
                gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = Math.Round(dblAmtAfterDis, 2)

            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function TruncateToDecimalPlaces(value As Decimal, decimalPlaces As Integer) As Double
        Dim factor As Decimal = Math.Pow(10, decimalPlaces)
        Return Math.Truncate(value * factor) / factor
    End Function
    Private Sub UpdateCurrentRow1(ByVal IntRowNo As Integer)
        Try
            'SetTax(gv1.Rows(IntRowNo).Cells(colICode).Value)
            'Dim arrTaxableAuth As New List(Of String)
            'Dim arrTaxableAuth1 As New List(Of String)
            'Dim dblFAmt As Decimal = 0
            GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
            If Not GSTStatus Then
                strExcise = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Excisable from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'")) = "T", True, False)
            End If
            'Dim dblAlterQty As Decimal = 0
            Dim strICode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)
            Dim strUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblQty As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colQty).Value)
            'Dim dblRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colRate).Value)
            Dim dblItemBasicPrice As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colItemBasicPrice).Value)
            Dim wt_unit As String = 0
            'Dim wt_qty As Decimal = 0
            'Dim Item_Weight As Decimal = 0
            Dim TotalItem_Weight As Decimal = 0
            'Dim TotalItem_WeightMetric As Decimal = 0
            If clsCommon.myLen(strICode) > 0 Then
                wt_unit = clsItemMaster.GetItemWeightUnit(strICode, Nothing)
                TotalItem_Weight = clsItemMaster.getTotalItemWeight(strICode, strUnit, dblQty, Nothing)
            End If
            'Dim strOrgUnit As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)
            Dim dblBasicAmt As Decimal = dblQty * dblItemBasicPrice
            Dim dblAmt As Decimal = (dblQty * dblItemBasicPrice) ''+ dblFAmt
            gv1.Rows(IntRowNo).Cells(colAmt).Value = dblAmt
            ''''' to calculate customer disc
            'Dim dt As New DataTable
            'Dim dblOrderQty As Decimal = 0
            'Dim dblCustDiscQty As Decimal = 0
            'Dim dblCustDiscAmt As Decimal = 0
            'Dim dblCustDiscPercentage As Decimal = 0
            'Dim dblApplyCustDisc As Decimal = 0
            'Dim dblTotCustDisc As Decimal = 0
            GetDCDetails(IntRowNo)
            Dim dblTotalDCAmt As Decimal = 0
            Dim dblTotalTCAmt As Decimal = 0
            Dim dblTotTaxRate As Decimal = GetCurrentRowTotalTaxRate(IntRowNo)
            ' Dim dblDisPer As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colDisPer).Value)
            'Dim dblSCRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColSCRate).Value)
            Dim dblDisAmt As Decimal = 0
            'Dim dblTCAmt As Decimal = 0
            'Dim dblSCAmt As Decimal = dblAmt * (dblSCRate / 100)
            'gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = dblSCAmt
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(ColDCRate).Value), "") <> CompairStringResult.Equal AndAlso clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColDCRate).Value) >= 0 Then
                If Not ApplyCommissionRateWithTax Then
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = gv1.Rows(IntRowNo).Cells(ColDCRate).Value
                Else
                    gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value = Math.Round(gv1.Rows(IntRowNo).Cells(ColDCRate).Value * 100 / (100 + dblTotTaxRate), 4)

                End If
                gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value = (gv1.Rows(IntRowNo).Cells(colQty).Value * gv1.Rows(IntRowNo).Cells(ColDCUnitCF).Value) / gv1.Rows(IntRowNo).Cells(ColDCCFUOM).Value
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRate).Value
                Else
                    gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value
                End If
                If gv1.Rows(IntRowNo).Cells(ColTCRate).Value IsNot Nothing AndAlso clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColTCRate).Value) >= 0 Then
                    gv1.Rows(IntRowNo).Cells(ColTCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColTCRate).Value
                End If
                ' gv1.Rows(IntRowNo).Cells(ColDCAmt).Value = gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value * gv1.Rows(IntRowNo).Cells(ColDCRateWithTax).Value
                gv1.Rows(IntRowNo).Cells(ColSCAmt).Value = clsCommon.myCDecimal(clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColDCQtyinSU).Value) * clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColSCRate).Value))
                dblTotalDCAmt = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColDCAmt).Value)
                dblTotalTCAmt = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(ColTCAmt).Value)
                If dblTotalDCAmt > 0 AndAlso ApplyCommission Then
                    'If ApplyCommission Then
                    dblDisAmt = dblDisAmt + dblTotalDCAmt
                    'End If
                End If
                If dblTotalTCAmt > 0 AndAlso ApplyTPT Then
                    'If ApplyTPT Then
                    dblDisAmt = dblDisAmt + dblTotalTCAmt
                    'End If
                End If
            End If
            Dim dblTotDiscAmt As Decimal = 0
            Dim dblAmtAfterDis As Decimal = 0
            ''richa agarwal 06 Aug,2019 calculate discoiunt amount in case of scheme item TEC/06/08/19-000981
            dblTotDiscAmt = dblDisAmt
            dblAmtAfterDis = dblAmt - dblDisAmt
            gv1.Rows(IntRowNo).Cells(colDisAmt).Value = dblTotDiscAmt
            ' gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = dblAmtAfterDis
            'Dim dblTotalNonTabxableRate As Decimal = 0
            'Dim dblTotalNonTabxableAmount As Decimal = 0
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
                Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colTax + clsCommon.myCstr(ii)).Value)
                If clsCommon.myLen(strTaxCode) > 0 Then
                    If clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "K") = CompairStringResult.Equal Then
                        dblKKFTaxRate = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("colTax_Rate" + clsCommon.myCstr(ii)).Value)
                    ElseIf clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "M") = CompairStringResult.Equal Then
                        dblMNDTaxRate = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("colTax_Rate" + clsCommon.myCstr(ii)).Value)
                    ElseIf clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "GST") = CompairStringResult.Equal Then
                        dblGSTTaxRate.Add(clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("colTax_Rate" + clsCommon.myCstr(ii)).Value))
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
            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)
                If rbtnTaxCalAutomatic.IsChecked Then
                    Dim strTaxCode As String = clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value)
                    If clsCommon.myLen(strTaxCode) > 0 Then
                        'Dim dblTaxRate As Decimal = clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Rate" + Strii)).Value)
                        'Dim dblTaxBaseAmt As Decimal = gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value
                        'Dim IsTaxonBaseAmount As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(colIsTaxOnBaseAmount + clsCommon.myCstr(ii)).Value)
                        'Dim dblBaseAmt As Decimal = dblAmtAfterDis
                        'Dim dblTaxAmt As Decimal = 0
                        'Dim dblOtherTaxAmt As Decimal = 0
                        'If Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") <> CompairStringResult.Equal Then
                        '    dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)
                        '    dblBaseAmt += dblOtherTaxAmt
                        'ElseIf Not IsTaxonBaseAmount AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "TCS") = CompairStringResult.Equal Then
                        '    dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth1)
                        '    dblBaseAmt += dblOtherTaxAmt
                        'End If
                        'If Not IsTaxonBaseAmount AndAlso clsCommon.myCDecimal(txttcstaxbaseamount.Value) > 0 Then
                        '    ' lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblAmtAfterDis + dblOtherTaxAmt)
                        '    Dim dblTotalBasicPrice As Decimal = 0
                        '    For n As Integer = 0 To gv1.Rows.Count - 1
                        '        If clsCommon.myLen(gv1.Rows(n).Cells(colICode).Value) > 0 Then
                        '            dblTotalBasicPrice = dblTotalBasicPrice + clsCommon.myCDecimal(gv1.Rows(n).Cells(colAmtAfterDis).Value)
                        '        End If
                        '    Next
                        '    If dblTotalBasicPrice > 0 Then
                        '        dblBaseAmt = (clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value) * clsCommon.myCDecimal(txttcstaxbaseamount.Value)) / dblTotalBasicPrice
                        '    End If
                        'End If
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" + Strii)).Value = Math.Round(dblBaseAmt, 2)
                        'dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                        'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + Strii)).Value = Math.Round(dblTaxAmt, 2)
                        ''gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" + Strii)).Value = Math.Round(dblTaxBaseAmt, 2)
                        ''gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + Strii)).Value = Math.Round((dblTaxBaseAmt * dblTaxRate) / 100, 2)
                        'If ((rbtnTaxable.IsChecked AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper()))) AndAlso (clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "CGST") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value, "SGST") <> CompairStringResult.Equal) Then
                        '    arrTaxableAuth.Add(strTaxCode.ToUpper())
                        'End If
                        'If (rbtnTaxable.IsChecked AndAlso Not arrTaxableAuth1.Contains(strTaxCode.ToUpper())) Then
                        '    arrTaxableAuth1.Add(strTaxCode.ToUpper())
                        'End If
                        If clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "K") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = dblKKFTaxValue
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(dblProductValue, 2, 4)
                        ElseIf clsCommon.CompairString(clsTaxCalculation.GetTaxType(strTaxCode, Nothing), "M") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = dblMNDTaxValue
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(dblProductValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "CGST") = CompairStringResult.Equal OrElse clsCommon.CompairString(strTaxCode, "IGST") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = dblGSTTaxValue1
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "SGST") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = dblGSTTaxValue2
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                            'ElseIf clsCommon.CompairString(strTaxCode, "IGST") = CompairStringResult.Equal Then
                            '    gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = dblGSTTaxValue1
                            '    gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" + Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "EXEMPTED") = CompairStringResult.Equal Then
                            gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = 0
                            gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(dblTaxableValue, 2, 4)
                        ElseIf clsCommon.CompairString(strTaxCode, "TCS") = CompairStringResult.Equal Then
                            If dblOutstandingAmount > AmountToCheckCustomerOutstandingForTCSTax Then
                                gv1.Rows(IntRowNo).Cells("colTax_Amt" & clsCommon.myCstr(ii)).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces((dblTaxableValue + dblGSTTaxValue1 + dblGSTTaxValue2) * (clsCommon.myCDecimal(gv1.Rows(IntRowNo).Cells("colTax_Rate" + clsCommon.myCstr(ii)).Value) / 100), 3), 2, 4)
                                gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblTaxableValue + dblGSTTaxValue1 + dblGSTTaxValue2, 3), 2, 4)
                                'TCSTaxRatio(IntRowNo)
                            End If

                        End If
                    Else
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" & Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Rate" & Strii)).Value = Nothing
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" & Strii)).Value = Nothing
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
                            dblCurrCalTax = clsCommon.myRoundOFF(clsCommon.myCDecimal(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 3, 4)
                        End If
                        gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + Strii)).Value = dblCurrCalTax
                    End If
                End If
            Next
            'If dblQty > 0 Then
            '    Dim dblNetPrice As Decimal = dblAmtAfterDis / dblQty
            'End If
            Dim dblTotTaxAmt As Decimal = GetCurrentRowTotalTaxAmt(IntRowNo)
            Dim dblTotTCSTaxAmt As Decimal = GetCurrentRowTotalTCSTaxAmt(IntRowNo)
            gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblAmtAfterDis - (dblTotTaxAmt - dblTotTCSTaxAmt), 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colTBaseAmt).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblAmtAfterDis - dblTotTaxAmt, 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colTTaxAmt).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblTotTaxAmt, 3), 2, 4)
            gv1.Rows(IntRowNo).Cells(colAmountWithTax).Value = clsCommon.myRoundOFF(TruncateToDecimalPlaces(dblAmtAfterDis, 3), 2, 4)
            If chkSampling.Checked Then
                gv1.Rows(IntRowNo).Cells(colDisAmt).Value = clsCommon.myRoundOFF(dblDisAmt, 2, 4) + clsCommon.myRoundOFF(dblAmtAfterDis, 2, 4)

            Else
                gv1.Rows(IntRowNo).Cells(colDisAmt).Value = clsCommon.myRoundOFF(dblDisAmt, 2, 4)

            End If
            'gv1.Rows(IntRowNo).Cells(colAmtAfterDis).Value = Math.Round(dblAmtAfterDis, 2)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Function GetCurrentRowTotalTCSTaxAmt(ByVal IntRowNo As Integer) As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value), "TCS") = CompairStringResult.Equal Then
                If IntRowNo < 0 Then
                    dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Amt" + strii)).Value)
                Else
                    dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + strii)).Value)
                End If
            End If

        Next
        Return dblTotTax
    End Function
    Private Function GetCurrentRowTotalTaxAmt(ByVal IntRowNo As Integer) As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            'If Not clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value), "TCS") = CompairStringResult.Equal Then
            If IntRowNo < 0 Then
                dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Amt" + strii)).Value)
            Else
                dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + strii)).Value)
            End If
            'End If

        Next
        Return dblTotTax
    End Function
    Private Function GetCurrentRowTotalTaxRate(ByVal IntRowNo As Integer) As Double
        Dim dblTotRate As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value), "TCS") <> CompairStringResult.Equal Then
                If IntRowNo < 0 Then
                    dblTotRate = dblTotRate + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Rate" & strii)).Value)
                Else
                    dblTotRate = dblTotRate + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Rate" & strii)).Value)
                End If
            End If


        Next
        Return dblTotRate
    End Function
    Private Function GetCurrentRowSurTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" & strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Amt" & strii)).Value)
                End If
            Else
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" & strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" & strii)).Value)
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
                        dblRet = dblRet + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Amt" + strii)).Value)
                    End If
                Else
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + strii)).Value)
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
                gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" & strII)).Value = Nothing
                gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Base_Amt" & strII)).Value = Nothing
                If isBlankRate Then
                    gv1.CurrentRow.Cells(clsCommon.myCstr("colTax_Rate" & strII)).Value = Nothing
                End If
            Else
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" & strII)).Value = Nothing
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" & strII)).Value = Nothing
                If isBlankRate Then
                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" & strII)).Value = Nothing
                End If
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Amt" & strII)).Value = Nothing
            End If
        Next
    End Sub
    Private Sub UpdateAllTotals()
        dblNetAmt = 0
        TotalQuantity = 0
        TotalCan = 0
        TotalBox = 0
        TotalCrate = 0
        Dim dblTotalDocAmt As Decimal = 0
        Dim dblTotalTcsAmt As Decimal = 0
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
        Dim dblTaxTotAmt As Double = 0
        Dim dblDisAmt As Double = 0
        Dim dblCommAmt As Double = 0
        Dim dblTCAmt As Double = 0
        Dim dblSCAmt As Double = 0
        Dim dblTotalQtyinKG As Double = 0
        'Dim qryCrate As String
        'Dim qryCAN As String
        'Dim qryBox As String
        For i As Integer = 0 To gv1.Rows.Count - 1
            dblNetAmt = dblNetAmt + clsCommon.myCdbl(gv1.Rows(i).Cells(colAmt).Value)
            TotalQuantity = TotalQuantity + clsCommon.myCdbl(gv1.Rows(i).Cells(colQty).Value)
            dblTotalDocAmt = dblTotalDocAmt + clsCommon.myCdbl(gv1.Rows(i).Cells(colAmountWithTax).Value)
            'qryCrate = clsDBFuncationality.getSingleValue("select * from tspl_unit_master where unit_code='" & clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value) & "'  and crate_type='Y'")
            'qryCAN = clsDBFuncationality.getSingleValue("select * from tspl_unit_master where unit_code='" & clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value) & "'  and can_type='Y'")
            'qryBox = clsDBFuncationality.getSingleValue("select * from tspl_unit_master where unit_code='" & clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value) & "'  and Box_type='Y'")
            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value), "Can") = CompairStringResult.Equal Then
                TotalCan = TotalCan + clsCommon.myCdbl(gv1.Rows(i).Cells(colQty).Value)
            ElseIf clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value), "Box") = CompairStringResult.Equal Then
                TotalBox = TotalBox + clsCommon.myCdbl(gv1.Rows(i).Cells(colQty).Value)
            ElseIf clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(i).Cells(colUnit).Value), "Crate") = CompairStringResult.Equal Then
                TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(i).Cells(colQty).Value)
            End If
            dblTotalQtyinKG = dblTotalQtyinKG + clsCommon.myCdbl(gv1.Rows(i).Cells(colQtyinKG).Value)
            dblDisAmt = dblDisAmt + clsCommon.myCdbl(gv1.Rows(i).Cells(colDisAmt).Value)
            dblCommAmt = dblCommAmt + clsCommon.myCdbl(gv1.Rows(i).Cells(ColDCAmt).Value)
            dblTCAmt = dblTCAmt + clsCommon.myCdbl(gv1.Rows(i).Cells(ColTCAmt).Value)
            dblSCAmt = dblSCAmt + clsCommon.myCdbl(gv1.Rows(i).Cells(ColSCAmt).Value)
            dblTaxAmt1 = dblTaxAmt1 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(1))).Value), 3, 4)
            dblTaxAmt2 = dblTaxAmt2 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(2))).Value), 3, 4)
            dblTaxAmt3 = dblTaxAmt3 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(3))).Value), 3, 4)
            dblTaxAmt4 = dblTaxAmt4 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(4))).Value), 3, 4)
            dblTaxAmt5 = dblTaxAmt5 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(5))).Value), 3, 4)
            dblTaxAmt6 = dblTaxAmt6 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(6))).Value), 3, 4)
            dblTaxAmt7 = dblTaxAmt7 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(7))).Value), 3, 4)
            dblTaxAmt8 = dblTaxAmt8 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(8))).Value), 3, 4)
            dblTaxAmt9 = dblTaxAmt9 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(9))).Value), 3, 4)
            dblTaxAmt10 = dblTaxAmt10 + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(10))).Value), 3, 4)
            dblTaxBaseAmt1 = dblTaxBaseAmt1 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(1))).Value)
            dblTaxBaseAmt2 = dblTaxBaseAmt2 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(2))).Value)
            dblTaxBaseAmt3 = dblTaxBaseAmt3 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(3))).Value)
            dblTaxBaseAmt4 = dblTaxBaseAmt4 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(4))).Value)
            dblTaxBaseAmt5 = dblTaxBaseAmt5 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(5))).Value)
            dblTaxBaseAmt6 = dblTaxBaseAmt6 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(6))).Value)
            dblTaxBaseAmt7 = dblTaxBaseAmt7 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(7))).Value)
            dblTaxBaseAmt8 = dblTaxBaseAmt8 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(8))).Value)
            dblTaxBaseAmt9 = dblTaxBaseAmt9 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(9))).Value)
            dblTaxBaseAmt10 = dblTaxBaseAmt10 + clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(10))).Value)
            dblTaxTotAmt = dblTaxTotAmt + clsCommon.myRoundOFF(clsCommon.myCdbl(gv1.Rows(i).Cells(colTTaxAmt).Value), 3, 4)
            For j As Integer = 1 To 10
                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(i).Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(j))).Value), "TCS") = CompairStringResult.Equal Then
                    dblTotalTcsAmt = dblTotalTcsAmt + clsCommon.myRoundOFF(TruncateToDecimalPlaces(clsCommon.myCdbl(gv1.Rows(i).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(j))).Value), 2), 2, 4)
                End If
            Next

        Next
        If rbtnTaxCalAutomatic.IsChecked Then
            For ii As Integer = 1 To gv2.Rows.Count
                If rbtnTaxable.IsChecked OrElse (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code IN (select Tax_Code  from TSPL_TAX_GROUP_DETAILS WHERE TAX_GROUP_CODE='" & txtTaxGroup.Value & "') AND Is_TCS ='Y'")) > 0) Then
                    Select Case (ii)
                        Case 1
                            If chkisTCS.Checked Then
                                If (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then

                                    lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblTaxBaseAmt1)
                                    If clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
                                        dblTaxBaseAmt1 = clsCommon.myCdbl(txttcstaxbaseamount.Value)
                                    Else
                                        dblTaxBaseAmt1 = clsCommon.myCdbl(lblActualTCSTaxBaseAmt.Text)
                                    End If
                                    dblTaxAmt1 = dblTotalTcsAmt
                                End If
                            End If

                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = clsCommon.myRoundOFF(dblTaxAmt1, 2, 4)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = clsCommon.myRoundOFF(dblTaxBaseAmt1, 2, 4)

                            If dblTaxBaseAmt1 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt1 * 100) / dblTaxBaseAmt1, 2, 4)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 2
                            If chkisTCS.Checked AndAlso (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                'If (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblTaxBaseAmt1 + dblTaxAmt1)
                                If clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt2 = clsCommon.myCdbl(txttcstaxbaseamount.Value)
                                Else
                                    dblTaxBaseAmt2 = clsCommon.myCdbl(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt2 = dblTotalTcsAmt
                                'End If
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = clsCommon.myRoundOFF(dblTaxAmt2, 2, 4)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = clsCommon.myRoundOFF(dblTaxBaseAmt2, 2, 4)
                            If chkisTCS.Checked AndAlso clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                'If clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTotalTcsAmt
                                'End If
                            End If
                            If dblTaxBaseAmt2 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt2 * 100) / dblTaxBaseAmt2, 2, 4)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 3
                            If chkisTCS.Checked AndAlso (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                'If (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblTaxBaseAmt1 + dblTaxAmt1 + dblTaxAmt2)
                                If clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt3 = clsCommon.myCdbl(txttcstaxbaseamount.Value)
                                Else
                                    dblTaxBaseAmt3 = clsCommon.myCdbl(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt3 = dblTotalTcsAmt
                                'End If
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = clsCommon.myRoundOFF(dblTaxAmt3, 2, 4)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = clsCommon.myRoundOFF(dblTaxBaseAmt3, 2, 4)
                            If chkisTCS.Checked AndAlso clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                'If clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTotalTcsAmt
                                'End If
                            End If
                            If dblTaxBaseAmt3 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt3 * 100) / dblTaxBaseAmt3, 2, 4)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 4
                            If chkisTCS.Checked AndAlso (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                'If (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblTaxBaseAmt1 + dblTaxAmt1 + dblTaxAmt2 + dblTaxAmt3)
                                If clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt4 = clsCommon.myCdbl(txttcstaxbaseamount.Value)
                                Else
                                    dblTaxBaseAmt4 = clsCommon.myCdbl(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt4 = dblTotalTcsAmt
                                'End If
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = clsCommon.myRoundOFF(dblTaxAmt4, 2, 4)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = clsCommon.myRoundOFF(dblTaxBaseAmt4, 2, 4)
                            If chkisTCS.Checked AndAlso clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                'If clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTotalTcsAmt
                                'End If
                            End If
                            If dblTaxBaseAmt4 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt4 * 100) / dblTaxBaseAmt4, 2, 4)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 5
                            If chkisTCS.Checked AndAlso (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                'If (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code='" & clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value) & "' AND Is_TCS ='Y'")) > 0) Then
                                lblActualTCSTaxBaseAmt.Text = clsCommon.myFormat(dblTaxBaseAmt1 + dblTaxAmt1 + dblTaxAmt2 + dblTaxAmt3 + dblTaxAmt4)
                                If clsCommon.myCdbl(txttcstaxbaseamount.Value) > 0 Then
                                    dblTaxBaseAmt5 = clsCommon.myCdbl(txttcstaxbaseamount.Value)
                                Else
                                    dblTaxBaseAmt5 = clsCommon.myCdbl(lblActualTCSTaxBaseAmt.Text)
                                End If
                                dblTaxAmt5 = dblTotalTcsAmt
                                'End If
                            End If
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = clsCommon.myRoundOFF(dblTaxAmt5, 2, 4)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = clsCommon.myRoundOFF(dblTaxBaseAmt5, 2, 4)
                            If chkisTCS.Checked Then
                                If clsCommon.CompairString(clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTTaxAutCode).Value), "TCS") = CompairStringResult.Equal Then
                                    gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTotalTcsAmt
                                End If
                            End If
                            If dblTaxBaseAmt5 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt5 * 100) / dblTaxBaseAmt5, 2, 4)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 6
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt6, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt6, 2)
                            If dblTaxBaseAmt6 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt6 * 100) / dblTaxBaseAmt6, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 7
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt7, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt7, 2)
                            If dblTaxBaseAmt7 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt7 * 100) / dblTaxBaseAmt7, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 8
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt8, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt8, 2)
                            If dblTaxBaseAmt8 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt8 * 100) / dblTaxBaseAmt8, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 9
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt9, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt9, 2)
                            If dblTaxBaseAmt9 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt9 * 100) / dblTaxBaseAmt9, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                        Case 10
                            gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = Math.Round(dblTaxAmt10, 2)
                            gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = Math.Round(dblTaxBaseAmt10, 2)
                            If dblTaxBaseAmt10 <> 0 Then
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = Math.Round((dblTaxAmt10 * 100) / dblTaxBaseAmt10, 2)
                            Else
                                gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                            End If
                    End Select
                End If
            Next
        End If
        'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
        '    Dim amtwithdis As Double = Math.Round(clsCommon.myCdbl(dblNetAmt) - dblTaxTotAmt, 2)
        '    lblTotRAmt1.Text = Math.Round(clsCommon.myCdbl(dblNetAmt) - dblTaxTotAmt, 2)
        '    txtCan.Text = Math.Round(clsCommon.myCdbl(TotalCan), 2)
        '    txtBox.Text = Math.Round(clsCommon.myCdbl(TotalBox), 2)
        '    txtCrate.Text = Math.Round(clsCommon.myCdbl(TotalCrate), 2)
        '    lblAmtWithDiscount.Text = clsCommon.myFormat(lblTotRAmt1.Text)
        '    lblDiscountAmt.Text = clsCommon.myFormat(dblDisAmt)
        '    lblAmtAfterDiscount.Text = clsCommon.myFormat(Math.Round(clsCommon.myCdbl(amtwithdis), 2) - dblDisAmt)
        '    lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt)
        '    lblTotRAmt.Text = clsCommon.myFormat(dblNetAmt)
        '    txtDCAmt.Text = clsCommon.myFormat(dblCommAmt)
        '    txtTCAmt.Text = clsCommon.myFormat(dblTCAmt)
        '    txtSecurity.Text = clsCommon.myFormat(dblSCAmt)
        '    lblTotalDocAmt.Text = clsCommon.myFormat(dblNetAmt)
        'Else
        If chkSampling.Checked Then
            dblDisAmt = Math.Round(clsCommon.myCdbl(dblNetAmt), 2)
        End If

        txtCan.Text = Math.Round(clsCommon.myCdbl(TotalCan), 2)
        txtBox.Text = Math.Round(clsCommon.myCdbl(TotalBox), 2)
        txtCrate.Text = Math.Round(clsCommon.myCdbl(TotalCrate), 2)
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            lblAmtWithDiscount.Text = clsCommon.myRoundOFF(dblNetAmt - (dblTaxTotAmt - dblTotalTcsAmt), 2, 4)
            lblDiscountAmt.Text = clsCommon.myFormat(dblDisAmt)
            lblAmtAfterDiscount.Text = clsCommon.myFormat(Math.Round(clsCommon.myCdbl(dblNetAmt - (dblTaxTotAmt - dblTotalTcsAmt)), 2) - dblDisAmt)
        Else
            lblAmtWithDiscount.Text = dblNetAmt - (dblDisAmt)
            lblDiscountAmt.Text = clsCommon.myFormat(dblDisAmt)
            lblAmtAfterDiscount.Text = clsCommon.myFormat(Math.Round(clsCommon.myCdbl(dblNetAmt - dblDisAmt), 2))
        End If
        'lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt + dblTotalTcsAmt)
        lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt)
        If chkSampling.Checked Then
            lblTotRAmt1.Text = Math.Round(clsCommon.myCdbl(dblNetAmt - dblDisAmt), 2)
            'lblTotRAmt.Text = clsCommon.myFormat((dblNetAmt - dblDisAmt) + dblTotalTcsAmt + dblTaxTotAmt)
            'lblTotalDocAmt.Text = clsCommon.myFormat((dblNetAmt - dblDisAmt) + dblTotalTcsAmt + dblTaxTotAmt)
            lblTotRAmt.Text = clsCommon.myFormat((dblNetAmt - dblDisAmt) + dblTaxTotAmt)
            lblTotalDocAmt.Text = clsCommon.myFormat((dblNetAmt - dblDisAmt) + dblTaxTotAmt)
        Else

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                lblTotRAmt1.Text = Math.Round(clsCommon.myCdbl(dblTotalDocAmt + dblTotalTcsAmt), 2)
                lblTotRAmt.Text = clsCommon.myFormat(dblTotalDocAmt + dblTotalTcsAmt)
                lblTotalDocAmt.Text = clsCommon.myFormat(dblTotalDocAmt + dblTotalTcsAmt)
            Else
                lblTotRAmt1.Text = Math.Round(clsCommon.myCdbl(dblNetAmt), 2)
                lblTotRAmt.Text = clsCommon.myFormat(dblTotalDocAmt)
                lblTotalDocAmt.Text = clsCommon.myFormat(dblTotalDocAmt)
            End If
            'lblTotRAmt.Text = clsCommon.myFormat(dblTotalDocAmt + dblTotalTcsAmt)
            'lblTotalDocAmt.Text = clsCommon.myFormat(dblTotalDocAmt + dblTotalTcsAmt)

        End If
        txtDCAmt.Text = clsCommon.myFormat(dblCommAmt)
        txtTCAmt.Text = clsCommon.myFormat(dblTCAmt)
        txtSecurity.Text = clsCommon.myFormat(dblSCAmt)
        'End If
        If ApplyRoundOffZero OrElse chkBPL.Checked Then
            If Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0) > clsCommon.myCdbl(lblTotRAmt.Text) Then
                'TxtRoundoff.Text = Math.Round(clsCommon.myCdbl(clsCommon.myCdbl(clsCommon.myCdbl(lblTotRAmt1.Text)) - Math.Round(clsCommon.myCdbl(lblTotRAmt1.Text), 0)), 2)
                TxtRoundoff.Text = Math.Round(Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0) - clsCommon.myCdbl(lblTotRAmt.Text), 2)
                lblTotRAmt.Text = Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0)
                lblTotRAmt1.Text = Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0)
                lblTotalDocAmt.Text = Math.Round(clsCommon.myCdbl(lblTotalDocAmt.Text), 0)
            Else
                'TxtRoundoff.Text = Math.Round(clsCommon.myCdbl(clsCommon.myCdbl(lblTotRAmt1.Text) - Math.Round(clsCommon.myCdbl(lblTotRAmt1.Text))), 2)
                TxtRoundoff.Text = Math.Round(Math.Round(clsCommon.myCdbl(lblTotRAmt.Text)) - clsCommon.myCdbl(lblTotRAmt.Text), 2)
                lblTotRAmt.Text = Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0)
                lblTotRAmt1.Text = Math.Round(clsCommon.myCdbl(lblTotRAmt.Text), 0)
                lblTotalDocAmt.Text = Math.Round(clsCommon.myCdbl(lblTotalDocAmt.Text), 0)
            End If
        End If
        If DeductTPTFromDocAmt Then
            txtTPTComAmt.Text = clsCommon.myFormat(dblTCAmt)
            lblGrossAmount.Text = clsCommon.myFormat(clsCommon.myCdbl(lblTotRAmt.Text) - dblTCAmt)
        Else
            lblGrossAmount.Text = clsCommon.myFormat(clsCommon.myCdbl(lblTotRAmt.Text))
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            If dblTotalQtyinKG > FORPRICE Then
                If Not IsTotalQtyinKG Then
                    For ii As Integer = 0 To gv1.Rows.Count - 2
                        ItemPrice(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(ii).Cells(colQty).Value), ii, True)
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                            UpdateCurrentRow1(ii)
                        Else
                            UpdateCurrentRow(ii)
                        End If
                        'UpdateCurrentRow(ii)
                        IsTotalQtyinKG = True
                        IsExistsForPrice = True
                    Next
                    UpdateAllTotals()
                End If
            Else
                If IsExistsForPrice AndAlso Not IsTotalQtyinKG Then
                    'If Not IsTotalQtyinKG Then
                    For ii As Integer = 0 To gv1.Rows.Count - 2
                        ItemPrice(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value), clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), clsCommon.myCdbl(gv1.Rows(ii).Cells(colQty).Value), ii, False)
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                            UpdateCurrentRow1(ii)
                        Else
                            UpdateCurrentRow(ii)
                        End If
                        IsTotalQtyinKG = True
                        IsExistsForPrice = False
                    Next
                    UpdateAllTotals()
                    'End If
                End If
            End If
        End If
        'Try
        '    lblTCSAmount.Text = Math.Round(Math.Round(clsCommon.myCdbl(dblTotalDocAmt), 2) * GetTCSRate(txtVendorNo.Value) / 100, 2)
        'Catch ex As Exception
        'End Try
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
    'Private Function GetBaseTaxAmount(ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
    'For ii As Integer = 0 To intEndCol
    '    If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAutCode).Value)) = CompairStringResult.Equal Then
    '        Return clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
    '    End If
    'Next
    'Return 0
    'End Function
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        AddNew()
    End Sub
    Sub AddNew()
        If ApplyTPT AndAlso ApplyCommission Then
            clsCommon.MyMessageBoxShow(Me, "Apply TPT and Apply Commission Both setting is on, please disable one.", Me.Text)
            Me.Close()
        End If
        UsLock1.Status = ERPTransactionStatus.Pending
        UcAttachment1.BlankAllControls()
        'VendorCodeForChangeIndent = ""
        IsTotalQtyinKG = False
        rgbItemType.Visible = True
        rgbItemType.Enabled = True
        If ApplyDefaultTCSIsChecked Then
            chkTPT.Checked = True
        Else
            chkTPT.Checked = False
        End If
        chkExcludeKKFMND.Checked = False
        chkExcludeKKFMND.Enabled = True
        chkisTCS.Checked = False
        chkisTCS.Visible = False
        lblShiftType.Text = ""
        btnGatePassPrint.Visible = False
        lblCancelStatus.Text = ""
        lblCreatedDateAndTime.Text = ""
        chkDistributor.Checked = False
        chkDCS.Checked = False
        chkDCS.Visible = False
        chkSampling.Checked = False
        chkSampling.Enabled = True
        txtDCSDemandNo.Text = ""
        lblDCSDemand.Visible = False
        txtDCSDemandNo.Visible = False
        chkNoTranspoter.Checked = False
        'If DefaultEnableEWayBill Then
        '    chkIsEwayBill.Checked = True
        'Else
        '    chkIsEwayBill.Checked = False
        'End If
        chkIsEwayBill.Enabled = True
        chkIsEwayBill.Checked = False
        chkBPL.Checked = False
        chkAPS.Checked = False
        chkGhee.Checked = False
        chkGhee.Enabled = True
        txtCouponCode.Text = ""
        txtCouponDate.Text = clsCommon.GETSERVERDATE()
        txtBPLName.Text = ""
        txtBPLRemark.Text = ""
        lblUploadingDate.Text = ""
        Is_Cancelled = 0
        FlagCreateDo = False
        chkNoCrateIssue.Checked = False
        'btnCopy.Enabled = True
        DOStatus = 0
        BookingStatus = 0
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtSupplyDate.Value = txtDate.Value
        GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        lblOutstandingDesc.Text = ""
        lblReceiptAmtDesc.Text = ""
        lblUnbilledMilkAmt.Text = ""
        txtDate.Enabled = True
        'chkIsTaxable.Enabled = True
        BlankAllControls()
        LoadBlankGrid()
        LoadBlankGridTax()
        'LoadBlankGridAC()
        lblAmtWithDiscount.Text = ""
        lblAmtAfterDiscount.Text = ""
        lblDiscountAmt.Text = ""
        lblTaxAmt.Text = ""
        txtAdharNo.Text = ""
        txtMobileNo.Text = ""

        lblTotRAmt.Text = ""
        txtTCSTaxRate.Text = ""
        lblActualTCSTaxBaseAmt.Text = ""
        txtDCAmt.Text = ""
        txtSecurity.Text = ""
        TxtRoundoff.Text = ""
        txtComment.Text = ""
        txttcstaxbaseamount.Text = ""
        txtSubLocation.Enabled = False
        txtSubLocation.Value = ""
        lblSubLocation.Text = ""
        txtTaxGroup.Value = ""
        lblTaxGrpName.Text = ""
        If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                txtSubLocation.Enabled = True
            Else
                txtSubLocation.Enabled = False
            End If
            txtSubLocation.Value = ""
            lblSubLocation.Text = ""
        End If
        isNewEntry = True
        btnSave.Text = "Save"
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnDelete.Enabled = True
        btnCreateAndPrintInvoice.Enabled = False
        'btnCopy.Enabled = True
        txtDate.Focus()
        gv1.Rows.AddNew()
        'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem

        txtVendorNo.Enabled = True
        txtLocation.Enabled = True
        txtDocNo.Value = ""
        txtShipToLocation.Value = ""
        lblShipToLocation.Text = ""
        lblDONumber.Text = ""
        txtRouteNo.Value = ""
        lblRouteDesc.Text = ""
        txtVendorNo.Value = ""
        lblVendorName.Text = ""
        lblroutecode.Text = ""
        lblroutename.Text = ""
        lblvehiclecode.Text = ""
        lblvehicleName.Text = ""
        'txtLocation.Value = ""
        'lblLocation.Text = ""
        txtPriceCode.Text = ""
        txtPONo.Text = ""
        txtCan.Text = 0
        txtCrate.Text = 0
        txtBox.Text = 0
        txtCustPODate.Checked = False
        txtCustPODate.Value = clsCommon.GETSERVERDATE()
        txtReceipt.Value = ""
        txtEx_Factory_Date.Checked = False
        txtEx_Factory_Date.Value = clsCommon.GETSERVERDATE()
        LblUpdatedVehicleCode.Text = ""
        LblUpdatedVehicleDesc.Text = ""
        chkSampling.Checked = False
        ItemTypePanel.Enabled = True
        'LoadBookingType()
        cmbcashcredit.Text = "Select"
        'cmbGatePassType.Text = "Select"
        cmbGatePassType.Enabled = True
        lblBoothStation.Text = ""
        ''richa VIJ/03/12/19-000091
        lblCreditLimit.Text = ""
        lblAdvanceSecurity.Text = ""
        lblReverseAdvanceSec.Text = ""
        lblPendingDO.Text = ""
        lblLedgerOutstanding.Text = ""
        lblShortcloseDO.Text = ""
        lblRefund.Text = ""
        lblReverseRefund.Text = ""
        lblARSecurity.Text = ""
        lblTotalOutstansing.Text = ""
        lbltotalOutstanding1.Text = ""
        lblTotalSecurity11.Text = ""
        txtItemSearch.Text = ""
        lblCreatedByValue.Text = ""
        chkGatePass.Checked = False
        IsLoadBookingType = False
        chkGatePass.Enabled = True
        txtRouteCode1.Text = ""
        txtRouteName1.Text = ""
        lblPriceCodeDesc.Text = ""
        txtVehicleCode.Value = ""
        txtVehicleName.Text = ""
        txtRouteCode1.Enabled = False
        txtRouteName1.Enabled = False
        txtVehicleCode.Enabled = True
        txtVehicleName.Enabled = True
        btnGatepass.Enabled = False
        chkcashsale.Checked = False
        chkcashsale.Enabled = True
        txtTPTComAmt.Text = 0
        lblGrossAmount.Text = 0
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            chkcashsale.Text = "Against Receipt"

        End If
        txtChequeNo.Text = ""
        txtReceiverName.Text = ""
        cmbPaymentType.SelectedIndex = 0
        chkManualVehicle.Checked = True
        fndTransporter.Value = ""
        lblTransporter.Text = ""
        txtDescription.Text = ""
        txtManualVehicle.Text = ""
        If chkManualVehicle.Checked Then
            txtManualVehicle.Enabled = True
        Else
            txtManualVehicle.Enabled = False
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
            If Not chkcashsale.Checked Then
                txtChequeNo.Visible = False
                lblPaymentType.Visible = False
                cmbPaymentType.Visible = False
                lblChequeNo.Visible = False
            End If
        Else
            txtChequeNo.Visible = False
            lblPaymentType.Visible = False
            cmbPaymentType.Visible = False
            lblChequeNo.Visible = False
        End If
        ENABLEDISABLECONTROLS()
        If ShowBookingTypeDropDownonDairyBookingCustomer Then
            txtVendorNo.Focus()
        End If
        btnprinte_wayBill.Visible = False
        txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "' "))
        If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                txtSubLocation.Enabled = True
                txtSubLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Sub_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "'"))
                lblSubLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtSubLocation.Value & "'"))
            Else
                txtSubLocation.Enabled = False
                txtSubLocation.Value = ""
                lblSubLocation.Text = ""
            End If

        End If
        If HideOutstanding Then
            lblOutStanding.Visible = False
            lblOutstandingDesc.Visible = False
            lblUnbilledMilk.Visible = False
            lblUnbilledMilkAmt.Visible = False
        Else
            lblOutStanding.Visible = True
            lblOutstandingDesc.Visible = True
            lblUnbilledMilk.Visible = True
            lblUnbilledMilkAmt.Visible = True
        End If
        If ApplyTPT Then
            chkTPT.Visible = True
        Else
            chkTPT.Visible = False
        End If
        PaymentType()
        If DefaultEnableNoTransporter Then
            fndTransporter.Enabled = False
            lblTransporter.Enabled = False
            chkNoTranspoter.Checked = True
        Else
            fndTransporter.Enabled = True
            lblTransporter.Enabled = True
            chkNoTranspoter.Checked = False
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            lblReceipt.Visible = False
            txtReceipt.Visible = False
            lblReceiptAmt.Visible = False
            lblReceiptAmtDesc.Visible = False
        Else
            lblReceipt.Visible = True
            txtReceipt.Visible = True
            lblReceiptAmt.Visible = True
            lblReceiptAmtDesc.Visible = True
        End If

    End Sub
    Sub ENABLEDISABLECONTROLS()
        If ShowBookingTypeDropDownonDairyBookingCustomer Then
            'cmbcashcredit.Visible = False
            'lblCredit.Visible = False
            MyLabel2.Visible = True
            MyLabel2.Text = "Booth Station"
            txtSalesman.Visible = False
            lblSalesman.Visible = False
            lblBoothStation.Visible = True
            Panel1.Visible = False
            lbl_ExFactoryDate.Visible = False
            txtEx_Factory_Date.Visible = False
            Panel3.Visible = False
            Panel4.Visible = True
            btn_QtyReset.Visible = True
            chkGatePass.Visible = False
            lblGatePassType.Visible = False
            cmbGatePassType.Visible = False
            RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Visible
        Else
            'cmbcashcredit.Visible = True
            'lblCredit.Visible = True
            MyLabel2.Visible = True
            MyLabel2.Text = "Salesman"
            txtSalesman.Visible = True
            lblSalesman.Visible = True
            lblBoothStation.Visible = False
            Panel1.Visible = True
            lbl_ExFactoryDate.Visible = True
            txtEx_Factory_Date.Visible = True
            Panel3.Visible = True
            Panel4.Visible = False
            btn_QtyReset.Visible = False
            RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Collapsed
        End If
    End Sub
    Function IsValidPhoneNumber(phone As String) As Boolean
        ' Regex pattern for US phone numbers (you can adjust it as per the requirement)
        Dim pattern As String = "^\(?([0-9]{3})\)?[-.●]?([0-9]{3})[-.●]?([0-9]{4})$"
        Dim regex As New Regex(pattern)
        Return regex.IsMatch(phone)
    End Function
    Function AllowToSave(ByVal trans As SqlTransaction) As Boolean
        Try
            Xtra.TransactionValidity(txtDate.Value)
            'Sanjay Ticket No- ERO/12/07/18-000371
            If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso Is_Cancelled = 1 Then
                'If Is_Cancelled = 1 Then
                clsCommon.MyMessageBoxShow(Me, "Booking Cancelled,Can not Update", Me.Text)
                'Exit Function
                Return False
                'End If
            End If
            If clsCommon.myLen(txtMobileNo.Text) > 0 AndAlso Not IsValidPhoneNumber(txtMobileNo.Text) Then
                txtMobileNo.Focus()
                Throw New Exception("Please enter a valid 10-digit Mobile number.")
            End If

            If ServerDateTimeForTaxableInvoice Then
                Dim serverDateTime As DateTime = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MMM/yyyy")
                If clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") <> serverDateTime Then
                    Throw New Exception("Document Date and Supply Date Should be Server Date Time")
                End If
            End If

            'Sanjay Ticket No- ERO/12/07/18-000371
            If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer", Me.Text)
                txtVendorNo.Focus()
                Return False
            End If
            If clsCommon.myLen(txtVendorNo.Value) > 0 AndAlso Not chkDCS.Checked AndAlso clsCommon.myLen(strRoutecode) = 0 Then
                'If Not chkDCS.Checked Then
                'If clsCommon.myLen(strRoutecode) = 0 Then
                Throw New Exception("Please Map Route for customer ")
                blnSaveTotalQTy = False
                'Exit Function
                'End If
                'End If
            End If
            If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                If Not chkDCS.Checked AndAlso clsCommon.myLen(txtVehicleCode.Value) = 0 Then
                    'If clsCommon.myLen(txtVehicleCode.Value) = 0 Then
                    Throw New Exception("Please enter Vehicle ")
                    blnSaveTotalQTy = False
                    'Exit Function
                    'End If
                Else
                    If clsCommon.CompairString(cmbcashcredit.Text, "Cash") = CompairStringResult.Equal AndAlso clsCommon.myLen(txtReceipt.Value) = 0 Then
                        'If clsCommon.myLen(txtReceipt.Value) = 0 Then
                        Throw New Exception("Please Select Receipt No.")
                        blnSaveTotalQTy = False
                        'Exit Function
                        'End If
                    End If
                    If clsCommon.myLen(txtVehicleCode.Value) = 0 Then
                        Throw New Exception("Please enter Vehicle ")
                        blnSaveTotalQTy = False
                        'Exit Function
                    End If

                End If
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                If chkcashsale.Checked AndAlso clsCommon.myLen(txtChequeNo.Text) = 0 Then
                    Throw New Exception("Please Enter Reference/Cheque No.")
                    blnSaveTotalQTy = False

                End If
            End If
            If clsCommon.myLen(txtLocation.Value) = 0 Then
                Throw New Exception("Please enter Location ")
                'Exit Function
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(cmbGatePassType.Text, "Select") = CompairStringResult.Equal Then
                'If clsCommon.CompairString(cmbGatePassType.Text, "Select") = CompairStringResult.Equal Then
                Throw New Exception("Please Select GatePass Type")
                'Exit Function
                Return False
                'End If
            End If

            If AllowToCreateNoOfBookingPerDay > 0 AndAlso Not chkGatePass.Checked Then
                Dim STRSQL As String = "select count(distinct TSPL_BOOKING_MATSER.Document_No) as cc from TSPL_BOOKING_DETAIL left join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.cust_code=TSPL_BOOKING_DETAIL.cust_code where TSPL_BOOKING_MATSER.From_Screen_Code ='BOOK-DS-CU' and TSPL_BOOKING_MATSER.location_code ='" & txtLocation.Value & "' and TSPL_BOOKING_DETAIL.Cust_Code='" & txtVendorNo.Value & "' AND convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & txtDate.Value & "',103) and TSPL_BOOKING_MATSER.Document_No<>'" & clsCommon.myCstr(txtDocNo.Value) & "' and TSPL_BOOKING_MATSER.AgainstGatePass=0  "
                If ShowBookingTypeDropDownonDairyBookingCustomer Then
                    STRSQL += " And isnull(TSPL_CUSTOMER_MASTER.customer_category,'') not in ('Others','') and TSPL_BOOKING_MATSER.Booking_Type<>'CD' "
                End If
                Dim TempBookingExist As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(STRSQL, trans))
                If TempBookingExist >= AllowToCreateNoOfBookingPerDay Then
                    common.clsCommon.MyMessageBoxShow(Me, "Booking already exist for Date [" + clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") + "]", Me.Text)
                    'btnCreateAndPrintInvoice.Focus()
                    Return False
                End If
            End If
            If Not isNewEntry AndAlso clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Booking Not found to save")
                txtDocNo.Focus()
                Return False
            End If
            If ShowBookingTypeDropDownonDairyBookingCustomer AndAlso (clsCommon.CompairString(cmbcashcredit.Text, "Select") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbcashcredit.Text, "") = CompairStringResult.Equal) Then
                'If clsCommon.CompairString(cmbcashcredit.Text, "Select") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbcashcredit.Text, "") = CompairStringResult.Equal Then
                Throw New Exception("Please Select Booking Type ")
                'End If
            End If
            ''richa agarwal VIJ/28/11/19-000081 
            If EnableCustomerPODetailonDairyBooking = 1 AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_CUSTOMER_GROUP_MASTER.PONOMandatory  from TSPL_CUSTOMER_MASTER left outer join TSPL_CUSTOMER_GROUP_MASTER on TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code =TSPL_CUSTOMER_MASTER.Cust_Group_Code where TSPL_CUSTOMER_MASTER.cust_code='" & txtVendorNo.Value & "'", trans)), "1") = CompairStringResult.Equal Then
                'If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_CUSTOMER_GROUP_MASTER.PONOMandatory  from TSPL_CUSTOMER_MASTER left outer join TSPL_CUSTOMER_GROUP_MASTER on TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code =TSPL_CUSTOMER_MASTER.Cust_Group_Code where TSPL_CUSTOMER_MASTER.cust_code='" & txtVendorNo.Value & "'", trans)), "1") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtPONo.Text) = 0 Then
                    Throw New Exception("Please enter Customer PO No ")
                End If
                If txtCustPODate.Checked = False Then
                    Throw New Exception("Please select Customer PO Date ")
                End If
                'End If
            End If
            If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'", trans)), "Y") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "Others") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "") = CompairStringResult.Equal Then
                        If clsCommon.myLen(txtSubLocation.Value) <= 0 Then
                            Throw New Exception("Please select Sub Location")
                            txtSubLocation.Focus()
                            Return False
                        End If
                    End If
                End If
            End If

            If ApplyEWBThresholdLimit AndAlso rbtnTaxable.IsChecked Then

                Dim EWBThresholdLimtCat As String = ""
                Dim EWBThresholdLimtCatForCust As String = ""
                chkIsEwayBill.Checked = False
                Dim EwbNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EWayBillNo from TSPL_SD_SALE_INVOICE_HEAD where Document_Code in(select Sale_Invoice_No from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No ='" & txtDocNo.Value & "')"))
                If clsCommon.myLen(EwbNo) = 0 Then
                    If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "") = CompairStringResult.Equal Then
                                If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                                    EWBThresholdLimtCat = "select City_Code,State from TSPL_LOCATION_MASTER where Location_Code='" & txtSubLocation.Value & "'"

                                End If
                            End If
                        Else
                            EWBThresholdLimtCat = "select City_Code,State from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"
                        End If
                        EWBThresholdLimtCatForCust = "select City_Code,State from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"
                        Dim dt_FromLocation As DataTable = clsDBFuncationality.GetDataTable(EWBThresholdLimtCat)
                        Dim dt_CustLocation As DataTable = clsDBFuncationality.GetDataTable(EWBThresholdLimtCatForCust)
                        If clsCommon.CompairString(clsCommon.myCstr(dt_FromLocation(0)("City_Code")), clsCommon.myCstr(dt_CustLocation(0)("City_Code"))) = CompairStringResult.Equal Then
                            If clsCommon.myCdbl(lblTotRAmt1.Text) > EWBThresholdLimitForIntraCity Then
                                chkIsEwayBill.Checked = True
                            End If
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dt_FromLocation(0)("State")), clsCommon.myCstr(dt_CustLocation(0)("State"))) = CompairStringResult.Equal Then
                            If clsCommon.myCdbl(lblTotRAmt1.Text) > EWBThresholdLimitForIntraState Then
                                chkIsEwayBill.Checked = True
                            End If
                        Else
                            If clsCommon.myCdbl(lblTotRAmt1.Text) > EWBThresholdLimitForInterState Then
                                chkIsEwayBill.Checked = True
                            End If
                        End If
                    End If
                Else
                    chkIsEwayBill.Checked = True
                End If

            End If
            'If clsCommon.CompairString(cmbBookingType.SelectedValue, "FN") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbBookingType.SelectedValue, "PS") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbBookingType.SelectedValue, "UP") = CompairStringResult.Equal Then
            '    If chkGatePass.Checked = False Then
            '        Throw New Exception("Gate Pass checkbox should be checked. ")
            '    End If
            'End If
            ''richa 4 Aug,2021 optimization related
            If (RunBatchFifowisewithmodifyfunctionality AndAlso clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) <= 0) Then
                OpenBatchItem()
            ElseIf RunBatchFifowise = 1 AndAlso Not RunBatchFifowisewithmodifyfunctionality Then
                OpenBatchItem()
            End If
            Dim strBookingType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(BookingThrough,'')  from TSPL_BOOKING_MATSER   where Document_No ='" & txtDocNo.Value & "' ", trans))
            Dim strCustomercategory As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans))
            Dim dblQuantity As Double = 0
            For ii As Integer = 0 To gv1.Rows.Count - 1
                If (clsCommon.CompairString(gv1.Rows(ii).Cells(colICode).Value, Nothing) = CompairStringResult.Equal) Then
                    Continue For
                End If
                Dim taxGroup As String = txtTaxGroup.Value
                Dim strICode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                Dim strIName As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                Dim dblQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colQty).Value)
                Dim AvgQty As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(ColAvgQty).Value)
                Dim dblrate As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colRate).Value)
                Dim dblMRP As Double = clsCommon.myCdbl(gv1.Rows(ii).Cells(colMRP).Value)
                Dim strUOM As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                Dim strRemarks As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colRemarks).Value)
                Dim MainItemLen As Integer = clsCommon.myLen(gv1.Rows(ii).Cells(ColMainItem).Value)
                'Dim isBatchWise As Boolean = clsItemMaster.IsBatchItem(clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value), trans)
                'If isNewEntry = False AndAlso clsCommon.myLen(strICode) > 0 Then
                '    If RunBatchFifowisewithmodifyfunctionality = True AndAlso isBatchWise = True Then
                '        common.clsCommon.MyMessageBoxShow(Me, "Item " + strIName + " should be Batch Wise. At Line No" + clsCommon.myCstr(ii + 1))
                '        Return False
                '    End If
                'End If
                dblQuantity = dblQuantity + dblQty
                If (clsCommon.myLen(strICode) > 0) Then
                    If clsCommon.myLen(strUOM) <= 0 Then
                        common.clsCommon.MyMessageBoxShow(Me, "Please enter Booked Quantity UOM for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                        Return False
                    End If
                    If clsCommon.myLen(gv1.Rows(ii).Cells(colBatchNo).Value) <= 0 AndAlso dblQty > 0 AndAlso clsCommon.myCBool(gv1.Rows(ii).Cells(colIsBatchItem).Value) AndAlso clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value, trans) Then
                        'If dblQty > 0 AndAlso clsCommon.myCBool(gv1.Rows(ii).Cells(colIsBatchItem).Value) AndAlso clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value, trans) = True Then
                        'If RunBatchFifowise = 1 Then
                        '    gv1.CurrentRow = gv1.Rows(ii)
                        '    OpenBatchItem()
                        'End If
                        Dim arrBatchNo As List(Of clsBatchInventory) = TryCast(gv1.Rows(ii).Cells(colICode).Tag, List(Of clsBatchInventory))
                        If arrBatchNo Is Nothing Then
                            Throw New Exception("Please provide Batch no for item : " & strICode & " . At Line No" & clsCommon.myCstr(ii + 1))
                        Else
                            Dim tQty As Decimal = 0
                            For Each objBatch As clsBatchInventory In arrBatchNo
                                tQty += objBatch.Qty
                            Next
                            If tQty <> dblQty Then
                                Throw New Exception("Item : " & strICode & " Entered Qty " & clsCommon.myCstr(dblQty) & Environment.NewLine & "And Batchwise Qty " & clsCommon.myCstr(tQty) + " . At Line No" + clsCommon.myCstr(ii + 1))
                            End If
                        End If
                        'End If
                    End If

                    If ShowBookingTypeDropDownonDairyBookingCustomer AndAlso Not chkSampling.Checked AndAlso dblQty > 0 AndAlso dblrate <= 0 Then
                        'If chkSampling.Checked = False Then
                        'If dblQty > 0 AndAlso dblrate <= 0 Then
                        Throw New Exception("Please enter Booked Rate for " & strIName & ". At Line No" & clsCommon.myCstr(ii + 1))
                        Return False
                        'End If
                        'End If
                    Else
                        If chkSampling.Checked Then
                            If dblQty <= 0 Then
                                Throw New Exception("Please enter Booked Quantity for " & strIName & ". At Line No" & clsCommon.myCstr(ii + 1))
                                Return False
                            End If
                            'If dblrate <= 0 Then
                            '    clsCommon.MyMessageBoxShow(Me, "Please enter Booked Rate for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1))
                            '    Return False
                            'End If
                        Else
                            If dblQty <= 0 Then
                                Throw New Exception("Please enter Booked Quantity for " & strIName & ". At Line No" & clsCommon.myCstr(ii + 1))
                                Return False
                            End If
                            If dblrate <= 0 Then
                                Throw New Exception("Please enter Booked Rate for " & strIName & ". At Line No" & clsCommon.myCstr(ii + 1))
                                Return False
                            End If
                        End If

                    End If
                    'Sanjay Ticket No- BHA/28/06/18-000106 Client- Bharat Dairy, Setting for check Average Quantity ''richa this check will not be worked for those booking which are created through Mobile app ERO/27/08/19-001005
                    ' If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(BookingThrough,'')  from TSPL_BOOKING_MATSER   where Document_No ='" & txtDocNo.Value & "' ", trans)), "App") <> CompairStringResult.Equal Then
                    If clsCommon.CompairString(strBookingType, "App") <> CompairStringResult.Equal Then
                        If CheckAvgQtyOnDairyBooking = True Then
                            If (FlagCreateDo = False) And (FlagFirstRecord = False) Then 'And (AvgQty > 0)
                                ''  ERO/29/07/19-000971 by Balwinder on 30/07/2019
                                If Math.Abs(dblQty - AvgQty) > SettDairyBookingTolleranceQty Then
                                    If common.clsCommon.MyMessageBoxShow(Me, "Booking Quantity Should be in Average Quantity [" + clsCommon.myCstr(IIf(AvgQty - SettDairyBookingTolleranceQty < 0, 0, (AvgQty - SettDairyBookingTolleranceQty))) + " - " + clsCommon.myCstr(AvgQty + SettDairyBookingTolleranceQty) + " ] for " + strIName + ". At Line No" + clsCommon.myCstr(ii + 1) + vbNewLine + " Do you want to continue? ", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                                        Return False
                                    Else
                                        Dim frm As New FrmFreeTxtBox1
                                        frm.Text = "Remarks"
                                        frm.strRmks = strRemarks
                                        frm.ShowDialog()
                                        gv1.Rows(ii).Cells(colRemarks).Value = frm.strRmks
                                    End If
                                End If

                            End If
                        End If
                    End If
                End If
                Dim dblBalQty As Double = 0
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'", trans)), "Y") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "Others") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "") = CompairStringResult.Equal Then
                        dblBalQty = clsItemLocationDetails.getBalance(strICode, IIf(clsCommon.myLen(clsCommon.myCstr(txtSubLocation.Value)) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, trans, strUOM, dblMRP)
                    Else
                        dblBalQty = clsItemLocationDetails.getBalance(strICode, IIf(clsCommon.myLen(clsCommon.myCstr(txtSubLocation.Value)) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, trans, strUOM, dblMRP)
                    End If
                Else
                    dblBalQty = clsItemLocationDetails.getBalance(strICode, IIf(clsCommon.myLen(clsCommon.myCstr(txtSubLocation.Value)) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, trans, strUOM, dblMRP)
                End If
                'Dim dblBalQty As Double = clsItemLocationDetails.getBalance(strICode, IIf(clsCommon.myLen(clsCommon.myCstr(txtSubLocation.Value)) > 0, txtSubLocation.Value, txtBillToLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, strUOM, dblMRP)
                Dim dblEnteredQty As Double = dblQty
                If (clsCommon.myLen(strICode) > 0) Then
                    For jj As Integer = 0 To gv1.Rows.Count - 1
                        If clsCommon.myLen(gv1.Rows(jj).Cells(colICode).Value) > 0 AndAlso clsCommon.CompairString(taxGroup, clsCommon.myCstr(gv1.Rows(jj).Cells(colTaxGroup).Value)) <> CompairStringResult.Equal Then
                            'If clsCommon.CompairString(taxGroup, clsCommon.myCstr(gv1.Rows(jj).Cells(colTaxGroup).Value)) <> CompairStringResult.Equal Then
                            Throw New Exception("Tax group not matched at Line no [" + clsCommon.myCstr(ii) + "]")
                            'End If
                        End If
                        If jj = ii Then
                            Continue For
                        End If
                        Dim strInnerICode As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                        Dim strInnerUOM As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value)
                        Dim dblQtyInner As String = clsCommon.myCdbl(gv1.Rows(jj).Cells(colQty).Value)
                        If dblQtyInner > 0 AndAlso clsCommon.CompairString(strInnerICode, strICode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strInnerUOM, strUOM) = CompairStringResult.Equal Then
                            dblEnteredQty += dblQtyInner
                        End If
                        If clsCommon.CompairString(strICode, strInnerICode) = CompairStringResult.Equal AndAlso MainItemLen = 0 Then
                            ' If ShowMulMRPOfSameItemOnDairyBookingCustomer = True AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "Others") = CompairStringResult.Equal Then
                            If ShowMulMRPOfSameItemOnDairyBookingCustomer = True AndAlso clsCommon.CompairString(strCustomercategory, "Others") = CompairStringResult.Equal Then
                                Dim dblInnerRate As Double = clsCommon.myCdbl(gv1.Rows(jj).Cells(colRate).Value)
                                If clsCommon.CompairString(dblrate, dblInnerRate) = CompairStringResult.Equal Then
                                    Dim Msg As String = "Same Item Rate at Row No " + clsCommon.myCstr(ii + 1) + " And " + clsCommon.myCstr(jj + 1)
                                    common.clsCommon.MyMessageBoxShow(Me, Msg, Me.Text)
                                    Return False
                                End If
                            Else
                                Dim Msg As String = ""
                                If ApplyManualScheme Then
                                    If clsCommon.myLen(gv1.Rows(jj).Cells(ColMainItem).Value) <= 0 Then
                                        Msg = "Same Item Exist at Row No " + clsCommon.myCstr(ii + 1) & " And " & clsCommon.myCstr(jj + 1)
                                        common.clsCommon.MyMessageBoxShow(Me, Msg, Me.Text)
                                        Return False
                                    End If
                                Else
                                    Msg = "Same Item Exist at Row No " + clsCommon.myCstr(ii + 1) & " And " & clsCommon.myCstr(jj + 1)
                                    common.clsCommon.MyMessageBoxShow(Me, Msg, Me.Text)
                                    Return False
                                End If


                            End If
                        End If
                    Next
                    dblEnteredQty = Math.Round(dblEnteredQty, 2, MidpointRounding.ToEven)
                    If dblEnteredQty > dblBalQty Then ' AndAlso clsCommon.CompairString(strSchemeItem, "No") = CompairStringResult.Equal Then
                        Throw New Exception("Item - " + strICode + Environment.NewLine + "Entered Quantity - " + clsCommon.myCstr(dblEnteredQty) + " and Balance Quantity - " + clsCommon.myCstr(dblBalQty))
                        Return False
                    End If

                End If
            Next


            If ShowBookingTypeDropDownonDairyBookingCustomer AndAlso dblQuantity <= 0 AndAlso AllowZeroQtyOnDairyBooking Then
                'If dblQuantity <= 0 AndAlso AllowZeroQtyOnDairyBooking Then
                Throw New Exception("Please enter Qunatity at least in one row")
                Return False
                'End If
            End If
            '' 18 SEP,2020 NOT CHECKED CONDITION FOR CUSTOMER CATEGORY OTHERS AND NULL
            If PrintTruckSheetAfterGenerate AndAlso (clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "Others") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "") <> CompairStringResult.Equal) Then
                Dim isTruckSheetGenerated As Boolean = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select max(truckSheetGenerate) from TSPL_BOOKING_MATSER where convert(date,Document_Date ,103)=convert(date,'" & txtDate.Value & "',103)", trans)) = 1, True, False)
                If isTruckSheetGenerated AndAlso Not chkGatePass.Checked Then
                    Throw New Exception("Truck sheet has been generated for date " & txtDate.Value & "..Please create Gate Pass. ")
                    Return False
                End If
            End If
            If chkGatePass.Checked AndAlso (clsCommon.CompairString(cmbGatePassType.Text, "Select") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbGatePassType.Text, "") = CompairStringResult.Equal) Then
                Throw New Exception("Please select Gate Pass type(AM/PM)")
                cmbGatePassType.Focus()
                Return False
            End If
            If chkBPL.Checked Then
                If clsCommon.myLen(txtCouponCode.Text) <= 0 Then
                    Throw New Exception("Please enter Coupon Code")
                    txtCouponCode.Focus()
                    Return False
                End If
                'If clsCommon.myLen(txtBPLName.Text) <= 0 Then
                '    Throw New Exception("Please enter Name")
                '    txtCouponCode.Focus()
                '    Return False
                'End If
                UcAttachment1.AllowToSave()
            End If

            If isRCDFRateControl Then
                If gv1 IsNot Nothing AndAlso gv1.Rows.Count > 0 Then
                    Dim strItem As String = ""
                    Dim isCheck As Boolean = False
                    For Each grow As GridViewRowInfo In gv1.Rows
                        If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colICode).Value)) > 0 Then
                            Dim Qry As String = "Select TSPL_RCDF_RATE_CONTROL_DETAIL_ALL_UOM.Code,TSPL_RCDF_RATE_CONTROL_DETAIL.Item_Code,TSPL_RCDF_RATE_CONTROL_DETAIL_ALL_UOM.UOM,TSPL_RCDF_RATE_CONTROL_DETAIL_ALL_UOM.Min_Rate,TSPL_RCDF_RATE_CONTROL_DETAIL_ALL_UOM.Max_Rate from TSPL_RCDF_RATE_CONTROL_DETAIL_ALL_UOM
                                        Inner Join TSPL_RCDF_RATE_CONTROL_DETAIL On TSPL_RCDF_RATE_CONTROL_DETAIL.PK_Id=TSPL_RCDF_RATE_CONTROL_DETAIL_ALL_UOM.Against_PK_Id
                                        Where TSPL_RCDF_RATE_CONTROL_DETAIL.Item_Code='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "'"
                            Dim dtt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                            If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
                                For Each rows As DataRow In dtt.Rows
                                    If clsCommon.CompairString(clsCommon.myCstr(rows("Item_Code")), clsCommon.myCstr(grow.Cells(colICode).Value)) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(rows("UOM")), clsCommon.myCstr(grow.Cells(colUnit).Value)) = CompairStringResult.Equal Then
                                        If clsCommon.myCDecimal(grow.Cells(colRate).Value) < clsCommon.myCDecimal(rows("Min_Rate")) OrElse clsCommon.myCDecimal(grow.Cells(colRate).Value) > clsCommon.myCDecimal(rows("Max_Rate")) Then
                                            strItem += "Item : " + clsCommon.myCstr(grow.Cells(colICode).Value) + " " + clsCommon.myCstr(grow.Cells(colIName).Value) + " " + Environment.NewLine
                                            strItem += "Unit Cost : " + clsCommon.myCstr(grow.Cells(colRate).Value) + " " + Environment.NewLine
                                            strItem += "According to RCDF Rate Control unit cost should be in range (" + clsCommon.myCstr(rows("Min_Rate")) + " to " + clsCommon.myCstr(rows("Max_Rate")) + ") . " + Environment.NewLine
                                            isCheck = True
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                    If isCheck Then
                        clsCommon.MyMessageBoxShow(Me, strItem, Me.Text)
                        Return False
                    End If
                    isCheck = False
                End If
            End If
            For ii As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    UpdateCurrentRow1(ii)
                Else
                    UpdateCurrentRow(ii)
                End If
            Next
            UpdateAllTotals()

            If CheckCustomeroutStandingAmt Then
                If Not clsCustomerMaster.IsCreditCustomer(txtVendorNo.Value) Then
                    Dim custOutStanding As Double = clsCommon.myCdbl(lblOutstandingDesc.Text)
                    If custOutStanding <= 0 Then
                        Throw New Exception("Insufficient Balance.")
                    End If
                    custOutStanding = Math.Abs(custOutStanding)
                    Dim TotalDocAmt As Double = clsCommon.myCdbl(lblTotRAmt.Text)
                    If TotalDocAmt > custOutStanding Then
                        Throw New Exception("Insufficient Balance.")
                    End If
                End If
            End If
            'Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            'Return False
        End Try
        Return True
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        RecordCount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_BOOKING_MATSER "))
        If RecordCount = 0 Then
            FlagFirstRecord = True
        Else
            FlagFirstRecord = False
        End If
        SavingData(False)

    End Sub
    'Private Sub UpdateExcelDoc()
    '    Dim tran As SqlTransaction = clsDBFuncationality.GetTransactin()
    '    Try
    '        Dim gv As New UserControls.MyRadGridView
    '        Me.Controls.Add(gv)
    '        If transportSql.importExcel(gv, "DocumentNo", "Remarks", "DocDate") Then

    '            Try
    '                clsCommon.ProgressBarShow()
    '                For Each grow As GridViewRowInfo In gv.Rows
    '                    Dim DocNo As String = clsCommon.myCstr(grow.Cells("DocumentNo").Value)
    '                    Dim Remarks As String = clsCommon.myCstr(grow.Cells("Remarks").Value)
    '                    Dim DocDate As DateTime = clsCommon.myCDate(grow.Cells("DocDate").Value)
    '                    Dim strqry As String = "update TSPL_BOOKING_MATSER set Document_Date='" & clsCommon.GetPrintDate(DocDate, "dd/MMM/yyyy") & "',Description='" & Remarks & "' where Document_No='" & DocNo & "' "
    '                    clsDBFuncationality.ExecuteNonQuery(strqry, tran)
    '                    Dim invoiceno As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Sale_Invoice_No from tspl_sd_shipment_head where Against_Booking_No='" & DocNo & "'", tran))
    '                    Dim ShipCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from tspl_sd_shipment_head where Against_Booking_No='" & DocNo & "'", tran))
    '                    strqry = "update tspl_sd_shipment_head set Document_Date='" & clsCommon.GetPrintDate(DocDate, "dd/MMM/yyyy") & "',Description='" & Remarks & "' where Document_Code='" & ShipCode & "' "
    '                    clsDBFuncationality.ExecuteNonQuery(strqry, tran)
    '                    strqry = "update TSPL_SD_SALE_INVOICE_HEAD set Document_Date='" & clsCommon.GetPrintDate(DocDate, "dd/MMM/yyyy") & "',Description='" & Remarks & "' where Document_Code='" & invoiceno & "' "
    '                    clsDBFuncationality.ExecuteNonQuery(strqry, tran)
    '                    'tran.Commit()
    '                Next
    '                tran.Commit()
    '                clsCommon.ProgressBarHide()
    '                clsCommon.MyMessageBoxShow(Me, "Update Successfully", Me.Text)
    '            Catch ex As Exception

    '                tran.Rollback()
    '                clsCommon.ProgressBarHide()
    '                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '            End Try
    '        End If

    '    Catch ex As Exception

    '    End Try
    'End Sub
    Sub SavingData(ByVal ChekBtnPost As Boolean)
        If SaveData() AndAlso Not ChekBtnPost Then
            'If ChekBtnPost = False Then
            common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
            btnAddNew.Focus()
            'End If
        End If
    End Sub
    Private Function SaveData() As Boolean
        Try
            Dim qry As String = ""
            blnSaveTotalQTy = True
            BookingStatus = 0
            'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
            If (AllowToSave(Nothing)) Then
                Dim obj As New clsBookingEntryDairySale()
                ' done by priti ERO/12/07/18-000373 for sampling booking
                obj.IsSampling = IIf(chkSampling.Checked, 1, 0)
                obj.Document_No = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Supply_Date = txtSupplyDate.Value
                obj.location_code = txtLocation.Value
                obj.AdharNo = txtAdharNo.Text
                obj.MobileNO = txtMobileNo.Text

                obj.Is_CashSale = IIf(chkcashsale.Checked, "Y", "N")
                obj.IsEwaybill = IIf(chkIsEwayBill.Checked, 1, 0)
                obj.No_Transporter = IIf(chkNoTranspoter.Checked, 1, 0)
                If chkcashsale.Checked Then
                    obj.Payment_Terms = cmbPaymentType.Text
                    obj.ChequeNo = txtChequeNo.Text
                    'Else
                    '    obj.Payment_Terms = "CREDIT"
                End If
                obj.ReceiverName = txtReceiverName.Text
                If clsCommon.myLen(txtReceipt.Value) > 0 Then
                    obj.Against_Receipt_No = txtReceipt.Value
                End If
                'obj.Cust_Group_Code = txtCustGrp.Value
                Dim isDemandBooking1 = clsDBFuncationality.getSingleValue("select top 1 Against_DemandBooking_No from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'")
                If clsCommon.myLen(isDemandBooking1) > 0 Then
                    rgbItemType.Visible = False
                Else
                    rgbItemType.Visible = True
                    If rbtnTaxable.IsChecked Then
                        obj.Is_Taxable = 2
                    ElseIf rbtnNonTax.IsChecked Then
                        obj.Is_Taxable = 1
                    End If
                End If
                ' 2 for Taxable and NonTaxable item in a single booking
                If CreateCommonDairyDispatchforFreshAmbient = 1 Then
                    obj.TRANSACTION_TYPE = ""
                Else
                    obj.TRANSACTION_TYPE = IIf(rbtn_Fresh.IsChecked, "FS", "PS")
                End If
                If chkExcludeKKFMND.Checked Then
                    obj.Exclude_KKF_And_Mandi = 1
                Else
                    obj.Exclude_KKF_And_Mandi = 0
                End If
                obj.From_Screen_code = clsUserMgtCode.frmDairyBookingCustomer
                If EnableCustomerPODetailonDairyBooking = 1 Then
                    obj.SalesmanCode = txtSalesman.Value
                    obj.Cust_PO_No = txtPONo.Text
                    If txtCustPODate.Checked Then
                        obj.Podate = txtCustPODate.Value
                    End If
                End If
                If clsCommon.myLen(lblUploadingDate.Text) > 0 Then
                    obj.Uploading_date = lblUploadingDate.Text
                End If
                If clsCommon.myLen(txtIndentNo.Text) > 0 Then
                    obj.PO_Indent_No = txtIndentNo.Text
                    obj.PO_Indent_Date = txtIndentDate.Value
                End If
                obj.TotalCAN = txtCan.Text
                If chkNoCrateIssue.Checked Then
                    obj.TotalCrate = 0
                    obj.NoCrateIssue = 1
                Else
                    obj.TotalCrate = txtCrate.Text
                    obj.NoCrateIssue = 0
                End If
                obj.TotalBox = txtBox.Text
                obj.RoundOffAmount = clsCommon.myCdbl(TxtRoundoff.Text)
                If txtEx_Factory_Date.Checked Then
                    obj.Ex_Factory_Date = txtEx_Factory_Date.Value
                End If
                If clsCommon.CompairString(cmbcashcredit.Text, "Select") = CompairStringResult.Equal Then
                    obj.Booking_Type = ""
                ElseIf clsCommon.CompairString(cmbcashcredit.Text, "PARLOR SALES") = CompairStringResult.Equal Then
                    obj.Booking_Type = "PS"
                ElseIf clsCommon.CompairString(cmbcashcredit.Text, "FORENOON") = CompairStringResult.Equal Then
                    obj.Booking_Type = "FN"
                ElseIf clsCommon.CompairString(cmbcashcredit.Text, "UP COUNTRY") = CompairStringResult.Equal Then
                    obj.Booking_Type = "UP"
                Else
                    obj.Booking_Type = cmbcashcredit.Text
                End If
                If chkDCS.Checked Then
                    obj.Is_DCS = 1
                    obj.Against_DCSBooking_No = txtDCSDemandNo.Text
                    obj.Booking_Type = cmbcashcredit.Text
                    If clsCommon.myLen(txtLastCollectionDate.Text) > 0 Then
                        obj.LastCollectionDate = txtLastCollectionDate.Text
                    Else
                        obj.LastCollectionDate = Nothing
                    End If
                Else
                    obj.Booking_Type = ""
                    obj.LastCollectionDate = Nothing
                End If
                If clsCommon.CompairString(cmbGatePassType.Text, "Select") = CompairStringResult.Equal Then
                    obj.GatePass_Type = ""
                Else
                    obj.GatePass_Type = clsCommon.myCstr(cmbGatePassType.Text)
                End If
                ''  obj.Booking_Type = IIf(cmbBookingType.Text = "Select", "", cmbBookingType.SelectedValue)
                obj.AgainstGatePass = IIf(chkGatePass.Checked, 1, 0)
                obj.IS_TCS = IIf(chkisTCS.Checked, 1, 0)
                obj.Ship_To_Location = txtShipToLocation.Value
                If clsCommon.myLen(lblLoginUserZone.Text) > 0 Then
                    obj.Login_User_Zone_Code = lblLoginUserZone.Text
                End If
                obj.Manual_VehicleNo = txtManualVehicle.Text
                obj.Is_Manual_Vehicle = IIf(chkManualVehicle.Checked, "Y", "N")
                obj.Transport_Id = fndTransporter.Value
                obj.Description = txtDescription.Text
                obj.FAT_Per = clsCommon.myCdbl(txtFATPER.Text)
                obj.SNF_Per = clsCommon.myCdbl(txtSNFPER.Text)
                obj.Acidity = clsCommon.myCdbl(txtAcidity.Text)
                obj.Temperature = clsCommon.myCdbl(txtTemp.Text)
                obj.MBRT_Hours = clsCommon.myCdbl(txtMBRTHours.Text)
                obj.TCSAmount = clsCommon.myCdbl(lblTCSAmount.Text)
                obj.Sub_Location_code = txtSubLocation.Value
                obj.Total_Amt = clsCommon.myCdbl(lblTotalDocAmt.Text)
                obj.Is_Credit_Customer = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Credit_Customer from TSPL_CUSTOMER_MASTER where Cust_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "'"))
                obj.Arr = New List(Of clsBookingDetailDairySale)
                ''richa 4 Aug,2021 optimization related
                Dim dblBooking_Status As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select  Booking_Status from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'  and Cust_Code='" & txtVendorNo.Value & "'"))
                Dim intLine As Integer = 0
                Dim DCTotalAmt As Double = 0
                Dim TCTotalAmt As Double = 0
                Dim SCTotalAmt As Double = 0
                For Each grow As GridViewRowInfo In gv1.Rows
                    Dim objTr As New clsBookingDetailDairySale()
                    objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))

                    objTr.Booking_Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
                    If (objTr.Booking_Qty) > 0 OrElse (AllowZeroQtyOnDairyBooking AndAlso clsCommon.myCdbl(lblTotRAmt1.Text) <= 0 AndAlso clsCommon.myLen(clsCommon.myCstr(grow.Cells(colICode).Value)) > 0) Then
                        intLine += 1
                        objTr.Line_No = grow.Cells(colLineNo).Value
                        'If grow.Cells(colIsKKF).Value = True Then
                        '    objTr.IsKKFTax = "YES"
                        'Else
                        '    objTr.IsKKFTax = "NO"
                        'End If
                        'If grow.Cells(colIsMNDTax).Value = True Then
                        '    objTr.IsMNDTax = "YES"
                        'Else
                        '    objTr.IsMNDTax = "NO"
                        'End If
                        'Dim objBookingCust As clsBookingTemp = TryCast(gv1.Columns(ii).Tag, clsBookingTemp)
                        objTr.Cust_Code = txtVendorNo.Value
                        '================Added by preeti gupta Against Ticket no[BHA/01/08/18-000206]
                        objTr.Route_No = lblroutecode.Text
                        '===========================================================
                        objTr.Sampling = 0
                        objTr.Loc_Code = txtLocation.Value
                        objTr.Total_Qty = TotalQuantity
                        'Dim objBookingitem As clsBookingTemp = TryCast(grow.Cells(ii).Tag, clsBookingTemp)
                        objTr.Vehicle_Code = txtVehicleCode.Value
                        objTr.Item_Code = clsCommon.myCstr(grow.Cells(colICode).Value)
                        objTr.Scheme_Main_Item = clsCommon.myCstr(grow.Cells(ColMainItem).Value)
                        objTr.Short_Description = ""  'clsCommon.myCstr(grow.Cells(colIShortName).Value)
                        objTr.Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                        objTr.PreviousBookingQty = clsCommon.myCdbl(grow.Cells(colPreviousQty).Value)
                        objTr.Tax_On_Amount = clsCommon.myCdbl(grow.Cells(colTBaseAmt).Value)
                        objTr.Tax_Amount = clsCommon.myCdbl(grow.Cells(colTTaxAmt).Value)
                        If ConvertIntoBillingUOM Then
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
                            objTr.Billing_Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
                        End If
                        If Not isNewEntry AndAlso clsCommon.myLen(isDemandBooking1) > 0 Then
                            objTr.Against_DemandBooking_No = clsDBFuncationality.getSingleValue("select top 1 Against_DemandBooking_No from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'")
                            objTr.Against_DemandBooking_TR_Code = clsDBFuncationality.getSingleValue("select Against_DemandBooking_TR_Code from TSPL_BOOKING_DETAIL  where Document_No='" & txtDocNo.Value & "' and Line_No=" + clsCommon.myCstr(objTr.Line_No) + "")
                            obj.Against_DemandBooking_No = objTr.Against_DemandBooking_No
                        End If
                        objTr.Item_Rate = clsCommon.myCDecimal(grow.Cells(colRate).Value)
                        objTr.Disc_Scheme_Amount = clsCommon.myCdbl(grow.Cells(colDisc_Scheme_Amount).Value)
                        objTr.Disc_Scheme_Code = clsCommon.myCstr(grow.Cells(colDisc_Scheme_Code).Value)
                        objTr.Disc_Scheme_Pers = clsCommon.myCdbl(grow.Cells(colDisc_Scheme_Pers).Value)
                        objTr.SchemeType = clsCommon.myCstr(grow.Cells(colSchemeType).Value)
                        objTr.OrgRate = clsCommon.myCdbl(grow.Cells(colOrgRate).Value)
                        objTr.SellingPrice = clsCommon.myCdbl(grow.Cells(colSellingRate).Value)
                        objTr.Tax_NonTax = clsCommon.myCdbl(grow.Cells(colTax_NonTax).Value)
                        objTr.FreshAmbient = clsCommon.myCstr(grow.Cells(colFreshAmbient).Value)
                        'Sanjay Ticket No- ERO/12/07/18-000371
                        objTr.Remarks = clsCommon.myCstr(grow.Cells(colRemarks).Value)
                        'Sanjay Ticket No- ERO/12/07/18-000371
                        objTr.Price_with_Tax = clsCommon.myCdbl(grow.Cells(colItemBasicPrice).Value)
                        objTr.Amount_with_Tax = clsCommon.myCdbl(grow.Cells(colAmountWithTax).Value)
                        objTr.Item_Price_ID = clsCommon.myCdbl(grow.Cells(colPriceId).Value)
                        If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colPriceIDAppDate).Value)) > 0 Then
                            objTr.Price_IdStartDate = clsCommon.myCDate(grow.Cells(colPriceIDAppDate).Value)
                        End If
                        objTr.PricePlanNo = clsCommon.myCstr(grow.Cells(colPricePlanNo).Value)
                        If BookingStatus = 0 Then
                            objTr.Booking_Status = 1
                        End If
                        If clsCommon.myLen(txtDocNo.Value) > 0 Then
                            ' BookingStatus = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select  Booking_Status from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'  and Cust_Code='" & objTr.Cust_Code & "'"))
                            BookingStatus = dblBooking_Status
                        End If
                        If BookingStatus = 3 Then
                            objTr.Booking_Status = 3
                        ElseIf BookingStatus = 5 Then
                            objTr.Booking_Status = 5
                            '===========approval status will be not change after update====
                        ElseIf BookingStatus = 2 Then
                            objTr.Booking_Status = 2
                            '===============================================================
                        Else
                            objTr.Booking_Status = objTr.Booking_Status
                        End If
                        '''''''''
                        BookingStatus = objTr.Booking_Status
                        objTr.DocumentAmount = clsCommon.myCdbl(lblTotRAmt1.Text)
                        Dim DOCdateCurrent As Date? = Nothing
                        DOCdateCurrent = clsCommon.GETSERVERDATE()
                        ' Query to get scheme type of Item
                        ' Scheme Items disable 
                        'Dim qryScheme As String = "Select TSPL_SCHEME_MASTER_NEW.Scheme_Type from TSPL_SCHEME_MASTER_NEW left outer join TSPL_SCHEME_DETAIL_NEW on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_MASTER_NEW.Scheme_Code "
                        'qryScheme += " left outer join tspl_item_master on tspl_item_master.item_code=TSPL_SCHEME_DETAIL_NEW.Item_Code "
                        'qryScheme += " where TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select scheme_code from (select ROW_NUMBER() over (partition by scheme_type order by start_date desc) as sno, TSPL_SCHEME_MASTER_NEW.Scheme_Code from "
                        'qryScheme += " TSPL_SCHEME_MASTER_NEW where TSPL_SCHEME_MASTER_NEW.Scheme_Type='Quantitive' and TSPL_SCHEME_MASTER_NEW.Status='Active' and  TSPL_SCHEME_MASTER_NEW.Start_Date<='" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  and (TSPL_SCHEME_MASTER_NEW.End_Date >= '" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  or TSPL_SCHEME_MASTER_NEW.End_date is null) and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select TSPL_SCHEME_DETAIL_NEW.Scheme_Code from "
                        'qryScheme += " TSPL_SCHEME_DETAIL_NEW left outer join TSPL_SCHEME_BENEFICIARY on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code where MainItem_Code='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "' and Cust_Code='" & txtVendorNo.Value & "'))a where a.sno=1)"
                        'qryScheme += " and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select Scheme_Code from TSPL_SCHEME_BENEFICIARY where Cust_Code='" & objTr.Cust_Code & "') and TSPL_SCHEME_MASTER_NEW.Status='Active'"
                        'qryScheme += " and TSPL_SCHEME_DETAIL_NEW.MainItem_Code='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "' "
                        'qryScheme += " order by TSPL_SCHEME_MASTER_NEW.Scheme_Code"
                        'Dim SchemeType As String = clsDBFuncationality.getSingleValue(qryScheme)
                        'If clsCommon.myLen(clsCommon.myCstr(SchemeType)) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(SchemeType), "Quantitive") = CompairStringResult.Equal Then
                        '    Dim objD As clsSchemeApplyOnDairy = Nothing
                        '    objD = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(grow.Cells(colICode).Value), clsCommon.myCstr(grow.Cells(colUnit).Value), clsCommon.myCdbl(grow.Cells(colQty).Value), txtVendorNo.Value, clsCommon.myCstr(grow.Cells(colSchemeType).Value), Nothing, Nothing)
                        '    If objD IsNot Nothing AndAlso objD.Arr.Count > 0 Then
                        '        For Each objtrScheme As clsSchemeApplyOnDairy In objD.Arr
                        '            objTr.SchemeType = objtrScheme.schm_Type
                        '            objTr.Scheme_Item_Code = objtrScheme.Schm_Icode
                        '            objTr.Scheme_Qty = objtrScheme.Schm_Qty
                        '            objTr.Scheme_Item_UOM = objtrScheme.Schm_Item_Uom
                        '            objTr.Scheme_Code = objtrScheme.Schm_Code
                        '        Next
                        '    End If
                        'End If
                        'End of Scheme Type of Detail
                    End If
                    objTr.QtyinKg = clsCommon.myCdbl(grow.Cells(colQtyinKG).Value)
                    objTr.Disc_Amt = clsCommon.myCdbl(grow.Cells(colDisAmt).Value)
                    objTr.Amt_Less_Discount = clsCommon.myCdbl(grow.Cells(colAmtAfterDis).Value)
                    objTr.TAX_Group = grow.Cells(colTaxGroup).Value
                    objTr.TAX1 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(1))).Value
                    objTr.TAX1_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(1))).Value
                    objTr.TAX1_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(1))).Value
                    objTr.TAX1_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(1))).Value
                    objTr.TAX2 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(2))).Value
                    objTr.TAX2_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(2))).Value
                    objTr.TAX2_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(2))).Value
                    objTr.TAX2_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(2))).Value
                    objTr.TAX3 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(3))).Value
                    objTr.TAX3_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(3))).Value
                    objTr.TAX3_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(3))).Value
                    objTr.TAX3_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(3))).Value
                    objTr.TAX4 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(4))).Value
                    objTr.TAX4_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(4))).Value
                    objTr.TAX4_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(4))).Value
                    objTr.TAX4_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(4))).Value
                    objTr.TAX5 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(5))).Value
                    objTr.TAX5_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(5))).Value
                    objTr.TAX5_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(5))).Value
                    objTr.TAX5_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(5))).Value
                    objTr.TAX6 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(6))).Value
                    objTr.TAX6_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(6))).Value
                    objTr.TAX6_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(6))).Value
                    objTr.TAX6_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(6))).Value
                    objTr.TAX7 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(7))).Value
                    objTr.TAX7_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(7))).Value
                    objTr.TAX7_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(7))).Value
                    objTr.TAX7_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(7))).Value
                    objTr.TAX8 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(8))).Value
                    objTr.TAX8_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(8))).Value
                    objTr.TAX8_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(8))).Value
                    objTr.TAX8_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(8))).Value
                    objTr.TAX9 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(9))).Value
                    objTr.TAX9_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(9))).Value
                    objTr.TAX9_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(9))).Value
                    objTr.TAX9_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(9))).Value
                    objTr.TAX10 = grow.Cells(clsCommon.myCstr("colTax" & clsCommon.myCstr(10))).Value
                    objTr.TAX10_Base_Amt = grow.Cells(clsCommon.myCstr("colTax_Base_Amt" & clsCommon.myCstr(10))).Value
                    objTr.TAX10_Rate = grow.Cells(clsCommon.myCstr("colTax_Rate" & clsCommon.myCstr(10))).Value
                    objTr.TAX10_Amt = grow.Cells(clsCommon.myCstr("colTax_Amt" & clsCommon.myCstr(10))).Value
                    objTr.Distributor_Commission_PKID = clsCommon.myCstr(grow.Cells(ColDCPKID).Value)
                    objTr.Distributor_Commission_Rate = clsCommon.myCdbl(grow.Cells(ColDCRate).Value)
                    objTr.Distributor_Commission_Amt = clsCommon.myCdbl(grow.Cells(ColDCAmt).Value)
                    objTr.Transporter_Commission_Rate = clsCommon.myCdbl(grow.Cells(ColTCRate).Value)
                    objTr.Transporter_Commission_Amt = clsCommon.myCdbl(grow.Cells(ColTCAmt).Value)
                    objTr.Security_Amt = clsCommon.myCdbl(grow.Cells(ColSCAmt).Value)
                    objTr.Security_Rate = clsCommon.myCdbl(grow.Cells(ColSCRate).Value)
                    objTr.Security_Amt = clsCommon.myCdbl(grow.Cells(ColSCAmt).Value)
                    objTr.Batch_No = clsCommon.myCstr(grow.Cells(colBatchNo).Value)
                    DCTotalAmt += objTr.Distributor_Commission_Amt
                    TCTotalAmt += objTr.Transporter_Commission_Amt
                    SCTotalAmt += objTr.Security_Amt
                    objTr.Distributor_Commission_RateWithTax = clsCommon.myCdbl(grow.Cells(ColDCRateWithTax).Value)


                    'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + intLine)).Value = Nothing
                    'gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Base_Amt" + intLine)).Value = Nothing
                    '    gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Rate" + intLine)).Value = Nothing
                    '    gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax_Amt" + intLine)).Value = Nothing
                    If (clsCommon.myLen(objTr.Cust_Code) > 0) AndAlso (clsCommon.myLen(objTr.Item_Code) > 0) Then
                        obj.Arr.Add(objTr)
                    End If
                Next
                If (clsCommon.myCdbl(lblTotRAmt1.Text)) > 0 Then
                    If Not AllowWo_Outstanding AndAlso CheckOutstandingOnbooking = 1 AndAlso ((BookingStatus = 1 OrElse BookingStatus = 2) AndAlso (BookingStatus <> 5)) Then
                        'If CheckOutstandingOnbooking = 1 AndAlso ((BookingStatus = 1 OrElse BookingStatus = 2) AndAlso (BookingStatus <> 5)) Then
                        CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
                        'End If
                    End If
                End If
                obj.Distributor_Commission_TotalAmt = DCTotalAmt
                obj.Transporter_Commission_TotalAmt = TCTotalAmt
                obj.Security_TotalAmt = SCTotalAmt
                obj.Credit_Limit = clsCommon.myCdbl(lblCreditLimit.Text)
                obj.Advance_Security = clsCommon.myCdbl(lblAdvanceSecurity.Text)
                obj.Revese_Adv_Security = clsCommon.myCdbl(lblReverseAdvanceSec.Text)
                obj.AR_Credit_Security = clsCommon.myCdbl(lblARSecurity.Text)
                obj.Pending_Posted_DO = clsCommon.myCdbl(lblPendingDO.Text)
                obj.UnPostedDispatch = clsCommon.myCdbl(lblShortcloseDO.Text)
                obj.Ledger_Outstansing = clsCommon.myCdbl(lblLedgerOutstanding.Text)
                obj.Refund_Security = clsCommon.myCdbl(lblRefund.Text)
                obj.Reverse_Refund_Sec = clsCommon.myCdbl(lblReverseRefund.Text)
                obj.Total_Outstanding = clsCommon.myCdbl(lblTotalOutstansing.Text)
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Fill at list one Item", Me.Text)
                    Return False
                End If
                If chkBPL.Checked Then
                    obj.Is_BPL = 1
                    obj.BPL_Coupon_Code = txtCouponCode.Text
                    obj.BPL_Name = txtBPLName.Text
                    obj.BPL_Remark = txtBPLRemark.Text
                    obj.BPL_Coupon_Date = txtCouponDate.Value
                    obj.BPL_Category = txtCategory.Value
                Else
                    obj.Is_BPL = 0
                    obj.BPL_Coupon_Date = Nothing
                End If
                If chkAPS.Checked Then
                    obj.Is_APS = 1
                Else
                    obj.Is_APS = 0
                End If
                If chkGhee.Checked Then
                    obj.Is_GHEE = 1
                Else
                    obj.Is_GHEE = 0
                End If
                If chkDistributor.Checked Then
                    obj.Is_Distributor = 1
                Else
                    obj.Is_Distributor = 0
                End If
                obj.Tax_Group = txtTaxGroup.Value
                obj.TaxGroupName = lblTaxGrpName.Text
                If rbtnTaxCalAutomatic.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                ElseIf rbtnTaxCalManual.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                End If
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
                obj.Total_Tax_Amt = clsCommon.myCdbl(lblTaxAmt.Text)
                If (obj.SaveData(obj, isNewEntry)) Then
                    If Not isNewEntry AndAlso obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                        'If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                        For Each item As clsBookingDetailDairySale In obj.Arr
                            qry = "update TSPL_DEMAND_BOOKING_DETAIL set Qty=" + clsCommon.myCstr(item.Booking_Qty) + " where TR_Code='" & item.Against_DemandBooking_TR_Code & "'"
                            clsDBFuncationality.ExecuteNonQuery(qry)
                        Next
                        'End If
                    End If
                    'Dim intSampling As Integer = 0
                    'Dim dblQty As Double = 0
                    'Dim dblRate As Double = 0
                    'Dim dblAmount As Double = 0
                    'Dim dblTotal As Double = 0
                    If chkBPL.Checked Then
                        UcAttachment1.SaveData(obj.Document_No)
                    End If
                    qry = "Delete from TSPL_TRANSACTION_APPROVAL where Document_No='" & obj.Document_No & "' "
                    clsDBFuncationality.ExecuteNonQuery(qry)
                    'For ii As Integer = 8 To gv1.Columns.Count - 1
                    'Dim objBookingCust As clsBookingTemp = TryCast(gv1.Columns(ii).Tag, clsBookingTemp)
                    If (clsCommon.myCdbl(lblTotRAmt1.Text)) > 0 Then
                        '' Performa Invoice No save only in case of GKD client
                        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GKD") = CompairStringResult.Equal Then
                            Dim strPerformaInvoiceNo = clsERPFuncationality.GetNextCode(Nothing, clsCommon.myCDate(txtDate.Value), clsDocType.frmPerformaInvoiceBooking, "", txtLocation.Value)
                            qry = "Update TSPL_BOOKING_DETAIL set DocumentAmount=" & clsCommon.myCdbl(lblTotRAmt1.Text) & ",Performance_Invoice_no='" & strPerformaInvoiceNo & "' where   Document_No='" & obj.Document_No & "' and Cust_Code='" & txtVendorNo.Value & "' and    Loc_Code='" & clsCommon.myCstr(txtLocation.Value) & "' and isnull(Scheme_Item,'N')='N' and isnull(FOC_Item,0)=0 "
                            clsDBFuncationality.ExecuteNonQuery(qry)
                        End If
                        If Not AllowWo_Outstanding Then
                            'Booking STATUS 1 -open 2 - Park 3 - approved 4 - posted 5 - rejected
                            If CheckOutstandingOnbooking = 1 AndAlso ((BookingStatus = 1 OrElse BookingStatus = 2) AndAlso (BookingStatus <> 5)) Then
                                If Not CustomerOutstandingAmount(txtVendorNo.Value, Nothing) Then
                                    qry = "Update TSPL_BOOKING_DETAIL set CreditApproval_Reqd='Y',Booking_Status=2 where   Document_No='" & obj.Document_No & "' and Cust_Code='" & txtVendorNo.Value & "' and    Loc_Code='" & clsCommon.myCstr(txtLocation.Value) & "'"
                                    clsDBFuncationality.ExecuteNonQuery(qry)
                                    qry = "insert into TSPL_TRANSACTION_APPROVAL(Screen_Name,Program_Code,Document_No,Doc_Date,approval_type,Approve,Created_By,Created_Date,Modified_By,Modified_Date,Comp_Code,Cust_Code,Loc_Code) " &
                               "values ('Booking Dairy','" & clsUserMgtCode.frmbookingdairy & "','" & obj.Document_No & "', " &
                               "'" & clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy") & "','Credit Limit',0,'" & objCommonVar.CurrentUserCode & "', " &
                               "'" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") & "', " &
                               "'" & objCommonVar.CurrentUserCode & "','" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") & "', " &
                               "'" & objCommonVar.CurrentCompanyCode & "','" & txtVendorNo.Value & "','" & txtLocation.Value & "')"
                                    clsDBFuncationality.ExecuteNonQuery(qry)
                                    '' create sms content
                                    Dim dtSMSEmail As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,EMail_Text from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'")
                                    Dim strSMSContent As String = ""
                                    Dim strEMailContent As String = ""
                                    If dtSMSEmail.Rows.Count > 0 Then
                                        strSMSContent = clsCommon.myCstr(dtSMSEmail.Rows(0).Item("SMS_Text"))
                                        strEMailContent = clsCommon.myCstr(dtSMSEmail.Rows(0).Item("EMail_Text"))
                                    End If
                                    'SMSCode Start
                                    If clsCommon.myLen(strSMSContent) > 0 Then
                                        Dim objSMSH As New clsSMSHead()
                                        objSMSH.SMS_Text = strSMSContent
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(txtDocNo.Value))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.BookingAmount, clsCommon.myCdbl(lblTotRAmt1.Text))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Cust_Code, clsCommon.myCstr(txtVendorNo.Value))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Cust_Name, clsCommon.myCstr(lblVendorName.Text))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Loc_Code, clsCommon.myCstr(txtLocation.Value))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Loc_Name, clsCommon.myCstr(lblLocation.Text))
                                        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Approval_Type, "Approved")
                                        CreateSMSContent(objSMSH.SMS_Text, Nothing)
                                        'obj1.SMS_Content = objSMSH.SMS_Text
                                    End If
                                    'email content Start
                                    If clsCommon.myLen(strEMailContent) > 0 Then
                                        Dim objEmailH As New clsEMailHead
                                        objEmailH.Email_Text = strEMailContent
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.BookingAmount, clsCommon.myCdbl(lblTotRAmt1.Text))
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Cust_Code, clsCommon.myCstr(txtVendorNo.Value))
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Cust_Name, clsCommon.myCstr(lblVendorName.Text))
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Loc_Code, clsCommon.myCstr(txtLocation.Value))
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Loc_Name, clsCommon.myCstr(lblLocation.Text))
                                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Approval_Type, "Approved")
                                        CreateEmailContent(objEmailH.Email_Text, Nothing)
                                    End If
                                Else
                                    qry = "Update TSPL_BOOKING_DETAIL set CreditApproval_Reqd='N',Booking_Status=1 where   Document_No='" & obj.Document_No & "' and Cust_Code='" & txtVendorNo.Value & "' and    Loc_Code='" & clsCommon.myCstr(txtLocation.Value) & "' "
                                    clsDBFuncationality.ExecuteNonQuery(qry)
                                End If
                            End If
                        End If
                    End If
                    ''============25/04/2018 Send Notification Alert for Ex Factory Date Entry, It show alert date befor one day of EX_factory_Date===========
                    Dim strNotificationOn As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_On from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmDairyBookingCustomer & "'", Nothing))
                    If clsCommon.CompairString(strNotificationOn, "S") = CompairStringResult.Equal AndAlso txtEx_Factory_Date.Checked Then
                        'If txtEx_Factory_Date.Checked Then
                        Dim Booking_Id As String = obj.Document_No
                        Dim Booking_Date As DateTime = clsCommon.myCDate(txtDate.Value)
                        Dim Ex_Factory_Date As DateTime = Nothing
                        If clsCommon.myLen(txtEx_Factory_Date.Value) > 0 Then
                            Ex_Factory_Date = clsCommon.myCDate(txtEx_Factory_Date.Value)
                        End If
                        CreateNotificationContentEMP(Booking_Id, Booking_Date, Ex_Factory_Date, Nothing)
                        'End If
                    End If
                    ''=============================================
                    'Next
                    LoadData(obj.Document_No, NavigatorType.Current)
                    Return True
                End If
                'FlagCopy = False
            Else
                Return False
            End If
            blnSaveTotalQTy = True
        Catch ex As Exception
            blnSaveTotalQTy = True
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
        Return False
    End Function
    Public Shared Sub CreateSMSContent(ByVal strSMSContent As String, ByVal trans As SqlTransaction)
        If clsCommon.myLen(strSMSContent) > 0 Then
            Dim objSMSH As New clsSMSHead()
            objSMSH.SMS_Text = strSMSContent
            objSMSH.arrMobilNo = New List(Of String)()
            objSMSH.SaveData(clsUserMgtCode.frmbookingdairy, objSMSH, trans)
            objSMSH = Nothing
        End If
    End Sub
    Public Shared Sub CreateEmailContent(ByVal strEmailContent As String, ByVal trans As SqlTransaction)
        'MailCode Start
        If clsCommon.myLen(strEmailContent) > 0 Then
            'Dim qry As String = "SELECT EMail_Subject from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'"
            Dim EmailSubject As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT EMail_Subject from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'", trans))
            Dim objSMSH As New clsEMailHead()
            objSMSH.Email_Text = strEmailContent
            objSMSH.Email_Subject = EmailSubject
            objSMSH.arrEMail = New List(Of String)()
            objSMSH.SaveData(clsUserMgtCode.frmbookingdairy, objSMSH, trans)
            objSMSH = Nothing
        End If
    End Sub
    Private Function CustomerOutstandingAmount(ByVal strCustomer As String, ByVal trans As SqlTransaction) As Boolean
        'Ticket No-MIL/08/04/19-000062,sanjay, Pass transaction object in function
        Try
            Dim strCustomerCategory As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(strCustomer) & "' ", trans))
            If clsCommon.myLen(strCustomer) > 0 AndAlso (clsCommon.CompairString(strCustomerCategory, "Institution CR") = CompairStringResult.Equal OrElse clsCommon.CompairString(strCustomerCategory, "Institution SO") = CompairStringResult.Equal) Then
                Return True
            End If
            'If throughShipment = False Then
            'PendingAfterApproval = 0
            Dim dblOutstandingAmt As Double = 0
            Dim dblCreditLimit As Double = 0
            Dim dblSecurityAmount As Double = 0
            Dim dblRefundAmount As Double = 0
            Dim dblReverseSecurityAmount As Double = 0
            Dim dblReverseRefundAmount As Double = 0
            Dim dblPendingDeliveryAmt As Double = 0
            Dim dblARSecurityAmt As Double = 0
            Dim dblAmt As Double = 0
            Dim qry As String = ""
            Dim dblShortCloseDoDispatch As Double = 0
            Dim dblBookingAmt As Double = 0
            Dim strCreditLimit As String = ""
            If DoNotConsiderCustomerCreditLimit = 1 Then
                strCreditLimit = " and CheckCreditLimit=1 "
            End If
            Dim strIsParent As String = clsDBFuncationality.getSingleValue("select Parent_Customer_YN from TSPL_CUSTOMER_MASTER where Cust_Code='" & strCustomer & "' ", trans)
            Dim strParentCust As String = clsDBFuncationality.getSingleValue("select Parent_Customer_No from TSPL_CUSTOMER_MASTER where Parent_Customer_YN='N' and Parent_Customer_No <> '' and Cust_Code='" & strCustomer & "'   and Credit_Limit=0", trans)
            If clsCommon.CompairString(strIsParent, "Y") = CompairStringResult.Equal Then
                strParentCust = strCustomer
            End If
            Dim strcustomerfilter As String = String.Empty
            strcustomerfilter = "'" & strCustomer & "'"
            If clsCommon.myLen(strParentCust) > 0 OrElse clsCommon.CompairString(strIsParent, "Y") = CompairStringResult.Equal Then
                'Dim strDate As Date = txtDate.Value.AddDays(1)
                ' qry = "Select  ( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))  as BalAmt From ( " &
                qry = "Select  case when (( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt)) )>=0 then -abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) else abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) end  as BalAmt From ( " &
                    "Select MAX(TSPL_CUSTOMER_MASTER.Cust_Group_Code) as Cust_Group_Code, ACode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName, '' as CurrencyCode,  " &
                    "null as ConvRate, SUM(DrAmt* Final.ConvRate)-SUM(CrAmt) as OpngBal, 0 as DrAmt, 0 as CrAmt, 0 as [Sales], 0 as CollectionRefund, 0 as DrNote,  " &
                    "0 as CrNote, MAX(tspl_customer_master.Cust_Category_Code) as Cust_Category_Code,MAX(CUST_CATEGORY_DESC) as Cust_Category_Desc,  " &
                    "MAX(tspl_customer_master.Cust_Type_Code) As Cust_Type_Code,MAX(Cust_Type_Desc) As Cust_Type_Desc from   " &
                    "(" & clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, "ConvRate", strcustomerfilter, True, clsCommon.GetPrintDate(txtDate.Value.AddDays(1), "dd/MMM/yyyy"), "", False, False, True, trans, False) & "   " &
                    " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code LEFT OUTER JOIN TSPL_CUSTOMER_GROUP_MASTER ON TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code " &
                    "Left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No =Final.DocNo  LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=Final.Bank_Code " &
                    "where  CONVERT(DATE,final.DocDate,103) <= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND LEN(ACode)>0 and ACode in ('" & strCustomer & "')   AND TSPL_CUSTOMER_MASTER.Status='N' GROUP BY ACode " &
                    ") XXX GROUP BY ACode ORDER BY ACode"
                dblOutstandingAmt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
                dblCreditLimit = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Credit_Limit from TSPL_CUSTOMER_MASTER where Cust_Code='" & strCustomer & "' " & strCreditLimit & "", trans))
                If Not DonotIncludeSecurityInCustomerOutstanding Then
                    dblSecurityAmount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select case when isnull(SUM(Receipt_Amount),0)>=0 then -abs(isnull(SUM(Receipt_Amount),0)) else abs(isnull(SUM(Receipt_Amount),0)) end as SecurityAmount from TSPL_RECEIPT_HEADER where Receipt_Type='P' and  SecurityDeposit='Y'  and Posted='Y' and Cust_Code='" & strCustomer & "'", trans))
                End If
                qry = "select isnull(sum(TSPL_BOOKING_MATSER.Total_Amt),0) from TSPL_BOOKING_MATSER left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No 
where TSPL_BOOKING_DETAIL.CUST_CODE = '" & strCustomer & "' and TSPL_BOOKING_MATSER.Posted=1  AND NOT EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD  LEFT JOIN TSPL_SD_SHIPMENT_DETAIL ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE WHERE TSPL_SD_SHIPMENT_HEAD.Status = 1  AND TSPL_SD_SHIPMENT_HEAD.Customer_Code = '" & strCustomer & "'  AND TSPL_BOOKING_DETAIL.PK_ID = TSPL_SD_SHIPMENT_DETAIL.Against_Booking_PK_ID)"
                dblPendingDeliveryAmt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
                qry = "select SUM( isnull( TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt, 0 ) ) as OutStandingAmt, 1 as RI 
from TSPL_SD_SHIPMENT_HEAD left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.document_code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
left outer join TSPL_BOOKING_DETAIL on TSPL_SD_SHIPMENT_DETAIL.Against_Booking_PK_ID = TSPL_BOOKING_DETAIL.PK_ID 
where TSPL_BOOKING_DETAIL.Cust_Code = '" & strCustomer & "' and TSPL_SD_SHIPMENT_HEAD.Status = 0 
and TSPL_BOOKING_DETAIL.document_No in ( SELECT DISTINCT TSPL_BOOKING_DETAIL.Document_No FROM TSPL_BOOKING_DETAIL INNER JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.dOCUMENT_nO = TSPL_BOOKING_DETAIL.Document_No where TSPL_BOOKING_MATSER.From_Screen_Code <> 'BOOK-DS_FSH' AND TSPL_BOOKING_DETAIL.CUST_CODE = '" & strCustomer & "' )"
                dblShortCloseDoDispatch = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
            Else
                'qry = "Select  ( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))  as BalAmt From ( " &
                qry = "Select  case when (( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt)) )>=0 then -abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) else abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) end  as BalAmt From ( " &
                    "Select MAX(TSPL_CUSTOMER_MASTER.Cust_Group_Code) as Cust_Group_Code, ACode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName, '' as CurrencyCode,  " &
                    "null as ConvRate, SUM(DrAmt* Final.ConvRate)-SUM(CrAmt) as OpngBal, 0 as DrAmt, 0 as CrAmt, 0 as [Sales], 0 as CollectionRefund, 0 as DrNote,  " &
                    "0 as CrNote, MAX(tspl_customer_master.Cust_Category_Code) as Cust_Category_Code,MAX(CUST_CATEGORY_DESC) as Cust_Category_Desc,  " &
                    "MAX(tspl_customer_master.Cust_Type_Code) As Cust_Type_Code,MAX(Cust_Type_Desc) As Cust_Type_Desc from   " &
                    "(" & clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, "ConvRate", strcustomerfilter, True, clsCommon.GetPrintDate(txtDate.Value.AddDays(1), "dd/MMM/yyyy"), "", False, False, True, trans, False) & "   " &
                    " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code LEFT OUTER JOIN TSPL_CUSTOMER_GROUP_MASTER ON TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code " &
                    "Left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No =Final.DocNo  LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=Final.Bank_Code " &
                    "where  CONVERT(DATE,final.DocDate,103) <= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND LEN(ACode)>0 and ACode in ('" & strCustomer & "')   AND TSPL_CUSTOMER_MASTER.Status='N' GROUP BY ACode " &
                    ") XXX GROUP BY ACode ORDER BY ACode"
                dblOutstandingAmt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
                dblCreditLimit = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Credit_Limit from TSPL_CUSTOMER_MASTER where Cust_Code='" & strCustomer & "' " & strCreditLimit & "", trans))
                If Not DonotIncludeSecurityInCustomerOutstanding Then
                    dblSecurityAmount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select case when isnull(SUM(Receipt_Amount),0)>=0 then -abs(isnull(SUM(Receipt_Amount),0)) else abs(isnull(SUM(Receipt_Amount),0)) end as SecurityAmount  from TSPL_RECEIPT_HEADER where Receipt_Type='P' and SecurityDeposit='Y'  and Posted='Y' and Cust_Code='" & strCustomer & "'", trans))
                End If
                qry = "select isnull(sum(TSPL_BOOKING_MATSER.Total_Amt),0) from TSPL_BOOKING_MATSER left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No 
where TSPL_BOOKING_DETAIL.CUST_CODE = '" & strCustomer & "' and TSPL_BOOKING_MATSER.Posted=1 AND NOT EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD  LEFT JOIN TSPL_SD_SHIPMENT_DETAIL ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE WHERE TSPL_SD_SHIPMENT_HEAD.Status = 1  AND TSPL_SD_SHIPMENT_HEAD.Customer_Code = '" & strCustomer & "'  AND TSPL_BOOKING_DETAIL.PK_ID = TSPL_SD_SHIPMENT_DETAIL.Against_Booking_PK_ID)"
                '                qry = "select ISNULL(sum(TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Total_Amt),0) from TSPL_DELIVERY_NOTE_MASTER_FRESHSALE 
                'INNER  JOIN (SELECT DISTINCT TSPL_BOOKING_DETAIL.DELIVERY_NO,TSPL_BOOKING_MATSER.From_Screen_Code FROM TSPL_BOOKING_DETAIL INNER JOIN TSPL_BOOKING_MATSER ON  TSPL_BOOKING_MATSER.dOCUMENT_nO=TSPL_BOOKING_DETAIL.Document_No  where TSPL_BOOKING_MATSER.From_Screen_Code<>'BOOK-DS_FSH' AND TSPL_BOOKING_DETAIL.CUST_CODE='" & strCustomer & "' AND ISNULL(TSPL_BOOKING_DETAIL.DELIVERY_NO,'')<>'' ) CASHINDENTBOOKING ON CASHINDENTBOOKING.DELIVERY_NO= TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No
                'where TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.posted=1 and TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No not in (Select isnull(TSPL_SD_SHIPMENT_DETAIL.Delivery_Code,'') from TSPL_SD_SHIPMENT_HEAD left outer join TSPL_SD_SHIPMENT_DETAIL on 
                'TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE where TSPL_SD_SHIPMENT_HEAD.Status=1 and TSPL_SD_SHIPMENT_HEAD.Customer_Code='" & strCustomer & "' )
                'and TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.customer_code='" & strCustomer & "' and  
                'isnull(TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Short_Close,'N')='N' "
                dblPendingDeliveryAmt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
                qry = "select SUM( isnull( TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt, 0 ) ) as OutStandingAmt, 1 as RI 
from TSPL_SD_SHIPMENT_HEAD left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.document_code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
left outer join TSPL_BOOKING_DETAIL on TSPL_SD_SHIPMENT_DETAIL.Against_Booking_PK_ID = TSPL_BOOKING_DETAIL.PK_ID 
where TSPL_BOOKING_DETAIL.Cust_Code = '" & strCustomer & "' and TSPL_SD_SHIPMENT_HEAD.Status = 0 
and TSPL_BOOKING_DETAIL.document_No in ( SELECT DISTINCT TSPL_BOOKING_DETAIL.Document_No FROM TSPL_BOOKING_DETAIL INNER JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.dOCUMENT_nO = TSPL_BOOKING_DETAIL.Document_No where TSPL_BOOKING_MATSER.From_Screen_Code <> 'BOOK-DS_FSH' AND TSPL_BOOKING_DETAIL.CUST_CODE = '" & strCustomer & "' )"
                'qry = "select SUM(isnull(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt,0) ) as OutStandingAmt , 1 as RI from TSPL_SD_SHIPMENT_HEAD  " &
                '"left outer join TSPL_SD_SHIPMENT_DETAIL on TSPL_SD_SHIPMENT_HEAD.document_code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE " &
                '"left outer join TSPL_DELIVERY_NOTE_MASTER_FRESHSALE on TSPL_SD_SHIPMENT_DETAIL.Delivery_Code=TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No " &
                '"where  TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Customer_Code='" & strCustomer & "' and  " &
                '"isnull(TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Short_Close,'N')='Y' and  TSPL_SD_SHIPMENT_HEAD.Status=0  " &
                '" and TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.document_No in " &
                '" ( SELECT DISTINCT TSPL_BOOKING_DETAIL.DELIVERY_NO FROM TSPL_BOOKING_DETAIL INNER JOIN TSPL_BOOKING_MATSER ON  TSPL_BOOKING_MATSER.dOCUMENT_nO=TSPL_BOOKING_DETAIL.Document_No  where TSPL_BOOKING_MATSER.From_Screen_Code<>'BOOK-DS_FSH' AND TSPL_BOOKING_DETAIL.CUST_CODE='" & strCustomer & "' AND ISNULL(TSPL_BOOKING_DETAIL.DELIVERY_NO,'')<>'' )"
                dblShortCloseDoDispatch = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
            End If
            dblBookingAmt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select isnull( sum(TSPL_BOOKING_DETAIL.Booking_Qty * TSPL_BOOKING_DETAIL.Item_Rate) ,0 ) from TSPL_BOOKING_DETAIL INNER JOIN TSPL_BOOKING_MATSER ON  TSPL_BOOKING_MATSER.dOCUMENT_nO=TSPL_BOOKING_DETAIL.Document_No where TSPL_BOOKING_DETAIL.Booking_Status=4 and isnull(TSPL_BOOKING_DETAIL.DO_Posted,0)=0 and  TSPL_BOOKING_DETAIL.Cust_Code='" & strCustomer & "'  and " &
                            "TSPL_BOOKING_DETAIL.Document_No not in ('" & txtDocNo.Value & "') and TSPL_BOOKING_MATSER.From_Screen_Code<>'BOOK-DS_FSH' ", trans))
            dblAmt = dblCreditLimit + dblSecurityAmount + dblOutstandingAmt
            If Not isInsideLoadData Then
                lblCreditLimit.Text = dblCreditLimit
                lblAdvanceSecurity.Text = dblSecurityAmount
                lblReverseAdvanceSec.Text = dblReverseSecurityAmount
                lblPendingDO.Text = dblPendingDeliveryAmt
                lblLedgerOutstanding.Text = dblOutstandingAmt
                lblShortcloseDO.Text = dblShortCloseDoDispatch
                lblRefund.Text = dblRefundAmount
                lblReverseRefund.Text = dblReverseRefundAmount
                lblARSecurity.Text = dblARSecurityAmt
                lblTotalOutstansing.Text = dblAmt
                lbltotalOutstanding1.Text = dblAmt
                'lblOutstandingDesc.Text = dblOutstandingAmt
                lblTotalSecurity11.Text = dblSecurityAmount - dblReverseSecurityAmount - dblRefundAmount + dblReverseRefundAmount - dblARSecurityAmt

            End If
            lblOutstandingDesc.Text = dblOutstandingAmt
            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select  isnull(tspl_customer_master.CheckCreditLimit,0) as CheckCreditLimit from tspl_customer_master where cust_code='" & strCustomer & "'", trans)) = 0 Then
                Return True
            End If
            If dblAmt < clsCommon.myCdbl(lblTotRAmt1.Text) Then
                'Dim dblNewCredtitLimit = dblAmt - clsCommon.myCdbl(lblTotRAmt1.Text)
                Return False
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
        Return True
    End Function
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try
            Dim dblTotalDocAmt As Decimal = 0
            Dim qry As String = ""
            Dim DocDate As DateTime = clsCommon.GETSERVERDATE
            Dim obj As New clsBookingEntryDairySale
            'Dim intRow As Integer
            obj = clsBookingEntryDairySale.GetData(strCode, NavTyep, clsUserMgtCode.frmDairyBookingCustomer)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
                IsTotalQtyinKG = False
                chkSampling.Enabled = False

                rgbItemType.Visible = True
                If obj.Is_Taxable = 2 Then
                    rbtnTaxable.IsChecked = True
                ElseIf obj.Is_Taxable = 1 Then
                    rbtnNonTax.IsChecked = True
                End If
                rgbItemType.Enabled = False
                ' End If

                btnSave.Enabled = True
                btnPost.Enabled = True
                btnDelete.Enabled = True
                isInsideLoadData = True
                'btnCopy.Enabled = False
                isNewEntry = False
                btnSave.Text = "Update"
                BlankAllControls()
                LoadBlankGrid()
                chkSampling.Checked = IIf(obj.IsSampling = 1, True, False)
                chkIsEwayBill.Checked = IIf(obj.IsEwaybill = 1, True, False)
                chkGatePass.Checked = IIf(obj.AgainstGatePass = 1, True, False)
                chkisTCS.Checked = IIf(obj.IS_TCS = 1, True, False)
                chkDCS.Checked = IIf(obj.Is_DCS = 1, True, False)
                chkBPL.Checked = IIf(obj.Is_BPL = 1, True, False)
                chkGhee.Checked = IIf(obj.Is_GHEE = 1, True, False)
                chkAPS.Checked = IIf(obj.Is_APS = 1, True, False)
                chkNoTranspoter.Checked = IIf(obj.No_Transporter = 1, True, False)
                chkGhee.Enabled = False
                chkcashsale.Checked = IIf(obj.Is_CashSale = "Y", True, False)
                chkManualVehicle.Checked = IIf(obj.Is_Manual_Vehicle = "Y", True, False)
                fndTransporter.Value = obj.Transport_Id
                lblTransporter.Text = clsDBFuncationality.getSingleValue(" select Transporter_Name from TSPL_TRANSPORT_MASTER where Transport_Id='" & fndTransporter.Value & "'")
                txtDescription.Text = obj.Description
                txtManualVehicle.Text = obj.Manual_VehicleNo
                chkcashsale.Enabled = True
                txtReceiverName.Text = obj.ReceiverName
                If chkcashsale.Checked Then
                    txtChequeNo.Text = obj.ChequeNo
                    cmbPaymentType.Text = obj.Payment_Terms
                End If
                chkDistributor.Checked = IIf(obj.Is_Distributor = 1, True, False)
                If chkBPL.Checked Then
                    txtCouponCode.Text = obj.BPL_Coupon_Code
                    txtBPLName.Text = obj.BPL_Name
                    txtBPLRemark.Text = obj.BPL_Remark
                    txtCouponDate.Value = obj.BPL_Coupon_Date
                    txtCategory.Value = obj.BPL_Category
                End If
                If chkDCS.Checked AndAlso obj.LastCollectionDate IsNot Nothing Then
                    'GetOutStandingBal(txtVendorNo.Value, txtDate.Value)
                    txtLastCollectionDate.Text = obj.LastCollectionDate
                    'Else
                    'CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
                End If
                ' txtLocation.Enabled = False
                txtVendorNo.Enabled = False
                txtDocNo.Value = obj.Document_No
                txtDate.Value = obj.Document_Date
                DocDate = obj.Document_Date
                txtSupplyDate.Value = obj.Supply_Date
                'GetUnbilledAmt(obj.Document_Date, txtVendorNo.Value)
                If clsCommon.myLen(obj.Against_Receipt_No) > 0 Then
                    txtReceipt.Value = obj.Against_Receipt_No
                    lblReceiptAmtDesc.Text = clsDBFuncationality.getSingleValue("select Receipt_Amount from TSPL_RECEIPT_HEADER where Receipt_No='" & obj.Against_Receipt_No & "'")
                End If
                If EnableCustomerPODetailonDairyBooking = 1 Then
                    txtSalesman.Value = obj.SalesmanCode
                    If obj.Podate IsNot Nothing Then
                        txtCustPODate.Value = obj.Podate
                        txtCustPODate.Checked = True
                    End If
                    txtPONo.Text = obj.Cust_PO_No
                End If
                chkExcludeKKFMND.Checked = IIf(obj.Exclude_KKF_And_Mandi = 1, True, False)
                chkExcludeKKFMND.Enabled = False
                txtFATPER.Text = obj.FAT_Per
                txtSNFPER.Text = obj.SNF_Per
                txtAcidity.Text = obj.Acidity
                txtTemp.Text = obj.Temperature
                txtMBRTHours.Text = obj.MBRT_Hours
                TxtRoundoff.Text = clsCommon.myCstr(obj.RoundOffAmount)
                txtSubLocation.Value = clsCommon.myCstr(obj.Sub_Location_code)
                If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                    lblSubLocation.Text = clsLocation.GetName(txtSubLocation.Value, Nothing)
                Else
                    lblSubLocation.Text = ""
                End If
                If obj.PO_Indent_Date IsNot Nothing Then
                    txtIndentNo.Text = obj.PO_Indent_No
                    txtIndentDate.Value = obj.PO_Indent_Date

                End If
                lblCreatedByValue.Text = clsCommon.myCstr(obj.Created_By)
                lblDONumber.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(Delivery_No,'') from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'"))
                'DOStatus = clsDBFuncationality.getSingleValue("select isnull(DO_Posted,0) from TSPL_BOOKING_DETAIL where Document_No ='" & txtDocNo.Value & "' and cust_code='" & txtVendorNo.Value & "'")
                If obj.Posted = 1 Then
                    btnSave.Enabled = False
                    btnDelete.Enabled = False
                    btnPost.Enabled = False
                    btnGatepass.Enabled = True
                    UsLock1.Status = ERPTransactionStatus.Approved
                    'Dim isDemandBooking = clsDBFuncationality.getSingleValue("select top 1 Against_DemandBooking_No from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'")
                    'If clsCommon.myLen(isDemandBooking) > 0 Then
                    '    btnCreateAndPrintInvoice.Enabled = False
                    'Else
                    txtOpeningbal.Text = obj.OpeningBal
                    txtDrAmt.Text = obj.DrAmt
                    txtCrAmt.Text = obj.CrAmt
                    txtClosingBal.Text = obj.ClosingBal
                    btnCreateAndPrintInvoice.Enabled = True
                    chkIsEwayBill.Enabled = False
                    'End If
                    btnCreateDO.Enabled = True

                    Dim DOStatus1 = clsDBFuncationality.getSingleValue("select top 1  Document_No from TSPL_BOOKING_DETAIL where DO_Posted <> 4 and Document_No='" & txtDocNo.Value & "'")
                    If clsCommon.myLen(DOStatus1) = 0 Then
                        btnCreateDO.Enabled = False
                    End If
                Else
                    chkIsEwayBill.Enabled = True
                    btnCreateEWB.Enabled = False
                    btnCreateAndPrintInvoice.Enabled = False
                    UsLock1.Status = ERPTransactionStatus.Pending
                End If
                'If obj.TRANSACTION_TYPE = "FS" Then
                '    rbtn_Fresh.IsChecked = True
                'ElseIf obj.TRANSACTION_TYPE = "PS" Then
                '    rbtn_Ambient.IsChecked = True
                'End If
                If clsCommon.myLen(obj.Ex_Factory_Date) = 0 Then
                    txtEx_Factory_Date.Checked = False
                Else
                    txtEx_Factory_Date.Checked = True
                    txtEx_Factory_Date.Value = obj.Ex_Factory_Date
                End If
                lblCreatedDateAndTime.Text = obj.Created_Date
                ''        cmbBookingType.Text = IIf(obj.Booking_Type = "", "Select", obj.Booking_Type)
                txtLocation.Value = obj.location_code
                txtAdharNo.Text = obj.AdharNo
                txtMobileNo.Text = obj.MobileNO

                lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
                txtCan.Text = obj.TotalCAN
                If obj.NoCrateIssue = 1 Then
                    txtCrate.Text = 0
                Else
                    txtCrate.Text = obj.TotalCrate
                End If
                chkNoCrateIssue.Checked = IIf(obj.NoCrateIssue = 1, True, False)

                txtBox.Text = obj.TotalBox
                txtShipToLocation.Value = obj.Ship_To_Location
                lblShipToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_SHIP_TO_LOCATION.Ship_To_Desc FROM  TSPL_SHIP_TO_LOCATION WHERE Ship_To_Code  ='" & txtShipToLocation.Value & "'"))
                If obj.Uploading_date IsNot Nothing Then
                    lblUploadingDate.Text = obj.Uploading_date
                End If
                'Sanjay ERO/12/07/18-000371
                Is_Cancelled = obj.Is_Cancelled
                lblCancelStatus.Text = IIf(obj.Is_Cancelled = 1, "Cancel", "")
                If obj.Is_Cancelled = 1 Then
                    btnCancel.Enabled = False
                Else
                    btnCancel.Enabled = True
                End If
                GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
                lblCreditLimit.Text = clsCommon.myCdbl(obj.Credit_Limit)
                lblAdvanceSecurity.Text = clsCommon.myCdbl(obj.Advance_Security)
                lblReverseAdvanceSec.Text = clsCommon.myCdbl(obj.Revese_Adv_Security)
                lblARSecurity.Text = clsCommon.myCdbl(obj.AR_Credit_Security)
                lblPendingDO.Text = clsCommon.myCdbl(obj.Pending_Posted_DO)
                lblShortcloseDO.Text = clsCommon.myCdbl(obj.UnPostedDispatch)
                lblLedgerOutstanding.Text = clsCommon.myCdbl(obj.Ledger_Outstansing)
                lblRefund.Text = clsCommon.myCdbl(obj.Refund_Security)
                lblReverseRefund.Text = clsCommon.myCdbl(obj.Reverse_Refund_Sec)
                lblTotalOutstansing.Text = clsCommon.myCdbl(obj.Total_Outstanding)
                lbltotalOutstanding1.Text = clsCommon.myCdbl(obj.Total_Outstanding)
                lblTotalSecurity11.Text = clsCommon.myCdbl(obj.Advance_Security) - clsCommon.myCdbl(obj.Revese_Adv_Security) - clsCommon.myCdbl(obj.Refund_Security) + clsCommon.myCdbl(obj.Reverse_Refund_Sec) - clsCommon.myCdbl(obj.AR_Credit_Security)
                qry = "SELECT TSPL_BOOKING_DETAIL.*,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_ROUTE_MASTER.Route_Desc,tspl_item_master.item_desc,tspl_item_master.Short_Description as item_Short_Description,tspl_item_master.HSN_Code FROM TSPL_BOOKING_DETAIL LEFT JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code left outer join " &
                        " TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join tspl_item_master on tspl_item_master.Item_code=TSPL_BOOKING_DETAIL.Item_code WHERE Document_No='" & txtDocNo.Value & "' and scheme_item='N'"
                Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
                txtVendorNo.Value = clsCommon.myCstr(dt2.Rows(0)("Cust_Code"))
                lblVendorName.Text = clsCommon.myCstr(dt2.Rows(0)("Customer_Name"))
                txtVehicleCode.Value = clsCommon.myCstr(dt2.Rows(0)("Vehicle_Code"))
                txtVehicleName.Text = clsDBFuncationality.getSingleValue("select Number from TSPL_VEHICLE_MASTER where Vehicle_id='" & txtVehicleCode.Value & "'")
                If clsCommon.CompairString(obj.Booking_Type, "") = CompairStringResult.Equal Then
                    cmbcashcredit.Text = ""
                ElseIf clsCommon.CompairString(obj.Booking_Type, "PS") = CompairStringResult.Equal Then
                    cmbcashcredit.Text = "PARLOR SALES"
                ElseIf clsCommon.CompairString(obj.Booking_Type, "FN") = CompairStringResult.Equal Then
                    cmbcashcredit.Text = "FORENOON"
                ElseIf clsCommon.CompairString(obj.Booking_Type, "UP") = CompairStringResult.Equal Then
                    cmbcashcredit.Text = "UP COUNTRY"
                Else
                    cmbcashcredit.Text = obj.Booking_Type
                End If
                '====================Added by preeti Gupta Against Ticket no[BHA/01/08/18-000206]=
                lblBoothStation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(Zone_Code,'') +  case when len(oldname )>0 then ', ' +isnull(oldname,'') else '' end from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
                lblroutecode.Text = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
                txtRouteCode1.Text = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
                lblroutename.Text = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
                txtRouteName1.Text = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
                strRoutecode = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
                strRouteDesc = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
                lblTotRAmt1.Text = clsCommon.myCstr(dt2.Rows(0)("DocumentAmount"))
                BookingStatus = clsCommon.myCstr(dt2.Rows(0)("Booking_Status"))
                DOStatus = clsCommon.myCstr(dt2.Rows(0)("DO_Posted"))
                Dim objTaxGrpMaster As New clsTaxGroupMaster()
                objTaxGrpMaster = objTaxGrpMaster.GetDataForSale(obj.Tax_Group)
                If (objTaxGrpMaster IsNot Nothing) Then
                    lblTaxGrpName.Text = objTaxGrpMaster.Tax_Group_Desc
                End If
                txtTermCode.Value = obj.Terms_Code
                'lblTermName.Text = obj.Terms_Description
                If obj.Due_Date IsNot Nothing Then
                    txtDueDate.Value = obj.Due_Date
                End If
                If obj.Posted = 0 Then
                    GetOpeningClosingAndReceivedAmt(txtVendorNo.Value, txtDate.Value)
                End If
                lblAmtWithDiscount.Text = clsCommon.myFormat(obj.Discount_Base)
                lblDiscountAmt.Text = clsCommon.myFormat(obj.Discount_Amt)
                lblAmtAfterDiscount.Text = clsCommon.myFormat(obj.Amount_Less_Discount)
                lblTaxAmt.Text = clsCommon.myFormat(obj.Total_Tax_Amt)
                lblTotRAmt.Text = clsCommon.myFormat(obj.Total_Amt)
                txtDCAmt.Text = clsCommon.myFormat(obj.Distributor_Commission_TotalAmt)
                txtTCAmt.Text = clsCommon.myFormat(obj.Transporter_Commission_TotalAmt)
                txtSecurity.Text = clsCommon.myFormat(obj.Security_TotalAmt)
                If DeductTPTFromDocAmt Then
                    txtTPTComAmt.Text = obj.Transporter_Commission_TotalAmt
                    lblGrossAmount.Text = clsCommon.myFormat(clsCommon.myCdbl(lblTotRAmt.Text) - obj.Transporter_Commission_TotalAmt)
                End If
                txtTaxGroup.Value = obj.Tax_Group
                lblTaxGrpName.Text = obj.TaxGroupName
                lblTermName.Text = obj.TermsName
                LoadBlankGridTax()
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
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_TCS  from tspl_tax_master where tax_code ='" & clsCommon.myCstr(obj.TAX1) & "' ")), "Y") = CompairStringResult.Equal Then
                        txtTCSTaxRate.Value = clsCommon.myCdbl(gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value)
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
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_TCS  from tspl_tax_master where tax_code ='" & clsCommon.myCstr(obj.TAX2) & "' ")), "Y") = CompairStringResult.Equal Then
                        txtTCSTaxRate.Value = clsCommon.myCdbl(gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value)
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
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_TCS  from tspl_tax_master where tax_code ='" & clsCommon.myCstr(obj.TAX3) & "' ")), "Y") = CompairStringResult.Equal Then
                        txtTCSTaxRate.Value = clsCommon.myCdbl(gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value)
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
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_TCS  from tspl_tax_master where tax_code ='" & clsCommon.myCstr(obj.TAX4) & "' ")), "Y") = CompairStringResult.Equal Then
                        txtTCSTaxRate.Value = clsCommon.myCdbl(gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value)
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
                If obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic Then
                    rbtnTaxCalAutomatic.IsChecked = True
                ElseIf obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual Then
                    rbtnTaxCalManual.IsChecked = True
                End If
                ''richa agarwal ERO/21/05/19-000609 21 May,2019 add updated vehicle No according to DO
                LblUpdatedVehicleCode.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Vehicle_Code from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'"))
                LblUpdatedVehicleDesc.Text = clsCommon.myCstr(ClsVehicleMaster.GetName(LblUpdatedVehicleCode.Text, Nothing))
                'setRouteDetail(txtVendorNo.Value, lblroutecode.Text)
                'If CheckOutstandingOnbooking = 1 Then
                GetUnbilledAmt(obj.Document_Date, txtVendorNo.Value)
                If chkDCS.Checked Then
                    lblOutstandingDesc.Text = clsCommon.myCstr(clsDairyInvoice.GetOutStandingBal(txtVendorNo.Value, txtDate.Value))
                Else
                    CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
                End If
                'End If

                ''richa TEC/01/10/19-001025
                txtRouteNo.Value = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
                lblRouteDesc.Text = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
                If clsCommon.CompairString(obj.GatePass_Type, "") = CompairStringResult.Equal Then
                    cmbGatePassType.Text = "Select"
                Else
                    cmbGatePassType.Text = clsCommon.myCstr(obj.GatePass_Type)
                End If
                ' done by priti BHA/14/06/18-000052
                If clsCommon.myLen(BookingStatus) > 0 Then
                    If BookingStatus = 1 Then
                        txtBOstatus.Text = "Open"
                    ElseIf BookingStatus = 2 Then
                        txtBOstatus.Text = "Pending"
                    ElseIf BookingStatus = 3 Then
                        txtBOstatus.Text = "Approved"
                    ElseIf BookingStatus = 4 Then
                        txtBOstatus.Text = "Posted"
                    ElseIf BookingStatus = 5 Then
                        txtBOstatus.Text = "Rejected"
                    End If
                End If
                If clsCommon.myLen(BookingStatus) > 0 Then
                    If DOStatus = 1 Then
                        txtDOStatus.Text = "Open"
                    ElseIf DOStatus = 2 Then
                        txtDOStatus.Text = "Pending"
                    ElseIf DOStatus = 3 Then
                        txtDOStatus.Text = "Approved"
                    ElseIf DOStatus = 4 Then
                        txtDOStatus.Text = "Posted"
                    ElseIf DOStatus = 5 Then
                        txtDOStatus.Text = "Rejected"
                    End If
                End If
                If obj.AgainstGatePass = 1 AndAlso clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "TSDDCF") = CompairStringResult.Equal Then
                    btnGatePassPrint.Visible = True
                Else
                    btnGatePassPrint.Visible = False
                End If
                For jj As Integer = 0 To dt2.Rows.Count() - 1
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count 'clsCommon.myCdbl(dt2.Rows(jj)("Line_No"))
                    'gv1.Rows(gv1.Rows.Count - 1).Cells(colIsKKF).Value = (clsCommon.CompairString(clsCommon.myCstr(dt2.Rows(jj)("IsKKFTax")), "YES") = CompairStringResult.Equal)
                    'gv1.Rows(gv1.Rows.Count - 1).Cells(colIsMNDTax).Value = (clsCommon.CompairString(clsCommon.myCstr(dt2.Rows(jj)("IsMNDTax")), "YES") = CompairStringResult.Equal)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dt2.Rows(jj)("Item_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColMainItem).Value = clsCommon.myCstr(dt2.Rows(jj)("Scheme_Main_Item"))

                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dt2.Rows(jj)("item_desc"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dt2.Rows(jj)("item_Short_Description"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dt2.Rows(jj)("HSN_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIType).Value = clsDBFuncationality.getSingleValue("select TypeOfItm from TSPL_ITEM_MASTER where Item_Code='" & gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value & "' ")
                    If ShowAvailableQtyOnDairyBooking = True Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColAvailableQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")), IIf(clsCommon.myLen(txtSubLocation.Value) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dt2.Rows(jj)("Unit_Code")), 0)
                    End If
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("PreviousBookingQty"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColBillingQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("Billing_Qty"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("OrgRate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dt2.Rows(jj)("Unit_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBillingUOM).Value = clsCommon.myCstr(dt2.Rows(jj)("Billing_Unit_code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSellingRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Selling_Price"))
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = (clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Price_with_Tax"))) - (clsCommon.myCdbl(dt2.Rows(jj)("TAX1_Amt")) + clsCommon.myCdbl(dt2.Rows(jj)("TAX2_Amt")) + clsCommon.myCdbl(dt2.Rows(jj)("TAX3_Amt")) + clsCommon.myCdbl(dt2.Rows(jj)("TAX4_Amt")))
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = Math.Round(clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate")), 2)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Amt_Less_Discount"))

                    Else
                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = Math.Round(clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate")), 2)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Amt_Less_Discount")) '' 
                    End If
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colIsBatchItem).Value = clsItemMaster.IsBatchItem(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")))
                    'gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = Math.Round(clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate")), 2)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_On_Amount"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_Amount"))
                    UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Amount).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Amount"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Code).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Code"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Pers).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Pers"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Type).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Type"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeType).Value = clsCommon.myCstr(dt2.Rows(jj)("Scheme_Type"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeItem).Value = clsCommon.myCstr(dt2.Rows(jj)("Scheme_Item"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_NonTax"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dt2.Rows(jj)("FreshAmbient"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colRemarks).Value = clsCommon.myCstr(dt2.Rows(jj)("Remarks"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colItemBasicPrice).Value = clsCommon.myCdbl(dt2.Rows(jj)("Price_with_Tax"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAmountWithTax).Value = clsCommon.myCdbl(dt2.Rows(jj)("Amount_with_Tax"))
                    dblTotalDocAmt = dblTotalDocAmt + clsCommon.myCdbl(gv1.Rows(gv1.Rows.Count - 1).Cells(colAmountWithTax).Value)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceId).Value = clsCommon.myCstr(dt2.Rows(jj)("Item_Price_ID"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceIDAppDate).Value = clsCommon.myCstr(dt2.Rows(jj)("Price_IdStartDate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPricePlanNo).Value = clsCommon.myCstr(dt2.Rows(jj)("PricePlanNo"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPricePlanNo).Value = clsCommon.myCstr(dt2.Rows(jj)("PricePlanNo"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPricePlanNo).Value = clsCommon.myCstr(dt2.Rows(jj)("PricePlanNo"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPricePlanNo).Value = clsCommon.myCstr(dt2.Rows(jj)("PricePlanNo"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colQtyinKG).Value = clsCommon.myCdbl(dt2.Rows(jj)("QtyinKG"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterDis).Value = clsCommon.myCdbl(dt2.Rows(jj)("Amt_Less_Discount"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxGroup).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX_Group"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(1))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX1"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(1))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX1_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(1))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX1_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(1))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX1_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(2))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX2"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(2))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX2_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(2))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX2_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(2))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX2_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(3))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX3"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(3))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX3_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(3))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX3_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(3))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX3_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(4))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX4"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(4))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX4_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(4))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX4_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(4))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX4_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(5))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX5"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(5))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX5_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(5))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX5_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(5))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX5_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(6))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX6"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(6))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX6_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(6))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX6_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(6))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX6_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(7))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX7"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(7))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX7_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(7))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX7_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(7))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX7_Amt"))
                    'gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(7))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX2_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(8))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX8"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(8))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX8_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(8))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX8_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(8))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX8_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(9))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX9"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(9))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX9_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(9))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX9_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(9))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX9_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax" + clsCommon.myCstr(10))).Value = clsCommon.myCstr(dt2.Rows(jj)("TAX10"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Rate" + clsCommon.myCstr(10))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX10_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Base_Amt" + clsCommon.myCstr(10))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX10_Base_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(clsCommon.myCstr("colTax_Amt" + clsCommon.myCstr(10))).Value = clsCommon.myCdbl(dt2.Rows(jj)("TAX10_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCPKID).Value = clsCommon.myCstr(dt2.Rows(jj)("Distributor_Commission_PKID"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCRate).Value = clsCommon.myCstr(dt2.Rows(jj)("Distributor_Commission_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCRateWithTax).Value = clsCommon.myCstr(dt2.Rows(jj)("Distributor_Commission_RateWithTax"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColDCAmt).Value = clsCommon.myCstr(dt2.Rows(jj)("Distributor_Commission_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColTCRate).Value = clsCommon.myCstr(dt2.Rows(jj)("Transporter_Commission_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColTCAmt).Value = clsCommon.myCstr(dt2.Rows(jj)("Transporter_Commission_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColSCRate).Value = clsCommon.myCstr(dt2.Rows(jj)("Security_Rate"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColSCAmt).Value = clsCommon.myCstr(dt2.Rows(jj)("Security_Amt"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colPKID).Value = clsCommon.myCstr(dt2.Rows(jj)("PK_ID"))
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colBatchNo).Value = clsCommon.myCstr(dt2.Rows(jj)("Batch_No"))
                    ''RICHA 06 JUNE,2020
                    If chkSampling.Checked = False And DonotAllowtoChangeUOMinDairyBookingCustomer = True Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).ReadOnly = True
                    Else
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).ReadOnly = False
                    End If
                    ' UpdateCurrentRow(jj)
                Next
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    Dim introw As Integer = 0
                    For Each objTr As clsBookingDetailDairySale In obj.Arr
                        gv1.Rows(introw).Cells(colICode).Tag = objTr.arrBatchItem
                        introw += 1
                    Next
                End If
                lblTotalDocAmt.Text = obj.Total_Amt 'Math.Round(clsCommon.myCdbl(dblTotalDocAmt), 2)
                txtTCSBaseAmt.Text = obj.TCSBaseAmt
                lblTCSAmount.Text = obj.TCSAmount

                'Try
                '    lblTCSAmount.Text = Math.Round(Math.Round(clsCommon.myCdbl(dblTotalDocAmt), 2) * GetTCSRate(txtVendorNo.Value) / 100, 2)
                'Catch ex As Exception
                'End Try
                ''to show all items other than booking in case of customer type other than others 25 Feb,2020
                'If clsCommon.CompairString(txtBOstatus.Text, "Posted") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(txtBOstatus.Text, "Rejected") <> CompairStringResult.Equal Then
                '    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") <> CompairStringResult.Equal Then
                '        qry = "  select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,TSPL_ITEM_UOM_DETAIL.UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                '        " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,0 as Booking_Qty,tspl_item_master.Sku_Seq   from tspl_item_master " & Environment.NewLine &
                '        " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
                '        " where isnull(tspl_item_master.Chilled_Freezen,0)=1 and isnull(TSPL_ITEM_MASTER.item_type,'')='F'  and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1 " & Environment.NewLine &
                '        " and tspl_item_master.Item_Code not in ( select tspl_booking_detail.item_code from tspl_booking_detail" & Environment.NewLine &
                '        " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
                '        " where tspl_booking_detail.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "' AND tspl_booking_detail.Scheme_Item ='N') order by Sku_Seq" & Environment.NewLine
                '        dt2 = clsDBFuncationality.GetDataTable(qry)
                '        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                '            For Each dr As DataRow In dt2.Rows
                '                isCellValueChangedOpen = True
                '                gv1.Rows.AddNew()
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("UOM_Code"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
                '                If clsCommon.myCdbl(clsCommon.myCdbl(dr("Booking_Qty"))) > 0 Then
                '                    ItemPrice(clsCommon.myCstr(dr("item_code")), clsCommon.myCstr(dr("Unit_code")), clsCommon.myCdbl(dr("Booking_Qty")), gv1.Rows.Count - 1)
                '                End If
                '                isCellValueChangedOpen = False
                '            Next
                '        End If
                '    Else
                '        qry = " select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,TSPL_ITEM_UOM_DETAIL.UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                '      " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,0 as Booking_Qty,tspl_item_master.Marketing_Seq   from tspl_item_master " & Environment.NewLine &
                '      " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
                '      " where isnull(tspl_item_master.Item_Type,'')='F' and isnull(tspl_item_master.TypeOfItm,'')='A' and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1 " & Environment.NewLine &
                '      " and tspl_item_master.Item_Code not in ( select tspl_booking_detail.item_code from tspl_booking_detail" & Environment.NewLine &
                '      " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
                '      " where tspl_booking_detail.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "' AND tspl_booking_detail.Scheme_Item ='N')" & Environment.NewLine &
                '      " order by Marketing_Seq"
                '        dt2 = clsDBFuncationality.GetDataTable(qry)
                '        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                '            For Each dr As DataRow In dt2.Rows
                '                isCellValueChangedOpen = True
                '                gv1.Rows.AddNew()
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("UOM_Code"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
                '                gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
                '                If clsCommon.myCdbl(clsCommon.myCdbl(dr("Booking_Qty"))) > 0 Then
                '                    ItemPrice(clsCommon.myCstr(dr("item_code")), clsCommon.myCstr(dr("Unit_code")), clsCommon.myCdbl(dr("Booking_Qty")), gv1.Rows.Count - 1)
                '                End If
                '                isCellValueChangedOpen = False
                '            Next
                '        End If
                '    End If
                'End If
                ItemTypePanel.Enabled = False
                If chkBPL.Checked Then
                    UcAttachment1.LoadData(obj.Document_No)
                End If
                If obj.Posted = 1 Then
                    Dim strEWB As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EWayBillNo from TSPL_SD_SALE_INVOICE_head  where Against_Shipment_No in ( select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "') "))
                    If clsCommon.myLen(strEWB) > 0 Then
                        btnCreateEWB.Enabled = False
                        btnprinte_wayBill.Visible = True
                    Else
                        btnCreateEWB.Enabled = True
                        btnprinte_wayBill.Visible = False
                    End If
                Else
                    btnCreateEWB.Enabled = False
                    btnprinte_wayBill.Visible = False
                End If
            End If
            If (clsCommon.myCdbl(lblTotRAmt1.Text)) > 0 Then
                If AllowWo_Outstanding = False Then
                    If CheckOutstandingOnbooking = 1 AndAlso ((BookingStatus = 1 OrElse BookingStatus = 2) AndAlso (BookingStatus <> 5)) Then
                        CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
                    End If
                End If
            End If
            gv1.Rows.AddNew()
            If gv1.Rows.Count > 0 Then
                'gv1.Focus()
                gv1.CurrentRow = gv1.Rows(0)
                gv1.CurrentColumn = gv1.Columns(colQty)
            End If
            'End If
            ChkTaxNonTax()
            If ServerDateTimeForTaxableInvoice Then
                txtDate.Value = DocDate
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
        End Try
    End Sub
    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        CloseForm()
    End Sub
    Sub CloseForm()
        clsERPFuncationality.closeForm(Me)
    End Sub
    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        If clsCommon.myLen(txtDocNo.Value) > 0 Then
            If Is_Cancelled = 1 Then
                clsCommon.MyMessageBoxShow(Me, "Booking Cancelled,Can not Post.", Me.Text)
                Exit Sub
            End If
            If BookingStatus = 2 Then
                clsCommon.MyMessageBoxShow(Me, "Booking is pending for approval,Can not Post.", Me.Text)
                Exit Sub
            End If
            'Check, If item uom is crate type in Unit of measure then Item should be Crate type in Item Master
            If AutoCalculateCrate = 1 Then
                Dim intCrateTypeItemMaster As Boolean = False
                Dim intCrateTypeUnitMaster As Boolean = False
                Dim strICode As String = ""
                Dim strIName As String = ""
                Dim strIUOM As String = ""
                For ii As Integer = 0 To gv1.Rows.Count - 1
                    strICode = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                    strIName = clsCommon.myCstr(gv1.Rows(ii).Cells(colIName).Value)
                    strIUOM = clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)
                    If (clsCommon.myLen(strICode) > 0) Then
                        intCrateTypeUnitMaster = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Crate_Type from tspl_unit_master where unit_code ='" & strIUOM & "'")) = "Y", True, False)
                        If intCrateTypeUnitMaster Then
                            intCrateTypeItemMaster = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select tspl_item_master.Is_CrateType from tspl_item_master where item_code ='" & strICode & "'")) = 1, True, False)
                            If Not intCrateTypeItemMaster Then
                                common.clsCommon.MyMessageBoxShow(Me, "Item " & strIName & " should be Crate Type in Item master. At Line No " & clsCommon.myCstr(ii + 1))
                                Exit Sub
                            End If
                        End If
                    End If
                Next
            End If
        End If
        ''richa 14 Dec,2018 ERO/11/12/18-000434
        Dim strOrderBookingPosted As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when ISNULL(Route_Time,'')<>''  then CASE WHEN CAST(GETDATE() AS TIME)<=CAST(Route_Time  AS TIME) THEN 'Y' ELSE 'N' END else 'Y' end  from tspl_route_MASTER where Route_No ='" & clsCommon.myCstr(lblroutecode.Text) & "'"))
        If clsCommon.CompairString(strOrderBookingPosted, "N") = CompairStringResult.Equal Then
            clsCommon.MyMessageBoxShow(Me, "Booking time has exceeded for Route " & lblroutecode.Text & ", it cannot be posted.", Me.Text)
            Exit Sub
        End If
        ''--------------------------------------
        PostData()
    End Sub
    Sub PostData()
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If (myMessages.postConfirm()) Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Document_Date,location_code from TSPL_BOOKING_MATSER where Document_No='" & txtDocNo.Value & "'", trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleSaleDairy, clsUserMgtCode.frmDairyBookingCustomer, clsCommon.myCstr(dt.Rows(0)("location_code")), clsCommon.myCDate(dt.Rows(0)("Document_Date")), trans)
                End If
                If clsCommon.myLen(txtDocNo.Value) > 0 Then
                    'Dim msg As String = String.Empty
                    clsBookingEntryDairySale.CreateDebitNote(txtDocNo.Value, trans)
                    Dim qry = "Update TSPL_BOOKING_MATSER set Posted=1, " &
                     "OpeningBal='" & clsCommon.myCstr(txtOpeningbal.Text) & "', " &
                     "DrAmt='" & clsCommon.myCstr(txtDrAmt.Text) & "', " &
                     "CrAmt='" & clsCommon.myCstr(txtCrAmt.Text) & "', " &
                     "ClosingBal='" & clsCommon.myCstr(txtClosingBal.Text) & "', " &
                     "Modified_By='" & objCommonVar.CurrentUserCode & "', " &
                     "Modified_Date='" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy hh:mm tt") & "' " &
                     "where Document_No='" & txtDocNo.Value & "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, txtDocNo.Value, "TSPL_BOOKING_MATSER", "Document_No", trans)


                    ''For ii As Integer = 8 To gv1.Columns.Count - 1
                    ''Dim objBooking As clsBookingTemp = TryCast(gv1.Columns(ii).Tag, clsBookingTemp)
                    'Dim dblTotalAmt As Double = clsCommon.myCdbl(lblTotRAmt1.Text)
                    'Dim strBookingStatus As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select  Booking_Status from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'  and Cust_Code='" & txtVendorNo.Value & "'", trans))
                    'If (dblTotalAmt > 0 OrElse AllowZeroQtyOnDairyBooking = True) AndAlso (strBookingStatus = 1 OrElse strBookingStatus = 3) Then
                    '    qry = "Update TSPL_BOOKING_DETAIL set Booking_Status=4 where Cust_Code='" & txtVendorNo.Value & "' and Document_No='" & txtDocNo.Value & "'  and Booking_Status<>5"
                    '    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    '    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, txtDocNo.Value, "TSPL_BOOKING_MATSER", "Document_No", "TSPL_BOOKING_DETAIL", "Document_No", "TSPL_BOOKING_PAYMENT_MODE_DETAIL", "Document_No", trans)
                    'End If
                    ''Next
                    ''== Notification regarding
                    'Dim strNotificationOn As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_On from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmDairyBookingCustomer & "'", trans))
                    'If clsCommon.CompairString(strNotificationOn, "P") = CompairStringResult.Equal Then
                    '    If txtEx_Factory_Date.Checked = True Then
                    '        Dim Booking_Id As String = txtDocNo.Value
                    '        Dim Booking_Date As DateTime = clsCommon.myCDate(txtDate.Value)
                    '        Dim Ex_Factory_Date As DateTime = Nothing
                    '        If clsCommon.myLen(txtEx_Factory_Date.Value) > 0 Then
                    '            Ex_Factory_Date = clsCommon.myCDate(txtEx_Factory_Date.Value)
                    '        End If
                    '        CreateNotificationContentEMP(Booking_Id, Booking_Date, Ex_Factory_Date, trans)
                    '    End If
                    'End If
                    ''== Complete
                    'Dim msg As String = String.Empty
                    '' btnCreateDO.Enabled = True
                    'RecordCount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_BOOKING_MATSER ", trans))
                    'If RecordCount = 0 Then
                    '    FlagFirstRecord = True
                    'Else
                    '    FlagFirstRecord = False
                    'End If
                    'FlagCreateDo = True
                    'BookingStatus = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select  Booking_Status from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'  and Cust_Code='" & txtVendorNo.Value & "'", trans))
                    'If CreateDO(False, trans, txtDocNo.Value) Then
                    '    If clsCommon.myLen(DOmsg) > 0 Then
                    '        common.clsCommon.MyMessageBoxShow(Me, DOmsg, Me.Text)
                    '    End If
                    '    If DOCreated = True Then
                    '        msg = Nothing
                    '    End If
                    '    trans.Commit()
                    '    msg = "Successfully Posted"
                    '    common.clsCommon.MyMessageBoxShow(Me, msg, Me.Text)
                    '    LoadData(txtDocNo.Value, NavigatorType.Current)
                    'Else
                    '    trans.Rollback()
                    'End If
                    trans.Commit()


                    clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)

                Else
                    Throw New Exception("No Data found to Post")
                End If
            End If
        Catch ex As Exception
            trans.Rollback()
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        Finally
            FlagCreateDo = False
        End Try
    End Sub
    Private Shared Function CreateNotificationContentEMP(ByVal Booking_Id As String, ByVal Booking_Date As DateTime, ByVal Ex_Factory_Date As DateTime, ByVal trans As SqlTransaction) As Boolean
        Dim strNotifiContent As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_Text from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmDairyBookingCustomer & "'", trans))
        Dim strNotifi_DetalContent As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_Detail_Text from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmDairyBookingCustomer & "'", trans))
        Dim strNotifiCaption As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_Caption from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmDairyBookingCustomer & "'", trans))
        Dim strNotificationOn As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_On from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmDairyBookingCustomer & "'", trans))
        If clsCommon.myLen(strNotifiContent) > 0 Then
            Dim objNotification As New clsNotificationHead()
            objNotification.Notification_Text = strNotifiContent
            objNotification.Notification_Caption = strNotifiCaption
            objNotification.Notification_On = strNotificationOn
            objNotification.Notification_Detail_Text = strNotifi_DetalContent
            objNotification.Notification_ConditionDate = clsCommon.myCDate(Ex_Factory_Date)
            objNotification.Notification_Text = objNotification.Notification_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(Booking_Id))
            objNotification.Notification_Text = objNotification.Notification_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_Date, clsCommon.myCstr(Booking_Date))
            objNotification.Notification_Text = objNotification.Notification_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Ex_Factory_Date, clsCommon.myCstr(clsCommon.myCDate(Ex_Factory_Date)))
            objNotification.SaveData(clsUserMgtCode.frmDairyBookingCustomer, objNotification, trans)
            objNotification = Nothing
            Return True
        End If
        Return False
    End Function
    'Sub Export()
    '    If gv1.Rows.Count > 0 Then
    '        ExportToExcel()
    '    Else
    '        Common.clsCommon.MyMessageBoxShow("No Data Found to Display", Me.Text)
    '    End If
    'End Sub
    'Private Sub ExportToExcel()
    '    Try
    '        Dim strCreatedBy As String = clsDBFuncationality.getSingleValue("Select Created_By from TSPL_BOOKING_MATSER where Document_No='" & txtDocNo.Value & "'")
    '        Dim arrHeader As List(Of String) = New List(Of String)()
    '        Dim strtemp As String = "Booking No : " + txtDocNo.Value
    '        arrHeader.Add(strtemp)
    '        strtemp = "Booking Date : " + clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")
    '        arrHeader.Add(strtemp)
    '        strtemp = "Created By : " + strCreatedBy
    '        arrHeader.Add(strtemp)
    '        strtemp = "Transaction Type : " + IIf(rbtn_Fresh.IsChecked = True, "Fresh", "Ambient")
    '        arrHeader.Add(strtemp)
    '        clsCommon.MyExportToExcelGrid("Booking Entry", gv1, arrHeader, Me.Text)
    '    Catch ex As Exception
    '        Common.clsCommon.MyMessageBoxShow(ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
    '    End Try
    'End Sub
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
                If (clsBookingEntryDairySale.DeleteData(txtDocNo.Value)) Then
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
    Private Function funSetUserAccess() As Boolean
        Try
            btnSave.Visible = True
            btnDelete.Visible = True
            btnPost.Visible = True
            Dim strRights As String
            Dim strTemp() As String
            Dim strProgCode = "PO-ODR"
            strRights = enuUserRights.enuRead & "," & enuUserRights.enuModify & "," & enuUserRights.enuDelete & "," & enuUserRights.enuAuthorised
            strRights = modUserMgt.funGetPermissions(strRights, strProgCode)
            strTemp = Split(strRights, ",")
            If strTemp(0) = "0" Then
                MsgBox("Permission Denied", MsgBoxStyle.Critical, Me.Text)
                funSetUserAccess = False
                blnRead = False
                Me.Close()
                Exit Function
            Else
                blnRead = True
            End If
            If strTemp(1) = "0" Then 'Grant modify access
                btnSave.Visible = False
            End If
            If strTemp(2) = "0" Then 'Grant modify access
                btnDelete.Visible = False
            End If
            If strTemp(3) = "0" Then 'Grant Authorize access
                btnPost.Visible = False
            End If
            funSetUserAccess = True
        Catch er As Exception
        End Try
        Return funSetUserAccess
    End Function
    Private Sub txtDocNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocNo._MYNavigator
        Try
            Dim qry As String = "select count(*) from TSPL_BOOKING_MATSER where Document_No='" & txtDocNo.Value & "' and comp_code='" & objCommonVar.CurrentCompanyCode & "' and TSPL_BOOKING_MATSER.From_Screen_code='" & clsUserMgtCode.frmDairyBookingCustomer & "'"
            Dim check As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
            If check > 0 Then
                txtDocNo.MyReadOnly = True
            Else
                txtDocNo.MyReadOnly = False
            End If
            LoadData(txtDocNo.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    ' Ticket : TEC/05/09/19-001000 By Prabhakar
    Private Sub txtDocNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtDocNo._MYValidating
        Dim qry As String = "select distinct TSPL_BOOKING_MATSER.Document_No as DocumentNo,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No as [Invoice No],TSPL_BOOKING_MATSER.AdharNo,TSPL_BOOKING_MATSER.MobileNo,case when TSPL_SD_SALE_INVOICE_HEAD.Status=1 then 'posted' else 'Unposted' end as [Invoice Status] , CONVERT(VARCHAR(12), TSPL_BOOKING_MATSER.Document_date, 103) AS Document_date,TSPL_CUSTOMER_MASTER.Cust_Code as Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_BOOKING_MATSER.GatePass_Type as ShiftType,TSPL_BOOKING_MATSER.location_code as Location  ,case when isnull(TSPL_BOOKING_MATSER.Is_Cancelled,0)=1 then 'Cancel' when TSPL_BOOKING_MATSER.Posted=1 then 'posted' else 'Unposted' end as Posted ,case when isnull(TBL_DELIVERY_NO.Delivery_No,'')='' then NULL else TBL_DELIVERY_NO.Delivery_No end as [Delivery No],TSPL_CUSTOMER_MASTER.Cust_Category_Code as [Customer Category Code],TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY as [Booking Type],TSPL_BOOKING_MATSER.against_demandBooking_no as [Against Demand Booking No],TSPL_BOOKING_MATSER.BPL_Coupon_Code as [Coupon Code],TSPL_BOOKING_MATSER.BPL_Coupon_Date as [Coupon Date],Case When TSPL_BOOKING_MATSER.Is_Taxable=2 Then 'Taxable'Else 'Non-Taxable' End As [Taxable/Non-Taxable],TSPL_BOOKING_MATSER.Total_Amt As [Amount] from TSPL_BOOKING_MATSER" &
         " left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_DETAIL.Document_No=TSPL_BOOKING_MATSER.Document_No " &
         " left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code " &
         " left join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Against_Booking_No=TSPL_BOOKING_MATSER.Document_No left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No " &
         " left join (select distinct Document_No,isnull (Delivery_No,'') as Delivery_No from TSPL_BOOKING_DETAIL  ) as TBL_DELIVERY_NO on TBL_DELIVERY_NO.Document_No = TSPL_BOOKING_MATSER.Document_No "
        Dim whrClas As String = " From_Screen_code='" & clsUserMgtCode.frmDairyBookingCustomer & "'"
        '-------richa 17/12/2019 show customer according to customer permission Ticket No. ---------
        Dim strwherecls As String = ""
        strwherecls = Xtra.CustomerPermission()
        If clsCommon.myLen(strwherecls) > 0 Then
            whrClas += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" + strwherecls + ")"
        End If
        If ShowDemandDoc Then
            whrClas += " and TSPL_BOOKING_MATSER.Against_DemandBooking_No is null "
        End If
        Dim strDocNo As String = clsCommon.ShowSelectForm("PSBookingOrderNoFndd", qry, "DocumentNo", whrClas, txtDocNo.Value, "DocumentNo", isButtonClicked, "TSPL_BOOKING_MATSER.Document_date")
        If clsCommon.myLen(strDocNo) > 0 Then
            LoadData(strDocNo, NavigatorType.Current)
        End If
    End Sub
    Private Function CheckItemCapacityLimit(ByVal strItemCode As String, ByVal strUOM As String, ByVal qty As String, ByVal dblrows As Integer, ByVal dblcolumns As Integer) As Boolean
        Dim status As Boolean = False
        Try
            Dim strQry As String = "select TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Item_Code,TSPL_ITEM_CAPACITY_LIMIT_DETAIL.UOM,TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Qty  from TSPL_ITEM_CAPACITY_LIMIT_HEAD
left join TSPL_ITEM_CAPACITY_LIMIT_DETAIL on TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Document_No=TSPL_ITEM_CAPACITY_LIMIT_head.Document_No
                where TSPL_ITEM_CAPACITY_LIMIT_HEAD.Posted=1 and TSPL_ITEM_CAPACITY_LIMIT_HEAD.Document_No in(
select top 1 TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Document_No from TSPL_ITEM_CAPACITY_LIMIT_HEAD
left join TSPL_ITEM_CAPACITY_LIMIT_DETAIL on TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Document_No=TSPL_ITEM_CAPACITY_LIMIT_head.Document_No
where TSPL_ITEM_CAPACITY_LIMIT_head.From_Date<='" & clsCommon.GetPrintDate(txtDate.Value) & "' and TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Item_Code='" & strItemCode & "'  and 2=(Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date is null then 2 else (Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date>='" & clsCommon.GetPrintDate(txtDate.Value) & "' then 2 else 3 end) end) order by From_Date desc)
 and TSPL_ITEM_CAPACITY_LIMIT_head.From_Date<='" & clsCommon.GetPrintDate(txtDate.Value) & "' and TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Item_Code='" & strItemCode & "' and 2=(Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date is null then 2 else (Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date>='" & clsCommon.GetPrintDate(txtDate.Value) & "' then 2 else 3 end) end)"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(dt.Rows(0)("UOM")) & "'"))
                Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and tspl_unit_master.Crate_Type ='Y' "))
                Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(strUOM) & "' "))
                Dim ItemLimitCF As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(dt.Rows(0)("UOM")) & "' "))
                Dim ItemLimitCFPouch As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))
                Dim cellValue As String = clsCommon.myCstr(qty)
                If CrateConvFactor > 0 And ItemConvFactor > 0 Then
                    Dim DispatchQty As Decimal = (clsCommon.myCdbl(qty) * ItemConvFactor) / ItemLimitCF
                    If DispatchQty > clsCommon.myCdbl(dt.Rows(0)("Qty")) Then
                        gv1.Rows(dblrows).Cells(dblcolumns).Value = 0
                        Throw New Exception("The maximum allowed quantity for this item is [" + clsCommon.myCstr(clsCommon.myCDecimal(dt.Rows(0)("Qty")) / (ItemConvFactor / ItemLimitCF)) + " ] ." & Environment.NewLine & " Please reduce the quantity To [" + clsCommon.myCstr(clsCommon.myCDecimal(dt.Rows(0)("Qty")) / (ItemConvFactor / ItemLimitCF)) + "] Or less To proceed.")
                    Else
                        status = True
                    End If
                End If
            Else
                gv1.Rows(dblrows).Cells(dblcolumns).Value = 0
                Throw New Exception("Capacity Limit Not found For [ " + strItemCode + " ]")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return status
    End Function
    Private Sub FrmAPInvoiceEntry_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing AndAlso gv1.CurrentColumn Is gv1.Columns(colUnit) Then
            isCellValueChangedOpen = True
            'If chkItemwise.Checked Then
            gv1.CurrentColumn = gv1.Columns(colIName)
            OpenItemUOMList(True)
            gv1.CurrentColumn = gv1.Columns(colUnit)
            'Else
            '    gv1.CurrentColumn = gv1.Columns(colIName)
            '    OpenItemGrpUOMList(True)
            '    gv1.CurrentColumn = gv1.Columns(colUnit)
            'End If
            setGridFocus()
            isCellValueChangedOpen = False
        End If
        If e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
            AddNew()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso btnSave.Enabled AndAlso MyBase.isModifyFlag Then
            SavingData(False)
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnPost.Enabled AndAlso MyBase.isPostFlag Then
            'Sanjay Ticket No- ERO/12/07/18-000371
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                If Is_Cancelled = 1 Then
                    clsCommon.MyMessageBoxShow(Me, "Booking Cancelled,Can not Post.", Me.Text)
                    Exit Sub
                End If
                'Sanjay Ticket No- BHA/03/10/18-000589
                If BookingStatus = 2 Then
                    clsCommon.MyMessageBoxShow(Me, "Booking is pending for approval,Can not Post.", Me.Text)
                    Exit Sub
                End If
                'Sanjay Ticket No- BHA/03/10/18-000589
            End If
            'Sanjay Ticket No- ERO/12/07/18-000371
            PostData()
            'skg
            'ElseIf e.Alt AndAlso e.KeyCode = Keys.F AndAlso btnCreateDO.Enabled Then
            '    btnCreateDO.PerformClick()
            '    'skg
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso btnDelete.Enabled AndAlso MyBase.isDeleteFlag Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
            CloseForm()
        ElseIf e.Control AndAlso e.KeyCode = Keys.F7 Then
            'SelectRequistionItems()
        ElseIf e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.T Then
            'chkRateDefaultSetting.Visible = Not chkRateDefaultSetting.Visible
            'chkRateUserCustomer.Visible = Not chkRateUserCustomer.Visible
            'ElseIf e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.E Then
            '    btnprinte_wayBill.Visible = True
        ElseIf e.Control AndAlso e.KeyCode = Keys.F Then
            If PanelSearchItem.Visible Then
                PanelSearchItem.Visible = False
            Else
                PanelSearchItem.Visible = True
            End If
        ElseIf e.Control AndAlso e.KeyCode = Keys.H Then
            HideOutstanding = False
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control AndAlso e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnreverse.Visible = True
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
            End If
            'Add Tool tip Task No- TEC/18/05/18-000237
            ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                                  "TSPL_BOOKING_MATSER " + Environment.NewLine +
                                  "TSPL_BOOKING_DETAIL " + Environment.NewLine +
                                  "TSPL_GATEPASS_MASTER_DAIRYSALE (For Gate Pass Document) " + Environment.NewLine +
                                  "TSPL_GATEPASS_DETAIL_DAIRYSALE (For Gate Pass Document) " + Environment.NewLine +
                                  "Press Alt+F for Create DO/Post DO Trasnaction" + Environment.NewLine +
                                  "TSPL_DELIVERY_NOTE_MASTER_FRESHSALE " + Environment.NewLine +
                                  "TSPL_DELIVERY_NOTE_DETAIL_FRESHSALE " + Environment.NewLine +
                                  "TSPL_TRANSACTION_APPROVAL (For Approving Pending Document) ")
            'Add Tool tip Task No- TEC/18/05/18-000237
        ElseIf e.KeyCode = Keys.F5 Then
            If RunBatchFifowise = 0 OrElse RunBatchFifowisewithmodifyfunctionality = True Then
                If btnPost.Enabled Then
                    OpenBatchItem()
                End If
            Else
                OpenBatchItemIfFIFIOSettingON()
            End If

        End If
    End Sub
    Private Sub BlankControlOnCustomer()
        lblVendorName.Text = ""
        LoadBlankGrid()
        'LoadBlankGridTax()
        gv1.Rows.AddNew()
    End Sub
    Private Sub GetOpeningClosingAndReceivedAmt(ByVal CustCode As String, ByVal docDate As DateTime)
        Try
            Dim OpeningBal As Decimal = 0
            Dim DrAmt As Decimal = 0
            Dim CrAmt As Decimal = 0
            Dim ClosingBal As Decimal = 0
            Dim isSkipBal As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select is_Skip_Balance from tspl_customer_Master where Cust_Code='" & CustCode & "'"))
            If isSkipBal = 0 Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("EXEC SP_GetBalCustWise @Cust_Code = '" + clsCommon.myCstr(CustCode) + "',@DocDate='" + clsCommon.GetPrintDate(docDate, "dd/MMM/yyyy HH:mm:ss") + "'")
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    OpeningBal = dt.Rows(0)("OpngBal")
                    DrAmt = dt.Rows(0)("DrAmt")
                    CrAmt = dt.Rows(0)("CrAmt")
                    ClosingBal = dt.Rows(0)("BalAmt")
                End If
                If OpeningBal > 0 Then
                    txtOpeningbal.Text = clsCommon.myCstr(OpeningBal) & " DR"
                Else
                    txtOpeningbal.Text = clsCommon.myCstr(OpeningBal) & " CR"
                End If
                If DrAmt > 0 Then
                    txtDrAmt.Text = clsCommon.myCstr(DrAmt) & " DR"
                Else
                    txtDrAmt.Text = clsCommon.myCstr(DrAmt) & " CR"
                End If
                If CrAmt > 0 Then
                    txtCrAmt.Text = clsCommon.myCstr(CrAmt) & " CR"
                Else
                    txtCrAmt.Text = clsCommon.myCstr(CrAmt) & " DR"
                End If
                If ClosingBal > 0 Then
                    txtClosingBal.Text = clsCommon.myCstr(ClosingBal) & " DR"
                Else
                    txtClosingBal.Text = clsCommon.myCstr(ClosingBal) & " CR"
                End If
            End If


        Catch ex As Exception

        End Try
    End Sub
    Private Sub txtVendorNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtVendorNo._MYValidating
        Dim qry As String = ""
        Dim docdate As DateTime = txtDate.Value
        BlankControlOnCustomer()
        If SettTagMultipleRouteWithCustomer Then
            If clsCommon.myLen(txtRouteNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please first select route", Me.Text)
                txtRouteNo.Focus()
                Exit Sub
            End If
            qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER.Alies_Name as [Short Name],TSPL_CUSTOMER_MASTER.Route_No"
            qry += " ,TSPL_CUSTOMER_MASTER.Zone_Code as [Zone Code],TSPL_ZONE_MASTER.Description as [Zone Name]"
            qry += " from TSPL_CUSTOMER_MASTER "
            'qry += " inner join TSPL_Customer_Route_Master on TSPL_Customer_Route_Master.cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code"
            qry += " left join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_CUSTOMER_MASTER.Zone_Code"
            If chkDCS.Checked Then
                qry += " inner join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code"
            End If
            Dim WhrCls As String = " 2=2 "
            If clsCommon.myLen(txtRouteNo.Value) > 0 Then
                WhrCls += "  and TSPL_CUSTOMER_MASTER.Status='N' and TSPL_Customer_Route_Master.Route_No='" & txtRouteNo.Value & "' "
            Else
                WhrCls += "  and  TSPL_CUSTOMER_MASTER.Status='N' "
            End If
            If chkDistributor.Checked Then
                WhrCls += " and TSPL_CUSTOMER_MASTER.IsDistributor='Y'"
            End If
            txtVendorNo.Value = clsCommon.ShowSelectForm("CustomerFnder1", qry, "Code", WhrCls, txtVendorNo.Value, "Code", isButtonClicked)
            lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
            setRouteDetail(txtVendorNo.Value, txtRouteNo.Value)
            gv1.DataSource = Nothing
        Else
            If chkDCS.Checked Then
                qry = "select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [Uploader Code],TSPL_CUSTOMER_MASTER.Cust_Code as Code,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER.Alies_Name as [Short Name],TSPL_CUSTOMER_MASTER.Route_No"
            Else
                qry = "select TSPL_CUSTOMER_MASTER.Cust_Code as Code,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER.Alies_Name as [Short Name],TSPL_CUSTOMER_MASTER.Route_No"
            End If
            qry += " ,TSPL_CUSTOMER_MASTER.Zone_Code as [Zone Code],TSPL_ZONE_MASTER.Description as [Zone Name],isnull(TSPL_CUSTOMER_MASTER.Customer_Category,'') as [Customer Category],tspl_customer_master.Cust_Group_Code as [Customer Group Code] "
            qry += " from TSPL_CUSTOMER_MASTER "
            qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code"
            qry += " left join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_CUSTOMER_MASTER.Zone_Code"
            If chkDCS.Checked Then
                qry += " inner join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code 
                          left join TSPL_VLC_MASTER_HEAD on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code"
            End If
            Dim WhrCls As String = ""
            If clsCommon.myLen(txtRouteNo.Value) > 0 Then
                WhrCls = "  TSPL_CUSTOMER_MASTER.Status='N' and TSPL_CUSTOMER_MASTER.Route_No='" & txtRouteNo.Value & "' "
            Else
                WhrCls = "  TSPL_CUSTOMER_MASTER.Status='N' "
            End If
            If chkDistributor.Checked Then
                WhrCls += " and TSPL_CUSTOMER_MASTER.IsDistributor='Y'"
            End If
            '-------richa 17/12/2019 show customer according to custoer permission Ticket No. ---------
            Dim strwherecls As String = ""
            strwherecls = Xtra.CustomerPermission()
            'If clsCommon.myLen(strwherecls) > 0 Then
            '    WhrCls += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" + strwherecls + ")"
            'End If
            'If chkDCS.Checked Then
            '    WhrCls += " and TSPL_CUSTOMER_MASTER.Cust_Group_Code='DCS'"
            'End If
            txtVendorNo.Value = clsCommon.ShowSelectForm("CustomerFnder1", qry, "Code", WhrCls, txtVendorNo.Value, "Code", isButtonClicked)
            lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
            lblBoothStation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(Zone_Code,'') +  case when len(oldname )>0 then ', ' +isnull(oldname,'') else '' end from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
            If chkDCS.Checked Then
                txtLastCollectionDate.Text = clsCommon.GetPrintDate(clsDBFuncationality.getSingleValue("select top 1 DOC_DATE from TSPL_MILK_SRN_HEAD where VSP_CODE='" & txtVendorNo.Value & "' order by TSPL_MILK_SRN_HEAD.DOC_DATE desc"))
            End If
            setRouteDetail(txtVendorNo.Value, txtRouteNo.Value)
            gv1.DataSource = Nothing
            If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                ''richa VIJ/18/12/19-000123
                ' CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
                'If Not HideOutstanding Then
                If chkDCS.Checked Then
                    lblOutstandingDesc.Text = clsCommon.myCstr(clsDairyInvoice.GetOutStandingBal(txtVendorNo.Value, docdate))
                Else
                    CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
                End If
                GetUnbilledAmt(docdate, txtVendorNo.Value)
                'End If
                GetOpeningClosingAndReceivedAmt(txtVendorNo.Value, docdate)
                'If chkDCS.Checked Then
                '    LoadDCSData(txtVendorNo.Value, docdate)
                'Else
                '    LoadData(txtVendorNo.Value, docdate)
                'End If
                If ShowBookingTypeDropDownonDairyBookingCustomer Then
                    LoadBlankGrid()
                    LoadBookingType()
                    LoadItemsFromPreviousBooking()
                End If
                'Vijaya, Load booking if already exist
                If clsCommon.myLen(txtDocNo.Value) <= 0 AndAlso Not chkGatePass.Checked Then
                    Dim isDemandDoc = clsDBFuncationality.getSingleValue("select top 1 TSPL_BOOKING_MATSER.Against_DemandBooking_No from TSPL_BOOKING_MATSER left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No where convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & clsCommon.GetPrintDate(docdate) & "',103) and TSPL_BOOKING_DETAIL.Cust_Code='" & txtVendorNo.Value & "'")
                    If clsCommon.myLen(isDemandDoc) > 0 Then
                        Dim STRSQL As String = "select TSPL_BOOKING_DETAIL.Document_No from TSPL_BOOKING_DETAIL left join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.cust_code=TSPL_BOOKING_DETAIL.cust_code where TSPL_BOOKING_MATSER.From_Screen_Code ='BOOK-DS-CU' and TSPL_BOOKING_MATSER.location_code ='" & txtLocation.Value & "' and TSPL_BOOKING_DETAIL.Cust_Code='" & txtVendorNo.Value & "' AND convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & txtDate.Value & "',103) and TSPL_CUSTOMER_MASTER.customer_category<>'Others'"
                        Dim TempBookingExist As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(STRSQL))
                        If clsCommon.myLen(TempBookingExist) > 0 Then
                            LoadData(TempBookingExist, NavigatorType.Current)
                        End If
                    End If
                End If
                'Load booking if already exist
                'RadPageView1.SelectedPage = RadPageViewPage2
            Else
                txtVendorNo.Focus()
            End If
        End If
    End Sub
    Sub setRouteDetail(ByVal strVendorCode As String, ByVal strtRouteCode As String)
        Dim qry As String = ""
        If SettTagMultipleRouteWithCustomer Then
            qry = "select TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_ROUTE_MASTER.vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_CUSTOMER_MASTER.Zone_Code,TSPL_Customer_Route_Master.Route_No,TSPL_VEHICLE_MASTER.Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon " + Environment.NewLine +
            "from TSPL_CUSTOMER_MASTER" + Environment.NewLine +
            "inner join TSPL_Customer_Route_Master on TSPL_Customer_Route_Master.cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code" + Environment.NewLine +
            "left outer join TSPL_ROUTE_MASTER on TSPL_Customer_Route_Master.Route_No=TSPL_ROUTE_MASTER.Route_No " + Environment.NewLine +
            "left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id " + Environment.NewLine +
            "where TSPL_CUSTOMER_MASTER.Cust_Code='" & strVendorCode & "' and TSPL_Customer_Route_Master.Route_No='" & strtRouteCode & "' "
        Else
            qry = "select Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_CUSTOMER_MASTER.Zone_Code,TSPL_CUSTOMER_MASTER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER left outer join " &
            "TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join TSPL_VEHICLE_MASTER on " &
            "TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where Cust_Code='" & strVendorCode & "'"
        End If
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
            strVehicleCode = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
            lblvehiclecode.Text = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
            strVehicleDesc = clsCommon.myCstr(dt1.Rows(0)("Number"))
            lblvehicleName.Text = clsCommon.myCstr(dt1.Rows(0)("Number"))
            strRoutecode = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
            lblroutecode.Text = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
            txtRouteCode1.Text = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
            strRouteDesc = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
            lblroutename.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
            txtRouteName1.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
            txtRouteNo.Value = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
            lblRouteDesc.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
            ' If EnableLocation Then
            'txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Route_Master where Route_No='" & txtRouteNo.Value & "' "))
            '    txtLocation.Enabled = False
            'Else
            'txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "' "))
            'End If
            'If clsCommon.myLen(txtLocation.Value) > 0 Then
            '    lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
            'End If
            If chkDCS.Checked Then
                If clsCommon.CompairString(cmbcashcredit.Text, "CASH") = CompairStringResult.Equal Then
                    lblPriceCodeDesc.Text = clsDBFuncationality.getSingleValue("select VSP_Price_Code_Cash from TSPL_customer_group_master where Default_VSP=1")
                    'Price_code = lblPriceCodeDesc.Text
                    txtPriceCode.Text = lblPriceCodeDesc.Text
                ElseIf clsCommon.CompairString(cmbcashcredit.Text, "CREDIT") = CompairStringResult.Equal Then
                    lblPriceCodeDesc.Text = clsDBFuncationality.getSingleValue("select VSP_Price_Code_Credit from TSPL_customer_group_master where Default_VSP=1")
                    'Price_code = lblPriceCodeDesc.Text
                    'txtPriceCode.Text = lblPriceCodeDesc.Text
                End If
                txtRouteCode1.Enabled = True
                txtRouteName1.Enabled = True
                txtVehicleCode.Enabled = True
                txtVehicleName.Enabled = True
            Else
                'Price_code = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
                txtPriceCode.Text = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
                lblPriceCodeDesc.Text = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
                txtVehicleCode.Value = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
                txtVehicleName.Text = clsCommon.myCstr(dt1.Rows(0)("Number"))
                txtRouteCode1.Enabled = False
                txtRouteName1.Enabled = False
                txtVehicleCode.Enabled = True
                txtVehicleName.Enabled = True
                fndTransporter.Value = clsDBFuncationality.getSingleValue("select Transport_Id from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicleCode.Value & "'")
                lblTransporter.Text = clsDBFuncationality.getSingleValue("  select Transporter_Name from TSPL_TRANSPORT_MASTER where Transport_Id='" & fndTransporter.Value & "'")
            End If
        End If
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                myMessages.blankValue(Me, "Booking not found to Print", Me.Text)
            Else
                LoadData(txtDocNo.Value, NavigatorType.Current)
                Export()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub Export()
        If gv1.Rows.Count > 0 Then
            ExportToExcel()
        Else
            common.clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
        End If
    End Sub
    Private Sub ExportToExcel()
        Try
            Dim strCreatedBy As String = clsDBFuncationality.getSingleValue("Select Created_By from TSPL_BOOKING_MATSER where Document_No='" & txtDocNo.Value & "'")
            Dim arrHeader As List(Of String) = New List(Of String)()
            Dim strtemp As String = "Booking No : " & txtDocNo.Value
            arrHeader.Add(strtemp)
            strtemp = "Booking Date : " & clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")
            arrHeader.Add(strtemp)
            strtemp = "Customer Name : " & clsCommon.myCstr(lblVendorName.Text)
            arrHeader.Add(strtemp)
            strtemp = "Location : " & clsCommon.myCstr(lblLocation.Text)
            arrHeader.Add(strtemp)
            strtemp = "Document Amount : " & clsCommon.myCstr(clsCommon.myCdbl(lblTotRAmt1.Text))
            arrHeader.Add(strtemp)
            strtemp = "Created By : " & strCreatedBy
            arrHeader.Add(strtemp)
            'strtemp = "Transaction Type : " + IIf(rbtn_Fresh.IsChecked = True, "Fresh", "Ambient")
            'arrHeader.Add(strtemp)
            clsCommon.MyExportToExcelGrid("Booking Entry", gv1, arrHeader, Me.Text)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub
    'Private Sub funPrint()
    '    Try
    '        Dim dtBarCode As New DataTable
    '        dtBarCode.Columns.Add("BarCodeImage", GetType(Byte()))
    '        Dim bytes() As Byte
    '        Dim BitmapConverter As System.ComponentModel.TypeConverter = System.ComponentModel.TypeDescriptor.GetConverter(clsCommon.MyBarcodeImage(txtDocNo.Value, 1, False).[GetType]())
    '        bytes = DirectCast(BitmapConverter.ConvertTo(clsCommon.MyBarcodeImage(txtDocNo.Value, 1, False), GetType(Byte())), Byte())
    '        Dim Qry As String
    '        Dim FooterText As String
    '        Dim frm As New frmPurchaseOrder
    '        frm.strFormId = MyBase.Form_ID
    '        Qry = ""
    '        'Qry = "select Footer_Text  from TSPL_Crystal_Report_Footer_Setting where Frm_ID ='" & Form_ID & "' "
    '        'Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(Qry)
    '        'FooterText = dt1.Rows(0).Item("Footer_Text")
    '        FooterText = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Footer_Text  from TSPL_Crystal_Report_Footer_Setting where Frm_ID ='" & Form_ID & "'"))
    '        Qry = " select TSPL_BOOKING_DETAIL_PRODUCTSALE.line_no,'" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE, "dd/MM/yyyy") & "' as RunDate,TSPL_BOOKING_MASTER_PRODUCTSALE.Cust_PO_No,'" & FooterText & "' as FooterText , TSPL_BOOKING_MASTER_PRODUCTSALE.Abandonment_No, TSPL_BOOKING_MASTER_PRODUCTSALE.Dept_Desc ,TSPL_BOOKING_MASTER_PRODUCTSALE.Delivery_date ,TSPL_BOOKING_MASTER_PRODUCTSALE.Remarks ,(case when TSPL_BOOKING_MASTER_PRODUCTSALE .status =1 then TSPL_BOOKING_MASTER_PRODUCTSALE.Modify_By else '' end) as Modify_By ,TSPL_BOOKING_DETAIL_PRODUCTSALE.Landing_Cost,TSPL_BOOKING_MASTER_PRODUCTSALE.Created_By , "
    '        Qry += " TSPL_BOOKING_MASTER_PRODUCTSALE.Terms_Code,TSPL_SHIP_TO_LOCATION.Telphone as ShipLoc_telphone,TSPL_LOCATION_MASTER.Service_Tax_Reg_No ,TSPL_SHIP_TO_LOCATION.Ship_To_Type_Desc as ShipLoc_Name,TSPL_SHIP_TO_LOCATION.email as shipLoc_Email,TSPL_LOCATION_MASTER.cst_no as CST_LST,TSPL_LOCATION_MASTER.Registration_Number,TSPL_LOCATION_MASTER.Add1 as loc_add1 ,TSPL_LOCATION_MASTER.add4 as loc_add4,TSPL_LOCATION_MASTER.Add2 as loc_add2,TSPL_LOCATION_MASTER.Add3 as loc_add3,TSPL_LOCATION_MASTER.Email as loc_email,TSPL_LOCATION_MASTER.TAN_No as loc_Fax , TSPL_BOOKING_DETAIL_PRODUCTSALE.TAX1_Rate,TSPL_BOOKING_DETAIL_PRODUCTSALE.PrincipleDesc ,TSPL_COMPANY_MASTER.Add2 as comp_Add2 ,TSPL_COMPANY_MASTER.Add3 as comp_Add3 ,ISNULL(TSPL_COMPANY_MASTER.Phone1,'')+ Case When ISNULL(TSPL_COMPANY_MASTER.Phone2,'')<>'' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as Company_Phone,TSPL_COMPANY_MASTER.Email as Comp_Email,TSPL_COMPANY_MASTER.Fax as comp_Fax,ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')+ Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'' Then ', '+ TSPL_CUSTOMER_MASTER.Phone2 Else'' End as Cust_Phone ,TSPL_CUSTOMER_MASTER.Email as Cust_Email,TSPL_CUSTOMER_MASTER.Fax as cust_Fax,(TSPL_BOOKING_DETAIL_PRODUCTSALE.Total_Tax_Amt /(case when TSPL_BOOKING_DETAIL_PRODUCTSALE.Amt_Less_Discount<=0 then 1 else TSPL_BOOKING_DETAIL_PRODUCTSALE.Amt_Less_Discount end)) *100 as Tax,((((TSPL_BOOKING_DETAIL_PRODUCTSALE.Total_Tax_Amt /(case when TSPL_BOOKING_DETAIL_PRODUCTSALE.Amt_Less_Discount<=0 then 1 else TSPL_BOOKING_DETAIL_PRODUCTSALE.Amt_Less_Discount end)) *100)*TSPL_BOOKING_DETAIL_PRODUCTSALE.Item_Cost/100) +TSPL_BOOKING_DETAIL_PRODUCTSALE.Item_Cost) as landing_Rate,(TSPL_BOOKING_DETAIL_PRODUCTSALE.Total_Tax_Amt /(case when TSPL_BOOKING_DETAIL_PRODUCTSALE.Amt_Less_Discount<=0 then 1 else TSPL_BOOKING_DETAIL_PRODUCTSALE.Amt_Less_Discount end) *100) as Tax_Lan,TSPL_BOOKING_MASTER_PRODUCTSALE.Mode_Of_Transport as  ModeofTransport,TSPL_BOOKING_DETAIL_PRODUCTSALE .Specification as  specification,  "
    '        Qry += " TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code as DocNo ,TSPL_SHIP_TO_LOCATION.add1 +case when len(TSPL_SHIP_TO_LOCATION.add2)>0 then ', '+TSPL_SHIP_TO_LOCATION.add2 else '' end +case when LEN(isnull(TSPL_SHIP_TO_LOCATION.Add3,''))>0 then ', '+isnull(TSPL_SHIP_TO_LOCATION.Add3,'') else ' ' end  + case when len(TSPL_SHIP_TO_LOCATION.State )>0 then TSPL_SHIP_TO_LOCATION.State else '' end  as Ship_address, TSPL_SHIP_TO_LOCATION.Tin_No as Ship_ToLocation_Tin_No, convert(varchar(10),TSPL_BOOKING_MASTER_PRODUCTSALE .Document_Date,103)  as po_date ,case TSPL_BOOKING_MASTER_PRODUCTSALE .SalesOrder_Type when 'L'then 'Local' when 'I' then 'Import' when 'J' then 'Job Work' when 'O' then 'Open' when 'S' then 'Specific'else 'Null' end as po_type , "
    '        Qry += " TSPL_BOOKING_MASTER_PRODUCTSALE.Customer_Code, TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_BOOKING_DETAIL_PRODUCTSALE.Item_Net_Amt as Total_Item_Net_amt,TSPL_BOOKING_DETAIL_PRODUCTSALE.MRP ,TSPL_BOOKING_DETAIL_PRODUCTSALE.Item_Cost, TSPL_BOOKING_MASTER_PRODUCTSALE .Terms_Code as termscode ,TSPL_BOOKING_MASTER_PRODUCTSALE .Ref_No as ref_no ,TSPL_BOOKING_MASTER_PRODUCTSALE .Comments as comments , "
    '        Qry += " TSPL_BOOKING_MASTER_PRODUCTSALE .Discount_Amt as dis_amt,TSPL_BOOKING_DETAIL_PRODUCTSALE .Disc_Amt  as dis_amt1,TSPL_BOOKING_MASTER_PRODUCTSALE.Amount_Less_Discount  as aftrdiscount ,TSPL_BOOKING_MASTER_PRODUCTSALE .Total_Amt as Total_amount,TSPL_BOOKING_MASTER_PRODUCTSALE.Discount_Base as bfrdisc_amount, "
    '        Qry += " tax1.Tax_Code_Desc as tax1name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax1_amt,0) as txt1amt, "
    '        Qry += " tax2.Tax_Code_Desc as tax2name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax2_amt,0) as txt2amt, "
    '        Qry += " tax3.Tax_Code_Desc as tax3name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax3_amt,0) as txt3amt, "
    '        Qry += " tax4.Tax_Code_Desc as tax4name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax4_amt,0) as txt4amt, "
    '        Qry += " tax5.Tax_Code_Desc as tax5name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax5_amt,0) as txt5amt, "
    '        Qry += " tax6.Tax_Code_Desc as tax6name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax6_amt,0) as txt6amt, "
    '        Qry += " tax7.Tax_Code_Desc as tax7name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax7_amt,0) as txt7amt, "
    '        Qry += " tax8.Tax_Code_Desc as tax8name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax8_amt,0) as txt8amt,  "
    '        Qry += " tax9.Tax_Code_Desc as tax9name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax9_amt,0) as txt9amt, "
    '        Qry += " tax10.Tax_Code_Desc as tax10name,isnull (TSPL_BOOKING_MASTER_PRODUCTSALE.tax10_amt,0) as txt10amt, "
    '        Qry += " isnull(TSPL_BOOKING_MASTER_PRODUCTSALE .Total_Tax_Amt,0) as total_tax_amt, TSPL_BOOKING_MASTER_PRODUCTSALE.Total_Amt as DocAmt,  TSPL_COMPANY_MASTER.Comp_Name as compname,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,ISNULL(tspl_company_Master.ADD1,'') as address1,TSPL_BOOKING_DETAIL_PRODUCTSALE.item_code as item_code, TSPL_ITEM_MASTER.Item_Desc   as itemdesc, TSPL_BOOKING_DETAIL_PRODUCTSALE.Row_Type,TSPL_BOOKING_DETAIL_PRODUCTSALE.Qty as qty,TSPL_BOOKING_DETAIL_PRODUCTSALE.scheme_item as freeitem,TSPL_BOOKING_DETAIL_PRODUCTSALE.unit_code as uom,TSPL_BOOKING_DETAIL_PRODUCTSALE.item_cost as itemcost,TSPL_BOOKING_DETAIL_PRODUCTSALE.amount as amount ,"
    '        Qry += " TSPL_LOCATION_MASTER.TIN_No ,ISNULL(TSPL_LOCATION_MASTER.Phone1,'')+ Case When ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as Location_Phone ,(Select Add1+Case When ISNULL(Add2,'')='' Then '' else ', '+Add2+ Case When ISNULL(Add3,'')='' Then '' Else ', '+Add3+ Case When ISNULL(Pin_Code,'')='' Then '' else '-'+CONVERT(varchar, Pin_Code) End End End from TSPL_LOCATION_MASTER Where TSPL_LOCATION_MASTER.Location_Code = TSPL_BOOKING_MASTER_PRODUCTSALE.Bill_To_Location ) as [BillToAddress],TSPL_CUSTOMER_MASTER.add1 +case when len(TSPL_CUSTOMER_MASTER.add2)>0 then ', '+TSPL_CUSTOMER_MASTER.add2 else '' end +case when LEN(isnull(TSPL_CUSTOMER_MASTER.Add3,''))>0 then ', '+isnull(TSPL_CUSTOMER_MASTER.Add3,'') else ' ' end  + case when len(TSPL_CUSTOMER_MASTER.State )>0 then TSPL_CUSTOMER_MASTER.State else '' end  as Cust_address,TSPL_CUSTOMER_MASTER.Lst_No as cust_Lst_No,TSPL_CUSTOMER_MASTER.Tin_No as Cust_Tin_No  ,TSPL_CUSTOMER_MASTER.CST as Cust_Cst,"
    '        Qry += " (Select Add1+Case When ISNULL(Add2,'')='' Then '' else ', '+Add2+ Case When ISNULL(Add3,'')='' Then '' Else ', '+Add3+ Case When ISNULL(Pin_Code,'')='' Then '' else '-'+CONVERT(varchar, Pin_Code) End End End from TSPL_LOCATION_MASTER Where TSPL_LOCATION_MASTER.Location_Code = TSPL_BOOKING_MASTER_PRODUCTSALE.Ship_To_Location ) as [ShipToAddress], TSPL_BOOKING_MASTER_PRODUCTSALE.TAX1,TSPL_BOOKING_MASTER_PRODUCTSALE.TAX2,TSPL_BOOKING_MASTER_PRODUCTSALE.TAX3,TSPL_BOOKING_MASTER_PRODUCTSALE.TAX4,TSPL_BOOKING_MASTER_PRODUCTSALE.TAX5,TSPL_BOOKING_MASTER_PRODUCTSALE.Total_Add_Charge "
    '        Qry += " from TSPL_BOOKING_DETAIL_PRODUCTSALE "
    '        Qry += " left outer join TSPL_BOOKING_MASTER_PRODUCTSALE  on TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code  =TSPL_BOOKING_DETAIL_PRODUCTSALE.Document_Code  "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_BOOKING_MASTER_PRODUCTSALE.tax1 "
    '        Qry += " left outer join  TSPL_SHIP_TO_LOCATION on TSPL_BOOKING_MASTER_PRODUCTSALE.Ship_To_Location =TSPL_SHIP_TO_LOCATION.Ship_To_Code "
    '        Qry += " left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_BOOKING_MASTER_PRODUCTSALE.tax2 "
    '        Qry += " left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_BOOKING_MASTER_PRODUCTSALE .TAX3 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_BOOKING_MASTER_PRODUCTSALE .tax4 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_BOOKING_MASTER_PRODUCTSALE .tax5 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_BOOKING_MASTER_PRODUCTSALE .TAX6 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_BOOKING_MASTER_PRODUCTSALE .TAX7 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_BOOKING_MASTER_PRODUCTSALE .TAX8 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_BOOKING_MASTER_PRODUCTSALE .TAX9 "
    '        Qry += " left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_BOOKING_MASTER_PRODUCTSALE .TAX10    "
    '        Qry += " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_BOOKING_MASTER_PRODUCTSALE.comp_code "
    '        Qry += " left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_BOOKING_MASTER_PRODUCTSALE.Customer_Code  "
    '        Qry += " left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_BOOKING_MASTER_PRODUCTSALE.Bill_To_Location"
    '        Qry += " Left Outer Join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_BOOKING_DETAIL_PRODUCTSALE.Item_Code  where 2=2 "
    '        ''If txtReqNo.Value <> "" Then            Why you make check here of ReqNo--Balwinder to Pankaj
    '        Qry += "  and  TSPL_BOOKING_MASTER_PRODUCTSALE.Document_Code = '" & txtDocNo.Value & "'"
    '        '' End If
    '        'Qry = "Select *, Case When ISNULL(ShipToAddress,'')='' Then BillToAddress Else ShipToAddress End as [ShipToAddress] from ( " + Qry + " ) XXX"
    '        Qry = "Select *, Case When ISNULL(ShipToAddress,'')='' Then BillToAddress Else ShipToAddress End as [ShipToAddress] from ( " + Qry + " ) XXX,(select document_code,sum(TAX1_Amt) as t1,sum(TAX2_Amt) as t2,sum(TAX3_Amt) as t3,sum(TAX4_Amt) as t4,sum(TAX5_Amt) as t5,sum(TAX6_Amt) as t6,sum(TAX7_Amt) as t7,sum(TAX8_Amt) as t8,sum(TAX9_Amt) as t9,sum(TAX10_Amt) as t10,sum(Amount) as detamt,(sum(Disc_Amt)+sum(cust_discount)) as Disc_Amt,sum(Amt_Less_Discount) as Amt_Less_Discount,sum(Total_Tax_Amt) as totaltaxamt,sum(Item_Net_Amt) as Item_Net_Amt,sum(Total_Basic_Amt) as Total_Basic_Amt from TSPL_BOOKING_DETAIL_PRODUCTSALE where Scheme_Item='N' group by document_code)b where docno=b.document_code order by xxx.line_no"
    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
    '        dt.Columns.Add("BarCodeImage", GetType(Byte()))
    '        For Each dr As DataRow In dt.Rows
    '            dr("BarCodeImage") = bytes
    '        Next
    '        If dt.Rows.Count > 0 Then
    '            'SetItemWiseTax(dt, txtDocNo.Value)
    '            Dim frmCRV As New frmCrystalReportViewer()
    '            frmCRV.funreport(CrystalReportFolder.NewSalesReports, dt, "crptSalesOrderReport", "Sales Order")
    '            frmCRV = Nothing
    '        End If
    '    Catch ex As Exception
    '        common.clsCommon.MyMessageBoxShow(ex.Message)
    '    End Try
    'End Sub
    Private Sub gv1_UserDeletedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserDeletedRow
        UpdateAllTotals()
        For ii As Integer = 1 To gv1.Rows.Count
            gv1.Rows(ii - 1).Cells(colLineNo).Value = ii
        Next
        'RefreshReqNo()
    End Sub
    Private Sub gv1_UserDeletingRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        If common.clsCommon.MyMessageBoxShow(Me, "Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
            e.Cancel = True
            'ElseIf clsCommon.myCBool(gv1.CurrentRow.Cells(colItemUsedINGRN).Value) Then
            '    common.clsCommon.MyMessageBoxShow("Can't Delete The Current Row.This Item is Used In GRN")
            '    e.Cancel = True
        End If
    End Sub
    'Private Sub gv1_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv1.CellFormatting
    '    Try
    '    Catch ex As Exception
    '        'common.clsCommon.MyMessageBoxShow(ex.Message)
    '    End Try
    'End Sub
    Private Sub setGridFocus()
        Dim intCurrRow As Integer = gv1.CurrentRow.Index
        gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
        If intCurrRow = gv1.Rows.Count - 1 Then
            gv1.Rows.AddNew()
            'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
            gv1.CurrentRow = gv1.Rows(intCurrRow)
        End If
        If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
            If gv1.CurrentColumn Is gv1.Columns(colICode) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colQty)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colQty) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colRate)
            End If
        End If
    End Sub
    Private Sub PrintAmendment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'If clsCommon.myLen(txtDocNo.Value) <= 0 Then
        '    myMessages.blankValue("Purchase Order No not found to Print")
        'Else
        '    FrmPurchaseOrderReport.PrintAbandoment(txtDocNo.Value)
        'End If
    End Sub
    'Private Sub txtDocNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocNo.Load
    'End Sub
    'Private Sub btnAmendment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'End Sub
    'Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ''Printing the amendment
    '    'If clsCommon.myLen(txtDocNo.Value) <= 0 Then
    '    '    common.clsCommon.MyMessageBoxShow("Purchase Order No not found to Print")
    '    'Else
    '    '    FrmPurchaseOrderReport.PrintAbandoment(txtDocNo.Value)
    '    'End If
    'End Sub
    Dim i As Integer
    Private Sub gv1_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gv1.CurrentColumnChanged
        If gv1.RowCount > 0 Then
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
            If intCurrRow = gv1.Rows.Count - 1 Then
                gv1.Rows.AddNew()
                'gv1.Rows(gv1.Rows.Count - 1).Cells(colRowType).Value = RowTypeItem
                gv1.CurrentRow = gv1.Rows(intCurrRow)
            End If
        End If
    End Sub
    Private Sub gv1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles gv1.Click
        If gv1.CurrentRow IsNot Nothing Then
            'setBalance()
        End If
    End Sub
    Private Sub gv1_CurrentRowChanged(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.CurrentRowChangedEventArgs) Handles gv1.CurrentRowChanged
        'If gv1.CurrentRow IsNot Nothing AndAlso Not e.CurrentRow.Index < 0 Then
        '    setBalance()
        'End If
    End Sub
    'Private Sub txtDiscAmt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscAmt.Click
    'End Sub
    'Private Sub txtDiscPer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiscPer.Click
    'End Sub
    'Private Sub RadLabel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadLabel1.Click
    'End Sub
    '--------------------------------------BM00000002443 Done By Monika 30/04/2014----------------------------------'
#Region "New Mail System"
    Private Sub RadMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem3.Click
        Dim frm As New FrmMailSMSSettingNew2()
        frm.FormId = clsUserMgtCode.frmDairyBookingCustomer
        frm.ShowDialog()
    End Sub
    Private Sub btnpreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please Select Document No. First", Me.Text)
            txtDocNo.Focus()
            txtDocNo.Select()
            Return
        End If
        'attachqry = GetAtchmentPrintQuery(txtDocNo.Value)
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(attachqry)
        If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            System.Diagnostics.Process.Start(frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt1, "crptSalesOrderReport", "Sales Order"))
            frmCRV = Nothing
        End If
    End Sub
    'Private Sub btnsend_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
    '            clsCommon.MyMessageBoxShow("Please Select Document No. First", Me.Text)
    '            'txtReqNo.Focus()
    '            'txtReqNo.Select()
    '            Return
    '        End If
    '        If Not (common.clsCommon.MyMessageBoxShow("Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
    '            Return
    '        End If
    '        LoadData(txtDocNo.Value, NavigatorType.Current)
    '        Dim lstUsers As New List(Of String)
    '        lstUsers.Add(txtVendorNo.Value)
    '        SendSMSandEmail(lstUsers, False)
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(ex.Message)
    '    End Try
    'End Sub
    'Private Sub btnSendForApproval_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
    '            clsCommon.MyMessageBoxShow("Please Select Document No. First", Me.Text)
    '            'txtReqNo.Focus()
    '            'txtReqNo.Select()
    '            Return
    '        End If
    '        If Not (common.clsCommon.MyMessageBoxShow("Send E-Mail/SMS Of Respective Booking No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
    '            Return
    '        End If
    '        LoadData(txtDocNo.Value, NavigatorType.Current)
    '        Dim qry As String = "Select * from TSPL_APPROVAL_LEVEL_SCREEN where 1=1 and Trans_Code='" & MyBase.Form_ID & "' "
    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '        Dim lstUsers As New List(Of String)
    '        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '            For Each dr As DataRow In dt.Rows
    '                lstUsers.Add(dr("User_Code").ToString())
    '            Next
    '        End If
    '        If lstUsers.Count = 0 Then
    '            Throw New Exception("No Receiptent Found")
    '        End If
    '        SendSMSandEmail(lstUsers, True)
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(ex.Message)
    '    End Try
    'End Sub
    'Private Sub SendSMSandEmail(ByVal lstUsers As List(Of String), ByVal isSendForApproval As Boolean)
    '    Try
    '        Dim obj As clsEmailSMSSettingNew = clsEmailSMSSettingNew.GetData(clsUserMgtCode.frmSNSalesOrder)
    '        If obj Is Nothing Then
    '            clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
    '            Return
    '        End If
    '        If clsCommon.myLen(obj.mailsubjct) <= 0 Then
    '            clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
    '            Return
    '        End If
    '        Dim strContactPerson As String = ""
    '        Dim strSubject As String = obj.mailsubjct.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
    '        strSubject = strSubject.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
    '        Dim strbody As String = obj.mailbody.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
    '        strbody = strbody.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
    '        strbody = strbody.Replace(clsEmailSMSConstants.CustomerNo, txtVendorNo.Value)
    '        strbody = strbody.Replace(clsEmailSMSConstants.CustomerName, lblVendorName.Text)
    '        strbody = strbody.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
    '        strbody = strbody.Replace(clsEmailSMSConstants.Form_Code, MyBase.Form_ID)
    '        '------------------------code for attchament-------------------------------------
    '        Dim strRptPath As String = ""
    '        If obj.atchmnt = "Y" Then
    '            'attachqry = GetAtchmentPrintQuery(txtDocNo.Value)
    '            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(attachqry)
    '            If dt1.Rows.Count > 0 Then
    '                'SetItemWiseTax(dt1, txtDocNo.Value)
    '                Dim frmCRV As New frmCrystalReportViewer()
    '                strRptPath = frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt1, "crptSalesOrderReport", "Sales Order")
    '                frmCRV = Nothing
    '            End If
    '        End If
    '        '---------------------------------------------------------------------------
    '        For Each strUser As String In lstUsers
    '            'lstUsers.Add(dr("User_Code").ToString())
    '            Dim lstReceiptents As New List(Of String)
    '            Dim qry As String = ""
    '            Dim emailId As String = ""
    '            If isSendForApproval Then
    '                strContactPerson = strUser
    '                qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" & strUser & "') "
    '                emailId = clsDBFuncationality.getSingleValue(qry)
    '            Else
    '                strContactPerson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
    '                emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
    '            End If
    '            strbody = strbody.Replace(clsEmailSMSConstants.ContactPerson, strContactPerson)
    '            lstReceiptents.Add(emailId)
    '            Dim body As String = strbody.Replace(clsEmailSMSConstants.UserCode, strUser)
    '            clsMailViaOutlook.SendEmail(strSubject, body, lstReceiptents, Nothing, strRptPath)
    '        Next
    '        clsCommon.MyMessageBoxShow("E-Mail Send Successfully", Me.Text)
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Sub
    'Public Sub SetMailRight()
    'End Sub
#End Region
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
            ''stuti regarding memory leakage
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub
    Private Sub RadMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem4.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub
    Private Sub RadMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem5.Click
        Dim frm As New FrmCrptFooter
        frm.strFormId = MyBase.Form_ID
        frm.ShowDialog()
    End Sub
    Private Sub btnHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New FrmSaleHistory
        frm.strFormId = MyBase.Form_ID
        frm.strCustId = txtVendorNo.Value
        frm.strCustName = lblVendorName.Text
        'Dim strvendor As String = txtVendorNo.Value
        frm.ShowDialog()
        frm.WindowState = FormWindowState.Maximized
    End Sub
    'Private Sub lblTotRAmt1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'End Sub
    'Private Sub txtDate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs)
    '    'SETGSTControl()
    'End Sub
    'Private Function txtRefNo() As Object
    '    Throw New NotImplementedException
    'End Function
    Private Sub txtLocation__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocation._MYValidating
        Dim qry As String = " select Location_Code as [Code],Location_Desc as [Description],Loc_Short_Name as [Short Name],Add1,Add2,Add3,Add4,City_Code as [City Code],State,Pin_Code as [Pin Code],Country,Telphone,Email,Location_Type as [Location Type],Loc_Status as [Location Status],Status_Date as [Status Date],Excisable,Loc_Segment_Code as [Location Segment Code],Type,Purchase_Tax_Group as [Purchase Tax Group],Sales_Tax_Group as [Sales Tax Group],Ecc_Number as [ECC Number],Registration_Number as [Registration Number],Commissionerate as [Commission Rate],Range_Code as [Range Code],Range_Name as [Range Name],Range_Address as [Range Address],Division_Code as [Division Code],Division_Name as [Division Name],Division_Address as [Division Address],Created_By as [Created By],Created_Date as [Created Date],Modify_By as [Modify By],Modify_Date as [Modify Date],Comp_code as [Company Code],TIN_No as [TIN No],TAN_No as [TAN No],TCAN_No as [TCAN No],Service_Tax_Reg_No as [Service Tax Registration No],DutyPaid as [Duty Paid],Purchase_Tax_GroupIS as [Purchase Tax Group Inter State],Sales_Tax_GroupIS as [Sales Tax Group Inter State],Stock_Transfer_Filled_Ac as [Stock Transfer Filled Account],Stock_Transfer_Empty_Ac as [Stock Transfer Empty Account],GIT_Location as [GIT Location],GIT_Type as [GIT Type],Rejected_Type as [Rejected Type],Rejected_Location as [Rejected Location],CSA_Type as [CSA Type],Cust_Code as [Cust Code],MCC_Type as [MCC Type],CST_No as [CST No],Phone1,Phone2  from TSPL_Location_MASTER"
        Dim WhrCls As String = " Loc_Status='N' and Location_Category not in('MCC') and Location_Type='Physical' and Is_Section='N' and Is_Sub_Location='N' and CSA_Type <>'Y' and DutyPaid <>'Y' and Rejected_Type <>'Y' and GIT_Type<>'Y' or isSaleLocation=1 "
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            WhrCls += "  and  Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
        End If
        txtLocation.Value = clsCommon.ShowSelectForm("MulDS-BOLocFndr", qry, "Code", WhrCls, txtLocation.Value, "Code", isButtonClicked)
        lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
        gv1.DataSource = Nothing
        If gv1.Rows.Count > 0 Then
            gv1.Focus()
            gv1.Rows(0).Cells(1).BeginEdit()
        End If
        If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsSubLocationWise,'N') as  IsSubLocationWise from tspl_location_master where location_code='" & clsCommon.myCstr(txtLocation.Value) & "'")), "Y") = CompairStringResult.Equal Then
                txtSubLocation.Enabled = True
            Else
                txtSubLocation.Enabled = False
            End If
            txtSubLocation.Value = ""
            lblSubLocation.Text = ""
        End If
    End Sub
    Private Function CheckItemtaxType() As Boolean
        Throw New NotImplementedException
    End Function
    'Private Sub btnCreateDO_Click(sender As Object, e As EventArgs)
    '    RecordCount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_BOOKING_MATSER "))
    '    If RecordCount = 0 Then
    '        FlagFirstRecord = True
    '    Else
    '        FlagFirstRecord = False
    '    End If
    '    If Is_Cancelled = 1 Then
    '        clsCommon.MyMessageBoxShow(Me, "Booking Cancelled,Can not Creat/Post DO", Me.Text)
    '        Exit Sub
    '    End If
    '    If BookingStatus = 1 Then
    '        clsCommon.MyMessageBoxShow(Me, "Please Post booking before creating DO", Me.Text)
    '        Exit Sub
    '    ElseIf BookingStatus = 2 Then
    '        clsCommon.MyMessageBoxShow(Me, "Please Approve and post booking before creating DO", Me.Text)
    '        Exit Sub
    '    ElseIf BookingStatus = 3 Then
    '        clsCommon.MyMessageBoxShow(Me, "Please post booking before creating DO", Me.Text)
    '        Exit Sub
    '    End If
    '    If DOStatus = 2 Then
    '        clsCommon.MyMessageBoxShow(Me, "DO is pending for approval.", Me.Text)
    '        Exit Sub
    '    End If
    '    Dim qry As String = ""
    '    FlagCreateDo = True
    '    'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
    '    'Try
    '    '    If CreateDO(False, Nothing, txtDocNo.Value) Then
    '    '        'trans.Commit()
    '    '        If clsCommon.myLen(DOmsg) > 0 Then
    '    '            common.clsCommon.MyMessageBoxShow(Me, DOmsg, Me.Text)
    '    '        End If
    '    '        If DOCreated Then
    '    '            Dim msg = "Successfully created"
    '    '            common.clsCommon.MyMessageBoxShow(Me, msg, Me.Text)
    '    '            msg = Nothing
    '    '        End If
    '    '        LoadData(txtDocNo.Value, NavigatorType.Current)
    '    '    Else
    '    '        'trans.Rollback()
    '    '    End If
    '    'Catch ex As Exception
    '    '    'trans.Rollback()
    '    '    common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    'Finally
    '    '    qry = Nothing
    '    '    FlagCreateDo = False
    '    'End Try
    'End Sub
    ''richa ERO/03/12/19-001138
    'Private Function CreateDO(ByVal ChekPostBtn As Boolean, ByVal trans As SqlTransaction, ByVal strBookingNo As String)
    '    Try
    '        blnSaveTotalQTy = True
    '        Dim qry As String = String.Empty
    '        DOmsg = String.Empty
    '        Dim dblTotal_Qty As Double = 0
    '        Dim blnRatezero As Boolean = False
    '        'lblDONumber.Text = clsDBFuncationality.getSingleValue("select isnull(Delivery_No,'') from TSPL_BOOKING_DETAIL where Document_No ='" & txtDocNo.Value & "' and cust_code='" & txtVendorNo.Value & "'", trans)
    '        If (AllowToSave(trans)) Then
    '            'For ii As Integer = 8 To gv1.Columns.Count - 1
    '            'Dim objBookingCust As clsBookingTemp = TryCast(gv1.Columns(ii).Tag, clsBookingTemp)
    '            If clsCommon.myCdbl(lblTotRAmt1.Text) > 0 AndAlso BookingStatus = 4 And BookingStatus <> 5 Then
    '                If (clsCommon.myLen(lblDONumber.Text) > 0 AndAlso DOStatus <> 4) OrElse (clsCommon.myLen(lblDONumber.Text) = 0 And clsCommon.myCdbl(lblTotRAmt1.Text) > 0) Then
    '                    Dim dblCustTotalQty As Double = 0
    '                    For Each grow As GridViewRowInfo In gv1.Rows
    '                        If clsCommon.myCdbl(grow.Cells(colQty).Value) > 0 Then
    '                            dblCustTotalQty += grow.Cells(colQty).Value
    '                        End If
    '                    Next
    '                    If clsCommon.myCdbl(dblCustTotalQty) > 0 Then
    '                        If clsCommon.myLen(lblDONumber.Text) = 0 Then
    '                            Dim obj As New clsDeliveryNoteDairySale
    '                            dblTotal_Qty = 0
    '                            'obj.Price_code = clsDBFuncationality.getSingleValue("select price_CodeNon from tspl_customer_master where cust_code='" & clsCommon.myCstr(grow.Cells(1).Value) & "'", trans)
    '                            obj.Credit_Limit = 0
    '                            obj.Document_Date = txtDate.Value
    '                            obj.Customer_Code = txtVendorNo.Value
    '                            obj.Location_Code = txtLocation.Value
    '                            dblTotal_Qty = dblCustTotalQty
    '                            obj.Sampling = IIf(chkSampling.Checked, 1, 0)
    '                            obj.Booking_No = txtDocNo.Value
    '                            obj.Booking_Date = txtDate.Value
    '                            obj.Ship_To_Location = txtShipToLocation.Value
    '                            obj.Vehicle_Capacity = 0
    '                            obj.Lorry_No = clsDBFuncationality.getSingleValue("select Vehicle_Code from TSPL_BOOKING_DETAIL where Document_No ='" & txtDocNo.Value & "' and cust_code='" & txtVendorNo.Value & "'", trans)
    '                            obj.Route_No = lblroutecode.Text
    '                            obj.Transporter_Name = obj.Lorry_No
    '                            obj.Price_code = txtPriceCode.Text
    '                            obj.Freight = ""
    '                            obj.Freight_Amount = 0
    '                            obj.Comments = ""
    '                            obj.OnHold = "N"
    '                            obj.Short_Close = "N"
    '                            obj.Total_Amt = 0
    '                            If CreateCommonDairyDispatchforFreshAmbient = 1 Then
    '                                obj.TRANSACTION_TYPE = ""
    '                            Else
    '                                obj.TRANSACTION_TYPE = IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
    '                            End If
    '                            obj.From_Screen_code = clsUserMgtCode.frmDairyBookingCustomer
    '                            obj.SalesmanCode = txtSalesman.Value
    '                            obj.Cust_PO_No = txtPONo.Text
    '                            If txtCustPODate.Checked Then
    '                                obj.Podate = txtCustPODate.Value
    '                            End If
    '                            obj.Arr = New List(Of clsDeliveryNoteDairySaleDetail)
    '                            Dim intLineNo As Integer = 1
    '                            Dim dblTotal As Double = 0
    '                            blnRatezero = False
    '                            DOCreated = False
    '                            For Each grow As GridViewRowInfo In gv1.Rows
    '                                Dim objTr As New clsDeliveryNoteDairySaleDetail()
    '                                If (clsCommon.myCdbl(grow.Cells(colQty).Value)) > 0 Then
    '                                    objTr.Line_No = intLineNo
    '                                    objTr.Sampling = 0
    '                                    'Dim objBookingitem As clsBookingTemp = TryCast(grow.Cells(ii).Tag, clsBookingTemp)
    '                                    objTr.Item_Code = clsCommon.myCstr(grow.Cells(colICode).Value)
    '                                    objTr.Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
    '                                    objTr.Booking_No = clsCommon.myCstr(txtDocNo.Value)
    '                                    objTr.Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
    '                                    objTr.BookQty = clsCommon.myCdbl(grow.Cells(colQty).Value)
    '                                    objTr.Balance_Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
    '                                    objTr.OrgUnit_code = clsCommon.myCdbl(grow.Cells(colQty).Value)
    '                                    'Dim objBookingItemRate As clsBookingTemp = TryCast(grow.Cells(ii).Tag, clsBookingTemp)
    '                                    'sanjay
    '                                    Dim dblRate As Decimal = 0
    '                                    'dblRate = objBookingitem.ItemRate
    '                                    'Dim tax_on_amt As Decimal = 0
    '                                    Dim dt As New DataTable()
    '                                    'Dim dblTotal As Double = 0
    '                                    Dim Price_code As String = lblPriceCodeDesc.Text 'clsDBFuncationality.getSingleValue("select price_CodeNon from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'", trans)
    '                                    Dim strCustName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER  where Cust_Code='" & txtVendorNo.Value & "'", trans))
    '                                    qry = "Select RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price,Item_MRP,XXXE.TAX1_Rate, XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
    '                   "XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7,XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.Against_Plan_TR_Code from ( " &
    '                   "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
    '                   "Start_Date Desc) as RowNo, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
    '                   "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price,Item_MRP ,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
    '                     "TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
    '                     "TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
    '                     "TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7,TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
    '                   "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
    '                   "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code    where  Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null) and  " &
    '                   "TSPL_ITEM_PRICE_MASTER.Price_Code='" & Price_code & "' and UOM='" & clsCommon.myCstr(grow.Cells(colUnit).Value) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(grow.Cells(colICode).Value) & "' AND Location_Code='" & clsCommon.myCstr(txtLocation.Value) & "'  "
    '                                    If ShowMulMRPOfSameItemOnDairyBookingCustomer = True AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ", trans)), "Others") = CompairStringResult.Equal Then
    '                                        qry += " and round(TSPL_ITEM_PRICE_MASTER.Item_Selling_Price,2) =" & clsCommon.myCdbl(grow.Cells(colRate).Value) & " "
    '                                    End If
    '                                    qry += ") XXXE WHERE RowNo=1  "
    '                                    dt = clsDBFuncationality.GetDataTable(qry, trans)
    '                                    If dt.Rows.Count > 0 Then
    '                                        'If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GHO") = CompairStringResult.Equal Or CalculateTaxRatefromItemwsieTaxOnSale = True Then
    '                                        '    dblRate = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price"))
    '                                        '    If dblRate = 0 Then
    '                                        '        Throw New Exception("Please Fill Basic Price for Location " & txtLocation.Value & "  for item " & gv1.Rows(intRow).Cells(colICode).Value & ".")
    '                                        '        Exit Sub
    '                                        '    End If
    '                                        'Else
    '                                        ''--------
    '                                        dblRate = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
    '                                        If dblRate = 0 Then
    '                                            Throw New Exception("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
    '                                            Exit Function
    '                                        End If
    '                                        'End If
    '                                        objTr.MRP = clsCommon.myCdbl(dt.Rows(0).Item("Item_MRP"))
    '                                        objTr.Price_Date = clsCommon.myCstr(dt.Rows(0).Item("Start_Date"))
    '                                        objTr.Item_Price_ID = clsCommon.myCstr(dt.Rows(0).Item("Item_Price_ID"))
    '                                        objTr.Price_IdStartDate = clsCommon.myCstr(dt.Rows(0).Item("Start_date"))
    '                                        objTr.PricePlanNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Plan_Code  from TSPL_ITEM_PRICE_PLAN_DETAIL WHERE PLAN_TR_CODE='" & clsCommon.myCstr(dt.Rows(0).Item("Against_Plan_TR_Code")) & "'", trans))
    '                                    Else
    '                                        Throw New Exception("Please create Price chart for customer " & txtVendorNo.Value & " for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value) & ".")
    '                                        blnSaveTotalQTy = False
    '                                        Exit Function
    '                                    End If
    '                                    '''''''''''scheme
    '                                    Dim obj_Cash As clsSchemeApplyOnDairy = Nothing
    '                                    Dim tax As Double = 0
    '                                    Dim tax_on_amt As Decimal = 0
    '                                    obj_Cash = clsSchemeApplyOnDairy.GetDiscountSchemeData(clsCommon.myCstr(grow.Cells(colICode).Value), clsCommon.myCstr(grow.Cells(colUnit).Value), clsCommon.myCdbl(grow.Cells(colQty).Value), txtVendorNo.Value, Nothing, trans)
    '                                    If obj_Cash IsNot Nothing Then
    '                                        gv1.CurrentRow.Cells(colDisc_Scheme_Amount).Value = clsCommon.myCdbl(obj_Cash.Cash_Amt) 'objBookingitem.Disc_Scheme_Amount = obj_Cash.Cash_Amt
    '                                        gv1.CurrentRow.Cells(colDisc_Scheme_Pers).Value = clsCommon.myCdbl(obj_Cash.Cash_Pers)  'objBookingitem.Disc_Scheme_Pers = obj_Cash.Cash_Pers
    '                                        gv1.CurrentRow.Cells(colDisc_Scheme_Code).Value = obj_Cash.Schm_Code        ' objBookingitem.Disc_Scheme_Code = obj_Cash.Schm_Code
    '                                        If clsCommon.myCdbl(obj_Cash.Cash_Pers) <> 0 Then
    '                                            gv1.CurrentRow.Cells(colDisc_Scheme_Type).Value = "P"
    '                                            gv1.CurrentRow.Cells(colDisc_Scheme_Amount).Value = System.Math.Round((dblRate * obj_Cash.Cash_Pers) / 100, 2)
    '                                        ElseIf clsCommon.myCdbl(obj_Cash.Cash_Amt) <> 0 Then
    '                                            gv1.CurrentRow.Cells(colDisc_Scheme_Type).Value = "A"
    '                                        End If
    '                                        dblRate = dblRate - gv1.CurrentRow.Cells(colDisc_Scheme_Amount).Value
    '                                        tax_on_amt = dblRate
    '                                        'Dim Alltax As Double = System.Math.Round((clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX5_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX6_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX7_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX8_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX9_Rate")) / 100 + clsCommon.myCdbl(dt.Rows(0).Item("TAX10_Rate"))), 3)
    '                                        'tax = System.Math.Round((dblRate * Alltax), 3)
    '                                        'dblRate = dblRate + tax
    '                                    End If
    '                                    ''''''''''''scheme
    '                                    'dt = clsDBFuncationality.GetDataTable(qry)
    '                                    'If dt.Rows.Count > 0 Then
    '                                    ' done by priti ERO/07/05/18-000297
    '                                    gv1.CurrentRow.Cells(colSellingRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
    '                                    gv1.CurrentRow.Cells(colOrgRate).Value = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
    '                                    If ShowMulMRPOfSameItemOnDairyBookingCustomer = True Then
    '                                        isCellValueChangedOpen = True
    '                                    End If
    '                                    gv1.CurrentRow.Cells(colRate).Value = dblRate
    '                                    If ShowMulMRPOfSameItemOnDairyBookingCustomer = True Then
    '                                        isCellValueChangedOpen = False
    '                                    End If
    '                                    Dim DOCdateCurrent As Date? = Nothing
    '                                    DOCdateCurrent = clsCommon.GETSERVERDATE(trans)
    '                                    ' Query to get scheme type of Item
    '                                    Dim qryScheme As String = "Select TSPL_SCHEME_MASTER_NEW.Scheme_Type from TSPL_SCHEME_MASTER_NEW left outer join TSPL_SCHEME_DETAIL_NEW on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_MASTER_NEW.Scheme_Code "
    '                                    qryScheme += " left outer join tspl_item_master on tspl_item_master.item_code=TSPL_SCHEME_DETAIL_NEW.Item_Code "
    '                                    qryScheme += " where TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select scheme_code from (select ROW_NUMBER() over (partition by scheme_type order by start_date desc) as sno, TSPL_SCHEME_MASTER_NEW.Scheme_Code from "
    '                                    qryScheme += " TSPL_SCHEME_MASTER_NEW where TSPL_SCHEME_MASTER_NEW.Scheme_Type='Quantitive' and TSPL_SCHEME_MASTER_NEW.Status='Active' and  TSPL_SCHEME_MASTER_NEW.Start_Date<='" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  and (TSPL_SCHEME_MASTER_NEW.End_Date >= '" & clsCommon.GetPrintDate(DOCdateCurrent, "dd/MMM/yyyy") & "'  or TSPL_SCHEME_MASTER_NEW.End_date is null) and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select TSPL_SCHEME_DETAIL_NEW.Scheme_Code from "
    '                                    qryScheme += " TSPL_SCHEME_DETAIL_NEW left outer join TSPL_SCHEME_BENEFICIARY on TSPL_SCHEME_DETAIL_NEW.Scheme_Code=TSPL_SCHEME_BENEFICIARY.Scheme_Code where MainItem_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' and Cust_Code='" & txtVendorNo.Value & "'))a where a.sno=1)"
    '                                    qryScheme += " and TSPL_SCHEME_MASTER_NEW.Scheme_Code in (select Scheme_Code from TSPL_SCHEME_BENEFICIARY where Cust_Code='" & txtVendorNo.Value & "') and TSPL_SCHEME_MASTER_NEW.Status='Active'"
    '                                    qryScheme += " and TSPL_SCHEME_DETAIL_NEW.MainItem_Code='" & gv1.CurrentRow.Cells(colICode).Value & "' "
    '                                    qryScheme += " order by TSPL_SCHEME_MASTER_NEW.Scheme_Code"
    '                                    Dim SchemeType As String = clsDBFuncationality.getSingleValue(qryScheme, trans)
    '                                    If clsCommon.myLen(clsCommon.myCstr(SchemeType)) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(SchemeType), "Quantitive") = CompairStringResult.Equal Then
    '                                        Dim objD As clsSchemeApplyOnDairy = Nothing
    '                                        objD = clsSchemeApplyOnDairy.GetSchemeData(clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value), txtVendorNo.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colSchemeType).Value), Nothing, Nothing, trans)
    '                                        If objD IsNot Nothing AndAlso objD.Arr.Count > 0 Then
    '                                            For Each objtrScheme As clsSchemeApplyOnDairy In objD.Arr
    '                                                gv1.CurrentRow.Cells(colSchemeType).Value = objtrScheme.schm_Type
    '                                                objTr.Scheme_Item_Code = objtrScheme.Schm_Icode
    '                                                objTr.Scheme_Qty = objtrScheme.Schm_Qty
    '                                                objTr.Scheme_Item_UOM = objtrScheme.Schm_Item_Uom
    '                                                objTr.Scheme_Code = objtrScheme.Schm_Code
    '                                            Next
    '                                        End If
    '                                    End If
    '                                    objTr.Rate = clsCommon.myCdbl(grow.Cells(colRate).Value)
    '                                    'sanjay
    '                                    objTr.OrgRate = clsCommon.myCdbl(grow.Cells(colOrgRate).Value)
    '                                    objTr.SellingPrice = clsCommon.myCdbl(grow.Cells(colSellingRate).Value)
    '                                    objTr.Tax_NonTax = clsCommon.myCdbl(grow.Cells(colTax_NonTax).Value)
    '                                    objTr.FreshAmbient = clsCommon.myCstr(grow.Cells(colFreshAmbient).Value)
    '                                    objTr.Amount = clsCommon.myCdbl(clsCommon.myCdbl(grow.Cells(colAmt).Value))
    '                                    dblTotal += objTr.Amount
    '                                    'End If
    '                                    If objTr.Rate = 0 Then
    '                                        blnRatezero = True
    '                                        DOmsg += "Please create Price chart for customer " & txtVendorNo.Value & " for Location " & txtLocation.Value & "  for item " & objTr.Item_Code & "." + Environment.NewLine
    '                                    End If
    '                                    objTr.Price_Code = obj.Price_code
    '                                    objTr.OrgUnit_code = clsCommon.myCstr(clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value))
    '                                    If (clsCommon.myLen(objTr.Item_Code) > 0) Then
    '                                        obj.Arr.Add(objTr)
    '                                    End If
    '                                    objTr.OrgRate = objTr.OrgRate
    '                                    intLineNo += 1
    '                                End If
    '                                obj.Total_Amt = dblTotal
    '                                If ShowMulMRPOfSameItemOnDairyBookingCustomer Then
    '                                    qry = "Update TSPL_BOOKING_DETAIL set DO_Qty=" & objTr.Qty & ",Booking_Qty=" & objTr.Qty & ",Total_Qty=" & dblTotal_Qty & " where item_code='" & objTr.Item_Code & "' and Scheme_Item<>'Y' and vehicle_code='" & obj.Lorry_No & "' and  Document_No='" & txtDocNo.Value & "' and Cust_Code='" & obj.Customer_Code & "' and    Loc_Code='" & obj.Location_Code & "' and Sampling =" & obj.Sampling & " and Line_No='" & objTr.Line_No & "'"
    '                                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
    '                                Else
    '                                    qry = "Update TSPL_BOOKING_DETAIL set DO_Qty=" & objTr.Qty & ",Booking_Qty=" & objTr.Qty & ",Total_Qty=" & dblTotal_Qty & " where item_code='" & objTr.Item_Code & "' and Scheme_Item<>'Y' and vehicle_code='" & obj.Lorry_No & "' and  Document_No='" & txtDocNo.Value & "' and Cust_Code='" & obj.Customer_Code & "' and    Loc_Code='" & obj.Location_Code & "' and Sampling =" & obj.Sampling & ""
    '                                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
    '                                End If
    '                            Next
    '                            If (obj.Arr IsNot Nothing OrElse obj.Arr.Count > 0) Then
    '                                If blnRatezero = False Then
    '                                    If CheckOutstandingOnbooking = 0 Then
    '                                        If AllowWo_Outstanding = False Then
    '                                            If CustomerOutstandingAmount(txtVendorNo.Value, trans) = False Then
    '                                                obj.CreditApproval_Reqd = "Y"
    '                                                obj.Status = 2
    '                                            End If
    '                                        End If
    '                                    End If
    '                                    ' Dim trans1 As SqlTransaction = clsDBFuncationality.GetTransactin()
    '                                    If (obj.SaveData(obj, True, trans)) Then
    '                                        lblDONumber.Text = obj.Document_No
    '                                        qry = "Update TSPL_BOOKING_DETAIL set DO_Posted=" & obj.Status & ", Delivery_No='" & obj.Document_No & "',DocumentAmount=" & clsCommon.myCdbl(lblTotRAmt1.Text) & " where Document_No='" & txtDocNo.Value & "' and Cust_Code='" & txtVendorNo.Value & "' and    Loc_Code='" & txtLocation.Value & "' " ' and vehicle_code='" & lblvehiclecode.Text & "' and Sampling='" & 0 & "'
    '                                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
    '                                        qry = "Update TSPL_BOOKING_MATSER set CreateDO_Automatic=1 where Document_No='" & txtDocNo.Value & "' "
    '                                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
    '                                        DOCreated = True
    '                                        If (clsDeliveryNoteDairySale.PostData("DEL-NOTE-FS", obj.Document_No, trans, 1)) Then
    '                                            'common.clsCommon.MyMessageBoxShow("Data Saved Successfully")
    '                                        End If
    '                                        ' trans.Commit()
    '                                    Else
    '                                        ' trans.Rollback()
    '                                        Throw New Exception("Error - Deliverery Note Dairy Sales! ")
    '                                    End If
    '                                End If
    '                            End If
    '                            'DO STATUS 1 -open 2 - pending 3 - approved 4 - posted
    '                        ElseIf (DOStatus = 3) Then
    '                            If (clsDeliveryNoteDairySale.PostData("DEL-NOTE-FS", lblDONumber.Text, 1)) Then
    '                                'common.clsCommon.MyMessageBoxShow("Data Saved Successfully")
    '                                DOCreated = True
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '            'Next
    '        End If
    '        blnSaveTotalQTy = False
    '        'clsCommon.MyMessageBoxShow(msg)
    '        Return True
    '    Catch ex As Exception
    '        'common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '        'blnSaveTotalQTy = False
    '        'Return False
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function
    'Private Sub MyLabel2_Click(sender As Object, e As EventArgs)
    'End Sub
    Private Sub btnCopy_Click(sender As Object, e As EventArgs)
        'Try
        '    Dim PreviousDate As Date = clsCommon.GETSERVERDATE()
        '    PreviousDate = PreviousDate.AddDays(-1)
        '    Dim qry1 As String = "select distinct TSPL_BOOKING_MATSER.Document_No as DocumentNo,convert(varchar(12),TSPL_BOOKING_MATSER.Document_date,103) as Document_date,TSPL_CUSTOMER_MASTER.Customer_Name,(select isnull((Select distinct '['+TSPL_LOCATION_MASTER.Location_Desc +']  ' from TSPL_booking_detail left outer join TSPL_LOCATION_MASTER on TSPL_BOOKING_DETAIL.Loc_Code=TSPL_LOCATION_MASTER.Location_Code where TSPL_booking_detail.Document_No=TSPL_BOOKING_MATSER.Document_No   for xml path('')),'')  ) as Location  ,case when TSPL_BOOKING_MATSER.Posted=1 then 'posted' else 'Unposted' end as Posted from TSPL_BOOKING_MATSER"
        '    qry1 += " left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_DETAIL.Document_No=TSPL_BOOKING_MATSER.Document_No "
        '    qry1 += " left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code "
        '    Dim whrClas As String = " TSPL_BOOKING_MATSER.comp_code='" & objCommonVar.CurrentCompanyCode & "' and TSPL_BOOKING_MATSER.Is_Taxable=2 "
        '    whrClas += " and convert(date,TSPL_BOOKING_MATSER.Document_Date,103) = convert(date,('" & PreviousDate & "'),103)"
        '    If clsCommon.myLen(txtVendorNo.Value) > 0 Then
        '        whrClas += " and TSPL_BOOKING_DETAIL.Cust_Code='" & txtVendorNo.Value & "'"
        '    End If
        '    Dim strCode As String = ""
        '    strCode = clsCommon.ShowSelectForm("PSBookingOrderNoFndd1", qry1, "DocumentNo", whrClas, "", "DocumentNo", True)
        '    If clsCommon.myLen(strCode) > 0 Then
        '        'FlagCopy = True
        '        Dim qry As String = ""
        '        Dim obj As New clsBookingEntryDairySale
        '        'Dim intRow As Integer
        '        obj = clsBookingEntryDairySale.GetData(strCode, NavigatorType.Current)
        '        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
        '            btnSave.Enabled = True
        '            btnPost.Enabled = True
        '            btnDelete.Enabled = True
        '            isInsideLoadData = False
        '            ' btnCopy.Enabled = True
        '            isNewEntry = True
        '            'btnSave.Text = "Update"
        '            BlankAllControls()
        '            LoadBlankGrid()
        '            txtLocation.Enabled = False
        '            txtVendorNo.Enabled = False
        '            'If obj.TRANSACTION_TYPE = "FS" Then
        '            '    rbtn_Fresh.IsChecked = True
        '            'ElseIf obj.TRANSACTION_TYPE = "PS" Then
        '            '    rbtn_Ambient.IsChecked = True
        '            'End If
        '            txtLocation.Value = obj.location_code
        '            lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
        '            GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        '            qry = "SELECT TSPL_BOOKING_DETAIL.*,TSPL_CUSTOMER_MASTER.Customer_Name FROM TSPL_BOOKING_DETAIL LEFT JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code WHERE Document_No='" & strCode & "' and scheme_item='N '"
        '            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
        '            txtVendorNo.Value = clsCommon.myCstr(dt2.Rows(0)("Cust_Code"))
        '            lblVendorName.Text = clsCommon.myCstr(dt2.Rows(0)("Customer_Name"))
        '            lblTotRAmt1.Text = clsCommon.myCstr(dt2.Rows(0)("DocumentAmount"))
        '            BookingStatus = 0 'clsCommon.myCstr(dt2.Rows(0)("Booking_Status"))
        '            DOStatus = 0 'clsCommon.myCstr(dt2.Rows(0)("DO_Posted"))
        '            qry = "select Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,Zone_Code,TSPL_CUSTOMER_MASTER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER left outer join " &
        '                    "TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join TSPL_VEHICLE_MASTER on " &
        '                    "TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where Cust_Code='" & txtVendorNo.Value & "'"
        '            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
        '            If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
        '                strVehicleCode = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
        '                lblvehiclecode.Text = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
        '                strVehicleDesc = clsCommon.myCstr(dt1.Rows(0)("Number"))
        '                lblvehicleName.Text = clsCommon.myCstr(dt1.Rows(0)("Number"))
        '                strRoutecode = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
        '                lblroutecode.Text = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
        '                strRouteDesc = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
        '                lblroutename.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
        '                Price_code = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
        '                txtPriceCode.Text = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
        '            End If
        '            For jj As Integer = 0 To dt2.Rows.Count() - 1
        '                gv1.Rows.AddNew()
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count 'clsCommon.myCdbl(dt2.Rows(jj)("Line_No"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dt2.Rows(jj)("Item_Code"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct item_desc from tspl_item_master where item_code='" & clsCommon.myCstr(dt2.Rows(jj)("Item_Code")) & "'"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")), Nothing)
        '                If ShowAvailableQtyOnDairyBooking Then
        '                    gv1.Rows(gv1.Rows.Count - 1).Cells(ColAvailableQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")), txtLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dt2.Rows(jj)("Unit_Code")), 0)
        '                End If
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("OrgRate"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dt2.Rows(jj)("Unit_Code"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colSellingRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Selling_Price"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = Math.Round(clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate")), 2)
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_On_Amount"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_Amount"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Amount).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Amount"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Code).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Code"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Pers).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Pers"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Type).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Type"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeType).Value = clsCommon.myCstr(dt2.Rows(jj)("Scheme_Type"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dt2.Rows(jj)("Tax_NonTax"))
        '                gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCdbl(dt2.Rows(jj)("FreshAmbient"))
        '                'gv1.CurrentRow.Cells(ColAvgQty).Value = clsDBFuncationality.getSingleValue(qry)
        '                UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
        '            Next
        '        End If
        '        gv1.Rows.AddNew()
        '        If gv1.Rows.Count > 0 Then
        '            ' gv1.Focus()
        '            gv1.CurrentRow = gv1.Rows(0)
        '            gv1.CurrentColumn = gv1.Columns(colQty)
        '        End If
        '        'End If
        '    End If
        'Catch ex As Exception
        '    common.clsCommon.MyMessageBoxShow(ex.Message)
        'End Try
    End Sub
    Private Sub UpdateCurrentRowAvgQty(ByVal IntRowNo As Integer)
        If CheckAvgQtyOnDairyBooking Then
            Dim qry As String = ""
            If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value)) > 0 Then
                qry = "select isnull(sum(TSPL_BOOKING_DETAIL.Booking_Qty)/3,0) " &
                    " from TSPL_BOOKING_DETAIL left join TSPL_BOOKING_MATSER on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No " &
                     " left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_BOOKING_DETAIL.Item_Code " &
                    " where TSPL_BOOKING_DETAIL.Item_Code='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colICode).Value) & "'" &
                    " and TSPL_BOOKING_DETAIL.Unit_code='" & clsCommon.myCstr(gv1.Rows(IntRowNo).Cells(colUnit).Value) & "' " &
                    " and TSPL_BOOKING_DETAIL.Cust_Code='" & txtVendorNo.Value & "' " &
                    " and TSPL_BOOKING_DETAIL.Loc_Code='" & txtLocation.Value & "' " &
                    " and isnull(TSPL_BOOKING_MATSER.Is_Cancelled,0) = 0 " &
                    " and convert(date,TSPL_BOOKING_MATSER.Document_Date,103) >= convert(date,'" & txtDate.Value.AddDays(-3) & "',103)" &
                    " and convert(date,TSPL_BOOKING_MATSER.Document_Date,103) <= convert(date,'" & txtDate.Value.AddDays(-1) & "',103)" &
                    " and  not exists (select 1 from TSPL_DELIVERY_NOTE_MASTER_FRESHSALE " &
                    " where TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Document_No=isnull(TSPL_BOOKING_DETAIL.Delivery_No,'') " &
                    " and TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.Short_Close='Y') "
                gv1.Rows(IntRowNo).Cells(ColAvgQty).Value = Math.Round(clsDBFuncationality.getSingleValue(qry), 2)
            End If
        End If
    End Sub
    Private Sub txtSalesman__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean)
        Dim qry As String = "select EMP_CODE as Code,Emp_Name as Name from TSPL_EMPLOYEE_MASTER"
        Dim whrcls As String = "Emp_type='Salesman'"
        txtSalesman.Value = clsCommon.ShowSelectForm("DBC-SNOSaleman", qry, "Code", whrcls, txtSalesman.Value, "Code", isButtonClicked)
        lblSalesman.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Emp_Name as Name from TSPL_EMPLOYEE_MASTER where EMP_CODE ='" & txtSalesman.Value & "' and Emp_type='Salesman'"))
    End Sub
    Private Sub txtLocation_Leave(sender As Object, e As EventArgs) Handles txtLocation.Leave
        If clsCommon.myLen(txtLocation.Value) > 0 Then
            If gv1.Rows.Count > 0 Then
                ' gv1.Focus()
                gv1.CurrentRow = gv1.Rows(0)
                gv1.CurrentColumn = gv1.Columns(colICode)
            End If
            'gv1.Rows(0).Cells(colICode).IsSelected = True
        End If
    End Sub
    'Private Sub MyLabel12_Click(sender As Object, e As EventArgs)
    'End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Try
            'Dim isSaved As Boolean = True
            'If clsCommon.myLen(txtDocNo.Value) > 0 Then
            '    If clsCommon.myLen(lblDONumber.Text) > 0 Then
            '        clsCommon.MyMessageBoxShow(Me, "Booking Can not cancelled, DO already created!", Me.Text)
            '        Exit Sub
            '    End If
            '    If common.clsCommon.MyMessageBoxShow("Do you want to cancel the Booking?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            '        Dim qrySaveCancel = "Update tspl_booking_matser set Is_Cancelled=1 where Document_No='" & txtDocNo.Value & "'"
            '        isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qrySaveCancel)
            '        If isSaved = True Then
            '            clsCommon.MyMessageBoxShow(Me, "Booking cancelled successfully!", Me.Text)
            '            LoadData(txtDocNo.Value, NavigatorType.Current)
            '        End If
            '    End If
            'End If
            Dim frm1 As New FrmPWD(Nothing)
            frm1.strType = clsFixedParameterType.Transactionupdate
            frm1.strCode = clsFixedParameterCode.BookingCancel
            frm1.ShowDialog()
            If frm1.isPasswordCorrect Then
                If common.clsCommon.MyMessageBoxShow("Are you sure to Cancel the Record?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    CancelData()
                    OneTimeCheck = True
                End If

            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub CancelData()
        Try
            'Dim strQry As String = ""
            'Dim DocNo As String = ""
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Document Not Found!")
            End If
            clsBookingEntryDairySale.CancelData(Me.Form_ID, txtDocNo.Value, NavigatorType.Current)
            clsCommon.MyMessageBoxShow(Me, "Successfully Cancelled", Me.Text)
            AddNew()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    ''richa ERO/11/12/18-000431 -------------- 11 Dec,2018
    Private Sub btnreverse_Click(sender As Object, e As EventArgs) Handles btnreverse.Click
        Try
            If common.clsCommon.MyMessageBoxShow("Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                '' REASON FOR DELETE 
                Dim Reason As String = ""
                Dim frm As New FrmFreeTxtBox1
                frm.Text = "Remarks for Reverse"
                frm.ShowDialog()
                If clsCommon.myLen(frm.strRmks) <= 0 Then
                    Exit Sub
                Else
                    Reason = frm.strRmks
                End If
                If clsBookingEntryDairySale.ReverseAndUnpost(txtDocNo.Value) Then
                    saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    '====added by preeti gupta Against ticket no[TEC/05/06/19-000520,TEC/30/07/19-000968]
    Private Sub rbtn_Fresh_ToggleStateChanged(sender As Object, args As StateChangedEventArgs)
        LoadBlankGrid()
        gv1.Rows.AddNew()
    End Sub
    Private Sub rbtn_Ambient_ToggleStateChanged(sender As Object, args As StateChangedEventArgs)
        LoadBlankGrid()
        gv1.Rows.AddNew()
    End Sub
    ''richa 16 Sep,2019 ERO/11/09/19-001027
    Private Sub chkSampling_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkSampling.ToggleStateChanged
        For i As Integer = 0 To gv1.Rows.Count - 1
            If Not chkSampling.Checked AndAlso DonotAllowtoChangeUOMinDairyBookingCustomer Then
                gv1.Rows(i).Cells(colUnit).ReadOnly = True
            Else
                gv1.Rows(i).Cells(colUnit).ReadOnly = False
            End If
        Next
        If chkSampling.Checked Then
            chkSampling.Enabled = False
        Else
            chkSampling.Enabled = True

        End If
    End Sub
    Private Sub cmbBookingType_Leave(sender As Object, e As EventArgs)
        If clsCommon.CompairString(cmbcashcredit.Text, "Select") <> CompairStringResult.Equal AndAlso gv1.Rows.Count > 0 Then
            'If gv1.Rows.Count > 0 Then
            'gv1.Focus()
            gv1.CurrentRow = gv1.Rows(0)
            gv1.CurrentColumn = gv1.Columns(colQty)
            'End If
        End If
    End Sub
    'Private Sub cmbBookingType_SelectedIndexChanged(sender As Object, e As UI.Data.PositionChangedEventArgs)
    '    Try
    '        ''richa VIJ/09/12/19-000099
    '        If clsCommon.myLen(txtVendorNo.Value) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") = CompairStringResult.Equal Then
    '            txtDate.Value = clsCommon.GETSERVERDATE()
    '        ElseIf clsCommon.CompairString(cmbcashcredit.Text, "Select") <> CompairStringResult.Equal Then
    '            txtDate.Value = clsCommon.GETSERVERDATE().AddDays(1)
    '        Else
    '            txtDate.Value = clsCommon.GETSERVERDATE()
    '        End If
    '        If IsLoadBookingType = False Then
    '            'if Customer Category is other then do not refresh item grid
    '            If clsCommon.myLen(txtVendorNo.Value) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") = CompairStringResult.Equal Then
    '            Else
    '                LoadItemsFromPreviousBooking()
    '            End If
    '        End If
    '        If IsLoadBookingType = False Then
    '            If clsCommon.CompairString(cmbcashcredit.SelectedValue, "FN") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbcashcredit.SelectedValue, "PS") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbcashcredit.SelectedValue, "UP") = CompairStringResult.Equal Then
    '                chkGatePass.Checked = True
    '            Else
    '                chkGatePass.Checked = False
    '            End If
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    ''richa agarwal VIJ/28/11/19-000081 29 Nov,2019
    Sub LoadItemsFromPreviousBooking()
        Try
            If clsCommon.myLen(txtLocation.Value) <= 0 Then
                Throw New Exception("Please select Location")
            End If
            If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
                Throw New Exception("Please select Customer")
            End If
            Dim qry As String = "select TOP 1 TSPL_BOOKING_MATSER.Document_No  from TSPL_BOOKING_MATSER" & Environment.NewLine &
            " LEFT OUTER JOIN TSPL_BOOKING_DETAIL ON TSPL_BOOKING_MATSER.Document_No =TSPL_BOOKING_DETAIL.Document_No " & Environment.NewLine &
            " where TSPL_BOOKING_MATSER.From_Screen_Code ='BOOK-DS-CU'  and TSPL_BOOKING_MATSER.location_code ='" & txtLocation.Value & "' and TSPL_BOOKING_MATSER.Booking_Type ='" & clsCommon.myCstr(cmbcashcredit.SelectedValue) & "' and TSPL_BOOKING_DETAIL .Cust_Code ='" & txtVendorNo.Value & "'" & Environment.NewLine &
            " order by TSPL_BOOKING_MATSER.Document_Date desc"
            Dim strBookingNo As String = clsDBFuncationality.getSingleValue(qry, Nothing)
            If clsCommon.CompairString(clsCommon.myCstr(cmbcashcredit.SelectedValue), "FESTIVE ORDER") = CompairStringResult.Equal Then
                strBookingNo = Nothing
            End If
            If PrintTruckSheetAfterGenerate AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") <> CompairStringResult.Equal Then
                Dim isTruckSheetGenerated As Boolean = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select max(truckSheetGenerate) from TSPL_BOOKING_MATSER where convert(date,Document_Date ,103)=convert(date,'" & txtDate.Value & "',103)")) = 1, True, False)
                If isTruckSheetGenerated Then
                    strBookingNo = Nothing
                End If
            End If
            If clsCommon.myLen(strBookingNo) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") <> CompairStringResult.Equal Then
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") <> CompairStringResult.Equal Then
                    qry = "SELECT FINAL.Item_Code,FINAL.Item_Desc,FINAL.Short_Description,FINAL.Unit_code,FINAL.IsTaxable,FINAL.IsFreshAmbient,FINAL.HSN_Code,FINAL.Booking_Qty FROM ( " &
                        "select CASE WHEN Z.Booking_Qty>0 THEN 1 ELSE 0 END AS HasQty,* from ( select tspl_booking_detail.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,tspl_booking_detail.Unit_code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,tspl_booking_detail.Booking_Qty,tspl_item_master.Sku_Seq from tspl_booking_detail " & Environment.NewLine &
                " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE  " & Environment.NewLine &
                " where tspl_booking_detail.Document_No='" & strBookingNo & "' AND tspl_booking_detail.Scheme_Item ='N'" & Environment.NewLine &
                " union " & Environment.NewLine &
                " select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,TSPL_ITEM_UOM_DETAIL.UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,0 as Booking_Qty,tspl_item_master.Sku_Seq   from tspl_item_master " & Environment.NewLine &
                " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
                " where isnull(tspl_item_master.Chilled_Freezen,0)=1 and isnull(TSPL_ITEM_MASTER.item_type,'')='F'  and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1" & Environment.NewLine &
                " and tspl_item_master.Item_Code not in ( select tspl_booking_detail.item_code from tspl_booking_detail" & Environment.NewLine &
                " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
                " where tspl_booking_detail.Document_No='" & strBookingNo & "' AND tspl_booking_detail.Scheme_Item ='N'))z" & Environment.NewLine &
                "  )FINAL order by FINAL.HasQty desc,FINAL.Sku_Seq asc"
                Else
                    'qry = "select tspl_booking_detail.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,tspl_booking_detail.Unit_code ,tspl_item_master.IsTaxable," & Environment.NewLine & _
                    '" case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,tspl_booking_detail.Booking_Qty   from tspl_booking_detail" & Environment.NewLine & _
                    '" LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine & _
                    '" where tspl_booking_detail.Document_No='" & strBookingNo & "' AND tspl_booking_detail.Scheme_Item ='N'"
                    qry = "SELECT FINAL.Item_Code,FINAL.Item_Desc,FINAL.Short_Description,FINAL.Unit_code,FINAL.IsTaxable,FINAL.IsFreshAmbient,FINAL.HSN_Code,FINAL.Booking_Qty FROM ( " &
                         "select CASE WHEN Z.Booking_Qty>0 THEN 1 ELSE 0 END AS HasQty,* from ( select tspl_booking_detail.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,tspl_booking_detail.Unit_code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                  " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,tspl_booking_detail.Booking_Qty,tspl_item_master.Marketing_Seq   from tspl_booking_detail" & Environment.NewLine &
                  " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
                  " where tspl_booking_detail.Document_No='" & strBookingNo & "' AND tspl_booking_detail.Scheme_Item ='N'" & Environment.NewLine &
                  " union " & Environment.NewLine &
                  " select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,TSPL_ITEM_UOM_DETAIL.UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                  " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,0 as Booking_Qty,tspl_item_master.Marketing_Seq   from tspl_item_master " & Environment.NewLine &
                  " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
                  " where isnull(tspl_item_master.Item_Type,'')='F' and isnull(tspl_item_master.TypeOfItm,'')='A' and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1 " & Environment.NewLine &
                  " and tspl_item_master.Item_Code not in ( select tspl_booking_detail.item_code from tspl_booking_detail" & Environment.NewLine &
                  " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
                  " where tspl_booking_detail.Document_No='" & strBookingNo & "' AND tspl_booking_detail.Scheme_Item ='N'))z" & Environment.NewLine &
                  "  )FINAL order by FINAL.HasQty desc,FINAL.Marketing_Seq asc"
                End If
                'qry = "select tspl_booking_detail.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,tspl_booking_detail.Unit_code ,tspl_item_master.IsTaxable," & Environment.NewLine & _
                '" case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,tspl_booking_detail.Booking_Qty   from tspl_booking_detail" & Environment.NewLine & _
                '" LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine & _
                '" where tspl_booking_detail.Document_No='" & strBookingNo & "' AND tspl_booking_detail.Scheme_Item ='N'"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Nothing)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    LoadBlankGrid()
                    isInsideLoadData = True
                    Dim i As Integer = 0
                    For Each dr As DataRow In dt.Rows
                        isCellValueChangedOpen = True
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i + 1
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit_code"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
                        If clsCommon.myCdbl(clsCommon.myCdbl(dr("Booking_Qty"))) > 0 Then
                            'ItemPrice(clsCommon.myCstr(dr("item_code")), clsCommon.myCstr(dr("Unit_code")), clsCommon.myCdbl(dr("Booking_Qty")), gv1.Rows.Count - 1)
                        End If
                        isCellValueChangedOpen = False
                        i = i + 1
                    Next
                    isInsideLoadData = False
                    UpdateAllTotals()
                Else
                    gv1.DataSource = Nothing
                End If
            Else
                If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                    Dim strCustomerCategory As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "'"))
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") <> CompairStringResult.Equal Then
                        qry = "select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,case when len (isnull(TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Default_UOM ,'')) > 0 then TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Default_UOM else TSPL_ITEM_UOM_DETAIL.UOM_Code end as UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                  " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code  from tspl_item_master " & Environment.NewLine &
                  " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
                  " left outer join TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS on TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Item_Code = tspl_item_master.Item_Code and TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Customer_Category = '" & strCustomerCategory & "' " & Environment.NewLine &
                  " where isnull(tspl_item_master.Chilled_Freezen,0)=1 and isnull(TSPL_ITEM_MASTER.item_type,'')='F'  and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1  order by tspl_item_master.Sku_Seq"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Nothing)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            LoadBlankGrid()
                            isInsideLoadData = True
                            Dim i As Integer = 0
                            For Each dr As DataRow In dt.Rows
                                isCellValueChangedOpen = True
                                gv1.Rows.AddNew()
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i + 1
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("UOM_Code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
                                isCellValueChangedOpen = False
                                i = i + 1
                            Next
                            isInsideLoadData = False
                            UpdateAllTotals()
                        Else
                            gv1.DataSource = Nothing
                        End If
                    Else
                        qry = " select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,case when len (isnull(TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Default_UOM ,'')) > 0 then TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Default_UOM else TSPL_ITEM_UOM_DETAIL.UOM_Code end as UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
                " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code  from tspl_item_master " & Environment.NewLine &
                " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
                " left outer join TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS on TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Item_Code = tspl_item_master.Item_Code and TSPL_CUST_CATEGORY_DEFAULT_UOM_DETAILS.Customer_Category = '" & strCustomerCategory & "' " & Environment.NewLine &
                " where isnull(tspl_item_master.Item_Type,'')='F' and isnull(tspl_item_master.TypeOfItm,'')='A' and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1 order by Marketing_Seq"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Nothing)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            LoadBlankGrid()
                            isInsideLoadData = True
                            Dim i As Integer = 0
                            For Each dr As DataRow In dt.Rows
                                isCellValueChangedOpen = True
                                gv1.Rows.AddNew()
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i + 1
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("UOM_Code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
                                isCellValueChangedOpen = False
                                i = i + 1
                            Next
                            isInsideLoadData = False
                            UpdateAllTotals()
                        Else
                            gv1.DataSource = Nothing
                        End If
                    End If
                End If
            End If
            If gv1.Rows.Count > 0 Then
                'gv1.Focus()
                gv1.CurrentRow = gv1.Rows(0)
                gv1.CurrentColumn = gv1.Columns(colQty)
            End If
            'cmbBookingType.Focus()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub PaymentType()
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Payment_Code from TSPL_PAYMENT_CODE")
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                cmbPaymentType.DataSource = Nothing
                cmbPaymentType.Items.Clear()
                cmbPaymentType.DataSource = dt
                cmbPaymentType.DisplayMember = "Payment_Code"
                cmbPaymentType.ValueMember = "Payment_Code"
                cmbPaymentType.SelectedIndex = -1
            Else
                cmbPaymentType.DataSource = Nothing
                cmbPaymentType.Items.Clear()
                clsCommon.MyMessageBoxShow(Me, "Payment Code not found !", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub LoadBookingType()
        Try
            IsLoadBookingType = True
            'If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
            '    Throw New Exception("Please select Customer First")
            'End If
            Dim CustomerCategory As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' "))
            Dim dt As DataTable = New DataTable()
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Name", GetType(String))
            Dim dr As DataRow = dt.NewRow()
            If clsCommon.CompairString(CustomerCategory, "Vendor") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "CASH"
                dr("Name") = "CASH"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "FESTIVE ORDER"
                dr("Name") = "FESTIVE ORDER"
                dt.Rows.Add(dr)
            ElseIf clsCommon.CompairString(CustomerCategory, "Institution SO") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "SO"
                dr("Name") = "SO"
                dt.Rows.Add(dr)
            ElseIf clsCommon.CompairString(CustomerCategory, "Institution CR") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "CR"
                dr("Name") = "CR"
                dt.Rows.Add(dr)
            ElseIf clsCommon.CompairString(CustomerCategory, "Distributor") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "CASH"
                dr("Name") = "CASH"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "FN"
                dr("Name") = "FORENOON"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "PS"
                dr("Name") = "PARLOR SALES"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "UP"
                dr("Name") = "UP COUNTRY"
                dt.Rows.Add(dr)
            ElseIf clsCommon.CompairString(CustomerCategory, "FORENOON") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "FN"
                dr("Name") = "FORENOON"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "FESTIVE ORDER"
                dr("Name") = "FESTIVE ORDER"
                dt.Rows.Add(dr)
            ElseIf clsCommon.CompairString(CustomerCategory, "UP COUNTRY") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "UP"
                dr("Name") = "UP COUNTRY"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "FESTIVE ORDER"
                dr("Name") = "FESTIVE ORDER"
                dt.Rows.Add(dr)
            ElseIf clsCommon.CompairString(CustomerCategory, "PARLOR SALES") = CompairStringResult.Equal Then
                dr = dt.NewRow()
                dr("Code") = "PS"
                dr("Name") = "PARLOR SALES"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "FESTIVE ORDER"
                dr("Name") = "FESTIVE ORDER"
                dt.Rows.Add(dr)
            Else
                dr("Code") = ""
                dr("Name") = "Select"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "CASH"
                dr("Name") = "CASH"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "FESTIVE ORDER"
                dr("Name") = "FESTIVE ORDER"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "CR"
                dr("Name") = "CR"
                dt.Rows.Add(dr)
                dr = dt.NewRow()
                dr("Code") = "SO"
                dr("Name") = "SO"
                dt.Rows.Add(dr)
            End If
            cmbcashcredit.DataSource = dt
            cmbcashcredit.ValueMember = "Code"
            cmbcashcredit.DisplayMember = "Name"
            IsLoadBookingType = False
            If clsCommon.CompairString(cmbcashcredit.SelectedValue, "FN") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbcashcredit.SelectedValue, "PS") = CompairStringResult.Equal OrElse clsCommon.CompairString(cmbcashcredit.SelectedValue, "UP") = CompairStringResult.Equal Then
                chkGatePass.Checked = True
                chkGatePass.Enabled = False
            Else
                chkGatePass.Checked = False
                chkGatePass.Enabled = True
            End If
        Catch ex As Exception
            IsLoadBookingType = False
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub btn_ChangeIndent_Click(sender As Object, e As EventArgs) Handles btnCreateAndPrintInvoice.Click
    '    'If clsCommon.myLen(txtDocNo.Value) <= 0 Then
    '    If clsCommon.myLen(txtVendorNo.Value) <= 0 Then
    '        common.clsCommon.MyMessageBoxShow("Please select Customer and then click on Change Indent.")
    '        txtVendorNo.Focus()
    '        Exit Sub
    '    End If
    '    Dim STRSQL As String = ""
    '    'If ShowBookingTypeDropDownonDairyBookingCustomer = True AndAlso clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") = CompairStringResult.Equal Then
    '    '    STRSQL = "select TOP 1 TSPL_BOOKING_MATSER.Document_No from TSPL_BOOKING_DETAIL left join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.cust_code=TSPL_BOOKING_DETAIL.cust_code where TSPL_BOOKING_MATSER.From_Screen_Code ='BOOK-DS-CU' and TSPL_BOOKING_MATSER.location_code ='" & txtLocation.Value & "' and TSPL_BOOKING_DETAIL.Cust_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' AND convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & txtDate.Value & "',103) and TSPL_CUSTOMER_MASTER.customer_category<>'Others' order by TSPL_BOOKING_MATSER.Document_Date desc"
    '    'Else
    '    STRSQL = "select TOP 1 TSPL_BOOKING_MATSER.Document_No from TSPL_BOOKING_DETAIL left join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.cust_code=TSPL_BOOKING_DETAIL.cust_code where TSPL_BOOKING_MATSER.From_Screen_Code ='BOOK-DS-CU' and TSPL_BOOKING_MATSER.location_code ='" & txtLocation.Value & "' and TSPL_BOOKING_DETAIL.Cust_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' AND convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & txtDate.Value & "',103) and TSPL_CUSTOMER_MASTER.customer_category<>'Others' order by TSPL_BOOKING_MATSER.Document_Date desc"
    '    'End If
    '    Dim TempBookingNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(STRSQL))
    '    If clsCommon.myLen(TempBookingNo) > 0 Then
    '        LoadData(TempBookingNo, NavigatorType.Current)
    '    Else
    '        common.clsCommon.MyMessageBoxShow("Booking Not found for Date [" + clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") + "]", Me.Text)
    '    End If
    '    'End If
    'End Sub
    'Private Sub txtVendorNo_Leave(sender As Object, e As EventArgs) Handles txtVendorNo.Leave
    '    If clsCommon.myLen(txtVendorNo.Value) > 0 Then
    '        cmbBookingType.Focus()
    '    Else
    '        txtVendorNo.Focus()
    '    End If
    'End Sub
    Private Sub txtShipToLocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtShipToLocation._MYValidating
        If clsCommon.myLen(txtLocation.Value) = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select Location first", Me.Text)
            txtLocation.Focus()
            Exit Sub
        End If
        If clsCommon.myLen(txtVendorNo.Value) = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select Customer first", Me.Text)
            txtVendorNo.Focus()
            Exit Sub
        End If
        Dim qry As String = " select distinct TSPL_SHIP_TO_LOCATION.Ship_To_Code as [Code],TSPL_SHIP_TO_LOCATION.Ship_To_Desc as [Description],TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code as[Customer Code] ,TSPL_SHIP_TO_LOCATION.Ship_To_Type_Desc  as [Customer Discription], replace(case when ISNULL (TSPL_CUSTOMER_MASTER.Add1,'')='' then '' else TSPL_CUSTOMER_MASTER.add1 +',' end + case when ISNULL (TSPL_CUSTOMER_MASTER.Add2,'')='' then '' else TSPL_CUSTOMER_MASTER.add2 +',' end + case when ISNULL (TSPL_CUSTOMER_MASTER.Add3,'')='' then '' else TSPL_CUSTOMER_MASTER.add3 +',' end  ,',,',',') as [Customer Address] ,replace(case when ISNULL (TSPL_SHIP_TO_LOCATION.Add1,'')='' then '' else TSPL_SHIP_TO_LOCATION.add1 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add2,'')='' then '' else TSPL_SHIP_TO_LOCATION.add2 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add3,'')='' then '' else TSPL_SHIP_TO_LOCATION.add3 +',' end + case when ISNULL (TSPL_SHIP_TO_LOCATION.Add4,'')='' then '' else TSPL_SHIP_TO_LOCATION.add4 +',' end ,',,',',') as [Ship to Address],TSPL_SHIP_TO_LOCATION.CST_No as [CST NO],TSPL_SHIP_TO_LOCATION.Tin_No as [TIN No]  from TSPL_SHIP_TO_LOCATION left outer join TSPL_SHIP_TO_LOCATION_LOCATIONS on TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SHIP_TO_LOCATION_LOCATIONS.Ship_To_Code left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SHIP_TO_LOCATION.Ship_To_Type_Code "
        txtShipToLocation.Value = clsCommon.ShowSelectForm("DbS-ShipShindrlter", qry, "Code", "Ship_To_Type_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "' and (TSPL_SHIP_TO_LOCATION_LOCATIONS.Loc_Code ='" & txtLocation.Value & "' or  TSPL_SHIP_TO_LOCATION.Loc_Code='" & txtShipToLocation.Value & "' )", txtShipToLocation.Value, "Code", isButtonClicked)
        lblShipToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Ship_To_Desc from TSPL_SHIP_TO_LOCATION where Ship_To_Code='" & txtShipToLocation.Value & "'"))
    End Sub
    Private Function GetTCSRate(ByVal strCustomerCode As String) As Double
        Dim dblTCSRate As Double = 0
        'dblTCSRate = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select   case when len( isnull (PAN,'')) > 0 OR LEN (ISNULL(Collectorate,'')) >0  then  (select Description from TSPL_FIXED_PARAMETER where Type = 'TCSRateforCustomerWithPanNo' and code = 'TCSRateforCustomerWithPanNo') else (select Description from TSPL_FIXED_PARAMETER where Type = 'TCSRateforCustomerWithoutPanNo' and code = 'TCSRateforCustomerWithoutPanNo')  end TCSRate  from TSPL_CUSTOMER_MASTER  where  CUSTOMER_CATEGORY = 'Others' and  isnull (IsTCSnotApplicable,0) = 0 and Cust_Code = '" & strCustomerCode & "'"))
        dblTCSRate = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select   case when len( isnull (PAN,'')) > 0 OR LEN (ISNULL(Collectorate,'')) >0  then  (select Description from TSPL_FIXED_PARAMETER where Type = 'TCSRateforCustomerWithPanNo' and code = 'TCSRateforCustomerWithPanNo') else (select Description from TSPL_FIXED_PARAMETER where Type = 'TCSRateforCustomerWithoutPanNo' and code = 'TCSRateforCustomerWithoutPanNo')  end TCSRate  from TSPL_CUSTOMER_MASTER  where  isnull (IsTCSnotApplicable,0) = 0 and Cust_Code = '" & strCustomerCode & "'"))
        Return dblTCSRate
    End Function
    Private Sub ChkGatePass_ToggleStateChanged(sender As Object, args As StateChangedEventArgs)
        If chkGatePass.Checked Then
            cmbGatePassType.Enabled = True
            txtDate.Value = clsCommon.GETSERVERDATE()
        Else
            cmbGatePassType.Text = "Select"
            cmbGatePassType.Enabled = False
        End If
    End Sub
    Private Sub BtnSerach_Click(sender As Object, e As EventArgs)
        Try
            For jj As Integer = 0 To gv1.Rows.Count - 1
                gv1.ClearSelection()
                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colIShortName).Value), txtItemSearch.Text.Trim()) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colIName).Value), txtItemSearch.Text.Trim()) = CompairStringResult.Equal Then
                    gv1.Rows(jj).Cells(colQty).IsSelected = True
                    gv1.Rows(jj).IsCurrent = True
                    gv1.Columns(colQty).IsCurrent = True
                    gv1.Focus()
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub Btn_QtyReset_Click(sender As Object, e As EventArgs) Handles btn_QtyReset.Click
        Try
            If chkGatePass.Checked Then
                For jj As Integer = 0 To gv1.Rows.Count - 1
                    If clsCommon.myCdbl(gv1.Rows(jj).Cells(colQty).Value) > 0 Then
                        gv1.Rows(jj).Cells(colQty).Value = 0
                    End If
                Next
                gv1.Rows(0).Cells(colQty).IsSelected = True
                gv1.Rows(0).IsCurrent = True
                gv1.Columns(colQty).IsCurrent = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles btnGatePassPrint.Click
    '    Try
    '        'tspl_booking_matser.AgainstGatePass=1 and tspl_booking_matser.GatePass_Type = '" & cmbGatePassType.Text & "'
    '        GetTruckSheetQry(" tspl_booking_matser.AgainstGatePass=1 and tspl_booking_matser.GatePass_Type = '" & cmbGatePassType.Text & "' and tspl_booking_matser.Document_No = '" & txtDocNo.Value & "' ", True)
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    'Public Sub GetTruckSheetQry(ByVal strCondition As String, Optional ByVal isGatePassTruckSheet As Boolean = False)
    '    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "TSDDCF") = CompairStringResult.Equal Then
    '        Dim whr As String = " and " & strCondition & " "
    '        Dim whrCustCategory As String = ""
    '        Dim whrVehicle As String = ""
    '        Dim whrRoute As String = ""
    '        Dim whrShipment As String = ""
    '        Dim isTrueFalse As String = " where 2=2"
    '        If isGatePassTruckSheet Then
    '            isTrueFalse = " where 2=3 "
    '        End If
    '        'If clsCommon.myLen(clsCommon.myCstr(txtLorryNo.Value)) > 0 Then
    '        '    whr += "  and TSPL_ROUTE_MASTER.Vehicle_Code in ('" & txtLorryNo.Value & "') "
    '        '    whrVehicle = " and xx.Vehicle_Code in  ('" & txtLorryNo.Value & "') "
    '        '    whrShipment = " and TSPL_ROUTE_MASTER.Vehicle_Code in ('" & txtLorryNo.Value & "')  "
    '        'End If
    '        'If txtRouteNo.arrValueMember IsNot Nothing AndAlso txtRouteNo.arrValueMember.Count > 0 Then
    '        '    whr += "  and TSPL_BOOKING_DETAIL.route_No in (" + clsCommon.GetMulcallString(txtRouteNo.arrValueMember) + ") "
    '        '    whrRoute = " and TSPL_CUSTOMER_MASTER.Route_No in (" + clsCommon.GetMulcallString(txtRouteNo.arrValueMember) + ") "
    '        '    whrShipment = " and TSPL_SD_SHIPMENT_HEAD.Route_No in  (" + clsCommon.GetMulcallString(txtRouteNo.arrValueMember) + ") "
    '        'Else
    '        '    If chkShowEarlyRoute.Checked = True Then
    '        '        whr += "  and TSPL_BOOKING_DETAIL.route_No in (" + clsCommon.myCstr(StrEarlyRoute) + ") "
    '        '        whrRoute = " and TSPL_CUSTOMER_MASTER.Route_No in (" + clsCommon.myCstr(StrEarlyRoute) + ") "
    '        '        whrShipment = " and TSPL_SD_SHIPMENT_HEAD.Route_No in  (" + clsCommon.myCstr(StrEarlyRoute) + ") "
    '        '    Else
    '        '        whr += "  and TSPL_BOOKING_DETAIL.route_No in (" + clsCommon.myCstr(StrNonEarlyRoute) + ") "
    '        '        whrRoute = " and TSPL_CUSTOMER_MASTER.Route_No in (" + clsCommon.myCstr(StrNonEarlyRoute) + ") "
    '        '        whrShipment = " and TSPL_SD_SHIPMENT_HEAD.Route_No in  (" + clsCommon.myCstr(StrNonEarlyRoute) + ") "
    '        '    End If
    '        'End If
    '        'If TxtMultiCustomerCategory.arrValueMember IsNot Nothing AndAlso TxtMultiCustomerCategory.arrValueMember.Count > 0 Then
    '        '    whr += " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(TxtMultiCustomerCategory.arrValueMember) + ")" + Environment.NewLine
    '        '    whrCustCategory = " and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(TxtMultiCustomerCategory.arrValueMember) + ") "
    '        '    whrShipment = " and   TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY in (" + clsCommon.GetMulcallString(TxtMultiCustomerCategory.arrValueMember) + ") "
    '        'End If
    '        Dim dtCreateDetailsDate As DateTime = txtDate.Value.AddDays(-1) 'New Date(TSP_Date.Value.Day, TSP_Date.Value.Day, -1)
    '        Dim strCreateDetails As String = " With CTETemp as ( Select convert(varchar,Doc_Date,103) as Doc_Date, Route_No,Route_Desc,Vehicle_Code,Vehicle_Name,Customer_Code,Customer_Name,Customer_Phone, ZSM , ZSM_NAME,ZSM_Phone,ASM,ASM_Name,ASM_Phone,ASO , ASO_Name,ASO_Phone ,OpencrateQty, OpenJaaliQty, OpenBoxQty,OpenCanQty, CrateQtyRecd, JaaliQtyRecd, BoxQtyRecd ,CanQtyRecd,CrateOutQty, jaaliOutQty,boxOutQty  ,CanOutQty, CrateAdjQty , JaaliAdjQty  , BoxAdjQty ,CanAdjQty,SUM(CrateQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code,Route_No) as CrateQtyClosing,  SUM(JaaliQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code,Route_No) as JaaliQtyClosing,  SUM(BoxQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code,Route_No) as BoxQtyClosing, SUM(CanQtyClosing) OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code,Route_No) as CanQtyClosing , Row_Number() OVER (Partition BY Customer_Code ORDER BY Customer_Code,Vehicle_Code,Route_No) as RowNo from(  " &
    '                                         " select  pp.Doc_Date  as Doc_Date,TSPL_CUSTOMER_MASTER.Route_No,TSPL_CUSTOMER_MASTER.Route_Desc,pp.Vehicle_Code,tspl_vehicle_master.Number as Vehicle_Name ,pp.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Phone1 as Customer_Phone ,   TSPL_CUSTOMER_MASTER.ZSM,TSPL_EMPLOYEE_MASTER_ZSM.Emp_Name as ZSM_NAME,case when len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PHONE,'')) > 0 And  len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO,'')) > 0  then '['+ TSPL_EMPLOYEE_MASTER_ZSM.PHONE + '],['  + TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO +']' when  len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PHONE,'')) > 0 then TSPL_EMPLOYEE_MASTER_ZSM.PHONE when len (isNull (TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO,'')) > 0 then TSPL_EMPLOYEE_MASTER_ZSM.PRESENT_MOBILE_NO else '' end  as ZSM_Phone ,TSPL_CUSTOMER_MASTER.ASM, TSPL_EMPLOYEE_MASTER_ASM.Emp_Name as ASM_NAME, case when len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PHONE,'')) > 0 And  len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO,'')) > 0  then '['+ TSPL_EMPLOYEE_MASTER_ASM.PHONE + '],['  + TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO +']' when  len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PHONE,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASM.PHONE when len (isNull (TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASM.PRESENT_MOBILE_NO else '' end  as ASM_Phone,TSPL_CUSTOMER_MASTER.ASO,TSPL_EMPLOYEE_MASTER_ASO.Emp_Name as ASO_NAME,   case when len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PHONE,'')) > 0 And  len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO,'')) > 0  then '['+ TSPL_EMPLOYEE_MASTER_ASO.PHONE + '],['  + TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO +']' when  len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PHONE,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASO.PHONE when len (isNull (TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO,'')) > 0 then TSPL_EMPLOYEE_MASTER_ASO.PRESENT_MOBILE_NO else '' end  as ASO_Phone,pp.OpencrateQty as OpencrateQty,pp.OpenJaaliQty  as OpenJaaliQty,pp.OpenBoxQty  as OpenBoxQty,pp.OpenCanQty  as OpenCanQty,pp.CrateQtyRecd  as CrateQtyRecd,pp.JaaliQtyRecd  as JaaliQtyRecd,pp.BoxQtyRecd  as BoxQtyRecd ,pp.CanQtyRecd  as CanQtyRecd,pp.CrateOutQty  as CrateOutQty,pp.jaaliOutQty  as jaaliOutQty ,pp.boxOutQty  as boxOutQty,pp.CanOutQty  as CanOutQty ,pp.CrateQtyClosing as CrateQtyClosing, pp.JaaliQtyClosing as  JaaliQtyClosing, pp.BoxQtyClosing as BoxQtyClosing, pp.CanQtyClosing as CanQtyClosing,pp.CrateAdjQty , pp.JaaliAdjQty  , pp.BoxAdjQty , pp.CanAdjQty from ( " &
    '                                         " select  max(convert(date,Doc_Date,103))  as Doc_Date,max(xx.Vehicle_Code) as Vehicle_Code,xx.Customer_Code,sum(xx.OpencrateQty) as OpencrateQty,sum(xx.OpenJaaliQty ) as OpenJaaliQty,sum(xx.OpenBoxQty )  as OpenBoxQty,sum(xx.OpenCanQty )  as OpenCanQty,sum(xx.CrateQtyRecd) as CrateQtyRecd,sum(xx.JaaliQtyRecd) as JaaliQtyRecd,sum(xx.BoxQtyRecd) as BoxQtyRecd,sum(xx.CanQtyRecd) as CanQtyRecd,sum(xx.CrateOutQty ) as CrateOutQty,sum(xx.jaaliOutQty ) as jaaliOutQty ,sum(xx.boxOutQty ) as boxOutQty,sum(xx.CanOutQty ) as CanOutQty, sum(xx.CrateAdjQty ) as CrateAdjQty ,sum(xx.JaaliAdjQty )as JaaliAdjQty  ,sum(xx.BoxAdjQty )  as BoxAdjQty ,sum(xx.CanAdjQty )  as CanAdjQty,(sum(xx.OpencrateQty)+sum(xx.CrateOutQty )-sum(xx.CrateQtyRecd)-sum(xx.CrateAdjQty )) as CrateQtyClosing, (sum(xx.OpenJaaliQty)+sum(xx.jaaliOutQty)-sum(xx.JaaliQtyRecd)-sum(xx.JaaliAdjQty )) as JaaliQtyClosing, (sum(xx.OpenBoxQty)+sum(xx.boxOutQty )-sum(xx.BoxQtyRecd)-sum(xx.BoxAdjQty )) as BoxQtyClosing  , (sum(xx.OpenCanQty)+sum(xx.CanOutQty )-sum(xx.CanQtyRecd)-sum(xx.CanAdjQty )) as CanQtyClosing  from (select  max(convert(date,'" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtCreateDetailsDate), "dd/MMM/yyyy") & "',103)) as Doc_Date,max(Opening.Vehicle_Code) as Vehicle_Code , Opening.Customer_Code ,sum(Opening.OpencrateQty*Type) as OpencrateQty,sum(Opening.OpenJaaliQty*Type ) as OpenJaaliQty,sum(Opening.OpenBoxQty *Type)  as OpenBoxQty,sum(Opening.OpenCanQty *Type)  as OpenCanQty,sum(Opening.CrateQtyRecd *Type) as CrateQtyRecd,sum(Opening.JaaliQtyRecd*Type ) as JaaliQtyRecd,sum(Opening.BoxQtyRecd*Type ) as BoxQtyRecd,sum(Opening.CanQtyRecd*Type ) as CanQtyRecd,sum(Opening.CrateOutQty*Type ) as CrateOutQty,sum(Opening.jaaliOutQty*Type ) as jaaliOutQty ,sum(Opening.boxOutQty*Type ) as boxOutQty,sum(Opening.CanOutQty*Type ) as CanOutQty,sum(Opening.CrateAdjQty*Type ) as CrateAdjQty,sum(Opening.JaaliAdjQty*Type ) as JaaliAdjQty,sum(Opening.BoxAdjQty *Type) as  BoxAdjQty ,sum(Opening.CanAdjQty *Type) as  canAdjQty from  ( select TSPL_SD_SHIPMENT_HEAD.Document_Date    as Document_Date,(case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end) as Vehicle_Code ,TSPL_SD_SHIPMENT_HEAD.customer_code  ,1 as Type  ,'O' as Type1,TSPL_SD_SHIPMENT_HEAD.Crate as OpencrateQty,TSPL_SD_SHIPMENT_HEAD.jaali as OpenJaaliQty, TSPL_SD_SHIPMENT_HEAD.box  as OpenBoxQty , TSPL_SD_SHIPMENT_HEAD.ShippedCAN  as OpenCanQty  ,0 as CrateQtyRecd,0  as JaaliQtyRecd ,0 as BoxQtyRecd ,0 as CanQtyRecd ,0 as CrateOutQty,0 as jaaliOutQty,0 as boxOutQty ,0 as CanOutQty,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty  ,0 as  CanAdjQty  from TSPL_SD_SHIPMENT_HEAD  where TSPL_SD_SHIPMENT_HEAD.screen_type='DS'  AND TSPL_SD_SHIPMENT_HEAD.Status =1  union all select TSPL_sd_SALE_RETURN_HEAD.Document_Date    as Document_Date,TSPL_sd_SALE_RETURN_HEAD.Vehicle_Code ,TSPL_sd_SALE_RETURN_HEAD.customer_code  ,-1 as Type  ,'I' as Type1,TSPL_sd_SALE_RETURN_HEAD.CrateQty as OpencrateQty,TSPL_sd_SALE_RETURN_HEAD.jaali as OpenJaaliQty, TSPL_sd_SALE_RETURN_HEAD.box  as OpenBoxQty , TSPL_sd_SALE_RETURN_HEAD.ShippedCAN  as OpenCanQty  ,0 as CrateQtyRecd,0  as JaaliQtyRecd ,0 as BoxQtyRecd ,0 as CanQtyRecd ,0 as CrateOutQty,0 as jaaliOutQty,0 as boxOutQty ,0 as CanOutQty,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty  ,0 as  CanAdjQty  from TSPL_sd_SALE_RETURN_HEAD  where TSPL_sd_SALE_RETURN_HEAD.screen_type='DS'  AND TSPL_sd_SALE_RETURN_HEAD.Status =1  union all  select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  , TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as OpencrateQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty   as OpenjaaliQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty  as OpenboxQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty as OpenCanQty , 0 as CrateQtyRecd,0 JaaliQtyRecd , 0 BoxQtyRecd , 0 CanQtyRecd , 0 as CrateOutQty, 0 jaaliOutQty, 0 boxoutqty, 0 Canoutqty, 0  as CrateAdjQty, 0  as JaaliAdjQty, 0  as BoxAdjQty, 0  as CanAdjQty from TSPL_CRATE_RECEIVED_detail_FRESHSALE left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='O' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1  union all select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,-1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  , isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment,0)  as OpencrateQty, isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment,0)    as OpenjaaliQty , isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment,0)  as OpenboxQty, isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec,0) + isnull(TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment,0)  as OpenCanQty, 0 as CrateQtyRecd,0 JaaliQtyRecd , 0 BoxQtyRecd , 0 CanQtyRecd , 0 as CrateOutQty, 0 jaaliOutQty, 0 boxoutqty, 0 Canoutqty, 0  as CrateAdjQty, 0  as JaaliAdjQty, 0  as BoxAdjQty, 0  as CanAdjQty from TSPL_CRATE_RECEIVED_detail_FRESHSALE left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='I' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1   )as Opening WHERE convert(date,Document_Date,103)<(convert(date,'" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtCreateDetailsDate), "dd/MMM/yyyy") & "',103)) group by Customer_Code  " &
    '                                         " UNION All " &
    '                                         " select Document_Date,Vehicle_Code,Customer_Code,0 as OpencrateQty,0 as OpenjaaliQty ,0 as OpenboxQty,0 as OpenCanQty,Case When [Type]=1 Then CrateQtyRecd Else 0 End as CrateQtyRecd,Case When [Type]=1 Then JaaliQtyRecd Else 0 End as JaaliQtyRecd,Case When [Type]=1 Then BoxQtyRecd Else 0 End as BoxQtyRecd,Case When [Type]=1 Then CANQtyRec Else 0 End as CANQtyRec, " &
    '                                         " Case When [Type]=-1 Then CrateOutQty Else 0 End as CrateOutQty,Case When [Type]=-1 Then jaaliOutQty Else 0 End as jaaliOutQty,Case When [Type]=-1 Then boxOutQty Else 0 End as boxOutQty,Case When [Type]=-1 Then CanOutQty Else 0 End as CanOutQty,Case When [Type]=1 Then CrateAdjQty Else 0 End as CrateAdjQty,Case When [Type]=1 Then JaaliAdjQty Else 0 End as JaaliAdjQty,Case When [Type]=1 Then BoxAdjQty Else 0 End as BoxAdjQty,Case When [Type]=1 Then CANAdjustment Else 0 End as CANAdjustment  " &
    '                                         " from ((select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyManual as OpencrateQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaali  as OpenjaaliQty ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.box as OpenboxQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQty as OpenCanQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd as CrateQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as CrateOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment  as CrateAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment  as JaaliAdjQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment  as BoxAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment  as CANAdjustment from TSPL_CRATE_RECEIVED_detail_FRESHSALE " &
    '                                         " left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  " &
    '                                         " where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='I' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1 ) union all  " &
    '                                         " (select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date as Document_Date,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Vehicle_Code ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code  ,-1 as Type,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type as Type1  ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyManual as OpencrateQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaali  as OpenjaaliQty ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.box as OpenboxQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQty as OpenCANQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CrateQtyRecd as CrateQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.JaaliQtyRecd ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.BoxQtyRecd,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANQtyRec ,TSPL_CRATE_RECEIVED_detail_FRESHSALE.OutQty as CrateOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxOutQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANOutQty , TSPL_CRATE_RECEIVED_detail_FRESHSALE.Adjustment  as CrateAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.jaaliAdjustment  as JaaliAdjQty, TSPL_CRATE_RECEIVED_detail_FRESHSALE.boxAdjustment  as BoxAdjQty,TSPL_CRATE_RECEIVED_detail_FRESHSALE.CANAdjustment  as CANAdjustment from TSPL_CRATE_RECEIVED_detail_FRESHSALE " &
    '                                         " left join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No =TSPL_CRATE_RECEIVED_detail_FRESHSALE.Document_No  " &
    '                                         " where TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Type ='O' AND TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted  =1  union all  " &
    '                                         " select TSPL_SD_SHIPMENT_HEAD.Document_Date   as Document_Date,(case when isnull(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')='' then case when isnull(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' then TSPL_SD_SHIPMENT_HEAD.AlternateVehicle else TSPL_SD_SHIPMENT_HEAD.Vehicle_Code end else TSPL_SD_SHIPMENT_HEAD.ManualVehicle end) as Vehicle_Code ,TSPL_SD_SHIPMENT_HEAD.customer_code  ,-1 as Type  ,'O' as Type1,0 as OpencrateQty,0  as OpenBoxQty ,0 as OpenJaaliQty,0 as OpenCANQty, 0 as CrateQtyRecd, 0 JaaliQtyRecd ,  0  as BoxQtyRecd,0 CanQtyRecd  ,TSPL_SD_SHIPMENT_HEAD.Crate as CrateOutQty,jaali  as jaaliOutQty,TSPL_SD_SHIPMENT_HEAD.Box as boxOutQty,TSPL_SD_SHIPMENT_HEAD.ShippedCAN as CanOutQty ,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty,0 as  CanAdjQty    from TSPL_SD_SHIPMENT_HEAD  where TSPL_SD_SHIPMENT_HEAD.screen_type='DS' AND TSPL_SD_SHIPMENT_HEAD.Status =1 union all  " &
    '                                         " select TSPL_sd_SALE_RETURN_HEAD.Document_Date   as Document_Date,TSPL_sd_SALE_RETURN_HEAD.Vehicle_Code ,TSPL_sd_SALE_RETURN_HEAD.customer_code  ,1 as Type  ,'I' as Type1,0 as OpencrateQty,0  as OpenBoxQty ,0 as OpenJaaliQty,0 as OpenCANQty, TSPL_sd_SALE_RETURN_HEAD.CrateQty as CrateQtyRecd, TSPL_sd_SALE_RETURN_HEAD.jaali as JaaliQtyRecd ,  TSPL_sd_SALE_RETURN_HEAD.Box  as BoxQtyRecd,TSPL_sd_SALE_RETURN_HEAD.ShippedCAN CanQtyRecd  ,0 as CrateOutQty,0 as jaaliOutQty,0 as boxOutQty,0 as CanOutQty ,0 as CrateAdjQty,0 as JaaliAdjQty,0 as  BoxAdjQty,0 as  CanAdjQty    from TSPL_sd_SALE_RETURN_HEAD  where TSPL_sd_SALE_RETURN_HEAD.screen_type='DS' AND TSPL_sd_SALE_RETURN_HEAD.Status =1 )  ) as Closing  " &
    '                                         " WHERE convert(date,Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtCreateDetailsDate), "dd/MMM/yyyy") & "',103) AND convert(date,Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtCreateDetailsDate), "dd/MMM/yyyy") & "',103)  " &
    '                                         " ) as xx where 2=2   " &
    '                                         " and xx.Customer_Code in (Select distinct TSPL_BOOKING_DETAIL.Cust_Code  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  " &
    '                                         " left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code " &
    '                                         " where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  and " + strCondition + "  " &
    '                                         " " + whrCustCategory + " " &
    '                                         " ) " &
    '                                         " " + whrVehicle + " " &
    '                                         " GROUP BY Customer_Code " &
    '                                         " ) as pp   left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =pp.Customer_Code  " &
    '                                         " left outer Join  TSPL_EMPLOYEE_MASTER as TSPL_EMPLOYEE_MASTER_ZSM on TSPL_EMPLOYEE_MASTER_ZSM.EMP_CODE = TSPL_CUSTOMER_MASTER.ZSM  " &
    '                                         " left outer Join  TSPL_EMPLOYEE_MASTER as TSPL_EMPLOYEE_MASTER_ASM on TSPL_EMPLOYEE_MASTER_ASM.EMP_CODE = TSPL_CUSTOMER_MASTER.ASM  " &
    '                                         " left outer Join  TSPL_EMPLOYEE_MASTER as TSPL_EMPLOYEE_MASTER_ASO on TSPL_EMPLOYEE_MASTER_ASO.EMP_CODE = TSPL_CUSTOMER_MASTER.ASO  " &
    '                                         " left join tspl_vehicle_master on tspl_vehicle_master.vehicle_id=pp.vehicle_code where 2=3   " + whrRoute + " " &
    '                                         " ) YYY )   "
    '        Dim qry = " " + strCreateDetails + " " &
    '                 " select XXXXXFinal.Document_Date,XXXXXFinal.Cust_Code,Case when  XXXXXFinal.DataSNO = 1 and No_Of_Item =1 then SUBSTRING((XXXXXFinal.Cust_Code + ' ('+XXXXXFinal.Document_Date + ')'),0,19) else  SUBSTRING ( XXXXXFinal.Customer_Name,0,19) end as Customer_Name, XXXXXFinal.route_no,XXXXXFinal.Route_Desc,XXXXXFinal.Zone_Code,XXXXXFinal.Zone_Name,XXXXXFinal.Item_Code,SUBSTRING (XXXXXFinal.Alies_Name,0,8) as Alies_Name, Convert (integer,XXXXXFinal.CR) as CR,Convert (Integer ,XXXXXFinal.CD) as CD,convert (Integer,XXXXXFinal.SO) as SO,Convert (Integer ,XXXXXFinal.Cash,103) as Cash ,Convert (Integer,XXXXXFinal.Total) as Total,XXXXXFinal.CrateQty_New as CrateQty_New , XXXXXFinal.PendingPcsQty_New as PendingPcsQty_New, XXXXXFinal.Amount_with_Tax, XXXXXFinal.  DataSNO,   XXXXXFinal. Unit_code,TBL_Count.No_Of_Item ,  XXXXXFinal.LR_PQ_ES ,isnull (XXXXXFinal.OpencrateQty,0) as OpencrateQty , isnull ( XXXXXFinal.CrateQtyRecd,0) as CrateQtyRecd ,isnull( XXXXXFinal.CrateOutQty,0) as CrateOutQty, isnull ( XXXXXFinal.CrateAdjQty,0) as CrateAdjQty , isnull (XXXXXFinal.CrateQtyClosing,0) as CrateQtyClosing, Total_Scheme_Star, isnull ( Qty_in_KG,0) as Qty_in_KG , isnull (Qty_in_Ltr,0) as Qty_in_Ltr  from (  " &
    '                 " select XXXXX.Document_Date, XXXXX.Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,XXXXX.CrateQty_New as CrateQty_New , XXXXX.PendingPcsQty_New as PendingPcsQty_New,isnull (XXXXX.Amount_with_Tax,0) as Amount_with_Tax, '1' as  DataSNO, '' as  Unit_code,0 as  LR_PQ_ES , 0 as OpencrateQty , 0 as CrateQtyRecd , 0 as CrateOutQty , 0 as CrateAdjQty , 0 as CrateQtyClosing, case when  isnull (XXXXX.Total_Scheme,0) > 0 then '*' else '' end as Total_Scheme_Star,0 as Qty_in_KG , 0 as Qty_in_Ltr  from (  Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity,(StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty , Convert (int,CrateQty) as CrateQty_New , (XXXFinal.Total) - ((StockPcs.Conversion_Factor * Convert (int,CrateQty))) as PendingPcsQty_New, (StockPcs.Conversion_Factor * Convert (int,CrateQty)) as TotalCrateCapicity_New from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select XFinal.Document_Date,max (XFinal.Comp_Code) as Comp_Code,XFinal.Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,Zone_Code as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code, isnull([Cash],0) as [Cash],isnull([CD],0) as [CD],isnull([CR],0) as [CR] ,isnull([SO],0) as [SO] ,isnull ([Cash],0)+isnull([CD],0)+isnull([CR],0)+isnull([SO],0) as Total,0 as [Cash_Scheme],0 as [CD_Scheme],0 as [CR_Scheme],0 as [SO_Scheme],0 as Total_Scheme , Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.OldName as Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "  )  Final  pivot  (  max(Booking_Qty) for Booking_Type in ([Cash],[CD],[CR],[SO])   )as pivots  Union All  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code,0 as Cash, 0 as CD, 0 as CR,0 as SO,0 as  Total,   isnull([Cash_Scheme],0) as [Cash_Scheme] ,isnull([CD_Scheme],0) as [CD_Scheme] ,isnull([CR_Scheme],0) as [CR_Scheme] ,isnull([SO_Scheme],0) as [SO_Scheme] ,isnull ([Cash_Scheme],0)+isnull([CD_Scheme],0)+isnull([CR_Scheme],0)+isnull([SO_Scheme],0) as Total_Scheme,Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end+'_Scheme' as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name as Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  where Scheme_Item = 'Y' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "   )  Final   pivot  (  max(Booking_Qty) for Booking_Type in ([Cash_Scheme],[CD_Scheme],[CR_Scheme],[SO_Scheme])   )as pivots  ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no, XFinal.Cust_Code,XFinal.Item_Code,XFinal.Unit_code  ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal  left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code   ) XXXXX   " &
    '                 " Union All " &
    '                 " select  XXXXX.Document_Date, XXXXX.Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax, '2' as  DataSNO, XXXXX.Unit_code as  Unit_code,  0 as  LR_PQ_ES , 0 as OpencrateQty , 0 as CrateQtyRecd , 0 as CrateOutQty , 0 as CrateAdjQty , 0 as CrateQtyClosing , '' as Total_Scheme_Star , 0 as Qty_in_KG , 0 as Qty_in_Ltr from (   Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity,(StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty ,  'Scheme '+XXXFinal.Alies_Name as  Alies_Name_on_Name from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select XFinal.Document_Date,max (XFinal.Comp_Code) as Comp_Code,XFinal.Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,Zone_Code as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code,0 as Cash, 0 as CD, 0 as CR,0 as SO,0 as  Total,   [Cash_Scheme],[CD_Scheme],[CR_Scheme],[SO_Scheme],isnull ([Cash_Scheme],0)+isnull([CD_Scheme],0)+isnull([CR_Scheme],0)+isnull([SO_Scheme],0) as Total_Scheme,Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end+'_Scheme' as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No   left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code   left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code   left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No   left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code   where Scheme_Item = 'Y' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "   )  Final   pivot  (   max(Booking_Qty) for Booking_Type in ([Cash_Scheme],[CD_Scheme],[CR_Scheme],[SO_Scheme])   )as pivots  ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no, XFinal.Cust_Code,XFinal.Item_Code,XFinal.Unit_code   ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code ) XXXXX " &
    '                 " Union All " &
    '                 " select  XXXXX.Document_Date, 'Route Total' as Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,XXXXX.CEILING_Crate_Qty  as CrateQty_New , PendingPcsQty as PendingPcsQty_New,XXXXX.Amount_with_Tax, '3' as  DataSNO, '' as  Unit_code ,  0 as  LR_PQ_ES , 0 as OpencrateQty , 0 as CrateQtyRecd , 0 as CrateOutQty , 0 as CrateAdjQty , 0 as CrateQtyClosing, '' as Total_Scheme_Star , 0 as Qty_in_KG , 0 as Qty_in_Ltr from (    Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty, Convert (int,CrateQty) as CrateQty_New , (XXXFinal.Total) - ((StockPcs.Conversion_Factor * Convert (int,CrateQty))) as PendingPcsQty_New, (StockPcs.Conversion_Factor * Convert (int,CrateQty)) as TotalCrateCapicity_New  from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select max(XFinal.Document_Date) as Document_Date,max (XFinal.Comp_Code) as Comp_Code,max(XFinal.Cust_Code) as Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,max(Zone_Code) as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code, isnull([Cash],0) as [Cash],isnull([CD],0) as [CD],isnull([CR],0) as [CR] ,isnull([SO],0) as [SO] ,isnull ([Cash],0)+isnull([CD],0)+isnull([CR],0)+isnull([SO],0) as Total,0 as [Cash_Scheme],0 as [CD_Scheme],0 as [CR_Scheme],0 as [SO_Scheme],0 as Total_Scheme , Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  where Scheme_Item = 'N'  and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "  )  Final  pivot  (  max(Booking_Qty) for Booking_Type in ([Cash],[CD],[CR],[SO])   )as pivots    ) XFinal Group by  XFinal.route_no, XFinal.Item_Code,Unit_code  ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal  left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code  )XXXXX " &
    '                 "  " &
    '                 " Union all  " &
    '                 " select  XXXXX.Document_Date, XXXXX.Cust_Code as Cust_Code,'Customer Leakage' as Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax, '22' as  DataSNO, XXXXX.Unit_code as  Unit_code  , LR_PQ_ES , 0 as OpencrateQty , 0 as CrateQtyRecd , 0 as CrateOutQty , 0 as CrateAdjQty , 0 as CrateQtyClosing, '' as Total_Scheme_Star, 0 as Qty_in_KG , 0 as Qty_in_Ltr from ( Select Document_Date,Cust_Code, max(Customer_Name) as Customer_Name ,Route_No,max(Route_Desc) as Route_Desc,Zone_Code,max(Zone_Name) as Zone_Name  , Item_Code, max (Alies_Name) as Alies_Name , sum (CR) as CR,sum (CD) as CD,Sum(SO) as SO ,sum (Cash) as Cash, sum(Total) as Total ,sum( CrateQty_New) as CrateQty_New,sum( PendingPcsQty_New) as PendingPcsQty_New, sum (Amount_with_Tax) as Amount_with_Tax,Unit_code,sum (LR_PQ_ES) as LR_PQ_ES  from ( Select convert (varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as Document_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Code as Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_SD_SHIPMENT_HEAD.Route_No ,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name , TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name  , 0 as CR,0 as CD,0 as SO,0 as Cash,0 as Total,0  as CrateQty_New , 0 as PendingPcsQty_New,0 as Amount_with_Tax, 'PCS' as  Unit_code , Convert (decimal(18,2), ( isnull( Source_Conv.Conversion_Factor,0) / nullif ( isnull(Target_Conv.Conversion_Factor,0),0) ) *  TSPL_SD_SHIPMENT_DETAIL.Qty) as LR_PQ_ES from TSPL_SD_SHIPMENT_DETAIL Left  Outer Join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.Document_Code  Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code left Outer Join TSPL_ROUTE_MASTER on TSPL_SD_SHIPMENT_HEAD.Route_No=TSPL_ROUTE_MASTER.Route_No   left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code left Outer Join TSPL_ITEM_MASTER on   TSPL_ITEM_MASTER.Item_code = TSPL_SD_SHIPMENT_DETAIL.Item_Code left Outer Join TSPL_ITEM_UOM_DETAIL as Source_Conv on Source_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Source_Conv.UOM_Code = TSPL_SD_SHIPMENT_DETAIL.Unit_Code left Outer join TSPL_ITEM_UOM_DETAIL as Target_Conv on Target_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Target_Conv.UOM_Code ='PCS' inner join ( Select distinct TSPL_BOOKING_DETAIL.Cust_Code  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No    left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  and " + strCondition + "  " + whr + " ) TBL_VALID_Cust on TBL_VALID_Cust.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code where  TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and TSPL_SD_SHIPMENT_HEAD.IsReplacement = 1  and Convert(date, TSPL_SD_SHIPMENT_HEAD.Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) and TSPL_SD_SHIPMENT_HEAD.status = 1   " + whrShipment + " ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no, XFinal.Cust_Code,XFinal.Item_Code,XFinal.Unit_code ) XXXXX  " + isTrueFalse + "  " &
    '                 " Union all " &
    '                 " select  XXXXX.Document_Date, XXXXX.Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax, '4' as  DataSNO, XXXXX.Unit_code as  Unit_code  , LR_PQ_ES , 0 as OpencrateQty , 0 as CrateQtyRecd , 0 as CrateOutQty , 0 as CrateAdjQty , 0 as CrateQtyClosing, '' as Total_Scheme_Star , 0 as Qty_in_KG , 0 as Qty_in_Ltr from ( Select Document_Date, 'Total Route Leakage' as Cust_Code, 'Total Route Leakage' as Customer_Name ,Route_No,max(Route_Desc) as Route_Desc,Zone_Code,max(Zone_Name) as Zone_Name  , Item_Code, max (Alies_Name) as Alies_Name , sum (CR) as CR,sum (CD) as CD,Sum(SO) as SO ,sum (Cash) as Cash, sum(Total) as Total ,sum( CrateQty_New) as CrateQty_New,sum( PendingPcsQty_New) as PendingPcsQty_New, sum (Amount_with_Tax) as Amount_with_Tax,Unit_code,sum (LR_PQ_ES) as LR_PQ_ES  from ( Select convert (varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as Document_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Code as Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_SD_SHIPMENT_HEAD.Route_No ,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name , TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name  , 0 as CR,0 as CD,0 as SO,0 as Cash,0 as Total,0  as CrateQty_New , 0 as PendingPcsQty_New,0 as Amount_with_Tax, 'PCS' as  Unit_code , Convert (decimal(18,2), ( isnull( Source_Conv.Conversion_Factor,0) / nullif ( isnull(Target_Conv.Conversion_Factor,0),0) ) *  TSPL_SD_SHIPMENT_DETAIL.Qty) as LR_PQ_ES  from TSPL_SD_SHIPMENT_DETAIL Left  Outer Join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.Document_Code  Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code left Outer Join TSPL_ROUTE_MASTER on TSPL_SD_SHIPMENT_HEAD.Route_No=TSPL_ROUTE_MASTER.Route_No   left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code left Outer Join TSPL_ITEM_MASTER on   TSPL_ITEM_MASTER.Item_code = TSPL_SD_SHIPMENT_DETAIL.Item_Code left Outer Join TSPL_ITEM_UOM_DETAIL as Source_Conv on Source_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Source_Conv.UOM_Code = TSPL_SD_SHIPMENT_DETAIL.Unit_Code left Outer join TSPL_ITEM_UOM_DETAIL as Target_Conv on Target_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Target_Conv.UOM_Code ='PCS' inner join ( Select distinct TSPL_BOOKING_DETAIL.Cust_Code  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No    left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  and " + strCondition + "  " + whr + ") TBL_VALID_Cust on TBL_VALID_Cust.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code where  TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others') and TSPL_SD_SHIPMENT_HEAD.IsReplacement = 1  and Convert(date, TSPL_SD_SHIPMENT_HEAD.Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) and TSPL_SD_SHIPMENT_HEAD.status = 1   " + whrShipment + " ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no,XFinal.Item_Code,XFinal.Unit_code ) XXXXX  " + isTrueFalse + "  " &
    '                 " Union all " &
    '                 " Select  XXXXX2.Document_Date,max( XXXXX2.Cust_Code) as Cust_Code ,max (XXXXX2.Customer_Name) as  Customer_Name,XXXXX2.route_no,max(XXXXX2.Route_Desc) as Route_Desc,max (XXXXX2.Zone_Code) as Zone_Code,max(XXXXX2.Zone_Name) as Zone_Name, max(XXXXX2.Item_Code) as Item_Code,max(XXXXX2.Alies_Name) as Alies_Name,sum(XXXXX2.CR) as CR ,Sum(XXXXX2.CD) as CD,sum(XXXXX2.SO) as SO,sum(XXXXX2.Cash) as Cash,sum(XXXXX2.Total) as Total,sum(XXXXX2.CrateQty_New)  as CrateQty_New , sum(XXXXX2.PendingPcsQty_New) as PendingPcsQty_New,sum(XXXXX2.Amount_with_Tax) as Amount_with_Tax, max(XXXXX2.DataSNO) as  DataSNO, max(XXXXX2.Unit_code) as  Unit_code, max(XXXXX2.LR_PQ_ES) as  LR_PQ_ES  , 0 as OpencrateQty , 0 as CrateQtyRecd , 0 as CrateOutQty , 0 as CrateAdjQty , 0 as CrateQtyClosing , '' as Total_Scheme_Star,  sum(XXXXX2.Qty_in_KG) as  Qty_in_KG, sum (Qty_in_Ltr) as Qty_in_Ltr from (  select  XXXXX.Document_Date, 'Total Route Summary' as Cust_Code,'Total Route Summary'as Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,'' as Item_Code,'' as Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,XXXXX.CrateQty_New  as CrateQty_New , PendingPcsQty as PendingPcsQty_New,XXXXX.Amount_with_Tax, '5' as  DataSNO, '' as  Unit_code, 0 as  LR_PQ_ES, Total_KG as Qty_in_KG , Total_LTR as  Qty_in_Ltr from (    Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty  ,Convert (int,CEILING_Crate_Qty) as CrateQty_New , (XXXFinal.Total) - ((StockPcs.Conversion_Factor * Convert (int,CrateQty))) as PendingPcsQty_New, (StockPcs.Conversion_Factor * Convert (int,CrateQty)) as TotalCrateCapicity_New   from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select max(XFinal.Document_Date) as Document_Date,max (XFinal.Comp_Code) as Comp_Code,max(XFinal.Cust_Code) as Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,max(Zone_Code) as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,max (Unit_Code_For_Create) as Unit_Code_For_Create,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item, sum (Total_KG) as Total_KG, sum (Total_LTR) as Total_LTR from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code,Unit_Code_For_Create as Unit_Code_For_Create,  isnull([Cash],0) as [Cash],isnull([CD],0) as [CD],isnull([CR],0) as [CR] ,isnull([SO],0) as [SO] , isnull ([Cash],0)+isnull([CD],0)+isnull([CR],0)+isnull([SO],0)  as Total,0 as [Cash_Scheme],0 as [CD_Scheme],0 as [CR_Scheme],0 as [SO_Scheme],0 as Total_Scheme , Amount_with_Tax,Scheme_Item, isnull ([Cash_KG],0)+isnull([CD_KG],0)+isnull([CR_KG],0)+isnull([SO_KG],0) as Total_KG,  isnull ([Cash_LTR],0)+isnull([CD_LTR],0)+isnull([CR_LTR],0)+isnull([SO_LTR],0) as Total_LTR from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code, case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end+'_KG' as  Booking_Type_KG,	case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end+'_LTR'	as  Booking_Type_LTR, case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end as Booking_Type,TSPL_BOOKING_DETAIL.Booking_Qty ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name, Convert (decimal(18,2), round( ((isnull (Source_Conv.Conversion_Factor,0)/ nullif ( isnull (Target_Conv.Conversion_Factor,0),0) ) * TSPL_BOOKING_DETAIL.Booking_Qty),2,1)) as Booking_Qty_Ltr, case when isnull (Target_Conv.Conversion_Factor,0) = 0 then Convert (decimal(18,2),  round(((isnull (Source_Conv.Conversion_Factor,0)/ nullif ( isnull (Target_Conv_KG.Conversion_Factor,0),0) ) * TSPL_BOOKING_DETAIL.Booking_Qty),2,1)) end as Booking_Qty_Kg ,  TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,'LTR' as Unit_Code,TSPL_BOOKING_DETAIL.Unit_Code as Unit_Code_For_Create,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  Left outer Join TSPL_ITEM_UOM_DETAIL as Source_Conv on 	Source_Conv.Item_Code = TSPL_BOOKING_DETAIL.Item_Code and 	Source_Conv.UOM_Code = TSPL_BOOKING_DETAIL.Unit_code Left Outer Join TSPL_ITEM_UOM_DETAIL as Target_Conv on   Target_Conv.Item_Code = TSPL_BOOKING_DETAIL.Item_Code and Target_Conv.Uom_code = 'LTR'  Left Outer Join TSPL_ITEM_UOM_DETAIL as Target_Conv_KG on   Target_Conv_KG.Item_Code = TSPL_BOOKING_DETAIL.Item_Code and Target_Conv_KG.Uom_code = 'KG'  where Scheme_Item = 'N'  and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)   and " + strCondition + " " + whr + "  )  Final  pivot  (  max(Booking_Qty) for Booking_Type in ([Cash],[CD],[CR],[SO])   )as pivotss  pivot  (  max(Booking_Qty_Kg) for Booking_Type_KG in ([Cash_KG],[CD_KG],[CR_KG],[SO_KG])   )as pivots  pivot  (  max(Booking_Qty_Ltr) for Booking_Type_LTR in ([Cash_LTR],[CD_LTR],[CR_LTR],[SO_LTR])   )as pivotsLtr ) XFinal Group by  XFinal.route_no, XFinal.Item_Code,Unit_code  ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_Code_For_Create =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal  left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code  )  XXXXX  ) XXXXX2 group by Document_Date, route_no  " &
    '                 " Union All  " &
    '                 "  select max (XXXXX.Date) as Document_Date, max([Customer Code]) as  Cust_Code, max ([Customer Name]) as Customer_Name , XXXXX.[Route Code] as route_no ,max( XXXXX.[Route Name]) as Route_Desc,'' as Zone_Code,'' as Zone_Name,'ItemCode' as Item_Code,'' as Alies_Name,0 as CR,0 as CD,0 as SO,0 as Cash,0 as Total,0  as CrateQty_New , 0 as PendingPcsQty_New,0 as Amount_with_Tax, '6' as  DataSNO, '' as    Unit_code  , 1 as LR_PQ_ES     , sum (OpencrateQty) as OpencrateQty , sum (CrateQtyRecd) as CrateQtyRecd , sum (CrateOutQty) as CrateOutQty ,sum (CrateAdjQty) as CrateAdjQty , sum (CrateQtyClosing) as CrateQtyClosing ,'' as Total_Scheme_Star, 0 as Qty_in_KG , 0 as Qty_in_Ltr  from (   Select convert(varchar,Doc_Date,103) as Date,Route_No as [Route Code],Route_Desc as [Route Name], Vehicle_Code as [Vehicle Code],Vehicle_Name as [Vehicle Name], 'XCrates Route Details' as [Customer Code],'XCrates Route Details' as [Customer Name],Customer_Phone as [Customer Phone], ZSM as [ZSM Code], ZSM_Name as [ZSM Name],ZSM_Phone as [ZSM Phone],ASM as [ASM Code],ASM_NAME as [ASM Name],ASM_Phone as [ASM Phone],ASO as [ASO Code],ASO_Name as [ASO Name],ASO_Phone as [ASO Phone], OpencrateQty,OpenJaaliQty,OpenBoxQty,OpenCanQty, CrateQtyRecd, JaaliQtyRecd,BoxQtyRecd, CanQtyRecd  ,CrateOutQty, jaaliOutQty,boxOutQty,CanOutQty, CrateAdjQty , JaaliAdjQty  , BoxAdjQty,CanAdjQty, OpencrateQty+CrateOutQty -CrateQtyRecd-CrateAdjQty as CrateQtyClosing, OpenCanQty+CanOutQty-CanQtyRecd-CanAdjQty as CanQtyClosing , OpenJaaliQty+jaaliOutQty-JaaliQtyRecd-JaaliAdjQty as JaaliQtyClosing, OpenBoxQty+boxOutQty-BoxQtyRecd-BoxAdjQty as BoxQtyClosing, (OpencrateQty+CrateOutQty -CrateQtyRecd-CrateAdjQty)* 0 as CrateValueClosing, (OpenJaaliQty+jaaliOutQty-JaaliQtyRecd-JaaliAdjQty)* 0 as JaaliValueClosing, (OpenBoxQty+boxOutQty-BoxQtyRecd-BoxAdjQty)* 0 as BoxValueClosing ,  (OpenCanQty+CanOutQty-CanQtyRecd-CanAdjQty)* 0 as CanValueClosing ,OpencrateQtySum,OpenCanQtySum ,OpencrateQtySumOfVehicleWise,OpencanQtySumOfVehicleWise  from (Select  case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) else 0 end OpencrateQtySum, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencanQty+ISNULL(CT1.CanQtyClosing,0) else 0 end OpencanQtySum, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) else 0 end OpencrateQtySumOfVehicleWise, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencanQty+ISNULL(CT1.CanQtyClosing,0) else 0 end OpencanQtySumOfVehicleWise, CTETemp.RowNo,CTETemp.Doc_Date, CTETemp.Route_No,CTETemp.Route_Desc ,CTETemp.Vehicle_Code,CTETemp.Vehicle_Name ,CTETemp.Customer_Code,CTETemp.Customer_Name, CTETemp.Customer_Phone, CTETemp.ZSM,CTETemp.ZSM_NAME,CTETemp.ZSM_Phone,CTETemp.ASM,CTETemp.ASM_NAME,CTETemp.ASM_Phone,CTETemp.ASO,CTETemp.ASO_NAME,CTETemp.ASO_Phone, CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) as OpencrateQty,  CTETemp.OpenJaaliQty+ISNULL(CT1.JaaliQtyClosing,0) as OpenJaaliQty,  CTETemp.OpenBoxQty+ISNULL(CT1.BoxQtyClosing,0) as OpenBoxQty, CTETemp.OpenCanQty+ISNULL(CT1.CanQtyClosing,0) as OpenCanQty,  CTETemp.CrateQtyRecd, CTETemp.JaaliQtyRecd, CTETemp.BoxQtyRecd,CTETemp.CanQtyRecd, CTETemp.CrateOutQty, CTETemp.jaaliOutQty, CTETemp.boxOutQty, CTETemp.CanOutQty, CTETemp.CrateAdjQty , CTETemp.JaaliAdjQty  , CTETemp.BoxAdjQty, CTETemp.CanAdjQty  from CTETemp LEFt OUTER JOIN CTETemp CT1 ON  CT1.Customer_Code=CTETemp.Customer_Code and  CT1.Vehicle_Code = CTETemp.Vehicle_Code  and CT1.Route_No  = CTETemp.Route_No  AND (CTETemp.RowNo-CT1.RowNo)=1 ) ZZZ   ) XXXXX  " + isTrueFalse + " Group by  XXXXX.[Route Code]   " &
    '                 "  " &
    '                 " )XXXXXFinal left Outer Join " &
    '                 " ( " &
    '                 "  select XXXXXFinal.DataSNO ,XXXXXFinal.route_no,XXXXXFinal.Cust_Code , Count (Item_Code) as No_Of_Item  from ( " &
    '                 "  select '1' as  DataSNO, XXXXX.Document_Date, XXXXX.Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,XXXXX.CrateQty_New as CrateQty_New , XXXXX.PendingPcsQty_New as PendingPcsQty_New,isnull (XXXXX.Amount_with_Tax,0) as Amount_with_Tax,'' as  Unit_code from (  Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity,(StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty , Convert (int,CrateQty) as CrateQty_New , (XXXFinal.Total) - ((StockPcs.Conversion_Factor * Convert (int,CrateQty))) as PendingPcsQty_New, (StockPcs.Conversion_Factor * Convert (int,CrateQty)) as TotalCrateCapicity_New from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select XFinal.Document_Date,max (XFinal.Comp_Code) as Comp_Code,XFinal.Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,Zone_Code as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code, isnull([Cash],0) as [Cash],isnull([CD],0) as [CD],isnull([CR],0) as [CR] ,isnull([SO],0) as [SO] ,isnull ([Cash],0)+isnull([CD],0)+isnull([CR],0)+isnull([SO],0) as Total,0 as [Cash_Scheme],0 as [CD_Scheme],0 as [CR_Scheme],0 as [SO_Scheme],0 as Total_Scheme , Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.OldName as Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "  )  Final  pivot  (  max(Booking_Qty) for Booking_Type in ([Cash],[CD],[CR],[SO])   )as pivots  Union All  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code,0 as Cash, 0 as CD, 0 as CR,0 as SO,0 as  Total,   isnull([Cash_Scheme],0) as [Cash_Scheme] ,isnull([CD_Scheme],0) as [CD_Scheme] ,isnull([CR_Scheme],0) as [CR_Scheme] ,isnull([SO_Scheme],0) as [SO_Scheme] ,isnull ([Cash_Scheme],0)+isnull([CD_Scheme],0)+isnull([CR_Scheme],0)+isnull([SO_Scheme],0) as Total_Scheme,Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end+'_Scheme' as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name as Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  where Scheme_Item = 'Y' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  " + whr + "  )  Final   pivot  (  max(Booking_Qty) for Booking_Type in ([Cash_Scheme],[CD_Scheme],[CR_Scheme],[SO_Scheme])   )as pivots  ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no, XFinal.Cust_Code,XFinal.Item_Code,XFinal.Unit_code  ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal  left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code   ) XXXXX  " &
    '                 " Union All " &
    '                 " select '2' as  DataSNO, XXXXX.Document_Date, XXXXX.Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax, XXXXX.Unit_code as  Unit_code from ( " &
    '                 " Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity,(StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty ,  'Scheme '+XXXFinal.Alies_Name as  Alies_Name_on_Name from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select XFinal.Document_Date,max (XFinal.Comp_Code) as Comp_Code,XFinal.Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,Zone_Code as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code,0 as Cash, 0 as CD, 0 as CR,0 as SO,0 as  Total,   [Cash_Scheme],[CD_Scheme],[CR_Scheme],[SO_Scheme],isnull ([Cash_Scheme],0)+isnull([CD_Scheme],0)+isnull([CR_Scheme],0)+isnull([SO_Scheme],0) as Total_Scheme,Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end+'_Scheme' as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No   left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code   left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code   left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No   left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code   where Scheme_Item = 'Y' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "   )  Final   pivot  (   max(Booking_Qty) for Booking_Type in ([Cash_Scheme],[CD_Scheme],[CR_Scheme],[SO_Scheme])   )as pivots  ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no, XFinal.Cust_Code,XFinal.Item_Code,XFinal.Unit_code   ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code " &
    '                 " ) XXXXX " &
    '                 " Union All " &
    '                 " select '3' as  DataSNO, XXXXX.Document_Date, 'Route Total' as Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax, '' as  Unit_code from ( " &
    '                 " Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty  from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select max(XFinal.Document_Date) as Document_Date,max (XFinal.Comp_Code) as Comp_Code,max(XFinal.Cust_Code) as Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,max(Zone_Code) as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code, isnull([Cash],0) as [Cash],isnull([CD],0) as [CD],isnull([CR],0) as [CR] ,isnull([SO],0) as [SO] ,isnull ([Cash],0)+isnull([CD],0)+isnull([CR],0)+isnull([SO],0) as Total,0 as [Cash_Scheme],0 as [CD_Scheme],0 as [CR_Scheme],0 as [SO_Scheme],0 as Total_Scheme , Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name,TSPL_BOOKING_DETAIL.Booking_Qty, TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,TSPL_BOOKING_DETAIL.Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  where Scheme_Item = 'N'  and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) " + whr + "  )  Final  pivot  (  max(Booking_Qty) for Booking_Type in ([Cash],[CD],[CR],[SO])   )as pivots    ) XFinal Group by  XFinal.route_no, XFinal.Item_Code,Unit_code  ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal  left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code  )   XXXXX  " &
    '                 "   " &
    '                 " Union all  " &
    '                 " select '22' as  DataSNO,  XXXXX.Document_Date, XXXXX.Cust_Code as Cust_Code,'Customer Leakage' as Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax , XXXXX.Unit_code as  Unit_code  from ( Select Document_Date,Cust_Code, max(Customer_Name) as Customer_Name ,Route_No,max(Route_Desc) as Route_Desc,Zone_Code,max(Zone_Name) as Zone_Name  , Item_Code, max (Alies_Name) as Alies_Name , sum (CR) as CR,sum (CD) as CD,Sum(SO) as SO ,sum (Cash) as Cash, sum(Total) as Total ,sum( CrateQty_New) as CrateQty_New,sum( PendingPcsQty_New) as PendingPcsQty_New, sum (Amount_with_Tax) as Amount_with_Tax,Unit_code,sum (LR_PQ_ES) as LR_PQ_ES  from ( Select convert (varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as Document_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Code as Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_SD_SHIPMENT_HEAD.Route_No ,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name , TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name  , 0 as CR,0 as CD,0 as SO,0 as Cash,0 as Total,0  as CrateQty_New , 0 as PendingPcsQty_New,0 as Amount_with_Tax, 'PCS' as  Unit_code , Convert (decimal(18,2), ( isnull( Source_Conv.Conversion_Factor,0) / nullif ( isnull(Target_Conv.Conversion_Factor,0),0) ) *  TSPL_SD_SHIPMENT_DETAIL.Qty) as LR_PQ_ES from TSPL_SD_SHIPMENT_DETAIL Left  Outer Join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.Document_Code  Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code left Outer Join TSPL_ROUTE_MASTER on TSPL_SD_SHIPMENT_HEAD.Route_No=TSPL_ROUTE_MASTER.Route_No   left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code left Outer Join TSPL_ITEM_MASTER on   TSPL_ITEM_MASTER.Item_code = TSPL_SD_SHIPMENT_DETAIL.Item_Code left Outer Join TSPL_ITEM_UOM_DETAIL as Source_Conv on Source_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Source_Conv.UOM_Code = TSPL_SD_SHIPMENT_DETAIL.Unit_Code left Outer join TSPL_ITEM_UOM_DETAIL as Target_Conv on Target_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Target_Conv.UOM_Code ='PCS' inner join ( Select distinct TSPL_BOOKING_DETAIL.Cust_Code  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No   left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  and " + strCondition + " " + whr + " ) TBL_VALID_Cust on TBL_VALID_Cust.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code where  TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others') and TSPL_SD_SHIPMENT_HEAD.IsReplacement = 1  and Convert(date, TSPL_SD_SHIPMENT_HEAD.Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) and TSPL_SD_SHIPMENT_HEAD.status = 1  " + whrShipment + "  ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no, XFinal.Cust_Code,XFinal.Item_Code,XFinal.Unit_code ) XXXXX  " + isTrueFalse + "  " &
    '                 " Union all  " &
    '                 " select  '4' as  DataSNO,  XXXXX.Document_Date, XXXXX.Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,XXXXX.Item_Code,XXXXX.Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,0  as CrateQty_New , 0 as PendingPcsQty_New,XXXXX.Amount_with_Tax,  XXXXX.Unit_code as  Unit_code from ( Select Document_Date, 'Total Route Leakage' as Cust_Code, max(Customer_Name) as Customer_Name ,Route_No,max(Route_Desc) as Route_Desc,Zone_Code,max(Zone_Name) as Zone_Name  , Item_Code, max (Alies_Name) as Alies_Name , sum (CR) as CR,sum (CD) as CD,Sum(SO) as SO ,sum (Cash) as Cash, sum(Total) as Total ,sum( CrateQty_New) as CrateQty_New,sum( PendingPcsQty_New) as PendingPcsQty_New, sum (Amount_with_Tax) as Amount_with_Tax,Unit_code,sum (LR_PQ_ES) as LR_PQ_ES  from ( Select convert (varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as Document_Date,TSPL_SD_SHIPMENT_HEAD.Customer_Code as Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_SD_SHIPMENT_HEAD.Route_No ,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name , TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name  , 0 as CR,0 as CD,0 as SO,0 as Cash,0 as Total,0  as CrateQty_New , 0 as PendingPcsQty_New,0 as Amount_with_Tax, 'PCS' as  Unit_code  , Convert (decimal(18,2), ( isnull( Source_Conv.Conversion_Factor,0) / nullif ( isnull(Target_Conv.Conversion_Factor,0),0) ) *  TSPL_SD_SHIPMENT_DETAIL.Qty) as LR_PQ_ES from TSPL_SD_SHIPMENT_DETAIL Left  Outer Join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.Document_Code  Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code left Outer Join TSPL_ROUTE_MASTER on TSPL_SD_SHIPMENT_HEAD.Route_No=TSPL_ROUTE_MASTER.Route_No   left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code left Outer Join TSPL_ITEM_MASTER on   TSPL_ITEM_MASTER.Item_code = TSPL_SD_SHIPMENT_DETAIL.Item_Code left Outer Join TSPL_ITEM_UOM_DETAIL as Source_Conv on Source_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Source_Conv.UOM_Code = TSPL_SD_SHIPMENT_DETAIL.Unit_Code left Outer join TSPL_ITEM_UOM_DETAIL as Target_Conv on Target_Conv.Item_Code = TSPL_SD_SHIPMENT_DETAIL.Item_Code and Target_Conv.UOM_Code ='PCS' inner join ( Select distinct TSPL_BOOKING_DETAIL.Cust_Code  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  where Scheme_Item = 'N' and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  and " + strCondition + "  " + whr + " ) TBL_VALID_Cust on TBL_VALID_Cust.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code where  TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others') and TSPL_SD_SHIPMENT_HEAD.IsReplacement = 1  and Convert(date, TSPL_SD_SHIPMENT_HEAD.Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103) and TSPL_SD_SHIPMENT_HEAD.status = 1  " + whrShipment + "  ) XFinal Group by XFinal.Document_Date ,XFinal.Zone_Code,XFinal.route_no,XFinal.Item_Code,XFinal.Unit_code ) XXXXX  " + isTrueFalse + " " &
    '                 " Union all  " &
    '                 " Select  max(XXXXX2.DataSNO) as  DataSNO, XXXXX2.Document_Date,max( XXXXX2.Cust_Code) as Cust_Code ,max (XXXXX2.Customer_Name) as  Customer_Name,XXXXX2.route_no,max(XXXXX2.Route_Desc) as Route_Desc,max (XXXXX2.Zone_Code) as Zone_Code,max(XXXXX2.Zone_Name) as Zone_Name, max(XXXXX2.Item_Code) as Item_Code,max(XXXXX2.Alies_Name) as Alies_Name,sum(XXXXX2.CR) as CR ,Sum(XXXXX2.CD) as CD,sum(XXXXX2.SO) as SO,sum(XXXXX2.Cash) as Cash,sum(XXXXX2.Total) as Total,sum(XXXXX2.CrateQty_New)  as CrateQty_New , sum(XXXXX2.PendingPcsQty_New) as PendingPcsQty_New,sum(XXXXX2.Amount_with_Tax) as Amount_with_Tax, max(XXXXX2.Unit_code) as  Unit_code   from (  select  XXXXX.Document_Date, 'Total Route Summary' as Cust_Code,XXXXX.Customer_Name,XXXXX.route_no,XXXXX.Route_Desc,XXXXX.Zone_Code,XXXXX.Zone_Name,'' as Item_Code,'' as Alies_Name,XXXXX.CR,XXXXX.CD,XXXXX.SO,XXXXX.Cash,XXXXX.Total,XXXXX.CrateQty_New  as CrateQty_New , PendingPcsQty as PendingPcsQty_New,XXXXX.Amount_with_Tax, '5' as  DataSNO, '' as  Unit_code, 0 as  LR_PQ_ES from (    Select XXXFinal.*, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) as TotalCrateCapicity, (StockPcs.Conversion_Factor * CEILING_Crate_Qty) -  (XXXFinal.Total) as PendingPcsQty  ,Convert (int,CrateQty) as CrateQty_New , (XXXFinal.Total) - ((StockPcs.Conversion_Factor * Convert (int,CrateQty))) as PendingPcsQty_New, (StockPcs.Conversion_Factor * Convert (int,CrateQty)) as TotalCrateCapicity_New  from (  Select XXFinal.*, (case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end) as CrateQty , CEILING ((case when coalesce(StockCrate.Conversion_Factor,0)=0 then 0 else cast((XXFinal.Total)* (Stock_SU.Conversion_Factor)/(coalesce(StockCrate.Conversion_Factor,1)) as numeric(18,3)) end)) as CEILING_Crate_Qty  from (  select max(XFinal.Document_Date) as Document_Date,max (XFinal.Comp_Code) as Comp_Code,max(XFinal.Cust_Code) as Cust_Code, max( XFinal.Customer_Name ) as Customer_Name , XFinal.route_no, max(XFinal.Route_Desc) as Route_Desc,max(Zone_Code) as Zone_Code ,max(Zone_Name) as Zone_Name ,max(XFinal.Vehicle_Code) as Vehicle_Code , XFinal.Item_Code,max (XFinal.Alies_Name) as Alies_Name , XFinal.Unit_code,  sum (XFinal.Cash_Scheme) as Cash_Scheme, sum(XFinal.CD_Scheme)  as CD_Scheme, sum(XFinal.CR_Scheme) as CR_Scheme , sum (XFinal.SO_Scheme) as SO_Scheme, sum (XFinal.Total_Scheme) as Total_Scheme,  sum (XFinal.Cash) as Cash, sum(XFinal.CD)  as CD, sum(XFinal.CR) as CR , sum (XFinal.SO) as SO, sum (XFinal.Total) + sum (XFinal.Total_Scheme) as Total ,sum (XFinal.Amount_with_Tax)  as Amount_with_Tax , max (Scheme_Item ) as Scheme_Item from (  select Document_No,Document_Date, Comp_Code,From_Screen_Code,Cust_Code,Customer_Name,Route_No,Route_Desc,Zone_Code,Zone_Name,Vehicle_Code,Loc_Code,Item_Code,Alies_Name, Unit_Code, isnull([Cash],0) as [Cash],isnull([CD],0) as [CD],isnull([CR],0) as [CR] ,isnull([SO],0) as [SO] ,isnull ([Cash],0)+isnull([CD],0)+isnull([CR],0)+isnull([SO],0) as Total,0 as [Cash_Scheme],0 as [CD_Scheme],0 as [CR_Scheme],0 as [SO_Scheme],0 as Total_Scheme , Amount_with_Tax,Scheme_Item from (  Select TSPL_BOOKING_MATSER.Document_No,Convert (varchar,TSPL_BOOKING_MATSER.Document_Date,103) as Document_Date,TSPL_BOOKING_MATSER.Comp_Code,TSPL_BOOKING_MATSER.From_Screen_Code,case when  TSPL_BOOKING_MATSER.Booking_Type in ('FN','PS','UP') then 'Cash' else TSPL_BOOKING_MATSER.Booking_Type end as Booking_Type ,TSPL_BOOKING_DETAIL.Route_No,TSPL_ROUTE_MASTER.Route_Desc ,TSPL_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_BOOKING_DETAIL.Loc_Code, TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name, Convert (decimal(18,2),(isnull (Source_Conv.Conversion_Factor,0)/ nullif ( isnull (Target_Conv.Conversion_Factor,0),0) ) * TSPL_BOOKING_DETAIL.Booking_Qty) as Booking_Qty,  TSPL_BOOKING_DETAIL.Price_with_Tax,Case when TSPL_BOOKING_MATSER.Booking_Type in ('CR','CD','SO') then 0 else TSPL_BOOKING_DETAIL.Amount_with_Tax end as Amount_with_Tax,'LTR' as Unit_Code,TSPL_BOOKING_DETAIL.Vehicle_Code ,TSPL_CUSTOMER_MASTER.Zone_Code ,TSPL_Zone_Master.Description as Zone_Name, TSPL_BOOKING_DETAIL.Scheme_Item  from TSPL_BOOKING_DETAIL  left Outer Join TSPL_BOOKING_MATSER on TSPL_BOOKING_DETAIL.Document_No = TSPL_BOOKING_MATSER.Document_No  left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_BOOKING_DETAIL.Item_Code  left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_BOOKING_DETAIL.Cust_Code  left Outer Join TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No  left Outer Join TSPL_Zone_Master on TSPL_Zone_Master.Zone_Code = TSPL_CUSTOMER_MASTER.Zone_Code  Left outer Join TSPL_ITEM_UOM_DETAIL as Source_Conv on 	Source_Conv.Item_Code = TSPL_BOOKING_DETAIL.Item_Code and 	Source_Conv.UOM_Code = TSPL_BOOKING_DETAIL.Unit_code Left Outer Join TSPL_ITEM_UOM_DETAIL as Target_Conv on   Target_Conv.Item_Code = TSPL_BOOKING_DETAIL.Item_Code and Target_Conv.Uom_code = 'LTR' where Scheme_Item = 'N'  and TSPL_CUSTOMER_MASTER.CUSTOMER_CATEGORY not in ('Others')  and Convert(date, Document_Date,103) =  Convert(date, '" & txtDate.Value & "',103)  and " + strCondition + "  " + whr + " )  Final  pivot  (  max(Booking_Qty) for Booking_Type in ([Cash],[CD],[CR],[SO])   )as pivots    ) XFinal Group by  XFinal.route_no, XFinal.Item_Code,Unit_code  ) XXFinal   left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXFinal.Item_Code =Stock_SU.Item_Code and XXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockCrate on XXFinal.Item_Code =StockCrate.Item_Code  ) XXXFinal  left join (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL ) as Stock_SU on XXXFinal.Item_Code =Stock_SU.Item_Code and XXXFinal.Unit_code =Stock_SU.UOM_Code  left outer join  (select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where UOM_Code='CRATE') as StockPcs on XXXFinal.Item_Code =StockPcs.Item_Code  )  XXXXX  ) XXXXX2 group by Document_Date, route_no   " &
    '                 " Union All  " &
    '                 " select max (XXXXX.Date) as Document_Date, max([Customer Code]) as  Cust_Code, max ([Customer Name]) as Customer_Name , XXXXX.[Route Code] as route_no ,max( XXXXX.[Route Name]) as Route_Desc, '' as Zone_Code,'' as Zone_Name,'ItemCode' as Item_Code,'' as Alies_Name,0 as CR,0 as CD,0 as SO,0 as Cash,0 as Total,0  as CrateQty_New , 0 as PendingPcsQty_New,0 as Amount_with_Tax, '6' as  DataSNO, '' as    Unit_code    from (   Select convert(varchar,Doc_Date,103) as Date,Route_No as [Route Code],Route_Desc as [Route Name], Vehicle_Code as [Vehicle Code],Vehicle_Name as [Vehicle Name], 'XCrates Route Details' as [Customer Code],Customer_Name as [Customer Name],Customer_Phone as [Customer Phone], ZSM as [ZSM Code], ZSM_Name as [ZSM Name],ZSM_Phone as [ZSM Phone],ASM as [ASM Code],ASM_NAME as [ASM Name],ASM_Phone as [ASM Phone],ASO as [ASO Code],ASO_Name as [ASO Name],ASO_Phone as [ASO Phone], OpencrateQty,OpenJaaliQty,OpenBoxQty,OpenCanQty, CrateQtyRecd, JaaliQtyRecd,BoxQtyRecd, CanQtyRecd  ,CrateOutQty, jaaliOutQty,boxOutQty,CanOutQty, CrateAdjQty , JaaliAdjQty  , BoxAdjQty,CanAdjQty, OpencrateQty+CrateOutQty -CrateQtyRecd-CrateAdjQty as CrateQtyClosing, OpenCanQty+CanOutQty-CanQtyRecd-CanAdjQty as CanQtyClosing , OpenJaaliQty+jaaliOutQty-JaaliQtyRecd-JaaliAdjQty as JaaliQtyClosing, OpenBoxQty+boxOutQty-BoxQtyRecd-BoxAdjQty as BoxQtyClosing, (OpencrateQty+CrateOutQty -CrateQtyRecd-CrateAdjQty)* 0 as CrateValueClosing, (OpenJaaliQty+jaaliOutQty-JaaliQtyRecd-JaaliAdjQty)* 0 as JaaliValueClosing, (OpenBoxQty+boxOutQty-BoxQtyRecd-BoxAdjQty)* 0 as BoxValueClosing ,  (OpenCanQty+CanOutQty-CanQtyRecd-CanAdjQty)* 0 as CanValueClosing ,OpencrateQtySum,OpenCanQtySum ,OpencrateQtySumOfVehicleWise,OpencanQtySumOfVehicleWise  from (Select  case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) else 0 end OpencrateQtySum, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencanQty+ISNULL(CT1.CanQtyClosing,0) else 0 end OpencanQtySum, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) else 0 end OpencrateQtySumOfVehicleWise, case when CTETemp.rowno= MIN(CTETemp.RowNo ) OVER (PARTITION BY CTETemp.Customer_Code) then CTETemp.OpencanQty+ISNULL(CT1.CanQtyClosing,0) else 0 end OpencanQtySumOfVehicleWise, CTETemp.RowNo,CTETemp.Doc_Date, CTETemp.Route_No,CTETemp.Route_Desc ,CTETemp.Vehicle_Code,CTETemp.Vehicle_Name ,CTETemp.Customer_Code,CTETemp.Customer_Name, CTETemp.Customer_Phone, CTETemp.ZSM,CTETemp.ZSM_NAME,CTETemp.ZSM_Phone,CTETemp.ASM,CTETemp.ASM_NAME,CTETemp.ASM_Phone,CTETemp.ASO,CTETemp.ASO_NAME,CTETemp.ASO_Phone, CTETemp.OpencrateQty+ISNULL(CT1.CrateQtyClosing,0) as OpencrateQty,  CTETemp.OpenJaaliQty+ISNULL(CT1.JaaliQtyClosing,0) as OpenJaaliQty,  CTETemp.OpenBoxQty+ISNULL(CT1.BoxQtyClosing,0) as OpenBoxQty, CTETemp.OpenCanQty+ISNULL(CT1.CanQtyClosing,0) as OpenCanQty,  CTETemp.CrateQtyRecd, CTETemp.JaaliQtyRecd, CTETemp.BoxQtyRecd,CTETemp.CanQtyRecd, CTETemp.CrateOutQty, CTETemp.jaaliOutQty, CTETemp.boxOutQty, CTETemp.CanOutQty, CTETemp.CrateAdjQty , CTETemp.JaaliAdjQty  , CTETemp.BoxAdjQty, CTETemp.CanAdjQty  from CTETemp LEFt OUTER JOIN CTETemp CT1 ON  CT1.Customer_Code=CTETemp.Customer_Code and  CT1.Vehicle_Code = CTETemp.Vehicle_Code  and CT1.Route_No  = CTETemp.Route_No  AND (CTETemp.RowNo-CT1.RowNo)=1 ) ZZZ   ) XXXXX  " + isTrueFalse + "  Group by  XXXXX.[Route Code]   " &
    '                 "   " &
    '                 "   " &
    '                 "   " &
    '                 " )XXXXXFinal  group by XXXXXFinal.route_no, XXXXXFinal.Cust_Code, XXXXXFinal.DataSNO   " &
    '                 " ) " &
    '                 " TBL_Count on TBL_Count.DataSNO = XXXXXFinal.DataSNO and TBL_Count.route_no = XXXXXFinal.route_no and TBL_Count.Cust_Code = XXXXXFinal.Cust_Code " &
    '                 " left Outer  Join  TSPL_ITEM_MASTER  on  TSPL_ITEM_MASTER.Item_Code = XXXXXFinal.Item_Code " &
    '                 " Order By XXXXXFinal.route_no,XXXXXFinal.Cust_Code,XXXXXFinal.DataSNO , case when TSPL_ITEM_MASTER.SKU_Seq = 0 then 10000 else  TSPL_ITEM_MASTER.SKU_Seq end  "
    '        Dim dtMain As DataTable = clsDBFuncationality.GetDataTable(qry)
    '        If dtMain IsNot Nothing AndAlso dtMain.Rows.Count > 0 Then
    '            Dim subPath As String = clsCommon.myCstr(objCommonVar.ImportExportDrive) + ":\ERPTempFolder" + "\" + objCommonVar.CurrDatabase + "\" + objCommonVar.CurrentUserCode + "\Downloads"
    '            Dim IsExists As Boolean = System.IO.Directory.Exists(subPath)
    '            If (Not IsExists) Then
    '                System.IO.Directory.CreateDirectory(subPath)
    '            End If
    '            subPath += "\\" & clsCommon.GetPrintDate(DateTime.Now, "yyyyMMddhhmmsss") & ".Txt"
    '            IsExists = System.IO.File.Exists(subPath)
    '            If IsExists Then
    '                System.IO.File.Delete(subPath)
    '            End If
    '            WriteDataToFile(dtMain, subPath, True)
    '            Process.Start(subPath)
    '        Else
    '            clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
    '        End If
    '    End If
    'End Sub
    'Sub WriteDataToFile(ByVal submittedDataTable As DataTable, ByVal submittedFilePath As String, Optional isGatePass As Boolean = False)
    '    Try
    '        Dim i As Integer = 0
    '        'Dim ii As Integer = 0
    '        Dim iH As Integer = 0
    '        Dim PageNo As Integer = 1
    '        Dim OnePageTotalLine As Integer = 60
    '        Dim TotalLineNoPrint As Integer = 0
    '        Dim NextTotalNoOnPage As Integer = 0
    '        Dim PrvTotalNoOnPage As Integer = 0
    '        Dim sw As StreamWriter = Nothing
    '        sw = New StreamWriter(submittedFilePath, False)
    '        Dim strPrvRouteNo As String = ""
    '        'Dim strPrvRouteDesc As String = ""
    '        Dim strPrvZoneCode As String = ""
    '        'Dim strPrvZoneDesc As String = ""
    '        Dim strPrvCustomerNo As String = ""
    '        Dim strPrvCustomerDesc As String = ""
    '        Dim strPrvCustomerDescForRoutetotal As String = ""
    '        'Dim strPrvDocDate As String = ""
    '        Dim SumOfCrate As Integer = 0
    '        Dim SumOfAmount As Double = 0.0
    '        Dim SumOfTCSAmount As Double = 0.0
    '        Dim SumOfTotlQty As Double = 0.0
    '        Dim DataSNO As String = ""
    '        Dim strCustomerLR_PQ_ES As String = ""
    '        Dim strTotalRouteLR_PQ_ES As String = ""
    '        Dim strTotalRouteSummary As String = ""
    '        Dim strXCratesRouteDetails As String = ""
    '        Dim rowsCount As Integer = 0
    '        Dim rowsCustLR_PQ_ES_Count As Integer = 0
    '        Dim strCustomerScheme As String = ""
    '        Dim rowsCustomerSchemeCount As Integer = 0
    '        Dim checkTotalRoutWise As Boolean = False
    '        Dim qry As String
    '        Dim dblTotalTCS As Decimal = 0
    '        Dim strTransporterName As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("  select SUBSTRING ( TSPL_TRANSPORT_MASTER.Transporter_Name,0,19) as Transporter_Name from TSPL_ROUTE_MASTER left outer Join TSPL_VEHICLE_MASTER  on TSPL_VEHICLE_MASTER.Vehicle_Id = TSPL_ROUTE_MASTER.vehicle_code left outer join TSPL_TRANSPORT_MASTER  on  TSPL_TRANSPORT_MASTER.Transport_Id = TSPL_VEHICLE_MASTER.Transport_Id where Route_No = '" & txtRouteNo.Value & "'"))
    '        Dim strTransportercode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("  select TSPL_VEHICLE_MASTER.Transport_Id from TSPL_ROUTE_MASTER left outer Join TSPL_VEHICLE_MASTER  on TSPL_VEHICLE_MASTER.Vehicle_Id = TSPL_ROUTE_MASTER.vehicle_code left outer join TSPL_TRANSPORT_MASTER  on  TSPL_TRANSPORT_MASTER.Transport_Id = TSPL_VEHICLE_MASTER.Transport_Id where Route_No = '" & txtRouteNo.Value & "'"))
    '        Dim strDocDate As String = clsCommon.myCstr(txtDate.Text)
    '        Dim strVehicleCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("  select TSPL_ROUTE_MASTER.vehicle_code from TSPL_ROUTE_MASTER left outer Join TSPL_VEHICLE_MASTER  on TSPL_VEHICLE_MASTER.Vehicle_Id = TSPL_ROUTE_MASTER.vehicle_code left outer join TSPL_TRANSPORT_MASTER  on  TSPL_TRANSPORT_MASTER.Transport_Id = TSPL_VEHICLE_MASTER.Transport_Id where Route_No = '" & txtRouteNo.Value & "'"))
    '        Dim strDocTime As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("    select  format (Document_Date, 'hh:mm tt') as DocTime from TSPL_BOOKING_MATSER where Document_No = '" & txtDocNo.Value & "'"))
    '        Dim countRecordRouteTotal As Integer = 0
    '        For Each rowH As DataRow In submittedDataTable.Rows
    '            Dim arrayH As Object() = rowH.ItemArray
    '            'Dim strItemCode As String = ""
    '            Dim strItemDesc As String = ""
    '            Dim strCR As String = ""
    '            Dim strCD As String = ""
    '            Dim strSO As String = ""
    '            Dim strCash As String = ""
    '            Dim strTotal As String = ""
    '            'Dim strCrate As String = ""
    '            'Dim strAmount As String = ""
    '            Dim strtotalCrateQty As Integer = 0
    '            Dim strTotalPendingPcsQty As Integer = 0
    '            Dim strTotalAmount As Double = 0.0
    '            Dim strTCSTotalAmount As Double = 0.0
    '            Dim strCurRouteNo As String = ""
    '            Dim strCurRouteDesc As String = ""
    '            Dim strCurZoneCode As String = ""
    '            Dim strCurZoneDesc As String = ""
    '            Dim strCurCustomerNo As String = ""
    '            Dim strCurCustomerDesc As String = ""
    '            Dim strCurDocDate As String = ""
    '            Dim strCurUnitCode As String = ""
    '            Dim strCurItemRowNo As String = ""
    '            Dim strLR_PQ_ES As String = ""
    '            Dim strOpencrateQty As String = ""
    '            Dim strCrateQtyRecd As String = ""
    '            Dim strCrateOutQty As String = ""
    '            Dim strCrateAdjQty As String = ""
    '            Dim strCrateQtyClosing As String = ""
    '            Dim strSchemeStar As String = ""
    '            Dim strTotal_Qty_In_LTR As Double = 0
    '            Dim strTotal_Qty_In_KG As Double = 0
    '            For iH = 0 To arrayH.Length - 1
    '                If iH = 3 Then
    '                    strCurRouteNo = arrayH(3).ToString()
    '                ElseIf iH = 4 Then
    '                    strCurRouteDesc = arrayH(4).ToString()
    '                ElseIf iH = 5 Then
    '                    strCurZoneCode = arrayH(5).ToString()
    '                ElseIf iH = 6 Then
    '                    strCurZoneDesc = arrayH(6).ToString()
    '                ElseIf iH = 1 Then
    '                    strCurCustomerNo = arrayH(1).ToString()
    '                ElseIf iH = 2 Then
    '                    strCurCustomerDesc = arrayH(2).ToString()
    '                ElseIf iH = 0 Then
    '                    strCurDocDate = arrayH(0).ToString()
    '                ElseIf iH = 8 Then
    '                    strItemDesc = arrayH(8).ToString()
    '                ElseIf iH = 9 Then
    '                    strCR = arrayH(9).ToString()
    '                ElseIf iH = 10 Then
    '                    strCD = arrayH(10).ToString()
    '                ElseIf iH = 11 Then
    '                    strSO = arrayH(11).ToString()
    '                ElseIf iH = 12 Then
    '                    strCash = arrayH(12).ToString()
    '                ElseIf iH = 13 Then
    '                    strTotal = arrayH(13).ToString()
    '                ElseIf iH = 14 Then
    '                    strtotalCrateQty = clsCommon.myCdbl(arrayH(14).ToString())
    '                ElseIf iH = 15 Then
    '                    strTotalPendingPcsQty = clsCommon.myCdbl(arrayH(15).ToString())
    '                ElseIf iH = 16 Then
    '                    strTotalAmount = 0 'clsCommon.myCdbl(arrayH(16).ToString())
    '                ElseIf iH = 17 Then
    '                    DataSNO = arrayH(17).ToString()
    '                ElseIf iH = 18 Then
    '                    strCurUnitCode = arrayH(18).ToString()
    '                ElseIf iH = 19 Then
    '                    strCurItemRowNo = arrayH(19).ToString()
    '                ElseIf iH = 20 Then
    '                    strLR_PQ_ES = arrayH(20).ToString()
    '                ElseIf iH = 21 Then
    '                    strOpencrateQty = arrayH(21).ToString()
    '                ElseIf iH = 22 Then
    '                    strCrateQtyRecd = arrayH(22).ToString()
    '                ElseIf iH = 23 Then
    '                    strCrateOutQty = arrayH(23).ToString()
    '                ElseIf iH = 24 Then
    '                    strCrateAdjQty = arrayH(24).ToString()
    '                ElseIf iH = 25 Then
    '                    strCrateQtyClosing = arrayH(25).ToString()
    '                ElseIf iH = 26 Then
    '                    strSchemeStar = arrayH(26).ToString()
    '                ElseIf iH = 27 Then
    '                    strTotal_Qty_In_KG = arrayH(27).ToString()
    '                ElseIf iH = 28 Then
    '                    strTotal_Qty_In_LTR = arrayH(28).ToString()
    '                ElseIf iH = 29 Then
    '                    strTCSTotalAmount = clsCommon.myCdbl(arrayH(29).ToString())
    '                End If
    '            Next
    '            '=============================================================================================================================
    '            If clsCommon.CompairString(strCurCustomerNo, strPrvCustomerNo) <> CompairStringResult.Equal Then
    '                If NextTotalNoOnPage = 0 And clsCommon.CompairString(DataSNO, "1") = CompairStringResult.Equal Then
    '                    NextTotalNoOnPage = NextTotalNoOnPage + 3 + 2 + 3 + clsCommon.myCdbl(strCurItemRowNo) + 1 + 1 ' add 1 for Line 
    '                    PrvTotalNoOnPage = 3 + 2 + 3 + clsCommon.myCdbl(strCurItemRowNo) + 1 + 1 ' add 1 for Line 
    '                ElseIf clsCommon.CompairString(DataSNO, "1") = CompairStringResult.Equal Then
    '                    NextTotalNoOnPage = NextTotalNoOnPage + 0 + 1 + clsCommon.myCdbl(strCurItemRowNo) + 1 ' 3 replace 2
    '                    PrvTotalNoOnPage = 0 + 1 + clsCommon.myCdbl(strCurItemRowNo) + 1  ' 3 replace 2
    '                ElseIf clsCommon.CompairString(DataSNO, "2") = CompairStringResult.Equal Then
    '                    NextTotalNoOnPage = NextTotalNoOnPage + 1
    '                    PrvTotalNoOnPage = 1
    '                ElseIf clsCommon.CompairString(DataSNO, "3") = CompairStringResult.Equal Then
    '                    NextTotalNoOnPage = NextTotalNoOnPage + 0 + 1 + clsCommon.myCdbl(strCurItemRowNo) + 1 + 6 + 2 ' 6 for additional 5 line below the route Total plus one dot line, 2 for blank line and (Q.P.S        LOADED BY          SECURITY               RMRD)
    '                    PrvTotalNoOnPage = 0 + 1 + clsCommon.myCdbl(strCurItemRowNo) + 1 + 6 + 2 ' 6 for additional 5 line below the route Total plus one dot line , 2 for blank line and (Q.P.S        LOADED BY          SECURITY               RMRD)
    '                End If
    '            End If
    '            If clsCommon.CompairString(strCurCustomerNo, strPrvCustomerNo) <> CompairStringResult.Equal AndAlso clsCommon.CompairString(DataSNO, "2") = CompairStringResult.Equal Then
    '                NextTotalNoOnPage = NextTotalNoOnPage + clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " /4) + case when  (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " %4) > 0  then 1 else 0 end"))
    '                PrvTotalNoOnPage = PrvTotalNoOnPage + clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " /4) + case when  (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " %4) > 0  then 1 else 0 end"))
    '            End If
    '            If clsCommon.CompairString(strCurCustomerDesc, "Customer Leakage") = CompairStringResult.Equal AndAlso clsCommon.CompairString(strCurCustomerNo, strPrvCustomerNo) = CompairStringResult.Equal AndAlso clsCommon.CompairString(strCurCustomerDesc, strPrvCustomerDesc) <> CompairStringResult.Equal AndAlso clsCommon.CompairString(DataSNO, "22") = CompairStringResult.Equal Then
    '                NextTotalNoOnPage = NextTotalNoOnPage + clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " /4) + case when  (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " %4) > 0  then 1 else 0 end")) '1
    '                PrvTotalNoOnPage = PrvTotalNoOnPage + clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " /4) + case when  (" + clsCommon.myCstr(clsCommon.myCdbl(strCurItemRowNo)) + " %4) > 0  then 1 else 0 end")) '1
    '            End If
    '            If NextTotalNoOnPage > OnePageTotalLine OrElse (clsCommon.CompairString(strCurRouteNo, strPrvRouteNo) <> CompairStringResult.Equal AndAlso clsCommon.myLen(strPrvRouteNo) > 0) Then
    '                Dim BlankSpace As Integer = PrvTotalNoOnPage - (NextTotalNoOnPage - OnePageTotalLine)
    '                TotalLineNoPrint = 0
    '                NextTotalNoOnPage = PrvTotalNoOnPage + 7  ' for 3 heading Total
    '                ' Frist total
    '                If (SumOfAmount > 0 OrElse SumOfCrate > 0) OrElse SumOfTotlQty > 0 Then
    '                    qry = "select IsTCSnotApplicable,PAN from TSPL_CUSTOMER_MASTER where Cust_Code='" & strPrvCustomerNo & "'"
    '                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '                    qry = ""
    '                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '                        If clsCommon.myCdbl(dt.Rows(0)("IsTCSnotApplicable")) = 0 Then
    '                            Dim TCSAmt As Decimal = 0
    '                            If clsCommon.myLen(dt.Rows(0)("PAN")) > 0 Then
    '                                TCSAmt = Math.Round(SumOfTCSAmount * settTCSRateforCustomerWithPanNo / 100, 0)
    '                                qry = "( Including TCS Amount Rs: " & clsCommon.myCstr(TCSAmt) & " )"
    '                            Else
    '                                TCSAmt = Math.Round(SumOfTCSAmount * settTCSRateforCustomerWithoutPanNo / 100, 0)
    '                                qry = "( Including TCS Amount Rs: " & clsCommon.myCstr(TCSAmt) & " )"
    '                            End If
    '                            SumOfAmount += TCSAmt
    '                            dblTotalTCS += TCSAmt
    '                        End If
    '                    End If
    '                    If checkTotalRoutWise Then
    '                        If ApplyIncludeTCSAmountInRouteTotalOnTruckSheet Then
    '                            SumOfAmount += dblTotalTCS
    '                            strTotalAmount += dblTotalTCS
    '                            qry = "( Including TCS Amount Rs: " & clsCommon.myCstr(dblTotalTCS) & " )"
    '                        Else
    '                            qry = "( Route TCS Amount Rs: " & clsCommon.myCstr(dblTotalTCS) & " )"
    '                        End If
    '                        'sw.Write(GetFormateColumnValue(qry, 35) + GetFormateColumnValue(SumOfCrate, 30, "R") + "-0" + GetFormateColumnValue(SumOfAmount.ToString("N2").Replace(",", String.Empty), 12, "R") + "  ")
    '                        dblTotalTCS = 0
    '                    Else
    '                        'sw.Write(GetFormateColumnValue(qry, 35) + GetFormateColumnValue(SumOfCrate, 30, "R") + "-0" + GetFormateColumnValue(Math.Round(SumOfAmount, 0), 12, "R") + "  ")
    '                    End If
    '                    'sw.WriteLine()
    '                    SumOfCrate = 0
    '                    SumOfAmount = 0.0
    '                    SumOfTCSAmount = 0.0
    '                    SumOfTotlQty = 0
    '                    If clsCommon.myLen(strCustomerScheme) > 0 Then
    '                        sw.Write(strCustomerScheme)
    '                        sw.WriteLine()
    '                        strCustomerScheme = ""
    '                        rowsCustomerSchemeCount = 0
    '                    End If
    '                    If clsCommon.myLen(strCustomerLR_PQ_ES) > 0 Then
    '                        sw.Write(strCustomerLR_PQ_ES)
    '                        sw.WriteLine()
    '                        strCustomerLR_PQ_ES = ""
    '                        rowsCustLR_PQ_ES_Count = 0
    '                    End If
    '                End If
    '                If clsCommon.myLen(strPrvRouteNo) > 0 Then
    '                    Dim iii As Integer = 0
    '                    For iii = 0 To BlankSpace - 1
    '                    Next
    '                End If
    '            End If
    '            If TotalLineNoPrint = 0 Then ' TotalLineNoPrint = 0
    '                'Dim iHeader As Integer = 0
    '                If PageNo <> 1 AndAlso clsCommon.CompairString(strCurCustomerNo, "Route Total") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(strPrvCustomerNo, "Route Total") <> CompairStringResult.Equal Then
    '                    NextTotalNoOnPage = NextTotalNoOnPage + 7 + 1 ' add 1 for Line  ' 7 for Logo space
    '                End If
    '                sw.Write("                                                              DES/FRT-05/REV-00")
    '                sw.WriteLine()
    '                sw.Write("                                                                01-04-11/Pg 1/1")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                If PageNo = 1 Then
    '                    sw.Write(".           THE TELANGANA STATE DAIRY DEV. CO-OP. FEDN. LTD -MPF: HYDERABAD")
    '                Else
    '                    sw.Write(".           THE TELANGANA STATE DAIRY DEV. CO-OP. FEDN. LTD -MPF: HYDERABAD")
    '                End If
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                'sw.Write("Page: " + clsCommon.myCstr(PageNo) + " ")
    '                'sw.WriteLine()
    '                'TotalLineNoPrint = TotalLineNoPrint + 1
    '                PageNo = PageNo + 1
    '                If isGatePass Then
    '                    sw.Write("ROUTE : " + clsCommon.myCstr(strCurRouteDesc) + "     DISTRIBUTOR WISE " + IIf(cmbGatePassType.Text = "AM", "MORNING", "EVENING") + " LMS SUPPLY TRUCK SHEET OF " + clsCommon.myCstr(strCurDocDate) + "")
    '                Else
    '                    sw.Write("ROUTE : " + clsCommon.myCstr(strCurRouteDesc) + "                MORNING TRUCK SHEET OF " + clsCommon.myCstr(strCurDocDate) + " Time :")
    '                End If
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                sw.Write("Time: " + clsCommon.myCstr(strDocTime) + "  Supply No:     Vehicle No:" + strVehicleCode + "  Invoice No:" + txtDocNo.Value + " ")
    '                'sw.Write("ZONE : " + clsCommon.myCstr(strCurZoneDesc) + "                   DRIVER:            CASHIER: ")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                sw.Write("----------------------------------------------------------------------------------")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                For i = 0 To submittedDataTable.Columns.Count - 1
    '                    Dim strColumn_Name As String = ""
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "Customer_Name") = CompairStringResult.Equal Then
    '                        strColumn_Name = "Booth Name          " '20 8+7 +4
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "Alies_Name") = CompairStringResult.Equal Then
    '                        strColumn_Name = "Type    " '8
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "CR") = CompairStringResult.Equal Then
    '                        strColumn_Name = "    CR" '7
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "CD") = CompairStringResult.Equal Then
    '                        strColumn_Name = "    CD" '7
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "SO") = CompairStringResult.Equal Then
    '                        strColumn_Name = "   S.O" '7
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "CASH") = CompairStringResult.Equal Then
    '                        strColumn_Name = "  Cash" '7
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "Total") = CompairStringResult.Equal Then
    '                        strColumn_Name = "   Total" '8
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "CrateQty_New") = CompairStringResult.Equal Then
    '                        strColumn_Name = "  Crates" '8
    '                    End If
    '                    If clsCommon.CompairString(submittedDataTable.Columns(i).ColumnName.Trim(), "Amount_with_Tax") = CompairStringResult.Equal Then
    '                        strColumn_Name = "     Amount" '11
    '                    End If
    '                    If clsCommon.myLen(strColumn_Name) > 0 Then
    '                        sw.Write(strColumn_Name)
    '                    End If
    '                Next
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                sw.Write("----------------------------------------------------------------------------------")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '            End If
    '            ' Second total
    '            If clsCommon.CompairString(strCurCustomerNo, strPrvCustomerNo) <> CompairStringResult.Equal Then  'clsCommon.CompairString(strCurRouteNo, strPrvRouteNo) <> CompairStringResult.Equal AndAlso
    '                If SumOfAmount > 0 OrElse SumOfTotlQty > 0 Then
    '                    qry = "select IsTCSnotApplicable,PAN from TSPL_CUSTOMER_MASTER where Cust_Code='" & strPrvCustomerNo & "'"
    '                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '                    qry = ""
    '                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '                        If clsCommon.myCdbl(dt.Rows(0)("IsTCSnotApplicable")) = 0 Then
    '                            Dim TCSAmt As Decimal = 0
    '                            If clsCommon.myLen(dt.Rows(0)("PAN")) > 0 Then
    '                                TCSAmt = Math.Round(SumOfTCSAmount * settTCSRateforCustomerWithPanNo / 100, 0)
    '                                qry = "( Including TCS Amount Rs: " + clsCommon.myCstr(TCSAmt) + " )"
    '                            Else
    '                                TCSAmt = Math.Round(SumOfTCSAmount * settTCSRateforCustomerWithoutPanNo / 100, 0)
    '                                qry = "( Including TCS Amount Rs: " + clsCommon.myCstr(TCSAmt) + " )"
    '                            End If
    '                            SumOfAmount += TCSAmt
    '                            dblTotalTCS += TCSAmt
    '                        End If
    '                    End If
    '                    If checkTotalRoutWise Then
    '                        If ApplyIncludeTCSAmountInRouteTotalOnTruckSheet Then
    '                            SumOfAmount += dblTotalTCS
    '                            strTotalAmount += dblTotalTCS
    '                            qry = "( Including TCS Amount Rs: " & clsCommon.myCstr(dblTotalTCS) & " )"
    '                        Else
    '                            qry = "( Route TCS Amount Rs: " & clsCommon.myCstr(dblTotalTCS) & " )"
    '                        End If
    '                        'sw.Write(GetFormateColumnValue(qry, 35) + GetFormateColumnValue(SumOfCrate, 30, "R") + "-0" + GetFormateColumnValue(SumOfAmount.ToString("N2").Replace(",", String.Empty), 12, "R") + "  ")
    '                        dblTotalTCS = 0
    '                    Else
    '                        'sw.Write(GetFormateColumnValue(qry, 35) + GetFormateColumnValue(SumOfCrate, 30, "R") + "-0" + GetFormateColumnValue(Math.Round(SumOfAmount, 0), 12, "R") + "  ")
    '                    End If
    '                    'sw.WriteLine()
    '                    SumOfCrate = 0
    '                    SumOfAmount = 0.0
    '                    SumOfTCSAmount = 0
    '                    SumOfTotlQty = 0
    '                    If clsCommon.myLen(strCustomerScheme) > 0 Then
    '                        sw.Write(strCustomerScheme)
    '                        sw.WriteLine()
    '                        strCustomerScheme = ""
    '                        rowsCustomerSchemeCount = 0
    '                    End If
    '                    If clsCommon.myLen(strCustomerLR_PQ_ES) > 0 Then
    '                        sw.Write(strCustomerLR_PQ_ES)
    '                        sw.WriteLine()
    '                        strCustomerLR_PQ_ES = ""
    '                        rowsCustLR_PQ_ES_Count = 0
    '                    End If
    '                    If Not checkTotalRoutWise Then
    '                        sw.Write("----------------------------------------------------------------------------------")
    '                        sw.WriteLine()
    '                        TotalLineNoPrint = TotalLineNoPrint + 1
    '                    End If
    '                End If
    '            End If
    '            If clsCommon.CompairString(DataSNO, "2") = CompairStringResult.Equal Then
    '                If clsCommon.myLen(strCustomerScheme) <= 0 Then
    '                    strCustomerScheme = "Scheme "
    '                End If
    '                If rowsCustomerSchemeCount = 4 OrElse rowsCustomerSchemeCount = 8 OrElse rowsCustomerSchemeCount = 12 OrElse rowsCustomerSchemeCount = 16 OrElse rowsCustomerSchemeCount = 20 Then
    '                    rowsCustomerSchemeCount += Environment.NewLine
    '                End If
    '                If rowsCustomerSchemeCount = 0 Then
    '                    strCustomerScheme += "" & strItemDesc & " " & strCurUnitCode & " : " & strTotal & "" & " "
    '                Else
    '                    strCustomerScheme += "," & strItemDesc & " " & strCurUnitCode & " : " & strTotal & "" & " "
    '                End If
    '                rowsCustomerSchemeCount += 1
    '            ElseIf clsCommon.CompairString(DataSNO, "22") = CompairStringResult.Equal Then
    '                If clsCommon.myLen(strCustomerLR_PQ_ES) <= 0 Then
    '                    strCustomerLR_PQ_ES = "LR/PQ/ES "
    '                End If
    '                If rowsCustLR_PQ_ES_Count = 4 OrElse rowsCustLR_PQ_ES_Count = 8 OrElse rowsCustLR_PQ_ES_Count = 12 OrElse rowsCustLR_PQ_ES_Count = 16 OrElse rowsCustLR_PQ_ES_Count = 20 Then
    '                    strCustomerLR_PQ_ES += Environment.NewLine
    '                End If
    '                strCustomerLR_PQ_ES += "" & strItemDesc & "  : " & strLR_PQ_ES & "" & " "
    '                rowsCustLR_PQ_ES_Count += 1
    '            ElseIf clsCommon.CompairString(DataSNO, "4") = CompairStringResult.Equal Then
    '                If clsCommon.myLen(strTotalRouteLR_PQ_ES) <= 0 Then
    '                    strTotalRouteLR_PQ_ES = "Leakage Total "
    '                End If
    '                If rowsCount = 4 OrElse rowsCount = 8 OrElse rowsCount = 12 OrElse rowsCount = 16 OrElse rowsCount = 20 Then
    '                    strTotalRouteLR_PQ_ES += Environment.NewLine
    '                End If
    '                strTotalRouteLR_PQ_ES += "" & strItemDesc & "  : " & strLR_PQ_ES & "" & " "
    '                rowsCount += 1
    '                If strCurItemRowNo = rowsCount Then
    '                    sw.Write(strTotalRouteLR_PQ_ES)
    '                    sw.WriteLine()
    '                    strTotalRouteLR_PQ_ES = ""
    '                    rowsCount = 0
    '                End If
    '            ElseIf clsCommon.CompairString(DataSNO, "5") = CompairStringResult.Equal Then
    '                strTotalRouteSummary = ""
    '                sw.Write("----------------------------------------------------------------------------------")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                Dim dblTotalLTRAndKg As Double = clsCommon.myCdbl(strTotal_Qty_In_LTR) + clsCommon.myCdbl(strTotal_Qty_In_KG)
    '                'strTotalRouteSummary += "TOT LTRS: " + "" + strTotal_Qty_In_LTR.ToString("N2") + "   TOT KG: " + strTotal_Qty_In_KG.ToString("N2") + "   TOT CRATES:" + clsCommon.myCstr(strtotalCrateQty) + "-0   " + "" + "   TOT CASH: " + clsCommon.myCstr(strTotalAmount)
    '                strTotalRouteSummary += "TOT CRATES:" + clsCommon.myCstr(strtotalCrateQty) + "-0   " + "" + "            TOT LTRS: " + "" + dblTotalLTRAndKg.ToString("N2") + ""
    '                sw.Write(strTotalRouteSummary)
    '                sw.WriteLine()
    '                sw.Write("----------------------------------------------------------------------------------")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                '===================================
    '                'sw.Write("|LoadingTmgs(Min) |Veh Bottom Cond.| Top| Logo| Self| Light| No.of Loaders|Other|")
    '                'sw.WriteLine()
    '                'sw.Write("| Strt| End| TmTkn|")
    '                'sw.WriteLine()
    '                'sw.Write("----------------------------------------------------------------------------------")
    '                'sw.WriteLine()
    '                'sw.Write("|      |    |      |                |    |        |   |     |               |      |")
    '                'sw.WriteLine()
    '                'sw.Write("----------------------------------------------------------------------------------")
    '                'sw.WriteLine()
    '                'TotalLineNoPrint = TotalLineNoPrint + 5
    '                sw.WriteLine()
    '                sw.Write("LOADED BY        SECURITY       Shift-In Charge         RMD")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 2
    '                '==================================
    '            ElseIf clsCommon.CompairString(DataSNO, "6") = CompairStringResult.Equal Then
    '                strXCratesRouteDetails = ""
    '                Dim dbloutStdAM As Double = clsCommon.myCdbl(strCrateQtyRecd) - clsCommon.myCdbl(strCrateOutQty)
    '                strXCratesRouteDetails += "Date= : " + "" + strCurDocDate + "   Crts-sent=:" + clsCommon.myCstr(strCrateOutQty) + " " + "   Recvd: " + clsCommon.myCstr(strCrateQtyRecd) + "    outStdAM= " + clsCommon.myCstr(dbloutStdAM) + "     Bal=0"
    '                sw.Write(strXCratesRouteDetails)
    '                sw.WriteLine()
    '                sw.Write("")
    '                sw.WriteLine()
    '                sw.Write("Q.P.S        LOADED BY          SECURITY               RMRD")
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 3
    '            Else
    '                If (clsCommon.CompairString(strCurCustomerNo, strPrvCustomerNo) = CompairStringResult.Equal OrElse clsCommon.myLen(strPrvCustomerNo) <= 0) Then
    '                    If clsCommon.myLen(strPrvCustomerNo) <= 0 Then
    '                    ElseIf clsCommon.myLen(strPrvCustomerDesc) <= 0 Then
    '                        strCurCustomerDesc = ""
    '                    ElseIf clsCommon.CompairString(strPrvCustomerDesc, strCurCustomerNo + " (" + strCurDocDate + ")") <> CompairStringResult.Equal Then
    '                        strCurCustomerDesc = strCurCustomerNo + " (" + strCurDocDate + ")"
    '                    Else
    '                        strCurCustomerDesc = ""
    '                    End If
    '                End If
    '                If clsCommon.CompairString(DataSNO, "3") = CompairStringResult.Equal Then
    '                    If (clsCommon.CompairString(strCurRouteNo, strPrvRouteNo) = CompairStringResult.Equal) Then
    '                        countRecordRouteTotal = countRecordRouteTotal + 1
    '                        If (clsCommon.CompairString(strCurCustomerNo, strPrvCustomerNo) <> CompairStringResult.Equal AndAlso clsCommon.CompairString(strCurCustomerNo, "Route Total") = CompairStringResult.Equal) Then
    '                            If strCurItemRowNo = 1 Then
    '                                strCurCustomerDesc = strTransportercode + "(" + strDocDate + ")"
    '                            Else
    '                                strCurCustomerDesc = strTransporterName '"ROUTE TOTAL"
    '                            End If
    '                            'ElseIf clsCommon.CompairString(strPrvCustomerDescForRoutetotal, "ROUTE TOTAL") <> CompairStringResult.Equal Then
    '                            '    strCurCustomerDesc = "ROUTE TOTAL"
    '                        Else
    '                            If countRecordRouteTotal = 2 Then
    '                                strCurCustomerDesc = strTransportercode + "(" + strDocDate + ")"
    '                            Else
    '                                strCurCustomerDesc = ""
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '                sw.Write(GetFormateColumnValue(strCurCustomerDesc, 20))
    '                sw.Write(GetFormateColumnValue(strItemDesc, 8))
    '                sw.Write(GetFormateColumnValue(strCR.Replace(",", String.Empty), 6, "R"))
    '                sw.Write(GetFormateColumnValue(strCD.Replace(",", String.Empty), 6, "R"))
    '                sw.Write(GetFormateColumnValue(strSO.Replace(",", String.Empty), 6, "R"))
    '                sw.Write(GetFormateColumnValue(strCash.Replace(",", String.Empty), 6, "R"))
    '                sw.Write(GetFormateColumnValue(strTotal.Replace(",", String.Empty), 8, "R"))
    '                'strSchemeStar
    '                sw.Write(GetFormateColumnValue(strSchemeStar, 1, "R"))
    '                'sw.Write(GetFormateColumnValue(clsCommon.myCstr(strtotalCrateQty) + "-" + GetFormateColumnValue(clsCommon.myCstr(strTotalPendingPcsQty), 2, "R"), 8, "R"))
    '                'sw.Write(GetFormateColumnValue(strTotalAmount.ToString("N2"), 11, "R"))
    '                If clsCommon.CompairString(DataSNO, "3") = CompairStringResult.Equal Then
    '                    sw.Write(GetFormateColumnValue(clsCommon.myCstr(strtotalCrateQty) + "-" + GetFormateColumnValue(clsCommon.myCstr(strTotalPendingPcsQty), 2, "R"), 7, "R"))
    '                Else
    '                    'sw.Write(GetFormateColumnValue(clsCommon.myCstr(strtotalCrateQty) + "." + GetFormateColumnValue(strTotalPendingPcsQty.ToString("00"), 2, "R"), 7, "R"))
    '                    sw.Write(GetFormateColumnValue(clsCommon.myCstr(strtotalCrateQty) + "." + GetFormateColumnValue(clsCommon.myCstr(strTotalPendingPcsQty), 2, "R"), 7, "R"))
    '                End If
    '                sw.Write(GetFormateColumnValue(strTotalAmount.ToString("N2").Replace(",", String.Empty), 11, "R"))
    '                sw.WriteLine()
    '                TotalLineNoPrint = TotalLineNoPrint + 1
    '                SumOfCrate += strtotalCrateQty
    '                SumOfAmount += strTotalAmount
    '                SumOfTCSAmount += strTCSTotalAmount
    '                SumOfTotlQty += clsCommon.myCdbl(strTotal)
    '                If clsCommon.CompairString(DataSNO, "3") = CompairStringResult.Equal Then
    '                    checkTotalRoutWise = True
    '                Else
    '                    checkTotalRoutWise = False
    '                End If
    '            End If
    '            strPrvRouteNo = strCurRouteNo
    '            strPrvZoneCode = strCurZoneCode
    '            strPrvCustomerNo = strCurCustomerNo
    '            strPrvCustomerDesc = strCurCustomerDesc
    '            If clsCommon.CompairString(DataSNO, "3") = CompairStringResult.Equal Then
    '                strPrvCustomerDescForRoutetotal = strCurCustomerDesc
    '            End If
    '        Next
    '        ' Third total
    '        If SumOfAmount > 0 OrElse SumOfTotlQty > 0 Then
    '            If Not checkTotalRoutWise Then
    '                sw.Write(GetFormateColumnValue(SumOfCrate, 65, "R") + "-0" + GetFormateColumnValue(Math.Round(SumOfAmount, 0), 12, "R") + "  ")
    '            Else
    '                sw.Write(GetFormateColumnValue(SumOfCrate, 65, "R") + "-0" + GetFormateColumnValue(SumOfAmount.ToString("N2").Replace(",", String.Empty), 12, "R") + "  ")
    '            End If
    '            sw.WriteLine()
    '            SumOfCrate = 0
    '            SumOfAmount = 0.0
    '            SumOfTotlQty = 0.0
    '            If clsCommon.myLen(strCustomerScheme) > 0 Then
    '                sw.Write(strCustomerScheme)
    '                sw.WriteLine()
    '                strCustomerScheme = ""
    '                rowsCustomerSchemeCount = 0
    '            End If
    '            If clsCommon.myLen(strCustomerLR_PQ_ES) > 0 Then
    '                sw.Write(strCustomerLR_PQ_ES)
    '                sw.WriteLine()
    '                strCustomerLR_PQ_ES = ""
    '                rowsCustLR_PQ_ES_Count = 0
    '            End If
    '        End If
    '        '=========================================
    '        sw.Close()
    '    Catch ex As Exception
    '        common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Public Shared Function GetFormateColumnValue(ByVal strValue As String, ByVal colFixLength As Integer, Optional strSide As String = "L") As String
        Dim collenth As Integer = clsCommon.myLen(strValue)
        Dim typeLength As Integer = colFixLength
        Dim collString = strValue
        If collenth < typeLength Then
            Dim balanceSpace As Integer = typeLength - collenth
            For ii = 0 To balanceSpace - 1
                If clsCommon.CompairString(strSide, "L") = CompairStringResult.Equal Then
                    collString = collString + " "
                Else
                    collString = " " + collString
                End If
            Next
        End If
        Return collString
    End Function
    Private Sub txtVendorNo__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVendorNo._MYOpenMasterForm
        Try
            clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.CustomerMaster, "1")
            ' txtVendorNo.Value = clsDBFuncationality.getSingleValue("select Top 1 Cust_Code from TSPL_CUSTOMER_MASTER order by Created_Date desc")
            'lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.ToString)
        End Try
    End Sub
    'Public Sub LoadData(ByVal VendorCode As String, ByVal DocDate As DateTime)
    '    Dim dblTotalDocAmt As Decimal = 0
    '    Dim obj As clsBookingEntryDairySale = Nothing
    '    Try
    '        If ShowDemandDoc Then
    '            Dim qry As String = "select distinct TSPL_BOOKING_MATSER.Against_DemandBooking_No,TSPL_BOOKING_MATSER.Ship_To_Location,TSPL_BOOKING_MATSER.Created_Date,TSPL_BOOKING_MATSER.AdvanceAmount,TSPL_BOOKING_MATSER.Against_Receipt_No,TSPL_BOOKING_MATSER.Against_Booking_No,TSPL_BOOKING_MATSER.Payment_Mode,TSPL_BOOKING_MATSER.Reference_No,TSPL_BOOKING_MATSER.Counter_No,TSPL_BOOKING_MATSER.IsSampling,TSPL_BOOKING_MATSER.AgainstGatePass,TSPL_BOOKING_MATSER.Document_No,TSPL_BOOKING_MATSER.Document_Date,TSPL_BOOKING_MATSER.Posted,CreateDO_Automatic,TSPL_BOOKING_MATSER.location_code,Cust_Group_Code,Is_Taxable,TRANSACTION_TYPE,Ex_Factory_Date,isnull(CustPO_No,'') as CustPO_No,custpo_date,isnull(SalesmanCode,'') as SalesmanCode,Total_Can,total_Box,Total_Crate,isnull(Is_Cancelled,0) as Is_Cancelled, isnull(Booking_Type,'') as Booking_Type,isnull(Card_SALE_No,'') as Card_SALE_No,CardSale_FROM_DATE,CardSale_TO_DATE,Uploading_date ,isnull(Credit_Limit,0) as Credit_Limit,isnull(Advance_Security,0) as Advance_Security,isnull(Revese_Adv_Security,0) as Revese_Adv_Security,isnull(AR_Credit_Security,0) as AR_Credit_Security,isnull(Pending_Posted_DO,0) as Pending_Posted_DO,isnull(UnPostedDispatch,0) as UnPostedDispatch,isnull(Ledger_Outstansing,0) as Ledger_Outstansing,isnull(Refund_Security,0) as Refund_Security,isnull(Reverse_Refund_Sec,0) as Reverse_Refund_Sec,isnull(Total_Outstanding,0) as Total_Outstanding, isnull(GatePass_Type,'') as GatePass_Type,Created_By,comp_code,Is_DCS,Is_BPL,Is_GHEE,BPL_Coupon_Code,BPL_Name,BPL_Remark,Is_Distributor,BPL_Category,BPL_Coupon_Date,TSPL_BOOKING_DETAIL.Vehicle_Code from TSPL_BOOKING_MATSER left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No where comp_code='" & objCommonVar.CurrentCompanyCode & "' and convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & clsCommon.GetPrintDate(DocDate) & "',103) and TSPL_BOOKING_DETAIL.Cust_Code='" & VendorCode & "' "
    '            Dim isDemandDoc = clsDBFuncationality.getSingleValue("select top 1 TSPL_BOOKING_DETAIL.Against_DemandBooking_No from TSPL_BOOKING_MATSER left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No where convert(date,TSPL_BOOKING_MATSER.Document_Date ,103)=convert(date,'" & clsCommon.GetPrintDate(DocDate) & "',103) and TSPL_BOOKING_DETAIL.Cust_Code='" & VendorCode & "'")
    '            If clsCommon.myLen(isDemandDoc) > 0 Then
    '                qry += "  and TSPL_BOOKING_MATSER.GatePass_Type='" & clsCommon.myCstr(cmbGatePassType.Text) & "'"
    '            End If
    '            If clsCommon.myLen(isDemandDoc) > 0 Then
    '                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Nothing)
    '                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '                    obj = New clsBookingEntryDairySale
    '                    obj.IsSampling = clsCommon.myCdbl(dt.Rows(0)("IsSampling"))
    '                    obj.AgainstGatePass = clsCommon.myCdbl(dt.Rows(0)("AgainstGatePass"))
    '                    obj.Document_No = clsCommon.myCstr(dt.Rows(0)("Document_No"))
    '                    obj.Against_DemandBooking_No = clsCommon.myCstr(dt.Rows(0)("Against_DemandBooking_No"))
    '                    obj.Document_Date = clsCommon.myCDate(dt.Rows(0)("Document_Date"))
    '                    obj.Created_Date = clsCommon.myCDate(dt.Rows(0)("Created_Date"))
    '                    obj.Posted = clsCommon.myCdbl(dt.Rows(0)("Posted"))
    '                    obj.CreateDO_Automatic = clsCommon.myCdbl(dt.Rows(0)("CreateDO_Automatic"))
    '                    obj.location_code = clsCommon.myCstr(dt.Rows(0)("location_code"))
    '                    obj.Ship_To_Location = clsCommon.myCstr(dt.Rows(0)("Ship_To_Location"))
    '                    obj.Cust_Group_Code = clsCommon.myCstr(dt.Rows(0)("Cust_Group_Code"))
    '                    obj.Is_Taxable = clsCommon.myCstr(dt.Rows(0)("Is_Taxable"))
    '                    obj.TRANSACTION_TYPE = clsCommon.myCstr(dt.Rows(0)("TRANSACTION_TYPE"))
    '                    obj.Booking_Type = clsCommon.myCstr(dt.Rows(0)("Booking_Type"))
    '                    If clsCommon.myLen(dt.Rows(0)("Ex_Factory_Date")) > 0 Then
    '                        obj.Ex_Factory_Date = clsCommon.myCDate(dt.Rows(0)("Ex_Factory_Date"))
    '                    End If
    '                    ''richa agarwal 16 Oct,2019
    '                    obj.Card_SALE_No = clsCommon.myCstr(dt.Rows(0)("Card_SALE_No"))
    '                    If clsCommon.myLen(dt.Rows(0)("CardSale_TO_DATE")) > 0 Then
    '                        obj.CardSale_TO_DATE = clsCommon.myCDate(dt.Rows(0)("CardSale_TO_DATE"))
    '                    End If
    '                    If clsCommon.myLen(dt.Rows(0)("CardSale_FROM_DATE")) > 0 Then
    '                        obj.CardSale_FROM_DATE = clsCommon.myCDate(dt.Rows(0)("CardSale_FROM_DATE"))
    '                    End If
    '                    If clsCommon.myLen(dt.Rows(0)("Uploading_date")) > 0 Then
    '                        obj.Uploading_date = clsCommon.myCDate(dt.Rows(0)("Uploading_date"))
    '                    End If
    '                    obj.Reference_No = clsCommon.myCstr(dt.Rows(0)("Reference_No"))
    '                    obj.Counter_No = clsCommon.myCstr(dt.Rows(0)("Counter_No"))
    '                    obj.Payment_Mode = clsCommon.myCstr(dt.Rows(0)("Payment_Mode"))
    '                    obj.Against_Booking_No = clsCommon.myCstr(dt.Rows(0)("Against_Booking_No"))
    '                    obj.Against_Receipt_No = clsCommon.myCstr(dt.Rows(0)("Against_Receipt_No"))
    '                    obj.AdvanceAmount = clsCommon.myCdbl(dt.Rows(0)("AdvanceAmount"))
    '                    obj.Is_DCS = clsCommon.myCdbl(dt.Rows(0)("Is_DCS"))
    '                    obj.SalesmanCode = clsCommon.myCstr(dt.Rows(0)("SalesmanCode"))
    '                    obj.Cust_PO_No = clsCommon.myCstr(dt.Rows(0)("CustPO_No"))
    '                    If clsCommon.myLen(dt.Rows(0)("custpo_date")) > 0 Then
    '                        obj.Podate = clsCommon.myCDate(dt.Rows(0)("custpo_date"))
    '                    End If
    '                    obj.TotalCrate = clsCommon.myCdbl(dt.Rows(0)("Total_Crate"))
    '                    obj.TotalCAN = clsCommon.myCdbl(dt.Rows(0)("Total_CAN"))
    '                    obj.TotalBox = clsCommon.myCdbl(dt.Rows(0)("Total_Box"))
    '                    'Sanjay Ticket No- ERO/12/07/18-000371
    '                    obj.Is_Cancelled = clsCommon.myCdbl(dt.Rows(0)("Is_Cancelled"))
    '                    'Sanjay Ticket No- ERO/12/07/18-000371
    '                    obj.Credit_Limit = clsCommon.myCdbl(dt.Rows(0)("Credit_Limit"))
    '                    obj.Advance_Security = clsCommon.myCdbl(dt.Rows(0)("Advance_Security"))
    '                    obj.Revese_Adv_Security = clsCommon.myCdbl(dt.Rows(0)("Revese_Adv_Security"))
    '                    obj.AR_Credit_Security = clsCommon.myCdbl(dt.Rows(0)("AR_Credit_Security"))
    '                    obj.Pending_Posted_DO = clsCommon.myCdbl(dt.Rows(0)("Pending_Posted_DO"))
    '                    obj.UnPostedDispatch = clsCommon.myCdbl(dt.Rows(0)("UnPostedDispatch"))
    '                    obj.Ledger_Outstansing = clsCommon.myCdbl(dt.Rows(0)("Ledger_Outstansing"))
    '                    obj.Refund_Security = clsCommon.myCdbl(dt.Rows(0)("Refund_Security"))
    '                    obj.Reverse_Refund_Sec = clsCommon.myCdbl(dt.Rows(0)("Reverse_Refund_Sec"))
    '                    obj.Total_Outstanding = clsCommon.myCdbl(dt.Rows(0)("Total_Outstanding"))
    '                    obj.GatePass_Type = clsCommon.myCstr(dt.Rows(0)("GatePass_Type"))
    '                    obj.Created_By = clsCommon.myCstr(dt.Rows(0)("Created_By"))
    '                    obj.Is_BPL = clsCommon.myCdbl(dt.Rows(0)("Is_BPL"))
    '                    obj.Is_GHEE = clsCommon.myCdbl(dt.Rows(0)("Is_GHEE"))
    '                    obj.Is_Distributor = clsCommon.myCdbl(dt.Rows(0)("Is_Distributor"))
    '                    obj.BPL_Coupon_Code = clsCommon.myCstr(dt.Rows(0)("BPL_Coupon_Code"))
    '                    If dt.Rows(0)("BPL_Coupon_Date") IsNot DBNull.Value Then
    '                        obj.BPL_Coupon_Date = clsCommon.myCDate(dt.Rows(0)("BPL_Coupon_Date"))
    '                    Else
    '                        obj.BPL_Coupon_Date = Nothing
    '                    End If
    '                    obj.BPL_Name = clsCommon.myCstr(dt.Rows(0)("BPL_Name"))
    '                    obj.BPL_Category = clsCommon.myCstr(dt.Rows(0)("BPL_Category"))
    '                    obj.BPL_Remark = clsCommon.myCstr(dt.Rows(0)("BPL_Remark"))
    '                    txtVehicleCode.Value = clsCommon.myCstr(dt.Rows(0)("Vehicle_Code"))
    '                    txtVehicleName.Text = clsDBFuncationality.getSingleValue("select Vehicle_Name from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicleCode.Value & "'")
    '                    obj.arrBookingDetailDairySalePaymentMode = clsBookingDetailDairySalePaymentMode.getData(obj.Document_No, Nothing)
    '                End If
    '                If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
    '                    'Dim isDemandBooking1 = clsDBFuncationality.getSingleValue("select top 1 Against_DemandBooking_No from TSPL_BOOKING_DETAIL  where Document_No='" & obj.Document_No & "'")
    '                    If clsCommon.myLen(isDemandDoc) > 0 Then
    '                        rgbItemType.Visible = False
    '                        If clsCommon.CompairString(obj.GatePass_Type, "AM") = CompairStringResult.Equal Then
    '                            lblShiftType.Text = "Morning"
    '                        Else
    '                            lblShiftType.Text = "Evening"
    '                        End If
    '                    Else
    '                        rgbItemType.Visible = True
    '                    End If
    '                    btnSave.Enabled = True
    '                    btnPost.Enabled = True
    '                    btnDelete.Enabled = True
    '                    isInsideLoadData = True
    '                    'btnCC.Enabled = False
    '                    isNewEntry = False
    '                    btnSave.Text = "Update"
    '                    BlankAllControls()
    '                    LoadBlankGrid()
    '                    chkSampling.Checked = IIf(obj.IsSampling = 1, True, False)
    '                    chkGatePass.Checked = IIf(obj.AgainstGatePass = 1, True, False)
    '                    chkDCS.Checked = IIf(obj.Is_DCS = 1, True, False)
    '                    chkBPL.Checked = IIf(obj.Is_BPL = 1, True, False)
    '                    chkGhee.Checked = IIf(obj.Is_GHEE = 1, True, False)
    '                    chkGhee.Enabled = False
    '                    chkDistributor.Checked = IIf(obj.Is_Distributor = 1, True, False)
    '                    'txtLocation.Enabled = False
    '                    txtVendorNo.Enabled = False
    '                    If chkBPL.Checked Then
    '                        txtCouponCode.Text = obj.BPL_Coupon_Code
    '                        txtBPLName.Text = obj.BPL_Name
    '                        txtBPLRemark.Text = obj.BPL_Remark
    '                        txtCouponDate.Value = obj.BPL_Coupon_Date
    '                        txtCategory.Value = obj.BPL_Category
    '                    End If
    '                    txtDocNo.Value = obj.Document_No
    '                    txtDate.Value = obj.Document_Date
    '                    If clsCommon.myLen(obj.Against_Receipt_No) > 0 Then
    '                        txtReceipt.Value = obj.Against_Receipt_No
    '                        lblReceiptAmtDesc.Text = clsDBFuncationality.getSingleValue("select Receipt_Amount from TSPL_RECEIPT_HEADER where Receipt_No='" & obj.Against_Receipt_No & "'")
    '                    End If
    '                    If EnableCustomerPODetailonDairyBooking = 1 Then
    '                        txtSalesman.Value = obj.SalesmanCode
    '                        If obj.Podate IsNot Nothing Then
    '                            txtCustPODate.Value = obj.Podate
    '                            txtCustPODate.Checked = True
    '                        End If
    '                        txtPONo.Text = obj.Cust_PO_No
    '                    End If
    '                    If obj.Is_Taxable = 2 Then
    '                        rbtnTaxable.IsChecked = True
    '                    Else
    '                        rbtnNonTax.IsChecked = True
    '                    End If
    '                    lblCreatedByValue.Text = clsCommon.myCstr(obj.Created_By)
    '                    lblDONumber.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(Delivery_No,'') from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'"))
    '                    'DOStatus = clsDBFuncationality.getSingleValue("select isnull(DO_Posted,0) from TSPL_BOOKING_DETAIL where Document_No ='" & txtDocNo.Value & "' and cust_code='" & txtVendorNo.Value & "'")
    '                    If obj.Posted = 1 Then
    '                        btnSave.Enabled = False
    '                        btnDelete.Enabled = False
    '                        btnPost.Enabled = False
    '                        'btnCreateDO.Enabled = True
    '                        'Dim isDemandBooking = clsDBFuncationality.getSingleValue("select top 1 Against_DemandBooking_No from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "'")
    '                        If clsCommon.myLen(isDemandDoc) > 0 Then
    '                            btnCreateAndPrintInvoice.Enabled = False
    '                        Else
    '                            btnCreateAndPrintInvoice.Enabled = True
    '                        End If
    '                        Dim DOStatus1 = clsDBFuncationality.getSingleValue("select top 1  Document_No from TSPL_BOOKING_DETAIL where DO_Posted <> 4 and Document_No='" & txtDocNo.Value & "'")
    '                        If clsCommon.myLen(DOStatus1) = 0 Then
    '                            btnCreateDO.Enabled = False
    '                        End If
    '                    End If
    '                    'If obj.TRANSACTION_TYPE = "FS" Then
    '                    '    rbtn_Fresh.IsChecked = True
    '                    'ElseIf obj.TRANSACTION_TYPE = "PS" Then
    '                    '    rbtn_Ambient.IsChecked = True
    '                    'End If
    '                    If clsCommon.myLen(obj.Ex_Factory_Date) = 0 Then
    '                        txtEx_Factory_Date.Checked = False
    '                    Else
    '                        txtEx_Factory_Date.Checked = True
    '                        txtEx_Factory_Date.Value = obj.Ex_Factory_Date
    '                    End If
    '                    lblCreatedDateAndTime.Text = obj.Created_Date
    '                    ''        cmbBookingType.Text = IIf(obj.Booking_Type = "", "Select", obj.Booking_Type)
    '                    txtLocation.Value = obj.location_code
    '                    lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
    '                    txtCan.Text = obj.TotalCAN
    '                    txtCrate.Text = obj.TotalCrate
    '                    txtBox.Text = obj.TotalBox
    '                    txtShipToLocation.Value = obj.Ship_To_Location
    '                    lblShipToLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_SHIP_TO_LOCATION.Ship_To_Desc FROM  TSPL_SHIP_TO_LOCATION WHERE Ship_To_Code  ='" & txtShipToLocation.Value & "'"))
    '                    If obj.Uploading_date IsNot Nothing Then
    '                        lblUploadingDate.Text = obj.Uploading_date
    '                    End If
    '                    'Sanjay ERO/12/07/18-000371
    '                    Is_Cancelled = obj.Is_Cancelled
    '                    lblCancelStatus.Text = IIf(obj.Is_Cancelled = 1, "Cancel", "")
    '                    If obj.Is_Cancelled = 1 Then
    '                        btnCancel.Enabled = False
    '                    Else
    '                        btnCancel.Enabled = True
    '                    End If
    '                    GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
    '                    lblCreditLimit.Text = clsCommon.myCdbl(obj.Credit_Limit)
    '                    lblAdvanceSecurity.Text = clsCommon.myCdbl(obj.Advance_Security)
    '                    lblReverseAdvanceSec.Text = clsCommon.myCdbl(obj.Revese_Adv_Security)
    '                    lblARSecurity.Text = clsCommon.myCdbl(obj.AR_Credit_Security)
    '                    lblPendingDO.Text = clsCommon.myCdbl(obj.Pending_Posted_DO)
    '                    lblShortcloseDO.Text = clsCommon.myCdbl(obj.UnPostedDispatch)
    '                    lblLedgerOutstanding.Text = clsCommon.myCdbl(obj.Ledger_Outstansing)
    '                    lblRefund.Text = clsCommon.myCdbl(obj.Refund_Security)
    '                    lblReverseRefund.Text = clsCommon.myCdbl(obj.Reverse_Refund_Sec)
    '                    lblTotalOutstansing.Text = clsCommon.myCdbl(obj.Total_Outstanding)
    '                    lbltotalOutstanding1.Text = clsCommon.myCdbl(obj.Total_Outstanding)
    '                    lblTotalSecurity11.Text = clsCommon.myCdbl(obj.Advance_Security) - clsCommon.myCdbl(obj.Revese_Adv_Security) - clsCommon.myCdbl(obj.Refund_Security) + clsCommon.myCdbl(obj.Reverse_Refund_Sec) - clsCommon.myCdbl(obj.AR_Credit_Security)
    '                    qry = "SELECT TSPL_BOOKING_DETAIL.*,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_ROUTE_MASTER.Route_Desc,tspl_item_master.item_desc,tspl_item_master.Short_Description as item_Short_Description,tspl_item_master.HSN_Code FROM TSPL_BOOKING_DETAIL LEFT JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code left outer join " &
    '                            " TSPL_ROUTE_MASTER on TSPL_BOOKING_DETAIL.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join tspl_item_master on tspl_item_master.Item_code=TSPL_BOOKING_DETAIL.Item_code WHERE Document_No='" & txtDocNo.Value & "' and scheme_item='N'"
    '                    Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
    '                    txtVendorNo.Value = clsCommon.myCstr(dt2.Rows(0)("Cust_Code"))
    '                    lblVendorName.Text = clsCommon.myCstr(dt2.Rows(0)("Customer_Name"))
    '                    If clsCommon.CompairString(obj.Booking_Type, "") = CompairStringResult.Equal Then
    '                        cmbcashcredit.Text = ""
    '                    ElseIf clsCommon.CompairString(obj.Booking_Type, "PS") = CompairStringResult.Equal Then
    '                        cmbcashcredit.Text = "PARLOR SALES"
    '                    ElseIf clsCommon.CompairString(obj.Booking_Type, "FN") = CompairStringResult.Equal Then
    '                        cmbcashcredit.Text = "FORENOON"
    '                    ElseIf clsCommon.CompairString(obj.Booking_Type, "UP") = CompairStringResult.Equal Then
    '                        cmbcashcredit.Text = "UP COUNTRY"
    '                    Else
    '                        cmbcashcredit.Text = obj.Booking_Type
    '                    End If
    '                    '====================Added by preeti Gupta Against Ticket no[BHA/01/08/18-000206]=
    '                    lblBoothStation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(Zone_Code,'') +  case when len(oldname )>0 then ', ' +isnull(oldname,'') else '' end from TSPL_CUSTOMER_MASTER where Cust_Code='" & txtVendorNo.Value & "'"))
    '                    lblroutecode.Text = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
    '                    lblroutename.Text = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
    '                    strRoutecode = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
    '                    strRouteDesc = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
    '                    lblTotRAmt1.Text = clsCommon.myCstr(dt2.Rows(0)("DocumentAmount"))
    '                    BookingStatus = clsCommon.myCstr(dt2.Rows(0)("Booking_Status"))
    '                    DOStatus = clsCommon.myCstr(dt2.Rows(0)("DO_Posted"))
    '                    ''richa agarwal ERO/21/05/19-000609 21 May,2019 add updated vehicle No according to DO
    '                    LblUpdatedVehicleCode.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Lorry_No from TSPL_DELIVERY_NOTE_MASTER_FRESHSALE where Booking_No ='" & txtDocNo.Value & "'"))
    '                    LblUpdatedVehicleDesc.Text = clsCommon.myCstr(ClsVehicleMaster.GetName(LblUpdatedVehicleCode.Text, Nothing))
    '                    setRouteDetail(txtVendorNo.Value, lblroutecode.Text)
    '                    GetUnbilledAmt(obj.Document_Date, txtVendorNo.Value)
    '                    If chkDCS.Checked Then
    '                        GetOutStandingBal(txtVendorNo.Value, txtDate.Value)
    '                    Else
    '                        CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
    '                    End If
    '                    ''richa TEC/01/10/19-001025
    '                    txtRouteNo.Value = clsCommon.myCstr(dt2.Rows(0)("Route_No"))
    '                    lblRouteDesc.Text = clsCommon.myCstr(dt2.Rows(0)("Route_Desc"))
    '                    If clsCommon.CompairString(obj.GatePass_Type, "") = CompairStringResult.Equal Then
    '                        cmbGatePassType.Text = "Select"
    '                    Else
    '                        cmbGatePassType.Text = clsCommon.myCstr(obj.GatePass_Type)
    '                    End If
    '                    ' done by priti BHA/14/06/18-000052
    '                    If clsCommon.myLen(BookingStatus) > 0 Then
    '                        If BookingStatus = 1 Then
    '                            txtBOstatus.Text = "Open"
    '                        ElseIf BookingStatus = 2 Then
    '                            txtBOstatus.Text = "Pending"
    '                        ElseIf BookingStatus = 3 Then
    '                            txtBOstatus.Text = "Approved"
    '                        ElseIf BookingStatus = 4 Then
    '                            txtBOstatus.Text = "Posted"
    '                        ElseIf BookingStatus = 5 Then
    '                            txtBOstatus.Text = "Rejected"
    '                        End If
    '                    End If
    '                    If clsCommon.myLen(BookingStatus) > 0 Then
    '                        If DOStatus = 1 Then
    '                            txtDOStatus.Text = "Open"
    '                        ElseIf DOStatus = 2 Then
    '                            txtDOStatus.Text = "Pending"
    '                        ElseIf DOStatus = 3 Then
    '                            txtDOStatus.Text = "Approved"
    '                        ElseIf DOStatus = 4 Then
    '                            txtDOStatus.Text = "Posted"
    '                        ElseIf DOStatus = 5 Then
    '                            txtDOStatus.Text = "Rejected"
    '                        End If
    '                    End If
    '                    If obj.AgainstGatePass = 1 AndAlso clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "TSDDCF") = CompairStringResult.Equal Then
    '                        btnGatePassPrint.Visible = True
    '                    Else
    '                        btnGatePassPrint.Visible = False
    '                    End If
    '                    For jj As Integer = 0 To dt2.Rows.Count() - 1
    '                        gv1.Rows.AddNew()
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count 'clsCommon.myCdbl(dt2.Rows(jj)("Line_No"))
    '                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colIsKKF).Value = (clsCommon.CompairString(clsCommon.myCstr(dt2.Rows(jj)("IsKKFTax")), "YES") = CompairStringResult.Equal)
    '                        'gv1.Rows(gv1.Rows.Count - 1).Cells(colIsMNDTax).Value = (clsCommon.CompairString(clsCommon.myCstr(dt2.Rows(jj)("IsMNDTax")), "YES") = CompairStringResult.Equal)
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dt2.Rows(jj)("Item_Code"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dt2.Rows(jj)("item_desc"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dt2.Rows(jj)("item_Short_Description"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dt2.Rows(jj)("HSN_Code"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIType).Value = clsDBFuncationality.getSingleValue("select TypeOfItm from TSPL_ITEM_MASTER where Item_Code='" & gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value & "' ")
    '                        If ShowAvailableQtyOnDairyBooking = True Then
    '                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColAvailableQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")), txtLocation.Value, txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dt2.Rows(jj)("Unit_Code")), 0)
    '                        End If
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("PreviousBookingQty"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("OrgRate"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dt2.Rows(jj)("Unit_Code"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSellingRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Selling_Price"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = Math.Round(clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate")), 2)
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_On_Amount"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_Amount"))
    '                        UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Amount).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Amount"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Code).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Code"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Pers).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Pers"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Type).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Type"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeType).Value = clsCommon.myCstr(dt2.Rows(jj)("Scheme_Type"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_NonTax"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dt2.Rows(jj)("FreshAmbient"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRemarks).Value = clsCommon.myCstr(dt2.Rows(jj)("Remarks"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colItemBasicPrice).Value = clsCommon.myCdbl(dt2.Rows(jj)("Price_with_Tax"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmountWithTax).Value = clsCommon.myCdbl(dt2.Rows(jj)("Amount_with_Tax"))
    '                        dblTotalDocAmt = dblTotalDocAmt + clsCommon.myCdbl(gv1.Rows(gv1.Rows.Count - 1).Cells(colAmountWithTax).Value)
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceId).Value = clsCommon.myCstr(dt2.Rows(jj)("Item_Price_ID"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPriceIDAppDate).Value = clsCommon.myCstr(dt2.Rows(jj)("Price_IdStartDate"))
    '                        gv1.Rows(gv1.Rows.Count - 1).Cells(colPricePlanNo).Value = clsCommon.myCstr(dt2.Rows(jj)("PricePlanNo"))
    '                        ''RICHA 06 JUNE,2020
    '                        If chkSampling.Checked = False And DonotAllowtoChangeUOMinDairyBookingCustomer = True Then
    '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).ReadOnly = True
    '                        Else
    '                            gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).ReadOnly = False
    '                        End If
    '                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
    '                            UpdateCurrentRow1(gv1.CurrentRow.Index)
    '                        Else
    '                            UpdateCurrentRow(gv1.CurrentRow.Index)
    '                        End If
    '                    Next
    '                    lblTotalDocAmt.Text = Math.Round(clsCommon.myCdbl(dblTotalDocAmt), 2)
    '                    'Try
    '                    '    lblTCSAmount.Text = Math.Round(Math.Round(clsCommon.myCdbl(dblTotalDocAmt), 2) * GetTCSRate(txtVendorNo.Value) / 100, 2)
    '                    'Catch ex As Exception
    '                    'End Try
    '                    ''to show all items other than booking in case of customer type other than others 25 Feb,2020
    '                    'If clsCommon.CompairString(txtBOstatus.Text, "Posted") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(txtBOstatus.Text, "Rejected") <> CompairStringResult.Equal Then
    '                    '    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Customer_category,'') from tspl_customer_master where cust_code='" & clsCommon.myCstr(txtVendorNo.Value) & "' ")), "Others") <> CompairStringResult.Equal Then
    '                    Dim DemandDoc = clsDBFuncationality.getSingleValue("select top 1 Against_DemandBooking_No from TSPL_BOOKING_DETAIL where Document_No='" & obj.Document_No & "'")
    '                    If clsCommon.myLen(DemandDoc) > 0 Then
    '                        qry = "  select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,TSPL_ITEM_UOM_DETAIL.UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
    '                  " case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,0 as Booking_Qty,tspl_item_master.Sku_Seq   from tspl_item_master " & Environment.NewLine &
    '                  " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
    '                  " where isnull(tspl_item_master.Chilled_Freezen,0)=1 and isnull(TSPL_ITEM_MASTER.item_type,'')='F'  and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1 " & Environment.NewLine &
    '                  " and tspl_item_master.Item_Code not in ( select tspl_booking_detail.item_code from tspl_booking_detail" & Environment.NewLine &
    '                  " LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
    '                  " where tspl_booking_detail.Document_No='" & clsCommon.myCstr(obj.Document_No) & "' AND tspl_booking_detail.Scheme_Item ='N') order by Sku_Seq" & Environment.NewLine
    '                        dt2 = clsDBFuncationality.GetDataTable(qry)
    '                        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
    '                            For Each dr As DataRow In dt2.Rows
    '                                isCellValueChangedOpen = True
    '                                gv1.Rows.AddNew()
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("UOM_Code"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
    '                                gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
    '                                If clsCommon.myCdbl(clsCommon.myCdbl(dr("Booking_Qty"))) > 0 Then
    '                                    'ItemPrice(clsCommon.myCstr(dr("item_code")), clsCommon.myCstr(dr("Unit_code")), clsCommon.myCdbl(dr("Booking_Qty")), gv1.Rows.Count - 1)
    '                                End If
    '                                isCellValueChangedOpen = False
    '                            Next
    '                        End If
    '                        '  Else
    '                        '      qry = " select tspl_item_master.item_code,tspl_item_master.Item_Desc,tspl_item_master.Short_Description ,TSPL_ITEM_UOM_DETAIL.UOM_Code ,tspl_item_master.IsTaxable," & Environment.NewLine &
    '                        '" case when tspl_item_master.Is_Ambient=1 then 'PS' WHEN tspl_item_master.Is_FreshItem=1 THEN 'FS' ELSE '' END  IsFreshAmbient,tspl_item_master.HSN_Code,0 as Booking_Qty,tspl_item_master.Marketing_Seq   from tspl_item_master " & Environment.NewLine &
    '                        '" left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code =tspl_item_master.Item_Code " & Environment.NewLine &
    '                        '" where isnull(tspl_item_master.Item_Type,'')='F' and isnull(tspl_item_master.TypeOfItm,'')='A' and tspl_item_master.Active=1 and TSPL_ITEM_UOM_DETAIL.Default_UOM=1 " & Environment.NewLine &
    '                        '" and tspl_item_master.Item_Code not in ( select tspl_booking_detail.item_code from tspl_booking_detail" & Environment.NewLine &
    '                        '" LEFT OUTER JOIN tspl_item_master ON tspl_item_master.item_code=tspl_booking_detail.iTEM_CODE " & Environment.NewLine &
    '                        '" where tspl_booking_detail.Document_No='" & clsCommon.myCstr(obj.Document_No) & "' AND tspl_booking_detail.Scheme_Item ='N')" & Environment.NewLine &
    '                        '" order by Marketing_Seq"
    '                        '      dt2 = clsDBFuncationality.GetDataTable(qry)
    '                        '      If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
    '                        '          For Each dr As DataRow In dt2.Rows
    '                        '              isInsideLoadData = True
    '                        '              isCellValueChangedOpen = True
    '                        '              gv1.Rows.AddNew()
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dr("item_code"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(dr("Item_Desc"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colIShortName).Value = clsCommon.myCstr(dr("Short_Description"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsCommon.myCstr(dr("HSN_Code"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dr("UOM_Code"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dr("IsTaxable"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCstr(dr("IsFreshAmbient"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colPreviousQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
    '                        '              gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dr("Booking_Qty"))
    '                        '              If clsCommon.myCdbl(clsCommon.myCdbl(dr("Booking_Qty"))) > 0 Then
    '                        '                  ItemPrice(clsCommon.myCstr(dr("item_code")), clsCommon.myCstr(dr("Unit_code")), clsCommon.myCdbl(dr("Booking_Qty")), gv1.Rows.Count - 1)
    '                        '              End If
    '                        '              isCellValueChangedOpen = False
    '                        '          Next
    '                        '      End If
    '                    End If
    '                    '    End If
    '                    'End If
    '                    ItemTypePanel.Enabled = False
    '                End If
    '                If (clsCommon.myCdbl(lblTotRAmt1.Text)) > 0 Then
    '                    If AllowWo_Outstanding = False Then
    '                        If CheckOutstandingOnbooking = 1 AndAlso ((BookingStatus = 1 OrElse BookingStatus = 2) AndAlso (BookingStatus <> 5)) Then
    '                            CustomerOutstandingAmount(txtVendorNo.Value, Nothing)
    '                        End If
    '                    End If
    '                End If
    '            End If
    '            'CustomerOutstandingAmount(VendorCode, Nothing)
    '            gv1.Rows.AddNew()
    '            If gv1.Rows.Count > 0 Then
    '                'gv1.Focus()
    '                gv1.CurrentRow = gv1.Rows(0)
    '                gv1.CurrentColumn = gv1.Columns(colQty)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Private Sub txtSalesman1__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean)
        Dim qry As String = "select EMP_CODE as Code,Emp_Name as Name from TSPL_EMPLOYEE_MASTER"
        Dim whrcls As String = "Emp_type='Salesman'"
        txtSalesman1.Value = clsCommon.ShowSelectForm("DBC-SNOSaleman", qry, "Code", whrcls, txtSalesman1.Value, "Code", isButtonClicked)
        lblSalesmandesc1.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Emp_Name as Name from TSPL_EMPLOYEE_MASTER where EMP_CODE ='" & txtSalesman1.Value & "' and Emp_type='Salesman'"))
    End Sub
    Private Sub txtReceipt__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtReceipt._MYValidating
        If clsCommon.myLen(txtVendorNo.Value) > 0 Then
            Dim qry As String = "select TSPL_RECEIPT_HEADER.Receipt_No as Code,TSPL_RECEIPT_HEADER.Receipt_Amount as Amount,TSPL_RECEIPT_HEADER.Receipt_Date,TSPL_RECEIPT_HEADER.Receipt_Post_Date,TSPL_RECEIPT_HEADER.Receipt_Type,TSPL_RECEIPT_HEADER.Balance_Amt from TSPL_RECEIPT_HEADER"
            Dim whrcls As String = " TSPL_RECEIPT_HEADER.Cust_Code='" & clsCommon.myCstr(txtVendorNo.Value) & "'"
            whrcls += " and not exists (select Against_Receipt_No from TSPL_BOOKING_MATSER where TSPL_RECEIPT_HEADER.Receipt_No =TSPL_BOOKING_MATSER.Against_Receipt_No ) and TSPL_RECEIPT_HEADER.Posted='Y'"
            txtReceipt.Value = clsCommon.ShowSelectForm("DBC-Receipt", qry, "Code", whrcls, txtReceipt.Value, "Code", isButtonClicked)
            lblReceiptAmtDesc.Text = clsDBFuncationality.getSingleValue("select Receipt_Amount from TSPL_RECEIPT_HEADER where Receipt_No='" & clsCommon.myCstr(txtReceipt.Value) & "'")
        Else
            clsCommon.MyMessageBoxShow(Me, "Please Select Customer", Me.Text)
        End If
    End Sub
    Private Sub chkDCS_CheckStateChanged(sender As Object, e As EventArgs)
        If chkDCS.Checked Then
            lblCredit.Visible = True
            cmbcashcredit.Visible = True
            chkBPL.Enabled = False
            chkDistributor.Checked = False
            chkDistributor.Enabled = False
            txtLastCollectionDate.Visible = True
            lblLastCollectionDate.Visible = True
            lblDCSDemand.Visible = True
            txtDCSDemandNo.Visible = True
        Else
            lblCredit.Visible = False
            cmbcashcredit.Visible = False
            chkBPL.Enabled = True
            chkDistributor.Checked = True
            chkDistributor.Enabled = True
            txtLastCollectionDate.Visible = False
            lblLastCollectionDate.Visible = False
            lblDCSDemand.Visible = False
            txtDCSDemandNo.Visible = False
        End If
    End Sub
    Public Sub GetUnbilledAmt(ByVal dtDoc As DateTime, ByVal VendorNo As String)
        Dim qry As String = "select TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per, TSPL_VLC_MASTER_HEAD.MCC,TSPL_PAYMENT_CYCLE_MASTER.PC_CODE,TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE" + Environment.NewLine +
        "from TSPL_VENDOR_MASTER" + Environment.NewLine +
        "left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_VENDOR_MASTER.Vendor_Code" + Environment.NewLine +
        "left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_VENDOR_MASTER.Vendor_Code" + Environment.NewLine +
        "left outer join TSPL_MCC_MASTER on  TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC" + Environment.NewLine +
        "left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle" + Environment.NewLine +
        "where TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" & clsCommon.myCstr(VendorNo) & "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, Nothing)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            If clsCommon.myCdbl(dt.Rows(0)("Credit_Limit_On_Milk_Receipt_Per") >= 0) Then
                'Dim dtDoc As Date = obj.Document_Date
                Dim PaymentCycleType As String = clsCommon.myCstr(dt.Rows(0)("PC_TYPE"))
                Dim PaymentCycleValue As Integer = clsCommon.myCdbl(dt.Rows(0)("PC_VALUE"))
                Dim dtFrom As Date = dtDoc
                Dim dtTo As Date = dtDoc
                If clsCommon.CompairString(PaymentCycleType, "Day") = CompairStringResult.Equal Then
                    Dim ModResut As Integer = dtDoc.Day Mod PaymentCycleValue
                    If ModResut = 0 Then
                        ModResut = PaymentCycleValue
                    End If
                    dtFrom = New Date(dtDoc.Year, dtDoc.Month, (dtDoc.Day - (ModResut) + 1))
                    dtTo = dtFrom.AddDays(PaymentCycleValue - 1)
                    Dim dtNxtPay As DateTime = dtTo.AddDays(Math.Ceiling(PaymentCycleValue / 2.0))
                    If dtFrom.Month <> dtNxtPay.Month Then
                        dtTo = New Date(dtFrom.Year, dtFrom.Month, 1).AddMonths(1).AddDays(-1)
                    End If
                ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
                    If clsCommon.myCdbl(clsCommon.GetPrintDate(dtFrom, "dd")) <> 1 Then
                        dtFrom = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                        dtTo = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                    End If
                    dtTo = DateAdd(DateInterval.Month, PaymentCycleValue, dtFrom)
                ElseIf clsCommon.CompairString(PaymentCycleType, "Year") = CompairStringResult.Equal Then
                    If clsCommon.myCdbl(clsCommon.GetPrintDate(dtFrom, "dd")) <> 1 Then
                        dtFrom = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                        dtTo = "01/" & DatePart(DateInterval.Month, dtDoc) & "/" & DatePart(DateInterval.Year, dtDoc)
                    End If
                    dtTo = DateAdd(DateInterval.Year, PaymentCycleValue, dtFrom)
                ElseIf clsCommon.CompairString(PaymentCycleType, "Week") = CompairStringResult.Equal Then
                    Dim today As Date = dtFrom
                    Dim dayDiff As Integer = today.DayOfWeek - IIf(PaymentCycleValue = 1, DayOfWeek.Sunday, IIf(PaymentCycleValue = 2, DayOfWeek.Monday, IIf(PaymentCycleValue = 3, DayOfWeek.Tuesday, IIf(PaymentCycleValue = 4, DayOfWeek.Wednesday, IIf(PaymentCycleValue = 5, DayOfWeek.Thursday, IIf(PaymentCycleValue = 6, DayOfWeek.Friday, DayOfWeek.Saturday))))))
                    dtFrom = today.AddDays(-dayDiff)
                    dtTo = dtFrom.AddDays(6)
                End If
                qry = "select convert(decimal(18,2), isnull( sum(Amount*RI),0)) as Amount from (" + Environment.NewLine +
                    "select TSPL_MILK_SRN_HEAD.DOC_CODE,(AMOUNT*TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per/100) as Amount,1 as RI,1 as Chk from TSPL_MILK_SRN_DETAIL" + Environment.NewLine +
                    "left outer join TSPL_MILK_SRN_HEAD	on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE" + Environment.NewLine +
                    "left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_MILK_SRN_HEAD.VSP_CODE" + Environment.NewLine +
                    "left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_MILK_SRN_HEAD.VSP_CODE" + Environment.NewLine +
                    "where TSPL_MILK_SRN_HEAD.Posted=1 and TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" & clsCommon.myCstr(VendorNo) & "' and TSPL_MILK_SRN_HEAD.DOC_DATE>='" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") & "' and TSPL_MILK_SRN_HEAD.DOC_DATE<='" & clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") & "' and TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per>=0" + Environment.NewLine +
                    "union all " + Environment.NewLine +
                    "select TSPL_SD_SHIPMENT_HEAD.Document_Code as DOC_CODE,TSPL_SD_SHIPMENT_HEAD.Total_Amt as Amount,-1 as RI,0 as chk  from TSPL_SD_SHIPMENT_HEAD where Trans_Type='MCC' and   TSPL_SD_SHIPMENT_HEAD.Customer_Code='" & clsCommon.myCstr(VendorNo) & "' " + Environment.NewLine +
                    "and TSPL_SD_SHIPMENT_HEAD.Document_Date>='" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") & "' and TSPL_SD_SHIPMENT_HEAD.Document_Date<='" & clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") & "'" & Environment.NewLine +
                    "union all" + Environment.NewLine +
                    "select TSPL_SD_SALE_RETURN_HEAD.Document_Code,TSPL_SD_SALE_RETURN_HEAD.Total_Amt,1 as RI,0 as chk from TSPL_SD_SALE_RETURN_HEAD where Trans_Type='MCC' and TSPL_SD_SALE_RETURN_HEAD.Customer_Code='" & clsCommon.myCstr(VendorNo) & "' and TSPL_SD_SALE_RETURN_HEAD.Status=1" + Environment.NewLine +
                    "and not exists(select 1 from TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN where TSPL_PAYMENT_PROCESS_MCC_SALE_RETURN.Return_Doc_No=TSPL_SD_SALE_RETURN_HEAD.Document_Code)" + Environment.NewLine +
                    "Union all" + Environment.NewLine +
                    "select TSPL_MILK_REJECT_HEAD.DOC_CODE,cast (((case when TSPL_MILK_REJECT_DETAIL.Is_Return = 0 and TSPL_MILK_REJECT_DETAIL.Defaulter='VSP' then TSPL_MILK_REJECT_DETAIL.SNF_Deduction_Amount+(case when TSPL_MILK_REJECT_DETAIL.Reject_Type='Curd' then TSPL_MILK_REJECT_DETAIL.FAT_Deduction_Amount else 0 end) else (case when TSPL_MILK_REJECT_DETAIL.Is_Return = 1 then TSPL_MILK_REJECT_DETAIL.Amount else (case when TSPL_MILK_REJECT_DETAIL.Is_Return = 3 then TSPL_MILK_REJECT_DETAIL.Amount else 0 end ) end ) end )*TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per/100) as decimal(18,2)) as Amount ,-1 as RI,0 as chk" + Environment.NewLine +
                    "from TSPL_MILK_REJECT_DETAIL" + Environment.NewLine +
                    "left outer join TSPL_MILK_REJECT_HEAD on TSPL_MILK_REJECT_HEAD.DOC_CODE=TSPL_MILK_REJECT_DETAIL.DOC_CODE" + Environment.NewLine +
                    "left outer join TSPL_CUSTOMER_VENDOR_MAPPING on TSPL_CUSTOMER_VENDOR_MAPPING.Vendor_Code=TSPL_MILK_REJECT_DETAIL.VSP_CODE" + Environment.NewLine +
                    "left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_MILK_REJECT_DETAIL.VSP_CODE" + Environment.NewLine +
                    "where TSPL_MILK_REJECT_HEAD.Posted=1 and TSPL_CUSTOMER_VENDOR_MAPPING.Cust_Code='" & clsCommon.myCstr(VendorNo) & "' and TSPL_MILK_REJECT_HEAD.DOC_DATE>='" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm:ss tt") & "' and TSPL_MILK_REJECT_HEAD.DOC_DATE<='" & clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtTo), "dd/MMM/yyyy hh:mm:ss tt") & "' and TSPL_VENDOR_MASTER.Credit_Limit_On_Milk_Receipt_Per>=0" + Environment.NewLine +
                     ")xx  "
                Dim dblBal As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, Nothing))
                lblUnbilledMilkAmt.Text = dblBal
            End If
        End If
    End Sub

    Private Sub btnCC_Click(sender As Object, e As EventArgs) Handles btnCC.Click
        Try
            Dim PreviousDate As Date = clsCommon.GETSERVERDATE()
            PreviousDate = PreviousDate.AddDays(-1)
            Dim qry1 As String = "select distinct TSPL_BOOKING_MATSER.Document_No as DocumentNo,convert(varchar(12),TSPL_BOOKING_MATSER.Document_date,103) as Document_date,TSPL_CUSTOMER_MASTER.Customer_Name,(select isnull((Select distinct '['+TSPL_LOCATION_MASTER.Location_Desc +']  ' from TSPL_booking_detail left outer join TSPL_LOCATION_MASTER on TSPL_BOOKING_DETAIL.Loc_Code=TSPL_LOCATION_MASTER.Location_Code where TSPL_booking_detail.Document_No=TSPL_BOOKING_MATSER.Document_No   for xml path('')),'')  ) as Location  ,case when TSPL_BOOKING_MATSER.Posted=1 then 'posted' else 'Unposted' end as Posted from TSPL_BOOKING_MATSER"
            qry1 += " left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_DETAIL.Document_No=TSPL_BOOKING_MATSER.Document_No "
            qry1 += " left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code "
            Dim whrClas As String = " TSPL_BOOKING_MATSER.comp_code='" & objCommonVar.CurrentCompanyCode & "' and TSPL_BOOKING_MATSER.Is_Taxable=2 "
            whrClas += " and convert(date,TSPL_BOOKING_MATSER.Document_Date,103) = convert(date,('" & PreviousDate & "'),103)"
            If clsCommon.myLen(txtVendorNo.Value) > 0 Then
                whrClas += " and TSPL_BOOKING_DETAIL.Cust_Code='" & txtVendorNo.Value & "'"
            End If
            Dim strCode As String = ""
            strCode = clsCommon.ShowSelectForm("PSBookingOrderNoFndd1", qry1, "DocumentNo", whrClas, "", "DocumentNo", True)
            If clsCommon.myLen(strCode) > 0 Then
                'FlagCopy = True
                Dim qry As String = ""
                Dim obj As New clsBookingEntryDairySale
                'Dim intRow As Integer
                obj = clsBookingEntryDairySale.GetData(strCode, NavigatorType.Current)
                If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
                    btnSave.Enabled = True
                    btnPost.Enabled = True
                    btnDelete.Enabled = True
                    isInsideLoadData = False
                    ' btnCopy.Enabled = True
                    isNewEntry = True
                    'btnSave.Text = "Update"
                    BlankAllControls()
                    LoadBlankGrid()
                    ' txtLocation.Enabled = False
                    txtVendorNo.Enabled = False
                    'If obj.TRANSACTION_TYPE = "FS" Then
                    '    rbtn_Fresh.IsChecked = True
                    'ElseIf obj.TRANSACTION_TYPE = "PS" Then
                    '    rbtn_Ambient.IsChecked = True
                    'End If
                    txtLocation.Value = obj.location_code
                    lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
                    GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
                    qry = "SELECT TSPL_BOOKING_DETAIL.*,TSPL_CUSTOMER_MASTER.Customer_Name FROM TSPL_BOOKING_DETAIL LEFT JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code WHERE Document_No='" & strCode & "' and scheme_item='N '"
                    Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
                    txtVendorNo.Value = clsCommon.myCstr(dt2.Rows(0)("Cust_Code"))
                    lblVendorName.Text = clsCommon.myCstr(dt2.Rows(0)("Customer_Name"))
                    lblTotRAmt1.Text = clsCommon.myCstr(dt2.Rows(0)("DocumentAmount"))
                    BookingStatus = 0 'clsCommon.myCstr(dt2.Rows(0)("Booking_Status"))
                    DOStatus = 0 'clsCommon.myCstr(dt2.Rows(0)("DO_Posted"))
                    qry = "select Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,Zone_Code,TSPL_CUSTOMER_MASTER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER left outer join " &
                            "TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join TSPL_VEHICLE_MASTER on " &
                            "TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where Cust_Code='" & txtVendorNo.Value & "'"
                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                        strVehicleCode = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
                        lblvehiclecode.Text = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
                        strVehicleDesc = clsCommon.myCstr(dt1.Rows(0)("Number"))
                        lblvehicleName.Text = clsCommon.myCstr(dt1.Rows(0)("Number"))
                        strRoutecode = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
                        lblroutecode.Text = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
                        strRouteDesc = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
                        lblroutename.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
                        Price_code = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
                        txtPriceCode.Text = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
                    End If
                    For jj As Integer = 0 To dt2.Rows.Count() - 1
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = gv1.Rows.Count 'clsCommon.myCdbl(dt2.Rows(jj)("Line_No"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colICode).Value = clsCommon.myCstr(dt2.Rows(jj)("Item_Code"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct item_desc from tspl_item_master where item_code='" & clsCommon.myCstr(dt2.Rows(jj)("Item_Code")) & "'"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")), Nothing)
                        If ShowAvailableQtyOnDairyBooking Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColAvailableQty).Value = clsItemLocationDetails.getBalance(clsCommon.myCstr(dt2.Rows(jj)("Item_Code")), IIf(clsCommon.myLen(txtSubLocation.Value) > 0, txtSubLocation.Value, txtLocation.Value), txtDocNo.Value, txtDate.Value, Nothing, clsCommon.myCstr(dt2.Rows(jj)("Unit_Code")), 0)
                        End If
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colOrgRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("OrgRate"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colUnit).Value = clsCommon.myCstr(dt2.Rows(jj)("Unit_Code"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSellingRate).Value = clsCommon.myCdbl(dt2.Rows(jj)("Item_Selling_Price"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = Math.Round(clsCommon.myCdbl(dt2.Rows(jj)("Booking_Qty")) * clsCommon.myCdbl(dt2.Rows(jj)("Item_Rate")), 2)
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_On_Amount"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt2.Rows(jj)("Tax_Amount"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Amount).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Amount"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Code).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Code"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Pers).Value = clsCommon.myCdbl(dt2.Rows(jj)("Disc_Scheme_Pers"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDisc_Scheme_Type).Value = clsCommon.myCstr(dt2.Rows(jj)("Disc_Scheme_Type"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colSchemeType).Value = clsCommon.myCstr(dt2.Rows(jj)("Scheme_Type"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colTax_NonTax).Value = clsCommon.myCstr(dt2.Rows(jj)("Tax_NonTax"))
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colFreshAmbient).Value = clsCommon.myCdbl(dt2.Rows(jj)("FreshAmbient"))
                        'gv1.CurrentRow.Cells(ColAvgQty).Value = clsDBFuncationality.getSingleValue(qry)
                        UpdateCurrentRowAvgQty(gv1.CurrentRow.Index)
                    Next
                End If
                gv1.Rows.AddNew()
                If gv1.Rows.Count > 0 Then
                    ' gv1.Focus()
                    gv1.CurrentRow = gv1.Rows(0)
                    gv1.CurrentColumn = gv1.Columns(colQty)
                End If
                'End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtVehicleCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVehicleCode._MYValidating
        Try
            Dim qry As String = "select Vehicle_Id as Code,Vehicle_Name,Model,Number,Description,Capacity from TSPL_VEHICLE_MASTER"
            Dim Whrcls As String = ""
            txtVehicleCode.Value = clsCommon.ShowSelectForm("DBC-Vehicle", qry, "Code", Whrcls, txtVehicleCode.Value, "Code", isButtonClicked)
            txtVehicleName.Text = clsDBFuncationality.getSingleValue("select Vehicle_Name from TSPL_VEHICLE_MASTER where Vehicle_id='" & txtVehicleCode.Value & "'")
            fndTransporter.Value = clsDBFuncationality.getSingleValue("select Transport_Id from TSPL_VEHICLE_MASTER where Vehicle_Id='" & txtVehicleCode.Value & "'")
            lblTransporter.Text = clsDBFuncationality.getSingleValue("  select Transporter_Name from TSPL_TRANSPORT_MASTER where Transport_Id='" & fndTransporter.Value & "'")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub cmbcashcredit_TextChanged(sender As Object, e As EventArgs) Handles cmbcashcredit.TextChanged
        Try
            If chkDCS.Checked Then
                If clsCommon.CompairString(cmbcashcredit.Text, "CASH") = CompairStringResult.Equal Then
                    lblPriceCodeDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select VSP_Price_Code_Cash from TSPL_customer_group_master where Default_VSP=1"))
                    Price_code = lblPriceCodeDesc.Text
                    txtPriceCode.Text = lblPriceCodeDesc.Text
                ElseIf clsCommon.CompairString(cmbcashcredit.Text, "CREDIT") = CompairStringResult.Equal Then
                    lblPriceCodeDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select VSP_Price_Code_Credit from TSPL_customer_group_master where Default_VSP=1"))
                    Price_code = lblPriceCodeDesc.Text
                    txtPriceCode.Text = lblPriceCodeDesc.Text
                End If
                txtRouteCode1.Enabled = True
                txtRouteName1.Enabled = True
                txtVehicleCode.Enabled = True
                txtVehicleName.Enabled = True
            Else
                txtRouteCode1.Enabled = False
                txtRouteName1.Enabled = False
                txtVehicleCode.Enabled = True
                txtVehicleName.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnCreateAndPrintInvoice_Click(sender As Object, e As EventArgs) Handles btnCreateAndPrintInvoice.Click


        'Dim Against_Delivery_Code As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Delivery_No from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "' "))
        Dim DocCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "'"))
        'Dim DocCode As String = txtDocNo.Value
        If clsCommon.myLen(DocCode) <= 0 Then


            Try

                If AllowToSave(Nothing) Then
                    Try
                        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()


                        Dim obj As New clsPSShipmentHead()
                        Dim str As String = String.Empty
                        Dim DCTotalAmt As Decimal = 0
                        Dim TCTotalAmt As Decimal = 0
                        Dim SCTotalAmt As Decimal = 0
                        obj.Vehicle_Code = clsCommon.myCstr(txtVehicleCode.Value)
                        If chkManualVehicle.Checked Then
                            obj.VehicleNo = txtManualVehicle.Text
                        Else
                            obj.VehicleNo = clsDBFuncationality.getSingleValue("select Number from TSPL_VEHICLE_MASTER where Vehicle_id='" & txtVehicleCode.Value & "'", trans)
                        End If
                        obj.Description = txtDescription.Text
                        obj.AdharNo = txtAdharNo.Text
                        obj.MobileNO = txtMobileNo.Text

                        obj.Transport_Id = fndTransporter.Value
                        obj.Transporter_Name = lblTransporter.Text
                        obj.IsSampling = IIf(chkSampling.Checked, 1, 0)
                        obj.IsEwaybill = IIf(chkIsEwayBill.Checked, 1, 0)
                        obj.No_Transporter = IIf(chkNoTranspoter.Checked, 1, 0)
                        'obj.ShippedCAN = txtCan.Value
                        obj.TotalCAN = txtCan.Text
                        obj.CrateQty = txtCrate.Text
                        obj.Crate = txtCrate.Text
                        obj.Box = txtBox.Text
                        obj.Exclude_KKF_And_Mandi = IIf(chkExcludeKKFMND.Checked, 1, 0)
                        'obj.Against_Delivery_Code = Against_Delivery_Code 'clsDBFuncationality.getSingleValue("select Delivery_No from TSPL_BOOKING_DETAIL where Document_No='" & txtDocNo.Value & "' ", trans)
                        obj.Against_Booking_No = txtDocNo.Value
                        obj.Payment_Terms = cmbPaymentType.Text
                        obj.Vehicle_Manual_No = txtManualVehicle.Text
                        obj.ReceiverName = txtReceiverName.Text
                        obj.Discount_Base = clsCommon.myCdbl(lblAmtWithDiscount.Text)
                        obj.Discount_Amt = clsCommon.myCdbl(lblDiscountAmt.Text)
                        obj.Amount_Less_Discount = clsCommon.myCdbl(lblAmtAfterDiscount.Text)
                        obj.RoundOffAmount = clsCommon.myCdbl(TxtRoundoff.Text)
                        obj.Sub_Location_code = txtSubLocation.Value
                        obj.FAT_Per = clsCommon.myCdbl(txtFATPER.Text)
                        obj.SNF_Per = clsCommon.myCdbl(txtSNFPER.Text)
                        obj.Acidity = clsCommon.myCdbl(txtAcidity.Text)
                        obj.Temperature = clsCommon.myCdbl(txtTemp.Text)
                        obj.MBRT_Hours = clsCommon.myCdbl(txtMBRTHours.Text)
                        obj.Item_Type = "S"
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                            If chkcashsale.Checked AndAlso clsCommon.CompairString(cmbPaymentType.Text, "Cash") = CompairStringResult.Equal Then
                                obj.Is_CashSale = "Y"
                            End If
                        Else
                            If chkcashsale.Checked Then
                                obj.Is_CashSale = "Y"
                            End If
                        End If


                        'obj.can
                        obj.Screen_Type = "DS"
                        ' obj.Scheme_Tax_Group = txtSchemeTaxGroup.Value
                        If rbtnTaxable.IsChecked Then
                            obj.DO_Item_Type = "T"
                            obj.Invoice_Type = "T"
                            obj.Is_Taxable = 1
                        Else
                            obj.DO_Item_Type = "NT"
                            obj.Invoice_Type = "R"
                            obj.Is_Taxable = 0
                        End If
                        obj.Route_No = txtRouteNo.Value
                        obj.Route_Desc = lblRouteDesc.Text
                        'obj.Price_Group_Code = txtPriceGroupCode.Text
                        obj.Price_Code = txtPriceCode.Text
                        obj.Document_Date = txtDate.Value
                        obj.Dispatch_date = txtDate.Value
                        obj.Challan_Date = clsCommon.GETSERVERDATE(trans)
                        obj.Inv_Date = clsCommon.GETSERVERDATE(trans)
                        obj.Customer_Code = txtVendorNo.Value
                        obj.Customer_Name = lblVendorName.Text
                        obj.Bill_To_Location = txtLocation.Value
                        obj.Ship_To_Location = txtShipToLocation.Value
                        obj.Trans_Type = "FS"
                        'obj.Against_Delivery_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No from TSPL_DELIVERY_NOTE_MASTER_FRESHSALE where Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "'", trans))
                        obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                        obj.Is_Create_Auto_Invoice = 1
                        obj.Supply_Date = txtSupplyDate.Value
                        If clsCommon.CompairString(cmbGatePassType.Text, "Select") = CompairStringResult.Equal Then
                            obj.Shift_Type = ""
                        Else
                            obj.Shift_Type = clsCommon.myCstr(cmbGatePassType.Text)
                        End If
                        obj.Tax_Group = txtTaxGroup.Value
                        obj.TaxGroupName = lblTaxGrpName.Text
                        obj.Terms_Code = txtTermCode.Value
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
                        obj.Arr = New List(Of clsPSShipmentHeadDetail)
                        Dim ii As Integer = 0
                        For Each grow As GridViewRowInfo In gv1.Rows
                            ii += 1
                            If grow.Cells(colICode).Value IsNot Nothing Then

                                Dim objTr As New clsPSShipmentHeadDetail()
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(1)).Value)) > 0) Then
                                    objTr.TAX1 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(1)).Value)
                                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(1)).Value)
                                    objTr.TAX1_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(1)).Value)
                                    objTr.TAX1_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(1)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(2)).Value)) > 0) Then
                                    objTr.TAX2 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(2)).Value)
                                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(2)).Value)
                                    objTr.TAX2_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(2)).Value)
                                    objTr.TAX2_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(2)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(3)).Value)) > 0) Then
                                    objTr.TAX3 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(3)).Value)
                                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(3)).Value)
                                    objTr.TAX3_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(3)).Value)
                                    objTr.TAX3_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(3)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(4)).Value)) > 0) Then
                                    objTr.TAX4 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(4)).Value)
                                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(4)).Value)
                                    objTr.TAX4_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(4)).Value)
                                    objTr.TAX4_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(4)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(5)).Value)) > 0) Then
                                    objTr.TAX5 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(5)).Value)
                                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(5)).Value)
                                    objTr.TAX5_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(5)).Value)
                                    objTr.TAX5_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(5)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(6)).Value)) > 0) Then
                                    objTr.TAX6 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(6)).Value)
                                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(6)).Value)
                                    objTr.TAX6_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(6)).Value)
                                    objTr.TAX6_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(6)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(7)).Value)) > 0) Then
                                    objTr.TAX7 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(7)).Value)
                                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(7)).Value)
                                    objTr.TAX7_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(7)).Value)
                                    objTr.TAX7_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(7)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(8)).Value)) > 0) Then
                                    objTr.TAX8 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(8)).Value)
                                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(8)).Value)
                                    objTr.TAX8_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(8)).Value)
                                    objTr.TAX8_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(8)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(9)).Value)) > 0) Then
                                    obj.Tax_Group = clsCommon.myCstr(grow.Cells(colTaxGroup).Value)
                                    objTr.TAX9 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(9)).Value)
                                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(9)).Value)
                                    objTr.TAX9_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(9)).Value)
                                    objTr.TAX9_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(9)).Value)
                                End If
                                If (clsCommon.myLen(clsCommon.myCdbl(grow.Cells(colTax + clsCommon.myCstr(10)).Value)) > 0) Then
                                    objTr.TAX10 = clsCommon.myCstr(grow.Cells(colTax + clsCommon.myCstr(10)).Value)
                                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(grow.Cells(colTax_Base_Amt + clsCommon.myCstr(10)).Value)
                                    objTr.TAX10_Rate = clsCommon.myCdbl(grow.Cells(colTax_Rate + clsCommon.myCstr(10)).Value)
                                    objTr.TAX10_Amt = clsCommon.myCdbl(grow.Cells(colTax_Amt + clsCommon.myCstr(10)).Value)
                                End If
                                objTr.Sub_Location_code = txtSubLocation.Value
                                objTr.Customer_Code = txtVendorNo.Value
                                objTr.Item_Code = clsCommon.myCstr(grow.Cells(colICode).Value)
                                objTr.Item_Desc = clsCommon.myCstr(grow.Cells(colIName).Value)
                                If ApplyManualScheme Then
                                    objTr.Scheme_Main_Item = clsCommon.myCstr(grow.Cells(ColMainItem).Value)
                                End If
                                objTr.Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                                objTr.Billing_Qty = clsCommon.myCDecimal(grow.Cells(ColBillingQty).Value)
                                objTr.Crate = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                                If chkSampling.Checked = True Then
                                    objTr.Item_Cost = clsCommon.myCDecimal(0)
                                Else
                                    objTr.Item_Cost = clsCommon.myCDecimal(grow.Cells(colOrgRate).Value)
                                End If
                                'objTr.Item_Cost = clsCommon.myCDecimal(grow.Cells(colOrgRate).Value)
                                objTr.Amount = clsCommon.myCDecimal(grow.Cells(colAmt).Value)
                                objTr.Unit_code = clsCommon.myCstr(grow.Cells(colUnit).Value)
                                objTr.Billing_Unit_code = clsCommon.myCstr(grow.Cells(colBillingUOM).Value)
                                objTr.Sampling = IIf(chkSampling.Checked, 1, 0)
                                objTr.Line_No = clsCommon.myCdbl(grow.Cells(colLineNo).Value)
                                objTr.Location = txtLocation.Value
                                objTr.MRP = clsDBFuncationality.getSingleValue("select Item_MRP from TSPL_ITEM_PRICE_MASTER where Item_Price_ID=" + clsCommon.myCstr(grow.Cells(colPriceId).Value) + "", trans)
                                objTr.Total_Tax_Amt = clsCommon.myCdbl(grow.Cells(colTTaxAmt).Value)
                                objTr.Total_MRP_Amt = clsCommon.myCdbl(grow.Cells(colAmountWithTax).Value)
                                objTr.Total_Basic_Amt = clsCommon.myCDecimal(grow.Cells(colAmt).Value)
                                objTr.Row_Type = "Item"
                                objTr.Trip_No = 1
                                objTr.Scheme_Item = clsCommon.myCstr(grow.Cells(colSchemeItem).Value)
                                objTr.Amt_Less_Discount = clsCommon.myCDecimal(grow.Cells(colAmtAfterDis).Value)
                                objTr.Item_Net_Amt = objTr.Total_Tax_Amt + objTr.Amt_Less_Discount
                                objTr.Distributor_Commission_PKID = clsCommon.myCstr(grow.Cells(ColDCPKID).Value)
                                objTr.Distributor_Commission_Rate = clsCommon.myCdbl(grow.Cells(ColDCRate).Value)
                                objTr.Distributor_Commission_Amt = clsCommon.myCdbl(grow.Cells(ColDCAmt).Value)
                                objTr.Transporter_Commission_Rate = clsCommon.myCdbl(grow.Cells(ColTCRate).Value)
                                objTr.Transporter_Commission_Amt = clsCommon.myCdbl(grow.Cells(ColTCAmt).Value)
                                objTr.Security_Amt = clsCommon.myCdbl(grow.Cells(ColSCAmt).Value)
                                objTr.Security_Rate = clsCommon.myCdbl(grow.Cells(ColSCRate).Value)
                                objTr.Against_Booking_PK_ID = clsCommon.myCdbl(grow.Cells(colPKID).Value)
                                objTr.Distributor_Commission_RateWithTax = clsCommon.myCdbl(grow.Cells(ColDCRateWithTax).Value)
                                objTr.Disc_Amt = clsCommon.myCdbl(grow.Cells(colDisAmt).Value)
                                objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))
                                'objTr.Amt_Less_Discount = clsCommon.myCdbl(grow.Cells(colAmtAfterDis).Value)
                                DCTotalAmt += objTr.Distributor_Commission_Amt
                                TCTotalAmt += objTr.Transporter_Commission_Amt
                                SCTotalAmt += objTr.Security_Amt

                                If clsCommon.myLen(txtDocNo.Value) > 0 Then

                                    Dim strQry As String = "INSERT INTO TSPL_BATCH_ITEM_CUSTOMER_BOOKING SELECT * FROM TSPL_BATCH_ITEM  where Document_Code='" & txtDocNo.Value & "' and Item_Code='" & objTr.Item_Code & "' and UOM='" & objTr.Unit_code & "'"
                                    clsDBFuncationality.ExecuteNonQuery(strQry, trans)
                                    strQry = "delete TSPL_BATCH_ITEM  where Document_Code='" & txtDocNo.Value & "' and Item_Code='" & objTr.Item_Code & "' and UOM='" & objTr.Unit_code & "'"
                                    clsDBFuncationality.ExecuteNonQuery(strQry, trans)
                                End If
                                If (clsCommon.myLen(objTr.Item_Code) > 0) Then
                                    obj.Arr.Add(objTr)
                                End If
                            End If
                        Next
                        obj.Distributor_Commission_TotalAmt = DCTotalAmt
                        obj.Transporter_Commission_TotalAmt = TCTotalAmt
                        obj.Security_TotalAmt = SCTotalAmt
                        obj.Total_Tax_Amt = clsCommon.myCdbl(lblTaxAmt.Text)
                        obj.Total_Amt = clsCommon.myCdbl(lblTotRAmt.Text)
                        If DeductTPTFromDocAmt Then
                            obj.TotalSubsidyAmt = clsCommon.myCdbl(txtTPTComAmt.Text)
                            obj.Gross_Amount = clsCommon.myCdbl(lblGrossAmount.Text)
                        Else
                            obj.Gross_Amount = clsCommon.myCdbl(lblGrossAmount.Text)
                        End If
                        ' obj.Discount_Amt = DCTotalAmt
                        isNewEntry = True
                        obj.Document_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & obj.Against_Booking_No & "'  and Customer_Code='" & txtVendorNo.Value & "'", trans))
                        DocCode = obj.Document_Code
                        If clsCommon.myLen(obj.Document_Code) <= 0 Then

                            If (clsPSShipmentHead.SaveData(obj, isNewEntry, trans, True)) Then
                                Dim strupdate As String = "update TSPL_SD_SHIPMENT_HEAD set ParentDocNo='" & obj.Document_Code & "' where Document_Code='" & obj.Document_Code & "'"
                                clsDBFuncationality.ExecuteNonQuery(strupdate, trans)
                                'trans.Commit()
                                'clsPSShipmentHead.PostData("DISPATCH-DS", obj.Document_Code, trans, Nothing, True, "")
                                'clsPSShipmentHead.PostData(MyBase.Form_ID, obj.Document_Code)
                                DocCode = obj.Document_Code
                                'Dim QryBatchUpdate As String = "update  TSPL_BATCH_ITEM set In_Out_Type='N' where Document_Type='FS-SH' and  Document_Code='" & txtDocNo.Value & "'"
                                'clsDBFuncationality.ExecuteNonQuery(QryBatchUpdate, trans)
                                trans.Commit()
                                'clsCommon.MyMessageBoxShow(Me, "Document Created Successfully.", Me.Text)
                            Else
                                DocCode = obj.Document_Code
                                trans.Rollback()
                            End If
                        Else
                            trans.Commit()
                            'clsCommon.MyMessageBoxShow(Me, "Document Already Exists!", Me.Text)
                        End If
                    Catch ex As Exception
                        'trans.Rollback()
                        Throw New Exception(ex.Message)
                    End Try
                End If

            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)

            End Try

        End If
        If clsCommon.myLen(DocCode) > 0 Then
            Try
                Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                Dim DocCode1 As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "' and status =0", trans))
                If clsCommon.myLen(DocCode1) > 0 Then
                    Try
                        clsPSShipmentHead.PostData("DISPATCH-DS", DocCode1, trans, Nothing, True, "")
                        trans.Commit()
                        clsCommon.MyMessageBoxShow(Me, "Document Created Successfully.", Me.Text)
                    Catch ex As Exception
                        trans.Rollback()
                        Throw New Exception(ex.Message)
                    End Try
                    Try
                        GetOpeningClosingAndReceivedAmt(txtVendorNo.Value, txtDate.Value)
                        Dim balqry As String = "Update TSPL_BOOKING_MATSER set OpeningBal='" & clsCommon.myCstr(txtOpeningbal.Text) & "', " &
                            "DrAmt='" & clsCommon.myCstr(txtDrAmt.Text) & "', " &
                            "CrAmt='" & clsCommon.myCstr(txtCrAmt.Text) & "'," &
                            "ClosingBal='" & clsCommon.myCstr(txtClosingBal.Text) & "' " &
                         " where Document_No='" & txtDocNo.Value & "'"
                        clsDBFuncationality.ExecuteNonQuery(balqry)
                    Catch ex As Exception

                    End Try
                Else
                    trans.Commit()
                    ' clsCommon.MyMessageBoxShow(Me, "Document Not Found!", Me.Text)
                End If

            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try
        End If
        printInvoice(Nothing, DocCode, txtDate.Value)
    End Sub

    Sub printInvoice(ByVal BookingCode As String, ByVal DocCode As String, Optional ByVal chkDate As Date = Nothing, Optional ByVal chkCancel As Boolean = False)
        Try
            Dim DocWithIRN As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "' and status =1 and Is_Taxable=1"))
            Dim isTaxable As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Is_Taxable from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "'"))
            Dim flag As Boolean = False
            If clsCommon.myLen(DocCode) > 0 Then
                If isTaxable = 1 Then
                    If clsCommon.myLen(DocWithIRN) > 0 Then
                        If common.clsCommon.MyMessageBoxShow(" E-Invoice generated, Do you want to print ", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                            flag = True
                        End If
                    Else
                        If common.clsCommon.MyMessageBoxShow("E-Invoice not generated, Do you want to print", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                            flag = True
                        End If

                    End If
                Else
                    If common.clsCommon.MyMessageBoxShow("Print Invoice", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                        flag = True
                    End If
                End If

                If flag Then
                    '' Print '''''''''''
                    Dim frmCRV As New frmCrystalReportViewer()
                    Dim objMultPrintInvoice As New FrmPrintFreshInvoice
                    Dim SaleInvoiceNo As New List(Of String)
                    SaleInvoiceNo.Add(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SALE_INVOICE_head where Against_Shipment_No ='" & DocCode & "'"))
                    Dim Qry As String = Nothing
                    If chkCancel Then
                        Qry = objMultPrintInvoice.PrintInvoiceForAll("'" & clsCommon.myCstr(DocCode) & "'", chkDate, Nothing, "Y", "Cancel")
                        'Qry = objMultPrintInvoice.PrintInvoiceForAll(clsCommon.GetMulcallString(SaleInvoiceNo), chkDate, Nothing, "Y", chkCancel)

                    Else
                        Qry = objMultPrintInvoice.PrintInvoiceForAll(clsCommon.GetMulcallString(SaleInvoiceNo), chkDate, txtVendorNo.Value)
                    End If
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal AndAlso dt.Rows(0)("TaxableNonTaxable").ToString() = "T" Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceBKN", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceNAG1", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())

                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("TaxableNonTaxable")), "T") = CompairStringResult.Equal Then
                            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceBKN", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        Else
                            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptMultipleInvoicePrintBKN", "Bill Of Supply")
                        End If
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceGNG", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceBAR", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJPR", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceRJS", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())

                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceSKR", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                        'frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJDH", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJDH", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceBHR", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal AndAlso dt.Rows(0)("TaxableNonTaxable").ToString() = "T" Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceALW1", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceALW1", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
                        'frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceSWM1", "Bill of Supply", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceTNK", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceCHU1", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJALNEW", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceTNKNEW", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                        If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("TaxableNonTaxable")), "T") = CompairStringResult.Equal Then
                            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceAJMNew", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        Else
                            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoice_AJM", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                        End If
                    Else
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", clsCommon.GetPrintDate(chkDate), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    End If
                    frmCRV = Nothing
                    ''   end of print Invoice '''''''
                End If
            Else
                If chkCancel Then
                    Throw New Exception("Invoice No not created against Booking (" & BookingCode & ").")
                Else
                    Throw New Exception("Invoice Not Found!")
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub chkBPL_CheckStateChanged(sender As Object, e As EventArgs) Handles chkBPL.CheckStateChanged
        If chkBPL.Checked Then
            BPLController(True)
        Else
            BPLController(False)
        End If
    End Sub
    Public Sub BPLController(ByVal flag As Boolean)
        lblCouponCode.Visible = flag
        txtCouponCode.Visible = flag
        lblBPLName.Visible = flag
        txtBPLName.Visible = flag
        lblBPLRemark.Visible = flag
        txtBPLRemark.Visible = flag
        lblCouponDate.Visible = flag
        txtCouponDate.Visible = flag
        txtCategory.Visible = flag
        lblCategory.Visible = flag
        lblIndentNo.Visible = Not flag
        txtIndentNo.Visible = Not flag
        lblIndentDate.Visible = Not flag
        txtIndentDate.Visible = Not flag
        If flag Then
            RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Visible
            chkDCS.Enabled = False
            chkDistributor.Checked = False
            chkDistributor.Enabled = False
            chkAPS.Checked = False
            chkAPS.Enabled = False
            txtCouponDate.Value = clsCommon.GETSERVERDATE()
            lblIndentNo.Location = New Point(633, 190)
            txtIndentNo.Location = New Point(710, 189)
            lblIndentDate.Location = New Point(830, 192)
            txtIndentDate.Location = New Point(915, 191)
        Else
            RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
            chkDCS.Enabled = True
            chkDistributor.Checked = False
            chkDistributor.Enabled = True
            chkAPS.Checked = False
            chkAPS.Enabled = True
            txtCouponDate.Value = Nothing
            lblIndentNo.Location = New Point(429, 47)
            txtIndentNo.Location = New Point(506, 46)
            lblIndentDate.Location = New Point(429, 70)
            txtIndentDate.Location = New Point(512, 69)
        End If
    End Sub
    'Public Function GetTax(ByVal obj As clsTaxCalculation, ByVal rowno As Integer, ByVal qty As Decimal) As Decimal
    '    Dim dbltotalTaxAmt As Decimal = 0
    '    Try
    '        If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
    '            gv1.Rows(rowno).Cells(colTaxGroup).Value = clsCommon.myCstr(obj.Tax_Group_Code)
    '            For count As Integer = 1 To 10
    '                gv1.Rows(rowno).Cells(colTax + clsCommon.myCstr(count)).Value = clsCommon.myCstr(obj.Arr(count - 1).Tax_Code)
    '                gv1.Rows(rowno).Cells(colTax_Base_Amt + clsCommon.myCstr(count)).Value = clsCommon.myCdbl(obj.Arr(count - 1).Tax_BaseAmt) * qty
    '                gv1.Rows(rowno).Cells(colTax_Rate + clsCommon.myCstr(count)).Value = clsCommon.myCdbl(obj.Arr(count - 1).Tax_Rate)
    '                gv1.Rows(rowno).Cells(colTax_Amt + clsCommon.myCstr(count)).Value = clsCommon.myCdbl(clsCommon.myCdbl(obj.Arr(count - 1).Tax_BaseAmt) * qty) * (clsCommon.myCdbl(obj.Arr(count - 1).Tax_Rate) / 100)
    '                dbltotalTaxAmt = dbltotalTaxAmt + clsCommon.myCdbl(gv1.Rows(rowno).Cells(colTax_Amt + clsCommon.myCstr(count)).Value)
    '            Next
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    '    Return dbltotalTaxAmt
    'End Function
    Private Sub txtCategory__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtCategory._MYValidating
        Try
            Dim qry As String = "select CAST_CATEGORY_CODE as Code,CAST_CATEGORY_NAME as [Category Name] from TSPL_CAST_CATEGORY_MASTER"
            txtCategory.Value = clsCommon.ShowSelectForm("DSRouteFinder", qry, "Code", "", txtCategory.Value, "", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub BtnRecieptEntry_Click(sender As Object, e As EventArgs) Handles BtnRecieptEntry.Click
        Try
            Dim chk As Boolean = True
            Dim CustRoute As String = clsDBFuncationality.getSingleValue("Select  TSPL_BOOKING_DETAIL.Cust_Code + ',' + CONVERT(VARCHAR ,TSPL_BOOKING_MATSER.Document_Date , 103) AS Cust_Code from TSPL_BOOKING_DETAIL
            Left OUTER JOIN TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No = TSPL_BOOKING_DETAIL.Document_No
            WHERE TSPL_BOOKING_DETAIL.CUST_CODE = '" & txtVendorNo.Value & "' AND TSPL_BOOKING_DETAIL.route_no = '" & txtRouteNo.Value & "' ")
            clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.ReceiptEntry, CustRoute)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnGatepass_Click(sender As Object, e As EventArgs) Handles btnGatepass.Click
        Try

            Dim frm As New frmDairyGatePass
            frm.routeno = txtRouteNo.Value
            frm.txtlocation = txtLocation.Value
            frm.vehicleno = txtVehicleCode.Value
            If chkManualVehicle.Checked Then
                frm.vehicle_desc = txtManualVehicle.Text
            Else
                frm.vehicle_desc = txtVehicleName.Text
            End If
            frm.DriverName = txtReceiverName.Text
            frm.docdate = txtDate.Value
            frm.Supplydate = txtSupplyDate.Value
            frm.Shifttype = cmbGatePassType.Text
            If GenerateCustomerWiseGatePass Then
                frm.GenerateCustomerWiseGatePass = True
                Dim strShipmentDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "' and Status=1 "))
                If clsCommon.myLen(strShipmentDocNo) > 0 Then
                    frm.ShipmentDocNo = strShipmentDocNo
                Else
                    Throw New Exception("Please Create Invoice.")
                End If
            Else
                frm.GenerateCustomerWiseGatePass = False
            End If
            Dim qry As String = (" Select Credit_Customer,IsDistributor from tspl_customer_master where cust_code =  '" & txtVendorNo.Value & "' ")
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt.Rows.Count > 0 Then
                Dim creditCustomerFlag As String = dt.Rows(0)("Credit_Customer").ToString()
                Dim isDistributorFlag As String = dt.Rows(0)("IsDistributor").ToString()
                If creditCustomerFlag = "Y" AndAlso isDistributorFlag = "N" Then
                    frm.CreditCustomer = lblVendorName.Text
                ElseIf creditCustomerFlag = "N" AndAlso isDistributorFlag = "Y" Then
                    frm.CreditCustomer = ""
                ElseIf creditCustomerFlag = "N" AndAlso isDistributorFlag = "N" Then
                    frm.CreditCustomer = lblVendorName.Text
                End If
            End If
            'Dim qry As String = clsDBFuncationality.getSingleValue("  Select Credit_Customer from tspl_customer_master where cust_code =  '" & txtVendorNo.Value & "' ")
            'If qry = "Y" Then
            '    frm.CreditCustomer = lblVendorName.Text
            'Else
            '    frm.CreditCustomer = ""
            'End If
            'frm.CreditCustomer = lblVendorName.Text
            frm.ShowDialog()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub RadPageView1_SelectedPageChanged(sender As Object, e As EventArgs) Handles RadPageView1.SelectedPageChanged
    End Sub
    Private Sub rbtnTaxCalAutomatic_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnTaxCalAutomatic.ToggleStateChanged
        If Not isInsideLoadData Then
            If rbtnTaxCalAutomatic.IsChecked Then
                'SetTaxDetails()
            ElseIf rbtnTaxCalManual.IsChecked Then
                For intRowNo As Integer = 0 To gv2.Rows.Count - 1
                    gv2.Rows(intRowNo).Cells(colTTaxRate).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTTaxAmt).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTBaseAmt).Value = Nothing
                Next
                For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                    For ii As Integer = 1 To 10
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = Nothing
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Amt" + strII)).Value = Nothing
                    Next
                Next
            End If
        End If
    End Sub
    Sub SetTaxDetails(ByVal ICode As String, ByVal intRow As Integer)
        LoadBlankGridTax()
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsByLocation(txtTaxGroup.Value, "S", txtVendorNo.Value, txtLocation.Value)
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
                                    txttcstaxbaseamount.Value = 0
                                    dblOutstandingAmount = clsCommon.myCdbl(clsCustomerMaster.GetCustomerOutstandingForTCSTaxApplicableOnFY(txtVendorNo.Value, txtDate.Value))
                                    If dblOutstandingAmount < AmountToCheckCustomerOutstandingForTCSTax Then
                                        dblOutstandingAmount = dblOutstandingAmount + clsCommon.myCdbl(clsCommon.myFormat(lblActualTCSTaxBaseAmt.Text))
                                        If dblOutstandingAmount > AmountToCheckCustomerOutstandingForTCSTax Then
                                            If clsCommon.myCdbl(clsCommon.myFormat(lblActualTCSTaxBaseAmt.Text)) > 0 Then
                                                txttcstaxbaseamount.Value = Math.Round(clsCommon.myCdbl(dblOutstandingAmount - AmountToCheckCustomerOutstandingForTCSTax), 2)
                                            End If
                                        End If
                                    End If
                                    If dblOutstandingAmount > AmountToCheckCustomerOutstandingForTCSTax Then
                                        If EnableTCSRateValidityFrom01July2021 Then
                                            Dim Is_ITR_Filled_And_TCSAmountGreater50K As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue(" SELECT CASE WHEN ISNULL(IsTCSGreaterthan50K,0)=1 AND ISNULL(IsITRfilledinLast2Years,0)=1 THEN 1 ELSE 0 END FROM TSPL_CUSTOMER_MASTER WHERE Cust_Code='" & txtVendorNo.Value & "'")) = 1, True, False)
                                            If Is_ITR_Filled_And_TCSAmountGreater50K = True Then
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                                            Else
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                                            End If
                                        Else
                                            Dim panno As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(pan,'')+isnull(Additional3 ,'') as PanNoAdhar from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'"))
                                            If clsCommon.myLen(panno) > 0 Then
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                                            Else
                                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                                            End If
                                        End If
                                    Else
                                        gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = 0
                                    End If
                                End If
                            Else
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = dr("TaxRate")
                            End If
                        Else
                            gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = 0
                        End If
                        txtTCSTaxRate.Value = clsCommon.myCdbl(gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value)
                    Else
                        If chkExcludeKKFMND.Checked AndAlso (clsCommon.CompairString(clsTaxCalculation.GetTaxType(clsCommon.myCstr(dr("Tax_Code")), Nothing), "M") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsTaxCalculation.GetTaxType(clsCommon.myCstr(dr("Tax_Code")), Nothing), "K") = CompairStringResult.Equal) Then
                            gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = 0
                        Else
                            gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = dr("TaxRate")
                        End If
                    End If
                End If
            Next
            SetitemWiseTaxSetting(True, ICode, intRow)
        Else
            lblTaxGrpName.Text = ""
            For ii As Integer = 0 To gv1.Rows.Count - 1
                BlankTaxDetails(ii)
            Next
        End If
        ' For ii As Integer = 0 To gv1.Rows.Count - 1
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            UpdateCurrentRow1(intRow)
        Else
            UpdateCurrentRow(intRow)
        End If
        ' Next
        UpdateAllTotals()
    End Sub
    Private Sub SetTax(ByVal Item_Code As String, ByVal intRow As Integer)

        GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
        If GSTStatus = False OrElse (rbtnTaxable.IsChecked AndAlso GSTStatus) Then
            If CalculateTaxRatefromItemwsieTaxOnSale Then
                If clsCommon.myLen(txtLocation.Value) > 0 Then
                    Dim strTaxType As String = clsLocationWiseTax.TaxType(txtLocation.Value, txtVendorNo.Value, "S", txtDate.Value, Nothing)
                    If GSTStatus = True AndAlso clsCommon.CompairString(strTaxType, "L") = CompairStringResult.Equal Then
                        txtTaxGroup.Value = clsItemWiseTaxAuthority.GetTaxGroupItemWise("L", "S", txtDate.Value, Item_Code)
                    Else
                        txtTaxGroup.Value = clsItemWiseTaxAuthority.GetTaxGroupItemWise("I", "S", txtDate.Value, Item_Code)
                    End If
                Else
                    txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
                End If
            Else
                txtTaxGroup.Value = clsLocationWiseTax.GetDefaultTaxGroup(txtLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
                lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, Nothing)
            End If
        Else
            If rbtnNonTax.IsChecked Then
                txtTaxGroup.Value = clsLocationWiseTax.GetExempedDefaultTaxGroup(True, txtLocation.Value, txtVendorNo.Value, "S", txtDate.Value)
                lblTaxGrpName.Text = clsTaxGroupMaster.GetNameOfSaleType(txtTaxGroup.Value, Nothing)
            End If
        End If
        'SetTaxDetails(Item_Code, intRow)
    End Sub
    Sub SetitemWiseTaxSetting(ByVal isChangeRate As Boolean, ByVal ICode As String, ByVal intRowNo As Integer)
        Dim strCustomer As String = ""
        Try
            strCustomer = clsCommon.myCstr(txtVendorNo.Value)
        Catch ex As Exception
        End Try
        If clsCommon.myLen(strCustomer) <= 0 Then
            strCustomer = txtVendorNo.Value
        End If
        Dim IsTaxable As Integer = 0
        Dim dt As DataTable = clsTaxCalculation.GetTaxDetailsByLocation(txtTaxGroup.Value, "S", txtVendorNo.Value, txtLocation.Value, ICode, clsCommon.GetPrintDate(txtDate.Value))
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            'For intRowNo As Integer = 0 To gv1.Rows.Count - 1
            If (clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode).Value) > 0) Then
                BlankTaxDetails(intRowNo, isChangeRate)
                'IsTaxable = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsTaxable from TSPL_ITEM_MASTER where item_code='" & clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colICode).Value) & "'"))
                IsTaxable = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select top 1 Is_Taxable from TSPL_ITEM_MASTER_TAXABLE where Item_Code='" & clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_MASTER_TAXABLE.EFFECTIVE_DATE <= '" & clsCommon.GetPrintDate(txtDate.Value) & "' ORDER BY TSPL_ITEM_MASTER_TAXABLE.EFFECTIVE_DATE DESC "))
                If ((clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode).Value) > 0 And IsTaxable = 1) OrElse (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) FROM TSPL_TAX_MASTER WHERE Tax_Code IN (select Tax_Code  from TSPL_TAX_GROUP_DETAILS WHERE TAX_GROUP_CODE='" & txtTaxGroup.Value & "') AND Is_TCS ='Y'")) > 0)) Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                        '' == Changes by Parteek 21/09/2017
                        'If CalculateTaxRatefromItemwsieTaxOnSale = 0 Then
                        '    If isChangeRate Then
                        If chkExcludeKKFMND.Checked AndAlso (clsCommon.CompairString(clsTaxCalculation.GetTaxType(clsCommon.myCstr(dr("Tax_Code")), Nothing), "M") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsTaxCalculation.GetTaxType(clsCommon.myCstr(dr("Tax_Code")), Nothing), "K") = CompairStringResult.Equal) Then
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = 0
                        Else
                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))

                        End If
                        '    End If
                        'End If
                        ''tcs tax rate
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Is_TCS  from tspl_tax_master where tax_code ='" & clsCommon.myCstr(dr("Tax_Code")) & "' ")), "Y") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(IsTCSnotApplicable ,0) from TSPL_CUSTOMER_MASTER where Cust_Code ='" & txtVendorNo.Value & "'")), "0") = CompairStringResult.Equal Then
                                If chkisTCS.Checked Then
                                    If AmountToCheckCustomerOutstandingForTCSTax > 0 Then
                                        Dim dblOutstandingAmount As Double = clsCommon.myCdbl(clsCustomerMaster.GetCustomerOutstandingForTCSTaxApplicableOnFY(txtVendorNo.Value, txtDate.Value))
                                        If dblOutstandingAmount <= AmountToCheckCustomerOutstandingForTCSTax Then
                                            dblOutstandingAmount = dblOutstandingAmount + clsCommon.myCdbl(clsCommon.myFormat(lblActualTCSTaxBaseAmt.Text))
                                            If dblOutstandingAmount >= AmountToCheckCustomerOutstandingForTCSTax Then
                                                If clsCommon.myCdbl(clsCommon.myFormat(lblActualTCSTaxBaseAmt.Text)) > 0 Then
                                                    txttcstaxbaseamount.Value = Math.Round(clsCommon.myCdbl(dblOutstandingAmount - AmountToCheckCustomerOutstandingForTCSTax), 2)
                                                End If
                                            End If
                                        End If
                                        If dblOutstandingAmount >= AmountToCheckCustomerOutstandingForTCSTax Then
                                            If EnableTCSRateValidityFrom01July2021 Then
                                                Dim Is_ITR_Filled_And_TCSAmountGreater50K As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue(" SELECT CASE WHEN ISNULL(IsTCSGreaterthan50K,0)=1 AND ISNULL(IsITRfilledinLast2Years,0)=1 THEN 1 ELSE 0 END FROM TSPL_CUSTOMER_MASTER WHERE Cust_Code='" & txtVendorNo.Value & "'")) = 1, True, False)
                                                If Is_ITR_Filled_And_TCSAmountGreater50K = True Then
                                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                                                Else
                                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                                                End If
                                            Else
                                                Dim panno As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(pan,'')+isnull(Additional3 ,'') as PanNoAdhar from tspl_customer_master where cust_code='" & txtVendorNo.Value & "'"))
                                                If clsCommon.myLen(panno) > 0 Then
                                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                                                Else
                                                    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                                                End If
                                            End If
                                        Else
                                            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = 0
                                        End If
                                    Else
                                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = dr("TaxRate")
                                    End If
                                End If
                            Else
                                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax_Rate" + strII)).Value = 0
                            End If
                        End If
                        ''end oftcs tax rate
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("RECOVERTABLETAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_Recoverable")), "Y") = CompairStringResult.Equal)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr(colIsTaxOnBaseAmount + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_On_Base_Amount")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                End If
            End If
            ' Next
        End If
    End Sub
    Sub SetitemWiseTaxOnlySetting()
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsscrap(txtTaxGroup.Value)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            Dim IsTaxable As Integer = 0
            For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                'IsTaxable = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IsTaxable from TSPL_ITEM_MASTER where item_code='" & clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colICode).Value) & "'"))
                IsTaxable = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select top 1 Is_Taxable from TSPL_ITEM_MASTER_TAXABLE where Item_Code='" & clsCommon.myCstr(gv1.Rows(intRowNo).Cells(colICode).Value) & "' and TSPL_ITEM_MASTER_TAXABLE.EFFECTIVE_DATE <= '" & clsCommon.GetPrintDate(txtDate.Value) & "' ORDER BY TSPL_ITEM_MASTER_TAXABLE.EFFECTIVE_DATE DESC "))
                If clsCommon.myLen(gv1.Rows(intRowNo).Cells(colICode).Value) > 0 AndAlso IsTaxable = 1 Then
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
    Private Sub txtTaxGroup__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTaxGroup._MYValidating
        Try
            If clsCommon.myLen(txtLocation.Value) > 0 Then
                Dim strCustomer As String = ""
                Try
                    strCustomer = txtVendorNo.Value
                Catch ex As Exception
                End Try
                If clsCommon.myLen(strCustomer) <= 0 Then
                    strCustomer = txtVendorNo.Value
                End If
                Dim qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
                txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroup(txtLocation.Value, strCustomer, "S", txtTaxGroup.Value, isButtonClicked)
                ' SetTaxDetails()
            Else
                Throw New Exception("Please select Location First")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtSchemeTaxGroup__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtSchemeTaxGroup._MYValidating
        Try
            If clsCommon.myLen(txtLocation.Value) > 0 Then
                Dim qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
                txtSchemeTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroup(txtLocation.Value, txtVendorNo.Value, "S", txtSchemeTaxGroup.Value, isButtonClicked)
                lblTaxGroupScheme.Text = clsTaxGroupMaster.GetNameOfSaleType(txtSchemeTaxGroup.Value, Nothing)
            Else
                Throw New Exception("Please select Location First")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtTermCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTermCode._MYValidating
        txtTermCode.Value = clsPaymentTerms.getFinder("", txtTermCode.Value, isButtonClicked)
        lblTermName.Text = clsDBFuncationality.getSingleValue("select isnull(Terms_Desc,'') from TSPL_TERMS_MASTER where Terms_Code='" & txtTermCode.Value & "'")
        SetTermDetails()
    End Sub
    Sub SetTermDetails()
        Dim qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_TERMS_MASTER where Terms_Code='" & txtTermCode.Value & "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Description"))
            txtDueDate.Value = txtDate.Value.AddDays(clsCommon.myCdbl(dt.Rows(0)("No Of Days")))
        Else
            lblTermName.Text = ""
        End If
    End Sub
    Private Sub gv2_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles gv2.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedTaxOpen Then
                    isCellValueChangedTaxOpen = True
                    If (e.Column Is (gv2.Columns(colTTaxAmt)) AndAlso rbtnTaxCalManual.IsChecked) Then
                        For ii As Integer = 0 To gv1.Rows.Count - 1
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                                UpdateCurrentRow1(ii)
                            Else
                                UpdateCurrentRow(ii)
                            End If
                        Next
                        UpdateAllTotals()
                    End If
                    isCellValueChangedTaxOpen = False
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtSubLocation__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtSubLocation._MYValidating
        Try
            If clsCommon.myLen(clsCommon.myCstr(txtLocation.Value)) > 0 Then
                txtSubLocation.Value = clsLocation.getFinder(" (isnull(is_sub_location,'N')='Y' or isnull(Is_Section,'N')='Y') and Main_Location_Code='" & txtLocation.Value & "'", txtSubLocation.Value, isButtonClicked)
                If clsCommon.myLen(txtSubLocation.Value) > 0 Then
                    lblSubLocation.Text = clsLocation.GetName(txtSubLocation.Value, Nothing)
                Else
                    lblSubLocation.Text = ""
                End If
            Else
                txtLocation.Value = ""
                lblSubLocation.Text = ""
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub GetBatchQty()
        Try
            Dim GetBatchFun As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_SCHEMA = 'dbo'   AND ROUTINE_NAME = 'fn_GetBatchQty'   AND ROUTINE_TYPE = 'FUNCTION'"))
            If GetBatchFun = 0 Then
                Dim qry = "CREATE or Alter FUNCTION dbo.fn_GetBatchQty(@item_code NVARCHAR(50),@Document_Code NVARCHAR(50))
                            RETURNS NVARCHAR(200)
AS
BEGIN
    DECLARE @result NVARCHAR(200)    
    SELECT @result =  ISNULL(CAST(Batch_No AS NVARCHAR(100)), '') + ',(' + ISNULL(CAST(Qty AS NVARCHAR(100)), '')+')'
    FROM TSPL_BATCH_ITEM
    WHERE Item_Code = @item_code and Document_Code=@Document_Code
    RETURN ISNULL(@result, '--')
END "
                clsDBFuncationality.ExecuteNonQuery(qry)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnPrintChallan_Click(sender As Object, e As EventArgs) Handles btnPrintChallan.Click
        Try
            Dim frmCRV As New frmCrystalReportViewer()


            Dim Qry As String = ""
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                GetBatchQty()
                Dim IsInvoiceCreated As Boolean = clsCommon.myCBool(IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(Document_Code) as NoofDoc from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'")) = 0, False, True))

                If IsInvoiceCreated Then
                    Qry = "select TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1+TSPL_COMPANY_MASTER.Add2+TSPL_COMPANY_MASTER.Add3 as Comp_Address,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.GSTReg_No,
TSPL_CUSTOMER_MASTER.Customer_Name,
TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No,TSPL_SD_SHIPMENT_HEAD.Supply_Date,TSPL_SD_SALE_INVOICE_HEAD.Inv_Date,
TSPL_SD_SHIPMENT_HEAD.VehicleNo,
TSPL_VENDOR_MASTER.Transporter,isnull(TSPL_VENDOR_MASTER.Phone1,isnull(TSPL_VENDOR_MASTER.Phone2,isnull(TSPL_VENDOR_MASTER.Contact_Person_Phone,''))) as Phone_No,
TSPL_BOOKING_DETAIL.route_no,TSPL_STATE_MASTER.STATE_NAME,
TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,
TSPL_BOOKING_DETAIL.Unit_code,TSPL_BOOKING_MATSER.document_no,TSPL_BOOKING_MATSER.Description,TSPL_VENDOR_MASTER.Vendor_Name as transporter_name,
TSPL_BOOKING_DETAIL.Booking_Qty as Qty,
TSPL_BOOKING_DETAIL.TAX1 AS ITAX1,TSPL_BOOKING_DETAIL.TAX1_RATE AS ITAX1_RATE,TSPL_BOOKING_DETAIL.TAX1_Amt AS ITAX1_Amt,TSPL_BOOKING_DETAIL.TAX1_Base_Amt AS ITAX1_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX2 AS ITAX2,TSPL_BOOKING_DETAIL.TAX2_RATE AS ITAX2_RATE,TSPL_BOOKING_DETAIL.TAX2_Amt AS ITAX2_Amt,TSPL_BOOKING_DETAIL.TAX2_Base_Amt AS ITAX2_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX3 AS ITAX3,TSPL_BOOKING_DETAIL.TAX3_Rate AS ITAX3_Rate,TSPL_BOOKING_DETAIL.TAX3_Amt AS ITAX3_Amt,TSPL_BOOKING_DETAIL.TAX3_Base_Amt AS ITAX3_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX4 AS ITAX4,TSPL_BOOKING_DETAIL.TAX4_RATE AS ITAX4_RATE,TSPL_BOOKING_DETAIL.TAX4_Amt AS ITAX4_Amt,TSPL_BOOKING_DETAIL.TAX4_Base_Amt AS ITAX4_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX5 AS ITAX5,TSPL_BOOKING_DETAIL.TAX5_RATE AS ITAX5_RATE,TSPL_BOOKING_DETAIL.TAX5_Amt AS ITAX5_Amt,TSPL_BOOKING_DETAIL.TAX5_Base_Amt AS ITAX5_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX6 AS ITAX6,TSPL_BOOKING_DETAIL.TAX6_RATE AS ITAX6_RATE,TSPL_BOOKING_DETAIL.TAX6_Amt AS ITAX6_Amt,TSPL_BOOKING_DETAIL.TAX6_Base_Amt AS ITAX6_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX7 AS ITAX7,TSPL_BOOKING_DETAIL.TAX7_Rate AS ITAX7_Rate,TSPL_BOOKING_DETAIL.TAX7_Amt AS ITAX7_Amt,TSPL_BOOKING_DETAIL.TAX7_Base_Amt AS ITAX7_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX8 AS ITAX8,TSPL_BOOKING_DETAIL.TAX8_RATE AS ITAX8_RATE,TSPL_BOOKING_DETAIL.TAX8_Amt AS ITAX8_Amt,TSPL_BOOKING_DETAIL.TAX8_Base_Amt AS ITAX8_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX9 AS ITAX9,TSPL_BOOKING_DETAIL.TAX9_Rate AS ITAX9_Rate,TSPL_BOOKING_DETAIL.TAX9_Amt AS ITAX9_Amt,TSPL_BOOKING_DETAIL.TAX9_Base_Amt AS ITAX9_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX10 AS ITAX10 ,TSPL_BOOKING_DETAIL.TAX10_RATE AS ITAX10_RATE,TSPL_BOOKING_DETAIL.TAX10_Amt AS ITAX10_Amt,TSPL_BOOKING_DETAIL.TAX10_Base_Amt AS ITAX10_Base_Amt,

dbo.fn_GetBatchQty(TSPL_BOOKING_DETAIL.item_code,TSPL_SD_SHIPMENT_HEAD.Document_Code) as Batch_No,TSPL_ITEM_UOM_DETAIL.Conversion_Factor as Conversion_Factor,
ITEMDETAIL1.Conversion_Factor as CF,
((TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/(ITEMDETAIL1.Conversion_Factor)/TSPL_BOOKING_DETAIL.Booking_Qty) as PackingQty,
((TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ITEMDETAIL1.Conversion_Factor) as TotalSaleQty,
TSPL_BOOKING_DETAIL.Tax_On_Amount as Amount_With_Tax

from TSPL_BOOKING_MATSER
left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
left join TSPL_CUSTOMER_MASTER on TSPL_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
left join TSPL_ITEM_MASTER on TSPL_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
left join TSPL_LOCATION_MASTER  on TSPL_BOOKING_MATSER.location_code=TSPL_LOCATION_MASTER.Location_Code
left outer join TSPL_STATE_MASTER   On TSPL_STATE_MASTER.STATE_CODE =TSPL_LOCATION_MASTER.state 
left join TSPL_COMPANY_MASTER on TSPL_BOOKING_MATSER.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code
left join TSPL_SD_SHIPMENT_HEAD on TSPL_BOOKING_DETAIL.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
left join TSPL_BATCH_ITEM on TSPL_BOOKING_DETAIL.Item_Code=TSPL_BATCH_ITEM.Item_Code and TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_BATCH_ITEM.Document_Code
left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code
left join TSPL_VENDOR_MASTER on TSPL_BOOKING_MATSER.Transport_Id=TSPL_VENDOR_MASTER.Vendor_Code

left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_BOOKING_DETAIL.Item_Code And   
TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_BOOKING_DETAIL.Unit_code 


left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code in (select UOM_Code  from TSPL_item_uom_detail where Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code and TSPL_item_uom_detail.Print_UOM=1)and TSPL_item_uom_detail.Print_UOM=1   ) as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_BOOKING_DETAIL.Item_Code 

where TSPL_BOOKING_MATSER.Document_No='" & txtDocNo.Value & "' and TSPL_SD_SHIPMENT_HEAD.Status=1"
                Else

                    Qry = "select TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1+TSPL_COMPANY_MASTER.Add2+TSPL_COMPANY_MASTER.Add3 as Comp_Address,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.GSTReg_No,
TSPL_CUSTOMER_MASTER.Customer_Name,
'To be Created..' as Sale_Invoice_No,'' as Supply_Date,'' as Inv_Date,
'' as VehicleNo,

TSPL_VENDOR_MASTER.Transporter,isnull(TSPL_VENDOR_MASTER.Phone1,isnull(TSPL_VENDOR_MASTER.Phone2,isnull(TSPL_VENDOR_MASTER.Contact_Person_Phone,''))) as Phone_No,
TSPL_BOOKING_DETAIL.route_no,TSPL_STATE_MASTER.STATE_NAME,
TSPL_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.HSN_Code,
TSPL_BOOKING_DETAIL.Unit_code,TSPL_BOOKING_MATSER.document_no,TSPL_BOOKING_MATSER.Description,TSPL_VENDOR_MASTER.Vendor_Name as transporter_name,
TSPL_BOOKING_DETAIL.Booking_Qty as Qty,ItemCF.Conversion_Factor as CurrCF,
TSPL_BOOKING_DETAIL.TAX1 AS ITAX1,TSPL_BOOKING_DETAIL.TAX1_RATE AS ITAX1_RATE,TSPL_BOOKING_DETAIL.TAX1_Amt AS ITAX1_Amt,TSPL_BOOKING_DETAIL.TAX1_Base_Amt AS ITAX1_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX2 AS ITAX2,TSPL_BOOKING_DETAIL.TAX2_RATE AS ITAX2_RATE,TSPL_BOOKING_DETAIL.TAX2_Amt AS ITAX2_Amt,TSPL_BOOKING_DETAIL.TAX2_Base_Amt AS ITAX2_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX3 AS ITAX3,TSPL_BOOKING_DETAIL.TAX3_Rate AS ITAX3_Rate,TSPL_BOOKING_DETAIL.TAX3_Amt AS ITAX3_Amt,TSPL_BOOKING_DETAIL.TAX3_Base_Amt AS ITAX3_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX4 AS ITAX4,TSPL_BOOKING_DETAIL.TAX4_RATE AS ITAX4_RATE,TSPL_BOOKING_DETAIL.TAX4_Amt AS ITAX4_Amt,TSPL_BOOKING_DETAIL.TAX4_Base_Amt AS ITAX4_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX5 AS ITAX5,TSPL_BOOKING_DETAIL.TAX5_RATE AS ITAX5_RATE,TSPL_BOOKING_DETAIL.TAX5_Amt AS ITAX5_Amt,TSPL_BOOKING_DETAIL.TAX5_Base_Amt AS ITAX5_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX6 AS ITAX6,TSPL_BOOKING_DETAIL.TAX6_RATE AS ITAX6_RATE,TSPL_BOOKING_DETAIL.TAX6_Amt AS ITAX6_Amt,TSPL_BOOKING_DETAIL.TAX6_Base_Amt AS ITAX6_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX7 AS ITAX7,TSPL_BOOKING_DETAIL.TAX7_Rate AS ITAX7_Rate,TSPL_BOOKING_DETAIL.TAX7_Amt AS ITAX7_Amt,TSPL_BOOKING_DETAIL.TAX7_Base_Amt AS ITAX7_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX8 AS ITAX8,TSPL_BOOKING_DETAIL.TAX8_RATE AS ITAX8_RATE,TSPL_BOOKING_DETAIL.TAX8_Amt AS ITAX8_Amt,TSPL_BOOKING_DETAIL.TAX8_Base_Amt AS ITAX8_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX9 AS ITAX9,TSPL_BOOKING_DETAIL.TAX9_Rate AS ITAX9_Rate,TSPL_BOOKING_DETAIL.TAX9_Amt AS ITAX9_Amt,TSPL_BOOKING_DETAIL.TAX9_Base_Amt AS ITAX9_Base_Amt,
                    TSPL_BOOKING_DETAIL.TAX10 AS ITAX10 ,TSPL_BOOKING_DETAIL.TAX10_RATE AS ITAX10_RATE,TSPL_BOOKING_DETAIL.TAX10_Amt AS ITAX10_Amt,TSPL_BOOKING_DETAIL.TAX10_Base_Amt AS ITAX10_Base_Amt,

dbo.fn_GetBatchQty(TSPL_BOOKING_DETAIL.item_code,TSPL_BOOKING_DETAIL.Document_No) as Batch_No,TSPL_ITEM_UOM_DETAIL.Conversion_Factor as Conversion_Factor,
ITEMDETAIL1.Conversion_Factor as CF,
((TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/(ITEMDETAIL1.Conversion_Factor)/TSPL_BOOKING_DETAIL.Booking_Qty) as PackingQty,
((TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ITEMDETAIL1.Conversion_Factor) as TotalSaleQty,
TSPL_BOOKING_DETAIL.Tax_On_Amount as Amount_With_Tax

from TSPL_BOOKING_MATSER
left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
left join TSPL_CUSTOMER_MASTER on TSPL_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
left join TSPL_ITEM_MASTER on TSPL_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
left join TSPL_LOCATION_MASTER  on TSPL_BOOKING_MATSER.location_code=TSPL_LOCATION_MASTER.Location_Code
left outer join TSPL_STATE_MASTER   On TSPL_STATE_MASTER.STATE_CODE =TSPL_LOCATION_MASTER.state 
left join TSPL_COMPANY_MASTER on TSPL_BOOKING_MATSER.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code

left join TSPL_VENDOR_MASTER on TSPL_BOOKING_MATSER.Transport_Id=TSPL_VENDOR_MASTER.Vendor_Code
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_Code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code ) as ItemCF on ItemCF.Item_code=TSPL_BOOKING_DETAIL.Item_Code and ItemCF.UOM_Code=TSPL_BOOKING_DETAIL.Unit_code
left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_BOOKING_DETAIL.Item_Code And   
TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_BOOKING_DETAIL.Unit_code 
--LEFT OUTER JOIN TSPL_ITEM_MASTER  ON  TSPL_ITEM_MASTER.Item_Code =TSPL_BOOKING_DETAIL.Item_Code
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code in (select UOM_Code  from TSPL_item_uom_detail where Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code and TSPL_item_uom_detail.Print_UOM=1)and TSPL_item_uom_detail.Print_UOM=1   ) as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_BOOKING_DETAIL.Item_Code  
where TSPL_BOOKING_MATSER.Document_No='" & txtDocNo.Value & "' "
                End If


            Else
                Qry = "Select * from (select 1 As CopyType,TSPL_BOOKING_MATSER.Is_BPL,TSPL_BOOKING_MATSER.BPL_Coupon_Code,TSPL_BOOKING_MATSER.BPL_Coupon_Date,TSPL_BOOKING_MATSER.BPL_Name,
                       TSPL_BOOKING_MATSER.BPL_Remark,TSPL_BOOKING_MATSER.BPL_Category,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    Qry += " TabBatch.BatchNo,TabBatch.Batch_Qty,'' as Comp_Phone2,'' as Comp_Add3,'' as Comp_Add2, "
                Else
                    Qry += " TSPL_COMPANY_MASTER.Phone2 As Comp_Phone2, TSPL_COMPANY_MASTER.Add3 As Comp_Add3,TSPL_COMPANY_MASTER.Add2 As Comp_Add2,
"
                End If
                Qry += " InLtr.Conversion_factor As [ConversionInLtr],InCrate.Conversion_factor As [ConversionInCrate],InPouch.Conversion_factor As [ConversionInPouch],TSPL_BOOKING_MATSER.Document_Date,
                        TSPL_BOOKING_MATSER.Is_Taxable,Case When TSPL_BOOKING_MATSER.GatePass_Type='AM' OR TSPL_BOOKING_MATSER.GatePass_Type='M' Then '[M]' Else '[E]' End As Shift,
                        TSPL_BOOKING_DETAIL.*,TSPL_COMPANY_MASTER.Access_Officer,
                        TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.HSN_Code,
                        case when coalesce(InLtr.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InLtr.Conversion_factor,1)) end as QtyInLtr, 
                        case when coalesce(InCrate.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InCrate.Conversion_factor,1)) end as QtyInCrate,
                        case when coalesce(InPouch.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  TSPL_BOOKING_DETAIL.Booking_Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(InPouch.Conversion_factor,1)) end as QtyInPouch,
                        TSPL_BOOKING_MATSER.FAT_Per,TSPL_BOOKING_MATSER.SNF_Per,TSPL_BOOKING_MATSER.Acidity,TSPL_BOOKING_MATSER.Temperature,TSPL_BOOKING_MATSER.MBRT_Hours, 
                        TSPL_Route_Master.Route_Desc,TSPL_VEHICLE_MASTER.Vehicle_Id,case when Manual_VehicleNo > ''  then  Manual_VehicleNo else TSPL_VEHICLE_MASTER.Number  end As Vehicle_Number,
                        TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 As Cust_Add1,TSPL_CUSTOMER_MASTER.Add2 As Cust_Add2,TSPL_CUSTOMER_MASTER.Add3 As Cust_Add3,TSPL_CUSTOMER_MASTER.PIN_Code As Cust_PINCode,
                        TSPL_CUSTOMER_MASTER.Phone1 As Cust_Phone1,TSPL_CUSTOMER_MASTER.Phone2 As Cust_Phone2,
                        TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img As Comp_Logo1,TSPL_COMPANY_MASTER.Logo_Img2 As Comp_Logo2,TSPL_COMPANY_MASTER.Add1 As Comp_Add1,
                       

TSPL_COMPANY_MASTER.City_Code As Comp_City,TSPL_COMPANY_MASTER.State As Comp_State,TSPL_COMPANY_MASTER.GSTReg_No As Comp_GSTReg_No,TSPL_COMPANY_MASTER.Pan_No As Comp_PanNo,
                        TSPL_COMPANY_MASTER.Email As Comp_Email,TSPL_COMPANY_MASTER.Pincode As Comp_Pincode,TSPL_COMPANY_MASTER.Phone1 As Comp_Phone1,




                        TSPL_STATE_MASTER.GST_STATE_Code As State_Code,TSPL_STATE_MASTER.STATE_NAME
                        from  TSPL_BOOKING_DETAIL
                        Left Outer Join TSPL_BOOKING_MATSER On TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
                        Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOKING_DETAIL.Cust_Code
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_BOOKING_MATSER.location_code 
                        Left Outer Join TSPL_Route_Master On TSPL_Route_Master.Route_No=TSPL_BOOKING_DETAIL.Route_No 
                        Left Outer Join TSPL_VEHICLE_MASTER On TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_BOOKING_DETAIL.Vehicle_Code 
                        left outer join TSPL_CITY_MASTER    On  TSPL_CITY_MASTER.City_Code =TSPL_LOCATION_MASTER.City_Code  
                        left outer join TSPL_STATE_MASTER   On TSPL_STATE_MASTER.STATE_CODE =TSPL_LOCATION_MASTER.state  
                        Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_BOOKING_DETAIL.Item_Code
                        left outer join TSPL_ITEM_UOM_DETAIL  on TSPL_ITEM_UOM_DETAIL.Item_Code =TSPL_ITEM_MASTER.Item_Code  and   TSPL_ITEM_UOM_DETAIL.UOM_Code =TSPL_BOOKING_DETAIL.Unit_code
                        left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.LTR_TYPE='Y') as InLtr on InLtr.Item_code=TSPL_ITEM_MASTER.Item_Code  
                        left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL 	  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.Crate_type='Y') as InCrate on InCrate.Item_code=TSPL_ITEM_MASTER.Item_Code  
                        left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.Packet_Type='Y' ) as InPouch on InPouch.Item_code=TSPL_ITEM_MASTER.Item_Code  "
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    Qry += " left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Against_Booking_No=TSPL_BOOKING_MATSER.Document_No

left outer join (select Document_Code, STRING_AGG(Batch_No,CHAR(10)) as BatchNo , STRING_AGG(CAST(Qty AS INT), CHAR(10)) as Batch_Qty
,Item_Code,UOM
                from(
SELECT TSPL_BATCH_ITEM.Document_Code, Batch_No, Qty, Parent_Line_No,Item_Code,UOM FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='FS-SH'
)x group by Document_Code,Parent_Line_No,Item_Code,UOM         
)TabBatch
On TabBatch.Document_Code= TSPL_SD_SHIPMENT_HEAD.Document_Code And  TabBatch.Item_Code=TSPL_BOOKING_DETAIL.Item_Code  and TabBatch.UOM=TSPL_BOOKING_DETAIL.Unit_code"
                End If

                Qry += "  
                        Left outer join TSPL_COMPANY_MASTER on  TSPL_COMPANY_MASTER.Comp_Code1 = '" & objCommonVar.CurrComp_Code1 & "'
                    where TSPL_BOOKING_MATSER.Document_No ='" & txtDocNo.Value & "')xxx
                        Left OUTER JOIN (Select 1 As COL1, 1 As COL2,  'ORIGINAL COPY' as CopyType1 UNION Select 1 as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XXX.CopyType ORDER BY Line_No,YYY.COL2 "
            End If

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt.Rows.Count > 0 Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "BatchWise_PrintChallan", "Challan", "", "rptCompanyAddress.rpt")
                Else

                    Dim IsTaxable As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Is_Taxable from TSPL_BOOKING_MATSER where TSPL_BOOKING_MATSER.Document_No='" & txtDocNo.Value & "'"))
                    If IsTaxable = 1 Then
                        'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
                        '    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptBookingNonTaxableChallanSWM", "Challan", "", "rptCompanyAddress.rpt")
                        'Else
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptBookingNonTaxableChallan", "Challan", "", "rptCompanyAddress.rpt")
                        'End If
                    ElseIf IsTaxable = 2 Then
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
                            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptBookingTaxableChallanSWM", "Challan", "", "rptCompanyAddress.rpt")
                        Else
                            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptBookingTaxableChallan", "Challan", "", "rptCompanyAddress.rpt")
                        End If
                    End If

                End If
                frmCRV = Nothing
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not Found to Print", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub ChkTaxNonTax()
        If rbtnTaxable.IsChecked Then
            rgbTaxNonTax.Visible = False
            If ServerDateTimeForTaxableInvoice Then
                txtDate.Value = clsCommon.GETSERVERDATE
                txtDate.Enabled = False
            Else
                txtDate.Enabled = True
            End If

        Else
            rgbTaxNonTax.Visible = True
            If ServerDateTimeForTaxableInvoice Then
                txtDate.Value = clsCommon.GETSERVERDATE
                txtDate.Enabled = False
            Else
                txtDate.Enabled = True
            End If

        End If
    End Sub
    Private Sub rbtnNonTax_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnNonTax.ToggleStateChanged
        Try
            ChkTaxNonTax()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnTaxable_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnTaxable.ToggleStateChanged
        Try
            ChkTaxNonTax()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub LoadDCSData(ByVal strVendorno As String, ByVal strDate As DateTime)
        Dim intRow As Integer = 0
        Try


            Dim VLC_Uploader As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select VLC_Code_VLC_Uploader from TSPL_VLC_master_Head  where VSP_Code='" & strVendorno & "'"))
            Dim strQry As String = "select * from TSPL_DCS_DEMAND_BOOKING_MASTER
left join TSPL_DCS_DEMAND_BOOKING_DETAIL on TSPL_DCS_DEMAND_BOOKING_MASTER.Document_No=TSPL_DCS_DEMAND_BOOKING_DETAIL.Document_No
where TSPL_DCS_DEMAND_BOOKING_DETAIL.VLC_Uploader='" & VLC_Uploader & "' and convert(date,TSPL_DCS_DEMAND_BOOKING_MASTER.Document_Date,103)='" & clsCommon.GetPrintDate(strDate) & "' and Posted=1 and Categories='G' and not exists(select 1 from TSPL_BOOKING_MATSER 
left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
where TSPL_BOOKING_MATSER.Against_DCSBooking_No=TSPL_DCS_DEMAND_BOOKING_MASTER.Document_No and TSPL_BOOKING_DETAIL.Cust_Code='" & strVendorno & "' and convert(date,TSPL_BOOKING_MATSER.Document_Date,103)='" & clsCommon.GetPrintDate(strDate) & "')"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then

                For Each dr As DataRow In dt.Rows
                    txtRouteNo.Value = clsCommon.myCstr(dr("Route_No"))
                    lblRouteDesc.Text = clsRouteMaster.GetName(txtRouteNo.Value, Nothing)
                    txtLocation.Value = clsCommon.myCstr(dr("Location"))
                    lblLocation.Text = clsLocation.GetName(txtLocation.Value, Nothing)
                    lblOutstandingDesc.Text = clsCommon.myCdbl(dr("OutStandingAmt"))
                    lblUnbilledMilkAmt.Text = clsCommon.myCdbl(dr("UnbilledMilkAmt"))
                    txtLastCollectionDate.Text = clsCommon.myCstr(dr("LastMilkDate"))
                    txtDCSDemandNo.Text = clsCommon.myCstr(dr("Document_No"))
                    cmbcashcredit.Text = clsCommon.myCstr(dr("CreditType"))
                    gv1.Rows(intRow).Cells(colLineNo).Value = intRow + 1
                    gv1.Rows(intRow).Cells(colICode).Value = clsCommon.myCstr(dr("Item_Code"))
                    gv1.Rows(intRow).Cells(colUnit).Value = clsCommon.myCstr(dr("Unit_code"))
                    gv1.Rows(intRow).Cells(colQty).Value = clsCommon.myCdbl(dr("Qty"))
                    ItemPrice(gv1.Rows(intRow).Cells(colICode).Value, gv1.Rows(intRow).Cells(colUnit).Value, clsCommon.myCdbl(gv1.Rows(intRow).Cells(colQty).Value), intRow, False)
                    SetTax(gv1.Rows(intRow).Cells(colICode).Value, intRow)
                    SetTaxDetails(gv1.Rows(intRow).Cells(colICode).Value, intRow)
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        'UpdateCurrentRow1(gv1.CurrentRow.Index)
                    Else
                        UpdateCurrentRow(intRow)
                    End If
                    intRow += 1
                    gv1.Rows.AddNew()
                Next
            Else
                strQry = "select top 1 TSPL_BOOKING_MATSER. Document_No from TSPL_BOOKING_MATSER 
left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
where  TSPL_BOOKING_DETAIL.Cust_Code='" & strVendorno & "' and convert(date,TSPL_BOOKING_MATSER.Document_Date,103)='" & clsCommon.GetPrintDate(strDate) & "'"
                Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strQry))
                If clsCommon.myLen(strDocNo) > 0 Then
                    LoadData(strDocNo, NavigatorType.Current)
                Else
                    Throw New Exception("Data Not Found!")
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub OpenBatchItem()
        Dim TransType_Str As String = ""
        Dim blnBatchqty As Boolean = False
        Dim isNewDocumentorExistingdoc As Boolean = True
        Dim isSubLocation As Boolean = False
        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsSubLocationWise from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'", Nothing)), "Y") = CompairStringResult.Equal Then
            isSubLocation = True
        Else
            isSubLocation = False
        End If

        If clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
            If RunBatchFifowisewithmodifyfunctionality = True Then
                If clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) > 0 Then
                    isNewDocumentorExistingdoc = False
                Else
                    isNewDocumentorExistingdoc = True
                End If
                If isNewDocumentorExistingdoc = True Then
                    TransType_Str = "FS"
                    TransType_Str = TransType_Str & "-SH"
                    For ii As Integer = 0 To gv1.Rows.Count - 1
                        If gv1.Rows(ii).Cells(colIsBatchItem).Value = True Then
                            Dim strBatchunion As String = ""
                            'If RunBatchFifowise = 1 Then
                            If ii > 0 Then
                                Dim strICodeOuter As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colICode).Value)
                                'For jj As Integer = 0 To ii - 1
                                'Dim strICodeInner As String = clsCommon.myCstr(gv1.Rows(jj).Cells(colICode).Value)
                                'If clsCommon.CompairString(strICodeOuter, strICodeInner) = CompairStringResult.Equal Then
                                Dim arr As List(Of clsBatchInventory) = Nothing
                                arr = TryCast(gv1.Rows(ii).Cells(colICode).Tag, List(Of clsBatchInventory))
                                If arr IsNot Nothing Then
                                    For Each obj As clsBatchInventory In arr
                                        Dim dblqty As Double = obj.Qty
                                        'If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(jj).Cells(colUnit).Value), clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value)) <> CompairStringResult.Equal Then
                                        dblqty = GetConvQuantity(strICodeOuter, clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value), obj.Qty)
                                        'End If
                                        strBatchunion += " union all select '" & clsCommon.myCstr(obj.Batch_No) & "' as Batch_No, " &
                                                    "'" & clsCommon.myCstr(obj.Manual_BatchNo) & "' as Manual_BatchNo,'O' as In_Out_Type, " &
                                                    "'" & clsCommon.myCstr(gv1.Rows(ii).Cells(colUnit).Value) & "' as OrgUOM," & dblqty & " as OrgQty,0 as OrgMRP, " &
                                                    "'" & clsCommon.GetPrintDate(obj.Expiry_Date, "dd/MMM/yyyy") & "' as Expiry_Date, " &
                                                    "'" & clsCommon.GetPrintDate(obj.Manufacture_Date, "dd/MMM/yyyy") & "' as Manufacture_Date, " &
                                                    "" & dblqty & " as Qty, 0 as MRP,'' as RALNo,'' as VehicleNo  "
                                    Next
                                End If
                                'End If
                                'Next
                            End If
                            gv1.CurrentRow = gv1.Rows(ii)
                            Dim frm As frmBatchItemOut = New frmBatchItemOut()
                            frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                            frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                            If isSubLocation Then
                                frm.strLocationCode = txtSubLocation.Value
                            Else
                                frm.strLocationCode = txtLocation.Value
                            End If

                            frm.strCurrDocNo = txtDocNo.Value
                            frm.strCurrDocType = TransType_Str
                            '"PS-SH"
                            frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) comment by balwinder on UDL on 02/12/2016
                            frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
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
                        TransType_Str = IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
                        TransType_Str = TransType_Str & "-SH"
                        Dim frm As frmBatchItemOut = New frmBatchItemOut()
                        frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                        frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                        If isSubLocation Then
                            frm.strLocationCode = txtSubLocation.Value
                        Else
                            frm.strLocationCode = txtLocation.Value
                        End If
                        frm.strCurrDocNo = txtDocNo.Value
                        frm.strCurrDocType = TransType_Str
                        '"PS-SH"
                        frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                        'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) comment by balwinder On UDL On 02/12/2016
                        frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
                        If checkstockmrpwise Then
                            frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
                        End If
                        frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                        frm.ShowDialog()
                        If Not frm.isCencelButtonClicked Then
                            gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                            gv1.CurrentRow.Cells(colBatchNo).Value = frm.arr(0).Batch_No
                            If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso btnPost.Enabled Then
                                Dim strQry As String = "delete TSPL_BATCH_ITEM  where Document_Code='" & txtDocNo.Value & "' and Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' and UOM='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) & "'"
                                clsDBFuncationality.ExecuteNonQuery(strQry)
                                'clsBatchInventory.SaveData(TransType_Str, txtDocNo.Value, txtDate.Value, "O", clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtLocation.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colLineNo).Value), 0, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), gv1.CurrentRow.Cells(colICode).Tag, Nothing)
                            End If
                        End If
                    End If
                End If
            ElseIf RunBatchFifowise = 0 Then
                ' w/o fifo start here
                If clsCommon.myCBool(gv1.CurrentRow.Cells(colIsBatchItem).Value) Then
                    TransType_Str = IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
                    TransType_Str = TransType_Str & "-SH"
                    Dim frm As frmBatchItemOut = New frmBatchItemOut()
                    frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value)
                    frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colIName).Value)
                    If isSubLocation Then
                        frm.strLocationCode = txtSubLocation.Value
                    Else
                        frm.strLocationCode = txtLocation.Value
                    End If
                    frm.strCurrDocNo = txtDocNo.Value
                    frm.strCurrDocType = TransType_Str
                    '"PS-SH"
                    frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                    'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) comment by balwinder On UDL On 02/12/2016
                    frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
                    If checkstockmrpwise Then
                        frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value)
                    End If
                    frm.arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    frm.ShowDialog()
                    If Not frm.isCencelButtonClicked Then
                        gv1.CurrentRow.Cells(colICode).Tag = frm.arr
                        gv1.CurrentRow.Cells(colBatchNo).Value = frm.arr(0).Batch_No
                        If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso btnPost.Enabled Then
                            Dim strQry As String = "delete TSPL_BATCH_ITEM  where Document_Code='" & txtDocNo.Value & "' and Item_Code='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value) & "' and UOM='" & clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value) & "'"
                            clsDBFuncationality.ExecuteNonQuery(strQry)
                            'clsBatchInventory.SaveData(TransType_Str, txtDocNo.Value, txtDate.Value, "O", clsCommon.myCstr(gv1.CurrentRow.Cells(colICode).Value), txtLocation.Value, clsCommon.myCstr(gv1.CurrentRow.Cells(colLineNo).Value), 0, clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value), gv1.CurrentRow.Cells(colICode).Tag, Nothing)
                        End If
                    End If
                End If
                ' w/o fifo ends here
            Else
                ' fifo start
                TransType_Str = IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
                TransType_Str = TransType_Str & "-SH"
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
                                frm.strLocationCode = txtLocation.Value
                            End If
                            frm.strCurrDocNo = txtDocNo.Value
                            frm.strCurrDocType = TransType_Str
                            '"PS-SH"
                            frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                            'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) comment by balwinder on UDL on 02/12/2016
                            frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
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
                'frm.OpenSerialList(0, "")
                'gv1.CurrentRow.Cells(colICode).Tag = frm.arr
            End If
        End If
    End Sub

    Public Sub OpenBatchItemIfFIFIOSettingON()
        If clsERPFuncationality.GetBatchWiseApplicableStatus(txtDate.Value) = True Then
            Dim arr As List(Of clsBatchInventory) = Nothing
            Dim strBatchunion As String = ""
            If clsCommon.myLen(gv1.CurrentRow.Cells(colICode).Value) > 0 Then
                arr = TryCast(gv1.CurrentRow.Cells(colICode).Tag, List(Of clsBatchInventory))
            End If
            If Not arr Is Nothing Then
                If arr.Count > 0 Then
                    For Each obj As clsBatchInventory In arr
                        strBatchunion += " Batch No - " & clsCommon.myCstr(obj.Batch_No) & "  Unit - " & clsCommon.myCstr(obj.UOM) & "        Qty - " & clsCommon.myCstr(obj.Qty) + Environment.NewLine
                    Next
                    clsCommon.MyMessageBoxShow(Me, strBatchunion, Me.Text)
                End If
            End If
        End If
    End Sub
    Public Sub ItemBatchFIFOWiseOnly()
        Dim TransType_Str As String = ""
        Dim blnBatchqty As Boolean = False
        Dim isNewDocumentorExistingdoc As Boolean = True
        Dim isSubLocation As Boolean = False
        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsSubLocationWise from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'", Nothing)), "Y") = CompairStringResult.Equal Then
            isSubLocation = True
        Else
            isSubLocation = False
        End If
        TransType_Str = IIf(rbtn_Fresh.IsChecked = True, "FS", "PS")
        TransType_Str = TransType_Str & "-SH"
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
                        frm.strLocationCode = txtLocation.Value
                    End If
                    frm.strCurrDocNo = txtDocNo.Value
                    frm.strCurrDocType = TransType_Str
                    '"PS-SH"
                    frm.strUOM = clsCommon.myCstr(gv1.CurrentRow.Cells(colUnit).Value)
                    'frm.dblMRP = clsCommon.myCdbl(gv1.CurrentRow.Cells(colMRP).Value) comment by balwinder on UDL on 02/12/2016
                    frm.dblqty = clsCommon.myCdbl(gv1.CurrentRow.Cells(colQty).Value)
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
    End Sub

    Public Shared Function GetConvQuantity(ByVal strItem As String, ByVal strCurrentUnit As String, ByVal strConvertedUnit As String, ByVal dblQty As Double) As Double
        Dim dblCurrentConvF As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strCurrentUnit & "'"))
        Dim dblConvQty As Double = 0
        If clsCommon.myLen(strConvertedUnit) > 0 Then
            Dim dblOrgConvF = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor from tspl_item_uom_detail where Item_Code='" & strItem & "' and UOM_Code='" & strConvertedUnit & "'"))
            If dblCurrentConvF > 0 Then
                dblConvQty = Math.Round(Math.Round((dblOrgConvF / dblCurrentConvF), 2) * dblQty, 2)
            End If
        End If
        Return dblConvQty
    End Function

    Private Sub chkcashsale_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkcashsale.ToggleStateChanged
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
            If chkcashsale.Checked Then
                lblPaymentType.Visible = True
                cmbPaymentType.Visible = True
                txtChequeNo.Visible = True
                lblChequeNo.Visible = True
            Else
                lblPaymentType.Visible = False
                cmbPaymentType.Visible = False
                txtChequeNo.Visible = False
                lblChequeNo.Visible = False

            End If
        Else
            lblPaymentType.Visible = False
            cmbPaymentType.Visible = False
            txtChequeNo.Visible = False
            lblChequeNo.Visible = False

        End If
    End Sub
    Private Sub chkManualVehicle_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkManualVehicle.ToggleStateChanged
        If chkManualVehicle.Checked Then
            txtManualVehicle.Enabled = True
        Else
            txtManualVehicle.Enabled = False
        End If
    End Sub
    Private Sub fndTransporter__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndTransporter._MYValidating
        Dim qry As String = "select Transport_Id as [Transport Id],Transporter_Name as [Transporter Name] from TSPL_TRANSPORT_MASTER"
        fndTransporter.Value = clsCommon.ShowSelectForm("RoutMastrCodFND", qry, "Transport Id", "", fndTransporter.Value, "", isButtonClicked)
        lblTransporter.Text = clsDBFuncationality.getSingleValue(" select Transporter_Name from TSPL_TRANSPORT_MASTER where Transport_Id='" & fndTransporter.Value & "'")
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

    Private Sub btnprinte_wayBill_Click(sender As Object, e As EventArgs) Handles btnprinte_wayBill.Click
        Try
            Dim strEWB As String = ""
            strEWB = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EWayBillNo from TSPL_SD_SALE_INVOICE_head  where Against_Shipment_No in ( select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "') "))
            If clsCommon.myLen(strEWB) > 0 Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(clsPSInvoiceHead.PrintEWayBill(strEWB, txtLocation.Value))
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rpte-waybill", "E-WayBill", clsCommon.GetPrintDate(txtDate.Value), "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    frmCRV = Nothing
                Else
                    Throw New Exception("No Data Found ")
                End If
            Else
                Throw New Exception("E-Way Bill not Found!")
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Function Create_EWB(ByVal trans As SqlTransaction) As Boolean
        Try
            Dim strInvoiceNO As String = ""
            strInvoiceNO = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SALE_INVOICE_head  where Against_Shipment_No in ( select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "') ", trans))
            If clsCommon.myLen(strInvoiceNO) > 0 Then
                Dim EwbNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select EWayBillNo from TSPL_SD_SALE_INVOICE_HEAD where  document_code = '" & strInvoiceNO & "' ", trans))
                If clsCommon.myLen(EwbNo) = 0 Then
                    If rbtnTaxable.IsChecked AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value, trans) Then
                        Dim EInvoiceCancelTimeValid As Int64 = 0
                        EInvoiceCancelTimeValid = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,EInvoice_Posting_Date,GETDATE()),0) as PostedHours from tspl_sd_sale_invoice_head where  document_code = '" & strInvoiceNO & "'", trans))
                        If EInvoiceCancelTimeValid >= 24 Then
                            Throw New Exception("E-Way Bill can not be Generated.It has been more than 24 hours.")
                        End If
                    End If
                    If clsCommon.myLen(GetEWayBillNo(strInvoiceNO, trans)) <= 0 Then
                        clsPSInvoiceHead.EWayBill_Implementation(strInvoiceNO, txtLocation.Value, trans, True)
                        clsDBFuncationality.ExecuteNonQuery("update TSPL_SD_SALE_INVOICE_HEAD set IsEwaybill=1 where Document_Code='" & strInvoiceNO & "'", trans)
                        clsDBFuncationality.ExecuteNonQuery("update TSPL_SD_SHIPMENT_HEAD set IsEwaybill=1 where Against_Booking_No='" & txtDocNo.Value & "'", trans)
                        clsDBFuncationality.ExecuteNonQuery("update TSPL_BOOKING_MATSER set IsEwaybill=1 where Document_No='" & txtDocNo.Value & "'", trans)

                        If clsCommon.myLen(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" & strInvoiceNO & "'", trans)) <= 0 Then
                            Throw New Exception("E-Way Bill For Sales Invoice No [" + strInvoiceNO + "] is not generated")
                        End If
                    End If
                Else
                    Throw New Exception("Ewb no [" & clsCommon.myCstr(EwbNo) & "] is already exists")
                End If
            Else
                Throw New Exception("Invoice Not Found!")
            End If


        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function GetEWayBillNo(strDocNo As String, trans As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  isnull(EWayBillNo,'') from TSPL_SD_SALE_INVOICE_head where Document_Code='" + strDocNo + "'", trans))
    End Function
    Private Sub btnHistory_Click_1(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocNo.Value, "Document_No", "TSPL_BOOKING_MATSER", "TSPL_BOOKING_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub chkAPS_CheckStateChanged(sender As Object, e As EventArgs) Handles chkAPS.CheckStateChanged
        If chkAPS.Checked Then
            chkBPL.Enabled = False
            chkDistributor.Checked = False
            chkDistributor.Enabled = False

        End If
    End Sub

    Private Sub btnCreateEWB_Click(sender As Object, e As EventArgs) Handles btnCreateEWB.Click
        Dim tran As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Create_EWB(tran)
            tran.Commit()
            clsCommon.MyMessageBoxShow(Me, "EWB Created Successfully", Me.Text)
            LoadData(txtDocNo.Value, NavigatorType.Current)
        Catch ex As Exception
            tran.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnUpdateVehicle_Click(sender As Object, e As EventArgs) Handles btnUpdateVehicle.Click
        Try
            Dim DocCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "'"))

            If clsCommon.myLen(DocCode) > 0 AndAlso UsLock1.Status = ERPTransactionStatus.Approved Then
                Dim frm As New FrmCommonUpdatesForEWB
                frm.strDocNo = txtDocNo.Value
                frm.strCustCode = txtVendorNo.Value
                frm.strTransId = fndTransporter.Value
                frm.strTransName = lblTransporter.Text
                frm.strVehicleNo = txtManualVehicle.Text
                frm.strScreenType = "CustomerBooking"
                frm.ShowDialog()
            Else
                Throw New Exception("Unable to proceed. You must select a Document Number that has already been created Disptach and approved.")
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkNoTranspoter_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkNoTranspoter.ToggleStateChanged
        If chkNoTranspoter.Checked Then
            fndTransporter.Enabled = False
            lblTransporter.Enabled = False
            fndTransporter.Value = ""
            lblTransporter.Text = ""
        Else
            fndTransporter.Enabled = True
            lblTransporter.Enabled = True

        End If
    End Sub

    'Private Sub txtMobileNo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMobileNo.KeyPress
    '    If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(8) Then
    '        e.Handled = True
    '    End If
    'End Sub

    'Private Sub txtMobileNo_TextChanged(sender As Object, e As EventArgs) Handles txtMobileNo.TextChanged
    '    If txtMobile.Text.Length > 10 Then
    '        txtMobile.Text = txtMobile.Text.Substring(0, 10)
    '        txtMobile.SelectionStart = txtMobile.Text.Length
    '    End If
    'End Sub




    'Private Function UpdateVehicleNo(ByVal trans As SqlTransaction) As Boolean
    '    Try
    '        Dim strInvoiceNO As String = ""
    '        strInvoiceNO = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SALE_INVOICE_head  where Against_Shipment_No in ( select Document_Code from TSPL_SD_SHIPMENT_HEAD where Against_Booking_No='" & txtDocNo.Value & "'  and Customer_Code='" & txtVendorNo.Value & "') ", trans))
    '        Dim strEwb As String = GetEWayBillNo(strInvoiceNO, Nothing)
    '        If clsCommon.myLen(strEwb) = 0 Then
    '            If clsCommon.MyMessageBoxShow(Me, "Do you want to update vehicle no?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
    '                Dim Reason As String = ""
    '                Dim frm As New FrmFreeTxtBox1
    '                frm.Text = "Enter Vehicle No."
    '                frm.ShowDialog()
    '                If clsCommon.myLen(frm.strRmks) <= 0 Then
    '                    Throw New Exception("Please enter Vehicle no.")
    '                Else
    '                    Reason = frm.strRmks
    '                    clsDBFuncationality.ExecuteNonQuery("update TSPL_SD_SALE_INVOICE_HEAD set VehicleNo='" & Reason & "' where Document_Code='" & strInvoiceNO & "'", trans)
    '                    clsDBFuncationality.ExecuteNonQuery("update TSPL_SD_SHIPMENT_HEAD set VehicleNo='" & Reason & "' where Against_Booking_No='" & txtDocNo.Value & "'", trans)
    '                    clsDBFuncationality.ExecuteNonQuery("update TSPL_BOOKING_MATSER set  Is_Manual_Vehicle='Y',Manual_VehicleNo='" & Reason & "' where Document_No='" & txtDocNo.Value & "'", trans)
    '                End If
    '            End If

    '        Else
    '            Throw New Exception("e-way bill already exists [" & strEwb & "]")
    '        End If
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    '    Return True
    'End Function
End Class
