''Created By------------------------------
''Pankaj Kumar Chaudhary----------------
''---on- 25/09/2011---------------------

''''07/06/2012---Updation by --[Pankaj kumar]-- Display Location, Amount, Sale Invoice no in ShipMent, Amount in Adhustment@MaterialManagement, Display Amount in Receipt, Display Amount in Payment
''''08/06/2012---Updation by --[Pankaj kumar]-- Shipment is Renmaed to (Shipment/Sale Invoice), Added a Label for COunting The No of Displayed Records, Added Location Filter, IN One Time Maxilum 100 records  Could be Posted
''''16/06/2012---Updation by --[Pankaj kumar]-- Posting Should be a Continuous Process If There is any Exception Than Jump that Document And Post Other And Generate an Error Record That Can Be saved AS Document
''''18/07/2012---Updation by --[Pankaj kumar]-- Added 3 new transactions [Empty Transactions, Production Entry, Store Adjustment] In Module Material Management With Viewing and Posting utilities----by--Ranjana mam
''''26/10/2012-12:27PM--Updation by --[Pankaj kumar]-- Show (BankCode, Location) In ReceiptEntry, Show (CustomerName, AMount, Location) in AdjustmentEntry@Receivable, Show (Location) In PaymentEntry In Payables-----From-Rakesh Sir
''''28/11/2012-10:31AM--Updation by --[Pankaj kumar]-- Added New Transaction [VCGL ENTRY] in Module [General Ledger]-----From-Manoj Sir
''''05/12/2012-12:57PM--Updation by --[Pankaj kumar]-- Added New Transaction [Empty Transactions] in Module [Material Management]-----From-Ranjana Sinha
''''06/12/2012-01:27PM--Updation by --[Pankaj kumar]-- Added New Transaction [Sale Return(Inter Company)] in Module [Sales N Distrobution---From-Ranjana Sinha
'--Updation By Pankaj Kumar Chaudhary Against Ticket [BM00000003922 - 29/09/2014]
''richa agarwal 12/05/2015 against ticket no.BM00000006520,BM00000008072
'====update by preet gupta Against ticket no[BHA/07/03/19-000833]
Imports common
Imports System.Data.SqlClient



Public Class FrmPendingAproval
    Inherits FrmMainTranScreen
    Dim trnsLstCustomer As New List(Of String)
    Dim strCustomerCode As String = Nothing
    Dim dt1 As DataTable = New DataTable()
    Dim Isrefreshed As Boolean = False    '' Variable for Validate the btnPost(Enable/Disable) and GridView
    Dim IsSelected As Boolean = False     '' Variable for Validate the btnSelectAll(ChangeText)
    Dim qry As String
    Dim dt As DataTable
    Dim count As Integer = 0
    Dim strNoOfRecord As String
    Dim trnsLst As New List(Of String)
    Dim arrUser As New ArrayList()
    Dim arrSelectedUser As New ArrayList()
    Dim strDocNo As String = Nothing
    Dim countPostedDoc As Integer = 0
    Public IsPostBack As Boolean = False
    Dim DtError As DataTable
    Dim dr As DataRow
    Public fromdate As DateTime
    Public Todate As DateTime
    Public ModuleName As String = ""
    Public Transaction As String = ""
    Public IsOpenPsted As Boolean
    Dim ButtonToolTip As New ToolTip()
    Dim isInsideLoad As Boolean = False
    '' List for Storing The Data(Which Is Selected) For Bulk Posting
    '==Sanjeet==========================
    Dim dtAuthen As DataTable
    Dim StrQuery As String = Nothing
    Dim ChkAllowBulkPosting As Double
    Dim arrLoc As String = Nothing
    Dim IsInsideLoadData As Boolean = True
    Dim ShowDairySaleModuleOnBulkPosting As Integer
    Dim RecordCount As Integer = 0
    Dim CreateProvisionOfTransporterInDairyDispatch As Boolean = False
    Dim FlagAllSelectWorking As Boolean = False
    Dim SettSeprateDemandForMorningEveningShift As Boolean = False
    Dim SettApplyMergeForDCSMultipleDays As Boolean = False
    Dim EnableProductSaleForJPR As Boolean = False


    Private Sub FrmPendingAproval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EnableProductSaleForJPR = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)) = 1, True, False)

        SettApplyMergeForDCSMultipleDays = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyMergeForDCSMultipleDays, clsFixedParameterCode.ApplyMergeForDCSMultipleDays, Nothing)) > 0)
        SettSeprateDemandForMorningEveningShift = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SeprateDemandForMorningEveningShift, clsFixedParameterCode.SeprateDemandForMorningEveningShift, Nothing)) = 1)
        ShowDairySaleModuleOnBulkPosting = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowDairySaleModuleOnBulkPosting, clsFixedParameterCode.ShowDairySaleModuleOnBulkPosting, Nothing))
        CreateProvisionOfTransporterInDairyDispatch = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateProvisionOfTransporterInDairyDispatch, clsFixedParameterCode.CreateProvisionOfTransporterInDairyDispatch, Nothing)) = 0, False, True)
        LOCATIONRIGTHS()
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C for Closing The Window")
        LoadLocation()
        chkLocAll.CheckState = CheckState.Checked
        ChkUserAll.CheckState = CheckState.Checked
        If clsCommon.CompairString(objCommonVar.CurrentUserCode, "ADMIN") <> CompairStringResult.Equal Then
            Dim dtUser As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_USER_MAPPING_DETAIL where User_Code ='" & objCommonVar.CurrentUserCode & "'")
            If dtUser IsNot Nothing AndAlso dtUser.Rows.Count > 0 Then
                ChkUserAll.CheckState = CheckState.Unchecked
                chkUserSelect.CheckState = CheckState.Checked
            End If
        End If

        'cbgLocation.CheckedAll()
        LoadUsers()
        '=========Sanjeet(05/01/2017)===================
        ChkAllowBulkPosting = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowBulkPostingofAllDocuments, clsFixedParameterCode.AllowBulkPostingofAllDocuments, Nothing))
        '==========================================
        arrUser = FrmUserMaster.GetSubbordinateUsers(objCommonVar.CurrentUserCode)
        SetUserMgmtNew()
        rbtnStatusPending.IsChecked = True
        btnPost.Enabled = False
        'If objCommonVar.RCDFCFP = True Then
        '    LoadModuleTypePO()
        '    LoadModuleNIRQC()
        'Else
        LoadModuleType()
            LaodModuleCommonServices()
        'End If
        LoadBlankGrid()
        'LoadModuleType()
        dtpFromDate.Value = clsCommon.GETSERVERDATE()
        dtpToDate.Value = clsCommon.GETSERVERDATE()
        'LaodModuleCommonServices()
        If clsCommon.myLen(ModuleName) > 0 Then
            cboModule.SelectedValue = ModuleName
            cboTransaction.SelectedValue = Transaction
            dtpFromDate.Value = fromdate
            dtpToDate.Value = Todate
            If IsOpenPsted = False Then
                rbtnStatusPending.IsChecked = True
            Else
                rbtnStatusPosted.IsChecked = True
            End If
            ShowData()
        End If
        ChkMilkType.Visible = False
        txtGrandTotal.Text = ""

    End Sub
    '===============================update by Preeti Gupta====================
    Private Sub LOCATIONRIGTHS()
        Dim obj As New clsMCCCodes()
        Try
            obj = clsMCCCodes.GetData()

            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Default_LocCode) > 0 Then
                Dim check As Integer = 0
                check = clsDBFuncationality.getSingleValue("select count(*) from tspl_location_master where type='Plant' and location_code='" + obj.Default_LocCode + "'")
                If check > 0 Then

                End If
            End If
            If obj.arrLocCodes IsNot Nothing AndAlso clsCommon.myLen(obj.arrLocCodes) > 0 Then
                arrLoc = obj.arrLocCodes
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            obj = Nothing
        End Try
    End Sub
    '============================================================================


    'Public Sub UserInformation()
    '    Dim qry As String
    '    Dim CurrentUser As String = objCommonVar.CurrentUserCode
    '    If (clsCommon.CompairString(clsCommon.myCstr(CurrentUser), "ADMIN") = CompairStringResult.Equal) Then
    '        qry = " select User_Code as Code, User_Name as Name from TSPL_USER_MASTER "
    '    Else
    '        'qry = " select User_Code as Code, User_Name as Name from TSPL_USER_MASTER where user_code='" + CurrentUser + "'"
    '        qry = " SELECT T1.Lvl1 as Code,T2.User_Name as Name FROM  ("
    '        qry += " SELECT '" + CurrentUser + "' AS 'Lvl1','" + CurrentUser + "' AS  'Lvl2',NULL AS  'Lvl3',NULL AS  'Lvl4',NULL AS  'Lvl5'  UNION ALL "
    '        qry += " SELECT "
    '        qry += " e1.User_Code AS 'Lvl1', "
    '        qry += " e2.User_Code AS 'Lvl2',"
    '        qry += " e3.User_Code AS 'Lvl3',"
    '        qry += " e4.User_Code AS 'Lvl4',"
    '        qry += " e5.User_Code AS 'Lvl5'"
    '        qry += " FROM TSPL_USER_MASTER AS e1"
    '        qry += " LEFT OUTER JOIN TSPL_USER_MASTER AS e2 ON e2.User_Code = e1.Level4_Code"
    '        qry += " LEFT OUTER JOIN TSPL_USER_MASTER AS e3 ON e3.User_Code = e2.Level4_Code"
    '        qry += " LEFT OUTER JOIN TSPL_USER_MASTER AS e4 ON e4.User_Code = e3.Level4_Code"
    '        qry += " LEFT OUTER JOIN TSPL_USER_MASTER AS e5 ON e5.User_Code = e4.Level4_Code) AS T1 inner join TSPL_USER_MASTER  T2 ON T1.Lvl1=T2.USER_CODE "
    '        qry += " WHERE t1.Lvl2='" + CurrentUser + "' or t1.Lvl3='" + CurrentUser + "' or t1.Lvl4='" + CurrentUser + "' or t1.Lvl5='" + CurrentUser + "'"
    '        'cbgUser.Enabled = False
    '    End If
    '    cbgUser.DataSource = clsDBFuncationality.GetDataTable(qry)
    '    cbgUser.ValueMember = "Code"
    '    cbgUser.DisplayMember = "Name"

    'End Sub
    Private Sub SetUserMgmtNew()
        '' Anubhooti 31-July-2014 BM00000003131
        'MyBase.SetUserMgmt(clsUserMgtCode.mbtnPendingApproval1)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
            'Me.Close()
            'Exit Function

        End If
        btnPost.Visible = MyBase.isPostFlag
        'btnsave.Visible = MyBase.isModifyFlag
        'btndelete.Visible = MyBase.isDeleteFlag
    End Sub
    ''
    ''Loads The Item In Combo Box (Module)
    ''
    Public Sub LoadModuleTypePO()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))

        dr = dt.NewRow()
        dr("Code") = "Purchase Order"
        dr("Name") = "Purchase Order"
        dt.Rows.Add(dr)

        cboModule.DataSource = dt
        cboModule.DisplayMember = "Name"
        cboModule.ValueMember = "Code"

    End Sub
    Public Sub LoadModuleType()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))

        dr = dt.NewRow()
        dr("Code") = "Common Services"
        dr("Name") = "Common Services"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Receivables"
        dr("Name") = "Receivables"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Payables"
        dr("Name") = "Payables"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Tax Deducted At Source"
        dr("Name") = "Tax Deducted At Source"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "General Ledger"
        dr("Name") = "General Ledger"
        dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "Sales And Distribution"
        'dr("Name") = "Sales And Distribution"
        'dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Product Sale"
        dr("Name") = "Product Sale"
        dt.Rows.Add(dr)


        If ShowDairySaleModuleOnBulkPosting = 1 Then
            dr = dt.NewRow()
            dr("Code") = "Dairy Sale"
            dr("Name") = "Dairy Sale"
            dt.Rows.Add(dr)
        Else
            dr = dt.NewRow()
            dr("Code") = "Fresh Sale"
            dr("Name") = "Fresh Sale"
            dt.Rows.Add(dr)
        End If

        dr = dt.NewRow()
        dr("Code") = "Material Management"
        dr("Name") = "Material Management"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Purchase Order"
        dr("Name") = "Purchase Order"
        dt.Rows.Add(dr)

        '' NEW ADDED MODULES 
        dr = dt.NewRow()
        dr("Code") = "HR and Payroll"
        dr("Name") = "HR and Payroll"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Production"
        dr("Name") = "Production"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Fixed Assets"
        dr("Name") = "Fixed Assets"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Bulk Sale"
        dr("Name") = "Bulk Sale"
        dt.Rows.Add(dr)
        '===Sanjeet(27/12/2016)=====
        dr = dt.NewRow()
        dr("Code") = "MCC Procurement"
        dr("Name") = "MCC Procurement"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Milk Procurement Bulk"
        dr("Name") = "Milk Procurement Bulk"
        dt.Rows.Add(dr)
        '===================================
        '===============Added by preeti Gupta[25/05/17]==========
        dr = dt.NewRow()
        dr("Code") = "Farmer Payment"
        dr("Name") = "Farmer Payment"
        dt.Rows.Add(dr)
        '===================================
        '========================================================

        dr = dt.NewRow()
        dr("Code") = "Dairy Production"
        dr("Name") = "Dairy Production"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Complaint"
        dr("Name") = "Complaint"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Sales And Distribution"
        dr("Name") = "Sales And Distribution"
        dt.Rows.Add(dr)


        cboModule.DataSource = dt
        cboModule.DisplayMember = "Name"
        cboModule.ValueMember = "Code"

    End Sub
    '================Added by Preeti Gupta=================
    Sub LoadModuleFarmerPayment()

        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))


        dr = dt1.NewRow()
        dr("Code") = "MCC Material Sale Farmer"
        dr("Name") = "MCC Material Sale Farmer"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "MCC Material Sale Farmer Return"
        dr("Name") = "MCC Material Sale Farmer Return"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Farmer Payment Adjustment"
        dr("Name") = "Farmer Payment Adjustment"
        dt1.Rows.Add(dr)




        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"

    End Sub

    Sub LoadModuleSalesAndDistribution()

        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))


        dr = dt1.NewRow()
        dr("Code") = "Shipment"
        dr("Name") = "Shipment"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Sale Invoice"
        dr("Name") = "Sale Invoice"
        dt1.Rows.Add(dr)

        If EnableProductSaleForJPR = True Then
            dr = dt1.NewRow()
            dr("Code") = "Product Dispatch"
            dr("Name") = "Product Dispatch"
            dt1.Rows.Add(dr)
        End If

        'dr = dt1.NewRow()
        'dr("Code") = "Product Demand Booking"
        'dr("Name") = "Product Demand Booking"
        'dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"

    End Sub

    '======================================================
    ''
    ''Loads The Transaction of Module(Purchase Order) In Combo Box (Transaction)
    ''
    Sub LoadModulePO()

        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))


        dr = dt1.NewRow()
        dr("Code") = "Purchase Requisition"
        dr("Name") = "Purchase Requisition"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Purchase Order"
        dr("Name") = "Purchase Order"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Gate Receipt Note"
        dr("Name") = "Gate Receipt Note"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Material Receipt Note"
        dr("Name") = "Material Receipt Note"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Store Receipt Note"
        dr("Name") = "Store Receipt Note"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Purchase Invoice"
        dr("Name") = "Purchase Invoice"
        dt1.Rows.Add(dr)
        ''''Added By ---Pankaj Kumar Chaudhary ----on 12/10/2011
        dr = dt1.NewRow()
        dr("Code") = "Purchase Return"
        dr("Name") = "Purchase Return"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "RGP/NRGP"
        dr("Name") = "RGP/NRGP"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Issue/Return/Transfer"
        dr("Name") = "Issue/Return/Transfer"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Scrap LoadOut"
        dr("Name") = "Scrap LoadOut"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Scrap Invoice"
        dr("Name") = "Scrap Invoice"
        dt1.Rows.Add(dr)



        dr = dt1.NewRow()
        dr("Code") = "NIR QC"
        dr("Name") = "NIR QC"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"

    End Sub

    Sub LoadModuleNIRQC()

        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "NIR QC"
        dr("Name") = "NIR QC"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    ''
    ''Loads The Transactions of Module(Material Management) In Combo Box (Transaction)
    ''
    Sub LaodModuleMaterialMgmt()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))


        dr = dt1.NewRow()
        dr("Code") = "Transfer(Load-In)"
        dr("Name") = "Transfer(Load-In)"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Transfer(Load-Out)"
        dr("Name") = "Transfer(Load-Out)"
        dt1.Rows.Add(dr)


        'dr = dt1.NewRow()
        'dr("Code") = "Adjustment Entry"
        'dr("Name") = "Adjustment Entry"
        'dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Empty Transactions"
        dr("Name") = "Empty Transactions"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Entry"
        dr("Name") = "Production Entry"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Store Adjustment"
        dr("Name") = "Store Adjustment"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    ''
    ''Loads The Transactions of Module(Sales And Distribution) In Combo Box (Transaction)
    ''

    Sub LaodModuleSalesNDistribution()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Sale Order"
        dr("Name") = "Sale Order"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Shipment"
        dr("Name") = "Shipment"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Sale Invoice"
        dr("Name") = "Sale Invoice"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Sale Return"
        dr("Name") = "Sale Return"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Sale Return (Inter Company)"
        dr("Name") = "Sale Return (Inter Company)"
        dt1.Rows.Add(dr)

        'dr = dt1.NewRow()
        'dr("Code") = "Quick Settlement"
        'dr("Name") = "Quick Settlement"
        'dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    Sub LaodModuleProductSale()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Shipment/Invoice"
        dr("Name") = "Shipment/Invoice"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    Sub LaodModuleFreshSale()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Shipment/Invoice"
        dr("Name") = "Shipment/Invoice"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    Sub LaodModuleDairySale()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))
        If SettSeprateDemandForMorningEveningShift Then
            dr = dt1.NewRow()
            dr("Code") = "Demand"
            dr("Name") = "Demand"
            dt1.Rows.Add(dr)
        End If

        dr = dt1.NewRow()
        dr("Code") = "Booking/DO"
        dr("Name") = "Booking/DO"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Fresh Booking/DO"
        dr("Name") = "Fresh Booking/DO"
        dt1.Rows.Add(dr)


        dr = dt1.NewRow()
        dr("Code") = "Shipment/Invoice"
        dr("Name") = "Shipment/Invoice"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Crate Received"
        dr("Name") = "Crate Received"
        dt1.Rows.Add(dr)

        ''richa BHA/27/12/18-000763 28 Dec,2018
        If CreateProvisionOfTransporterInDairyDispatch = True Then
            dr = dt1.NewRow()
            dr("Code") = "Dairy GatePass"
            dr("Name") = "Dairy GatePass"
            dt1.Rows.Add(dr)
        End If

        dr = dt1.NewRow()
        dr("Code") = "Product Demand Booking"
        dr("Name") = "Product Demand Booking"
        dt1.Rows.Add(dr)
        ''-----------------------

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    ''
    ''Loads The Transaction of Module(Genertal Ledger) In Combo Box (Transaction)
    ''

    Sub LaodModuleGeneralLedger()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Journal Entry"
        dr("Name") = "Journal Entry"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "VCGL Entry"
        dr("Name") = "VCGL Entry"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    ''
    ''Loads The Transaction of Module(Tax Deducted At Source) In Combo Box (Transaction)
    ''

    Sub LaodModuleTaxDeductedAtSource()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        'dr("Code") = "Create Remittance"
        'dr("Name") = "Create Remittance"
        'dt1.Rows.Add(dr)

        'dr = dt1.NewRow()
        dr("Code") = "Remittance Entry"
        dr("Name") = "Remittance Entry"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    ''
    ''Loads The Transaction of Module(Payables) In Combo Box (Transaction)
    ''

    Sub LaodModulePayables()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))


        dr = dt1.NewRow()
        dr("Code") = "Payment Entry"
        dr("Name") = "Payment Entry"
        dt1.Rows.Add(dr)


        dr = dt1.NewRow()
        dr("Code") = "AP Invoice Entry"
        dr("Name") = "AP Invoice Entry"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    ''
    ''Loads The Transaction of Module(Receivables) In Combo Box (Transaction)
    ''

    Sub LaodModuleReceivables()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Receipt Entry"
        dr("Name") = "Receipt Entry"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Adjustment Entry"
        dr("Name") = "Adjustment Entry"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "AR Invoice Entry"
        dr("Name") = "AR Invoice Entry"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    ''
    ''Loads The Transaction of Module(Commmon Services) In Combo Box (Transaction)
    ''

    Sub LaodModuleCommonServices()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Bank Transfer"
        dr("Name") = "Bank Transfer"
        dt1.Rows.Add(dr)


        dr = dt1.NewRow()
        dr("Code") = "Reverse Transaction"
        dr("Name") = "Reverse Transaction"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    Sub LaodModuleHRAndPayroll()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Employee Salary"
        dr("Name") = "Employee Salary"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Hourly Attendance"
        dr("Name") = "Hourly Attendance"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Daily Attendance"
        dr("Name") = "Daily Attendance"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Monthly Attendance"
        dr("Name") = "Monthly Attendance"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Allowance"
        dr("Name") = "Allowance"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Deduction"
        dr("Name") = "Deduction"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Bonus"
        dr("Name") = "Bonus"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Salary Adjustment"
        dr("Name") = "Salary Adjustment"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Reimbursement"
        dr("Name") = "Reimbursement"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Loan Application"
        dr("Name") = "Loan Application"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Loan Adjustment"
        dr("Name") = "Loan Adjustment"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Leave Application"
        dr("Name") = "Leave Application"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Leave Adjustment"
        dr("Name") = "Leave Adjustment"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Employee Increment"
        dr("Name") = "Employee Increment"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "OT Sheet"
        dr("Name") = "OT Sheet"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    Sub LaodModuleProduction()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Batch Order"
        dr("Name") = "Batch Order"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Bill Of Material"
        dr("Name") = "Bill Of Material"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Manufacturing Order"
        dr("Name") = "Manufacturing Order"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Planning"
        dr("Name") = "Production Planning"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Receipt"
        dr("Name") = "Production Receipt"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Entry"
        dr("Name") = "Production Entry"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Requisition"
        dr("Name") = "Production Requisition"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Return"
        dr("Name") = "Production Return"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Store Issue"
        dr("Name") = "Store Issue"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Silo Milk Transfer"
        dr("Name") = "Silo Milk Transfer"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    '======Sanjeet(27/12/2016)===============
    Sub LoadMccProcurement()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "VSP Item Issue"
        dr("Name") = "VSP Item Issue"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "VSP Asset Issue"
        dr("Name") = "VSP Asset Issue"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "MCC Material Sale"
        dr("Name") = "MCC Material Sale"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "MCC Material Sale Return"
        dr("Name") = "MCC Material Sale Return"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Procurement Deduction"
        dr("Name") = "Procurement Deduction"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Multiple Deduction"
        dr("Name") = "Multiple Deduction"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Acknowledgement Entry"
        dr("Name") = "Acknowledgement Entry"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Milk Shift Uploader"
        dr("Name") = "Milk Shift Uploader"
        dt1.Rows.Add(dr)
        dr = dt1.NewRow()
        dr("Code") = "Milk Procurement Uploader"
        dr("Name") = "Milk Procurement Uploader"
        dt1.Rows.Add(dr)
        dr = dt1.NewRow()
        dr("Code") = "MCC Milk Collection"
        dr("Name") = "MCC Milk Collection"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "DCS Milk Collection"
        dr("Name") = "DCS Milk Collection"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "BMC Mobile Entry"
        dr("Name") = "BMC Mobile Entry"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub

    Sub LoadMilkProcurementBulk()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Purchase Invoice"
        dr("Name") = "Purchase Invoice"
        dt1.Rows.Add(dr)

        If Not SettApplyMergeForDCSMultipleDays Then
            dr = dt1.NewRow()
            dr("Code") = "DCS Milk Collection Multiple Days"
            dr("Name") = "DCS Milk Collection Multiple Days"
            dt1.Rows.Add(dr)
        End If


        'dr = dt1.NewRow()
        'dr("Code") = "MCC Milk Collection Multiple Days"
        'dr("Name") = "MCC Milk Collection Multiple Days"
        'dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    Sub LoadModuleComplaint()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Customer Complaint"
        dr("Name") = "Customer Complaint"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    '=============================
    ''richa agarwal 12/05/2015 against ticket no.BM00000006520
    Sub LaodModuleBulkSale()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Can Sale"
        dr("Name") = "Can Sale"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Bulk Dispatch"
        dr("Name") = "Bulk Dispatch"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    ''---------------------------------
    Private Sub cboModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs) Handles cboModule.SelectedIndexChanged

        If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Sale") = CompairStringResult.Equal Then
            'If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Demand Booking") = CompairStringResult.Equal Then
            '    lblCustomer.Visible = False
            '    txtCustomer.Visible = False
            'Else
            lblCustomer.Visible = True
            txtCustomer.Visible = True
            'End If
            'lblCustomer.Visible = True
            'txtCustomer.Visible = True
        Else
            lblCustomer.Visible = False
            txtCustomer.Visible = False
        End If


        LoadTrnsListOfSelectedModeule()
    End Sub

    Public Sub LoadTrnsListOfSelectedModeule()
        'If objCommonVar.RCDFCFP = True Then
        If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Purchase Order") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            'LoadModuleNIRQC()
            'Else
            LoadModulePO()
            'End If

            'btnPost.Enabled = False
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Material Management") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleMaterialMgmt()
            'btnPost.Enabled = False
            'ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Sales And Distribution") = CompairStringResult.Equal Then
            '    cboTransaction.DataSource = Nothing
            '    LaodModuleSalesNDistribution()
            '    'btnPost.Enabled = False
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "General Ledger") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleGeneralLedger()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Product Sale") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleProductSale()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fresh Sale") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleFreshSale()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Sale") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleDairySale()
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
                lblCustomer.Text = "Route"
                gbStatus.Text = "Shift "
                rbtnStatusPending.Text = "Morning"
                rbtnStatusPosted.Text = "Evening"
                RadLabel3.Visible = False
                dtpToDate.Visible = False

            Else
                lblCustomer.Text = "Customer"
                gbStatus.Text = "Status"
                rbtnStatusPending.Text = "Pending"
                rbtnStatusPosted.Text = "Posted"
                RadLabel3.Visible = True
                dtpToDate.Visible = True
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Tax Deducted At Source") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleTaxDeductedAtSource()
            'btnPost.Enabled = False
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Payables") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModulePayables()
            'btnPost.Enabled = False
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Receivables") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleReceivables()
            'btnPost.Enabled = False
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Common Services") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleCommonServices()
            'btnPost.Enabled = False
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "HR and Payroll") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleHRAndPayroll()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Production") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleProduction()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Bulk Sale") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleBulkSale()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "MCC Procurement") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LoadMccProcurement()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Milk Procurement Bulk") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LoadMilkProcurementBulk()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Farmer Payment") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LoadModuleFarmerPayment()

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fixed Assets") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodModuleFixedAsset()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Production") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LaodDairyProduction()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Complaint") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LoadModuleComplaint()
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Sales And Distribution") = CompairStringResult.Equal Then
            cboTransaction.DataSource = Nothing
            LoadModuleSalesAndDistribution()
        End If

    End Sub

    Sub LaodDairyProduction()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Production Issue Entry"
        dr("Name") = "Production Issue Entry"
        dt1.Rows.Add(dr)


        dr = dt1.NewRow()
        dr("Code") = "Production Standardization"
        dr("Name") = "Production Standardization"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Production Standardization Final QC"
        dr("Name") = "Production Standardization Final QC"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Stage Process"
        dr("Name") = "Stage Process"
        dt1.Rows.Add(dr)


        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    Sub LoadBlankGrid()
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        gv1.AllowAddNewRow = False

    End Sub

    ''
    ''Format The GridView
    ''
    Sub gv1Format()
        Me.gv1.MasterTemplate.Columns("Status").Width = 50      ''First Column
        Me.gv1.MasterTemplate.Columns("Document Id").Width = 150    ''Second Column
        Dim count As Integer = gv1.MasterTemplate.Columns.Count
        For i As Integer = 2 To count - 2
            Me.gv1.MasterTemplate.Columns(i).Width = 120
        Next i
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "NIR QC") = CompairStringResult.Equal Then
            Me.gv1.MasterTemplate.Columns("Amount").IsVisible = False    ''Third Column

        End If
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "BMC Mobile Entry") = CompairStringResult.Equal Then
            Me.gv1.MasterTemplate.Columns("Date").Width = 200    ''Third Column
            Me.gv1.MasterTemplate.Columns("BMC Code").Width = 200    ''Fourth Column
        Else
            If Not (clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Standardization") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Standardization Final QC") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Stage Process") = CompairStringResult.Equal) Then
                Me.gv1.MasterTemplate.Columns("Description").Width = 200    ''Last Column
            End If
        End If

        For j As Integer = 1 To count - 1
            Me.gv1.MasterTemplate.Columns(j).ReadOnly = True
        Next
        Dim summaryRowItem As New GridViewSummaryRowItem()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VCGL Entry") = CompairStringResult.Equal Then
            Dim item2 As New GridViewSummaryItem("Debit Amount", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item2)
            Dim item3 As New GridViewSummaryItem("Credit Amount", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item3)
            Dim item4 As New GridViewSummaryItem("Amount", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item4)
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bulk Dispatch") = CompairStringResult.Equal Then
            Dim item2 As New GridViewSummaryItem("Amount", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item2)
            Dim item3 As New GridViewSummaryItem("Tare Weight", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item3)
            Dim item4 As New GridViewSummaryItem("Gross Weight", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item4)
            Dim item5 As New GridViewSummaryItem("Net Weight", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item5)
        Else
            'Dim item1 As New GridViewSummaryItem("Amount", "{0:F0}", GridAggregateFunction.Sum)
            'summaryRowItem.Add(item1)
        End If

        gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Private Sub gv1_ValueChanging(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.ValueChangingEventArgs) Handles gv1.ValueChanging
        Try
            If FlagAllSelectWorking = False Then
                TotalAmount(e.NewValue, gv1)
            End If
        Catch ex As Exception

        End Try
    End Sub
    ''richa ERO/16/04/19-000559
    Private Sub TotalAmount(ByVal Xselect As Boolean, ByVal gv1 As RadGridView)
        Dim Status As Double = 0
        Dim total As Double = 0
        For Each grow As GridViewRowInfo In gv1.Rows
            If clsCommon.myCBool(grow.Cells("Status").Value) = True AndAlso ((grow.Index <> gv1.CurrentRow.Index) Or grow.Index <= -1) Then
                If gv1.Columns.IndexOf("Amount") >= 0 Then
                    total += clsCommon.myCdbl(grow.Cells("Amount").Value)
                End If
            End If
        Next
        'Ticket No- GKD/11/09/18-000157,Check Amount column exist or not
        If gv1.Columns.IndexOf("Amount") >= 0 Then
            Status = gv1.CurrentRow.Cells("Amount").Value
        End If
        If Xselect = True Then
            total += Status
        End If
        txtGrandTotal.Text = total
    End Sub

#Region "Showing Details on GRID"
    'done by stuti on 18/10/2016 against ticket no - BM00000010089
    ''-------------------------------------------------------------------
    '' Function For Filling --------Module(Purchase Order)---------------
    ''-------------------------------------------------------------------

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
            dtpToDate.Value = dtpFromDate.Value
        End If
        If clsCommon.myCDate(dtpFromDate.Value, "dd/MMM/yyyy") > clsCommon.myCDate(dtpToDate.Value, "dd/MMM/yyyy") Then
            common.clsCommon.MyMessageBoxShow(Me, "'From date' Cann't Be Greater Than 'To Date'", Me.Text)
        Else
            qry = Nothing
            ShowData()
            txtGrandTotal.Text = ""
        End If


    End Sub

    Sub ShowData()
        Try
            If cbgUser.CheckedValue.Count > 0 Then
                arrSelectedUser = cbgUser.CheckedValue
            Else
                arrSelectedUser = arrUser
            End If

            If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Purchase Order") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillPurchaseOrder()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Product Sale") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillSalesNDistribution()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Sales And Distribution") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillSalesAndDistributionNew()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fresh Sale") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillFreshSale()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Sale") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillDairySale()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Material Management") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillMaterialManagement()
                'ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Sales And Distribution") = CompairStringResult.Equal Then
                '    gv1.DataSource = Nothing
                '    FillSalesNDistribution()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "General Ledger") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillGeneralLedger()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Tax Deducted At Source") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillTaxDdctdAtSource()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Payables") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillPayables()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Receivables") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillReceivables()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Common Services") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillCommonServices()
                If gv1.DataSource IsNot Nothing Then
                    gv1.MasterTemplate.Columns("Description").IsVisible = False
                End If
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "HR and Payroll") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillHRAndPayroll()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Production") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillProduction()
                ''richa agarwal 12/05/2015 against ticket no.BM00000006520
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Bulk Sale") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillBulkSale()
                'gv1.MasterTemplate.Columns("Description").IsVisible = False
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "MCC Procurement") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillMccProcurement()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Milk Procurement Bulk") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillMilkProcurementBulk()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Farmer Payment") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillMCCMaterialSaleFarmer()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fixed Assets") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillFixeedAsset()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Production") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillDairyProduction()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Complaint") = CompairStringResult.Equal Then
                gv1.DataSource = Nothing
                FillComplaint()
            End If
            btnSlctAll.Text = "Select TOP 100"
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                gv1.DataSource = Nothing
                If rbtnStatusPending.IsChecked = True Then
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow(Me, "There Is No '" + rbtnStatusPending.Text + "' Data For Date '" + dtpFromDate.Value.Date + "'", caption:="Morning")

                    Else
                        common.clsCommon.MyMessageBoxShow(Me, "There Is No '" + rbtnStatusPending.Text + "' Data Between The Dates '" + dtpFromDate.Value.Date + "' And '" + dtpToDate.Value.Date + "' ", caption:="Pending")

                    End If
                Else
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
                        common.clsCommon.MyMessageBoxShow(Me, "There Is No '" + rbtnStatusPosted.Text + "' Data For Date '" + dtpFromDate.Value.Date + "' ", caption:="Evening")

                    Else
                        common.clsCommon.MyMessageBoxShow(Me, "There Is No '" + rbtnStatusPosted.Text + "' Data Between The Dates '" + dtpFromDate.Value.Date + "' And '" + dtpToDate.Value.Date + "' ", caption:="Posted")

                    End If

                End If
                Return
            End If
            IsPostBack = False

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub FillDairyProduction()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Issue Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProcessProductionIssueEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST((TSPL_PP_ISSUE_HEAD.is_Post)as BIT) as Status,TSPL_PP_ISSUE_HEAD.Issue_Code as [Document Id],convert(varchar,TSPL_PP_ISSUE_HEAD.Issue_Date ,103) as [Document Date],isnull(TSPL_PP_ISSUE_HEAD.Batch_Code,0) as [Batch Code],Main_Location_Code as [Location],case when TSPL_ITEM_MASTER.Item_Type='F' then 'Finished Goods' when TSPL_ITEM_MASTER.Item_Type='S' then 'Semi Finished Goods' else '' end [Item Type] ,TSPL_PP_ISSUE_HEAD.Created_By as [Created By],convert(varchar,TSPL_PP_ISSUE_HEAD.Created_Date,103) as [Created Date],TSPL_PP_ISSUE_HEAD.Description from TSPL_PP_ISSUE_HEAD  " &
                    " left outer join TSPL_PP_BATCH_ORDER_BOM_DETAIL on TSPL_PP_BATCH_ORDER_BOM_DETAIL.Batch_Code =TSPL_PP_ISSUE_HEAD.Batch_Code " &
                    " left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_code=TSPL_PP_BATCH_ORDER_BOM_DETAIL.item_Code " &
                    " WHERE convert(date,TSPL_PP_ISSUE_HEAD.Issue_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PP_ISSUE_HEAD.Issue_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PP_ISSUE_HEAD.is_Post = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_PP_ISSUE_HEAD.is_Post = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PP_ISSUE_HEAD.Main_Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_PP_ISSUE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PP_ISSUE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_PP_ISSUE_HEAD.Issue_Date, TSPL_PP_ISSUE_HEAD.Issue_Code "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Standardization") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProcessProductionStandardization)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST((TSPL_PP_STANDARDIZATION_HEAD.Posted)as BIT) as Status,TSPL_PP_STANDARDIZATION_HEAD.Standardization_Code as [Document Id],convert(varchar,TSPL_PP_STANDARDIZATION_HEAD.Standardization_Date ,103) as [Document Date],isnull(TSPL_PP_STANDARDIZATION_HEAD.Main_Batch_Code,0) as [Batch Code],Loaction_Code as [Location],TSPL_PP_STANDARDIZATION_HEAD.Created_By as [Created By],convert(varchar,TSPL_PP_STANDARDIZATION_HEAD.Created_Date,103) as [Created Date] from TSPL_PP_STANDARDIZATION_HEAD " &
                    " WHERE  convert(date,TSPL_PP_STANDARDIZATION_HEAD.Standardization_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PP_STANDARDIZATION_HEAD.Standardization_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PP_STANDARDIZATION_HEAD.Posted = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_PP_STANDARDIZATION_HEAD.Posted = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PP_STANDARDIZATION_HEAD.Loaction_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_PP_STANDARDIZATION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PP_STANDARDIZATION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_PP_STANDARDIZATION_HEAD.Standardization_Date, TSPL_PP_STANDARDIZATION_HEAD.Standardization_Code "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Standardization Final QC") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.ProcessProductionStandardizationFinalQC)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST((TSPL_PP_STD_FINALQC_HEAD.Posted)as BIT) as Status,TSPL_PP_STD_FINALQC_HEAD.QC_Code as [Document Id],convert(varchar,TSPL_PP_STD_FINALQC_HEAD.QC_Date ,103) as [Document Date],isnull(TSPL_PP_STD_FINALQC_HEAD.Main_Batch_Code,0) as [Batch Code],Loaction_Code as [Location],TSPL_PP_STD_FINALQC_HEAD.Created_By as [Created By],convert(varchar,TSPL_PP_STD_FINALQC_HEAD.Created_Date,103) as [Created Date] from TSPL_PP_STD_FINALQC_HEAD " &
                    " WHERE  convert(date,TSPL_PP_STD_FINALQC_HEAD.QC_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PP_STD_FINALQC_HEAD.QC_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PP_STD_FINALQC_HEAD.Posted = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_PP_STD_FINALQC_HEAD.Posted = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PP_STD_FINALQC_HEAD.Loaction_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_PP_STD_FINALQC_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PP_STD_FINALQC_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_PP_STD_FINALQC_HEAD.QC_Date, TSPL_PP_STD_FINALQC_HEAD.QC_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Stage Process") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProcessProductionStageProcess)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST((TSPL_PP_STAGE_PROCESS_HEAD.Posted )as BIT) as Status,TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_CODE as [Document Id],convert(varchar,TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_DATE ,103) as [Document Date],isnull(TSPL_PP_STAGE_PROCESS_HEAD.Main_Batch_Code,0) as [Batch Code],TSPL_PP_STAGE_PROCESS_HEAD.Created_By as [Created By],convert(varchar,TSPL_PP_STAGE_PROCESS_HEAD.Created_Date,103) as [Created Date] from TSPL_PP_STAGE_PROCESS_HEAD  " &
                    " WHERE  convert(date,TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_DATE,103)<= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PP_STAGE_PROCESS_HEAD.Posted  = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_PP_STAGE_PROCESS_HEAD.Posted  = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_PP_STAGE_PROCESS_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PP_STAGE_PROCESS_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_DATE, TSPL_PP_STAGE_PROCESS_HEAD.STAGE_PROCESS_CODE  "

                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then
            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub

    Sub FillPurchaseOrder()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Requisition") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnPurchaseRequistion)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_REQUISITION_HEAD.Status as BIT ) as Status, TSPL_REQUISITION_HEAD.Requisition_Id as [Document Id], TSPL_REQUISITION_HEAD.Requisition_Date as [Document Date],isnull(TSPL_REQUISITION_HEAD.Total_RQ_Amt,0) as [Amount], TSPL_REQUISITION_HEAD.Expire_Date as [Exipe Date], TSPL_REQUISITION_HEAD.Require_Date  as [Require Date], TSPL_REQUISITION_HEAD.Description as Description, CAST(TSPL_REQUISITION_HEAD.On_Hold as bit) as Hold,TSPL_REQUISITION_HEAD.Created_By AS 'Created By' FROM TSPL_REQUISITION_HEAD WHERE  convert(date,TSPL_REQUISITION_HEAD.Requisition_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_REQUISITION_HEAD.Requisition_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_REQUISITION_HEAD.Status=0 "
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_REQUISITION_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then


                        qry += " and TSPL_REQUISITION_HEAD.Location  IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "

                    End If

                    If ChkAllowBulkPosting.Equals(0) Then ''when setting is OFF then only created user can post the document.
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_REQUISITION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ") "
                        Else
                            qry += " and TSPL_REQUISITION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ") "
                        End If
                    End If

                    qry += " ORDER BY TSPL_REQUISITION_HEAD.Requisition_Date , TSPL_REQUISITION_HEAD.Requisition_Id  "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Order") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnPurchaseOrder)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_PURCHASE_ORDER_HEAD.Status as BIT ) as Status, TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No as [Document Id], TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Date as [Document Date],isnull(TSPL_PURCHASE_ORDER_HEAD.PO_Total_Amt,0) as [Amount], TSPL_PURCHASE_ORDER_HEAD.Vendor_Code as [Vendor Code], TSPL_PURCHASE_ORDER_HEAD.Vendor_Name as [Vendor Name], TSPL_PURCHASE_ORDER_HEAD.Description as Description, CAST(TSPL_PURCHASE_ORDER_HEAD.On_Hold as bit) as Hold,TSPL_PURCHASE_ORDER_HEAD.Created_By as 'Created By' FROM TSPL_PURCHASE_ORDER_HEAD   where  convert(date,TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PURCHASE_ORDER_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_PURCHASE_ORDER_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then

                        qry += " and TSPL_PURCHASE_ORDER_HEAD.bill_to_location IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_PURCHASE_ORDER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PURCHASE_ORDER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Date , TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Gate Receipt Note") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnGRN)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_GRN_HEAD.Status as BIT ) as Status, TSPL_GRN_HEAD.GRN_No as [Document Id], TSPL_GRN_HEAD.GRN_Date as [Document Date],isnull(TSPL_GRN_HEAD.GRN_Total_Amt,0)  as [Amount], TSPL_GRN_HEAD.Vendor_Code as [Vendor Code], TSPL_GRN_HEAD.Vendor_Name as [Vendor Name], TSPL_GRN_HEAD.Description as Description, CAST(TSPL_GRN_HEAD.On_Hold as bit) as Hold,TSPL_GRN_HEAD.Created_By as 'Created By' FROM TSPL_GRN_HEAD   where  convert(date,TSPL_GRN_HEAD.GRN_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_GRN_HEAD.GRN_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_GRN_HEAD.Status=0 "
                        'qry += " and TSPL_GRN_HEAD.Status=0 " + IIf(clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal, IIf(clsCommon.CompairString(objCommonVar.SelectedUser, "All") = CompairStringResult.Equal, Nothing, " and TSPL_GRN_HEAD.Created_By in (" + objCommonVar.SelectedUser + ")"), " and TSPL_GRN_HEAD.Created_By='" + objCommonVar.CurrentUserCode + "'") + " "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_GRN_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_GRN_HEAD.Bill_To_Location IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and  TSPL_GRN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and  TSPL_GRN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_GRN_HEAD.GRN_Date , TSPL_GRN_HEAD.GRN_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Material Receipt Note") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnMRN)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_MRN_HEAD.Status as BIT ) as Status, TSPL_MRN_HEAD.MRN_No as [Document Id], TSPL_MRN_HEAD.MRN_Date as [Document Date],isnull(TSPL_MRN_HEAD.MRN_Total_Amt,0) as [Amount], TSPL_MRN_HEAD.Vendor_Code as [Vendor Code], TSPL_MRN_HEAD.Vendor_Name as [Vendor Name], TSPL_MRN_HEAD.Description, CAST(TSPL_MRN_HEAD.On_Hold as bit) as Hold,TSPL_MRN_HEAD.Created_By as 'Created By' FROM TSPL_MRN_HEAD   where  convert(date,TSPL_MRN_HEAD.MRN_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MRN_HEAD.MRN_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MRN_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_MRN_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MRN_HEAD.Bill_To_Location IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MRN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MRN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MRN_HEAD.MRN_Date , TSPL_MRN_HEAD.MRN_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Receipt Note") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnSRN)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_SRN_HEAD.Status as BIT ) as Status, TSPL_SRN_HEAD.SRN_No as [Document Id], TSPL_SRN_HEAD.SRN_Date as [Document Date],isnull(TSPL_SRN_HEAD.SRN_Total_Amt,0) as [Amount], TSPL_SRN_HEAD.Vendor_Code as [Vendor Code], TSPL_SRN_HEAD.Vendor_Name as [Vendor Name], TSPL_SRN_HEAD.Description as Description, CAST(TSPL_SRN_HEAD.On_Hold as bit) as Hold,TSPL_SRN_HEAD.Created_By as 'Created BY' FROM TSPL_SRN_HEAD   where  convert(date,TSPL_SRN_HEAD.SRN_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SRN_HEAD.SRN_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SRN_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SRN_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SRN_HEAD.Bill_To_Location IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SRN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SRN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_SRN_HEAD.SRN_Date , TSPL_SRN_HEAD.SRN_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnPurchaseInvoice)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_PI_HEAD.Status as BIT ) as Status, TSPL_PI_HEAD.PI_No as [Document Id], TSPL_PI_HEAD.PI_Date as [Document Date],isnull(TSPL_PI_HEAD.PI_Total_Amt,0) as [Amount], TSPL_PI_HEAD.Vendor_Code as [Vendor Code], TSPL_PI_HEAD.Vendor_Name as [Vendor Name], TSPL_PI_HEAD.Description as Description, CAST(TSPL_PI_HEAD.On_Hold as bit) as Hold,TSPL_PI_HEAD.Created_By as 'Created By' FROM TSPL_PI_HEAD   where  convert(date,TSPL_PI_HEAD.PI_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PI_HEAD.PI_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PI_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_PI_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PI_HEAD.Bill_To_Location  IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_PI_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ") "
                        Else
                            qry += " and TSPL_PI_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ") "
                        End If

                    End If
                    qry += " ORDER BY TSPL_PI_HEAD.PI_Date , TSPL_PI_HEAD.PI_No "
                End If
            End If
            ''Added on 12/10/2011
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Return") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnPurchaseReturn)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_PR_HEAD.Status as BIT ) as Status, TSPL_PR_HEAD.PR_No as [Document Id], TSPL_PR_HEAD.PR_Date as [Document Date],isnull(TSPL_PR_HEAD.PR_Total_Amt,0) as [Amount], TSPL_PR_HEAD.Vendor_Code as [Vendor Code], TSPL_PR_HEAD.Vendor_Name as [Vendor Name], TSPL_PR_HEAD.Description as Description, CAST(TSPL_PR_HEAD.On_Hold as bit) as Hold,TSPL_PR_HEAD.Created_By as 'Created By' FROM TSPL_PR_HEAD   where  convert(date,TSPL_PR_HEAD.PR_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PR_HEAD.PR_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)  "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PR_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_PR_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PR_HEAD.Bill_To_Location   IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_PR_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PR_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_PR_HEAD.PR_Date , TSPL_PR_HEAD.PR_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "RGP/NRGP") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnRGP_NRGP_Rpt)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_RGP_HEAD.Status as BIT ) as Status, TSPL_RGP_HEAD.RGP_No as [Document Id], TSPL_RGP_HEAD.RGP_Date as [Document Date],isnull(TSPL_RGP_HEAD.Document_Amount,0) as [Amount], TSPL_RGP_HEAD.Vendor_Code as [Vendor Code], TSPL_RGP_HEAD.Vendor_Name as [Vendor Name], TSPL_RGP_HEAD.Remarks as Description, CAST(TSPL_RGP_HEAD.On_Hold as bit) as Hold,TSPL_RGP_HEAD.Created_By as 'Created BY' FROM TSPL_RGP_HEAD   where  convert(date,TSPL_RGP_HEAD.RGP_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_RGP_HEAD.RGP_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_RGP_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_RGP_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_RGP_HEAD.Location     IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_RGP_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_RGP_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_RGP_HEAD.RGP_Date , TSPL_RGP_HEAD.RGP_No "
                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Issue/Return/Transfer") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnIssueReturn)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_IssueReturn_HEAD.Status as BIT ) as Status, TSPL_IssueReturn_HEAD.Doc_No as [Document Id], TSPL_IssueReturn_HEAD.Doc_Date as [Document Date], TSPL_IssueReturn_HEAD.Doc_Type as [Document Type],isnull(TSPL_IssueReturn_HEAD.Doc_Amt,0) as [Amount], TSPL_IssueReturn_HEAD.From_Location as [From Location],TSPL_IssueReturn_HEAD.To_Location as [To Location], TSPL_IssueReturn_HEAD.Remarks as Description, CAST(TSPL_IssueReturn_HEAD.On_Hold as bit) as Hold,TSPL_IssueReturn_HEAD.Created_By as 'Created By' FROM TSPL_IssueReturn_HEAD   where  convert(date,TSPL_IssueReturn_HEAD.Doc_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_IssueReturn_HEAD.Doc_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_IssueReturn_HEAD.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_IssueReturn_HEAD.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_IssueReturn_HEAD.From_Location     IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_IssueReturn_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_IssueReturn_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_IssueReturn_HEAD.Doc_Date , TSPL_IssueReturn_HEAD.Doc_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Scrap LoadOut") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.ScrapSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_SCRAPSALE_HEAD.ispost =1 then 1 else 0 end)as BIT) as Status, TSPL_SCRAPSALE_HEAD.shipment_No as [Document Id], TSPL_SCRAPSALE_HEAD.shipment_Date as [Document Date],isnull(TSPL_SCRAPSALE_HEAD.Doc_Amt,0) as [Amount],TSPL_SCRAPSALE_HEAD.cust_Code as [Customer Code],TSPL_SCRAPSALE_HEAD.cust_Name as [Customer Name], TSPL_SCRAPSALE_HEAD.Description as [Description],TSPL_SCRAPSALE_HEAD.Created_By as 'Created By' FROM TSPL_SCRAPSALE_HEAD WHERE  convert(date,TSPL_SCRAPSALE_HEAD.shipment_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SCRAPSALE_HEAD.shipment_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_SCRAPSALE_HEAD.ispost is NULL OR TSPL_SCRAPSALE_HEAD.ispost = '0') "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SCRAPSALE_HEAD.ispost = '1' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SCRAPSALE_HEAD.Loc_Code     IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_SCRAPSALE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SCRAPSALE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += "ORDER BY TSPL_SCRAPSALE_HEAD.shipment_Date , TSPL_SCRAPSALE_HEAD.shipment_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Scrap Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.ScrapSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_SCRAPINVOICE_HEAD.ispost =1 then 1 else 0 end)as BIT) as Status, TSPL_SCRAPINVOICE_HEAD.invoice_No as [Document Id], TSPL_SCRAPINVOICE_HEAD.shipment_Date as [Document Date],isnull(TSPL_SCRAPINVOICE_HEAD.Doc_Amt,0) as [Amount],TSPL_SCRAPINVOICE_HEAD.cust_Code as [Customer Code],TSPL_SCRAPINVOICE_HEAD.cust_Name as [Customer Name], TSPL_SCRAPINVOICE_HEAD.Description as [Description],TSPL_SCRAPINVOICE_HEAD.Created_By as 'Created By' FROM TSPL_SCRAPINVOICE_HEAD WHERE  convert(date,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SCRAPINVOICE_HEAD.shipment_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_SCRAPINVOICE_HEAD.ispost is NULL OR TSPL_SCRAPINVOICE_HEAD.ispost = '0') "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SCRAPINVOICE_HEAD.ispost = '1' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SCRAPINVOICE_HEAD.Loc_Code      IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SCRAPINVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SCRAPINVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += "ORDER BY TSPL_SCRAPINVOICE_HEAD.shipment_Date , TSPL_SCRAPINVOICE_HEAD.invoice_No "
                End If
            End If

            '---VARSHA ADD NIRQC
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "NIR QC") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnPendingApproval1)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    '                    qry = "SELECT CAST(TSPL_NIR_QC.Status as BIT ) as Status, TSPL_NIR_QC.Document_No as [Document Id],
                    'TSPL_NIR_QC.Document_Date as [Document Date],
                    'isnull(TSPL_MRN_HEAD.MRN_Total_Amt,0) as [Amount], TSPL_MRN_HEAD.Vendor_Code as [Vendor Code], TSPL_MRN_HEAD.Vendor_Name as [Vendor Name], TSPL_MRN_HEAD.Description, CAST(TSPL_MRN_HEAD.On_Hold as bit) as Hold,
                    'TSPL_NIR_QC.Created_By AS 'Created By' 
                    'FROM TSPL_NIR_QC 
                    'LEFT OUTER JOIN TSPL_MRN_HEAD ON TSPL_MRN_HEAD.MRN_No=TSPL_NIR_QC.MRN_No 
                    'WHERE
                    'convert(date,TSPL_NIR_QC.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_NIR_QC.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    qry = "select CAST(TSPL_NIR_QC.Status as BIT ) as Status, TSPL_NIR_QC.Document_No as [Document Id],
                    TSPL_NIR_QC.Document_Date as [Document Date],
isnull(TSPL_MRN_HEAD.MRN_Total_Amt,0) as [Amount],TSPL_MRN_HEAD.VehicleNo AS [Vehicle No],
CAST(TSPL_MRN_HEAD.On_Hold as bit) as Hold,
TSPL_MRN_HEAD.Against_GRN as [GRN No],FORMAT(TSPL_GRN_HEAD.GRN_Date, 'dd/MM/yyyy')as [GRN Date],
TSPL_PO_WEIGHTMENT_HEAD.Weighment_Code as [Weighment Code],FORMAT(TSPL_PO_WEIGHTMENT_HEAD.Weighment_Date, 'dd/MM/yyyy') as [Weighment Date],
TSPL_PURCHASE_ORDER_HEAD.RefTendorNo as [RAL NO],TSPL_MRN_HEAD.Vendor_Code as [Vendor Code],TSPL_MRN_HEAD.Vendor_Name as [Vendor Name],
TSPL_LOCATION_MASTER.Location_Desc AS [Location Desc],TSPL_MRN_DETAIL.Item_Code AS [Item Code],TSPL_ITEM_MASTER.Item_Desc AS[Description]
from TSPL_NIR_QC
left outer join TSPL_MRN_HEAD  on TSPL_MRN_HEAD.MRN_No=TSPL_NIR_QC.MRN_No
left outer join TSPL_MRN_DETAIL on TSPL_MRN_DETAIL.MRN_No=TSPL_MRN_HEAD.MRN_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_MRN_DETAIL.Item_Code 
left outer join TSPL_GRN_HEAD on TSPL_GRN_HEAD.GRN_No=TSPL_MRN_HEAD.Against_GRN
left outer join TSPL_PO_WEIGHTMENT_HEAD on TSPL_PO_WEIGHTMENT_HEAD.Against_GRN_No=TSPL_MRN_HEAD.Against_GRN
left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_GRN_HEAD.Against_PO
left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_MRN_HEAD.Bill_To_Location
                    WHERE
                    convert(date,TSPL_NIR_QC.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_NIR_QC.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_NIR_QC.Status=0 "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_NIR_QC.Status=1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MRN_HEAD.Bill_To_Location    IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_NIR_QC.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_NIR_QC.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_NIR_QC.Document_Date , TSPL_NIR_QC.Document_No "
                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------

    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling--------Module(Material Management)-----------
    ''-------------------------------------------------------------------

    Sub FillMaterialManagement()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Transfer(Load-In)") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.Transfer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    If rbtnStatusPending.IsChecked = True Then
                        qry = "SELECT  CAST((case when TSPL_TRANSFER_HEAD.Post ='Y' then 1 else 0 end)as BIT) as Status, TSPL_TRANSFER_HEAD.Transfer_No  as [Document Id], TSPL_TRANSFER_HEAD.Transfer_Date as [Document Date], isnull(TSPL_TRANSFER_HEAD.Total_Transfer_Amount,0) as Amount, TSPL_TRANSFER_HEAD.Transfer_Type as [Transfer Type], TSPL_TRANSFER_HEAD.Load_Out_No, TSPL_TRANSFER_HEAD.Reference as [Reference], TSPL_TRANSFER_HEAD.Description as [Description],TSPL_TRANSFER_HEAD.Created_By as 'Created By', TSPL_TRANSFER_HEAD.FromLoc_Desc AS [To Location], TSPL_TRANSFER_HEAD.ToLoc_Desc as [From Location] FROM TSPL_TRANSFER_HEAD WHERE  convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) And TSPL_TRANSFER_HEAD.Transfer_Type='LI' and (TSPL_TRANSFER_HEAD.Post = 'N' OR TSPL_TRANSFER_HEAD.Post IS NULL) and TSPL_TRANSFER_HEAD.is_Auto_Created_Trans=0"
                        Isrefreshed = False
                    Else
                        qry = "SELECT  CAST((case when TSPL_TRANSFER_HEAD.Post ='Y' then 1 else 0 end)as BIT) as Status, TSPL_TRANSFER_HEAD.Transfer_No  as [Document Id], TSPL_TRANSFER_HEAD.Transfer_Date as [Document Date], isnull(TSPL_TRANSFER_HEAD.Total_Transfer_Amount,0) as Amount, TSPL_TRANSFER_HEAD.Transfer_Type as [Transfer Type], TSPL_TRANSFER_HEAD.Posting_Date  as [Posting Date], TSPL_TRANSFER_HEAD.Load_Out_No, TSPL_TRANSFER_HEAD.Reference as [Reference], TSPL_TRANSFER_HEAD.Description as [Description],TSPL_TRANSFER_HEAD.Created_By as 'Created By',TSPL_TRANSFER_HEAD.FromLoc_Desc AS [To Location], TSPL_TRANSFER_HEAD.ToLoc_Desc as [From Location]  FROM TSPL_TRANSFER_HEAD WHERE  convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) And TSPL_TRANSFER_HEAD.Transfer_Type='LI' and TSPL_TRANSFER_HEAD.Post = 'Y'"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_TRANSFER_HEAD.To_Location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += "  AND Created_By in (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ") "
                        Else
                            qry += "  AND Created_By in (" + clsCommon.GetMulcallString(arrSelectedUser) + ") "
                        End If
                    End If
                    qry += " and  TSPL_TRANSFER_HEAD.Trans_Type<>'Route'"
                    qry += " ORDER BY TSPL_TRANSFER_HEAD.Transfer_Date , TSPL_TRANSFER_HEAD.Transfer_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Transfer(Load-Out)") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.Transfer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    If rbtnStatusPending.IsChecked = True Then
                        qry = "SELECT CAST((case when TSPL_TRANSFER_HEAD.Post ='Y' then 1 else 0 end)as BIT) as Status, TSPL_TRANSFER_HEAD.Transfer_No  as [Document Id], TSPL_TRANSFER_HEAD.Transfer_Date as [Document Date], isnull(TSPL_TRANSFER_HEAD.Total_Transfer_Amount,0) as Amount, TSPL_TRANSFER_HEAD.Transfer_Type as [Transfer Type], TSPL_TRANSFER_HEAD.Reference as [Reference], TSPL_TRANSFER_HEAD.Description as [Description],TSPL_TRANSFER_HEAD.Created_By as 'Created By',TSPL_TRANSFER_HEAD.FromLoc_Desc AS [From Location], TSPL_TRANSFER_HEAD.ToLoc_Desc as [To Location] FROM TSPL_TRANSFER_HEAD WHERE  convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) And TSPL_TRANSFER_HEAD.Transfer_Type='LO' and (TSPL_TRANSFER_HEAD.Post = 'N' OR TSPL_TRANSFER_HEAD.Post IS NULL )and TSPL_TRANSFER_HEAD.is_Auto_Created_Trans=0"
                        Isrefreshed = False
                    Else
                        qry = "SELECT CAST((case when TSPL_TRANSFER_HEAD.Post ='Y' then 1 else 0 end)as BIT) as Status, TSPL_TRANSFER_HEAD.Transfer_No  as [Document Id], TSPL_TRANSFER_HEAD.Transfer_Date as [Document Date], isnull(TSPL_TRANSFER_HEAD.Total_Transfer_Amount,0) as Amount, TSPL_TRANSFER_HEAD.Transfer_Type as [Transfer Type],TSPL_TRANSFER_HEAD.Posting_Date  as [Posting Date], TSPL_TRANSFER_HEAD.Reference as [Reference], TSPL_TRANSFER_HEAD.Description as [Description],TSPL_TRANSFER_HEAD.Created_By as 'Created By', TSPL_TRANSFER_HEAD.FromLoc_Desc AS [From Location], TSPL_TRANSFER_HEAD.ToLoc_Desc as [To Location]   FROM TSPL_TRANSFER_HEAD WHERE  convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_TRANSFER_HEAD.Transfer_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) And TSPL_TRANSFER_HEAD.Transfer_Type='LO' and TSPL_TRANSFER_HEAD.Post = 'Y'"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_TRANSFER_HEAD.To_Location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_TRANSFER_HEAD.From_Location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += "  AND Created_By in (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ") "
                        Else
                            qry += "  AND Created_By in (" + clsCommon.GetMulcallString(arrSelectedUser) + ") "
                        End If
                    End If
                End If
            End If
        Else
            'Load_Authorisation("Transfer(Load-In)")
            'If dtAuthen.Rows.Count > 0 Then
            '    If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
            qry = "Select CAST((case when Max(Status) ='Y' then 1 else 0 end)as BIT) as Status, MAX([Document Id]) as [Document Id], MAX([Document Date]) as [Document Date], SUM(Amount) as Amount, MAX(ItemType) as [Item Type], MAX([Referenced Document]) as [Referenced Document], MAX([Transfer No]) as [Transfer No], MAX(Description) as Description from (SELECT TSPL_ADJUSTMENT_HEADER.Posted as Status, TSPL_ADJUSTMENT_HEADER.Adjustment_No as [Document Id], TSPL_ADJUSTMENT_HEADER.Adjustment_Date as [Document Date], isnull(Item_Cost,0) as Amount, TSPL_ADJUSTMENT_HEADER.ItemType as ItemType, TSPL_ADJUSTMENT_HEADER.Reference_Document as [Referenced Document], " &
            " (case when  TSPL_ADJUSTMENT_HEADER.Reference_Document ='Sale Invoice' then TSPL_SHIPMENT_MASTER.Transfer_No else  case when TSPL_ADJUSTMENT_HEADER.Reference_Document ='Load Out/Transfer' then  TSPL_ADJUSTMENT_HEADER.Document_No end end ) as [Transfer No], TSPL_ADJUSTMENT_HEADER.Description as Description,TSPL_ADJUSTMENT_HEADER.Created_By as 'Created By'   FROM TSPL_ADJUSTMENT_HEADER Left Outer Join TSPL_ADJUSTMENT_DETAIL on TSPL_ADJUSTMENT_DETAIL.Adjustment_No=TSPL_ADJUSTMENT_HEADER.Adjustment_No " &
         " LEFT OUTER JOIN TSPL_SALE_INVOICE_HEAD on TSPL_SALE_INVOICE_HEAD.Sale_Invoice_No=TSPL_ADJUSTMENT_HEADER.Document_No  left outer join TSPL_SHIPMENT_MASTER on TSPL_SHIPMENT_MASTER .Shipment_No =TSPL_SALE_INVOICE_HEAD.Shipment_No  WHERE  convert(date,TSPL_ADJUSTMENT_HEADER.Adjustment_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_ADJUSTMENT_HEADER.Adjustment_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Empty Transactions") = CompairStringResult.Equal Then
                Load_Authorisation(clsUserMgtCode.mbtnEmptyTrans)
                If dtAuthen.Rows.Count > 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                        qry += " AND (TSPL_ADJUSTMENT_HEADER.ItemType='E' AND (ISNULL(Reference_Document, '')='' OR TSPL_SALE_INVOICE_HEAD.Shipment_Type='Sale')) "
                    End If
                End If
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Entry") = CompairStringResult.Equal Then
                Load_Authorisation(clsUserMgtCode.mbtnEmptyTrans)
                If dtAuthen.Rows.Count > 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                        qry += " AND TSPL_ADJUSTMENT_HEADER.ItemType IN ('FT', 'FM')"
                    End If
                End If

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Adjustment") = CompairStringResult.Equal AndAlso ChkMilkType.Checked = True Then
                Load_Authorisation(clsUserMgtCode.mbtnStoreAdjustment)
                If dtAuthen.Rows.Count > 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                        qry += " AND TSPL_ADJUSTMENT_HEADER.IsMilkType='1'"
                    Else
                        qry += " AND TSPL_ADJUSTMENT_HEADER.IsMilkType='0'"
                        'qry += " AND TSPL_ADJUSTMENT_HEADER.ItemType='OT'"
                    End If
                End If
            End If

            If rbtnStatusPending.IsChecked = True Then
                qry += " and TSPL_ADJUSTMENT_HEADER.Posted = 'N'"
                Isrefreshed = False
            ElseIf rbtnStatusPosted.IsChecked = True Then
                qry += " and TSPL_ADJUSTMENT_HEADER.Posted = 'Y'"
                Isrefreshed = True
            End If
            If ChkAllowBulkPosting.Equals(0) Then
                If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                    qry += " and TSPL_ADJUSTMENT_HEADER.Created_By in (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                Else
                    qry += " and TSPL_ADJUSTMENT_HEADER.Created_By in (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                End If
            End If
            qry += " ) AAA Group By [Document Id] "
            qry += " order by [Document Date], [Document Id] "
            ''-------------------------------------------Code  Ends Here-----------------------------------------------
        End If
        '    End If
        'End If

        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = Nothing
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling--------Module(Sales & Distribution)----------
    ''-------------------------------------------------------------------

    Sub FillSalesNDistribution()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Order") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmSalesOrderProductSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_SALES_ORDER_HEAD.Is_Post =1 then 1 else 0 end)as BIT) as Status, TSPL_SALES_ORDER_HEAD.Order_No as [Document Id], TSPL_SALES_ORDER_HEAD.Order_Date as [Document Date],isnull(TSPL_SALES_ORDER_HEAD.Total_Order_Amt,0) as [Amount], TSPL_SALES_ORDER_HEAD.Cust_Code as [Customer Code], TSPL_SALES_ORDER_HEAD.Cust_Name as [Customer Name], TSPL_SALES_ORDER_HEAD.Description as [Description], CAST((case when TSPL_SALES_ORDER_HEAD.On_Hold  =1 then 1 else 0 end)as BIT) as Hold,TSPL_SALES_ORDER_HEAD.Created_By as 'Created By' FROM TSPL_SALES_ORDER_HEAD WHERE  convert(date,TSPL_SALES_ORDER_HEAD.Order_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SALES_ORDER_HEAD.Order_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SALES_ORDER_HEAD.Is_Post = 0 "
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_SALES_ORDER_HEAD.Is_Post = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SALES_ORDER_HEAD.location      IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SALES_ORDER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SALES_ORDER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += "ORDER BY TSPL_SALES_ORDER_HEAD.Order_Date , TSPL_SALES_ORDER_HEAD.Order_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment/Invoice") = CompairStringResult.Equal Then
            gv1.DataSource = Nothing
            Dim strTransType As String = ""
            If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Product Sale") = CompairStringResult.Equal Then
                Load_Authorisation(clsUserMgtCode.frmShipmentProductSale)
                strTransType = "PS"
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fresh Sale") = CompairStringResult.Equal Then
                Load_Authorisation(clsUserMgtCode.FrmDispatchFreshSale)
                strTransType = "FS"
            End If
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then

                    qry = "SELECT  CAST(TSPL_SD_SHIPMENT_HEAD.Status as BIT) as Status, TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document Id], TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date], TSPL_SD_SALE_INVOICE_HEAD.Document_Code as [Sale Invoice], isnull(TSPL_SD_SHIPMENT_HEAD.Total_Amt,0) as [Amount], TSPL_SD_SHIPMENT_HEAD.Bill_To_Location As Location , TSPL_LOCATION_MASTER.Location_Desc as [Location Desc], TSPL_SD_SHIPMENT_HEAD.Customer_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], TSPL_SD_SHIPMENT_HEAD.Description as [Description], CAST(TSPL_SD_SHIPMENT_HEAD.On_Hold as BIT) as Hold,TSPL_SD_SHIPMENT_HEAD.Created_By as 'Created By' FROM TSPL_SD_SHIPMENT_HEAD" &
                        " Left Outer Join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code" &
                        " Left Outer Join TSPL_LOCATION_MASTER on TSPL_SD_SHIPMENT_HEAD.Bill_To_Location=TSPL_LOCATION_MASTER.Location_Code" &
                        " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code WHERE convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status <> 1"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SD_SHIPMENT_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " and TSPL_SD_SHIPMENT_HEAD.Trans_Type='" & strTransType & "'"
                    qry += "ORDER BY TSPL_SD_SHIPMENT_HEAD.Document_Date , TSPL_SD_SHIPMENT_HEAD.Document_Code"
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmSaleInvoiceProductSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_SD_SALE_INVOICE_HEAD.Status as BIT) as Status, TSPL_SD_SALE_INVOICE_HEAD.Document_Code as [Document Id], TSPL_SD_SALE_INVOICE_HEAD.Document_Date as [Document Date],isnull(TSPL_SD_SALE_INVOICE_HEAD.Total_Amt,0) as [Amount], TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No as [Shipment No], TSPL_SD_SALE_INVOICE_HEAD.Customer_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], TSPL_SD_SALE_INVOICE_HEAD.Description as [Description], CAST(TSPL_SD_SALE_INVOICE_HEAD.On_Hold as BIT) as Hold,TSPL_SD_SALE_INVOICE_HEAD.Created_By as 'Created By'" &
                    " FROM TSPL_SD_SALE_INVOICE_HEAD" &
                    " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code" &
                    " WHERE  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 0"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SD_SALE_INVOICE_HEAD.bill_to_location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += "ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Code "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Return") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmSaleReturnProductSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_SD_SALE_RETURN_HEAD.Status as BIT) as Status, TSPL_SD_SALE_RETURN_HEAD.Document_Code as [Document Id], TSPL_SD_SALE_RETURN_HEAD.Document_Date as [Document Date],isnull(TSPL_SD_SALE_RETURN_HEAD.Total_Amt,0) as [Amount], TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No As Invoice_No, TSPL_SD_SALE_RETURN_HEAD.Customer_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_SD_SALE_RETURN_HEAD.Description as [Description],TSPL_SD_SALE_RETURN_HEAD.Created_By as 'Created By' FROM TSPL_SD_SALE_RETURN_HEAD LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_RETURN_HEAD.Customer_Code WHERE  convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_RETURN_HEAD.Status = 1"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_RETURN_HEAD.Status = 0"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location  in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SALE_RETURN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SALE_RETURN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += "ORDER BY TSPL_SALE_RETURN_HEAD.Sale_Return_Date , TSPL_SALE_RETURN_HEAD.Sale_Return_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Return (Inter Company)") = CompairStringResult.Equal Then
            Load_Authorisation("Sale Return (Inter Company)")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_SALE_RETURN_INTER_HEAD.Is_Post =1 then 1 else 0 end)as BIT) as Status, TSPL_SALE_RETURN_INTER_HEAD.Document_No  as [Document Id], TSPL_SALE_RETURN_INTER_HEAD.Document_Date as [Document Date],isnull(TSPL_SALE_RETURN_INTER_HEAD.Total_Order_Amt,0) as [Amount], TSPL_SALE_RETURN_INTER_HEAD.Cust_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_SALE_RETURN_INTER_HEAD.Description as [Description],TSPL_SALE_RETURN_INTER_HEAD.Created_By as 'Created By' FROM TSPL_SALE_RETURN_INTER_HEAD Left Outer Join TSPL_CUSTOMER_MASTER ON TSPL_SALE_RETURN_INTER_HEAD.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code Left Outer  Join TSPL_LOCATION_MASTER ON TSPL_SALE_RETURN_INTER_HEAD.Location=TSPL_LOCATION_MASTER.Location_Code WHERE  convert(date,TSPL_SALE_RETURN_INTER_HEAD.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SALE_RETURN_INTER_HEAD.Document_Date  ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " AND (ISNULL(TSPL_SALE_RETURN_INTER_HEAD.Is_Post,0)=0 OR TSPL_SALE_RETURN_INTER_HEAD.Is_Post=0)"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " AND TSPL_SALE_RETURN_INTER_HEAD.Is_Post=1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SALE_RETURN_INTER_HEAD.location  in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SALE_RETURN_INTER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SALE_RETURN_INTER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_SALE_RETURN_INTER_HEAD.Document_Date , TSPL_SALE_RETURN_INTER_HEAD.Document_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Quick Settlement") = CompairStringResult.Equal Then
            Load_Authorisation("Quick Settlement")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT Distinct CAST((case when tspl_QuickSettleMent.Post ='Y' then 1 else 0 end)as BIT) as Status, tspl_QuickSettleMent.Quick_SettleMent_Id as [Document Id], tspl_QuickSettleMent.Quick_Settlement_Date as [Document Date],isnull(tspl_QuickSettleMent.Transfer_Amount,0) as [Amount],tspl_QuickSettleMent.Transfer_Number, tspl_QuickSettleMent.Comments as [Description],tspl_QuickSettleMent.Created_By as 'Created By' FROM tspl_QuickSettleMent WHERE  convert(date,tspl_QuickSettleMent.Quick_Settlement_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,tspl_QuickSettleMent.Quick_Settlement_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (tspl_QuickSettleMent.Post is NULL OR tspl_QuickSettleMent.Post = 'N')"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and tspl_QuickSettleMent.Post = 'Y'"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and tspl_QuickSettleMent.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and tspl_QuickSettleMent.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += "ORDER BY tspl_QuickSettleMent.Quick_Settlement_Date , tspl_QuickSettleMent.Quick_SettleMent_Id "

                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling --------Module(General Ledger)---------------
    ''-------------------------------------------------------------------
    Sub FillSalesAndDistributionNew()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment") = CompairStringResult.Equal Then
            gv1.DataSource = Nothing
            'Dim strTransType As String = "P"
            Load_Authorisation(clsUserMgtCode.frmSNShipment)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then

                    qry = "SELECT  CAST(TSPL_SD_SHIPMENT_HEAD.Status as BIT) as Status, TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document Id], TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date], TSPL_SD_SALE_INVOICE_HEAD.Document_Code as [Sale Invoice], isnull(TSPL_SD_SHIPMENT_HEAD.Total_Amt,0) as [Amount], TSPL_SD_SHIPMENT_HEAD.Bill_To_Location As Location , TSPL_LOCATION_MASTER.Location_Desc as [Location Desc], TSPL_SD_SHIPMENT_HEAD.Customer_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], TSPL_SD_SHIPMENT_HEAD.Description as [Description], CAST(TSPL_SD_SHIPMENT_HEAD.On_Hold as BIT) as Hold,TSPL_SD_SHIPMENT_HEAD.Created_By as 'Created By' FROM TSPL_SD_SHIPMENT_HEAD" &
                        " Left Outer Join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code" &
                        " Left Outer Join TSPL_LOCATION_MASTER on TSPL_SD_SHIPMENT_HEAD.Bill_To_Location=TSPL_LOCATION_MASTER.Location_Code" &
                        " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code 
                        WHERE convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status <> 1"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SD_SHIPMENT_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    'qry += " and TSPL_SD_SHIPMENT_HEAD.Trans_Type='" & strTransType & "'"
                    qry += "ORDER BY TSPL_SD_SHIPMENT_HEAD.Document_Date , TSPL_SD_SHIPMENT_HEAD.Document_Code"
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmSNSaleInvoice)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_SD_SALE_INVOICE_HEAD.Status as BIT) as Status, TSPL_SD_SALE_INVOICE_HEAD.Document_Code as [Document Id], TSPL_SD_SALE_INVOICE_HEAD.Document_Date as [Document Date],isnull(TSPL_SD_SALE_INVOICE_HEAD.Total_Amt,0) as [Amount], TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No as [Shipment No], TSPL_SD_SALE_INVOICE_HEAD.Customer_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], TSPL_SD_SALE_INVOICE_HEAD.Description as [Description], CAST(TSPL_SD_SALE_INVOICE_HEAD.On_Hold as BIT) as Hold,TSPL_SD_SALE_INVOICE_HEAD.Created_By as 'Created By'" &
                    " FROM TSPL_SD_SALE_INVOICE_HEAD" &
                    " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code" &
                    " WHERE  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type = 'P' "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 0"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SD_SALE_INVOICE_HEAD.bill_to_location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += "ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Code "
                End If
            End If

            '        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Demand Booking") = CompairStringResult.Equal Then
            '            Load_Authorisation(clsUserMgtCode.FrmProductDemandBooking)
            '            If dtAuthen.Rows.Count > 0 Then
            '                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
            '                    gv1.DataSource = Nothing

            '                    qry = "SELECT  CAST(TSPL_Product_DEMAND_BOOKING_MASTER.posted as BIT) as Status, 

            ''' as [Description], '' as Hold,
            'TSPL_Product_DEMAND_BOOKING_MASTER.Document_no as [Document Id],
            '                         TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date as [Document Date],  isnull(TSPL_Product_DEMAND_BOOKING_DETAIL.ItemNetAmount,0) as [Amount], 
            '                         TSPL_Product_DEMAND_BOOKING_MASTER.Location_Code As Location , TSPL_LOCATION_MASTER.Location_Desc as [Location Desc],
            '                         TSPL_Product_DEMAND_BOOKING_DETAIL.Cust_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], 
            '                                      TSPL_Product_DEMAND_BOOKING_MASTER.Created_By as 'Created By' FROM TSPL_Product_DEMAND_BOOKING_MASTER" &
            '                        " Left Outer Join TSPL_Product_DEMAND_BOOKING_DETAIL on TSPL_Product_DEMAND_BOOKING_DETAIL.Document_No=TSPL_Product_DEMAND_BOOKING_MASTER.Document_No" &
            '                        " Left Outer Join TSPL_LOCATION_MASTER on TSPL_Product_DEMAND_BOOKING_MASTER.Location_Code=TSPL_LOCATION_MASTER.Location_Code" &
            '                        " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_Product_DEMAND_BOOKING_DETAIL.Cust_Code 
            '                        WHERE   convert(date,TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"
            '                    If rbtnStatusPending.IsChecked = True Then
            '                        qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.posted = 0"
            '                        Isrefreshed = False
            '                    ElseIf rbtnStatusPosted.IsChecked = True Then
            '                        qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.posted = 1"
            '                        Isrefreshed = True
            '                    End If
            '                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
            '                        qry += " AND TSPL_Product_DEMAND_BOOKING_MASTER.Location_Code in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
            '                    End If
            '                    If ChkAllowBulkPosting.Equals(0) Then
            '                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
            '                            qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
            '                        Else
            '                            qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
            '                        End If
            '                    End If
            '                    qry += "ORDER BY TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date, TSPL_Product_DEMAND_BOOKING_MASTER.Document_No "
            '                End If
            '            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Dispatch") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FrmProductDispatch)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then

                    qry = "SELECT  CAST(TSPL_SD_SHIPMENT_HEAD.Status as BIT) as Status, TSPL_SD_SHIPMENT_HEAD.Document_Code as [Document Id], TSPL_SD_SHIPMENT_HEAD.Document_Date as [Document Date], TSPL_SD_SALE_INVOICE_HEAD.Document_Code as [Sale Invoice], isnull(TSPL_SD_SHIPMENT_HEAD.Total_Amt,0) as [Amount], TSPL_SD_SHIPMENT_HEAD.Bill_To_Location As Location , TSPL_LOCATION_MASTER.Location_Desc as [Location Desc], TSPL_SD_SHIPMENT_HEAD.Customer_Code as [Customer Code], TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], TSPL_SD_SHIPMENT_HEAD.Description as [Description], CAST(TSPL_SD_SHIPMENT_HEAD.On_Hold as BIT) as Hold,TSPL_SD_SHIPMENT_HEAD.Created_By as 'Created By' FROM TSPL_SD_SHIPMENT_HEAD" &
                        " Left Outer Join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD.Document_Code" &
                        " Left Outer Join TSPL_LOCATION_MASTER on TSPL_SD_SHIPMENT_HEAD.Bill_To_Location=TSPL_LOCATION_MASTER.Location_Code" &
                        " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code 
                        WHERE TSPL_SD_SHIPMENT_HEAD.Item_Type in('P','I')  and convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status <> 1"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_SD_SHIPMENT_HEAD.Bill_To_Location in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    'qry += " and TSPL_SD_SHIPMENT_HEAD.Trans_Type='" & strTransType & "'"
                    qry += "ORDER BY TSPL_SD_SHIPMENT_HEAD.Document_Date , TSPL_SD_SHIPMENT_HEAD.Document_Code"
                End If
            End If


        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub
    Sub FillGeneralLedger()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Journal Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.journalEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_JOURNAL_MASTER.Authorized ='A' then 1 else 0 end)as BIT) as Status, TSPL_JOURNAL_MASTER.Voucher_No as [Document Id], TSPL_JOURNAL_MASTER.Voucher_Date as [Document Date],isnull(TSPL_JOURNAL_MASTER.Total_Debit_Amt,0) as [Amount], TSPL_JOURNAL_MASTER.Posting_Date as [Posting Date],TSPL_JOURNAL_MASTER.Source_Doc_No as [Source Document Code], TSPL_JOURNAL_MASTER.Source_Code as [Source Code],TSPL_JOURNAL_MASTER.CustVend_Code as [Customer/Vendor Code],TSPL_JOURNAL_MASTER.CustVend_Name as [Customer/Vendor Name], TSPL_JOURNAL_MASTER.Voucher_Desc as [Description],TSPL_JOURNAL_MASTER.Created_By as 'Created By' FROM TSPL_JOURNAL_MASTER WHERE  convert(date,TSPL_JOURNAL_MASTER.Voucher_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_JOURNAL_MASTER.Voucher_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_JOURNAL_MASTER.Authorized = 'N' OR TSPL_JOURNAL_MASTER.Authorized IS NULL)"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_JOURNAL_MASTER.Authorized = 'A' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_JOURNAL_MASTER.Segment_code in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_JOURNAL_MASTER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_JOURNAL_MASTER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If

                    qry += " ORDER BY TSPL_JOURNAL_MASTER.Voucher_Date , TSPL_JOURNAL_MASTER.Voucher_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VCGL Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnVCGLEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((TSPL_VCGL_Head.Status)as BIT) as Status, TSPL_VCGL_Head.Document_No as [Document Id], CONVERT(Varchar,TSPL_VCGL_Head.Document_Date, 103) as [Document Date], CONVERT(Varchar,TSPL_VCGL_Head.Posting_Date, 103) as [Posting Date],TSPL_VCGL_Head.VC_Code as [VC Code], TSPL_VCGL_Head.VC_Name as [VC Name],isnull(TSPL_VCGL_Head.Tot_Dr_Amount,0) as [Debit Amount],isnull(TSPL_VCGL_Head.Tot_Cr_Amount,0) as [Credit Amount], isnull(TSPL_VCGL_Head.Amount,0) as [Amount], TSPL_VCGL_Head.Remarks as [Description],TSPL_VCGL_Head.Created_By as 'Created By' FROM TSPL_VCGL_Head WHERE  convert(date,TSPL_VCGL_Head.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_VCGL_Head.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " AND ISNULL(TSPL_VCGL_Head.Status,0) = 0 "
                        Isrefreshed = False
                    Else
                        qry += " AND ISNULL(TSPL_VCGL_Head.Status,0) = 1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_VCGL_Head.Location_Segment  in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_VCGL_Head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_VCGL_Head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_VCGL_Head.Document_Date  , TSPL_VCGL_Head.Document_No "
                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling --------Module(Tax Deducted At Source)-------
    ''-------------------------------------------------------------------

    Sub FillTaxDdctdAtSource()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Remittance Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.remittanceentry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_REMITTANCE_ENTRY.Posted ='Y' then 1 else 0 end)as BIT) as Status, TSPL_REMITTANCE_ENTRY.Remittance_Code as [Document Id], TSPL_REMITTANCE_ENTRY.Remittance_Date as [Document Date],isnull(TSPL_REMITTANCE_ENTRY.Amt_To_Remit,0) as [Amount], TSPL_REMITTANCE_ENTRY.Bank_Code as [Bank Code],TSPL_REMITTANCE_ENTRY.Payment_Code as [Payment Code],TSPL_REMITTANCE_ENTRY.Cheque_No as [Cheque No], TSPL_REMITTANCE_ENTRY.Description  as [Description],TSPL_REMITTANCE_ENTRY.Created_By as 'Created By' FROM TSPL_REMITTANCE_ENTRY WHERE  convert(date,TSPL_REMITTANCE_ENTRY.Remittance_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_REMITTANCE_ENTRY.Remittance_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_REMITTANCE_ENTRY.Posted = 'N' OR TSPL_REMITTANCE_ENTRY.Posted IS NULL)"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_REMITTANCE_ENTRY.Posted = 'Y'"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_REMITTANCE_ENTRY.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_REMITTANCE_ENTRY.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_REMITTANCE_ENTRY.Remittance_Date , TSPL_REMITTANCE_ENTRY.Remittance_Code "
                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then
            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling--------Module(Payables)----------------------
    ''-------------------------------------------------------------------

    Sub FillPayables()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Payment Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.PaymentEntryNew)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_PAYMENT_HEADER.Posted ='P' then 1 else 0 end)as BIT) as Status, TSPL_PAYMENT_HEADER.Payment_No as [Document Id], TSPL_PAYMENT_HEADER.Payment_Date as [Document Date], isnull(TSPL_PAYMENT_HEADER.Payment_Amount,0) as [Amount], TSPL_PAYMENT_HEADER.Vendor_Code as [Vendor Code],TSPL_PAYMENT_HEADER.Vendor_Name as [Vendor Name], TSPL_PAYMENT_HEADER.Bank_Code as [Bank Code], right(TSPL_BANK_MASTER.BANKACC,3) as Location, TSPL_PAYMENT_HEADER.Payment_Code as [Payment Code],TSPL_PAYMENT_HEADER.Payment_Type as [Payment Type],TSPL_PAYMENT_HEADER.Cheque_No as [Cheque No], TSPL_PAYMENT_HEADER.Narration as [Description],TSPL_PAYMENT_HEADER.Created_By as 'Created BY' FROM TSPL_PAYMENT_HEADER Left Outer Join TSPL_BANK_MASTER on TSPL_PAYMENT_HEADER.Bank_Code=TSPL_BANK_MASTER.BANK_CODE WHERE  convert(date,TSPL_PAYMENT_HEADER.Payment_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_PAYMENT_HEADER.Payment_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and ISNULL(TSPL_PAYMENT_HEADER.Posted,0) = '0'"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_PAYMENT_HEADER.Posted = '1'"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PAYMENT_HEADER.Location_Code  in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_PAYMENT_HEADER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PAYMENT_HEADER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_PAYMENT_HEADER.Payment_Date , TSPL_PAYMENT_HEADER.Payment_No "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "AP Invoice Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnAPInvoiceEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date, '')= '' then 0 else 1 end)as BIT) as Status, TSPL_VENDOR_INVOICE_HEAD.Document_No as [Document Id], TSPL_VENDOR_INVOICE_HEAD.invoice_entry_date as [Document Date],isnull(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as [Amount], TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as [Vendor Code],TSPL_VENDOR_INVOICE_HEAD.Vendor_Name as [Vendor Name], TSPL_VENDOR_INVOICE_HEAD.RefDocNo, TSPL_VENDOR_INVOICE_HEAD.Description as " &
                          "[Description], CAST((case when TSPL_VENDOR_INVOICE_HEAD.On_Hold ='Y' then 1 else 0 end)as BIT) as Hold,TSPL_VENDOR_INVOICE_HEAD.Created_By as 'Created BY' FROM TSPL_VENDOR_INVOICE_HEAD WHERE  convert(date,TSPL_VENDOR_INVOICE_HEAD.invoice_entry_date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_VENDOR_INVOICE_HEAD.invoice_entry_date ,103) <= convert(date,'" + dtpToDate.Value + "',103) and TSPL_VENDOR_INVOICE_HEAD .ISProcurementDeduction =0 "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " And ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date, '')= '' "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date, '')<> '' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then

                        qry += " and TSPL_VENDOR_INVOICE_HEAD.Loc_Code   in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"

                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_VENDOR_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_VENDOR_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date , TSPL_VENDOR_INVOICE_HEAD.Document_No "

                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "BMC Mobile Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnAPInvoiceEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when ISNULL(TSPL_MILK_COLLECTION_BMCDCS.Posted_Date, '')= '' then 0 else 1 end)as BIT) as Status, TSPL_MILK_COLLECTION_BMCDCS.PK_ID as [Document Id], FORMAT(TSPL_MILK_COLLECTION_BMCDCS.IDate, 'dd/MM/yyyy') AS [Date], TSPL_MILK_COLLECTION_BMCDCS.MCC_Code as [BMC Code] FROM TSPL_MILK_COLLECTION_BMCDCS WHERE  convert(date,TSPL_MILK_COLLECTION_BMCDCS.IDate) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MILK_COLLECTION_BMCDCS.IDate ,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " And ISNULL(TSPL_MILK_COLLECTION_BMCDCS.Posted_Date, '')= '' "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and ISNULL(TSPL_MILK_COLLECTION_BMCDCS.Posted_Date, '')<> '' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then

                        qry += " (select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"

                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MILK_COLLECTION_BMCDCS.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_COLLECTION_BMCDCS.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MILK_COLLECTION_BMCDCS.IDate , TSPL_MILK_COLLECTION_BMCDCS.PK_ID "

                End If
            End If

        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "MCC Procurement") <> CompairStringResult.Equal Then
                gv1.MasterTemplate.Columns("Description").IsVisible = False
            End If


            'Me.gv1.Columns("RefDocNo").IsVisible = False
            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling --------Module(Receivable)-------------------
    ''-------------------------------------------------------------------

    Sub FillReceivables()

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Receipt Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.ReceiptEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_RECEIPT_HEADER.Posted ='Y' then 1 else 0 end)as BIT) as Status, TSPL_RECEIPT_HEADER.Receipt_No as [Document Id], TSPL_RECEIPT_HEADER.Receipt_Date as [Document Date], isnull(TSPL_RECEIPT_HEADER.Receipt_Amount,0) as [Amount], TSPL_RECEIPT_HEADER.Cust_Code as [Customer Code],TSPL_RECEIPT_HEADER.Customer_Name as [Customer Name], TSPL_BANK_MASTER.BANK_CODE as [Bank Code], right(TSPL_BANK_MASTER.BANKACC,3) as Location, TSPL_RECEIPT_HEADER.Payment_Code as [Payment Code],TSPL_RECEIPT_HEADER.Cheque_No as [Cheque No],TSPL_RECEIPT_HEADER.Cheque_Date as [Cheque Date], TSPL_RECEIPT_HEADER.Entry_Desc as [Description],TSPL_RECEIPT_HEADER.Created_By as 'Created By' FROM TSPL_RECEIPT_HEADER Left Outer Join TSPL_BANK_MASTER on TSPL_RECEIPT_HEADER.Bank_Code=TSPL_BANK_MASTER.BANK_CODE WHERE  convert(date,TSPL_RECEIPT_HEADER.Receipt_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_RECEIPT_HEADER.Receipt_Date  ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_RECEIPT_HEADER.Posted IS NULL OR TSPL_RECEIPT_HEADER.Posted = 'N')"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_RECEIPT_HEADER.Posted = 'Y'"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_RECEIPT_HEADER.Location_GL_Code   in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_RECEIPT_HEADER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_RECEIPT_HEADER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_RECEIPT_HEADER.Receipt_Date , TSPL_RECEIPT_HEADER.Receipt_No "
                End If

            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Adjustment Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.ReceiptAdjustmentEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_Receipt_Adjustment_Header.Is_Post='Y' then 1 else 0 end)as BIT) as Status, TSPL_Receipt_Adjustment_Header.Adjustment_No as [Document Id], TSPL_Receipt_Adjustment_Header.Adjustment_Date as [Document Date], isnull(TSPL_Receipt_Adjustment_Header.Adjustment_Amount,0) as [Amount], TSPL_SALE_INVOICE_HEAD.Location , TSPL_Receipt_Adjustment_Header.Customer_No as [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name], TSPL_Receipt_Adjustment_Header.Description  as [Description],TSPL_Receipt_Adjustment_Header.Created_By as 'Created By' FROM TSPL_Receipt_Adjustment_Header Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_Receipt_Adjustment_Header.Customer_No=TSPL_CUSTOMER_MASTER.Cust_Code Left Outer Join TSPL_SALE_INVOICE_HEAD ON TSPL_Receipt_Adjustment_Header.Doc_No=TSPL_SALE_INVOICE_HEAD.Sale_Invoice_No WHERE  convert(date,TSPL_Receipt_Adjustment_Header.Adjustment_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_Receipt_Adjustment_Header.Adjustment_Date   ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_Receipt_Adjustment_Header.Is_Post is NULL OR TSPL_Receipt_Adjustment_Header.Is_Post = 'N')"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_Receipt_Adjustment_Header.Is_Post = 'Y'"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += "and TSPL_Receipt_Adjustment_Header.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += "and TSPL_Receipt_Adjustment_Header.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_Receipt_Adjustment_Header.Adjustment_Date , TSPL_Receipt_Adjustment_Header.Adjustment_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "AR Invoice Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.mbtnARInvoiceEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_Customer_Invoice_Head.status =1 then 1 else 0 end)as BIT) as Status, TSPL_Customer_Invoice_Head.Document_No  as [Document Id],TSPL_Customer_Invoice_Head.Document_Date as [Document Date],isnull(TSPL_Customer_Invoice_Head.Document_Total,0) as [Amount], TSPL_Customer_Invoice_Head.Customer_code as [Customer Code],TSPL_Customer_Invoice_Head.Customer_Name  as [Customer Name], TSPL_Customer_Invoice_Head.Description  as [Description],TSPL_Customer_Invoice_Head.Created_By as 'Created By' FROM TSPL_Customer_Invoice_Head WHERE  convert(date,TSPL_Customer_Invoice_Head.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_Customer_Invoice_Head.Document_Date   ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_Customer_Invoice_Head.status is NULL OR TSPL_Customer_Invoice_Head.status = 0) "
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_Customer_Invoice_Head.status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_Customer_Invoice_Head.Loc_Code  in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_Customer_Invoice_Head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_Customer_Invoice_Head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_Customer_Invoice_Head.Document_Date , TSPL_Customer_Invoice_Head.Document_No "


                End If
            End If
        End If

        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If


        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Filling --------Module(Common Services)--------------
    ''-------------------------------------------------------------------

    Sub FillCommonServices()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bank Transfer") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.bankTransfer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_BANK_TRANSFER.Post ='P' then 1 else 0 end)as BIT) as Status, TSPL_BANK_TRANSFER.Transfer_No as [Document Id], TSPL_BANK_TRANSFER.Transfer_Date as [Document Date],isnull(TSPL_BANK_TRANSFER.Transfer_Amount,0) as [Amount], TSPL_BANK_TRANSFER.From_Bank_Code as [From-Bank Code], TSPL_BANK_TRANSFER.From_Bank_Name as [From-Bank Name], TSPL_BANK_TRANSFER.From_Bank_Acc_No as [From-Bank A/C No], TSPL_BANK_TRANSFER.To_Bank_Code as [To-Bank Code], TSPL_BANK_TRANSFER.To_Bank_Name as [To-Bank Name], TSPL_BANK_TRANSFER.From_Bank_Acc_No as [From-Bank A/C No], TSPL_BANK_TRANSFER.Description as [Description],TSPL_BANK_TRANSFER.Created_By as 'Created By' FROM TSPL_BANK_TRANSFER WHERE  convert(date,TSPL_BANK_TRANSFER.Transfer_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_BANK_TRANSFER.Transfer_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_BANK_TRANSFER.Post = 'N'"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_BANK_TRANSFER.Post = 'P'"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_BANK_TRANSFER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_BANK_TRANSFER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"

                        End If
                    End If
                    qry += " ORDER BY TSPL_BANK_TRANSFER.Transfer_Date, TSPL_BANK_TRANSFER.Transfer_No "

                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Reverse Transaction") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.reverseTransaction)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when TSPL_BANK_REVERSE.Post ='P' then 1 else 0 end)as BIT) as Status, TSPL_BANK_REVERSE.Reverse_Code as [Document Id], TSPL_BANK_REVERSE.Reversal_Date as [Document Date],isnull(TSPL_BANK_REVERSE.Amount,0) as [Amount], TSPL_BANK_REVERSE.Bank_Code as [Bank Code], TSPL_BANK_REVERSE.Back_Acc_No as [Back A/C No], TSPL_BANK_REVERSE.Vendor_Code as [Vendor Code], TSPL_BANK_REVERSE.Vendor_Name as [Vendor Name], TSPL_BANK_REVERSE.Cust_Code as [Customer Code], TSPL_BANK_REVERSE.Cust_Name as [Customer Name], TSPL_BANK_REVERSE.Cheque_No as [Cheque No],'' as [Description],TSPL_BANK_REVERSE.Created_By as 'Created BY' FROM TSPL_BANK_REVERSE WHERE  convert(date,TSPL_BANK_REVERSE.Reversal_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_BANK_REVERSE.Reversal_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and (TSPL_BANK_REVERSE.Post = 'N' OR TSPL_BANK_REVERSE.Post IS NULL)"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_BANK_REVERSE.Post = 'P'"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_BANK_REVERSE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_BANK_REVERSE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_BANK_REVERSE.Reversal_Date, TSPL_BANK_REVERSE.Reverse_Code "

                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    Sub FillHRAndPayroll()
        '' fill Employee Salary
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Employee Salary") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmEmpSalary)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,EMP_SAL_CODE as [Document Id],APPLICABLE_FROM as [Document Date],'' AS [Description],TSPL_EMPLOYEE_SALARY.EMP_CODE as [Employee Code]," &
                          " TSPL_EMPLOYEE_MASTER.Emp_Name AS [Employee Name],REVISION_NO as [Revision No],SALARY_STRUCTURE_CODE as [Salary Structure], " &
                          " TSPL_EMPLOYEE_SALARY.Created_By as [Created By],TSPL_EMPLOYEE_SALARY.Created_Date as [Created Date] " &
                          " FROM TSPL_EMPLOYEE_SALARY left join TSPL_EMPLOYEE_MASTER ON TSPL_EMPLOYEE_SALARY.EMP_CODE=TSPL_EMPLOYEE_MASTER.EMP_CODE " &
                          " WHERE  convert(date,TSPL_EMPLOYEE_SALARY.APPLICABLE_FROM ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_EMPLOYEE_SALARY.APPLICABLE_FROM,103) <= convert(date,'" + dtpToDate.Value + "',103)"

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_EMPLOYEE_SALARY.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_EMPLOYEE_SALARY.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_EMPLOYEE_SALARY.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_EMPLOYEE_SALARY.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_EMPLOYEE_SALARY.APPLICABLE_FROM, TSPL_EMPLOYEE_SALARY.EMP_SAL_CODE "

                End If
            End If

            '' fill Employee Hourly Attendance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Hourly Attendance") = CompairStringResult.Equal Then
            Load_Authorisation("HOURL_ATTEND")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,DLA_CODE as [Document Id],'' AS [Description],PAY_PERIOD_CODE as [Pay Period Code],REGISTER_TYPE as [Register Type]," &
                          " TSPL_HOURLY_ATTENDANCE.Created_Date as [Document Date],ENTEREDBY_EMP_CODE as [Entered By],TSPL_HOURLY_ATTENDANCE.Created_By as [Created By],TSPL_HOURLY_ATTENDANCE.Created_Date as [Created Date] FROM TSPL_HOURLY_ATTENDANCE  " &
                          " WHERE  convert(date,TSPL_HOURLY_ATTENDANCE.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_HOURLY_ATTENDANCE.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_HOURLY_ATTENDANCE.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_HOURLY_ATTENDANCE.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_HOURLY_ATTENDANCE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_HOURLY_ATTENDANCE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_HOURLY_ATTENDANCE.Created_Date, TSPL_HOURLY_ATTENDANCE.DLA_CODE "

                End If
            End If

            '' fill Employee Daily Attendance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Daily Attendance") = CompairStringResult.Equal Then
            Load_Authorisation("DAILY_ATTEND")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,DLA_CODE as [Document Id],'' AS [Description],PAY_PERIOD_CODE as [Pay Period Code],REGISTER_TYPE as [Register Type]," &
                          " TSPL_DAILY_ATTENDANCE.Created_Date as [Document Date],ENTEREDBY_EMP_CODE as [Entered By],TSPL_DAILY_ATTENDANCE.Created_By as [Created By],TSPL_DAILY_ATTENDANCE.Created_Date as [Created Date] FROM TSPL_DAILY_ATTENDANCE  " &
                          " WHERE  convert(date,TSPL_DAILY_ATTENDANCE.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_DAILY_ATTENDANCE.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_DAILY_ATTENDANCE.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_DAILY_ATTENDANCE.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_DAILY_ATTENDANCE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_DAILY_ATTENDANCE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_DAILY_ATTENDANCE.Created_Date, TSPL_DAILY_ATTENDANCE.DLA_CODE "

                End If
            End If

            '' fill Employee Monthly Attendance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Monthly Attendance") = CompairStringResult.Equal Then
            Load_Authorisation("MONTH_ATTEND")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,MTA_CODE as [Document Id],'' AS [Description],PAY_PERIOD_CODE as [Pay Period Code],REGISTER_TYPE as [Register Type]," &
                          " TSPL_MONTHLY_ATTENDANCE.Created_Date as [Document Date],ENTEREDBY_EMP_CODE as [Entered By],TSPL_MONTHLY_ATTENDANCE.Created_By as [Created By],TSPL_MONTHLY_ATTENDANCE.Created_Date as [Created Date] FROM TSPL_MONTHLY_ATTENDANCE  " &
                          " WHERE  convert(date,TSPL_MONTHLY_ATTENDANCE.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MONTHLY_ATTENDANCE.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MONTHLY_ATTENDANCE.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MONTHLY_ATTENDANCE.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MONTHLY_ATTENDANCE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MONTHLY_ATTENDANCE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MONTHLY_ATTENDANCE.Created_Date, TSPL_MONTHLY_ATTENDANCE.MTA_CODE "

                End If
            End If
            '' fill Employee OT Sheet
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "OT Sheet") = CompairStringResult.Equal Then
            Load_Authorisation("OT-SHEET")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,OT_Sheet_Code as [Document Id],'' AS [Description],PAY_PERIOD_CODE as [Pay Period Code]," &
                          " tspl_ot_sheet.Created_Date as [Document Date],tspl_ot_sheet.Created_By as [Created By],tspl_ot_sheet.Created_Date as [Created Date] FROM tspl_ot_sheet  " &
                          " WHERE  convert(date,tspl_ot_sheet.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,tspl_ot_sheet.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and tspl_ot_sheet.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and tspl_ot_sheet.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and tspl_ot_sheet.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and tspl_ot_sheet.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY tspl_ot_sheet.Created_Date, tspl_ot_sheet.OT_Sheet_Code "

                End If
            End If
            '' fill Employee Allowance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Allowance") = CompairStringResult.Equal Then
            Load_Authorisation("Allowance")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT TSPL_ALLOWANCE.POSTED as Status,TSPL_ALLOWANCE.ALLOWANCE_CODE as [Document Id],TSPL_ALLOWANCE.PAY_PERIOD_CODE as [Pay Period Code], " &
                          " TSPL_ALLOWANCE.ALLOWANCE_DATE as [Document Date],TSPL_ALLOWANCE.EMP_CODE as [Employee Code],TSPL_ALLOWANCE.ALLOWANCE_REMARKS as [Description],  " &
                          " SUM(isnull(TSPL_ALLOWANCE_DETAIL.Allowance_Amount,0)) as Amount,TSPL_ALLOWANCE.Created_By as [Created By],TSPL_ALLOWANCE.Created_Date as [Created Date] FROM TSPL_ALLOWANCE LEFT OUTER JOIN TSPL_ALLOWANCE_DETAIL ON TSPL_ALLOWANCE.ALLOWANCE_CODE=TSPL_ALLOWANCE_DETAIL.ALLOWANCE_CODE " &
                          " WHERE  convert(date,TSPL_ALLOWANCE.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_ALLOWANCE.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_ALLOWANCE.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_ALLOWANCE.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_ALLOWANCE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ALLOWANCE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " GROUP BY TSPL_ALLOWANCE.ALLOWANCE_CODE,TSPL_ALLOWANCE.POSTED,TSPL_ALLOWANCE.PAY_PERIOD_CODE,TSPL_ALLOWANCE.ALLOWANCE_DATE, TSPL_ALLOWANCE.EMP_CODE,TSPL_ALLOWANCE.ALLOWANCE_REMARKS,TSPL_ALLOWANCE.Created_By,TSPL_ALLOWANCE.Created_Date"
                    qry += " ORDER BY TSPL_ALLOWANCE.ALLOWANCE_DATE, TSPL_ALLOWANCE.ALLOWANCE_CODE "
                End If
            End If

            '' fill Employee Deduction
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Deduction") = CompairStringResult.Equal Then
            Load_Authorisation("Deduction")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT TSPL_DEDUCTION.POSTED as Status,TSPL_DEDUCTION.DEDUCTION_CODE as [Document Id],TSPL_DEDUCTION.PAY_PERIOD_CODE as [Pay Period Code], " &
                          " TSPL_DEDUCTION.DEDUCTION_DATE as [Document Date],TSPL_DEDUCTION.EMP_CODE as [Employee Code],TSPL_DEDUCTION.DEDUCTION_REMARKS as [Description], " &
                          " SUM(isnull(TSPL_DEDUCTION_DETAIL.DEDUCTION_AMOUNT ,0)) as Amount,TSPL_DEDUCTION.Created_By as [Created By],TSPL_DEDUCTION.Created_Date as [Created Date] FROM TSPL_DEDUCTION LEFT OUTER JOIN TSPL_DEDUCTION_DETAIL ON TSPL_DEDUCTION.DEDUCTION_CODE=TSPL_DEDUCTION_DETAIL.DEDUCTION_CODE " &
                          " WHERE  convert(date,TSPL_DEDUCTION.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_DEDUCTION.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_DEDUCTION.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_DEDUCTION.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_DEDUCTION.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_DEDUCTION.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " GROUP BY TSPL_DEDUCTION.DEDUCTION_CODE,TSPL_DEDUCTION.POSTED,TSPL_DEDUCTION.PAY_PERIOD_CODE,TSPL_DEDUCTION.DEDUCTION_DATE, TSPL_DEDUCTION.EMP_CODE,TSPL_DEDUCTION.DEDUCTION_REMARKS,TSPL_DEDUCTION.Created_By,TSPL_DEDUCTION.Created_Date"
                    qry += " ORDER BY TSPL_DEDUCTION.DEDUCTION_DATE, TSPL_DEDUCTION.DEDUCTION_CODE "
                End If
            End If

            '' fill Employee Bonus
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bonus") = CompairStringResult.Equal Then
            Load_Authorisation("Bonus")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT TSPL_EMPLOYEE_BONUS.POSTED as Status,TSPL_EMPLOYEE_BONUS.EMP_BONUS_CODE as [Document Id],TSPL_EMPLOYEE_BONUS.FROM_PAY_PERIOD_CODE as [From Pay Period],TSPL_EMPLOYEE_BONUS.TO_PAY_PERIOD_CODE as [To Pay Period], " &
                          " TSPL_EMPLOYEE_BONUS.PAYABLE_PAY_PERIOD_CODE as [Payable Pay Period],TSPL_EMPLOYEE_BONUS.Created_Date as [Document Date],TSPL_EMPLOYEE_BONUS.DESCRIPTION as [Description], SUM(isnull(TSPL_BONUS_GENERATION_DETAIL.ACTUAL_AMOUNT  ,0)) as Amount," &
                          " TSPL_EMPLOYEE_BONUS.Created_By as [Created By],TSPL_EMPLOYEE_BONUS.Created_Date as [Created Date] FROM TSPL_EMPLOYEE_BONUS LEFT OUTER JOIN TSPL_BONUS_GENERATION_DETAIL ON TSPL_EMPLOYEE_BONUS.EMP_BONUS_CODE=TSPL_BONUS_GENERATION_DETAIL.EMP_BONUS_CODE" &
                          " WHERE  convert(date,TSPL_EMPLOYEE_BONUS.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_EMPLOYEE_BONUS.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_EMPLOYEE_BONUS.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_EMPLOYEE_BONUS.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_EMPLOYEE_BONUS.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_EMPLOYEE_BONUS.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"

                        End If
                    End If
                    qry += " GROUP BY TSPL_EMPLOYEE_BONUS.EMP_BONUS_CODE,TSPL_EMPLOYEE_BONUS.POSTED,TSPL_EMPLOYEE_BONUS.FROM_PAY_PERIOD_CODE,TSPL_EMPLOYEE_BONUS.TO_PAY_PERIOD_CODE, TSPL_EMPLOYEE_BONUS.PAYABLE_PAY_PERIOD_CODE,TSPL_EMPLOYEE_BONUS.Created_Date,TSPL_EMPLOYEE_BONUS.DESCRIPTION,TSPL_EMPLOYEE_BONUS.Created_By"
                    qry += " ORDER BY TSPL_EMPLOYEE_BONUS.Created_Date, TSPL_EMPLOYEE_BONUS.EMP_BONUS_CODE "
                End If
            End If

            '' fill Employee Salary Adjustment
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Salary Adjustment") = CompairStringResult.Equal Then

            Load_Authorisation("Salary Adjustment")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT TSPL_ADJUSTMENT_VOUCHER.POSTED as Status,TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_CODE as [Document Id],TSPL_ADJUSTMENT_VOUCHER.PAY_PERIOD_CODE as [Pay Period Code], " &
                          " TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_DATE as [Document Date],TSPL_ADJUSTMENT_VOUCHER.EMP_CODE as [Employee Code], " &
                          " TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_REMARK as [Description], SUM(isnull(TSPL_EMPADJUSTMENT_DETAIL.ADJUSTMENT_PLUS ,0)+isnull(TSPL_EMPADJUSTMENT_DETAIL.ADJUSTMENT_MINUS ,0)) as Amount," &
                          " TSPL_ADJUSTMENT_VOUCHER.Created_By as [Created By],TSPL_ADJUSTMENT_VOUCHER.Created_Date as [Created Date] FROM TSPL_ADJUSTMENT_VOUCHER LEFT OUTER JOIN TSPL_EMPADJUSTMENT_DETAIL ON TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_CODE=TSPL_EMPADJUSTMENT_DETAIL.ADJUSTMENT_CODE" &
                          " WHERE  convert(date,TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_ADJUSTMENT_VOUCHER.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_ADJUSTMENT_VOUCHER.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_ADJUSTMENT_VOUCHER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ADJUSTMENT_VOUCHER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " GROUP BY TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_CODE,TSPL_ADJUSTMENT_VOUCHER.POSTED,TSPL_ADJUSTMENT_VOUCHER.PAY_PERIOD_CODE, TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_DATE, tSPL_ADJUSTMENT_VOUCHER.EMP_CODE, TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_REMARK, TSPL_ADJUSTMENT_VOUCHER.Created_By, TSPL_ADJUSTMENT_VOUCHER.Created_Date"
                    qry += " ORDER BY TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_DATE, TSPL_ADJUSTMENT_VOUCHER.ADJUSTMENT_CODE "
                End If
            End If

            '' fill Employee Reimbursement
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Reimbursement") = CompairStringResult.Equal Then
            Load_Authorisation("Reimbursement")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,REIMBURSEMENT_CODE as [Document Id],PAY_PERIOD_CODE as [Pay Period Code], " &
                          " TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_DATE as [Document Date],TSPL_EMP_REIMBURSEMENT.EMP_CODE as [Employee Code], " &
                          " isnull(TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_AMOUNT,0) as [Amount],TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_REMARK as [Description], " &
                          " TSPL_EMP_REIMBURSEMENT.Created_By as [Created By],TSPL_EMP_REIMBURSEMENT.Created_Date as [Created Date] FROM TSPL_EMP_REIMBURSEMENT " &
                          " WHERE  convert(date,TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_EMP_REIMBURSEMENT.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_EMP_REIMBURSEMENT.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_EMP_REIMBURSEMENT.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_EMP_REIMBURSEMENT.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"

                        End If
                    End If
                    qry += " ORDER BY TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_DATE, TSPL_EMP_REIMBURSEMENT.REIMBURSEMENT_CODE "

                End If
            End If
            '' fill Employee Loan Application
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Loan Application") = CompairStringResult.Equal Then
            Load_Authorisation("Loan Application")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,LOAN_CODE as [Document Id],TSPL_LOAN_APPLICATION.LOAN_DATE as [Document Date],TSPL_LOAN_APPLICATION.EMP_CODE as [Employee Code], " &
                          " TSPL_LOAN_APPLICATION.LOAN_TYPE as [Loan Type],isnull(TSPL_LOAN_APPLICATION.LOAN_AMOUNT,0) as [Amount], " &
                          " TSPL_LOAN_APPLICATION.PAYMENT_STARTDATE as [Payment Start Date],TSPL_LOAN_APPLICATION.NO_OF_EMI as [No of EMI], " &
                          " TSPL_LOAN_APPLICATION.INTEREST_APPLIED as [Interest Applied],TSPL_LOAN_APPLICATION.INTEREST_TYPE as [Interest Type], " &
                          " TSPL_LOAN_APPLICATION.INTEREST_RATE as [Interest Rate],isnull(TSPL_LOAN_APPLICATION.INTEREST_AMOUNT,0) as [Interest Amount]," &
                          " isnull(TSPL_LOAN_APPLICATION.TOTALPAYABLE_AMOUNT,0) as [Total Payable Amount],TSPL_LOAN_APPLICATION.LOAN_DESCRIPTION as [Description]," &
                          " TSPL_LOAN_APPLICATION.Created_By as [Created By],TSPL_LOAN_APPLICATION.Created_Date as [Created Date] FROM TSPL_LOAN_APPLICATION " &
                          " WHERE  convert(date,TSPL_LOAN_APPLICATION.LOAN_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_LOAN_APPLICATION.LOAN_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_LOAN_APPLICATION.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_LOAN_APPLICATION.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_LOAN_APPLICATION.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_LOAN_APPLICATION.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If

                    qry += " ORDER BY TSPL_LOAN_APPLICATION.LOAN_DATE, TSPL_LOAN_APPLICATION.LOAN_CODE "

                End If
            End If

            '' fill Employee Loan Adjustment
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Loan Adjustment") = CompairStringResult.Equal Then

            Load_Authorisation("Loan Adjustment")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,LOANADJUSTMENT_CODE as [Document Id],TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_DATE as [Document Date], " &
                          " TSPL_LOAN_ADJUSTMENT.EMP_CODE as [Employee Code],TSPL_LOAN_ADJUSTMENT.PAY_PERIOD_CODE as [Pay Period], " &
                          " isnull(TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_PLUS,0) as [Adjustment Plus],isnull(TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_MINUS,0) as [Adjustment Minus], " &
                          " TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_REASON as [Description], " &
                          " TSPL_LOAN_ADJUSTMENT.Created_By as [Created By],TSPL_LOAN_ADJUSTMENT.Created_Date as [Created Date] FROM TSPL_LOAN_ADJUSTMENT " &
                          " WHERE  convert(date,TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_LOAN_ADJUSTMENT.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_LOAN_ADJUSTMENT.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_LOAN_ADJUSTMENT.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_LOAN_ADJUSTMENT.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_LOAN_ADJUSTMENT.ADJUSTMENT_DATE, TSPL_LOAN_ADJUSTMENT.LOANADJUSTMENT_CODE "

                End If
            End If

            '' fill Employee Leave Application
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Leave Application") = CompairStringResult.Equal Then
            Load_Authorisation("Leave Application")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,LVAPPLICATION_CODE as [Document Id],TSPL_LEAVE_APPLICATION.APPLICATION_DATE as [Document Date], " &
                          " TSPL_LEAVE_APPLICATION.EMP_CODE as [Employee Code],TSPL_LEAVE_APPLICATION.PAY_PERIOD_CODE as [Pay Period], " &
                          " TSPL_LEAVE_APPLICATION.LEAVE_CODE as [Leave Code],TSPL_LEAVE_APPLICATION.FROM_DATE as [From Date],TSPL_LEAVE_APPLICATION.TO_DATE as [To Date], " &
                          " TSPL_LEAVE_APPLICATION.TOTAL_DAYS as [Total Days],TSPL_LEAVE_APPLICATION.LEAVE_REASON as [Description], " &
                          " TSPL_LEAVE_APPLICATION.Created_By as [Created By],TSPL_LEAVE_APPLICATION.Created_Date as [Created Date] FROM TSPL_LEAVE_APPLICATION  " &
                          " WHERE  convert(date,TSPL_LEAVE_APPLICATION.APPLICATION_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_LEAVE_APPLICATION.APPLICATION_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_LEAVE_APPLICATION.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_LEAVE_APPLICATION.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_LEAVE_APPLICATION.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_LEAVE_APPLICATION.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_LEAVE_APPLICATION.APPLICATION_DATE, TSPL_LEAVE_APPLICATION.LVAPPLICATION_CODE "

                End If
            End If

            '' fill Employee Leave Adjustment
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Leave Adjustment") = CompairStringResult.Equal Then
            Load_Authorisation("Leave Adjustment")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " SELECT POSTED as Status,LVADJUSTMENT_CODE as [Document Id],TSPL_LEAVE_ADJUSTMENT.ADJUSTMENT_DATE as [Document Date], " &
                          " TSPL_LEAVE_ADJUSTMENT.EMP_CODE as [Employee Code],TSPL_LEAVE_ADJUSTMENT.PAY_PERIOD_CODE as [Pay Period], " &
                          " TSPL_LEAVE_ADJUSTMENT.LEAVE_CODE as [Leave Code],TSPL_LEAVE_ADJUSTMENT.ADJUST_ALLOTED as [Adjustment in Alloted], " &
                          " TSPL_LEAVE_ADJUSTMENT.ADJUST_AVAILED as [Adjustment in Availed],TSPL_LEAVE_ADJUSTMENT.LEAVE_REASON as [Description], " &
                          " TSPL_LEAVE_ADJUSTMENT.Created_By as [Created By],TSPL_LEAVE_ADJUSTMENT.Created_Date as [Created Date] FROM TSPL_LEAVE_ADJUSTMENT " &
                          " WHERE  convert(date,TSPL_LEAVE_ADJUSTMENT.ADJUSTMENT_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_LEAVE_ADJUSTMENT.ADJUSTMENT_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_LEAVE_ADJUSTMENT.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_LEAVE_ADJUSTMENT.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_LEAVE_ADJUSTMENT.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_LEAVE_ADJUSTMENT.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_LEAVE_ADJUSTMENT.ADJUSTMENT_DATE, TSPL_LEAVE_ADJUSTMENT.LVADJUSTMENT_CODE "
                End If
            End If
            '==Shivani Tyagi===against[BM00000007827]
            '' fill Employee Increment
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Employee Increment") = CompairStringResult.Equal Then
            Load_Authorisation("Employee Increment")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select TSPL_EMPLOYEE_INCREMENT_HEAD.POSTED as status ,TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_CODE as [Document Id] ,convert(varchar,TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_DATE,103) as [Document Date],'' as [Description],TSPL_EMPLOYEE_INCREMENT_HEAD.EMP_CODE as [Employee Code],TSPL_EMPLOYEE_MASTER.Emp_Name as [Employee Name] ,TSPL_EMPLOYEE_INCREMENT_HEAD.REVISION_NO as [Revision No],TSPL_EMPLOYEE_INCREMENT_HEAD.SALARY_STRUCTURE_CODE as [Salary Structure],SUM(convert(decimal,isnull(TSPL_EMPLOYEE_INCREMENT_DETAIL.IncrementAmt,0))) AS Amount,TSPL_EMPLOYEE_INCREMENT_HEAD.Created_By as [Created By] ,convert(varchar,TSPL_EMPLOYEE_INCREMENT_HEAD.Created_Date,103) as [Created Date]  from TSPL_EMPLOYEE_INCREMENT_HEAD left join tspl_employee_master on tspl_employee_master.EMP_CODE =TSPL_EMPLOYEE_INCREMENT_HEAD.EMP_CODE LEFT OUTER JOIN TSPL_EMPLOYEE_INCREMENT_DETAIL ON TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_CODE=TSPL_EMPLOYEE_INCREMENT_DETAIL.INCREMENT_CODE  " &
                          " WHERE  convert(date,TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103)"

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_EMPLOYEE_INCREMENT_HEAD.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_EMPLOYEE_INCREMENT_HEAD.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_EMPLOYEE_INCREMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_EMPLOYEE_INCREMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " GROUP BY TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_CODE,TSPL_EMPLOYEE_INCREMENT_HEAD.POSTED,TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_DATE, TSPL_EMPLOYEE_INCREMENT_HEAD.EMP_CODE,TSPL_EMPLOYEE_MASTER.Emp_Name,TSPL_EMPLOYEE_INCREMENT_HEAD.REVISION_NO,TSPL_EMPLOYEE_INCREMENT_HEAD.SALARY_STRUCTURE_CODE,TSPL_EMPLOYEE_INCREMENT_HEAD.Created_By,TSPL_EMPLOYEE_INCREMENT_HEAD.Created_Date"
                    qry += " ORDER BY TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_DATE, TSPL_EMPLOYEE_INCREMENT_HEAD.INCREMENT_CODE "
                End If
            End If
        End If

        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

    Sub FillProduction()
        '' fill Employee Salary
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Batch Order") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmBatchOrderSTD)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " Select POSTED as Status,BO_CODE as[Document Id],BO_Date as [Document Date], Description," &
                           "APPROVED_BY [Approved By],Created_By as [Created By],Created_Date as [Created Date] " &
                           "from TSPL_MF_BATCH_ORDER" &
                          " WHERE  convert(date,BO_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,BO_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_BATCH_ORDER.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_BATCH_ORDER.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MF_BATCH_ORDER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_BATCH_ORDER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_BATCH_ORDER.BO_Date, TSPL_MF_BATCH_ORDER.BO_CODE "

                End If
            End If
            '' fill Employee Hourly Attendance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bill Of Material") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmBillOfMaterialCosting)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " select POSTED as Status,BOM_CODE as[Document Id],BOM_Date as [Document Date],'' as Description," &
                          " REVISION_NO [Revision No],PROD_ITEM_CODE as [Item Code],Prod_quantity as [Quatinty]," &
                          " prod_item_unit_code as [Item Unit Code]," &
                          " Created_By as [Created By],Created_Date as [Created Date]" &
                          " from TSPL_MF_BOM_HEAD " &
                          " WHERE  convert(date,TSPL_MF_BOM_HEAD.BOM_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_BOM_HEAD.BOM_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_BOM_HEAD.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_BOM_HEAD.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MF_BOM_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_BOM_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_BOM_HEAD.Created_Date, TSPL_MF_BOM_HEAD.BOM_CODE "

                End If
            End If

            '' fill Employee Daily Attendance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Manufacturing Order") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmManufacturingOrder)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "  select POSTED as Status,MO_CODE as[Document Id],MO_DATE as [Document Date],'' as Description," &
                          " ITEM_CODE as [Item Code],QTY_ORDERED as [Qty Ordered],UNIT_CODE as [Unit Code], " &
                          " QTY_ORDERED_STOCK as [Qty Ordered Stock],UNIT_CODE_STOCK as [Unit Code Stock], " &
                          " Description,BOM_CODE as [Batch of Material Code],MO_DUE_DATE as [MO Due Date], " &
                          " PRODUCTION_AREA as [Production Area],PLANNER Planner,IN_CHARGE as [In Charge], " &
                          " Created_By as [Created By],Created_Date as [Created Date]  " &
                          " from TSPL_MF_MANUFACTURING_ORDER " &
                          " WHERE  convert(date,TSPL_MF_MANUFACTURING_ORDER.MO_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_MANUFACTURING_ORDER.MO_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_MANUFACTURING_ORDER.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_MANUFACTURING_ORDER.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MF_MANUFACTURING_ORDER.LOCATION_CODE in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MF_MANUFACTURING_ORDER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_MANUFACTURING_ORDER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_MANUFACTURING_ORDER.Created_Date, TSPL_MF_MANUFACTURING_ORDER.MO_CODE "

                End If
            End If

            '' fill Employee Monthly Attendance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Planning") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProductionPlanningSTD)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " select POSTED as Status,PROD_PLAN_CODE as[Document Id],PLAN_FOR_DATE as [Document Date]," &
                          " DESCRIPTION as [Description],COMMENTS as [Comments]," &
                          " Created_By as [Created By],Created_Date as [Created Date] " &
                          " from TSPL_MF_PRODUCTION_PLAN_HEAD " &
                          " WHERE  convert(date,TSPL_MF_PRODUCTION_PLAN_HEAD.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_PRODUCTION_PLAN_HEAD.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_PRODUCTION_PLAN_HEAD.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_PRODUCTION_PLAN_HEAD.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MF_PRODUCTION_PLAN_HEAD.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MF_PRODUCTION_PLAN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_PRODUCTION_PLAN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_PRODUCTION_PLAN_HEAD.Created_Date, TSPL_MF_PRODUCTION_PLAN_HEAD.PROD_PLAN_CODE "

                End If
            End If

            '' fill Employee Allowance
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Receipt") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProductionReceiptSTD)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " select POSTED as Status,RECEIPT_CODE as[Document Id],BATCH_DATE [Document Date]," &
                          " DESCRIPTION as [Description],BO_CODE [BO Code],RECEIPT_DATE [Receipt Date], " &
                          " Received_by [Received By],LOCATION_CODE [Location Code],COMMENTS as [Comments]," &
                          " Created_By as [Created By],Created_Date as [Created Date] " &
                          " from TSPL_MF_RECEIPT " &
                          " WHERE  convert(date,TSPL_MF_RECEIPT.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_RECEIPT.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_RECEIPT.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_RECEIPT.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MF_RECEIPT.LOCATION_CODE  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MF_RECEIPT.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_RECEIPT.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_RECEIPT.RECEIPT_DATE, TSPL_MF_RECEIPT.RECEIPT_CODE "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Entry") = CompairStringResult.Equal Then

            'Load_Authorisation(clsUserMgtCode.frmProductionEntryWithoutBatch)
            Load_Authorisation(clsUserMgtCode.frmProductionEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " select POSTED as Status,PROD_ENTRY_CODE as[Document Id],BATCH_DATE [Document Date]," &
                          " DESCRIPTION as [Description],Batch_Code [BO Code],PROD_DATE [Receipt Date], " &
                          " Received_by [Received By],LOCATION_CODE [Location Code],COMMENTS as [Comments], " &
                          " Created_By as [Created By],Created_Date as [Created Date] " &
                          " from TSPL_PP_PRODUCTION_ENTRY " &
                          " WHERE  convert(date,TSPL_PP_PRODUCTION_ENTRY.PROD_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_PP_PRODUCTION_ENTRY.PROD_DATE ,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_PP_PRODUCTION_ENTRY.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_PP_PRODUCTION_ENTRY.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_PP_PRODUCTION_ENTRY.LOCATION_CODE  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_PP_PRODUCTION_ENTRY.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_PP_PRODUCTION_ENTRY.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_PP_PRODUCTION_ENTRY.PROD_DATE, TSPL_PP_PRODUCTION_ENTRY.PROD_ENTRY_CODE "

                End If
            End If
            '' fill Employee Deduction
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Requisition") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProductionRequisition)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "  select POSTED as Status,REQ_CODE as[Document Id],REQ_DATE as [Document Date] ," &
                          " EXP_Date as [Expire Date],REQUESTED_BY as [Requested By]," &
                          " DESCRIPTION as [Description],LOCATION_CODE [Location Code],COMMENTS as [Comments]," &
                          " Created_By as [Created By],Created_Date as [Created Date] " &
                          " from TSPL_MF_REQUISITION " &
                          " WHERE  convert(date,TSPL_MF_REQUISITION.Created_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_REQUISITION.Created_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_REQUISITION.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_REQUISITION.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MF_REQUISITION.LOCATION_CODE  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MF_REQUISITION.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_REQUISITION.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_REQUISITION.REQ_DATE, TSPL_MF_REQUISITION.REQ_CODE "

                End If
            End If

            '' fill Employee Bonus
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Return") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProductionReturnSTD)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " select POSTED as Status,RETURN_CODE as[Document Id],RETURN_DATE as [Document Date] ," &
                          " EXP_Date as [Expire Date],RETURNED_BY as [Returned By],RETURNED_TO as [Returned To]," &
                          " DESCRIPTION as [Description],LOCATION_CODE [Location Code],COMMENTS as [Comments]," &
                          " Created_By as [Created By],Created_Date as [Created Date] " &
                          " from TSPL_MF_RETURN " &
                          " WHERE  convert(date,TSPL_MF_RETURN.RETURN_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_RETURN.RETURN_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_RETURN.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_RETURN.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MF_RETURN.LOCATION_CODE    in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MF_RETURN.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_RETURN.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_RETURN.RETURN_DATE, TSPL_MF_RETURN.RETURN_CODE "
                End If
            End If

            '' fill Employee Salary Adjustment
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Issue") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmStoreIssueSTD)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = " select POSTED as Status,ISSUE_CODE as[Document Id],ISSUE_DATE as [Document Date] ," &
                          " EXP_Date as [Expire Date],ISSUED_BY as [Issue By],ISSUED_TO as [Issue To]," &
                          " DESCRIPTION as [Description],LOCATION_CODE [Location Code],COMMENTS as [Comments]," &
                          " Created_By as [Created By],Created_Date as [Created Date] " &
                          " from TSPL_MF_ISSUE" &
                          " WHERE  convert(date,TSPL_MF_ISSUE.ISSUE_DATE ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_MF_ISSUE.ISSUE_DATE,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MF_ISSUE.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_MF_ISSUE.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MF_ISSUE.LOCATION_CODE    in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MF_ISSUE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MF_ISSUE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MF_ISSUE.ISSUE_DATE, TSPL_MF_ISSUE.ISSUE_CODE "

                End If
            End If
            ''richa GKD/02/09/19-000185
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Silo Milk Transfer") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmSiloMilkTransfer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "  select CAST(TSPL_SILO_MILK_TRANSFER_HEAD.POSTED as BIT )  as Status,Document_Code  as [Document Id],Document_Date as [Document Date] , DESCRIPTION as [Description],MainLocation_Code as [Location Code]," &
                        " Silo_Code as Silo,Item_Code as Item, Created_By as [Created By],Created_Date as [Created Date]  from TSPL_SILO_MILK_TRANSFER_HEAD " &
                          " WHERE  convert(date,TSPL_SILO_MILK_TRANSFER_HEAD.Document_Date  ,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                          " and convert(date,TSPL_SILO_MILK_TRANSFER_HEAD.Document_Date  ,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SILO_MILK_TRANSFER_HEAD.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += " and TSPL_SILO_MILK_TRANSFER_HEAD.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SILO_MILK_TRANSFER_HEAD.MainLocation_Code    in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_SILO_MILK_TRANSFER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SILO_MILK_TRANSFER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_SILO_MILK_TRANSFER_HEAD.Document_Date, TSPL_SILO_MILK_TRANSFER_HEAD.Document_Code "

                End If
            End If

        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            '----Added by--Pankaj Kummar--on-08/06/2012--For Counting Records on Each Refresh Operation-----
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub
    '===========Sanjeet(27/12/2016)============================
    Sub FillMccProcurement()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VSP Item Issue") = CompairStringResult.Equal Then
            'Load_Authorisation(clsUserMgtCode.frmVSPItemIssue)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(Status AS BIT) as Status,DOC_NO as [Document Id],DOC_DATE as [Document Date],To_Location as [Location Code],Remarks as [Description],Comment ,On_Hold AS Hold, Created_By as [Created By],Created_Date as [Created Date]  from TSPL_VSPItem_HEAD " &
                        " where  convert(date,DOC_DATE,103)>= convert(date,'" + dtpFromDate.Value + "',103) " &
                    "and convert(date,DOC_DATE,103) <=convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_VSPItem_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_VSPItem_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_VSPItem_HEAD.From_Location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_VSPItem_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_VSPItem_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_VSPItem_HEAD.DOC_DATE, TSPL_VSPItem_HEAD.DOC_NO "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VSP Asset Issue") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmVSPAssetIssue)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(Status AS BIT) as Status,DOC_NO as [Document Id],DOC_DATE as [Document Date],To_Location as [Location Code],Remarks as [Description],Comment ,On_Hold AS Hold, Created_By as [Created By],Created_Date as [Created Date]  from TSPL_VSPAsset_HEAD " &
                    " where  convert(date,DOC_DATE,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                    "and convert(date,DOC_DATE,103) <=convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_VSPAsset_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_VSPAsset_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_VSPAsset_HEAD.From_Location in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_VSPAsset_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_VSPAsset_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_VSPAsset_HEAD.DOC_DATE, TSPL_VSPAsset_HEAD.DOC_NO "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmMCCMaterial)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(Status AS BIT) AS Status,Document_Code as [Document Id],Document_Date as [Document Date],Total_Amt as [Amount],DESCRIPTION as [Description], " &
                    "Bill_To_Location as [Location Code],Created_By as [Created By],Created_Date as [Created Date]  from TSPL_SD_SHIPMENT_HEAD " &
                    "where Trans_Type='MCC' and convert(date,Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) " &
                    "and convert(date,Document_Date,103) <=convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_SD_SHIPMENT_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Bill_To_Location in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_SD_SHIPMENT_HEAD.Document_Date, TSPL_SD_SHIPMENT_HEAD.Document_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale Return") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmMCCMaterialSaleReturn)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(Status AS BIT) AS Status,Document_Code as [Document Id],Document_Date as [Document Date],DESCRIPTION as [Description]," &
                            "Bill_To_Location as [Location Code],Created_By as [Created By],Created_Date as [Created Date]  from TSPL_SD_SALE_RETURN_HEAD " &
                            "where Trans_Type='MCC' and convert(date,Document_Date,103)>= convert(date,'" + dtpFromDate.Value + "',103) " &
                    " and convert(date,Document_Date,103) <=convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_RETURN_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_SD_SALE_RETURN_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_SD_SALE_RETURN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SALE_RETURN_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_SD_SALE_RETURN_HEAD.Document_Date, TSPL_SD_SALE_RETURN_HEAD.Document_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Procurement Deduction") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmProcurementDeduction)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date, '')= '' then 0 else 1 end)as BIT) as Status, TSPL_VENDOR_INVOICE_HEAD.Document_No as [Document Id], TSPL_VENDOR_INVOICE_HEAD.invoice_entry_date as [Document Date],isnull(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as [Amount], TSPL_VENDOR_INVOICE_HEAD.Vendor_Code as [Vendor Code],TSPL_VENDOR_INVOICE_HEAD.Vendor_Name as [Vendor Name], TSPL_VENDOR_INVOICE_HEAD.RefDocNo, TSPL_VENDOR_INVOICE_HEAD.Description as " &
                          "[Description], CAST((case when TSPL_VENDOR_INVOICE_HEAD.On_Hold ='Y' then 1 else 0 end)as BIT) as Hold,TSPL_VENDOR_INVOICE_HEAD.Created_By as 'Created BY' FROM TSPL_VENDOR_INVOICE_HEAD WHERE  convert(date,TSPL_VENDOR_INVOICE_HEAD.invoice_entry_date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_VENDOR_INVOICE_HEAD.invoice_entry_date ,103) <= convert(date,'" + dtpToDate.Value + "',103) and TSPL_VENDOR_INVOICE_HEAD .ISProcurementDeduction =1 "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " And ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date, '')= '' "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and ISNULL(TSPL_VENDOR_INVOICE_HEAD.Posting_Date, '')<> '' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then

                        qry += " and TSPL_VENDOR_INVOICE_HEAD.Loc_Code   in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"

                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_VENDOR_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_VENDOR_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date , TSPL_VENDOR_INVOICE_HEAD.Document_No "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Multiple Deduction") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmMultipleProcDeduction)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST((case when ISNULL(TSPL_MULTIPLE_DEDUCTION_HEAD.Posting_Date, '')= '' then 0 else 1 end)as BIT) as Status, TSPL_MULTIPLE_DEDUCTION_HEAD.Document_No as [Document Id], TSPL_MULTIPLE_DEDUCTION_HEAD.Document_date as [Document Date], TSPL_MULTIPLE_DEDUCTION_HEAD.Loc_Code as [Loc Code],TSPL_MULTIPLE_DEDUCTION_HEAD.MCC_Code as [MCC Code],TSPL_MULTIPLE_DEDUCTION_HEAD.MCC_Name as [MCC Name],TSPL_MULTIPLE_DEDUCTION_HEAD.Remarks as " &
                          "[Description],TSPL_MULTIPLE_DEDUCTION_HEAD.Created_By as 'Created BY' FROM TSPL_MULTIPLE_DEDUCTION_HEAD WHERE  convert(date,TSPL_MULTIPLE_DEDUCTION_HEAD.Document_date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MULTIPLE_DEDUCTION_HEAD.Document_date ,103) <= convert(date,'" + dtpToDate.Value + "',103)  "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " And ISNULL(TSPL_MULTIPLE_DEDUCTION_HEAD.Posting_Date, '')= '' "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and ISNULL(TSPL_MULTIPLE_DEDUCTION_HEAD.Posting_Date, '')<> '' "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then

                        qry += " and TSPL_MULTIPLE_DEDUCTION_HEAD.Loc_Code   in (select Loc_Segment_Code   from TSPL_LOCATION_MASTER where Location_Code IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") and Is_Sub_Location='N' and Is_Section='N')"

                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MULTIPLE_DEDUCTION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MULTIPLE_DEDUCTION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MULTIPLE_DEDUCTION_HEAD.Document_Date , TSPL_MULTIPLE_DEDUCTION_HEAD.Document_No "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Acknowledgement Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmAcknowledgementEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "SELECT CAST(TSPL_ACKNOWLEDGENT_ENTRY_Header.isPosted as BIT) as Status, TSPL_ACKNOWLEDGENT_ENTRY_Header.Document_No as [Document Id], TSPL_ACKNOWLEDGENT_ENTRY_Header.Document_Date as [Document Date],TSPL_ACKNOWLEDGENT_ENTRY_Header.Dispatch_Document_No as [Tanker Dispatch Doc No],TSPL_ACKNOWLEDGENT_ENTRY_Header.Tanker_No , TSPL_ACKNOWLEDGENT_ENTRY_Header.Created_By as 'Created BY','' as Description FROM TSPL_ACKNOWLEDGENT_ENTRY_Header WHERE  convert(date,TSPL_ACKNOWLEDGENT_ENTRY_Header.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_ACKNOWLEDGENT_ENTRY_Header.Document_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_ACKNOWLEDGENT_ENTRY_Header .isPosted =0  "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_ACKNOWLEDGENT_ENTRY_Header .isPosted =1  "
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_ACKNOWLEDGENT_ENTRY_Header.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ACKNOWLEDGENT_ENTRY_Header.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_ACKNOWLEDGENT_ENTRY_Header.Document_Date , TSPL_ACKNOWLEDGENT_ENTRY_Header.Document_No "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Milk Shift Uploader") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkShiftUploader)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST(TSPL_MILK_SHIFT_UPLOADER_HEAD.Status as BIT) as Status, TSPL_MILK_SHIFT_UPLOADER_HEAD.Document_No as [Document Id], TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date as [Document Date],TSPL_MILK_SHIFT_UPLOADER_HEAD.MCC_Code as [MCC],TSPL_MCC_MASTER.MCC_NAME as [MCC Name] , TSPL_MILK_SHIFT_UPLOADER_HEAD.Created_By as 'Created BY',TSPL_MILK_SHIFT_UPLOADER_HEAD.Description FROM TSPL_MILK_SHIFT_UPLOADER_HEAD 
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_SHIFT_UPLOADER_HEAD.MCC_Code
WHERE  convert(date,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Status=0  "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Status=1  "
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_SHIFT_UPLOADER_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MILK_SHIFT_UPLOADER_HEAD.Shift_Date, TSPL_MILK_SHIFT_UPLOADER_HEAD.Document_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Milk Procurement Uploader") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkProcurementUploader)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST(TSPL_MILK_PROCUREMENT_UPLOADER_head.Status as BIT) as Status, TSPL_MILK_PROCUREMENT_UPLOADER_head.Document_No as [Document Id], TSPL_MILK_PROCUREMENT_UPLOADER_head.Document_Date as [Document Date],TSPL_MILK_PROCUREMENT_UPLOADER_head.MCC_Code as [MCC],TSPL_MCC_MASTER.MCC_NAME as [MCC Name] , TSPL_MILK_PROCUREMENT_UPLOADER_head.Created_By as 'Created BY',TSPL_MILK_PROCUREMENT_UPLOADER_head.Description FROM TSPL_MILK_PROCUREMENT_UPLOADER_head 
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_MILK_PROCUREMENT_UPLOADER_head.MCC_Code
WHERE  convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_head.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MILK_PROCUREMENT_UPLOADER_head.Document_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_head.Status=0  "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_head.Status=1  "
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_PROCUREMENT_UPLOADER_head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MILK_PROCUREMENT_UPLOADER_head.Document_Date, TSPL_MILK_PROCUREMENT_UPLOADER_head.Document_No "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Milk Collection") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkCollectionMCC)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST(TSPL_MILK_COLLECTION_MCC.Status as BIT) as Status, TSPL_MILK_COLLECTION_MCC.Document_No as [Document Id], TSPL_MILK_COLLECTION_MCC.Document_Date as [Document Date], TSPL_MILK_COLLECTION_MCC.Created_By as 'Created BY',TSPL_MILK_COLLECTION_MCC.Description FROM TSPL_MILK_COLLECTION_MCC 
                          WHERE  convert(date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MILK_COLLECTION_MCC.Document_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_MCC.Status=0  "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_MCC.Status=1  "
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MILK_COLLECTION_MCC.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_COLLECTION_MCC.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MILK_COLLECTION_MCC.Document_Date, TSPL_MILK_COLLECTION_MCC.Document_No "
                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "DCS Milk Collection") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkCollectionDCS)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST(TSPL_MILK_COLLECTION_DCS.Status as BIT) as Status, TSPL_MILK_COLLECTION_DCS.Document_No as [Document Id], TSPL_MILK_COLLECTION_DCS.Document_Date as [Document Date], TSPL_MILK_COLLECTION_DCS.Created_By as 'Created BY',TSPL_MILK_COLLECTION_DCS.Description FROM TSPL_MILK_COLLECTION_DCS 
                          WHERE  convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MILK_COLLECTION_DCS.Document_Date ,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_DCS.Status=0  "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_DCS.Status=1  "
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MILK_COLLECTION_DCS.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_COLLECTION_DCS.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MILK_COLLECTION_DCS.Document_Date, TSPL_MILK_COLLECTION_DCS.Document_No "
                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "BMC Mobile Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkCollectionDCS)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST(TSPL_MILK_COLLECTION_BMCDCS.Status as BIT) as Status, TSPL_MILK_COLLECTION_BMCDCS.PK_ID as [Document Id], FORMAT(TSPL_MILK_COLLECTION_BMCDCS.IDate, 'dd/MM/yyyy') AS [Date], TSPL_MILK_COLLECTION_BMCDCS.MCC_Code as [BMC Code] FROM TSPL_MILK_COLLECTION_BMCDCS 
                          WHERE convert(date,TSPL_MILK_COLLECTION_BMCDCS.IDate,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MILK_COLLECTION_BMCDCS.IDate,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_BMCDCS.Status=0  "
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_BMCDCS.Status=1  "
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_MILK_COLLECTION_BMCDCS.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_COLLECTION_BMCDCS.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += " ORDER BY TSPL_MILK_COLLECTION_BMCDCS.IDate, TSPL_MILK_COLLECTION_BMCDCS.PK_ID "
                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub


    Sub FillMilkProcurementBulk()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmBulkMilkPurchaseInvoice)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(isPosted AS BIT) as Status,DOC_NO as [Document Id],DOC_DATE as [Document Date],isnull(Total_AMT,0) as [Amount], '' as [Description], Vendor_Code as [Vendor Code],loc_code as [Location Code] , Created_By as [Created By],Created_Date as [Created Date]  from TSPL_BULK_MILK_PURCHASE_INVOICE_head " &
                        " where  convert(date,DOC_DATE,103)>= convert(date,'" + dtpFromDate.Value + "',103) " &
                    " and convert(date,DOC_DATE,103) <=convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_BULK_MILK_PURCHASE_INVOICE_head.isPosted = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_BULK_MILK_PURCHASE_INVOICE_head.isPosted = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_BULK_MILK_PURCHASE_INVOICE_head.loc_code in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_BULK_MILK_PURCHASE_INVOICE_head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_BULK_MILK_PURCHASE_INVOICE_head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_BULK_MILK_PURCHASE_INVOICE_head.DOC_DATE, TSPL_BULK_MILK_PURCHASE_INVOICE_head.DOC_NO "
                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "DCS Milk Collection Multiple Days") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkCollectionDCSMultipleDays)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(Status AS BIT) as Status,Document_No as [Document Id],Document_Date as [Document Date],  Created_By as [Created By],Description from TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS  " &
                        " where  convert(date,Document_Date,103)>= convert(date,'" + dtpFromDate.Value + "',103) " &
                    " and convert(date,Document_Date,103) <=convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Status = 1"
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_Date, TSPL_MILK_COLLECTION_DCS_MULTIPLE_DAYS.Document_No "
                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Milk Collection Multiple Days") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.MilkCollectionDCSMultipleDays)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "select CAST(Status AS BIT) as Status,Document_No as [Document Id],Document_Date as [Document Date],  Created_By as [Created By],Description from TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS  " &
                        " where  convert(date,Document_Date,103)>= convert(date,'" + dtpFromDate.Value + "',103) " &
                    " and convert(date,Document_Date,103) <=convert(date,'" + dtpToDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS.Status = 1"
                        Isrefreshed = True
                    End If

                    If ChkAllowBulkPosting.Equals(0) Then

                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS.Document_Date, TSPL_MILK_COLLECTION_MCC_MULTIPLE_DAYS.Document_No "
                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub

    Sub FillComplaint()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Customer Complaint") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmCustomerComplaint)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((case when TSPL_CUSTOMER_COMPLAINT_HEAD.isPosted='Y' then 1 else 0 end)as BIT) as Status,TSPL_CUSTOMER_COMPLAINT_HEAD.Complaint_No as[Document Id],convert(varchar,TSPL_CUSTOMER_COMPLAINT_HEAD.Complaint_Date,103) as [Document Date],TSPL_CUSTOMER_COMPLAINT_HEAD.Invoice_No as [Invoice No],TSPL_CUSTOMER_COMPLAINT_HEAD.Cust_Code as [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer],TSPL_CUSTOMER_COMPLAINT_HEAD.Created_By as [Created By],convert(varchar,TSPL_CUSTOMER_COMPLAINT_HEAD.Created_Date,103) as [Created Date],TSPL_CUSTOMER_COMPLAINT_HEAD.Remarks as Description from TSPL_CUSTOMER_COMPLAINT_HEAD 
Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_COMPLAINT_HEAD.Cust_Code =TSPL_CUSTOMER_MASTER.Cust_Code " &
                    " WHERE isnull(TSPL_CUSTOMER_COMPLAINT_HEAD.Invoice_No,'')='' and convert(date,TSPL_CUSTOMER_COMPLAINT_HEAD.Complaint_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_CUSTOMER_COMPLAINT_HEAD.Complaint_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and isnull(TSPL_CUSTOMER_COMPLAINT_HEAD.isPosted,'N')='N'"
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_CUSTOMER_COMPLAINT_HEAD.isPosted,'N')='Y'"
                        Isrefreshed = True
                    End If
                    'If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    '    qry += " and TSPL_Dispatch_BulkSale.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    'End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_CUSTOMER_COMPLAINT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_CUSTOMER_COMPLAINT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_CUSTOMER_COMPLAINT_HEAD.Complaint_Date, TSPL_CUSTOMER_COMPLAINT_HEAD.Complaint_No "

                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub
    '============================================

    ''richa agarwal 12/05/2015 against ticket no.BM00000006520
    Sub FillBulkSale()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bulk Dispatch") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FrmDispatchBulkSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_Dispatch_BulkSale.Posted)as BIT) as Status,TSPL_Dispatch_BulkSale.Document_No as[Document Id],convert(varchar,TSPL_Dispatch_BulkSale.Document_Date,103) as [Document Date]," &
                    " isnull(TSPL_Dispatch_BulkSale.Total_Amt,0) as [Amount],TSPL_Dispatch_BulkSale.QC_Code as [QC No],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer],TSPL_Dispatch_BulkSale.Tare_Weight as [Tare Weight], " &
                    " TSPL_Dispatch_BulkSale.Gross_Weight as [Gross Weight],TSPL_Dispatch_BulkSale.Net_Weight as [Net Weight] ,TSPL_Dispatch_BulkSale.Created_By as [Created By], " &
                    " convert(varchar,TSPL_Dispatch_BulkSale.Created_Date,103) as [Created Date],'' as Description from TSPL_Dispatch_BulkSale " &
                    " Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_Dispatch_BulkSale.Customer_Code =TSPL_CUSTOMER_MASTER.Cust_Code " &
                    " WHERE  convert(date,TSPL_Dispatch_BulkSale.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_Dispatch_BulkSale.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_Dispatch_BulkSale.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_Dispatch_BulkSale.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_Dispatch_BulkSale.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_Dispatch_BulkSale.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_Dispatch_BulkSale.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_Dispatch_BulkSale.Document_Date, TSPL_Dispatch_BulkSale.Document_No "

                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Can Sale") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FrmCanSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_CAN_SALE_HEAD.Posted)as BIT) as Status,TSPL_CAN_SALE_HEAD.Document_No as[Document Id],convert(varchar,TSPL_CAN_SALE_HEAD.Document_Date,103) as [Document Date]," &
                    " isnull(TSPL_CAN_SALE_HEAD.DocumentAmount,0) as [Amount],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer], " &
                    " TSPL_CAN_SALE_HEAD.Created_By as [Created By], " &
                    " convert(varchar,TSPL_CAN_SALE_HEAD.Created_Date,103) as [Created Date],'' as Description from TSPL_CAN_SALE_HEAD " &
                    " Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CAN_SALE_HEAD.Customer_Code =TSPL_CUSTOMER_MASTER.Cust_Code " &
                    " WHERE  convert(date,TSPL_CAN_SALE_HEAD.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_CAN_SALE_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_CAN_SALE_HEAD.POSTED = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_CAN_SALE_HEAD.POSTED = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_CAN_SALE_HEAD.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_CAN_SALE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_CAN_SALE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_CAN_SALE_HEAD.Document_Date, TSPL_CAN_SALE_HEAD.Document_No "

                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub


    Sub FillFreshSale()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment/Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FrmDispatchFreshSale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "Select CAST((TSPL_SD_SALE_INVOICE_HEAD.Status)as BIT) as Status,Against_Shipment_No as [Document Id],TSPL_SD_SALE_INVOICE_HEAD.Document_Code as InvoiceNo, " &
                        "convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as [Document Date], isnull(TSPL_SD_SALE_INVOICE_HEAD.Total_Amt,0) as [Amount], " &
                        "'' as [QC No],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer],TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location as Location, " &
                        "TSPL_SD_SALE_INVOICE_HEAD.Created_By as [Created By],  convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Created_Date,103) as [Created Date] ,TSPL_SD_SALE_INVOICE_HEAD.description" &
                        " from TSPL_SD_SALE_INVOICE_HEAD  Left Outer Join TSPL_CUSTOMER_MASTER on " &
                        "TSPL_SD_SALE_INVOICE_HEAD.Customer_Code =TSPL_CUSTOMER_MASTER.Cust_Code  " &
                        "Left Outer Join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER.Location_Code " &
                    " WHERE  convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and " &
                    "convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_SD_SALE_INVOICE_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='FS' "
                    qry += " ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Code "

                End If
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub

    Sub FillDairySale()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment/Invoice") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmSaleDispatchDairy)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "Select CAST((TSPL_SD_SHIPMENT_HEAD.Status)as BIT) as Status,Document_Code as [Document Id], " &
                        "(select isnull((Select distinct '['+TSPL_SD_SALE_INVOICE_HEAD.Document_Code+']  ' from TSPL_SD_SHIPMENT_HEAD a left outer join TSPL_SD_SALE_INVOICE_HEAD on a.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No where  a.Document_Code= TSPL_SD_SHIPMENT_HEAD.Document_Code  for xml path('')),'') ) as InvoiceNo, " &
                        "convert(varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as [Document Date], isnull(TSPL_SD_SHIPMENT_HEAD.Total_Amt,0) as [Amount], " &
                        "'' as [QC No],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer],TSPL_SD_SHIPMENT_HEAD.Bill_To_Location as Location,  " &
                        "TSPL_SD_SHIPMENT_HEAD.Created_By as [Created By],  convert(varchar,TSPL_SD_SHIPMENT_HEAD.Created_Date,103) as [Created Date] ,   " &
                        "TSPL_SD_SHIPMENT_HEAD.description from TSPL_SD_SHIPMENT_HEAD  Left Outer Join TSPL_CUSTOMER_MASTER on  " &
                        "TSPL_SD_SHIPMENT_HEAD.Customer_Code =TSPL_CUSTOMER_MASTER.Cust_Code  Left Outer Join TSPL_LOCATION_MASTER on  " &
                        "TSPL_SD_SHIPMENT_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER.Location_Code " &
                    " WHERE  convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and " &
                    "convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_SD_SHIPMENT_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_SD_SHIPMENT_HEAD.Bill_To_Location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_SD_SHIPMENT_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                        qry += " and TSPL_CUSTOMER_MASTER.Cust_Code  in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") " + Environment.NewLine
                    End If
                    'qry += " and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type='FS' "
                    qry += " and TSPL_SD_SHIPMENT_HEAD.Screen_Type='DS' "
                    qry += " ORDER BY TSPL_SD_SHIPMENT_HEAD.Document_Date, TSPL_SD_SHIPMENT_HEAD.Document_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmDemandBooking)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST(TSPL_DEMAND_BOOKING_MASTER.Posted as BIT) as Status " &
                        " ,TSPL_DEMAND_BOOKING_MASTER.Document_No as [Document Id],TSPL_DEMAND_BOOKING_MASTER.ShiftType " &
                    ",TSPL_DEMAND_BOOKING_MASTER.Route_No as [Route Code],TSPL_ROUTE_MASTER.Route_Desc as [Route],TSPL_DEMAND_BOOKING_MASTER.Location_Code as [Location Code] " &
                    ",TSPL_LOCATION_MASTER.Location_Desc as [Location], TSPL_DEMAND_BOOKING_MASTER.Created_By as [Created By] " &
                    ", convert(varchar,TSPL_DEMAND_BOOKING_MASTER.Created_Date,103) as [Created Date] " &
                    ", '' as [Description] " &
                    "from TSPL_DEMAND_BOOKING_MASTER " &
                    "Left Outer Join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No =TSPL_DEMAND_BOOKING_MASTER.Route_No   " &
                    "Left Outer Join TSPL_LOCATION_MASTER on TSPL_DEMAND_BOOKING_MASTER.Location_Code=TSPL_LOCATION_MASTER.Location_Code" &
                    " WHERE  convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and " &
                    "convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= convert(date,'" + dtpFromDate.Value + "',103) "
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_DEMAND_BOOKING_MASTER.ShiftType ='Morning' and TSPL_DEMAND_BOOKING_MASTER.Posted = 0 "
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_DEMAND_BOOKING_MASTER.ShiftType = 'Evening' and TSPL_DEMAND_BOOKING_MASTER.Posted = 0 "
                        Isrefreshed = False
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_DEMAND_BOOKING_MASTER.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If txtCustomer.arrValueMember IsNot Nothing Then
                        qry += " and TSPL_DEMAND_BOOKING_MASTER.Route_No in(" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_DEMAND_BOOKING_MASTER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_DEMAND_BOOKING_MASTER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY [Created Date], [Document Id] "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Booking/DO") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmbookingdairy)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST((case when TSPL_BOOKING_DETAIL.Booking_Status=4 then 1 else 0 end)as BIT) as Status " &
                        " ,TSPL_BOOKING_DETAIL.Document_No as [Document Id] " &
                    ",TSPL_BOOKING_DETAIL.DocumentAmount as [Amount],TSPL_CUSTOMER_MASTER.Cust_Code as [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name] " &
                    ",TSPL_LOCATION_MASTER.Location_Desc as [Location], TSPL_BOOKING_MATSER.Created_By as [Created By] " &
                    ", convert(varchar,TSPL_BOOKING_MATSER.Created_Date,103) as [Created Date] " &
                    ", '' as [Description] " &
                    "from TSPL_BOOKING_MATSER inner join TSPL_BOOKING_DETAIL on TSPL_BOOKING_DETAIL.Document_No=TSPL_BOOKING_MATSER.Document_No " &
                    "Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_BOOKING_DETAIL.Cust_Code =TSPL_CUSTOMER_MASTER.Cust_Code   " &
                    "Left Outer Join TSPL_LOCATION_MASTER on TSPL_BOOKING_DETAIL.Loc_Code =TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE FOC_Item=0 and convert(date,TSPL_BOOKING_MATSER.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and " &
                    "convert(date,TSPL_BOOKING_MATSER.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) " &
                    " and isnull(TSPL_BOOKING_MATSER.From_Screen_code,'')<>'" & "" & "' "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_BOOKING_DETAIL.Booking_Status in (1,3) "
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_BOOKING_DETAIL.Booking_Status = 4"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_BOOKING_DETAIL.loc_code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_BOOKING_MATSER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_BOOKING_MATSER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If

                    If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                        qry += " and TSPL_CUSTOMER_MASTER.Cust_Code  in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") " + Environment.NewLine
                    End If

                    qry += " group by Booking_Status,TSPL_BOOKING_DETAIL.Document_No,DocumentAmount,TSPL_CUSTOMER_MASTER.Cust_Code,Customer_Name,Location_Desc,TSPL_BOOKING_MATSER.Created_By,TSPL_BOOKING_MATSER.Created_Date"
                    qry += " ORDER BY [Created Date], [Document Id] "
                End If
            End If
            ''richa agarwal 19 Nov,2019
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Fresh Booking/DO") = CompairStringResult.Equal Then
            Load_Authorisation("")
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select CAST((case when TSPL_BOOKING_DETAIL.Booking_Status=4 then 1 else 0 end)as BIT) as Status " &
                        " ,TSPL_BOOKING_DETAIL.Document_No as [Document Id] " &
                    ",TSPL_BOOKING_DETAIL.DocumentAmount as [Amount],TSPL_CUSTOMER_MASTER.Cust_Code as [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name] " &
                    ",TSPL_LOCATION_MASTER.Location_Desc as [Location], TSPL_BOOKING_MATSER.Created_By as [Created By] " &
                    ", convert(varchar,TSPL_BOOKING_MATSER.Created_Date,103) as [Created Date] " &
                    ", '' as [Description] " &
                    "from TSPL_BOOKING_MATSER inner join TSPL_BOOKING_DETAIL on TSPL_BOOKING_DETAIL.Document_No=TSPL_BOOKING_MATSER.Document_No " &
                    "Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_BOOKING_DETAIL.Cust_Code =TSPL_CUSTOMER_MASTER.Cust_Code   " &
                    "Left Outer Join TSPL_LOCATION_MASTER on TSPL_BOOKING_DETAIL.Loc_Code =TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE FOC_Item=0 and convert(date,TSPL_BOOKING_MATSER.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and " &
                    "convert(date,TSPL_BOOKING_MATSER.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) " &
                    " and isnull(TSPL_BOOKING_MATSER.From_Screen_code,'')='" & "" & "' "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_BOOKING_DETAIL.Booking_Status in (1,3) "
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_BOOKING_DETAIL.Booking_Status = 4"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_BOOKING_DETAIL.loc_code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_BOOKING_MATSER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_BOOKING_MATSER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If

                    If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                        qry += " and TSPL_CUSTOMER_MASTER.Cust_Code  in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") " + Environment.NewLine
                    End If

                    qry += " group by Booking_Status,TSPL_BOOKING_DETAIL.Document_No,DocumentAmount,TSPL_CUSTOMER_MASTER.Cust_Code,Customer_Name,Location_Desc,TSPL_BOOKING_MATSER.Created_By,TSPL_BOOKING_MATSER.Created_Date"
                    qry += " ORDER BY [Created Date], [Document Id] "
                End If
            End If
            ''richa BHA/27/12/18-000763 28 Dec,2018
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Dairy GatePass") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmDairyGatePass)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "Select CAST((case when isnull(TSPL_DAIRYSALE_GATEPASS_MASTER.Post,'N')='N' then 0 else 1 end )as BIT) as Status,GPCode as [Document Id]," &
                    " TSPL_DAIRYSALE_GATEPASS_MASTER.Transporter , convert(varchar,TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate,103) as [Document Date]," &
                    " isnull(TSPL_DAIRYSALE_GATEPASS_MASTER.TotalCAN ,0) as [Total CAN],isnull(TSPL_DAIRYSALE_GATEPASS_MASTER.TotalCrate  ,0) as [Total Crate], " &
                    " TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Id  as [Vehicle Id],TSPL_DAIRYSALE_GATEPASS_MASTER.Vehicle_Number as [Vehicle No]," &
                    " TSPL_DAIRYSALE_GATEPASS_MASTER.Location_Code  as Location, TSPL_DAIRYSALE_GATEPASS_MASTER.Created_By as [Created By]," &
                    " convert(varchar,TSPL_DAIRYSALE_GATEPASS_MASTER.Created_Date,103) as [Created Date] , " &
                    " TSPL_DAIRYSALE_GATEPASS_MASTER.Remarks as description from TSPL_DAIRYSALE_GATEPASS_MASTER " &
                    " Left Outer Join TSPL_LOCATION_MASTER on  TSPL_DAIRYSALE_GATEPASS_MASTER.Location_Code =TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  " &
                    " and convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate,103) <= convert(date,'" + dtpToDate.Value + "',103)"

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and isnull(TSPL_DAIRYSALE_GATEPASS_MASTER.Post,'N')='N'  "
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_DAIRYSALE_GATEPASS_MASTER.Post,'N')='Y'"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If

                    qry += " ORDER BY TSPL_DAIRYSALE_GATEPASS_MASTER.GPDate, TSPL_DAIRYSALE_GATEPASS_MASTER.GPCode  "
                End If
            End If
            '================Added by preeti Gupta Against ticket no[BHA/26/06/19-000911][26/06/2019]
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Crate Received") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmCrateReceviedDairySale)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing
                    qry = "select isnull(CAST((TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted)as BIT),0) as Status  " &
                   " ,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No as [Document Id] , convert(varchar,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date,103) as [Document Date],TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.type" &
                    " ,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.totalcrateQty as [Crate Qty],TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.TotalCanQty as [Can Qty]" &
                   " ,TSPL_LOCATION_MASTER.Location_Desc as [Location], TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Created_By as [Created By] " &
                   " , convert(varchar,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Created_Date,103) as [Created Date] " &
                   " ,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Comments  as [Description] " &
                   " from TSPL_CRATE_RECEIVED_HEAD_FRESHSALE" &
                " Left Outer Join TSPL_LOCATION_MASTER on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Location_Code =TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and " &
                    "convert(date,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and  isnull(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted,0) =0  "
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Posted,0) =1 "
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If

                    'If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                    '    qry += " and TSPL_CUSTOMER_MASTER.Cust_Code  in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") " + Environment.NewLine
                    'End If

                    'qry += " group by Booking_Status,TSPL_BOOKING_DETAIL.Document_No,DocumentAmount,TSPL_CUSTOMER_MASTER.Cust_Code,Customer_Name,Location_Desc,TSPL_BOOKING_MATSER.Created_By,TSPL_BOOKING_MATSER.Created_Date"
                    'qry += " ORDER BY [Created Date], [Document Id] "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Demand Booking") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FrmProductDemandBooking)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = "SELECT  CAST(TSPL_Product_DEMAND_BOOKING_MASTER.posted as BIT) as Status, 

            '' as [Description], '' as Hold,
            TSPL_Product_DEMAND_BOOKING_MASTER.Document_no as [Document Id],
                                     TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date as [Document Date],  isnull(TSPL_PRODUCT_DEMAND_BOOKING_MASTER.DocumentAmount,0) as [Amount], 
                                     TSPL_Product_DEMAND_BOOKING_MASTER.Location_Code As Location , TSPL_LOCATION_MASTER.Location_Desc as [Location Desc],
                                     TSPL_Product_DEMAND_BOOKING_MASTER.Created_By as 'Created By' FROM TSPL_Product_DEMAND_BOOKING_MASTER" &
                                   " Left Outer Join TSPL_LOCATION_MASTER on TSPL_Product_DEMAND_BOOKING_MASTER.Location_Code=TSPL_LOCATION_MASTER.Location_Code" &
                        "  WHERE   convert(date,TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)"
                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.posted = 0"
                        Isrefreshed = False
                    ElseIf rbtnStatusPosted.IsChecked = True Then
                        qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.posted = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " AND TSPL_Product_DEMAND_BOOKING_MASTER.Location_Code in (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ")"
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                            qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_Product_DEMAND_BOOKING_MASTER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If
                    End If
                    qry += "ORDER BY TSPL_Product_DEMAND_BOOKING_MASTER.Document_Date, TSPL_Product_DEMAND_BOOKING_MASTER.Document_No "
                End If
            End If
            '======================================================================
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub
    '=====================Addede by preeti gupta=======[25/05/2017]

    Sub FillMCCMaterialSaleFarmer()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale Farmer") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmMCCMaterialFarmer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_MCC_Sale_Farmer_Head.Status)as BIT) as Status,TSPL_MCC_Sale_Farmer_Head.Document_Code as[Document Id],convert(varchar,TSPL_MCC_Sale_Farmer_Head.Document_Date,103) as [Document Date]," &
                    " isnull(TSPL_MCC_Sale_Farmer_Head.Total_Amt,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_MCC_Sale_Farmer_Head.Created_By as [Created By], " &
                    " convert(varchar,TSPL_MCC_Sale_Farmer_Head.Created_Date,103) as [Created Date],TSPL_MCC_Sale_Farmer_Head.Description from TSPL_MCC_Sale_Farmer_Head " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_MCC_Sale_Farmer_Head.bill_to_location=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_MCC_Sale_Farmer_Head.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MCC_Sale_Farmer_Head.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MCC_Sale_Farmer_Head.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_MCC_Sale_Farmer_Head.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MCC_Sale_Farmer_Head.bill_to_location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MCC_Sale_Farmer_Head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MCC_Sale_Farmer_Head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_MCC_Sale_Farmer_Head.Document_Date, TSPL_MCC_Sale_Farmer_Head.Document_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale Farmer Return") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmMCCMaterialFarmer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_MCC_SALE_RETURN_HEAD_FARMER.Status)as BIT) as Status,TSPL_MCC_SALE_RETURN_HEAD_FARMER.Document_Code as[Document Id],convert(varchar,TSPL_MCC_SALE_RETURN_HEAD_FARMER.Document_Date,103) as [Document Date]," &
                    " isnull(TSPL_MCC_SALE_RETURN_HEAD_FARMER.Total_Amt,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_MCC_SALE_RETURN_HEAD_FARMER.Created_By as [Created By], " &
                    " convert(varchar,TSPL_MCC_SALE_RETURN_HEAD_FARMER.Created_Date,103) as [Created Date],TSPL_MCC_SALE_RETURN_HEAD_FARMER.Description from TSPL_MCC_SALE_RETURN_HEAD_FARMER " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_MCC_SALE_RETURN_HEAD_FARMER.bill_to_location=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_MCC_SALE_RETURN_HEAD_FARMER.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MCC_SALE_RETURN_HEAD_FARMER.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_MCC_SALE_RETURN_HEAD_FARMER.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_MCC_SALE_RETURN_HEAD_FARMER.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MCC_SALE_RETURN_HEAD_FARMER.bill_to_location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MCC_SALE_RETURN_HEAD_FARMER.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MCC_SALE_RETURN_HEAD_FARMER.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_MCC_SALE_RETURN_HEAD_FARMER.Document_Date, TSPL_MCC_SALE_RETURN_HEAD_FARMER.Document_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Farmer Payment Adjustment") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmMCCMaterialFarmer)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_MP_Pay_Adj_Head.Is_Post)as BIT) as Status,TSPL_MP_Pay_Adj_Head.Adjustment_No as[Document Id],convert(varchar,TSPL_MP_Pay_Adj_Head.Adjustment_Date,103) as [Document Date]," &
                    " isnull(TSPL_MP_Pay_Adj_Head.Doc_Amount ,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_MP_Pay_Adj_Head.Created_By as [Created By], " &
                    " convert(varchar,TSPL_MP_Pay_Adj_Head.Created_Date,103) as [Created Date],TSPL_MP_Pay_Adj_Head.Description from TSPL_MP_Pay_Adj_Head " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_MP_Pay_Adj_Head.MCC_Code=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_MP_Pay_Adj_Head.Adjustment_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_MP_Pay_Adj_Head.Adjustment_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and isnull(TSPL_MP_Pay_Adj_Head.Is_Post,'') = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_MP_Pay_Adj_Head.Is_Post,'') = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_MP_Pay_Adj_Head.MCC_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_MP_Pay_Adj_Head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_MP_Pay_Adj_Head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_MP_Pay_Adj_Head.Adjustment_No, TSPL_MP_Pay_Adj_Head.Adjustment_Date "

                End If
            End If
        End If


        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub


    '===============================================================
    ''-----------------------------------
#End Region


#Region "Posting"

    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        trnsLst = New List(Of String)
        Dim trnsNo As Integer
        countPostedDoc = 0
        trnsLstCustomer = New List(Of String)
        '' CODE BY PANCH RAJ ON 24.AUG-2013
        DtError = New DataTable
        DtError.Columns.Add("Code", GetType(String))
        DtError.Columns.Add("Error", GetType(String))

        If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Sale") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Booking/DO") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Fresh Booking/DO") = CompairStringResult.Equal) Then
            For trnsNo = 0 To gv1.ChildRows.Count - 1
                If gv1.ChildRows(trnsNo).Cells("Status").Value = True Then
                    trnsLst.Add(gv1.ChildRows(trnsNo).Cells("Document Id").Value)  '' Insert The Document_Id in The StringList
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
                        trnsLstCustomer.Add(IIf(clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(trnsNo).Cells("ShiftType").Value), "Morning") = CompairStringResult.Equal, "1", IIf(clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(trnsNo).Cells("ShiftType").Value), "Evening") = CompairStringResult.Equal, "2", "3")))
                    Else
                        trnsLstCustomer.Add(gv1.Rows(trnsNo).Cells("Customer Code").Value) '' Insert The Customer Code in The StringList
                    End If

                End If
            Next
        Else
            For trnsNo = 0 To gv1.ChildRows.Count - 1
                If gv1.ChildRows(trnsNo).Cells("Status").Value = True Then
                    trnsLst.Add(gv1.ChildRows(trnsNo).Cells("Document Id").Value)  '' Insert The Document_Id in The StringList
                End If
            Next
        End If

        If trnsLst.Count <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Select Atleast One Document", Me.Text)
        Else

            If myMessages.postConfirm Then
                Try
                    clsCommon.ProgressBarPercentShow()
                    If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Purchase Order") = CompairStringResult.Equal Then
                        PostDataPO()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Payables") = CompairStringResult.Equal Then
                        PostPayable()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Common Services") = CompairStringResult.Equal Then
                        PostCommon()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Receivables") = CompairStringResult.Equal Then
                        PostReceivable()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Tax Deducted At Source") = CompairStringResult.Equal Then
                        PostTDS()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "General Ledger") = CompairStringResult.Equal Then
                        PostgnrlLedger()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Product Sale") = CompairStringResult.Equal Then
                        PostSalesNDistribution()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fresh Sale") = CompairStringResult.Equal Then
                        PostFreshSale()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Sale") = CompairStringResult.Equal Then
                        PostDairySale()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Material Management") = CompairStringResult.Equal Then
                        PostMaterialMngmt()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "HR and Payroll") = CompairStringResult.Equal Then
                        PostHRAndPayroll()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Production") = CompairStringResult.Equal Then
                        PostProduction()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fixed Assets") = CompairStringResult.Equal Then
                        PostFixedAsset()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Bulk Sale") = CompairStringResult.Equal Then
                        PostBulkSale()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "MCC Procurement") = CompairStringResult.Equal Then
                        PostMccProcurement()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Milk Procurement Bulk") = CompairStringResult.Equal Then
                        PostMilkProcurementBulk()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Farmer Payment") = CompairStringResult.Equal Then
                        PostFarmerPayment()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Dairy Production") = CompairStringResult.Equal Then
                        PostDataDairyProduction()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Complaint") = CompairStringResult.Equal Then
                        PostComplaint()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Sales And Distribution") = CompairStringResult.Equal Then
                        PostSalesAndDistributionNew()
                    End If
                    clsCommon.ProgressBarPercentHide()
                    common.clsCommon.MyMessageBoxShow(Me, "" + clsCommon.myCstr(countPostedDoc) + " Document Posted Successfully", Me.Text)
                Catch ex As Exception
                    clsCommon.ProgressBarPercentHide()
                    common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                Finally
                    clsCommon.ProgressBarPercentHide()
                    '----Added by---Pankaj Kumar-on-14/06/2012-----For Openning a Exceptions Collection Window-------------
                    If DtError IsNot Nothing AndAlso DtError.Rows.Count > 0 Then
                        Dim frm As New FrmFreeGrid()
                        frm.strFormName = "Errors In Posting " + cboTransaction.Text + "'s Documents"
                        frm.dt = DtError
                        frm.ReportID = "BulkPostingErrors"
                        frm.ShowDialog()
                    End If
                    '-----------------------------------------Code Ends Here----------------------------------------------
                End Try
            End If
        End If
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Purchase Order)---------------
    ''-------------------------------------------------------------------

    Sub PostDataDairyProduction()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Issue Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsProcessProductionIssueEntry.PostData(clsUserMgtCode.frmProcessProductionIssueEntry, False, "", strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Standardization") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsProcessProductionStandardization.PostData(clsUserMgtCode.frmProcessProductionStandardization, strDocNo, "")
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Standardization Final QC") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsPPStdFinalQCHead.PostData(clsUserMgtCode.ProcessProductionStandardizationFinalQC, strDocNo, "")
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Stage Process") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsProcessProductionStageProcess.PostData(clsUserMgtCode.frmProcessProductionStageProcess, strDocNo, "")
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next


            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillDairyProduction()
    End Sub

    Sub PostDataPO()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Requisition") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsRequistionHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Order") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        Dim objPO As New clsPurchaseOrderHead()
                        objPO.PostData("", strDocNo, True, False)
                        objPO = Nothing
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Gate Receipt Note") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsGRNHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Material Receipt Note") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMRNHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

                '---VARSHA ADD NIRQC
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "NIR QC") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsNIRQC.postdata(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next


            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Receipt Note") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSRNHead.PostData("", strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Invoice") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsPurchaseInvoiceHead.PostData("", strDocNo, "", Nothing)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Return") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsPurchasReturnHead.PostData("", strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "RGP/NRGP") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsRGPHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Issue/Return/Transfer") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsIssueReturnHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Scrap LoadOut") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        ClsScrapSaleHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Scrap Invoice") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        ClsScrapInvoiceHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try

        gv1.DataSource = Nothing
        FillPurchaseOrder()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Payable)----------------------
    ''-------------------------------------------------------------------

    Sub PostPayable()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Payment Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsPaymentHeader.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "AP Invoice Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        Dim strRefDocNo As String = clsDBFuncationality.getSingleValue("Select TSPL_VENDOR_INVOICE_HEAD.RefDocNo from TSPL_VENDOR_INVOICE_HEAD Where TSPL_VENDOR_INVOICE_HEAD.Document_No ='" + strDocNo + "'")
                        clsVedorInvoiceHead.PostData("", strDocNo, strRefDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If
            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        gv1.DataSource = Nothing
        FillPayables()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Common Services)--------------
    ''-------------------------------------------------------------------

    Sub PostCommon()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bank Transfer") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsBankTrasnferNew.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Reverse Transaction") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsBankReverse.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents, And Creating Error Record-18/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try

                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)

        End Try
        gv1.DataSource = Nothing
        FillCommonServices()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Receivables)------------------
    ''-------------------------------------------------------------------
    ''against[BM00000007770]
    Sub UpdateGainAndLossAmtForRecieptEntry(ByVal strDocNo As String)
        Try
            Dim EXCHANGE_GAIN_ACCOUNT As String = String.Empty
            Dim EXCHANGE_LOSS_ACCOUNT As String = String.Empty
            Dim RECEIVED_AMOUNT_BASE_CURRENCY As Double = 0
            Dim EXCHANGE_LOSS_AMT As Double = 0
            Dim EXCHANGE_GAIN_AMT As Double = 0

            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("Select Cust_Code,TSPL_RECEIPT_HEADER.Receipt_Type ,TSPL_RECEIPT_DETAIL.Document_No,TSPL_RECEIPT_HEADER.RECEIVED_AMOUNT_BASE_CURRENCY,TSPL_RECEIPT_DETAIL.Applied_Amount,TSPL_RECEIPT_HEADER.Receipt_Date  from TSPL_RECEIPT_HEADER left outer join TSPL_RECEIPT_DETAIL on TSPL_RECEIPT_HEADER.Receipt_No =TSPL_RECEIPT_DETAIL.Receipt_No where TSPL_RECEIPT_HEADER.Receipt_No ='" & strDocNo & "'")

            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dt1.Rows(0)("Receipt_Type")), "R") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(dt1.Rows(0)("Receipt_Type")), "A") = CompairStringResult.Equal Then

                    Dim dt As DataTable
                    dt = clsRcptEntryHeader.GetExchangeDetailDt(clsCommon.myCstr(dt1.Rows(0)("Cust_Code")))
                    If dt.Rows.Count > 0 Then
                        EXCHANGE_GAIN_ACCOUNT = clsCommon.myCstr(dt.Rows(0).Item("EXCHANGE_GAIN_ACCOUNT"))
                        EXCHANGE_LOSS_ACCOUNT = clsCommon.myCstr(dt.Rows(0).Item("EXCHANGE_LOSS_ACCOUNT"))
                    Else
                        EXCHANGE_GAIN_ACCOUNT = Nothing
                        EXCHANGE_LOSS_ACCOUNT = Nothing
                    End If
                    Dim dtLastRate As DataTable
                    '' gather conv rate and amount of transaction to calculate exchange loss and gain
                    Dim strInvoiceNo As String = String.Empty
                    Dim lossorgainamount As Double = 0
                    Dim Totallossorgainamount As Double = 0
                    For i As Integer = 0 To dt1.Rows.Count - 1
                        strInvoiceNo = clsCommon.myCstr(dt1.Rows(i)("Document_No"))
                        dtLastRate = clsRcptEntryHeader.GetExchangeRateAmount(strInvoiceNo, clsCommon.myCDate(dt1.Rows(0)("Receipt_Date")))
                        lossorgainamount = clsCommon.myCdbl(dt1.Rows(i)("Applied_Amount")) * dtLastRate.Rows(0).Item("ConvRate")
                        Totallossorgainamount = Totallossorgainamount + lossorgainamount
                    Next
                    RECEIVED_AMOUNT_BASE_CURRENCY = clsCommon.myCdbl(dt1.Rows(0)("RECEIVED_AMOUNT_BASE_CURRENCY"))
                    Dim diff As Double = 0.0
                    If Totallossorgainamount <> 0 Then
                        diff = RECEIVED_AMOUNT_BASE_CURRENCY - Totallossorgainamount
                    End If
                    If diff = 0 Then
                        EXCHANGE_LOSS_AMT = 0
                        EXCHANGE_GAIN_AMT = 0
                    ElseIf diff < 0 Then
                        If clsCommon.myLen(EXCHANGE_LOSS_ACCOUNT) = 0 Then
                            Throw New Exception("Exchange Loss Account not defined.")
                        End If
                        EXCHANGE_LOSS_AMT = -diff
                        EXCHANGE_GAIN_AMT = 0
                    Else
                        If clsCommon.myLen(EXCHANGE_GAIN_ACCOUNT) = 0 Then
                            Throw New Exception("Exchange Gain Account not defined.")
                        End If
                        EXCHANGE_LOSS_AMT = 0
                        EXCHANGE_GAIN_AMT = diff
                    End If
                    clsDBFuncationality.ExecuteNonQuery("Update TSPL_RECEIPT_HEADER set EXCHANGE_GAIN_ACCOUNT='" & clsCommon.myCstr(EXCHANGE_GAIN_ACCOUNT) & "' ,EXCHANGE_LOSS_ACCOUNT='" & clsCommon.myCstr(EXCHANGE_LOSS_ACCOUNT) & "' ,EXCHANGE_GAIN_AMT=" & clsCommon.myCdbl(EXCHANGE_GAIN_AMT) & " ,EXCHANGE_LOSS_AMT=" & clsCommon.myCdbl(EXCHANGE_LOSS_AMT) & " where Receipt_No='" & clsCommon.myCstr(strDocNo) & "' ")
                Else
                    EXCHANGE_LOSS_AMT = 0
                    EXCHANGE_GAIN_AMT = 0
                    EXCHANGE_GAIN_ACCOUNT = Nothing
                    EXCHANGE_LOSS_ACCOUNT = Nothing
                End If

                'clsDBFuncationality.ExecuteNonQuery("Update TSPL_RECEIPT_HEADER set EXCHANGE_GAIN_ACCOUNT='" & clsCommon.myCstr(EXCHANGE_GAIN_ACCOUNT) & "' ,EXCHANGE_LOSS_ACCOUNT='" & clsCommon.myCstr(EXCHANGE_LOSS_ACCOUNT) & "' ,EXCHANGE_GAIN_AMT=" & clsCommon.myCdbl(EXCHANGE_GAIN_AMT) & " ,EXCHANGE_LOSS_AMT=" & clsCommon.myCdbl(EXCHANGE_LOSS_AMT) & " where Receipt_No='" & clsCommon.myCstr(strDocNo) & "' ")

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    '====Sanjeet(27/12/2016)=======================

    Sub PostMccProcurement()
        Try

            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VSP Item Issue") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsVSPItemIssue.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VSP Asset Issue") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMCCIssueReturnHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        'clsPaymentHeader.PostData(strDocNo)
                        clsMCCMaterialSale.PostData(clsUserMgtCode.frmMCCMaterial, strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale Return") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        'clsPaymentHeader.PostData(strDocNo)
                        clsMccMaterialSaleReturn.PostData(clsUserMgtCode.frmMCCMaterialSaleReturn, strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Procurement Deduction") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        Dim strRefDocNo As String = clsDBFuncationality.getSingleValue("Select TSPL_VENDOR_INVOICE_HEAD.RefDocNo from TSPL_VENDOR_INVOICE_HEAD Where TSPL_VENDOR_INVOICE_HEAD.Document_No ='" + strDocNo + "'")
                        clsVedorInvoiceHead.PostData("", strDocNo, strRefDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Multiple Deduction") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMultipleProcDeductionHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Acknowledgement Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsAcknowledgementEntry.PostData("", strDocNo)

                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Milk Shift Uploader") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkShiftUploaderHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Milk Procurement Uploader") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkProcurementUploaderHead.PostData(strDocNo)

                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Milk Collection") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkCollectionMCC.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "DCS Milk Collection") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkCollectionDCS.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "BMC Mobile Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkCollectionBMCDCS.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If
            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        gv1.DataSource = Nothing
        FillPayables()
    End Sub
    '===================Added by preeti gupta=================


    Sub PostFarmerPayment()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale Farmer") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        'clsVSPItemIssue.PostData(strDocNo)
                        clsMCCMaterialSaleFarmer.PostData(clsUserMgtCode.frmMCCMaterialFarmer, strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Material Sale Farmer Return") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMccMaterialSaleReturnFarmer.PostData(clsUserMgtCode.frmMCCMaterialSaleReturnFarmer, strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "farmer Payment Adjustment") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsFarmerPaymentAdjustmentEntry.FunPost(clsUserMgtCode.frmFarmerPaymentAdjustment, strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            End If
            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        gv1.DataSource = Nothing
        FillMCCMaterialSaleFarmer()
    End Sub


    '==========================================================
    Sub PostMilkProcurementBulk()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Purchase Invoice") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        'clsVSPItemIssue.PostData(strDocNo)
                        'clsMilkPurchaseInvoiceHead.postData(clsUserMgtCode.frmBulkMilkPurchaseInvoice, strDocNo)
                        clsMilkPurchaseInvoiceHead.postData(strDocNo, clsUserMgtCode.frmBulkMilkPurchaseInvoice)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "DCS Milk Collection Multiple Days") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkCollectionDCSMulipleDays.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "MCC Milk Collection Multiple Days") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsMilkCollectionMCCMulipleDays.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If
            GC.Collect()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        gv1.DataSource = Nothing
        FillPayables()
    End Sub

    '=============================================

    Sub PostReceivable()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Receipt Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        UpdateGainAndLossAmtForRecieptEntry(strDocNo)
                        clsRcptEntryHeader.funRcptPost(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-15/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Adjustment Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsAdjustmentEntryReceivables.FunPost(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-15/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "AR Invoice Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        'Dim RefDoc As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select RefDocNo  from TSPL_Customer_Invoice_Head"))
                        clsCustomerInvoiceHead.PostData("", strDocNo, "")
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-15/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillReceivables()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Tax Deducted At Source)-------
    ''-------------------------------------------------------------------

    Sub PostTDS()
        Try
            'If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Create Remittance") = CompairStringResult.Equal Then
            '    For j As Integer = 0 To trnsLst.Count - 1
            '        Dim strDocNo As String = trnsLst.Item(j)
            '        clsRemittance.PostRemit(trnsLst1)
            '        ''Add "Posting Method" Here
            '    Next

            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Remittance Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))

                        strDocNo = trnsLst.Item(j)
                        clsRemitanceEntry.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-15/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillTaxDdctdAtSource()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(General Ledger)----------
    ''-------------------------------------------------------------------

    Sub PostgnrlLedger()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Journal Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsGLEntry.funGLPOST(strDocNo, clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MM/yyyy"))
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-14/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "VCGL Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsVCGLHead.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillGeneralLedger()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Sales And Distribution)-------
    ''-------------------------------------------------------------------

    Sub PostSalesNDistribution()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Order") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSaleOrder.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-14/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment/Invoice") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        If clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Product Sale") = CompairStringResult.Equal Then
                            clsPSShipmentHead.PostData("Shipment", strDocNo)
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboModule.SelectedValue), "Fresh Sale") = CompairStringResult.Equal Then
                            clsDispatchNoteFreshSale.PostData("Shipment", strDocNo)
                        End If

                        'qry = "select Document_Code from TSPL_SD_SALE_INVOICE_HEAD where TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No='" + strDocNo + "'"
                        'qry = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                        'If clsCommon.myLen(qry) > 0 Then
                        '    clsSNInvoiceHead.PostData("Sale Invoice", qry, trans) 'Here Variable(qry) works like Sale Invoice no.
                        'End If
                        countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-14/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Invoice") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSNInvoiceHead.PostData("Sale Invoice", strDocNo)
                        countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-26/09/2014
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Return") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSNSalesReturnHead.PostData("Sale Return", strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-14/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Return (Inter Company)") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSaleReturnInterCompany.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Quick Settlement") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        ClsQuickSettlement.PostData(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-14/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillSalesNDistribution()
    End Sub

    Sub PostSalesAndDistributionNew()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSNShipmentHead.PostData(clsUserMgtCode.frmSNShipment, strDocNo, True)
                        countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-14/06/2012
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Sale Invoice") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsSNInvoiceHead.PostData("", strDocNo)
                        countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-26/09/2014
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Dispatch") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsPSInvoiceHead.PostData("", strDocNo)
                        countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-26/09/2014
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Demand Booking") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsProductDemandBookingSale.PostData("", strDocNo)
                        countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-26/09/2014
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillSalesAndDistributionNew()
    End Sub
    ''-------------------------------------------------------------------
    '' Function For Posting --------Module(Material management)----------
    ''-------------------------------------------------------------------

    Sub PostMaterialMngmt()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Transfer(Load-In)") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        clsTransferMaster.postTransfer(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-
                    Catch ex As Exception
                        ''''Added By-[Pankaj Kumar]--For Storing Doc no and Exception(Error) in a data table--------
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                        '-----------------------------------COde Ends Here-----------------------------------------
                    End Try
                Next

            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Transfer(Load-Out)") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)                   '
                        clsTransferMaster.postTransfer(strDocNo)
                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next

                ''--18/07/2012---Added by Pankaj Kuamr(This Code Posts the Data From [Adjustment Entry, Empty Transactions, Production Entry, Store Adjustment])---By-Ranjana Mam
            Else 'If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Adjustment Entry") = CompairStringResult.Equal Then
                For j As Integer = 0 To trnsLst.Count - 1
                    Try
                        clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                        strDocNo = trnsLst.Item(j)
                        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Adjustment Entry") = CompairStringResult.Equal Then
                            ClsAdjustments.PostData(strDocNo, "Adjustment Entry")
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Empty Transactions") = CompairStringResult.Equal Then
                            ClsAdjustments.PostData(strDocNo, "Empty Transactions")
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Entry") = CompairStringResult.Equal Then
                            ClsAdjustments.PostData(strDocNo, "Production Entry")
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Adjustment") = CompairStringResult.Equal Then
                            ClsAdjustments.PostData(strDocNo, "Store Adjustment")
                        End If

                        countPostedDoc = countPostedDoc + 1 '  Added By [Pankaj Kumar]-for Counting Posted Documents-
                    Catch ex As Exception
                        dr = DtError.NewRow()
                        dr("Code") = strDocNo
                        dr("Error") = ex.Message
                        DtError.Rows.Add(dr)
                    End Try
                Next
                ''--------------------------------------------------------Code Ends Here---------------------------------------------------------------------------------------
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillMaterialManagement()
    End Sub

    ''-------------------------------------------------------------------
    ''Function for Posting-------Module(Material Management)-------------
    ''---------------BY ASHWANI RAGHAV-----------------------------------

    Sub PostHRAndPayroll()
        Try
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Employee Salary") = CompairStringResult.Equal Then
                        clsEmployeeSalary.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Hourly Attendance") = CompairStringResult.Equal Then
                        clsHourlyAttendance.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Daily Attendance") = CompairStringResult.Equal Then
                        clsDailyAttendance.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Monthly Attendance") = CompairStringResult.Equal Then
                        clsMonthAttendance.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "OT Sheet") = CompairStringResult.Equal Then
                        clsOTSheet.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Allowance") = CompairStringResult.Equal Then
                        clsAllowanceDetails.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Deduction") = CompairStringResult.Equal Then
                        clsDeductionDetails.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bonus") = CompairStringResult.Equal Then
                        clsBonus.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Salary Adjustment") = CompairStringResult.Equal Then
                        clsAdjustmentVoucher.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Reimbursement") = CompairStringResult.Equal Then
                        clsReimbursementDetails.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Loan Application") = CompairStringResult.Equal Then
                        clsApplyLoan.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Loan Adjustment") = CompairStringResult.Equal Then
                        clsLoanAdjustment.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Leave Application") = CompairStringResult.Equal Then
                        clsLeaveApplication.PostData("LEAVE_APPLI", strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Leave Adjustment") = CompairStringResult.Equal Then
                        clsLeaveAdjustment.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Employee Increment") = CompairStringResult.Equal Then
                        ClsEmpIncrement.PostData(strDocNo, False)
                    End If

                    countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
            ''--------------------------------------------------------Code Ends Here---------------------------------------------------------------------------------------
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillHRAndPayroll()
    End Sub

    Sub PostProduction()
        Try
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Batch Order") = CompairStringResult.Equal Then
                        clsBatchOrder.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bill Of Material") = CompairStringResult.Equal Then
                        clsBillOfMaterial.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Manufacturing Order") = CompairStringResult.Equal Then
                        clsManufacturingOrder.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Planning") = CompairStringResult.Equal Then
                        clsProductionPlanning.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Receipt") = CompairStringResult.Equal Then
                        Dim trans As SqlClient.SqlTransaction
                        trans = clsDBFuncationality.GetTransactin
                        clsProductionRM.UpdateInventoryMovement(strDocNo, trans)
                        clsProductionReceipt.PostData(strDocNo, False, trans)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Entry") = CompairStringResult.Equal Then
                        clsProductionEntryWithoutBatch.PostData("PROD_ENTRY", strDocNo, "", True)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Requisition") = CompairStringResult.Equal Then
                        clsProductionRequisition.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Production Return") = CompairStringResult.Equal Then
                        clsProductionReturn.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Issue") = CompairStringResult.Equal Then
                        clsProductionIssue.PostData(strDocNo, False)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Silo Milk Transfer") = CompairStringResult.Equal Then
                        ClsSiloMilkTransfer.PostData(strDocNo)
                    End If

                    countPostedDoc = countPostedDoc + 1 'Added By [Pankaj Kumar]-for Counting Posted Documents-
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
            ''--------------------------------------------------------Code Ends Here---------------------------------------------------------------------------------------
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillProduction()
    End Sub
    ''richa agarwal 12/05/2015 against ticket no.BM00000006520
    Sub PostBulkSale()
        Try
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Bulk Dispatch") = CompairStringResult.Equal Then
                        ClsDispatchBulkSale.PostData("DISPATCH-BS", "", strDocNo)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Can Sale") = CompairStringResult.Equal Then
                        ClsCanSale.PostData("CAN-SALE", "", strDocNo)
                    End If
                    countPostedDoc = countPostedDoc + 1
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillBulkSale()
    End Sub


    Sub PostFreshSale()
        Try
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment/Invoice") = CompairStringResult.Equal Then
                        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                        Try
                            clsDispatchNoteFreshSale.PostData("", strDocNo, trans)
                            'clsSaleInvoiceFreshSale.PostData("", strDocNo, trans)
                            trans.Commit()
                        Catch ex As Exception
                            trans.Rollback()
                            Throw New Exception(ex.Message)
                        End Try
                    End If


                    countPostedDoc = countPostedDoc + 1
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)

            Return
        End Try
        gv1.DataSource = Nothing
        FillFreshSale()
    End Sub
    Sub PostComplaint()
        clsCommon.ProgressBarHide()
        clsCommon.ProgressBarPercentShow()
        Try
            Dim x As Integer = 0
            RecordCount = trnsLst.Count
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Customer Complaint") = CompairStringResult.Equal Then
                        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                        Try
                            x = x + 1
                            clsCommon.ProgressBarPercentUpdate(x / RecordCount * 100, " Posting Record(s) " & j + 1 & " of Total " & RecordCount)
                            clsCustomerComplainHead.PostData_FromBulkPosting(clsUserMgtCode.frmLeakageReplacementUploader, strDocNo, trans)
                            trans.Commit()
                        Catch ex As Exception
                            trans.Rollback()
                            Throw New Exception(ex.Message)
                        End Try
                    End If
                    countPostedDoc = countPostedDoc + 1
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)

            Return
        End Try
        clsCommon.ProgressBarPercentHide()
        clsCommon.ProgressBarShow()
        gv1.DataSource = Nothing
        FillComplaint()
    End Sub

    Sub PostDairySale()
        Try
            RecordCount = trnsLst.Count
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Shipment/Invoice") = CompairStringResult.Equal Then
                        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                        Try
                            clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                            clsPSShipmentHead.PostData(clsUserMgtCode.frmSaleDispatchDairy, strDocNo, trans, "", True)
                            'clsSaleInvoiceFreshSale.PostData("", strDocNo, trans)
                            trans.Commit()
                        Catch ex As Exception
                            trans.Rollback()
                            Throw New Exception(ex.Message)
                        End Try
                        ''richa VIJ/20/11/19-000070
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
                        strCustomerCode = trnsLstCustomer.Item(j)
                        Try
                            clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                            Dim CashCustCount As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(TSPL_DEMAND_BOOKING_DETAIL.Cust_Code) as noofCust from TSPL_DEMAND_BOOKING_DETAIL
left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
where Document_No='" & strDocNo & "' and TSPL_CUSTOMER_MASTER.Cash_Customer='Y'
group by TSPL_DEMAND_BOOKING_DETAIL.Cust_Code"))
                            If CashCustCount = 0 Then
                                clsDemandBookingSale.PostData(clsUserMgtCode.frmDemandBooking, strDocNo, strCustomerCode)
                            Else
                                Throw New Exception("Cash Customer Document should be posted through Demand Screen")
                            End If
                            'If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, Nothing)) = 1 Or clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, Nothing)) = 1 Then
                            '    Dim isNewEntry As Boolean = False
                            '    Dim obj As clsDemandBookingSale = clsDemandBookingSale.GetData(strDocNo, NavigatorType.Current)
                            '    obj.Document_Date = clsCommon.myCDate(obj.Document_Date).AddDays(1)
                            '    obj.Document_No = clsDBFuncationality.getSingleValue("select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No='" + clsCommon.myCstr(obj.Route_No) + "' and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 103 )='" + clsCommon.GetPrintDate(obj.Document_Date) + "') and location_code='" + clsCommon.myCstr(obj.Location_Code) + "' and ShiftType='" + obj.ShiftType + "' and IsIndividualCustomer=0 ")
                            '    If clsCommon.myLen(obj.Document_No) > 0 Then
                            '        isNewEntry = False
                            '    Else
                            '        isNewEntry = True
                            '    End If
                            '    If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, Nothing)) = 1 Then
                            '        For ii As Integer = 0 To obj.Arr.Count - 1
                            '            If clsCommon.CompairString(obj.Arr(ii).Cust_Code, clsDBFuncationality.getSingleValue("Select Cust_Code from TSPL_CUSTOMER_MASTER where cust_code='" + clsCommon.myCstr(obj.Arr(ii).Cust_Code) + "' and IsReorder=1")) = CompairStringResult.Equal Then
                            '                obj.Arr(ii).Qty = obj.Arr(ii).Qty
                            '            Else
                            '                obj.Arr(ii).Qty = 0
                            '            End If
                            '        Next
                            '    End If
                            '    obj.SaveData(obj, isNewEntry)
                            'End If
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Booking/DO") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Fresh Booking/DO") = CompairStringResult.Equal Then
                        strCustomerCode = trnsLstCustomer.Item(j)
                        Try
                            clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                            clsBulkPostingDairySale.PostingAndDOCreation(strDocNo, strCustomerCode)
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try
                        ''richa BHA/27/12/18-000763 28 Dec,2018
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Dairy GatePass") = CompairStringResult.Equal Then
                        Try
                            clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                            clsDairyGatePassEntry.PostData(clsUserMgtCode.frmDairyGatePass, strDocNo)
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try

                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Crate Received") = CompairStringResult.Equal Then
                        Try
                            clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                            clsCrateReceivedHead.PostData(clsUserMgtCode.frmDairyGatePass, strDocNo)
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try
                    End If


                    countPostedDoc = countPostedDoc + 1
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try

        gv1.DataSource = Nothing
        FillDairySale()
    End Sub
    ''------------------------------------------------------
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        closeForm()
    End Sub


    Sub closeForm()
        Me.Close()
    End Sub

    ''-------------------------------------------------------------------
    '' Function For Validating----------GridView Column And Buttons------
    ''-------------------------------------------------------------------

    Sub Validation()
        If Isrefreshed = True Then
            Me.gv1.MasterTemplate.Columns("Status").ReadOnly = True    '' Column First Is Set To be ReadOnly
            btnPost.Enabled = False
            btnSlctAll.Enabled = False
        ElseIf Isrefreshed = False Then
            Me.gv1.MasterTemplate.Columns("Status").ReadOnly = False   '' Column First Is set to be Editable
            btnPost.Enabled = True
            btnSlctAll.Enabled = True
        End If
    End Sub


    Private Sub btnSlctAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSlctAll.Click
        SelectUnselectAll()
    End Sub

    ''---------------------------------------------------------------
    '' Function For Operating Select/UnSelect (btnSlctAll)----------
    ''---------------------------------------------------------------
    Sub SelectUnselectAll()
        For i As Integer = 0 To gv1.ChildRows.Count - 1
            gv1.ChildRows(i).Cells("Status").Value = False
        Next
        If btnSlctAll.Text = "Select TOP 100" Then
            'If gv1.ChildRows.Count <= 100 Then
            For i As Integer = 0 To gv1.ChildRows.Count - 1
                If i > 99 Then
                    Exit For
                End If
                gv1.ChildRows(i).Cells("Status").Value = True
            Next
            IsSelected = True
            btnSlctAll.Text = "UnSelect Top 100"
            'TotalAmount(True, gv1.CurrentRow.Index)
            TotalAmount(True, gv1)
            'Else
            '    For i As Integer = 0 To 99
            '        gv1.ChildRows(i).Cells("Status").Value = True
            '    Next
            '    IsSelected = False
            '    btnSlctAll.Text = "UnSelect Top 100"
            'End If
        ElseIf btnSlctAll.Text = "UnSelect Top 100" Then
            IsSelected = False
            btnSlctAll.Text = "Select TOP 100"
            txtGrandTotal.Text = ""
        End If


    End Sub



    'This will check the authorization of user to access the screen.If authorize then it will allow user to access the screen.
    'Private Function funSetUserAccess() As Boolean
    '    Try

    '        Dim strRights As String
    '        Dim strTemp() As String
    '        Dim strProgCode = "PENAPP"
    '        strRights = enuUserRights.enuRead & "," & enuUserRights.enuModify & "," & enuUserRights.enuDelete
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
    '            'btnsave.Enabled = False
    '        End If
    '        If strTemp(2) = "0" Then 'Grant modify access
    '            'btndelete.Enabled = False
    '        End If

    '        funSetUserAccess = True
    '    Catch er As Exception
    '        myMessages.myExceptions(er)
    '    End Try
    'End Function

    Private Sub FrmPendingAproval_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt And e.KeyCode = Keys.C Then
            closeForm()
        End If
    End Sub

    ''''--Addde By--Pankaj Kumar---on--08/06/2012-------Location Filter-----------
    Private Sub chkLOcAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs)
        cbgLocation.Enabled = False
    End Sub

    Private Sub chkLOcSelect_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs)
        cbgLocation.Enabled = True
    End Sub
    Public Sub LoadLocation()
        Dim qry As String = Nothing
        If clsCommon.myLen(arrLoc) > 0 Then
            qry = " Select Location_Code as Code, Location_Desc as Description from TSPL_LOCATION_MASTER Where Location_Type='Physical' and Location_Code in (" + arrLoc + ")"
        End If
        cbgLocation.DataSource = clsDBFuncationality.GetDataTable(qry)
        cbgLocation.ValueMember = "Code"
        cbgLocation.DisplayMember = "Description"
    End Sub

    Public Sub LoadUsers()
        Dim qry As String = clsUserMaster.GetSubbordinateUsersQry(objCommonVar.CurrentUserCode)
        cbgUser.DataSource = clsDBFuncationality.GetDataTable(qry)
        cbgUser.ValueMember = "User_Code"
        cbgUser.DisplayMember = "User_Name"
        If clsCommon.CompairString(objCommonVar.CurrentUserCode, "ADMIN") <> CompairStringResult.Equal Then
            Dim dtUser As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_USER_MAPPING_DETAIL where User_Code ='" & objCommonVar.CurrentUserCode & "'")
            If dtUser IsNot Nothing AndAlso dtUser.Rows.Count > 0 Then
                cbgUser.CheckedAll()
            End If
        End If
    End Sub

    Public Sub Load_Authorisation(ByVal ProgramName As String)

        ''StrQuery = "select max(TSPL_GROUP_PROGRAM_MAPPING.Authorized_Flag) as Authorized_Flag from TSPL_GROUP_PROGRAM_MAPPING " & _
        ' " inner join TSPL_Program_Master on TSPL_Program_Master.Program_Code=TSPL_GROUP_PROGRAM_MAPPING.Program_Code " & _
        '  " where TSPL_Program_Master.Program_Code='" + ProgramName + "' and " & _
        '  "TSPL_GROUP_PROGRAM_MAPPING.Group_Code in (select Group_Code  from TSPL_USER_GROUP_MAPPING where User_Code ='" + objCommonVar.CurrentUserCode + "') and TSPL_GROUP_PROGRAM_MAPPING.Authorized_Flag=1 "

        StrQuery = "select max(TSPL_GROUP_PROGRAM_MAPPING.Authorized_Flag) as Authorized_Flag from TSPL_GROUP_PROGRAM_MAPPING " &
        " inner join TSPL_Program_Master on TSPL_Program_Master.Program_Code=TSPL_GROUP_PROGRAM_MAPPING.Program_Code " &
         " where TSPL_Program_Master.Program_Code='" + ProgramName + "' and " &
         "TSPL_GROUP_PROGRAM_MAPPING.Group_Code in (select Group_Code  from TSPL_USER_GROUP_MAPPING where User_Code in (select '" + objCommonVar.CurrentUserCode + "' union  select TSPL_USER_MAPPING_DETAIL.Mapped_UserCode AS User_Code  from TSPL_USER_MAPPING_DETAIL Where TSPL_USER_MAPPING_DETAIL.User_Code='" + objCommonVar.CurrentUserCode + "' ) ) and TSPL_GROUP_PROGRAM_MAPPING.Authorized_Flag=1 "


        dtAuthen = clsDBFuncationality.GetDataTable(StrQuery)

    End Sub
    ''''----------------------------Code Ends Here--------------------------------
    '===============added by shivani tyagi only for store adjustment against ticket no [BM00000009002]
    Private Sub cboTransaction_SelectedValueChanged(sender As Object, e As EventArgs) Handles cboTransaction.SelectedValueChanged
        If isInsideLoad Then
            Exit Sub
        End If
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Store Adjustment") = CompairStringResult.Equal Then
            ChkMilkType.Visible = True
        Else
            ChkMilkType.Visible = False
        End If
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Demand") = CompairStringResult.Equal Then
            lblCustomer.Text = "Route"
            gbStatus.Text = "Shift "
            rbtnStatusPending.Text = "Morning"
            rbtnStatusPosted.Text = "Evening"
            RadLabel3.Visible = False
            dtpToDate.Visible = False

        Else
            lblCustomer.Text = "Customer"
            gbStatus.Text = "Status"
            rbtnStatusPending.Text = "Pending"
            rbtnStatusPosted.Text = "Posted"
            RadLabel3.Visible = True
            dtpToDate.Visible = True
        End If
    End Sub

    Private Sub chkLocAll_ToggleStateChanged_1(sender As Object, args As StateChangedEventArgs) Handles chkLocAll.ToggleStateChanged
        cbgLocation.Enabled = Not chkLocAll.IsChecked
    End Sub

    Private Sub ChkUserAll_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles ChkUserAll.ToggleStateChanged
        cbgUser.Enabled = Not ChkUserAll.IsChecked
    End Sub

    Private Sub txtCustomer__My_Click(sender As Object, e As EventArgs) Handles txtCustomer._My_Click
        Dim strQry As String = ""
        If clsCommon.CompairString(cboTransaction.Text, "Demand") = CompairStringResult.Equal Then
            strQry = " Select TSPL_ROUTE_MASTER.Route_No as Code,Route_Desc as Description,Type,Employee_Code as 'Employee Code',Off_Day as 'Off Day' from TSPL_ROUTE_MASTER"
            txtCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("TransTypeMulSel", strQry, "Code", "Name", txtCustomer.arrValueMember, txtCustomer.arrDispalyMember)
        Else
            strQry = " select Cust_Code as [code],Customer_Name as [Name] from TSPL_CUSTOMER_MASTER"
            txtCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("TransTypeMulSel", strQry, "Code", "Name", txtCustomer.arrValueMember, txtCustomer.arrDispalyMember)
        End If

    End Sub


    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpFromDate.Value = clsCommon.GETSERVERDATE()
        dtpToDate.Value = clsCommon.GETSERVERDATE()
        chkLocAll.CheckState = CheckState.Checked
        ChkUserAll.CheckState = CheckState.Checked
        rbtnStatusPending.IsChecked = True
        txtCustomer.arrValueMember = Nothing
        gv1.DataSource = Nothing
    End Sub
    '================Added by preeti Gupta Against Ticket no[][17/04/2019]
    Sub LaodModuleFixedAsset()
        Dim dt1 As DataTable = New DataTable()
        dt1.Columns.Add("Code", GetType(String))
        dt1.Columns.Add("Name", GetType(String))

        dr = dt1.NewRow()
        dr("Code") = "Acquisition Entry"
        dr("Name") = "Acquisition Entry"
        dt1.Rows.Add(dr)


        dr = dt1.NewRow()
        dr("Code") = "Disposal Entry"
        dr("Name") = "Disposal Entry"
        dt1.Rows.Add(dr)



        dr = dt1.NewRow()
        dr("Code") = "Asset Store Requisition"
        dr("Name") = "Asset Store Requisition"
        dt1.Rows.Add(dr)



        dr = dt1.NewRow()
        dr("Code") = "Issue Items to Assemble Asset"
        dr("Name") = "Issue Items to Assemble Asset"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Asset Work Expanses"
        dr("Name") = "Asset Work Expanses"
        dt1.Rows.Add(dr)

        dr = dt1.NewRow()
        dr("Code") = "Asset Merging Entry"
        dr("Name") = "Asset Merging Entry"
        dt1.Rows.Add(dr)

        cboTransaction.DataSource = dt1
        cboTransaction.DisplayMember = "Name"
        cboTransaction.ValueMember = "Code"
    End Sub
    Sub PostFixedAsset()
        Try
            For j As Integer = 0 To trnsLst.Count - 1
                Try
                    clsCommon.ProgressBarPercentUpdate((((j + 1) * 100) / (trnsLst.Count)), "Document Posting " + clsCommon.myCstr(j + 1) + "/" + clsCommon.myCstr(trnsLst.Count))
                    strDocNo = trnsLst.Item(j)
                    If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Acquisition Entry") = CompairStringResult.Equal Then
                        clsAcquisitionHead.PostData("FA-ACQE", strDocNo, True, Nothing)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Disposal Entry") = CompairStringResult.Equal Then
                        clsAssetScrapSaleHead.PostData(strDocNo)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Asset Store Requisition") = CompairStringResult.Equal Then
                        clsRequistionHead.PostData(strDocNo)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Issue Items to Assemble Asset") = CompairStringResult.Equal Then
                        clsItemIssueToAssembledAsset.PostData(strDocNo)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Asset Work Expanses") = CompairStringResult.Equal Then
                        Dim objAssetWorkHead As New clsAssetWorkHead()
                        objAssetWorkHead.PostData("FA-WORK", strDocNo)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Asset Merging Entry") = CompairStringResult.Equal Then
                        clsFAMergeHead.PostData("FA-MRG-ACE", strDocNo, True, "NEW")
                    End If

                    countPostedDoc = countPostedDoc + 1
                Catch ex As Exception
                    dr = DtError.NewRow()
                    dr("Code") = strDocNo
                    dr("Error") = ex.Message
                    DtError.Rows.Add(dr)
                End Try
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return
        End Try
        gv1.DataSource = Nothing
        FillFixeedAsset()
    End Sub
    '===============Added by preeti Gupta Against ticket no[ERO/16/04/19-000558]
    Sub FillFixeedAsset()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Acquisition Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FAAcquisitionEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_ACQUISITION_head.Status)as BIT) as Status,TSPL_ACQUISITION_head.Acquisition_Code as[Document Id],convert(varchar,TSPL_ACQUISITION_head.acquisition_date,103) as [Document Date]," &
                    " isnull(TSPL_ACQUISITION_head.net_amt,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_ACQUISITION_head.Created_By as [Created By], " &
                    " convert(varchar,TSPL_ACQUISITION_head.Created_Date,103) as [Created Date],TSPL_ACQUISITION_head.Description from TSPL_ACQUISITION_head " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_ACQUISITION_head.Loc_Code=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE Acquisition_Type<>'Merge' AND  convert(date,TSPL_ACQUISITION_head.acquisition_date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_ACQUISITION_head.acquisition_date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_ACQUISITION_head.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_ACQUISITION_head.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_ACQUISITION_head.Loc_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_ACQUISITION_head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ACQUISITION_head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_ACQUISITION_head.acquisition_date, TSPL_ACQUISITION_head.Acquisition_Code "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Disposal Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FADisposalEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_ASSET_SCRAP_HEAD.Status)as BIT) as Status,TSPL_ASSET_SCRAP_HEAD.Document_No as[Document Id],convert(varchar,TSPL_ASSET_SCRAP_HEAD.Document_Date,103) as [Document Date]," &
                    " isnull(TSPL_ASSET_SCRAP_HEAD.Doc_Amt,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_ASSET_SCRAP_HEAD.Created_By as [Created By], " &
                    " convert(varchar,TSPL_ASSET_SCRAP_HEAD.Created_Date,103) as [Created Date],TSPL_ASSET_SCRAP_HEAD.Description from TSPL_ASSET_SCRAP_HEAD " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_ASSET_SCRAP_HEAD.Loc_Code=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_ASSET_SCRAP_HEAD.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_ASSET_SCRAP_HEAD.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_ASSET_SCRAP_HEAD.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_ASSET_SCRAP_HEAD.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_ASSET_SCRAP_HEAD.Loc_code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_ASSET_SCRAP_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ASSET_SCRAP_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_ASSET_SCRAP_HEAD.Document_Date, TSPL_ASSET_SCRAP_HEAD.Document_No "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Asset Store Requisition") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmAssetStoreRequistion)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_REQUISITION_HEAD.status)as BIT) as Status,TSPL_REQUISITION_HEAD.Requisition_Id as[Document Id],convert(varchar,TSPL_REQUISITION_HEAD.Requisition_Date,103) as [Document Date]," &
                    " isnull(TSPL_REQUISITION_HEAD.total_RQ_Amt ,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_REQUISITION_HEAD.Created_By as [Created By], " &
                    " convert(varchar,TSPL_REQUISITION_HEAD.Created_Date,103) as [Created Date],TSPL_REQUISITION_HEAD.Description from TSPL_REQUISITION_HEAD " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_REQUISITION_HEAD.Location=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_REQUISITION_HEAD.Requisition_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_REQUISITION_HEAD.Requisition_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and isnull(TSPL_REQUISITION_HEAD.status,'') = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_REQUISITION_HEAD.status,'') = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_REQUISITION_HEAD.Location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_REQUISITION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_REQUISITION_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_REQUISITION_HEAD.Requisition_Id, TSPL_REQUISITION_HEAD.Requisition_Date "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Issue Items to Assemble Asset") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.frmIssueItemsToAsset)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_IssueItemToAssembledAsset_Head.status)as BIT) as Status,TSPL_IssueItemToAssembledAsset_Head.doc_no as[Document Id],convert(varchar,TSPL_IssueItemToAssembledAsset_Head.doc_date,103) as [Document Date]," &
                    " isnull(TSPL_IssueItemToAssembledAsset_Head.Doc_Amt ,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_IssueItemToAssembledAsset_Head.Created_By as [Created By], " &
                    " convert(varchar,TSPL_IssueItemToAssembledAsset_Head.Created_Date,103) as [Created Date],TSPL_IssueItemToAssembledAsset_Head.REMARKS AS Description from TSPL_IssueItemToAssembledAsset_Head " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_IssueItemToAssembledAsset_Head.From_Location=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_IssueItemToAssembledAsset_Head.doc_date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_IssueItemToAssembledAsset_Head.doc_date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and isnull(TSPL_IssueItemToAssembledAsset_Head.status,'') = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_IssueItemToAssembledAsset_Head.status,'') = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_IssueItemToAssembledAsset_Head.From_Location  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_IssueItemToAssembledAsset_Head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_IssueItemToAssembledAsset_Head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_IssueItemToAssembledAsset_Head.doc_no, TSPL_IssueItemToAssembledAsset_Head.doc_date "

                End If
            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Asset Work Expanses") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FAAssetWork)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_ASSET_WORK_HEAD.status)as BIT) as Status,TSPL_ASSET_WORK_HEAD.document_code as[Document Id],convert(varchar,TSPL_ASSET_WORK_HEAD.document_date,103) as [Document Date]," &
                    " isnull(TSPL_ASSET_WORK_HEAD.Net_Amt ,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_ASSET_WORK_HEAD.Created_By as [Created By], " &
                    " convert(varchar,TSPL_ASSET_WORK_HEAD.Created_Date,103) as [Created Date],TSPL_ASSET_WORK_HEAD.Description from TSPL_ASSET_WORK_HEAD " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_ASSET_WORK_HEAD.Location_Code=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE  convert(date,TSPL_ASSET_WORK_HEAD.document_date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_ASSET_WORK_HEAD.document_date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and isnull(TSPL_ASSET_WORK_HEAD.status,'') = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and isnull(TSPL_ASSET_WORK_HEAD.status,'') = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_ASSET_WORK_HEAD.Location_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_ASSET_WORK_HEAD.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ASSET_WORK_HEAD.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_ASSET_WORK_HEAD.document_code, TSPL_ASSET_WORK_HEAD.document_date "

                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Asset Merging Entry") = CompairStringResult.Equal Then
            Load_Authorisation(clsUserMgtCode.FAMergeAcquisitionEntry)
            If dtAuthen.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtAuthen.Rows(0)("Authorized_Flag")), "1") = CompairStringResult.Equal Then
                    gv1.DataSource = Nothing

                    qry = " Select CAST((TSPL_ACQUISITION_head.Status)as BIT) as Status,TSPL_ACQUISITION_head.Acquisition_Code as[Document Id],convert(varchar,TSPL_ACQUISITION_head.acquisition_date,103) as [Document Date]," &
                    " isnull(TSPL_ACQUISITION_head.net_amt,0) as [Amount],TSPL_LOCATION_MASTER.Location_Desc as [Location], " &
                    " TSPL_ACQUISITION_head.Created_By as [Created By], " &
                    " convert(varchar,TSPL_ACQUISITION_head.Created_Date,103) as [Created Date],TSPL_ACQUISITION_head.Description from TSPL_ACQUISITION_head " &
                    " left join TSPL_LOCATION_MASTER  on TSPL_ACQUISITION_head.Loc_Code=TSPL_LOCATION_MASTER.Location_Code  " &
                    " WHERE Acquisition_Type='Merge' AND  convert(date,TSPL_ACQUISITION_head.acquisition_date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) and convert(date,TSPL_ACQUISITION_head.acquisition_date,103) <= convert(date,'" + dtpToDate.Value + "',103) "

                    If rbtnStatusPending.IsChecked = True Then
                        qry += " and TSPL_ACQUISITION_head.Status = 0"
                        Isrefreshed = False
                    Else
                        qry += "  and TSPL_ACQUISITION_head.Status = 1"
                        Isrefreshed = True
                    End If
                    If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                        qry += " and TSPL_ACQUISITION_head.Loc_Code  in   (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
                    End If
                    If ChkAllowBulkPosting.Equals(0) Then
                        If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then

                            qry += " and TSPL_ACQUISITION_head.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                        Else
                            qry += " and TSPL_ACQUISITION_head.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                        End If

                    End If
                    qry += " ORDER BY TSPL_ACQUISITION_head.acquisition_date, TSPL_ACQUISITION_head.Acquisition_Code "

                End If
            End If
        End If


        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            Validation()
            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" + strNoOfRecord + " Records Found"
            End If
        End If
    End Sub

    Private Sub btnSel1000_Click(sender As Object, e As EventArgs) Handles btnSel1000.Click
        For i As Integer = 0 To gv1.ChildRows.Count - 1
            gv1.ChildRows(i).Cells("Status").Value = False
        Next
        If btnSel1000.Text = "Select TOP 1000" Then
            For i As Integer = 0 To gv1.ChildRows.Count - 1
                If i > 999 Then
                    Exit For
                End If
                gv1.ChildRows(i).Cells("Status").Value = True
            Next
            IsSelected = True
            btnSel1000.Text = "UnSelect Top 1000"
            TotalAmount(True, gv1)
        ElseIf btnSel1000.Text = "UnSelect Top 1000" Then
            IsSelected = False
            btnSel1000.Text = "Select TOP 1000"
            txtGrandTotal.Text = ""
        End If
    End Sub


    Private Sub BtnSelAll_Click(sender As Object, e As EventArgs) Handles BtnSelAll.Click
        'Try
        '    FlagAllSelectWorking = True
        '    Dim TempTotal As Double = 0
        '    For i As Integer = 0 To gv1.Rows.Count - 1
        '        gv1.Rows(i).Cells("Status").Value = True
        '        TempTotal = TempTotal + clsCommon.myCdbl(gv1.Rows(i).Cells("Amount"))
        '    Next
        '    txtGrandTotal.Text = TempTotal
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        'Finally
        '    FlagAllSelectWorking = False
        'End Try
        Try
            FlagAllSelectWorking = True
            If clsCommon.CompairString(BtnSelAll.Text, "Select All") = CompairStringResult.Equal Then
                Dim TempTotal As Double = 0
                For i As Integer = 0 To gv1.ChildRows.Count - 1
                    gv1.ChildRows(i).Cells("Status").Value = True
                    TempTotal = TempTotal + clsCommon.myCdbl(gv1.ChildRows(i).Cells("Amount"))
                Next
                txtGrandTotal.Text = TempTotal
                IsSelected = True
                BtnSelAll.Text = "UnSelect All"


                'For i As Integer = 0 To gv1.Rows.Count - 1
                '    'gv1.ChildRows(i).IsVisible = True
                '    If i > gv1.Rows.Count - 1 Then
                '        Exit For
                '    End If
                '    gv1.Rows(i).Cells("Status").Value = False
                'Next


            ElseIf clsCommon.CompairString(BtnSelAll.Text, "UnSelect All") = CompairStringResult.Equal Then
                For i As Integer = 0 To gv1.ChildRows.Count - 1
                    'gv1.ChildRows(i).IsVisible = True
                    'If i > gv1.Rows.Count - 1 Then
                    '    Exit For
                    'End If
                    gv1.ChildRows(i).Cells("Status").Value = False
                Next
                IsSelected = False
                BtnSelAll.Text = "Select All"
                txtGrandTotal.Text = ""
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            FlagAllSelectWorking = False
        End Try

        'For i As Integer = 0 To gv1.ChildRows.Count - 1
        '    'gv1.ChildRows(i).IsVisible = True
        '    If i > gv1.ChildRows.Count - 1 Then
        '        Exit For
        '    End If
        '    gv1.ChildRows(i).Cells("Status").Value = True
        'Next
        '    IsSelected = True
        '    BtnSelAll.Text = "UnSelect All"
        ''    TotalAmount(True, gv1)
        'ElseIf clsCommon.CompairString(BtnSelAll.Text, "UnSelect All") = CompairStringResult.Equal Then
        'IsSelected = False
        'BtnSelAll.Text = "Select All"
        'txtGrandTotal.Text = ""
        'End If
    End Sub

    Private Sub gv1_FilterChanged(sender As Object, e As GridViewCollectionChangedEventArgs) Handles gv1.FilterChanged
        lblNoOfRecords.Text = "" + gv1.ChildRows.Count.ToString + " Records Found"
    End Sub

    Private Sub cboTransaction_SelectedIndexChanged(sender As Object, e As UI.Data.PositionChangedEventArgs) Handles cboTransaction.SelectedIndexChanged
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), "Product Demand Booking") = CompairStringResult.Equal Then
            lblCustomer.Visible = False
            txtCustomer.Visible = False
        End If
    End Sub
End Class
