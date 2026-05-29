''Created By--richa agarwal-TEC/10/09/19-001006----on- 10/09/2019---------------------
Imports common
Imports System.Data.SqlClient



Public Class frmCancelledTransactions_DairySale
    Inherits FrmMainTranScreen
    Dim arrUser As New ArrayList()
    Dim ButtonToolTip As New ToolTip()
    Public ModuleName As String = ""
    Public Transaction As String = ""
    Public fromdate As DateTime
    Public Todate As DateTime
    Dim dr As DataRow
    Dim dt As DataTable
    Dim strNoOfRecord As String
    Dim qry As String
    Dim arrSelectedUser As New ArrayList()
    Dim arrLoc As String = Nothing
    Dim RecordCount As Integer = 0

    Private Sub FrmPendingAproval_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LOCATIONRIGTHS()
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C for Closing The Window")
        LoadLocation()
        chkLocAll.CheckState = CheckState.Checked
        ChkUserAll.CheckState = CheckState.Checked
        LoadUsers()
        arrUser = GetSubbordinateUsers(objCommonVar.CurrentUserCode)
        SetUserMgmtNew()
        LoadBlankGrid()
        LoadModuleType()
        dtpFromDate.Value = clsCommon.GETSERVERDATE()
        dtpToDate.Value = clsCommon.GETSERVERDATE()
        LoadModuleProduction()
        If clsCommon.myLen(ModuleName) > 0 Then
            cboModule.SelectedValue = ModuleName
            cboTransaction.SelectedValue = Transaction
            dtpFromDate.Value = fromdate
            dtpToDate.Value = Todate
            ShowData()
        End If
        MSGShowAndHide()
    End Sub
    Public Sub LoadModuleType()
        Dim Qry As String = "select Distinct TBL_MODULE.Program_Code As [Module Code],case when len (isnull(TBL_MODULE.Re_Name,'')) > 0 then TBL_MODULE.Re_Name else  TSPL_PROGRAM_MASTER.Program_Name end  as [Module Name] 
from TSPL_PROGRAM_MASTER
left outer join (select Program_Code, Program_Name,Parent_Code,case when len (isnull(TSPL_PROGRAM_MASTER.Re_Name,'')) > 0 then TSPL_PROGRAM_MASTER.Re_Name else  TSPL_PROGRAM_MASTER.Program_Name end As Re_Name from TSPL_PROGRAM_MASTER where Type in ('SM')) as TBL_SMODULE on TBL_SMODULE.Program_Code = TSPL_PROGRAM_MASTER.Parent_Code
left outer join (select Program_Code, Program_Name,Parent_Code,case when len (isnull(TSPL_PROGRAM_MASTER.Re_Name,'')) > 0 then TSPL_PROGRAM_MASTER.Re_Name else  TSPL_PROGRAM_MASTER.Program_Name end As Re_Name from TSPL_PROGRAM_MASTER where Type in ('M')) as TBL_MODULE on TBL_MODULE.Program_Code = TBL_SMODULE.Parent_Code
Where TBL_MODULE.Program_Code in (select  distinct Module_Name from TSPL_MODULE_PERMISSION ) and  not TSPL_PROGRAM_MASTER.Type in ('M','SM') 
and TBL_MODULE.Program_Code in ('" & clsUserMgtCode.ModulePurchase & "','" & clsUserMgtCode.ModuleReceivable & "','" & clsUserMgtCode.ModuleSaleDairy & "','" & clsUserMgtCode.ModuleSalesNew & "','" & clsUserMgtCode.SubModuleSaleNewTransaction & "','" & clsUserMgtCode.ModuleMCCMilkProcurement & "','" & clsUserMgtCode.ModulePurchase & "','" & clsUserMgtCode.ModulePayable & "','" & clsUserMgtCode.ModuleMaterial & "') 
and TBL_SMODULE.Program_Name in ('Transaction','MCC Transaction','Bulk Transaction') 
 "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            'dr = dt.NewRow()
            'dr("Module Code") = dt.Rows(0)("Module Code")
            'dr("Module Name") = dt.Rows(0)("Module Name")
            'dt.Rows.Add(dr)
            cboModule.DataSource = dt
            cboModule.DisplayMember = "Module Name"
            cboModule.ValueMember = "Module Code"
        End If

    End Sub

    Sub LoadModuleProduction()
        Try
            Dim Qry As String = "select TSPL_PROGRAM_MASTER.Program_Code as Code,case when len (isnull(TSPL_PROGRAM_MASTER.Re_Name,'')) > 0 then TSPL_PROGRAM_MASTER.Re_Name else  TSPL_PROGRAM_MASTER.Program_Name end  as Name 
from TSPL_PROGRAM_MASTER
left outer join (select Program_Code, Program_Name,Parent_Code,case when len (isnull(TSPL_PROGRAM_MASTER.Re_Name,'')) > 0 then TSPL_PROGRAM_MASTER.Re_Name else  TSPL_PROGRAM_MASTER.Program_Name end As Re_Name from TSPL_PROGRAM_MASTER where Type in ('SM')) as TBL_SMODULE on TBL_SMODULE.Program_Code = TSPL_PROGRAM_MASTER.Parent_Code
left outer join (select Program_Code, Program_Name,Parent_Code,case when len (isnull(TSPL_PROGRAM_MASTER.Re_Name,'')) > 0 then TSPL_PROGRAM_MASTER.Re_Name else  TSPL_PROGRAM_MASTER.Program_Name end As Re_Name from TSPL_PROGRAM_MASTER where Type in ('M')) as TBL_MODULE on TBL_MODULE.Program_Code = TBL_SMODULE.Parent_Code
Where TBL_MODULE.Program_Code in (select  distinct Module_Name from TSPL_MODULE_PERMISSION ) and  not TSPL_PROGRAM_MASTER.Type in ('M','SM') 
and TBL_SMODULE.Parent_Code In ('" & clsCommon.myCstr(cboModule.SelectedValue) & "') 
and TBL_SMODULE.Program_Name in ('Transaction','MCC Transaction','Bulk Transaction') And TSPL_PROGRAM_MASTER.Program_Code In ('" & clsUserMgtCode.frmSaleDispatchDairy & "','" & clsUserMgtCode.frmSNSaleInvoice & "','" & clsUserMgtCode.frmDairyGatePass & "','" & clsUserMgtCode.frmMCCMaterial & "','" & clsUserMgtCode.frmDairyBookingCustomer & "','" & clsUserMgtCode.frmWreckageBooking & "','" & clsUserMgtCode.frmSNPOS & "','" & clsUserMgtCode.frmGatePassDairy & "','" & clsUserMgtCode.ScrapSale & "','" & clsUserMgtCode.FrmVendorService & "','" & clsUserMgtCode.frmSaleReturndairy & "','" & clsUserMgtCode.frmSaleInvoicedairy & "','" & clsUserMgtCode.Transfer & "','" + clsUserMgtCode.mbtnAPInvoiceEntry + "','" + clsUserMgtCode.mbtnARInvoiceEntry + "','" + clsUserMgtCode.ScrapSaleRetrun + "')
 "
            dt = clsDBFuncationality.GetDataTable(Qry)
            'dr = dt.NewRow()
            'dr("Code") = dt.Rows(0)("Code")
            'dr("Name") = dt.Rows(0)("Name")
            'dt.Rows.Add(dr)

            cboTransaction.DataSource = dt
            cboTransaction.DisplayMember = "Name"
            cboTransaction.ValueMember = "Code"
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub LOCATIONRIGTHS()
        Dim obj As New clsMCCCodes()
        Try
            obj = clsMCCCodes.GetData()

            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Default_LocCode) > 0 Then
                Dim check As Integer = 0
                check = clsDBFuncationality.getSingleValue("select count(*) from tspl_location_master where type='Plant' and location_code='" & obj.Default_LocCode & "'")
                'If check > 0 Then

                'End If
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
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
    End Sub


    Private Sub cboModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs) Handles cboModule.SelectedIndexChanged
        Try
            LoadTrnsListOfSelectedModeule()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Public Sub LoadTrnsListOfSelectedModeule()
        cboTransaction.DataSource = Nothing
        LoadModuleProduction()
    End Sub

    Sub LoadBlankGrid()
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        gv1.AllowAddNewRow = False
    End Sub
    Sub gv1Format()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleInvoicedairy) = CompairStringResult.Equal Then
            Me.gv1.MasterTemplate.Columns("IsMultipleInvoice").IsVisible = False
        End If

        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmDairyGatePass) = CompairStringResult.Equal Then
            Me.gv1.MasterTemplate.Columns("GatePass No").Width = 120      ''First Column
            Me.gv1.MasterTemplate.Columns("Gate Pass Date").Width = 150
        Else
            Me.gv1.MasterTemplate.Columns("Document Id").Width = 120      ''First Column
            Me.gv1.MasterTemplate.Columns("Document Date").Width = 150    ''Second Column
        End If

        Dim count As Integer = gv1.MasterTemplate.Columns.Count
        For i As Integer = 2 To count - 1
            Me.gv1.MasterTemplate.Columns(i).Width = 120
        Next
        Me.gv1.MasterTemplate.Columns("Description").Width = 200    ''Last Column
        For j As Integer = 0 To count - 1
            Me.gv1.MasterTemplate.Columns(j).ReadOnly = True
        Next
        Dim summaryRowItem As New GridViewSummaryRowItem()
        gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

#Region "Showing Details on GRID"
    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        If dtpFromDate.Value > dtpToDate.Value Then
            common.clsCommon.MyMessageBoxShow(Me, "'From date' Can't Be Greater Than 'To Date'", Me.Text)
        Else
            qry = Nothing
            ShowData()
        End If
    End Sub
    Function GetSubbordinateUsers(ByVal strUserCode As String) As ArrayList
        Dim arrUser As New ArrayList
        Try
            Dim qry As String
            If clsCommon.CompairString(objCommonVar.CurrentUserCode, "ADMIN") = CompairStringResult.Equal Then
                qry = "Select User_Code from TSPL_User_MASTER Where 1=1"
            Else
                qry = "Select User_Code from TSPL_User_MASTER Where Level4_Code='" & strUserCode & "' OR User_Code='" & strUserCode & "'"
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            For Each dr As DataRow In dt.Rows
                arrUser.Add(clsCommon.myCstr(dr("User_Code")))
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return arrUser
    End Function
    Sub ShowData()
        Try
            'If cbgUser.CheckedValue.Count > 0 Then
            '    arrSelectedUser = cbgUser.CheckedValue
            'Else
            '    arrSelectedUser = arrUser
            'End If
            gv1.DataSource = Nothing
            FillDairySale()
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub FillDairySale()
        If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleDispatchDairy) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmMCCMaterial) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "  Select TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code as [Document Id],convert(varchar,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_date ,103) as [Document Date]," &
                    " TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_delivery_Code as [Against Delivery No], TSPL_CUSTOMER_MASTER.Cust_Code As [Distributor Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Distributor Name] , TSPL_SD_SHIPMENT_HEAD_Delete_Data.Route_No As [Route No],TSPL_SD_SHIPMENT_HEAD_Delete_Data.Route_Desc As [Route Name],Convert(varchar,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Supply_Date,103) As [Supply Date]," &
                    " Case When TSPL_SD_SHIPMENT_HEAD_Delete_Data.Shift_Type='AM' Then 'Morning' Else 'Evening' End As [Supply Shift],Case When TSPL_SD_SHIPMENT_HEAD_Delete_Data.Is_Taxable=0 Then 'Non-Taxable' Else 'Taxable' End As [Invoice Type],TSPL_SD_SHIPMENT_HEAD_Delete_Data.Bill_To_Location as [Location Code]," &
                    " TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_SD_SHIPMENT_HEAD_Delete_Data.Created_By as [Created By]," &
                    " (convert(varchar,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Created_Date,103)+' '+convert(varchar,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Created_Date,108)) as [Created Date],'' as Description,Sale_Invoice_No as [Invoice No],
CONVERT(varchar, TSPL_SD_SHIPMENT_HEAD_Delete_Data.Sale_Invoice_Date,103) AS [Invoice Date],TSPL_SD_SHIPMENT_HEAD_Delete_Data.Total_Amt as [Invoice Amount],IRN_No as[IRN No] ,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Ack_No as [Ack No],
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Ack_Date as [Ack Date],TSPL_SD_SHIPMENT_HEAD_Delete_Data.Delete_By AS [Deleted By],(CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Delete_On,103)+' '+CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Delete_On,108)) AS [Deleted Date]
                 from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data 
                    left outer join TSPL_SD_SHIPMENT_HEAD_Delete_Data on TSPL_SD_SHIPMENT_HEAD_Delete_Data.Sale_Invoice_No=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code " &
                    " Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD_Delete_Data.Customer_Code
 Left Outer Join TSPL_LOCATION_MASTER  on TSPL_SD_SHIPMENT_HEAD_Delete_Data.Bill_To_Location  =TSPL_LOCATION_MASTER.Location_Code " &
                " WHERE "
                If rbtnCancelDate.IsChecked Then
                    qry += "convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) " &
    " and convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) " &
    " and convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmMCCMaterial) = CompairStringResult.Equal Then
                    qry += " and  TSPL_SD_SHIPMENT_HEAD_Delete_Data.Trans_Type IN ('MCC')"
                Else
                    qry += " and  TSPL_SD_SHIPMENT_HEAD_Delete_Data.Trans_Type IN ('FS', 'PS') and TSPL_SD_SHIPMENT_HEAD_Delete_Data.Screen_Type='DS'"
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SD_SHIPMENT_HEAD_Delete_Data.Bill_To_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SHIPMENT_HEAD_Delete_Data.Created_By IN (" & clsCommon.GetMulcallString(cbgUser.CheckedValue) & ")"
                'Else
                '    qry += " and TSPL_SD_SHIPMENT_HEAD_Delete_Data.Created_By IN (" & clsCommon.GetMulcallString(arrSelectedUser) & ")"
                'End If
                qry += " ORDER BY TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_date,TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code "
            Else
                gv1.DataSource = Nothing
                qry = "  Select TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code as [Document Id],convert(varchar,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_date ,103) as [Document Date]," &
                " TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_delivery_Code as [Against Delivery No], TSPL_CUSTOMER_MASTER.Cust_Code As [Distributor Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Distributor Name] , TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Route_No As [Route No],TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Route_Desc As [Route Name],Convert(varchar,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Supply_Date,103) As [Supply Date]," &
                " Case When TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Shift_Type='AM' Then 'Morning' Else 'Evening' End As [Supply Shift],Case When TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Is_Taxable=0 Then 'Non-Taxable' Else 'Taxable' End As [Invoice Type],TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Bill_To_Location as [Location Code]," &
                " TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Created_By as [Created By]," &
                " (convert(varchar,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Created_Date,103)+' '+convert(varchar,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Created_Date,108)) as [Created Date],'' as Description,Sale_Invoice_No as [Invoice No],
CONVERT(varchar, TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Sale_Invoice_Date,103) AS [Invoice Date],TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Total_Amt as [Invoice Amount],IRN_No as[IRN No] ,TSPL_SD_SALE_INVOICe_HEAD_Cancel_Data.Ack_No as [Ack No],
TSPL_SD_SALE_INVOICe_HEAD_Cancel_Data.Ack_Date as [Ack Date],TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Cancel_By AS [Canceled By],(CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Cancel_On,103)+' '+CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Cancel_On,108)) AS [Canceled Date]
                 from TSPL_SD_SHIPMENT_HEAD_Cancel_Data left outer join TSPL_SD_SALE_INVOICe_HEAD_Cancel_Data on TSPL_SD_SALE_INVOICe_HEAD_Cancel_Data.Document_Code=TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Sale_Invoice_No " &
                " Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Customer_Code
 Left Outer Join TSPL_LOCATION_MASTER  on TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Bill_To_Location  =TSPL_LOCATION_MASTER.Location_Code " &
            " WHERE "
                If rbtnCancelDate.IsChecked Then
                    qry += "convert(date,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Cancel_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) " &
" and convert(date,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "convert(date,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) " &
" and convert(date,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmMCCMaterial) = CompairStringResult.Equal Then
                    qry += " and  TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Trans_Type IN ('MCC')"
                Else
                    qry += " and  TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Trans_Type IN ('FS', 'PS') and TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Screen_Type='DS'"
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Bill_To_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
                qry += " ORDER BY TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_date,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code "

            End If


        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.mbtnAPInvoiceEntry) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "Select TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Document_No  as [Document Id],convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Invoice_Entry_Date ,103) as [Document Date] 
                  ,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Vendor_Code as [Customer Code],TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Vendor_Name as [Customer Name] 
                  , TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Document_Type as [Document Type] 
                  , TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Document_Total as [Document Amount]
                  ,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Created_By as [Created By] 
                  , convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Created_Date,103) as [Created Date], Description,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Delete_BY as [Deleted By],convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Delete_On,103) as [Delete Date] from TSPL_VENDOR_INVOICE_HEAD_Delete_Data 
            WHERE  (Against_POInvoice_No IS NULL OR Against_POInvoice_No = '')
AND (Against_PurchaseReturn_No IS NULL OR Against_PurchaseReturn_No = '')
AND (Against_Acquisition IS NULL OR Against_Acquisition = '')
AND (Against_MillkPurchaseInvoice_No IS NULL OR Against_MillkPurchaseInvoice_No = '')
AND (Against_BulkMillkPurchaseInvoice_No IS NULL OR Against_BulkMillkPurchaseInvoice_No = '')
AND (Against_VSPItemIssue_No IS NULL OR Against_VSPItemIssue_No = '')
AND (Against_MCC_Material_Sale IS NULL OR Against_MCC_Material_Sale = '')
AND (Against_VSP_Asset_Issue IS NULL OR Against_VSP_Asset_Issue = '')
AND (Against_MCC_Material_Sale_Return IS NULL OR Against_MCC_Material_Sale_Return = '')
AND (Against_TransferToSavingPKID IS NULL OR Against_TransferToSavingPKID = '') and  "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Delete_On ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Delete_On,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                Else
                    qry += "convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Invoice_Entry_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Invoice_Entry_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)  "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_VENDOR_INVOICE_HEAD_Delete_Data.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += "ORDER BY TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Invoice_Entry_Date,TSPL_VENDOR_INVOICE_HEAD_Delete_Data.Document_No"

            Else
                gv1.DataSource = Nothing
                qry = "Select TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Document_No  as [Document Id],convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Invoice_Entry_Date ,103) as [Document Date] 
                  ,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Vendor_Code as [Customer Code],TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Vendor_Name as [Customer Name] 
                  , TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Document_Type as [Document Type] 
                  , TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Document_Total as [Document Amount]
                  ,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Created_By as [Created By] 
                  , convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Created_Date,103) as [Created Date], Description,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Cancel_By as [Cancelled By],convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Cancel_On,103) as [Cancelled Date] from TSPL_VENDOR_INVOICE_HEAD_Cancel_Data 
            WHERE (Against_POInvoice_No IS NULL OR Against_POInvoice_No = '')
AND (Against_PurchaseReturn_No IS NULL OR Against_PurchaseReturn_No = '')
AND (Against_Acquisition IS NULL OR Against_Acquisition = '')
AND (Against_MillkPurchaseInvoice_No IS NULL OR Against_MillkPurchaseInvoice_No = '')
AND (Against_BulkMillkPurchaseInvoice_No IS NULL OR Against_BulkMillkPurchaseInvoice_No = '')
AND (Against_VSPItemIssue_No IS NULL OR Against_VSPItemIssue_No = '')
AND (Against_MCC_Material_Sale IS NULL OR Against_MCC_Material_Sale = '')
AND (Against_VSP_Asset_Issue IS NULL OR Against_VSP_Asset_Issue = '')
AND (Against_MCC_Material_Sale_Return IS NULL OR Against_MCC_Material_Sale_Return = '')
AND (Against_TransferToSavingPKID IS NULL OR Against_TransferToSavingPKID = '') and "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_VENDOR_INVOICE_HEAD_CANCEL_DATA.cancel_on ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.cancel_on,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                Else
                    qry += "  convert(date,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Invoice_Entry_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Invoice_Entry_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += " ORDER BY TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Invoice_Entry_Date,TSPL_VENDOR_INVOICE_HEAD_Cancel_Data.Document_No"

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
            End If





        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.mbtnARInvoiceEntry) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "Select TSPL_Customer_Invoice_Head_Delete_DATA.Document_No  as [Document Id],convert(varchar,TSPL_Customer_Invoice_Head_Delete_DATA.Document_Date ,103) as [Document Date] " &
                  ",TSPL_Customer_Invoice_Head_Delete_DATA.Customer_Code as [Customer Code],TSPL_Customer_Invoice_Head_Delete_DATA.Customer_Name as [Customer Name] " &
                  ", TSPL_Customer_Invoice_Head_Delete_DATA.Document_Type as [Document Type] " &
                  ", TSPL_Customer_Invoice_Head_Delete_DATA.Document_Total as [Document Amount] " &
                  ",TSPL_Customer_Invoice_Head_Delete_DATA.Created_By as [Created By] " &
                  ", convert(varchar,TSPL_Customer_Invoice_Head_Delete_DATA.Created_Date,103) as [Created Date], Description,TSPL_Customer_Invoice_Head_Delete_DATA.Delete_By as [Delete By],convert(varchar,TSPL_Customer_Invoice_Head_Delete_DATA.Delete_On,103) as [Delete Date] from TSPL_Customer_Invoice_Head_Delete_DATA " &
                  " WHERE    AgainstServiceInvoice ='Y' and  "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_Customer_Invoice_Head_Delete_Data.Delete_On ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_Customer_Invoice_Head_Delete_Data.Delete_On,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                Else
                    qry += "convert(date,TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)  "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_Customer_Invoice_Head_Delete_DATA.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += "ORDER BY TSPL_Customer_Invoice_Head_Delete_DATA.Document_Date,TSPL_Customer_Invoice_Head_Delete_Data.Document_No"
            Else
                gv1.DataSource = Nothing
                qry = "Select TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_No  as [Document Id],convert(varchar,TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Date ,103) as [Document Date] " &
                  ",TSPL_Customer_Invoice_Head_CANCEL_DATA.Customer_Code as [Customer Code],TSPL_Customer_Invoice_Head_CANCEL_DATA.Customer_Name as [Customer Name] " &
                  ", TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Type as [Document Type] " &
                  ", TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Total as [Document Amount] " &
                  ",TSPL_Customer_Invoice_Head_CANCEL_DATA.Created_By as [Created By] " &
                  ", convert(varchar,TSPL_Customer_Invoice_Head_CANCEL_DATA.Created_Date,103) as [Created Date], Description,TSPL_Customer_Invoice_Head_CANCEL_DATA.Cancel_By as [Cancelled By],convert(varchar,TSPL_Customer_Invoice_Head_CANCEL_DATA.Cancel_On,103) as [Cancelled Date] from TSPL_Customer_Invoice_Head_CANCEL_DATA " &
                  " WHERE  AgainstServiceInvoice ='Y' and "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_Customer_Invoice_Head_CANCEL_DATA.cancel_on ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_Customer_Invoice_Head_CANCEL_DATA.cancel_on,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                Else
                    qry += "  convert(date,TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_Customer_Invoice_Head_CANCEL_DATA.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += " ORDER BY TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_Date,TSPL_Customer_Invoice_Head_CANCEL_DATA.Document_No"

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
            End If



        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.FrmVendorService) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "Select TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Document_No as [Document Id] " &
                 ",convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Vendor_Invoice_Date ,103) as [Document Date], " &
                 "TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Loc_Code as [Location Code], TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Created_By as [Created By] " &
                 ", convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Created_Date,103) as [Created Date] ,Description,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Delete_By as [Delete By],convert(varchar,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Delete_On,103) as [Delete Date]  " &
                 " From TSPL_VENDOR_INVOICE_HEAD_Delete_DATA 
                 Left Outer Join TSPL_LOCATION_MASTER  On TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Loc_Code =TSPL_LOCATION_MASTER.Location_Code " &
                 " Where  "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Delete_On ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Delete_On,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                Else
                    qry += "convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Vendor_Invoice_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Vendor_Invoice_Date,103) <= convert(date,'" + dtpToDate.Value + "',103)  "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Loc_Code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += "ORDER BY TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Vendor_Invoice_Date,TSPL_VENDOR_INVOICE_HEAD_Delete_DATA.Document_No"

            Else
                gv1.DataSource = Nothing
                qry = "Select TSPL_VENDOR_INVOICE_HEAD_cancel_data.Document_No as [Document Id] " &
                 ",convert(varchar,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Vendor_Invoice_Date ,103) as [Document Date], " &
                 "TSPL_VENDOR_INVOICE_HEAD_cancel_data.Loc_Code as [Location Code], TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_VENDOR_INVOICE_HEAD_cancel_data.Created_By as [Created By] " &
                 ", convert(varchar,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Created_Date,103) as [Created Date] ,Description,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Cancel_By as [Cancelled By],convert(varchar,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Cancel_On,103) as [Cancelled Date]  " &
                 " From TSPL_VENDOR_INVOICE_HEAD_cancel_data   Left Outer Join TSPL_LOCATION_MASTER  On TSPL_VENDOR_INVOICE_HEAD_cancel_data.Loc_Code =TSPL_LOCATION_MASTER.Location_Code " &
                 " Where  "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_VENDOR_INVOICE_HEAD_cancel_data.cancel_on ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_cancel_data.cancel_on,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                Else
                    qry += "  convert(date,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Vendor_Invoice_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103)  and convert(date,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Vendor_Invoice_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_VENDOR_INVOICE_HEAD_cancel_data.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += "ORDER BY TSPL_VENDOR_INVOICE_HEAD_cancel_data.Due_Date,TSPL_VENDOR_INVOICE_HEAD_cancel_data.Document_No"

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
            End If




        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.ScrapSaleRetrun) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Document_No  as [Document Id],convert(varchar,TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Due_Date ,103) as [Document Date],  
                TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.shipment_No as [Against Shipment No],TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Loc_Code  as [Location Code], 
                TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Created_By as [Created By], 
                 convert(varchar,TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Created_Date,103) as [Created Date],'' as Description,TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Delete_By as [Delete By],convert(varchar,TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Delete_On,103) as [Delete Date] from TSPL_SCRAPSALE_HEAD_RETURN_Delete_data  
                 Left Outer Join TSPL_LOCATION_MASTER  on TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Loc_Code  =TSPL_LOCATION_MASTER.Location_Code
                WHERE "
                If rbtnCancelDate.IsChecked Then
                    qry += "  Convert(Date, TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Delete_On,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
                          " and convert(date,TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += " convert(date,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Due_Date ,103) >= convert(date,'" + dtpFromDate.Value + "',103) 
                and convert(date,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Due_Date,103) <= convert(date,'" + dtpToDate.Value + "',103) "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += " ORDER BY TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Due_Date,TSPL_SCRAPSALE_HEAD_RETURN_Delete_data.Document_No"
            Else
                gv1.DataSource = Nothing
                qry = "Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Document_No  as [Document Id],convert(varchar,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Due_Date ,103) as [Document Date],  
                TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.shipment_No as [Against Shipment No],TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Loc_Code  as [Location Code], 
                TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Created_By as [Created By], 
                 convert(varchar,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Created_Date,103) as [Created Date],'' as Description,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Cancel_By as [Cancelled By],convert(varchar,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Cancel_On,103) as [Cancelled Date] from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data  
                 Left Outer Join TSPL_LOCATION_MASTER  on TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Loc_Code  =TSPL_LOCATION_MASTER.Location_Code
                WHERE "
                If rbtnCancelDate.IsChecked Then
                    qry += "  Convert(Date, TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Cancel_On,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
                      " and convert(date,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "  Convert(Date, TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Due_Date,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
                      " and convert(date,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Due_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.loc_code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                qry += " ORDER BY TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Due_Date,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data.Document_No"

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
            End If


        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSNSaleInvoice) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleInvoicedairy) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "  Select TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_code  as [Document Id],convert(varchar,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_date ,103) as [Document Date],  " &
                " TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No as [Against Shipment No],TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Bill_to_Location  as [Location Code], " &
                " TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Created_By as [Created By], " &
                " convert(varchar,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Created_Date,103) as [Created Date],'' as Description,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.IsMultipleInvoice from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data  " &
                " Left Outer Join TSPL_LOCATION_MASTER  on TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Bill_to_Location  =TSPL_LOCATION_MASTER.Location_Code " &
                  " WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += "  Convert(Date, TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_On,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
      " and convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "  Convert(Date, TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_date,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
      " and convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                qry += " and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type IN ('FS','PS') and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Screen_Type='DS' "
                If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleInvoicedairy) = CompairStringResult.Equal Then
                    qry += " "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Bill_to_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If

                qry += "  ORDER BY TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_code "
            Else
                gv1.DataSource = Nothing
                qry = "  Select TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Document_code  as [Document Id],convert(varchar,TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Document_date ,103) as [Document Date],  " &
                " TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Against_Shipment_No as [Against Shipment No],TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Bill_to_Location  as [Location Code], " &
                " TSPL_LOCATION_MASTER.Location_Desc as [Location Name],TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By as [Created By], " &
                " convert(varchar,TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_Date,103) as [Created Date],'' as Description,TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.IsMultipleInvoice from TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA  " &
                " Left Outer Join TSPL_LOCATION_MASTER  on TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Bill_to_Location  =TSPL_LOCATION_MASTER.Location_Code " &
                  " WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += "  Convert(Date, TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Cancel_On,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
      " and convert(date,TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "  Convert(Date, TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Document_date,103) >= convert(Date,'" & dtpFromDate.Value & "',103) " &
      " and convert(date,TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Document_date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Trans_Type IN ('FS','PS') and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Screen_Type='DS' "
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Bill_to_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If

                qry += "  ORDER BY TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Document_date,TSPL_SD_SALE_INVOICE_HEAD_CANCEL_DATA.Document_code "

            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmDairyGatePass) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = "select GPCode as [GatePass No],convert(varchar,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.GPDate ,103) as [Gate Pass Date], Vehicle_Number as [Vehicle Number],Loading_Slip as [Loading Slip],TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Created_By as [Created By] 
,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Created_Date  as [Created Date] ,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Location_Code as [Location Code],Location_Desc  as [Location Name],Delete_By as [Deleted By] ,Delete_On as [Deleted Date] ,Remarks as Description,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Route_No,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.DistributorNamE as [Distributor Name]
 from TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data 
            left outer join  TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.ROUTE_NO = TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Route_No 
            WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += "convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Delete_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103)
                     and convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.GatePass_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) 
                   and convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.GatePass_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Location_Code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If

                qry += "order by TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.GPDate , TSPL_DAIRYSALE_GATEPASS_MASTER_Delete_Data.GPCode"
            Else
                gv1.DataSource = Nothing
                qry = "select GPCode as [GatePass No],convert(varchar,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.GPDate ,103) as [Gate Pass Date],TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Item_Type As [Item Type], Vehicle_Number as [Vehicle Number],Loading_Slip as [Loading Slip],TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Created_By as [Created By] 
,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Created_Date  as [Created Date] ,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Location_Code as [Location Code],Location_Desc  as [Location Name],Cancel_By as [Canceled By] ,Cancel_On as [Canceled Date] ,Remarks as Description,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Route_No,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.DistributorNamE as [Distributor Name]
 from TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data 
            left outer join  TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.ROUTE_NO = TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Route_No 
            WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += "convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Cancel_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103)
                     and convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.GatePass_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) 
                   and convert(date,TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.GatePass_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Location_Code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If

                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If

                qry += "order by TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.GPDate , TSPL_DAIRYSALE_GATEPASS_MASTER_Cancel_Data.GPCode"

            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmDairyBookingCustomer) = CompairStringResult.Equal Then
        If RdbDelete.IsChecked Then
            gv1.DataSource = Nothing
            qry = "  select TSPL_BOOKING_MATSER_Delete_Data.Document_No  as [Document Id] , convert(varchar,TSPL_BOOKING_MATSER_Delete_Data.Document_date ,103) as [Document Date],TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code As [Invoice No],Convert(Varchar(10),TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) As [Invoice Date],
TSPL_BOOKING_MATSER_Delete_Data.location_code as [Location Code],TSPL_BOOKING_MATSER_Delete_Data.sub_Location_Code As [Sub Location Code], 
TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Customer_Code As [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Customer Name],Convert(varchar,TSPL_BOOKING_MATSER_Delete_Data.Supply_Date,103) As [Supply Date], Case When TSPL_BOOKING_MATSER_Delete_Data.Is_Taxable=0 Then 'Non-Taxable' Else 'Taxable' End As [Invoice Type], TSPL_BOOKING_MATSER_Delete_Data.Created_By as [Created By], (convert(varchar,TSPL_BOOKING_MATSER_Delete_Data.Created_Date,103)+' '+convert(varchar,TSPL_BOOKING_MATSER_Delete_Data.Created_Date,108)) as [Created Date],TSPL_BOOKING_MATSER_Delete_Data.Delete_By AS [Deleted By],(CONVERT(varchar,TSPL_BOOKING_MATSER_Delete_Data.Delete_On,103)+' '+CONVERT(varchar,TSPL_BOOKING_MATSER_Delete_Data.Delete_On,108)) AS [Deleted Date]
,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Description from TSPL_BOOKING_MATSER_Delete_Data  
Left Outer Join TSPL_SD_SHIPMENT_HEAD_Cancel_Data On TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No=TSPL_BOOKING_MATSER_Delete_Data.Document_No
Left Outer Join TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data On TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code
Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Customer_Code WHERE "
            If rbtnCancelDate.IsChecked Then
                qry += " convert(date,TSPL_BOOKING_MATSER_Delete_Data.Delete_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                    convert(date,TSPL_BOOKING_MATSER_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
            Else
                qry += " convert(date,TSPL_BOOKING_MATSER_Delete_Data.Document_date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                  convert(date,TSPL_BOOKING_MATSER_Delete_Data.Document_date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
            End If
            If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                qry += " and TSPL_BOOKING_MATSER_Delete_Data.Location_Code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
            End If
            'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
            '    qry += " and TSPL_BOOKING_MATSER_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
            'Else
            '    qry += " and TSPL_BOOKING_MATSER_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
            'End If
        Else
            gv1.DataSource = Nothing
            qry = " select TSPL_BOOKING_MATSER_Cancel_Data.Document_No as [Document Id] , convert(varchar,TSPL_BOOKING_MATSER_Cancel_Data.Document_date ,103) as [Document Date],TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code As [Invoice No],Convert(Varchar(10),TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) As [Invoice Date],
               TSPL_BOOKING_MATSER_Cancel_Data.location_code as [Location Code],TSPL_BOOKING_MATSER_Cancel_Data.sub_Location_Code As [Sub Location Code], 
               TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Customer_Code As [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Customer Name],
               Convert(varchar,TSPL_BOOKING_MATSER_Cancel_Data.Supply_Date,103) As [Supply Date], Case When TSPL_BOOKING_MATSER_Cancel_Data.Is_Taxable=0 Then 'Non-Taxable' Else 'Taxable' End As [Invoice Type],TSPL_BOOKING_MATSER_Cancel_Data.Created_By as [Created By], (convert(varchar,TSPL_BOOKING_MATSER_Cancel_Data.Created_Date,103)+' '+convert(varchar,TSPL_BOOKING_MATSER_Cancel_Data.Created_Date,108)) as [Created Date],TSPL_BOOKING_MATSER_Cancel_Data.Cancel_By AS [Canceled By],(CONVERT(varchar,TSPL_BOOKING_MATSER_Cancel_Data.Cancel_On,103)+' '+CONVERT(varchar,TSPL_BOOKING_MATSER_Cancel_Data.Cancel_On,108)) AS [Canceled Date]
               ,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Description from TSPL_BOOKING_MATSER_Cancel_Data 
               Left Outer Join TSPL_SD_SHIPMENT_HEAD_Cancel_Data On TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No=TSPL_BOOKING_MATSER_Cancel_Data.Document_No
               Left Outer Join TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data On TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No=TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code
               Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Customer_Code WHERE "
            If rbtnCancelDate.IsChecked Then
                qry += "convert(date,TSPL_BOOKING_MATSER_Cancel_Data.Cancel_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                    convert(date,TSPL_BOOKING_MATSER_Cancel_Data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
            Else
                qry += "convert(date,TSPL_BOOKING_MATSER_Cancel_Data.Document_date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                  convert(date,TSPL_BOOKING_MATSER_Cancel_Data.Document_date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
            End If
            If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                qry += " and TSPL_BOOKING_MATSER_Cancel_Data.Location_Code  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
            End If
            'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
            '    qry += " and TSPL_BOOKING_MATSER_Cancel_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
            'Else
            '    qry += " and TSPL_BOOKING_MATSER_Cancel_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
            'End If

        End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSNPOS) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = " select TSPL_SD_POS_HEAD_Delete_Data.Document_Code  as [Document Id] , convert(varchar,TSPL_SD_POS_HEAD_Delete_Data.Document_Date ,103) as [Document Date],
TSPL_SD_POS_HEAD_Delete_Data.Bill_To_Location as [Location Code], Convert(varchar,TSPL_SD_POS_HEAD_Delete_Data.Document_Date,103) As [Supply Date], 
TSPL_SD_POS_HEAD_Delete_Data.Created_By as [Created By], (convert(varchar,TSPL_SD_POS_HEAD_Delete_Data.Created_Date,103)+' '+convert(varchar,TSPL_SD_POS_HEAD_Delete_Data.Created_Date,108)) as [Created Date],TSPL_SD_POS_HEAD_Delete_Data.Delete_By AS [Deleted By],(CONVERT(varchar,TSPL_SD_POS_HEAD_Delete_Data.Delete_On,103)+' '+CONVERT(varchar,TSPL_SD_POS_HEAD_Delete_Data.Delete_On,108)) AS [Deleted Date]
,'' as Description from TSPL_SD_POS_HEAD_Delete_Data   
            WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += "  convert(date,TSPL_SD_POS_HEAD_Delete_Data.Delete_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                    convert(date,TSPL_SD_POS_HEAD_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "  convert(date,TSPL_SD_POS_HEAD_Delete_Data.Document_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                    convert(date,TSPL_SD_POS_HEAD_Delete_Data.Document_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SD_POS_HEAD_Delete_Data.Bill_To_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_POS_HEAD_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_POS_HEAD_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
            Else
                gv1.DataSource = Nothing
                qry = " select TSPL_SD_POS_HEAD_cancel_data.Document_Code  as [Document Id] , convert(varchar,TSPL_SD_POS_HEAD_cancel_data.Document_Date ,103) as [Document Date],
TSPL_SD_POS_HEAD_cancel_data.Bill_To_Location as [Location Code], Convert(varchar,TSPL_SD_POS_HEAD_cancel_data.Document_Date,103) As [Supply Date], 
TSPL_SD_POS_HEAD_cancel_data.Created_By as [Created By], (convert(varchar,TSPL_SD_POS_HEAD_cancel_data.Created_Date,103)+' '+convert(varchar,TSPL_SD_POS_HEAD_cancel_data.Created_Date,108)) as [Created Date],TSPL_SD_POS_HEAD_cancel_data.Cancel_By AS [Canceled By],(CONVERT(varchar,TSPL_SD_POS_HEAD_cancel_data.Cancel_On,103)+' '+CONVERT(varchar,TSPL_SD_POS_HEAD_cancel_data.Cancel_On,108)) AS [Canceled Date]
,'' as Description from TSPL_SD_POS_HEAD_cancel_data   
            WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += "  convert(date,TSPL_SD_POS_HEAD_cancel_data.Cancel_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                    convert(date,TSPL_SD_POS_HEAD_cancel_data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += "  convert(date,TSPL_SD_POS_HEAD_cancel_data.Document_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                    convert(date,TSPL_SD_POS_HEAD_cancel_data.Document_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If
                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_SD_POS_HEAD_cancel_data.Bill_To_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_SD_POS_HEAD_cancel_data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_SD_POS_HEAD_cancel_data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If

            End If

        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmGatePassDairy) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                gv1.DataSource = Nothing
                qry = " select TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Document_No  as [Document Id] , convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Document_Date ,103) as [Document Date],
TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Location_Code as [Location Code], Convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Document_Date,103) As [Supply Date], 
TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Created_By as [Created By], (convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Created_Date,103)+' '+convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Created_Date,108)) as [Created Date],TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Delete_By AS [Deleted By],(CONVERT(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Delete_On,103)+' '+CONVERT(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Delete_On,108)) AS [Deleted Date]
,'' as Description from TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data   
            WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Delete_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += " convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Document_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Document_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Bill_To_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(cbgUser.CheckedValue) + ")"
                'Else
                '    qry += " and TSPL_GATEPASS_MASTER_DAIRYSALE_Delete_Data.Created_By IN (" + clsCommon.GetMulcallString(arrSelectedUser) + ")"
                'End If
            Else
                gv1.DataSource = Nothing
                qry = " select TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Document_No  as [Document Id] , convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Document_Date ,103) as [Document Date],
TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Location_Code as [Location Code], Convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Document_Date,103) As [Supply Date], 
TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Created_By as [Created By], (convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Created_Date,103)+' '+convert(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Created_Date,108)) as [Created Date],TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Cancel_By AS [Canceled By],(CONVERT(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Cancel_On,103)+' '+CONVERT(varchar,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Cancel_On,108)) AS [Canceled Date]
,'' as Description from TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data   
            WHERE  "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Cancel_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += " convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Document_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Document_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If

                If chkLocSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count > 0 Then
                    qry += " and TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Bill_To_Location  in   (" & clsCommon.GetMulcallString(cbgLocation.CheckedValue) & ") "
                End If
                'If chkUserSelect.IsChecked AndAlso cbgUser.CheckedValue.Count > 0 Then
                '    qry += " and TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Created_By IN (" & clsCommon.GetMulcallString(cbgUser.CheckedValue) & ")"
                'Else
                '    qry += " and TSPL_GATEPASS_MASTER_DAIRYSALE_cancel_data.Created_By IN (" & clsCommon.GetMulcallString(arrSelectedUser) & ")"
                'End If

            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.ScrapSale) = CompairStringResult.Equal Then
            If rdbCancel.IsChecked Then
                qry = "Select TSPL_SCRAPSALE_HEAD_Cancel_Data.shipment_No As [Document ID],Convert(Varchar(10),TSPL_SCRAPSALE_HEAD_Cancel_Data.shipment_Date,103) As [Document Date],TSPL_SCRAPINVOICE_HEAD_Cancel_Data.invoice_No As [Invoice No],
TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Ack_No As [Ack No],Convert(Varchar(10),TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Ack_Date,103) As [Ack Date],
TSPL_SCRAPSALE_HEAD_Cancel_Data.cust_Code As [Customer Code],TSPL_SCRAPSALE_HEAD_Cancel_Data.cust_Name As [Customer Name],Case When TSPL_SCRAPSALE_HEAD_Cancel_Data.Invoice_Type='T' Then 'Taxable' Else 'Non-Taxable' End As [Invoice Type],
TSPL_SCRAPSALE_HEAD_Cancel_Data.Loc_Code As [Location Code],TSPL_SCRAPSALE_HEAD_Cancel_Data.Loc_Name As [Location Name],
TSPL_SCRAPSALE_HEAD_Cancel_Data.Doc_Amt As [Amount],TSPL_SCRAPSALE_HEAD_Cancel_Data.Description,
TSPL_SCRAPSALE_HEAD_Cancel_Data.Created_By As [Created By],Convert(Varchar(10),TSPL_SCRAPSALE_HEAD_Cancel_Data.Created_Date,103) As [Created Date],
TSPL_SCRAPSALE_HEAD_Cancel_Data.Cancel_By As [Canceled By],Convert(Varchar(10),TSPL_SCRAPSALE_HEAD_Cancel_Data.Cancel_On,103) As [Canceled Date] 
from TSPL_SCRAPSALE_HEAD_Cancel_Data
left outer join TSPL_SCRAPINVOICE_HEAD_Cancel_Data on TSPL_SCRAPINVOICE_HEAD_Cancel_Data.shipment_No=TSPL_SCRAPSALE_HEAD_Cancel_Data.shipment_No
Inner Join TSPL_USER_MASTER On TSPL_USER_MASTER.User_Code=TSPL_SCRAPSALE_HEAD_Cancel_Data.Cancel_By Where "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Cancel_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Cancel_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += " convert(date,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.shipment_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.shipment_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If
            Else
                qry = "Select TSPL_SCRAPSALE_HEAD_Delete_Data.shipment_No As [Document ID],Convert(Varchar(10),TSPL_SCRAPSALE_HEAD_Delete_Data.shipment_Date,103) As [Document Date],TSPL_SCRAPINVOICE_HEAD_Delete_Data.invoice_No As [Invoice No],
TSPL_SCRAPINVOICE_HEAD_Delete_Data.Ack_No As [Ack No],Convert(Varchar(10),TSPL_SCRAPINVOICE_HEAD_Delete_Data.Ack_Date,103) As [Ack Date],
TSPL_SCRAPSALE_HEAD_Delete_Data.cust_Code As [Customer Code],TSPL_SCRAPSALE_HEAD_Delete_Data.cust_Name As [Customer Name],Case When TSPL_SCRAPSALE_HEAD_Delete_Data.Invoice_Type='T' Then 'Taxable' Else 'Non-Taxable' End As [Invoice Type],
TSPL_SCRAPSALE_HEAD_Delete_Data.Loc_Code As [Location Code],TSPL_SCRAPSALE_HEAD_Delete_Data.Loc_Name As [Location Name],
TSPL_SCRAPSALE_HEAD_Delete_Data.Doc_Amt As [Amount],TSPL_SCRAPSALE_HEAD_Delete_Data.Description,
TSPL_SCRAPSALE_HEAD_Delete_Data.Created_By As [Created By],Convert(Varchar(10),TSPL_SCRAPSALE_HEAD_Delete_Data.Created_Date,103) As [Created Date],
TSPL_SCRAPSALE_HEAD_Delete_Data.Delete_By As [Deleted By],Convert(Varchar(10),TSPL_SCRAPSALE_HEAD_Delete_Data.Delete_On,103) As [Deleted Date] 
from TSPL_SCRAPSALE_HEAD_Delete_Data
left outer join TSPL_SCRAPINVOICE_HEAD_Delete_Data on TSPL_SCRAPINVOICE_HEAD_Delete_Data.shipment_No=TSPL_SCRAPSALE_HEAD_Delete_Data.shipment_No
Inner Join TSPL_USER_MASTER On TSPL_USER_MASTER.User_Code=TSPL_SCRAPSALE_HEAD_Delete_Data.Delete_By Where "
                If rbtnCancelDate.IsChecked Then
                    qry += " convert(date,TSPL_SCRAPSALE_HEAD_Delete_Data.Delete_On ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_SCRAPSALE_HEAD_Delete_Data.Delete_On,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                Else
                    qry += " convert(date,TSPL_SCRAPSALE_HEAD_Delete_Data.shipment_Date ,103) >= convert(date,'" & dtpFromDate.Value & "',103) and 
                     convert(date,TSPL_SCRAPSALE_HEAD_Delete_Data.shipment_Date,103) <= convert(date,'" & dtpToDate.Value & "',103) "
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.FrmVendorService) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                qry = "Select Document_No As [Document ID],Convert(Varchar(10),Invoice_Entry_Date,103) As [Document Date],Vendor_Code As [Vendor Code],Vendor_Name As [Vendor Name],Document_Total As [Total],Description,Delete_By As [Deleted By],Convert(varchar(10),Delete_On,103) As [Deleted Date] from TSPL_VENDOR_INVOICE_HEAD_Delete_Data Where "
                If rbtnCancelDate.IsChecked Then
                    qry += " Convert(Date,Delete_On,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Delete_On,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                Else
                    qry += " Convert(Date,Invoice_Entry_Date,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Invoice_Entry_Date,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                End If
            Else
                qry = "Select Document_No As [Document ID],Invoice_Entry_Date As [Document Date],Vendor_Code As [Vendor Code],Vendor_Name As [Vendor Name],Document_Total As [Total],Description,Cancel_By As [Canceled By],Convert(varchar(10),Cancel_On,103) As [Canceled Date] from TSPL_VENDOR_INVOICE_HEAD_Cancel_Data Where "
                If rbtnCancelDate.IsChecked Then
                    qry += " Convert(Date,Cancel_On,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Cancel_On,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                Else
                    qry += " Convert(Date,Invoice_Entry_Date,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Invoice_Entry_Date,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleReturndairy) = CompairStringResult.Equal Then
            If RdbDelete.IsChecked Then
                qry = "Select Document_Code As [Document ID],Convert(Varchar(10),Document_Date,103) As [Document Date],Against_Invoice_No As [Invoice No],Customer_Code As [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Customer Name],Total_Amt As [Total],Description,Delete_By As [Deleted By],Convert(varchar(10),Delete_On,103) As [Deleted Date] from TSPL_SD_SALE_RETURN_HEAD_Delete_Data Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Customer_Code where "
                If rbtnCancelDate.IsChecked Then
                    qry += " Convert(Date,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Delete_On,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Delete_On,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                Else
                    qry += " Convert(Date,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Document_Date,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Document_Date,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                End If
            Else
                qry = "Select Document_Code As [Document ID],Convert(Varchar(10),Document_Date,103) As [Document Date],Against_Invoice_No As [Invoice No],Customer_Code As [Customer Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Customer Name],Total_Amt As [Total],Description,Cancel_By As [Canceled By],Convert(varchar(10),Cancel_On,103) As [Canceled Date] from TSPL_SD_SALE_RETURN_HEAD_cancel_Data Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_RETURN_HEAD_cancel_Data.Customer_Code where "
                If rbtnCancelDate.IsChecked Then
                    qry += " Convert(Date,Cancel_On,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Cancel_On,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                Else
                    qry += " Convert(Date,Document_Date,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Document_Date,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                End If
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.transfer) = CompairStringResult.Equal Then
            If rdbCancel.IsChecked Then
                qry = " select Document_No As [Document ID],Convert(varchar(10),Document_Date,103) As [Document Date],From_Location As [From Location],To_Location As [To Location],DOC_Total_Amt As [Total Amount],'' As [Description],Created_By As [Created By],Convert(varchar(10),Created_Date,103) As [Created Date],IRN_No As [IRN No],Ack_No As [Ack No],Convert(varchar(10),Ack_Date,103) As [Ack Date],Cancel_By As [Canceled By],
Convert(varchar(10),Cancel_On,103) As [Canceled Date]  from TSPL_TRANSFER_ORDER_HEAD_Cancel_Data where 1=1 "
                If rbtnCancelDate.IsChecked Then
                    qry += " And Convert(Date,Cancel_On,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Cancel_On,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                Else
                    qry += " And Convert(Date,Document_Date,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Document_Date,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                End If
            Else
                qry = " select Document_No As [Document ID],Convert(varchar(10),Document_Date,103) As [Document Date],From_Location As [From Location],To_Location As [To Location],DOC_Total_Amt As [Total Amount],'' As [Description],Created_By As [Created By],Convert(varchar(10),Created_Date,103) As [Created Date],IRN_No As [IRN No],Ack_No As [Ack No],Convert(varchar(10),Ack_Date,103) As [Ack Date],Delete_By As [Deleted By],
Convert(varchar(10),Delete_On,103) As [Deleted Date] from TSPL_TRANSFER_ORDER_HEAD_Delete_Data "
                If rbtnCancelDate.IsChecked Then
                    qry += " Convert(Date,Delete_On,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Delete_On,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                Else
                    qry += " Convert(Date,Document_Date,103)>=Convert(Date,'" & dtpFromDate.Value & "',103) And Convert(Date,Document_Date,103)<=Convert(Date,'" & dtpToDate.Value & "',103)"
                End If

            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(qry), Nothing) <> CompairStringResult.Equal Then

            dt = clsDBFuncationality.GetDataTable(qry)
            LoadBlankGrid()
            gv1.DataSource = Nothing
            gv1.DataSource = dt
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1Format()

            If dt.Rows.Count <= 0 Then
                lblNoOfRecords.Text = "No Record Found"
            Else
                strNoOfRecord = clsCommon.myCstr(dt.Rows.Count)
                lblNoOfRecords.Text = "" & strNoOfRecord & " Records Found"
            End If
        End If
        '-------------------------------------Code Ends Here--------------------------------------------
    End Sub

#End Region


    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        closeForm()
    End Sub


    Sub closeForm()
        Me.Close()
    End Sub


    Private Sub FrmPendingAproval_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.C Then
            closeForm()
        End If
    End Sub

    Private Sub chkLOcAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs)
        cbgLocation.Enabled = False
    End Sub

    Private Sub chkLOcSelect_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs)
        cbgLocation.Enabled = True
    End Sub
    Public Sub LoadLocation()
        Try
            Dim qry As String = Nothing
            If clsCommon.myLen(arrLoc) > 0 Then
                qry = " Select Location_Code as Code, Location_Desc as Description from TSPL_LOCATION_MASTER Where Location_Type='Physical' and Location_Code in (" & arrLoc & ")"
            End If
            cbgLocation.DataSource = clsDBFuncationality.GetDataTable(qry)
            cbgLocation.ValueMember = "Code"
            cbgLocation.DisplayMember = "Description"
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub LoadUsers()
        Try
            Dim qry As String = clsUserMaster.GetSubbordinateUsersQry(objCommonVar.CurrentUserCode)
            cbgUser.DataSource = clsDBFuncationality.GetDataTable(qry)
            cbgUser.ValueMember = "User_Code"
            cbgUser.DisplayMember = "User_Name"
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub chkLocAll_ToggleStateChanged_1(sender As Object, args As StateChangedEventArgs) Handles chkLocAll.ToggleStateChanged
        cbgLocation.Enabled = Not chkLocAll.IsChecked
    End Sub

    Private Sub ChkUserAll_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles ChkUserAll.ToggleStateChanged
        cbgUser.Enabled = Not ChkUserAll.IsChecked
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        dtpFromDate.Value = clsCommon.GETSERVERDATE()
        dtpToDate.Value = clsCommon.GETSERVERDATE()
        chkLocAll.CheckState = CheckState.Checked
        ChkUserAll.CheckState = CheckState.Checked
        gv1.DataSource = Nothing
        rbtnDocumentDate.IsChecked = True
    End Sub

    Private Sub btnPrintCancel_Click(sender As Object, e As EventArgs) Handles btnPrintCancel.Click
        'If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmMCCMaterial) = CompairStringResult.Equal Then
        '    clsMCCMaterialSale.funPrint(True, clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
        'Else
        '    printCanceInvoice()
        'End If
    End Sub

    Sub printCanceInvoice()
        Try
            If clsCommon.CompairString(cboTransaction.SelectedValue, clsUserMgtCode.frmSaleDispatchDairy) = CompairStringResult.Equal Then
                Dim Doc_Code As String = Nothing
                Dim Doc_Date As DateTime = Nothing
                Dim Inv_Code As String = Nothing
                Dim Cust_Code As String = Nothing
                Dim objPrintInvoice As New frmShipmentDairy
                If gv1 IsNot Nothing AndAlso gv1.Rows.Count > 0 Then
                    Doc_Code = clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value)
                    Doc_Date = clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value)
                    Inv_Code = clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Invoice No").Value)
                    If rdbCancel.IsChecked Then
                        Cust_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Customer_Code from TSPL_SD_SHIPMENT_HEAD_Cancel_Data where Document_Code='" & Doc_Code & "'"))
                        objPrintInvoice.PrintInvoiveForAll(Doc_Code, Doc_Date, Inv_Code, "Cancel")
                    Else
                        Cust_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Customer_Code from TSPL_SD_SHIPMENT_HEAD_Delete_Data where Document_Code='" & Doc_Code & "'"))
                        objPrintInvoice.PrintInvoiveForAll(Doc_Code, Doc_Date, Inv_Code, "Delete")
                    End If
                Else
                    Throw New Exception("Data not found !")
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub MSGShowAndHide()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleDispatchDairy) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmMCCMaterial) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmDairyBookingCustomer) = CompairStringResult.Equal Then
                btnPrintCancel.Visible = False
                lblPrintMsg.Visible = True
            Else
                btnPrintCancel.Visible = False
                lblPrintMsg.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub cboTransaction_SelectedValueChanged(sender As Object, e As EventArgs) Handles cboTransaction.SelectedValueChanged
        MSGShowAndHide()
    End Sub

    Private Sub gv1_CellDoubleClick(sender As Object, e As GridViewCellEventArgs) Handles gv1.CellDoubleClick
        Try
            If gv1 IsNot Nothing AndAlso gv1.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmMCCMaterial) = CompairStringResult.Equal Then
                    clsMCCMaterialSale.funPrint(MyBase.Form_ID, True, clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), False)
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmDairyBookingCustomer) = CompairStringResult.Equal Then
                    'PrintInvoiceForAll("", MyBase.Form_ID, True, clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    Dim frm As New frmDairyBookingCustomer()

                    Dim InvoiceDocNo As String
                    Dim doccodeShip As String
                    'Dim InvoiceDocNo As New List(Of String)
                    doccodeShip = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SHIPMENT_HEAD_Cancel_Data  where Against_Booking_No ='" & clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value) & "'"))

                    InvoiceDocNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data where Against_Shipment_No ='" & doccodeShip & "'"))
                    'Dim Qry As String = Nothing

                    frm.printInvoice(clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), clsCommon.myCstr(InvoiceDocNo), clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), True)
                    'frm.printInvoice(clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), True)
                    frm = Nothing
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSNPOS) = CompairStringResult.Equal Then
                    clsSNPOSHead.funSNFPOSPrint(MyBase.Form_ID, True, clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmGatePassDairy) = CompairStringResult.Equal Then
                    clsMilkTransferIn.funGatepassdairyPrint(MyBase.Form_ID, True, clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), False, False, Nothing)
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmDairyGatePass) = CompairStringResult.Equal Then
                    Dim frmFree As New XpertERPEngine.FrmFreeComboBox()
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable("Select Code from (Select 'Print 1' As Code Union All Select 'Print 2' As Code)xyz")
                    frmFree.ComboSource = dt
                    frmFree.ComboValueMember = "Code"
                    frmFree.ComboDisplayMember = "Code"
                    frmFree.ShowDialog()
                    Dim Value As String = frmFree.strRetValue
                    frmFree = Nothing
                    If clsCommon.myLen(Value) <= 0 Then
                        Throw New Exception("Select Print !")
                    End If
                    Dim isFresh As Boolean = False
                    Dim isAmbient As Boolean = False
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Item Type").Value), "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Item Type").Value), "I") = CompairStringResult.Equal Then
                        isAmbient = True
                    Else
                        isFresh = True
                    End If
                    Dim frm As New frmDairyGatePass()
                    If clsCommon.CompairString(Value, "Print 1") = CompairStringResult.Equal Then
                        frm.GatepassWithFilePath(clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("GatePass No").Value), clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Gate Pass Date").Value), Nothing, isFresh, isAmbient, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Route_No").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Location Code").Value), False, True)
                    Else
                        frm.funPrint2(clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("GatePass No").Value), False, True)
                    End If
                    frm = Nothing
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.ScrapSale) = CompairStringResult.Equal Then
                    Dim frm As New frmScrapSale()
                    'Print(ByVal isPrint As Boolean, ByVal ischallan As Boolean, ByVal isPDFPath As Boolean, ByVal strCancelDelete As String)
                    Dim strDoc As String = clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value)
                    Dim strInvNo As String = clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Invoice No").Value)
                    Dim strLocCode As String = clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Location Code").Value)
                    If rdbCancel.IsChecked Then
                        frm.Print(True, True, False, "Cancel", strDoc, strInvNo, strLocCode)
                    Else
                        frm.Print(True, True, False, "Delete", strDoc, strInvNo, strLocCode)
                    End If
                    frm = Nothing
                    'ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.FrmVendorService) = CompairStringResult.Equal Then
                    '    If rdbCancel.IsChecked Then
                    '        clsVedorInvoiceHead.funVendorServicePrint(MyBase.Form_ID, "Cancel", clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    '    Else
                    '        clsVedorInvoiceHead.funVendorServicePrint(MyBase.Form_ID, "Delete", clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    '    End If
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleReturndairy) = CompairStringResult.Equal Then
                    If rdbCancel.IsChecked Then
                        clsDSSalesReturnHead.funsaleReturnDairyPrint(MyBase.Form_ID, "Cancel", clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), False)
                    Else
                        clsDSSalesReturnHead.funsaleReturnDairyPrint(MyBase.Form_ID, "Delete", clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), False)
                    End If
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSNSaleInvoice) = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.frmSaleInvoicedairy) = CompairStringResult.Equal Then
                    Dim Qry As String = Nothing
                    Dim dt As DataTable = Nothing
                    Dim frm As New frmMultipleInvoice()
                    If rdbCancel.IsChecked Then
                        frm.PrintMultipleInvoice(Me.Form_ID, Me.Text, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), "Cancel")
                    Else
                        frm.PrintMultipleInvoice(Me.Form_ID, Me.Text, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), "Delete")
                    End If
                    frm = Nothing
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.Transfer) = CompairStringResult.Equal Then
                    Dim frmFree As New XpertERPEngine.FrmFreeComboBox()
                    Dim strQry As String = "Select Code from ("
                    strQry &= " Select 'Print' As Code "
                    strQry &= " Union All "
                    strQry &= " Select 'Delivery Challan' As Code "
                    strQry &= " Union All "
                    strQry &= " Select 'STA Print' As Code "
                    strQry &= " Union All "
                    strQry &= " Select 'STA Milk Print' As Code "
                    strQry &= " Union All "
                    strQry &= " Select 'STA Product Print' As Code "
                    strQry &= " )xyz"
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
                    frmFree.ComboSource = dt
                    frmFree.ComboValueMember = "Code"
                    frmFree.ComboDisplayMember = "Code"
                    frmFree.ShowDialog()
                    Dim Value As String = frmFree.strRetValue
                    frmFree = Nothing
                    If clsCommon.myLen(Value) <= 0 Then
                        Throw New Exception("Select Print !")
                    End If
                    Dim isCancelOrDelete As String = Nothing
                    If rdbCancel.IsChecked Then
                        isCancelOrDelete = "Cancel"
                    Else
                        isCancelOrDelete = "Delete"
                    End If
                    Dim frm As New FrmTransferKDIL()
                    If clsCommon.CompairString(Value, "Print") = CompairStringResult.Equal Then
                        frm.PrintData(Me.Form_ID, Me.Text, isCancelOrDelete, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    ElseIf clsCommon.CompairString(Value, "Delivery Challan") = CompairStringResult.Equal Then
                        frm.PrintChallanReport(Me.Form_ID, Me.Text, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), isCancelOrDelete)
                    ElseIf clsCommon.CompairString(Value, "STA Print") = CompairStringResult.Equal Then
                        frm.STAPrint(Me.Form_ID, Me.Text, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), isCancelOrDelete)
                    ElseIf clsCommon.CompairString(Value, "STA Milk Print") = CompairStringResult.Equal Then
                        frm.STAMilkPrint(Me.Form_ID, Me.Text, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), isCancelOrDelete)
                    Else
                        frm.STAProductPrint(Me.Form_ID, Me.Text, clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), isCancelOrDelete)
                    End If
                    frm = Nothing
                    ' ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.mbtnAPInvoiceEntry) = CompairStringResult.Equal Then
                    '  clsVedorInvoiceHead.funAPInvoiceEntryPrint(MyBase.Form_ID, "Cancel", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), False)

                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.mbtnARInvoiceEntry) = CompairStringResult.Equal Then
                    ' clsReceiptInvoiceHead.funARInvoicePrint(MyBase.Form_ID, True, clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), Nothing, Nothing)
                    If rdbCancel.IsChecked Then
                        clsReceiptInvoiceHead.funPrintServiceInvoicePrint(MyBase.Form_ID, "Cancel", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), 0)
                    Else
                        clsReceiptInvoiceHead.funPrintServiceInvoicePrint(MyBase.Form_ID, "Delete", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), 0)
                    End If
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.mbtnAPInvoiceEntry) = CompairStringResult.Equal Then
                    If rdbCancel.IsChecked Then
                        clsVedorInvoiceHead.funVendorServicePrint(MyBase.Form_ID, "Cancel", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    Else
                        clsVedorInvoiceHead.funVendorServicePrint(MyBase.Form_ID, "Delete", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    End If
                    'clsVedorInvoiceHead.funAPInvoiceEntryPrint(MyBase.Form_ID, "Cancel", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), False)
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.FrmVendorService) = CompairStringResult.Equal Then
                    If rdbCancel.IsChecked Then
                        clsVedorInvoiceHead.funVendorServicePrint(MyBase.Form_ID, "Cancel", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))
                    Else
                        clsVedorInvoiceHead.funVendorServicePrint(MyBase.Form_ID, "Delete", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value))

                    End If
                ElseIf clsCommon.CompairString(clsCommon.myCstr(cboTransaction.SelectedValue), clsUserMgtCode.ScrapSaleRetrun) = CompairStringResult.Equal Then
                        Dim qry As String = ""
                        Dim dt As New DataTable
                        Dim locCode As String = ""

                        qry = " select Loc_Code from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_data where Document_No ='" + clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value) + "'"
                        dt = clsDBFuncationality.GetDataTable(qry)
                        If dt.Rows.Count > 0 Then
                            locCode = clsCommon.myCstr(dt.Rows(0)("Loc_Code"))
                        End If
                        If rdbCancel.IsChecked Then

                            ClsScrapSaleHeadReturn.frmScrapSaleReturnPrint(MyBase.Form_ID, "Cancel", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), True, (locCode))
                        Else
                            ClsScrapSaleHeadReturn.frmScrapSaleReturnPrint(MyBase.Form_ID, "Delete", clsCommon.myCDate(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document Date").Value), clsCommon.myCstr(gv1.Rows(gv1.CurrentCell.RowIndex).Cells("Document ID").Value), True, (locCode))

                        End If

                    Else
                        printCanceInvoice()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    'Private Sub cboTransaction_SelectedIndexChanged(sender As Object, e As UI.Data.PositionChangedEventArgs) Handles cboTransaction.SelectedIndexChanged

    'End Sub

    Private Sub RdbDelete_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles RdbDelete.ToggleStateChanged
        If RdbDelete.IsChecked Then
            rbtnCancelDate.Text = "Delete Date"
        Else
            rbtnCancelDate.Text = "Cancel Date"
        End If
    End Sub

    'Private Sub gv1_CellEditorInitialized(sender As Object, e As GridViewCellEventArgs) Handles gv1.CellEditorInitialized
    '    If TypeOf Me.gv1.CurrentColumn Is GridViewMultiComboBoxColumn Then
    '        Dim editor As RadMultiColumnComboBoxElement = DirectCast(Me.gv1.ActiveEditor, RadMultiColumnComboBoxElement)
    '        editor.AutoSizeDropDownTofrdfy = True
    '        editor.EditorControl.MasterTemplate.BestFitColumns()
    '        editor.DropDownStyle = RadDropDownStyle.DropDown
    '        editor.AutoFilter = True
    '        'If editor.EditorControl.MasterTemplate.FilterDescriptors.Count = 0 Then
    '        '    Dim autoFilter As FilterDescriptor = New FilterDescriptor("Description", FilterOperator.StartsWith, "")
    '        '    autoFilter.IsFilterEditor = True
    '        '    editor.EditorControl.FilterDescriptors.Add(autoFilter)
    '        'End If
    '    End If
    'End Sub
End Class




