Imports System.IO
Imports common
Public Class rptCrateRegister
    Inherits FrmMainTranScreen
    Private Sub rptCrateRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        LoadShiftFrom()
        LoadShiftTo()
    End Sub
    Sub LoadShiftFrom()
        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Code")
        dt.Columns.Add("Shift")

        Dim dr As DataRow = dt.NewRow
        dr("Code") = "M"
        dr("Shift") = "Morning"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "E"
        dr("Shift") = "Evening"
        dt.Rows.Add(dr)

        txtFromShift.DataSource = dt
        txtFromShift.ValueMember = "Code"
    End Sub

    Sub LoadShiftTo()
        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Code")
        dt.Columns.Add("Shift")

        Dim dr As DataRow = dt.NewRow
        dr("Code") = "M"
        dr("Shift") = "Morning"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "E"
        dr("Shift") = "Evening"
        dt.Rows.Add(dr)

        txtToShift.DataSource = dt
        txtToShift.ValueMember = "Code"

    End Sub
    Private Sub MultLocation__My_Click(sender As Object, e As EventArgs) Handles MultLocation._My_Click
        Try
            'Dim qry As String = "select Cust_Code as Code ,Customer_Name as  Name from TSPL_CUSTOMER_MASTER "

            'If txtRoute.arrValueMember IsNot Nothing Then
            '    qry += "where Route_No in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ") "
            'End If
            Dim qry As String = " select CUST_CODE,Customer_Name from TSPL_CUSTOMER_MASTER"
            MultLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("ItemMulSel", qry, "CUST_CODE", "CUST_CODE", MultLocation.arrValueMember, MultLocation.arrDispalyMember)
            ' MultLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("LedgerCustomer", qry, "Code", "Name", MultLocation.arrValueMember, MultLocation.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub MultArea__My_Click(sender As Object, e As EventArgs) Handles MultArea._My_Click
        Dim qry As String = " select Route_No,Route_Desc from tspl_route_master "
        MultArea.arrValueMember = clsCommon.ShowMultipleSelectForm("ItemMulSel", qry, "Route_No", "Route_No", MultArea.arrValueMember, MultArea.arrDispalyMember)
    End Sub

    Private Sub btnclose_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Griddata(False)
    End Sub
    Public Sub Griddata(ByVal print As Boolean)
        Try

            MultLocation.Enabled = False
            MultArea.Enabled = False
            txtFromDate.Enabled = False
            txtToDate.Enabled = False
            Dim FromDate As String = clsCommon.myCstr(txtFromDate.Text)
            Dim ToDate As String = clsCommon.myCstr(txtToDate.Text)
            Dim qry As String = Nothing
            Dim whrcls As String = ""
            Dim whrclsClosing As String = ""
            Dim whrclss As String = ""

            'whrclss = "    Document_Date >= '" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(FromDate), "dd/MMM/yyyy hh:mm:ss tt") & "' AND " & "Document_Date <= '" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(ToDate), "dd/MMM/yyyy hh:mm:ss tt") & "' "
            whrclss = " convert(date, Document_Date, 108) >= convert(date, '" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(FromDate), "dd/MMM/yyyy hh:mm:ss tt") & "', 108) " &
          "AND convert(date, Document_Date, 108) <= convert(date, '" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(ToDate), "dd/MMM/yyyy hh:mm:ss tt") & "', 108)"

            'whrcls = "where 2 = 2 and convert(date,Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103) "
            'whrclsClosing = "where 2 = 2 and convert(date,Document_Date,103)<convert(date,'" + txtFromDate.Value + "',103) "
            whrclsClosing = "  Document_Date <= '" & clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(FromDate), "dd/MMM/yyyy hh:mm:ss tt") & "'"


            Dim whrclsCust As String = Nothing
            Dim whrclsRoute As String = Nothing

            If MultLocation.arrValueMember IsNot Nothing AndAlso MultLocation.arrValueMember.Count > 0 Then
                whrclsCust += "  and TSPL_CRATE_RECEIVED_detail_FRESHSALE.Customer_Code in (" + clsCommon.GetMulcallString(MultLocation.arrValueMember) + ")  "

            End If

            If MultArea.arrValueMember IsNot Nothing AndAlso MultArea.arrValueMember.Count > 0 Then
                whrclsRoute += " and TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code in (" + clsCommon.GetMulcallString(MultArea.arrValueMember) + ")"
            End If
            qry = "  SELECT *
FROM(
select
'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + "' As FromDate, '" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy") + "'  As ToDate,
Route_No,max(Route_Desc)Route_Desc,Customer_Code, max(Customer_Name)Customer_Name,max(Comp_Name)Comp_Name,max(Location_Code)Location_Code,max(Location_Desc)Location_Desc, max(Vehicle_Id)Vehicle_Id,max(Vehicle_Number)Vehicle_Number "

            If clsCommon.CompairString(txtFromShift.SelectedValue, "M") = CompairStringResult.Equal Then
                qry += " ,sum(Qty * RI * CASE when  Document_Date<'" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'   then 1 else 0 end ) as  OP "
            ElseIf clsCommon.CompairString(txtFromShift.SelectedValue, "E") = CompairStringResult.Equal Then
                qry += " ,sum(Qty * RI * CASE when  Document_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and ShiftType='M' then 1 else 0 end ) as  OP "
            End If
            Dim whrclsShift As String = ""
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                whrclsShift += " and 2=( case when Cast(Document_Date as Date) >= '" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' and Cast(Document_Date as Date) <= '" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' and  shifttype='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                whrclsShift += " and 2=( case when Cast(Document_Date as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy") + "' and Cast(Document_Date as Date) <= '" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "'  and shifttype='E' then 3 else 2 end  )"
            End If

            qry += " ,sum(Qty * case when RI=1 then 1 else 0 end * CASE when  Document_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + whrclsShift + "  then 1 else 0 end ) as  Supply
,sum(Qty * case when RI=-1 then 1 else 0 end * CASE when  Document_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + whrclsShift + "  then 1 else 0 end ) as  [Return] "

            If clsCommon.CompairString(txtToShift.SelectedValue, "M") = CompairStringResult.Equal Then
                qry += " ,sum(Qty * RI * CASE when  Document_Date<'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' then 1 else 0 end ) as  CL "
            ElseIf clsCommon.CompairString(txtToShift.SelectedValue, "E") = CompairStringResult.Equal Then
                qry += " ,sum(Qty * RI * CASE when  Document_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and ShiftType='M' then 1 else 0 end ) as  CL "
            End If

            qry += " from (
select TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.type, TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_Date, TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No, TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Vehicle_Code AS Vehicle_Id,TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.VehicleNo AS Vehicle_Number, TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Route_code as Route_No,  TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.ShiftType,TSPL_route_master.Route_Desc,TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,CAST(TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Invoice_Date AS DATE) AS Sale_Invoice_Date,TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.CrateQtyRecd as Qty ,(case when TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.type='O' THEN 1 ELSE -1 END )as RI,TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Location_Code,TSPL_LOCATION_MASTER.Location_Desc,TSPL_COMPANY_MASTER.Comp_Name
From TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE
left outer join TSPL_CRATE_RECEIVED_HEAD_FRESHSALE on TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Document_No=TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Document_No
left outer join tspl_route_master on tspl_route_master.route_No=TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.route_Code 
left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Location_Code
LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=TSPL_CRATE_RECEIVED_HEAD_FRESHSALE.Comp_Code 
left outer join tspl_customer_master on tspl_customer_master.Cust_Code = TSPL_CRATE_RECEIVED_DETAIL_FRESHSALE.Customer_Code
where 2=2  " + whrclsRoute + " " + whrclsCust + "
)xx group by Customer_Code,Route_No) final
WHERE NOT (OP = 0 AND Supply = 0 AND [Return] = 0 AND CL = 0)
                               "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing OrElse dt.Rows.Count > 0 Then
                If print = False Then
                    gv2.DataSource = Nothing
                    gv2.Rows.Clear()
                    gv2.Columns.Clear()
                    gv2.GroupDescriptors.Clear()
                    gv2.MasterTemplate.SummaryRowsBottom.Clear()
                    gv2.MasterView.Refresh()
                    gv2.DataSource = dt
                    For ii As Integer = 0 To gv2.Columns.Count - 1
                        gv2.Columns(ii).ReadOnly = True
                    Next
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv2.EnableFiltering = True
                    SetGridFormat1()
                    gv2.BestFitColumns()
                Else
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptCrateRegister", "")

                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If

            'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '    gv2.DataSource = Nothing
            '    gv2.Rows.Clear()
            '    gv2.Columns.Clear()
            '    gv2.GroupDescriptors.Clear()
            '    gv2.MasterTemplate.SummaryRowsBottom.Clear()
            '    gv2.MasterView.Refresh()
            '    gv2.DataSource = dt
            '    For ii As Integer = 0 To gv2.Columns.Count - 1
            '        gv2.Columns(ii).ReadOnly = True
            '    Next
            '    RadPageView1.SelectedPage = RadPageViewPage2
            '    gv2.EnableFiltering = True
            '    SetGridFormat1()
            '    gv2.BestFitColumns()
            'Else
            '    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
            '    Exit Sub
            'End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub SetGridFormat1()
        gv2.TableElement.TableHeaderHeight = 40
        gv2.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To gv2.Columns.Count - 1
            gv2.Columns(ii).ReadOnly = True
            gv2.Columns(ii).IsVisible = True
            gv2.Columns("FromDate").IsVisible = False
            gv2.Columns("ToDate").IsVisible = False
            'gv1.Columns("Document_No").HeaderText = "Document No."

            'gv2.Columns("Document_No").IsVisible = False
            'gv2.Columns("Document_No").HeaderText = "Document No"
            'gv2.Columns("Document_Date").IsVisible = False
            'gv2.Columns("Document_Date").HeaderText = "Document Date"
            'gv2.Columns("DistributorCode").IsVisible = True
            gv2.Columns("Comp_Name").HeaderText = "Comp_Name"
            gv2.Columns("Comp_Name").IsVisible = False
            gv2.Columns("Location_Code").HeaderText = "Location_Code"
            gv2.Columns("Location_Code").IsVisible = False
            gv2.Columns("Location_Desc").HeaderText = "Location_Desc"
            gv2.Columns("Location_Desc").IsVisible = False
            gv2.Columns("Vehicle_Id").HeaderText = "Vehicle_Id"
            gv2.Columns("Vehicle_Id").IsVisible = False
            gv2.Columns("Vehicle_Number").HeaderText = "Vehicle_Number"
            gv2.Columns("Vehicle_Number").IsVisible = False

            gv2.Columns("Customer_Code").HeaderText = "Distributor Code"
            gv2.Columns("Customer_Code").IsVisible = True
            gv2.Columns("Customer_Name").HeaderText = "Distributor Name"
            gv2.Columns("Route_No").IsVisible = True
            gv2.Columns("Route_No").HeaderText = "Area Code"
            gv2.Columns("Route_Desc").IsVisible = False
            gv2.Columns("Route_Desc").HeaderText = "Area Name"
            gv2.Columns("Supply").IsVisible = True
            gv2.Columns("Supply").HeaderText = "Supply"
            gv2.Columns("Return").IsVisible = True
            gv2.Columns("Return").HeaderText = "Return"


            'gv2.Columns("Comp_Code").IsVisible = False
            'gv2.Columns("Comp_Code").HeaderText = "Comp_Code"
            gv2.Columns("Comp_Name").IsVisible = False
            gv2.Columns("Comp_Name").HeaderText = "Comp_Name"
            'gv2.Columns("FromDate").IsVisible = False
            'gv2.Columns("FromDate").HeaderText = "FromDate"
            'gv2.Columns("ToDate").IsVisible = False
            'gv2.Columns("ToDate").HeaderText = "ToDate"
            'gv2.Columns("Sale_Invoice_Date").IsVisible = False
            'gv2.Columns("Sale_Invoice_Date").HeaderText = "Sale Invoice Date"


            'gv2.Columns("Morning_Supply").IsVisible = False
            'gv2.Columns("Morning_Supply").HeaderText = "Morning_Supply"
            'gv2.Columns("Morning_Return").IsVisible = False
            'gv2.Columns("Morning_Return").HeaderText = "Morning_Return"
            'gv2.Columns("Evening_Supply").IsVisible = False
            'gv2.Columns("Evening_Supply").HeaderText = "Evening_Supply"
            'gv2.Columns("Evening_Return").IsVisible = False
            'gv2.Columns("Evening_Return").HeaderText = "Evening_Return"
            gv2.Columns("OP").IsVisible = True
            gv2.Columns("OP").HeaderText = "Opening Balance"
            gv2.Columns("CL").IsVisible = True
            gv2.Columns("CL").HeaderText = "Closing Balance"
        Next
        Dim summaryRowItemB As New GridViewSummaryRowItem()

        'Dim MilkTypeB As New GridViewSummaryItem("Payable_Amount", "{0:n0}", GridAggregateFunction.Sum)

        Dim Supply As New GridViewSummaryItem("Supply", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(Supply)
        Dim Returnqty As New GridViewSummaryItem("Return", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(Returnqty)
        Dim CloseingBal As New GridViewSummaryItem("CL", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(CloseingBal)
        Dim OpeningBal As New GridViewSummaryItem("OP", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(OpeningBal)
        gv2.MasterTemplate.SummaryRowsBottom.Add(summaryRowItemB)
        gv2.AutoSizeRows = True
        gv2.BestFitColumns()
        gv2.MasterTemplate.AutoExpandGroups = True
    End Sub
    Private Sub btnExportToExcel_Click(sender As Object, e As EventArgs) Handles btnExportToExcel.Click
        Export(EnumExportTo.Excel)
    End Sub
    Private Sub Export(ByVal exporter As EnumExportTo)
        Try
            If gv2.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Print Date (" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd-MMM-yyyy hh:mm:ss tt") + ")")
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptCrateRegister & "'"))
                arrHeader.Add("Date Range : " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + "  To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy"))

                If MultLocation.arrValueMember IsNot Nothing AndAlso MultLocation.arrValueMember.Count > 0 Then
                    arrHeader.Add("Distrutor Code : " & MultLocation.arrDispalyMember(0))
                End If

                If MultArea.arrValueMember IsNot Nothing AndAlso MultArea.arrValueMember.Count > 0 Then
                    arrHeader.Add("Area : " & MultArea.arrDispalyMember(0))

                End If
                ' transportSql.applyExportTemplate(gv2, PageSetupReport_ID)
                clsCommon.MyExportToExcelGrid(Me.Text, gv2, arrHeader, Me.Text, True)

                'If exporter = EnumExportTo.Excel Then
                ' transportSql.applyExportTemplate(gv2, PageSetupReport_ID)
                ' transportSql.exportdata(gv2, "", Me.Text, False, arrHeader, False, False, True)
                'End If
                'transportSql.QuickExportToExcel(gv1, "", Me.Text, , arrHeader)
            Else
                Throw New Exception("No data found to export.")

            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Griddata(True)
    End Sub
    Private Sub Reset()
        'txtMultBmc.Enabled=True
        MultLocation.arrValueMember = Nothing
        MultArea.arrValueMember = Nothing
        txtFromDate.Enabled = True
        txtToDate.Enabled = True
        gv2.DataSource = Nothing
        gv2.Rows.Clear()
        gv2.Columns.Clear()
        RadPageView1.SelectedPage = RadPageViewPage1
        btnGo.Enabled = True
        MultLocation.Enabled = True
        MultArea.Enabled = True
    End Sub
    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles BtnReset.Click
        Reset()
    End Sub
End Class