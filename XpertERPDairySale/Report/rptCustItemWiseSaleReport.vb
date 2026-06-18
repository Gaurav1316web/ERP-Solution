
Imports System.IO
Imports common

Public Class rptCustItemWiseSaleReport
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim dtTax As DataTable = New DataTable()
    Dim isPrint As Boolean = False
    Dim ShowAllCustomerItemWiseSaleReportOptions As Boolean = False
    Dim Report_ID As String = MyBase.Form_ID
    Dim BoothSaleQry As String = Nothing
#End Region
    Private Sub rptCustItemWiseSaleReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ShowAllCustomerItemWiseSaleReportOptions = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ShowAllCustomerItemWiseSaleReportOptions, clsFixedParameterCode.ShowAllCustomerItemWiseSaleReportOptions, Nothing)) > 0)
        If ShowAllCustomerItemWiseSaleReportOptions Then
            BKNGroupBox.Visible = True
        Else
            BKNGroupBox.Visible = False
        End If
        funreset()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        txtToDate1.Value = clsCommon.GETSERVERDATE()
        txtFromDate1.Value = clsCommon.GETSERVERDATE()
        LoadQtyConversionType()
        LoadDefaultReportUOMType()
        LoadShiftFrom()
        LoadShiftTo()
        LoadType()
        Dim subLocation As String = clsDBFuncationality.getSingleValue(" select ISNULL(Sub_Location,'')Sub_Location,* from tspl_user_master where User_Code= '" & objCommonVar.CurrentUserCode & "'  ")
        Dim arrList As ArrayList = New ArrayList
        arrList.Add(subLocation)
        If clsCommon.myLen(subLocation) > 0 Then
            txtLocation.arrValueMember = arrList
        End If
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

    Private Sub txtCustomer__My_Click(sender As Object, e As EventArgs) Handles txtCustomer._My_Click
        Try
            Dim qry As String = "select Cust_Code as Code,Customer_Name as Name from tspl_customer_master  "

            txtCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("CustFilter", qry, "Code", "Name", txtCustomer.arrValueMember, txtCustomer.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtItem__My_Click(sender As Object, e As EventArgs) Handles txtItem._My_Click
        Try
            Dim qry As String = "select  Item_Code as [Item Code] ,Item_Desc as  [Item Desc] ,Short_Description as [Short Description] from TSPL_ITEM_MASTER where Item_Type = 'F'"
            txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("HSNItem", qry, "Item Code", "Item Desc", txtItem.arrValueMember, txtItem.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        funreset()
    End Sub
    Sub GetReportID()
        Dim VarID As String = ""
        If BtnStcRegisterPartyandItemWiseSummary.IsChecked Then
            VarID += "_STRS"
        ElseIf BtnStcRegisterItemWiseSummary.IsChecked Then
            VarID += "_SRIS"
        ElseIf BtnProductWiseSaleQuantity.IsChecked Then
            VarID += "_PWSQ"
        ElseIf BtnMilkStcSummary.IsChecked Then
            VarID += "_MSS"
        ElseIf BtnPartySaleMilkProduct.IsChecked Then
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                VarID += "_PSMPDS"
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                VarID += "_PSMPIN"
            End If
            VarID += "_PSMPDS"
        ElseIf BtnProductSalesSummary.IsChecked Then
            If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal Then
                VarID += "_PSST"
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                VarID += "_PSNT"
            ElseIf ddlType.SelectedValue = "Both" Then
                VarID += "_PSNTNT"
            End If
        ElseIf BtnBillWiseSaleOfMilk.IsChecked Then
            VarID += "_BWSM"
        ElseIf BtnBillWiseSaleOfMilkSummary.IsChecked Then
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                VarID += "_BSMDS"
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                VarID += "_BSMSI"
            End If
        ElseIf BtnTransportationCharges.IsChecked Then
            VarID += "_TRCH"
        ElseIf BtnTcsSummary.IsChecked Then
            VarID += "_TCSSM"
        ElseIf BtnGheeReport.IsChecked Then
            VarID += "_GHRP"
        ElseIf BtnRouteWiseSale.IsChecked Then
            VarID += "_RWS"
        ElseIf BtnCreditPartyWiseSaleAmount.IsChecked Then
            VarID += "_CPWSA"
        ElseIf rbtnDistributorCollStatement.IsChecked Then
            VarID += "_DSTC"
        ElseIf rbtnMilkSale.IsChecked Then
            VarID += "_MLKS"
        ElseIf rbtnStockStatement.IsChecked Then
            VarID += "_STOCK"
        End If
        If rbtnDetail.IsChecked Then
            VarID += "_DE"
        ElseIf rbtnSummary.IsChecked Then
            VarID += "_SM"
        End If
        gv1.VarID = VarID
        Report_ID = MyBase.Form_ID + "_" + VarID
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        GetReportID()
        Dim viewBlank As New TableViewDefinition()
        gv1.ViewDefinition = viewBlank
        If BtnProductWiseSaleQuantity.IsChecked Then
            ProductWiseSale(False)
        ElseIf BtnBillWiseSaleOfMilkSummary.IsChecked Then
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                BillWisesaleSummaryDispatch(False)
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                BillWisesaleSummaryInvoice(False)
            End If
        ElseIf BtnPartySaleMilkProduct.IsChecked Then
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                PartySaleMilkProductDispatch(False)
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                PartySaleMilkProductInvoice(False)
            End If
        ElseIf BtnBillWiseSaleOfMilk.IsChecked Then
            BillwiseSaleOfMilk(False)
        ElseIf BtnProductSalesSummary.IsChecked Then
            Productsalesummarytaxablenontaxable(False)
        ElseIf BtnMilkStcSummary.IsChecked Then
            MilkStcSummary(False)
        ElseIf BtnStcRegisterItemWiseSummary.IsChecked Then
            STCRegisterItemwiseSummarytotal(False)
            'STCRegisterItemwiseSummarytotal(False)
        ElseIf BtnStcRegisterPartyandItemWiseSummary.IsChecked Then
            STCRegisterItemwiseSummaryPartyWise(False)
        ElseIf BtnTransportationCharges.IsChecked Then
            TransportationCharges(False)
        ElseIf BtnTcsSummary.IsChecked Then
            TcsSummary(False)
        ElseIf BtnGheeReport.IsChecked Then
            GheeReport(False)
        ElseIf BtnRouteWiseSale.IsChecked Then
            RouteWiseSale(False)
        ElseIf BtnCreditPartyWiseSaleAmount.IsChecked Then
            CreditPartyWiseSaleAmount(False)
        ElseIf rbtnDistributorCollStatement.IsChecked Then
            LoadDistributorCollStatementData(False)
        ElseIf rbtnMilkSale.IsChecked Then
            LoadMilkSaleData(False)
        ElseIf rbtnStockStatement.IsChecked Then
            If rbtnSupplyDate.IsChecked Then
                clsCommon.MyMessageBoxShow(Me, "Stock Statement report not working with Supply Date Option! ", Me.Text)
                Exit Sub
            Else
                LoadStockStatementData(False)
            End If
        ElseIf rbtnBoothSaleItemWise.IsChecked Then
            LoadBoothSaleItemWiseData(False)
        ElseIf rbtnSalesRegister.IsChecked Then
            LoadSalesRegisterData(False)
        Else
            LoadData()
        End If
    End Sub

    Sub LoadDistributorCollStatementData(ByVal print As Boolean)
        Try
            If ddlQtyConversionType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            Dim whrcls As String = ""
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                whrcls = " And TSPL_SD_SHIPMENT_DETAIL.Item_Code In (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtRoute.arrValueMember IsNot Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
                whrcls += " and TSPL_SD_SHIPMENT_HEAD.Route_No in ( " + clsCommon.GetMulcallString(txtRoute.arrValueMember) + " )"
            End If
            If txtCustomer.arrValueMember Is Nothing Then
                clsCommon.MyMessageBoxShow(Me, "You must select at least one Customer with Distributor Collection Statement option", Me.Text)
                Exit Sub
            Else
                If txtCustomer.arrValueMember.Count = 1 Then
                    whrcls += " and Customer_Code in ( " + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + " )"
                ElseIf txtCustomer.arrValueMember.Count > 1 Then
                    clsCommon.MyMessageBoxShow(Me, "You cannot select more than one Customer at a time with Distributor Collection Statement option", Me.Text)
                    Exit Sub
                End If
            End If

            Dim qry As String = ""
            Dim strDate As String = ""
            If rbtnDocumentDate.IsChecked Then
                strDate = "Document_Date"
            ElseIf rbtnSupplyDate.IsChecked Then
                strDate = "Supply_Date"
            End If
            qry = " Select "
            If print Then
                qry += "  'Qty. in " + ddlQtyConversionType.SelectedValue + "' as QtyConvType,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy") + "' as ToDate,max(CompName)CompName,max(Add1)Add1, "
            End If
            qry += " Customer_Name,max(Item_Desc)Item_Desc,sum([Qty(M)])[Qty(M)],sum([Qty(E)])[Qty(E)],sum(Total_Qty)Total_Qty,max(case when Amount=0 then 0 else Amount/Total_Qty end) as Rate,sum(Amount)Amount  from ( SELECT tspl_customer_master.Customer_Name,TSPL_ITEM_MASTER.Item_Code,TSPL_SD_SHIPMENT_HEAD.Customer_Code as Cust_Code,TSPL_SD_SHIPMENT_HEAD." + strDate + " as document_date, Shift_Type, TSPL_ITEM_MASTER.Item_Desc ,
                TSPL_SD_SHIPMENT_DETAIL.Unit_Code AS UOM ,isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) AS Qty,case when Shift_Type='AM' then round((isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I." + ddlQtyConversionType.SelectedValue + ",2)else 0 end as [Qty(M)],case when Shift_Type='PM' then round((isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I." + ddlQtyConversionType.SelectedValue + ",2)else 0 end as [Qty(E)],
			    round((isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I." + ddlQtyConversionType.SelectedValue + ",2) as Total_Qty,isnull(TSPL_SD_SHIPMENT_DETAIL.Amount,0)Amount,TSPL_COMPANY_MASTER.Comp_Name AS CompName,TSPL_COMPANY_MASTER.Add1
                FROM TSPL_SD_SHIPMENT_DETAIL left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE Left outer join tspl_customer_master on tspl_customer_master.cust_code= TSPL_SD_SHIPMENT_HEAD.Customer_Code left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code left join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code   and TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_Code 
                Left Outer Join TSPL_COMPANY_MASTER on 1=1 left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [Pouch],[LTR],[Crate] )) P ) I ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = I.item_code  where 2 = 2 and TSPL_SD_SHIPMENT_HEAD.Status = 1 and convert(date,TSPL_SD_SHIPMENT_HEAD." + strDate + ",103) >= CONVERT(DATE, '" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "', 103) and convert(date,TSPL_SD_SHIPMENT_HEAD." + strDate + ",103) <= CONVERT(DATE, '" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "', 103) and TSPL_ITEM_MASTER.IsTaxable=0 " + whrcls + " )xx group by Customer_Name,Item_Code order by Item_Desc"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv1.DataSource = Nothing
                gv1.Rows.Clear()
                gv1.Columns.Clear()
                gv1.GroupDescriptors.Clear()
                gv1.MasterView.Refresh()
                gv1.GroupDescriptors.Clear()
                gv1.EnableFiltering = True
                gv1.MasterTemplate.SummaryRowsBottom.Clear()
                gv1.DataSource = dt
                gv1.BestFitColumns()
                SetGridFormationn()
                ReStoreGridLayout()
                gv1.MasterTemplate.AutoExpandGroups = True
                RadPageView1.SelectedPage = RadPageViewPage2
                gv1.BestFitColumns()
                If print Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "rptDistributorCollectionStatement", "Distributor Collection Statement")
                    frmCRV = Nothing
                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadMilkSaleData(ByVal Print As Boolean)
        Try
            If ddlReportType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Report Type ", Me.Text)
                Exit Sub
            End If
            If ddlQtyConversionType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            Dim whrcls As String = ""

            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                whrcls = " And TSPL_SD_SHIPMENT_DETAIL.Item_Code In (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                whrcls += " and TSPL_SD_SHIPMENT_HEAD.Customer_Code in ( " + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + " )"
            End If

            If txtRoute.arrValueMember IsNot Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
                whrcls += " and TSPL_SD_SHIPMENT_HEAD.Route_No in ( " + clsCommon.GetMulcallString(txtRoute.arrValueMember) + " )"
            End If

            Dim qry As String = "( SELECT right(TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode,5) as GPCode,TSPL_DAIRYSALE_GATEPASS_MASTER.GPDATE,TSPL_SD_SHIPMENT_HEAD.shift_type,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Logo_Img,TSPL_ITEM_MASTER.IsTaxable,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 ,TSPL_SD_SHIPMENT_HEAD.Customer_Code  ,TSPL_ITEM_MASTER.Short_Description,TSPL_CUSTOMER_MASTER.Customer_Name as Booth, "
            qry += "TSPL_SD_SHIPMENT_HEAD.Route_No,TSPL_SD_SHIPMENT_HEAD.Document_Date, TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,
TSPL_SD_SHIPMENT_DETAIL.Amount as Amount,"
            qry += " TSPL_SD_SHIPMENT_DETAIL.Unit_code, round((isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I." + ddlQtyConversionType.SelectedValue + ",2) as Qty,round((isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/RepUOM.Conversion_Factor,2) as RepUOM_Qty,round((isnull(TSPL_SD_SHIPMENT_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/DefUOM.conversion_factor,2) as Def_UOM_Qty,TSPL_ITEM_MASTER.Sku_Seq,"
            qry += "TSPL_CUSTOMER_MASTER.Display_Seq FROM TSPL_SD_SHIPMENT_DETAIL "
            qry += "  left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.Document_Code Left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_SD_SHIPMENT_HEAD.Route_No 
            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code left join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code   and TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_Code 
               left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [Pouch],[LTR],[Crate] )) P ) I ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = I.item_code
LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE DEFAULT_UOM = 1 ) DefUOM ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = DefUOM.item_code  LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE Report_UOM = 1 ) RepUOM ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = RepUOM.item_code
LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_SD_SHIPMENT_HEAD.Comp_Code left outer join TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL  on TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.pk_id =TSPL_SD_SHIPMENT_detail.pk_id
left outer join TSPL_DAIRYSALE_GATEPASS_MASTER  on TSPL_DAIRYSALE_GATEPASS_MASTER.gpcode =TSPL_DAIRYSALE_GATEPASS_SHIPMENT_DETAIL.GPCode where 2 = 2  and TSPL_SD_SHIPMENT_HEAD.Status = 1  "
            qry += "" & whrcls & "  "
            Dim strDate As String = ""
            If rbtnDocumentDate.IsChecked Then
                strDate = "Document_Date"
            ElseIf rbtnSupplyDate.IsChecked Then
                strDate = "Supply_Date"
            End If
            qry += " and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"

            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "'  and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(ddlType.SelectedValue, "Milk") = CompairStringResult.Equal Then
                qry += " and  TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 0 "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Product") = CompairStringResult.Equal Then
                qry += " and (TSPL_ITEM_MASTER.Is_Ambient = 1 or TSPL_ITEM_MASTER.IsTaxable = 1)  "
            End If
            qry += " )"
            Dim dtPrint As DataTable = clsDBFuncationality.GetDataTable(qry + "  order by Sku_Seq")
            If dtPrint Is Nothing OrElse dtPrint.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow("No Data Found to Print")
                Exit Sub
            ElseIf dtPrint.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                    dtPrint = clsDBFuncationality.GetDataTable(" select '" & objCommonVar.CurrentUser & "' as UserName,max(Comp_Name)Comp_Name,max(add1)Add1,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtFromDate1.Value), "dd-MM-yyyy") + "' AS FromDate,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtToDate1.Value), "dd-MM-yyyy") + "' AS ToDate,Customer_Code,max(Booth)Booth,CASE WHEN ROW_NUMBER() OVER (PARTITION BY XX.GPCode ORDER BY Sku_Seq) = 1  THEN XX.GPCode ELSE '' END AS GPCode,format(GPDATE,'dd-MM-yyyy') as GPDATE,max(Short_Description)Short_Description,sum(RepUOM_Qty)LtrQty,sum(Def_UOM_Qty)Def_UOM_Qty from  ( " + qry + " )xx   group by customer_code,gpcode,gpdate,item_code,Sku_Seq order by booth,XX.GPCode,Sku_Seq ")
                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dtPrint, Nothing, "rptMilkSaleAsPerGatePassPartyWise", "Milk Sale As Per Gate Pass Party Wise")
                ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                    Dim dtItems As DataTable = clsDBFuncationality.GetDataTable("select Item_Code,max(Sku_Seq) as Sku_Seq,max(Short_Description) as Short_Description,max(Unit_code)Unit_code from (" + qry + ") x group by Item_Code order by Sku_Seq")
                    Dim FinalQuery As String = " With CTERawData as ( " + qry + "  )" + Environment.NewLine + Environment.NewLine
                    For ii As Integer = 1 To dtItems.Rows.Count Step 10
                        If ii > 1 Then
                            FinalQuery += Environment.NewLine + " Union all " + Environment.NewLine
                        End If
                        FinalQuery += " select '" & objCommonVar.CurrentUser & "' as UserName," + clsCommon.myCstr(ii) + " as Grp ,'" + txtFromShift.Text + "' as FromShift,'" + txtToShift.Text + "' as ToShift, ROW_NUMBER() over (order by max(Booth)) As SNo, 'Quantity. in " + ddlQtyConversionType.SelectedValue + "' as QtyConvType,  '" + clsCommon.GetPrintDate(clsCommon.myCDate(txtFromDate1.Value), "dd/MM/yyyy") + "' AS FromDate,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtToDate1.Value), "dd/MM/yyyy") + "' AS ToDate, max(Comp_Name) as Comp_Name,max(Add1) as Add1,max(Booth) as Booth,"
                        FinalQuery += "Customer_Code,max(Route_No) as Route_No,"
                        FinalQuery += "max(Document_Date) as Document_Date"
                        For jj As Integer = 1 To 10
                            Dim strJJ As String = clsCommon.myCstr(jj)
                            Dim strICODE As String = ""
                            Dim strIShortDesc As String = ""
                            Dim strIUnitCode As String
                            If (ii + jj - 1) > dtItems.Rows.Count Then
                                strICODE = ""
                                strIShortDesc = ""
                                strIUnitCode = "-"
                            Else
                                strICODE = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Item_Code"))
                                strIShortDesc = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Short_Description"))
                                strIUnitCode = "" + ddlQtyConversionType.SelectedValue + ""
                            End If
                            FinalQuery += " ,'" + strICODE + "' as Item_" + strJJ + " ,'" + strIShortDesc + "' as Item_Short_Description_" + strJJ + " ,'" + strIUnitCode + "' as Item_Unit_code_" + strJJ + "
    ,(sum(case when Item_Code='" + strICODE + "'  then Qty else 0 end )) as ItemQtyCrateTotal_" + strJJ + "
    ,Case when (sum(case when Item_Code='" + strICODE + "' then convert(decimal(18,2),Qty) else null end )) is null then '-' ELSE  convert(varchar,sum(case when Item_Code='" + strICODE + "'  then convert(decimal(18,2),Qty) else null end))
    End   As ItemQtyCrate_" + strJJ + ""
                        Next

                        FinalQuery += " ,sum(Amount*case when IsTaxable=0 then 1 else 0 end) as Amount ,sum(qty)TotalQty"
                        FinalQuery += ",max(Display_Seq) as Display_Seq from ( select xx.*	from CTERawData xx ) x group by Customer_Code "
                    Next
                    dtPrint = clsDBFuncationality.GetDataTable(FinalQuery + " order by grp,Booth")
                    frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dtPrint, Nothing, "rptMilkSaleAsPerGatePass", "Milk Sale As Per Gate Pass")
                End If
                frmCRV = Nothing
            End If
            Exit Sub
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadStockStatementData(ByVal print As Boolean)
        Try
            If ddlDefaultReportUOM.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            Dim whrcls As String = ""
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                whrcls = " And InventroyMovement.Item_Code In (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                whrcls = " And TSPL_LOCATION_MASTER.Location_Code In (" & clsCommon.GetMulcallString(txtLocation.arrValueMember) & ") "
            End If

            Dim qry As String = ""
            qry = " Select "
            If print Then
                qry += "  TSPL_COMPANY_MASTER.Comp_Name AS CompName,TSPL_COMPANY_MASTER.Add1,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy") + "' as ToDate ,"
            End If
            qry += " *,OPBal+Production_In_Qty+Other_In_Qty-Sale_Qty-STC_Qty-Production_Out_Qty-Other_Out_Qty as Closing_Qty, Inter_Union_Sale FROM ( select row_number() OVER( order by  max(Item_Desc)) as Sno,MAX(Item_Desc) as Item_Desc,max(Report_UOM)as Report_UOM , sum(Report_UOM_Qty * (CASE WHEN PUNCHING_DAte < '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' THEN 1.00 ELSE 0 end) * (case when InOut='I' then 1.00 else -1.00 end))  AS [OPBal]  , 
            sum (Report_UOM_Qty * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and (Trans_Type ='DRY-PRO-UPL' or Trans_Type= 'PROD_ENTRY') THEN 1.00 ELSE 0 end) * (case when InOut='I' then 1.00 else 0 end))  AS Production_In_Qty,SUM(Report_UOM_Qty * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and (Trans_Type ='IC-AD' or Trans_Type= 'SRN' OR Trans_Type='FS-SR' or Trans_Type ='ITransfer') THEN 1.00 ELSE 0 end) * (case when InOut='I' then 1.00 else 0 end))  AS Other_In_Qty  ,
            sum (Report_UOM_Qty * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and (Trans_Type ='FS-SH' ) THEN 1.00 ELSE 0 end) * (case when InOut='O' then 1.00 else 0 end))  AS Sale_Qty,sum (Report_UOM_Qty * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and (Trans_Type ='ITransfer' AND Transfer_Type <> 'T' ) THEN 1.00 ELSE 0 end) * (case when InOut='O' then 1.00 else 0 end))  AS STC_Qty,sum (Report_UOM_Qty * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and Trans_Type ='ITransfer'  AND Transfer_Type ='T' AND IsProdTransfer = 1  THEN 1.00 ELSE 0 end) * (case when InOut='O' then 1.00 else 0 end))  AS Production_Out_Qty,
            SUM(Report_UOM_Qty * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and Trans_Type ='ITransfer' AND Transfer_Type ='T' and IsProdTransfer =0 THEN 1.00 ELSE 0 end) * (case when InOut='O' then 1.00 else 0 end))  AS Other_Out_Qty,sum(Inter_Union_Sale)Inter_Union_Sale from (  " + Environment.NewLine + "  select 
   CASE WHEN EXISTS ( SELECT 1 FROM TSPL_INVENTORY_MOVEMENT IM2 INNER JOIN TSPL_LOCATION_MASTER LM2 
            ON LM2.Location_Code = IM2.Location_Code
        WHERE IM2.Source_Doc_No = InventroyMovement.Source_Doc_No
          AND IM2.Item_Code = InventroyMovement.Item_Code
          AND IM2.InOut = 'I'
          AND LM2.IsProduction = 1
    ) THEN 1 ELSE 0 END AS IsProdTransfer, (isnull((InventroyMovement.Stock_Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(Report_UOM.Conversion_Factor),0) * (CASE WHEN PUNCHING_DAte >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' AND PUNCHING_DAte <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and (isnull(tspl_customer_master.Inter_Union_Sale,0)=1) THEN 1.00 ELSE 0 end) )  AS Inter_Union_Sale ,qty,InventroyMovement.Trans_Id,InventroyMovement.Trans_Type, (CASE WHEN (InventroyMovement.Trans_Type='IC-AD' AND TSPL_ADJUSTMENT_HEADER.Reference_Document='JWO-SRN-JLO') THEN 'Jobwork Consumption' ELSE  TSPL_INVENTORY_SOURCE_CODE.Name END )as Trans_Type_Name,InventroyMovement.Source_Doc_No,InventroyMovement.Punching_Date, InventroyMovement.InOut,case when InventroyMovement.InOut='I' then 'In' else case when InventroyMovement.InOut='O' then 'Out' else '' end end as 'InOutView',
            case when TSPL_LOCATION_MASTER.Is_Section='N' and TSPL_LOCATION_MASTER.Is_Sub_Location='N' then TSPL_LOCATION_MASTER.Location_Code else TSPL_LOCATION_MASTER.Main_Location_Code end as Main_Location_Code,MainLocationTable.Location_Desc as MainLocationDesc, InventroyMovement.Location_Code,TSPL_LOCATION_MASTER.Location_Desc AS [Loc Desp],SourceCode,SourceName,SourceType ,InventroyMovement.Item_Code, InventroyMovement.MRP ,TSPL_ITEM_MASTER.Item_Desc,TSPL_LOCATION_MASTER.Is_Sub_Location, InventroyMovement.Stock_UOM,InventroyMovement.Stock_Qty,isnull((InventroyMovement.Stock_Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(Report_UOM.Conversion_Factor),0) As Report_UOM_Qty,Report_UOM.UOM_Code as Report_UOM, (case when TSPL_PURCHASE_ACCOUNTS.Costing_Method=3 then InventroyMovement.FIFO_Cost else case when TSPL_PURCHASE_ACCOUNTS.Costing_Method=2 then InventroyMovement.LIFO_Cost else InventroyMovement.Avg_Cost end end ) as Cost,TSPL_INVENTORY_SOURCE_CODE.In_Category,TSPL_INVENTORY_SOURCE_CODE.Out_Category,TSPL_INVENTORY_SOURCE_CODE.Code ,tspl_transfer_order_head.Transfer_Type,tspl_location_master.IsProduction
            from  ( select qty,Trans_Id,Trans_Type,Source_Doc_No,Punching_Date,InOut,Location_Code,Item_Code,UOM, MRP,Stock_UOM,Stock_Qty,FIFO_Cost,LIFO_Cost,Avg_Cost,0 as IsFromMilk,case when cust_code is not null and len(cust_code)>0 then cust_code else case when Vendor_Code is not null and len(Vendor_Code)>0 then Vendor_Code else Other_Location_Code end end as SourceCode,case when cust_code is not null and len(cust_code)>0 then Cust_Name else case when Vendor_Code is not null and len(Vendor_Code)>0 then Vendor_Name else Other_Location_Desc end end as SourceName, case when cust_code is not null and len(cust_code)>0 then 'C' else case when Vendor_Code is not null and len(Vendor_Code)>0 then 'V' else case when Other_Location_Code is not null and len(Other_Location_Code)>0 then 'L' else '' end end end as SourceType,'' as Custom_UOM,0 as Custom_Coversion_Factor  from TSPL_INVENTORY_MOVEMENT ) InventroyMovement  left outer join tspl_transfer_order_head on tspl_transfer_order_head.Document_No=InventroyMovement.Source_Doc_No
            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=InventroyMovement.Item_Code  left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code=TSPL_ITEM_MASTER.Purchase_Class_Code  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code = InventroyMovement.Location_Code  left outer join TSPL_LOCATION_MASTER as MainLocationTable on MainLocationTable.Location_Code =(case when TSPL_LOCATION_MASTER.Is_Section='N' and  TSPL_LOCATION_MASTER.Is_Sub_Location='N' then TSPL_LOCATION_MASTER.Location_Code else TSPL_LOCATION_MASTER.Main_Location_Code end) left outer join TSPL_ITEM_UOM_DETAIL ON tspl_item_uom_detail.Item_Code=InventroyMovement.Item_Code and tspl_item_uom_detail.UOM_Code= InventroyMovement.Stock_UOM  LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where " + ddlDefaultReportUOM.SelectedValue + " = 1 ) as  Report_UOM ON InventroyMovement.Item_Code = Report_UOM.item_code 
             left outer join TSPL_INVENTORY_SOURCE_CODE on TSPL_INVENTORY_SOURCE_CODE.code=InventroyMovement.Trans_Type  left outer join TSPL_ADJUSTMENT_HEADER ON TSPL_ADJUSTMENT_HEADER.Adjustment_No=InventroyMovement.Source_Doc_No   left outer join TSPL_GL_ACCOUNTS on TSPL_GL_ACCOUNTS.Account_Code =TSPL_PURCHASE_ACCOUNTS .Inv_Control_Account   left outer join TSPL_GL_ACCOUNTS gl1 on gl1.Account_Seg_Code1 =TSPL_GL_ACCOUNTS.Account_Seg_Code1  and gl1.Account_Seg_Code7 =  tspl_location_master.Loc_Segment_Code left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=InventroyMovement.Source_Doc_No
			left outer join tspl_customer_master  on tspl_customer_master.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code  Where 2=2  and TSPL_LOCATION_MASTER.GIT_Type<>'Y' and MainLocationTable.GIT_Type<>'Y' and tspl_item_master.Item_Type='F' " + whrcls + " ) xx group by Item_Code ) xxx Left Outer Join TSPL_COMPANY_MASTER on 1=1 order by Item_Desc"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv1.DataSource = Nothing
                gv1.Rows.Clear()
                gv1.Columns.Clear()
                gv1.GroupDescriptors.Clear()
                gv1.MasterView.Refresh()
                gv1.GroupDescriptors.Clear()
                gv1.ShowGroupPanel = False
                gv1.EnableFiltering = True
                gv1.MasterTemplate.SummaryRowsBottom.Clear()
                gv1.DataSource = dt
                gv1.BestFitColumns()
                View()
                SetGridFormationn()
                ReStoreGridLayout()
                gv1.MasterTemplate.AutoExpandGroups = True
                RadPageView1.SelectedPage = RadPageViewPage2
                'gv1.BestFitColumns()
                If print Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "rptStockStatementBKN", "Stock Statement")
                    frmCRV = Nothing
                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Function getBoothSaleBaseQry() As String
        Dim whrcls As String = ""
        If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
            whrcls = " And TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code In (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
        End If
        If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
            whrcls += " and TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booth_Code in ( " + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + " )"
        End If

        If txtRoute.arrValueMember IsNot Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
            whrcls += " and TSPL_SD_SHIPMENT_HEAD.Route_No in ( " + clsCommon.GetMulcallString(txtRoute.arrValueMember) + " )"
        End If
        If clsCommon.CompairString(ddlType.SelectedValue, "Milk") = CompairStringResult.Equal Then
            whrcls += " and  TSPL_ITEM_MASTER.Is_FreshItem = 1 and TSPL_ITEM_MASTER.IsTaxable = 0 "
        ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Product") = CompairStringResult.Equal Then
            whrcls += " and (TSPL_ITEM_MASTER.Is_Ambient = 1 or TSPL_ITEM_MASTER.IsTaxable = 1)  "
        End If
        Dim qry As String = "( SELECT TSPL_SD_SHIPMENT_HEAD.shift_type,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Logo_Img,TSPL_ITEM_MASTER.IsTaxable,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 ,TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booth_Code as Customer_Code  ,TSPL_ITEM_MASTER.Short_Description,TSPL_CUSTOMER_MASTER.Customer_Name as Booth, "
        qry += "TSPL_SD_SHIPMENT_HEAD.Route_No,TSPL_SD_SHIPMENT_HEAD.Document_Date, TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,"
        qry += " TSPL_SD_SHIPMENT_BOOKING_DETAIL.Unit_code, isnull((TSPL_SD_SHIPMENT_BOOKING_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(Report_UOM.Conversion_Factor),0) As Qty,Report_UOM.UOM_Code as Def_Rep_UOM,TSPL_ITEM_MASTER.Sku_Seq,"
        qry += "TSPL_CUSTOMER_MASTER.Display_Seq FROM TSPL_SD_SHIPMENT_BOOKING_DETAIL "
        qry += "  left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Document_Code Left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_SD_SHIPMENT_HEAD.Route_No 
            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code left join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code   and TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Unit_Code 
               LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where  " + ddlDefaultReportUOM.SelectedValue + " = 1 ) as  Report_UOM ON TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code = Report_UOM.item_code 
LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booth_Code
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_SD_SHIPMENT_HEAD.Comp_Code where 2 = 2  and TSPL_SD_SHIPMENT_HEAD.Status = 1 "
        qry += "" & whrcls & "  "
        Dim strDate As String = ""
        If rbtnDocumentDate.IsChecked Then
            strDate = "Document_Date"
        ElseIf rbtnSupplyDate.IsChecked Then
            strDate = "Supply_Date"
        End If
        qry += " and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"

        If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
            qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
        End If
        If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
            qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "'  and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
        End If
        qry += " )"
        Return qry
    End Function
    Sub LoadBoothSaleDateShiftWiseData(ByVal Print As Boolean)
        Try
            If ddlDefaultReportUOM.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            Dim qry As String = getBoothSaleBaseQry()
            Dim dtPrint As DataTable = clsDBFuncationality.GetDataTable(qry + "  order by Sku_Seq")
            If dtPrint Is Nothing OrElse dtPrint.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow("No Data Found to Print")
                Exit Sub
            ElseIf dtPrint.Rows.Count > 0 Then
                Dim dtItems As DataTable = clsDBFuncationality.GetDataTable("select Item_Code,max(Sku_Seq) as Sku_Seq,max(Short_Description) as Short_Description,max(Def_Rep_UOM)Def_Rep_UOM from (" + qry + ") x group by Item_Code order by Sku_Seq")
                Dim FinalQuery As String = " With CTERawData as ( " + qry + "  )" + Environment.NewLine + Environment.NewLine
                For ii As Integer = 1 To dtItems.Rows.Count Step 7
                    If ii > 1 Then
                        FinalQuery += Environment.NewLine + " Union all " + Environment.NewLine
                    End If
                    FinalQuery += " select " + clsCommon.myCstr(ii) + " as Grp ,'" + txtFromShift.Text + "' as FromShift,'" + txtToShift.Text + "' as ToShift, ROW_NUMBER() OVER (PARTITION BY Customer_Code ORDER BY Customer_Code,Document_Date,Shift_Type) SNo, 'Quantity. in " + dtItems.Rows(0)("Def_Rep_UOM") + "' as QtyConvType,  '" + clsCommon.GetPrintDate(clsCommon.myCDate(txtFromDate1.Value), "dd/MM/yyyy") + "' AS FromDate,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtToDate1.Value), "dd/MM/yyyy") + "' AS ToDate,'" & objCommonVar.CurrentUser & "' as UserName, max(Comp_Name) as Comp_Name,max(Add1) as Add1,max(Booth) as Booth,"
                    FinalQuery += "Customer_Code,max(Route_No) as Route_No,"
                    FinalQuery += "convert(varchar,Document_Date,103) as Document_Date,case when Shift_Type='AM' then 'M' when Shift_Type = 'PM' then 'E' end as Shift_Type"
                    For jj As Integer = 1 To 7
                        Dim strJJ As String = clsCommon.myCstr(jj)
                        Dim strICODE As String = ""
                        Dim strIShortDesc As String = ""
                        Dim strIUnitCode As String
                        If (ii + jj - 1) > dtItems.Rows.Count Then
                            strICODE = ""
                            strIShortDesc = ""
                            strIUnitCode = "-"
                        Else
                            strICODE = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Item_Code"))
                            strIShortDesc = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Short_Description"))
                            strIUnitCode = clsCommon.myCstr(dtItems.Rows(ii + jj - 2)("Def_Rep_UOM"))
                        End If
                        FinalQuery += " ,'" + strICODE + "' as Item_" + strJJ + " ,'" + strIShortDesc + "' as Item_Short_Description_" + strJJ + " ,'" + strIUnitCode + "' as Item_Unit_code_" + strJJ + "
    ,(sum(case when Item_Code='" + strICODE + "'  then Qty else 0 end )) as ItemQtyCrateTotal_" + strJJ + "
    ,Case when (sum(case when Item_Code='" + strICODE + "' then convert(decimal(18,2),Qty) else null end )) is null then '-' ELSE  convert(varchar,sum(case when Item_Code='" + strICODE + "'  then convert(decimal(18,2),Qty) else null end))
    End   As ItemQtyCrate_" + strJJ + ""
                    Next
                    If ii = 1 Then
                        FinalQuery += " ,sum(qty)TotalQty"
                    Else
                        FinalQuery += " ,0 as TotalQty"
                    End If
                    FinalQuery += ",max(Display_Seq) as Display_Seq from ( select xx.*	from CTERawData xx ) x group by Document_Date,Shift_Type,Customer_Code "
                Next
                dtPrint = clsDBFuncationality.GetDataTable(FinalQuery + " order by grp,Booth ,Document_Date,Shift_Type desc")
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funsubreportWithdt(Form_ID, CrystalReportFolder.SalesReport, dtPrint, Nothing, "rptBoothSaleDateShiftWise", "Booth Sale")
                frmCRV = Nothing
            End If
            Exit Sub
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadBoothSaleItemWiseData(ByVal Print As Boolean)
        Try
            If ddlDefaultReportUOM.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            Dim qry As String = getBoothSaleBaseQry()
            qry = " select comp.Logo_Img,comp.Add1,comp.Comp_Name,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtFromDate1.Value), "dd/MM/yyyy") + "' AS FromDate,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtToDate1.Value), "dd/MM/yyyy") + "' AS ToDate,'" + txtFromShift.Text + "' as FromShift,'" + txtToShift.Text + "' as ToShift,'" & objCommonVar.CurrentUser & "' as UserName,xxx.* from ( select ROW_NUMBER() OVER (PARTITION BY Customer_Code ORDER BY Customer_Code,Sku_Seq) as SNo, Customer_Code,max(Booth)Booth,Item_Code,max(Item_Desc)Short_Description,sum(Qty)Qty,max(Def_Rep_UOM)Def_Rep_UOM from " & qry & "  xx
 group by Customer_Code,Item_Code,Sku_Seq)xxx left outer join TSPL_COMPANY_MASTER as comp on 1=1 "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If Not Print Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    SetGridFormationn()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf Print Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "rptBoothSaleItemWise", "Booth Sale Item Wise")
                    frmCRV = Nothing
                End If
            Else
                Throw New Exception("No Data Found!")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub STCRegisterItemwiseSummaryPartyWise(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            Qry = "Select To_LocationCode,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,  Max(xxx.Item_Desc)Item_Desc,
                        SUM(xxx.item_Net_Amt) AS Total_Item_Net_Amt,
                        SUM(xxx.Quantity) AS Total_Quantity,
                        SUM(xxx.Report_UOM_Qty)Report_UOM_Qty
                        ,MAX(Report_UOM)Report_UOM,
                        max(xxx.UOM1) as UOM1,
                        max(xxx.UOM2) as UOM2,
                        SUM(xxx.Weight) AS Total_Weight,
                        SUM(xxx.TotalQty) AS Total_TotalQty,
                        Max(xxx.TotalQtyUOM) as TotalQtyUOM,
                        max(xxx.CompName) as CompName,
                        MAX(xxx.add1) as add1,
                        MAX(xxx.To_LocationName) as Location_Desc,
                        SUM(xxx.DOC_Total_Amt) AS Total_Doc_Total_Amt,
                        SUM(xxx.Amount) AS Total_Amount,
                        SUM(price.TotalAmount) AS Total_Taxable_Amount,
                        SUM(price.Product_value) AS Total_Product_Value,
                        SUM(price.PTax1_Amt) AS Total_PTax1_Amt,
                        SUM(price.PTax2_Amt) AS Total_PTax2_Amt,
                        SUM(price.PTax3_Amt) AS Total_PTax3_Amt,
                        SUM(price.PTax4_Amt) AS Total_PTax4_Amt,
                        SUM(price.PTax5_Amt) AS Total_PTax5_Amt,
                        SUM(price.PTax6_Amt) AS Total_PTax6_Amt,
                        SUM(price.PTax7_Amt) AS Total_PTax7_Amt,
                        SUM(price.PTax8_Amt) AS Total_PTax8_Amt,
                        SUM(price.PTax9_Amt) AS Total_PTax9_Amt,
                        SUM(price.PTax10_Amt) AS Total_PTax10_Amt from (select TSPL_TRANSFER_ORDER_HEAD.Price_Code,TSPL_TRANSFER_ORDER_DETAIL.Document_No, TSPL_TRANSFER_ORDER_HEAD.Document_Date,        TSPL_TRANSFER_ORDER_DETAIL.item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc , TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity,isnull((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Report_UOM_Qty,I.UOM_Code as Report_UOM ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1,TSPL_ITEM_MASTER.Weight_UOM as UOM2, TSPL_ITEM_MASTER.Weight_Value as Weight,((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinLTR.Conversion_Factor ) As TotalQty, CFinLTR.UOM_Code As TotalQtyUOM,TSPL_COMPANY_MASTER.Comp_Name AS CompName,TSPL_COMPANY_MASTER.Add1,TSPL_LOCATION_MASTER_2.Location_Code as To_LocationCode,TSPL_LOCATION_MASTER_2.Location_Desc as To_LocationName,TSPL_LOCATION_MASTER.Location_Code as FromLocation,
                        TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate 
                        from TSPL_TRANSFER_ORDER_DETAIL 
                        join TSPL_TRANSFER_ORDER_HEAD  on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No  
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location   
                        left outer join TSPL_LOCATION_MASTER AS TSPL_LOCATION_MASTER_1 on TSPL_LOCATION_MASTER_1.GIT_LOCATION =  TSPL_TRANSFER_ORDER_HEAD.To_Location  
                        left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                        INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE  
                        Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='LTR') As CFinLTR On CFinLTR.Item_Code=TSPL_ITEM_MASTER.Item_Code
                        Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code 
                        LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code 
                        LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State  
                        left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  
                        LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = I.item_code
                        left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code  
                        left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id  
                        LEFT OUTER JOIN tspl_location_master  AS tspl_location_master_For_Location ON tspl_location_master_For_Location.GIT_Location =TSPL_TRANSFER_ORDER_HEAD.To_Location  
                        left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State  
                        LEFT OUTER JOIN TSPL_STATE_MASTER StateMaster_ToLocation ON StateMaster_ToLocation.State_Code=TSPL_LOCATION_MASTER_1.State 
                        left join TSPL_TRANSFER_ORDER_HEAD GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno 
                        left join tspl_location_master as From_GP_Location on From_GP_Location.Location_Code =GP.From_Location 
                        left join TSPL_STATE_MASTER as From_GP_Location_State on From_GP_Location_State.STATE_CODE =From_GP_Location.State  
                        left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.GIT_Location =  GP.To_Location  
                        left join TSPL_STATE_MASTER as To_GP_Location_State on To_GP_Location_State.STATE_CODE =To_GP_Location.State  
                        left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1 
                        left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2    
                        left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3   
                        left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4   
                        left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5   
                        left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6    
                        left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7    
                        left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8  
                        left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9    
                        left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10    
                        left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL .tax1 
                        left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2    
                        left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .TAX3   
                        left outer join TSPL_TAX_MASTER as dtax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_DETAIL .tax4   
                        left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .tax5   
                        left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX6    
                        left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX7    
                        left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX8  
                        left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX9    
                        left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX10  
                        LEFT JOIN TSPL_TAX_GROUP_MASTER ON TSPL_TRANSFER_ORDER_HEAD.Tax_Group=TSPL_TAX_GROUP_MASTER.Tax_Group_Code and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T' where 2=2 and TSPL_ITEM_MASTER.IsTaxable =1  and TSPL_TRANSFER_ORDER_HEAD.Status = 1  )xxx
                        cross apply ( select top 1 ((xxx.Quantity)*Item_Basic_Price) as TotalAmount,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1 as PTax1,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1_Rate as PTax1_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1_Amt*xxx.Quantity) as PTax1_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2 as PTax2,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2_Rate as PTax2_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2_Amt*xxx.Quantity) as PTax2_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3 as PTax3,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3_Rate as PTax3_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3_Amt * xxx.Quantity) as PTax3_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4 as PTax4,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4_Rate as PTax4_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4_Amt * xxx.Quantity) as PTax4_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5 as PTax5,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5_Rate as PTax5_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5_Amt * xxx.Quantity) as PTax5_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6 as PTax6,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6_Rate as PTax6_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6_Amt * xxx.Quantity) as PTax6_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7 as PTax7,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7_Rate as PTax7_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7_Amt * xxx.Quantity) as PTax7_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8 as PTax8,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8_Rate as PTax8_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8_Amt * xxx.Quantity) as PTax8_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9 as PTax9,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9_Rate as PTax9_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9_Amt * xxx.Quantity) as PTax9_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10 as PTax10,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10_Rate as PTax10_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10_Amt * xxx.Quantity) as PTax10_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_MRP,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Basic_Price,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Selling_Price,((xxx.Quantity)*Item_Selling_Price) as Product_value
						,TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code
						from TSPL_ITEM_PRICE_PLAN_DETAIL
                  left outer join TSPL_ITEM_PRICE_PLAN_HEADER on TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Plan_Code 
                  where xxx.Price_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code and xxx.Item_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Code 
				  and TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date <= xxx.Document_Date 
                  ORDER BY TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date DESC 
                  ) as price  where "
            'If rbtnSupplyDate.IsChecked Then
            '    Qry += " convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' "
            'ElseIf rbtnDocumentDate.IsChecked Then
            Qry += " convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' "
            'End If

            If txtItem.arrValueMember IsNot Nothing Then
                Qry += " and Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                Qry += " And FromLocation In (" & clsCommon.GetMulcallString(txtLocation.arrValueMember) & ") "
            End If
            Qry += " group by price.Price_Code, Item_Code,To_LocationCode order by Item_Desc"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemWiseSummaryPartyWisee", "STS Register-Item Wise Summary Party Wise")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub STCRegisterItemwiseSummarytotal(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            If ddlReportType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Report Type ", Me.Text)
                Exit Sub
            End If
            If ddlDefaultReportUOM.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            Qry = "Select  "
            If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                Qry += "  To_LocationCode,Max(xxx.Item_Desc)Item_Desc "
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                Qry += " Item_Code,Max(xxx.Item_Desc)Item_Desc,right(Document_No,5) as BillNo,format(Document_Date,'dd-MM-yyyy') as Document_Date,(To_LocationCode) as To_LocationCode"
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Summary") = CompairStringResult.Equal Then
                Qry += "  Item_Code,Max(xxx.Item_Desc) "
            End If

            Qry += " ,'" & objCommonVar.CurrentUser & "' as UserName ,max(To_LocationName)To_LocationName,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,  Max(xxx.Item_Desc)Item_Desc,
                        SUM(xxx.item_Net_Amt) AS Total_Item_Net_Amt,
                        SUM(xxx.Quantity) AS Total_Quantity,
                        SUM(xxx.Report_UOM_Qty)Report_UOM_Qty
                        ,MAX(Report_UOM)Report_UOM,sum(DefRepUOM_Qty) as DefRepUOM_Qty,MAX(DefRepUOM)DefRepUOM,
                        max(xxx.UOM1) as UOM1,
                        max(xxx.UOM2) as UOM2,
                        SUM(xxx.Weight) AS Total_Weight,
                        SUM(xxx.TotalQty) AS Total_TotalQty,
                        Max(xxx.TotalQtyUOM) as TotalQtyUOM,
                        max(xxx.CompName) as CompName,
                        MAX(xxx.add1) as add1,
                        MAX(xxx.To_LocationName) as Location_Desc,
                        SUM(xxx.DOC_Total_Amt) AS Total_Doc_Total_Amt,
                        SUM(xxx.Amount) AS Total_Amount,max(price.Item_MRP) AS Item_MRP,
                        SUM(price.TotalAmount) AS Total_Taxable_Amount,
                        SUM(price.Product_value) AS Total_Product_Value,
                        SUM(price.PTax1_Amt) AS Total_PTax1_Amt,
                        SUM(price.PTax2_Amt) AS Total_PTax2_Amt,
                        SUM(price.PTax3_Amt) AS Total_PTax3_Amt,
                        SUM(price.PTax4_Amt) AS Total_PTax4_Amt,
                        SUM(price.PTax5_Amt) AS Total_PTax5_Amt,
                        SUM(price.PTax6_Amt) AS Total_PTax6_Amt,
                        SUM(price.PTax7_Amt) AS Total_PTax7_Amt,
                        SUM(price.PTax8_Amt) AS Total_PTax8_Amt,
                        SUM(price.PTax9_Amt) AS Total_PTax9_Amt,
                        SUM(price.PTax10_Amt) AS Total_PTax10_Amt from (select TSPL_TRANSFER_ORDER_HEAD.Price_Code,TSPL_TRANSFER_ORDER_DETAIL.Document_No, TSPL_TRANSFER_ORDER_HEAD.Document_Date,TSPL_TRANSFER_ORDER_DETAIL.item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.item_code,TSPL_TRANSFER_ORDER_DETAIL.mrp,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc , TSPL_TRANSFER_ORDER_DETAIL.Item_Cost AS Item_Cost,  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Quantity,isnull((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Report_UOM_Qty,I.UOM_Code as Report_UOM,isnull((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(Report_UOM.Conversion_Factor),0)  As DefRepUOM_Qty,Report_UOM.UOM_Code as DefRepUOM ,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM1,TSPL_ITEM_MASTER.Weight_UOM as UOM2, TSPL_ITEM_MASTER.Weight_Value as Weight,((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty*tspl_item_uom_detail.Conversion_Factor)/CFinLTR.Conversion_Factor ) As TotalQty, CFinLTR.UOM_Code As TotalQtyUOM,TSPL_COMPANY_MASTER.Comp_Name AS CompName,TSPL_COMPANY_MASTER.Add1,TSPL_LOCATION_MASTER_2.Location_Code as To_LocationCode,TSPL_LOCATION_MASTER_2.Location_Desc as To_LocationName,TSPL_LOCATION_MASTER.Location_Code as FromLocation,
                        TSPL_TRANSFER_ORDER_HEAD.DOC_Total_Amt,TSPL_TRANSFER_ORDER_DETAIL.Amount,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate 
                        from TSPL_TRANSFER_ORDER_DETAIL 
                        join TSPL_TRANSFER_ORDER_HEAD  on TSPL_TRANSFER_ORDER_HEAD.Document_No   =TSPL_TRANSFER_ORDER_DETAIL.Document_No  
                        left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER .Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location   
                        left outer join TSPL_LOCATION_MASTER AS TSPL_LOCATION_MASTER_1 on TSPL_LOCATION_MASTER_1.GIT_LOCATION =  TSPL_TRANSFER_ORDER_HEAD.To_Location  
                        left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                        INNER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.ITEM_CODE = TSPL_TRANSFER_ORDER_DETAIL.iTEM_CODE  
                        Left Outer Join (Select Item_Code,Conversion_Factor,UOM_Code from tspl_item_uom_detail where UOM_Code='LTR') As CFinLTR On CFinLTR.Item_Code=TSPL_ITEM_MASTER.Item_Code
                        Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code 
                        LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code 
                        LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State  
                        left outer join tspl_item_uom_detail on tspl_item_uom_detail.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and tspl_item_uom_detail.uom_code=TSPL_TRANSFER_ORDER_DETAIL.unit_code  
                        LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = I.item_code
                        LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where " + ddlDefaultReportUOM.SelectedValue + " = 1 ) as  Report_UOM ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = Report_UOM.item_code 
                        left outer join tspl_item_uom_detail alt_convrsn on alt_convrsn.item_code=TSPL_TRANSFER_ORDER_DETAIL.item_code and alt_convrsn.uom_code=TSPL_TRANSFER_ORDER_DETAIL.alt_unit_code  
                        left outer join TSPL_TRANSPORT_MASTER on TSPL_TRANSPORT_MASTER.transport_id=TSPL_TRANSFER_ORDER_HEAD.transport_id  
                        LEFT OUTER JOIN tspl_location_master  AS tspl_location_master_For_Location ON tspl_location_master_For_Location.GIT_Location =TSPL_TRANSFER_ORDER_HEAD.To_Location  
                        left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE=TSPL_LOCATION_MASTER.State  
                        LEFT OUTER JOIN TSPL_STATE_MASTER StateMaster_ToLocation ON StateMaster_ToLocation.State_Code=TSPL_LOCATION_MASTER_1.State 
                        left join TSPL_TRANSFER_ORDER_HEAD GP on GP.document_no=TSPL_TRANSFER_ORDER_HEAD.transferoutno 
                        left join tspl_location_master as From_GP_Location on From_GP_Location.Location_Code =GP.From_Location 
                        left join TSPL_STATE_MASTER as From_GP_Location_State on From_GP_Location_State.STATE_CODE =From_GP_Location.State  
                        left outer join TSPL_LOCATION_MASTER AS To_GP_Location on To_GP_Location.GIT_Location =  GP.To_Location  
                        left join TSPL_STATE_MASTER as To_GP_Location_State on To_GP_Location_State.STATE_CODE =To_GP_Location.State  
                        left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_TRANSFER_ORDER_HEAD.tax1 
                        left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_TRANSFER_ORDER_HEAD.tax2    
                        left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .TAX3   
                        left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_HEAD .tax4   
                        left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_TRANSFER_ORDER_HEAD .tax5   
                        left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX6    
                        left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX7    
                        left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX8  
                        left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX9    
                        left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_TRANSFER_ORDER_HEAD .TAX10    
                        left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_TRANSFER_ORDER_DETAIL .tax1 
                        left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_TRANSFER_ORDER_DETAIL.tax2    
                        left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .TAX3   
                        left outer join TSPL_TAX_MASTER as dtax4 on tax4.Tax_Code= TSPL_TRANSFER_ORDER_DETAIL .tax4   
                        left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_TRANSFER_ORDER_DETAIL .tax5   
                        left outer join TSPL_TAX_MASTER as dtax6 on dtax6.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX6    
                        left outer join TSPL_TAX_MASTER as dtax7 on dtax7.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX7    
                        left outer join TSPL_TAX_MASTER as dtax8 on dtax8.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX8  
                        left outer join TSPL_TAX_MASTER as dtax9 on dtax9.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX9    
                        left outer join TSPL_TAX_MASTER as dtax10 on dtax10.Tax_Code =TSPL_TRANSFER_ORDER_DETAIL .TAX10  
                        LEFT JOIN TSPL_TAX_GROUP_MASTER ON TSPL_TRANSFER_ORDER_HEAD.Tax_Group=TSPL_TAX_GROUP_MASTER.Tax_Group_Code and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='T' where 2=2 and TSPL_ITEM_MASTER.IsTaxable =1 and TSPL_TRANSFER_ORDER_HEAD.Status = 1 and TSPL_TRANSFER_ORDER_HEAD.Transfer_Type = 'O' )xxx
                        cross apply ( select top 1 ((xxx.Quantity)*Item_Basic_Price) as TotalAmount,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1 as PTax1,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1_Rate as PTax1_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX1_Amt*xxx.Quantity) as PTax1_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2 as PTax2,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2_Rate as PTax2_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX2_Amt*xxx.Quantity) as PTax2_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3 as PTax3,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3_Rate as PTax3_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX3_Amt * xxx.Quantity) as PTax3_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4 as PTax4,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4_Rate as PTax4_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX4_Amt * xxx.Quantity) as PTax4_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5 as PTax5,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5_Rate as PTax5_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX5_Amt * xxx.Quantity) as PTax5_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6 as PTax6,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6_Rate as PTax6_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX6_Amt * xxx.Quantity) as PTax6_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7 as PTax7,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7_Rate as PTax7_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX7_Amt * xxx.Quantity) as PTax7_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8 as PTax8,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8_Rate as PTax8_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX8_Amt * xxx.Quantity) as PTax8_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9 as PTax9,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9_Rate as PTax9_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX9_Amt * xxx.Quantity) as PTax9_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10 as PTax10,TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10_Rate as PTax10_Rate,(TSPL_ITEM_PRICE_PLAN_DETAIL.TAX10_Amt * xxx.Quantity) as PTax10_Amt,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_MRP,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Basic_Price,TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Selling_Price,((xxx.Quantity)*Item_Selling_Price) as Product_value
						,TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code
						from TSPL_ITEM_PRICE_PLAN_DETAIL
                  left outer join TSPL_ITEM_PRICE_PLAN_HEADER on TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Plan_Code 
                  where xxx.Price_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code and xxx.Item_Code = TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Code 
				  and TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date <= xxx.Document_Date 
                  ORDER BY TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date DESC 
                  ) as price  where
                  convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' "
            If txtItem.arrValueMember IsNot Nothing Then
                Qry += " and Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                Qry += " And FromLocation In (" & clsCommon.GetMulcallString(txtLocation.arrValueMember) & ") "
            End If
            If txtToLocation.arrValueMember IsNot Nothing AndAlso txtToLocation.arrValueMember.Count > 0 Then
                Qry += " And To_LocationCode In (" & clsCommon.GetMulcallString(txtToLocation.arrValueMember) & ") "
            End If
            If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                Qry += "  GROUP BY  To_LocationCode,Item_Code order by To_LocationName,max(xxx.Item_Desc)"
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                Qry += "  GROUP BY Item_Code,Document_No,Document_Date ,To_LocationCode order by max(xxx.Item_Desc),BillNo,Document_Date,To_LocationName "
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Summary") = CompairStringResult.Equal Then
                Qry += "  GROUP BY Item_Code order by max(xxx.Item_Desc)  "
            End If
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemWiseSummary", "STS Register Item Wise Summary")
                    ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemPartyWiseSummary", "STS Register Item Party Wise Summary")
                    ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Summary") = CompairStringResult.Equal Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-Summary", "STS Register Summary")
                    End If

                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub MilkStcSummary(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whrcls As String = ""
            If ddlReportType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Report Type ", Me.Text)
                Exit Sub
            End If
            If txtItem.arrValueMember IsNot Nothing Then
                whrcls += " and Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                whrcls += " And FromLoc.Location_Code In (" & clsCommon.GetMulcallString(txtLocation.arrValueMember) & ") "
            End If
            If txtToLocation.arrValueMember IsNot Nothing AndAlso txtToLocation.arrValueMember.Count > 0 Then
                whrcls += " And TSPL_LOCATION_MASTER_2.Location_Code In (" & clsCommon.GetMulcallString(txtToLocation.arrValueMember) & ") "
            End If
            Qry = " select "
            If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                Qry += "  To_LocationCode,Max(xxx.Item_Desc)Item_Desc "
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                Qry += " Item_Code,Max(xxx.Item_Desc)Item_Desc,right(Document_No,5) as BillNo,format(Document_Date,'dd-MM-yyyy') as Document_Date,(To_LocationCode) as To_LocationCode"
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Summary") = CompairStringResult.Equal Then
                Qry += " Item_Code,Max(xxx.Item_Desc)Item_Desc "
            End If

            Qry += ",'" & objCommonVar.CurrentUser & "' as UserName, max(To_LocationName)To_LocationName, SUM(Out_Qty) AS Out_Qty, SUM(QtyPouch) AS QtyPouch, sum(qtyLtr)[Qty(Ltr)],sum(Rep_UOM_Qty)Rep_UOM_Qty,sum(Def_UOM_Qty)Def_UOM_Qty,
    SUM(Amount) AS Amount, 
    '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' AS Fromdate, 
    '" + clsCommon.GetPrintDate(txtToDate.Value) + "' AS ToDate, 
    MAX(Add1) AS Add1, 
    MAX(Add2) AS Add2, 
    MAX(Add3) AS Add3, 
    MAX(City_Code) AS City_Code, 
    MAX(State) AS State
FROM (
    SELECT 
        xx.*, 
        (xx.Out_Qty * tab2.Rate) AS Amount, 
        tab2.Rate
    FROM (
        SELECT TSPL_TRANSFER_ORDER_HEAD.Document_NO, (TSPL_LOCATION_MASTER_2.Location_Code) as To_LocationCode,
						max(TSPL_LOCATION_MASTER_2.Location_Desc) as To_LocationName,(FromLoc.Location_Code) as FromLocation,
            TSPL_TRANSFER_ORDER_DETAIL.Item_Code, 
            MAX(TSPL_TRANSFER_ORDER_HEAD.Price_Code) AS Price_Code, 
            TSPL_TRANSFER_ORDER_HEAD.Document_Date, 
            '' AS Fromdate, 
            '' AS ToDate, 
            MAX(TSPL_COMPANY_MASTER.Comp_Name) AS Comp_Name, 
            MAX(TSPL_COMPANY_MASTER.Add1) AS Add1, 
            MAX(TSPL_COMPANY_MASTER.Add2) AS Add2, 
            MAX(TSPL_COMPANY_MASTER.Add3) AS Add3, 
            MAX(TSPL_LOCATION_MASTER.City_Code) AS City_Code, 
            MAX(TSPL_LOCATION_MASTER.State) AS State, 
            MAX(TSPL_TRANSFER_ORDER_DETAIL.Item_Desc) AS Item_Desc, 
            SUM(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty) AS Out_Qty, 
            MAX(TSPL_TRANSFER_ORDER_DETAIL.Unit_code) AS UoM, 
            MAX(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * 
                ItemConvReportUOM.Conversion_Factor / 
                ItemConvinUOMpouch.Conversion_Factor) AS QtyPouch, 
    MAX(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * 
                ItemConvReportUOM.Conversion_Factor / 
                ItemConvinUOMLtr.Conversion_Factor) AS QtyLtr,sum(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * ItemConvReportUOM.Conversion_Factor / RepUOM.Conversion_Factor ) as Rep_UOM_Qty,
				sum(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * ItemConvReportUOM.Conversion_Factor / DefUOM.Conversion_Factor) as Def_UOM_Qty,
            TSPL_TRANSFER_ORDER_DETAIL.Unit_code
        FROM TSPL_TRANSFER_ORDER_DETAIL 
        LEFT OUTER JOIN TSPL_TRANSFER_ORDER_HEAD 
        ON TSPL_TRANSFER_ORDER_HEAD.Document_No = TSPL_TRANSFER_ORDER_DETAIL.Document_No
        LEFT OUTER JOIN TSPL_ITEM_MASTER 
        ON TSPL_ITEM_MASTER.Item_Code = TSPL_TRANSFER_ORDER_DETAIL.Item_Code
        LEFT JOIN TSPL_ITEM_UOM_DETAIL AS ItemConvReportUOM
        ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = ItemConvReportUOM.Item_Code
        AND TSPL_TRANSFER_ORDER_DETAIL.Unit_code = ItemConvReportUOM.UOM_Code
        LEFT JOIN TSPL_ITEM_UOM_DETAIL AS ItemConvinUOMpouch ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = ItemConvinUOMpouch.Item_Code AND ItemConvinUOMpouch.UOM_Code = 'Pouch'
    	LEFT JOIN TSPL_ITEM_UOM_DETAIL AS ItemConvinUOMLtr  ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = ItemConvinUOMLtr.Item_Code AND ItemConvinUOMLtr.UOM_Code = 'LTR'
    LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE Report_UOM = 1 ) RepUOM ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = RepUOM.item_code
LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE DEFAULT_UOM = 1 ) DefUOM ON TSPL_TRANSFER_ORDER_DETAIL.Item_Code = DefUOM.item_code
        Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_TRANSFER_ORDER_HEAD.Comp_Code  LEFT OUTER JOIN TSPL_LOCATION_MASTER 
        ON TSPL_LOCATION_MASTER.Location_Code = TSPL_TRANSFER_ORDER_DETAIL.Location 
    left outer join TSPL_LOCATION_MASTER as FromLoc on FromLoc.Location_Code=  TSPL_TRANSFER_ORDER_HEAD.From_Location 
    left outer join  TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_2 on TSPL_LOCATION_MASTER_2.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
        WHERE  IsTaxable=0 and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " + whrcls + " GROUP BY TSPL_LOCATION_MASTER_2.Location_Code,FromLoc.Location_Code, TSPL_TRANSFER_ORDER_HEAD.Document_NO,  TSPL_TRANSFER_ORDER_HEAD.Document_Date, TSPL_TRANSFER_ORDER_DETAIL.Item_Code, TSPL_TRANSFER_ORDER_DETAIL.Unit_code
    ) xx
    OUTER APPLY ( 
        SELECT TOP 1 TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Basic_Price AS Rate
        FROM TSPL_ITEM_PRICE_PLAN_DETAIL
        LEFT JOIN TSPL_ITEM_PRICE_PLAN_HEADER 
        ON TSPL_ITEM_PRICE_PLAN_DETAIL.Plan_Code = TSPL_ITEM_PRICE_PLAN_HEADER.Plan_Code
        WHERE TSPL_ITEM_PRICE_PLAN_DETAIL.Price_Code = xx.Price_Code  
        AND TSPL_ITEM_PRICE_PLAN_DETAIL.Item_Code = xx.Item_Code 
        AND TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date <= xx.Document_Date 
        ORDER BY TSPL_ITEM_PRICE_PLAN_HEADER.Start_Date DESC 
    ) AS tab2
) xxx  "
            If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                Qry += "  GROUP BY  To_LocationCode,Item_Code order by To_LocationName,max(xxx.Item_Desc)"
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                Qry += "  GROUP BY Item_Code,Document_No,Document_Date ,To_LocationCode order by max(xxx.Item_Desc),Document_Date,max(xxx.To_LocationName) "
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Summary") = CompairStringResult.Equal Then
                Qry += "  GROUP BY Item_Code order by max(xxx.Item_Desc)  "
            End If
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptMilkStcSummaryPartyWise", "Milk Stc Summary Party Wise")
                    ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptMilkStcSummary", "Milk Stc Summary Item Wise")
                    ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Summary") = CompairStringResult.Equal Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptMilkStcItemSummary", "Milk Stc Summary")
                    End If

                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub Productsalesummarytaxablenontaxable(ByVal print As Boolean)
        Try
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            If ddlReportType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Report Type ", Me.Text)
                Exit Sub
            End If
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            Qry = " SELECT '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,'" & objCommonVar.CurrentUser & "' as UserName, "
            If clsCommon.CompairString(ddlType.SelectedValue, "Both") = CompairStringResult.Equal Then
                Qry += "'Taxable & Non Taxable' as TaxableNonTaxable, "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal Then
                Qry += "'Taxable' as TaxableNonTaxable, "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                Qry += "'Non Taxable' as TaxableNonTaxable, "
            End If

            Qry += " Sum(Distributor_Commission_TotalAmt) as Trp_Charge,max(Cust_Code)Cust_Code,MAX(Customer_Name) as Customer_Name,
                               MAX(Item_Desc) AS Item_Desc,Item_Code, MAX(Unit_code) AS Unit_code,SUM(Def_UOM_Qty) AS Qty,sum(Amount) as Amount,"

            If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlType.SelectedValue, "Both") = CompairStringResult.Equal Then
                Qry += "  sum([KKF Amt]) as KKF_Amt,SUM(Amount) + SUM([Mandi Tax Amt]) + SUM([KKF Amt]) AS Taxable_Value,SUM([Mandi Tax Amt]) as Mandi_Tax_Amt,SUM([CGST Amt]) as CGST_Amt,sum([SGST Amt]) as SGST_Amt,SUM([IGST Amt]) as IGST_Amt,SUM([TCS Amt]) as TCS_Amt,SUM(Amount) + SUM([Mandi Tax Amt]) + SUM([KKF Amt])+SUM([CGST Amt])+sum([SGST Amt])+SUM([IGST Amt]) as [Total Amount],"
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                Qry += " SUM(Amount) AS Taxable_Value,SUM([CGST Amt]) as CGST_Amt,sum([SGST Amt]) as SGST_Amt,SUM([IGST Amt]) as IGST_Amt,SUM(Amount)+SUM([CGST Amt])+sum([SGST Amt])+SUM([IGST Amt]) as [Total Amount],"
            End If
            Qry += " SUM(QtyAccToReportUOM) AS QtyAccToReportUOM,MAX(UOM_Code) AS UOM_Code,MAX(Comp_Name) AS Comp_Name,MAX(Add1) AS Add1,MAX(Add2) AS Add2,MAX(Add3) AS Add3,MAX(City_Code) AS City_Code,MAX(State) AS State, MAX(Document_Date) AS Document_Date
                                FROM (SELECT TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.Item_Code, 
                            TSPL_SD_SHIPMENT_DETAIL.Unit_code,TSPL_SD_SHIPMENT_DETAIL.Qty,cast(
                                    (
                                      TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / DefUOM.Conversion_Factor
                                    ) as Decimal(18, 2)
                                  ) as Def_UOM_Qty,
                            cast((TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor) as Decimal(18, 2)) as QtyAccToReportUOM,ItemConvReportUOM.UOM_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2, 
                            TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.City_Code,TSPL_COMPANY_MASTER.State,TSPL_SD_SHIPMENT_DETAIL.Amt_Less_Discount as Amount,TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No,TSPL_SD_SHIPMENT_HEAD.TAX5_Amt,TSPL_SD_SHIPMENT_HEAD.Distributor_Commission_TotalAmt,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,
							CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='KKF'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='KKF' THEN TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt else 0 END 
    				AS [KKF Amt],
                           CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='MANDITAX'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='MANDITAX' THEN TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt else 0 END 
    				AS  [Mandi Tax Amt],
                    CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='CGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='CGST' THEN TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt else 0 END  AS [CGST Amt],
                             CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='SGST'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='SGST' THEN TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt else 0 END  AS [SGST Amt]
					,CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='IGST'  THEN isnull (TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='IGST'  THEN isnull (TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='IGST'  THEN isnull (TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='IGST'  THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='IGST'  THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='IGST'  THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='IGST'  THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='IGST'  THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='IGST'  THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt,0)
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='IGST' THEN isnull(TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt,0) else 0 END AS [IGST Amt],
							 CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='TCS'  THEN TSPL_SD_SHIPMENT_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='TCS' THEN TSPL_SD_SHIPMENT_DETAIL.TAX10_Amt else 0 END  AS [TCS Amt],
                    Document_Date 
                    FROM TSPL_SD_SHIPMENT_DETAIL
                    LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
					left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
                    LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SHIPMENT_DETAIL.item_code
                    left join TSPL_ITEM_UOM_DETAIL as ItemConvReportUOM on TSPL_ITEM_master.Item_Code = ItemConvReportUOM.Item_Code and ItemConvReportUOM.Report_UOM = 1
LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE DEFAULT_UOM = 1 ) DefUOM ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = DefUOM.item_code
                    left join TSPL_ITEM_UOM_DETAIL as ItemConvinUOM on TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemConvinUOM.Item_Code and TSPL_SD_SHIPMENT_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
                        LEFT JOIN TSPL_COMPANY_MASTER ON 2 = 2
outer apply ( select top 1 IS_TAXABLE,ITEM_CODE from TSPL_ITEM_MASTER_TAXABLE 
where  ITEM_CODE = TSPL_SD_SHIPMENT_DETAIL.Item_Code and EFFECTIVE_DATE <= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' order by EFFECTIVE_DATE desc ) as ItemTaxable 
                            where 2=2  	"

            If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal Then
                Qry += " and ItemTaxable.IS_TAXABLE = 1  "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                Qry += " and TSPL_ITEM_MASTER.Is_Ambient = 1 AND ItemTaxable.IS_TAXABLE = 0 "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Both") = CompairStringResult.Equal Then
                Qry += " and ( ( Is_FreshItem = 1 and ItemTaxable.IS_TAXABLE = 1 ) or Is_Ambient=1 ) "
            End If
            If txtCustomer.arrValueMember IsNot Nothing Then
                Qry += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If
            If txtLocation.arrValueMember IsNot Nothing Then
                Qry += " and TSPL_SD_SHIPMENT_HEAD.Sub_Location_code in (" & clsCommon.GetMulcallString(txtLocation.arrValueMember) & ") "
            End If
            If txtItem.arrValueMember IsNot Nothing Then
                Qry += " and TSPL_SD_SHIPMENT_DETAIL.Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If rbtnSupplyDate.IsChecked Then
                Qry += " and convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "') xx "
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " and convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "') xx "
            End If


            If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                Qry += "  GROUP BY  Cust_Code,Item_Code order by Customer_Name,Item_Desc"
            ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                Qry += "  GROUP BY Item_Code order by Item_Desc"
            End If

            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If BtnProductSalesSummary.IsChecked Then
                        If clsCommon.CompairString(ddlReportType.SelectedValue, "Party Wise") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlType.SelectedValue, "Both") = CompairStringResult.Equal Then
                                frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductSaleSummaryPartyTaxableNonTaxable", "Product Sale Summary Party Wise Taxable NonTaxable")
                            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                                frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductSaleSummaryPartyNonTaxable", "Product Sale Summary Party Wise NonTaxable")
                            End If
                        ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Item Wise") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlType.SelectedValue, "Both") = CompairStringResult.Equal Then
                                frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductSaleSummaryTaxableNonTaxable", "Product Sale Summary Taxable NonTaxable")
                            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                                frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductSaleSummaryNonTaxable", "Product Sale Summary NonTaxable")
                            End If
                        End If

                    End If
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Sub BillwiseSaleOfMilk(ByVal print As Boolean)
        Try
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            Dim TableName As String = ""
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                TableName = "TSPL_SD_SHIPMENT"
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                TableName = "TSPL_SD_SALE_INVOICE"
            End If

            Qry = " SELECT '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                   
	                   " & IIf(clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal, "(" & TableName & "_DETAIL.DOCUMENT_CODE) as DOCUMENT_CODE", "" & TableName & "_DETAIL.DOCUMENT_CODE ") & " ,
    MAX(TSPL_ITEM_MASTER.Item_Desc) AS Item_Desc, 
    MAX(TSPL_ITEM_MASTER.Item_Code) AS Item_Code, 
    MAX(" & TableName & "_DETAIL.Unit_code) AS Unit_code, SUM(" & TableName & "_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvDefaultUOM.Conversion_Factor ) AS Qty,
    SUM(
        " & TableName & "_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
    ) AS QtyAccToReportUOM,
   SUM(Amount) - max(ISNULL(" & TableName & "_HEAD.Transporter_Commission_TotalAmt, 0)) + 
       sum( CASE 
            WHEN " & TableName & "_HEAD.TAX1 = 'TCS' THEN " & TableName & "_DETAIL.TAX1_Amt
            WHEN " & TableName & "_HEAD.TAX2 = 'TCS' THEN " & TableName & "_DETAIL.TAX2_Amt
            WHEN " & TableName & "_HEAD.TAX3 = 'TCS' THEN " & TableName & "_DETAIL.TAX3_Amt
            WHEN " & TableName & "_HEAD.TAX4 = 'TCS' THEN " & TableName & "_DETAIL.TAX4_Amt
            WHEN " & TableName & "_HEAD.TAX5 = 'TCS' THEN " & TableName & "_DETAIL.TAX5_Amt
            WHEN " & TableName & "_HEAD.TAX6 = 'TCS' THEN " & TableName & "_DETAIL.TAX6_Amt
            WHEN " & TableName & "_HEAD.TAX7 = 'TCS' THEN " & TableName & "_DETAIL.TAX7_Amt
            WHEN " & TableName & "_HEAD.TAX8 = 'TCS' THEN " & TableName & "_DETAIL.TAX8_Amt
            WHEN " & TableName & "_HEAD.TAX9 = 'TCS' THEN " & TableName & "_DETAIL.TAX9_Amt
            WHEN " & TableName & "_HEAD.TAX10 = 'TCS' THEN " & TableName & "_DETAIL.TAX10_Amt 
            ELSE 0 
        END)
    AS Total_Amount,
    MAX(ItemConvReportUOM.UOM_Code) AS UOM_Code, 
    MAX(TSPL_COMPANY_MASTER.Comp_Name) AS Comp_Name, 
    MAX(TSPL_COMPANY_MASTER.Add1) AS Add1, 
    MAX(TSPL_COMPANY_MASTER.Add2) AS Add2, 
    MAX(TSPL_COMPANY_MASTER.Add3) AS Add3, 
    MAX(TSPL_COMPANY_MASTER.City_Code) AS City_Code, 
    MAX(TSPL_COMPANY_MASTER.State) AS State,
    SUM(" & TableName & "_DETAIL.Amount) AS Amount, "

            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                Qry += " MAX(" & TableName & "_HEAD.Sale_Invoice_No) AS Sale_Invoice_No "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                Qry += " MAX(" & TableName & "_HEAD.DOCUMENT_CODE) AS Sale_Invoice_No "
            End If

            Qry += " ,SUM(" & TableName & "_HEAD.TAX5_Amt) AS TAX5_Amt,
    max(" & TableName & "_HEAD.Transporter_Commission_TotalAmt) AS Transporter_Commission_Amt,
    MAX(TSPL_CUSTOMER_MASTER.Customer_Name) AS Customer_Name,
    SUM(
        CASE 
            WHEN " & TableName & "_HEAD.TAX1 = 'TCS' THEN " & TableName & "_DETAIL.TAX1_Amt
            WHEN " & TableName & "_HEAD.TAX2 = 'TCS' THEN " & TableName & "_DETAIL.TAX2_Amt
            WHEN " & TableName & "_HEAD.TAX3 = 'TCS' THEN " & TableName & "_DETAIL.TAX3_Amt
            WHEN " & TableName & "_HEAD.TAX4 = 'TCS' THEN " & TableName & "_DETAIL.TAX4_Amt
            WHEN " & TableName & "_HEAD.TAX5 = 'TCS' THEN " & TableName & "_DETAIL.TAX5_Amt
            WHEN " & TableName & "_HEAD.TAX6 = 'TCS' THEN " & TableName & "_DETAIL.TAX6_Amt
            WHEN " & TableName & "_HEAD.TAX7 = 'TCS' THEN " & TableName & "_DETAIL.TAX7_Amt
            WHEN " & TableName & "_HEAD.TAX8 = 'TCS' THEN " & TableName & "_DETAIL.TAX8_Amt
            WHEN " & TableName & "_HEAD.TAX9 = 'TCS' THEN " & TableName & "_DETAIL.TAX9_Amt
            WHEN " & TableName & "_HEAD.TAX10 = 'TCS' THEN " & TableName & "_DETAIL.TAX10_Amt 
            ELSE 0 
        END
    ) AS [TCS Amt],
    MAX(" & TableName & "_HEAD.Document_Date) AS Document_Date
FROM " & TableName & "_DETAIL
LEFT OUTER JOIN " & TableName & "_HEAD 
    ON " & TableName & "_HEAD.Document_Code = " & TableName & "_DETAIL.DOCUMENT_CODE
LEFT OUTER JOIN TSPL_CUSTOMER_MASTER 
    ON TSPL_CUSTOMER_MASTER.Cust_Code = " & TableName & "_HEAD.Customer_Code "

            If clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                Qry += " left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No  "
            End If
            Qry += "     Left OUTER JOIN TSPL_ITEM_MASTER 
    On TSPL_ITEM_MASTER.item_code = " & TableName & "_DETAIL.item_code
Left Join TSPL_ITEM_UOM_DETAIL AS ItemConvReportUOM  ON TSPL_ITEM_MASTER.Item_Code = ItemConvReportUOM.Item_Code And ItemConvReportUOM.Report_UOM = 1
    Left Join TSPL_ITEM_UOM_DETAIL AS ItemConvDefaultUOM ON " & TableName & "_DETAIL.Item_Code = ItemConvDefaultUOM.Item_Code And ItemConvDefaultUOM.Default_UOM = 1
Left Join TSPL_ITEM_UOM_DETAIL AS ItemConvinUOM 
    On " & TableName & "_DETAIL.Item_Code = ItemConvinUOM.Item_Code And " & TableName & "_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
Left Join TSPL_COMPANY_MASTER 
    On 2 = 2 where 2=2 "

            If rbtnSupplyDate.IsChecked Then
                Qry += " and cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date)>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as date )<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " and cast(" & TableName & "_HEAD.Document_Date as date)>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and cast(" & TableName & "_HEAD.Document_Date as date)<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"
            End If
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                If rbtnSupplyDate.IsChecked Then
                    Qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
                ElseIf rbtnDocumentDate.IsChecked Then
                    Qry += " and 2=( case when Cast(" & TableName & "_HEAD.Document_Date as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(" & TableName & "_HEAD.Document_Date as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
                End If
            End If

            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                If rbtnSupplyDate.IsChecked Then
                    Qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "' and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
                ElseIf rbtnDocumentDate.IsChecked Then
                    Qry += " and 2=( case when Cast(" & TableName & "_HEAD.Document_Date as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(" & TableName & "_HEAD.Document_Date as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "' and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
                End If
            End If

            Qry += "  and  TSPL_ITEM_MASTER.IsTaxable=0 "
            If txtCustomer.arrValueMember IsNot Nothing Then
                Qry += " And " & TableName & "_HEAD.Customer_Code in (" & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If
            If txtItem.arrValueMember IsNot Nothing Then
                Qry += " And " & TableName & "_DETAIL.Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            Qry += " Group BY " & IIf(clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal, "Sale_Invoice_No,TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE order by Sale_Invoice_No ", TableName & "_DETAIL.DOCUMENT_CODE order by Document_Code ") & " , Item_Desc"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptBillWiseSaleOfMilk", "Bill Wise Sale Of Milk")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub PartySaleMilkProductDispatch(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whrcls As String = ""
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " And TSPL_CUSTOMER_MASTER.Cust_Code in (" & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If
            Qry =
                 " select '" & objCommonVar.CurrentUser & "' as UserName,*,GstAMt + Taxable_Amount+ Non_Taxable_Amount as Sale_Amt,((Taxable_Amount + Non_Taxable_Amount + GSTAmt + TCS_AMT) - ([Trip & other Charge]))as Bill_Amt from ( 
                select 
            (XXFinal.Customer_Code) ,MAX(XXFinal.FromDate) as FromDate,MAX(XXFinal.ToDate) as ToDate,MAX(XXFinal.Customer_Name) as Customer_Name,SUM(XXFinal.[CGST Amt]+[SGST Amt]) as GstAMt,
            MAX(XXFinal.GSTNO) as GSTNO, SUM(XXFinal.Taxable_Amount+[KKF_Amt]+[Mandi_Tax_Amt]) as Taxable_Amount ,SUM(XXFinal.Non_Taxable_Amount)as Non_Taxable_Amount,
            SUM(XXFinal.[KKF_Amt]) as KKF,SUM(XXFinal.[Mandi_Tax_Amt]) as MandiTax,
            SUM(XXFinal.TCS_AMT) as TCS_AMT,sum(XXFinal.Trp_othcharg) as [Trip & other Charge],
            MAX(XXFinal.Comp_Name) as Comp_Name,
            MAX(XXFinal.CompAddress) as CompAddress
            from(
            select xx.* ,Taxable_Amount - [CGST Amt Head] - [SGST Amt Head] as Taxable_Amount1  
            from(
                                            select 
                                            TSPL_CUSTOMER_MASTER.Customer_Name as Customer_Name,TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_SD_SHIPMENT_HEAD.Document_Code,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as FromDate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                                            TSPL_COMPANY_MASTER.Comp_Name,
                                            (TSPL_COMPANY_MASTER.Add1 + TSPL_COMPANY_MASTER.Add2+TSPL_COMPANY_MASTER.Add3) as CompAddress, 
                                            TSPL_CUSTOMER_MASTER.GSTNO as GSTNO,
                                            case when TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='T' Then TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount else 0 end as Taxable_Amount,
                                             case when TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='NT' Then TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount else 0 end as Non_Taxable_Amount,
                                             Case when TSPL_SD_SHIPMENT_HEAD.TAX1 = 'IGST' Then isnull(TSPL_SD_SHIPMENT_HEAD.TAX1_Amt,0) else(
                                            Case when ISNULL(TSPL_SD_SHIPMENT_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SHIPMENT_HEAD.tax2,'')='KKF' Then (case when TSPL_SD_SHIPMENT_HEAD.TAX3='IGST' Then (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt+TSPL_SD_SHIPMENT_HEAD.TAX2_Amt+TSPL_SD_SHIPMENT_HEAD.TAX3_Amt)else (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt+TSPL_SD_SHIPMENT_HEAD.TAX2_Amt+TSPL_SD_SHIPMENT_HEAD.TAX3_Amt +TSPL_SD_SHIPMENT_HEAD.TAX4_Amt) end) else 0 end) end as GSTAmt,

                                            CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX1_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX2_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX3_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX4_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX5_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX6_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX7_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX8_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX9_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='CGST' THEN TSPL_SD_SHIPMENT_HEAD.TAX10_Amt else 0 END  AS [CGST Amt],

                                           CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX1_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX2_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX3_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX4_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX5_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX6_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX7_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX8_Amt
                				            WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX9_Amt
               				                WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='SGST' THEN TSPL_SD_SHIPMENT_HEAD.TAX10_Amt else 0 END  AS [SGST Amt],

                                            CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='KKF'  THEN TSPL_SD_SHIPMENT_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='KKF' THEN TSPL_SD_SHIPMENT_HEAD.TAX10_Amt else 0 END 
                            				AS [KKF_Amt],
                                                  CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='MANDITAX'  THEN TSPL_SD_SHIPMENT_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='MANDITAX' THEN TSPL_SD_SHIPMENT_HEAD.TAX10_Amt else 0 END 
                            				AS  [Mandi_Tax_Amt],
                                            CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='CGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='CGST' THEN TSPL_SD_SHIPMENT_HEAD.TAX10_Amt else 0 END 
                            				AS [CGST Amt Head],
                                                  CASE WHEN TSPL_SD_SHIPMENT_HEAD.TAX1='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX2='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX3='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX4='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX5='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX6='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX7='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX8='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX9='SGST'  THEN TSPL_SD_SHIPMENT_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SHIPMENT_HEAD.TAX10='SGST' THEN TSPL_SD_SHIPMENT_HEAD.TAX10_Amt else 0 END 
                            				AS  [SGST Amt Head],
                                            Case when TSPL_SD_SHIPMENT_HEAD.TAX1 = 'IGST' Then isnull(TSPL_SD_SHIPMENT_HEAD.TAX2_Amt,0) else(
                                            Case when ISNULL(TSPL_SD_SHIPMENT_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SHIPMENT_HEAD.tax2,'')='KKF' Then (case when TSPL_SD_SHIPMENT_HEAD.TAX3='IGST' Then TSPL_SD_SHIPMENT_HEAD.TAX4_Amt else (case when    
                                            TSPL_SD_SHIPMENT_HEAD.TAX5='TCS' Then TSPL_SD_SHIPMENT_HEAD.TAX5_Amt else 0 end) end) else (case when TSPL_SD_SHIPMENT_HEAD.tax2='TCS' Then TSPL_SD_SHIPMENT_HEAD.TAX2_Amt else ((case when TSPL_SD_SHIPMENT_HEAD.tax3='TCS' Then TSPL_SD_SHIPMENT_HEAD.TAX3_Amt else 0 end)) end) end) end as TCS_AMT,
                                        isnull(TSPL_SD_SHIPMENT_HEAD.Transporter_Commission_TotalAmt,0) as Trp_othcharg
                                            from TSPL_SD_SHIPMENT_HEAD
                                            
                                            left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
                                            left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code1='"& objCommonVar.CurrComp_Code1 &"' "
            If rbtnSupplyDate.IsChecked Then
                Qry += " WHERE Convert(Date, Supply_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'and TSPL_SD_SHIPMENT_HEAD.Status=1  " + whrcls + ""
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " WHERE Convert(Date, Document_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'and TSPL_SD_SHIPMENT_HEAD.Status=1  " + whrcls + ""
            End If

            Qry +=" ) xx
            ) XXFinal 
            group by XXFinal.Customer_Code)xxx order by Customer_Name"



            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptPartySalesMilkandProducts", "Party Sales Milk and Products")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub PartySaleMilkProductInvoice(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whrcls As String = ""
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " and TSPL_CUSTOMER_MASTER.Cust_Code in (" & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If
            Qry =
                 " select '" & objCommonVar.CurrentUser & "' as UserName,*,GstAMt + Taxable_Amount+ Non_Taxable_Amount as Sale_Amt,((Taxable_Amount + Non_Taxable_Amount + GSTAmt + TCS_AMT) - ([Trip & other Charge]))as Bill_Amt from ( 
                select 
            (XXFinal.Customer_Code) ,MAX(XXFinal.FromDate) as FromDate,MAX(XXFinal.ToDate) as ToDate,MAX(XXFinal.Customer_Name) as Customer_Name,SUM(XXFinal.[CGST Amt]+[SGST Amt]) as GstAMt,
            MAX(XXFinal.GSTNO) as GSTNO, SUM(XXFinal.Taxable_Amount+[KKF_Amt]+[Mandi_Tax_Amt]) as Taxable_Amount ,SUM(XXFinal.Non_Taxable_Amount)as Non_Taxable_Amount,
            SUM(XXFinal.[KKF_Amt]) as KKF,SUM(XXFinal.[Mandi_Tax_Amt]) as MandiTax,
            SUM(XXFinal.TCS_AMT) as TCS_AMT,sum(XXFinal.Trp_othcharg) as [Trip & other Charge],
            MAX(XXFinal.Comp_Name) as Comp_Name,
            MAX(XXFinal.CompAddress) as CompAddress
            from(
            select xx.* ,Taxable_Amount - [CGST Amt Head] - [SGST Amt Head] as Taxable_Amount1  
            from(
                                            select 
                                            TSPL_CUSTOMER_MASTER.Customer_Name as Customer_Name,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Code,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as FromDate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                                            TSPL_COMPANY_MASTER.Comp_Name,
                                            (TSPL_COMPANY_MASTER.Add1 + TSPL_COMPANY_MASTER.Add2+TSPL_COMPANY_MASTER.Add3) as CompAddress, 
                                            TSPL_CUSTOMER_MASTER.GSTNO as GSTNO,
                                             case when TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable=1 Then TSPL_SD_SALE_INVOICE_HEAD.Amount_Less_Discount else 0 end as Taxable_Amount,
                                             case when TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable=0 Then TSPL_SD_SALE_INVOICE_HEAD.Amount_Less_Discount else 0 end as Non_Taxable_Amount,
                                             Case when TSPL_SD_SALE_INVOICE_HEAD.TAX1 = 'IGST' Then isnull(TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt,0) else(
                                            Case when ISNULL(TSPL_SD_SALE_INVOICE_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SALE_INVOICE_HEAD.tax2,'')='KKF' Then (case when TSPL_SD_SALE_INVOICE_HEAD.TAX3='IGST' Then (TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt+TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt+TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt)else (TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt+TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt+TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt +TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt) end) else 0 end) end as GSTAmt,

                                            CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX7_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX8_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX9_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_HEAD.TAX10_Amt else 0 END  AS [CGST Amt],

                                           CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX7_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX8_Amt
                				            WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX9_Amt
               				                WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_HEAD.TAX10_Amt else 0 END  AS [SGST Amt],

                                            CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_HEAD.TAX10_Amt else 0 END 
                            				AS [KKF_Amt],
                                                  CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='MANDITAX'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='MANDITAX' THEN TSPL_SD_SALE_INVOICE_HEAD.TAX10_Amt else 0 END 
                            				AS  [Mandi_Tax_Amt],
                                            CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_HEAD.TAX10_Amt else 0 END 
                            				AS [CGST Amt Head],
                                                  CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX7_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX8_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_HEAD.TAX9_Amt
                            				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_HEAD.TAX10_Amt else 0 END 
                            				AS  [SGST Amt Head],
                                            Case when TSPL_SD_SALE_INVOICE_HEAD.TAX1 = 'IGST' Then isnull(TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt,0) else(
                                            Case when ISNULL(TSPL_SD_SALE_INVOICE_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SALE_INVOICE_HEAD.tax2,'')='KKF' Then (case when TSPL_SD_SALE_INVOICE_HEAD.TAX3='IGST' Then TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt else (case when    
                                            TSPL_SD_SALE_INVOICE_HEAD.TAX5='TCS' Then TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt else 0 end) end) else (case when TSPL_SD_SALE_INVOICE_HEAD.tax2='TCS' Then TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt else ((case when TSPL_SD_SALE_INVOICE_HEAD.tax3='TCS' Then TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt else 0 end)) end) end) end as TCS_AMT,
                                       isnull(TSPL_SD_SALE_INVOICE_HEAD.Transporter_Commission_TotalAmt,0) as Trp_othcharg 
                                            from TSPL_SD_SALE_INVOICE_HEAD
                                            left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
                                            left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                                            left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code1='BKN' "
            If rbtnSupplyDate.IsChecked Then
                Qry += "  where Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' "
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += "  where Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'  "
            End If
            Qry += "  and TSPL_SD_SALE_INVOICE_HEAD.Status=1  " + whrcls + "
               
            ) xx
            ) XXFinal 
            group by XXFinal.Customer_Code)xxx order by Customer_Name"



            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptPartySalesMilkandProducts", "Party Sales Milk and Products")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub TransportationCharges(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim whrcls As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""

            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " and TSPL_CUSTOMER_MASTER.Cust_Code in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If

            If txtItem.arrValueMember IsNot Nothing Then
                whrcls += " and Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            Qry = " SELECT  
                    max(XXFinal.Customer_Name) as Customer_Name,MAX(XXFinal.FromDate) as FromDate,MAX(XXFinal.ToDate) as ToDate,
                    sum(XXFinal.Trp_othcharg) as Trp_othcharg,SUM(XXFinal.Trp_othcharg) as Total_Amount,
                    MAX(XXFinal.Comp_Name) as Comp_Name,
                    MAX(XXFinal.CompAddress) as CompAddress
                    from(
                    select xx.*, (xx.Taxable_Amount +xx.Non_Taxable_Amount+xx.GSTAmt) as Sale_Amt,((xx.Taxable_Amount+xx.Non_Taxable_Amount+xx.GSTAmt+xx.TCS_AMT) -(xx.Trp_othcharg)) as Bill_Amt 
                    from(
                    select 
                    TSPL_CUSTOMER_MASTER.Customer_Name as Customer_Name,TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_SD_SHIPMENT_HEAD.Document_Code,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                                                        TSPL_COMPANY_MASTER.Comp_Name,
                    (TSPL_COMPANY_MASTER.Add1 ) as CompAddress, 
                     case when TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='T' then TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount else 0 end as Taxable_Amount,
                     case when TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='NT' then TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount else 0 end as Non_Taxable_Amount,
                     Case when TSPL_SD_SHIPMENT_HEAD.TAX1 = 'IGST' then isnull(TSPL_SD_SHIPMENT_HEAD.TAX1_Amt,0) else(
                    Case when ISNULL(TSPL_SD_SHIPMENT_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SHIPMENT_HEAD.tax2,'')='KKF' then (case when TSPL_SD_SHIPMENT_HEAD.TAX3='IGST' then (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt+TSPL_SD_SHIPMENT_HEAD.TAX2_Amt+TSPL_SD_SHIPMENT_HEAD.TAX3_Amt)else (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt+TSPL_SD_SHIPMENT_HEAD.TAX2_Amt+TSPL_SD_SHIPMENT_HEAD.TAX3_Amt +TSPL_SD_SHIPMENT_HEAD.TAX4_Amt) end) else 0 end) end as GSTAmt,
                    Case when TSPL_SD_SHIPMENT_HEAD.TAX1 = 'IGST' then isnull(TSPL_SD_SHIPMENT_HEAD.TAX2_Amt,0) else(
                    Case when ISNULL(TSPL_SD_SHIPMENT_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SHIPMENT_HEAD.tax2,'')='KKF' then (case when TSPL_SD_SHIPMENT_HEAD.TAX3='IGST' then TSPL_SD_SHIPMENT_HEAD.TAX4_Amt else (case when    
                    TSPL_SD_SHIPMENT_HEAD.TAX5='TCS' then TSPL_SD_SHIPMENT_HEAD.TAX5_Amt else 0 end) end) else (case when TSPL_SD_SHIPMENT_HEAD.tax2='TCS' then TSPL_SD_SHIPMENT_HEAD.TAX2_Amt else ((case when TSPL_SD_SHIPMENT_HEAD.tax3='TCS' then TSPL_SD_SHIPMENT_HEAD.TAX3_Amt else 0 end)) end) end) end as TCS_AMT,

                    (isnull(TSPL_SD_SHIPMENT_HEAD.Distributor_Commission_TotalAmt,0) + isnull(TSPL_SD_SHIPMENT_HEAD.Transporter_Commission_TotalAmt,0) +isnull(TSPL_SD_SHIPMENT_HEAD.Security_TotalAmt,0) +isnull(TSPL_SD_SHIPMENT_HEAD.BoothSecurity_TotalAmt,0) ) as Trp_othcharg
                    from TSPL_SD_SHIPMENT_HEAD
                    left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
                    left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "' "
            If rbtnSupplyDate.IsChecked Then
                Qry += " WHERE Convert(Date, Supply_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' and TSPL_SD_SHIPMENT_HEAD.Status=1 " & whrcls & ""
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " WHERE Convert(Date, Document_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' and TSPL_SD_SHIPMENT_HEAD.Status=1 " & whrcls & ""
            End If


            Qry +=" ) xx
                        ) XXFinal 
                        where XXFinal.Trp_othcharg>0
                        group by XXFinal.Customer_Name order by XXFinal.Customer_Name"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptTransportationCharges", "Transportation Charges")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub TcsSummary(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim whrcls As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " and TSPL_CUSTOMER_MASTER.Cust_Code in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If

            Qry = " select 
                        max(XXFinal.Customer_Name) as Customer_Name,max(XXFinal.pan) as Pan,MAX(XXFinal.FromDate) as FromDate,
                        MAX(XXFinal.ToDate) as ToDate,
                        sum(XXFinal.Total_Amt-XXFinal.TCS_AMT) as Sale_Amount,
                        SUM(XXFinal.TCS_AMT) as TCS_AMT,
                        MAX(XXFinal.Comp_Name) as Comp_Name,
                        MAX(XXFinal.CompAddress) as CompAddress
                        from(
                        select xx.*, (xx.Taxable_Amount +xx.Non_Taxable_Amount+xx.GSTAmt) as Sale_Amt,((xx.Taxable_Amount+xx.Non_Taxable_Amount+xx.GSTAmt+xx.TCS_AMT) -(xx.Trp_othcharg)) as Bill_Amt
                        from(
                        select 
                        TSPL_CUSTOMER_MASTER.Customer_Name as Customer_Name,TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_SD_SHIPMENT_HEAD.Document_Code,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                         TSPL_COMPANY_MASTER.Comp_Name,

                        (TSPL_COMPANY_MASTER.Add1 ) as CompAddress,
                         case when TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='T' then TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount else 0 end as Taxable_Amount,
                         case when TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='NT' then TSPL_SD_SHIPMENT_HEAD.Amount_Less_Discount else 0 end as Non_Taxable_Amount,

                         Case when TSPL_SD_SHIPMENT_HEAD.TAX1 = 'IGST' then isnull(TSPL_SD_SHIPMENT_HEAD.TAX1_Amt,0) else(
                        Case when ISNULL(TSPL_SD_SHIPMENT_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SHIPMENT_HEAD.tax2,'')='KKF' then (case when TSPL_SD_SHIPMENT_HEAD.TAX3='IGST' then (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt+TSPL_SD_SHIPMENT_HEAD.TAX2_Amt+TSPL_SD_SHIPMENT_HEAD.TAX3_Amt)else (TSPL_SD_SHIPMENT_HEAD.TAX1_Amt+TSPL_SD_SHIPMENT_HEAD.TAX2_Amt+TSPL_SD_SHIPMENT_HEAD.TAX3_Amt +TSPL_SD_SHIPMENT_HEAD.TAX4_Amt) end) else 0 end) end as GSTAmt,
                        Case when TSPL_SD_SHIPMENT_HEAD.TAX1 = 'IGST' then isnull(TSPL_SD_SHIPMENT_HEAD.TAX2_Amt,0) else(
                        Case when ISNULL(TSPL_SD_SHIPMENT_HEAD.tax1,'')='KKF' or ISNULL(TSPL_SD_SHIPMENT_HEAD.tax2,'')='KKF' then (case when TSPL_SD_SHIPMENT_HEAD.TAX3='IGST' then TSPL_SD_SHIPMENT_HEAD.TAX4_Amt else (case when    
                        TSPL_SD_SHIPMENT_HEAD.TAX5='TCS' then TSPL_SD_SHIPMENT_HEAD.TAX5_Amt else 0 end) end) else (case when TSPL_SD_SHIPMENT_HEAD.tax2='TCS' then TSPL_SD_SHIPMENT_HEAD.TAX2_Amt else ((case when TSPL_SD_SHIPMENT_HEAD.tax3='TCS' then TSPL_SD_SHIPMENT_HEAD.TAX3_Amt else 0 end)) end) end) end as TCS_AMT,TSPL_SD_SHIPMENT_HEAD.Total_Amt,
                        (isnull(TSPL_SD_SHIPMENT_HEAD.Distributor_Commission_TotalAmt,0) + isnull(TSPL_SD_SHIPMENT_HEAD.Transporter_Commission_TotalAmt,0) +isnull(TSPL_SD_SHIPMENT_HEAD.Security_TotalAmt,0) +isnull(TSPL_SD_SHIPMENT_HEAD.BoothSecurity_TotalAmt,0) ) as Trp_othcharg,TSPL_CUSTOMER_MASTER.PAN
                        from TSPL_SD_SHIPMENT_HEAD
                        left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
                        left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code1='"& objCommonVar.CurrComp_Code1 &"' "
            If rbtnSupplyDate.IsChecked Then
                Qry += " WHERE convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'"
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " WHERE convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'"
            End If
            Qry += "  and TSPL_SD_SHIPMENT_HEAD.Status=1  " & whrcls & "
                        ) xx
                        ) XXFinal 
                        where XXFinal.TCS_AMT>0 group by XXFinal.Customer_Name order by XXFinal.Customer_Name"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptTcsSummaryReport", "Tcs Summary Report")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub GheeReport(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim whrcls As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""

            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " and TSPL_CUSTOMER_MASTER.Cust_Code in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If

            Qry = " select '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                        Max(Customer_Name) as Customer_Name,
						Max(Item_Desc) as Item_Desc,Item_Code ,
                         Max(Unit_code) AS Unit_code,Sum(Qty) AS Qty,Sum(QtyAccToReportUOM) AS QtyAccToReportUOM,Max(UOM_Code) AS UOM_Code,max(Comp_Name) AS Comp_Name,max(Add1) AS Add1
                    FROM (
                        SELECT 
	                    TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,
                        TSPL_CUSTOMER_MASTER.Customer_Name,
                        TSPL_SD_SHIPMENT_DETAIL.Unit_code, 
                        TSPL_SD_SHIPMENT_DETAIL.Qty,
						cast((TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
						) as Decimal(18, 2)
						) as QtyAccToReportUOM,
                        ItemConvReportUOM.UOM_Code, 
                        TSPL_COMPANY_MASTER.Comp_Name, 
                        TSPL_COMPANY_MASTER.Add1 ,
						TSPL_ITEM_MASTER.Item_Desc,
						TSPL_ITEM_MASTER.Item_Code
                        FROM TSPL_SD_SHIPMENT_DETAIL
                        LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
                        LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SHIPMENT_DETAIL.item_code
						left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Customer_Code
                        left join TSPL_ITEM_UOM_DETAIL as ItemConvReportUOM on TSPL_ITEM_master.Item_Code = ItemConvReportUOM.Item_Code 
                        and ItemConvReportUOM.Report_UOM = 1
                        left join TSPL_ITEM_UOM_DETAIL as ItemConvinUOM on TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemConvinUOM.Item_Code 
                        and TSPL_SD_SHIPMENT_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
						LEFT JOIN TSPL_COMPANY_MASTER ON 2 = 2 "
            If rbtnSupplyDate.IsChecked Then
                Qry += " WHERE convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'"
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " WHERE convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'"
            End If
            Qry += " and  TSPL_ITEM_MASTER.Is_Ambient=1  AND TSPL_ITEM_MASTER.Item_Desc LIKE '%Ghee%'  " & whrcls & " ) xx
						group By Customer_Name,Item_Code order by Customer_Name,Item_Desc"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptGheeReport", "Ghee Report")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Sub RouteWiseSale(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            Qry = " select '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                                                MAX(TSPL_ROUTE_MASTER.Route_Desc) AS Route_Desc,
                                                MAX(TSPL_SD_SHIPMENT_HEAD.Route_No) AS Route_No,
                            MAX(TSPL_SD_SHIPMENT_HEAD.Shift_Type) AS Shift_Type,
                            SUM(
                                TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
                            ) AS QtyAccToReportUOM,
                            SUM(
                                CASE 
                                    WHEN TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'Am' 
                                    THEN TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor 
                                    ELSE 0 
                                END
                            ) AS qty_Am,
                            SUM(
                                CASE 
                                    WHEN TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'PM' 
                                    THEN TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor 
                                    ELSE 0 
                                END
                            ) AS qty_PM,
                            MAX(ItemConvReportUOM.UOM_Code) AS UOM_Code, 
                            MAX(TSPL_COMPANY_MASTER.Comp_Name) AS Comp_Name, 
                            MAX(TSPL_COMPANY_MASTER.Add1) AS Add1
                        FROM TSPL_SD_SHIPMENT_DETAIL
                        LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD 
                            ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
                        LEFT OUTER JOIN TSPL_CUSTOMER_MASTER 
                            ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code
                        LEFT OUTER JOIN TSPL_ITEM_MASTER 
                            ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SHIPMENT_DETAIL.item_code
                        left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No=TSPL_SD_SHIPMENT_HEAD.Route_No
                        LEFT JOIN TSPL_ITEM_UOM_DETAIL AS ItemConvReportUOM 
                            ON TSPL_ITEM_MASTER.Item_Code = ItemConvReportUOM.Item_Code AND ItemConvReportUOM.Report_UOM = 1
                        LEFT JOIN TSPL_ITEM_UOM_DETAIL AS ItemConvinUOM 
                            ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemConvinUOM.Item_Code AND TSPL_SD_SHIPMENT_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
                        LEFT JOIN TSPL_COMPANY_MASTER 
                            ON 2 = 2 "
            If rbtnSupplyDate.IsChecked Then
                Qry += " WHERE convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'"
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " WHERE convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'"
            End If
            Qry += "  AND TSPL_ITEM_MASTER.IsTaxable = 0 AND TSPL_ITEM_MASTER.Is_FreshItem=1 
                        AND TSPL_SD_SHIPMENT_HEAD.Route_No <> 'DIRECT' 
	                    AND TSPL_SD_SHIPMENT_HEAD.Route_No <> 'DIRECT1' 
                    GROUP BY TSPL_SD_SHIPMENT_HEAD.Route_No,TSPL_ROUTE_MASTER.Route_Seq_No order by TSPL_ROUTE_MASTER.Route_Seq_No"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptRouteWiseSale", "Route Wise Sale")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub



    Sub CreditPartyWiseSaleAmount(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim whrcls As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " and TSPL_CUSTOMER_MASTER.Cust_Code in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If
            Qry = " select '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                            TSPL_CUSTOMER_MASTER.Customer_Name,
                            Max(TSPL_CUSTOMER_MASTER.cust_code) as cust_code,
                            Sum(CASE 
                                WHEN TSPL_ITEM_MASTER.IsTaxable = 0 THEN TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt  
                                ELSE 0 
                            END) AS Milk_Amount, 
                            Sum(CASE 
                                WHEN TSPL_ITEM_MASTER.IsTaxable = 1 THEN TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt 
                                ELSE 0 
                            END) AS Product_Amount,
	                        Sum(TSPL_SD_SHIPMENT_DETAIL.Item_Net_Amt) AS amount,
	                        MAX(TSPL_COMPANY_MASTER.Comp_Name) AS Comp_Name, 
                            MAX(TSPL_COMPANY_MASTER.Add1) AS Add1 
                        FROM TSPL_SD_SHIPMENT_HEAD 
                        LEFT OUTER JOIN TSPL_SD_SHIPMENT_DETAIL 
                            ON TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE = TSPL_SD_SHIPMENT_HEAD.Document_Code
                        LEFT OUTER JOIN TSPL_CUSTOMER_MASTER 
                            ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code
                        LEFT OUTER JOIN TSPL_ITEM_MASTER 
                            ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SHIPMENT_DETAIL.item_code
	                        LEFT JOIN TSPL_ITEM_UOM_DETAIL AS ItemConvinUOM 
                            ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemConvinUOM.Item_Code AND TSPL_SD_SHIPMENT_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
                        LEFT JOIN TSPL_COMPANY_MASTER ON 2 = 2 "
            If rbtnSupplyDate.IsChecked Then
                Qry += "  WHERE convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " & whrcls & ""
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += "  WHERE convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " & whrcls & ""
            End If
            Qry += "  AND TSPL_CUSTOMER_MASTER.Credit_Customer = 'Y'
                        GROUP BY 
                            TSPL_CUSTOMER_MASTER.Customer_Name order by TSPL_CUSTOMER_MASTER.Customer_Name"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptCreditPartyWiseSaleAmount", "Credit Party Wise Sale Amount")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub



    Sub BillWisesaleSummaryDispatch(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            Qry = " SELECT '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                                    MAX(Item_Desc) AS Item_Desc, 
                        Item_Code, MAX(Unit_code) AS Unit_code,SUM(Def_UOM_Qty) AS Qty,sum(Amount) as Amount,SUM(QtyAccToReportUOM) AS QtyAccToReportUOM,MAX(UOM_Code) AS UOM_Code,MAX(Comp_Name) AS Comp_Name,MAX(Add1) AS Add1,MAX(Add2) AS Add2,MAX(Add3) AS Add3,MAX(City_Code) AS City_Code,MAX(State) AS State, MAX(Document_Date) AS Document_Date 
                    FROM (
                        SELECT 
	                    TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,
  
                            TSPL_ITEM_MASTER.Item_Desc, 
                            TSPL_ITEM_MASTER.Item_Code, 
                            TSPL_SD_SHIPMENT_DETAIL.Unit_code, 
                            TSPL_SD_SHIPMENT_DETAIL.Qty,
                    cast(
                                    (
                                      TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
                                    ) as Decimal(18, 2)
                                  ) as QtyAccToReportUOM, cast(
                                    (
                                      TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / DefUOM.Conversion_Factor
                                    ) as Decimal(18, 2)
                                  ) as Def_UOM_Qty,
                            ItemConvReportUOM.UOM_Code, 
                            TSPL_COMPANY_MASTER.Comp_Name, 
                            TSPL_COMPANY_MASTER.Add1, 
                            TSPL_COMPANY_MASTER.Add2, 
                            TSPL_COMPANY_MASTER.Add3, 
                            TSPL_COMPANY_MASTER.City_Code, 
                            TSPL_COMPANY_MASTER.State,
		                    TSPL_SD_SHIPMENT_DETAIL.Amount,
                            Document_Date 
                        FROM TSPL_SD_SHIPMENT_DETAIL
                        LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
                        LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SHIPMENT_DETAIL.item_code
                     left join TSPL_ITEM_UOM_DETAIL as ItemConvReportUOM on TSPL_ITEM_master.Item_Code = ItemConvReportUOM.Item_Code 
                                        and ItemConvReportUOM.Report_UOM = 1
LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE DEFAULT_UOM = 1 ) DefUOM ON TSPL_SD_SHIPMENT_DETAIL.Item_Code = DefUOM.item_code
                                         left join TSPL_ITEM_UOM_DETAIL as ItemConvinUOM on TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemConvinUOM.Item_Code 
                                       and TSPL_SD_SHIPMENT_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
                        LEFT JOIN TSPL_COMPANY_MASTER ON 2 = 2 where 2=2 "

            Dim strDate As String = ""
            If rbtnDocumentDate.IsChecked Then
                strDate = "Document_Date"
            ElseIf rbtnSupplyDate.IsChecked Then
                strDate = "Supply_Date"
            End If
            Qry += " and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"

            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD." + strDate + " as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "'  and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
            End If

            Qry += "  and  TSPL_ITEM_MASTER.IsTaxable=0 ) xx "

            Qry +=" GROUP BY Item_Code order by Item_Desc"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptBillWiseSaleOfMilkSummary", "Bill Wise Sale Of Milk Summary")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub BillWisesaleSummaryInvoice(ByVal print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whr As String = ""
            Qry = " SELECT '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                                    MAX(Item_Desc) AS Item_Desc, 
                        Item_Code, MAX(Unit_code) AS Unit_code,SUM(Def_UOM_Qty) AS Qty,sum(Amount) as Amount,SUM(QtyAccToReportUOM) AS QtyAccToReportUOM,MAX(UOM_Code) AS UOM_Code,MAX(Comp_Name) AS Comp_Name,MAX(Add1) AS Add1,MAX(Add2) AS Add2,MAX(Add3) AS Add3,MAX(City_Code) AS City_Code,MAX(State) AS State, MAX(Document_Date) AS Document_Date 
                    FROM (
                        SELECT 
	                    TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE,
  
                            TSPL_ITEM_MASTER.Item_Desc, 
                            TSPL_ITEM_MASTER.Item_Code, 
                            TSPL_SD_SALE_INVOICE_DETAIL.Unit_code, 
                            TSPL_SD_SALE_INVOICE_DETAIL.Qty,
                    cast(
                                    (
                                      TSPL_SD_SALE_INVOICE_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
                                    ) as Decimal(18, 2)
                                  ) as QtyAccToReportUOM,cast(
                                    (
                                      TSPL_SD_SALE_INVOICE_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / DefUOM.Conversion_Factor
                                    ) as Decimal(18, 2)
                                  ) as Def_UOM_Qty,
                            ItemConvReportUOM.UOM_Code, 
                            TSPL_COMPANY_MASTER.Comp_Name, 
                            TSPL_COMPANY_MASTER.Add1, 
                            TSPL_COMPANY_MASTER.Add2, 
                            TSPL_COMPANY_MASTER.Add3, 
                            TSPL_COMPANY_MASTER.City_Code, 
                            TSPL_COMPANY_MASTER.State,
		                    TSPL_SD_SALE_INVOICE_DETAIL.Amount,
                            TSPL_SD_SALE_INVOICE_HEAD.Document_Date 
                        FROM TSPL_SD_SALE_INVOICE_DETAIL
                        LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE
                        LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SALE_INVOICE_DETAIL.item_code
left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
                     left join TSPL_ITEM_UOM_DETAIL as ItemConvReportUOM on TSPL_ITEM_master.Item_Code = ItemConvReportUOM.Item_Code 
                                        and ItemConvReportUOM.Report_UOM = 1
LEFT JOIN ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL WHERE DEFAULT_UOM = 1 ) DefUOM ON TSPL_SD_SALE_INVOICE_DETAIL.Item_Code = DefUOM.item_code
                                         left join TSPL_ITEM_UOM_DETAIL as ItemConvinUOM on TSPL_SD_SALE_INVOICE_DETAIL.Item_Code = ItemConvinUOM.Item_Code 
                                       and TSPL_SD_SALE_INVOICE_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
                        LEFT JOIN TSPL_COMPANY_MASTER ON 2 = 2 where 2=2 "

            If rbtnSupplyDate.IsChecked Then
                Qry += " and Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " and Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(clsCommon.myCDate(txtFromDate1.Value)), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(clsCommon.myCDate(txtToDate1.Value)), "dd/MMM/yyyy") + "'"
            End If
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                If rbtnSupplyDate.IsChecked Then
                    Qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
                ElseIf rbtnDocumentDate.IsChecked Then
                    Qry += " and 2=( case when Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) >= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) <= '" + clsCommon.GetPrintDate(txtFromDate1.Value, "dd/MMM/yyyy") + "' and  TSPL_SD_SHIPMENT_HEAD.shift_type='AM' then 3 else 2 end  )"
                End If
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                If rbtnSupplyDate.IsChecked Then
                    Qry += " and 2=( case when Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SHIPMENT_HEAD.Supply_Date as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "'  and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
                ElseIf rbtnDocumentDate.IsChecked Then
                    Qry += " and 2=( case when Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate1.Value), "dd/MMM/yyyy") + "' and Cast(TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date) <= '" + clsCommon.GetPrintDate(txtToDate1.Value, "dd/MMM/yyyy") + "'  and TSPL_SD_SHIPMENT_HEAD.shift_type='PM' then 3 else 2 end  )"
                End If
            End If


            Qry += " and  TSPL_ITEM_MASTER.IsTaxable=0 ) xx Group BY Item_Code order by Item_Desc"
            dt = clsDBFuncationality.GetDataTable(Qry)

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptBillWiseSaleOfMilkSummary", "Bill Wise Sale Of Milk Summary")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub ProductWiseSale(ByVal Print As Boolean)
        Try
            Dim Qry As String = ""
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            Dim whrcls As String = ""
            Dim whrcls2 As String = ""
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " And TSPL_SD_SHIPMENT_head.Customer_Code in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If

            If txtItem.arrValueMember IsNot Nothing Then
                whrcls += " and TSPL_ITEM_MASTER.Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls2 = " and tspl_transfer_order_head.To_Location in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If

            If txtItem.arrValueMember IsNot Nothing Then
                whrcls2 += " and TSPL_ITEM_MASTER.Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If
            Qry = " SELECT '" + clsCommon.GetPrintDate(txtFromDate.Value) + "' as Fromdate,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' as ToDate,
                MAX(Item_Desc) AS Item_Desc, 
                    Item_Code, MAX(Unit_code) AS Unit_code,SUM(Qty) AS Qty,SUM(QtyAccToReportUOM) AS QtyAccToReportUOM,MAX(UOM_Code) AS UOM_Code,MAX(Comp_Name) AS Comp_Name,MAX(Add1) AS Add1,MAX(Add2) AS Add2,MAX(Add3) AS Add3,MAX(City_Code) AS City_Code,MAX(State) AS State, MAX(Document_Date) AS Document_Date 
                FROM (
                    SELECT 
	                TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,
  
                        TSPL_ITEM_MASTER.Item_Desc, 
                        TSPL_ITEM_MASTER.Item_Code, 
                        ItemBulkUOM.UOM_Code as Unit_code, 
                        cast(
                                (
                                  TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemBulkUOM.Conversion_Factor
                                ) as Decimal(18, 2)
                              ) as Qty,
                cast(
                                (
                                  TSPL_SD_SHIPMENT_DETAIL.Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
                                ) as Decimal(18, 2)
                              ) as QtyAccToReportUOM,
                        ItemConvReportUOM.UOM_Code, Document_Date 
                    FROM TSPL_SD_SHIPMENT_DETAIL
                    LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE
                    LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.item_code = TSPL_SD_SHIPMENT_DETAIL.item_code
                 left join TSPL_ITEM_UOM_DETAIL as ItemConvReportUOM on TSPL_ITEM_master.Item_Code = ItemConvReportUOM.Item_Code 
                                    and ItemConvReportUOM.Report_UOM = 1
                                     left join TSPL_ITEM_UOM_DETAIL as ItemConvinUOM on TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemConvinUOM.Item_Code 
                                   and TSPL_SD_SHIPMENT_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
left join TSPL_ITEM_UOM_DETAIL as ItemBulkUOM on TSPL_SD_SHIPMENT_DETAIL.Item_Code = ItemBulkUOM.Item_Code 
                                  and ItemBulkUOM.Bulk_UOM = 1 "
            If rbtnSupplyDate.IsChecked Then
                Qry += " where convert(date,Supply_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Supply_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " + whrcls + ""
            ElseIf rbtnDocumentDate.IsChecked Then
                Qry += " where convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " + whrcls + ""
            End If
            'Qry += " where convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " + whrcls + ""
            '-----Add Transfer Out Type Data
            If rbtnDocumentDate.IsChecked Then
                Qry += " union all
					select tspl_transfer_order_head.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Desc,TSPL_TRANSFER_ORDER_DETAIL.Item_Code,ItemBulkUOM.UOM_Code as Unit_code,cast((TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * ItemConvinUOM.Conversion_Factor / ItemBulkUOM.Conversion_Factor) as Decimal(18, 2)) as Qty,
					cast(
                                (
                                  TSPL_TRANSFER_ORDER_DETAIL.Out_Qty * ItemConvinUOM.Conversion_Factor / ItemConvReportUOM.Conversion_Factor
                                ) as Decimal(18, 2)
                              ) as QtyAccToReportUOM,
                        ItemConvReportUOM.UOM_Code,Document_Date 
					 from TSPL_TRANSFER_ORDER_DETAIL left join tspl_transfer_order_head on tspl_transfer_order_head.Document_No = TSPL_TRANSFER_ORDER_DETAIL.Document_No
					LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.item_code = TSPL_TRANSFER_ORDER_DETAIL.item_code
                 left join TSPL_ITEM_UOM_DETAIL as ItemConvReportUOM on TSPL_ITEM_master.Item_Code = ItemConvReportUOM.Item_Code  and ItemConvReportUOM.Report_UOM = 1
                                     left join TSPL_ITEM_UOM_DETAIL as ItemConvinUOM on TSPL_TRANSFER_ORDER_DETAIL.Item_Code = ItemConvinUOM.Item_Code 
                                   and TSPL_TRANSFER_ORDER_DETAIL.Unit_code = ItemConvinUOM.UOM_Code
 left join TSPL_ITEM_UOM_DETAIL as ItemBulkUOM on TSPL_TRANSFER_ORDER_DETAIL.Item_Code = ItemBulkUOM.Item_Code  and ItemBulkUOM.Bulk_UOM = 1 "

                Qry += " where convert(date,Document_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "' " + whrcls2 + " and tspl_transfer_order_head.Transfer_Type = 'O' "


            End If

            Qry += " ) xx  LEFT JOIN TSPL_COMPANY_MASTER ON 2 = 2 GROUP BY Item_Code "
            dt = clsDBFuncationality.GetDataTable(" with cte as ( " & Qry & " ) select * from cte order by Item_Desc ")

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If Print = False Then
                    gv1.DataSource = Nothing
                    gv1.Rows.Clear()
                    gv1.Columns.Clear()
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterView.Refresh()
                    gv1.GroupDescriptors.Clear()
                    gv1.EnableFiltering = True
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.BestFitColumns()
                    SetGridFormationn()
                    'ReStoreGridLayout()
                    gv1.MasterTemplate.AutoExpandGroups = True
                    'EnableDisableControls(False)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.BestFitColumns()
                ElseIf Print = True Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductWiseSaleQuantity", "Product Wise Sale Quantity")
                    frmCRV = Nothing

                End If
            Else
                clsCommon.MyMessageBoxShow("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub SetGridFormationn()
        gv1.TableElement.TableHeaderHeight = 40
        gv1.MasterTemplate.ShowRowHeaderColumn = True
        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).IsVisible = True
            gv1.Columns(ii).Width = 80
            If rbtnStockStatement.IsChecked Then
                If clsCommon.CompairString(gv1.Columns(ii).Name, "SNO") <> CompairStringResult.Equal Then
                    gv1.Columns(ii).FormatString = "{0:n2}"
                End If
            Else
                gv1.Columns(ii).FormatString = "{0:n2}"
            End If

        Next
        If rbtnDistributorCollStatement.IsChecked Then
            If isPrint Then
                gv1.Columns("CompName").IsVisible = False
                gv1.Columns("Add1").IsVisible = False
                gv1.Columns("Fromdate").IsVisible = False
                gv1.Columns("ToDate").IsVisible = False
                gv1.Columns("QtyConvType").IsVisible = False
            End If

            gv1.Columns("Customer_Name").HeaderText = "Customer"
            gv1.Columns("Item_Desc").HeaderText = "Item"
            gv1.Columns("Total_Qty").HeaderText = "Total Qty"
            gv1.Columns("Total_Qty").FormatString = "{0:n2}"
            gv1.Columns("Rate").FormatString = "{0:n2}"
            gv1.Columns("Amount").FormatString = "{0:n2}"
            gv1.Columns("Qty(M)").FormatString = "{0:n2}"
            gv1.Columns("Qty(E)").FormatString = "{0:n2}"
            gv1.ShowGroupPanel = False
            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim QtyM As New GridViewSummaryItem("Qty(M)", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(QtyM)
            Dim QtyE As New GridViewSummaryItem("Qty(E)", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(QtyE)
            Dim Total_Qty As New GridViewSummaryItem("Total_Qty", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(Total_Qty)
            Dim Amount As New GridViewSummaryItem("Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(Amount)
            gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        ElseIf rbtnStockStatement.IsChecked Then
            gv1.Columns("Item_Desc").HeaderText = "Item"
            gv1.Columns("Report_UOM").HeaderText = "UNIT"
            gv1.Columns("OPBal").HeaderText = "Opening Stock"
            gv1.Columns("Production_In_Qty").HeaderText = "Production "
            gv1.Columns("Other_In_Qty").HeaderText = "Other"
            gv1.Columns("Production_Out_Qty").HeaderText = "Production "
            gv1.Columns("Other_Out_Qty").HeaderText = "Other"
            gv1.Columns("Sale_Qty").HeaderText = "Sale "
            gv1.Columns("STC_Qty").HeaderText = "STC"
            gv1.Columns("Closing_Qty").HeaderText = "Closing Stock"
            gv1.Columns("Inter_Union_Sale").HeaderText = "Inter Union Sale"
            If isPrint Then
                gv1.Columns("CompName").IsVisible = False
                gv1.Columns("Add1").IsVisible = False
                gv1.Columns("Fromdate").IsVisible = False
                gv1.Columns("ToDate").IsVisible = False
            End If
            Dim index As Integer = 0
            If isPrint Then
                index = 8
            Else
                index = 3
            End If
            Dim summaryRowItem As New GridViewSummaryRowItem()
            For ii As Integer = index To gv1.Columns.Count - 1
                summaryRowItem.Add(New GridViewSummaryItem(gv1.Columns(ii).Name, "{0:F2}", GridAggregateFunction.Sum))
            Next
            gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        ElseIf rbtnBoothSaleItemWise.IsChecked Then
            gv1.Columns("SNo").FormatString = ""
            gv1.Columns("Comp_Name").IsVisible = False
            gv1.Columns("Logo_Img").IsVisible = False
            gv1.Columns("FromShift").IsVisible = False
            gv1.Columns("ToShift").IsVisible = False
            gv1.Columns("Add1").IsVisible = False
            gv1.Columns("Fromdate").IsVisible = False
            gv1.Columns("ToDate").IsVisible = False
            gv1.Columns("Customer_Code").IsVisible = False
            gv1.Columns("Item_Code").IsVisible = False
            gv1.Columns("Item_Code").HeaderText = "Item Code"
            gv1.Columns("Customer_Code").HeaderText = "Customer Code"
            gv1.Columns("Booth").HeaderText = "Customer"
            gv1.Columns("Short_Description").HeaderText = "Item"
            gv1.Columns("Qty").HeaderText = "Quantity"
            gv1.Columns("Def_Rep_UOM").HeaderText = "Unit"
            Dim Descriptor1 As New GroupDescriptor()
            Descriptor1.GroupNames.Add("Booth", System.ComponentModel.ListSortDirection.Ascending)
            gv1.GroupDescriptors.Add(Descriptor1)
            Dim summaryRowItem As New GridViewSummaryRowItem()
            summaryRowItem.Add(New GridViewSummaryItem("Qty", "{0:F2}", GridAggregateFunction.Sum))
            gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        End If
    End Sub

    Sub View()
        If gv1.Rows.Count > 0 Then
            Dim view As New ColumnGroupsViewDefinition()

            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Sno").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Item_Desc").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("Report_UOM").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns("OPBal").Name)

            view.ColumnGroups.Add(New GridViewColumnGroup("Received From"))
            view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())

            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Production_In_Qty").Name)
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv1.Columns("Other_In_Qty").Name)
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())

            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("Sale_Qty").Name)
            view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv1.Columns("STC_Qty").Name)


            view.ColumnGroups.Add(New GridViewColumnGroup("Transfer back to"))
            view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())

            view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv1.Columns("Production_Out_Qty").Name)
            view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv1.Columns("Other_Out_Qty").Name)

            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())

            view.ColumnGroups(4).Rows(0).ColumnNames.Add(gv1.Columns("Closing_Qty").Name)
            view.ColumnGroups(4).Rows(0).ColumnNames.Add(gv1.Columns("Inter_Union_Sale").Name)
            gv1.ViewDefinition = view
        End If
    End Sub
    Sub funreset()
        Dim viewBlank As New TableViewDefinition()
        gv1.ViewDefinition = viewBlank
        RadGroupBox4.Visible = False
        BtnStcRegisterPartyandItemWiseSummary.IsChecked = False
        BtnStcRegisterItemWiseSummary.IsChecked = False
        BtnProductWiseSaleQuantity.IsChecked = False
        BtnMilkStcSummary.IsChecked = False
        BtnPartySaleMilkProduct.IsChecked = False
        BtnProductSalesSummary.IsChecked = False
        BtnBillWiseSaleOfMilk.IsChecked = False
        BtnBillWiseSaleOfMilkSummary.IsChecked = False
        BtnTransportationCharges.IsChecked = False
        BtnTcsSummary.IsChecked = False
        BtnGheeReport.IsChecked = False
        BtnRouteWiseSale.IsChecked = False
        BtnCreditPartyWiseSaleAmount.IsChecked = False
        EnableDisableControls(True)
        gv1.DataSource = Nothing
        rbtnDocumentDate.IsChecked = True
        rbtnDetail.IsChecked = True
        txtLocation.arrValueMember = Nothing
        ddlQtyConversionType.SelectedValue = ""
        txtRoute.arrValueMember = Nothing
        rbtnDistributorCollStatement.IsChecked = False
        rbtnMilkSale.IsChecked = False
        rbtnStockStatement.IsChecked = False
        rbtnBoothSaleDateShiftWise.IsChecked = False
        rbtnBoothSaleItemWise.IsChecked = False
        lblRoute.Visible = False
        txtRoute.Visible = False
        lblLocation.Visible = False
        txtLocation.Visible = False
        lblToLocation.Visible = False
        txtToLocation.Visible = False
        lblQtyConv.Visible = False
        ddlQtyConversionType.Visible = False
        ddlDefaultReportUOM.Visible = False
        RadPageView1.SelectedPage = RadPageViewPage1
        btnGo.Enabled = True
        RadSplitButton1.Enabled = True
    End Sub

    Private Sub EnableDisableControls(ByVal val As Boolean)
        RadGroupBox1.Enabled = val
    End Sub

    Private Sub LoadData()
        Try

            Dim qry As String = ""
            Dim groupby As String = ""
            Dim whrcls As String = ""
            Dim BaseQuery As String = ""
            Dim qry2 As String = ""
            If txtCustomer.arrValueMember IsNot Nothing Then
                whrcls = " and Cust_Code in (''," & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ") "
            End If

            If txtItem.arrValueMember IsNot Nothing Then
                whrcls += " and Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
            End If

            If txtTransaction.arrValueMember IsNot Nothing Then
                whrcls += " and Trans_Name in (" & clsCommon.GetMulcallString(txtTransaction.arrValueMember) & ") "
            End If

            qry = "---------------- VCGL---------------------------- 
            select 'VCGL' as Trans_Name,TSPL_VCGL_Head.Document_Date,'' as Cust_Code, TSPL_VCGL_Detail.Document_No,'' as Item_Code ,0 as Qty ,'' as UOM,case when TSPL_VCGL_Detail.Row_Type ='Customer' then TSPL_VCGL_Detail.Dr_Amount-TSPL_VCGL_Detail.Cr_Amount else 0 end as Item_Net_Amt ,0 as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,0 as Item_Rate from TSPL_VCGL_Detail left outer join TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No  =TSPL_VCGL_Detail.Document_No  where TSPL_VCGL_Detail.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_VCGL  from TSPL_Customer_Invoice_Head
            left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_VCGL ,'')<>''     and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_VCGL_Head.Status =1  
            union all 
            select 'VCGL' as Trans_Name,TSPL_VCGL_Head.Document_Date,'' as Cust_Code, TSPL_VCGL_Head.Document_No,'' as Item_Code ,0 as Qty ,'' as UOM,case when TSPL_VCGL_Head.Document_Type  ='C' then TSPL_VCGL_Head.Tot_Cr_Amount-TSPL_VCGL_Head.Tot_Dr_Amount else 0 end as Item_Net_Amt ,0 as Total_Tax_Amt, '' as Tax, 0 as Tax_Amt ,0 as Tax_Rate,0 as Item_Rate from TSPL_VCGL_Detail left outer join TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No  =TSPL_VCGL_Detail.Document_No  
            where TSPL_VCGL_Head.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_VCGL  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_VCGL ,'')<>''     and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_VCGL_Head.Status =1  
            Union all 
            ---------------- Sale Invoice ,BULK SALE,CAN SALE--------------------- 
            select 'DS'  as Trans_Name," & IIf(rbtnSupplyDate.IsChecked, " TSPL_SD_SHIPMENT_HEAD.Supply_Date as Document_Date", " TSPL_SD_SALE_INVOICE_HEAD.Document_Date") & " ,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code as Cust_Code, TSPL_SD_SALE_INVOICE_DETAIL.Document_Code as Document_No,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code ,TSPL_SD_SALE_INVOICE_DETAIL.Qty ,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as UOM, (TSPL_SD_SALE_INVOICE_DETAIL.Amount+TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt-Total_Disc_Amt)*(case when coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0)<=0  then 1 else coalesce(TSPL_SD_SALE_INVOICE_HEAD.convrate,0) end)*(case when len(TSPL_SD_SALE_INVOICE_HEAD.Invoice_No_For_Supplementary)>0 and TSPL_SD_SALE_INVOICE_HEAD.Supplementary_Type='C' then -1 else 1 end) as Item_Net_Amt ,TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2), TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost) as  Item_Rate   from TSPL_SD_SALE_INVOICE_DETAIL  left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.Document_Code left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
            left outer join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER.Location_Code  where TSPL_SD_SALE_INVOICE_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_Customer_INVOICE_HEAD.Against_Sale_No  left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No   left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and  convert(date, " & IIf(rbtnSupplyDate.IsChecked, "TSPL_SD_SHIPMENT_HEAD.Supply_Date", "TSPL_Customer_INVOICE_HEAD.Document_Date") & " ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date, " & IIf(rbtnSupplyDate.IsChecked, "TSPL_SD_SHIPMENT_HEAD.Supply_Date", "TSPL_Customer_INVOICE_HEAD.Document_Date") & ",103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT'))) and TSPL_SD_SALE_INVOICE_HEAD.Status =1   
            union all 
            select 'CAN-SALE' as Trans_Name,TSPL_CANSALE_INVOICE_HEAD.Document_Date,TSPL_CANSALE_INVOICE_HEAD.Customer_Code as Cust_Code, TSPL_CANSALE_INVOICE_detail.Document_No,TSPL_CANSALE_INVOICE_detail.ItemCode ,TSPL_CANSALE_INVOICE_detail.Qty ,TSPL_CANSALE_INVOICE_detail.UOM ,TSPL_CANSALE_INVOICE_detail.Item_Net_Amt ,TSPL_CANSALE_INVOICE_detail.Total_Tax_Amt,'' as Tax,  0 as Tax_Amt ,0 as Tax_Rate, convert(decimal(18,2),TSPL_CANSALE_INVOICE_detail.PriceRate) as Item_Rate  from TSPL_CANSALE_INVOICE_detail left outer join TSPL_CANSALE_INVOICE_HEAD on TSPL_CANSALE_INVOICE_HEAD .Document_no =TSPL_CANSALE_INVOICE_detail.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_CANSALE_INVOICE_HEAD.Document_No 
            left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_CANSALE_INVOICE_HEAD.Customer_Code where TSPL_CANSALE_INVOICE_HEAD.Posted =1 and TSPL_CANSALE_INVOICE_HEAD.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all 
            select  'INVOICE-BS' as Trans_Name,TSPL_INVOICE_MASTER_BULKSALE.Document_Date,TSPL_INVOICE_MASTER_BULKSALE.Customer_Code as Cust_Code,TSPL_INVOICE_DETAIL_BULKSALE.Document_No,TSPL_INVOICE_DETAIL_BULKSALE.item_code,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceQty as Qty,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as UOM,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount ,TSPL_INVOICE_DETAIL_BULKSALE.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,0 as Item_Rate from TSPL_INVOICE_DETAIL_BULKSALE left outer join TSPL_INVOICE_MASTER_BULKSALE on TSPL_INVOICE_MASTER_BULKSALE .Document_no =TSPL_INVOICE_DETAIL_BULKSALE.Document_No  left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_INVOICE_MASTER_BULKSALE.Document_No 
            where TSPL_INVOICE_MASTER_BULKSALE.Posted =1 and TSPL_INVOICE_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all  
            ---------------- Sale RETURN ,CAN SALE RETURN----------------------- 
            select 'DSR' AS Trans_Name,TSPL_SD_SALE_RETURN_HEAD.Document_Date,TSPL_SD_SALE_RETURN_HEAD.Customer_Code as Cust_Code, TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,-(TSPL_SD_SALE_RETURN_DETAIL.Qty)Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,-(TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2),TSPL_SD_SALE_RETURN_DETAIL.Item_Cost)  as Item_Rate from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code  
            where TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT'))  and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 
            union all 
            select 'BULK-SALE-RE' AS Trans_Name, TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code as Cust_Code, TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No,TSPL_SALE_RETURN_DETAIL_BULKSALE.item_code,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceQty as Qty,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code as UOM,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount ,TSPL_SALE_RETURN_DETAIL_BULKSALE.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,0 as Item_Rate from TSPL_SALE_RETURN_DETAIL_BULKSALE left outer join TSPL_SALE_RETURN_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE .Document_no =TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_Return_No =TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No where TSPL_SALE_RETURN_MASTER_BULKSALE.Posted =1 AND TSPL_SALE_RETURN_MASTER_BULKSALE.AGAINST='Bulk Invoice'  and TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all 
            ---------------- Scrap Sale Return---------------------- 
            select 'SCRAP-S-R' as Trans_Name,TSPL_SCRAPSALE_HEAD_RETURN.Return_ship_Date as Document_Date,TSPL_SCRAPSALE_HEAD_RETURN.cust_Code as Cust_Code, TSPL_SCRAPSALE_DETAIL_RETURN.Document_No,TSPL_SCRAPSALE_DETAIL_RETURN.Item_Code ,TSPL_SCRAPSALE_DETAIL_RETURN.shipped_Qty AS Qty ,TSPL_SCRAPSALE_DETAIL_RETURN.Unit_code as UOM,TSPL_SCRAPSALE_DETAIL_RETURN.TotalAmt as Item_Net_Amt ,TSPL_SCRAPSALE_DETAIL_RETURN.TotalTaxAmt as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2),TSPL_SCRAPSALE_DETAIL_RETURN.price)  as Item_Rate from TSPL_SCRAPSALE_DETAIL_RETURN left outer join TSPL_SCRAPSALE_HEAD_RETURN on TSPL_SCRAPSALE_HEAD_RETURN .Document_No =TSPL_SCRAPSALE_DETAIL_RETURN.Document_No  where TSPL_SCRAPSALE_DETAIL_RETURN.Document_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPSALE_HEAD_RETURN.ispost =1  
            union all   
            ---------------- Scrap Sale ----------------- 
            select 'SCRAP-SALE' as Trans_Name,TSPL_SCRAPINVOICE_HEAD.shipment_Date as Document_Date,TSPL_SCRAPINVOICE_HEAD.cust_Code as Cust_Code,  TSPL_SCRAPINVOICE_DETAIL.invoice_No AS Document_No ,TSPL_SCRAPINVOICE_DETAIL.Item_Code ,TSPL_SCRAPINVOICE_DETAIL.shipped_Qty AS Qty ,TSPL_SCRAPINVOICE_DETAIL.Unit_code as UOM,TSPL_SCRAPINVOICE_DETAIL.TotalAmt as Item_Net_Amt ,TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2),TSPL_SCRAPINVOICE_DETAIL.price) as Item_Rate from TSPL_SCRAPINVOICE_DETAIL left outer join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD.invoice_No =TSPL_SCRAPINVOICE_DETAIL.invoice_No  where TSPL_SCRAPINVOICE_DETAIL.invoice_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrap  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrap ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPINVOICE_HEAD.ispost =1  
            union all 
            ---------------- mcc Material Sale RETURN-------------------------------- 
            select 'M-Material-R' as Trans_Name, TSPL_SD_SALE_RETURN_HEAD.Document_Date ,TSPL_SD_SALE_RETURN_HEAD.Customer_Code as Cust_Code, TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,-(TSPL_SD_SALE_RETURN_DETAIL.Qty)Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,-(TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2),TSPL_SD_SALE_RETURN_DETAIL.Item_Cost) as Item_Rate from TSPL_SD_SALE_RETURN_DETAIL  left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code  where TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 AND TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' 
            union all 
            ---------------- Security Receipt--------------------------
            select 'AR-INVOICE' as Trans_Name,TSPL_Customer_INVOICE_HEAD.Document_Date,TSPL_Customer_INVOICE_HEAD.Customer_Code as Cust_Code,  TSPL_Customer_Invoice_Head.Against_Security_Receipt_No as Document_No,'' AS Item_Code ,0 AS Qty ,'' as UOM,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  (isnull(TSPL_Customer_Invoice_Detail.Total_Amount ,0))as Item_Net_Amt ,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  (isnull(TSPL_Customer_Invoice_Detail.Total_Tax  ,0)) as Total_Tax_Amt,'' as Tax,  0 as Tax_Amt ,0 as Tax_Rate, 0 as Item_Rate from TSPL_Customer_Invoice_Detail left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Document_No =TSPL_Customer_Invoice_Detail.Document_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Security_Receipt_No ,'')<>''  and TSPL_Customer_INVOICE_HEAD.Status =1    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 
            union all 
             ---------------- Transfer --------------------------
            select  'STO-TRANSFER' as Trans_Name,TSPL_TRANSFER_ORDER_HEAD.Document_Date,To_Location as Cust_Code,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM,  TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2), TSPL_TRANSFER_ORDER_DETAIL.Item_Cost) as  Item_Rate from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)  and  TSPL_TRANSFER_ORDER_HEAD.status = 1  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'O'
            union all
            ---------------- Transfer Return--------------------------
            select  'STO-TRANS-R' as Trans_Name,TSPL_TRANSFER_ORDER_HEAD.Document_Date,Location as Cust_Code,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,-(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty) as Qty,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM,  -(TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate ,convert(decimal(18,2), TSPL_TRANSFER_ORDER_DETAIL.Item_Cost) as  Item_Rate from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)  and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'T'"

            qry += " " & Environment.NewLine & " Union all " & Environment.NewLine & ""
            For ii As Integer = 1 To 6
                If ii > 1 Then
                    BaseQuery += " Union all " & Environment.NewLine & ""
                End If

                BaseQuery += "---------------- Sale Invoice ,BULK SALE,CAN SALE--------------------- 
            select 'DS'  as Trans_Name," & IIf(rbtnSupplyDate.IsChecked, " TSPL_SD_SHIPMENT_HEAD.Supply_Date as Document_Date", " TSPL_SD_SALE_INVOICE_HEAD.Document_Date") & " ,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code as Cust_Code, TSPL_SD_SALE_INVOICE_DETAIL.Document_Code as Document_No,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code ,0 as Qty ,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt,   isnull(TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & ",'') AS Tax,  TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & "_Amt as TAX_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & "_Rate as TAX_Rate ,0  as Item_Rate   from TSPL_SD_SALE_INVOICE_DETAIL  left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.Document_Code  left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
            left outer join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER .Location_Code  where TSPL_SD_SALE_INVOICE_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_Customer_INVOICE_HEAD.Against_Sale_No left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date, " & IIf(rbtnSupplyDate.IsChecked, "TSPL_SD_SHIPMENT_HEAD.Supply_Date", "TSPL_Customer_INVOICE_HEAD.Document_Date") & "  ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date, " & IIf(rbtnSupplyDate.IsChecked, "TSPL_SD_SHIPMENT_HEAD.Supply_Date", "TSPL_Customer_INVOICE_HEAD.Document_Date") & ",103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT'))  and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_INVOICE_HEAD.Status =1  and TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & "_Amt > 0   
            union all 
            
            select 'CAN-SALE' as Trans_Name,TSPL_CANSALE_INVOICE_HEAD.Document_Date,TSPL_CANSALE_INVOICE_HEAD.Customer_Code as Cust_Code,  TSPL_CANSALE_INVOICE_detail.Document_No,TSPL_CANSALE_INVOICE_detail.ItemCode ,0 as Qty ,TSPL_CANSALE_INVOICE_detail.UOM ,0 as Item_Net_Amt,TSPL_CANSALE_INVOICE_detail.Total_Tax_Amt," & IIf(ii = 6, " ''  ", " isnull(TSPL_CANSALE_INVOICE_detail.TAX" & ii & ",'')") & "  AS Tax," & IIf(ii = 6, " 0  ", " TSPL_CANSALE_INVOICE_detail.TAX" & ii & "_Amt ") & " as TAX_Amt," & IIf(ii = 6, " 0  ", " TSPL_CANSALE_INVOICE_detail.TAX" & ii & "_Rate ") & "  as TAX_Rate ,0  as Item_Rate  from TSPL_CANSALE_INVOICE_detail left outer join TSPL_CANSALE_INVOICE_HEAD on TSPL_CANSALE_INVOICE_HEAD .Document_no =TSPL_CANSALE_INVOICE_detail.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_CANSALE_INVOICE_HEAD.Document_No 
            left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_CANSALE_INVOICE_HEAD.Customer_Code where TSPL_CANSALE_INVOICE_HEAD.Posted =1 and TSPL_CANSALE_INVOICE_HEAD.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and  " & IIf(ii < 6, "  TSPL_CANSALE_INVOICE_detail.TAX" & ii & "_Amt>0 ", " 2=2") & "
            union all 
            select 'INVOICE-BS' as Trans_Name,TSPL_INVOICE_MASTER_BULKSALE.Document_Date,TSPL_INVOICE_MASTER_BULKSALE.Customer_Code as Cust_Code, TSPL_INVOICE_DETAIL_BULKSALE.Document_No,TSPL_INVOICE_DETAIL_BULKSALE.item_code,0 as Qty,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_INVOICE_DETAIL_BULKSALE.Total_Tax_Amt," & IIf(ii = 6, " ''  ", " isnull(TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & ",'') ") & "  AS Tax, " & IIf(ii = 6, " 0  ", "  TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & "_Amt ") & " as TAX_Amt, " & IIf(ii = 6, " 0  ", "  TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & "_Rate ") & " as TAX_Rate,0 as Item_Rate from TSPL_INVOICE_DETAIL_BULKSALE left outer join TSPL_INVOICE_MASTER_BULKSALE on TSPL_INVOICE_MASTER_BULKSALE .Document_no =TSPL_INVOICE_DETAIL_BULKSALE.Document_No  left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_INVOICE_MASTER_BULKSALE.Document_No 
            where TSPL_INVOICE_MASTER_BULKSALE.Posted =1 and TSPL_INVOICE_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and
            " & IIf(ii < 6, " TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & "_Amt>0", " 2=2") & "
            union all  
            ---------------- Sale RETURN ,CAN SALE RETURN----------------------- 
            select 'DSR' AS Trans_Name,TSPL_SD_SALE_RETURN_HEAD.Document_Date,TSPL_SD_SALE_RETURN_HEAD.Customer_Code as Cust_Code, TSPL_SD_SALE_RETURN_DETAIL.Document_Code as Document_No,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,0 as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt, isnull(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & ",'') AS Tax,  -(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt) as Tax_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Rate as Tax_Rate,0 as Item_Rate from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code  
            where TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 and TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt>0
            union all 
            select 'BULK-SALE-RE' as Trans_Name, TSPL_SALE_RETURN_MASTER_BULKSALE.Document_Date,TSPL_SALE_RETURN_MASTER_BULKSALE.Customer_Code as Cust_Code, TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No,TSPL_SALE_RETURN_DETAIL_BULKSALE.item_code,0 as Qty,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SALE_RETURN_DETAIL_BULKSALE.Total_Tax_Amt," & IIf(ii = 6, " ''  ", "  isnull(TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & ",'') ") & " AS Tax, " & IIf(ii = 6, " 0  ", " TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & "_Amt") & "  as TAX_Amt, " & IIf(ii = 6, " 0  ", " TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & "_Rate ") & "  as TAX_Rate,0 as Item_Rate  from TSPL_SALE_RETURN_DETAIL_BULKSALE left outer join TSPL_SALE_RETURN_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE .Document_no =TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_Return_No =TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No where TSPL_SALE_RETURN_MASTER_BULKSALE.Posted =1 AND TSPL_SALE_RETURN_MASTER_BULKSALE.AGAINST='Bulk Invoice' and  " & IIf(ii < 6, "  TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & "_Amt>0 ", " 2=2") & " and TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all 
            ---------------- Scrap Sale Return---------------------- 
            select 'SCRAP-S-R' AS Trans_Name,TSPL_SCRAPSALE_HEAD_RETURN.Return_ship_Date as Document_Date,TSPL_SCRAPSALE_HEAD_RETURN.cust_Code as Cust_Code, TSPL_SCRAPSALE_DETAIL_RETURN.Document_No,TSPL_SCRAPSALE_DETAIL_RETURN.Item_Code ,0 as Qty ,TSPL_SCRAPSALE_DETAIL_RETURN.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SCRAPSALE_DETAIL_RETURN.TotalTaxAmt as Total_Tax_Amt, isnull(TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & ",'') as Tax, TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & "_Amt as Tax_Amt,TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & "_Rate as TAX_Rate,0 as Item_Rate from TSPL_SCRAPSALE_DETAIL_RETURN left outer join TSPL_SCRAPSALE_HEAD_RETURN on TSPL_SCRAPSALE_HEAD_RETURN .Document_No =TSPL_SCRAPSALE_DETAIL_RETURN.Document_No  where TSPL_SCRAPSALE_DETAIL_RETURN.Document_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPSALE_HEAD_RETURN.ispost =1   and TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & "_Amt>0
            union all  
            ---------------- Scrap Sale ----------------- 
            select 'SCRAP-SALE' AS Trans_Name,TSPL_SCRAPINVOICE_HEAD.shipment_Date as Document_Date,TSPL_SCRAPINVOICE_HEAD.cust_Code as Cust_Code, TSPL_SCRAPINVOICE_DETAIL.invoice_No AS Document_No ,TSPL_SCRAPINVOICE_DETAIL.Item_Code ,0 as Qty ,TSPL_SCRAPINVOICE_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt as Total_Tax_Amt, isnull(TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & ",'') as Tax, TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & "_Amt as TAX_Amt,TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & "_Rate as TAX_Rate,0 as Item_Rate from TSPL_SCRAPINVOICE_DETAIL left outer join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD.invoice_No =TSPL_SCRAPINVOICE_DETAIL.invoice_No  where TSPL_SCRAPINVOICE_DETAIL.invoice_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrap  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrap ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPINVOICE_HEAD.ispost =1  and TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & "_Amt>0
            union all 
            ---------------- mcc Material Sale RETURN-------------------------------- 
            select 'M-Material-R' as Trans_Name, TSPL_SD_SALE_RETURN_HEAD.Document_Date ,TSPL_SD_SALE_RETURN_HEAD.Customer_Code as Cust_Code, TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,0 as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,  isnull(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & ",'') as Tax,-(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt) as Tax_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Rate as TAX_Rate,0 as Item_Rate from TSPL_SD_SALE_RETURN_DETAIL  left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code  where TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 AND TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' and TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt > 0
            union all 
            ---------------- Security Receipt--------------------------
            select 'AR-INVOICE' as Trans_Name,TSPL_Customer_INVOICE_HEAD.Document_Date,TSPL_Customer_INVOICE_HEAD.Customer_Code as Cust_Code, TSPL_Customer_Invoice_Head.Against_Security_Receipt_No as Document_No,'' AS Item_Code ,0 AS Qty ,'' as UOM,0 as Item_Net_Amt ,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  (isnull(TSPL_Customer_Invoice_Detail.Total_Tax  ,0)) as Total_Tax_Amt,TSPL_Customer_Invoice_Detail.TAX" & ii & " as Tax,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  isnull(TSPL_Customer_Invoice_Detail.TAX" & ii & "_Amt,0) as Tax_Amt ,TSPL_Customer_Invoice_Detail.TAX" & ii & "_Rate as Tax_Rate,0 as Item_Rate from TSPL_Customer_Invoice_Detail left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Document_No =TSPL_Customer_Invoice_Detail.Document_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Security_Receipt_No ,'')<>''  and TSPL_Customer_INVOICE_HEAD.Status =1    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0   and TSPL_Customer_Invoice_Detail.TAX" & ii & "_Amt >0
            union all 
             ---------------- Transfer --------------------------
           select 'STO-TRANSFER' AS Trans_Name, TSPL_TRANSFER_ORDER_HEAD.Document_Date,Location as Cust_Code, TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,0 as Qty,'' as UOM,  0 as Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  isnull(TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & ",'') as Tax,  TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt as TAX_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Rate as TAX_Rate, 0 as Item_Rate  from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt>0  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'O'   
           union all 
                   ---------------- Transfer Return--------------------------
           select 'STO-TRANS-R' AS Trans_Name, TSPL_TRANSFER_ORDER_HEAD.Document_Date,Location as Cust_Code, TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,0 as Qty,'' as UOM,  0 as Item_Net_Amt ,-(TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt)Total_Tax_Amt,  isnull(TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & ",'') as Tax,  -(TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt) as TAX_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Rate as TAX_Rate, 0 as Item_Rate from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt>0  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'T' "
            Next

            BaseQuery = " select Trans_Name,Document_Date,Cust_Code, Document_No,xx.Item_Code,UOM,(Qty)Qty, (Item_Net_Amt)Item_Net_Amt,Tax,Tax_Rate,Tax_Amt, Item_Rate from ( " & qry & " " & BaseQuery & " )xx  where 2= 2 " & whrcls & " ) xxx  left outer join TSPL_TAX_MASTER ON  TSPL_TAX_MASTER.Tax_Code = XXX.TAX "
            Dim TaxDesc As String = ""
            Dim TaxAmount As String = ""
            qry = ""

            dtTax = clsDBFuncationality.GetDataTable("SELECT * FROM ( select Tax_Code as Tax,Tax_Code_Desc,type,case  when type = 'M' then 1 when type = 'K' then 2 when type = 'SGST' then 3 when type = 'CGST' then 4  when type = 'IGST' then 5 end as Sequence_No from TSPL_TAX_MASTER where Type in ('SGST','CGST','IGST','M','K') )X Order by Sequence_No ")

            For ii As Integer = 0 To dtTax.Rows.Count - 1
                If isPrint Then
                    TaxDesc += " TaxAmount_" + clsCommon.myCstr(ii + 1) + " ,Tax_" + clsCommon.myCstr(ii + 1) + ", "
                    qry += " ,'" & dtTax.Rows(ii)("Tax_Code_Desc") & "' as  Tax_" + clsCommon.myCstr(ii + 1) + " "
                    qry += " ,sum(Case When xxx.Type ='" & dtTax.Rows(ii)("Type") & "' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_" + clsCommon.myCstr(ii + 1) + " "
                    TaxAmount += ",max(Tax_" + clsCommon.myCstr(ii + 1) + ") as  Tax_" + clsCommon.myCstr(ii + 1) + " , sum(TaxAmount_" + clsCommon.myCstr(ii + 1) + ") as TaxAmount_" + clsCommon.myCstr(ii + 1) + ""
                Else
                    TaxDesc += "[" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount] ,"
                    qry += " ,sum(case when xxx.Tax ='" & dtTax.Rows(ii)("Tax") & "' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as [" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount]"
                    TaxAmount += " sum([" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount])[" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount] ,"
                End If
            Next
            BaseQuery = " select xxx.*,Type,Is_TCS from (  " & BaseQuery & "  )xxx "
            qry += ",Type "

            If isPrint Then
                qry = " select   Cust_Code,max(Customer_Name)Customer_Name,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' as FromDate,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' as ToDate,max(Comp_Name)Comp_Name,max(Add1) as Add1,Item_Code,max(Short_Description)Short_Description,sum(Total_Qty) As Total_Qty,max(UOM) AS UOM,max(Item_Rate)Item_Rate 
               " & TaxAmount & ",max(Tax_6) as  Tax_6 ,sum(TaxAmount_6)TaxAmount_6 ,sum(Item_Net_Amt)Item_Net_Amt,case when max(IsTaxable) = 1 then 'Product' else 'Milk' end as IsProduct  from (
                Select  TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,IsTaxable,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, FINAL.Item_Code,TSPL_ITEM_MASTER.Short_Description,I.UOM_Code AS UOM, isnull((Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Total_Qty,Item_Net_Amt,   isnull(Tax_Rate,0)Tax_Rate,(Item_Net_Amt - Total_Tax_Amt) AS Taxable_Amt, " & TaxDesc & "Tax_6,TaxAmount_6,Total_Tax_Amt,kkfAmt,MandiAmt,final.Item_Rate  from (  select  xxx.Cust_Code,xxx.Item_Code,UOM,SUM(Qty)Qty, sum(Item_Net_Amt)Item_Net_Amt, sum(Tax_Amt)Total_Tax_Amt,max(item_rate)item_rate,  max( case when (xxx.Tax ='IGST' and isnull(Tax_Rate ,0) > 0)  then isnull(Tax_Rate ,0) else (case when xxx.Tax ='CGST' then 2 * isnull(Tax_Rate ,0) end ) end) as Tax_Rate,sum(case when xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as kkfAmt, sum(case when xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as MandiAmt  " & qry & ",sum(case when Is_TCS = 'Y' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_6, 'TCS'  as Tax_6 from ( "
                groupby = " Cust_Code "
            Else
                If rbtnDetail.IsChecked Then
                    groupby = " Document_No,Document_Date,Cust_Code "
                    qry2 = " select convert(varchar, Document_Date,103) Document_Date,Document_No,Cust_Code,max(Customer_Name)Customer_Name,"
                ElseIf rbtnSummary.IsChecked Then
                    groupby = " Cust_Code "
                    qry2 = " select Cust_Code,max(Customer_Name)Customer_Name,"
                End If
                qry = " " & qry2 & " Item_Code,max(Short_Description)Short_Description,sum(Total_Qty) As Total_Qty,max(UOM) AS UOM,max(Item_Rate)Item_Rate," & TaxAmount & " sum([TCS Amount])[TCS Amount],sum(Item_Net_Amt)Item_Net_Amt
              from ( select Document_Date,Document_No, TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,FINAL.Item_Code,TSPL_ITEM_MASTER.Short_Description,I.UOM_Code AS UOM, isnull((Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Total_Qty,Item_Net_Amt,   isnull(Tax_Rate,0)Tax_Rate,(Item_Net_Amt - Total_Tax_Amt) AS Taxable_Amt, " & TaxDesc & "[TCS Amount], Total_Tax_Amt,kkfAmt,MandiAmt ,final.Item_Rate from ( select xxx.Item_Code,xxx.Document_Date,xxx.Document_No,xxx.Cust_Code,UOM,SUM(Qty)Qty, sum(Item_Net_Amt)Item_Net_Amt, sum(Tax_Amt)Total_Tax_Amt,max(Item_Rate)Item_Rate,  max( case when (xxx.Tax ='IGST' and isnull(Tax_Rate ,0) > 0)  then isnull(Tax_Rate ,0) else (case when xxx.Tax ='CGST' then 2 * isnull(Tax_Rate ,0) end ) end) as Tax_Rate,sum(case when xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as kkfAmt, sum(case when xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as MandiAmt  " & qry & ",sum(case when Is_TCS = 'Y' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as [TCS Amount] from ( "

            End If

            BaseQuery = "with CTERawData as ( " & qry & " " & BaseQuery & " GROUP BY xxx.Document_No,xxx.Document_Date,xxx.Cust_Code,xxx.Item_Code,UOM,Type ) FINAL left outer join tspl_item_master on TSPL_ITEM_MASTER.Item_Code = FINAL.Item_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = final.Cust_Code LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = FINAL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = FINAL.UOM	LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON FINAL.Item_Code = I.item_code left outer join TSPL_COMPANY_MASTER on 2 =2  ) xxxxFinal group by " & groupby & ",Item_Code ) select CTERawData.* from CTERawData order by  " & IIf(rbtnDetail.IsChecked AndAlso Not isPrint, "Document_Date,", "") & " Cust_Code"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(BaseQuery)

            gv1.DataSource = Nothing
            gv1.Rows.Clear()
            gv1.Columns.Clear()
            gv1.GroupDescriptors.Clear()
            gv1.MasterView.Refresh()
            gv1.GroupDescriptors.Clear()
            gv1.EnableFiltering = True
            gv1.MasterTemplate.SummaryRowsBottom.Clear()

            If dt.Rows.Count > 0 Then
                gv1.DataSource = dt
                gv1.BestFitColumns()
                SetGridFormation()
                ReStoreGridLayout()
                gv1.MasterTemplate.AutoExpandGroups = True
                EnableDisableControls(False)
                RadPageView1.SelectedPage = RadPageViewPage2
                gv1.BestFitColumns()

                If isPrint Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If BtnStcRegisterItemWiseSummary.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemWiseSummary", "STS Register Item Wise Summary")
                    ElseIf BtnBillWiseSaleOfMilkSummary.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptBillWiseSaleOfMilkSummary", "Bill Wise Sale Of Milk Summary")
                    ElseIf BtnProductWiseSaleQuantity.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductWiseSaleQuantity", "Product Wise Sale Quantity")
                    ElseIf BtnBillWiseSaleOfMilk.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptBillWiseSaleOfMilk", "Bill Wise Sale Of Milk")
                    ElseIf BtnPartySaleMilkProduct.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptPartySalesMilkandProducts", "Party Sales Milk and Products")
                    ElseIf BtnProductSalesSummary.IsChecked = True Then
                        If ddlType.SelectedValue = "Taxable" OrElse ddlType.SelectedValue = "Both" Then
                            frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductSaleSummaryTaxableNonTaxable", "Product Sale Summary Taxable NonTaxable")
                        ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                            frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptProductSaleSummaryNonTaxable", "Product Sale Summary NonTaxable")
                        End If
                    ElseIf BtnMilkStcSummary.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptMilkStcSummary", "Milk Stc Summary")
                    ElseIf BtnStcRegisterPartyandItemWiseSummary.IsChecked = True Then
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemWiseSummaryPartyWise", "STS Register-Item Wise Summary Party Wise")
                    Else
                        frmCRV.funreport(Report_ID, CrystalReportFolder.SalesReport, dt, "rptCustItemWiseSale", "Customer Item Wise Sale")
                    End If
                    frmCRV = Nothing
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub

            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)

        End Try
    End Sub
    'Sub reportSTCPartyRegister(ByVal Print As Boolean)
    '    Try
    '        Dim qry As String = "
    '            Select  TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,IsTaxable,TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name, FINAL.Item_Code,TSPL_ITEM_MASTER.Short_Description,I.UOM_Code + '-' + I.UOM_Description AS UOM, isnull((Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Total_Qty,Item_Net_Amt,   isnull(Tax_Rate,0)Tax_Rate,(Item_Net_Amt - Total_Tax_Amt) AS Taxable_Amt,  TaxAmount_1 ,Tax_1,  TaxAmount_2 ,Tax_2,  TaxAmount_3 ,Tax_3,  TaxAmount_4 ,Tax_4,  TaxAmount_5 ,Tax_5,Tax_6,TaxAmount_6,Total_Tax_Amt,kkfAmt,MandiAmt,final.Item_Rate  from (  select  xxx.Cust_Code,xxx.Item_Code,UOM,SUM(Qty)Qty, sum(Item_Net_Amt)Item_Net_Amt, sum(Tax_Amt)Total_Tax_Amt,max(item_rate)item_rate,  max( case when (xxx.Tax ='IGST' and isnull(Tax_Rate ,0) > 0)  then isnull(Tax_Rate ,0) else (case when xxx.Tax ='CGST' then 2 * isnull(Tax_Rate ,0) end ) end) as Tax_Rate,sum(case when xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as kkfAmt, sum(case when xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as MandiAmt   ,'MANDITAX' as  Tax_1  ,sum(Case When xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_1  ,'KKF' as  Tax_2  ,sum(Case When xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_2  ,'State Goods Service Tax' as  Tax_3  ,sum(Case When xxx.Type ='SGST' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_3  ,'Central Goods Serivce Tax' as  Tax_4  ,sum(Case When xxx.Type ='CGST' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_4  ,'Integrated Goods and Services Tax' as  Tax_5  ,sum(Case When xxx.Type ='IGST' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_5 ,Type ,sum(case when Is_TCS = 'Y' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_6, 'TCS'  as Tax_6 from (   select xxx.*,Type,Is_TCS from (   select Trans_Name,Document_Date,Cust_Code, Document_No,xx.Item_Code,UOM,(Qty)Qty, (Item_Net_Amt)Item_Net_Amt,Tax,Tax_Rate,Tax_Amt, Item_Rate from ( ---------------- VCGL---------------------------- 
    '        select 'VCGL' as Trans_Name,TSPL_VCGL_Head.Document_Date,'' as Cust_Code, TSPL_VCGL_Detail.Document_No,'' as Item_Code ,0 as Qty ,'' as UOM,case when TSPL_VCGL_Detail.Row_Type ='Customer' then TSPL_VCGL_Detail.Dr_Amount-TSPL_VCGL_Detail.Cr_Amount else 0 end as Item_Net_Amt ,0 as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,0 as Item_Rate from TSPL_VCGL_Detail left outer join TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No  =TSPL_VCGL_Detail.Document_No  where TSPL_VCGL_Detail.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_VCGL  from TSPL_Customer_Invoice_Head
    '        left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_VCGL ,'')<>''     and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'01/Nov/2023',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'20/Dec/2024',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_VCGL_Head.Status =1  
    '        union all 
    '         ---------------- Transfer --------------------------
    '        select  'STO-TRANSFER' as Trans_Name,TSPL_TRANSFER_ORDER_HEAD.Document_Date,Location as Cust_Code,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM,  TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,convert(decimal(18,2), TSPL_TRANSFER_ORDER_DETAIL.Item_Cost) as  Item_Rate from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'01/Nov/2023',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'20/Dec/2024',103)  and  TSPL_TRANSFER_ORDER_HEAD.status = 1  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'O'
    '        union all
    '        ---------------- Transfer Return--------------------------
    '        select  'STO-TRANS-R' as Trans_Name,TSPL_TRANSFER_ORDER_HEAD.Document_Date,Location as Cust_Code,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,-(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty) as Qty,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM,  -(TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate ,convert(decimal(18,2), TSPL_TRANSFER_ORDER_DETAIL.Item_Cost) as  Item_Rate from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'01/Nov/2023',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'20/Dec/2024',103)  and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'T' 
    '        )xx  where 2= 2  ) xxx  left outer join TSPL_TAX_MASTER ON  TSPL_TAX_MASTER.Tax_Code = XXX.TAX   )xxx  GROUP BY xxx.Document_No,xxx.Document_Date,xxx.Cust_Code,xxx.Item_Code,UOM,Type ) FINAL left outer join tspl_item_master on TSPL_ITEM_MASTER.Item_Code = FINAL.Item_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = final.Cust_Code LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = FINAL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = FINAL.UOM	LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON FINAL.Item_Code = I.item_code left outer join TSPL_COMPANY_MASTER on 2 =2   where tspl_item_master.Item_Type = 'F'  "

    '        If BtnStcRegisterPartyandItemWiseSummary.IsChecked = True Then
    '            qry = qry
    '        ElseIf BtnStcRegisterItemWiseSummary.IsChecked = True Then
    '            qry = qry + ""
    '        End If

    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

    '        If dt.Rows.Count > 0 Then
    '            If Print = False Then
    '                gv1.DataSource = Nothing
    '                gv1.Rows.Clear()
    '                gv1.Columns.Clear()
    '                gv1.GroupDescriptors.Clear()
    '                gv1.MasterView.Refresh()
    '                gv1.GroupDescriptors.Clear()
    '                gv1.EnableFiltering = True
    '                gv1.MasterTemplate.SummaryRowsBottom.Clear()
    '                gv1.DataSource = dt
    '                gv1.BestFitColumns()
    '                'SetGridFormation()
    '                ReStoreGridLayout()
    '                gv1.MasterTemplate.AutoExpandGroups = True
    '                EnableDisableControls(False)
    '                RadPageView1.SelectedPage = RadPageViewPage2
    '                gv1.BestFitColumns()
    '            Else
    '                Dim frmCRV As New frmCrystalReportViewer()
    '                If BtnStcRegisterPartyandItemWiseSummary.IsChecked = True Then
    '                    frmCRV.funreport(CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemWiseSummaryPartyWise", "STS Register-Item Wise Summary Party Wise")
    '                ElseIf BtnStcRegisterItemWiseSummary.IsChecked = True Then
    '                    frmCRV.funreport(CrystalReportFolder.SalesReport, dt, "CrptSTSRegister-ItemWiseSummary", "STS Register Item Wise Summary")
    '                End If
    '            End If
    '        Else
    '            clsCommon.MyMessageBoxShow(Me, "No Data Found.", Me.Text)
    '            Exit Sub
    '        End If


    '    Catch ex As Exception

    '    End Try
    'End Sub
    Sub SetGridFormation()
        gv1.TableElement.TableHeaderHeight = 40
        gv1.MasterTemplate.ShowRowHeaderColumn = True
        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).IsVisible = True
            gv1.Columns(ii).Width = 80
        Next

        gv1.ShowGroupPanel = False
        If Not isPrint Then
            If rbtnDetail.IsChecked Then
                gv1.Columns("Document_Date").HeaderText = "Document Date"
                gv1.Columns("Document_No").HeaderText = "Document No"
            End If
        End If

        gv1.Columns("Cust_Code").HeaderText = "Customer Code"
        gv1.Columns("Cust_Code").IsVisible = False
        gv1.Columns("Customer_Name").HeaderText = "Customer Name"

        gv1.Columns("Short_Description").HeaderText = "Item Name"
        gv1.Columns("Short_Description").Width = 110
        gv1.Columns("Total_Qty").HeaderText = "Quantity"
        gv1.Columns("Item_Net_Amt").FormatString = "{0:n2}"
        gv1.Columns("Item_Net_Amt").HeaderText = "Gross Amount"
        gv1.Columns("Item_Rate").HeaderText = "Item Rate"
        gv1.Columns("Item_Code").HeaderText = "Item Code"
        gv1.Columns("Item_Code").IsVisible = False
        gv1.Columns("Item_Rate").FormatString = "{0:n2}"

        If isPrint Then
            gv1.Columns("Tax_1").IsVisible = False
            gv1.Columns("Tax_2").IsVisible = False
            gv1.Columns("Tax_3").IsVisible = False
            gv1.Columns("Tax_4").IsVisible = False
            gv1.Columns("Tax_5").IsVisible = False
            gv1.Columns("Tax_6").IsVisible = False
            gv1.Columns("IsProduct").IsVisible = False
            gv1.Columns("Comp_Name").IsVisible = False
            gv1.Columns("FromDate").IsVisible = False
            gv1.Columns("ToDate").IsVisible = False
            For ii As Integer = 0 To dtTax.Rows.Count - 1
                gv1.Columns("TaxAmount_" & clsCommon.myCstr(ii + 1) & "").HeaderText = dtTax.Rows(ii)("Tax_Code_Desc")
            Next
            gv1.Columns("TaxAmount_6").HeaderText = "TCS Amount"
        End If

        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim j As Integer = 0
        If isPrint Then
            j = 7
        Else
            If rbtnDetail.IsChecked Then
                j = 6
            Else
                j = 4
            End If
        End If

        For ii As Integer = j To gv1.Columns.Count - 1
            If clsCommon.CompairString(gv1.Columns(ii).Name, "UOM") = CompairStringResult.Equal OrElse clsCommon.CompairString(gv1.Columns(ii).Name, "Item_Rate") = CompairStringResult.Equal Then
            Else
                summaryRowItem.Add(New GridViewSummaryItem(gv1.Columns(ii).Name, "{0:F2}", GridAggregateFunction.Sum))
            End If
        Next
        gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(Report_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Report_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
                    Dim ii As Integer = 0
                    For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
                        gv1.Columns(ii).IsVisible = False
                        gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
                obj = Nothing
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        If clsCommon.myLen(Report_ID) > 0 Then
            gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = Report_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv1.SaveLayout(obj.GridLayout)
            obj.GridColumns = gv1.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(Report_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        Try
            If gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Date : " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "  To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy"))
                If txtCustomer.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Customer Code : " & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & "   Customer Name :" & clsCommon.GetMulcallString(txtCustomer.arrDispalyMember) & "")
                End If

                If txtItem.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Item Code : " & clsCommon.GetMulcallString(txtItem.arrValueMember) & "   Item Name :" & clsCommon.GetMulcallString(txtItem.arrDispalyMember) & "")
                End If
                clsCommon.MyExportToPDF(Me.Text, gv1, arrHeader, Me.Text)
            Else
                clsCommon.MyMessageBoxShow(Me, "No data found to export", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        CancelPressed()
    End Sub
    Sub CancelPressed()
        Me.Close()
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Try
            If gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_code='" & clsUserMgtCode.rptCustItemWiseSaleReport & "'"))
                arrHeader.Add("Date : " & clsCommon.myCDate(txtFromDate.Value) + "  To " + clsCommon.myCDate(txtToDate.Value))

                If txtCustomer.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Customer Code : " & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & "   Customer Name :" & clsCommon.GetMulcallString(txtCustomer.arrDispalyMember) & "")
                End If

                If txtItem.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Item Code : " & clsCommon.GetMulcallString(txtItem.arrValueMember) & "   Item Name :" & clsCommon.GetMulcallString(txtItem.arrDispalyMember) & "")
                End If
                clsCommon.MyExportToExcel(Me.Text, gv1, arrHeader, Me.Text)
                clsCommon.MyMessageBoxShow(Me, "Export Successfully", Me.Text)
            Else
                clsCommon.MyMessageBoxShow(Me, "No data found to export", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
        '    reportSTCPartyRegister(True)
        'Else
        GetReportID()
        isPrint = True
        If BtnProductWiseSaleQuantity.IsChecked Then
            ProductWiseSale(True)
        ElseIf BtnBillWiseSaleOfMilkSummary.IsChecked Then
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                BillWisesaleSummaryDispatch(True)
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                BillWisesaleSummaryInvoice(True)
            End If
        ElseIf BtnPartySaleMilkProduct.IsChecked Then
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type ", Me.Text)
                Exit Sub
            End If
            If clsCommon.CompairString(ddlType.SelectedValue, "Dispatch") = CompairStringResult.Equal Then
                PartySaleMilkProductDispatch(True)
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Invoice") = CompairStringResult.Equal Then
                PartySaleMilkProductInvoice(True)
            End If
        ElseIf BtnBillWiseSaleOfMilk.IsChecked Then
            BillwiseSaleOfMilk(True)
        ElseIf BtnProductSalesSummary.IsChecked Then
            Productsalesummarytaxablenontaxable(True)
        ElseIf BtnMilkStcSummary.IsChecked Then
            MilkStcSummary(True)
        ElseIf BtnStcRegisterItemWiseSummary.IsChecked Then
            STCRegisterItemwiseSummarytotal(True)
        ElseIf BtnStcRegisterPartyandItemWiseSummary.IsChecked Then
            STCRegisterItemwiseSummaryPartyWise(True)
        ElseIf BtnTransportationCharges.IsChecked Then
            TransportationCharges(True)
        ElseIf BtnTcsSummary.IsChecked Then
            TcsSummary(True)
        ElseIf BtnGheeReport.IsChecked Then
            GheeReport(True)
        ElseIf BtnRouteWiseSale.IsChecked Then
            RouteWiseSale(True)
        ElseIf BtnCreditPartyWiseSaleAmount.IsChecked Then
            CreditPartyWiseSaleAmount(True)
        ElseIf rbtnDistributorCollStatement.IsChecked Then
            LoadDistributorCollStatementData(True)
        ElseIf rbtnMilkSale.IsChecked Then
            LoadMilkSaleData(True)
        ElseIf rbtnStockStatement.IsChecked Then
            LoadStockStatementData(True)
        ElseIf rbtnBoothSaleDateShiftWise.IsChecked Then
            LoadBoothSaleDateShiftWiseData(True)
        ElseIf rbtnBoothSaleItemWise.IsChecked Then
            LoadBoothSaleItemWiseData(True)
        ElseIf rbtnSalesRegister.IsChecked Then
            LoadSalesRegisterData(True)
        Else
            isPrint = True
            LoadData()
            isPrint = False
        End If
        isPrint = False
        'End If
    End Sub

    Private Sub txtTransaction__My_Click(sender As Object, e As EventArgs) Handles txtTransaction._My_Click
        Try
            Dim qry As String = "SELECT Program_Code As Code,Program_Name as Name from TSPL_PROGRAM_MASTER where Program_Code in ( 'AR-INVOICE','INVOICE-BS','BULK-SALE-RE','CAN-SALE','VCGL','M-Material-R','STO-TRANSFER','STO-TRANS-R')
       UNION ALL " & Environment.NewLine & " Select 'DS' as Code,'Dairy Sale' as Name " &
                  " Union All " & Environment.NewLine &
                  " select 'DSR','Dairy Sale Return' as Name  Union All " & Environment.NewLine & "
        Select 'SCRAP-SALE' As Code,'Scrap Sale' as Name  Union All " & Environment.NewLine & "
                  Select 'SCRAP-S-R' As Code,'Scrap Sale Return' as Name  "

            txtTransaction.arrValueMember = clsCommon.ShowMultipleSelectForm("TransMulSel", qry, "Code", "Name", txtTransaction.arrValueMember, txtTransaction.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadQtyConversionType()
        Dim qry As String = "select '' as Code, '<--Select-->' as Name union all 
                             Select 'Ltr' as Code,'Qty in Ltr' as Name  union all 
 Select 'Pouch' as Code,'Qty in Pouch' as Name union all 
 Select 'Crate' as Code,'Qty in Crate' as Name "
        ddlQtyConversionType.DataSource = clsDBFuncationality.GetDataTable(qry)
        ddlQtyConversionType.ValueMember = "Code"
        ddlQtyConversionType.DisplayMember = "Name"
    End Sub
    Sub LoadDefaultReportUOMType()
        Dim qry As String = "select '' as Code, '<--Select-->' as Name union all 
                             Select 'Default_UOM' as Code,'Default UOM' as Name  union all 
         Select 'Report_UOM' as Code,'Report UOM' as Name  "
        ddlDefaultReportUOM.DataSource = clsDBFuncationality.GetDataTable(qry)
        ddlDefaultReportUOM.ValueMember = "Code"
        ddlDefaultReportUOM.DisplayMember = "Name"
    End Sub
    Sub LoadType()
        Try
            Dim qry As String = "select '' as Code, '<--Select-->' as Name   "
            If BtnProductSalesSummary.IsChecked Then
                qry += " union all Select 'Taxable' as Code,'Taxable' as Name  union all 
 Select 'Non Taxable' as Code,'Non Taxable' as Name union all 
 Select 'Both' as Code,'Both' as Name "
            ElseIf BtnPartySaleMilkProduct.IsChecked OrElse BtnBillWiseSaleOfMilkSummary.IsChecked OrElse BtnBillWiseSaleOfMilk.IsChecked Then
                qry += " union all Select 'Dispatch' as Code,'Dispatch' as Name  union all 
 Select 'Invoice' as Code,'Invoice' as Name "
            ElseIf rbtnBoothSaleDateShiftWise.IsChecked OrElse rbtnBoothSaleItemWise.IsChecked OrElse rbtnMilkSale.IsChecked Then
                qry += " union all Select 'Milk' as Code,'Milk' as Name  union all 
 Select 'Product' as Code,'Product' as Name 
            union all 
 Select 'Both' as Code,'Both' as Name "
            ElseIf rbtnSalesRegister.IsChecked Then
                qry += " union all Select 'Taxable' as Code,'Taxable' as Name  union all 
 Select 'Non Taxable' as Code,'Non Taxable' as Name  "
            End If
            ddlType.DataSource = clsDBFuncationality.GetDataTable(qry)
            ddlType.ValueMember = "Code"
            ddlType.DisplayMember = "Name"
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadReportType()
        Try
            Dim qry As String = "select '' as Code, '<--Select-->' as Name   "
            If BtnProductSalesSummary.IsChecked OrElse rbtnMilkSale.IsChecked OrElse BtnStcRegisterItemWiseSummary.IsChecked OrElse BtnMilkStcSummary.IsChecked Then
                qry += " union all Select 'Item Wise' as Code,'Item Wise' as Name  union all 
 Select 'Party Wise' as Code,'Party Wise' as Name "
                If BtnMilkStcSummary.IsChecked OrElse BtnStcRegisterItemWiseSummary.IsChecked Then
                    qry += " union all Select 'Summary' as Code,'Summary' as Name "
                End If
            ElseIf rbtnSalesRegister.IsChecked Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                    qry += " union all Select 'Milk' as Code,'Milk' as Name  union all 
 Select 'Product' as Code,'Product' as Name 
            union all 
 Select 'Both' as Code,'Both' as Name "
                End If
            End If

            ddlReportType.DataSource = clsDBFuncationality.GetDataTable(qry)
            ddlReportType.ValueMember = "Code"
            ddlReportType.DisplayMember = "Name"
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtLocation__My_Click(sender As Object, e As EventArgs) Handles txtLocation._My_Click
        Try
            Dim qry As String = "select TSPL_LOCATION_MASTER.Location_Code as Code,TSPL_LOCATION_MASTER.Location_Desc as Name from TSPL_LOCATION_MASTER  "
            txtLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("LocFilter", qry, "Code", "Name", txtLocation.arrValueMember, txtLocation.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnStockStatement_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnStockStatement.ToggleStateChanged, rbtnMilkSale.ToggleStateChanged, rbtnDistributorCollStatement.ToggleStateChanged, rbtnBoothSaleDateShiftWise.ToggleStateChanged, rbtnBoothSaleItemWise.ToggleStateChanged, BtnPartySaleMilkProduct.ToggleStateChanged, BtnProductSalesSummary.ToggleStateChanged, BtnBillWiseSaleOfMilkSummary.ToggleStateChanged, BtnStcRegisterPartyandItemWiseSummary.ToggleStateChanged, BtnStcRegisterItemWiseSummary.ToggleStateChanged, BtnMilkStcSummary.ToggleStateChanged, rbtnSalesRegister.ToggleStateChanged, BtnBillWiseSaleOfMilk.ToggleStateChanged
        If rbtnStockStatement.IsChecked Then
            txtLocation.Visible = True
            lblLocation.Visible = True
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            btnGo.Enabled = True
            RadSplitButton1.Enabled = True
            txtRoute.Visible = False
            lblRoute.Visible = False
            lblQtyConv.Visible = True
            ddlQtyConversionType.Visible = False
            ddlDefaultReportUOM.Visible = True
            RadGroupBox4.Visible = False
            lblType.Visible = False
            ddlType.Visible = False
            lblReportType.Visible = False
            ddlReportType.Visible = False
        ElseIf rbtnMilkSale.IsChecked Then
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            btnGo.Enabled = False
            RadSplitButton1.Enabled = False
            txtRoute.Visible = True
            lblRoute.Visible = True
            lblQtyConv.Visible = True
            ddlQtyConversionType.Visible = True
            ddlDefaultReportUOM.Visible = False
            RadGroupBox4.Visible = True
            lblType.Visible = True
            ddlType.Visible = True
            LoadType()
            lblReportType.Visible = True
            ddlReportType.Visible = True
            LoadReportType()
        ElseIf rbtnDistributorCollStatement.IsChecked Then
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            txtRoute.Visible = True
            lblRoute.Visible = True
            btnGo.Enabled = True
            RadSplitButton1.Enabled = True
            lblQtyConv.Visible = True
            ddlQtyConversionType.Visible = True
            ddlDefaultReportUOM.Visible = False
            RadGroupBox4.Visible = False
            lblType.Visible = False
            ddlType.Visible = False
            lblReportType.Visible = False
            ddlReportType.Visible = False
        ElseIf rbtnBoothSaleDateShiftWise.IsChecked Then
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            btnGo.Enabled = False
            RadSplitButton1.Enabled = False
            txtRoute.Visible = True
            lblRoute.Visible = True
            lblQtyConv.Visible = True
            ddlDefaultReportUOM.Visible = True
            RadGroupBox4.Visible = True
            lblType.Visible = True
            ddlType.Visible = True
            lblReportType.Visible = False
            ddlReportType.Visible = False
            LoadType()
        ElseIf rbtnBoothSaleItemWise.IsChecked Then
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            btnGo.Enabled = True
            RadSplitButton1.Enabled = False
            txtRoute.Visible = True
            lblRoute.Visible = True
            lblQtyConv.Visible = True
            ddlDefaultReportUOM.Visible = True
            RadGroupBox4.Visible = True
            lblType.Visible = True
            ddlType.Visible = True
            lblReportType.Visible = False
            ddlReportType.Visible = False
            LoadType()
        ElseIf BtnProductSalesSummary.IsChecked OrElse BtnPartySaleMilkProduct.IsChecked OrElse BtnBillWiseSaleOfMilkSummary.IsChecked OrElse BtnBillWiseSaleOfMilk.IsChecked Then
            btnGo.Enabled = True
            RadSplitButton1.Enabled = True
            txtRoute.Visible = False
            lblRoute.Visible = False
            lblQtyConv.Visible = False
            ddlQtyConversionType.Visible = False
            ddlDefaultReportUOM.Visible = False
            RadGroupBox4.Visible = False
            lblReportType.Visible = False
            ddlReportType.Visible = False
            lblType.Visible = True
            ddlType.Visible = True
            txtLocation.Visible = False
            lblLocation.Visible = False
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            If BtnProductSalesSummary.IsChecked Then
                lblReportType.Visible = True
                ddlReportType.Visible = True
                LoadReportType()
                txtLocation.Visible = True
                lblLocation.Visible = True
            End If
            LoadType()
            If BtnBillWiseSaleOfMilkSummary.IsChecked OrElse BtnBillWiseSaleOfMilk.IsChecked Then
                RadGroupBox4.Visible = True
            End If
        ElseIf BtnStcRegisterPartyandItemWiseSummary.IsChecked OrElse BtnStcRegisterItemWiseSummary.IsChecked OrElse BtnMilkStcSummary.IsChecked Then
            btnGo.Enabled = True
            RadSplitButton1.Enabled = True
            txtLocation.Visible = True
            lblLocation.Visible = True
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            lblReportType.Visible = False
            ddlReportType.Visible = False
            lblQtyConv.Visible = False
            ddlQtyConversionType.Visible = False
            ddlDefaultReportUOM.Visible = False
            If BtnStcRegisterItemWiseSummary.IsChecked Then
                lblLocation.Text = "From Location"
                lblToLocation.Visible = True
                txtToLocation.Visible = True
                lblReportType.Visible = True
                ddlReportType.Visible = True
                LoadReportType()
                lblQtyConv.Visible = True
                ddlDefaultReportUOM.Visible = True
                LoadDefaultReportUOMType()
            ElseIf BtnMilkStcSummary.IsChecked Then
                lblLocation.Text = "From Location"
                lblToLocation.Visible = True
                txtToLocation.Visible = True
                lblReportType.Visible = True
                ddlReportType.Visible = True
                LoadReportType()
            Else
                lblLocation.Text = "Location"
            End If
            txtRoute.Visible = False
            lblRoute.Visible = False
            RadGroupBox4.Visible = False
            lblType.Visible = False
            ddlType.Visible = False
        ElseIf rbtnSalesRegister.IsChecked Then
            btnGo.Enabled = True
            RadSplitButton1.Enabled = True
            txtLocation.Visible = False
            lblLocation.Visible = False
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                lblReportType.Visible = True
                lblReportType.Text = "Item Type"
                ddlReportType.Visible = True
                LoadReportType()
            Else
                lblReportType.Visible = False
                ddlReportType.Visible = False
            End If
            lblQtyConv.Visible = True
            ddlQtyConversionType.Visible = False
            ddlDefaultReportUOM.Visible = True
            LoadDefaultReportUOMType()
            txtRoute.Visible = False
            lblRoute.Visible = False
            RadGroupBox4.Visible = False
            lblType.Visible = True
            ddlType.Visible = True
            LoadType()
        Else
            lblToLocation.Visible = False
            txtToLocation.Visible = False
            btnGo.Enabled = True
            RadSplitButton1.Enabled = True
            txtLocation.Visible = False
            lblLocation.Visible = False
            txtRoute.Visible = False
            lblRoute.Visible = False
            lblQtyConv.Visible = False
            ddlQtyConversionType.Visible = False
            ddlDefaultReportUOM.Visible = False
            RadGroupBox4.Visible = False
            lblType.Visible = False
            ddlType.Visible = False
            lblReportType.Visible = False
            ddlReportType.Visible = False
        End If
    End Sub

    Private Sub txtRoute__My_Click(sender As Object, e As EventArgs) Handles txtRoute._My_Click
        Try
            Dim qry As String = "select Route_No as Code,Route_Desc as Name from TSPL_ROUTE_MASTER  "
            txtRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("RouteFilter", qry, "Code", "Name", txtRoute.arrValueMember, txtRoute.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtToLocation__My_Click(sender As Object, e As EventArgs) Handles txtToLocation._My_Click
        Try
            Dim qry As String = "select TSPL_LOCATION_MASTER.Location_Code as Code,TSPL_LOCATION_MASTER.Location_Desc as Name from TSPL_LOCATION_MASTER  "
            txtToLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("LocFilter", qry, "Code", "Name", txtToLocation.arrValueMember, txtToLocation.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadSalesRegisterData(ByVal Print As Boolean)
        Try
            Dim WhrCust As String = ""
            Dim Sublocn As String = ""
            Dim item As String = ""
            Dim dt As DataTable = Nothing
            Dim strtxtfDate As String = clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy")
            Dim strToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
            Dim qry As String = ""
            Dim Baseqry As String = ""
            If ddlDefaultReportUOM.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Quantity conversion type", Me.Text)
                Exit Sub
            End If
            If ddlType.SelectedIndex = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select Type", Me.Text)
                Exit Sub
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                If ddlReportType.SelectedIndex = 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Please select Item Type", Me.Text)
                    Exit Sub
                End If
            End If

            Baseqry = " select  
        CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD.Against_Booking_No IS NOT NULL)  And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC'  THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
		  LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD.Screen_Type= ('CT'))) And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' THEN 'APS SALES'
        WHEN TSPL_SD_SALE_INVOICE_HEAD.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'	  
ELSE 'DCS SALE' END AS Transcation_Type,
case when TSPL_SD_SALE_INVOICE_HEAD.Status=1 then 'Approved' else'Pending' end as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location AS [Location],
TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Invoice_Date,TSPL_SD_SHIPMENT_HEAD.Shift_Type,
TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,
TSPL_SD_SALE_INVOICE_HEAD.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_INVOICE_DETAIL.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],(TSPL_SD_SALE_INVOICE_DETAIL.item_cost * Report_UOM.Conversion_Factor/tspl_item_uom_detail.Conversion_Factor ) as [Rate],
TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as [Measure of Qty],
  TSPL_SD_SALE_INVOICE_DETAIL.Qty as [Product Qty],
  tspl_item_master.HSN_Code,TSPL_ITEM_MASTER.Is_FreshItem ,TSPL_ITEM_MASTER.Is_Ambient ,
  case when TSPL_SD_SALE_INVOICE_DETAIL.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate else  TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate + TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate else TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate +TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate end)end as [IGST Rate], 
     TSPL_SD_SALE_INVOICE_DETAIL.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_INVOICE_Detail.Item_Net_Amt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],
TSPL_SD_SALE_INVOICE_HEAD.Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SALE_INVOICE_HEAD.IRN_No,
Report_UOM.UOM_Code as Report_UOM,
(qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_INVOICE_HEAD.Created_By,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Created_Date,103) as Created_Date,TSPL_SD_SALE_INVOICE_HEAD.Total_Amt As [Bill Amount]

                         from TSPL_SD_SALE_INVOICE_HEAD
left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.Document_Code
left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_INVOICE_DETAIL.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where " + ddlDefaultReportUOM.SelectedValue + " = 1 ) as  Report_UOM ON TSPL_SD_SALE_INVOICE_DETAIL.Item_Code = Report_UOM.item_code 
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX10 "

             If rbtnSupplyDate.IsChecked Then
                Baseqry += "  where Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >=Convert( Date,'" + strtxtfDate + "',103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)<=Convert( Date,'" + strToDate + "',103) "
            ElseIf rbtnDocumentDate.IsChecked Then
                Baseqry += "  where Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value) + "' and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value) + "'  "
            End If

            If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                Baseqry += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code in(" + clsCommon.GetMulcallString(txtLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            Baseqry += "            union all

select  'MATERIAL SALE' as Transcation_Type,case when TSPL_SCRAPINVOICE_HEAD.ispost=1 then 'Approved' else'Pending' end as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SCRAPINVOICE_HEAD.Loc_Code AS [Location],
TSPL_SCRAPINVOICE_HEAD.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) as Invoice_Date,'' as Shift_Type,
TSPL_SCRAPINVOICE_HEAD.invoice_No as Invoice_No,
TSPL_ROUTE_MASTER.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SCRAPINVOICE_Detail.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],TSPL_SCRAPINVOICE_Detail.Price As [Rate],
TSPL_SCRAPINVOICE_Detail.Unit_code as [Measure of Qty],
  TSPL_SCRAPINVOICE_Detail.shipped_Qty as [Product Qty],
  tspl_item_master.HSN_Code,TSPL_ITEM_MASTER.Is_FreshItem ,TSPL_ITEM_MASTER.Is_Ambient ,
  case when TSPL_SCRAPINVOICE_Detail.Tax1='KKF' or TSPL_SCRAPINVOICE_Detail.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_Detail.tax3='IGST' then TSPL_SCRAPINVOICE_Detail.TAX3_Rate else  TSPL_SCRAPINVOICE_Detail.TAX3_Rate + TSPL_SCRAPINVOICE_Detail.TAX4_Rate end ) else (case when  TSPL_SCRAPINVOICE_Detail.tax1='IGST' then TSPL_SCRAPINVOICE_Detail.TAX1_Rate else TSPL_SCRAPINVOICE_Detail.TAX1_Rate +TSPL_SCRAPINVOICE_Detail.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SCRAPINVOICE_Detail.Tax1='KKF' or TSPL_SCRAPINVOICE_Detail.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_Detail.tax3='IGST' then TSPL_SCRAPINVOICE_Detail.TAX3_Base_Amt else  TSPL_SCRAPINVOICE_Detail.TAX3_Base_Amt end ) else (case when  TSPL_SCRAPINVOICE_Detail.tax1='IGST' then TSPL_SCRAPINVOICE_Detail.TAX1_Base_Amt else TSPL_SCRAPINVOICE_Detail.TAX1_Base_Amt end)end as [Basic Amt]
     TSPL_SCRAPINVOICE_HEAD.Discount_Base as[ItemBasic Amt],
    TSPL_SCRAPINVOICE_HEAD.Discount_Amt as[Margin Amt],
  TSPL_SCRAPINVOICE_HEAD.Amount_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SCRAPINVOICE_Detail.TAX1='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX2='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX3='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX4='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX5='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX6='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX7='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX8='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX9='KKF'  THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt 
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX10='KKF' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SCRAPINVOICE_Detail.TAX1='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX2='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX3='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX4='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX5='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX6='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX7='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX8='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX9='TCS'  THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX10='TCS' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SCRAPINVOICE_Detail.TAX1='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX2='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX3='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX4='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX5='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX6='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX7='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX8='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX9='CGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX10='CGST' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SCRAPINVOICE_Detail.TAX1='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Rate
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX2='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX3='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX4='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX5='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX6='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX7='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX8='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX9='SGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX10='SGST' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt else 0 END) AS [SGST Amt],
					CASE WHEN TSPL_SCRAPINVOICE_Detail.TAX1='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX2='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX3='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX4='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX5='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX6='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX7='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX8='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX9='IGST'  THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_Detail.TAX10='IGST' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt as [Total Tax Amt],
					TSPL_SCRAPINVOICE_Detail.TotalAmt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],
TSPL_SCRAPINVOICE_HEAD.Ack_No,TSPL_SCRAPINVOICE_HEAD.Ack_Date,TSPL_SCRAPINVOICE_HEAD.IRN_No
,Report_UOM.UOM_Code as Report_UOM,
(shipped_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty,0 as [Subsidy Amt], TSPL_SCRAPINVOICE_HEAD.Invoice_Type as [Invoice Type],
TSPL_SCRAPINVOICE_HEAD.EWayBillNo,TSPL_SCRAPINVOICE_HEAD.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SCRAPINVOICE_HEAD.Created_By,Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD.Created_Date,103) as Created_Date,TSPL_SCRAPINVOICE_HEAD.Doc_Amt As [Bill Amount]
                           from TSPL_SCRAPINVOICE_HEAD
                    left join TSPL_SCRAPINVOICE_Detail on TSPL_SCRAPINVOICE_Detail.invoice_No=TSPL_SCRAPINVOICE_HEAD.invoice_No
                    LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SCRAPINVOICE_Detail.Item_Code
                    left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SCRAPINVOICE_Detail.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SCRAPINVOICE_Detail.Unit_code
                        LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where " + ddlDefaultReportUOM.SelectedValue + " = 1 ) as  Report_UOM ON TSPL_SCRAPINVOICE_Detail.Item_Code = Report_UOM.item_code 
                    LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SCRAPINVOICE_HEAD.Loc_Code
                     LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SCRAPINVOICE_HEAD.cust_Code
                    left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SCRAPINVOICE_HEAD.cust_Code
                     LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
                     left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SCRAPINVOICE_HEAD.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SCRAPINVOICE_HEAD.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SCRAPINVOICE_HEAD .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SCRAPINVOICE_HEAD .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SCRAPINVOICE_HEAD .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SCRAPINVOICE_HEAD .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SCRAPINVOICE_HEAD .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SCRAPINVOICE_HEAD .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SCRAPINVOICE_HEAD .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SCRAPINVOICE_HEAD .TAX10 
                    where convert(date,shipment_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,shipment_Date,103)<=Convert( Date,'" + strToDate + "',103)  "

            If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
                Baseqry += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtLocation.arrValueMember IsNot Nothing AndAlso txtLocation.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SCRAPINVOICE_HEAD.Sub_Location_code in(" + clsCommon.GetMulcallString(txtLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SCRAPINVOICE_Detail.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            If Print Then
                qry &= " Select xx.*,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.State,TSPL_STATE_MASTER.STATE_NAME,TSPL_COMPANY_MASTER.Pincode from ("
                qry &= " Select ROW_NUMBER() OVER (PARTITION BY Invoice_No ORDER BY [Item Code]) AS SNo,Max(PrintBy)PrintBy,Max(FromDate)FromDate,Max(ToDate)ToDate,Max([Transcation Type])[Transcation Type],MAX([Customer Type])[Customer Type],Max([Doc Status])[Doc Status],Max(Location)Location,Max([Location GST])[Location GST],Max([Sub Location])[Sub Location],Invoice_Date,Shift,Invoice_No,Max([Invoice Type])[Invoice Type],Max([Route No])[Route No],([Party Code])[Party Code],Max([DCS Uploader])[DCS Uploader],Max([Party Name])[Party Name],
                    Max([Customer GSTNo])[Customer GSTNo],Max([Party State Code])[Party State Code],[Item Code],Max([Item Name])[Item Name],Max([HSN Code])[HSN Code],Max([Measure of Qty])[Measure of Qty],Sum([Product Qty])[Product Qty],Max([Report UOM])[Report UOM],Max(Rate)Rate,
					Sum(CAST([ReportUOM Qty] AS DECIMAL(18,2))) AS [ReportUOM Qty],
					Max(Cast(([GST Rate]) as decimal(10,2)))[GST Rate],Sum(cast(([ItemBasic Amt]) as decimal(10,2)))[ItemBasic Amt],
					Sum(cast(([Margin Amt]) as decimal(10,2)))[Margin Amt],Sum(Cast(([Basic Amt]) as Decimal(10,2)))[Basic Amt],Sum(Cast((KKF) as decimal(10,2)))KKF,
					Sum(Cast(([Mandi Tax Amt]) as Decimal(10,2)))[Mandi Tax Amt],Sum(Cast(([Party TCS Amt]) as Decimal(10,2)))[Party TCS Amt],
                        Sum(Cast(([CGST Amt]) as Decimal(10,2)))[CGST Amt],SUm(Cast(([SGST Amt]) as Decimal(10,2)))[SGST Amt],Sum(Cast(([IGST Amt]) as decimal(10,2)))[IGST Amt],Sum(cast(([Total Tax Amt]) as decimal(10,2)))[Total Tax Amt],Sum(cast(([Total Amt]) as decimal(10,2)))[Total Amt],Sum(([Subsidy Amt]))[Subsidy Amt],Max([B2B/B2C])[B2B/B2C],Max([Ack No])[Ack No],Max([Ack Date])[Ack Date],Max([IRN No])[IRN No],Max(EWayBillNo)EWayBillNo,Max(EWayBillDate)EWayBillDate
                    ,Max([Created By])[Created By],Max([Created Date])[Created Date],Max([Bill Amount])[Bill Amount] from ("
            End If
            qry &= "Select '" & objCommonVar.CurrentUser & "' As PrintBy,ROW_NUMBER() OVER (PARTITION BY xx.Invoice_No ORDER BY xx.[Item Code]) AS SNo,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd-MM-yyyy") & "' As FromDate,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd-MM-yyyy") & "' As ToDate,(Transcation_Type)[Transcation Type],([Customer Type])[Customer Type],(Doc_Status)[Doc Status],(Location)Location,([GST No])[Location GST],([Sub Location])[Sub Location],(Invoice_Date)Invoice_Date, case when Shift_Type = 'AM' then 'M' when Shift_Type = 'PM' then 'E' end as Shift,Invoice_No,([Invoice Type])[Invoice Type],([Route No])[Route No],([Party Code])[Party Code],VLC_Code_VLC_Uploader as [DCS Uploader],([Party Name])[Party Name],
                    ([Customer GSTNo])[Customer GSTNo],([State Code])[Party State Code],[Item Code],[Item Name],HSN_Code AS [HSN Code],([Measure of Qty])[Measure of Qty],([Product Qty])[Product Qty],(Report_UOM)[Report UOM],Rate,CAST(ReportUOM_Qty AS DECIMAL(18,2)) AS [ReportUOM Qty],Cast(([IGST Rate]) as decimal(10,2))[GST Rate],cast(([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast((KKF) as decimal(10,2))KKF,Cast(([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(([IGST Amt]) as decimal(10,2))[IGST Amt],cast(([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(([Total Amt]) as decimal(10,2))[Total Amt],([Subsidy Amt])[Subsidy Amt],([B2B/B2C])[B2B/B2C],(Ack_No)[Ack No],(Ack_Date)[Ack Date],(IRN_No)[IRN No],EWayBillNo,EWayBillDate,CASE WHEN ROW_NUMBER() OVER (PARTITION BY Invoice_No ORDER BY [Item Code] DESC) = 1 THEN [Bill Amount] Else 0 End As [Bill Amount]
                    ,(Created_By)[Created By],(Created_Date)[Created Date]
                    from (" & Baseqry & ")XX   "
            qry &= " Where 1=1 "

            If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal Then
                qry &= " And XX.[Invoice Type]='Taxable' "
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                qry &= " And XX.[Invoice Type]='Non-Taxable' "
            End If
            If txtTransaction.arrValueMember IsNot Nothing AndAlso txtTransaction.arrValueMember.Count > 0 Then
                qry += " And XX.Transcation_Type In(" + clsCommon.GetMulcallString(txtTransaction.arrValueMember) + ")" + Environment.NewLine
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                If clsCommon.CompairString(ddlReportType.SelectedValue, "Milk") = CompairStringResult.Equal Then
                    qry += " and  Is_FreshItem = 1  "
                ElseIf clsCommon.CompairString(ddlReportType.SelectedValue, "Product") = CompairStringResult.Equal Then
                    qry += " and Is_Ambient = 1  "
                End If
            End If

            If Print Then
                qry &= " )xyz "
                qry &= " Group By [Party Code],Invoice_No,Invoice_Date,shift,[Item Code] "
                qry &= " )xx Left Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
					          Left Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State"
                qry += " order by xx.[Transcation Type], xx.Invoice_Date,shift desc "
            End If

            dt = clsDBFuncationality.GetDataTable(qry)
            gv1.GroupDescriptors.Clear()
            gv1.MasterTemplate.SummaryRowsBottom.Clear()
            gv1.DataSource = Nothing
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                If Print Then
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Print", Me.Text)
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                End If
                Exit Sub
            Else
                If Print Then
                    Dim frm As New frmCrystalReportViewer()
                    If clsCommon.CompairString(ddlType.SelectedValue, "Taxable") = CompairStringResult.Equal Then
                        frm.funreport(Form_ID, CrystalReportFolder.SalesReport, dt, "crptSaleRegisterTaxable", "Sales Register (Bill Of Supply)")
                    ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Non Taxable") = CompairStringResult.Equal Then
                        frm.funreport(Form_ID, CrystalReportFolder.SalesReport, dt, "crptSaleRegisterNonTaxable", "Sales Register (Tax Invoice)")
                    End If
                    frm = Nothing
                Else
                    RadPageView1.SelectedPage = RadPageViewPage2
                    gv1.GroupDescriptors.Clear()
                    gv1.MasterTemplate.SummaryRowsBottom.Clear()
                    gv1.DataSource = dt
                    gv1.AutoExpandGroups = True
                    gv1.ShowGroupPanel = True
                    gv1.ShowRowHeaderColumn = False
                    gv1.AllowAddNewRow = False
                    gv1.AllowDeleteRow = False
                    gv1.EnableFiltering = True
                    gv1.ShowFilteringRow = True
                    'SetGridFormation()
                    'EnableDisableControls(False)
                    gv1.BestFitColumns()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        ReStoreGridLayout()
    End Sub
End Class


