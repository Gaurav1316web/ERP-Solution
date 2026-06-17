Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Imports common
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime
'' ERO/05/04/19-000545 (work on only on report format in case of taxable invoice)
Public Class FrmPrintDistributerInvoiceStatement
    Inherits FrmMainTranScreen
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim isInsideLoadData As Boolean = False
    Dim strQry As String = ""
    'Dim Refresh As Boolean = False
    Public Shared ArrInvoice_Arr As New ArrayList()
    Dim ShowShipToPartyInDairyDispatch As Integer = 0
    Dim AllowSeperateSchemeItemOnPrint As Boolean = False
    Dim strRptPath As String = ""
    Dim ApplyMilkPouchPrint As Boolean = False
    Dim EnableProductSaleForJPR As Boolean = False
    Dim lstinvNo As List(Of String)
    Dim chkbtnEmailSMS As Boolean = False


    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.FrmPrintDistributerInvoiceStatement)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        RadSplitButton1.Visible = MyBase.isExport
    End Sub

    Function ReturnLoadReportQry() As String
        Dim sQuery As String = Nothing
        Dim WhrCls As String = " and 2=2 "
        Dim Whr As String = ""

        If rbtnEvening.Checked Then
            Whr += "And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'PM' "
        ElseIf rbtnMorning.Checked Then
            Whr += "And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' "
        End If

        If EnableProductSaleForJPR Then
            If rbtnMilk.Checked Then
                Whr += " and TSPL_SD_SALE_INVOICE_HEAD.item_type IN ('S','') "
            ElseIf rbtnProduct.Checked Then
                Whr += " and TSPL_SD_SALE_INVOICE_HEAD.item_type='P' "
            ElseIf rbtnIceCream.Checked Then
                Whr += " and TSPL_SD_SALE_INVOICE_HEAD.item_type='I' "
            End If
        End If

        Dim ItemType As String = ""
        If EnableProductSaleForJPR Then
            ItemType = " case when  TSPL_SD_SALE_INVOICE_HEAD.item_type = 'M' then 'Milk' when TSPL_SD_SALE_INVOICE_HEAD.item_type = 'P' then 'Product' when TSPL_SD_SALE_INVOICE_HEAD.item_type = 'I' then 'Ice Cream' end as [Item Type] ,"
        Else
            ItemType = "'' as [Item Type] ,"
        End If

        'sQuery += "  select Cast(0 as BIT) as 'Check', TSPL_SD_SALE_INVOICE_HEAD.Document_Code," & ItemType & " convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date,
        '                        TSPL_SD_SALE_INVOICE_HEAD.Route_No as Route_Code,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,Customer_Name  ,
        '                        Location_Desc ,TSPL_SD_SALE_INVOICE_HEAD.Total_Amt,
        '                     CASE 
        '                        WHEN Shift_type = 'AM' THEN 'Morning'
        '                        ELSE 'Evening'
        '                        END AS Shift_type,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) as Supply_Date,TSPL_CUSTOMER_MASTER.Email as Email,TSPL_CUSTOMER_MASTER.Phone1 as Mobile_no
        '                        from TSPL_SD_SALE_INVOICE_HEAD
        '                        left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and TSPL_SD_SALE_INVOICE_DETAIL.Line_No=1
        '                        left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
        '                        left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
        '                        left join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE =TSPL_LOCATION_MASTER.State 
        '                        left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
        '                        where  TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND TSPL_SD_SALE_INVOICE_HEAD.Screen_Type='DS'
        '                        and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' ,103) " + Whr + " "

        sQuery += " select Cast(0 as BIT) as 'Check', TSPL_SD_SALE_INVOICE_HEAD.Document_Code," & ItemType & "convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date,TSPL_ROUTE_MASTER.Zone_Code,TSPL_SD_SALE_INVOICE_HEAD.Route_No as Route_Code,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,Customer_Name ,Location_Desc ,TSPL_SD_SALE_INVOICE_HEAD.Total_Amt ,CASE WHEN Shift_type = 'AM' THEN 'Morning' ELSE 'Evening' END AS Shift_type,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) as Supply_Date,TSPL_CUSTOMER_MASTER.Email as Email,TSPL_CUSTOMER_MASTER.Phone1 as Mobile_no "
        'If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "AL") = CompairStringResult.Equal Then
        sQuery += ",Isnull(TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable,0)Is_Taxable"
        'End If

        sQuery += " from TSPL_SD_SALE_INVOICE_HEAD
left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and TSPL_SD_SALE_INVOICE_DETAIL.Line_No=1 
left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
left join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
left join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE =TSPL_LOCATION_MASTER.State 
left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
where  1=1 "
        If chkDCSSale.Checked Then
            sQuery &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS','MCC') "
        Else
            sQuery &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND TSPL_SD_SALE_INVOICE_HEAD.Screen_Type='DS' "
        End If

        If rbtnDocumentDate.Checked Then
            sQuery += " and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' ,103) " + Whr + " "
        End If
        If rbtnSupplyDate.Checked Then
            sQuery += " and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' ,103) " + Whr + " "
        End If

        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "T") = CompairStringResult.Equal Then
            sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=1 "
            'sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=1 and TSPL_LOCATION_MASTER.State=TSPL_CUSTOMER_MASTER.State  and 0= (select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax_Group and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y'))"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "LWM") = CompairStringResult.Equal Then
            sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=1 and 1 <= (select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax_Group and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y'))"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "UT") = CompairStringResult.Equal Then
            sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=1  and isnull(TSPL_STATE_MASTER.Is_GST_UT,0)=1"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "IT") = CompairStringResult.Equal Then
            sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=1 and TSPL_LOCATION_MASTER.State<>TSPL_CUSTOMER_MASTER.State and 0= (select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax_Group and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y'))"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "IWM") = CompairStringResult.Equal Then
            sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=1 and TSPL_LOCATION_MASTER.State<>TSPL_CUSTOMER_MASTER.State and 1 <= (select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax_Group and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y'))"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "NT") = CompairStringResult.Equal Then
            sQuery += " and TSPL_SD_SALE_INVOICE_HEAD.is_taxable=0 "
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "AL") = CompairStringResult.Equal Then
            '   sQuery += "  and TSPL_LOCATION_MASTER.State=TSPL_CUSTOMER_MASTER.State  "
            'sQuery += "  and TSPL_LOCATION_MASTER.State=TSPL_CUSTOMER_MASTER.State  and 0= (select count(*) from TSPL_TAX_GROUP_DETAILS where Tax_Group_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax_Group and Tax_Code in(select Tax_Code from TSPL_TAX_MASTER where Is_Mandi_Tax='Y'))"
        End If

        'sQuery += " and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' ,103)"
        '==============update by preeti gupta Against Ticket No [BM00000005410]

        'If chkLocationSelect.IsChecked And cbgLocation.CheckedValue.Count > 0 Then
        '    sQuery += " and TSPL_LOCATION_MASTER. Location_Code   IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
        'End If
        If txtMultCustomer.arrValueMember IsNot Nothing AndAlso txtMultCustomer.arrValueMember.Count > 0 Then
            WhrCls += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultCustomer.arrValueMember) + ") "
        End If

        ''RICHA AGARWAL 19 jUNE,2019 SHOW ROUTE NO FROM TSPL_SD_SALE_INVOICE_HEAD INSTEAD OF CUSTOMER MASTER ERO/18/06/19-000646
        If txtMultRoute.arrValueMember IsNot Nothing AndAlso txtMultRoute.arrValueMember.Count > 0 Then
            WhrCls += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in (" + clsCommon.GetMulcallString(txtMultRoute.arrValueMember) + ") "
        End If

        If txtMultLocation.arrValueMember IsNot Nothing AndAlso txtMultLocation.arrValueMember.Count > 0 Then
            WhrCls += " and TSPL_LOCATION_MASTER. Location_Code   IN (" + clsCommon.GetMulcallString(txtMultLocation.arrValueMember) + ") "
        End If

        If TxtItem.arrValueMember IsNot Nothing AndAlso TxtItem.arrValueMember.Count > 0 Then
            WhrCls += " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code  IN (" + clsCommon.GetMulcallString(TxtItem.arrValueMember) + ") "
        End If

        sQuery += WhrCls
        'If Not rbtnPartyWise.Checked Then
        '    sQuery += " order by TSPL_SD_SALE_INVOICE_HEAD.Document_Date "
        'End If

        Return sQuery
    End Function

    Public Sub loadReport()
        Try

            Dim FinalQry As String = Nothing
            If clsCommon.myCDate(txtFromDate.Value) > clsCommon.myCDate(txtToDate.Value) Then
                common.clsCommon.MyMessageBoxShow(Me, "From date can not be greater then to Date", Me.Text)
                txtFromDate.Focus()
                Exit Sub
            End If

            If clsCommon.myCDate(txtFromDate.Value) >= objCommonVar.GSTApplicableDate AndAlso clsCommon.myCDate(txtToDate.Value) >= objCommonVar.GSTApplicableDate Then
                If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "") = CompairStringResult.Equal Then
                    clsCommon.MyMessageBoxShow(Me, "Please Select Report Type.", Me.Text)
                    cboReportType.Focus()
                    Exit Sub
                End If
            End If
            'If chkLocationSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count = 0 Then
            '    clsCommon.MyMessageBoxShow("Please select atleast single Location or select all.")
            '    Exit Sub
            'End If
            Dim sQuery As String = ReturnLoadReportQry()
            If rbtnPartyWise.Checked Then
                FinalQry = "WITH FilteredData AS (" + sQuery + ") "
                FinalQry += " SELECT CAST(0 AS BIT) AS [Check],Supply_Date,Shift_type,Customer_Code,MAX(Customer_Name) AS Customer_Name,(SELECT STRING_AGG(Route_Code, ', ')FROM (SELECT DISTINCT Route_Code FROM FilteredData r WHERE r.Customer_Code = f.Customer_Code AND r.Supply_Date = f.Supply_Date AND r.Shift_type = f.Shift_type) AS UniqueRoutes) AS Route_Code,MAX(Location_Desc) AS Location_Desc,SUM(Total_Amt) AS Total_Amt,MAX(Email) AS Email,MAX(Mobile_no) AS Mobile_no FROM FilteredData f GROUP BY Customer_Code, Supply_Date, Shift_type"
            Else
                FinalQry = "Select Cast(0 as BIT) as 'Check',Document_Code,Max([Item Type])[Item Type],Max(Document_Date)Document_Date,Max(Zone_Code)Zone_Code,Max(Route_Code)Route_Code,Max(Customer_Code)Customer_Code,Max(Customer_Name)Customer_Name,Max(Location_Desc)Location_Desc,Max(Total_Amt)Total_Amt,Max(Shift_type)Shift_type,Max(Supply_Date)Supply_Date,Max(Email)Email,Max(Mobile_no)Mobile_no,Max(Is_Taxable)Is_Taxable from (" + sQuery + ")BaseQry Group By Document_Code  order by Document_Date "
            End If


            Dim dtgv As New DataTable
            dtgv = clsDBFuncationality.GetDataTable(FinalQry)
            If dtgv IsNot Nothing AndAlso dtgv.Rows.Count > 0 Then
                gv.DataSource = Nothing
                gv.Rows.Clear()
                gv.Columns.Clear()
                gv.DataSource = dtgv
                gv.GroupDescriptors.Clear()
                gv.MasterTemplate.SummaryRowsBottom.Clear()
                FormatGrid()

                RadPageView1.SelectedPage = RadPageViewPage2
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
            End If
            ReStoreGridLayout()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gv.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gv.Columns.Count - 1 Step ii + 1
                        gv.Columns(ii).IsVisible = False
                        gv.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub
    Sub FormatGrid()
        gv.TableElement.TableHeaderHeight = 25
        gv.MasterTemplate.ShowRowHeaderColumn = False
        For ii As Integer = 0 To gv.Columns.Count - 1
            gv.Columns(ii).ReadOnly = True
            gv.Columns(ii).IsVisible = False
        Next

        gv.Columns("Check").IsVisible = True
        gv.Columns("Check").Width = 100
        gv.Columns("Check").HeaderText = " "
        gv.Columns("Check").ReadOnly = False

        If EnableProductSaleForJPR Then
            If Not rbtnPartyWise.Checked Then
                gv.Columns("Item Type").IsVisible = True
            End If
        End If

        If Not rbtnPartyWise.Checked Then
            gv.Columns("Document_Code").IsVisible = True
            gv.Columns("Document_Code").Width = 100
            gv.Columns("Document_Code").HeaderText = "Sale Invoice No."

            gv.Columns("Document_Date").IsVisible = True
            gv.Columns("Document_Date").Width = 100
            gv.Columns("Document_Date").HeaderText = " Date"
            'gv.Columns("Document_Date").FormatString = "{0:d}"

            gv.Columns("Zone_Code").IsVisible = True
            gv.Columns("Zone_Code").Width = 100
            gv.Columns("Zone_Code").HeaderText = "Zone Code"
        End If


        gv.Columns("Route_Code").IsVisible = True
        gv.Columns("Route_Code").Width = 100
        gv.Columns("Route_Code").HeaderText = "Route Code"

        gv.Columns("Location_Desc").IsVisible = True
        gv.Columns("Location_Desc").Width = 100
        gv.Columns("Location_Desc").HeaderText = "Location"

        gv.Columns("Customer_Code").IsVisible = True
        gv.Columns("Customer_Code").Width = 100
        gv.Columns("Customer_Code").HeaderText = "Customer Code"

        gv.Columns("Customer_Name").IsVisible = True
        gv.Columns("Customer_Name").Width = 150
        gv.Columns("Customer_Name").HeaderText = "Customer"

        gv.Columns("Total_Amt").IsVisible = True
        gv.Columns("Total_Amt").Width = 100
        gv.Columns("Total_Amt").HeaderText = "Amount"

        gv.Columns("Supply_Date").IsVisible = True
        gv.Columns("Supply_Date").Width = 100
        gv.Columns("Supply_Date").HeaderText = "Supply Date"

        gv.Columns("Shift_Type").IsVisible = True
        gv.Columns("Shift_Type").Width = 100
        gv.Columns("Shift_Type").HeaderText = "Shift"

        gv.Columns("Email").IsVisible = True
        gv.Columns("Email").Width = 200
        gv.Columns("Email").HeaderText = "Email"

        gv.Columns("Mobile_no").IsVisible = True
        gv.Columns("Mobile_no").Width = 100
        gv.Columns("Mobile_no").HeaderText = "Mobile no"

        'If Not rbtnInvoiceWise.Checked AndAlso clsCommon.CompairString(cboReportType.SelectedValue, "AL") = CompairStringResult.Equal Then
        '    gv.Columns("IsTaxable").IsVisible = True
        '    gv.Columns("IsTaxable").Width = 100
        '    gv.Columns("IsTaxable").HeaderText = "Taxable/Non-Taxable"
        'End If

        'gv.Columns("DespatchDocumentNo").IsVisible = False
        'gv.Columns("DespatchDocumentNo").Width = 100
        'gv.Columns("DespatchDocumentNo").HeaderText = "DespatchDocumentNo"

        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim intCount As Integer = 0

        Dim item1 As New GridViewSummaryItem("Total_Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item1)

        gv.ShowGroupPanel = False
        gv.MasterTemplate.AutoExpandGroups = True
        gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Try
            PageSetupReport_ID = MyBase.Form_ID
            TemplateGridview = gv
            loadReport()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub Reset()
        rbtnDocumentDate.Checked = True
        RadGroupBox2.Visible = True
        RadGroupBox4.Visible = False
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = txtToDate.Value.AddMonths(-1)
        'LoadLocation()
        'chkLocationAll.CheckState = CheckState.Checked
        txtMultCustomer.arrValueMember = Nothing
        txtMultRoute.arrValueMember = Nothing
        txtMultLocation.arrValueMember = Nothing
        'cboReportType.SelectedValue = ""
        btnPrePrintFormat.Visible = False
        gv.DataSource = Nothing
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "BHAD") = CompairStringResult.Equal Then
            btnBatchWiseInvoice.Visible = True
        Else
            btnBatchWiseInvoice.Visible = False
        End If
        fndCustom.Value = ""
        lblCustomer.Text = ""
        If ApplyMilkPouchPrint = True Then
            MyLabel1.Visible = False
            txtMultCustomer.Visible = False
            MyLabel4.Visible = False
            cboReportType.Visible = False
            lblLocation.Visible = False
            txtMultLocation.Visible = False
            lblRoute.Visible = False
            txtMultRoute.Visible = False
            btnGo.Visible = False
            btnUnSelect.Visible = False
            RadSplitButton1.Visible = False
            BtnPrintChallan.Visible = False
            BtnEmailSms.Visible = False
            btnPrePrintFormat.Visible = False
            btnBatchWiseInvoice.Visible = False
            btnCombinedInvoice.Visible = False
            BtnPrint.Visible = True
            BtnPrint.Enabled = True
            BtnReset.Visible = True
            BtnReset.Enabled = True
        Else
            MyLabel2.Visible = False
            fndCustom.Visible = False
            lblCustomer.Visible = False
        End If
        RadPageView1.SelectedPage = RadPageViewPage1
    End Sub
    Private Sub BtnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReset.Click
        Reset()
    End Sub

    Sub LoadReportType()
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))

        Dim dr As DataRow = dt.NewRow()
        dr = dt.NewRow()
        dr("Code") = ""
        dr("Name") = "-Select-"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "T"
        dr("Name") = "Taxable"
        dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "LWM"
        'dr("Name") = "Local With MandiTax"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "IT"
        'dr("Name") = "Interstate Taxable"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "IWM"
        'dr("Name") = "Interstate With MandiTax"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "UT"
        'dr("Name") = "UT Taxable"
        'dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "NT"
        dr("Name") = "NonTaxable"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "AL"
        dr("Name") = "ALL"
        dt.Rows.Add(dr)

        cboReportType.DataSource = dt
        cboReportType.ValueMember = "Code"
        cboReportType.DisplayMember = "Name"
        cboReportType.SelectedValue = ""
    End Sub

    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        Try
            If ApplyMilkPouchPrint Then
                Try
                    'If txtFromDate.ToString("yyyyMM") < txtToDate.ToString("yyyyMM") Then
                    '    clsCommon.MyMessageBoxShow("From Date and To Date should be same month of year.", Me.Text)
                    '    Return
                    'ElseIf txtFromDate.ToString("yyyyMM") > txtToDate.ToString("yyyyMM") Then
                    '    clsCommon.MyMessageBoxShow("From Date and To Date should be same month of year.", Me.Text)
                    '    Return
                    'End If
                    'clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy")
                    If clsCommon.CompairString(clsCommon.GetPrintDate(txtFromDate.Value, "MMM-yyyy"), clsCommon.GetPrintDate(txtToDate.Value, "MMM-yyyy")) <> CompairStringResult.Equal Then
                        clsCommon.MyMessageBoxShow(Me, "From Date and To Date should be same month of year.", Me.Text)
                        Return
                    End If

                    If clsCommon.myLen(fndCustom.Value) <= 0 Then
                        clsCommon.MyMessageBoxShow(Me, "Please Select Customer First", Me.Text)
                        Return
                    End If

                    Dim strMonth As String = clsCommon.GetPrintDate(txtFromDate.Value, "MMM-yyyy")
                    Dim strCurrentDate As String = clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd-MMM-yyyy")
                    Dim strFromDate As String = clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy")
                    Dim strToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy")
                    Dim strMonthlyQry As String = " select top 1 isnull (TSPL_SD_SALE_INVOICE_HEAD.MonthlySaleInvoiceNo,'')   FROM   TSPL_SD_SALE_INVOICE_DETAIL 
                                                left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code  
                                                inner join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code  
                                                where  1=1 "
                    If chkDCSSale.Checked Then
                        strMonthlyQry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS','MCC') "
                    Else
                        strMonthlyQry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND Screen_Type='DS' "
                    End If
                    strMonthlyQry &= " And convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103) and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code= '" + fndCustom.Value + "' and TSPL_ITEM_MASTER.Is_Milk_Pouch =1  and TSPL_SD_SALE_INVOICE_HEAD.Status =1 and len( isnull (TSPL_SD_SALE_INVOICE_HEAD.MonthlySaleInvoiceNo,'')) > 0 "
                    Dim ExistMonthlySaleInvoiceNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strMonthlyQry))
                    If clsCommon.myLen(ExistMonthlySaleInvoiceNo) > 0 Then
                    Else
                        ExistMonthlySaleInvoiceNo = clsERPFuncationality.GetNextCode(Nothing, clsCommon.GETSERVERDATE(), clsDocType.rptMonthlySalesInvoice, "", "")

                    End If
                    Dim qryUpdate As String = "  update  TSPL_SD_SALE_INVOICE_HEAD  set MonthlySaleInvoiceNo = '" + ExistMonthlySaleInvoiceNo + "'  where  Document_Code in (
                                             select distinct TSPL_SD_SALE_INVOICE_HEAD.Document_Code   FROM   TSPL_SD_SALE_INVOICE_DETAIL 
                                             left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code  
                                             inner join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code  
                                             where  1=1 "
                    If chkDCSSale.Checked Then
                        qryUpdate &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS','MCC') "
                    Else
                        qryUpdate &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND Screen_Type='DS'  "
                    End If
                    qryUpdate &=" And convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103) and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code= '" + fndCustom.Value + "' and TSPL_ITEM_MASTER.Is_Milk_Pouch =1  and TSPL_SD_SALE_INVOICE_HEAD.Status =1 and  len (isnull (MonthlySaleInvoiceNo,'')) <=0 ) "
                    clsDBFuncationality.ExecuteNonQuery(qryUpdate)
                    ' strMonth, strCurrentDate , strFromDate,strToDate 
                    Dim qry As String = "  select '" + strMonth + "' as Month , '" + strCurrentDate + "' as CurrentDate , '" + strFromDate + "' as FromDate , '" + strToDate + "' as ToDate, '" + ExistMonthlySaleInvoiceNo + "' as MonthlySaleInvoiceNo,  max(Comp_Code) as Comp_Code,  max(Comp_Name ) as Comp_Name ,  max(Comp_Add1) as Comp_Add1, max(Comp_Add2) as Comp_Add2,max( Comp_Add3) as Comp_Add3, max(Comp_City_Code ) as Comp_City_Code, max(Comp_PinCode) as Comp_PinCode ,max(Comp_Phone1) as Comp_Phone1 ,max(Comp_Phone2) as Comp_Phone2 ,max(Comp_GSTINNo) as Comp_GSTINNo, max(Comp_StateCode) as Comp_StateCode ,max( Comp_StateName) as Comp_StateName , max(Cust_Code) as Cust_Code ,max(Customer_Name) as Customer_Name , max(Cust_Add1) as Cust_Add1 ,max(Cust_Add2) as Cust_Add2,max(Cust_Add3) as Cust_Add3 ,max(Cust_GSTNO) as Cust_GSTNO, max(Cust_StateCode) as  Cust_StateCode,  max(Cust_State_Name) as Cust_State_Name , max(Route_No) as Route_No,  max(Route_Desc) as Route_Desc, max(HSN_Code) as HSN_Code,Doc_Date  ,
 
                                         sum (isnull([Item1_Qty],0)) as [Item1_Qty],
                                         sum(isnull([Item1_Amt],0)) as [Item1_Amt],
                                         isnull(convert (decimal(18,2), (sum(isnull([Item1_Amt],0)) / nullif (sum (isnull([Item1_Qty],0)),0))),0) as [Item1_Rate],
 
                                         sum (isnull([Item2_Qty],0)) as [Item2_Qty],
                                         sum(isnull([Item2_Amt],0)) as [Item2_Amt],
                                         isnull(convert (decimal(18,2), (sum(isnull([Item2_Amt],0)) / nullif (sum (isnull([Item2_Qty],0)),0))),0) as [Item2_Rate],
 
                                         sum (isnull([Item3_Qty],0)) as [Item3_Qty],
                                         sum(isnull([Item3_Amt],0)) as [Item3_Amt],
                                         isnull(convert (decimal(18,2), (sum(isnull([Item3_Amt],0)) / nullif (sum (isnull([Item3_Qty],0)),0))),0) as [Item3_Rate],
 
                                          sum (isnull([Item4_Qty],0)) as [Item4_Qty],
                                         sum(isnull([Item4_Amt],0)) as [Item4_Amt],
                                         isnull( convert (decimal(18,2), (sum(isnull([Item4_Amt],0)) / nullif (sum (isnull([Item4_Qty],0)),0))),0) as [Item4_Rate],

                                          sum (isnull([Item5_Qty],0)) as [Item5_Qty],
                                         sum(isnull([Item5_Amt],0)) as [Item5_Amt],
                                         isnull( convert (decimal(18,2), (sum(isnull([Item5_Amt],0)) / nullif (sum (isnull([Item3_Qty],0)),0))),0) as [Item5_Rate]

                                         ,max(TTT.Item1) as Item1 , max(TTT.Item2) as Item2, max(TTT.Item3 ) as Item3, max(TTT.Item4) as Item4 , max(TTT.Item5) as Item5
 
 
                                         from (
                                         select 1 as SNO, Comp_Code,  Comp_Name ,  Comp_Add1, Comp_Add2, Comp_Add3, Comp_City_Code , Comp_PinCode ,Comp_Phone1 ,Comp_Phone2 ,Comp_GSTINNo, Comp_StateCode , Comp_StateName , Cust_Code ,Customer_Name , Cust_Add1,Cust_Add2,Cust_Add3,Cust_GSTNO, Cust_StateCode ,  Cust_State_Name , Route_No,  Route_Desc, HSN_Code,Doc_Date  ,
 
                                         [Item1_Qty],[Item2_Qty],[Item3_Qty],[Item4_Qty] ,[Item5_Qty], 
 
                                         [Item1_Amt],[Item2_Amt],[Item3_Amt],[Item4_Amt],[Item5_Amt]
 
 
                                         from (
                                         select max(XXXFinal.Comp_Code) as Comp_Code, max( XXXFinal.Comp_Name) as Comp_Name , max( Comp_Add1) as Comp_Add1,max(Comp_Add2) as Comp_Add2,max(Comp_Add3) as Comp_Add3,max(Comp_City_Code) as Comp_City_Code ,max(Comp_PinCode) as Comp_PinCode ,max(Comp_Phone1) as Comp_Phone1 ,max(Comp_Phone2) as Comp_Phone2 ,max(Comp_GSTINNo) as Comp_GSTINNo, max(Comp_StateCode) as Comp_StateCode , max( Comp_StateName) as Comp_StateName , Cust_Code ,max( Customer_Name) as Customer_Name , max(Cust_Add1) as Cust_Add1,max( Cust_Add2 ) as Cust_Add2,max(Cust_Add3) as Cust_Add3,max(Cust_GSTNO) as Cust_GSTNO, max(Cust_StateCode) as Cust_StateCode , max(Cust_State_Name) as Cust_State_Name ,max( Route_No) as Route_No, max(Route_Desc) as Route_Desc, Doc_Date, Item_Code , max(Item_Desc) as Item_Desc, max(Item_Desc2) as Item_Desc2 , max(HSN_Code) as HSN_Code,sum( QtyInLtr) as QtyInLtr , sum (Item_Net_Amt) as Item_Net_Amt , convert (decimal(18,2), sum (Item_Net_Amt) / nullif (sum( QtyInLtr),0) )  as Rate from (
                                         select  TSPL_COMPANY_MASTER.Comp_Code , TSPL_COMPANY_MASTER.Comp_Name , TSPL_COMPANY_MASTER.Add1 as Comp_Add1, TSPL_COMPANY_MASTER.Add2 as Comp_Add2, TSPL_COMPANY_MASTER.Add3 as Comp_Add3, TSPL_COMPANY_MASTER.City_Code as Comp_City_Code, TSPL_COMPANY_MASTER.Pincode as Comp_PinCode ,TSPL_COMPANY_MASTER.Phone1 as Comp_Phone1, TSPL_COMPANY_MASTER.Phone2 as Comp_Phone2,TSPL_COMPANY_MASTER.GSTINNo as Comp_GSTINNo,TSPL_STATE_MASTER_comp.GST_STATE_Code as Comp_StateCode, TSPL_STATE_MASTER_comp.STATE_NAME as Comp_StateName, TSPL_CUSTOMER_MASTER.Cust_Code, TSPL_CUSTOMER_MASTER.Customer_Name , TSPL_CUSTOMER_MASTER.Add1 as Cust_Add1 , TSPL_CUSTOMER_MASTER.Add2 as Cust_Add2, TSPL_CUSTOMER_MASTER.Add3 as Cust_Add3, TSPL_CUSTOMER_MASTER.GSTNO as Cust_GSTNO, TSPL_STATE_MASTER_Cust.GST_STATE_Code as Cust_StateCode, TSPL_STATE_MASTER_Cust.STATE_NAME as Cust_State_Name,TSPL_CUSTOMER_MASTER.Route_No, TSPL_ROUTE_MASTER.Route_Desc , convert (varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,6) as Doc_Date,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code ,XXFinal.Item + '_Qty' as Item_Desc, XXFinal.Item + '_Amt' as Item_Desc2 ,TSPL_item_master.HSN_Code,TSPL_SD_SALE_INVOICE_DETAIL.Qty, TSPL_SD_SALE_INVOICE_DETAIL.Unit_code , convert (decimal(18,2), TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull( StockUOM.Conversion_Factor,0) / nullif( isnull (TargetUOM.Conversion_Factor,0),0) ) as QtyInLtr , TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt , TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt / (TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull( StockUOM.Conversion_Factor,0) / nullif (isnull (TargetUOM.Conversion_Factor,0),0) ) as Rate
                                         from TSPL_SD_SALE_INVOICE_DETAIL 
                                         left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
                                         left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = TSPL_SD_SALE_INVOICE_HEAD.Comp_Code
                                         left outer join TSPL_STATE_MASTER as TSPL_STATE_MASTER_comp on TSPL_STATE_MASTER_comp.STATE_CODE = TSPL_COMPANY_MASTER.State
                                         left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                                         left outer join TSPL_STATE_MASTER as TSPL_STATE_MASTER_Cust on TSPL_STATE_MASTER_Cust.STATE_CODE = TSPL_CUSTOMER_MASTER.State
                                         left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_CUSTOMER_MASTER.Route_No
                                         left outer join TSPL_ITEM_UOM_DETAIL as StockUOM on StockUOM.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and  StockUOM.UOM_Code  = TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
                                         left outer join TSPL_ITEM_UOM_DETAIL as TargetUOM on TargetUOM.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TargetUOM.UOM_Code = 'Ltr'
                                         inner join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
 
                                         left outer join (  
                                         select Case when ROW_NUMBER = 1 then 'Item1'   when ROW_NUMBER = 2 then 'Item2'  when ROW_NUMBER = 3 then 'Item3' when ROW_NUMBER = 4 then 'Item4' when ROW_NUMBER = 5 then 'Item5' end as Item, Item_Desc , Item_Code     from (
                                         select ROW_NUMBER() OVER(ORDER BY Item_Desc) AS Row_Number , Item_Desc, Item_Code  from (
                                         SELECT Distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code, TSPL_ITEM_MASTER .Alies_Name as Item_Desc   FROM   TSPL_SD_SALE_INVOICE_DETAIL 
                                         left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
                                         left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = TSPL_SD_SALE_INVOICE_HEAD.Comp_Code
                                         left outer join TSPL_STATE_MASTER as TSPL_STATE_MASTER_comp on TSPL_STATE_MASTER_comp.STATE_CODE = TSPL_COMPANY_MASTER.State
                                         left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                                         left outer join TSPL_STATE_MASTER as TSPL_STATE_MASTER_Cust on TSPL_STATE_MASTER_Cust.STATE_CODE = TSPL_CUSTOMER_MASTER.State
                                         left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_CUSTOMER_MASTER.Route_No
                                         left outer join TSPL_ITEM_UOM_DETAIL as StockUOM on StockUOM.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and  StockUOM.UOM_Code  = TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
                                         left outer join TSPL_ITEM_UOM_DETAIL as TargetUOM on TargetUOM.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TargetUOM.UOM_Code = 'Ltr'
                                         inner join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code  
                                         where 1=1 "
                    If chkDCSSale.Checked Then
                        qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS','MCC') "
                    Else
                        qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND Screen_Type='DS' "
                    End If
                    qry &= " and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103) and TSPL_CUSTOMER_MASTER.Cust_Code= '" + fndCustom.Value + "' and TSPL_ITEM_MASTER.Is_Milk_Pouch =1  and TSPL_SD_SALE_INVOICE_HEAD.Status =1
                                         ) Final ) XFinal ) XXFinal  on XXFinal.Item_Desc = TSPL_ITEM_MASTER.Alies_Name and XXFinal.Item_Code = TSPL_ITEM_MASTER.Item_Code
                                         where 1=1 "
                    If chkDCSSale.Checked Then
                        qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS','MCC') "
                    Else
                        qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND Screen_Type='DS' "
                    End If
                    qry &= " and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103) and TSPL_CUSTOMER_MASTER.Cust_Code= '" + fndCustom.Value + "' and TSPL_ITEM_MASTER.Is_Milk_Pouch =1  and TSPL_SD_SALE_INVOICE_HEAD.Status =1
                                         ) XXXFinal group by XXXFinal.Doc_Date, XXXFinal.Cust_Code, XXXFinal.Item_Code
                                         ) Final
                                         pivot ( sum(QtyInLtr) for Item_Desc in ([Item1_Qty],[Item2_Qty],[Item3_Qty],[Item4_Qty],[Item5_Qty]) ) QtyPivot 

                                         pivot ( sum(Item_Net_Amt) for Item_Desc2 in ([Item1_Amt],[Item2_Amt],[Item3_Amt],[Item4_Amt],[Item5_Amt]) ) QtyPivot2

                                          ) XFinal 

 
                                          left outer join  ( select 1 as SNO, max(isnull(Item1,'')) as Item1, max(isnull(Item2,'')) as Item2 , max(isnull(Item3,'')) as Item3, max(isnull(Item4,'')) as Item4, max(isnull(Item5,'')) as Item5 from (
                                         select Case when ROW_NUMBER = 1 then Item_Desc end as Item1,Case when ROW_NUMBER = 2 then Item_Desc end as Item2,Case when ROW_NUMBER = 3 then Item_Desc end as Item3, Case when ROW_NUMBER = 4 then Item_Desc end as Item4, Case when ROW_NUMBER = 5 then Item_Desc end as Item5     from (
                                         select ROW_NUMBER() OVER(ORDER BY Item_Desc) AS Row_Number , Item_Desc from (
                                         SELECT Distinct TSPL_ITEM_MASTER.Alies_Name as  Item_Desc  FROM   TSPL_SD_SALE_INVOICE_DETAIL 
                                         left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
                                         left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = TSPL_SD_SALE_INVOICE_HEAD.Comp_Code
                                         left outer join TSPL_STATE_MASTER as TSPL_STATE_MASTER_comp on TSPL_STATE_MASTER_comp.STATE_CODE = TSPL_COMPANY_MASTER.State
                                         left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                                         left outer join TSPL_STATE_MASTER as TSPL_STATE_MASTER_Cust on TSPL_STATE_MASTER_Cust.STATE_CODE = TSPL_CUSTOMER_MASTER.State
                                         left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_CUSTOMER_MASTER.Route_No
                                         left outer join TSPL_ITEM_UOM_DETAIL as StockUOM on StockUOM.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and  StockUOM.UOM_Code  = TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
                                         left outer join TSPL_ITEM_UOM_DETAIL as TargetUOM on TargetUOM.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TargetUOM.UOM_Code = 'Ltr'
                                         inner join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code 
                                          where 1=1 "
                    If chkDCSSale.Checked Then
                        qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS','MCC') "
                    Else
                        qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type IN ('FS','PS') AND Screen_Type='DS' "
                    End If
                    qry &= "  and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103) and TSPL_CUSTOMER_MASTER.Cust_Code= '" + fndCustom.Value + "' and TSPL_ITEM_MASTER.Is_Milk_Pouch =1  and TSPL_SD_SALE_INVOICE_HEAD.Status =1
                                         ) Final ) xfinal ) as XXFinal) TTT on TTT.SNO = XFinal.SNO
                                          group by Doc_Date, Cust_Code "

                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dt.Rows.Count > 0 Then
                        Dim frmCRV As New frmCrystalReportViewer()
                        frmCRV.funreport(MyBase.Form_ID, False, CrystalReportFolder.SalesReport, dt, "rptMonthlyInvoicePrint", "Customer Monthly Sales")
                    Else
                        common.clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
                    End If
                Catch ex As Exception
                    common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error, Me.Text)
                End Try
            Else
                checkSendEmailSMS(False)
                Printing()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function CheckEmailOrSMS(ByVal InvCode As String)
        Dim strChkEmailSMS As String = "Select SUM(AlreadySendEmailParty)AlreadySendEmailParty,SUM(RemainingParty)RemainingParty,SUM(AlreadySendSMSParty)AlreadySendSMSParty,SUM(RemainingSMSParty)RemainingSMSParty from(Select Customer_Code,Sum(Distinct Case When IsNUll(Send_Email,0)>0 Then 1 Else 0 End)AlreadySendEmailParty,Sum(Distinct Case When IsNUll(Send_Email,0)=0 Then 1 Else 0 End)RemainingParty,Sum(Distinct Case When IsNUll(Send_SMS,0)>0 Then 1 Else 0 End)AlreadySendSMSParty,Sum(Distinct Case When IsNUll(Send_SMS,0)=0 Then 1 Else 0 End)RemainingSMSParty from( Select TSPL_SD_SALE_INVOICE_HEAD.Document_Code,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_SD_SALE_INVOICE_HEAD.Send_Email,TSPL_SD_SALE_INVOICE_HEAD.Send_SMS from TSPL_SD_SALE_INVOICE_HEAD
Inner Join TSPL_SD_SHIPMENT_HEAD On TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No=TSPL_SD_SALE_INVOICE_HEAD.Document_Code
Where TSPL_SD_SALE_INVOICE_HEAD.Document_Code In (" + InvCode + "))xyz Group By Customer_Code)final"
        Return strChkEmailSMS
    End Function

    Private Function Printing(Optional ByVal DocNo As String = "", Optional ByVal isPdf As Boolean = False, Optional ByVal isPrePrintFormat As Boolean = False, Optional ByVal isPdfForMail As Boolean = False, Optional ByVal SupplyDate As String = Nothing, Optional ByVal Shift As String = Nothing, Optional ByVal isTaxable As String = Nothing, Optional ByVal strPartyCode As String = Nothing) As String
        Dim pdfPath As String = Nothing
        Dim ii As Integer = 1
        Dim Total As Integer = 0

        Dim Qry As String = Nothing
        'Dim frmCRV As New frmCrystalReportViewer()
        Dim objMultPrintInvoice As New FrmPrintFreshInvoice
        ' Dim strInvoice As String
        ' If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then

        If DocNo IsNot Nothing AndAlso clsCommon.myLen(DocNo) > 0 AndAlso isPdfForMail = True Then
            lstinvNo = New List(Of String)
            Total = 1
            lstinvNo.Add(DocNo)
        Else
            For Each grow As GridViewRowInfo In gv.Rows
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    Total += 1
                End If
            Next
            If rbtnInvoiceWise.Checked Then
                lstinvNo = New List(Of String)
                For Each grow As GridViewRowInfo In gv.Rows
                    'clsCommon.ProgressBarPercentUpdate((ii) * 100 / gv.Rows.Count, " Send Email " & (ii) & " Of " & gv.Rows.Count)
                    If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                        'clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Processing..." & (ii) & " Of " & Total)
                        lstinvNo.Add(clsCommon.myCstr(grow.Cells("Document_Code").Value))
                        ii += 1
                    End If
                    'ii += 1
                Next
            Else
                If Not isPdf Then

                End If
            End If
        End If

        Try
            If lstinvNo Is Nothing AndAlso lstinvNo.Count <= 0 Then
                clsCommon.ProgressBarPercentHide()
                clsCommon.MyMessageBoxShow(Me, "Invoice not found to Print", Me.Text)
            Else
                Dim dtDocdate As Date?
                dtDocdate = Nothing
                Dim InvoiceNO As String = clsCommon.GetMulcallString(lstinvNo)
                Dim StrSql As String = "Select Document_Code,Document_Date,Customer_Code,Bill_To_Location,is_taxable,Tax_Group from TSPL_SD_SALE_INVOICE_HEAD where Document_Code in(" + InvoiceNO + ")"
                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(StrSql)
                If dt1.Rows.Count > 0 Then
                    'IsTaxable = clsCommon.myCdbl(dt1.Rows(0)("is_taxable"))
                    dtDocdate = clsCommon.myCDate(dt1.Rows(0)("Document_Date"))
                End If
                Dim frmCRV As New frmCrystalReportViewer()
                ' Dim InvoiceNO As String = clsCommon.GetMulcallString(lstinvNo)
                Dim ItemType As String = ""
                If EnableProductSaleForJPR Then
                    If rbtnMilk.Checked Then
                        ItemType = " 'S','' "
                    ElseIf rbtnProduct.Checked Then
                        ItemType = " 'P' "
                    ElseIf rbtnIceCream.Checked Then
                        ItemType = " 'I' "
                    End If
                End If

                If rbtnPartyWise.Checked Then
                    Dim BaseQry As String = "Select Max(Report_Status) As Report_Status,Max(Is_Distributor) As Is_Distributor,Max(Is_BPL) As Is_BPL,Max(Is_CashSale) As Is_CashSale,Max(Is_DCS) As Is_DCS,Max(Booking_Type) As Booking_Type,Max(xyz.CST_LST) As CST_LST,Max(Manual_VehicleNo) As Manual_VehicleNo,Max(Payment_Terms) As Payment_Terms,Max(ReceiverName) As ReceiverName,Sum(Security_TotalAmt) As Security_TotalAmt,Max(Supply_Date) As Supply_Date,Max(Shift_Type) As Shift_Type,Sum(QTY_LTRKG) As QTY_LTRKG,"
                    BaseQry += " Case When Max(ITAX1) In ('KKF','MNDTAX') Then MAX(ITAX1_Base_Amt) Else 0 End +
Case When Max(ITAX2) In ('KKF','MNDTAX') Then MAX(ITAX2_Base_Amt) Else 0 End +
Case When Max(ITAX3) In ('KKF','MNDTAX') Then MAX(ITAX3_Base_Amt) Else 0 End +
Case When Max(ITAX4) In ('KKF','MNDTAX') Then MAX(ITAX4_Base_Amt) Else 0 End +
Case When Max(ITAX5) In ('KKF','MNDTAX') Then MAX(ITAX5_Base_Amt) Else 0 End +
Case When Max(ITAX6) In ('KKF','MNDTAX') Then MAX(ITAX6_Base_Amt) Else 0 End +
Case When Max(ITAX7) In ('KKF','MNDTAX') Then MAX(ITAX7_Base_Amt) Else 0 End +
Case When Max(ITAX8) In ('KKF','MNDTAX') Then MAX(ITAX8_Base_Amt) Else 0 End +
Case When Max(ITAX9) In ('KKF','MNDTAX') Then MAX(ITAX9_Base_Amt) Else 0 End +
Case When Max(ITAX10) In ('KKF','MNDTAX') Then MAX(ITAX10_Base_Amt) Else 0 End As 'BaseAmt',
Case When Max(ITAX1)='KKF' Then Sum(ITAX1_Amt) Else 0 End +
Case When Max(ITAX2)='KKF' Then Sum(ITAX2_Amt) Else 0 End +
Case When Max(ITAX3)='KKF' Then Sum(ITAX3_Amt) Else 0 End +
Case When Max(ITAX4)='KKF' Then Sum(ITAX4_Amt) Else 0 End +
Case When Max(ITAX5)='KKF' Then Sum(ITAX5_Amt) Else 0 End +
Case When Max(ITAX6)='KKF' Then Sum(ITAX6_Amt) Else 0 End +
Case When Max(ITAX7)='KKF' Then Sum(ITAX7_Amt) Else 0 End +
Case When Max(ITAX8)='KKF' Then Sum(ITAX8_Amt) Else 0 End +
Case When Max(ITAX9)='KKF' Then Sum(ITAX9_Amt) Else 0 End +
Case When Max(ITAX10)='KKF' Then Sum(ITAX10_Amt) Else 0 End As 'KKF',
Case When Max(ITAX1)='MNDTAX' Then Sum(ITAX1_Amt) Else 0 End +
Case When Max(ITAX2)='MNDTAX' Then Sum(ITAX2_Amt) Else 0 End +
Case When Max(ITAX3)='MNDTAX' Then Sum(ITAX3_Amt) Else 0 End +
Case When Max(ITAX4)='MNDTAX' Then Sum(ITAX4_Amt) Else 0 End +
Case When Max(ITAX5)='MNDTAX' Then Sum(ITAX5_Amt) Else 0 End +
Case When Max(ITAX6)='MNDTAX' Then Sum(ITAX6_Amt) Else 0 End +
Case When Max(ITAX7)='MNDTAX' Then Sum(ITAX7_Amt) Else 0 End +
Case When Max(ITAX8)='MNDTAX' Then Sum(ITAX8_Amt) Else 0 End +
Case When Max(ITAX9)='MNDTAX' Then Sum(ITAX9_Amt) Else 0 End +
Case When Max(ITAX10)='MNDTAX' Then Sum(ITAX10_Amt) Else 0 End As 'MNDTAX',
Case When Max(ITAX1) In ('SGST','CGST','IGST') Then Max(ITAX1_Base_Amt)
When Max(ITAX2) In ('SGST','CGST','IGST') Then Max(ITAX2_Base_Amt)
When Max(ITAX3) In ('SGST','CGST','IGST') Then Max(ITAX3_Base_Amt)
When Max(ITAX4) In ('SGST','CGST','IGST') Then Max(ITAX4_Base_Amt)
When Max(ITAX5) In ('SGST','CGST','IGST') Then Max(ITAX5_Base_Amt)
When Max(ITAX6) In ('SGST','CGST','IGST') Then Max(ITAX6_Base_Amt)
When Max(ITAX7) In ('SGST','CGST','IGST') Then Max(ITAX7_Base_Amt)
When Max(ITAX8) In ('SGST','CGST','IGST') Then Max(ITAX8_Base_Amt)
When Max(ITAX9) In ('SGST','CGST','IGST') Then Max(ITAX9_Base_Amt)
When Max(ITAX10) In ('SGST','CGST','IGST') Then Max(ITAX10_Base_Amt) Else 0 End As 'TaxableAmount', 
Case When Max(ITAX1)='SGST' Then Sum(ITAX1_Amt) Else 0 End +
Case When Max(ITAX2)='SGST' Then Sum(ITAX2_Amt) Else 0 End +
Case When Max(ITAX3)='SGST' Then Sum(ITAX3_Amt) Else 0 End +
Case When Max(ITAX4)='SGST' Then Sum(ITAX4_Amt) Else 0 End +
Case When Max(ITAX5)='SGST' Then Sum(ITAX5_Amt) Else 0 End +
Case When Max(ITAX6)='SGST' Then Sum(ITAX6_Amt) Else 0 End +
Case When Max(ITAX7)='SGST' Then Sum(ITAX7_Amt) Else 0 End +
Case When Max(ITAX8)='SGST' Then Sum(ITAX8_Amt) Else 0 End +
Case When Max(ITAX9)='SGST' Then Sum(ITAX9_Amt) Else 0 End +
Case When Max(ITAX10)='SGST' Then Sum(ITAX10_Amt) Else 0 End As 'SGST',
Case When Max(ITAX1)='CGST' Then Sum(ITAX1_Amt) Else 0 End +
Case When Max(ITAX2)='CGST' Then Sum(ITAX2_Amt) Else 0 End +
Case When Max(ITAX3)='CGST' Then Sum(ITAX3_Amt) Else 0 End +
Case When Max(ITAX4)='CGST' Then Sum(ITAX4_Amt) Else 0 End +
Case When Max(ITAX5)='CGST' Then Sum(ITAX5_Amt) Else 0 End +
Case When Max(ITAX6)='CGST' Then Sum(ITAX6_Amt) Else 0 End +
Case When Max(ITAX7)='CGST' Then Sum(ITAX7_Amt) Else 0 End +
Case When Max(ITAX8)='CGST' Then Sum(ITAX8_Amt) Else 0 End +
Case When Max(ITAX9)='CGST' Then Sum(ITAX9_Amt) Else 0 End +
Case When Max(ITAX10)='CGST' Then Sum(ITAX10_Amt) Else 0 End As 'CGST',
Case When Max(ITAX1)='IGST' Then Sum(ITAX1_Amt) Else 0 End +
Case When Max(ITAX2)='IGST' Then Sum(ITAX2_Amt) Else 0 End +
Case When Max(ITAX3)='IGST' Then Sum(ITAX3_Amt) Else 0 End +
Case When Max(ITAX4)='IGST' Then Sum(ITAX4_Amt) Else 0 End +
Case When Max(ITAX5)='IGST' Then Sum(ITAX5_Amt) Else 0 End +
Case When Max(ITAX6)='IGST' Then Sum(ITAX6_Amt) Else 0 End +
Case When Max(ITAX7)='IGST' Then Sum(ITAX7_Amt) Else 0 End +
Case When Max(ITAX8)='IGST' Then Sum(ITAX8_Amt) Else 0 End +
Case When Max(ITAX9)='IGST' Then Sum(ITAX9_Amt) Else 0 End +
Case When Max(ITAX10)='IGST' Then Sum(ITAX10_Amt) Else 0 End As 'IGST',
Case When Max(ITAX1)='TCS' Then Sum(ITAX1_Amt) Else 0 End +
Case When Max(ITAX2)='TCS' Then Sum(ITAX2_Amt) Else 0 End +
Case When Max(ITAX3)='TCS' Then Sum(ITAX3_Amt) Else 0 End +
Case When Max(ITAX4)='TCS' Then Sum(ITAX4_Amt) Else 0 End +
Case When Max(ITAX5)='TCS' Then Sum(ITAX5_Amt) Else 0 End +
Case When Max(ITAX6)='TCS' Then Sum(ITAX6_Amt) Else 0 End +
Case When Max(ITAX7)='TCS' Then Sum(ITAX7_Amt) Else 0 End +
Case When Max(ITAX8)='TCS' Then Sum(ITAX8_Amt) Else 0 End +
Case When Max(ITAX9)='TCS' Then Sum(ITAX9_Amt) Else 0 End +
Case When Max(ITAX10)='TCS' Then Sum(ITAX10_Amt) Else 0 End As 'TCS',"
                    BaseQry += " Max(ITAX1) As ITAX1,Max(ITAX1_RATE) As ITAX1_RATE,Sum(ITAX1_Amt) As ITAX1_Amt,Max(ITAX2) As ITAX2,Max(ITAX2_RATE) As ITAX2_RATE,Sum(ITAX2_Amt) As ITAX2_Amt,Max(ITAX3) As ITAX3,Max(ITAX3_Rate) As ITAX3_Rate,Sum(ITAX3_Amt) As ITAX3_Amt,Max(ITAX4) As ITAX4,Max(ITAX4_RATE) As ITAX4_RATE,Sum(ITAX4_Amt) As ITAX4_Amt,Max(ITAX5) As ITAX5,Max(ITAX5_RATE) As ITAX5_RATE,Sum(ITAX5_Amt) As ITAX5_Amt,Max(ITAX6) As ITAX6,Max(ITAX6_RATE) As ITAX6_RATE,Sum(ITAX6_Amt) As ITAX6_Amt,Max(ITAX7) As ITAX7,Max(ITAX7_Rate) As ITAX7_Rate,Sum(ITAX7_Amt) As ITAX7_Amt,Max(ITAX8) As ITAX8,Max(ITAX8_RATE) As ITAX8_RATE,Sum(ITAX8_Amt) As ITAX8_Amt,Max(ITAX9) As ITAX9,Max(ITAX9_Rate) As ITAX9_Rate,Sum(ITAX9_Amt) As ITAX9_Amt,Max(ITAX10) As ITAX10,Max(ITAX10_RATE) As ITAX10_RATE,Sum(ITAX10_Amt) As ITAX10_Amt,Max(Zone_Code) As Zone_Code,Max(CF) As CF,Max(ConversionFactor) As ConversionFactor,Max(EInvoice_Type) As EInvoice_Type,Max(LeakageDeduction_Freshsale) As LeakageDeduction_Freshsale,Max(LeakageDeduction) As LeakageDeduction,Max(Location_Desc) As Location_Desc,Max(Loc_Short_Name) As Loc_Short_Name,Max(Loc_Pin) As Loc_Pin,Max(Loc_Phone) As Loc_Phone,Max(Loc_Eamil) As Loc_Eamil,Max(Loc_Website) As Loc_Website,Max(xyz.ISO_No) As ISO_No,Max(Invoice_No) As Invoice_No,Max(Invoice_Date) As Invoice_Date,Max(Cust_City) As Cust_City,Max(Against_Shipment_No) As Against_Shipment_No,Max(Cust_Gst_StateCode) As Cust_Gst_StateCode,Max(Electronic_Ref_No) As Electronic_Ref_No,Max(CustGSTNo) As CustGSTNo,Max(gst_state_code) As gst_state_code,Max(LocGstNo) As LocGstNo,Max(EWayBillNo) As EWayBillNo,Max(EWayBillDate) As EWayBillDate,Max(HSN_Code) As HSN_Code,Max(InvRemarks) As InvRemarks,Max(Delivery_Code) As Delivery_Code,Max(Conversion_factor) As Conversion_factor,Sum(QTY_Box) As QTY_Box,Max(Sale_Invoice_No) As Sale_Invoice_No,Max(vehicleNo) As vehicleNo,Max(Sale_Invoice_Date) As Sale_Invoice_Date,Sum(RoundOffAmount) As RoundOffAmount,Max(Loc_ADd1) As Loc_ADd1,Max(LOC_ADD2) As LOC_ADD2,Max(LOC_ADD3) As LOC_ADD3,Max(LocationState) As LocationState,Max(LOCPhone) As LOCPhone,Max(Loc_TIN_NO) As Loc_TIN_NO,Max(Document_Code) As Document_Code,Max(Document_Date) As Document_Date,Max(Description) As Description,Max(Lorry_No) As Lorry_No,Max(Sku_Seq) As Sku_Seq,Max(Item_Code) As Item_Code,Max(Line_No) As Line_No,Max(Item_Desc) As Item_Desc,Sum(QtyCrates) As QtyCrates,Max(ConvFactInCrate) As ConvFactInCrate,Max(ConvQtyInCrate) As ConvQtyInCrate,Max(Unit_code) As Unit_code,Sum(Qty_Default) As Qty_Default,Max(Rate_Default) As Rate_Default,Sum(QtyPCS) As QtyPCS,Sum(free_qty) As free_qty,Max(FreeSchemeInLitres) As FreeSchemeInLitres,Max(RatePerPcs) As RatePerPcs,Sum(valueInRs) As valueInRs,Max(comp_add2) As comp_add2,Max(comp_add3) As comp_add3,Max(CompPhone) As CompPhone,Max(Cash_Scheme_Amount) As Cash_Scheme_Amount,Max(schemeInCrates) As schemeInCrates,Max(GrandTotalCrates) As GrandTotalCrates,Max(xyz.Comp_Code) As Comp_Code,Max(xyz.Comp_Name) As Comp_Name,Max(comp_add1) As comp_add1,Max(comp_Fax) As comp_Fax,Max(comp_Email) As comp_Email,Max(comp_tinNo) As comp_tinNo,Max(xyz.cust_Code) As cust_Code,Max(Customer_Name) As Customer_Name,Max(cust_add1) As cust_add1,Max(cust_add2) As cust_add2,Max(cust_add3) As cust_add3,Max(CustPhone) As CustPhone,Max(cust_fax) As cust_fax,Max(Cust_state) As Cust_state,Max(cust_Statename) As cust_Statename,Max(cust_Email) As cust_Email,Max(cust_website) As cust_website,Max(Customer_Pan) As Customer_Pan,Max(Ack_No) As Ack_No,Max(Ack_Date) As Ack_Date,Max(TaxableNonTaxable) As TaxableNonTaxable,Max(TAX1) As TAX1,Max(TaxType1) As TaxType1,Sum(TAX1_Amt) As TAX1_Amt,Max(TAX1_Rate) As TAX1_Rate,Sum(TAX1Amt) As TAX1Amt,Max(TaxType2) As TaxType2,Max(TAX2) As TAX2,Sum(TAX2_Amt) As TAX2_Amt,Max(TAX2_Rate) As TAX2_Rate,Sum(TAX2Amt) As TAX2Amt,Max(TaxType3) As TaxType3,Max(TAX3) As TAX3,Sum(TAX3_Amt) As TAX3_Amt,Max(TAX3_Rate) As TAX3_Rate,Sum(TAX3Amt) As TAX3Amt,Max(TaxType4) As TaxType4,Max(TAX4) As TAX4,Sum(TAX4_Amt) As TAX4_Amt,Max(TAX4_Rate) As TAX4_Rate,Sum(TAX4Amt) As TAX4Amt,Max(TaxType5) As TaxType5,Max(TAX5) As TAX5,Sum(TAX5_Amt) As TAX5_Amt,Max(TaxType6) As TaxType6,Max(TAX6) As TAX6,Sum(TAX6_Amt) As TAX6_Amt,STRING_AGG(Route_No, ', ') As Route_No,STRING_AGG(Route_Desc, ', ') As Route_Desc,Max(Distributor_Commission_TotalAmt) As Distributor_Commission_TotalAmt,Max(Transporter_Commission_TotalAmt) As Transporter_Commission_TotalAmt,Max(Against_Delivery_Code) As Against_Delivery_Code,Max(batchNO) As batchNO,Max(Credit_Customer) As Credit_Customer,Max(Ship_To_Code) As Ship_To_Code,Max(Ship_To_Desc) As Ship_To_Desc,Max(Ship_Address) As Ship_Address,Max(Ship_City) As Ship_City,Max(Ship_State) As Ship_State,Max(Ship_Pin_Code) As Ship_Pin_Code,Max(Ship_PAN) As Ship_PAN,Max(Ship_GSTNO) As Ship_GSTNO,Sum(Booth_Security_Amt) As Booth_Security_Amt,Max(Brand) As Brand,Max(BRANDDESC) As BRANDDESC,Max(Particulars) As Particulars,Max(Crate_No) As Crate_No,Max(CopyType) As CopyType,Max(SellerGST) As SellerGST,Max(xyz.Pan_No) As Pan_No,Max(RateLtr) As RateLtr from (" + objMultPrintInvoice.PrintInvoiceForAll(InvoiceNO, txtFromDate.Value, "", "Y", False, ItemType) + ")xyz "
                    If isPdf Then
                        BaseQry += " where convert(Date,Supply_Date,103)=Convert(Date,'" + SupplyDate + "',103) And Shift_Type='" + Shift + "' And cust_Code ='" + strPartyCode + "'" ' And TaxableNonTaxable='" + isTaxable + "'"
                    Else
                        BaseQry += " And TaxableNonTaxable='" + isTaxable + "'"
                    End If
                    BaseQry += " Group By Convert(DAte,Supply_Date,103),Shift_Type,TaxableNonTaxable,cust_Code,Item_Code "
                    Qry = "Select TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,* from (" + BaseQry + ")xyzFinal left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=xyzFinal.comp_code"
                Else
                    Qry = objMultPrintInvoice.PrintInvoiceForAll(InvoiceNO, txtFromDate.Value, "", "Y", False, ItemType)
                End If

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt Is Nothing AndAlso dt.Rows.Count <= 0 Then
                    Throw New Exception("Data not found !")
                End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceTNK", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    If rbtnPartyWise.Checked Then
                        pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJPR_PartyWise", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    Else
                        pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJPR", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    End If
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal AndAlso dt.Rows(0)("TaxableNonTaxable").ToString() = "T" Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceBKN", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceBKN", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceNew", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
                    '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceSKRPrintDistribution", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceCHU1", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                    'pdfPath = frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceCHU1", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJAL", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())

                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceGNG1", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceNAG", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())

                Else
                    pdfPath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                End If
                'frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJPR", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
                frmCRV = Nothing
            End If
            'clsCommon.ProgressBarPercentHide()
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            Throw New Exception(ex.Message)
        End Try
        'Else
        'For Each grow As GridViewRowInfo In gv.Rows
        '    If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
        '        strInvoice = "'" + clsCommon.myCstr(grow.Cells("Document_Code").Value) + "'"
        '        Try
        '            If clsCommon.myLen(strInvoice) <= 0 Then
        '                myMessages.blankValue(Me, "Invoice not found to Print", Me.Text)
        '            Else
        '                Dim dtDocdate As Date?
        '                dtDocdate = Nothing
        '                Dim StrSql As String = "Select Document_Code,Document_Date,Customer_Code,Bill_To_Location,is_taxable,Tax_Group from TSPL_SD_SALE_INVOICE_HEAD where Document_Code=" + strInvoice
        '                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(StrSql)
        '                If dt1.Rows.Count > 0 Then
        '                    'IsTaxable = clsCommon.myCdbl(dt1.Rows(0)("is_taxable"))
        '                    dtDocdate = clsCommon.myCDate(dt1.Rows(0)("Document_Date"))
        '                End If
        '                Dim frmCRV As New frmCrystalReportViewer()
        '                Qry = objMultPrintInvoice.PrintInvoiceForAll(strInvoice, clsCommon.GetPrintDate(grow.Cells("Document_Date").Value), clsCommon.myCstr(grow.Cells("Customer_Code").Value))
        '                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
        '                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal Then
        '                    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceTNK", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
        '                    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceJPR", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal AndAlso dt.Rows(0)("TaxableNonTaxable").ToString() = "T" Then
        '                    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoiceBKN", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
        '                    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptNonTaxableInvoiceBKN", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                    'ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
        '                    '    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                Else
        '                    frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                End If
        '                'frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptTaxableNonTaxableInvoice", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                frmCRV = Nothing
        '            End If
        '        Catch ex As Exception
        '            clsCommon.MyMessageBoxShow(Me, "No data found", Me.Text)
        '        End Try
        '    End If
        'Next
        ' End If




        'loadData(1)
        'ArrInvoice_Arr = New ArrayList
        'Dim InvoiceNo As String = ""
        'Dim frmCRV As New frmCrystalReportViewer()
        'If isPdf = True Then
        '    InvoiceNo = DocNo
        'Else
        '    For Each grow As GridViewRowInfo In gv.Rows
        '        If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
        '            InvoiceNo = InvoiceNo + "','" + clsCommon.myCstr(grow.Cells("Document_Code").Value)
        '        End If
        '    Next

        '    If clsCommon.myLen(InvoiceNo) > 0 AndAlso clsCommon.myCstr(InvoiceNo).Substring(0, 3) = "','" Then
        '        InvoiceNo = InvoiceNo.Substring(3, InvoiceNo.Length - 3)

        '    End If
        'End If
        'If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AllowToPrintInvoiceAfterPosting & "'")) = 1 Then
        '    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select distinct status from tspl_sd_sale_invoice_head where document_code in ('" & InvoiceNo & "')")) = 0 Then
        '        frmCRV.ShowCystalReportToolbar = False
        '    End If
        'End If
        'If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "WHOLLY") = CompairStringResult.Equal Then

        '    Dim Qry As String = " WITH A AS( " & _
        '    " SELECT IT.ITEM_DESC AS Item_Desc,CRATE_QTY=SUM(Qty),PCS_QTY=(SUM(Qty) *(ISNULL(U.Conversion_Factor,0))),Scheme_Qty=0,SCHEME_PCS_QTY=0,Item_Cost= ROUND((case when ISNULL(U.Conversion_Factor,0)>0 THEN ( CASE WHEN d.Unit_code!='PP' THEN  Item_Cost/ISNULL(U.Conversion_Factor,0) ELSE Item_Cost END ) else Item_Cost end) ,2) ,Amount=SUM(Amount),d.Unit_code,Amt_Less_Discount=SUM(Amt_Less_Discount) FROM TSPL_SD_SALE_INVOICE_DETAIL D  " & _
        '    " LEFT JOIN TSPL_ITEM_MASTER IT ON D.Item_Code=IT.Item_Code " & _
        '    " LEFT JOIN TSPL_ITEM_UOM_DETAIL U ON U.Item_Code=D.Item_Code AND D.Unit_code=U.UOM_Code  " & _
        '    " WHERE  D.Document_Code='" + InvoiceNo + "'  and Scheme_Item='N' " & _
        '    " GROUP BY IT.ITEM_DESC,Item_Cost,U.Conversion_Factor,d.Unit_code " & _
        '    " UNION ALL " & _
        '    " SELECT IT.ITEM_DESC AS Item_Desc,CRATE_QTY=0,PCS_QTY=0,Scheme_Qty=SUM(Qty),SCHEME_PCS_QTY=(case when d.Unit_code='CRATES' then (SUM(Qty) *(ISNULL(U.Conversion_Factor,0))) else 0 end),Item_Cost= ROUND((case when ISNULL(U.Conversion_Factor,0)>0 THEN ( CASE WHEN d.Unit_code!='PP' THEN  Item_Cost/ISNULL(U.Conversion_Factor,0) ELSE Item_Cost END ) else Item_Cost end) ,2) ,Amount=SUM(Amount),d.Unit_code,Amt_Less_Discount=SUM(Amt_Less_Discount) FROM TSPL_SD_SALE_INVOICE_DETAIL D  " & _
        '    " LEFT JOIN TSPL_ITEM_MASTER IT ON D.Item_Code=IT.Item_Code " & _
        '    " LEFT JOIN TSPL_ITEM_UOM_DETAIL U ON U.Item_Code=D.Item_Code AND D.Unit_code=U.UOM_Code " & _
        '    " WHERE D. Document_Code='" + InvoiceNo + "'  and Scheme_Item='Y' " & _
        '    " GROUP BY IT.ITEM_DESC,Item_Cost,U.Conversion_Factor,d.Unit_code ,d.Unit_code ) " & _
        '    " SELECT Comp_Name,COMP_ADDRESS=(CM.Add1+' '+CM.Add2+' '+CM.Add3), STATE=(CM.Phone1 ),CM.Tin_No, " & _
        '    " LOCATION_ADDRESS=(L.Add1+' '+L.Add2+' '+L.Add3), " & _
        '    " C.Customer_Name,CUST_ADD=(C.Add1+' '+C.Add2+' '+C.Add3),CT.City_Name,C.PAN, " & _
        '    " I.Document_Code, " & _
        '    " Document_Date=CONVERT(VARCHAR(100),Document_Date,103),Cust_PO_Date=CONVERT(VARCHAR(100),Cust_PO_Date,103),I.Against_Shipment_No ,Challan_Date=CONVERT(VARCHAR(100),I.Challan_Date,103),I.VehicleNo,i.discount_amt, " & _
        '    " A.*,TOTAL_SCHEME_AMOUNT=(SELECT ISNULL(SUM(AMOUNT),0) FROM TSPL_SD_SALE_INVOICE_DETAIL I WHERE  I.Document_Code='" + InvoiceNo + "' AND  Scheme_Item='Y'),CM.Logo_Img,I.Created_By " & _
        '    " FROM TSPL_SD_SALE_INVOICE_HEAD I " & _
        '    " JOIN TSPL_CUSTOMER_MASTER C ON C.Cust_Code=I.Customer_Code " & _
        '    " LEFT JOIN TSPL_CITY_MASTER CT ON C.City_Code=CT.City_Code " & _
        '    " JOIN TSPL_LOCATION_MASTER L ON L.Location_Code=I.Bill_To_Location " & _
        '    " JOIN TSPL_COMPANY_MASTER CM ON CM.Comp_Code=I.Comp_Code " & _
        '    " JOIN A A ON 1=1 " & _
        '    " WHERE I.Document_Code='" + InvoiceNo + "' "

        '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)

        '    strRptPath = frmCRV.funreport(isPdf, CrystalReportFolder.SalesReport, dt, "rptProductionSaleInvoiceWHC", "Sale Report")
        'Else
        '    Dim IsTaxable As Double = 0
        '    Dim dtDocdate As Date?
        '    dtDocdate = Nothing
        '    Dim StrSql = "Select Document_Date,Customer_Code,Bill_To_Location,is_taxable,Tax_Group from TSPL_SD_SALE_INVOICE_HEAD where Document_Code in ('" & InvoiceNo & "')"
        '    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(StrSql)
        '    If dt1.Rows.Count > 0 Then
        '        IsTaxable = clsCommon.myCdbl(dt1.Rows(0)("is_taxable"))
        '        dtDocdate = clsCommon.myCDate(dt1.Rows(0)("Document_Date"))
        '    End If
        '    Dim objInvoice As New frmSaleInvoiceProductSale
        '    If IsTaxable = 1 Then
        '        strRptPath = ""
        '        objInvoice.funPrint(InvoiceNo, False, Nothing, Nothing, False, isPdf)
        '        If isPdf = True Then
        '            strRptPath = objInvoice.strrptpath
        '        End If
        '    Else
        '        Dim Qry As String = Nothing
        '        Dim objMultPrintInvoice As New FrmPrintFreshInvoice

        '        If AllowSeperateSchemeItemOnPrint Then
        '            objInvoice.GetSeperateSchemeItemPrintQry(InvoiceNo)
        '        Else

        '            '====================Added by richa agarwal 23 Nov,2018=======================
        '            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then
        '                Dim CreateFreshInvoiceOnDispatchSave As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateFreshInvoiceOnDispatchSave, clsFixedParameterCode.CreateFreshInvoiceOnDispatchSave, Nothing))
        '                If CreateFreshInvoiceOnDispatchSave = 0 Then
        '                    Dim dt3 As DataTable = clsDBFuncationality.GetDataTable("select document_code from TSPL_SD_SHIPMENT_HEAD where Sale_Invoice_No in ('" & InvoiceNo & "')")
        '                    Dim dispatchno As String = String.Empty
        '                    For Each dr As DataRow In dt3.Rows
        '                        dispatchno = dispatchno + "','" + clsCommon.myCstr(dr("document_code"))
        '                    Next

        '                    If clsCommon.myLen(dispatchno) > 0 AndAlso clsCommon.myCstr(dispatchno).Substring(0, 3) = "','" Then
        '                        dispatchno = dispatchno.Substring(3, dispatchno.Length - 3)

        '                    End If
        '                    Qry = objMultPrintInvoice.LoadPrintQuery(dispatchno)
        '                Else
        '                    ''richa 2 Apr,2019 ERO/04/04/19-000543
        '                    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then

        '                        Dim dt3 As DataTable = clsDBFuncationality.GetDataTable("select document_code from TSPL_SD_SHIPMENT_HEAD where Sale_Invoice_No in ('" & InvoiceNo & "')")
        '                        Dim dispatchno As String = String.Empty
        '                        For Each dr As DataRow In dt3.Rows
        '                            dispatchno = dispatchno + "','" + clsCommon.myCstr(dr("document_code"))
        '                        Next

        '                        If clsCommon.myLen(dispatchno) > 0 AndAlso clsCommon.myCstr(dispatchno).Substring(0, 3) = "','" Then
        '                            dispatchno = dispatchno.Substring(3, dispatchno.Length - 3)

        '                        End If
        '                        Qry = objMultPrintInvoice.LoadPrintQuery(dispatchno)

        '                    Else
        '                        Qry = objMultPrintInvoice.LoadPrintQuery(InvoiceNo)
        '                    End If

        '                End If
        '            Else
        '                Qry = objMultPrintInvoice.LoadPrintQuery(InvoiceNo)
        '            End If
        '            '================================================
        '            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then
        '                If isPdf = True Then
        '                    Qry = " Select * from ( " + Qry + " ) XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  '' as CopyType1 ) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2 ,xxx.Line_No "
        '                ElseIf isPrePrintFormat = True Then
        '                    ' Qry = " Select *,count(*) over ( partition by Sale_invoice_no ) as TotalNoOfItem from ( " + Replace(Qry, " upper(SUBSTRING(Batch_No,1,7))", " case when len ( isnull (Batch_No,'') ) > 0 then SUBSTRING (Batch_No,1,7) else '' end as Batch_No ") + " ) XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL' as CopyType1  ) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2 ,xxx.Line_No "
        '                    Qry = " Select *,count(*) over ( partition by Sale_invoice_no ) as TotalNoOfItem from ( " + Qry + " ) XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL' as CopyType1  ) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2 ,xxx.Line_No "
        '                Else
        '                    Qry = " Select * from ( " + Qry + " ) XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE' as CopyType1 ) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2 ,xxx.Line_No "
        '                End If
        '            End If



        '            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)

        '            If dt.Rows.Count > 0 Then
        '                If ShowShipToPartyInDairyDispatch = 1 OrElse clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then
        '                    ' done by richa agarwal 23 Nov,2018 add customer dashbord into existing report for Erode client ERO/01/04/19-000534
        '                    If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then
        '                        Qry = "select distinct customer_code from tspl_sd_sale_invoice_head where document_code in ('" & InvoiceNo & "')"
        '                        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(Qry)
        '                        Dim CustomerCode As String = ""
        '                        If (dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0) Then
        '                            For Each dr As DataRow In dt2.Rows
        '                                CustomerCode = CustomerCode + "','" + clsCommon.myCstr(dr("customer_code"))
        '                            Next
        '                            If clsCommon.myLen(CustomerCode) > 0 AndAlso clsCommon.myCstr(CustomerCode).Substring(0, 3) = "','" Then
        '                                CustomerCode = CustomerCode.Substring(3, CustomerCode.Length - 3)

        '                            End If
        '                        End If
        '                        Dim dtCustomerOutstanding As DataTable = Nothing
        '                        dtCustomerOutstanding = clsCustomerMaster.getCustomerOutstandingOfAmt_Can_Crate("'" & CustomerCode & "'", clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_date")).AddDays(-1), "dd/MMM/yyyy"), clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(0)("Document_date")), "dd/MMM/yyyy"))
        '                        ' Pre-Format Print for EROD
        '                        If isPrePrintFormat = True Then
        '                            strRptPath = frmCRV.funsubreportWithdt(isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptFreshSaleInvoicePartyForPreFormat", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader(), "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
        '                        Else
        '                            strRptPath = frmCRV.funsubreportWithdt(isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptFreshSaleInvoiceParty", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader(), "rptCustomerOutstandingErode.rpt", dtCustomerOutstanding)
        '                        End If


        '                    Else
        '                        strRptPath = frmCRV.funsubreportWithdt(isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptFreshSaleInvoiceParty", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                    End If

        '                Else
        '                    strRptPath = frmCRV.funsubreportWithdt(isPdf, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptFreshSaleInvoice(New)", "Bill of Supply", dtDocdate, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
        '                End If

        '            End If
        '        End If
        '    End If
        'End If
        'frmCRV = Nothing
        Return pdfPath
    End Function

    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub

    Private Sub FrmPrintDistributerInvoiceStatement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            RadGroupBox2.Visible = True
            RadGroupBox4.Visible = False
            EnableProductSaleForJPR = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)) = 1, True, False)
            If EnableProductSaleForJPR Then
                RadGroupBox6.Visible = True
            Else
                RadGroupBox6.Visible = False
            End If
            'Dim coll As Dictionary(Of String, String)
            'coll = New Dictionary(Of String, String)
            'coll.Add("MonthlySaleInvoiceNo", "varchar(30) NULL")
            'clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_SD_SALE_INVOICE_HEAD", coll, Nothing, False, True, "", "Document_Code", "Document_Date")
            isInsideLoadData = True
            SetUserMgmtNew()
            ApplyMilkPouchPrint = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyMilkPouchPrint, clsFixedParameterCode.ApplyMilkPouchPrint, Nothing)) = 0, False, True)
            AllowSeperateSchemeItemOnPrint = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowSeprateSchemeItemPrintDairySaleInvoice, clsFixedParameterCode.AllowSeprateSchemeItemPrintDairySaleInvoice, Nothing)) = 0, False, True)
            ShowShipToPartyInDairyDispatch = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.ShowShipToPartyInDairyDispatch & "'")) = 0, 0, 1)
            ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
            ButtonToolTip.SetToolTip(btnGo, "Press Alt+R Refresh ")
            ButtonToolTip.SetToolTip(BtnPrint, "Press Alt+P For Print")
            ButtonToolTip.SetToolTip(BtnReset, "Press Alt+N Adding New")
            LoadReportType()
            btnUnSelect.Text = "Select All"
            Reset()
            isInsideLoadData = False
        Catch ex As Exception
            isInsideLoadData = False
        End Try

    End Sub
    Sub loadData(ByVal printType As Integer)


        ArrInvoice_Arr = New ArrayList


        Dim InvoiceNo As String = ""
        Dim DispatchNo As String = ""

        For Each grow As GridViewRowInfo In gv.Rows
            If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                InvoiceNo = InvoiceNo + "','" + clsCommon.myCstr(grow.Cells("Document_Code").Value)

            End If
        Next

        If clsCommon.myLen(InvoiceNo) > 0 AndAlso clsCommon.myCstr(InvoiceNo).Substring(0, 3) = "','" Then
            InvoiceNo = InvoiceNo.Substring(3, InvoiceNo.Length - 3)

        End If




        If txtFromDate.Value > txtToDate.Value Then
            common.clsCommon.MyMessageBoxShow(Me, "From date can not be greater then to Date", Me.Text)
            txtFromDate.Focus()
            Exit Sub
        End If


        Dim objInvoice As New clsPSShipmentHead
        Dim Qry As String = objInvoice.GetQuery()
        Qry += " where TSPL_SD_SALE_INVOICE_HEAD.Document_Code in ('" + InvoiceNo + "')"
        ''===============================================================================
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "WHOLLY") = CompairStringResult.Equal Then

            Qry = "select max(Remarks) as Remarks,sum(isnull(Roundoff,0)) as Roundoff,max(Transporter_Name) as Transporter_Name,sum(crate) as crate,sum(jaali) as jaali,sum(box) as box,max(GP_Invoice_Type) as GP_Invoice_Type,max(Loc_Tin_No) as Loc_Tin_No,max(Location_Address_GP) as Location_Address_GP,max(comp_code) as comp_code,max(Comp_Add_GP) as Comp_Add_GP,max(GP_Division)as GP_Division,max(GP_ECC_No) as GP_ECC_No,max(GP_CE_Range) as GP_CE_Range,max(ITF_CODE) as ITF_CODE,max(Loc_State_Name) as Loc_State_Name,max(City_Name) as City_Name,max(Loc_state_code) as Loc_state_code,max(HOAdd1) as HOAdd1,max(HOAdd2) as HOAdd2,max(Comp_CSt_LST) as Comp_CSt_LST,max(Transport_Id) as Transport_Id,max(GRNo) as GRNo,max(GR_Date) as GR_Date,max(VehicleNo) as VehicleNo,sum(isnull(Alter_UnitQty,0)) as Alter_UnitQty,sum(isnull(HeadDisc_Amt,0)) as HeadDisc_Amt,max(Chap_Desc) as Chap_Desc,sum(isnull(HeadDisc_PerAmt,0)) as HeadDisc_PerAmt " &
          " ,max(Registration_Number) as Registration_Number,max(Payment_Terms) as Payment_Terms,max(Modify_By) as Modify_By,max(Comp_PANNO) as Comp_PANNO,max(Road_Permit_No) as Road_Permit_No,max(Cheapter_Heads) as Cheapter_Heads,max(RoadPermit_Date)as RoadPermit_Date,max(Location_Address) as Location_Address,max(Loc_CSTNo) as Loc_CSTNo,max(loc_Excisable) as loc_Excisable,max(Loc_range_Add) as Loc_range_Add,max(loc_Division_Address) as loc_Division_Address,max(Loc_Commissionerate) as Loc_Commissionerate " &
          " ,max(Challan_No) as Challan_No,max(Challan_Date) as Challan_Date	,max(Removal_Date) as Removal_Date,sum(isnull(total_add_charge,0))as total_add_charge,max(WayBillNo) as WayBillNo,max(Comp_Address) as Comp_Address,max(Loc_Add1) as Loc_Add1,max(Loc_ADd2) as Loc_ADd2,max(Loc_Add3) as Loc_Add3,max(Loc_Pin_Code) as Loc_Pin_Code,max(Loc_TinNo) as Loc_TinNo,max(Loc_Phn) as Loc_Phn,max(Loc_Email) as Loc_Email,max(Invoice_Type) as Invoice_Type,sum(isnull(Alternet_Qty,0)) as Alternet_Qty,max(Alternate_UOM) as Alternate_UOM,sum(isnull(Scheme_Qty,0)) as Scheme_Qty,max(Scheme_Item_UOM) as Scheme_Item_UOM,sum(isnull(Discount_Base,0)) as Discount_Base,max(Dis_Doc_No) as Dis_Doc_No,max([Description]) as [Description],sum(isnull(Print_Discount_Amt,0)) as Print_Discount_Amt,max(Comp_State) as Comp_State,max(Buyer_order_no) as Buyer_order_no,max(Buyer_order_date) as Buyer_order_date,max(Terms_of_delivery) as Terms_of_delivery,max(Loc_Short_Name)Loc_Short_Name,InvoiceNo,Distributer_Dispatch_No,max(Date_Time_Invoice) as Date_Time_Invoice,max(InvoiceDate)as InvoiceDate,ShipmentNo,sum(isnull(Alt_Qty,0)) as Alt_Qty,max(Alt_UOM) as Alt_UOM,max(ShipmentDate) as ShipmentDate,DeliveryOrderNo,max(TermCondition) as TermCondition,max(Location_Desc) as Location_Desc,max(CompName) as CompName,max(CompPhone) as CompPhone,max(CompFax) as CompFax,max(ComMail) as ComMail,max(CompCinNo) as CompCinNo,max(ComPanNo) as ComPanNo,max(CompCSTLST) as CompCSTLST,max(ComPINCode) as ComPINCode,max(ComTinNO) as ComTinNO,max(Compaddress2) as Compaddress2,max(Compaddress3) as Compaddress3,max(P_Add1) as P_Add1,max(P_Add2) as P_Add2,max(P_Add3) as P_Add3,max(P_PinNo) as P_PinNo,max(P_CstNo) as P_CstNo,max(P_TinNo) as P_TinNo,max(P_Email) as P_Email,max(P_Fax) as P_Fax,max(P_LstNo) as P_LstNo,P_CustCode,max(P_Cust_Name) as P_Cust_Name,max(P_City_Name) as P_City_Name,max(P_State_Name) as P_State_Name,max(P_Cust_Phn) as P_Cust_Phn,max(P_Pan) as P_Pan,Cust_Code,max(Customer_Name) as Customer_Name,max(Cust_Add1) as Cust_Add1,max(Cust_add2) as Cust_add2,max(cust_add3) as cust_add3,max(Cust_Phn) as Cust_Phn,max(Cust_TinNo) as Cust_TinNo,max(Cust_CSTNo) as Cust_CSTNo,max(Cust_LSTNo) as Cust_LSTNo,max(Cust_Email) as Cust_Email,max(Cust_PinCode) as Cust_PinCode,max(Cust_City_Name) as Cust_City_Name,max(Cust_Fax) as Cust_Fax,max(Cust_State_Name) as Cust_State_Name,max(Customer_Pan) as Customer_Pan,item_code,max(itemdesc)as itemdesc,sum(isnull(qty,0)) as qty,mrp,sum(isnull(amount,0)) as amount,sum(isnull(Amt_Less_Discount,0)) as Amt_Less_Discount,sum(Scheme_Amount) as Scheme_Amount,SUM(CRATES_QTY) AS CRATES_QTY,SUM(PCS_QTY) AS PCS_QTY,SUM(SCH_PCS_QTY) AS SCH_PCS_QTY,sum(Item_Cost_Main) as Item_Cost_Main,max(uom) as uom,max(RATE_UOM) as RATE_UOM,itemcost,max(tax1name) as tax1name,sum(isnull(txt1amt,0)) as txt1amt,max(tax2name) as tax2name,sum(isnull(txt2amt,0)) as txt2amt,max(tax3name) as tax3name,sum(isnull(txt3amt,0)) as txt3amt,max(tax4name) as tax4name,sum(isnull(txt4amt,0)) as txt4amt,max(tax5name) as tax5name,sum(isnull(txt5amt,0)) as txt5amt,max(tax6name) as tax6name,sum(isnull(txt6amt,0)) as txt6amt, max(tax7name) as tax7name,sum(isnull(txt7amt,0)) as txt7amt,max(tax8name) as tax8name,sum(isnull(txt8amt,0)) as txt8amt, max(tax9name) as tax9name,sum(isnull(txt9amt,0)) as txt9amt,max(tax10name) as tax10name,sum(isnull(txt10amt,0)) as txt10amt " &
          " ,max(TAX1_Rate) as TAX1_Rate ,max(TAX2_Rate) as TAX2_Rate ,max(TAX3_Rate) as TAX3_Rate,max(TAX4_Rate) as TAX4_Rate,max(TAX5_Rate) as TAX5_Rate,max(TAX6_Rate) as TAX6_Rate,max(TAX7_Rate) as TAX7_Rate,max(TAX8_Rate) as TAX8_Rate,max(TAX9_Rate) as TAX9_Rate,max(TAX10_Rate) as TAX10_Rate,max(Disc_Per) as Disc_Per,sum(isnull( Discount_Amt,0)) as  Discount_Amt,sum(isnull(Amount_Less_Discount,0)) as Amount_Less_Discount,sum(isnull(Total_Tax_Amt,0)) as Total_Tax_Amt,sum(isnull(Total_Amt,0)) as Total_Amt,max(CopyType) as CopyType ,max(Add_Charge_Name1) as Add_Charge_Name1,sum(isnull(Add_Charge_Amt1,0)) as Add_Charge_Amt1,max(Add_Charge_Name2) as Add_Charge_Name2,sum(isnull(Add_Charge_Amt2,0)) as Add_Charge_Amt2,max(Add_Charge_Name3) as Add_Charge_Name3,sum(isnull(Add_Charge_Amt3,0)) as Add_Charge_Amt3,max(Add_Charge_Name4) as Add_Charge_Name4,sum(isnull(Add_Charge_Amt4,0)) as Add_Charge_Amt4,max(Add_Charge_Name5) as Add_Charge_Name5,sum(isnull(Add_Charge_Amt5,0)) as Add_Charge_Amt5,max(Add_Charge_Name6) as Add_Charge_Name6,sum(isnull(Add_Charge_Amt6,0)) as Add_Charge_Amt6,max(Add_Charge_Name7) as Add_Charge_Name7,sum(isnull(Add_Charge_Amt7,0)) as Add_Charge_Amt7,max(Add_Charge_Name8) as Add_Charge_Name8,sum(isnull(Add_Charge_Amt8,0)) as Add_Charge_Amt8,max(Add_Charge_Name9) as Add_Charge_Name9,sum(isnull(Add_Charge_Amt9,0)) as Add_Charge_Amt9,max(Add_Charge_Name10) as Add_Charge_Name10,sum(isnull(Add_Charge_Amt10,0)) as Add_Charge_Amt10  ,sum(isnull(RoundOffAmount,0)) as RoundOffAmount,Booking_No,sum(isnull(Distr_Commision,0)) as Distr_Commision,Line_No from  " &
          " (" & Qry & ")final group by InvoiceNo,Distributer_Dispatch_No,ShipmentNo,DeliveryOrderNo,P_CustCode,Cust_Code,item_code,mrp,itemcost,Booking_No,Line_No "
            ''=========================================================================


            Qry = "Select cast(TSPL_COMPANY_MASTER.logo_img as image) as logo_img,* from (" & Qry & ") XXX left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.comp_code=xxx.comp_code "

            If printType = 2 Then
                Qry += " LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2,Item_Code,Line_No asc"
            Else
                Qry += " LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2,Item_Code,Line_No asc"
            End If
        Else

            Qry = "select max(Remarks) as Remarks,sum(isnull(Roundoff,0)) as Roundoff,max(Transporter_Name) as Transporter_Name,sum(crate) as crate,sum(jaali) as jaali,sum(box) as box,max(GP_Invoice_Type) as GP_Invoice_Type,max(Loc_Tin_No) as Loc_Tin_No,max(Location_Address_GP) as Location_Address_GP,max(comp_code) as comp_code,max(Comp_Add_GP) as Comp_Add_GP,max(GP_Division)as GP_Division,max(GP_ECC_No) as GP_ECC_No,max(GP_CE_Range) as GP_CE_Range,max(ITF_CODE) as ITF_CODE,max(Loc_State_Name) as Loc_State_Name,max(City_Name) as City_Name,max(Loc_state_code) as Loc_state_code,max(HOAdd1) as HOAdd1,max(HOAdd2) as HOAdd2,max(Comp_CSt_LST) as Comp_CSt_LST,max(Transport_Id) as Transport_Id,max(GRNo) as GRNo,max(GR_Date) as GR_Date,max(VehicleNo) as VehicleNo,sum(isnull(Alter_UnitQty,0)) as Alter_UnitQty,sum(isnull(HeadDisc_Amt,0)) as HeadDisc_Amt,max(Chap_Desc) as Chap_Desc,sum(isnull(HeadDisc_PerAmt,0)) as HeadDisc_PerAmt " &
                " ,max(Registration_Number) as Registration_Number,max(Payment_Terms) as Payment_Terms,max(Modify_By) as Modify_By,max(Comp_PANNO) as Comp_PANNO,max(Road_Permit_No) as Road_Permit_No,max(Cheapter_Heads) as Cheapter_Heads,max(RoadPermit_Date)as RoadPermit_Date,max(Location_Address) as Location_Address,max(Loc_CSTNo) as Loc_CSTNo,max(loc_Excisable) as loc_Excisable,max(Loc_range_Add) as Loc_range_Add,max(loc_Division_Address) as loc_Division_Address,max(Loc_Commissionerate) as Loc_Commissionerate " &
                " ,max(Challan_No) as Challan_No,max(Challan_Date) as Challan_Date	,max(Removal_Date) as Removal_Date,sum(isnull(total_add_charge,0))as total_add_charge,max(WayBillNo) as WayBillNo,max(Comp_Address) as Comp_Address,max(Loc_Add1) as Loc_Add1,max(Loc_ADd2) as Loc_ADd2,max(Loc_Add3) as Loc_Add3,max(Loc_Pin_Code) as Loc_Pin_Code,max(Loc_TinNo) as Loc_TinNo,max(Loc_Phn) as Loc_Phn,max(Loc_Email) as Loc_Email,max(Invoice_Type) as Invoice_Type,sum(isnull(Alternet_Qty,0)) as Alternet_Qty,max(Alternate_UOM) as Alternate_UOM,sum(isnull(Scheme_Qty,0)) as Scheme_Qty,max(Scheme_Item_UOM) as Scheme_Item_UOM,sum(isnull(Discount_Base,0)) as Discount_Base,max(Dis_Doc_No) as Dis_Doc_No,max([Description]) as [Description],sum(isnull(Print_Discount_Amt,0)) as Print_Discount_Amt,max(Comp_State) as Comp_State,max(Buyer_order_no) as Buyer_order_no,max(Buyer_order_date) as Buyer_order_date,max(Terms_of_delivery) as Terms_of_delivery,max(Loc_Short_Name)Loc_Short_Name,InvoiceNo,Distributer_Dispatch_No,max(Date_Time_Invoice) as Date_Time_Invoice,max(InvoiceDate)as InvoiceDate,ShipmentNo,sum(isnull(Alt_Qty,0)) as Alt_Qty,max(Alt_UOM) as Alt_UOM,max(ShipmentDate) as ShipmentDate,DeliveryOrderNo,max(TermCondition) as TermCondition,max(Location_Desc) as Location_Desc,max(CompName) as CompName,max(CompPhone) as CompPhone,max(CompFax) as CompFax,max(ComMail) as ComMail,max(CompCinNo) as CompCinNo,max(ComPanNo) as ComPanNo,max(CompCSTLST) as CompCSTLST,max(ComPINCode) as ComPINCode,max(ComTinNO) as ComTinNO,max(Compaddress2) as Compaddress2,max(Compaddress3) as Compaddress3,max(P_Add1) as P_Add1,max(P_Add2) as P_Add2,max(P_Add3) as P_Add3,max(P_PinNo) as P_PinNo,max(P_CstNo) as P_CstNo,max(P_TinNo) as P_TinNo,max(P_Email) as P_Email,max(P_Fax) as P_Fax,max(P_LstNo) as P_LstNo,P_CustCode,max(P_Cust_Name) as P_Cust_Name,max(P_City_Name) as P_City_Name,max(P_State_Name) as P_State_Name,max(P_Cust_Phn) as P_Cust_Phn,max(P_Pan) as P_Pan,Cust_Code,max(Customer_Name) as Customer_Name,max(Cust_Add1) as Cust_Add1,max(Cust_add2) as Cust_add2,max(cust_add3) as cust_add3,max(Cust_Phn) as Cust_Phn,max(Cust_TinNo) as Cust_TinNo,max(Cust_CSTNo) as Cust_CSTNo,max(Cust_LSTNo) as Cust_LSTNo,max(Cust_Email) as Cust_Email,max(Cust_PinCode) as Cust_PinCode,max(Cust_City_Name) as Cust_City_Name,max(Cust_Fax) as Cust_Fax,max(Cust_State_Name) as Cust_State_Name,max(Customer_Pan) as Customer_Pan,item_code,max(itemdesc)as itemdesc,sum(isnull(qty,0)) as qty,mrp,sum(isnull(amount,0)) as amount,sum(isnull(Amt_Less_Discount,0)) as Amt_Less_Discount,max(uom) as uom,max(RATE_UOM) as RATE_UOM,itemcost,max(tax1name) as tax1name,sum(isnull(txt1amt,0)) as txt1amt,max(tax2name) as tax2name,sum(isnull(txt2amt,0)) as txt2amt,max(tax3name) as tax3name,sum(isnull(txt3amt,0)) as txt3amt,max(tax4name) as tax4name,sum(isnull(txt4amt,0)) as txt4amt,max(tax5name) as tax5name,sum(isnull(txt5amt,0)) as txt5amt,max(tax6name) as tax6name,sum(isnull(txt6amt,0)) as txt6amt, max(tax7name) as tax7name,sum(isnull(txt7amt,0)) as txt7amt,max(tax8name) as tax8name,sum(isnull(txt8amt,0)) as txt8amt, max(tax9name) as tax9name,sum(isnull(txt9amt,0)) as txt9amt,max(tax10name) as tax10name,sum(isnull(txt10amt,0)) as txt10amt " &
                " ,max(TAX1_Rate) as TAX1_Rate ,max(TAX2_Rate) as TAX2_Rate ,max(TAX3_Rate) as TAX3_Rate,max(TAX4_Rate) as TAX4_Rate,max(TAX5_Rate) as TAX5_Rate,max(TAX6_Rate) as TAX6_Rate,max(TAX7_Rate) as TAX7_Rate,max(TAX8_Rate) as TAX8_Rate,max(TAX9_Rate) as TAX9_Rate,max(TAX10_Rate) as TAX10_Rate,max(Disc_Per) as Disc_Per,sum(isnull( Discount_Amt,0)) as  Discount_Amt,sum(isnull(Amount_Less_Discount,0)) as Amount_Less_Discount,sum(isnull(Total_Tax_Amt,0)) as Total_Tax_Amt,sum(isnull(Total_Amt,0)) as Total_Amt,max(CopyType) as CopyType ,max(Add_Charge_Name1) as Add_Charge_Name1,sum(isnull(Add_Charge_Amt1,0)) as Add_Charge_Amt1,max(Add_Charge_Name2) as Add_Charge_Name2,sum(isnull(Add_Charge_Amt2,0)) as Add_Charge_Amt2,max(Add_Charge_Name3) as Add_Charge_Name3,sum(isnull(Add_Charge_Amt3,0)) as Add_Charge_Amt3,max(Add_Charge_Name4) as Add_Charge_Name4,sum(isnull(Add_Charge_Amt4,0)) as Add_Charge_Amt4,max(Add_Charge_Name5) as Add_Charge_Name5,sum(isnull(Add_Charge_Amt5,0)) as Add_Charge_Amt5,max(Add_Charge_Name6) as Add_Charge_Name6,sum(isnull(Add_Charge_Amt6,0)) as Add_Charge_Amt6,max(Add_Charge_Name7) as Add_Charge_Name7,sum(isnull(Add_Charge_Amt7,0)) as Add_Charge_Amt7,max(Add_Charge_Name8) as Add_Charge_Name8,sum(isnull(Add_Charge_Amt8,0)) as Add_Charge_Amt8,max(Add_Charge_Name9) as Add_Charge_Name9,sum(isnull(Add_Charge_Amt9,0)) as Add_Charge_Amt9,max(Add_Charge_Name10) as Add_Charge_Name10,sum(isnull(Add_Charge_Amt10,0)) as Add_Charge_Amt10  ,sum(isnull(RoundOffAmount,0)) as RoundOffAmount,Booking_No,sum(isnull(Distr_Commision,0)) as Distr_Commision from  " &
                " (" & Qry & ")final group by InvoiceNo,Distributer_Dispatch_No,ShipmentNo,DeliveryOrderNo,P_CustCode,Cust_Code,item_code,mrp,itemcost,Booking_No "
            ''=========================================================================

            'scheme_Rate

            Qry = "Select cast(TSPL_COMPANY_MASTER.logo_img as image) as logo_img,isnull((select sum(qty*Amt_Less_Discount) from TSPL_SD_SALE_INVOICE_DETAIL where  document_code in ('" + InvoiceNo + "')  and Scheme_Item='Y'),0) as Scheme_Amount,* from (" & Qry & ") XXX left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.comp_code=xxx.comp_code "

            If printType = 2 Then
                Qry += " LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2,Item_Code desc"
            Else
                Qry += " LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2,Item_Code "
            End If
        End If

        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)

        Dim count As Integer = 0
        If dt IsNot Nothing And dt.Rows.Count > 0 Then
            Dim UOMKG As String = String.Empty

            For i As Integer = 0 To dt.Rows.Count - 1
                If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(i)("UOM")), clsCommon.myCstr(dt.Rows(0)("UOM"))) = CompairStringResult.Equal Then
                    count = count + 1

                End If
            Next
        End If
        dt.Columns.Add("UOMKG")
        If dt.Rows.Count = count Then
            For i As Integer = 0 To dt.Rows.Count - 1
                dt.Rows(i)("UOMKG") = "T"
            Next
        Else
            For i As Integer = 0 To dt.Rows.Count - 1
                dt.Rows(i)("UOMKG") = "F"
            Next
        End If



        If dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            If printType = 1 Then
                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDistributerProductSaleInvoice", "Distributer Invoice Statement", "")
            ElseIf printType = 2 Then
                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDistributerProductSaleInvoice_Challan", "Challan", "rptCompanyAddress.rpt")
            End If
            frmCRV = Nothing

        End If
        'Next

    End Sub
    Private Sub FrmPrintDistributerInvoiceStatement_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            loadReport()
        ElseIf e.Alt And e.KeyCode = Keys.C Then
            Me.Close()
        ElseIf e.Alt And e.KeyCode = Keys.R Then
            Reset()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P Then
            loadData(1)
        End If
    End Sub

    Private Sub btnUnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnSelect.Click
        If clsCommon.CompairString(btnUnSelect.Text, "UnSelect All") = CompairStringResult.Equal Then
            If gv IsNot Nothing AndAlso gv.ChildRows.Count > 0 Then
                For i As Integer = 0 To gv.ChildRows.Count - 1
                    gv.ChildRows(i).Cells(0).Value = False
                Next
            End If
        Else
            If gv IsNot Nothing AndAlso gv.ChildRows.Count > 0 Then
                For i As Integer = 0 To gv.ChildRows.Count - 1
                    gv.ChildRows(i).Cells(0).Value = True
                Next
            End If
        End If
        'If clsCommon.CompairString(btnUnSelect.Text, "UnSelect All") = CompairStringResult.Equal Then
        '    For Each grow As GridViewRowInfo In gv.Rows
        '        grow.Cells(0).Value = False

        '    Next
        '    btnUnSelect.Text = "Select All"
        'Else
        '    For Each grow As GridViewRowInfo In gv.Rows
        '        If clsCommon.myCstr(grow.Cells("Email").Value) IsNot Nothing And clsCommon.myLen(grow.Cells("Email").Value) > 0 Then
        '            grow.Cells(0).Value = True
        '        Else
        '            grow.Cells(0).Value = False
        '        End If
        '    Next
        '    btnUnSelect.Text = "UnSelect All"
        'End If
    End Sub
    Private Sub ExportToExcel(ByVal exporter As EnumExportTo)
        Try

            Dim arrHeader As List(Of String) = New List(Of String)()
            Dim strtemp As String = "Date Range : " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")
            arrHeader.Add(strtemp)
            arrHeader.Add("Company : " + objCommonVar.CurrentCompanyName)
            arrHeader.Add("Invoice Type : " & cboReportType.Text)
            arrHeader.Add("Customer : " & IIf(clsCommon.myLen(clsCommon.GetMulcallString(txtMultCustomer.arrDispalyMember)) = 2, "All", clsCommon.GetMulcallString(txtMultCustomer.arrDispalyMember)))
            arrHeader.Add("Location : " & IIf(clsCommon.myLen(clsCommon.GetMulcallString(txtMultLocation.arrDispalyMember)) = 2, "All", clsCommon.GetMulcallString(txtMultLocation.arrDispalyMember)))
            arrHeader.Add("Route : " & IIf(clsCommon.myLen(clsCommon.GetMulcallString(txtMultRoute.arrDispalyMember)) = 2, "All", clsCommon.GetMulcallString(txtMultRoute.arrDispalyMember)))


            'If chkLocationSelect.IsChecked Then
            '    Dim stVSPName As String = ""
            '    For Each StrName As String In cbgLocation.CheckedDisplayMember
            '        If clsCommon.myLen(stVSPName) > 0 Then
            '            stVSPName += ", "
            '        End If
            '        stVSPName += StrName
            '    Next
            '    Dim strVSPCode As String = ""
            '    For Each StrCode As String In cbgLocation.CheckedValue
            '        If clsCommon.myLen(strVSPCode) > 0 Then
            '            strVSPCode += ", "
            '        End If
            '        strVSPCode += StrCode
            '    Next
            '    arrHeader.Add(("Location: " + stVSPName + " "))
            'End If
            transportSql.applyExportTemplate(gv, PageSetupReport_ID)
            If exporter = EnumExportTo.Excel Then
                clsCommon.MyExportToExcelGrid("Product Sale Invoice", gv, arrHeader, Me.Text)

            Else
                clsCommon.MyExportToPDF("Product Sale Invoice", gv, arrHeader, "Product Sale Invoice", PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error, Me.Text)
        End Try
    End Sub
    Private Sub rmExcel_Click(sender As Object, e As EventArgs) Handles rmExcel.Click
        'If gv.Rows.Count > 0 Then
        '    ExportToExcel(EnumExportTo.Excel)
        'Else
        '    RadMessageBox.Show("No Data Found to Display", Me.Text)
        'End If
        Try

            Dim arrHeader As List(Of String) = New List(Of String)()

            arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")) + " ")
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.FrmPrintDistributerInvoiceStatement & "'"))
            arrHeader.Add("Invoice Type : " & cboReportType.Text)
            arrHeader.Add("Customer : " & IIf(clsCommon.myLen(clsCommon.GetMulcallString(txtMultCustomer.arrDispalyMember)) = 2, "All", clsCommon.GetMulcallString(txtMultCustomer.arrDispalyMember)))
            arrHeader.Add("Location : " & IIf(clsCommon.myLen(clsCommon.GetMulcallString(txtMultLocation.arrDispalyMember)) = 2, "All", clsCommon.GetMulcallString(txtMultLocation.arrDispalyMember)))
            arrHeader.Add("Route : " & IIf(clsCommon.myLen(clsCommon.GetMulcallString(txtMultRoute.arrDispalyMember)) = 2, "All", clsCommon.GetMulcallString(txtMultRoute.arrDispalyMember)))

            'Dim sfd As SaveFileDialog = New SaveFileDialog()
            'Dim filePath As String
            'sfd.FileName = Me.Text
            'sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
            'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            '    filePath = sfd.FileName
            'Else
            '    Exit Sub
            'End If
            transportSql.applyExportTemplate(gv, PageSetupReport_ID)
            transportSql.QuickExportToExcel(gv, "", Me.Text, , arrHeader)
            'transportSql.exportdataChilRows(gv, filePath, filePath.Substring(filePath.LastIndexOf("\") + 1, filePath.Length - filePath.LastIndexOf("\") - 1), , arrHeader)
            'common.clsCommon.MyMessageBoxShow("Exported Successfully.")
            'Process.Start(filePath)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs)
        Try
            Dim obj As New clsPSShipmentHead
            obj.CreateEmailContent("ORDBOOK00008", Nothing)
            'For Each grw As GridViewRowInfo In gv.Rows
            '    ''Distributer_Dispatch_No
            '    obj.CreateEmailContent("ORDBOOK00008", Nothing)
            'Next
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub BtnPrintChallan_Click(sender As Object, e As EventArgs) Handles BtnPrintChallan.Click
        Try
            loadData(2)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtMultCustomer__My_Click(sender As Object, e As EventArgs) Handles txtMultCustomer._My_Click
        Try
            strQry = " select Cust_Code as [Code] , Customer_Name as [Name] from TSPL_CUSTOMER_MASTER "
            txtMultCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("txtMultCustomer", strQry, "Code", "Name", txtMultCustomer.arrValueMember, txtMultCustomer.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtMultLocation__My_Click(sender As Object, e As EventArgs) Handles txtMultLocation._My_Click
        Try
            strQry = "select Location_Code  as [Code],Location_Desc as [Name] from TSPL_LOCATION_MASTER where Type='PLANT' "
            txtMultLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("txtMultLocation", strQry, "Code", "Name", txtMultLocation.arrValueMember, txtMultLocation.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtMultRoute__My_Click(sender As Object, e As EventArgs) Handles txtMultRoute._My_Click
        Try
            strQry = "select TSPL_ROUTE_MASTER.Route_No as Code,TSPL_ROUTE_MASTER.Route_Desc as Name from TSPL_ROUTE_MASTER "
            txtMultRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("txtMultRoute", strQry, "Code", "Name", txtMultRoute.arrValueMember, txtMultRoute.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    'Private Sub RadPageViewPage3_Paint(sender As Object, e As PaintEventArgs) Handles RadPageViewPage3.Paint

    'End Sub

    'Private Sub RadGroupBox3_Click(sender As Object, e As EventArgs) Handles RadGroupBox3.Click

    'End Sub

    'Private Sub MyLabel1_Click(sender As Object, e As EventArgs) Handles MyLabel1.Click

    'End Sub

    'Private Sub lblRoute_Click(sender As Object, e As EventArgs) Handles lblRoute.Click

    'End Sub

    'Private Sub RadPageViewPage1_Paint(sender As Object, e As PaintEventArgs) Handles RadPageViewPage1.Paint

    'End Sub

    Private Sub PDF_Click(sender As Object, e As EventArgs) Handles PDF.Click
        Try
            If gv.Rows.Count > 0 Then
                ExportToExcel(EnumExportTo.PDF)
            Else
                RadMessageBox.Show("No Data Found to Display", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiSaveLayout_Click(sender As Object, e As EventArgs) Handles rmiSaveLayout.Click
        Try
            If clsCommon.myLen(PageSetupReport_ID) > 0 Then
                gv.MasterTemplate.FilterDescriptors.Clear()
                Dim obj As New clsGridLayout()
                obj.ReportID = PageSetupReport_ID
                obj.UserID = objCommonVar.CurrentUserCode
                obj.GridLayout = New MemoryStream()
                gv.SaveLayout(obj.GridLayout)
                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                obj.GridColumns = gv.ColumnCount
                If obj.SaveData() Then
                    common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
                End If

                obj.GridLayout.Close()
                obj.GridLayout.Dispose()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmiDeleteLayout.Click
        Try
            clsGridLayout.DeleteData(PageSetupReport_ID, objCommonVar.CurrentUserCode)
            common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub checkSendEmailSMS(Optional ByVal isPrintOrSend As Boolean = True)
        Dim strDate As New List(Of String)
        Dim strCustCode As New List(Of String)
        Dim strTaxableNonTaxable As String = Nothing

        If rbtnPartyWise.Checked Then
            For Each grow As GridViewRowInfo In gv.Rows
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    strDate.Add("Convert(Date,'" + clsCommon.myCstr(grow.Cells("Supply_Date").Value) + "',103)")
                    strCustCode.Add(clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                    'If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("IsTaxable").Value), "Non-Taxable") = CompairStringResult.Equal Then
                    '    If clsCommon.myLen(strTaxableNonTaxable) > 0 AndAlso Not strTaxableNonTaxable.Contains("0") Then
                    '        strTaxableNonTaxable += ",'0'"
                    '    ElseIf clsCommon.myLen(strTaxableNonTaxable) <= 0 Then
                    '        strTaxableNonTaxable = "'0'"
                    '    End If
                    'End If
                    'If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("IsTaxable").Value), "Taxable") = CompairStringResult.Equal Then
                    '    If clsCommon.myLen(strTaxableNonTaxable) > 0 AndAlso Not strTaxableNonTaxable.Contains("1") Then
                    '        strTaxableNonTaxable += ",'1'"
                    '    ElseIf clsCommon.myLen(strTaxableNonTaxable) <= 0 Then
                    '        strTaxableNonTaxable = "'1'"
                    '    End If
                    'End If
                End If
            Next

            Dim strQry As String = ReturnLoadReportQry()
            strQry += " and CONVERT(date,Supply_Date,103) in (" + clsCommon.GetMulcallStringWithComma(strDate) + ") and TSPL_CUSTOMER_MASTER.Cust_Code In (" + clsCommon.GetMulcallString(strCustCode) + ") "
            'strQry +=" And TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable in (" + strTaxableNonTaxable + ") "
            Dim dtt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
                lstinvNo = New List(Of String)
                For Each rows In dtt.Rows
                    'clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Printing " & (ii) & " Of " & Total)
                    lstinvNo.Add(clsCommon.myCstr(rows("Document_Code")))
                    'ii += 1
                Next
            End If

            If isPrintOrSend Then
                Dim ChkQry As String = CheckEmailOrSMS(clsCommon.GetMulcallString(lstinvNo))
                dtt = Nothing
                dtt = clsDBFuncationality.GetDataTable(ChkQry)

                If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
                    If chkbtnEmailSMS Then
                        If clsCommon.MyMessageBoxShow(Me, "Already send email is " + clsCommon.myCstr(dtt.Rows(0)("AlreadySendEmailParty")) + " and remaining for send email is " + clsCommon.myCstr(dtt.Rows(0)("RemainingParty")) + "." + Environment.NewLine + "Are you sure to send email for all ?", Me.Text, MessageBoxButtons.YesNo) = DialogResult.No Then
                            strQry += " and TSPL_SD_SALE_INVOICE_HEAD.Send_Email is Null"
                            dtt = Nothing
                            dtt = clsDBFuncationality.GetDataTable(strQry)
                            If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
                                lstinvNo = Nothing
                                lstinvNo = New List(Of String)
                                For Each rows In dtt.Rows
                                    'clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Printing " & (ii) & " Of " & Total)
                                    lstinvNo.Add(clsCommon.myCstr(rows("Document_Code")))
                                    'ii += 1
                                Next
                            Else
                                Throw New Exception("Data not found to Send !")
                            End If
                        End If
                    Else
                        If clsCommon.MyMessageBoxShow(Me, "Already send sms is " + clsCommon.myCstr(dtt.Rows(0)("AlreadySendSMSParty")) + " and remaining for send sms is " + clsCommon.myCstr(dtt.Rows(0)("RemainingSMSParty")) + "." + Environment.NewLine + "Are you sure to send sms for all ?", Me.Text, MessageBoxButtons.YesNo) = DialogResult.No Then
                            strQry += " and TSPL_SD_SALE_INVOICE_HEAD.Send_SMS is Null"
                            dtt = Nothing
                            dtt = clsDBFuncationality.GetDataTable(strQry)
                            If dtt IsNot Nothing AndAlso dtt.Rows.Count > 0 Then
                                lstinvNo = Nothing
                                lstinvNo = New List(Of String)
                                For Each rows In dtt.Rows
                                    'clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Printing " & (ii) & " Of " & Total)
                                    lstinvNo.Add(clsCommon.myCstr(rows("Document_Code")))
                                    'ii += 1
                                Next
                            Else
                                Throw New Exception("Data not found to Send !")
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    'Ticket No-ERO/03/09/19-001018,Send Email with Invoice PDF attachment
    Private Sub BtnEmailSms_Click(sender As Object, e As EventArgs) Handles BtnEmailSms.Click
        Try
            If Not clsCommon.CompairString(cboReportType.SelectedValue, "AL") = CompairStringResult.Equal Then
                clsCommon.MyMessageBoxShow(Me, "Select All in invoice type", Me.Text)
                Exit Sub
            End If
            chkbtnEmailSMS = True
            Dim dtContent As DataTable
            If rbtnPartyWise.Checked Then
                dtContent = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "2'", Nothing)
            Else
                dtContent = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "1'", Nothing)
            End If
            If dtContent Is Nothing OrElse dtContent.Rows.Count = 0 Then
                clsCommon.MyMessageBoxShow(Me, "First do email setting", Me.Text)
                Exit Sub
            End If
            clsCommon.ProgressBarPercentShow()
            Try
                checkSendEmailSMS()
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
            Dim ii As Integer = 1
            Dim Total As Integer = 0
            For Each grow As GridViewRowInfo In gv.Rows
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    'clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Processing..." & (ii) & " Of " & Total)
                    Total += 1
                End If
                'ii += 1
            Next
            Dim istaxable As String = Nothing
            For Each grow As GridViewRowInfo In gv.Rows
                clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Processing..." & (ii) & " Of " & Total)
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    Dim objEmailH As New clsEMailHead()
                    objEmailH.arrEMail = New List(Of String)()
                    If rbtnInvoiceWise.Checked Then
                        Dim routeName = clsDBFuncationality.getSingleValue("SELECT Route_Desc from TSPL_ROUTE_MASTER where Route_No='" + clsCommon.myCstr(grow.Cells("Route_Code").Value) + "'", Nothing)

                        objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(grow.Cells("Document_Code").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(grow.Cells("Document_Date").Value, "dd/MMM/yyyy"))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells("Customer_Name").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells("Total_Amt").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells("Location_Desc").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells("Shift_type").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells("Supply_Date").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells("Route_Code").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.RouteName, clsCommon.myCstr(routeName))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)

                        objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(grow.Cells("Document_Code").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(grow.Cells("Document_Date").Value, "dd/MMM/yyyy"))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells("Customer_Name").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells("Total_Amt").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells("Location_Desc").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells("Shift_type").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells("Supply_Date").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells("Route_Code").Value))

                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.RouteName, clsCommon.myCstr(routeName))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)
                        '------------------------code for attchament-------------------------------------
                        strRptPath = ""
                        strRptPath = Printing(clsCommon.myCstr(grow.Cells("Document_Code").Value), True, False, True)
                    Else
                        Dim strCustGrp As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Group_Code from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(grow.Cells("Customer_Code").Value) + "'"))

                        objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells("Customer_Name").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells("Total_Amt").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells("Location_Desc").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells("Shift_type").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells("Supply_Date").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells("Route_Code").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.CustomerGroup, strCustGrp)
                        'objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.InvoiceType, clsCommon.myCstr(grow.Cells("IsTaxable").Value))
                        objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)

                        objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells("Customer_Name").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells("Total_Amt").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells("Location_Desc").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells("Shift_type").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells("Supply_Date").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells("Route_Code").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.CustomerGroup, strCustGrp)
                        'objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.InvoiceType, clsCommon.myCstr(grow.Cells("IsTaxable").Value))
                        objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)
                        '------------------------code for attchament-------------------------------------
                        'istaxable = Nothing
                        'If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("IsTaxable").Value), "Taxable") = CompairStringResult.Equal Then
                        '    istaxable = "T"
                        'Else
                        '    istaxable = "NT"
                        'End If
                        strRptPath = ""
                        strRptPath = Printing(Nothing, True, False, True, clsCommon.myCstr(grow.Cells("Supply_Date").Value), clsCommon.myCstr(grow.Cells("Shift_type").Value), istaxable, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                    End If

                    objEmailH.Attachment_1_Path = strRptPath
                    '---------------------------------------------------------------------------
                    'Dim Data As Attachment = New Attachment(objEmailH.Attachment_1_Path, MediaTypeNames.Application.Octet)

                    Dim emailId As String = ""
                    emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & clsCommon.myCstr(grow.Cells("Customer_Code").Value) & "' ")

                    If clsCommon.myLen(emailId) > 0 Then
                        objEmailH.arrEMail.Add(emailId)
                    End If

                    If rbtnInvoiceWise.Checked Then
                        objEmailH.SaveData(clsUserMgtCode.FrmPrintDistributerInvoiceStatement, objEmailH, Nothing)
                    Else
                        objEmailH.SaveData(clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "2", objEmailH, Nothing)
                        clsDBFuncationality.ExecuteNonQuery("Update TSPL_SD_SALE_INVOICE_HEAD Set Send_Email=1 Where Document_Code In (" + clsCommon.GetMulcallString(lstinvNo) + ") and Customer_Code='" + clsCommon.myCstr(grow.Cells("Customer_Code").Value) + "' and Is_Taxable='" + clsCommon.myCstr(IIf(clsCommon.CompairString(istaxable, "T") = CompairStringResult.Equal, 1, 0)) + "'")
                    End If
                    'Dim MailMsg As New MailMessage()
                    'MailMsg.Subject = objEmailH.Email_Subject
                    'MailMsg.From = New MailAddress("mohdsuhail0677@gmail.com")
                    'MailMsg.To.Add(clsCommon.GetMulcallStringWithComma(objEmailH.arrEMail))
                    'MailMsg.Body = objEmailH.Email_Text
                    'MailMsg.Priority = MailPriority.High
                    'MailMsg.IsBodyHtml = False
                    'MailMsg.Attachments.Add(Data)
                    'Dim SmtpMail As New SmtpClient("smtp.gmail.com")
                    'SmtpMail.Port = clsCommon.myCdbl(587)
                    'SmtpMail.Credentials = New System.Net.NetworkCredential("mohdsuhail0677@gmail.com", "civd wxfw eehy vsoz")
                    'SmtpMail.EnableSsl = True

                    'SmtpMail.Send(MailMsg)
                    'objEmailH = Nothing
                    ii += 1
                End If
            Next
            'clsCommon.ProgressBarUpdate("Gathering information regarding PF Rules...")
            clsCommon.ProgressBarPercentHide()
            clsCommon.MyMessageBoxShow(Me, "E-Mail Send Successfully", Me.Text)
            chkbtnEmailSMS = False
        Catch ex As Exception
            chkbtnEmailSMS = False
            clsCommon.ProgressBarPercentHide()
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            'common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error, Me.Text)
        End Try
    End Sub

    Private Sub cboReportType_SelectedIndexChanged(sender As Object, e As UI.Data.PositionChangedEventArgs) Handles cboReportType.SelectedIndexChanged
        Try
            If isInsideLoadData = False Then
                If clsCommon.CompairString(cboReportType.SelectedValue, "NT") = CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then
                    btnPrePrintFormat.Visible = True
                Else
                    btnPrePrintFormat.Visible = False
                End If
                If (clsCommon.CompairString(cboReportType.SelectedValue, "NT") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboReportType.SelectedValue, "T") = CompairStringResult.Equal OrElse clsCommon.CompairString(cboReportType.SelectedValue, "IT") = CompairStringResult.Equal) AndAlso clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "BHAD") = CompairStringResult.Equal Then
                    btnBatchWiseInvoice.Visible = True
                Else
                    btnBatchWiseInvoice.Visible = False
                End If
                If clsCommon.CompairString(cboReportType.SelectedValue, "AL") = CompairStringResult.Equal Then
                    rbtnPartyWise.Enabled = True
                Else
                    rbtnPartyWise.Enabled = False
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrePrintFormat_Click(sender As Object, e As EventArgs) Handles btnPrePrintFormat.Click
        Try
            Printing("", False, True)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnBatchWiseInvoice_Click(sender As Object, e As EventArgs) Handles btnBatchWiseInvoice.Click
        Try
            Dim InvoiceNo As String = ""
            For Each grow As GridViewRowInfo In gv.Rows
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    InvoiceNo = InvoiceNo + "','" + clsCommon.myCstr(grow.Cells("Document_Code").Value)
                End If
            Next

            If clsCommon.myLen(InvoiceNo) > 0 AndAlso clsCommon.myCstr(InvoiceNo).Substring(0, 3) = "','" Then
                InvoiceNo = InvoiceNo.Substring(3, InvoiceNo.Length - 3)

            End If
            clsPSShipmentPrint.PrintDataBatchWiseInvoice(MyBase.Form_ID, InvoiceNo)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnCombinedInvoice_Click(sender As Object, e As EventArgs) Handles btnCombinedInvoice.Click
        Try
            If txtFromDate.Value > txtToDate.Value Then
                common.clsCommon.MyMessageBoxShow(Me, "From date can not be greater then to Date", Me.Text)
                txtFromDate.Focus()
                Exit Sub
            End If

            Dim WhrCls As String = ""
            If txtMultCustomer.arrValueMember IsNot Nothing AndAlso txtMultCustomer.arrValueMember.Count > 0 Then
                WhrCls += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultCustomer.arrValueMember) + ") "
            End If
            If txtMultRoute.arrValueMember IsNot Nothing AndAlso txtMultRoute.arrValueMember.Count > 0 Then
                WhrCls += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in (" + clsCommon.GetMulcallString(txtMultRoute.arrValueMember) + ") "
            End If
            If txtMultLocation.arrValueMember IsNot Nothing AndAlso txtMultLocation.arrValueMember.Count > 0 Then
                WhrCls += " and TSPL_LOCATION_MASTER. Location_Code   IN (" + clsCommon.GetMulcallString(txtMultLocation.arrValueMember) + ") "
            End If
            Dim Qry As String = " Select TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE + ' Customer :'+TSPL_SD_SALE_INVOICE_HEAD.Customer_Code + '('+TSPL_CUSTOMER_MASTER.Customer_Name +')' + ' Route :'+ TSPL_SD_SALE_INVOICE_HEAD.Route_No + '('+TSPL_ROUTE_MASTER.Route_Desc+')' as Group_column1, TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE , Convert (varchar, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_DETAIL.Location,TSPL_LOCATION_MASTER.Location_Desc, TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.Add1 as Customer_Add1, TSPL_CUSTOMER_MASTER.Add2 as Customer_Add2, TSPL_CUSTOMER_MASTER.Add3 as Customer_Add3,TSPL_SD_SALE_INVOICE_HEAD.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_SD_SALE_INVOICE_DETAIL.Line_No, TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_MASTER.ITEM_DESC,TSPL_SD_SALE_INVOICE_DETAIL.Qty, TSPL_SD_SALE_INVOICE_DETAIL.Item_Cost,TSPL_SD_SALE_INVOICE_DETAIL.Amount, TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1 as Comp_Add1,TSPL_COMPANY_MASTER.Add2 as Comp_Add2 ,TSPL_COMPANY_MASTER.Add3 as Comp_Add3, TSPL_COMPANY_MASTER.City_Code as Comp_City_Code,TSPL_COMPANY_MASTER.Fax as Comp_Fax,TSPL_COMPANY_MASTER.Email as Comp_Email,TSPL_COMPANY_MASTER.Pincode as Comp_Pincode ,TSPL_COMPANY_MASTER.State,TSPL_COMPANY_MASTER.Pan_No as Comp_Pan_No, TSPL_COMPANY_MASTER.GSTReg_No as Comp_GSTIN_No,TSPL_COMPANY_MASTER.Access_Officer as Comp_FASSAI_NO , TSPL_COMPANY_MASTER.Phone1 as Comp_Phone1 , TSPL_COMPANY_MASTER.Phone2 as Comp_Phone2 " &
                                " from TSPL_SD_SALE_INVOICE_DETAIL " &
                                " Left Outer Join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE " &
                                " Left Outer Join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.ITEM_CODE = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code " &
                                " Left Outer Join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code " &
                                " Left Outer Join TSPL_ROUTE_MASTER on  TSPL_ROUTE_MASTER.Route_No = TSPL_SD_SALE_INVOICE_HEAD.Route_No " &
                                " Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = TSPL_SD_SALE_INVOICE_HEAD.Comp_Code " &
                                " Left Outer Join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_DETAIL.Location " &
                                " where TSPL_SD_SALE_INVOICE_HEAD.Status = 1 and TSPL_SD_SALE_INVOICE_HEAD.Screen_Type = 'DS'  and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' ,103)    " + WhrCls + " " &
                                " Order By TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE , TSPL_SD_SALE_INVOICE_DETAIL.Line_No  "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim obj As clsDosPrint = New clsDosPrint()
                obj.ReportName = objCommonVar.CurrentCompanyName
                obj.ReportName1 = "Customer Copy"
                obj.ShowPageNo = True
                obj.arrFilter = New List(Of clsDosPrintHeaderFilter)()
                obj.arrFilter.Add(clsDosPrintHeaderFilter.GetObject("Phone No", clsCommon.myCstr(dt.Rows(0)("Comp_Phone1"))))
                obj.arrFilter.Add(clsDosPrintHeaderFilter.GetObject("Email", clsCommon.myCstr(dt.Rows(0)("Comp_Email"))))
                obj.arrFilter.Add(clsDosPrintHeaderFilter.GetObject("GSTIN NO", clsCommon.myCstr(dt.Rows(0)("Comp_GSTIN_No"))))
                obj.arrFilter.Add(clsDosPrintHeaderFilter.GetObject("FASSAI NO", clsCommon.myCstr(dt.Rows(0)("Comp_FASSAI_NO"))))
                obj.arrFilter.Add(clsDosPrintHeaderFilter.GetObject("FASSAI NO", clsCommon.myCstr(dt.Rows(0)("Comp_FASSAI_NO"))))

                obj.arrGroup = New List(Of clsDosPrintGroup)()
                obj.arrGroup.Add(clsDosPrintGroup.GetObject("DOCUMENT_CODE", "Invoice", ""))
                'obj.arrGroup.Add(clsDosPrintGroup.GetObject("Group_column1", "Invoice", ""))


                obj.arrColumn = New List(Of clsDosPrintColumn)()
                obj.arrColumn.Add(clsDosPrintColumn.SetColumn("Line_No", "SNo", True, DosPrintAlignment.Left, 5, False, DecimalPlaces.NA))
                obj.arrColumn.Add(clsDosPrintColumn.SetColumn("Item_Code", "Product", True, DosPrintAlignment.Left, 15, False, DecimalPlaces.NA))
                obj.arrColumn.Add(clsDosPrintColumn.SetColumn("ITEM_DESC", "Description", True, DosPrintAlignment.Left, 20, False, DecimalPlaces.NA))
                obj.arrColumn.Add(clsDosPrintColumn.SetColumn("Qty", "Quantity", True, DosPrintAlignment.Left, 20, False, DecimalPlaces.NA))
                obj.arrColumn.Add(clsDosPrintColumn.SetColumn("Item_Cost", "UnitPrice(Rs)", False, DosPrintAlignment.Right, 20, True, DecimalPlaces.NA))
                obj.arrColumn.Add(clsDosPrintColumn.SetColumn("Amount", "Amount(Rs)", False, DosPrintAlignment.Right, 20, True, DecimalPlaces.Two))
                obj.Print(obj, dt, PageSetup.Potrate)
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not Found!", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndCustom__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndCustom._MYValidating
        Try
            Dim strQuery As String = " select Cust_Code as Code , Customer_Name as Name from TSPL_CUSTOMER_MASTER "
            fndCustom.Value = clsCommon.ShowSelectForm("PrintSaleInvoiceMonthly", strQuery, "Code", "", fndCustom.Value, "Code", isButtonClicked)
            lblCustomer.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code = '" + fndCustom.Value + "' "))
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnSupplyDate_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnSupplyDate.CheckedChanged
        Try
            If rbtnSupplyDate.Checked Then
                RadGroupBox4.Visible = True
            End If
            If rbtnDocumentDate.Checked Then
                RadGroupBox4.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub BtnSMS_Click(sender As Object, e As EventArgs) Handles BtnSMS.Click
        Try
            If Not clsCommon.CompairString(cboReportType.SelectedValue, "AL") = CompairStringResult.Equal Then
                clsCommon.MyMessageBoxShow(Me, "Select All in invoice type", Me.Text)
                Exit Sub
            End If

            chkbtnEmailSMS = False
            'Dim dtContent1 As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "'", Nothing)

            Dim dtContent1 As DataTable
            If rbtnPartyWise.Checked Then
                dtContent1 = clsDBFuncationality.GetDataTable("SELECT SMS_Text from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "2'", Nothing)
            Else
                dtContent1 = clsDBFuncationality.GetDataTable("SELECT SMS_Text from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "1'", Nothing)
            End If

            If dtContent1 Is Nothing OrElse dtContent1.Rows.Count = 0 Then
                'clsCommon.ProgressBarPercentHide()
                clsCommon.MyMessageBoxShow(Me, "First do SMS setting", Me.Text)
                Exit Sub
            End If
            clsCommon.ProgressBarPercentShow()
            Try
                checkSendEmailSMS()
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
            Dim ii As Integer = 1
            Dim Total As Integer = 0
            For Each grow As GridViewRowInfo In gv.Rows
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    'clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Processing..." & (ii) & " Of " & Total)
                    Total += 1
                End If
                'ii += 1
            Next

            'For Each grow As GridViewRowInfo In gv.Rows
            '    If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
            '        Dim objSMSH As New clsSMSHead()
            '        objSMSH.arrSMS = New List(Of String)()
            '        objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(grow.Cells(("Document_Code")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(grow.Cells(("Document_Date")).Value, "dd/MMM/yyyy"))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells(("Customer_Code")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells(("Customer_Name")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells(("Total_Amt")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells(("Location_Desc")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells(("Shift_type")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells(("Supply_Date")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells(("Route_Code")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.RouteName, clsCommon.myCstr(grow.Cells(("Route_Name")).Value))
            '        objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)
            '        'strRptPath = Printing(clsCommon.myCstr(grow.Cells("Document_Code").Value), True)
            '        Dim SMS_Text As String = ""
            '        SMS_Text = clsDBFuncationality.getSingleValue("select email from TSPL_customer_MASTER where cust_code ='" & clsCommon.myCstr(grow.Cells("Customer_Code").Value) & "' ")

            '        If clsCommon.myLen(SMS_Text) > 0 Then
            '            objSMSH.arrEMail.Add(SMS_Text)
            '        End If
            '        objSMSH.SendSMS(clsUserMgtCode.FrmPrintDistributerInvoiceStatement, objSMSH, Nothing)
            '    End If
            'Next
            ''clsCommon.ProgressBarUpdate("Gathering information regarding PF Rules...")
            'clsCommon.MyMessageBoxShow("SMS Send Successfully", Me.Text)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Shaurya Rajput
            Dim istaxable As String = Nothing
            For Each grow As GridViewRowInfo In gv.Rows
                clsCommon.ProgressBarPercentUpdate((ii) * 100 / Total, " Processing..." & (ii) & " Of " & Total)
                If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                    Dim strContactPerson As String = ""
                    Dim strotherno As String = Nothing
                    strotherno = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Phone1 from TSPL_customer_MASTER where cust_code ='" & clsCommon.myCstr(grow.Cells("Customer_Code").Value) & "'"))

                    'Dim emailId As String = ""
                    'emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & clsCommon.myCstr(grow.Cells("Customer_Code").Value) & "'")

                    'If clsCommon.myLen(emailId) > 0 Then
                    '    objEmailH.arrEMail.Add(emailId)
                    'End If


                    Dim dtContent As DataTable
                    If rbtnPartyWise.Checked Then
                        dtContent = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "2'", Nothing)
                    Else
                        dtContent = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "1'", Nothing)
                    End If
                    Dim objSMSH As New clsSMSHead()
                    objSMSH.arrMobilNo = New List(Of String)()
                    objSMSH.arrMobilNo.Add(clsCommon.myCstr(strotherno))
                    If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
                        If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                            objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))
                            If rbtnInvoiceWise.Checked Then
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(grow.Cells("Document_Code").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(grow.Cells("Document_Date").Value, "dd/MMM/yyyy"))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells("Customer_Name").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells("Total_Amt").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells("Location_Desc").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells("Shift_type").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells("Supply_Date").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells("Route_Code").Value))
                                Dim routeName = clsDBFuncationality.getSingleValue("SELECT Route_Desc from TSPL_ROUTE_MASTER where Route_No='" + clsCommon.myCstr(grow.Cells("Route_Code").Value) + "'", Nothing)
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.RouteName, clsCommon.myCstr(routeName))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)
                            Else
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerNo, clsCommon.myCstr(grow.Cells("Customer_Code").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerName, clsCommon.myCstr(grow.Cells("Customer_Name").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.TotalAmount, clsCommon.myCstr(grow.Cells("Total_Amt").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.LocationName, clsCommon.myCstr(grow.Cells("Location_Desc").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SupplyShift, clsCommon.myCstr(grow.Cells("Shift_type").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.SupplyDate, clsCommon.myCstr(grow.Cells("Supply_Date").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Route, clsCommon.myCstr(grow.Cells("Route_Code").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.CustomerGroup, clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Group_Code from TSPL_CUSTOMER_MASTER where Cust_Code='" + clsCommon.myCstr(grow.Cells("Customer_Code").Value) + "'")))
                                'objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.InvoiceType, clsCommon.myCstr(grow.Cells("IsTaxable").Value))
                                objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Form_Code, Me.Form_ID)

                                'istaxable = Nothing
                                'If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("IsTaxable").Value), "Taxable") = CompairStringResult.Equal Then
                                '    istaxable = "T"
                                'Else
                                '    istaxable = "NT"
                                'End If
                            End If
                        End If
                        If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                            If rbtnInvoiceWise.Checked Then
                                objSMSH.SaveData(clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "1", objSMSH, Nothing)
                            Else
                                objSMSH.SaveData(clsUserMgtCode.FrmPrintDistributerInvoiceStatement + "2", objSMSH, Nothing)
                                clsDBFuncationality.ExecuteNonQuery("Update TSPL_SD_SALE_INVOICE_HEAD Set Send_SMS=1 Where Document_Code In (" + clsCommon.GetMulcallString(lstinvNo) + ") and Customer_Code='" + clsCommon.myCstr(grow.Cells("Customer_Code").Value) + "' ")
                            End If
                            objSMSH = Nothing
                        End If
                    End If
                    ii += 1
                End If
            Next
            clsCommon.ProgressBarPercentHide()
            clsCommon.MyMessageBoxShow(Me, "SMS Send Successfully", Me.Text)
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub TxtItem__My_Click(sender As Object, e As EventArgs) Handles TxtItem._My_Click
        Try
            Dim qry As String = " Select Item_Code,Item_Desc,Short_Description,Structure_Code from tspl_item_master "
            TxtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("ItemDMulSel", qry, "Item_Code", "Short_Description", TxtItem.arrValueMember, TxtItem.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnPartyWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnPartyWise.CheckedChanged
        Try
            InvoiceAndPartyWiseCheck()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnInvoiceWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnInvoiceWise.CheckedChanged
        Try
            InvoiceAndPartyWiseCheck()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub InvoiceAndPartyWiseCheck()
        If rbtnPartyWise.Checked Then
            rbtnDocumentDate.Enabled = False
            rbtnSupplyDate.Checked = True
        Else
            rbtnDocumentDate.Enabled = True
        End If
    End Sub
End Class
