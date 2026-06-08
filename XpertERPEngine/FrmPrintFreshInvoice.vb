'--Rollback and Add Sanjeet work by balwinder on 29/01/2018

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Imports common
Imports System.IO
Public Class FrmPrintFreshInvoice
    Inherits FrmMainTranScreen
    Dim CreateFreshInvoiceOnDispatchSave As Integer = 0
    Dim LeakageDeduction_Freshsale As Double = 0
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim isInsideLoadData As Boolean = False
    'Dim Refresh As Boolean = False
    Public Shared ArrInvoice_Arr As New ArrayList()
    'update by preeti gupta against ticket no[BM00000008457,BHA/18/12/18-000759]
    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.FrmPrintFreshInvoice)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        RadSplitButton1.Visible = MyBase.isExport
        BtnPrint.Visible = MyBase.isPrintFlag
        'btnQuickExport.Visible = MyBase.isQuickExportFlag

    End Sub
    Private Sub FrmPrintFreshInvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetUserMgmtNew()
        CreateFreshInvoiceOnDispatchSave = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateFreshInvoiceOnDispatchSave, clsFixedParameterCode.CreateFreshInvoiceOnDispatchSave, Nothing))


        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnGo, "Press Alt+R Refresh ")
        ButtonToolTip.SetToolTip(BtnPrint, "Press Alt+P For Print")
        ButtonToolTip.SetToolTip(BtnReset, "Press Alt+N Adding New")
        Reset()
    End Sub
    Sub LoadLocation()
        Dim qry As String = "select Location_Code as [Code] ,Location_Desc as [Name] from TSPL_LOCATION_MASTER  "
        cbgLocation.DataSource = clsDBFuncationality.GetDataTable(qry)
        cbgLocation.ValueMember = "Code"
        cbgLocation.DisplayMember = "Name"

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
    Private Sub rmSaveLayout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rmSaveLayout.Click
        If clsCommon.myLen(MyBase.Form_ID) > 0 Then
            gv.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = MyBase.Form_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv.SaveLayout(obj.GridLayout)
            obj.GridColumns = gv.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            ''stuti regarding memory leakage
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub rmDeleteLayout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        PageSetupReport_ID = MyBase.Form_ID
        TemplateGridview = gv
        loadReport()
    End Sub
    Sub Reset()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = txtToDate.Value.AddMonths(-1)
        'LoadLocation()
        'chkLocationAll.CheckState = CheckState.Checked
        RadPageView1.SelectedPage = RadPageViewPage1
        gv.DataSource = Nothing
        gv.Rows.Clear()

    End Sub
    Private Sub BtnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnReset.Click
        Reset()
    End Sub
    Sub print(ByVal exporter As EnumExportTo)
        Try
            Dim arrHeader As List(Of String) = New List(Of String)()
            arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")) + " ")
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            'If chkLocationSelect.IsChecked Then
            '    Dim strLocationName As String = ""
            '    For Each StrName As String In cbgLocation.CheckedDisplayMember
            '        If clsCommon.myLen(strLocationName) > 0 Then
            '            strLocationName += ", "
            '        End If
            '        strLocationName += StrName
            '    Next
            '    Dim strLocationCode As String = ""
            '    For Each StrCode As String In cbgLocation.CheckedValue
            '        If clsCommon.myLen(strLocationCode) > 0 Then
            '            strLocationCode += ", "
            '        End If
            '        strLocationCode += StrCode
            '    Next
            '    arrHeader.Add((" Location Name: " + strLocationName + " "))
            'End If
            If txtLocation1Mult.arrValueMember IsNot Nothing AndAlso txtLocation1Mult.arrValueMember.Count > 0 Then
                Dim strLocationName As String = clsCommon.GetMulcallStringWithComma(txtLocation1Mult.arrValueMember)
                arrHeader.Add((" Location : " + strLocationName + " "))
            Else
                arrHeader.Add((" Location : All"))
            End If
            If txtCustomerMult.arrValueMember IsNot Nothing AndAlso txtCustomerMult.arrValueMember.Count > 0 Then
                arrHeader.Add(" Customer : " + clsCommon.GetMulcallStringWithComma(txtCustomerMult.arrValueMember))
            Else
                arrHeader.Add((" Customer: All"))
            End If
            If txtCustomerCategoryMult.arrValueMember IsNot Nothing AndAlso txtCustomerCategoryMult.arrValueMember.Count > 0 Then
                arrHeader.Add("Customer Category : " + clsCommon.GetMulcallStringWithComma(txtCustomerCategoryMult.arrValueMember))
            Else
                arrHeader.Add(("Customer Category: All"))
            End If

            If exporter = EnumExportTo.Excel Then
                clsCommon.MyExportToExcelGrid("Fresh Invoice", gv, arrHeader, Me.Text)
            Else
                clsCommon.MyExportToPDF("Fresh Invoice", gv, arrHeader, Me.Text, True)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub
    Private Sub rmExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rmExcel.Click
        Try

            Dim arrHeader As List(Of String) = New List(Of String)()

            arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")) + " ")
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.FrmPrintFreshInvoice & "'"))

            If txtLocation1Mult.arrValueMember IsNot Nothing AndAlso txtLocation1Mult.arrValueMember.Count > 0 Then
                Dim strLocationName As String = clsCommon.GetMulcallStringWithComma(txtLocation1Mult.arrValueMember)
                arrHeader.Add((" Location : " + strLocationName + " "))
            Else
                arrHeader.Add((" Location : All"))
            End If
            If txtCustomerMult.arrValueMember IsNot Nothing AndAlso txtCustomerMult.arrValueMember.Count > 0 Then
                arrHeader.Add(" Customer : " + clsCommon.GetMulcallStringWithComma(txtCustomerMult.arrDispalyMember))
            End If
            If txtCustomerCategoryMult.arrValueMember IsNot Nothing AndAlso txtCustomerCategoryMult.arrValueMember.Count > 0 Then
                arrHeader.Add(" Customer Category : " + clsCommon.GetMulcallStringWithComma(txtCustomerCategoryMult.arrDispalyMember))
            End If
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

    Public Sub loadReport()
        ''changes by Shivani [BM00000007756],BM00000008110
        If txtFromDate.Value > txtToDate.Value Then
            common.clsCommon.MyMessageBoxShow(Me, "From date can not be greater then to Date", Me.Text)
            txtFromDate.Focus()
            Exit Sub
        End If
        If chkLocationSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select atleast single Location or select all.", Me.Text)
            Exit Sub
        End If
        Dim sQuery As String = " select  Cast(1 as BIT) as 'Check',case when TSPL_SD_SHIPMENT_HEAD.ManualVehicle <> '' then '' else TSPL_SD_SALE_INVOICE_HEAD.Vehicle_Code end as Vehicle_Code,case when TSPL_SD_SHIPMENT_HEAD.ManualVehicle <> '' then ManualVehicle else  TSPL_SD_SALE_INVOICE_HEAD.vehicleNo end as vehicleNo, TSPL_COMPANY_MASTER.Comp_Name,TSPL_LOCATION_MASTER.Add1 as Loc_Add1,TSPL_LOCATION_MASTER.Add2  as Loc_Add2,TSPL_LOCATION_MASTER.Add3  as Loc_Add3,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2    , TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_CITY_MASTER_For_Comp.City_Name )>0 then ', '+TSPL_CITY_MASTER_For_Comp.City_Name else ' ' end + case when len(TSPL_STATE_MASTER.STATE_NAME  )>0 then TSPL_STATE_MASTER.STATE_NAME  else '' end +    case when LEN(TSPL_COMPANY_MASTER.Tin_No  )>0 then ', '+TSPL_COMPANY_MASTER.Tin_No else ' ' end  as Comp_address  ,ISNULL(tspl_company_Master.ADD2,'')   as Compaddress2,ISNULL(tspl_company_Master.ADD3,'') as Compaddress3,  TSPL_SD_SALE_INVOICE_HEAD.Document_Code as InvoiceNo,convert(varchar ,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as InvoiceDate ,TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location, TSPL_LOCATION_MASTER.Location_Desc , TSPL_COMPANY_MASTER.Comp_Name as CompName,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,  TSPL_CUSTOMER_MASTER.Customer_Name as CustName,cust_category_desc,TSPL_CUSTOMER_MASTER.Add1 as Customer_Add1,TSPL_CUSTOMER_MASTER.add2 as   customer_Add2,TSPL_CUSTOMER_MASTER.Add3 as customer_Add3,TSPL_CUSTOMER_MASTER.Cust_Group_Code as Customer_Group,TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Desc   , TSPL_SD_SALE_INVOICE_HEAD.Total_Amt as DocAmt ,Against_Shipment_No   from TSPL_SD_SALE_INVOICE_HEAD left outer join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER.Location_Code   left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_SD_SALE_INVOICE_HEAD.comp_code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code   left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code =TSPL_CUSTOMER_MASTER.City_Code  left outer join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE =TSPL_COMPANY_MASTER.State  left outer join TSPL_CITY_MASTER  as TSPL_CITY_MASTER_For_Comp on TSPL_CITY_MASTER_For_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code  left join TSPL_CUSTOMER_CATEGORY_MASTER on TSPL_CUSTOMER_CATEGORY_MASTER.cust_category_code =TSPL_CUSTOMER_MASTER.cust_category_code left join TSPL_CUSTOMER_GROUP_MASTER on TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code =TSPL_CUSTOMER_MASTER.Cust_Group_Code left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SHIPMENT_HEAD.Sale_Invoice_No  where  TSPL_SD_SALE_INVOICE_HEAD.Trans_Type ='FS' "

        sQuery += " and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103)"
        'If chkLocationSelect.IsChecked And cbgLocation.CheckedValue.Count > 0 Then
        '    sQuery += " and TSPL_LOCATION_MASTER. Location_Code   IN (" + clsCommon.GetMulcallString(cbgLocation.CheckedValue) + ") "
        'End If
        If txtLocation1Mult.arrValueMember IsNot Nothing AndAlso txtLocation1Mult.arrValueMember.Count > 0 Then
            sQuery += " and TSPL_LOCATION_MASTER. Location_Code   IN (" + clsCommon.GetMulcallString(txtLocation1Mult.arrValueMember) + ") "
        End If
        If txtCustomerMult.arrValueMember IsNot Nothing AndAlso txtCustomerMult.arrValueMember.Count > 0 Then
            sQuery += "and TSPL_SD_SALE_INVOICE_HEAD.customer_code in (" + clsCommon.GetMulcallString(txtCustomerMult.arrValueMember) + ") " + Environment.NewLine
        End If
        'If txtCustomerCategoryMult.arrValueMember IsNot Nothing AndAlso txtCustomerCategoryMult.arrValueMember.Count > 0 Then
        '    sQuery += " AND TSPL_CUSTOMER_MASTER.cust_category_code in (" + clsCommon.GetMulcallString(txtCustomerCategoryMult.arrValueMember) + ") " + Environment.NewLine
        'End If
        If TxtMultiCustomerGroup.arrValueMember IsNot Nothing AndAlso TxtMultiCustomerGroup.arrValueMember.Count > 0 Then
            sQuery += " AND TSPL_CUSTOMER_MASTER.Cust_Group_Code in (" + clsCommon.GetMulcallString(TxtMultiCustomerGroup.arrValueMember) + ") " + Environment.NewLine
        End If
        sQuery += " order by convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) "
        Dim dtgv As New DataTable
        dtgv = clsDBFuncationality.GetDataTable(sQuery)
        If dtgv IsNot Nothing And dtgv.Rows.Count > 0 Then
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

        gv.Columns("vehicleNo").IsVisible = True
        gv.Columns("vehicleNo").Width = 100
        gv.Columns("vehicleNo").HeaderText = "Vehicle No "
        gv.Columns("vehicleNo").ReadOnly = False

        gv.Columns("Comp_Name").IsVisible = False
        gv.Columns("Comp_Name").Width = 100
        gv.Columns("Comp_Name").HeaderText = "Comp Name"

        gv.Columns("Comp_address").IsVisible = False
        gv.Columns("Comp_address").Width = 100
        gv.Columns("Comp_address").HeaderText = "Comp address"


        gv.Columns("InvoiceNo").IsVisible = True
        gv.Columns("InvoiceNo").Width = 100
        gv.Columns("InvoiceNo").HeaderText = "Sale Invoice No."



        gv.Columns("InvoiceDate").IsVisible = True
        gv.Columns("InvoiceDate").Width = 100
        gv.Columns("InvoiceDate").HeaderText = " Date"
        gv.Columns("InvoiceDate").FormatString = "{0:d}"

        gv.Columns("Bill_To_Location").IsVisible = True
        gv.Columns("Bill_To_Location").Width = 100
        gv.Columns("Bill_To_Location").HeaderText = "Location Code"


        gv.Columns("Location_Desc").IsVisible = True
        gv.Columns("Location_Desc").Width = 100
        gv.Columns("Location_Desc").HeaderText = "Location Name"

        gv.Columns("Customer_Code").IsVisible = True
        gv.Columns("Customer_Code").Width = 150
        gv.Columns("Customer_Code").HeaderText = "Customer Code"


        gv.Columns("CustName").IsVisible = True
        gv.Columns("CustName").Width = 150
        gv.Columns("CustName").HeaderText = "Customer Name"

        gv.Columns("Customer_Add1").IsVisible = True
        gv.Columns("Customer_Add1").Width = 100
        gv.Columns("Customer_Add1").HeaderText = " Customer Address"

        gv.Columns("DocAmt").IsVisible = True
        gv.Columns("DocAmt").Width = 100
        gv.Columns("DocAmt").HeaderText = "Amount"

        gv.Columns("Loc_Add1").IsVisible = False
        gv.Columns("Loc_Add1").Width = 100
        gv.Columns("Loc_Add1").HeaderText = "Location Address1"
        gv.Columns("Loc_Add2").IsVisible = False
        gv.Columns("Loc_Add2").Width = 100
        gv.Columns("Loc_Add2").HeaderText = "Location Address2" '

        gv.Columns("Loc_Add3").IsVisible = False
        gv.Columns("Loc_Add3").Width = 100
        gv.Columns("Loc_Add3").HeaderText = "Location Address3"

        gv.Columns("Logo_Img").IsVisible = False
        gv.Columns("Logo_Img").Width = 100
        gv.Columns("Logo_Img").HeaderText = "Logo Img1"
        gv.Columns("Logo_Img2").IsVisible = False
        gv.Columns("Logo_Img2").Width = 100
        gv.Columns("Logo_Img2").HeaderText = "Logo Img2"

        gv.Columns("Compaddress2").IsVisible = False
        gv.Columns("Compaddress2").Width = 100
        gv.Columns("Compaddress2").HeaderText = "Company address2"
        gv.Columns("Compaddress3").IsVisible = False
        gv.Columns("Compaddress3").Width = 100
        gv.Columns("Compaddress3").HeaderText = "Company address3"

        gv.Columns("cust_category_desc").IsVisible = False
        gv.Columns("cust_category_desc").Width = 150
        gv.Columns("cust_category_desc").HeaderText = "Customer Category"


        gv.Columns("customer_Add2").IsVisible = False
        gv.Columns("customer_Add2").Width = 100
        gv.Columns("customer_Add2").HeaderText = "Customer address2"


        gv.Columns("customer_Add3").IsVisible = False
        gv.Columns("customer_Add3").Width = 100
        gv.Columns("customer_Add3").HeaderText = "Customer address3"


        gv.Columns("Customer_Group").IsVisible = True
        gv.Columns("Customer_Group").Width = 150
        gv.Columns("Customer_Group").HeaderText = "Customer Group Code"

        gv.Columns("Cust_Group_Desc").IsVisible = True
        gv.Columns("Cust_Group_Desc").Width = 150
        gv.Columns("Cust_Group_Desc").HeaderText = "Customer Group Name"

        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim intCount As Integer = 0

        Dim item1 As New GridViewSummaryItem("DocAmt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item1)



        gv.ShowGroupPanel = False
        gv.MasterTemplate.AutoExpandGroups = True

        gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Private Sub BtnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint.Click
        loadData()

    End Sub



    Sub loadData()
        If clsCommon.myCDate(txtFromDate.Value) < objCommonVar.GSTApplicableDate AndAlso clsCommon.myCDate(txtToDate.Value) > objCommonVar.GSTApplicableDate Then
            clsCommon.MyMessageBoxShow(Me, "Please Select From Date and To date range without GST or within GST", Me.Text)
            Exit Sub
        End If

        ArrInvoice_Arr = New ArrayList

        Dim InvoiceNo As String = ""
        For Each grow As GridViewRowInfo In gv.Rows
            If clsCommon.CompairString(clsCommon.myCBool(grow.Cells(0).Value), True) = CompairStringResult.Equal Then
                If CreateFreshInvoiceOnDispatchSave = 0 Then
                    InvoiceNo = InvoiceNo + "','" + clsCommon.myCstr(grow.Cells("Against_Shipment_No").Value)
                Else
                    InvoiceNo = InvoiceNo + "','" + clsCommon.myCstr(grow.Cells("InvoiceNo").Value)
                End If
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
        If chkLocationSelect.IsChecked AndAlso cbgLocation.CheckedValue.Count = 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please select atleast single Location or select all.", Me.Text)
            Exit Sub
        End If

        Dim Qry As String = LoadPrintQuery(InvoiceNo)

        'Qry += "   ('" + InvoiceNo + "')  and TSPL_SD_SHIPMENT_DETAIL.Scheme_Item='N'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)

        If dt.Rows.Count > 0 Then

            'KwalitySalesReportViewer.funreport(dt, "crptFreshInvoicePrintStatement", "Fresh Invoice Statement")
            ''KwalitySalesReportViewer.funreport(dt, "ggg", "Fresh Invoice Statement")
            'KwalitySalesReportViewer.funsubreportWithdt(dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptFreshSaleInvoice(New)", "Fresh Invoice Statement", "rptCompanyAddress.rpt")
            Dim frmCRV As New frmCrystalReportViewer()
            frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "crptFreshSaleInvoice(New)", "Fresh Invoice Statement", txtFromDate.Value, "rptCompanyAddress.rpt", "FreshHeader.rpt", clsERPFuncationality.CompanyAddresInvoiceHeader())
            frmCRV = Nothing
        End If



    End Sub
    '==============Changed By Rohit on Apr 03,2014==========================
    ''changes by Shivani [BM00000007756]
    '=================preeti gupta Ticket no[BM00000008544,ERO/20/06/18-000358,ERO/29/06/18-000362,ERO/29/06/18-000379]
    Public Function LoadPrintQuery(ByVal strinvoiceNo) As String
        Dim Qry As String = Nothing
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDP") = CompairStringResult.Equal Then
            Qry = GetQuerySwadesh(strinvoiceNo)
            Return Qry
        End If
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "GMD") = CompairStringResult.Equal Then
            Qry = GetQueryGMD(strinvoiceNo)
            Return Qry
        End If
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "Alpha") = CompairStringResult.Equal Then
            Qry = GetQueryAlpha(strinvoiceNo)
            Return Qry
        End If
        CreateFreshInvoiceOnDispatchSave = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateFreshInvoiceOnDispatchSave, clsFixedParameterCode.CreateFreshInvoiceOnDispatchSave, Nothing))
        Dim ShowSchemeItemRate As Integer = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowSchemeItemRateonDairyDispatch, clsFixedParameterCode.ShowSchemeItemRateonDairyDispatch, Nothing))
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then
            If clsCommon.CompairString(CreateFreshInvoiceOnDispatchSave, "1") = CompairStringResult.Equal Then
                CreateFreshInvoiceOnDispatchSave = 0
            End If
        End If
        If CreateFreshInvoiceOnDispatchSave = 0 Then
            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "SPMMD") = CompairStringResult.Equal Then  'For Batch wise,Rate per Ltr/Kg in printout
                Dim Bank_Name As String = ""
                Dim IFSC_Code As String = ""
                Dim BANKACCNUMBER As String = ""

                Qry = "select isnull(DESCRIPTION,'') as Bank_Name,isnull(IFSC_Code,'') as IFSC_Code,isnull(BANKACCNUMBER,'') as BANKACCNUMBER from TSPL_BANK_MASTER" &
                " left outer join TSPL_BANK_BRANCH_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_BANK_BRANCH_MASTER.Bank_CODE " &
                " where Default_Bank = 1 "

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)

                If dt.Rows.Count > 0 Then
                    Bank_Name = clsCommon.myCstr(dt.Rows(0).Item("Bank_Name"))
                    IFSC_Code = clsCommon.myCstr(dt.Rows(0).Item("IFSC_Code"))
                    BANKACCNUMBER = clsCommon.myCstr(dt.Rows(0).Item("BANKACCNUMBER"))
                End If
                Qry = " select xx.*,TSPL_COMPANY_MASTER.Bank_Name as Comp_Bank_Name , TSPL_COMPANY_MASTER.BankAccountNo as Comp_BankAccountNo , TSPL_COMPANY_MASTER.BankIFSCCode as Comp_BankIFSCCode , TSPL_COMPANY_MASTER.BankBranchAddress as Comp_BankBranchAddress, 'KARUR VYSYA BANK LTD' as Bank_Name,'' as IFSC_Code,'1129135000009595' as BANKACCNUMBER, tspl_company_master.CINNo as Comp_CINNo,tspl_company_master.Pan_No  as Com_PAN_No, '' as Party,'' as Partycust_add1 , '' as Partycust_add2, '' as Partycust_add3 , '' as Party_Gst_StateCode, '' as PartyGSTNo , '' as Party_Pan , '' as Party_Phone ,TSPL_COMPANY_MASTER.Logo_Img from ( select '1' as CopyType,max(TCS_Rate) as TCS_Rate,max(TCS_Amount) as TCS_Amount,max(gst_state_code) as gst_state_code,max(InvRemarks) as InvRemarks,max(Date_Time_Invoice) as Date_Time_Invoice,sum(SchemeValue) as SchemeValue,max(shipName) as shipName,max(ship_Add1) as ship_Add1,max(ship_add2) as ship_add2,max(ship_add3) as ship_add3,max(Ship_Pin_Code) as Ship_Pin_Code,max(Ship_GST_State_Code) as Ship_GST_State_Code,max(Ship_GSTNo) as Ship_GSTNo,max(Ship_PANNo) as Ship_PANNo,max(Ship_FSSAI_No) as Ship_FSSAI_No,max(SHip_Phn) as SHip_Phn ,max(terms_code) as terms_code,max(Payment_code) as Payment_code,max(Payment_desc) as Payment_desc,max(Pin_No) as Pin_No,sum(Amt_Less_discount) as Amt_Less_discount,max(Total_Amt) as Total_Amt,max(Description) as Description,max(Cust_city) as Cust_city,max(Against_Shipment_No) as Against_Shipment_No,max(Cust_Gst_StateCode) as Cust_Gst_StateCode,max(Electronic_Ref_No) as Electronic_Ref_No,max(CustGSTNo) as CustGSTNo,max(LocGstNo) as LocGstNo,max(EWayBillNo) as EWayBillNo,max(EWayBillDate) as EWayBillDate,max(HSN_Code) as HSN_Code,max(Delivery_Code) as Delivery_Code,max(Conversion_factor) as Conversion_factor,sum(QTY_Box) as QTY_Box ,sum(QTY_CAN) as QTY_CAN,sum(FreeQtyInLtr) as FreeQtyInLtr ,max(kgConFactor) as kgConFactor,max(LtrConvFactor) as LtrConvFactor,sum(FreeQtyInkg) as FreeQtyInkg,Sale_Invoice_No,max(vehicleNo) as vehicleNo,max(Sale_Invoice_Date) as Sale_Invoice_Date,max(RoundOffAmount) as RoundOffAmount,max(Loc_ADd1) as Loc_ADd1,max(LOC_ADD2) as LOC_ADD2,max(LOC_ADD3) as LOC_ADD3,max(LocationState) as LocationState,max(LOCPhone) as LOCPhone,max(Loc_TIN_NO) as Loc_TIN_NO,max(Document_Code) as Document_Code,max(Document_Date) as Document_Date,max(Lorry_No) as Lorry_No,max(Sku_Seq) as Sku_Seq,Item_Code ,max(Line_No) as Line_No,max(Particulars) as Particulars,sum(QtyCrates) as QtyCrates,max(Unit_code) as Unit_code,sum(QtyPCS) as QtyPCS ,sum(free_qty) as free_qty,max(FreeSchemeInLitres) as FreeSchemeInLitres,max(RatePerPcs) as RatePerPcs,max(RatePerLtrKg) as RatePerLtrKg ,sum(valueInRs) as valueInRs, 
                         upper(SUBSTRING(Batch_No,1,7)) as Batch_No
                         ,sum(Cash_Scheme_Amount) as Cash_Scheme_Amount ,sum(schemeInCrates) as schemeInCrates ,max(GrandTotalCrates) as GrandTotalCrates,max(FSSAI) as FSSAI , max(Comp_Code) as Comp_Code,max(Comp_Name) as Comp_Name,max(comp_add1) as comp_add1,max(comp_add2) as comp_add2,max(comp_add3) as comp_add3 ,max(comp_Fax) as comp_Fax,max(comp_Email) as comp_Email,max(CompPhone) as CompPhone,max(comp_tinNo) as comp_tinNo,max(cust_Code) as cust_Code ,max(Customer_Name) as Customer_Name,max(cust_add1) as cust_add1,max(cust_add2) as cust_add2,max(cust_add3) as cust_add3,max(Cust_Pin) as Cust_Pin,max(Sales_Name) as Sales_Name,max(Cust_FSSAI_LIC_NO) as Cust_FSSAI_LIC_NO,max(CustPhone) as CustPhone,max(cust_fax) as cust_fax,max(cust_Email) as cust_Email,max(cust_website) as cust_website,max(Customer_Pan) as Customer_Pan from  
                         (select  Case when dtax1.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX1_Rate when  dtax2.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX2_Rate when dtax3.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX3_Rate  when dtax4.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX4_Rate  when dtax5.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX5_Rate end as TCS_Rate   ,Case when dtax1.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX1_Amt when  dtax2.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX2_Amt when dtax3.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX3_Amt when dtax4.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX4_Amt  when dtax5.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX5_Amt end  as TCS_Amount, TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date_Time_Invoice,
                         isnull(convert(Decimal(18,2), ( (case when (ROW_NUMBER() OVER(PARTITION BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code
                         ,BI.Batch_No ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                          ))=1 then Sub_qty else 0 end)/StockUnit.Conversion_Factor)*Sub_Stocking_Con*Suborgconversion_factor),0)* isnull(CONVERT(DECIMAL(18,2),(TSPL_SD_SHIPMENT_DETAIL.item_cost/TSPL_ITEM_UOM_DETAIL.Conversion_Factor)* StockUnit.Conversion_Factor),0) as SchemeValue
                         ,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Customer_Name else TSPL_SHIP_TO_LOCATION.Ship_To_Desc end  as shipName,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Add1 else TSPL_SHIP_TO_LOCATION.add1 end as ship_Add1,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Add2 else  TSPL_SHIP_TO_LOCATION.Add2 end as ship_add2 ,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Add3 else  TSPL_SHIP_TO_LOCATION.Add3 end  as ship_add3  , case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then tspl_customer_master.pin_no else TSPL_SHIP_TO_LOCATION.Pin_Code end as Ship_Pin_Code,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then CUSTOMER_STATE_MASTER.GST_STATE_Code else Ship_State.gst_state_code end as Ship_GST_State_Code,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then Tspl_customer_master.gstno else TSPL_SHIP_TO_LOCATION.gstNo  end as Ship_GSTNo,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.pan else  TSPL_SHIP_TO_LOCATION.PAN end as Ship_PANNo,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then isnull(TSPL_CUSTOMER_MASTER.FSSAI_NO,'') else '' end as Ship_FSSAI_No,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End else TSPL_SHIP_TO_LOCATION.telphone end as SHip_Phn, tspl_customer_master.terms_code,TSPL_TERMS_MASTER.terms_desc,tspl_customer_master.payment_code,tspl_payment_code.payment_desc,tspl_customer_master.pin_no
                         , (case when BI.Qty is not null then convert(Decimal(18,2),(TSPL_SD_SHIPMENT_detail.amt_less_discount/TSPL_SD_SHIPMENT_detail.Qty)*BI.Qty)
                         else TSPL_SD_SHIPMENT_detail.amt_less_discount end) as amt_less_discount
                         ,TSPL_SD_SHIPMENT_HEAD.Total_Amt, TSPL_SD_SALE_INVOICE_HEAD.Description,customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SHIPMENT_HEAD.Electronic_Ref_No,Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SHIPMENT_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor
                         ,isnull(case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),
                         coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty) *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end ,0) as QTY_Box
                         , isnull(case when coalesce(ITEMDETAILCAN.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),
                        coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty)*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAILCAN.Conversion_factor,1)) end ,0) as QTY_CAN 
                         ,isnull(case when coalesce(ITEMDETAILLTR.Conversion_factor,0)=0 then 0 else convert (Decimal(18,2),((case when (ROW_NUMBER() OVER(PARTITION BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code ,BI.Batch_No
                         ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                         ))=1 then Sub_qty else 0 end)/ITEMDETAILLTR.Conversion_factor)*Sub_Stocking_Con*Suborgconversion_factor) end ,0) as FreeQtyInLtr
                         ,ITEMDETAILKG.Conversion_factor AS kgConFactor,ITEMDETAILLTR.Conversion_factor as LtrConvFactor, isnull(case when coalesce(ITEMDETAILKG.Conversion_factor,0)=0 then 0 else convert (Decimal(18,2),((case when (ROW_NUMBER() OVER(PARTITION BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                          ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                          ))=1 then Sub_qty else 0 end)/ITEMDETAILKG.Conversion_factor)*Sub_Stocking_Con*Suborgconversion_factor) end ,0) as FreeQtyInkg
                         ,TSPL_SD_SALE_INVOICE_HEAD.document_code as Sale_Invoice_No, Case When ISNULL(TSPL_SD_SHIPMENT_Head.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_Head.ManualVehicle WHEN ISNULL(TSPL_SD_SHIPMENT_Head.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SHIPMENT_Head.vehicleNo End as vehicleNo, TSPL_SD_SHIPMENT_Head.Sale_Invoice_Date, TSPL_SD_SHIPMENT_HEAD.RoundOffAmount, TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO,  TSPL_SD_SHIPMENT_HEAD.Document_Code ,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as Document_Date , TSPL_SD_SHIPMENT_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_SHIPMENT_DETAIL.Item_Code ,TSPL_SD_SHIPMENT_DETAIL.Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars
                         ,CASE WHEN TSPL_SD_SHIPMENT_DETAIL.Unit_code IN ('Ltr','Crate') then coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty) else 0 end as QtyCrates 
                         ,TSPL_SD_SHIPMENT_DETAIL.Unit_code 
                         ,isnull(convert(Decimal(18,2), (coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty)
                         /StockUnit.Conversion_Factor)*StockingUnit.Conversion_Factor*TSPL_ITEM_UOM_DETAIL.Conversion_Factor),0)  as QtyPCS
                         , isnull(convert(Decimal(18,2), ((case when (ROW_NUMBER() OVER(PARTITION BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                           ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                           ))=1 then Sub_qty else 0 end)/StockUnit.Conversion_Factor)*Sub_Stocking_Con*Suborgconversion_factor),0) as free_qty, TSPL_SD_SHIPMENT_DETAIL.Scheme_Item as FreeSchemeInLitres , isnull(CONVERT(DECIMAL(18,2),(TSPL_SD_SHIPMENT_DETAIL.item_cost/TSPL_ITEM_UOM_DETAIL.Conversion_Factor)* StockUnit.Conversion_Factor),0) as RatePerPcs , isnull(CONVERT(DECIMAL(18,2),case when isnull(ITEMDETAILLTR.Conversion_Factor,0)>0 then (TSPL_SD_SHIPMENT_DETAIL.item_cost/TSPL_ITEM_UOM_DETAIL.Conversion_Factor )*ITEMDETAILLTR.Conversion_Factor else (TSPL_SD_SHIPMENT_DETAIL.item_cost/TSPL_ITEM_UOM_DETAIL.Conversion_Factor )*ITEMDETAILKG.Conversion_Factor end),0) as RatePerLtrKg 
                          ,(case when TSPL_SD_SHIPMENT_DETAIL.Sampling=1 then 0 else   coalesce(BI.Qty*convert(DECIMAL(18,2),(case when TSPL_SD_SHIPMENT_DETAIL.Sampling=1then 0 else  TSPL_SD_SHIPMENT_DETAIL.Amount end)/ (TSPL_SD_SHIPMENT_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)),TSPL_SD_SHIPMENT_DETAIL.amount)  end) as valueInRs
                          ,BI.Batch_No, coalesce(TSPL_SD_SHIPMENT_DETAIL.Cash_Scheme_Amount,0)+coalesce(TSPL_SD_SHIPMENT_DETAIL.HeadDiscAmt,0) as Cash_Scheme_Amount,case when (ROW_NUMBER() OVER(PARTITION BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                           ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No,TSPL_SD_SHIPMENT_DETAIL.Item_Code,BI.Batch_No
                           ))=1 then isnull(schemeInCrates,0) else 0 end as schemeInCrates
                          ,   '' GrandTotalCrates,tspl_company_master.Access_Officer as FSSAI , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo ,TSPL_SD_SHIPMENT_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,isnull(TSPL_CUSTOMER_MASTER.pin_Code,'') as Cust_Pin,(select top 1 tspl_employee_master.emp_name from tspl_salesman_detail left join tspl_employee_master on tspl_employee_master.emp_code=tspl_salesman_detail.salesman_code where route_code=tspl_route_master.route_no) as  Sales_Name,isnull(TSPL_CUSTOMER_MASTER.FSSAI_NO,'') as Cust_FSSAI_LIC_NO ,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan   from TSPL_SD_SHIPMENT_DETAIL   LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD .Document_Code =TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE  left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON TSPL_ITEM_MASTER.Item_Code =TSPL_SD_SHIPMENT_DETAIL.Item_Code LEFT JOIN TSPL_BATCH_ITEM as BI ON BI.Document_Code=TSPL_SD_SHIPMENT_DETAIL.Document_Code AND BI.Parent_Line_No=TSPL_SD_SHIPMENT_DETAIL.Line_No AND BI.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE  where TSPL_UNIT_MaSTER.BOX_TYPE = 'Y') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL   left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE  where TSPL_UNIT_MaSTER.LTR_TYPE='Y') as ITEMDETAILLTR on ITEMDETAILLTR.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='KG') as ITEMDETAILKG on ITEMDETAILKG.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code   left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL 	  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE  where TSPL_UNIT_MaSTER.CAN_TYPE='Y') as ITEMDETAILCAN on ITEMDETAILCAN.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code        left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_SD_SHIPMENT_DETAIL .tax1   left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_SD_SHIPMENT_DETAIL.tax2   left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_SD_SHIPMENT_DETAIL .TAX3  left outer join TSPL_TAX_MASTER as dtax4 on dtax4.Tax_Code= TSPL_SD_SHIPMENT_DETAIL .tax4   left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code=TSPL_SD_SHIPMENT_DETAIL .tax5   left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SHIPMENT_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SHIPMENT_HEAD .Customer_Code left join tspl_payment_code on tspl_payment_code.payment_code=tspl_customer_master.payment_code  left join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.terms_code=tspl_customer_master.terms_code  left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SHIPMENT_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code  left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE  left outer join TSPL_SD_SALE_INVOICE_HEAD  on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No   
                          Full join  (select inn.DOCUMENT_CODE,inn.Item_Code as Scheme_Item_Code,SUM(coalesce(BI.Qty,INN.Qty) ) AS SUB_QTY,SUM(Crate) AS schemeInCrates,max(TSPL_ITEM_UOM_DETAIL.conversion_factor) as SubOrgconversion_factor ,max(StockingUnit.conversion_factor) as Sub_Stocking_Con,BI.Batch_No
                           from TSPL_SD_SHIPMENT_DETAIL  as inn  
                          LEFT JOIN TSPL_BATCH_ITEM as BI ON BI.Document_Code=inn.Document_Code AND BI.Parent_Line_No=inn.Line_No AND BI.Item_Code=inn.Item_Code 
                          left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=inn.Item_Code 
                         and TSPL_ITEM_UOM_DETAIL .UOM_Code=inn.Unit_code LEFT OUTER JOIN tspl_item_uom_detail as StockingUnit on StockingUnit.Item_Code=inn.Item_Code and StockingUnit.stocking_unit='Y' where inn.DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by inn.DOCUMENT_CODE,inn.Item_Code,BI.Batch_No
                         )  TSPL_SD_SHIPMENT_DETAIL_Sub on TSPL_SD_SHIPMENT_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE and TSPL_SD_SHIPMENT_DETAIL_sub.Scheme_Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  
                         and BI.Batch_No =TSPL_SD_SHIPMENT_DETAIL_Sub.Batch_No
                        LEFT OUTER JOIN TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle LEFT OUTER JOIN tspl_item_uom_detail as StockUnit on StockUnit.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and StockUnit.pieces=1  LEFT OUTER JOIN tspl_item_uom_detail as StockingUnit on StockingUnit.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and StockingUnit.stocking_unit='Y' LEFT OUTER JOIN TSPL_VEHICLE_MASTER as Vehicle on Vehicle.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.Vehicle_Code left join TSPL_DELIVERY_NOTE_MASTER_FRESHSALE on TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.document_no=TSPL_SD_SHIPMENT_HEAD.against_delivery_code left join tspl_route_master on tspl_route_master.route_no= TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.route_no  left outer join  TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code =TSPL_SD_SHIPMENT_HEAD .Ship_To_Location  left outer join TSPL_CITY_MASTER  as Ship_City on Ship_City.City_Code =TSPL_SHIP_TO_LOCATION.City_Code  left join tspl_state_master as Ship_State on Ship_State.state_code=TSPL_SHIP_TO_LOCATION.state where 2=2 and  TSPL_SD_SHIPMENT_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_SHIPMENT_DETAIL.Scheme_Item='N'   
                        UNION aLL 
                        select  Case when dtax1.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX1_Rate when  dtax2.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX2_Rate when dtax3.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX3_Rate  when dtax4.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX4_Rate  when dtax5.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX5_Rate end as TCS_Rate , Case when dtax1.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX1_Amt when  dtax2.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX2_Amt when dtax3.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX3_Amt when dtax4.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX4_Amt when dtax5.Is_TCS = 'Y' then TSPL_SD_SHIPMENT_HEAD.TAX5_Amt end  as TCS_Amount,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_SALE_INVOICE_HEAD.Document_Date as Date_Time_Invoice,TSPL_SD_SHIPMENT_DETAIL.item_Net_Amt as SchemeValue, case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Customer_Name else TSPL_SHIP_TO_LOCATION.Ship_To_Desc end  as shipName,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Add1 else TSPL_SHIP_TO_LOCATION.add1 end as ship_Add1,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Add2 else  TSPL_SHIP_TO_LOCATION.Add2 end as ship_add2 ,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.Add3 else  TSPL_SHIP_TO_LOCATION.Add3 end  as ship_add3  , case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then tspl_customer_master.pin_no else TSPL_SHIP_TO_LOCATION.Pin_Code end as Ship_Pin_Code,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then CUSTOMER_STATE_MASTER.GST_STATE_Code else Ship_State.gst_state_code end as Ship_GST_State_Code,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then Tspl_customer_master.gstno else TSPL_SHIP_TO_LOCATION.gstNo  end as Ship_GSTNo,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then TSPL_CUSTOMER_MASTER.pan else  TSPL_SHIP_TO_LOCATION.PAN end as Ship_PANNo,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then isnull(TSPL_CUSTOMER_MASTER.FSSAI_NO,'') else '' end as Ship_FSSAI_No,case when isnull(TSPL_SD_SHIPMENT_HEAD .Ship_To_Location,'')='' then case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End else TSPL_SHIP_TO_LOCATION.telphone end as SHip_Phn,tspl_customer_master.terms_code,TSPL_TERMS_MASTER.terms_desc,tspl_customer_master.payment_code,tspl_payment_code.payment_desc,tspl_customer_master.pin_no,0 as amt_less_discount, 0 as Total_Amt,TSPL_SD_SALE_INVOICE_HEAD.Description, customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode,TSPL_SD_SHIPMENT_HEAD.Electronic_Ref_No,Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SHIPMENT_DETAIL.Delivery_Code , ITEMDETAIL.Conversion_factor,case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box, case when coalesce(ITEMDETAILCAN.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty) 
                         *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAILCAN.Conversion_factor,1)) end as QTY_CAN
                         ,case when coalesce(ITEMDETAILLTR.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  coalesce( case when TSPL_SD_SHIPMENT_DETAIL.Sampling=1 then 0 else 
                        coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty) end,0)*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAILLTR.Conversion_factor,1)) end as FreeQtyInLtr
                          ,ITEMDETAILKG.Conversion_factor AS kgConFactor,ITEMDETAILLTR.Conversion_factor as LtrConvFactor,  case when coalesce(ITEMDETAILKG.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),  coalesce( case when TSPL_SD_SHIPMENT_DETAIL.Sampling=1 then 0 else  
                          coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty)
                          end,0)*TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAILKG.Conversion_factor,1)) end as FreeQtyInkg,TSPL_SD_SALE_INVOICE_HEAD.document_code as Sale_Invoice_No, Case When ISNULL(TSPL_SD_SHIPMENT_Head.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_Head.ManualVehicle WHEN ISNULL(TSPL_SD_SHIPMENT_Head.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SHIPMENT_Head.vehicleNo End as vehicleNo, TSPL_SD_SHIPMENT_Head.Sale_Invoice_Date as Sale_Invoice_Date, TSPL_SD_SHIPMENT_HEAD.RoundOffAmount, TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO,  TSPL_SD_SHIPMENT_HEAD.Document_Code ,convert(varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) as Document_Date , TSPL_SD_SHIPMENT_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars, 0 as QtyCrates ,TSPL_SD_SHIPMENT_DETAIL.Unit_code ,0 as QtyPCS  , coalesce( case when TSPL_SD_SHIPMENT_DETAIL.Sampling=1 then 0 else 
                         coalesce(BI.Qty,TSPL_SD_SHIPMENT_DETAIL.Qty)
                           end,0) as free_qty,  TSPL_SD_SHIPMENT_DETAIL.Scheme_Item as FreeSchemeInLitres ,0 as RatePerPcs,0 as RatePerLtrKg,0  as valueInRs,
                           BI.Batch_No
                            ,0 as Cash_Scheme_Amount,isnull(TSPL_SD_SHIPMENT_DETAIL.Crate,0)as schemeInCrates,   '' GrandTotalCrates ,tspl_company_master.Access_Officer as FSSAI, TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo ,TSPL_SD_SHIPMENT_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,isnull(TSPL_CUSTOMER_MASTER.pin_Code,'') as Cust_Pin,(select top 1 tspl_employee_master.emp_name from tspl_salesman_detail left join tspl_employee_master on tspl_employee_master.emp_code=tspl_salesman_detail.salesman_code where route_code=tspl_route_master.route_no) as Sales_Name,isnull(TSPL_CUSTOMER_MASTER.FSSAI_NO,'') as Cust_FSSAI_LIC_NO,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan   from TSPL_SD_SHIPMENT_DETAIL   LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD .Document_Code =TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE  left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON TSPL_ITEM_MASTER.Item_Code =TSPL_SD_SHIPMENT_DETAIL.Item_Code LEFT JOIN TSPL_BATCH_ITEM as BI ON BI.Document_Code=TSPL_SD_SHIPMENT_DETAIL.Document_Code AND BI.Parent_Line_No=TSPL_SD_SHIPMENT_DETAIL.Line_No AND BI.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL   left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.BOX_TYPE = 'Y' ) as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.LTR_TYPE='Y') as ITEMDETAILLTR on ITEMDETAILLTR.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code  left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='KG') as ITEMDETAILKG on ITEMDETAILKG.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code   left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL 	  left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.UNIT_CODE= TSPL_ITEM_UOM_DETAIL.UOM_CODE where TSPL_UNIT_MaSTER.CAN_TYPE='Y') as ITEMDETAILCAN on ITEMDETAILCAN.Item_code=TSPL_SD_SHIPMENT_DETAIL.Item_Code      left outer join TSPL_TAX_MASTER as dtax1 on dtax1.tax_code =TSPL_SD_SHIPMENT_HEAD .tax1   left outer join tspl_tax_master as dtax2 on dtax2.tax_code = TSPL_SD_SHIPMENT_HEAD.tax2   left outer join tspl_tax_master as dtax3 on dtax3.Tax_Code=TSPL_SD_SHIPMENT_HEAD .TAX3  left outer join TSPL_TAX_MASTER as dtax4 on dtax4.Tax_Code= TSPL_SD_SHIPMENT_HEAD .tax4  left outer join TSPL_TAX_MASTER as dtax5 on dtax5.Tax_Code= TSPL_SD_SHIPMENT_HEAD .tax5  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SHIPMENT_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SHIPMENT_HEAD .Customer_Code left join tspl_payment_code on tspl_payment_code.payment_code=tspl_customer_master.payment_code  left join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.terms_code=tspl_customer_master.terms_code  left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SHIPMENT_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE  left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code  left outer join TSPL_SD_SALE_INVOICE_HEAD  on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No   LEFT OUTER JOIN TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle LEFT OUTER JOIN TSPL_VEHICLE_MASTER as Vehicle on Vehicle.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.Vehicle_Code LEFT OUTER JOIN tspl_item_uom_detail as StockUnit on StockUnit.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and StockUnit.pieces=1 left join TSPL_DELIVERY_NOTE_MASTER_FRESHSALE on TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.document_no=TSPL_SD_SHIPMENT_HEAD.against_delivery_code left join tspl_route_master on tspl_route_master.route_no= TSPL_DELIVERY_NOTE_MASTER_FRESHSALE.route_no  left outer join  TSPL_SHIP_TO_LOCATION on TSPL_SHIP_TO_LOCATION.Ship_To_Code =TSPL_SD_SHIPMENT_HEAD .Ship_To_Location  left outer join TSPL_CITY_MASTER  as Ship_City on Ship_City.City_Code =TSPL_SHIP_TO_LOCATION.City_Code  left join tspl_state_master as Ship_State on Ship_State.state_code=TSPL_SHIP_TO_LOCATION.state where 2=2 and  TSPL_SD_SHIPMENT_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_SHIPMENT_DETAIL.Scheme_Item='y'    and not exists (select 1 from  
                           (select  TSPL_SD_SHIPMENT_DETAIL.Item_Code,TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE,TSPL_SD_SHIPMENT_DETAIL.Line_No,BI.Batch_no    from TSPL_SD_SHIPMENT_DETAIL   LEFT OUTER JOIN TSPL_SD_SHIPMENT_HEAD ON TSPL_SD_SHIPMENT_HEAD .Document_Code =TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
                           LEFT JOIN TSPL_BATCH_ITEM as BI ON BI.Document_Code=TSPL_SD_SHIPMENT_DETAIL.Document_Code AND BI.Parent_Line_No=TSPL_SD_SHIPMENT_DETAIL.Line_No AND BI.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code
                            left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_SHIPMENT_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON TSPL_ITEM_MASTER.Item_Code =TSPL_SD_SHIPMENT_DETAIL.Item_Code left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SHIPMENT_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SHIPMENT_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SHIPMENT_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE  left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code  Full join  
                           (select inn.DOCUMENT_CODE,inn.Item_Code as Scheme_Item_Code,SUM(coalesce(BI.Qty,INN.Qty)) AS SUB_QTY,SUM(Crate) AS schemeInCrates,Delivery_Code
                           ,BI.Batch_No
                            from TSPL_SD_SHIPMENT_DETAIL as inn
                          LEFT JOIN TSPL_BATCH_ITEM as BI ON BI.Document_Code=inn.Document_Code AND BI.Parent_Line_No=inn.Line_No AND BI.Item_Code=inn.Item_Code 
                           where inn.DOCUMENT_CODE in ('" + strinvoiceNo + "') and Scheme_Item='Y'   group by inn.DOCUMENT_CODE,inn.Item_Code,inn.Delivery_Code
                        ,BI.Batch_No
                        )  TSPL_SD_SHIPMENT_DETAIL_Sub on TSPL_SD_SHIPMENT_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.DOCUMENT_CODE and TSPL_SD_SHIPMENT_DETAIL_sub.Scheme_Item_Code=TSPL_SD_SHIPMENT_DETAIL.Item_Code and  TSPL_SD_SHIPMENT_DETAIL_sub.Delivery_Code=TSPL_SD_SHIPMENT_DETAIL .Delivery_Code
                         and BI.Batch_No =TSPL_SD_SHIPMENT_DETAIL_Sub.Batch_No
                         where 2=2 and  TSPL_SD_SHIPMENT_HEAD.Document_Code in   ('" + strinvoiceNo + "')  and TSPL_SD_SHIPMENT_DETAIL.Scheme_Item='N')xx where xx.Item_Code=TSPL_SD_SHIPMENT_DETAIL.item_CODE 
                         and xx.DOCUMENT_CODE=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE 
                         and xx.Batch_no=BI.Batch_no ))
                         as final   group by Sale_Invoice_No ,final.item_code,final.Batch_No
                         ) as xx left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=xx.comp_code "

            End If
        Else
            Dim ShowShipToPartyInDairyDispatch As Integer = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.ShowShipToPartyInDairyDispatch & "'")) = 0, 0, 1)
            LeakageDeduction_Freshsale = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.leakagededuction_freshsale & "'"))
            If ShowShipToPartyInDairyDispatch = 1 Then
                Dim Bank_Name As String = ""
                Dim IFSC_Code As String = ""
                Dim BANKACCNUMBER As String = ""
                Qry = "select isnull(DESCRIPTION,'') as Bank_Name,isnull(IFSC_Code,'') as IFSC_Code,isnull(BANKACCNUMBER,'') as BANKACCNUMBER from TSPL_BANK_MASTER" &
                " left outer join TSPL_BANK_BRANCH_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_BANK_BRANCH_MASTER.Bank_CODE " &
                " where Default_Bank = 1 "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt.Rows.Count > 0 Then
                    Bank_Name = clsCommon.myCstr(dt.Rows(0).Item("Bank_Name"))
                    IFSC_Code = clsCommon.myCstr(dt.Rows(0).Item("IFSC_Code"))
                    BANKACCNUMBER = clsCommon.myCstr(dt.Rows(0).Item("BANKACCNUMBER"))
                End If
                Qry = " select '" + Bank_Name + "' as Bank_Name,'" + IFSC_Code + "' as IFSC_Code,'" + BANKACCNUMBER + "' as BANKACCNUMBER" + ",case when Line_No=1 then  TSPL_COMPANY_MASTER.Logo_Img else null end as Logo_Img,tt.* from (select DISTINCT * from ( " &
                    "select TSPL_TERMS_MASTER.Terms_Desc,customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor, " &
                    "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2), " &
                    "TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end as QTY_Box, " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                    "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                    "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                    "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                    "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, " &
                    "TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date , " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No,  " &
                    " (case when TSPL_SD_sale_invoice_DETAIL.Row_Type='Item' then TSPL_ITEM_MASTER.Item_Desc else tspl_Additional_Charges.Description end) as Particulars ," &
                    " TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code , " &
                    "convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor) as QtyPCS , " &
                    "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , " &
                    "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,2),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, " &
                    "(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs, " &
                    "coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount " &
                    ",isnull(schemeInCrates,0)as schemeInCrates, " &
                    "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , " &
                    "TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                    "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
                     ",coalesce(TCM.Customer_Name,TSPL_CUSTOMER_MASTER.Customer_Name) Party " &
                     ",coalesce(TCM.Add1,TSPL_CUSTOMER_MASTER.Add1) as Partycust_add1 " &
                     ",coalesce(TCM.Add2,TSPL_CUSTOMER_MASTER. Add2) as Partycust_add2 " &
                     ",coalesce(TCM.Add3,TSPL_CUSTOMER_MASTER.Add3 ) Partycust_add3 " &
                     ",coalesce(PARTY_STATE_MASTER.GST_STATE_Code,CUSTOMER_STATE_MASTER.GST_STATE_Code) AS Party_Gst_StateCode " &
                     ",coalesce(TCM.gstno,Tspl_customer_master.gstno) as PartyGSTNo " &
                     ",coalesce(TCM.pan,TSPL_CUSTOMER_MASTER.pan) as Party_Pan " &
                     ",coalesce(TCM.phone1,TSPL_CUSTOMER_MASTER.phone1) as Party_Phone " &
                     " from TSPL_SD_sale_invoice_DETAIL  " &
                    "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                    "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                    "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and  " &
                    "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                    "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
                    " left outer join TSPL_CUSTOMER_MASTER TCM on tcm.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Ship_To_Party_Parent left outer join TSPL_STATE_MASTER as PARTY_STATE_MASTER on TCM.State= PARTY_STATE_MASTER.STATE_CODE " &
                    " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                    " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                    " left outer join TSPL_TERMS_MASTER on TSPL_CUSTOMER_MASTER.Terms_Code=TSPL_TERMS_MASTER.Terms_Code " &
                    " left outer join tspl_Additional_Charges on tspl_Additional_Charges.Code=TSPL_SD_sale_invoice_DETAIL.Item_Code " &
                    " Full join  (select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                    "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                    "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                    "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                    "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle where 2=2 and " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N' " &
                    "UNION aLL " &
                    " select TSPL_TERMS_MASTER.Terms_Desc,customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code , ITEMDETAIL.Conversion_factor, " &
                    "case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                    "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                    "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                    "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,  TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                    "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date , " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No, " &
                    " (case when TSPL_SD_sale_invoice_DETAIL.Row_Type='Item' then TSPL_ITEM_MASTER.Item_Desc else tspl_Additional_Charges.Description end) as Particulars ," &
                    "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code ,0 as QtyPCS  , " &
                    "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, " &
                    "TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres "

                If ShowSchemeItemRate = 1 Then
                    Qry += ",case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,2),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, " &
                        "(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs, " &
                        "coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount "
                Else
                    Qry += ",0 as RatePerPcs,0  as valueInRs,0 as Cash_Scheme_Amount"
                End If

                Qry += ",isnull(schemeInCrates,0)as schemeInCrates, " &
                    "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                    "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
                   ",coalesce(TCM.Customer_Name,TSPL_CUSTOMER_MASTER.Customer_Name) Party " &
                     ",coalesce(TCM.Add1,TSPL_CUSTOMER_MASTER.Add1) as Partycust_add1 " &
                     ",coalesce(TCM.Add2,TSPL_CUSTOMER_MASTER. Add2) as Partycust_add2 " &
                     ",coalesce(TCM.Add3,TSPL_CUSTOMER_MASTER.Add3 ) Partycust_add3 " &
                     ",coalesce(PARTY_STATE_MASTER.GST_STATE_Code,CUSTOMER_STATE_MASTER.GST_STATE_Code) AS Party_Gst_StateCode " &
                     ",coalesce(TCM.gstno,Tspl_customer_master.gstno) as PartyGSTNo " &
                     ",coalesce(TCM.pan,TSPL_CUSTOMER_MASTER.pan) as Party_Pan " &
                      ",coalesce(TCM.phone1,TSPL_CUSTOMER_MASTER.phone1) as Party_Phone " &
                     " from TSPL_SD_sale_invoice_DETAIL  " &
                    "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                    "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                    "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                    "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                    "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join " &
                    "(select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX' ) as ITEMDETAIL on " &
                    "ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on " &
                    "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                    "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                    "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location " &
                    " left outer join TSPL_CUSTOMER_MASTER TCM on tcm.Cust_Code=TSPL_SD_SHIPMENT_HEAD.Ship_To_Party_Parent left outer join TSPL_STATE_MASTER as PARTY_STATE_MASTER on TCM.State= PARTY_STATE_MASTER.STATE_CODE " &
                     " LEFT OUTER JOIN  TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
                   " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                    " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                    " left outer join TSPL_TERMS_MASTER on TSPL_CUSTOMER_MASTER.Terms_Code=TSPL_TERMS_MASTER.Terms_Code " &
                    " left outer join tspl_Additional_Charges on tspl_Additional_Charges.Code=TSPL_SD_sale_invoice_DETAIL.Item_Code " &
                    " Full join " &
                    "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn " &
                    "where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                    "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                    "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                    "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle where 2=2 and  TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')   " &
                    "and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='y'    and not exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, " &
                    "TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    from TSPL_SD_sale_invoice_DETAIL " &
                    "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                    "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                    "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                    "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left outer join TSPL_COMPANY_MASTER on " &
                    "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                    "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                    "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN  " &
                    "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State Full join " &
                    "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                    "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                    "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                    "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  where 2=2 and " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')  and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N')xx " &
                    "where xx.Item_Code=TSPL_SD_sale_invoice_DETAIL.item_CODE and xx.DOCUMENT_CODE=TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE ) " &
                    ") as final)tt left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=tt.Comp_Code  ORDER BY Line_No"
            Else
                Qry = "  select  Main_Final.*,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.GSTINNo As SellerGST,TSPL_COMPANY_MASTER.Pan_No from ( select DISTINCT final.*,tbl_Brand.Brand,tbl_Brand.BRANDDESC from ( " &
                    "select "
                If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
                    Qry += "TSPL_SD_SALE_INVOICE_HEAD.Amount_Less_Discount*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
                Else
                    Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
                End If
                Qry += " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, " &
                    " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                    " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN,  TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
                  " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
                  " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
                  " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
                  " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
                  " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
                    " customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor, " &
                    "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2), " &
                    "TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end as QTY_Box, " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                    "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                    "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                    "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                    "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, " &
                    "TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No,TSPL_ITEM_MASTER.Item_Desc as Particulars,  " &
                    "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code, convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty ) as Qty_Default,  convert(Decimal(18,2), case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty )) else 0 end ) as Rate_Default, " &
                    "convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor) as QtyPCS , " &
                    "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , " &
                    "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, " &
                    "(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs, " &
                    "coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
                    "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , " &
                    "TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                    "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SHIPMENT_HEAD.DO_Item_Type As TaxableNonTaxable   " &
                    "from TSPL_SD_sale_invoice_DETAIL  " &
                    "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                    "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                    "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and  " &
                    "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                    "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
                     " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                    " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                    " Full join  (select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                    "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                    "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                    "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                    "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
                    " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code " &
                    " and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location " &
                    " and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
                    " where 2=2 and " &
                    "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N' " &
                    "UNION aLL " &
                    " select "
                If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
                    Qry += "TSPL_SD_SALE_INVOICE_HEAD.Total_Amt*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
                Else
                    Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
                End If
                Qry += " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, " &
                        " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                        " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN, TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
                      " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
                      " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
                      " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
                      " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
                      " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
                        "customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code , ITEMDETAIL.Conversion_factor, " &
                        "case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                        "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                        "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                        "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,  TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                        "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                        "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
                        "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_sale_invoice_DETAIL.Item_Code,0 as Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars, " &
                        "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code ,0 as QtyPCS , 0 as Qty_Default , 0 as Rate_Default  , " &
                        "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, " &
                        "TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres ,0 as RatePerPcs,0  as valueInRs,0 as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
                        "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                        "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SHIPMENT_HEAD.DO_Item_Type As TaxableNonTaxable  " &
                        "from TSPL_SD_sale_invoice_DETAIL  " &
                        "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                        "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                        "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                        "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                        "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join " &
                        "(select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX' ) as ITEMDETAIL on " &
                        "ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on " &
                        "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                        "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                        "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN " &
                        "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
                       " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                        " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                        " Full join " &
                        "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn " &
                        "where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                        "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                        "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                        "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
                        " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code " &
                        " and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location " &
                        " and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code   left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
                        " where 2=2 and  TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')   " &
                        "and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='y'    and not exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, " &
                        "TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    from TSPL_SD_sale_invoice_DETAIL " &
                        "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                        "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                        "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                        "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left outer join TSPL_COMPANY_MASTER on " &
                        "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                        "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                        "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN  " &
                        "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State Full join " &
                        "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                        "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                        "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                        "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code " &
                        " where 2=2 and " &
                        "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')  and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N')xx " &
                        "where xx.Item_Code=TSPL_SD_sale_invoice_DETAIL.item_CODE and xx.DOCUMENT_CODE=TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE ) " &
                        ") as final left outer join " &
                        " ( select Item_Code,max([CATEGORY RM]) as [CATEGORY RM],max([BRAND]) as [BRAND],max([SUB BRAND]) as [SUB BRAND],max([DESCRP]) as [DESCRP],max([PACK]) as [PACK]," &
                        " max([PACK SIZE]) as [PACK SIZE],max([CATEGORY OT]) as [CATEGORY OT],max([CATEGORY FA]) as [CATEGORY FA],max([P TYPE]) as [P TYPE],max([L TYPE]) as [L TYPE],max([JW]) as [JW]," &
                        " max([SCRAP]) as [SCRAP],max([CATEGORY RMDESC]) as [CATEGORY RMDESC],max([BRANDDESC]) as [BRANDDESC],max([SUB BRANDDESC]) as [SUB BRANDDESC],max([DESCRPDESC]) as [DESCRPDESC],max([PACKDESC]) as [PACKDESC]," &
                        " max([PACK SIZEDESC]) as [PACK SIZEDESC],max([CATEGORY OTDESC]) as [CATEGORY OTDESC],max([CATEGORY FADESC]) as [CATEGORY FADESC],max([P TYPEDESC]) as [P TYPEDESC],max([L TYPEDESC]) as [L TYPEDESC],max([JWDESC]) as [JWDESC]," &
                        " max([SCRAPDESC]) as [SCRAPDESC]  from ( select * from (   select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code   ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+'DESC' as Item_Category_CodeDesc   ," &
                        " TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  from  TSPL_ITEM_MASTER    left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code " &
                        " left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values   where 2=2 )xx " &
                        " Pivot   ( max(Item_Cagetory_Values) for Item_Category_Code   in ( [CATEGORY RM],[BRAND],[SUB BRAND],[DESCRP],[PACK],[PACK SIZE],[CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE],[JW],[SCRAP])  ) Pivt   Pivot  ( max(Category_Value_Desc) for Item_Category_CodeDesc in ([CATEGORY RMDESC]," &
                        " [BRANDDESC],[SUB BRANDDESC],[DESCRPDESC],[PACKDESC],[PACK SIZEDESC],[CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC],[JWDESC],[SCRAPDESC])  ) Pivt1 ) xxx  group by Item_Code )  as tbl_Brand on tbl_Brand.Item_Code=final.item_Code  ) AS Main_Final left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=Main_Final.comp_code " &
                            " order  by Main_Final.Sku_Seq ,Main_Final.Line_No"
            End If
        End If
        Return Qry
    End Function
    Public Function PrintInvoiceForAll(ByVal strinvoiceNo As String, ByVal docDate As DateTime, ByVal CustCode As String) As String
        Return PrintInvoiceForAll(strinvoiceNo, docDate, CustCode, "Y")
    End Function

    Public Function PrintInvoiceForAll(ByVal strinvoiceNo As String, ByVal docDate As DateTime, ByVal CustCode As String, ByVal isMainPrint As String) As String
        Return PrintInvoiceForAll(strinvoiceNo, docDate, CustCode, "Y", False)
    End Function
    Public Function PrintInvoiceForAll(ByVal strinvoiceNo As String, ByVal docDate As DateTime, ByVal CustCode As String, ByVal isMainPrint As String, ByVal isCancel As String) As String
        Return PrintInvoiceForAll(strinvoiceNo, docDate, CustCode, "Y", isCancel, Nothing)
    End Function
    ' Public Function PrintInvoiceForAll(ByVal strinvoiceNo As String, ByVal docDate As DateTime, ByVal CustCode As String) As String
    Public Function PrintInvoiceForAll(ByVal strinvoiceNo As String, ByVal docDate As DateTime, ByVal CustCode As String, ByVal isMainPrint As String, ByVal isCancel As String, ByVal ItemType As String) As String
        Dim Qry As String = Nothing

        If clsCommon.myLen(strinvoiceNo) > 0 Then

            Dim SD_SALE_INVOICE_DETAIL As String = Nothing
            Dim SD_SALE_INVOICE_HEAD As String = Nothing
            Dim SD_SHIPMENT_HEAD As String = Nothing
            Dim SD_SHIPMENT_DETAIL As String = Nothing
            If clsCommon.CompairString(isCancel, "Cancel") = CompairStringResult.Equal Then
                SD_SALE_INVOICE_DETAIL = " TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data As TSPL_SD_sale_invoice_DETAIL "
                SD_SALE_INVOICE_HEAD = " TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data As TSPL_SD_SALE_INVOICE_HEAD "
                SD_SHIPMENT_HEAD = " TSPL_SD_SHIPMENT_HEAD_Cancel_Data As TSPL_SD_SHIPMENT_HEAD "
                SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL_Cancel_Data As TSPL_SD_SHIPMENT_DETAIL "
            ElseIf clsCommon.CompairString(isCancel, "Delete") = CompairStringResult.Equal Then
                SD_SALE_INVOICE_DETAIL = " TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data As TSPL_SD_sale_invoice_DETAIL "
                SD_SALE_INVOICE_HEAD = " TSPL_SD_SALE_INVOICE_HEAD_Delete_Data As TSPL_SD_SALE_INVOICE_HEAD "
                SD_SHIPMENT_HEAD = " TSPL_SD_SHIPMENT_HEAD_Delete_Data As TSPL_SD_SHIPMENT_HEAD "
                SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL_Delete_Data As TSPL_SD_SHIPMENT_DETAIL "
            Else
                SD_SALE_INVOICE_DETAIL = " TSPL_SD_SALE_INVOICE_DETAIL"
                SD_SALE_INVOICE_HEAD = " TSPL_SD_SALE_INVOICE_HEAD "
                SD_SHIPMENT_HEAD = " TSPL_SD_SHIPMENT_HEAD "
                SD_SHIPMENT_DETAIL = " TSPL_SD_SHIPMENT_DETAIL "
            End If

            Dim OpeningBal As Decimal = 0
            Dim ClosingBal As Decimal = 0
            Dim AmtReceived As Decimal = 0
            Dim paymentMode As New List(Of String)
            Dim ManualScheme As Integer = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AutoSchemeOn & "'")) = 0, False, True)
            Dim ShowShipToPartyInDairyDispatch As Integer = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.ShowShipToPartyInDairyDispatch & "'")) = 0, 0, 1)
            LeakageDeduction_Freshsale = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.leakagededuction_freshsale & "'"))
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                'dt = clsDBFuncationality.GetDataTable(clsDairyInvoice.GetBalCustWise(clsCommon.GetPrintDate(docDate), CustCode))
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("EXEC SP_GetBalCustWise @Cust_Code = '" + clsCommon.myCstr(CustCode) + "',@DocDate='" + clsCommon.GetPrintDate(docDate) + "'")
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    OpeningBal = dt.Rows(0)("OpngBal")
                    ClosingBal = dt.Rows(0)("BalAmt")
                End If

            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                Dim dtchu As DataTable = clsDBFuncationality.GetDataTable("select Balance_Amt,Payment_Code from TSPL_RECEIPT_HEADER where Cust_Code='" + clsCommon.myCstr(CustCode) + "' and Receipt_Date='" + clsCommon.GetPrintDate(docDate) + "' and Posted='Y'")
                If dtchu IsNot Nothing AndAlso dtchu.Rows.Count > 0 Then
                    For Each dr As DataRow In dtchu.Rows
                        AmtReceived += clsCommon.myCdbl(dr("Balance_Amt"))
                        paymentMode.Add(clsCommon.myCstr(dr("Payment_Code")))
                    Next
                End If
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry = " select * from( "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry += " select "
            Else
                Qry = " select "
            End If
            'Qry = "   select "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                Qry += " 1 As CopyType, "
            End If
            If clsCommon.CompairString(isCancel, "Cancel") = CompairStringResult.Equal OrElse clsCommon.CompairString(isCancel, "Delete") = CompairStringResult.Equal Then
                Qry += " 'Cancelled' As Report_Status, "
            End If
            Qry += "'" & objCommonVar.CurrentUser & "' As PrintBy, "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                If clsCommon.CompairString(isCancel, "Cancel") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(isCancel, "Delete") <> CompairStringResult.Equal Then
                    Qry += " '' As Report_Status, "
                End If
                Qry += " ISNULL( CAST((SUM(Total_Basic_Amt) + SUM(Total_Tax_Amt)) / NULLIF(SUM(Qty_Default), 0)AS DECIMAL(18,2)),0) AS Basic_Rate, '" & objCommonVar.CurrentUserCode & "' as UserName,  max(final.PaymentTerms)PaymentTerms,max(Is_Distributor)Is_Distributor,max(Is_BPL)Is_BPL,max(Is_CashSale)Is_CashSale,
                            max(BPL_Coupon_Code)BPL_Coupon_Code,max(BPL_Name)BPL_Name,max(BPL_Remark)BPL_Remark,max(BPL_Coupon_Date)BPL_Coupon_Date,max(BPL_Category)BPL_Category,max(PO_Indent_No)PO_Indent_No,max(PO_Indent_Date)PO_Indent_Date,max(Booking_OpeningBal)Booking_OpeningBal,max(Booking_DrAmt)Booking_DrAmt,max(Booking_CrAmt)Booking_CrAmt,max(Booking_ClosingBal)Booking_ClosingBal,
                            max(Booking_ChequeNo)Booking_ChequeNo,max(Is_DCS)Is_DCS,max(Booking_Type)Booking_Type,max(CST_LST)CST_LST,max(DocumentTime)DocumentTime,max(Manual_VehicleNo)Manual_VehicleNo,max(Payment_Terms)Payment_Terms,max(ReceiverName)ReceiverName,sum(Amt_Less_Discount)Amt_Less_Discount,max(Dispatch_OpeningBal)Dispatch_OpeningBal,max(Dispatch_DrAmt)Dispatch_DrAmt,max(Dispatch_CrAmt)Dispatch_CrAmt,max(Dispatch_ClosingBal)Dispatch_ClosingBal,MAX(CrAmt)CrAmt,max(Security_TotalAmt)Security_TotalAmt,max(Supply_Date)Supply_Date,max(Shift_Type)Shift_Type,sum(QTY_LTRKG)QTY_LTRKG,
                            max(ITAX1)ITAX1,max(ITAX1_RATE)ITAX1_RATE,sum(ITAX1_Amt)ITAX1_Amt,sum(ITAX1_Base_Amt)ITAX1_Base_Amt,max(ITAX2)ITAX2,max(ITAX2_RATE)ITAX2_RATE,sum(ITAX2_Amt)ITAX2_Amt,sum(ITAX2_Base_Amt)ITAX2_Base_Amt,max(ITAX3)ITAX3,max(ITAX3_RATE)ITAX3_RATE,sum(ITAX3_Amt)ITAX3_Amt,sum(ITAX3_Base_Amt)ITAX3_Base_Amt,max(ITAX4)ITAX4,max(ITAX4_RATE)ITAX4_RATE,sum(ITAX4_Amt)ITAX4_Amt,sum(ITAX4_Base_Amt)ITAX4_Base_Amt,max(ITAX5)ITAX5,max(ITAX5_RATE)ITAX5_RATE,sum(ITAX5_Amt)ITAX5_Amt,sum(ITAX5_Base_Amt)ITAX5_Base_Amt,max(ITAX6)ITAX6,max(ITAX6_RATE)ITAX6_RATE,sum(ITAX6_Amt)ITAX6_Amt,sum(ITAX6_Base_Amt)ITAX6_Base_Amt,
                            max(ITAX7)ITAX7,max(ITAX7_RATE)ITAX7_RATE,sum(ITAX7_Amt)ITAX7_Amt,sum(ITAX7_Base_Amt)ITAX7_Base_Amt,max(ITAX8)ITAX8,max(ITAX8_RATE)ITAX8_RATE,sum(ITAX8_Amt)ITAX8_Amt,sum(ITAX8_Base_Amt)ITAX8_Base_Amt,max(ITAX9)ITAX9,max(ITAX9_RATE)ITAX9_RATE,sum(ITAX9_Amt)ITAX9_Amt,sum(ITAX9_Base_Amt)ITAX9_Base_Amt,max(ITAX10)ITAX10,max(ITAX10_RATE)ITAX10_RATE,sum(ITAX10_Amt)ITAX10_Amt,sum(ITAX10_Base_Amt)ITAX10_Base_Amt,max(IRN_No)IRN_No,max(Zone_Code)Zone_Code,max(CF)CF,max(ConversionFactor)ConversionFactor,max(EInvoice_Type)EInvoice_Type,max(LeakageDeduction_Freshsale)LeakageDeduction_Freshsale,max(LeakageDeduction)LeakageDeduction,max(SCM)SCM,max(DIS_MARGIN)DIS_MARGIN,max(NearestCity)NearestCity,
	                    	max(Location_Desc)Location_Desc,max(Loc_Short_Name)Loc_Short_Name,max(Loc_Pin)Loc_Pin,max(Loc_Phone)Loc_Phone,max(Loc_Eamil)Loc_Eamil,max(Loc_Website)Loc_Website,max(ISO_No)ISO_No,max(Invoice_No)Invoice_No,max(Invoice_Date)Invoice_Date,max(Cust_City)Cust_City,max(Against_Shipment_No)Against_Shipment_No,max(Cust_Gst_StateCode)Cust_Gst_StateCode,max(Electronic_Ref_No)Electronic_Ref_No,max(CustGSTNo)CustGSTNo,max(Area_Code)Area_Code,max(gst_state_code)gst_state_code,max(LocGstNo)LocGstNo,max(EWayBillNo)EWayBillNo,max(EWayBillDate)EWayBillDate,max(HSN_Code)HSN_Code, 	max(InvRemarks)InvRemarks,max(Delivery_Code)Delivery_Code,max(Conversion_factor)Conversion_factor,sum(QTY_Box)QTY_Box,max(Sale_Invoice_No)Sale_Invoice_No,max(vehicleNo)vehicleNo,max(Sale_Invoice_Date)Sale_Invoice_Date,max(RoundOffAmount)RoundOffAmount,max(Total_Amt)Total_Amt,max(Loc_ADd1)Loc_ADd1,max(LOC_ADD2)LOC_ADD2,max(LOC_ADD3)LOC_ADD3,max(LocationState)LocationState,max(LOCPhone)LOCPhone,max(Loc_TIN_NO)Loc_TIN_NO,(Document_Code)Document_Code,max(Document_Date)Document_Date,max(Description)Description,max(Lorry_No)Lorry_No,max(Sku_Seq)Sku_Seq,Item_Code,max(Line_No)Line_No,max(Item_Desc)Item_Desc,sum(QtyCrates)QtyCrates,max(ConvFactInCrate)ConvFactInCrate,sum(ConvQtyInCrate)ConvQtyInCrate,max(Unit_code)Unit_code,sum(Qty_Default)Qty_Default,
                            max(Rate_Default)Rate_Default,sum(QtyPCS)QtyPCS,sum(free_qty)free_qty, max(FreeSchemeInLitres)FreeSchemeInLitres,max(IsMultipleInvoice)IsMultipleInvoice,max(RatePerPcs)RatePerPcs,sum(valueInRs)valueInRs,max(comp_add2)comp_add2,max(comp_add3)comp_add3,max(CompPhone)CompPhone,max(Cash_Scheme_Amount)Cash_Scheme_Amount,max(schemeInCrates)schemeInCrates,max(GrandTotalCrates)GrandTotalCrates,max(Comp_Code)Comp_Code,max(Comp_Name)Comp_Name,max(comp_add1)comp_add1,max(comp_Fax)comp_Fax,max(comp_Email)comp_Email,max(comp_tinNo)comp_tinNo,max(cust_Code)cust_Code,max(Customer_Name)Customer_Name,
	                    	max(cust_add1)cust_add1,max(cust_add2)cust_add2,max(cust_add3)cust_add3,max(CustPhone)CustPhone,max(cust_fax)cust_fax,max(Cust_state)Cust_state,max(cust_Statename)cust_Statename,max(cust_Email)cust_Email,max(cust_website)cust_website,max(Customer_Pan)Customer_Pan,max(Ack_No)Ack_No,max(Ack_Date)Ack_Date,max(TaxableNonTaxable)TaxableNonTaxable,max(TAX1)TAX1,max(TaxType1)TaxType1,sum(TAX1_Amt)TAX1_Amt,max(TAX1_Rate)TAX1_Rate,sum(TAX1Amt)TAX1Amt,max(TAX2)TAX2,max(TaxType2)TaxType2,sum(TAX2_Amt)TAX2_Amt,max(TAX2_Rate)TAX2_Rate,sum(TAX2Amt)TAX2Amt,max(TAX3)TAX3,max(TaxType3)TaxType3,sum(TAX3_Amt)TAX3_Amt,max(TAX3_Rate)TAX3_Rate,sum(TAX3Amt)TAX3Amt,max(TAX4)TAX4,max(TaxType4)TaxType4,sum(TAX4_Amt)TAX4_Amt,max(TAX4_Rate)TAX4_Rate,sum(TAX4Amt)TAX4Amt,max(TAX5)TAX5,max(TaxType5)TaxType5,sum(TAX5_Amt)TAX5_Amt,max(TAX6)TAX6,max(TaxType6)TaxType6,sum(TAX6_Amt)TAX6_Amt,max(Route_No)Route_No,max(Route_Desc)Route_Desc,max(Distributor_Commission_TotalAmt)Distributor_Commission_TotalAmt,
                            max(Transporter_Commission_TotalAmt)Transporter_Commission_TotalAmt,max(Transport_Id)Transport_Id,max(Transporter_Name)Transporter_Name,max(Against_Delivery_Code)Against_Delivery_Code,max(batchNO)batchNO,max(Credit_Customer)Credit_Customer,max(Ship_To_Code)Ship_To_Code,max(Ship_To_Desc)Ship_To_Desc,max(Ship_Address)Ship_Address,max(Ship_City)Ship_City,max(Ship_State)Ship_State,max(Ship_Pin_Code)Ship_Pin_Code,max(Ship_PAN)Ship_PAN,max(Ship_GSTNO)Ship_GSTNO,max(Booth_Security_Amt)Booth_Security_Amt,max(Billing_Unit_code)Billing_Unit_code,sum(Billing_Qty)Billing_Qty,max(BulkCF)BulkCF,sum(Total_Basic_Amt)Total_Basic_Amt,Sum(Total_Tax_Amt)Total_Tax_Amt,max(Brand)Brand,max(BRANDDESC)BRANDDESC,max(Particulars)Particulars,
	                    	sum(Crate_No)Crate_No,max(CopyType)CopyType,max(SellerGST)SellerGST,max(Pan_No)Pan_No,max(Bank_Name)Bank_Name,max(BankAccountNo)BankAccountNo,max(BankBranchAddress)BankBranchAddress,max(BankIFSCCode)BankIFSCCode,max(Tcan_No)Tcan_No,max(RateLtr)RateLtr,max(Company_Name)Company_Name,max(Address2)Address2,max(Regn_No)Regn_No,max(FSSAI_NO)FSSAI_NO,max(Receipt_No)Receipt_No,max(Receipt_Date)Receipt_Date,max(Receipt_Amount)Receipt_Amount,max(Payment_Code)Payment_Code,max(cheque_No)cheque_No,max(Cheque_Date)Cheque_Date,max(OpeningBal)OpeningBal,max(ClosingBal)ClosingBal,max(cast(Logo_Img as varbinary(max))) as Logo_Img  "
            Else
                Qry += " '' As Report_Status, "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Qry += ", max(cast(BarCode_Img as varbinary(max))) as BarCode_Img "
            End If
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            '    Qry += " YYY.CopyType1,"
            'End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry += " from ( select Main_Final.*," + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal, "TSPL_COMPANY_MASTER.Access_Officer as FSSAI_NO,", "") + "TSPL_COMPANY_MASTER.Logo_Img,"
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                Qry += " Max(Final.PaymentTerms)PaymentTerms,
 Max(Final.Booking_OpeningBal)Booking_OpeningBal,Max(Final.Booking_DrAmt)Booking_DrAmt,
 Max(Final.Booking_CrAmt)Booking_CrAmt,Max(Final.Booking_ClosingBal)Booking_ClosingBal,
 max(cast(BarCode_Img as varbinary(max))) as BarCode_Img,
 --Max(Final.BarCode_Img)BarCode_Img,
 Max(Final.DocumentTime)DocumentTime,Max(Final.Manual_VehicleNo)Manual_VehicleNo,Max(Final.Payment_Terms)Payment_Terms,Max(Final.ReceiverName)ReceiverName,Sum(Final.Amt_Less_Discount)Amt_Less_Discount,Sum(Final.Security_TotalAmt)Security_TotalAmt,Max(Final.Supply_Date)Supply_Date,Max(Final.Shift_Type)Shift_Type,Sum(Final.QTY_LTRKG)QTY_LTRKG,SUM(Final.QtyInPouch)QtyInPouch,Sum(Final.QtyInPack)QtyInPack,Sum(Final.TCAmt)TCAmt,SUM(TCRate)TCRate,Max(Final.Rate_Per_Pouch)Rate_Per_Pouch,Max(Final.Rate_Per_Pack)Rate_Per_Pack,
 Max(Final.ITAX1)ITAX1,Max(Final.ITAX1_RATE)ITAX1_RATE,Sum(Final.ITAX1_Amt)ITAX1_Amt,Sum(Final.ITAX1_Base_Amt)ITAX1_Base_Amt,
 Max(Final.ITAX2)ITAX2,Max(Final.ITAX2_RATE)ITAX2_RATE,Sum(Final.ITAX2_Amt)ITAX2_Amt,Sum(Final.ITAX2_Base_Amt)ITAX2_Base_Amt,
 Max(Final.ITAX3)ITAX3,Max(Final.ITAX3_Rate)ITAX3_Rate,Sum(Final.ITAX3_Amt)ITAX3_Amt,Sum(Final.ITAX3_Base_Amt)ITAX3_Base_Amt,
 Max(Final.ITAX4)ITAX4,Max(Final.ITAX4_RATE)ITAX4_RATE,Sum(Final.ITAX4_Amt)ITAX4_Amt,Sum(Final.ITAX4_Base_Amt)ITAX4_Base_Amt,
 Max(Final.ITAX5)ITAX5,Max(Final.ITAX5_RATE)ITAX5_RATE,Sum(Final.ITAX5_Amt)ITAX5_Amt,Sum(Final.ITAX5_Base_Amt)ITAX5_Base_Amt,
 Max(Final.ITAX6)ITAX6,Max(Final.ITAX6_RATE)ITAX6_RATE,Sum(Final.ITAX6_Amt)ITAX6_Amt,Sum(Final.ITAX6_Base_Amt)ITAX6_Base_Amt,
 Max(Final.ITAX7)ITAX7,Max(Final.ITAX7_RATE)ITAX7_RATE,Sum(Final.ITAX7_Amt)ITAX7_Amt,Sum(Final.ITAX7_Base_Amt)ITAX7_Base_Amt,
 Max(Final.ITAX8)ITAX8,Max(Final.ITAX8_RATE)ITAX8_RATE,Sum(Final.ITAX8_Amt)ITAX8_Amt,Sum(Final.ITAX8_Base_Amt)ITAX8_Base_Amt,
 Max(Final.ITAX9)ITAX9,Max(Final.ITAX9_Rate)ITAX9_Rate,Sum(Final.ITAX9_Amt)ITAX9_Amt,Sum(Final.ITAX9_Base_Amt)ITAX9_Base_Amt,
 Max(Final.ITAX10)ITAX10,Max(Final.ITAX10_RATE)ITAX10_RATE,Sum(Final.ITAX10_Amt)ITAX10_Amt,Sum(Final.ITAX10_Base_Amt)ITAX10_Base_Amt,
 Max(Final.IRN_No)IRN_No,Max(Final.Zone_Code)Zone_Code,Max(Final.CF)CF,Max(Final.CF2)CF2,Max(Final.ConversionFactor)ConversionFactor,Max(Final.EInvoice_Type)EInvoice_Type,Max(Final.LeakageDeduction_Freshsale)LeakageDeduction_Freshsale,Max(Final.LeakageDeduction)LeakageDeduction,Max(Final.SCM)SCM,Max(Final.DIS_MARGIN)DIS_MARGIN,Max(Final.Location_Desc)Location_Desc,Max(Final.Loc_Short_Name)Loc_Short_Name,Max(Final.Loc_Pin)Loc_Pin,Max(Final.Loc_Phone)Loc_Phone,Max(Final.Loc_Eamil)Loc_Eamil,Max(Final.Loc_Website)Loc_Website,Max(Final.ISO_No)ISO_No,Max(Final.Invoice_No)Invoice_No,Max(Final.Invoice_Date)Invoice_Date,Max(Final.Cust_City)Cust_City,Max(Final.Against_Shipment_No)Against_Shipment_No,Max(Final.Cust_Gst_StateCode)Cust_Gst_StateCode,Max(Final.Electronic_Ref_No)Electronic_Ref_No,Max(Final.CustGSTNo)CustGSTNo,Max(Final.Area_Code)Area_Code,Max(Final.GST_STATE_Code)GST_STATE_Code,Max(Final.LocGstNo)LocGstNo,Max(Final.EWayBillNo)EWayBillNo,Max(Final.EWayBillDate)EWayBillDate,Max(Final.HSN_Code)HSN_Code,Max(Final.InvRemarks)InvRemarks,Max(Final.Conversion_Factor)Conversion_Factor,Sum(Final.QTY_Box)QTY_Box,Max(Final.Sale_Invoice_No)Sale_Invoice_No,Max(Final.vehicleNo)vehicleNo,Max(Final.Sale_Invoice_Date)Sale_Invoice_Date,Max(Final.RoundOffAmount)RoundOffAmount,Sum(Final.Total_Amt)Total_Amt,Max(Final.Loc_ADd1)Loc_ADd1,Max(Final.LOC_ADD2)LOC_ADD2,Max(Final.LOC_ADD3)LOC_ADD3,Max(Final.LocationState)LocationState,Max(Final.LOCPhone)LOCPhone,Max(Final.Loc_TIN_NO)Loc_TIN_NO,Final.Document_Code,Max(Final.Document_Date)Document_Date,Max(Final.Description)Description,
 
 Max(Final.Sku_Seq)Sku_Seq,
 Final.Item_Code,Max(Final.Print_Sequence)Print_Sequence,Max(Final.Line_No)Line_No,Max(Final.Item_Desc)Item_Desc,Sum(Final.QtyCrates)QtyCrates
 ,Max(Final.ConvFactInCrate)ConvFactInCrate,Max(Final.convQtyInCrate)convQtyInCrate,Max(final.Unit_code)Unit_code,Sum(Final.Qty_Default)Qty_Default,Max(Final.Rate_Default)Rate_Default,Sum(Final.QtyPCS)QtyPCS,Sum(Final.valueInRs)valueInRs,Max(Final.comp_add2)comp_add2,Max(Final.comp_add3)comp_add3,Max(Final.CompPhone)CompPhone,Max(Final.Cash_Scheme_Amount)Cash_Scheme_Amount,Max(Final.schemeInCrates)schemeInCrates
 ,--Sum(Final.GrandTotalCrates)GrandTotalCrates,max(cast(Logo_Img as varbinary(max))) As Logo_Img, max(cast(Logo_Img2 as varbinary(max))) As Logo_Img2,
 Max(Final.Comp_Code)Comp_Code,Max(Final.Comp_Name)Comp_Name,Max(Final.comp_add1)comp_add1,Max(Final.comp_Fax)comp_Fax,Max(Final.comp_Email)comp_Email,Max(Final.comp_tinNo)comp_tinNo,Max(Final.cust_Code)cust_Code,Max(Final.Customer_Name)Customer_Name,Max(Final.cust_add1)cust_add1,Max(Final.cust_add2)cust_add2
 ,Max(Final.cust_add3)cust_add3,Max(Final.CustPhone)CustPhone,Max(Final.cust_fax)cust_fax,Max(Final.Cust_state)Cust_state,Max(Final.cust_Statename)cust_Statename,Max(Final.cust_Email)cust_Email,Max(Final.cust_website)cust_website,Max(Final.Customer_Pan)Customer_Pan,Max(Final.Ack_No)Ack_No,Max(Final.Ack_Date)Ack_Date,Max(Final.TaxableNonTaxable)TaxableNonTaxable,
 max(TAX2)TAX2,max(TaxType2)TaxType2,sum(TAX2Amt)TAX2_Amt,max(TAX2_Rate)TAX2_Rate,
 max(TAX3)TAX3,max(TaxType3)TaxType3,sum(TAX3Amt)TAX3_Amt,max(TAX3_Rate)TAX3_Rate,
 max(TAX4)TAX4,max(TaxType4)TaxType4,sum(TAX4Amt)TAX4_Amt,max(TAX4_Rate)TAX4_Rate,
 max(TAX5)TAX5,max(TaxType5)TaxType5,sum(TAX5_Amt)TAX5_Amt,
 max(TAX6)TAX6,max(TaxType6)TaxType6,sum(TAX6_Amt)TAX6_Amt,
 max(TAX1)TAX1,max(TaxType1)TaxType1,sum(TAX1Amt)TAX1_Amt,max(TAX1_Rate)TAX1_Rate,

 
Max(Final.Route_No)Route_No,Max(Final.Route_Desc)Route_Desc,Sum(Final.Distributor_Commission_TotalAmt)Distributor_Commission_TotalAmt,
Sum(Final.Transporter_Commission_TotalAmt)Transporter_Commission_TotalAmt,Max(Final.Transport_Id)Transport_Id,Max(Final.Transporter_Name)Transporter_Name,Max(Final.Against_Delivery_Code)Against_Delivery_Code,Max(Final.batchNO)batchNO,Max(Final.Batchqty)Batchqty,Max(Final.Credit_Customer)Credit_Customer,Max(Final.Ship_To_Code)Ship_To_Code,Max(Final.Ship_To_Desc)Ship_To_Desc,
Max(ISNULL(Final.Ship_Address,'')) AS Ship_Address,max(Ship_City)Ship_City,Max(Final.Ship_State)Ship_State,Max(Final.Ship_Pin_Code)Ship_Pin_Code,Max(Final.Ship_PAN)Ship_PAN,Max(Final.Ship_GSTNO)Ship_GSTNO,Max(Final.Booth_Security_Amt)Booth_Security_Amt,Max(Final.Billing_Unit_code)Billing_Unit_code,Sum(Final.Billing_Qty)Billing_Qty,Max(Final.BulkCF)BulkCF,Max(Final.Total_Basic_Amt)Total_Basic_Amt,
Sum(Final.Total_Tax_Amt)Total_Tax_Amt,Max(Final.ActualRate)ActualRate,Sum(Final.Item_Net_Amt)Item_Net_Amt,Max(Final.Particulars)Particulars,
Max(Final.CopyType)CopyType,Max(Final.UserName)UserName,Max(Final.SellerGST)SellerGST,Max(Final.Pan_No)Pan_No,Max(Final.Bank_Name)Bank_Name,Max(Final.BankAccountNo)BankAccountNo,Max(Final.BankBranchAddress)BankBranchAddress,Max(Final.BankIFSCCode)BankIFSCCode,Max(Final.Tcan_No)Tcan_No,Max(Final.RateLtr)RateLtr
 
 from ( select Main_Final.*,"
            Else
                Qry += " Final.*  from ( select Main_Final.*," + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal, "TSPL_COMPANY_MASTER.Access_Officer as FSSAI_NO,", "") + "TSPL_COMPANY_MASTER.Logo_Img,"
            End If
            'Qry += " Final.*  from ( select Main_Final.*," + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal, "TSPL_COMPANY_MASTER.Access_Officer as FSSAI_NO,", "") + "TSPL_COMPANY_MASTER.Logo_Img,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
            Else
                Qry += "1 As CopyType,"
            End If


            Qry += " '" + objCommonVar.CurrentUserCode + "' as UserName, TSPL_COMPANY_MASTER.GSTReg_No As SellerGST,TSPL_COMPANY_MASTER.Pan_No,TSPL_COMPANY_MASTER.Bank_Name,TSPL_COMPANY_MASTER.BankAccountNo,TSPL_COMPANY_MASTER.BankBranchAddress,TSPL_COMPANY_MASTER.BankIFSCCode,TSPL_COMPANY_MASTER.Tcan_No, "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BHR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                Qry += " Convert(decimal(18,2),(valueInRs/((Qty_Default*ConversionFactor)/CF))) As RateLtr,Convert(decimal(18,2),(valueInRs/((Qty_Default*ConversionFactor)/CF2))) As Rate2 "
            Else
                Qry += " Rate_Default As RateLtr"

            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                Qry += ",TSPL_BOOKING_MATSER.ChequeNo"
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry += " , isnull(TSPL_COMPANY_MASTER.Comp_Name,'') as Company_Name,  isnull(TSPL_COMPANY_MASTER.Add2,'') as Address2,TSPL_COMPANY_MASTER.Regn_No,
  isnull(TSPL_COMPANY_MASTER.Access_Officer,'') as FSSAI_NO,

TSPL_RECEIPT_HEADER.Receipt_No,TSPL_RECEIPT_HEADER.Receipt_Date,TSPL_RECEIPT_HEADER.Receipt_Amount,

TSPL_RECEIPT_HEADER.Payment_Code,TSPL_RECEIPT_HEADER.cheque_No,TSPL_RECEIPT_HEADER.Cheque_Date, '" + clsCommon.myCstr(OpeningBal) + "' as OpeningBal,'" + clsCommon.myCstr(ClosingBal) + "' as ClosingBal "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                Qry += " ,'" + clsCommon.myCstr(AmtReceived) + "' as AmtReceived,'" + clsCommon.myCstr(clsCommon.GetMulcallStringWithComma(paymentMode)) + "' as PaymentMode,'" + clsCommon.myCstr(OpeningBal) + "' as OpeningBal,'" + clsCommon.myCstr(ClosingBal) + "' as ClosingBal  "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                Qry += " ,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 "
            End If
            Qry += "  from ( select final.*,tbl_Brand.Brand,tbl_Brand.BRANDDESC,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then

                Qry += " Item_Desc+'   '+isnull('HSN '+ HSN_Code,'') as Particulars"
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry += " Item_Desc as Particulars,FLOOR(Qty_Default / NULLIF(COALESCE(ConvFactInCrate, 0), 0)) AS Crate_No "
            Else
                'Qry += " Item_Desc+'   '+isnull(batchNO,'') as Particulars,FLOOR(Qty_Default / NULLIF(COALESCE(ConvFactInCrate, 0), 0)) AS Crate_No,isnull(batchNO,'') as PBatchNo "
                Qry += " Item_Desc as Particulars,FLOOR(Qty_Default / NULLIF(COALESCE(ConvFactInCrate, 0), 0)) AS Crate_No,isnull(batchNO,'') as PBatchNo "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                Qry += " From ( select TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Rate,"
            Else
                Qry += " from( Select "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                Qry += "  CINNo, "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
                Qry += " TSPL_SD_SALE_INVOICE_detail.item_cost,  "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                Qry += "  TSPL_BOOKING_MATSER.MOBILENO,TSPL_BOOKING_MATSER.AdharNo,   "

            End If
            Qry += "  case when TSPL_BOOKING_MATSER.Is_CashSale='Y' then TSPL_SD_SHIPMENT_HEAD.Payment_Terms else 'CREDIT' END AS PaymentTerms,TSPL_BOOKING_MATSER.Is_Distributor,TSPL_BOOKING_MATSER.Is_BPL,TSPL_BOOKING_MATSER.Is_CashSale,TSPL_BOOKING_MATSER.BPL_Coupon_Code,TSPL_BOOKING_MATSER.BPL_Name,TSPL_BOOKING_MATSER.BPL_Remark,TSPL_BOOKING_MATSER.BPL_Coupon_Date,TSPL_BOOKING_MATSER.BPL_Category,TSPL_BOOKING_MATSER.PO_Indent_No,Convert(varchar(12),TSPL_BOOKING_MATSER.PO_Indent_Date,103) as PO_Indent_Date,TSPL_BOOKING_MATSER.OpeningBal as Booking_OpeningBal,TSPL_BOOKING_MATSER.DrAmt as Booking_DrAmt,TSPL_BOOKING_MATSER.CrAmt as Booking_CrAmt,TSPL_BOOKING_MATSER.ClosingBal as Booking_ClosingBal,TSPL_BOOKING_MATSER.ChequeNo as Booking_ChequeNo , TSPL_BOOKING_MATSER.Is_DCS,TSPL_BOOKING_MATSER.Booking_Type, TSPL_COMPANY_MASTER.CST_LST,(Case When TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='T' Then cast(TSPL_SD_SALE_INVOICE_HEAD.BarCode_Img as image) End) as BarCode_Img,FORMAT(TSPL_SD_SHIPMENT_HEAD.Dispatch_date,'hh:mm:ss tt') As DocumentTime,
TSPL_SD_SHIPMENT_HEAD.ManualVehicle as Manual_VehicleNo,TSPL_SD_SHIPMENT_HEAD.Payment_Terms,TSPL_SD_SHIPMENT_HEAD.ReceiverName,TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount,
TSPL_SD_SHIPMENT_HEAD.OpeningBal as Dispatch_OpeningBal,TSPL_SD_SHIPMENT_HEAD.DrAmt as Dispatch_DrAmt,TSPL_SD_SHIPMENT_HEAD.CrAmt as Dispatch_CrAmt,TSPL_SD_SHIPMENT_HEAD.ClosingBal as Dispatch_ClosingBal,TSPL_SD_SHIPMENT_HEAD.CrAmt,
TSPL_SD_SHIPMENT_HEAD.Security_TotalAmt,convert(varchar(12),TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)Supply_Date,case when TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 'Morning' else 'Evening' end as Shift_Type, "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                Qry += " TSPL_SD_sale_invoice_DETAIL.Balance_Qty, "
                'TSPL_ITEM_PRICE_MASTER.Item_MRP as MRP ,
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " ((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SALE_INVOICE_DETAIL.Qty)/ITEMDETAIL3.Conversion_Factor) as QTY_LTRKG,"
            Else
                Qry += " Case when tspl_item_master.Is_FreshItem = 1 then (( TSPL_SD_sale_invoice_DETAIL.Qty *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1)) / coalesce(ITEMDETAIL3.LTR,1))  when tspl_item_master.Is_Ambient = 1 then ((TSPL_SD_sale_invoice_DETAIL.Qty *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1)) / coalesce(ITEMDETAIL3.kg,1)) end  as QTY_LTRKG,"
            End If
            Qry += " Case when tspl_item_master.Is_FreshItem = 1 then 'LTR' When tspl_item_master.Is_Ambient = 1 then 'KG' End As Qty_LTRKG_UOM,((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SALE_INVOICE_DETAIL.Qty)/ItemConversionInPouch.Conversion_Factor) as QtyInPouch,
((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SALE_INVOICE_DETAIL.Qty)/ItemConversionInPACK.Conversion_Factor) as QtyInPack,ISNull(TSPL_SD_SALE_INVOICE_DETAIL.Transporter_Commission_Rate,0.00) As TCRate,isnull(TSPL_SD_SALE_INVOICE_HEAD.Transporter_Commission_TotalAmt,0.00) as TCAmt,(TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt/((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SALE_INVOICE_DETAIL.Qty)/ItemConversionInPouch.Conversion_Factor)) as Rate_Per_Pouch,
(TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt/((TSPL_ITEM_UOM_DETAIL.Conversion_Factor * TSPL_SD_SALE_INVOICE_DETAIL.Qty)/ItemConversionInPACK.Conversion_Factor)) as Rate_Per_Pack,ItemConversionInPouch.UOM_Code As CPPouch_UOM_Code,ItemConversionInPACK.UOM_Code  As CPack_UOM_Code, "

            Qry += "TSPL_SD_SALE_INVOICE_DETAIL.TAX1 AS ITAX1,TSPL_SD_SALE_INVOICE_DETAIL.TAX1_RATE AS ITAX1_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt AS ITAX1_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Base_Amt AS ITAX1_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX2 AS ITAX2,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_RATE AS ITAX2_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt AS ITAX2_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Base_Amt AS ITAX2_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX3 AS ITAX3,TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate AS ITAX3_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt AS ITAX3_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Base_Amt AS ITAX3_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX4 AS ITAX4,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_RATE AS ITAX4_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt AS ITAX4_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Base_Amt AS ITAX4_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX5 AS ITAX5,TSPL_SD_SALE_INVOICE_DETAIL.TAX5_RATE AS ITAX5_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt AS ITAX5_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Base_Amt AS ITAX5_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX6 AS ITAX6,TSPL_SD_SALE_INVOICE_DETAIL.TAX6_RATE AS ITAX6_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt AS ITAX6_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Base_Amt AS ITAX6_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX7 AS ITAX7,TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate AS ITAX7_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt AS ITAX7_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Base_Amt AS ITAX7_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX8 AS ITAX8,TSPL_SD_SALE_INVOICE_DETAIL.TAX8_RATE AS ITAX8_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt AS ITAX8_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Base_Amt AS ITAX8_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX9 AS ITAX9,TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate AS ITAX9_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt AS ITAX9_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Base_Amt AS ITAX9_Base_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX10 AS ITAX10 ,TSPL_SD_SALE_INVOICE_DETAIL.TAX10_RATE AS ITAX10_RATE,TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt AS ITAX10_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Base_Amt AS ITAX10_Base_Amt, "

            Qry += "    (Case When TSPL_SD_SHIPMENT_HEAD.DO_Item_Type='T' Then 'IRN : '+ TSPL_SD_SALE_INVOICE_HEAD.IRN_No End) As IRN_No,TSPL_ROUTE_MASTER.Zone_Code,ITEMDETAIL1.Conversion_Factor As CF,ITEMDETAIL4.Conversion_Factor As CF2,TSPL_ITEM_UOM_DETAIL.Conversion_Factor As ConversionFactor,"
            If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
                Qry += "TSPL_SD_SALE_INVOICE_HEAD.Amount_Less_Discount*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
            Else
                Qry += " TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type,0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                If ManualScheme <> 0 Then
                    Qry += "case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, 
                case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN,"
                End If
            Else
                Qry += "case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0)   when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, 
                case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0)  when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN,"
            End If
            Qry += "  TSPL_LOCATION_MASTER.NearestCity, TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name ,  
TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin, (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone , 
TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website,TSPL_COMPANY_MASTER.ISO_No,TSPL_COMPANY_MASTER.ISO_Date,TSPL_COMPANY_MASTER.Access_Officer, 
TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date, 
customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,Tspl_customer_master.Area_Code,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor,  
(CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else (case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2),TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end) end) as QTY_Box ,  
TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No,  Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN  
ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo,  Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount,TSPL_SD_SALE_INVOICE_HEAD.Total_Amt,  
TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3,
TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, 
TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description,  
TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then TSPL_SD_sale_invoice_DETAIL.Item_Code + ' -Scheme' else
TSPL_SD_sale_invoice_DETAIL.Item_Code end as Item_Code,TSPL_ITEM_MASTER.Print_Sequence,(CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else (TSPL_SD_sale_invoice_DETAIL.Line_No) end) as Line_No,CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then  TSPL_ITEM_MASTER.Item_Desc + ' -Scheme' else TSPL_ITEM_MASTER.Item_Desc  end as Item_Desc ,   
TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,ITEMDETAIL2.Conversion_Factor As ConvFactInCrate,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " Cast((TSPL_SD_sale_invoice_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ITEMDETAIL2.Conversion_Factor As Int) As ConvQtyInCrate,"
            Else
                Qry += " CEILING((TSPL_SD_sale_invoice_DETAIL.Qty*TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ITEMDETAIL2.Conversion_Factor) As ConvQtyInCrate,"
            End If
            Qry += " ITEMDETAIL2.UOM_Code As CCrateUOM,TSPL_SD_sale_invoice_DETAIL.Unit_code, convert(Decimal(18,3), TSPL_SD_sale_invoice_DETAIL.Qty ) as Qty_Default
,(CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else (convert(Decimal(18,2), case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty )) else 0 end )) end) as Rate_Default,  
(CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else (convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) end) as QtyPCS , 
coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , TSPL_SD_sale_invoice_Head.IsMultipleInvoice, 
(CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else (case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end) end) as RatePerPcs, "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                Qry += "  (CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else ((case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else  TSPL_SD_sale_invoice_DETAIL.TAX1_Base_Amt end)) end)  as valueInRs,'' as  comp_add2,'' as comp_add3,TSPL_COMPANY_MASTER.Phone1 as CompPhone ,"
            Else
                Qry += "  (CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else ((case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else  TSPL_SD_sale_invoice_DETAIL.Amt_Less_Discount end)) end)  as valueInRs,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 , case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,"
            End If
            Qry += "(CASE when TSPL_SD_SALE_INVOICE_DETAIL.Scheme_Item='Y' then 0 else (coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0)) end) as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates,  
'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code1 As Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,  
TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email,  TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , 
TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SHIPMENT_HEAD.DO_Item_Type As TaxableNonTaxable,  
TSPL_SD_SALE_INVOICE_HEAD.TAX1,(select type from TSPL_TAX_MASTER where Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX1) as TaxType1, IsNull(TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt,0.00) As TAX1_Amt,TSPL_SD_SALE_INVOICE_HEAD.TAX1_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt as TAX1Amt, 
(select type from TSPL_TAX_MASTER where Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX2) as TaxType2,TSPL_SD_SALE_INVOICE_HEAD.TAX2, IsNull(TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt,0.00) As TAX2_Amt,TSPL_SD_SALE_INVOICE_HEAD.TAX2_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt as TAX2Amt,
(select type from TSPL_TAX_MASTER where Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX3) as TaxType3,TSPL_SD_SALE_INVOICE_HEAD.TAX3, IsNull(TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt,0.00) As TAX3_Amt,TSPL_SD_SALE_INVOICE_HEAD.TAX3_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt as TAX3Amt,
(select type from TSPL_TAX_MASTER where Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX4) as TaxType4,TSPL_SD_SALE_INVOICE_HEAD.TAX4, IsNull(TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt,0.00) As TAX4_Amt,TSPL_SD_SALE_INVOICE_HEAD.TAX4_Rate,TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt as TAX4Amt,
(select type from TSPL_TAX_MASTER where Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX5) as TaxType5,TSPL_SD_SALE_INVOICE_HEAD.TAX5, IsNull(TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt,0.00) As TAX5_Amt,
(select type from TSPL_TAX_MASTER where Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX6) as TaxType6,TSPL_SD_SALE_INVOICE_HEAD.TAX6, IsNull(TSPL_SD_SALE_INVOICE_HEAD.TAX6_Amt,0.00) As TAX6_Amt,TSPL_SD_SALE_INVOICE_HEAD.Route_No,TSPL_SD_SALE_INVOICE_HEAD.Route_Desc,TSPL_SD_SHIPMENT_HEAD.Distributor_Commission_TotalAmt,IsNull(TSPL_SD_SHIPMENT_HEAD.Transporter_Commission_TotalAmt,0.00) As Transporter_Commission_TotalAmt,TSPL_SD_SHIPMENT_HEAD.Transport_Id,TSPL_SD_SHIPMENT_HEAD.Transporter_Name, isnull(TSPL_SD_SHIPMENT_HEAD.Against_booking_no,'') as Against_Delivery_Code
,TabBatch.Batch_No as batchNO,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                Qry += " TabBatch.Batch_No1 as batchno1, "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                Qry += " TabBatch.BATCH_QTY AS Batchqty,"
            End If
            Qry += "
Case when TSPL_CUSTOMER_MASTER.Credit_Customer='Y' THEN 'CREDIT' else 'CASH' end as Credit_Customer,
TSPL_SHIP_TO_LOCATION.Ship_To_Code,TSPL_SHIP_TO_LOCATION.Ship_To_Desc,TSPL_SHIP_TO_LOCATION.Ship_Address,TSPL_SHIP_TO_LOCATION.Ship_City,
TSPL_SHIP_TO_LOCATION.Ship_State,Convert(Varchar,TSPL_SHIP_TO_LOCATION.Ship_Pin_Code)Ship_Pin_Code,TSPL_SHIP_TO_LOCATION.Ship_PAN,TSPL_SHIP_TO_LOCATION.Ship_GSTNO ,IsNull(TSPL_SD_SHIPMENT_DETAIL.Booth_Security_Amt,0)Booth_Security_Amt"

            ' If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            Qry += " ,TSPL_SD_SHIPMENT_DETAIL.Billing_Unit_code,TSPL_SD_SHIPMENT_DETAIL.Billing_Qty,ITEMBulkCF.Conversion_Factor as BulkCF,TSPL_SD_SHIPMENT_DETAIL.Total_Basic_Amt , TSPL_SD_SHIPMENT_DETAIL.Total_Tax_Amt,TSPL_SD_SALE_INVOICE_DETAIL.ActualRate,TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt  "
            'End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                Qry += " ,ITEMBulkCF.UOM_Code as BulkUOM,TSPL_CUSTOMER_TENDER_ORDER.Ref_No,convert(varchar(12),TSPL_CUSTOMER_TENDER_ORDER.Ref_Date,103) as Ref_Date "
            End If
            Qry += "from " + SD_SALE_INVOICE_DETAIL + "   
LEFT OUTER JOIN " + SD_SALE_INVOICE_HEAD + " ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE  
left outer join " + SD_SHIPMENT_HEAD + " on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No  
left outer join " + SD_SHIPMENT_DETAIL + " on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE  and TSPL_SD_SHIPMENT_DETAIL.Line_No=TSPL_SD_sale_invoice_DETAIL.Line_No
left outer join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
            '    Qry += " left outer join TSPL_BOOKING_DETAIL on TSPL_BOOKING_DETAIL.Document_no=TSPL_BOOKING_MATSER.Document_no "
            'End If
            Qry += " left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code And   
TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON  TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code
left join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_SD_SHIPMENT_HEAD.Route_No "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                Qry += " left join TSPL_CUSTOMER_TENDER_ORDER on TSPL_CUSTOMER_TENDER_ORDER.Document_Code=TSPL_SD_SHIPMENT_HEAD.Against_Cust_Order "
            End If
            Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code
left join ( select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code,TSPL_ITEM_UOM_DETAIL.UOM_Code from TSPL_ITEM_UOM_DETAIL where UOM_code in ( select UOM_Code from TSPL_item_uom_detail where Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code and TSPL_item_uom_detail.Bulk_UOM = 1 ) and TSPL_item_uom_detail.Bulk_UOM = 1 ) as ITEMBulkCF on ITEMBulkCF.Item_code = TSPL_SD_sale_invoice_DETAIL.Item_Code
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code in"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SKR") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JAL") = CompairStringResult.Equal Then
                Qry += " (select UOM_Code  from TSPL_item_uom_detail where Item_Code = TSPL_ITEM_UOM_DETAIL.Item_code and TSPL_item_uom_detail.Print_UOM=1)and TSPL_item_uom_detail.Print_UOM=1  "
            Else
                Qry += " (select case when Is_FreshItem=1 then 'LTR' else 'KG' end from TSPL_ITEM_MASTER where Item_Code=TSPL_ITEM_UOM_DETAIL.Item_code) "
            End If
            Qry += " ) as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JDH") = CompairStringResult.Equal Then
                Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,TSPL_ITEM_UOM_DETAIL.UOM_Code from TSPL_ITEM_UOM_DETAIL 
            where TSPL_ITEM_UOM_DETAIL.Print_UOM=1  ) as ITEMDETAIL4 on ITEMDETAIL4.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  "
            Else
                Qry += "   left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,TSPL_ITEM_UOM_DETAIL.UOM_Code from TSPL_ITEM_UOM_DETAIL 
where TSPL_ITEM_UOM_DETAIL.Print_UOM=1  ) as ITEMDETAIL4 on ITEMDETAIL4.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code And ITEMDETAIL4.UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code "
            End If




            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code And Print_UOM=1) as ITEMDETAIL3 on ITEMDETAIL3.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code "
            Else
                Qry += " left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) ITEMDETAIL3 ON TSPL_SD_sale_invoice_DETAIL.Item_Code = ITEMDETAIL3.item_code "
            End If
            'Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code And UOM_code IN ('Crate','CFC','TIN','JAR','Carton','Box')) as ITEMDETAIL2 on ITEMDETAIL2.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code And UOM_code IN ('Crate','CFC','TIN','JAR','Carton','Box')) as ITEMDETAIL2 on ITEMDETAIL2.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code "
            Else
                Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_code from TSPL_ITEM_UOM_DETAIL where ITEM_Code=TSPL_ITEM_UOM_DETAIL.Item_Code And UOM_code='Crate') as ITEMDETAIL2 on ITEMDETAIL2.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code "
            End If
            Qry += " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_Code from TSPL_ITEM_UOM_DETAIL where UOM_code IN ('CUP','Pouch')) as ItemConversionInPouch on ItemConversionInPouch.Item_code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code,UOM_Code from TSPL_ITEM_UOM_DETAIL where UOM_code In ('Matka','PACK')) as ItemConversionInPACK on ItemConversionInPACK.Item_code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code1 ='" & objCommonVar.CurrComp_Code1 & "'  
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code 
left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location 
LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State  
left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE  
left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code  
Left Outer Join (select Distinct Ship_To_Code,Ship_To_Desc,Ship_To_Type_Desc,(TSPL_SHIP_TO_LOCATION.Add1 +' '+TSPL_SHIP_TO_LOCATION.Add2+' '+TSPL_SHIP_TO_LOCATION.Add3+' '+TSPL_SHIP_TO_LOCATION.Add4) As Ship_Address,TSPL_CITY_MASTER.City_Name As Ship_City,TSPL_STATE_MASTER.STATE_NAME As Ship_State,Pin_Code As Ship_Pin_Code,PAN As Ship_PAN,GSTNO As Ship_GSTNO from TSPL_SHIP_TO_LOCATION
Left Outer Join TSPL_CITY_MASTER On TSPL_CITY_MASTER.City_Code=TSPL_SHIP_TO_LOCATION.City_Code
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_SHIP_TO_LOCATION.State) As TSPL_SHIP_TO_LOCATION ON TSPL_SHIP_TO_LOCATION.Ship_To_Code=TSPL_SD_SALE_INVOICE_HEAD.Ship_To_Location"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                'Qry += " LEFT OUTER JOIN TSPL_DISTRIBUTOR_COMMISSION_DETAIL ON TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Distributor_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code 
                '         AND TSPL_DISTRIBUTOR_COMMISSION_DETAIL.Route_Code = TSPL_SD_SHIPMENT_HEAD.Route_No"
                Qry += " LEFT OUTER JOIN TSPL_DISTRIBUTOR_COMMISSION_DETAIL ON TSPL_DISTRIBUTOR_COMMISSION_DETAIL.pk_id  =TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_PKID 		"
            End If
            Qry += " Full join  (select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from "
            If clsCommon.CompairString(isCancel, "Cancel") = CompairStringResult.Equal Then
                Qry += " TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data "
            ElseIf clsCommon.CompairString(isCancel, "Delete") = CompairStringResult.Equal Then
                Qry += " TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data "
            Else
                Qry += " TSPL_SD_SALE_INVOICE_DETAIL "
            End If
            Qry += " as inn "
            Qry += " where DOCUMENT_CODE in (" + strinvoiceNo + ") " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal, "", " and inn.Scheme_Item='Y'") + "   group by DOCUMENT_CODE,Item_Code)  
TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and  
TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on  
TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle"


            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                '                Qry += " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code  and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_shipment_head.Bill_To_Location 
                'and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and TSPL_ITEM_PRICE_MASTER.UOM=TSPL_SD_sale_invoice_DETAIL.Unit_code
                'and  len(isnull( TSPL_SD_shipment_head.Price_code  ,''))>0 "
            Else
                Qry += " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code  and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location
and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and  len(isnull( TSPL_SD_sale_invoice_DETAIL.Price_code  ,''))>0  "
            End If


            Qry += "
left outer join (select Document_Code,Parent_Line_No,"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                Qry += " STRING_AGG(Batch_No,CHAR(10)) as Batch_No , STRING_AGG(Qty,CHAR(10)) as Batch_Qty"
            Else
                Qry += " STRING_AGG(Batch_No +'('+convert(varchar(8),Qty)+')', ',') as Batch_No "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                Qry += " ,STRING_AGG(Batch_No,CHAR(10)) as Batch_No1 "
            End If
            Qry += "
                from(
SELECT Document_Code, Batch_No, Qty, Parent_Line_No FROM TSPL_BATCH_ITEM WHERE TSPL_BATCH_ITEM.Document_Type='FS-SH'
)x group by Document_Code,Parent_Line_No"
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
            '    Qry += " ,qty "
            'End If
            Qry += "        
)TabBatch On TabBatch.Document_Code= TSPL_SD_SHIPMENT_HEAD.Document_Code And TabBatch.Parent_Line_No = TSPL_SD_SALE_INVOICE_DETAIL.Line_No
                where 2 = 2 And TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   (" + strinvoiceNo + ")  "

            If clsCommon.myLen(ItemType) > 0 Then
                Qry += " and TSPL_SD_SALE_INVOICE_HEAD.item_type IN (" & ItemType & ") "
            End If
            Qry += "  And  exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    from " + SD_SALE_INVOICE_DETAIL + " 
Left OUTER JOIN " + SD_SALE_INVOICE_HEAD + " ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE 
Left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code And  TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code 
Left OUTER JOIN TSPL_ITEM_MASTER  ON  TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code 
Left outer join TSPL_COMPANY_MASTER on  TSPL_COMPANY_MASTER.Comp_Code1 ='" & objCommonVar.CurrComp_Code1 & "'  
Left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code 
Left outer join TSPL_LOCATION_MASTER  on  TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location 
Left OUTER JOIN   TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State 
Full Join(select DOCUMENT_CODE, Item_Code As Scheme_Item_Code, SUM(Qty) As SUB_QTY, SUM(Crate) AS schemeInCrates 
From "
            If clsCommon.CompairString(isCancel, "Cancel") = CompairStringResult.Equal Then
                Qry += " TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data "
            ElseIf clsCommon.CompairString(isCancel, "Delete") = CompairStringResult.Equal Then
                Qry += " TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data "
            Else
                Qry += " TSPL_SD_SALE_INVOICE_DETAIL "
            End If
            Qry += " as inn "
            Qry += " Where DOCUMENT_CODE In (" + strinvoiceNo + ") " + IIf(clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal, "", " and inn.Scheme_Item='Y'") + "   group by DOCUMENT_CODE,Item_Code)  TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and  TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code 
where 2=2 And  TSPL_SD_SALE_INVOICE_HEAD.Document_Code In   (" + strinvoiceNo + ") "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") <> CompairStringResult.Equal Then
                Qry += "   And TSPL_SD_sale_invoice_DETAIL.Scheme_Item ='N'  "
            End If
            Qry += " )xx where xx.Item_Code = TSPL_SD_sale_invoice_DETAIL.item_CODE And xx.DOCUMENT_CODE = TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE 
)
) as final 
Left outer join ( Select Item_Code, max([CATEGORY RM]) As [CATEGORY RM], max([BRAND]) As [BRAND], max([SUB BRAND]) As [SUB BRAND], max([DESCRP]) As [DESCRP], max([PACK]) As [PACK], max([PACK SIZE]) As [PACK SIZE], max([CATEGORY OT]) As [CATEGORY OT], max([CATEGORY FA]) As [CATEGORY FA], max([P TYPE]) As [P TYPE], max([L TYPE]) As [L TYPE], max([JW]) As [JW], max([SCRAP]) As [SCRAP], max([CATEGORY RMDESC]) As [CATEGORY RMDESC], max([BRANDDESC]) As [BRANDDESC], max([SUB BRANDDESC]) As [SUB BRANDDESC], max([DESCRPDESC]) As [DESCRPDESC], max([PACKDESC]) As [PACKDESC], max([PACK SIZEDESC]) As [PACK SIZEDESC], max([CATEGORY OTDESC]) As [CATEGORY OTDESC], max([CATEGORY FADESC]) As [CATEGORY FADESC], max([P TYPEDESC]) As [P TYPEDESC], max([L TYPEDESC]) As [L TYPEDESC], max([JWDESC]) As [JWDESC], max([SCRAPDESC]) As [SCRAPDESC]  from ( Select * from (   Select TSPL_ITEM_MASTER.Item_Code, TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code, TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code +'DESC' as Item_Category_CodeDesc,TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  
From TSPL_ITEM_MASTER    
Left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code 
Left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code And TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values   where 2=3 )xx  
Pivot(max(Item_Cagetory_Values) For Item_Category_Code   In ( [CATEGORY RM], [BRAND],[SUB BRAND], [DESCRP], [PACK],[PACK SIZE],[CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE], [JW], [SCRAP])  ) Pivt   Pivot  ( max(Category_Value_Desc) For Item_Category_CodeDesc In ([CATEGORY RMDESC], 
[BRANDDESC],[SUB BRANDDESC], [DESCRPDESC], [PACKDESC],[PACK SIZEDESC],[CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC], [JWDESC], [SCRAPDESC])  ) Pivt1 ) xxx  group by Item_Code )  as tbl_Brand on tbl_Brand.Item_Code=final.item_Code  ) AS Main_Final left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code1=Main_Final.comp_code "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry += "  left join TSPL_BOOKING_DETAIL on Main_Final.Against_Delivery_Code=TSPL_BOOKING_DETAIL.Document_No AND TSPL_BOOKING_DETAIL.Line_No=1
  Left Join TSPL_BOOKING_Matser on TSPL_BOOKING_DETAIL.document_No=TSPL_BOOKING_Matser.document_NO
  Left Join TSPL_RECEIPT_HEADER on TSPL_BOOKING_Matser.Against_Receipt_No=TSPL_RECEIPT_HEADER.Receipt_No  "
            End If
            Qry += " ) Final "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                Qry += " Group By Document_Code,Item_Code "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                Qry += " Order By Sku_Seq "
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                Qry += " Order By Print_Sequence "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RJS") = CompairStringResult.Equal Then
                Qry += " )xxx Left OUTER JOIN (Select 1 As COL1, 1 As COL2,  'ORIGINAL COPY' as CopyType1 UNION Select 1 as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=xxx.CopyType "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BAR") = CompairStringResult.Equal Then
                Qry += " group by Document_Code,Item_Code )XY 
                        LEFT JOIN ( SELECT Item_Code AS I1, BarCode_Img
                            FROM
                            ( SELECT TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_SD_SALE_INVOICE_HEAD.BarCode_Img,
                                    ROW_NUMBER() OVER (PARTITION BY TSPL_SD_SALE_INVOICE_DETAIL.Item_Code ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Date DESC ) AS RN
                                FROM TSPL_SD_SALE_INVOICE_DETAIL
                                JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.Document_Code
                                WHERE 2=2  AND TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   (" + strinvoiceNo + ")  ) X WHERE RN = 1 ) BC ON BC.I1 = XY.Item_Code "
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Qry += " group by Document_Code,Item_Code )XY  "
            End If
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            '    Qry += " Left OUTER JOIN (Select 1 As COL1, 1 As COL2,  'ORIGINAL COPY' as CopyType1 UNION Select 1 as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 3 as COL2, 
            '             'TRIPLICATE COPY' as CopyType1 UNION Select 1 as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=Final.CopyType ORDER BY Line_No,YYY.COL2 "
            'End If
            'Qry += " ) Final  order  by Final.Line_No asc,Final.Sku_Seq"
        End If
        Return Qry

    End Function

    Public Function PrintInvoiceForTruckSheetReport(ByVal FromDate, ByVal ToDate, ByVal whrcls, ByVal whrcls2) As String
        Dim Qry As String = Nothing
        Qry = " Select '" + FromDate + "' As 'From_Date',Convert(Varchar(10),'" + ToDate + "') As 'To_Date',*, TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.GSTReg_No As SellerGST,TSPL_COMPANY_MASTER.Pan_No,(Convert(decimal(18,2),(Item_Net_Amt/((QtyCrates*ConversionFactor)/CF)))) As RateLtr, Convert(decimal(18,2),(valueInRs/((Qty_Default*ConversionFactor)/CF2))) As Rate2  from( select Max(Zone_Code)Zone_Code,Max(Structure_Code)Structure_Code,Max(Main_Final.EInvoice_Type)EInvoice_Type,Max(Main_Final.CF)CF,	Max(Main_Final.ConversionFactor)ConversionFactor,Max(Main_Final.Location_Desc)Location_Desc,	Max(Main_Final.Loc_Short_Name)Loc_Short_Name,	Max(Main_Final.Loc_Pin)Loc_Pin,	
                    Max(Main_Final.Loc_Phone)Loc_Phone, Max(Main_Final.Loc_Eamil)Loc_Eamil,	Max(Main_Final.Loc_Website)Loc_Website,	Max(Main_Final.Invoice_Date)Invoice_Date,	Max(Main_Final.Cust_City)Cust_City,	Max(Main_Final.Cust_Gst_StateCode)Cust_Gst_StateCode,	Max(Main_Final.CustGSTNo)CustGSTNo,	Max(Main_Final.gst_state_code)gst_state_code,	Max(LocGstNo)LocGstNo,	Max(HSN_Code)HSN_Code,	
                    Max(Main_Final.InvRemarks)InvRemarks,	Sum(Main_Final.QTY_Box)QTY_Box,	Max(Main_Final.vehicleNo)vehicleNo,	Max(Main_Final.Sale_Invoice_Date)Sale_Invoice_Date,Sum(Main_Final.RoundOffAmount)RoundOffAmount,	Max(Main_Final.Loc_ADd1)Loc_ADd1,	Max(Main_Final.LOC_ADD2)LOC_ADD2,	Max(Main_Final.LOC_ADD3)LOC_ADD3,	Max(Main_Final.LocationState)LocationState,	Max(Main_Final.LOCPhone)LOCPhone,	Max(Main_Final.Loc_TIN_NO)Loc_TIN_NO,Max(Main_Final.Description)Description,Max(Main_Final.Sku_Seq)Sku_Seq,	Main_Final.Item_Code,	Max(Main_Final.Line_No)Line_No,	Max(Main_Final.Particulars)Particulars,	Sum(Main_Final.QtyCrates)QtyCrates,	Max(Main_Final.Unit_code)Unit_code,	Max(Main_Final.Qty_Default)Qty_Default,Max(Main_Final.Rate_Default)Rate_Default,	Max(Main_Final.QtyPCS)QtyPCS,	Max(Main_Final.free_qty)free_qty,
                    Max(Main_Final.FreeSchemeInLitres)FreeSchemeInLitres,	Max(Main_Final.RatePerPcs)RatePerPcs,Sum(Main_Final.valueInRs)valueInRs,(Sum(Main_Final.Item_Net_Amt)-Sum(Main_Final.TotalTaxAmt))Item_Net_Amt,Sum(Main_Final.Cash_Scheme_Amount)Cash_Scheme_Amount,	Sum(Main_Final.schemeInCrates)schemeInCrates,
                    Max(Main_Final.GrandTotalCrates)GrandTotalCrates,	Max(Main_Final.Comp_Code)Comp_Code,	Max(Main_Final.Comp_Name)Comp_Name,	Max(Main_Final.comp_add1)comp_add1,	Max(Main_Final.comp_add2)comp_add2,	Max(Main_Final.comp_add3)comp_add3,	Max(Main_Final.comp_Fax)comp_Fax,	Max(Main_Final.comp_Email)comp_Email,	Max(Main_Final.CompPhone)CompPhone,	Max(Main_Final.comp_tinNo)comp_tinNo,	Max(Main_Final.cust_Code)cust_Code,	Max(Main_Final.Customer_Name)Customer_Name,Max(Main_Final.cust_add1)cust_add1,	Max(cust_add2)cust_add2,	Max(Main_Final.cust_add3)cust_add3,	Max(Main_Final.CustPhone)CustPhone,	Max(Main_Final.cust_fax)cust_fax,	Max(Main_Final.Cust_state)Cust_state,	Max(Main_Final.cust_Statename)cust_Statename,	Max(Main_Final.cust_Email)cust_Email,	Max(Main_Final.cust_website)cust_website,	Max(Main_Final.Customer_Pan)Customer_Pan,	Max(Main_Final.Ack_No)Ack_No,	Max(Main_Final.Ack_Date)Ack_Date,Max(Main_Final.TaxableNonTaxable)TaxableNonTaxable,	Max(Main_Final.TAX1)TAX1,	Sum(Main_Final.TAX1_Amt)TAX1_Amt,	Max(Main_Final.TAX2)TAX2,	Sum(Main_Final.TAX2_Amt)TAX2_Amt,	Max(Main_Final.TAX3)TAX3,	Sum(Main_Final.TAX3_Amt)TAX3_Amt,	Max(Main_Final.TAX4)TAX4,	Sum(Main_Final.TAX4_Amt)TAX4_Amt,	Max(Main_Final.TAX5)TAX5,	Sum(Main_Final.TAX5_Amt)TAX5_Amt,	Max(Main_Final.TAX6)TAX6,	Sum(Main_Final.TAX6_Amt)TAX6_Amt,Sum(TotalTaxAmt)TotalTaxAmt,		
                    Max(Main_Final.Brand)Brand,Max(Main_Final.BRANDDESC)BRANDDESC
                    from( select final.*, tbl_Brand.Brand, tbl_Brand.BRANDDESC 
                    from( select  TSPL_ITEM_MASTER.Structure_Code, Zone_Code, ITEMDETAIL1.Conversion_Factor As CF, TSPL_ITEM_UOM_DETAIL.Conversion_Factor As ConversionFactor, TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type, 0 As LeakageDeduction_Freshsale, 0 As LeakageDeduction, TSPL_LOCATION_MASTER.Location_Desc, TSPL_LOCATION_MASTER.Loc_Short_Name, TSPL_LOCATION_MASTER.Pin_Code As Loc_Pin, (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone , TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date, customer_city_master.city_name as Cust_City,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,ITEMDETAIL.Conversion_factor, case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end as QTY_Box,  Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No,TSPL_ITEM_MASTER.Item_Desc as Particulars,  TSPL_SD_sale_invoice_DETAIL.Qty as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code, convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty ) as Qty_Default,  convert(Decimal(18,2), case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty )) else 0 end ) as Rate_Default, convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor) as QtyPCS , coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, (case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs,(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Item_Net_Amt end)  as Item_Net_Amt, coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, '' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SHIPMENT_HEAD.DO_Item_Type As TaxableNonTaxable, 
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX1, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt, 0.00) As TAX1_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX2, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt, 0.00) As TAX2_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX3, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt, 0.00) As TAX3_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX4, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt, 0.00) As TAX4_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX5, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt, 0.00) As TAX5_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX6, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt, 0.00) As TAX6_Amt,
                    (IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt,0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt, 0.00)) As TotalTaxAmt from TSPL_SD_sale_invoice_DETAIL  
				    Left OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE 
                    Left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No 
                    Left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code And  TSPL_ITEM_UOM_DETAIL.UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code 
                    Left OUTER JOIN TSPL_ITEM_MASTER  ON TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code 
                    Left Join(select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  
                Left Join(select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='LTR') as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  
                Left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State  left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE  left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code  
                    Full Join(select Item_Code as Scheme_Item_Code, SUM(Qty) As SUB_QTY, SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn
                    where inn.Scheme_Item ='Y'   group by Item_Code) TSPL_SD_sale_invoice_DETAIL_Sub on 
                    TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code = TSPL_SD_sale_invoice_DETAIL.Item_Code  
                    Left OUTER JOIN TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle  left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code  And TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location  And TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code 
                    where 2 = 2 And TSPL_SD_sale_invoice_DETAIL.Scheme_Item ='N' " + whrcls2 + "				
                    UNION aLL  				
                    Select Case TSPL_ITEM_MASTER.Structure_Code,Zone_Code,ITEMDETAIL1.Conversion_Factor As CF,TSPL_ITEM_UOM_DETAIL.Conversion_Factor As ConversionFactor, TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type,0 As LeakageDeduction_Freshsale,0 As LeakageDeduction,     TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name ,  TSPL_LOCATION_MASTER.Pin_Code As Loc_Pin, (Case When isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone , TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website, convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date, customer_city_master.city_name as Cust_City,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks, ITEMDETAIL.Conversion_factor, case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box, Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,  TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, TSPL_SD_SALE_INVOICE_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_sale_invoice_DETAIL.Item_Code,0 as Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars, TSPL_SD_sale_invoice_DETAIL.Qty as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code ,0 as QtyPCS , 0 as Qty_Default , 0 as Rate_Default  , coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres ,0 as RatePerPcs,0  as valueInRs,0 as Item_Net_Amt,0 as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, '' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan,IsNull(TSPL_SD_SALE_INVOICE_HEAD.Ack_No,'NA') AS Ack_No,TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SHIPMENT_HEAD.DO_Item_Type As TaxableNonTaxable,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX1, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt, 0.00) As TAX1_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX2, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt, 0.00) As TAX2_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX3, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt, 0.00) As TAX3_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX4, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt, 0.00) As TAX4_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX5, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt, 0.00) As TAX5_Amt,
                    TSPL_SD_SALE_INVOICE_DETAIL.TAX6, IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt, 0.00) As TAX6_Amt,
                    (IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt,0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt, 0.00)+
					IsNull(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt, 0.00)) As TotalTaxAmt from TSPL_SD_sale_invoice_DETAIL  
					 Left OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE 
                     Left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No 
                     Left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code And TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code 
                     Left OUTER JOIN TSPL_ITEM_MASTER  ON TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code 
                     Left Join(select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code   
                    Left Join(select Conversion_factor, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='LTR') as ITEMDETAIL1 on ITEMDETAIL1.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code   
                    Left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  
                     Left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code 
                     Left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State  
                     Left Join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE  left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code  
                     Full Join(select Item_Code as Scheme_Item_Code, SUM(Qty) As SUB_QTY, SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn 
                     where  inn.Scheme_Item ='Y' group by Item_Code) TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  
                     Left OUTER JOIN TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle  
                     Left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code  And TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location  And TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code   left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code  
                     where 2 = 2 And TSPL_SD_sale_invoice_DETAIL.Scheme_Item ='y'    and not exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    
                            From TSPL_SD_sale_invoice_DETAIL 
                            Left OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE 
                            Left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code And TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code 
                            Left OUTER JOIN TSPL_ITEM_MASTER  ON TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code 
                            Left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  
                            Left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code 
                            Left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location 
                            Left OUTER JOIN  TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State 
                            Full Join(select Item_Code as Scheme_Item_Code, SUM(Qty) As SUB_QTY, SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn 
                            where inn.Scheme_Item ='Y'   group by Item_Code) TSPL_SD_sale_invoice_DETAIL_Sub on 
                            TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code = TSPL_SD_sale_invoice_DETAIL.Item_Code 
                            where 2 = 2 And TSPL_SD_sale_invoice_DETAIL.Scheme_Item ='N' " + whrcls2 + ")xx where xx.Item_Code=TSPL_SD_sale_invoice_DETAIL.item_CODE and xx.DOCUMENT_CODE=TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE ) ) as final left outer join  ( select Item_Code,max([CATEGORY RM]) as [CATEGORY RM],max([BRAND]) as [BRAND],max([SUB BRAND]) as [SUB BRAND],max([DESCRP]) as [DESCRP],max([PACK]) as [PACK], max([PACK SIZE]) as [PACK SIZE],max([CATEGORY OT]) as [CATEGORY OT],max([CATEGORY FA]) as [CATEGORY FA],max([P TYPE]) as [P TYPE],max([L TYPE]) as [L TYPE],max([JW]) as [JW], max([SCRAP]) as [SCRAP],max([CATEGORY RMDESC]) as [CATEGORY RMDESC],max([BRANDDESC]) as [BRANDDESC],max([SUB BRANDDESC]) as [SUB BRANDDESC],max([DESCRPDESC]) as [DESCRPDESC],max([PACKDESC]) as [PACKDESC], max([PACK SIZEDESC]) as [PACK SIZEDESC],max([CATEGORY OTDESC]) as [CATEGORY OTDESC],max([CATEGORY FADESC]) as [CATEGORY FADESC],max([P TYPEDESC]) as [P TYPEDESC],max([L TYPEDESC]) as [L TYPEDESC],max([JWDESC]) as [JWDESC], max([SCRAPDESC]) as [SCRAPDESC]  from ( select * from (   select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code   ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+'DESC' as Item_Category_CodeDesc   , TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  from  TSPL_ITEM_MASTER    left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code  left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values   where 2=2 )xx  Pivot   ( max(Item_Cagetory_Values) for Item_Category_Code   in ( [CATEGORY RM],[BRAND],[SUB BRAND],[DESCRP],[PACK],[PACK SIZE],[CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE],[JW],[SCRAP])  ) Pivt   Pivot  ( max(Category_Value_Desc) for Item_Category_CodeDesc in ([CATEGORY RMDESC], [BRANDDESC],[SUB BRANDDESC],[DESCRPDESC],[PACKDESC],[PACK SIZEDESC],[CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC],[JWDESC],[SCRAPDESC])  ) Pivt1 ) xxx  group by Item_Code )  as tbl_Brand on tbl_Brand.Item_Code=final.item_Code  ) AS Main_Final 
                            " + whrcls + "
                            Group By Main_Final.Item_Code
                            ) xyz
							Left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=xyz.comp_code
                            order  by Line_No asc"
        Return Qry
    End Function



    'Ticket -MIL/29/04/19-000071
    Public Function GetQueryGMD(Optional ByVal strinvoiceNo As String = Nothing) As String
        Dim Qry As String = ""
        Qry = "  Select Main_Final.*, TSPL_COMPANY_MASTER.Logo_Img, TSPL_COMPANY_MASTER.Logo_Img2 from ( Select DISTINCT final.*, tbl_Brand.Brand, tbl_Brand.BRANDDESC from ( " &
                     "select "
        If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
            Qry += "TSPL_SD_SALE_INVOICE_HEAD.Amount_Less_Discount*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
        Else
            Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
        End If

        Qry += " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, " &
        " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
        " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN,  TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
      " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
      " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
      " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
      " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
      " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
        " customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor, " &
        "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2), " &
        "TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end as QTY_Box, " &
        "TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
        "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
        "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
        "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
        "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, " &
        "TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
        "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
        "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No,TSPL_ITEM_MASTER.Item_Desc as Particulars,  " &
        "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code , " &
        "convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor) as QtyPCS , " &
        "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , " &
        "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, " &
        "(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs " &
        ",convert(DECIMAL(18,5),CASE WHEN TSPL_UNIT_MASTER.Crate_Type='Y' THEN TSPL_ITEM_UOM_DETAIL.Conversion_Factor ELSE 0 END * " &
         " case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0  " &
         " else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end) AS RatePerCrate " &
        ",coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
        "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , " &
        "TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
        "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
        "from TSPL_SD_sale_invoice_DETAIL  " &
        "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
        "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
        "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and  " &
        "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
        "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code " &
        "left join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_ITEM_UOM_DETAIL.UOM_Code and TSPL_UNIT_MASTER.Crate_Type='Y'" &
        " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
         " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
        " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
        " Full join  (select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
        "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
        "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
        "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
        "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
        " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code " &
        " and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location " &
        " and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
        " where 2=2 and " &
        "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N' " &
        "UNION aLL " &
        " select "

        If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
            Qry += "TSPL_SD_SALE_INVOICE_HEAD.Total_Amt*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
        Else
            Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
        End If


        Qry += " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, " &
                " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN, TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
              " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
              " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
              " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
              " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
              " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
                "customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code , ITEMDETAIL.Conversion_factor, " &
                "case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,  TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
                "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_sale_invoice_DETAIL.Item_Code,0 as Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars, " &
                "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code ,0 as QtyPCS  , " &
                "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, " &
                "TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres ,0 as RatePerPcs,0 as RatePerCrate,0  as valueInRs,0 as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
                "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
                "from TSPL_SD_sale_invoice_DETAIL  " &
                "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join " &
                "(select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX' ) as ITEMDETAIL on " &
                "ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on " &
                "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN " &
                "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
               " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                " Full join " &
                "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn " &
                "where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
                " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code " &
                " and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location " &
                " and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code   left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
                " where 2=2 and  TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')   " &
                "and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='y'    and not exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, " &
                "TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    from TSPL_SD_sale_invoice_DETAIL " &
                "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left outer join TSPL_COMPANY_MASTER on " &
                "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN  " &
                "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State Full join " &
                "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code " &
                " where 2=2 and " &
                "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')  and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N')xx " &
                "where xx.Item_Code=TSPL_SD_sale_invoice_DETAIL.item_CODE and xx.DOCUMENT_CODE=TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE ) " &
                ") as final left outer join " &
                " ( select Item_Code,max([CATEGORY RM]) as [CATEGORY RM],max([BRAND]) as [BRAND],max([SUB BRAND]) as [SUB BRAND],max([DESCRP]) as [DESCRP],max([PACK]) as [PACK]," &
                " max([PACK SIZE]) as [PACK SIZE],max([CATEGORY OT]) as [CATEGORY OT],max([CATEGORY FA]) as [CATEGORY FA],max([P TYPE]) as [P TYPE],max([L TYPE]) as [L TYPE],max([JW]) as [JW]," &
                " max([SCRAP]) as [SCRAP],max([CATEGORY RMDESC]) as [CATEGORY RMDESC],max([BRANDDESC]) as [BRANDDESC],max([SUB BRANDDESC]) as [SUB BRANDDESC],max([DESCRPDESC]) as [DESCRPDESC],max([PACKDESC]) as [PACKDESC]," &
                " max([PACK SIZEDESC]) as [PACK SIZEDESC],max([CATEGORY OTDESC]) as [CATEGORY OTDESC],max([CATEGORY FADESC]) as [CATEGORY FADESC],max([P TYPEDESC]) as [P TYPEDESC],max([L TYPEDESC]) as [L TYPEDESC],max([JWDESC]) as [JWDESC]," &
                " max([SCRAPDESC]) as [SCRAPDESC]  from ( select * from (   select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code   ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+'DESC' as Item_Category_CodeDesc   ," &
                " TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  from  TSPL_ITEM_MASTER    left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code " &
                " left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values   where 2=2 )xx " &
                " Pivot   ( max(Item_Cagetory_Values) for Item_Category_Code   in ( [CATEGORY RM],[BRAND],[SUB BRAND],[DESCRP],[PACK],[PACK SIZE],[CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE],[JW],[SCRAP])  ) Pivt   Pivot  ( max(Category_Value_Desc) for Item_Category_CodeDesc in ([CATEGORY RMDESC]," &
                " [BRANDDESC],[SUB BRANDDESC],[DESCRPDESC],[PACKDESC],[PACK SIZEDESC],[CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC],[JWDESC],[SCRAPDESC])  ) Pivt1 ) xxx  group by Item_Code )  as tbl_Brand on tbl_Brand.Item_Code=final.item_Code  ) AS Main_Final left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=Main_Final.comp_code " &
                    " order  by Main_Final.Sku_Seq ,Main_Final.Line_No"

        Return Qry
    End Function
    '========Added by preeti gupta Against ticket no[ALF/08/05/19-000104]
    Public Function GetQueryAlpha(Optional ByVal strinvoiceNo As String = Nothing) As String
        Dim Qry As String = ""
        Qry = " select  final.*,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from ( " &
                     "select   "
        If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
            Qry += "TSPL_SD_SALE_INVOICE_HEAD.Amount_Less_Discount*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
        Else
            Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
        End If

        Qry += " TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
          " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
          " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
          " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
          " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
          " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
            " customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor, " &
            "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2), " &
            "TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end as QTY_Box, " &
            "TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
            "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
            "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
            "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
            "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, " &
            "TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
            "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
            "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No,TSPL_ITEM_MASTER.Item_Desc as Particulars,  " &
            "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code , " &
            "convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor) as QtyPCS , " &
            "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , " &
            "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, " &
            "(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs " &
            ",convert(DECIMAL(18,5),CASE WHEN TSPL_UNIT_MASTER.Crate_Type='Y' THEN TSPL_ITEM_UOM_DETAIL.Conversion_Factor ELSE 0 END * " &
             " case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0  " &
             " else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end) AS RatePerCrate " &
            ",coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
            "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , " &
            "TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
            "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
            "from TSPL_SD_sale_invoice_DETAIL  " &
            "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
            "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
            "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and  " &
            "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
            "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code " &
            "left join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_ITEM_UOM_DETAIL.UOM_Code and TSPL_UNIT_MASTER.Crate_Type='Y'" &
            " left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
             " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
            " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
            " Full join  (select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
            "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
            "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
            "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
            "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
            "  left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
            " where 2=2 and " &
            "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N' " &
            "UNION aLL " &
            " select "

        If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
            Qry += "TSPL_SD_SALE_INVOICE_HEAD.Total_Amt*'" & LeakageDeduction_Freshsale & "'/100 as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
        Else
            Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
        End If


        Qry += " TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
              " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
              " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
              " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
              " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
              " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
                "customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code , ITEMDETAIL.Conversion_factor, " &
                "case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,  TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
                "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_sale_invoice_DETAIL.Item_Code,0 as Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars, " &
                "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code ,0 as QtyPCS  , " &
                "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, " &
                "TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres ,0 as RatePerPcs,0 as RatePerCrate,0  as valueInRs,0 as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
                "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
                "from TSPL_SD_sale_invoice_DETAIL  " &
                "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join " &
                "(select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX' ) as ITEMDETAIL on " &
                "ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on " &
                "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN " &
                "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
               " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                " Full join " &
                "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn " &
                "where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
                "   left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
                " where 2=2 and  TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')   " &
                "and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='y'    and not exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, " &
                "TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    from TSPL_SD_sale_invoice_DETAIL " &
                "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left outer join TSPL_COMPANY_MASTER on " &
                "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN  " &
                "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State Full join " &
                "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code " &
                " where 2=2 and " &
                "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')  and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N')xx " &
                "where xx.Item_Code=TSPL_SD_sale_invoice_DETAIL.item_CODE and xx.DOCUMENT_CODE=TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE ) " &
                ") as final  left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=final.comp_code " &
                    " order  by final.Sku_Seq ,final.Line_No"

        Return Qry
    End Function
    Public Function GetQuerySwadesh(Optional ByVal strinvoiceNo As String = Nothing) As String
        'Sanjay Ticket no-SWA/01/04/19-000065 add Crate received
        LeakageDeduction_Freshsale = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.leakagededuction_freshsale & "'"))
        Dim Qry As String = ""
        Qry = "  select Main_Final.*,TSPL_COMPANY_MASTER.Logo_Img from ( select DISTINCT final.*,tbl_Brand.Brand,tbl_Brand.BRANDDESC from ( select TSPL_SD_SALE_INVOICE_HEAD.Total_Amt," &
                   " (select TOP 1 ISNULL(CrateQtyRecd,0) from tspl_crate_received_detail_freshsale " &
            " where tspl_crate_received_detail_freshsale.Customer_Code = TSPL_SD_sale_invoice_head.Customer_Code " &
            " and tspl_crate_received_detail_freshsale.Sale_Invoice_Date<TSPL_SD_sale_invoice_head.Document_Date ORDER BY Sale_Invoice_Date desc) as CrateQtyRecd," &
            " isnull(TSPL_SD_SALE_RETURN_HEAD.Is_Cancelled,0) as Is_Cancelled,"

        If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
            Qry += "TSPL_SD_SALE_INVOICE_HEAD.LeakageAmount as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
        Else
            Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
        End If

        Qry += " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, " &
            " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
            " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN,  TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
          " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
          " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
          " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
          " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
          " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
            " customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code ,ITEMDETAIL.Conversion_factor, " &
            "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then 0 else convert(Decimal(18,2), " &
            "TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor/coalesce(ITEMDETAIL.Conversion_factor,1)) end else 0 end as QTY_Box, " &
            "TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
            "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
            "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
            "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date, TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
            "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, " &
            "TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
            "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
            "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No,TSPL_ITEM_MASTER.Sku_Seq ,TSPL_SD_sale_invoice_DETAIL.Item_Code,TSPL_SD_sale_invoice_DETAIL.Line_No,TSPL_ITEM_MASTER.Item_Desc as Particulars,  " &
            "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code , " &
            "convert(Decimal(18,2), TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor) as QtyPCS , " &
            "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres , " &
            "case when TSPL_SD_sale_invoice_DETAIL.Qty > 0 then convert(DECIMAL(18,5),(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)/ (TSPL_SD_sale_invoice_DETAIL.Qty *TSPL_ITEM_UOM_DETAIL.Conversion_Factor)) else 0 end as RatePerPcs, " &
            "(case when TSPL_SD_sale_invoice_DETAIL.Sampling=1then 0 else  TSPL_SD_sale_invoice_DETAIL.Amount end)  as valueInRs, " &
            "coalesce(TSPL_SD_sale_invoice_DETAIL.Cash_Scheme_Amount,0) as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
            "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 , " &
            "TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
            "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
            "from TSPL_SD_sale_invoice_DETAIL  " &
            "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
            "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
            "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and  " &
            "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
            "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join (select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX') as ITEMDETAIL on ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
             " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
            " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
            " Full join  (select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
            "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
            "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
            "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
            "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
            " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code " &
            " and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location " &
            " and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
            " left join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No=TSPL_SD_SALE_INVOICE_HEAD.document_code where 2=2 and " &
            "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')    and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N' " &
            "UNION aLL " &
            " select TSPL_SD_SALE_INVOICE_HEAD.Total_Amt, " &
             " (select TOP 1 ISNULL(CrateQtyRecd,0) from tspl_crate_received_detail_freshsale " &
            " where tspl_crate_received_detail_freshsale.Customer_Code = TSPL_SD_sale_invoice_head.Customer_Code " &
            " and tspl_crate_received_detail_freshsale.Sale_Invoice_Date<TSPL_SD_sale_invoice_head.Document_Date ORDER BY Sale_Invoice_Date desc) as CrateQtyRecd," &
            " isnull(TSPL_SD_SALE_RETURN_HEAD.Is_Cancelled,0) as Is_Cancelled,"

        If clsCommon.myCdbl(LeakageDeduction_Freshsale) > 0 Then
            Qry += "TSPL_SD_SALE_INVOICE_HEAD.LeakageAmount as LeakageDeduction_Freshsale,1 as LeakageDeduction, "
        Else
            Qry += " 0 as LeakageDeduction_Freshsale,0 as LeakageDeduction, "
        End If


        Qry += " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='SCM' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as SCM, " &
                " case when TSPL_ITEM_PRICE_MASTER.Price_Comp1='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount1,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp2='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount2,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp3='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount3,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp4='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount4,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp5='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount5,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp6='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount6,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp7='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount7,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp8='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount8,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp9='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount9,0) " &
                " when TSPL_ITEM_PRICE_MASTER.Price_Comp10='DIS-MARGIN' then isnull(TSPL_SD_sale_invoice_DETAIL.Price_Amount10,0) end as DIS_MARGIN, TSPL_LOCATION_MASTER.Location_Desc,TSPL_LOCATION_MASTER.Loc_Short_Name , " &
              " TSPL_LOCATION_MASTER.Pin_Code AS Loc_Pin," &
              " (case when isnull(TSPL_LOCATION_MASTER.Phone1,'')<>'' then TSPL_LOCATION_MASTER.Phone1  when  isnull(TSPL_LOCATION_MASTER.Phone2,'')<>'' then + ', '+ TSPL_LOCATION_MASTER.Phone2 end) as Loc_Phone ," &
              " TSPL_LOCATION_MASTER.Email as Loc_Eamil,'' as Loc_Website," &
              " TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,convert(varchar(12),TSPL_SD_SALE_INVOICE_HEAD.Document_date,103) as Invoice_Date," &
              " TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No as DO_No,convert(varchar(12),TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_Date,103) as Do_Date," &
                "customer_city_master.city_name as Cust_City,TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No ,CUSTOMER_STATE_MASTER.GST_STATE_Code AS Cust_Gst_StateCode, TSPL_SD_SALE_INVOICE_HEAD.Electronic_Ref_No, Tspl_customer_master.gstno as CustGSTNo,TSPL_STATE_MASTER.gst_state_code,tspl_location_master.gstno as LocGstNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.Remarks as InvRemarks,TSPL_SD_sale_invoice_DETAIL.Delivery_Code , ITEMDETAIL.Conversion_factor, " &
                "case when coalesce(ITEMDETAIL.Conversion_factor,0)=0 then null else 0 end as QTY_Box,TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Sale_Invoice_No, " &
                "Case When ISNULL(TSPL_SD_SHIPMENT_HEAD.ManualVehicle,'')<>'' Then TSPL_SD_SHIPMENT_HEAD.ManualVehicle WHEN " &
                "ISNULL(TSPL_SD_SHIPMENT_HEAD.AlternateVehicle,'')<>'' Then TSPL_VEHICLE_MASTER.Number Else TSPL_SD_SALE_INVOICE_HEAD.vehicleNo End as vehicleNo, " &
                "Convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Sale_Invoice_Date,103) as Sale_Invoice_Date,  TSPL_SD_SALE_INVOICE_HEAD.RoundOffAmount, " &
                "TSPL_LOCATION_MASTER.Add1  as Loc_ADd1,TSPL_LOCATION_MASTER.Add2  as LOC_ADD2,TSPL_LOCATION_MASTER.Add3 as LOC_ADD3, TSPL_STATE_MASTER.State_Name as LocationState, case when ISNULL(TSPL_LOCATION_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_LOCATION_MASTER.Phone1 end +  Case When   ISNULL(TSPL_LOCATION_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_LOCATION_MASTER.Phone2 Else'' End as LOCPhone,TSPL_LOCATION_MASTER.TIN_No as Loc_TIN_NO, " &
                "TSPL_SD_SALE_INVOICE_HEAD.Document_Code ,convert(varchar,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Document_Date ,TSPL_SD_SALE_INVOICE_HEAD.Description, " &
                "TSPL_SD_SALE_INVOICE_HEAD.Lorry_No, TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_sale_invoice_DETAIL.Item_Code,0 as Line_No ,TSPL_ITEM_MASTER.Item_Desc as Particulars, " &
                "TSPL_SD_sale_invoice_DETAIL.Crate as QtyCrates ,TSPL_SD_sale_invoice_DETAIL.Unit_code ,0 as QtyPCS  , " &
                "coalesce( case when TSPL_SD_sale_invoice_DETAIL.Sampling=1 then 0 else SUB_QTY end,0) as free_qty, " &
                "TSPL_SD_sale_invoice_DETAIL.Scheme_Item as FreeSchemeInLitres ,0 as RatePerPcs,0  as valueInRs,0 as Cash_Scheme_Amount,isnull(schemeInCrates,0)as schemeInCrates, " &
                "'' GrandTotalCrates , TSPL_COMPANY_MASTER.Comp_Code ,TSPL_COMPANY_MASTER.Comp_Name ,TSPL_COMPANY_MASTER.Add1 as comp_add1 ,TSPL_COMPANY_MASTER.Add2 as  comp_add2 ,TSPL_COMPANY_MASTER.Add3 as comp_add3 ,TSPL_COMPANY_MASTER.Fax as comp_Fax ,TSPL_COMPANY_MASTER.Email as comp_Email, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone , TSPL_COMPANY_MASTER.Tin_No as comp_tinNo , " &
                "TSPL_SD_SALE_INVOICE_HEAD. Customer_Code  as cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Add1 as cust_add1 ,TSPL_CUSTOMER_MASTER. Add2 as cust_add2 ,TSPL_CUSTOMER_MASTER.Add3 cust_add3,case when ISNULL(TSPL_CUSTOMER_MASTER.Phone1,'')='(+__)__________' then ''  else TSPL_CUSTOMER_MASTER.Phone1 end +  Case When ISNULL(TSPL_CUSTOMER_MASTER.Phone2,'')<>'(+__)__________' Then ', '+  TSPL_CUSTOMER_MASTER.Phone2 Else'' End as CustPhone ,TSPL_CUSTOMER_MASTER.Fax as cust_fax ,TSPL_CUSTOMER_MASTER.State as Cust_state,CUSTOMER_STATE_MASTER.STATE_NAME as cust_Statename,TSPL_CUSTOMER_MASTER.Email as  cust_Email,TSPL_CUSTOMER_MASTER.WebSite as cust_website,TSPL_CUSTOMER_MASTER.pan as Customer_Pan  " &
                "from TSPL_SD_sale_invoice_DETAIL  " &
                "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No " &
                "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left join " &
                "(select Conversion_factor,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code='BOX' ) as ITEMDETAIL on " &
                "ITEMDETAIL.Item_code=TSPL_SD_sale_invoice_DETAIL.Item_Code  left outer join TSPL_COMPANY_MASTER on " &
                "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN " &
                "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State " &
               " left join TSPL_STATE_MASTER as CUSTOMER_STATE_MASTER on TSPL_CUSTOMER_MASTER.State= CUSTOMER_STATE_MASTER.STATE_CODE " &
                " left outer join TSPL_CITY_MASTER as customer_city_master on TSPL_CUSTOMER_MASTER.city_code=customer_city_master.City_Code " &
                " Full join " &
                "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from TSPL_SD_sale_invoice_DETAIL as inn " &
                "where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code  LEFT OUTER JOIN TSPL_VEHICLE_MASTER on " &
                "TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_SD_SHIPMENT_HEAD.AlternateVehicle " &
                " left outer join TSPL_ITEM_PRICE_MASTER on TSPL_ITEM_PRICE_MASTER.Price_Code=TSPL_SD_sale_invoice_DETAIL.Price_code " &
                " and TSPL_ITEM_PRICE_MASTER.Location_Code=TSPL_SD_sale_invoice_DETAIL.Location " &
                " and TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code   left outer join TSPL_DELIVERY_NOTE_master_FRESHSALE on TSPL_DELIVERY_NOTE_master_FRESHSALE.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Delivery_Code " &
                " left join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No=TSPL_SD_SALE_INVOICE_HEAD.document_code where 2=2 and  TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')   " &
                "and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='y'    and not exists (select 1 from  (select  TSPL_SD_sale_invoice_DETAIL.Item_Code, " &
                "TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE,TSPL_SD_sale_invoice_DETAIL.Line_No    from TSPL_SD_sale_invoice_DETAIL " &
                "LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE " &
                "left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code and " &
                "TSPL_ITEM_UOM_DETAIL .UOM_Code=TSPL_SD_sale_invoice_DETAIL.Unit_code LEFT OUTER JOIN TSPL_ITEM_MASTER  ON " &
                "TSPL_ITEM_MASTER.Item_Code =TSPL_SD_sale_invoice_DETAIL.Item_Code left outer join TSPL_COMPANY_MASTER on " &
                "TSPL_COMPANY_MASTER.Comp_Code =TSPL_SD_SALE_INVOICE_HEAD.Comp_Code  left outer join TSPL_CUSTOMER_MASTER  on " &
                "TSPL_CUSTOMER_MASTER .Cust_Code  =TSPL_SD_SALE_INVOICE_HEAD .Customer_Code left outer join TSPL_LOCATION_MASTER  on " &
                "TSPL_LOCATION_MASTER.Location_Code =TSPL_SD_SALE_INVOICE_HEAD .Bill_To_Location LEFT OUTER JOIN  " &
                "TSPL_STATE_MASTER On TSPL_STATE_MASTER.State_Code=TSPL_LOCATION_MASTER.State Full join " &
                "(select DOCUMENT_CODE,Item_Code as Scheme_Item_Code,SUM(Qty ) AS SUB_QTY,SUM(Crate) AS schemeInCrates from " &
                "TSPL_SD_sale_invoice_DETAIL as inn where DOCUMENT_CODE in ('" + strinvoiceNo + "') and inn.Scheme_Item='Y'   group by DOCUMENT_CODE,Item_Code) " &
                "TSPL_SD_sale_invoice_DETAIL_Sub on TSPL_SD_sale_invoice_DETAIL_sub.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE and " &
                "TSPL_SD_sale_invoice_DETAIL_sub.Scheme_Item_Code=TSPL_SD_sale_invoice_DETAIL.Item_Code " &
                " where 2=2 and " &
                "TSPL_SD_SALE_INVOICE_HEAD.Document_Code in   ('" + strinvoiceNo + "')  and TSPL_SD_sale_invoice_DETAIL.Scheme_Item='N')xx " &
                "where xx.Item_Code=TSPL_SD_sale_invoice_DETAIL.item_CODE and xx.DOCUMENT_CODE=TSPL_SD_sale_invoice_DETAIL.DOCUMENT_CODE ) " &
                ") as final left outer join " &
                " ( select Item_Code,max([CATEGORY RM]) as [CATEGORY RM],max([BRAND]) as [BRAND],max([SUB BRAND]) as [SUB BRAND],max([DESCRP]) as [DESCRP],max([PACK]) as [PACK]," &
                " max([PACK SIZE]) as [PACK SIZE],max([CATEGORY OT]) as [CATEGORY OT],max([CATEGORY FA]) as [CATEGORY FA],max([P TYPE]) as [P TYPE],max([L TYPE]) as [L TYPE],max([JW]) as [JW]," &
                " max([SCRAP]) as [SCRAP],max([CATEGORY RMDESC]) as [CATEGORY RMDESC],max([BRANDDESC]) as [BRANDDESC],max([SUB BRANDDESC]) as [SUB BRANDDESC],max([DESCRPDESC]) as [DESCRPDESC],max([PACKDESC]) as [PACKDESC]," &
                " max([PACK SIZEDESC]) as [PACK SIZEDESC],max([CATEGORY OTDESC]) as [CATEGORY OTDESC],max([CATEGORY FADESC]) as [CATEGORY FADESC],max([P TYPEDESC]) as [P TYPEDESC],max([L TYPEDESC]) as [L TYPEDESC],max([JWDESC]) as [JWDESC]," &
                " max([SCRAPDESC]) as [SCRAPDESC]  from ( select * from (   select TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code   ,TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code+'DESC' as Item_Category_CodeDesc   ," &
                " TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values  ,TSPL_ITEM_CATEGORY_LEVEL_VALUES.DESCRIPTION as Category_Value_Desc  from  TSPL_ITEM_MASTER    left outer join TSPL_ITEM_MASTER_CATEGORY on  TSPL_ITEM_MASTER_CATEGORY.Item_code = TSPL_ITEM_MASTER.Item_code " &
                " left outer join TSPL_ITEM_CATEGORY_LEVEL_VALUES on TSPL_ITEM_CATEGORY_LEVEL_VALUES.ITEM_CATEGORY_CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Category_Code and TSPL_ITEM_CATEGORY_LEVEL_VALUES.CODE=TSPL_ITEM_MASTER_CATEGORY.Item_Cagetory_Values   where 2=2 )xx " &
                " Pivot   ( max(Item_Cagetory_Values) for Item_Category_Code   in ( [CATEGORY RM],[BRAND],[SUB BRAND],[DESCRP],[PACK],[PACK SIZE],[CATEGORY OT],[CATEGORY FA],[P TYPE],[L TYPE],[JW],[SCRAP])  ) Pivt   Pivot  ( max(Category_Value_Desc) for Item_Category_CodeDesc in ([CATEGORY RMDESC]," &
                " [BRANDDESC],[SUB BRANDDESC],[DESCRPDESC],[PACKDESC],[PACK SIZEDESC],[CATEGORY OTDESC],[CATEGORY FADESC],[P TYPEDESC],[L TYPEDESC],[JWDESC],[SCRAPDESC])  ) Pivt1 ) xxx  group by Item_Code )  as tbl_Brand on tbl_Brand.Item_Code=final.item_Code  ) AS Main_Final left outer join TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.comp_code=Main_Final.comp_code " &
                    " order  by Main_Final.Sku_Seq ,Main_Final.Line_No"
        Return Qry
    End Function

    '========================================
    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub

    'Private Sub chkLocationAll_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkLocationAll.ToggleStateChanged
    '    cbgLocation.Enabled = Not chkLocationAll.IsChecked
    'End Sub

    Private Sub FrmPrintFreshInvoice_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            loadReport()
        ElseIf e.Alt And e.KeyCode = Keys.C Then
            Me.Close()
        ElseIf e.Alt And e.KeyCode = Keys.R Then
            Reset()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P Then
            loadData()
        End If
    End Sub

    Private Sub btnUnSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUnSelect.Click
        If clsCommon.CompairString(btnUnSelect.Text, "UnSelect All") = CompairStringResult.Equal Then
            For Each grow As GridViewRowInfo In gv.MasterView.Rows
                grow.Cells(0).Value = False

            Next
            btnUnSelect.Text = "Select All"
        Else
            For Each grow As GridViewRowInfo In gv.MasterView.Rows
                grow.Cells(0).Value = True
            Next
            btnUnSelect.Text = "UnSelect All"
        End If
    End Sub


    Private Sub txtCustomerMult__My_Click(sender As Object, e As EventArgs) Handles txtCustomerMult._My_Click
        Dim qry As String = "select Cust_Code as Code,Customer_Name as Name from TSPL_CUSTOMER_master order by Cust_Code"
        txtCustomerMult.arrValueMember = clsCommon.ShowMultipleSelectForm("Cust", qry, "Code", "Name", txtCustomerMult.arrValueMember, txtCustomerMult.arrDispalyMember)
    End Sub

    Private Sub txtCustomerCategoryMult__My_Click(sender As Object, e As EventArgs) Handles txtCustomerCategoryMult._My_Click
        Dim qry As String = "select CUST_CATEGORY_CODE as Code,CUST_CATEGORY_DESC as Name from TSPL_CUSTOMER_CATEGORY_MASTER order by CUST_CATEGORY_CODE"
        txtCustomerCategoryMult.arrValueMember = clsCommon.ShowMultipleSelectForm("Custcategory", qry, "Code", "Name", txtCustomerCategoryMult.arrValueMember, txtCustomerCategoryMult.arrDispalyMember)
    End Sub

    Private Sub txtLocation1Mult__My_Click(sender As Object, e As EventArgs) Handles txtLocation1Mult._My_Click
        Dim qry As String = "select Location_Code as [Code] ,Location_Desc as [Name] from TSPL_LOCATION_MASTER  "
        txtLocation1Mult.arrValueMember = clsCommon.ShowMultipleSelectForm("Cust", qry, "Code", "Name", txtLocation1Mult.arrValueMember, txtLocation1Mult.arrDispalyMember)
    End Sub
    Private Sub TxtMultiCustomerGroup__My_Click(sender As Object, e As EventArgs) Handles TxtMultiCustomerGroup._My_Click
        Dim qry As String = "select Cust_Group_Code as [Code],Cust_Group_Desc as [Name] from TSPL_CUSTOMER_GROUP_MASTER"
        TxtMultiCustomerGroup.arrValueMember = clsCommon.ShowMultipleSelectForm("CustGroup", qry, "Code", "Name", TxtMultiCustomerGroup.arrValueMember, TxtMultiCustomerGroup.arrDispalyMember)
    End Sub

    Private Sub rmPDF_Click(sender As Object, e As EventArgs) Handles rmPDF.Click
        Try

            Dim arrHeader As List(Of String) = New List(Of String)()

            arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")) + " ")
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.FrmPrintFreshInvoice & "'"))

            If txtLocation1Mult.arrValueMember IsNot Nothing AndAlso txtLocation1Mult.arrValueMember.Count > 0 Then
                Dim strLocationName As String = clsCommon.GetMulcallStringWithComma(txtLocation1Mult.arrValueMember)
                arrHeader.Add((" Location : " + strLocationName + " "))
            Else
                arrHeader.Add((" Location : All"))
            End If
            If txtCustomerMult.arrValueMember IsNot Nothing AndAlso txtCustomerMult.arrValueMember.Count > 0 Then
                arrHeader.Add(" Customer : " + clsCommon.GetMulcallStringWithComma(txtCustomerMult.arrDispalyMember))
            End If
            If txtCustomerCategoryMult.arrValueMember IsNot Nothing AndAlso txtCustomerCategoryMult.arrValueMember.Count > 0 Then
                arrHeader.Add(" Customer Category : " + clsCommon.GetMulcallStringWithComma(txtCustomerCategoryMult.arrDispalyMember))
            End If
            transportSql.applyExportTemplate(gv, PageSetupReport_ID)
            clsCommon.MyExportToPDF(Me.Text, gv, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class
