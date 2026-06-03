'---------------SALEINVOICESTATUSREPORT----------------------------------
Imports common
Imports System.IO


Public Class rptSaleInvoiceStatusReport
    Inherits FrmMainTranScreen

    Dim EnableProductSaleForJPR As Boolean = False
    Dim isPrint As Boolean = False

    Private Sub rptSaleInvoiceStatusReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtfDate.Value = clsCommon.GETSERVERDATE()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtItem.Visible = False
        MyLabel4.Visible = False
        EnableProductSaleForJPR = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)) = 1, True, False)
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        reset()
    End Sub

    Private Sub reset()
        gvdata.DataSource = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
        EnableDisableCntrl(True)
    End Sub

    Sub EnableDisableCntrl(ByVal val As Boolean)
        txtfDate.Enabled = val
        txtToDate.Enabled = val
        txtMultiCustomer.Enabled = val
        TxtCustomerType.Enabled = val
        txtItem.Enabled = val
        TxtSubLocation.Enabled = val
        TxtTransaction.Enabled = val
        RadGroupBox2.Enabled = val
        RadGroupBox1.Enabled = val
        RadGroupBox5.Enabled = val
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gvdata.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gvdata.Columns.Count - 1 Step ii + 1
                        gvdata.Columns(ii).IsVisible = False
                        gvdata.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gvdata.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Private Sub rmSaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        If clsCommon.myLen(PageSetupReport_ID) > 0 Then
            gvdata.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = PageSetupReport_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gvdata.SaveLayout(obj.GridLayout)
            obj.GridColumns = gvdata.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow("Layout saved successfully", "Information")
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(PageSetupReport_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow("Layout Delete successfully", "Information")
    End Sub

    Private Sub txtItem__My_Click(sender As Object, e As EventArgs) Handles txtItem._My_Click
        Dim qry As String = " select Item_Code as [Code], Item_Desc as [Name] from TSPL_ITEM_MASTER "
        txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", txtItem.arrValueMember, txtItem.arrDispalyMember)

    End Sub

    Private Sub txtMultiCustomer__My_Click(sender As Object, e As EventArgs) Handles txtMultiCustomer._My_Click
        Dim qry As String = " select cust_code as [Code], Customer_Name as [Name] from tspl_customer_master  "
        If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
            qry += " Where Cust_Type_Code In (" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")"
        End If
        txtMultiCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", txtMultiCustomer.arrValueMember, txtMultiCustomer.arrDispalyMember)
    End Sub

    Private Sub TxtCustomerType__My_Click(sender As Object, e As EventArgs) Handles TxtCustomerType._My_Click
        Dim qry As String = " SELECT  [Cust_Type_Code] as Code,[Cust_Type_Desc] as Name FROM [TSPL_CUSTOMER_TYPE_MASTER] Where 2=2  "
        TxtCustomerType.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", txtMultiCustomer.arrValueMember, txtMultiCustomer.arrDispalyMember)
        txtMultiCustomer.arrValueMember = Nothing
    End Sub

    Private Sub TxtTransaction__My_Click(sender As Object, e As EventArgs) Handles TxtTransaction._My_Click
        Dim qry As String = ""
        If rdbSaleReturn.IsChecked Then
            qry = " Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.saleReturn + "'
union all
Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.frmMCCMaterialSaleReturn + "'
union all
Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.ScrapSaleRetrun + "' "
        Else
            qry = " Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.frmDairyBookingCustomer + "'
                              union all
                              Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.frmSaleDispatchDairy + "'
                              union all
                              Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.FrmSalesOrderDispatch + "'
                              union all
                              Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.frmMCCMaterial + "'"
            If EnableProductSaleForJPR Then
                qry += "          Union all 
                              Select coalesce(Re_Name,Program_Name)Code,Program_Code as Name from TSPL_PROGRAM_MASTER where Program_Code='" + clsUserMgtCode.FrmProductDispatch + "'"
            End If
        End If
        TxtTransaction.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", TxtTransaction.arrValueMember, TxtTransaction.arrDispalyMember)
    End Sub

    Private Sub TxtSubLocation__My_Click(sender As Object, e As EventArgs) Handles TxtSubLocation._My_Click
        Dim qry As String = " Select Location_Code AS Code,Location_Desc as Name,Is_Sub_Location from TSPL_LOCATION_MASTER where Is_Sub_Location='Y' "
        TxtSubLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", TxtSubLocation.arrValueMember, TxtSubLocation.arrDispalyMember)
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        If rdbItemWiseCustomer.IsChecked Then
            LoadDataItemWiseCustomer()
        ElseIf rdbInvoiceCount.IsChecked Then
            LoadDataInvoiceCount()
        ElseIf rdbProductSummary.IsChecked Then
            LoadProductSummary()
        Else
            LoadData()
        End If
    End Sub
    Sub LoadProductSummary()
        Try
            Dim qry As String = ""
            If rbtnDetail.IsChecked Then
                qry = " Select ([Sub Location])[Sub Location] ,"
            Else
                qry = " Select "
            End If
            qry += " max([Item Name])as [Product Name],max(HSN_Code)[HSN Code],max(Report_UOM)[Report UOM],Cast(SUM(ReportUOM_Qty) as decimal (18,2))[Report Qty], sum([ItemBasic Amt])[ItemBasic Amt],sum([Margin Amt])[Margin Amt],
sum([Basic Amt])[Basic Amt],sum([KKF Amt])[KKF Amt],Cast(sum([Mandi Tax Amt]) as decimal(18,2))[Mandi Tax Amt],Cast(sum([Party TCS Amt]) as decimal(18,2))[TCS Amt],cast(sum([CGST Amt]) as decimal(18,2))[CGST Amt],cast(sum([SGST Amt]) as decimal(18,2))[SGST Amt],cast(sum([IGST Amt]) as decimal(18,2))[IGST Amt],sum([Total Tax Amt])[Total Tax Amt],
sum([Total Amt])[Total Amt],cast(Sum([Subsidy Amt]) as decimal(18,2))[Subsidy Amt] from (	select    
						TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location AS [Location],
TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code AS [Sub Location],
						TSPL_SD_SALE_INVOICE_DETAIL.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
  tspl_item_master.HSN_Code,
       		  Report_UOM.UOM_Code as Report_UOM,
(Billing_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty,
     TSPL_SD_SALE_INVOICE_DETAIL.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount as [Basic Amt]
 , case when TSPL_SD_SALE_INVOICE_DETAIL.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate else  TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate + TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate else TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate +TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate end)end as [GST Rate]
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0  END) AS [KKF Amt],
					CASE WHEN tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN tax2.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END AS [Mandi Tax Amt],
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
					 TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt as [Subsidy Amt]
                         from TSPL_SD_SALE_INVOICE_HEAD
left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.Document_Code
left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_INVOICE_DETAIL.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_INVOICE_DETAIL.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
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
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX10 
where convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=Convert( Date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103) 
and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)<=Convert( Date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) 
--and TSPL_ITEM_MASTER.TypeOfItm='P' 

 union all

select TSPL_SCRAPINVOICE_HEAD.Loc_Code AS [Location],
TSPL_SCRAPINVOICE_HEAD.Sub_Location_code AS [Sub Location],
   TSPL_SCRAPINVOICE_Detail.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
  tspl_item_master.HSN_Code,
  Report_UOM.UOM_Code as Report_UOM,
(shipped_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty,
 TSPL_SCRAPINVOICE_HEAD.Discount_Base as[ItemBasic Amt],
    TSPL_SCRAPINVOICE_HEAD.Discount_Amt as[Margin Amt],
  TSPL_SCRAPINVOICE_HEAD.Amount_Less_Discount as [Basic Amt] ,
  case when TSPL_SCRAPINVOICE_Detail.Tax1='KKF' or TSPL_SCRAPINVOICE_Detail.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_Detail.tax3='IGST' then TSPL_SCRAPINVOICE_Detail.TAX3_Rate else  TSPL_SCRAPINVOICE_Detail.TAX3_Rate + TSPL_SCRAPINVOICE_Detail.TAX4_Rate end ) else (case when  TSPL_SCRAPINVOICE_Detail.tax1='IGST' then TSPL_SCRAPINVOICE_Detail.TAX1_Rate else TSPL_SCRAPINVOICE_Detail.TAX1_Rate +TSPL_SCRAPINVOICE_Detail.TAX2_Rate end)end as [GST Rate]
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
					CASE WHEN tax1.Type='M'  THEN TSPL_SCRAPINVOICE_Detail.TAX1_Amt
    				WHEN tax2.Type='M'  THEN TSPL_SCRAPINVOICE_Detail.TAX2_Amt
    				WHEN tax3.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX3_Amt
    				WHEN tax4.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX4_Amt
    				WHEN tax5.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX5_Amt
    				WHEN tax6.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX6_Amt
    				WHEN tax7.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX7_Amt
    				WHEN tax8.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX8_Amt
    				WHEN tax9.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX9_Amt
    				WHEN tax10.Type='M' THEN TSPL_SCRAPINVOICE_Detail.TAX10_Amt else 0 END AS [Mandi Tax Amt],
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
					TSPL_SCRAPINVOICE_HEAD.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SCRAPINVOICE_Detail.ItemNetAmt as [Total Amt],0 as [Subsidy Amt]
                           from TSPL_SCRAPINVOICE_HEAD
                    left join TSPL_SCRAPINVOICE_Detail on TSPL_SCRAPINVOICE_Detail.invoice_No=TSPL_SCRAPINVOICE_HEAD.invoice_No
                    LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SCRAPINVOICE_Detail.Item_Code
                    left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SCRAPINVOICE_Detail.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SCRAPINVOICE_Detail.Unit_code
					left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SCRAPINVOICE_Detail.item_code and Report_UOM.Report_UOM=1
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
                    where convert(date,shipment_Date,103)>=Convert( Date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103) 
                    and convert(date,shipment_Date,103)<=Convert( Date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)
)XX "

            If rbtnDetail.IsChecked Then
                If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                    qry += "  where xx.[Sub Location]  in (" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")"
                End If
                qry += "   Group by XX.[Sub Location],XX.[Item Code]    "
            ElseIf rbtnSummary.IsChecked Then
                qry += "   Group by XX.[Item Code]    "
            End If
            Dim dt As DataTable = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry)
            gvdata.DataSource = Nothing
            gvdata.Rows.Clear()
            gvdata.Columns.Clear()
            gvdata.GroupDescriptors.Clear()
            gvdata.MasterView.Refresh()
            Dim viewBlank As New TableViewDefinition()
            gvdata.ViewDefinition = viewBlank
            If dt.Rows.Count > 0 Then
                gvdata.DataSource = dt
                gvdata.GroupDescriptors.Clear()
                gvdata.EnableFiltering = True
                gvdata.MasterTemplate.SummaryRowsBottom.Clear()
                SetGridFormationProductSummary()
                EnableDisableCntrl(False)
                gvdata.MasterTemplate.AutoExpandGroups = True
                RadPageView1.SelectedPage = RadPageViewPage3
                gvdata.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function BaseQryLoadDataInvoiceCount(ByVal fromDate As String, ByVal toDate As String, ByVal strLocation As String) As String
        Dim Qry As String = " Select *,(Total_Invoice-Total_CancelInvoice-Total_DeleteInvoice)as Active_invoice 
from(Select LEFT(XX.Document_Code, CHARINDEX('/', XX.Document_Code) - 1) As Prefix,Transcation_Type, Invoice_Tax_Type, MIN(XX.Document_Code) First_Invoice, MAX(XX.Document_Code)Last_Invoice, count(Distinct XX.Document_Code)Total_Invoice,
COUNT(Distinct XX.Cancel_DocumentCode) As Total_CancelInvoice,COUNT(Distinct XX.Delete_DocumentCode) As Total_DeleteInvoice
from(Select   CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
        Left Join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          And TSPL_SD_SHIPMENT_HEAD.Against_Booking_No Is Not NULL)  And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC'  THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
        Left Join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          And TSPL_SD_SHIPMENT_HEAD.Against_Booking_No Is NULL And ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          And TSPL_SD_SHIPMENT_HEAD.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
        Left Join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          And TSPL_SD_SHIPMENT_HEAD.Against_Booking_No Is NULL And ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          And TSPL_SD_SHIPMENT_HEAD.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
		  Left Join TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
          And TSPL_SD_SHIPMENT_HEAD.Against_Booking_No Is NULL And (TSPL_BOOKING_MATSER.Is_APS=1 Or TSPL_SD_SHIPMENT_HEAD.Screen_Type= ('CT'))) And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' THEN 'APS SALES'
WHEN TSPL_SD_SALE_INVOICE_HEAD.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'		  
Else 'DCS SALE' END AS Transcation_Type,   
CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'     
	 WHEN TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable = 1 THEN 'TAXABLE (LOCAL)'     
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable,TSPL_SD_SALE_INVOICE_HEAD.Document_Code,NULL as Cancel_DocumentCode,Null as Delete_DocumentCode
	from TSPL_SD_SALE_INVOICE_HEAD  
Left Outer Join ( Select xyz.* from (
Select Document_Code,TAX1 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax1
Where Tax1.Type='IGST'
Union All
Select Document_Code,TAX2 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax2
Where Tax2.Type='IGST'
Union All
Select Document_Code,TAX3 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax3
Where Tax3.Type='IGST'
Union All
Select Document_Code,TAX4 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax4
Where Tax4.Type='IGST'
Union All
Select Document_Code,TAX5 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax5
Where Tax5.Type='IGST'
Union All
Select Document_Code,TAX6 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax6
Where Tax6.Type='IGST'
Union All
Select Document_Code,TAX7 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax7
Where Tax7.Type='IGST'
Union All
Select Document_Code,TAX8 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax8
Where Tax8.Type='IGST'
Union All
Select Document_Code,TAX9 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.Tax9
Where Tax9.Type='IGST'
Union All
Select Document_Code,TAX10 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX10 Where Tax10.Type='IGST' ) xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Document_Code
WHERE CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location = '" & clsCommon.myCstr(strLocation) & "'"
        End If

        Qry &= " Union all 

  Select  CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NOT NULL)  And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC'  THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Cancel_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Cancel_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
		  LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER_Cancel_Data.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Screen_Type= ('CT'))) And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'APS SALES'
	WHEN TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'	  
  ELSE 'DCS SALE' END AS Transcation_Type ,
 CASE WHEN TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'     
	 WHEN TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Is_Taxable,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code as Document_Code,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code as Cancel_DocumentCode,NULL as Delete_DocumentCode from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join ( Select xyz.* from (
Select Document_Code,TAX1 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax1 Where Tax1.Type='IGST'
Union All
Select Document_Code,TAX2 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax2 Where Tax2.Type='IGST'
Union All
Select Document_Code,TAX3 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax3 Where Tax3.Type='IGST'
Union All
Select Document_Code,TAX4 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax4 Where Tax4.Type='IGST'
Union All
Select Document_Code,TAX5 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax5 Where Tax5.Type='IGST'
Union All
Select Document_Code,TAX6 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax6 Where Tax6.Type='IGST'
Union All
Select Document_Code,TAX7 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax7 Where Tax7.Type='IGST'
Union All
Select Document_Code,TAX8 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax8 Where Tax8.Type='IGST'
Union All
Select Document_Code,TAX9 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Tax9 Where Tax9.Type='IGST'
Union All
Select Document_Code,TAX10 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.TAX10 Where Tax10.Type='IGST' ) xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code
WHERE CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Bill_To_Location = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " Union all

  Select  CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Delete_Data ON TSPL_BOOKING_MATSER_Delete_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NOT NULL)  And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC'  THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Delete_Data ON TSPL_BOOKING_MATSER_Delete_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Delete_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Delete_Data ON TSPL_BOOKING_MATSER_Delete_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Delete_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
		  LEFT JOIN TSPL_BOOKING_MATSER_Delete_Data ON TSPL_BOOKING_MATSER_Delete_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER_Delete_Data.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD_Delete_Data.Screen_Type= ('CT'))) And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'APS SALES'
	WHEN TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'	  
  ELSE 'DCS SALE' END AS Transcation_Type ,
CASE WHEN TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
	 WHEN TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Is_Taxable,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code AS Document_Code,NULL as Cancel_DocumentCode,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code as Delete_DocumentCode from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
 
Left Outer Join ( Select xyz.* from (
Select Document_Code,TAX1 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax1 Where Tax1.Type='IGST'
Union All
Select Document_Code,TAX2 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax2 Where Tax2.Type='IGST'
Union All
Select Document_Code,TAX3 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax3 Where Tax3.Type='IGST'
Union All
Select Document_Code,TAX4 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax4 Where Tax4.Type='IGST'
Union All
Select Document_Code,TAX5 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax5 Where Tax5.Type='IGST'
Union All
Select Document_Code,TAX6 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax6 Where Tax6.Type='IGST'
Union All
Select Document_Code,TAX7 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax7 Where Tax7.Type='IGST'
Union All
Select Document_Code,TAX8 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax8 Where Tax8.Type='IGST'
Union All
Select Document_Code,TAX9 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Tax9 Where Tax9.Type='IGST'
Union All
Select Document_Code,TAX10 As Tax,Type from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.TAX10 Where Tax10.Type='IGST') xyz ---Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code
WHERE CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Bill_To_Location = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
  Select 'MATERIAL SALE'  AS Transcation_Type ,
		CASE WHEN TSPL_SCRAPINVOICE_HEAD.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'     
	         WHEN TSPL_SCRAPINVOICE_HEAD.Is_Taxable = 1 THEN 'TAXABLE (LOCAL)'     
        ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SCRAPINVOICE_HEAD.Invoice_Type,TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,NULL as Cancel_DocumentCode,NULL AS Delete_DocumentCode from TSPL_SCRAPINVOICE_HEAD
 Left Outer Join ( Select xyz.* from (
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX1 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax1 Where Tax1.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX2 As Tax,Type from TSPL_SCRAPINVOICE_HEAD 
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax2 Where Tax2.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX3 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax3 Where Tax3.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX4 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax4 Where Tax4.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX5 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax5 Where Tax5.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX6 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax6 Where Tax6.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX7 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax7 Where Tax7.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX8 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax8 Where Tax8.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX9 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SCRAPINVOICE_HEAD.Tax9 Where Tax9.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD.invoice_No AS Document_Code,TAX10 As Tax,Type from TSPL_SCRAPINVOICE_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SCRAPINVOICE_HEAD.TAX10 Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SCRAPINVOICE_HEAD.invoice_No
WHERE CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SCRAPINVOICE_HEAD.Loc_Code = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
  Select 'MATERIAL SALE'  AS Transcation_Type ,
		  CASE WHEN TSPL_SCRAPINVOICE_HEAD_cancel_data.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
	 WHEN TSPL_SCRAPINVOICE_HEAD_cancel_data.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SCRAPINVOICE_HEAD_cancel_data.Invoice_Type,TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No as Cancel_DocumentCode,NULL AS Delete_DocumentCode from TSPL_SCRAPINVOICE_HEAD_cancel_data
 	Left Outer Join ( Select xyz.* from (
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX1 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax1 Where Tax1.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX2 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax2 Where Tax2.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX3 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax3 Where Tax3.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX4 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax4 Where Tax4.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX5 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax5 Where Tax5.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX6 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax6 Where Tax6.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX7 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax7 Where Tax7.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX8 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax8 Where Tax8.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX9 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.Tax9 Where Tax9.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No AS Document_Code,TAX10 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_cancel_data
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.TAX10 Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SCRAPINVOICE_HEAD_cancel_data.invoice_No
WHERE CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD_cancel_data.shipment_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD_cancel_data.shipment_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103)"
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SCRAPINVOICE_HEAD_cancel_data.Loc_Code = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
  Select 'MATERIAL SALE'  AS Transcation_Type ,
		  CASE WHEN TSPL_SCRAPINVOICE_HEAD_Delete_data.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
	 WHEN TSPL_SCRAPINVOICE_HEAD_Delete_data.Is_Taxable = 1   THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SCRAPINVOICE_HEAD_Delete_data.Invoice_Type,TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,NULL as Cancel_DocumentCode,TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Delete_DocumentCode from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join ( Select xyz.* from (
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX1 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax1 Where Tax1.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX2 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax2 Where Tax2.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX3 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax3 Where Tax3.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX4 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax4 Where Tax4.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX5 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax5 Where Tax5.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX6 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax6 Where Tax6.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX7 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax7 Where Tax7.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX8 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax8 Where Tax8.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX9 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.Tax9 Where Tax9.Type='IGST'
Union All
Select TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No AS Document_Code,TAX10 As Tax,Type from TSPL_SCRAPINVOICE_HEAD_Delete_data
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.TAX10 Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SCRAPINVOICE_HEAD_Delete_data.invoice_No
WHERE CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD_Delete_data.shipment_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD_Delete_data.shipment_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SCRAPINVOICE_HEAD_Delete_data.Loc_Code = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
  Select case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' then 'DCS SALE RETURN' else 'SALES RETURN' end  AS Transcation_Type ,
		CASE WHEN TSPL_SD_SALE_RETURN_HEAD.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
	 WHEN TSPL_SD_SALE_RETURN_HEAD.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SD_SALE_RETURN_HEAD.Is_Taxable,TSPL_SD_SALE_RETURN_HEAD.Document_Code AS Document_Code,NULL as Cancel_DocumentCode,NULL AS Delete_DocumentCode from TSPL_SD_SALE_RETURN_HEAD
 Left Outer Join ( Select xyz.* from (
Select Document_Code,TAX1 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax1 Where Tax1.Type='IGST'
Union All
Select Document_Code,TAX2 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax2 Where Tax2.Type='IGST'
Union All
Select Document_Code,TAX3 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax3 Where Tax3.Type='IGST'
Union All
Select Document_Code,TAX4 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax4 Where Tax4.Type='IGST'
Union All
Select Document_Code,TAX5 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax5 Where Tax5.Type='IGST'
Union All
Select Document_Code,TAX6 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax6 Where Tax6.Type='IGST'
Union All
Select Document_Code,TAX7 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax7 Where Tax7.Type='IGST'
Union All
Select Document_Code,TAX8 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax8 Where Tax8.Type='IGST'
Union All
Select Document_Code,TAX9 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.Tax9 Where Tax9.Type='IGST'
Union All
Select Document_Code,TAX10 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SD_SALE_RETURN_HEAD.TAX10  Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SD_SALE_RETURN_HEAD.Document_Code
WHERE CONVERT(DATE,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
  Select case when TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Trans_Type='MCC' then 'DCS SALE RETURN' else 'SALES RETURN' end  AS Transcation_Type ,
		  CASE WHEN TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
	 WHEN TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Is_Taxable,TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Document_Code AS Document_Code,TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Document_Code as Cancel_DocumentCode,NULL AS Delete_DocumentCode from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
 Left Outer Join ( Select xyz.* from (
Select Document_Code,TAX1 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax1 Where Tax1.Type='IGST'
Union All
Select Document_Code,TAX2 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax2 Where Tax2.Type='IGST'
Union All
Select Document_Code,TAX3 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax3 Where Tax3.Type='IGST'
Union All
Select Document_Code,TAX4 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax4 Where Tax4.Type='IGST'
Union All
Select Document_Code,TAX5 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax5 Where Tax5.Type='IGST'
Union All
Select Document_Code,TAX6 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax6 Where Tax6.Type='IGST'
Union All
Select Document_Code,TAX7 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA 
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax7 Where Tax7.Type='IGST'
Union All 
Select Document_Code,TAX8 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax8 Where Tax8.Type='IGST'
Union All
Select Document_Code,TAX9 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Tax9 Where Tax9.Type='IGST'
Union All
Select Document_Code,TAX10 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.TAX10 Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Document_Code
WHERE CONVERT(DATE,TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103)"
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SD_SALE_RETURN_HEAD_CANCEL_DATA.Bill_To_Location = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
  Select case when TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Trans_Type='MCC' then 'DCS SALE RETURN' else 'SALES RETURN' end  AS Transcation_Type ,
		  CASE WHEN TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'     
	 WHEN TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'     
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Is_Taxable,TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Document_Code AS Document_Code,NULL as Cancel_DocumentCode,TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Document_Code AS Delete_DocumentCode from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
 Left Outer Join ( Select xyz.* from (
Select Document_Code,TAX1 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax1 Where Tax1.Type='IGST'
Union All
Select Document_Code,TAX2 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax2 Where Tax2.Type='IGST'
Union All
Select Document_Code,TAX3 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax3 Where Tax3.Type='IGST'
Union All
Select Document_Code,TAX4 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax4 Where Tax4.Type='IGST'
Union All
Select Document_Code,TAX5 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax5 Where Tax5.Type='IGST'
Union All
Select Document_Code,TAX6 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA 
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax6 Where Tax6.Type='IGST'
Union All
Select Document_Code,TAX7 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax7 Where Tax7.Type='IGST'
Union All
Select Document_Code,TAX8 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax8 Where Tax8.Type='IGST'
Union All
Select Document_Code,TAX9 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Tax9 Where Tax9.Type='IGST'
Union All
Select Document_Code,TAX10 As Tax,Type from TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.TAX10  Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Document_Code
WHERE CONVERT(DATE,TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103)"
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SD_SALE_RETURN_HEAD_DELETE_DATA.Bill_To_Location = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " union all
    Select 'MATERIAL SALE RETURN'   AS Transcation_Type ,
		 CASE WHEN TSPL_SCRAPSALE_HEAD_RETURN.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'     
	 WHEN TSPL_SCRAPSALE_HEAD_RETURN.Is_Taxable = 1  THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SCRAPSALE_HEAD_RETURN.Is_Taxable,TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,NULL as Cancel_DocumentCode,NULL AS Delete_DocumentCode from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join ( Select xyz.* from (
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX1 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax1 Where Tax1.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX2 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax2 Where Tax2.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX3 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax3 Where Tax3.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX4 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax4 Where Tax4.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX5 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax5 Where Tax5.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX6 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax6 Where Tax6.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX7 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax7 Where Tax7.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX8 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax8 Where Tax8.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX9 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.Tax9 Where Tax9.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN.Document_No AS Document_Code,TAX10 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN.TAX10 Where Tax10.Type='IGST') xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SCRAPSALE_HEAD_RETURN.Document_No
WHERE CONVERT(DATE,TSPL_SCRAPSALE_HEAD_RETURN.Return_ship_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SCRAPSALE_HEAD_RETURN.Return_ship_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103)"
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SCRAPSALE_HEAD_RETURN.Loc_Code = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " UNION ALL
  Select 'MATERIAL SALE RETURN'   AS Transcation_Type ,
		  CASE WHEN TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
	 WHEN TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Is_Taxable = 1 THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Is_Taxable,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No AS Document_Code,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Cancel_DocumentCode,NULL AS Delete_DocumentCode from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
 Left Outer Join ( Select xyz.* from (
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX1 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax1 Where Tax1.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX2 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax2 Where Tax2.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX3 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax3 Where Tax3.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX4 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax4 Where Tax4.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX5 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax5 Where Tax5.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX6 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax6 Where Tax6.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX7 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax7 Where Tax7.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX8 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax8 Where Tax8.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX9 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Tax9 Where Tax9.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No as Document_Code,TAX10 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.TAX10 Where Tax10.Type='IGST' ) xyz --Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Document_No
WHERE CONVERT(DATE,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Return_ship_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Return_ship_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103)"
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SCRAPSALE_HEAD_RETURN_Cancel_Data.Loc_Code = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " UNION ALL
    Select 'MATERIAL SALE RETURN'   AS Transcation_Type ,
		  CASE WHEN TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Is_Taxable = 1 And Tax_Master.Type ='IGST' THEN 'TAXABLE (IGST)'
     	 WHEN TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Is_Taxable = 1 THEN 'TAXABLE (LOCAL)'
     ELSE 'NON TAXABLE (LOCAL)' END AS Invoice_Tax_Type,
	TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Is_Taxable,TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,NULL as Cancel_DocumentCode,TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Delete_DocumentCode from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
  Left Outer Join ( Select xyz.* from (
Select  TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX1 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax1 On Tax1.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax1 Where Tax1.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX2 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax2 On Tax2.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax2 Where Tax2.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX3 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax3 On Tax3.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax3 Where Tax3.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX4 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax4 On Tax4.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax4 Where Tax4.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX5 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax5 On Tax5.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax5 Where Tax5.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX6 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax6 On Tax6.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax6 Where Tax6.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX7 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax7 On Tax7.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax7 Where Tax7.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX8 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax8 On Tax8.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax8 Where Tax8.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX9 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax9 On Tax9.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Tax9 Where Tax9.Type='IGST'
Union All
Select TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No AS Document_Code,TAX10 As Tax,Type from TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data
Left Outer Join TSPL_TAX_MASTER As Tax10 On Tax10.Tax_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.TAX10 Where Tax10.Type='IGST' ) xyz ---Where Xyz.Type='IGST'
) Tax_Master On Tax_Master.Document_Code=TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Document_No
 WHERE CONVERT(DATE,TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Return_ship_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(fromDate, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Return_ship_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(toDate, "dd/MMM/yyyy") & "',103) "
        If clsCommon.myLen(strLocation) > 0 Then
            Qry &= " And TSPL_SCRAPSALE_HEAD_RETURN_Delete_Data.Loc_Code = '" & clsCommon.myCstr(strLocation) & "'"
        End If
        Qry &= " ) XX group by LEFT(XX.Document_Code, CHARINDEX('/', XX.Document_Code) - 1),XX.Transcation_Type,XX.Invoice_Tax_Type )YY "
        Return Qry
    End Function
    Sub LoadDataInvoiceCount()
        Try
            Dim qry As String = BaseQryLoadDataInvoiceCount(txtfDate.Value, txtToDate.Value, Nothing)
            Dim dt As DataTable = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry & " order by YY.Transcation_Type ")
            gvdata.DataSource = Nothing
            gvdata.Rows.Clear()
            gvdata.Columns.Clear()
            gvdata.GroupDescriptors.Clear()
            gvdata.MasterView.Refresh()
            Dim viewBlank As New TableViewDefinition()
            gvdata.ViewDefinition = viewBlank
            If dt.Rows.Count > 0 Then
                gvdata.DataSource = dt
                gvdata.GroupDescriptors.Clear()
                gvdata.EnableFiltering = True
                gvdata.MasterTemplate.SummaryRowsBottom.Clear()
                SetGridFormationInvoiceCount()
                EnableDisableCntrl(False)
                gvdata.MasterTemplate.AutoExpandGroups = True
                RadPageView1.SelectedPage = RadPageViewPage3
                gvdata.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        ReStoreGridLayout()
    End Sub
    Sub LoadDataItemWiseCustomer()
        Try
            Dim itemsWithZeroUOM As New List(Of String)
            Dim qry1 As String = Nothing
            Dim qry2 As String = "SELECT D.Item_Code,max(IM.Item_Desc)Item_Desc FROM TSPL_SD_SALE_INVOICE_DETAIL D
LEFT JOIN TSPL_SD_SALE_INVOICE_HEAD H ON D.DOCUMENT_CODE = H.Document_Code
LEFT JOIN TSPL_ITEM_MASTER IM ON IM.Item_Code = D.Item_Code
LEFT JOIN TSPL_ITEM_UOM_DETAIL U ON U.Item_Code = D.Item_Code
WHERE CONVERT(date, H.Document_Date, 103)  BETWEEN CONVERT(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103) AND CONVERT(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103)
GROUP BY D.Item_Code HAVING SUM(CASE WHEN U.Report_UOM = 1 THEN 1 ELSE 0 END) = 0 "
            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry2)
            If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                Dim itemList As New List(Of String)
                For Each dr As DataRow In dt2.Rows
                    itemList.Add(dr("Item_Code").ToString())
                Next
                Dim errorMessage As String = "Report UOM not define for Item Code(s): " & String.Join(", ", itemList.ToArray())
                clsCommon.MyMessageBoxShow(Me, errorMessage, Me.Text)
                Exit Sub
            End If
            Dim SumItemCode As String = Nothing
            Dim ItemCode As String = Nothing
            Dim ItemDesc As String = Nothing
            Dim itemqry As String = "  Select distinct Item_Code,item_Desc,Report_UOM,UOM_Code FROM  ( Select  TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,isnull(TSPL_ITEM_UOM_DETAIL.Report_UOM,0)Report_UOM,TSPL_ITEM_UOM_DETAIL.UOM_Code
                                        from TSPL_SD_SALE_INVOICE_DETAIL 
                                        Left outer join  TSPL_SD_SALE_INVOICE_HEAD ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
		                                Left outer join  TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
		                                Left outer join  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code   
		 	                            where 2=2  and CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
                                        AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) and TSPL_ITEM_UOM_DETAIL.Report_UOM = 1 "
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                itemqry += "  and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
            End If
            itemqry += "  Union all 
                            Select  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code,TSPL_ITEM_MASTER.Item_Desc,isnull(TSPL_ITEM_UOM_DETAIL.Report_UOM,0)Report_UOM,TSPL_ITEM_UOM_DETAIL.UOM_Code
                                        from TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data 
                                        Left outer join  TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data ON TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code
		                                Left outer join  TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code
		                                Left outer join  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code   
		 	                            where 2=2  and CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
                                        AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) and TSPL_ITEM_UOM_DETAIL.Report_UOM = 1 "
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                itemqry += "  and TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Customer_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
            End If
            itemqry += " Union all
                                        Select  TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code,TSPL_ITEM_MASTER.Item_Desc,isnull(TSPL_ITEM_UOM_DETAIL.Report_UOM,0)Report_UOM,TSPL_ITEM_UOM_DETAIL.UOM_Code
                                        from TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data 
                                        Left outer join  TSPL_SD_SALE_INVOICE_HEAD_Delete_Data ON TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code
		                                Left outer join  TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code
		                                Left outer join  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code   
		 	                            where 2=2  and CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
                                        AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) and TSPL_ITEM_UOM_DETAIL.Report_UOM = 1 "
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                itemqry += "  and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Customer_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
            End If

            itemqry += " Union all
Select  TSPL_SCRAPINVOICE_Detail.Item_Code,TSPL_ITEM_MASTER.Item_Desc,isnull(TSPL_ITEM_UOM_DETAIL.Report_UOM,0)Report_UOM,TSPL_ITEM_UOM_DETAIL.UOM_Code
                                        from TSPL_SCRAPINVOICE_Detail 
                                        Left outer join  TSPL_SCRAPINVOICE_HEAD ON TSPL_SCRAPINVOICE_HEAD.invoice_No = TSPL_SCRAPINVOICE_Detail.invoice_No
		                                Left outer join  TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code = TSPL_SCRAPINVOICE_Detail.Item_Code
		                                Left outer join  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SCRAPINVOICE_Detail.Item_Code   
		 	                            where 2=2  and CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
                                        AND CONVERT(DATE,TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) and TSPL_ITEM_UOM_DETAIL.Report_UOM = 1 "
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                itemqry += "  and TSPL_SCRAPINVOICE_HEAD.cust_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
            End If
            itemqry += "  )XX order by Item_Desc "
            Dim dtitem As DataTable = clsDBFuncationality.GetDataTable(itemqry)
            If dtitem.Rows.Count > 0 Then
                For i As Integer = 0 To dtitem.Rows.Count - 1
                    If i = 0 Then
                        SumItemCode += " Sum(IsNull([" + clsCommon.myCstr(dtitem.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + " Qty (" + clsCommon.myCstr(dtitem.Rows(i)("UOM_Code")) + ")],Sum(IsNull([" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + "],0)) As [" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + " Amt]"
                        ItemCode += "[" + clsCommon.myCstr(dtitem.Rows(i)("Item_Code")) + "] "
                        ItemDesc += "[" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + "]"
                    Else
                        SumItemCode += ", Sum(IsNull([" + clsCommon.myCstr(dtitem.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + " Qty (" + clsCommon.myCstr(dtitem.Rows(i)("UOM_Code")) + ")],Sum(IsNull([" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + "],0)) As [" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + " Amt]"
                        ItemCode += ", [" + clsCommon.myCstr(dtitem.Rows(i)("Item_Code")) + "] "
                        ItemDesc += ", [" + clsCommon.myCstr(dtitem.Rows(i)("Item_Desc")) + "]"
                    End If
                Next
            End If
            Dim whrcls As String = Nothing
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                whrcls += "  and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                whrcls += "  and TSPL_SD_SALE_INVOICE_DETAIL.Item_code in (" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")"
            End If
            Dim qry As String = ""
            If dtitem.Rows.Count > 0 Then

                qry = "   With CTE_DATA AS ( 
SELECT 
TSPL_SD_SALE_INVOICE_HEAD.Document_Code AS Document_Code,CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD 
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
    RIGHT(TSPL_SD_SALE_INVOICE_HEAD.Document_Code, 6) AS BillNo,
    (CONVERT(varchar(20), TSPL_SD_SALE_INVOICE_HEAD.Document_Date, 103)) AS BillDate,
     (TSPL_CUSTOMER_MASTER.Cust_Code) as Party_Code,
    (TSPL_CUSTOMER_MASTER.Customer_Name) AS PARTY_Name,
	(TSPL_CUSTOMER_MASTER.Cust_Type_Code) as Customer_Type,
Case when (TSPL_SD_SALE_INVOICE_HEAD.Status)= 1 then 'POSTED' else 'UNPOSTED' end as Status,
(TSPL_SD_SALE_INVOICE_DETAIL.Item_Code)Item_Code,
(TSPL_ITEM_MASTER.Item_Desc)Item_Code1,
(CAST((TSPL_SD_SALE_INVOICE_DETAIL.Qty * IMCF) / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS DECIMAL(18,2))) as QtyUom,
    TSPL_SD_SALE_INVOICE_DETAIL.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL.disc_Amt as[Margin Amt],
TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount,
TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount as Amt_Less_Discount1,
Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0  END) AS [KKF],
					CASE WHEN tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN tax2.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt  else 0 END) AS [Party TCS Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END) AS [CGST Amt],
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
					
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt  else 0 END )AS [IGST Amt],
					 TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_INVOICE_Detail.Item_Net_Amt as [Total Amt],TSPL_SD_SALE_INVOICE_HEAD.Created_By,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Created_Date,103) as Created_Date
FROM TSPL_SD_SALE_INVOICE_HEAD
LEFT JOIN TSPL_SD_SALE_INVOICE_DETAIL
       ON TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD.Document_Code
LEFT JOIN TSPL_ITEM_MASTER 
       ON TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
LEFT JOIN TSPL_ITEM_UOM_DETAIL
       ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM = 1  
LEFT JOIN TSPL_CUSTOMER_MASTER
       ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
	  left JOIN (
        SELECT Item_Code, UOM_Code, Conversion_Factor as IMCF
        FROM TSPL_ITEM_UOM_DETAIL
    ) AS ItemCF 
         ON ItemCF.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code 
        AND ItemCF.UOM_Code = TSPL_SD_SALE_INVOICE_DETAIL.Billing_Unit_code
  left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX10 
WHERE CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
                If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
                End If
                If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_SD_SALE_INVOICE_DETAIL.Item_code in (" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")"
                End If
                qry += " Union all
                      
  SELECT 
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code AS Document_Code,CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_cancel_data 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_cancel_data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No IS NOT NULL)  And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC'  THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_cancel_data 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_cancel_data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_cancel_data.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_cancel_data 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_cancel_data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_cancel_data.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_cancel_data 
		  LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_cancel_data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_cancel_data.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD_cancel_data.Screen_Type= ('CT'))) And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'APS SALES'
        WHEN TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'	  
ELSE 'DCS SALE' END AS Transcation_Type,
    RIGHT(TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code, 6) AS BillNo,
    (CONVERT(varchar(20), TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date, 103)) AS BillDate,
     (TSPL_CUSTOMER_MASTER.Cust_Code) as Party_Code,
    (TSPL_CUSTOMER_MASTER.Customer_Name) AS PARTY_Name,
	(TSPL_CUSTOMER_MASTER.Cust_Type_Code) as Customer_Type,
'CANCEL'  as Status,
(TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code)Item_Code,
(TSPL_ITEM_MASTER.Item_Desc)Item_Code1,
(CAST((TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Qty * IMCF) / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS DECIMAL(18,2))) as QtyUom,
    TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.disc_Amt as[Margin Amt],
TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Amt_Less_Discount,
TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Amt_Less_Discount as Amt_Less_Discount1,
Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE WHEN tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN tax2.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt  else 0 END) AS [Party TCS Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END) AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt  else 0 END )AS [IGST Amt],
					 TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Net_Amt as [Total Amt],TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Cancel_By as Created_By,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Cancel_On,103) as Created_Date
FROM TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
LEFT JOIN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data
       ON TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code
LEFT JOIN TSPL_ITEM_MASTER 
       ON TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code
LEFT JOIN TSPL_ITEM_UOM_DETAIL
       ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM = 1  
LEFT JOIN TSPL_CUSTOMER_MASTER
       ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Customer_Code
	  left JOIN (
        SELECT Item_Code, UOM_Code, Conversion_Factor as IMCF
        FROM TSPL_ITEM_UOM_DETAIL
    ) AS ItemCF 
         ON ItemCF.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code 
        AND ItemCF.UOM_Code = TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Billing_Unit_code
  left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX10 
WHERE CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
                If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
                End If
                If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_code in (" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")"
                End If

                qry += " union all
                     SELECT 
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code AS Document_Code,CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NOT NULL)  And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC'  THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
        LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Delete_Data 
		  LEFT JOIN TSPL_BOOKING_MATSER ON TSPL_BOOKING_MATSER.Document_No = TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Delete_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Delete_Data.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD_Delete_Data.Screen_Type= ('CT'))) And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'APS SALES'
        WHEN TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'	  
ELSE 'DCS SALE' END AS Transcation_Type,
    RIGHT(TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code, 6) AS BillNo,
    (CONVERT(varchar(20), TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date, 103)) AS BillDate,
     (TSPL_CUSTOMER_MASTER.Cust_Code) as Party_Code,
    (TSPL_CUSTOMER_MASTER.Customer_Name) AS PARTY_Name,
	(TSPL_CUSTOMER_MASTER.Cust_Type_Code) as Customer_Type,
'DELETE'  as Status,
(TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code)Item_Code,
(TSPL_ITEM_MASTER.Item_Desc)Item_Code1,
(CAST((TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Qty * IMCF) / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS DECIMAL(18,2))) as QtyUom,
    TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.disc_Amt as[Margin Amt],
TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Amt_Less_Discount,
TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Amt_Less_Discount as Amt_Less_Discount1,
Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE WHEN tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN tax2.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt  else 0 END) AS [Party TCS Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END) AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt  else 0 END )AS [IGST Amt],
					 TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Net_Amt as [Total Amt],TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_By as Created_By,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_On,103) as Created_Date
FROM TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
LEFT JOIN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data
       ON TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.DOCUMENT_CODE = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code
LEFT JOIN TSPL_ITEM_MASTER 
       ON TSPL_ITEM_MASTER.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code
LEFT JOIN TSPL_ITEM_UOM_DETAIL
       ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM = 1  
LEFT JOIN TSPL_CUSTOMER_MASTER
       ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Customer_Code
	  left JOIN (
        SELECT Item_Code, UOM_Code, Conversion_Factor as IMCF
        FROM TSPL_ITEM_UOM_DETAIL
    ) AS ItemCF 
         ON ItemCF.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code 
        AND ItemCF.UOM_Code = TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Billing_Unit_code
  left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX10 
WHERE CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) >= convert(date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103)
  AND CONVERT(DATE,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) <=convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
                If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
                End If
                If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_code in (" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")"
                End If

                qry += "   ----MATERIAL SALE ------"
                qry += " 
              union all 
               select  TSPL_SCRAPINVOICE_HEAD.invoice_No,'MATERIAL SALE' as Transcation_Type,
   RIGHT(TSPL_SCRAPINVOICE_HEAD.invoice_No, 6) AS BillNo,
    (CONVERT(varchar(20), TSPL_SCRAPINVOICE_HEAD.shipment_Date, 103)) AS BillDate,
     (TSPL_CUSTOMER_MASTER.Cust_Code) as Party_Code,
    (TSPL_CUSTOMER_MASTER.Customer_Name) AS PARTY_Name,
	(TSPL_CUSTOMER_MASTER.Cust_Type_Code) as Customer_Type,
Case when (TSPL_SCRAPINVOICE_HEAD.ispost)= 1 then 'POSTED' else 'UNPOSTED' end as Status,
(TSPL_SCRAPINVOICE_Detail.Item_Code)Item_Code   ,
(TSPL_ITEM_MASTER.Item_Desc)Item_Code1,
(CAST((TSPL_SCRAPINVOICE_Detail.shipped_Qty * IMCF) / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS DECIMAL(18,2))) as QtyUom,
           TSPL_SCRAPINVOICE_HEAD.Discount_Base as[ItemBasic Amt],
    TSPL_SCRAPINVOICE_HEAD.Discount_Amt as[Margin Amt],
  TSPL_SCRAPINVOICE_HEAD.Amount_Less_Discount as Amt_Less_Discount,
  TSPL_SCRAPINVOICE_HEAD.Amount_Less_Discount as Amt_Less_Discount1

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
					TSPL_SCRAPINVOICE_HEAD.Created_By,Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD.Created_Date,103) as Created_Date
                           from TSPL_SCRAPINVOICE_HEAD
                    left join TSPL_SCRAPINVOICE_Detail on TSPL_SCRAPINVOICE_Detail.invoice_No=TSPL_SCRAPINVOICE_HEAD.invoice_No
                    LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SCRAPINVOICE_Detail.Item_Code
                    left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SCRAPINVOICE_Detail.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SCRAPINVOICE_Detail.Unit_code
					 left JOIN (
        SELECT Item_Code, UOM_Code, Conversion_Factor as IMCF
        FROM TSPL_ITEM_UOM_DETAIL
    ) AS ItemCF 
         ON ItemCF.Item_Code = TSPL_SCRAPINVOICE_Detail.Item_Code 
        AND ItemCF.UOM_Code = TSPL_SCRAPINVOICE_Detail.Unit_code
					left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SCRAPINVOICE_Detail.item_code and Report_UOM.Report_UOM=1
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
                    where convert(date,shipment_Date,103)>=Convert( Date,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") & "',103) and convert(date,shipment_Date,103)<=Convert( Date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "

                If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")"
                End If
                If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                    qry += "  and TSPL_SCRAPINVOICE_Detail.Item_code in (" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")"
                End If

                qry += " ) "

                qry += " Select "
                If rbtnDetail.IsChecked Then
                    qry += "  Document_Code,max(Transcation_Type)[Transcation Type],max(BillNo)BillNo,max(BillDate)BillDate,max(Created_By)[Created By],max(Created_Date)[Created Date],max(Party_Code)Party_Code,max(PARTY_Name)PARTY_Name "
                ElseIf rbtnSummary.IsChecked Then
                    qry += " Party_Code,max(PARTY_Name)PARTY_Name "
                End If
                qry += "  , max(ISNULL((Customer_Type),''))Customer_Type,max(Status)Status,sum([ItemBasic Amt])[ItemBasic Amt],
sum([Margin Amt])[Margin Amt], " & SumItemCode & ",sum(Amt_Less_Discount1)[Total Basic Amt],sum(KKF)[KKF Amt],SUM([Mandi Tax Amt])[Mandi Tax Amt],sum([Party TCS Amt])[Party TCS Amt],sum([CGST Amt])[CGST Amt],sum([SGST Amt])[SGST Amt],sum([IGST Amt])[IGST Amt],sum([Total Tax Amt])[Total Tax Amt],
Sum([Total Amt])[Total Amt]
from ( SELECT *
    FROM CTE_DATA
) XX "
                qry += " PIVOT (SUM(QtyUom)  For Item_Code In (" & ItemCode & ") ) As pivot_Code
  PIVOT (SUM(Amt_Less_Discount)  For Item_Code1 In (" & ItemDesc & ") ) As pivot_Desc "
                If rbtnDetail.IsChecked Then
                    qry += "  GROUP BY Document_Code order by max(Transcation_Type) "
                ElseIf rbtnSummary.IsChecked Then
                    qry += " GROUP BY Party_Code order by max(PARTY_Name)   "
                End If
                Dim dt As DataTable = New DataTable()
                dt = clsDBFuncationality.GetDataTable(qry)
                gvdata.DataSource = Nothing
                gvdata.Rows.Clear()
                gvdata.Columns.Clear()
                gvdata.GroupDescriptors.Clear()
                'gvData.MasterView.Refresh()
                'Dim viewBlank As New TableViewDefinition()
                'gvData.ViewDefinition = viewBlank
                If dt.Rows.Count > 0 Then

                    Try
                        gvdata.DataSource = dt
                        gvdata.GroupDescriptors.Clear()
                        gvdata.EnableFiltering = True
                        gvdata.MasterTemplate.SummaryRowsBottom.Clear()
                        SetGridFormation()
                        EnableDisableCntrl(False)
                        ' gvData.MasterTemplate.AutoExpandGroups = True
                        'RadPageView1.SelectedPage = RadPageViewPage2
                        RadPageView1.SelectedPage = RadPageViewPage3
                        'gvdata.BestFitColumns()

                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                    Exit Sub
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        ReStoreGridLayout()
    End Sub

    Sub SetGridFormationProductSummary()
        gvdata.AutoExpandGroups = True
        gvdata.ShowGroupPanel = True
        gvdata.ShowRowHeaderColumn = False
        gvdata.AllowAddNewRow = False
        gvdata.AllowDeleteRow = False
        gvdata.EnableFiltering = True
        gvdata.ShowFilteringRow = True
        For ii As Integer = 0 To gvdata.Columns.Count - 1
            gvdata.Columns(ii).ReadOnly = True
            gvdata.Columns(ii).BestFit()
        Next
        gvdata.Columns("ItemBasic Amt").IsVisible = False
        gvdata.Columns("ItemBasic Amt").VisibleInColumnChooser = True
        gvdata.Columns("Margin Amt").IsVisible = False
        gvdata.Columns("Margin Amt").VisibleInColumnChooser = True
    End Sub
    Sub SetGridFormationInvoiceCount()
        gvdata.AutoExpandGroups = True
        gvdata.ShowGroupPanel = True
        gvdata.ShowRowHeaderColumn = False
        gvdata.AllowAddNewRow = False
        gvdata.AllowDeleteRow = False
        gvdata.EnableFiltering = True
        gvdata.ShowFilteringRow = True
        For ii As Integer = 0 To gvdata.Columns.Count - 1
            gvdata.Columns(ii).ReadOnly = True
            gvdata.Columns(ii).BestFit()
        Next
        gvdata.Columns("Transcation_Type").HeaderText = "Transaction"
        gvdata.Columns("Invoice_Tax_Type").HeaderText = "Invoice Type"
        gvdata.Columns("First_Invoice").HeaderText = "First Invoice No"
        gvdata.Columns("Last_Invoice").HeaderText = "Last Invoice No"
        gvdata.Columns("Total_Invoice").HeaderText = "Total Invoice Issued"
        gvdata.Columns("Total_CancelInvoice").HeaderText = "No Of Invoice Cancel "
        gvdata.Columns("Total_DeleteInvoice").HeaderText = "No Of Invoice Delete"
        gvdata.Columns("Active_invoice").HeaderText = "Active Invoice"
        gvdata.ShowGroupPanel = True
        gvdata.MasterTemplate.AutoExpandGroups = True
    End Sub
    Sub SetGridFormation()
        gvdata.AutoExpandGroups = True
        gvdata.ShowGroupPanel = True
        gvdata.ShowRowHeaderColumn = False
        gvdata.AllowAddNewRow = False
        gvdata.AllowDeleteRow = False
        gvdata.EnableFiltering = True
        gvdata.ShowFilteringRow = True
        For ii As Integer = 0 To gvdata.Columns.Count - 1
            gvdata.Columns(ii).ReadOnly = True
            gvdata.Columns(ii).BestFit()
        Next
        If rbtnDetail.IsChecked Then
            gvdata.Columns("Document_Code").HeaderText = "Invoice No"
            gvdata.Columns("Document_Code").IsVisible = False
            gvdata.Columns("Document_Code").VisibleInColumnChooser = True
            gvdata.Columns("BillNo").HeaderText = "Bill No"
            gvdata.Columns("BillDate").HeaderText = "Bill Date"
        End If

        gvdata.Columns("ItemBasic Amt").IsVisible = False
        gvdata.Columns("ItemBasic Amt").VisibleInColumnChooser = True
        gvdata.Columns("Margin Amt").IsVisible = False
        gvdata.Columns("Margin Amt").VisibleInColumnChooser = True
        gvdata.Columns("PARTY_Name").HeaderText = "PARTY Name"
        gvdata.Columns("Party_Code").HeaderText = "PARTY Code"
        gvdata.Columns("Customer_Type").HeaderText = "Customer Type"

        gvdata.ShowGroupPanel = True
        gvdata.MasterTemplate.AutoExpandGroups = True
        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim index As Integer
        If rbtnDetail.IsChecked Then
            index = 10
        ElseIf rbtnSummary.IsChecked Then
            index = 4
        End If
        For ii As Integer = index To gvdata.Columns.Count - 1
            If clsCommon.CompairString(gvdata.Columns(ii).Name, "Created By") <> CompairStringResult.Equal OrElse clsCommon.CompairString(gvdata.Columns(ii).Name, "Created Date") <> CompairStringResult.Equal Then
                summaryRowItem.Add(New GridViewSummaryItem(gvdata.Columns(ii).Name, "{0:F2}", GridAggregateFunction.Sum))
            End If
            'summaryRowItem.Add(New GridViewSummaryItem(gvdata.Columns(ii).Name, "{0:F2}", GridAggregateFunction.Sum))
        Next
        gvdata.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gvdata.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Sub LoadData()
        Try
            Dim WhrCust As String = ""
            Dim Sublocn As String = ""
            Dim item As String = ""
            Dim dt As DataTable = Nothing
            Dim strtxtfDate As String = clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy")
            Dim strToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
            Dim qry As String = ""
            Dim Baseqry As String = ""
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
Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Invoice_Date,
TSPL_SD_SALE_INVOICE_HEAD.Document_Code as Invoice_No,
TSPL_SD_SALE_INVOICE_HEAD.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_INVOICE_DETAIL.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],TSPL_SD_SALE_INVOICE_DETAIL.item_cost As [Rate],
TSPL_SD_SALE_INVOICE_DETAIL.Billing_Unit_code as [Measure of Qty],
  TSPL_SD_SALE_INVOICE_DETAIL.Billing_Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SD_SALE_INVOICE_DETAIL.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate else  TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate + TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate else TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate +TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SD_SALE_INVOICE_DETAIL.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Base_Amt else  TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Base_Amt end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Base_Amt else TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Base_Amt end)end as [Basic Amt]
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
(Billing_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SHIPMENT_HEAD.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_INVOICE_HEAD.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_INVOICE_HEAD.Created_By,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Created_Date,103) as Created_Date,TSPL_SD_SALE_INVOICE_HEAD.Total_Amt As [Bill Amount]

                         from TSPL_SD_SALE_INVOICE_HEAD
left join TSPL_SD_SALE_INVOICE_DETAIL on TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD.Document_Code
left  join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_INVOICE_DETAIL.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_INVOICE_DETAIL.item_code and Report_UOM.Report_UOM=1
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
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD .TAX10 
where convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103)<=Convert( Date,'" + strToDate + "',103)"
            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                Baseqry += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                Baseqry += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SD_SALE_INVOICE_DETAIL.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            Baseqry += "            union all

select  'MATERIAL SALE' as Transcation_Type,case when TSPL_SCRAPINVOICE_HEAD.ispost=1 then 'Approved' else'Pending' end as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SCRAPINVOICE_HEAD.Loc_Code AS [Location],
TSPL_SCRAPINVOICE_HEAD.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD.shipment_Date,103) as Invoice_Date,
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
  tspl_item_master.HSN_Code,
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
					left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SCRAPINVOICE_Detail.item_code and Report_UOM.Report_UOM=1
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
            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                Baseqry += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                Baseqry += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SCRAPINVOICE_HEAD.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_SCRAPINVOICE_Detail.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            If rbtnDetail.IsChecked AndAlso rdbGstInvoice.IsChecked Then
                If isPrint Then
                    qry &= " Select xx.*,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.State,TSPL_STATE_MASTER.STATE_NAME,TSPL_COMPANY_MASTER.Pincode from ("
                    qry &= " Select ROW_NUMBER() OVER (PARTITION BY Invoice_No ORDER BY [Item Code]) AS SNo,Max(PrintBy)PrintBy,Max(FromDate)FromDate,Max(ToDate)ToDate,Max([Transcation Type])[Transcation Type],MAX([Customer Type])[Customer Type],Max([Doc Status])[Doc Status],Max(Location)Location,Max([Location GST])[Location GST],Max([Sub Location])[Sub Location],Invoice_Date,Invoice_No,Max([Invoice Type])[Invoice Type],Max([Route No])[Route No],([Party Code])[Party Code],Max([DCS Uploader])[DCS Uploader],Max([Party Name])[Party Name],
                    Max([Customer GSTNo])[Customer GSTNo],Max([Party State Code])[Party State Code],[Item Code],Max([Item Name])[Item Name],Max([HSN Code])[HSN Code],Max([Measure of Qty])[Measure of Qty],Sum([Product Qty])[Product Qty],Max([Report UOM])[Report UOM],Max(Rate)Rate,
					Sum(CAST([ReportUOM Qty] AS DECIMAL(18,2))) AS [ReportUOM Qty],
					Max(Cast(([GST Rate]) as decimal(10,2)))[GST Rate],Sum(cast(([ItemBasic Amt]) as decimal(10,2)))[ItemBasic Amt],
					Sum(cast(([Margin Amt]) as decimal(10,2)))[Margin Amt],Sum(Cast(([Basic Amt]) as Decimal(10,2)))[Basic Amt],Sum(Cast((KKF) as decimal(10,2)))KKF,
					Sum(Cast(([Mandi Tax Amt]) as Decimal(10,2)))[Mandi Tax Amt],Sum(Cast(([Party TCS Amt]) as Decimal(10,2)))[Party TCS Amt],
                        Sum(Cast(([CGST Amt]) as Decimal(10,2)))[CGST Amt],SUm(Cast(([SGST Amt]) as Decimal(10,2)))[SGST Amt],Sum(Cast(([IGST Amt]) as decimal(10,2)))[IGST Amt],Sum(cast(([Total Tax Amt]) as decimal(10,2)))[Total Tax Amt],Sum(cast(([Total Amt]) as decimal(10,2)))[Total Amt],Sum(([Subsidy Amt]))[Subsidy Amt],Max([B2B/B2C])[B2B/B2C],Max([Ack No])[Ack No],Max([Ack Date])[Ack Date],Max([IRN No])[IRN No],Max(EWayBillNo)EWayBillNo,Max(EWayBillDate)EWayBillDate
                    ,Max([Created By])[Created By],Max([Created Date])[Created Date],Max([Bill Amount])[Bill Amount] from ("
                End If
                qry &= "Select '" & objCommonVar.CurrentUser & "' As PrintBy,ROW_NUMBER() OVER (PARTITION BY xx.Invoice_No ORDER BY xx.[Item Code]) AS SNo,'" & clsCommon.GetPrintDate(txtfDate.Value, "dd-MM-yyyy") & "' As FromDate,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd-MM-yyyy") & "' As ToDate,(Transcation_Type)[Transcation Type],([Customer Type])[Customer Type],(Doc_Status)[Doc Status],(Location)Location,([GST No])[Location GST],([Sub Location])[Sub Location],(Invoice_Date)Invoice_Date,Invoice_No,([Invoice Type])[Invoice Type],([Route No])[Route No],([Party Code])[Party Code],VLC_Code_VLC_Uploader as [DCS Uploader],([Party Name])[Party Name],
                    ([Customer GSTNo])[Customer GSTNo],([State Code])[Party State Code],[Item Code],[Item Name],HSN_Code AS [HSN Code],([Measure of Qty])[Measure of Qty],([Product Qty])[Product Qty],(Report_UOM)[Report UOM],Rate,CAST(ReportUOM_Qty AS DECIMAL(18,2)) AS [ReportUOM Qty],Cast(([IGST Rate]) as decimal(10,2))[GST Rate],cast(([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast((KKF) as decimal(10,2))KKF,Cast(([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(([IGST Amt]) as decimal(10,2))[IGST Amt],cast(([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(([Total Amt]) as decimal(10,2))[Total Amt],([Subsidy Amt])[Subsidy Amt],([B2B/B2C])[B2B/B2C],(Ack_No)[Ack No],(Ack_Date)[Ack Date],(IRN_No)[IRN No],EWayBillNo,EWayBillDate,CASE WHEN ROW_NUMBER() OVER (PARTITION BY Invoice_No ORDER BY [Item Code] DESC) = 1 THEN [Bill Amount] Else 0 End As [Bill Amount]
                    ,(Created_By)[Created By],(Created_Date)[Created Date]
                    from (" & Baseqry & ")XX   "
                qry &= " Where 1=1 "
                If rbtnNonTaxable.IsChecked Then
                    qry &= " And XX.[Invoice Type]='Non-Taxable' "
                Else
                    qry &= " And XX.[Invoice Type]='Taxable' "
                End If
                If TxtTransaction.arrValueMember IsNot Nothing AndAlso TxtTransaction.arrValueMember.Count > 0 Then
                    qry += " And XX.Transcation_Type In(" + clsCommon.GetMulcallString(TxtTransaction.arrValueMember) + ")" + Environment.NewLine
                End If

                If isPrint Then
                    qry &= " )xyz "
                    qry &= " Group By [Party Code],Invoice_No,Invoice_Date,[Item Code] "
                    qry &= " )xx Left Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
					          Left Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State"
                    qry += " order by xx.[Transcation Type], xx.Invoice_Date "
                Else
                    qry += " order by xx.Transcation_Type, xx.Invoice_Date   "
                End If
            ElseIf rbtnSummary.IsChecked AndAlso rdbGstInvoice.IsChecked Then
                qry = " Select max(Transcation_Type)[Transcation Type],max([Customer Type])[Customer Type],max(Doc_Status)[Doc Status],max(Location)Location,max([GST No])[Location GST],max([Sub Location])[Sub Location],max(Invoice_Date)Invoice_Date,Invoice_No,max([Invoice Type])[Invoice Type],MAX([Route No])[Route No],max([Party Code])[Party Code],max(VLC_Code_VLC_Uploader) as [DCS Uploader],max([Party Name])[Party Name],
                        max([Customer GSTNo])[Customer GSTNo],max([State Code])[Party State Code],cast(sum([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(sum([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(sum([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast(sum(KKF) as decimal(10,2))KKF,Cast(sum([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(sum([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(sum([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(sum([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(sum([IGST Amt]) as decimal(10,2))[IGST Amt],cast(sum([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(sum([Total Amt]) as decimal(10,2))[Total Amt],sum([Subsidy Amt])[Subsidy Amt],max([B2B/B2C])[B2B/B2C],max(Ack_No)[Ack No],max(Ack_Date)[Ack Date],max(IRN_No)[IRN No],max(EWayBillNo)EWayBillNo,max(EWayBillDate)EWayBillDate
                        ,max(Created_By)[Created By],max(Created_Date)[Created Date]
                        from ( " + Baseqry + " )XX  "
                If TxtTransaction.arrValueMember IsNot Nothing AndAlso TxtTransaction.arrValueMember.Count > 0 Then
                    qry += " where XX.Transcation_Type In(" + clsCommon.GetMulcallString(TxtTransaction.arrValueMember) + ")" + Environment.NewLine
                End If
                qry += " Group by xx.Invoice_No "
                qry += " order by max(xx.Transcation_Type), max(xx.Invoice_Date) "
            End If

            Dim BaseQryCancel As String = ""
            BaseQryCancel = " 
select  CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NOT NULL) 
        And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Cancel_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Cancel_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
		  LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER_Cancel_Data.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Screen_Type= ('CT')))
         And TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Trans_Type <> 'MCC' THEN 'APS SALES'
		  WHEN TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'
ELSE 'DCS SALE' END AS Transcation_Type,
'Cancel' as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Bill_To_Location AS [Location],
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103) as Invoice_Date,
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code as Invoice_No,
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Billing_Unit_code as [Measure of Qty],
  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Billing_Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Rate else  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Rate + TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Rate end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Rate else TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Rate +TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt else  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt else TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt end)end as [Basic Amt]
  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Amt_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Net_Amt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Ack_No,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Ack_Date,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.IRN_No,
Report_UOM.UOM_Code as Report_UOM,
(Billing_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Created_By,
 Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Created_Date,103) as Created_Date,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Cancel_By as [Cancel By],
TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Cancel_On as [Cancel Date]
                              from TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data
left join TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data on TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Code
left  join TSPL_SD_SHIPMENT_HEAD_Cancel_Data on TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Against_Shipment_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Customer_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Customer_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data .TAX10 
where convert(date,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Document_Date,103)<=Convert( Date,'" + strToDate + "',103)"
            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            BaseQryCancel += "  union all
select  CASE WHEN EXISTS ( SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NOT NULL) 
        And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'CUSTOMER BOOKING'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Cancel_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Item_Type = 'S') And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
        LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND ISNULL(TSPL_BOOKING_MATSER_Cancel_Data.Is_APS,0) = 0
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Item_Type In ('P','I')) And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'PRODUCT DISPATCH'
		  WHEN EXISTS (SELECT 1 FROM TSPL_SD_SHIPMENT_HEAD_Cancel_Data 
		  LEFT JOIN TSPL_BOOKING_MATSER_Cancel_Data ON TSPL_BOOKING_MATSER_Cancel_Data.Document_No = TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No
        WHERE TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
          AND TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Against_Booking_No IS NULL AND (TSPL_BOOKING_MATSER_Cancel_Data.Is_APS=1 OR TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Screen_Type= ('CT')))
         And TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Trans_Type <> 'MCC' THEN 'APS SALES'
WHEN TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.IsMultipleInvoice = 1 THEN 'MULTIPLE INVOICE'
		  ELSE 'DCS SALE' END AS Transcation_Type,
'Delete' as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Bill_To_Location AS [Location],
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103) as Invoice_Date,
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code as Invoice_No,
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Billing_Unit_code as [Measure of Qty],
  TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Billing_Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Rate else  TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Rate + TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Rate end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Rate else TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Rate +TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt else  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt end ) else (case when  TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt else TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt end)end as [Basic Amt]
  TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Amt_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt 
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Net_Amt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Ack_No,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Ack_Date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.IRN_No,
Report_UOM.UOM_Code as Report_UOM,
(Billing_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SHIPMENT_HEAD_Cancel_Data.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Created_By,
 Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Created_Date ,103) as Created_Date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_By as [Cancel By],
TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Delete_On as [Cancel Date]

                              from TSPL_SD_SALE_INVOICE_HEAD_Delete_Data
left join TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data on TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Code
left  join TSPL_SD_SHIPMENT_HEAD_Cancel_Data on TSPL_SD_SHIPMENT_HEAD_Cancel_Data.Document_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Against_Shipment_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Customer_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Customer_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD_Delete_Data .TAX10 
where convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103)>=Convert( Date,'" + strtxtfDate + "',103)
and convert(date,TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Document_Date,103)<=Convert( Date,'" + strToDate + "',103) "

            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SD_SALE_INVOICE_HEAD_Delete_Data.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SD_SALE_INVOICE_DETAIL_Delete_Data.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            BaseQryCancel += " union all

select  'MATERIAL SALE' as Transcation_Type,'Cancel' as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Loc_Code AS [Location],
TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD_Cancel_Data.shipment_Date,103) as Invoice_Date,
TSPL_SCRAPINVOICE_HEAD_Cancel_Data.invoice_No as Invoice_No,
 TSPL_ROUTE_MASTER.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Unit_code as [Measure of Qty],
  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.shipped_Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Rate else  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Rate + TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Rate end ) else (case when  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Rate else TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Rate +TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt else  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt end ) else (case when  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt else TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt end)end as [Basic Amt]
   TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Discount_Base as[ItemBasic Amt],
    TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Discount_Amt as[Margin Amt],
  TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Amount_Less_Discount as [Basic Amt]
,Convert(decimal(18,2),CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10='KKF' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10='TCS' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10='CGST' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Rate
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10='SGST' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10='IGST' THEN TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TotalTaxAmt as [Total Tax Amt],
					TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TotalAmt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C]
,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Ack_No,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Ack_Date,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.IRN_No
,Report_UOM.UOM_Code as Report_UOM,
(shipped_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty,0 as [Subsidy Amt], TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Invoice_Type as [Invoice Type],
TSPL_SCRAPINVOICE_HEAD_Cancel_Data.EWayBillNo,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Created_By,
Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Created_Date  ,103) as Created_Date,TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Cancel_By as [Cancel By],
TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Cancel_On as [Cancel Date]

                         from TSPL_SCRAPINVOICE_HEAD_Cancel_Data
left join TSPL_SCRAPINVOICE_DETAIL_Cancel_Data on TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.invoice_No=TSPL_SCRAPINVOICE_HEAD_Cancel_Data.invoice_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Loc_Code
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SCRAPINVOICE_HEAD_Cancel_Data.cust_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SCRAPINVOICE_HEAD_Cancel_Data.cust_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SCRAPINVOICE_HEAD_Cancel_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SCRAPINVOICE_HEAD_Cancel_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Cancel_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SCRAPINVOICE_HEAD_Cancel_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Cancel_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Cancel_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Cancel_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Cancel_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Cancel_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Cancel_Data .TAX10 
where convert(date,shipment_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,shipment_Date,103)<=Convert( Date,'" + strToDate + "',103) "
            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SCRAPINVOICE_HEAD_Cancel_Data.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            BaseQryCancel += " union all

select  'MATERIAL SALE' as Transcation_Type,'Delete' as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SCRAPINVOICE_HEAD_Delete_Data.Loc_Code AS [Location],
TSPL_SCRAPINVOICE_HEAD_Delete_Data.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD_Delete_Data.shipment_Date,103) as Invoice_Date,
TSPL_SCRAPINVOICE_HEAD_Delete_Data.invoice_No as Invoice_No,
 TSPL_ROUTE_MASTER.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Unit_code as [Measure of Qty],
  TSPL_SCRAPINVOICE_DETAIL_Delete_Data.shipped_Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Tax1='KKF' or TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_DETAIL_Delete_Data.tax3='IGST' then TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Rate else  TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Rate + TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Rate end ) else (case when  TSPL_SCRAPINVOICE_DETAIL_Delete_Data.tax1='IGST' then TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Rate else TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Rate +TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt else  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX3_Base_Amt end ) else (case when  TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt else TSPL_SCRAPINVOICE_DETAIL_Cancel_Data.TAX1_Base_Amt end)end as [Basic Amt]
   TSPL_SCRAPINVOICE_HEAD_Delete_Data.Discount_Base as[ItemBasic Amt],
    TSPL_SCRAPINVOICE_HEAD_Delete_Data.Discount_Amt as[Margin Amt],
  TSPL_SCRAPINVOICE_HEAD_Delete_Data.Amount_Less_Discount as [Basic Amt]
,Convert(decimal(18,2),CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9='KKF'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9_Amt 
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10='KKF' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9='TCS'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10='TCS' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9='CGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10='CGST' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Rate
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9='SGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10='SGST' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9='IGST'  THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10='IGST' THEN TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TotalTaxAmt as [Total Tax Amt],
					TSPL_SCRAPINVOICE_DETAIL_Delete_Data.TotalAmt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C]
,TSPL_SCRAPINVOICE_HEAD_Delete_Data.Ack_No,TSPL_SCRAPINVOICE_HEAD_Delete_Data.Ack_Date,TSPL_SCRAPINVOICE_HEAD_Delete_Data.IRN_No
,Report_UOM.UOM_Code as Report_UOM,
(shipped_Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty,0 as [Subsidy Amt], TSPL_SCRAPINVOICE_HEAD_Delete_Data.Invoice_Type as [Invoice Type],
TSPL_SCRAPINVOICE_HEAD_Delete_Data.EWayBillNo,TSPL_SCRAPINVOICE_HEAD_Delete_Data.EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SCRAPINVOICE_HEAD_Delete_Data.Created_By,
 Convert(varchar(20),TSPL_SCRAPINVOICE_HEAD_Delete_Data.Created_Date  ,103) as Created_Date,TSPL_SCRAPINVOICE_HEAD_Delete_Data.Delete_By as [Cancel By],
TSPL_SCRAPINVOICE_HEAD_Delete_Data.Delete_On as [Cancel Date]

                         from TSPL_SCRAPINVOICE_HEAD_Delete_Data
left join TSPL_SCRAPINVOICE_DETAIL_Delete_Data on TSPL_SCRAPINVOICE_DETAIL_Delete_Data.invoice_No=TSPL_SCRAPINVOICE_HEAD_Delete_Data.invoice_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SCRAPINVOICE_DETAIL_Delete_Data.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SCRAPINVOICE_DETAIL_Delete_Data.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SCRAPINVOICE_HEAD_Delete_Data.Loc_Code
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SCRAPINVOICE_HEAD_Delete_Data.cust_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SCRAPINVOICE_HEAD_Delete_Data.cust_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SCRAPINVOICE_HEAD_Delete_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SCRAPINVOICE_HEAD_Delete_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SCRAPINVOICE_HEAD_Delete_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SCRAPINVOICE_HEAD_Delete_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Delete_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Delete_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Delete_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Delete_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SCRAPINVOICE_HEAD_Delete_Data .TAX10 
where convert(date,shipment_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,shipment_Date,103)<=Convert( Date,'" + strToDate + "',103)   "
            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                BaseQryCancel += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SCRAPINVOICE_HEAD_Delete_Data.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                BaseQryCancel += " and TSPL_SCRAPINVOICE_DETAIL_Delete_Data.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If
            If rbtnDetail.IsChecked AndAlso rdbCancelInvoice.IsChecked Then
                qry = "Select (Transcation_Type)[Transcation Type],([Customer Type])[Customer Type],(Doc_Status)[Doc Status],(Location)Location,([GST No])[Location GST],([Sub Location])[Sub Location],(Invoice_Date)Invoice_Date,Invoice_No,([Invoice Type])[Invoice Type],([Route No])[Route No],([Party Code])[Party Code],VLC_Code_VLC_Uploader as [DCS Uploader],([Party Name])[Party Name],
                       ([Customer GSTNo])[Customer GSTNo],([State Code])[Party State Code],[Item Code],[Item Name],HSN_Code AS [HSN Code],([Measure of Qty])[Measure of Qty],([Product Qty])[Product Qty],(Report_UOM)[Report UOM],CAST(ReportUOM_Qty AS DECIMAL(18,2)) AS [ReportUOM Qty],Cast(([IGST Rate]) as decimal(10,2))[GST Rate],cast(([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast((KKF) as decimal(10,2))KKF,Cast(([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(([IGST Amt]) as decimal(10,2))[IGST Amt],cast(([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(([Total Amt]) as decimal(10,2))[Total Amt],([Subsidy Amt])[Subsidy Amt],([B2B/B2C])[B2B/B2C],(Ack_No)[Ack No],(Ack_Date)[Ack Date],(IRN_No)[IRN No],EWayBillNo,EWayBillDate
                        ,(Created_By)[Created By],(Created_Date)[Created Date],[Cancel By],[Cancel Date] from (" + BaseQryCancel + ")XX   "
                If TxtTransaction.arrValueMember IsNot Nothing AndAlso TxtTransaction.arrValueMember.Count > 0 Then
                    qry += " where XX.Transcation_Type In(" + clsCommon.GetMulcallString(TxtTransaction.arrValueMember) + ")" + Environment.NewLine
                End If
                qry += " order by xx.Transcation_Type, xx.Invoice_Date   "
            ElseIf rbtnSummary.IsChecked AndAlso rdbCancelInvoice.IsChecked Then
                qry = " Select max(Transcation_Type)[Transcation Type],max([Customer Type])[Customer Type],max(Doc_Status)[Doc Status],max(Location)Location,max([GST No])[Location GST],max([Sub Location])[Sub Location],max(Invoice_Date)Invoice_Date,Invoice_No,max([Invoice Type])[Invoice Type],MAX([Route No])[Route No],max([Party Code])[Party Code],max(VLC_Code_VLC_Uploader) as [DCS Uploader],max([Party Name])[Party Name],
                        max([Customer GSTNo])[Customer GSTNo],max([State Code])[Party State Code],cast(sum([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(sum([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(sum([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast(sum(KKF) as decimal(10,2))KKF,Cast(sum([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(sum([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(sum([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(sum([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(sum([IGST Amt]) as decimal(10,2))[IGST Amt],cast(sum([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(sum([Total Amt]) as decimal(10,2))[Total Amt],sum([Subsidy Amt])[Subsidy Amt],max([B2B/B2C])[B2B/B2C],max(Ack_No)[Ack No],max(Ack_Date)[Ack Date],max(IRN_No)[IRN No],max(EWayBillNo)EWayBillNo,max(EWayBillDate)EWayBillDate
                        ,max(Created_By)[Created By],max(Created_Date)[Created Date],max([Cancel By])[Cancel By],max([Cancel Date])[Cancel Date] from (" + BaseQryCancel + " )XX  "
                If TxtTransaction.arrValueMember IsNot Nothing AndAlso TxtTransaction.arrValueMember.Count > 0 Then
                    qry += " where XX.Transcation_Type In(" + clsCommon.GetMulcallString(TxtTransaction.arrValueMember) + ")" + Environment.NewLine
                End If
                qry += " Group by xx.Invoice_No "
                qry += " order by max(xx.Transcation_Type), max(xx.Invoice_Date)   "
            End If
            Dim qryreturn As String = ""
            qryreturn = " select  case when TSPL_SD_SALE_RETURN_HEAD.Screen_Type='CT' then 'APS' else (case when TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' then 'MCC' else 'Dairy Sale'end)end as Transaction_Type,
case when TSPL_SD_SALE_RETURN_HEAD.Status=1 then 'Approved' else'Pending' end as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location AS [Location],
TSPL_SD_SALE_RETURN_HEAD.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) as Invoice_Date,
TSPL_SD_SALE_RETURN_HEAD.Document_Code as Invoice_No,
TSPL_SD_SALE_RETURN_HEAD.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_RETURN_DETAIL.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
  TSPL_SD_SALE_RETURN_DETAIL.Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SD_SALE_RETURN_DETAIL.Tax1='KKF' or TSPL_SD_SALE_RETURN_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_RETURN_DETAIL.tax3='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate else  TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate + TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate end ) else (case when  TSPL_SD_SALE_RETURN_DETAIL.tax1='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate else TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate +TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SD_SALE_RETURN_DETAIL.Tax1='KKF' or TSPL_SD_SALE_RETURN_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_RETURN_DETAIL.tax3='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt else  TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt end ) else (case when  TSPL_SD_SALE_RETURN_DETAIL.tax1='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt else TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt end)end as [Basic Amt]
 TSPL_SD_SALE_RETURN_DETAIL.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_RETURN_DETAIL.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_RETURN_DETAIL.Amt_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX1='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX2='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX3='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX4='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX5='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX6='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX7='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX8='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX9='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX10='KKF' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M' Then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX1='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX2='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX3='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX4='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX5='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX6='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX7='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX8='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX9='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX10='TCS' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX1='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX2='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX3='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX4='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX5='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX6='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX7='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX8='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX9='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX10='CGST' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX1='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX2='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX3='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX4='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX5='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX6='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX7='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX8='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX9='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX10='SGST' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX1='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX2='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX3='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX4='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX5='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX6='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX7='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX8='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX9='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL.TAX10='IGST' THEN TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],
TSPL_SD_SALE_RETURN_HEAD.Ack_No,TSPL_SD_SALE_RETURN_HEAD.Ack_Date,TSPL_SD_SALE_RETURN_HEAD.IRN_No,
Report_UOM.UOM_Code as Report_UOM,
(Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SALE_RETURN_HEAD.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_RETURN_HEAD.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
'' AS EWayBillNo,'' AS EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Invoice_Date1,TSPL_SD_SALE_RETURN_HEAD.Created_By,
Convert(varchar(20),TSPL_SD_SALE_RETURN_HEAD.Created_Date,103) as Created_Date

                           from TSPL_SD_SALE_RETURN_HEAD
left join TSPL_SD_SALE_RETURN_DETAIL on TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Document_Code
left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_DETAIL.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_RETURN_DETAIL.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_RETURN_DETAIL.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_RETURN_DETAIL.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_RETURN_HEAD.Customer_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SD_SALE_RETURN_HEAD.Customer_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_RETURN_HEAD.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_RETURN_HEAD.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_RETURN_HEAD .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_RETURN_HEAD .TAX10 
where convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,TSPL_SD_SALE_RETURN_HEAD.Document_Date,103)<=Convert( Date,'" + strToDate + "',103) "

            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                qryreturn += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                qryreturn += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                qryreturn += " and TSPL_SD_SALE_RETURN_HEAD.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                qryreturn += " and TSPL_SD_SALE_RETURN_DETAIL.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            qryreturn += " union all

select  case when TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Screen_Type='CT' then 'APS' else (case when TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Trans_Type='MCC' then 'MCC' else 'Dairy Sale'end)end as Transaction_Type,
'Cancel' as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Bill_To_Location AS [Location],
TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Document_Date,103) as Invoice_Date,
TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Document_Code as Invoice_No,
TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
  TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Tax1='KKF' or TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Tax2='KKF' then (case when TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.tax3='IGST' then TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Rate else  TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Rate + TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Rate end ) else (case when  TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.tax1='IGST' then TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Rate else TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Rate +TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SD_SALE_RETURN_DETAIL.Tax1='KKF' or TSPL_SD_SALE_RETURN_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_RETURN_DETAIL.tax3='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt else  TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt end ) else (case when  TSPL_SD_SALE_RETURN_DETAIL.tax1='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt else TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt end)end as [Basic Amt]
 TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Amt_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10='KKF' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10='TCS' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10='CGST' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Rate
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10='SGST' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10='IGST' THEN TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Item_Net_Amt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],
TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Ack_No,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Ack_Date,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.IRN_No,
Report_UOM.UOM_Code as Report_UOM,
(Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
'' AS EWayBillNo,'' AS EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Against_Invoice_No,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Invoice_Date1,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Cancel_By as Created_By,
Convert(varchar(20),TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Cancel_On,103) as Created_Date

                           from TSPL_SD_SALE_RETURN_HEAD_Cancel_Data
left join TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data on TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Document_Code
left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Against_Invoice_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Customer_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Customer_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Cancel_Data .TAX10 
where convert(date,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Document_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Document_Date,103)<=Convert( Date,'" + strToDate + "',103) "

            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                qryreturn += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                qryreturn += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                qryreturn += " and TSPL_SD_SALE_RETURN_HEAD_Cancel_Data.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                qryreturn += " and TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If

            qryreturn += " Union all
                            select  case when TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Screen_Type='CT' then 'APS' else (case when TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Trans_Type='MCC' then 'MCC' else 'Dairy Sale'end)end as Transaction_Type,
'Delete' as Doc_Status,
TSPL_CUSTOMER_MASTER.Cust_Type_Code as[Customer Type],
TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Bill_To_Location AS [Location],
TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Sub_Location_code AS [Sub Location],
Convert(varchar(20),TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Document_Date,103) as Invoice_Date,
TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Document_Code as Invoice_No,
TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Route_No as [Route No],
TSPL_CUSTOMER_MASTER.Cust_Code as[Party Code],
TSPL_CUSTOMER_MASTER.Customer_Name AS [Party Name],
 TSPL_LOCATION_MASTER.GSTNO AS [GST No],
 TSPL_CUSTOMER_MASTER.GSTNO as [Customer GSTNo],
  TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
   TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Item_Code as [Item Code],
tspl_item_master.Item_Desc as [Item Name],
  TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Qty as [Product Qty],
  tspl_item_master.HSN_Code,
  case when TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Tax1='KKF' or TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Tax2='KKF' then (case when TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.tax3='IGST' then TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Rate else  TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Rate + TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Rate end ) else (case when  TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.tax1='IGST' then TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Rate else TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Rate +TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Rate end)end as [IGST Rate], 
  --case when TSPL_SD_SALE_RETURN_DETAIL.Tax1='KKF' or TSPL_SD_SALE_RETURN_DETAIL.Tax2='KKF' then (case when TSPL_SD_SALE_RETURN_DETAIL.tax3='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt else  TSPL_SD_SALE_RETURN_DETAIL.TAX3_Base_Amt end ) else (case when  TSPL_SD_SALE_RETURN_DETAIL.tax1='IGST' then TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt else TSPL_SD_SALE_RETURN_DETAIL.TAX1_Base_Amt end)end as [Basic Amt]
 TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Amount as[ItemBasic Amt],
    TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.disc_Amt as[Margin Amt],
  TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Amt_Less_Discount as [Basic Amt] 
,Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9='KKF'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9_Amt 
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10='KKF' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10_Amt else 0  END) AS [KKF],
					CASE When Tax1.Type='M'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Amt
    				WHEN Tax2.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Amt
    				WHEN Tax3.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Amt
    				WHEN Tax4.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Amt
    				WHEN Tax5.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5_Amt
    				WHEN Tax6.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6_Amt
    				WHEN Tax7.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7_Amt
    				WHEN Tax8.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8_Amt
    				WHEN Tax9.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9_Amt
    				WHEN Tax10.Type='M' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [Mandi Tax Amt],
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9='TCS'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10='TCS' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10_Amt  else 0 END AS [Party TCS Amt],
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9='CGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10='CGST' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10_Amt else 0 END AS [CGST Amt],
					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Rate
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9='SGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10='SGST' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10_Amt else 0 END) AS [SGST Amt],
					
					CASE WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX1_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX2_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX3_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX4_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX5_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX6_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX7_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX8_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9='IGST'  THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX9_Amt
    				WHEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10='IGST' THEN TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.TAX10_Amt  else 0 END AS [IGST Amt],
                    TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Total_Tax_Amt as [Total Tax Amt],
					TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Item_Net_Amt as [Total Amt],
					 case  when TSPL_CUSTOMER_MASTER.GSTNO is null or  TSPL_CUSTOMER_MASTER.GSTNO = ''  then  'B2C' else 'B2B' end as [B2B/B2C],
TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Ack_No,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Ack_Date,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.IRN_No,
Report_UOM.UOM_Code as Report_UOM,
(Qty * tspl_item_uom_detail.Conversion_Factor/Report_UOM.Conversion_Factor ) as ReportUOM_Qty
,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.TotalSubsidyAmt as [Subsidy Amt],Case when TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Is_Taxable=1 then 'Taxable' else 'Non-Taxable' end as [Invoice Type],
'' AS EWayBillNo,'' AS EWayBillDate,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Against_Invoice_No,Convert(varchar(20),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) as Invoice_Date1,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Delete_By as Created_By,
Convert(varchar(20),TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Delete_On,103) as Created_Date

                           from TSPL_SD_SALE_RETURN_HEAD_Delete_Data
left join TSPL_SD_SALE_RETURN_DETAIL_Delete_Data on TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Document_Code
left join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE=TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Against_Invoice_No
LEFT JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Item_Code
left join tspl_item_uom_detail  on tspl_item_uom_detail.item_code=TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.item_code and tspl_item_uom_detail.UOM_Code=TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Unit_code
left join tspl_item_uom_detail Report_UOM  on Report_UOM.item_code=TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.item_code and Report_UOM.Report_UOM=1
LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Bill_To_Location
 LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Customer_Code
left outer join TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code= TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Customer_Code
 LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
 left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_RETURN_HEAD_Delete_Data.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_RETURN_HEAD_Delete_Data.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_Delete_Data .TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_RETURN_HEAD_Delete_Data .tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_RETURN_HEAD_Delete_Data .tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Delete_Data .TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Delete_Data .TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Delete_Data .TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Delete_Data .TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_RETURN_HEAD_Delete_Data .TAX10 
where convert(date,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Document_Date,103)>=Convert( Date,'" + strtxtfDate + "',103) and convert(date,TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Document_Date,103)<=Convert( Date,'" + strToDate + "',103) "

            If TxtCustomerType.arrValueMember IsNot Nothing AndAlso TxtCustomerType.arrValueMember.Count > 0 Then
                qryreturn += " and tspl_customer_master.Cust_Type_Code in(" + clsCommon.GetMulcallString(TxtCustomerType.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                qryreturn += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If
            If TxtSubLocation.arrValueMember IsNot Nothing AndAlso TxtSubLocation.arrValueMember.Count > 0 Then
                qryreturn += " and TSPL_SD_SALE_RETURN_HEAD_Delete_Data.Sub_Location_code in(" + clsCommon.GetMulcallString(TxtSubLocation.arrValueMember) + ")" + Environment.NewLine
            End If
            If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                qryreturn += " and TSPL_SD_SALE_RETURN_DETAIL_Delete_Data.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
            End If
            If rbtnDetail.IsChecked AndAlso rdbSaleReturn.IsChecked Then
                qry = "Select (Transaction_Type)[Transcation Type],([Customer Type])[Customer Type],(Doc_Status)[Doc Status],(Location)Location,([GST No])[Location GST],([Sub Location])[Sub Location],(Invoice_Date)[Return Date],Invoice_No,([Invoice Type])[Invoice Type],([Route No])[Route No],([Party Code])[Party Code],VLC_Code_VLC_Uploader as [DCS Uploader],([Party Name])[Party Name],([GST No])[GST No],
                       ([Customer GSTNo])[Customer GSTNo],([State Code])[Party State Code],[Item Code],[Item Name],HSN_Code AS [HSN Code],([Product Qty])[Product Qty],(Report_UOM)[Report UOM],CAST(ReportUOM_Qty AS DECIMAL(18,2)) AS [ReportUOM Qty],Cast(([IGST Rate]) as decimal(10,2))[GST Rate],cast(([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast((KKF) as decimal(10,2))KKF,Cast(([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(([IGST Amt]) as decimal(10,2))[IGST Amt],cast(([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(([Total Amt]) as decimal(10,2))[Total Amt],([Subsidy Amt])[Subsidy Amt],([B2B/B2C])[B2B/B2C],(Ack_No)[Ack No],(Ack_Date)[Ack Date],(IRN_No)[IRN No],EWayBillNo,EWayBillDate,Against_Invoice_No as[Against Invoice No],Invoice_Date1 as Invoice_Date
                        ,(Created_By)[Created By],(Created_Date)[Created Date] from (" + qryreturn + ")XX  "
                If TxtTransaction.arrValueMember IsNot Nothing AndAlso TxtTransaction.arrValueMember.Count > 0 Then
                    qry += " where XX.Transaction_Type In(" + clsCommon.GetMulcallString(TxtTransaction.arrValueMember) + ")" + Environment.NewLine
                End If
                qry += " order by xx.Transaction_Type, xx.Invoice_Date   "
            ElseIf rbtnSummary.IsChecked AndAlso rdbSaleReturn.IsChecked Then
                qry = " Select  max(Transaction_Type)[Transcation Type],max([Customer Type])[Customer Type],max(Doc_Status)[Doc Status],max(Location)Location,max([GST No])[Location GST],max([Sub Location])[Sub Location],max(Invoice_Date)[  Return Date],Invoice_No,max([Invoice Type])[Invoice Type],MAX([Route No])[Route No],max([Party Code])[Party Code],max(VLC_Code_VLC_Uploader) as [DCS Uploader],max([Party Name])[Party Name],
                        max([Customer GSTNo])[Customer GSTNo],max([State Code])[Party State Code],cast(Sum([ItemBasic Amt]) as decimal(10,2))[ItemBasic Amt],cast(Sum([Margin Amt]) as decimal(10,2))[Margin Amt],Cast(sum([Basic Amt]) as Decimal(10,2))[Basic Amt],Cast(sum(KKF) as decimal(10,2))KKF,Cast(sum([Mandi Tax Amt]) as Decimal(10,2))[Mandi Tax Amt],Cast(sum([Party TCS Amt]) as Decimal(10,2))[Party TCS Amt],
                        Cast(sum([CGST Amt]) as Decimal(10,2))[CGST Amt],Cast(sum([SGST Amt]) as Decimal(10,2))[SGST Amt],Cast(sum([IGST Amt]) as decimal(10,2))[GST Amt],cast(sum([Total Tax Amt]) as decimal(10,2))[Total Tax Amt],cast(sum([Total Amt]) as decimal(10,2))[Total Amt],sum([Subsidy Amt])[Subsidy Amt],max([B2B/B2C])[B2B/B2C],max(Ack_No)[Ack No],max(Ack_Date)[Ack Date],max(IRN_No)[IRN No],Max(EWayBillNo)EWayBillNo,max(EWayBillDate)EWayBillDate,max(Against_Invoice_No) as[Against Invoice No],max(Invoice_Date1) as Invoice_Date 
                        ,max(Created_By)[Created By],max(Created_Date)[Created Date] from (" + qryreturn + " )XX  "
                If TxtTransaction.arrValueMember IsNot Nothing AndAlso TxtTransaction.arrValueMember.Count > 0 Then
                    qry += " where XX.Transaction_Type In(" + clsCommon.GetMulcallString(TxtTransaction.arrValueMember) + ")" + Environment.NewLine
                End If
                qry += "  group by xx.Invoice_No   "
                qry += " order by max(xx.Transaction_Type), max(xx.Invoice_Date)   "
            End If
            dt = clsDBFuncationality.GetDataTable(qry)
            gvdata.GroupDescriptors.Clear()
            gvdata.MasterTemplate.SummaryRowsBottom.Clear()
            gvdata.DataSource = Nothing
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                If isPrint Then
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Print", Me.Text)
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                End If
                Exit Sub
            Else
                If isPrint AndAlso rbtnDetail.IsChecked AndAlso rdbGstInvoice.IsChecked Then
                    Dim frm As New frmCrystalReportViewer()
                    If rbtnNonTaxable.IsChecked Then
                        frm.funreport(Form_ID, CrystalReportFolder.SalesReport, dt, "crptSaleInvoiceNonTaxable", "Sales Register (Bill Of Supply)")
                    Else
                        frm.funreport(Form_ID, CrystalReportFolder.SalesReport, dt, "crptSaleInvoiceTaxable", "Sales Register (Tax Invoice)")
                    End If
                    frm = Nothing
                Else
                    RadPageView1.SelectedPage = RadPageViewPage3
                    gvdata.GroupDescriptors.Clear()
                    gvdata.MasterTemplate.SummaryRowsBottom.Clear()
                    gvdata.DataSource = dt
                    gvdata.AutoExpandGroups = True
                    gvdata.ShowGroupPanel = True
                    gvdata.ShowRowHeaderColumn = False
                    gvdata.AllowAddNewRow = False
                    gvdata.AllowDeleteRow = False
                    gvdata.EnableFiltering = True
                    gvdata.ShowFilteringRow = True
                    SetGridFormat()
                    EnableDisableCntrl(False)
                    gvdata.BestFitColumns()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        ReStoreGridLayout()
    End Sub


    Sub SetGridFormat()
        Try
            gvdata.AutoExpandGroups = True
            gvdata.ShowGroupPanel = True
            gvdata.ShowRowHeaderColumn = False
            gvdata.AllowAddNewRow = False
            gvdata.AllowDeleteRow = False
            gvdata.EnableFiltering = True
            gvdata.ShowFilteringRow = True
            For ii As Integer = 0 To gvdata.Columns.Count - 1
                gvdata.Columns(ii).ReadOnly = True
                gvdata.Columns(ii).BestFit()
            Next
            If rbtnDetail.IsChecked AndAlso Not rdbCancelInvoice.IsChecked Then
                gvdata.Columns("PrintBy").IsVisible = False
                gvdata.Columns("SNo").IsVisible = False
                gvdata.Columns("FromDate").IsVisible = False
                gvdata.Columns("ToDate").IsVisible = False
                gvdata.Columns("Bill Amount").IsVisible = False
            End If
            gvdata.Columns("Doc Status").HeaderText = "Status"
            gvdata.Columns("Invoice_Date").HeaderText = "Invoice Date"
            gvdata.Columns("EWayBillNo").IsVisible = False
            gvdata.Columns("EWayBillNo").VisibleInColumnChooser = True
            gvdata.Columns("EWayBillDate").IsVisible = False
            gvdata.Columns("EWayBillDate").VisibleInColumnChooser = True
            gvdata.Columns("Invoice_No").HeaderText = "Invoice No"
            gvdata.Columns("ItemBasic Amt").IsVisible = False
            gvdata.Columns("ItemBasic Amt").VisibleInColumnChooser = True
            gvdata.Columns("Margin Amt").IsVisible = False
            gvdata.Columns("Margin Amt").VisibleInColumnChooser = True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ExportToExcel(ByVal exporter As EnumExportTo)
        Try
            Dim arrHeader As List(Of String) = New List(Of String)()
            Dim strtemp As String = "Date Range : " + clsCommon.GetPrintDate(txtfDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")
            arrHeader.Add(strtemp)
            arrHeader.Add("Company : " + objCommonVar.CurrentCompanyName)
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                arrHeader.Add(" Customer : " + clsCommon.GetMulcallStringWithComma(txtMultiCustomer.arrDispalyMember))
            End If
            If exporter = EnumExportTo.Excel Then
                clsCommon.MyExportToExcelGrid("Sale Invoice Status Report", gvdata, arrHeader, Me.Text)
            Else
                clsCommon.MyExportToPDF("Sale Invoice Status Report", gvdata, arrHeader, "Sale Invoice Status Report", PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub

    Private Sub rmenuExport_Click(sender As Object, e As EventArgs) Handles rmenuExport.Click
        If gvdata.Rows.Count > 0 Then
            ExportToExcel(EnumExportTo.Excel)
        Else
            RadMessageBox.Show("No Data Found to Display", Me.Text)
        End If
    End Sub

    Private Sub rmenuPDF_Click(sender As Object, e As EventArgs) Handles rmenuPDF.Click
        If gvdata.Rows.Count > 0 Then
            ExportToExcel(EnumExportTo.PDF)
        Else
            RadMessageBox.Show("No Data Found to Display", Me.Text)
        End If
    End Sub

    Private Sub rbtnDetail_CheckStateChanged(sender As Object, e As EventArgs) Handles rbtnDetail.CheckStateChanged
        If rbtnDetail.IsChecked Then
            txtItem.Visible = True
            MyLabel4.Visible = True
        ElseIf rbtnDetail.IsChecked = False Then
            txtItem.Visible = False
            MyLabel4.Visible = False
        End If
    End Sub

    Private Sub rdbItemWiseCustomer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rdbItemWiseCustomer.ToggleStateChanged
        If rdbItemWiseCustomer.IsChecked Then
            RadGroupBox2.Enabled = True
            MyLabel1.Visible = False
            TxtSubLocation.Visible = False
            MyLabel2.Visible = False
            TxtCustomerType.Visible = False
            MyLabel10.Visible = False
            TxtTransaction.Visible = False
        Else
            RadGroupBox2.Enabled = True
            RadGroupBox2.Enabled = True
            MyLabel1.Visible = True
            TxtSubLocation.Visible = True
            MyLabel2.Visible = True
            TxtCustomerType.Visible = True
            MyLabel10.Visible = True
            TxtTransaction.Visible = True
        End If
    End Sub

    Private Sub rdbInvoiceCount_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rdbInvoiceCount.ToggleStateChanged
        If rdbInvoiceCount.IsChecked Then
            RadGroupBox2.Enabled = False
        Else
            RadGroupBox2.Enabled = True
        End If
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            isPrint = True
            If rbtnDetail.IsChecked Then
                LoadData()
            End If
            isPrint = False
        Catch ex As Exception
            isPrint = False
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class