Imports common
Imports System.IO
Imports System.Globalization

Public Class rptHSNWiseSaleReport
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim dtTax As DataTable = New DataTable()
    Dim isPrint As Boolean = False
    Dim isHSNWise As Boolean = False
#End Region
    Private Sub rptHSNWiseSaleReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        funreset()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        chkKKFMandi.Visible = False
    End Sub

    Private Sub txtTaxCode__My_Click(sender As Object, e As EventArgs) Handles txtTaxCode._My_Click
        Try
            Dim qry As String = "select Tax_Code as [Tax Code],Tax_Code_Desc as [Tax Name] from TSPL_TAX_MASTER  "
            If chkKKFMandi.Checked Then
            Else
                qry += " where Type in ('SGST','CGST','IGST') OR Is_TCS = 'Y' "
            End If
            txtTaxCode.arrValueMember = clsCommon.ShowMultipleSelectForm("HSNTax", qry, "Tax Code", "Tax Name", txtTaxCode.arrValueMember, txtTaxCode.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtItem__My_Click(sender As Object, e As EventArgs) Handles txtItem._My_Click
        Try
            Dim qry As String = "select  Item_Code as [Item Code] ,Item_Desc as  [Item Desc] ,Short_Description as [Short Description] from TSPL_ITEM_MASTER where 1=1 "
            If cboItemType.SelectedValue IsNot Nothing AndAlso clsCommon.myLen(cboItemType.SelectedValue) > 0 Then
                qry &= " And Item_Type = '" & clsCommon.myCstr(cboItemType.SelectedValue) & "'"
            End If
            txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("HSNItem", qry, "Item Code", "Item Desc", txtItem.arrValueMember, txtItem.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        funreset()
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        LoadData()
    End Sub

    Sub funreset()
        EnableDisableControls(True)
        LoadItemType()
        gv1.DataSource = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
    End Sub

    Private Sub EnableDisableControls(ByVal val As Boolean)
        RadGroupBox1.Enabled = val
    End Sub

    Function ReturnQry(ByVal isAcc As Boolean, ByVal FromDate As String, ByVal ToDate As String, ByVal strLocation As String, ByVal UOMType As String, ByRef strOutTaxColumnSum As String) As String
        strOutTaxColumnSum = ""
        Dim qry As String = ""
        Dim whrcls As String = ""
        Dim BaseQuery As String = ""
        Dim FinalQuery As String = ""
        ' whrcls = " where "
        If txtTaxCode.arrValueMember IsNot Nothing Then
            whrcls = " and tax in (''," & clsCommon.GetMulcallString(txtTaxCode.arrValueMember) & ") "
        End If

        If txtItem.arrValueMember IsNot Nothing Then
            whrcls = " and Item_Code in (" & clsCommon.GetMulcallString(txtItem.arrValueMember) & ") "
        End If

        If txtTransaction.arrValueMember IsNot Nothing Then
            whrcls = " and Trans_Name in (" & clsCommon.GetMulcallString(txtTransaction.arrValueMember) & ") "
        End If

        Dim ItemType As String = " 1=1 "
        If cboItemType.SelectedValue IsNot Nothing AndAlso clsCommon.myLen(cboItemType.SelectedValue) > 0 Then
            ItemType &= " And tspl_item_master.Item_Type = '" & clsCommon.myCstr(cboItemType.SelectedValue) & "' "
        End If

        If Not isAcc AndAlso isHSNWise Then
            qry = "---------------- VCGL---------------------------- 
            select 'VCGL' as Trans_Name, TSPL_VCGL_Detail.Document_No,'' as Item_Code ,0 as Qty ,'' as UOM,case when TSPL_VCGL_Detail.Row_Type ='Customer' then TSPL_VCGL_Detail.Dr_Amount-TSPL_VCGL_Detail.Cr_Amount else 0 end as Item_Net_Amt ,0 as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_VCGL_Detail left outer join TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No  =TSPL_VCGL_Detail.Document_No Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_VCGL_Head.VC_Code And TSPL_VCGL_Head.Document_Type='C' where " & IIf(clsCommon.myLen(strLocation) > 0, "TSPL_VCGL_Head.Location_Segment='" & strLocation & "' And", Nothing) & " TSPL_VCGL_Detail.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_VCGL  from TSPL_Customer_Invoice_Head
            left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_VCGL ,'')<>''     and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_VCGL_Head.Status =1  
            union all 
            select 'VCGL' as Trans_Name, TSPL_VCGL_Head.Document_No,'' as Item_Code ,0 as Qty ,'' as UOM,case when TSPL_VCGL_Head.Document_Type  ='C' then TSPL_VCGL_Head.Tot_Cr_Amount-TSPL_VCGL_Head.Tot_Dr_Amount else 0 end as Item_Net_Amt ,0 as Total_Tax_Amt, '' as Tax, 0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_VCGL_Detail left outer join TSPL_VCGL_Head on TSPL_VCGL_Head.Document_No  =TSPL_VCGL_Detail.Document_No  Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_VCGL_Head.VC_Code And TSPL_VCGL_Head.Document_Type='C'
            where " & IIf(clsCommon.myLen(strLocation) > 0, "TSPL_VCGL_Head.Location_Segment='" & strLocation & "' And", Nothing) & " TSPL_VCGL_Head.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_VCGL  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_VCGL ,'')<>''     and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_VCGL_Head.Status =1  
            Union all "
        End If
        If isAcc OrElse (isHSNWise AndAlso Not isAcc) Then
            qry &= "---------------- Sale Invoice ,BULK SALE,CAN SALE--------------------- 
            select 'DS'  as Trans_Name, TSPL_SD_SALE_INVOICE_DETAIL.Document_Code as Document_No,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code ,TSPL_SD_SALE_INVOICE_DETAIL.Qty ,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as UOM,TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt ,TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered " & IIf(isAcc, ",TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type", Nothing) & "   from TSPL_SD_SALE_INVOICE_DETAIL  left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.Document_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
            left outer join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER .Location_Code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & " TSPL_SD_SALE_INVOICE_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) "
            If Not isAcc Then
                qry &= " and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 "
            End If
            qry &= " ) and TSPL_SD_SALE_INVOICE_HEAD.Status =1  "
        End If
        If isHSNWise Then
            qry &= " union all 
            select 'CAN-SALE' as Trans_Name, TSPL_CANSALE_INVOICE_detail.Document_No,TSPL_CANSALE_INVOICE_detail.ItemCode ,TSPL_CANSALE_INVOICE_detail.Qty ,TSPL_CANSALE_INVOICE_detail.UOM ,TSPL_CANSALE_INVOICE_detail.Item_Net_Amt ,TSPL_CANSALE_INVOICE_detail.Total_Tax_Amt,'' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered  from TSPL_CANSALE_INVOICE_detail left outer join TSPL_CANSALE_INVOICE_HEAD on TSPL_CANSALE_INVOICE_HEAD .Document_no =TSPL_CANSALE_INVOICE_detail.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_CANSALE_INVOICE_HEAD.Document_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_CANSALE_INVOICE_HEAD.Customer_Code
            where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_CANSALE_INVOICE_HEAD.Location_Code='" & strLocation & "' And", Nothing) & " TSPL_CANSALE_INVOICE_HEAD.Posted =1 and TSPL_CANSALE_INVOICE_HEAD.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all 
            select  'INVOICE-BS' as Trans_Name,TSPL_INVOICE_DETAIL_BULKSALE.Document_No,TSPL_INVOICE_DETAIL_BULKSALE.item_code,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceQty as Qty,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as UOM,TSPL_INVOICE_DETAIL_BULKSALE.InvoiceAmount ,TSPL_INVOICE_DETAIL_BULKSALE.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_INVOICE_DETAIL_BULKSALE left outer join TSPL_INVOICE_MASTER_BULKSALE on TSPL_INVOICE_MASTER_BULKSALE .Document_no =TSPL_INVOICE_DETAIL_BULKSALE.Document_No  left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_INVOICE_MASTER_BULKSALE.Document_No left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_INVOICE_HEAD.Customer_Code
            where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_INVOICE_MASTER_BULKSALE.Location_Code='" & strLocation & "' And", Nothing) & " TSPL_INVOICE_MASTER_BULKSALE.Posted =1 and TSPL_INVOICE_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all  
            ---------------- Sale RETURN ,CAN SALE RETURN----------------------- 
            select 'DSR' AS Trans_Name, TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,-(TSPL_SD_SALE_RETURN_DETAIL.Qty)Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,-(TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code  Left Outer Join TSPl_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_SD_SALE_RETURN_HEAD.Customer_Code
            where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & " TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT'))  and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 
            union all 
            select 'BULK-SALE-RE' AS Trans_Name, TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No,TSPL_SALE_RETURN_DETAIL_BULKSALE.item_code,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceQty as Qty,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code as UOM,TSPL_SALE_RETURN_DETAIL_BULKSALE.InvoiceAmount ,TSPL_SALE_RETURN_DETAIL_BULKSALE.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SALE_RETURN_DETAIL_BULKSALE left outer join TSPL_SALE_RETURN_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE .Document_no =TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_Return_No =TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No Left Outer Join TSPL_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_Customer_INVOICE_HEAD.Customer_Code where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code='" & strLocation & "' And", Nothing) & " TSPL_SALE_RETURN_MASTER_BULKSALE.Posted =1 AND TSPL_SALE_RETURN_MASTER_BULKSALE.AGAINST='Bulk Invoice'  and TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all 
            ---------------- Scrap Sale Return---------------------- 
            select 'SCRAP-S-R' as Trans_Name, TSPL_SCRAPSALE_DETAIL_RETURN.Document_No,TSPL_SCRAPSALE_DETAIL_RETURN.Item_Code ,TSPL_SCRAPSALE_DETAIL_RETURN.shipped_Qty AS Qty ,TSPL_SCRAPSALE_DETAIL_RETURN.Unit_code as UOM,TSPL_SCRAPSALE_DETAIL_RETURN.TotalAmt as Item_Net_Amt ,TSPL_SCRAPSALE_DETAIL_RETURN.TotalTaxAmt as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SCRAPSALE_DETAIL_RETURN left outer join TSPL_SCRAPSALE_HEAD_RETURN on TSPL_SCRAPSALE_HEAD_RETURN .Document_No =TSPL_SCRAPSALE_DETAIL_RETURN.Document_No Left Outer Join TSPL_Customer_Master On TSPL_Customer_master.cust_code=TSPL_SCRAPSALE_HEAD_RETURN.cust_code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SCRAPSALE_HEAD_RETURN.Loc_Code='" & strLocation & "' And", Nothing) & " TSPL_SCRAPSALE_DETAIL_RETURN.Document_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where  isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPSALE_HEAD_RETURN.ispost =1  
            union all   
            ---------------- Scrap Sale ----------------- 
            select 'SCRAP-SALE' as Trans_Name, TSPL_SCRAPINVOICE_DETAIL.invoice_No AS Document_No ,TSPL_SCRAPINVOICE_DETAIL.Item_Code ,TSPL_SCRAPINVOICE_DETAIL.shipped_Qty AS Qty ,TSPL_SCRAPINVOICE_DETAIL.Unit_code as UOM,TSPL_SCRAPINVOICE_DETAIL.TotalAmt as Item_Net_Amt ,TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt as Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SCRAPINVOICE_DETAIL left outer join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD.invoice_No =TSPL_SCRAPINVOICE_DETAIL.invoice_No Left Outer Join TSPL_Customer_Master On TSpl_customer_master.cust_code=TSPL_SCRAPINVOICE_HEAD.cust_code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SCRAPINVOICE_HEAD.Loc_Code='" & strLocation & "' And", Nothing) & " TSPL_SCRAPINVOICE_DETAIL.invoice_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrap  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrap ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPINVOICE_HEAD.ispost =1  
            union all 
            ---------------- mcc Material Sale RETURN-------------------------------- 
            select 'M-Material-R' as Trans_Name,  TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,-(TSPL_SD_SALE_RETURN_DETAIL.Qty)Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,-(TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt, '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SD_SALE_RETURN_DETAIL  left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code Left Outer Join TSPL_Customer_Master On TSPL_Customer_master.cust_code=TSPL_SD_SALE_RETURN_HEAD.customer_code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & " TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 AND TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' 
            union all 
            ---------------- Security Receipt--------------------------
            select 'AR-INVOICE' as Trans_Name, TSPL_Customer_Invoice_Head.Against_Security_Receipt_No as Document_No,'' AS Item_Code ,0 AS Qty ,'' as UOM,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  (isnull(TSPL_Customer_Invoice_Detail.Total_Amount ,0))as Item_Net_Amt ,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  (isnull(TSPL_Customer_Invoice_Detail.Total_Tax  ,0)) as Total_Tax_Amt,'' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_Customer_Invoice_Detail left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Document_No =TSPL_Customer_Invoice_Detail.Document_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where  isnull(TSPL_Customer_INVOICE_HEAD.Against_Security_Receipt_No ,'')<>''  and TSPL_Customer_INVOICE_HEAD.Status =1    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & "  TSPL_Customer_Invoice_Head.Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 
            union all 
             ---------------- Transfer --------------------------
            select  'STO-TRANSFER' as Trans_Name,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM,  TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,0 As GST_Registered from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_TRANSFER_ORDER_HEAD.From_Location='" & strLocation & "' And", Nothing) & "  convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)  and  TSPL_TRANSFER_ORDER_HEAD.status = 1  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'O'
            union all
            ---------------- Transfer Return--------------------------
            select  'STO-TRANS-R' as Trans_Name,TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,-(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty) as Qty,TSPL_TRANSFER_ORDER_DETAIL.Unit_code as UOM,  -(TSPL_TRANSFER_ORDER_DETAIL.Item_Net_Amt)Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  '' as Tax,  0 as Tax_Amt ,0 as Tax_Rate,0 As GST_Registered from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_TRANSFER_ORDER_HEAD.From_Location='" & strLocation & "' And", Nothing) & " convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)  and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'T'"
        End If
        qry += " " & Environment.NewLine & " Union all " & Environment.NewLine & ""
        For ii As Integer = 1 To 6
            If ii > 1 Then
                BaseQuery += " Union all " & Environment.NewLine & ""
            End If
            If isAcc OrElse (isHSNWise AndAlso Not isAcc) Then
                BaseQuery += "---------------- Sale Invoice ,BULK SALE,CAN SALE--------------------- 
            select 'DS'  as Trans_Name, TSPL_SD_SALE_INVOICE_DETAIL.Document_Code,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code ,0 as Qty ,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt,   isnull(TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & ",'') AS Tax,  TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & "_Amt as TAX_Amt,TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & "_Rate as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered " & IIf(isAcc, ",TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type", Nothing) & "    from TSPL_SD_SALE_INVOICE_DETAIL  left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD .Document_Code =TSPL_SD_SALE_INVOICE_DETAIL.Document_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
            left outer join TSPL_LOCATION_MASTER on TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location =TSPL_LOCATION_MASTER .Location_Code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_LOCATION_MASTER.Location_Code='" & strLocation & "' And", Nothing) & " TSPL_SD_SALE_INVOICE_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT'))  and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_INVOICE_HEAD.Status =1  and TSPL_SD_SALE_INVOICE_DETAIL.TAX" & ii & "_Amt > 0 "
            End If
            If isHSNWise Then
                BaseQuery &= " union all 

            Select 'CAN-SALE' as Trans_Name, TSPL_CANSALE_INVOICE_detail.Document_No,TSPL_CANSALE_INVOICE_detail.ItemCode ,0 as Qty ,TSPL_CANSALE_INVOICE_detail.UOM ,0 as Item_Net_Amt,TSPL_CANSALE_INVOICE_detail.Total_Tax_Amt," & IIf(ii = 6, " ''  ", " isnull(TSPL_CANSALE_INVOICE_detail.TAX" & ii & ",'')") & "  AS Tax," & IIf(ii = 6, " 0  ", " TSPL_CANSALE_INVOICE_detail.TAX" & ii & "_Amt ") & " as TAX_Amt," & IIf(ii = 6, " 0  ", " TSPL_CANSALE_INVOICE_detail.TAX" & ii & "_Rate ") & "  as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered  from TSPL_CANSALE_INVOICE_detail left outer join TSPL_CANSALE_INVOICE_HEAD on TSPL_CANSALE_INVOICE_HEAD .Document_no =TSPL_CANSALE_INVOICE_detail.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_CANSALE_INVOICE_HEAD.Document_No 
            Left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_CANSALE_INVOICE_HEAD.Customer_Code where TSPL_CANSALE_INVOICE_HEAD.Posted =1 and TSPL_CANSALE_INVOICE_HEAD.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_CANSALE_INVOICE_HEAD.Location_Code='" & strLocation & "' And", Nothing) & " isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and  " & IIf(ii < 6, "  TSPL_CANSALE_INVOICE_detail.TAX" & ii & "_Amt>0 ", " 2=2") & "
            union all 
            select 'INVOICE-BS' as Trans_Name, TSPL_INVOICE_DETAIL_BULKSALE.Document_No,TSPL_INVOICE_DETAIL_BULKSALE.item_code,0 as Qty,TSPL_INVOICE_DETAIL_BULKSALE.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_INVOICE_DETAIL_BULKSALE.Total_Tax_Amt," & IIf(ii = 6, " ''  ", " isnull(TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & ",'') ") & "  AS Tax, " & IIf(ii = 6, " 0  ", "  TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & "_Amt ") & " as TAX_Amt, " & IIf(ii = 6, " 0  ", "  TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & "_Rate ") & " as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_INVOICE_DETAIL_BULKSALE left outer join TSPL_INVOICE_MASTER_BULKSALE on TSPL_INVOICE_MASTER_BULKSALE .Document_no =TSPL_INVOICE_DETAIL_BULKSALE.Document_No  left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_No =TSPL_INVOICE_MASTER_BULKSALE.Document_No Left Outer Join TSPL_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_Customer_INVOICE_HEAD.Customer_Code
            where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_INVOICE_MASTER_BULKSALE.Location_Code='" & strLocation & "' And", Nothing) & " TSPL_INVOICE_MASTER_BULKSALE.Posted =1 and TSPL_INVOICE_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_No  from TSPL_Customer_Invoice_Head  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_No ,'')<>''   and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and
            " & IIf(ii < 6, " TSPL_INVOICE_DETAIL_BULKSALE.TAX" & ii & "_Amt>0", " 2=2") & "
            union all  
            ---------------- Sale RETURN ,CAN SALE RETURN----------------------- 
            select 'DSR' AS Trans_Name, TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,0 as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt, isnull(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & ",'') AS Tax,  -(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt) as Tax_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Rate as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SD_SALE_RETURN_DETAIL left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code  Left Outer Join TSPL_Customer_Master On TSPl_customer_Master.Cust_Code=TSPL_SD_SALE_RETURN_HEAD.Customer_Code
            where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & " TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 and TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt>0
            union all 
            select 'BULK-SALE-RE' as Trans_Name, TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No,TSPL_SALE_RETURN_DETAIL_BULKSALE.item_code,0 as Qty,TSPL_SALE_RETURN_DETAIL_BULKSALE.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SALE_RETURN_DETAIL_BULKSALE.Total_Tax_Amt," & IIf(ii = 6, " ''  ", "  isnull(TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & ",'') ") & " AS Tax, " & IIf(ii = 6, " 0  ", " TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & "_Amt") & "  as TAX_Amt, " & IIf(ii = 6, " 0  ", " TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & "_Rate ") & "  as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered  from TSPL_SALE_RETURN_DETAIL_BULKSALE left outer join TSPL_SALE_RETURN_MASTER_BULKSALE on TSPL_SALE_RETURN_MASTER_BULKSALE .Document_no =TSPL_SALE_RETURN_DETAIL_BULKSALE.Document_No left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Against_Sale_Return_No =TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No Left Outer Join TSPl_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_Customer_INVOICE_HEAD.Customer_Code where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SALE_RETURN_MASTER_BULKSALE.Location_Code='" & strLocation & "' And", Nothing) & " TSPL_SALE_RETURN_MASTER_BULKSALE.Posted =1 AND TSPL_SALE_RETURN_MASTER_BULKSALE.AGAINST='Bulk Invoice' and  " & IIf(ii < 6, "  TSPL_SALE_RETURN_DETAIL_BULKSALE.TAX" & ii & "_Amt>0 ", " 2=2") & " and TSPL_SALE_RETURN_MASTER_BULKSALE.Document_No in (select  TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Sale_Return_No ,'')<>''    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) 
            union all 
            ---------------- Scrap Sale Return---------------------- 
            select 'SCRAP-S-R' AS Trans_Name, TSPL_SCRAPSALE_DETAIL_RETURN.Document_No,TSPL_SCRAPSALE_DETAIL_RETURN.Item_Code ,0 as Qty ,TSPL_SCRAPSALE_DETAIL_RETURN.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SCRAPSALE_DETAIL_RETURN.TotalTaxAmt as Total_Tax_Amt, isnull(TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & ",'') as Tax, TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & "_Amt as Tax_Amt,TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & "_Rate as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SCRAPSALE_DETAIL_RETURN left outer join TSPL_SCRAPSALE_HEAD_RETURN on TSPL_SCRAPSALE_HEAD_RETURN .Document_No =TSPL_SCRAPSALE_DETAIL_RETURN.Document_No Left Outer Join TSPL_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_SCRAPSALE_HEAD_RETURN.Cust_Code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SCRAPSALE_HEAD_RETURN.Loc_Code='" & strLocation & "' And", Nothing) & " TSPL_SCRAPSALE_DETAIL_RETURN.Document_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where  isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrapReturn ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPSALE_HEAD_RETURN.ispost =1   and TSPL_SCRAPSALE_DETAIL_RETURN.TAX" & ii & "_Amt>0
            union all  
            ---------------- Scrap Sale ----------------- 
            select 'SCRAP-SALE' AS Trans_Name, TSPL_SCRAPINVOICE_DETAIL.invoice_No AS Document_No ,TSPL_SCRAPINVOICE_DETAIL.Item_Code ,0 as Qty ,TSPL_SCRAPINVOICE_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SCRAPINVOICE_DETAIL.TotalTaxAmt as Total_Tax_Amt, isnull(TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & ",'') as Tax, TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & "_Amt as TAX_Amt,TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & "_Rate as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SCRAPINVOICE_DETAIL left outer join TSPL_SCRAPINVOICE_HEAD on TSPL_SCRAPINVOICE_HEAD.invoice_No =TSPL_SCRAPINVOICE_DETAIL.invoice_No Left Outer Join TSPL_Customer_Master On TSPl_Customer_Master.Cust_Code=TSPL_SCRAPINVOICE_HEAD.Cust_Code  where TSPL_SCRAPINVOICE_DETAIL.invoice_No in (select  TSPL_Customer_INVOICE_HEAD.AgainstScrap  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.AgainstScrap ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SCRAPINVOICE_HEAD.ispost =1  and TSPL_SCRAPINVOICE_DETAIL.TAX" & ii & "_Amt>0
            union all 
            ---------------- mcc Material Sale RETURN-------------------------------- 
            select 'M-Material-R' as Trans_Name, TSPL_SD_SALE_RETURN_DETAIL.Document_Code,TSPL_SD_SALE_RETURN_DETAIL.Item_Code ,0 as Qty ,TSPL_SD_SALE_RETURN_DETAIL.Unit_code as UOM,0 as Item_Net_Amt ,TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt,  isnull(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & ",'') as Tax,-(TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt) as Tax_Amt,TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Rate as TAX_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_SD_SALE_RETURN_DETAIL  left outer join TSPL_SD_SALE_RETURN_HEAD on TSPL_SD_SALE_RETURN_HEAD.Document_Code  =TSPL_SD_SALE_RETURN_DETAIL.Document_Code Left Outer Join TSPL_Customer_Master On TSPL_Customer_Master.Cust_Code=TSPL_SD_SALE_RETURN_HEAD.Customer_Code  where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & " TSPL_SD_SALE_RETURN_DETAIL.Document_Code in (select  TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return  from TSPL_Customer_Invoice_Head left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_MCC_Material_Sale_Return ,'')<>'' and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0 ) and TSPL_SD_SALE_RETURN_HEAD.Status =1 AND TSPL_SD_SALE_RETURN_HEAD.Trans_Type='MCC' and TSPL_SD_SALE_RETURN_DETAIL.TAX" & ii & "_Amt > 0
            union all 
            ---------------- Security Receipt--------------------------
            select 'AR-INVOICE' as Trans_Name, TSPL_Customer_Invoice_Head.Against_Security_Receipt_No as Document_No,'' AS Item_Code ,0 AS Qty ,'' as UOM,0 as Item_Net_Amt ,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  (isnull(TSPL_Customer_Invoice_Detail.Total_Tax  ,0)) as Total_Tax_Amt,TSPL_Customer_Invoice_Detail.TAX" & ii & " as Tax,case when TSPL_Customer_Invoice_Head.Document_Type ='C' then -1 else 1 end *  isnull(TSPL_Customer_Invoice_Detail.TAX" & ii & "_Amt,0) as Tax_Amt ,TSPL_Customer_Invoice_Detail.TAX" & ii & "_Rate as Tax_Rate,TSPL_CUSTOMER_MASTER.GST_Registered from TSPL_Customer_Invoice_Detail left outer join TSPL_Customer_INVOICE_HEAD on TSPL_Customer_INVOICE_HEAD .Document_No =TSPL_Customer_Invoice_Detail.Document_No  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_Customer_Invoice_Head.Customer_Code where isnull(TSPL_Customer_INVOICE_HEAD.Against_Security_Receipt_No ,'')<>''  and TSPL_Customer_INVOICE_HEAD.Status =1    and convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_Customer_INVOICE_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103)   AND TSPL_Customer_Invoice_Head.Against_Sale_No NOT IN (select  Document_Code  from TSPL_SD_SALE_INVOICE_HEAD where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location='" & strLocation & "' And", Nothing) & " Document_Type in ('EX','MT')) and TSPL_CUSTOMER_MASTER.GST_Registered =1 and isnull(TSPL_CUSTOMER_MASTER.GSTNO ,'')<>'' and isnull(TSPL_Customer_Invoice_Head.Total_Tax,0) >0   and TSPL_Customer_Invoice_Detail.TAX" & ii & "_Amt >0
            union all 
             ---------------- Transfer --------------------------
           select 'STO-TRANSFER' AS Trans_Name,  TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,0 as Qty,'' as UOM,  0 as Item_Net_Amt ,TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt,  isnull(TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & ",'') as Tax,  TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt as TAX_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Rate as TAX_Rate,0 As GST_Registered from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION='" & strLocation & "' And", Nothing) & " convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103) and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt>0  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'O'   
           union all 
                   ---------------- Transfer Return--------------------------
           select 'STO-TRANS-R' AS Trans_Name,  TSPL_TRANSFER_ORDER_DETAIL.Document_No,TSPL_TRANSFER_ORDER_DETAIL.Item_Code ,0 as Qty,'' as UOM,  0 as Item_Net_Amt ,-(TSPL_TRANSFER_ORDER_DETAIL.Total_Tax_Amt)Total_Tax_Amt,  isnull(TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & ",'') as Tax,  -(TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt) as TAX_Amt ,TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Rate as TAX_Rate,0 As GST_Registered from TSPL_TRANSFER_ORDER_DETAIL  left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No =TSPL_TRANSFER_ORDER_DETAIL.Document_No LEFT join TSPL_LOCATION_MASTER as  From_LOC ON TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION=From_LOC.LOCATION_CODE where " & IIf(clsCommon.myLen(strLocation) > 0, " TSPL_TRANSFER_ORDER_HEAD.FROM_LOCATION='" & strLocation & "' And", Nothing) & " convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date ,103)>=convert(date,'" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "',103) AND convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103)<=convert(date,'" & clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") & "',103) and  TSPL_TRANSFER_ORDER_HEAD.status = 1 and TSPL_TRANSFER_ORDER_DETAIL.TAX" & ii & "_Amt>0  and TSPL_TRANSFER_ORDER_HEAD.transfer_type = 'T' "
            End If
        Next

        BaseQuery = " select Trans_Name, Document_No,xx.Item_Code,UOM,(Qty)Qty, (Item_Net_Amt)Item_Net_Amt,Tax,Tax_Rate,Tax_Amt ,GST_Registered " & IIf(isAcc, ",EInvoice_Type", Nothing) & "  from ( " & qry & " " & BaseQuery & " )xx  where 2= 2 " & whrcls & " ) xxx  left outer join TSPL_TAX_MASTER ON  TSPL_TAX_MASTER.Tax_Code = XXX.TAX "
        Dim TaxDesc As String = ""
        Dim TaxAmount As String = ""
        qry = ""
        If chkKKFMandi.Checked Then
            dtTax = clsDBFuncationality.GetDataTable("with CTERawData as ( select Tax,max(Tax_Code_Desc)Tax_Code_Desc,Sequence_No,max(Type)Type from ( select tax,Tax_Code_Desc ,Type,case when type = 'M' then 1 when type = 'K' then 2 when type = 'SGST' then 3  when type = 'CGST' then 4 when type = 'IGST' then 5 when Is_TCS = 'Y' then 6  end as Sequence_No from ( " & BaseQuery & " left outer join TSPL_ITEM_MASTER ON  tspl_item_master.Item_Code = XXX.Item_Code where Tax<> '' and  " & ItemType & " ) XXXXX   group by Tax,Sequence_No )	select CTERawData.* from CTERawData order by Sequence_No")
        Else
            dtTax = clsDBFuncationality.GetDataTable("SELECT * FROM ( select Tax_Code,Tax_Code_Desc,type,case when type = 'SGST' then 1 when type = 'CGST' then 2  when type = 'IGST' then 3  end as Sequence_No from TSPL_TAX_MASTER where Type in ('SGST','CGST','IGST') )X Order by Sequence_No ")
        End If

        Dim VoucherCount As String = "0"
        If isPrint Then
            VoucherCount = clsDBFuncationality.getSingleValue("select COUNT(Document_No) Document_No  FROM ( select Document_No from ( " & BaseQuery & "  group by Document_No )XXX ")
        End If
        For ii As Integer = 0 To dtTax.Rows.Count - 1
            If chkKKFMandi.Checked Then
                If isPrint Then
                    TaxDesc += " TaxAmount_" + clsCommon.myCstr(ii + 1) + " ,Tax_" + clsCommon.myCstr(ii + 1) + ", "
                    qry += " ,'" & dtTax.Rows(ii)("Tax_Code_Desc") & "' as  Tax_" + clsCommon.myCstr(ii + 1) + " "
                    qry += " ,sum(Case When xxx.Tax ='" & dtTax.Rows(ii)("Tax") & "' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_" + clsCommon.myCstr(ii + 1) + " "
                    TaxAmount += ",max(Tax_" + clsCommon.myCstr(ii + 1) + ") as  Tax_" + clsCommon.myCstr(ii + 1) + " , sum(TaxAmount_" + clsCommon.myCstr(ii + 1) + ") as TaxAmount_" + clsCommon.myCstr(ii + 1) + ""

                Else
                    TaxDesc += "[" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount] ,"
                    qry += " ,sum(case when xxx.Tax ='" & dtTax.Rows(ii)("Tax") & "' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as [" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount]"
                    TaxAmount += " sum([" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount])[" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount] ,"
                End If
            Else
                If isPrint Then
                    TaxDesc += " TaxAmount_" + clsCommon.myCstr(ii + 1) + " ,Tax_" + clsCommon.myCstr(ii + 1) + ", "
                    qry += " ,'" & dtTax.Rows(ii)("Tax_Code_Desc") & "' as  Tax_" + clsCommon.myCstr(ii + 1) + " "
                    qry += " ,sum(Case When xxx.Type ='" & dtTax.Rows(ii)("Type") & "' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_" + clsCommon.myCstr(ii + 1) + " "
                    TaxAmount += ",max(Tax_" + clsCommon.myCstr(ii + 1) + ") as  Tax_" + clsCommon.myCstr(ii + 1) + " , sum(TaxAmount_" + clsCommon.myCstr(ii + 1) + ") as TaxAmount_" + clsCommon.myCstr(ii + 1) + ""
                Else
                    TaxDesc += "[" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount] ,"
                    qry += " ,sum(case when xxx.Type ='" & dtTax.Rows(ii)("Type") & "' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as [" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount]"
                    TaxAmount += " sum([" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount])[" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount] ,"

                    strOutTaxColumnSum += " ,sum([" & dtTax.Rows(ii)("Tax_Code_Desc") & " Amount]) as "
                    If clsCommon.CompairString(clsCommon.myCstr(dtTax.Rows(ii)("type")), "SGST") = CompairStringResult.Equal Then
                        strOutTaxColumnSum += " [IGST]"
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(dtTax.Rows(ii)("type")), "CGST") = CompairStringResult.Equal Then
                        strOutTaxColumnSum += " [CGST]"
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(dtTax.Rows(ii)("type")), "IGST") = CompairStringResult.Equal Then
                        strOutTaxColumnSum += " [S/UGST]"
                    End If
                End If
            End If

        Next
        BaseQuery = " select xxx.*,Type,Is_TCS from (  " & BaseQuery & "  )xxx "
        Dim isShowKFFMandiAdd As String = ""
        Dim isShowKFFMandiRemove As String
        If chkKKFMandi.Checked Then
            isShowKFFMandiRemove = "  sum(Total_Tax_Amt) Total_Tax_Amt"
            isShowKFFMandiAdd = " sum(Taxable_Amt) AS Taxable_Amt"
        Else
            isShowKFFMandiRemove = " sum(Total_Tax_Amt-kkfAmt-MandiAmt) Total_Tax_Amt "
            isShowKFFMandiAdd = " sum(Taxable_Amt +kkfAmt+MandiAmt)Taxable_Amt"
        End If
        If Not isAcc Then
            qry += ",Type "
        End If
        If isPrint Then
            qry = " select  " & VoucherCount & " As VoucherCount, '" & clsCommon.GetPrintDate(FromDate, "dd/MMM/yyyy") & "' as FromDate,'" + clsCommon.GetPrintDate(ToDate, "dd/MMM/yyyy") + "' as ToDate,max(TypeOfSupply)TypeOfSupply,max(Comp_Name)Comp_Name,max(HSN_Code)HSN_Code,Item_Code,max(Short_Description)Short_Description,max(UOM) AS UOM, sum(Total_Qty) As Total_Qty,sum(Item_Net_Amt)Item_Net_Amt,max(Tax_Rate)Tax_Rate," & isShowKFFMandiAdd & "
               " & TaxAmount & ",max(Tax_4) as  Tax_4 ,sum(TaxAmount_4)TaxAmount_4  ," & isShowKFFMandiRemove & " from (
                Select  case when TSPL_ITEM_MASTER.Item_Type = 'F' then 'Goods' else '' end AS TypeOfSupply, TSPL_COMPANY_MASTER.Comp_Name,  TSPL_ITEM_MASTER.HSN_Code,FINAL.Item_Code,TSPL_ITEM_MASTER.Short_Description,I.UOM_Code AS UOM, isnull((Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Total_Qty,Item_Net_Amt,   isnull(Tax_Rate,0)Tax_Rate,(Item_Net_Amt - Total_Tax_Amt) AS Taxable_Amt, " & TaxDesc & "TaxAmount_4,Tax_4, Total_Tax_Amt,kkfAmt,MandiAmt  from (  select xxx.Item_Code,UOM,SUM(Qty)Qty, sum(Item_Net_Amt)Item_Net_Amt, sum(Tax_Amt)Total_Tax_Amt,  max( case when (xxx.Tax ='IGST' and isnull(Tax_Rate ,0) > 0)  then isnull(Tax_Rate ,0) else (case when xxx.Tax ='CGST' then 2 * isnull(Tax_Rate ,0) end ) end) as Tax_Rate,sum(case when xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as kkfAmt, sum(case when xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as MandiAmt  " & qry & ",sum(case when Is_TCS = 'Y' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as TaxAmount_4, 'TCS'  as Tax_4 from ( "
        Else
            If isAcc Then
                qry = " select  Case When EInvoice_Type='BB' OR EInvoice_Type='B2B' Then 'B2B' Else 'B2C' End As [Supply Type],TSPL_ITEM_MASTER.HSN_Code,FINAL.Item_Code,TSPL_ITEM_MASTER.Short_Description,I.UOM_Code AS UOM, isnull((Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Total_Qty,Item_Net_Amt,   isnull(Tax_Rate,0)Tax_Rate,(Item_Net_Amt - Total_Tax_Amt) AS Taxable_Amt, " & TaxDesc & "[TCS Amount], Total_Tax_Amt,kkfAmt,MandiAmt,GST_Registered " & IIf(isAcc, ",EInvoice_Type", Nothing) & " from ( select xxx.Item_Code,UOM,SUM(Qty)Qty, sum(Item_Net_Amt)Item_Net_Amt, sum(Tax_Amt)Total_Tax_Amt,  max( case when (xxx.Tax ='IGST' and isnull(Tax_Rate ,0) > 0)  then isnull(Tax_Rate ,0) else (case when xxx.Tax ='CGST' then 2 * isnull(Tax_Rate ,0) end ) end) as Tax_Rate,sum(case when xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as kkfAmt, sum(case when xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as MandiAmt,Max(GST_Registered) As GST_Registered " & IIf(isAcc, ",Max(EInvoice_Type) As EInvoice_Type", Nothing) & "  " & qry & ",sum(case when Is_TCS = 'Y' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as [TCS Amount] from ( "
            Else
                qry = " select max(HSN_Code)HSN_Code,Item_Code,max(Short_Description)Short_Description,max(UOM) AS UOM, sum(Total_Qty) As Total_Qty,sum(Item_Net_Amt)Item_Net_Amt,max(Tax_Rate)Tax_Rate," & isShowKFFMandiAdd & "," & TaxAmount & " sum([TCS Amount])[TCS Amount]," & isShowKFFMandiRemove & " from ( select  TSPL_ITEM_MASTER.HSN_Code,FINAL.Item_Code,TSPL_ITEM_MASTER.Short_Description,I.UOM_Code + '-' + I.UOM_Description AS UOM, isnull((Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Total_Qty,Item_Net_Amt,   isnull(Tax_Rate,0)Tax_Rate,(Item_Net_Amt - Total_Tax_Amt) AS Taxable_Amt, " & TaxDesc & "[TCS Amount], Total_Tax_Amt,kkfAmt,MandiAmt  from ( select xxx.Item_Code,UOM,SUM(Qty)Qty, sum(Item_Net_Amt)Item_Net_Amt, sum(Tax_Amt)Total_Tax_Amt,  max( case when (xxx.Tax ='IGST' and isnull(Tax_Rate ,0) > 0)  then isnull(Tax_Rate ,0) else (case when xxx.Tax ='CGST' then 2 * isnull(Tax_Rate ,0) end ) end) as Tax_Rate,sum(case when xxx.Type ='K' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as kkfAmt, sum(case when xxx.Type ='M' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as MandiAmt  " & qry & ",sum(case when Is_TCS = 'Y' then 1 else 0 end *  (isnull(Tax_Amt ,0))) as [TCS Amount] from ( "
            End If
        End If

        If isAcc Then
            BaseQuery = "" & qry & " " & BaseQuery & " GROUP BY xxx.EInvoice_Type,xxx.Item_Code,UOM,Type ) FINAL left outer join tspl_item_master on TSPL_ITEM_MASTER.Item_Code = FINAL.Item_Code  LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = FINAL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = FINAL.UOM	LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL "
            If clsCommon.CompairString(UOMType, "Billing UOM") = CompairStringResult.Equal Then
                BaseQuery &= " where Billing_UOM = 1 ) "
            ElseIf clsCommon.CompairString(UOMType, "Default UOM") = CompairStringResult.Equal Then
                BaseQuery &= " where Default_UOM = 1 ) "
            Else
                BaseQuery &= " where Report_UOM = 1 ) "
            End If
            BaseQuery &= " as  I ON FINAL.Item_Code = I.item_code left outer join TSPL_COMPANY_MASTER on 2 =2   where " & ItemType & "  "
        Else
            BaseQuery = "with CTERawData as ( " & qry & " " & BaseQuery & " GROUP BY xxx.Item_Code,UOM,Type ) FINAL left outer join tspl_item_master on TSPL_ITEM_MASTER.Item_Code = FINAL.Item_Code  LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = FINAL.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code = FINAL.UOM	LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON FINAL.Item_Code = I.item_code left outer join TSPL_COMPANY_MASTER on 2 =2   where " & ItemType & " ) xxxxFinal group by Item_Code ) select CTERawData.* from CTERawData Where IsNull(HSN_Code,'') <>'' "
        End If
        Return BaseQuery
    End Function

    Private Sub LoadData()
        Try
            isHSNWise = True
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(ReturnQry(False, txtFromDate.Value, txtToDate.Value, Nothing, Nothing, ""))
            isHSNWise = False
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
                    If chkKKFMandi.Checked Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, "rptHSNWiseSale", "HSN Wise Sale")
                    Else
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, "rptHSNWiseSaleALW", "HSN Wise Sale")
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

    Private Sub gv1_CellFormatting(sender As Object, e As CellFormattingEventArgs) Handles gv1.CellFormatting
        If e.CellElement.Value IsNot Nothing AndAlso IsNumeric(e.CellElement.Value) AndAlso clsCommon.CompairString(e.Column.Name, "HSN_Code") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(e.Column.Name, "Short_Description") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(e.Column.Name, "UOM") <> CompairStringResult.Equal Then
            e.CellElement.Text = Convert.ToDecimal(e.CellElement.Value).ToString("N2", New CultureInfo("en-IN"))
        End If
    End Sub

    Sub LoadItemType()
        Try
            Dim dt As New DataTable()
            Dim Whr = " AND IS_NON_INVENTORY=0 "
            dt = clsItemMaster.getItemTypeQuery(Whr)
            cboItemType.DataSource = dt
            cboItemType.ValueMember = "Code"
            cboItemType.DisplayMember = "Name"
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub SetGridFormation()
        gv1.TableElement.TableHeaderHeight = 40
        gv1.MasterTemplate.ShowRowHeaderColumn = True
        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).IsVisible = True
            gv1.Columns(ii).Width = 80
            If chkKKFMandi.Checked = False Then
                'If gv1.Columns(ii).Name.Contains("KKF") OrElse gv1.Columns(ii).Name.Contains("Mandi") OrElse gv1.Columns(ii).Name.Contains("MANDI") Then
                '    gv1.Columns(ii).IsVisible = False
                'End If
            End If

            'For Each dr As DataRow In dtTax.Rows
            '    If clsCommon.CompairString(gv1.Columns(ii).Name, dr("Tax_Code_Desc") + " Amount") = CompairStringResult.Equal Then
            '        If gv1.Columns(ii).Name.Length > 14 Then
            '            gv1.Columns(ii).HeaderText = gv1.Columns(ii).Name.ToString().Substring(0, 14) + Environment.NewLine + gv1.Columns(ii).Name.ToString().Substring(14, gv1.Columns(ii).Name.Length - 14)
            '            Exit For
            '        End If
            '    End If
            'Next
        Next

        gv1.ShowGroupPanel = False
        gv1.Columns("HSN_Code").HeaderText = "HSN/SAC"
        gv1.Columns("Short_Description").HeaderText = "Description"
        gv1.Columns("Short_Description").Width = 110
        gv1.Columns("UOM").HeaderText = "UQC"
        gv1.Columns("Total_Qty").HeaderText = "Total Quantity"
        gv1.Columns("Item_Net_Amt").FormatString = "{0:n2}"
        gv1.Columns("Item_Net_Amt").HeaderText = "Total Value"
        gv1.Columns("Tax_Rate").HeaderText = "Tax Rate"
        gv1.Columns("Taxable_Amt").HeaderText = "Taxable Amount"
        gv1.Columns("Total_Tax_Amt").HeaderText = "Total Tax Amount"
        gv1.Columns("Item_Code").HeaderText = "Item Code"
        gv1.Columns("Item_Code").IsVisible = False
        gv1.Columns("Tax_Rate").FormatString = "{0:n2}"

        If isPrint Then
            gv1.Columns("Tax_1").IsVisible = False
            gv1.Columns("Tax_2").IsVisible = False
            gv1.Columns("Tax_3").IsVisible = False
            gv1.Columns("Tax_4").IsVisible = False

            gv1.Columns("VoucherCount").IsVisible = False
            gv1.Columns("Comp_Name").IsVisible = False
            gv1.Columns("TypeOfSupply").IsVisible = False
            gv1.Columns("FromDate").IsVisible = False
            gv1.Columns("ToDate").IsVisible = False
            For ii As Integer = 0 To dtTax.Rows.Count - 1
                gv1.Columns("TaxAmount_" & clsCommon.myCstr(ii + 1) & "").HeaderText = dtTax.Rows(ii)("Tax_Code_Desc")
            Next
            If chkKKFMandi.Checked = False Then
                gv1.Columns("TaxAmount_4").HeaderText = "TCS Amount"
            End If
        End If

        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim j As Integer = 0
        If isPrint Then
            j = 9
        Else
            j = 4
        End If
        For ii As Integer = j To gv1.Columns.Count - 1
            If clsCommon.CompairString(gv1.Columns(ii).Name, "Tax_Rate") = CompairStringResult.Equal Then
            Else
                summaryRowItem.Add(New GridViewSummaryItem(gv1.Columns(ii).Name, "{0:F2}", GridAggregateFunction.Sum))
            End If
        Next
        gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(MyBase.Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
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
                clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        Try
            If gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Date : " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "  To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy"))
                If txtTaxCode.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Tax Code : " & clsCommon.GetMulcallString(txtTaxCode.arrValueMember) & "   Tax Name :" & clsCommon.GetMulcallString(txtTaxCode.arrDispalyMember) & "")
                End If

                If txtItem.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Item Code : " & clsCommon.GetMulcallString(txtItem.arrValueMember) & "   Tax Name :" & clsCommon.GetMulcallString(txtItem.arrDispalyMember) & "")
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
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptHSNWiseSaleReport & "'"))
                arrHeader.Add("Date : " & clsCommon.myCDate(txtFromDate.Value) + "  To " + clsCommon.myCDate(txtToDate.Value))

                If txtTaxCode.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Tax Code : " & clsCommon.GetMulcallString(txtTaxCode.arrValueMember) & "   Tax Name :" & clsCommon.GetMulcallString(txtTaxCode.arrDispalyMember) & "")
                End If

                If txtItem.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Item Code : " & clsCommon.GetMulcallString(txtItem.arrValueMember) & "   Tax Name :" & clsCommon.GetMulcallString(txtItem.arrDispalyMember) & "")
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
        isPrint = True
        LoadData()
        isPrint = False
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

    Private Sub rptHSNWiseSaleReport_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.F12 Then
            chkKKFMandi.Visible = True
        End If
    End Sub
End Class

