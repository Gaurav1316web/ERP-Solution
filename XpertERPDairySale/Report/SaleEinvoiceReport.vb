'''''''' Created Report By Shaurya Rajput
Imports common
Imports System.IO
Public Class SaleEinvoiceReport
    Dim EnableProductSaleForJPR As Boolean = False
    Private Sub RmSecurityDeduction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            txtfDate.Value = clsCommon.GETSERVERDATE()
            txtToDate.Value = txtfDate.Value
            ChkBoth.Checked = True
            txtItem.Visible = False
            MyLabel4.Visible = False
            RadGroupBox3.Visible = False
            EnableProductSaleForJPR = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)) = 1, True, False)
            If EnableProductSaleForJPR Then
                RadGroupBox6.Visible = True
            Else
                RadGroupBox6.Visible = False
            End If
            reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        reset()
    End Sub
    Private Sub reset()
        chkDCSSale.Checked = False
        chkOnlyDCSSale.Checked = False
        txtMultiCustomer.arrValueMember = Nothing
        TxtRoute.arrValueMember = Nothing
        txtItem.arrValueMember = Nothing
        gvData.DataSource = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
        ChkBoth.Checked = True
        RadGroupBox3.Visible = False
        rbtnMilk.Checked = True
        chkBPL.Checked = False
        enableDisable(True)
        BtnBoth.IsChecked = True
    End Sub

    Sub enableDisable(ByVal isBool As Boolean)
        RadGroupBox1.Enabled = isBool
        RadGroupBox2.Enabled = isBool
        RadGroupBox3.Enabled = isBool
        RadGroupBox4.Enabled = isBool
        RadGroupBox5.Enabled = isBool
        RadGroupBox6.Enabled = isBool
        RadGroupBox7.Enabled = isBool
        RadGroupBox8.Enabled = isBool
    End Sub

    'Private Sub txtMultiCustomer__My_Click(sender As Object, e As EventArgs)
    '    Dim qry As String = " select cust_code as [Code], Customer_Name as [Name] from tspl_customer_master "
    '    txtMultiCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", txtMultiCustomer.arrValueMember, txtMultiCustomer.arrDispalyMember)
    'End Sub

    'Private Sub TxtRoute__My_Click(sender As Object, e As EventArgs)
    '    Dim qry As String = "Select TSPL_ROUTE_MASTER.Route_No AS Code,TSPL_ROUTE_MASTER.Route_Desc as Name from TSPL_ROUTE_MASTER  where 1=1  "
    '    TxtRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("RouteMulSel", qry, "Code", "Name", TxtRoute.arrValueMember, TxtRoute.arrDispalyMember)
    'End Sub

    Private Sub rmenuPDF_Click(sender As Object, e As EventArgs) Handles rmenuPDF.Click
        ExportToExcel(EnumExportTo.PDF)
    End Sub

    Private Sub rmExcel_Click(sender As Object, e As EventArgs) Handles rmenuExport.Click
        If gvData.Rows.Count > 0 Then
            ExportToExcel(EnumExportTo.Excel)
        Else
            RadMessageBox.Show("No Data Found to Display", Me.Text)
        End If
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
                clsCommon.MyExportToExcelGrid("Sale EInvoice Report", gvData, arrHeader, Me.Text)
            Else
                clsCommon.MyExportToPDF("Sale EInvoice Report", gvData, arrHeader, "Sale EInvoice Report", PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub

    Sub GetReportGridID()
        Dim VarID As String = ""
        If rbtnDetail.IsChecked Then
            VarID += "_DE"
        ElseIf rbtnSummary.IsChecked Then
            VarID += "_Su"
        End If
        If rbtnDocumentdate.IsChecked Then
            VarID += "_DD"
        ElseIf rbtnSupplydate.IsChecked Then
            VarID += "_SD"
        End If
        If BtnMorning.IsChecked Then
            VarID += "_MO"
        ElseIf BtnEvening.IsChecked Then
            VarID += "_EV"
        ElseIf BtnBoth.IsChecked Then
            VarID += "_BO"
        End If
        If chkB2C.Checked = True Then
            VarID += "_BC"
        ElseIf ChkB2B.Checked = True Then
            VarID += "_BB"
        ElseIf ChkBoth.Checked = True Then
            VarID += "_BO"
        End If
        gvData.VarID = VarID
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        PageSetupReport_ID = MyBase.Form_ID
        TemplateGridview = gvData
        LoadData()
        'ControlEnableDisable(False)
    End Sub

    Public Sub LoadData()
        Try
            Dim Whr As String = ""
            Dim dt As DataTable = Nothing
            Dim strtxtfDate As String = clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy")
            Dim strToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
            Dim qry As String = ""
            Dim whrDCSSale As String = ""
            Dim whrclsDate As String = ""

            If rbtnDocumentdate.IsChecked Then
                whrclsDate += " where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) >= Convert( Date,'" + strtxtfDate + "',103) AND 
                                Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) <= Convert(Date,'" + strToDate + "',103) and TSPL_SD_SALE_INVOICE_HEAD.Status='1' "
                whrDCSSale += " where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) >= Convert( Date,'" + strtxtfDate + "',103) AND 
                                Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) <= Convert(Date,'" + strToDate + "',103) and TSPL_SD_SALE_INVOICE_HEAD.Status='1' "
            Else
                whrDCSSale += " where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) >= Convert( Date,'" + strtxtfDate + "',103) AND 
                                Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) <= Convert(Date,'" + strToDate + "',103) and TSPL_SD_SALE_INVOICE_HEAD.Status='1' "
                whrclsDate += " where Convert( Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert( Date,'" + strtxtfDate + "',103) AND 
                                Convert( Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date,'" + strToDate + "',103) and TSPL_SD_SALE_INVOICE_HEAD.Status='1' "
            End If
            If BtnMorning.IsChecked Then
                whrDCSSale = " And TSPL_SD_SHIPMENT_HEAD.Shift_Type in( 'AM','')"
                whrclsDate += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' "
            ElseIf BtnEvening.IsChecked Then
                whrDCSSale += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type in('PM','') "
                whrclsDate += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'PM' "
            End If

            If ChkB2B.Checked Then
                Whr += " and TSPL_CUSTOMER_MASTER.GST_Registered=1 "
            ElseIf chkB2C.Checked Then
                Whr += " and TSPL_CUSTOMER_MASTER.GST_Registered=0 "
            End If
            If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
                Whr += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
            End If


            If chkBPL.Checked Then
                Whr += " and TSPL_BOOKING_MATSER.Is_BPL=1  "
            ElseIf chkAPS.Checked Then
                Whr += " and TSPL_BOOKING_MATSER.Is_APS=1 or TSPL_SD_SHIPMENT_HEAD.Screen_Type ='CT' "
            Else
                If TxtRoute.arrValueMember IsNot Nothing AndAlso TxtRoute.arrValueMember.Count > 0 Then
                    If chkDCSSale.Checked OrElse chkOnlyDCSSale.Checked Then
                        Whr += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in(" + clsCommon.GetMulcallString(TxtRoute.arrValueMember) + ",'')" + Environment.NewLine
                    Else
                        Whr += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in(" + clsCommon.GetMulcallString(TxtRoute.arrValueMember) + ")" + Environment.NewLine
                    End If
                End If

                If txtItem.arrValueMember IsNot Nothing AndAlso txtItem.arrValueMember.Count > 0 Then
                    Whr += " and TSPL_ITEM_MASTER.Item_Code in(" + clsCommon.GetMulcallString(txtItem.arrValueMember) + ")" + Environment.NewLine
                End If
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
            Dim Baseqry As String = ""
            If rbtnSummary.IsChecked Then
                GetReportGridID()
                qry = "SELECT  Max(TSPL_SD_SHIPMENT_HEAD.Document_Code) As [Dispatch Code],Convert(Varchar(10),Max(TSPL_SD_SHIPMENT_HEAD.Document_Date),103) As [Dispatch Date],
Max(TSPL_DCS_SALE_ENTRY_detail.DOCUMENT_CODE) As [DCS Sale Entry],max(CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD.Supply_Date, 103)) as [Supply Date],
                           CASE WHEN max(TSPL_SD_SHIPMENT_HEAD.Shift_Type) = 'AM' THEN 'Morning' WHEN max(TSPL_SD_SHIPMENT_HEAD.Shift_Type) = 'PM' THEN 'Evening'  end as [Shift Type],
                           max(TSPL_SD_SHIPMENT_HEAD.Bill_To_Location) AS [Location],
                           max(TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code) AS [Sub Location],
                           max(TSPL_LOCATION_MASTER.GSTNO) AS [GST No],
                           max(TSPL_STATE_MASTER.GST_STATE_Code) AS [State Code],
                           max(TSPL_CUSTOMER_MASTER.Cust_Code) AS [Customer Code],
                           max(TSPL_CUSTOMER_MASTER.Customer_Name) AS [Customer Name],"
                If chkDCSSale.Checked OrElse chkOnlyDCSSale.Checked Then
                    qry += " max(tspl_vlc_master_head.VLC_Code_VLC_Uploader) as [DCS Uploader Code], "
                End If

                qry += " maX(TSPL_STATE_MASTER.GST_STATE_Code) AS [Party State],
                           max(TSPL_CUSTOMER_MASTER.GSTNO) AS [Recipient Gst No],
                           max(TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type) as [E Invoice Type],"
                If EnableProductSaleForJPR Then
                    qry += "max(case when  TSPL_SD_SALE_INVOICE_HEAD.item_type = 'M' then 'Milk' when TSPL_SD_SALE_INVOICE_HEAD.item_type = 'P' then 'Product' when TSPL_SD_SALE_INVOICE_HEAD.item_type = 'I' then 'Ice Cream' end  )  as [Item Type],"
                End If
                qry += " max(Ack_No) AS [Ack No],
                           max(CONVERT(varchar,Ack_Date, 103)) AS [Ack Date],
                           max(IRN_No) AS [IRN No],
                           TSPL_SD_SALE_INVOICE_HEAD.Document_Code AS [Invoice No],
                           max(CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date, 103)) AS [Invoice Date],
                           CASE WHEN max(TSPL_SD_SHIPMENT_HEAD.DO_Item_Type) = 'T' THEN 'Taxable'  
                           WHEN max(TSPL_SD_SHIPMENT_HEAD.DO_Item_Type) ='NT' THEN 'Non-Taxable' 
                           END AS [Invoice Type],
                           max(TSPL_SD_SALE_INVOICE_HEAD.Total_Amt) AS [Invoice Value],Max(isnull(TSPL_ROUTE_MASTER.Zone_Code,''))as Zone_Code,
                           max(TSPL_SD_SALE_INVOICE_HEAD.Route_No) as [Route No],
                           max(TSPL_SD_SALE_INVOICE_HEAD.EwayBillNo) AS [EwayBillNo],
                           max(TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate) AS [EwayBillDate],
                           sum(TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount) As [Base Amount],
                              CASE WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX1)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX2)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX3)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX4)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX5)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX6)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX7)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX8)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX9)='KKF'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX10)='KKF' THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt) else 0 END 
    				AS [KKF Amt],

                          CASE WHEN max(tax1.Type)='M'  THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt)
    				WHEN max(tax2.Type)='M'  THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt)
    				WHEN max(tax3.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt)
    				WHEN max(tax4.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt)
    				WHEN max(tax5.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt)
    				WHEN max(tax6.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt)
    				WHEN max(tax7.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt)
    				WHEN max(tax8.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt)
    				WHEN max(tax9.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt)
    				WHEN max(tax10.Type)='M' THEN Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt) else 0 END AS [Mandi Tax Amt],

                    CASE WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX1)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX2)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX3)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX4)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX5)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX6)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX7)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX8)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX9)='CGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX10)='CGST' THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt) else 0 END  AS [CGST Amt],

                             CASE WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX1)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX2)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX3)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX4)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX5)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX6)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX7)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX8)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX9)='SGST'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX10)='SGST' THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt) else 0 END  AS [SGST Amt],CASE WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX1)='IGST'  THEN isnull(Sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX2)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX3)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX4)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX5)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX6)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX7)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX8)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX9)='IGST'  THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt),0)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX10)='IGST' THEN isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt),0) else 0 END AS [IGST Amt],
                    CASE WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX1)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX2)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX3)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX4)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX5)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX6)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX7)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX8)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX9)='TCS'  THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt)
    				WHEN max(TSPL_SD_SALE_INVOICE_HEAD.TAX10)='TCS' THEN sum(TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt) else 0 END  AS [TCS Amt],
Sum(TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt) As [Total Tax Amount],
                                Sum(TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt) as [Total Amount]
                           FROM TSPL_SD_SALE_INVOICE_HEAD
                           LEFT OUTER JOIN TSPL_LOCATION_MASTER 
                           ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
                           LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_DETAIL 
                           ON TSPL_SD_SALE_INVOICE_DETAIL.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE
                           
  left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD.tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX10 
                           LEFT OUTER JOIN TSPL_CUSTOMER_MASTER 
                           ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                           LEFT OUTER JOIN TSPL_STATE_MASTER 
                           ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
                           left join tspl_item_master on tspl_item_master.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code "
                 If chkDCSSale.Checked OrElse chkOnlyDCSSale.Checked Then
                    qry += " left outer join tspl_vlc_master_head on tspl_vlc_master_head.vsp_code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code "
                End If
                qry += "    left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
                           left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
						   Left outer Join TSPL_SD_SHIPMENT_DETAIL On TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_Detail.Shipment_Code  And   TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_SD_SALE_INVOICE_Detail.Item_Code
						  Left Join TSPL_DCS_SALE_ENTRY_detail On TSPL_DCS_SALE_ENTRY_detail.PK_Id=TSPL_SD_SHIPMENT_DETAIL.REF_PK_ID "
                If chkBPL.Checked OrElse chkAPS.Checked Then
                    qry += " left outer join TSPL_BOOKING_MATSER on TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                End If
                Baseqry = qry
                qry += "" + whrclsDate + " " + Whr + "  And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' group by TSPL_SD_SALE_INVOICE_HEAD.Document_Code  "
            ElseIf rbtnDetail.IsChecked Then
                qry = "SELECT  TSPL_SD_SHIPMENT_HEAD.Document_Code As [Dispatch Code],Convert(Varchar(10),TSPL_SD_SHIPMENT_HEAD.Document_Date,103) As [Dispatch Date],
TSPL_DCS_SALE_ENTRY_detail.DOCUMENT_CODE As [DCS Sale Entry],CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD.Supply_Date, 103) as [Supply Date],
                                CASE WHEN TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' THEN 'Morning' WHEN TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'PM' THEN 'Evening'  end as [Shift Type], 
                                TSPL_SD_SHIPMENT_HEAD.Bill_To_Location AS [Location],
                                TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code AS [Sub Location],
                                TSPL_LOCATION_MASTER.GSTNO AS [GST No],
                                TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
                                TSPL_CUSTOMER_MASTER.Cust_Code AS [Customer Code],
                                TSPL_CUSTOMER_MASTER.Customer_Name AS [Customer Name],"
                If chkDCSSale.Checked OrElse chkOnlyDCSSale.Checked Then
                    qry += " tspl_vlc_master_head.VLC_Code_VLC_Uploader as [DCS Uploader Code], "
                End If
                qry += " TSPL_STATE_MASTER.GST_STATE_Code AS [Party State],
								TSPL_CUSTOMER_MASTER.GSTNO AS [Recipient Gst No],
                                TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type as [E Invoice Type],"
                If EnableProductSaleForJPR Then
                    qry += "case when  TSPL_SD_SALE_INVOICE_HEAD.item_type = 'S' then 'Milk' when TSPL_SD_SALE_INVOICE_HEAD.item_type = 'P' then 'Product' when TSPL_SD_SALE_INVOICE_HEAD.item_type = 'I' then 'Ice Cream'  end as [Item Type],"
                End If
                qry += " Ack_No AS [Ack No],
                                CONVERT(varchar,Ack_Date, 103) AS [Ack Date],
								IRN_No AS [IRN No],
                                TSPL_SD_SALE_INVOICE_HEAD.Document_Code AS [Invoice No],
                                CONVERT(varchar,TSPL_SD_SHIPMENT_HEAD.Document_Date, 103) AS [Invoice Date],
                                CASE WHEN TSPL_SD_SHIPMENT_HEAD.DO_Item_Type = 'T' THEN 'Taxable'WHEN TSPL_SD_SHIPMENT_HEAD.DO_Item_Type = 'NT' THEN 'Non-Taxable' 
                                END AS [Invoice Type],isnull(TSPL_ROUTE_MASTER.Zone_Code,'')as Zone_Code,
                                TSPL_SD_SALE_INVOICE_HEAD.Route_No as [Route No],
                                tspl_item_master.Item_Code as [Item Code],
                                tspl_item_master.Item_Desc as [Item Name],"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    qry += " TSPL_SD_SALE_INVOICE_DETAIL.Billing_Unit_code as [Billing UOM],
  TSPL_SD_SALE_INVOICE_DETAIL.Billing_Qty as [Billing Qty],
   case when isnull(TSPL_SD_SALE_INVOICE_DETAIL.Billing_Qty,0)>0 then  Convert(decimal(18, 2),(TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount/  TSPL_SD_SALE_INVOICE_DETAIL.Billing_Qty)) else 0 end as [Billing Rate], "
                End If

                qry += "TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as [UOM],
								TSPL_SD_SALE_INVOICE_DETAIL.Qty as [Qty],
                                TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount AS [Item Amount],
                                tspl_item_master.HSN_Code as [HSN Code],                               
                                TSPL_SD_SALE_INVOICE_HEAD.EwayBillNo AS [EwayBillNo],
                                TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate AS [EwayBillDate], "
                If chkBPL.Checked Then
                    qry += " TSPL_BOOKING_MATSER.BPL_Coupon_Code,CONVERT(VARCHAR(10),TSPL_BOOKING_MATSER.BPL_Coupon_Date,103) AS BPL_Coupon_Date ,TSPL_BOOKING_MATSER.BPL_Name,TSPL_BOOKING_MATSER.BPL_Remark,TSPL_BOOKING_MATSER.BPL_Category,"
                End If
                qry += " Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate 
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate else 0  END) AS [KKF %],

                    CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='KKF'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='KKF' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END 
    				AS [KKF Amt],

					Convert(decimal(18,2),CASE WHEN tax1.Type='M'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate
    				WHEN tax2.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate
    				WHEN tax3.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate
    				WHEN tax4.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate
    				WHEN tax5.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate
    				WHEN tax6.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate
    				WHEN tax7.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate
    				WHEN tax8.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate
    				WHEN tax9.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate
    				WHEN tax10.Type='M' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate  else 0 END) AS [Mandi Tax %],

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
				
    				Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate else 0 END) AS [CGST %],

    				CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='CGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='CGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END AS [CGST Amt],

					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate  else 0 END) AS [IGST %],

    				CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='IGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='IGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt  else 0 END AS [IGST Amt],

					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate else 0 END) AS [SGST %],

    				CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='SGST'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='SGST' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt else 0 END AS [SGST Amt],

					Convert(decimal(18,2),CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate else 0 END) AS [TCS %],

    				CASE WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX1='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX2='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX3='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX4='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX5='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX6='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX7='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX8='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX9='TCS'  THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt
    				WHEN TSPL_SD_SALE_INVOICE_HEAD.TAX10='TCS' THEN TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt  else 0 END AS [TCS Amt],
TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt As [Total Tax Amount],
                                TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt as [Total Amount]
                                FROM TSPL_SD_SALE_INVOICE_HEAD
                                LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
                                LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_DETAIL ON TSPL_SD_SALE_INVOICE_DETAIL.Document_Code =TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE
                                left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_SD_SALE_INVOICE_HEAD.tax1
            left outer join tspl_tax_master as tax2 on tax2.tax_code = TSPL_SD_SALE_INVOICE_HEAD.tax2  
             left outer join tspl_tax_master as tax3 on tax3.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.TAX3  
            left outer join TSPL_TAX_MASTER as tax4 on tax4.Tax_Code= TSPL_SD_SALE_INVOICE_HEAD.tax4  
            left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code=TSPL_SD_SALE_INVOICE_HEAD.tax5  
            left outer join TSPL_TAX_MASTER as tax6 on tax6.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX6  
             left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX7 
              left outer join TSPL_TAX_MASTER as tax8 on tax8.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX8
            left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX9 
              left outer join TSPL_TAX_MASTER as tax10 on tax10.Tax_Code =TSPL_SD_SALE_INVOICE_HEAD.TAX10 

                                LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                                LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
                                left join tspl_item_master on tspl_item_master.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code "
                If chkDCSSale.Checked OrElse chkOnlyDCSSale.Checked Then
                    qry += " left outer join tspl_vlc_master_head on tspl_vlc_master_head.vsp_code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code "
                End If
                qry += "  left outer join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
                                left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
						  Left outer Join TSPL_SD_SHIPMENT_DETAIL On TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE=TSPL_SD_SALE_INVOICE_Detail.Shipment_Code  And   TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_SD_SALE_INVOICE_Detail.Item_Code
						  Left Join TSPL_DCS_SALE_ENTRY_detail On TSPL_DCS_SALE_ENTRY_detail.PK_Id=TSPL_SD_SHIPMENT_DETAIL.REF_PK_ID "
                If chkBPL.Checked OrElse chkAPS.Checked Then
                    qry += " left outer join TSPL_BOOKING_MATSER on TSPL_BOOKING_MATSER.Document_No=TSPL_SD_SHIPMENT_HEAD.Against_Booking_No "
                End If

                Baseqry = qry
                qry += " " + whrclsDate + " " + Whr + " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' " + IIf(chkAPS.Checked = False, " and TSPL_SD_SALE_INVOICE_HEAD.Screen_Type <>'CT' ", "")
            End If
            If chkDCSSale.Checked Then
                qry = "" + qry + " " + Environment.NewLine + " union all " + Environment.NewLine + " " + Baseqry + whrDCSSale + Whr + " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type = 'MCC'"
                If rbtnSummary.IsChecked Then
                    qry += "  group by TSPL_SD_SALE_INVOICE_HEAD.Document_Code "
                End If
            End If
            If chkOnlyDCSSale.Checked Then
                qry = Baseqry + whrDCSSale + Whr + " And TSPL_SD_SALE_INVOICE_HEAD.Trans_Type = 'MCC' "
                If rbtnSummary.IsChecked Then
                    qry += "  group by TSPL_SD_SALE_INVOICE_HEAD.Document_Code "
                End If
                qry += " ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Code"
            End If

            dt = clsDBFuncationality.GetDataTable(qry)
            gvData.GroupDescriptors.Clear()
            gvData.MasterTemplate.SummaryRowsBottom.Clear()
            gvData.DataSource = Nothing
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            Else
                RadPageView1.SelectedPage = RadPageViewPage2
                gvData.GroupDescriptors.Clear()
                gvData.MasterTemplate.SummaryRowsBottom.Clear()
                gvData.DataSource = dt
                gvData.AutoExpandGroups = True
                gvData.ShowGroupPanel = True
                gvData.ShowRowHeaderColumn = False
                gvData.AllowAddNewRow = False
                gvData.AllowDeleteRow = False
                gvData.EnableFiltering = True
                gvData.ShowFilteringRow = True
                SetGridFormat()
                SetGridFormationOFGV1()
                gvData.BestFitColumns()
                enableDisable(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        ReStoreGridLayout()
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gvData.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gvData.Columns.Count - 1 Step ii + 1
                        gvData.Columns(ii).IsVisible = False
                        gvData.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gvData.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Sub SetGridFormat()
        Try
            gvData.AutoExpandGroups = True
            gvData.ShowGroupPanel = True
            gvData.ShowRowHeaderColumn = False
            gvData.AllowAddNewRow = False
            gvData.AllowDeleteRow = False
            gvData.EnableFiltering = True
            gvData.ShowFilteringRow = True
            For ii As Integer = 0 To gvData.Columns.Count - 1
                gvData.Columns(ii).ReadOnly = True
                gvData.Columns(ii).BestFit()
            Next
            If rbtnSummary.IsChecked Then
                gvData.Columns("Dispatch Code").HeaderText = "Dispatch Code"
                gvData.Columns("Dispatch Code").Width = 100
                gvData.Columns("Dispatch Code").IsVisible = False

                gvData.Columns("Dispatch Date").HeaderText = "Dispatch Date"
                gvData.Columns("Dispatch Date").Width = 100
                gvData.Columns("Dispatch Date").IsVisible = False

                gvData.Columns("DCS Sale Entry").HeaderText = "DCS Sale Entry"
                gvData.Columns("DCS Sale Entry").Width = 100
                gvData.Columns("DCS Sale Entry").IsVisible = False

                gvData.Columns("Supply Date").HeaderText = "Supply Date"
                gvData.Columns("Supply Date").Width = 100
                gvData.Columns("Supply Date").IsVisible = True

                gvData.Columns("Shift Type").HeaderText = "Shift Type"
                gvData.Columns("Shift Type").Width = 100
                gvData.Columns("Shift Type").IsVisible = True

                gvData.Columns("Location").HeaderText = "Location"
                gvData.Columns("Location").Width = 100
                gvData.Columns("Location").IsVisible = True

                gvData.Columns("Sub Location").HeaderText = "Sub Location"
                gvData.Columns("Sub Location").Width = 100
                gvData.Columns("Sub Location").IsVisible = True

                gvData.Columns("GST No").HeaderText = "GST No"
                gvData.Columns("GST No").Width = 100
                gvData.Columns("GST No").IsVisible = True

                gvData.Columns("State Code").HeaderText = "State Code"
                gvData.Columns("State Code").Width = 100
                gvData.Columns("State Code").IsVisible = True

                gvData.Columns("Customer Code").HeaderText = "Customer Code"
                gvData.Columns("Customer Code").Width = 100
                gvData.Columns("Customer Code").IsVisible = True

                gvData.Columns("Customer Name").HeaderText = "Customer Name"
                gvData.Columns("Customer Name").Width = 150
                gvData.Columns("Customer Name").IsVisible = True

                gvData.Columns("Party State").HeaderText = "Party State"
                gvData.Columns("Party State").Width = 100
                gvData.Columns("Party State").IsVisible = True

                gvData.Columns("Recipient Gst No").HeaderText = "Recipient Gst No"
                gvData.Columns("Recipient Gst No").Width = 100
                gvData.Columns("Recipient Gst No").IsVisible = True

                gvData.Columns("E Invoice Type").HeaderText = "E Invoice Type"
                gvData.Columns("E Invoice Type").Width = 100
                gvData.Columns("E Invoice Type").IsVisible = True

                gvData.Columns("Ack No").HeaderText = "Ack No"
                gvData.Columns("Ack No").Width = 100
                gvData.Columns("Ack No").IsVisible = True

                gvData.Columns("Ack Date").HeaderText = "Ack Date"
                gvData.Columns("Ack Date").Width = 100
                gvData.Columns("Ack Date").IsVisible = True

                gvData.Columns("IRN No").HeaderText = "IRN No"
                gvData.Columns("IRN No").Width = 100
                gvData.Columns("IRN No").IsVisible = True

                gvData.Columns("Invoice No").HeaderText = "Invoice No"
                gvData.Columns("Invoice No").Width = 100
                gvData.Columns("Invoice No").IsVisible = True

                gvData.Columns("Invoice Date").HeaderText = "Invoice Date"
                gvData.Columns("Invoice Date").Width = 100
                gvData.Columns("Invoice Date").IsVisible = True

                gvData.Columns("Invoice Type").HeaderText = "Invoice Type"
                gvData.Columns("Invoice Type").Width = 100
                gvData.Columns("Invoice Type").IsVisible = True

                gvData.Columns("Invoice Value").HeaderText = "Invoice Value"
                gvData.Columns("Invoice Value").Width = 100
                gvData.Columns("Invoice Value").IsVisible = True
                gvData.Columns("Zone_Code").HeaderText = "Zone Code"
                gvData.Columns("Zone_Code").Width = 100
                gvData.Columns("Zone_Code").IsVisible = True
                gvData.Columns("Route No").HeaderText = "Route No"
                gvData.Columns("Route No").Width = 100
                gvData.Columns("Route No").IsVisible = True

                'gvData.Columns("Invoice Value").HeaderText = "Invoice Value"
                'gvData.Columns("Invoice Value").Width = 100
                'gvData.Columns("Invoice Value").IsVisible = True

                gvData.Columns("EwayBillNo").HeaderText = "EwayBillNo"
                gvData.Columns("EwayBillNo").Width = 100
                gvData.Columns("EwayBillNo").IsVisible = True

                gvData.Columns("EwayBillDate").HeaderText = "EwayBillDate"
                gvData.Columns("EwayBillDate").Width = 100
                gvData.Columns("EwayBillDate").IsVisible = True

                gvData.Columns("Base Amount").HeaderText = "Base Amount"
                gvData.Columns("Base Amount").Width = 100
                gvData.Columns("Base Amount").IsVisible = True

                gvData.Columns("KKF Amt").HeaderText = "KKF Amt"
                gvData.Columns("KKF Amt").Width = 100
                gvData.Columns("KKF Amt").IsVisible = True

                gvData.Columns("Mandi Tax Amt").HeaderText = "Mandi Tax Amt"
                gvData.Columns("Mandi Tax Amt").Width = 100
                gvData.Columns("Mandi Tax Amt").IsVisible = True

                gvData.Columns("CGST Amt").HeaderText = "CGST Amt"
                gvData.Columns("CGST Amt").Width = 100
                gvData.Columns("CGST Amt").IsVisible = True

                gvData.Columns("SGST Amt").HeaderText = "SGST Amt"
                gvData.Columns("SGST Amt").Width = 100
                gvData.Columns("SGST Amt").IsVisible = True

                gvData.Columns("IGST Amt").HeaderText = "IGST Amt"
                gvData.Columns("IGST Amt").Width = 100
                gvData.Columns("IGST Amt").IsVisible = True

                gvData.Columns("TCS Amt").HeaderText = "TCS Amt"
                gvData.Columns("TCS Amt").Width = 100
                gvData.Columns("TCS Amt").IsVisible = True

                gvData.Columns("Total Tax Amount").HeaderText = "Total Tax Amount"
                gvData.Columns("Total Tax Amount").Width = 100
                gvData.Columns("Total Tax Amount").IsVisible = True

                gvData.Columns("Total Amount").HeaderText = "Total Amount"
                gvData.Columns("Total Amount").Width = 100
                gvData.Columns("Total Amount").IsVisible = True

                'Dim summaryRowItem As New GridViewSummaryRowItem()
                'Dim item1 As New GridViewSummaryItem("INWARDQTYReportUom", "{0:n2}", GridAggregateFunction.Sum)
                'summaryRowItem.Add(item1)
                'Dim item2 As New GridViewSummaryItem("OUTWARDQTYReportUom", "{0:n2}", GridAggregateFunction.Sum)
                'summaryRowItem.Add(item2)
                'gvData.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                'gvData.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom

            ElseIf rbtnDetail.IsChecked Then
                gvData.Columns("Dispatch Code").HeaderText = "Dispatch Code"
                gvData.Columns("Dispatch Code").Width = 100
                gvData.Columns("Dispatch Code").IsVisible = False

                gvData.Columns("Dispatch Date").HeaderText = "Dispatch Date"
                gvData.Columns("Dispatch Date").Width = 100
                gvData.Columns("Dispatch Date").IsVisible = False

                gvData.Columns("DCS Sale Entry").HeaderText = "DCS Sale Entry"
                gvData.Columns("DCS Sale Entry").Width = 100
                gvData.Columns("DCS Sale Entry").IsVisible = False

                gvData.Columns("Supply Date").HeaderText = "Supply Date"
                gvData.Columns("Supply Date").Width = 100
                gvData.Columns("Supply Date").IsVisible = True

                gvData.Columns("Shift Type").HeaderText = "Shift Type"
                gvData.Columns("Shift Type").Width = 100
                gvData.Columns("Shift Type").IsVisible = True

                gvData.Columns("Location").HeaderText = "Location"
                gvData.Columns("Location").Width = 100
                gvData.Columns("Location").IsVisible = True

                gvData.Columns("Sub Location").HeaderText = "Sub Location"
                gvData.Columns("Sub Location").Width = 100
                gvData.Columns("Sub Location").IsVisible = True

                gvData.Columns("GST No").HeaderText = "GST No"
                gvData.Columns("GST No").Width = 100
                gvData.Columns("GST No").IsVisible = True

                gvData.Columns("State Code").HeaderText = "State Code"
                gvData.Columns("State Code").Width = 100
                gvData.Columns("State Code").IsVisible = True

                gvData.Columns("Customer Code").HeaderText = "Customer Code"
                gvData.Columns("Customer Code").Width = 100
                gvData.Columns("Customer Code").IsVisible = True

                gvData.Columns("Customer Name").HeaderText = "Customer Name"
                gvData.Columns("Customer Name").Width = 100
                gvData.Columns("Customer Name").IsVisible = True

                gvData.Columns("Party State").HeaderText = "Party State"
                gvData.Columns("Party State").Width = 100
                gvData.Columns("Party State").IsVisible = True

                gvData.Columns("Recipient Gst No").HeaderText = "Recipient Gst No"
                gvData.Columns("Recipient Gst No").Width = 100
                gvData.Columns("Recipient Gst No").IsVisible = True

                gvData.Columns("E Invoice Type").HeaderText = "E Invoice Type"
                gvData.Columns("E Invoice Type").Width = 100
                gvData.Columns("E Invoice Type").IsVisible = True

                gvData.Columns("Ack No").HeaderText = "Ack No"
                gvData.Columns("Ack No").Width = 100
                gvData.Columns("Ack No").IsVisible = True

                gvData.Columns("Ack Date").HeaderText = "Ack Date"
                gvData.Columns("Ack Date").Width = 100
                gvData.Columns("Ack Date").IsVisible = True

                gvData.Columns("IRN No").HeaderText = "IRN No"
                gvData.Columns("IRN No").Width = 100
                gvData.Columns("IRN No").IsVisible = True

                gvData.Columns("Invoice No").HeaderText = "Invoice No"
                gvData.Columns("Invoice No").Width = 100
                gvData.Columns("Invoice No").IsVisible = True

                gvData.Columns("Invoice Date").HeaderText = "Invoice Date"
                gvData.Columns("Invoice Date").Width = 100
                gvData.Columns("Invoice Date").IsVisible = True

                gvData.Columns("Invoice Type").HeaderText = "Invoice Type"
                gvData.Columns("Invoice Type").Width = 100
                gvData.Columns("Invoice Type").IsVisible = True
                gvData.Columns("Zone_Code").HeaderText = "Zone Code"
                gvData.Columns("Zone_Code").Width = 100
                gvData.Columns("Zone_Code").IsVisible = True

                gvData.Columns("Route No").HeaderText = "Route No"
                gvData.Columns("Route No").Width = 100
                gvData.Columns("Route No").IsVisible = True

                gvData.Columns("Item Code").HeaderText = "Item Code"
                gvData.Columns("Item Code").Width = 100
                gvData.Columns("Item Code").IsVisible = True

                gvData.Columns("Item Name").HeaderText = "Item Name"
                gvData.Columns("Item Name").Width = 100
                gvData.Columns("Item Name").IsVisible = True
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    gvData.Columns("Billing UOM").HeaderText = "Billing UOM"
                    gvData.Columns("Billing UOM").Width = 100
                    gvData.Columns("Billing UOM").IsVisible = True

                    gvData.Columns("Billing Qty").HeaderText = "Billing Qty"
                    gvData.Columns("Billing Qty").Width = 100
                    gvData.Columns("Billing Qty").IsVisible = True

                    gvData.Columns("Billing Rate").HeaderText = "Billing Rate"
                    gvData.Columns("Billing Rate").Width = 100
                    gvData.Columns("Billing Rate").IsVisible = True
                End If

                gvData.Columns("UOM").HeaderText = "UOM"
                gvData.Columns("UOM").Width = 100
                gvData.Columns("UOM").IsVisible = True

                gvData.Columns("Qty").HeaderText = "Qty"
                gvData.Columns("Qty").Width = 100
                gvData.Columns("Qty").IsVisible = True

                gvData.Columns("Item Amount").HeaderText = "Item Amount"
                gvData.Columns("Item Amount").Width = 100
                gvData.Columns("Item Amount").IsVisible = True

                gvData.Columns("HSN Code").HeaderText = "HSN Code"
                gvData.Columns("HSN Code").Width = 100
                gvData.Columns("HSN Code").IsVisible = True

                gvData.Columns("EwayBillNo").HeaderText = "EwayBillNo"
                gvData.Columns("EwayBillNo").Width = 100
                gvData.Columns("EwayBillNo").IsVisible = True

                gvData.Columns("EwayBillDate").HeaderText = "EwayBillDate"
                gvData.Columns("EwayBillDate").Width = 100
                gvData.Columns("EwayBillDate").IsVisible = True
                If chkBPL.Checked Then
                    gvData.Columns("BPL_Coupon_Code").HeaderText = "BPL Coupon Code"
                    gvData.Columns("BPL_Coupon_Code").Width = 100
                    gvData.Columns("BPL_Coupon_Code").IsVisible = True

                    gvData.Columns("BPL_Coupon_Date").HeaderText = "BPL Coupon Date"
                    gvData.Columns("BPL_Coupon_Date").Width = 100
                    gvData.Columns("BPL_Coupon_Date").IsVisible = True

                    gvData.Columns("BPL_Name").HeaderText = "BPL Name"
                    gvData.Columns("BPL_Name").Width = 100
                    gvData.Columns("BPL_Name").IsVisible = True

                    gvData.Columns("BPL_Remark").HeaderText = "BPL Remark"
                    gvData.Columns("BPL_Remark").Width = 100
                    gvData.Columns("BPL_Remark").IsVisible = True

                    gvData.Columns("BPL_Category").HeaderText = "BPL Category"
                    gvData.Columns("BPL_Category").Width = 100
                    gvData.Columns("BPL_Category").IsVisible = True
                End If

                gvData.Columns("KKF %").HeaderText = "KKF %"
                gvData.Columns("KKF %").Width = 100
                gvData.Columns("KKF %").IsVisible = True

                gvData.Columns("KKF Amt").HeaderText = "KKF Amt"
                gvData.Columns("KKF Amt").Width = 100
                gvData.Columns("KKF Amt").IsVisible = True

                gvData.Columns("Mandi Tax %").HeaderText = "Mandi Tax %"
                gvData.Columns("Mandi Tax %").Width = 100
                gvData.Columns("Mandi Tax %").IsVisible = True

                gvData.Columns("Mandi Tax Amt").HeaderText = "Mandi Tax Amt"
                gvData.Columns("Mandi Tax Amt").Width = 100
                gvData.Columns("Mandi Tax Amt").IsVisible = True

                gvData.Columns("CGST %").HeaderText = "CGST %"
                gvData.Columns("CGST %").Width = 100
                gvData.Columns("CGST %").IsVisible = True

                gvData.Columns("CGST Amt").HeaderText = "CGST Amt"
                gvData.Columns("CGST Amt").Width = 100
                gvData.Columns("CGST Amt").IsVisible = True

                gvData.Columns("IGST %").HeaderText = "IGST %"
                gvData.Columns("IGST %").Width = 100
                gvData.Columns("IGST %").IsVisible = True

                gvData.Columns("IGST Amt").HeaderText = "IGST Amt"
                gvData.Columns("IGST Amt").Width = 100
                gvData.Columns("IGST Amt").IsVisible = True

                gvData.Columns("SGST %").HeaderText = "SGST %"
                gvData.Columns("SGST %").Width = 100
                gvData.Columns("SGST %").IsVisible = True

                gvData.Columns("SGST Amt").HeaderText = "SGST Amt"
                gvData.Columns("SGST Amt").Width = 100
                gvData.Columns("SGST Amt").IsVisible = True

                gvData.Columns("TCS %").HeaderText = "TCS %"
                gvData.Columns("TCS %").Width = 100
                gvData.Columns("TCS %").IsVisible = True

                gvData.Columns("TCS Amt").HeaderText = "TCS Amt"
                gvData.Columns("TCS Amt").Width = 100
                gvData.Columns("TCS Amt").IsVisible = True

                gvData.Columns("Total Tax Amount").HeaderText = "Total Tax Amount"
                gvData.Columns("Total Tax Amount").Width = 100
                gvData.Columns("Total Tax Amount").IsVisible = True

                gvData.Columns("Total Amount").HeaderText = "Total Amount"
                gvData.Columns("Total Amount").Width = 100
                gvData.Columns("Total Amount").IsVisible = True

                'Dim summaryRowItem As New GridViewSummaryRowItem()
                'Dim item1 As New GridViewSummaryItem("OPBal", "{0:n2}", GridAggregateFunction.Sum)
                'summaryRowItem.Add(item1)
                'Dim item2 As New GridViewSummaryItem("Received_Qty", "{0:n2}", GridAggregateFunction.Sum)
                'summaryRowItem.Add(item2)
                'Dim item3 As New GridViewSummaryItem("Issued_Qty", "{0:n2}", GridAggregateFunction.Sum)
                'summaryRowItem.Add(item3)
                'Dim item4 As New GridViewSummaryItem("Balance_Qty", "{0:n2}", GridAggregateFunction.Sum)
                'summaryRowItem.Add(item4)
                'gvData.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                'gvData.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    'Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
    '    Dim dt As DataTable = Nothing

    '    If rbtnSummary.IsChecked Then

    '        Dim strtxtfDate As String = clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy")
    '        Dim strToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
    '        GetReportGridID()
    '        Dim qry As String = "SELECT DISTINCT 
    '                       Bill_To_Location AS [Location],
    '                       TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code AS [Sub Location],
    '                       TSPL_LOCATION_MASTER.GSTNO AS [GST No],
    '                       TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
    '                       TSPL_CUSTOMER_MASTER.Cust_Code AS [Customer Code],
    '                       TSPL_CUSTOMER_MASTER.Customer_Name AS [Customer Name],
    '                       TSPL_STATE_MASTER.GST_STATE_Code AS [Party State], 
    '                       Ack_No AS [Ack No],
    '                       CONVERT(varchar,Ack_Date, 103) AS [Ack Date],
    '                       TSPL_SD_SALE_INVOICE_HEAD.Document_Code AS [Invoice No],
    '                       CONVERT(varchar,Document_Date, 103) AS [Invoice Date],
    '                       CASE WHEN Is_Taxable = 1 THEN 'Taxable'  
    '                       WHEN Is_Taxable = 0 THEN 'Non-Taxable' 
    '                       END AS [Invoice Type],
    '                       TSPL_SD_SALE_INVOICE_HEAD.Route_No as [Route No],
    '                       TSPL_SD_SALE_INVOICE_HEAD.Total_Amt AS [Invoice Value],
    '                       TSPL_CUSTOMER_MASTER.GSTNO AS [Recipient Gst No], 
    '                       IRN_No AS [IRN No],
    '                       TSPL_SD_SALE_INVOICE_HEAD.EwayBillNo AS [EwayBillNo],
    '                       TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate AS [EwayBillDate],
    '                       TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt AS [KKF],
    '                       TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt AS [Mandi Tax],
    '                       TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt AS [CGST Amt],
    '                       TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt AS [SGST Amt],
    '                       TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt AS [TCS Amt]
    '                       FROM TSPL_SD_SALE_INVOICE_HEAD
    '                       LEFT OUTER JOIN TSPL_LOCATION_MASTER 
    '                       ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
    '                       LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_DETAIL 
    '                       ON TSPL_SD_SALE_INVOICE_DETAIL.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE
    '                       LEFT OUTER JOIN TSPL_TAX_MASTER 
    '                       ON TSPL_TAX_MASTER.Tax_Code = TSPL_SD_SALE_INVOICE_DETAIL.TAX1
    '                       LEFT OUTER JOIN TSPL_CUSTOMER_MASTER 
    '                       ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
    '                       LEFT OUTER JOIN TSPL_STATE_MASTER 
    '                       ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
    '                       where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) >= Convert( Date,'" + strtxtfDate + "',103) AND 
    '                       Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) <= Convert(Date,'" + strToDate + "',103)"

    '        If ChkB2B.Checked = True Then
    '            qry += " and TSPL_CUSTOMER_MASTER.GST_Registered=1 "
    '        ElseIf chkB2C.Checked = True Then
    '            qry += " and TSPL_CUSTOMER_MASTER.GST_Registered=0 "
    '        End If
    '        'qry += "ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Code"

    '        If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
    '            qry += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
    '        End If

    '        If TxtRoute.arrValueMember IsNot Nothing AndAlso TxtRoute.arrValueMember.Count > 0 Then
    '            qry += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in(" + clsCommon.GetMulcallString(TxtRoute.arrValueMember) + ")" + Environment.NewLine
    '        End If

    '        qry += "ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Code"

    '        dt = clsDBFuncationality.GetDataTable(qry)

    '    ElseIf rbtnDetail.IsChecked Then

    '        Dim strtxtfDate As String = clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy")
    '        Dim strToDate As String = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
    '        GetReportGridID()
    '        Dim qry1 As String = "SELECT DISTINCT Bill_To_Location AS [Location],
    '                            TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code AS [Sub Location],
    '                            TSPL_LOCATION_MASTER.GSTNO AS [GST No],
    '                            TSPL_STATE_MASTER.GST_STATE_Code AS [State Code],
    '                            TSPL_CUSTOMER_MASTER.Cust_Code AS [Customer Code],
    '                            TSPL_CUSTOMER_MASTER.Customer_Name AS [Customer Name],
    '                            TSPL_STATE_MASTER.GST_STATE_Code AS [Party State],
    '                            Ack_No AS [Ack No],
    '                            CONVERT(varchar,Ack_Date, 103) AS [Ack Date],
    '                            TSPL_SD_SALE_INVOICE_HEAD.Document_Code AS [Invoice No],
    '                            CONVERT(varchar,Document_Date, 103) AS [Invoice Date],
    '                            CASE WHEN Is_Taxable = 1 THEN 'Taxable'WHEN Is_Taxable = 0 THEN 'Non-Taxable' 
    '                            END AS [Invoice Type],
    '                            TSPL_SD_SALE_INVOICE_HEAD.Route_No as [Route No],
    '                            tspl_item_master.Item_Code as [Item Code],
    '                            tspl_item_master.Item_Desc as [Item Name],
    '                            tspl_item_master.HSN_Code as [HSN Code],
    '                            TSPL_SD_SALE_INVOICE_DETAIL.Amt_Less_Discount AS [Item Amount],
    '                            TSPL_SD_SALE_INVOICE_DETAIL.Qty as [Qty],
    '                            TSPL_SD_SALE_INVOICE_DETAIL.Unit_code as [UOM],
    '                            TSPL_CUSTOMER_MASTER.GSTNO AS [Recipient Gst No], 
    '                            IRN_No AS [IRN No],
    '                            TSPL_SD_SALE_INVOICE_HEAD.EwayBillNo AS [EwayBillNo],
    '                            TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate AS [EwayBillDate],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX1_Amt AS [KKF],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX1_Rate AS [KKF %],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX2_Amt AS [Mandi Tax],   
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX2_Rate as [Mandi Tax %], 
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX3_Amt AS [CGST Amt],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX3_Rate AS [CGST %],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX4_Amt AS [SGST Amt],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX4_Rate AS [SGST %],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX5_Amt AS [TCS Amt],
    '                            TSPL_SD_SALE_INVOICE_HEAD.TAX5_Rate AS [TCS %],
    '                            TSPL_SD_SALE_INVOICE_HEAD.Total_Tax_Amt as [Total Tax Amt],
    '                            TSPL_SD_SALE_INVOICE_HEAD.Total_Amt as [Total Amt]
    '                            FROM TSPL_SD_SALE_INVOICE_HEAD
    '                            LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code = TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location
    '                            LEFT OUTER JOIN TSPL_SD_SALE_INVOICE_DETAIL ON TSPL_SD_SALE_INVOICE_DETAIL.Document_Code =TSPL_SD_SALE_INVOICE_HEAD.DOCUMENT_CODE
    '                            LEFT OUTER JOIN TSPL_TAX_MASTER ON TSPL_TAX_MASTER.Tax_Code = TSPL_SD_SALE_INVOICE_DETAIL.TAX1
    '                            LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
    '                            LEFT OUTER JOIN TSPL_STATE_MASTER ON TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE
    '                            left join tspl_item_master on tspl_item_master.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
    '                            where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) >= Convert( Date,'" + strtxtfDate + "',103) AND 
    '                            Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.document_Date,103) <= Convert(Date,'" + strToDate + "',103) "

    '        If ChkB2B.Checked = True Then
    '            qry1 += " and TSPL_CUSTOMER_MASTER.GST_Registered=1 "
    '        ElseIf chkB2C.Checked = True Then
    '            qry1 += " and TSPL_CUSTOMER_MASTER.GST_Registered=0 "
    '        End If


    '        'If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
    '        '    qry += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
    '        'End If

    '        If txtMultiCustomer.arrValueMember IsNot Nothing AndAlso txtMultiCustomer.arrValueMember.Count > 0 Then
    '            qry1 += " and tspl_customer_master.cust_code in(" + clsCommon.GetMulcallString(txtMultiCustomer.arrValueMember) + ")" + Environment.NewLine
    '        End If

    '        If TxtRoute.arrValueMember IsNot Nothing AndAlso TxtRoute.arrValueMember.Count > 0 Then
    '            qry1 += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in(" + clsCommon.GetMulcallString(TxtRoute.arrValueMember) + ")" + Environment.NewLine
    '        End If
    '        qry1 += "ORDER BY TSPL_SD_SALE_INVOICE_HEAD.Document_Code"

    '        dt = clsDBFuncationality.GetDataTable(qry1)
    '    End If
    '    If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
    '        common.clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
    '        Exit Sub
    '    Else
    '        RadPageView1.SelectedPage = RadPageViewPage2
    '        gvData.GroupDescriptors.Clear()
    '        gvData.MasterTemplate.SummaryRowsBottom.Clear()
    '        gvData.DataSource = dt

    '        'gvData.Columns("Customer_Code").IsVisible = False
    '        'gvData.Columns("Customerqty").IsVisible = False
    '        'gvData.Columns("CAN_Qty").IsVisible = False
    '        'gvData.Columns("CRATE_Qty").IsVisible = False
    '        'gvData.Columns("BOX_Qty").IsVisible = False
    '        'gvData.Columns("CarteQtyLtr").IsVisible = False
    '        'gvData.Columns("CanQtyLtr").IsVisible = False


    '        'SetGridFormationOFGV1()
    '        gvData.AutoExpandGroups = True
    '        gvData.ShowGroupPanel = True
    '        gvData.ShowRowHeaderColumn = False
    '        gvData.AllowAddNewRow = False
    '        gvData.AllowDeleteRow = False
    '        gvData.EnableFiltering = True
    '        gvData.ShowFilteringRow = True
    '        gvData.BestFitColumns()
    '    End If
    'End Sub

    Sub SetGridFormationOFGV1()
        gvData.TableElement.TableHeaderHeight = 40
        gvData.MasterTemplate.ShowRowHeaderColumn = False
        For ii As Integer = 0 To gvData.Columns.Count - 1
            gvData.Columns(ii).ReadOnly = True

        Next
        Dim summaryRowItem As New GridViewSummaryRowItem()

        ' gvData.GroupDescriptors.Add(New GridGroupByExpression("Route as RouteName format ""{0}: {1}"" Group By Route"))

        'For i As Integer = 10 To gvData.Columns.Count - 1
        '    Dim aa = gvData.Columns(i).HeaderText()
        '    Dim item8 As New GridViewSummaryItem("[Total Amount]", "{0:F2}", GridAggregateFunction.Sum)
        '    summaryRowItem.Add(item8)

        'Next
        'Dim aa = gvData.Columns(i).HeaderText()
        Dim item82 As New GridViewSummaryItem("TCS Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item82)

        Dim item83 As New GridViewSummaryItem("SGST Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item83)

        Dim item84 As New GridViewSummaryItem("IGST Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item84)

        Dim item85 As New GridViewSummaryItem("CGST Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item85)

        Dim item86 As New GridViewSummaryItem("Mandi Tax Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item86)

        Dim item87 As New GridViewSummaryItem("KKF Amt", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item87)

        Dim item88 As New GridViewSummaryItem("Qty", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item88)

        Dim item89 As New GridViewSummaryItem("Item Amount", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item89)

        Dim item90 As New GridViewSummaryItem("Base Amount", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item90)

        Dim item91 As New GridViewSummaryItem("Invoice Value", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item91)

        Dim item92 As New GridViewSummaryItem("Total Tax Amount", "{0:N2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item92)

        Dim item93 As New GridViewSummaryItem("Total Amount", "{0:N2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item93)

        gvData.ShowGroupPanel = True
        gvData.MasterTemplate.AutoExpandGroups = True

        gvData.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gvData.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        gvData.MasterTemplate.ShowTotals = True
        'ReStoreGridLayout()

    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Close()
    End Sub

    'Private Sub chkB2C_CheckStateChanged(sender As Object, e As EventArgs)
    '    If chkB2C.Checked Then
    '        ChkB2B.Checked = False
    '        ChkBoth.Checked = False
    '    End If
    'End Sub

    'Private Sub ChkB2B_CheckStateChanged(sender As Object, e As EventArgs)
    '    If ChkB2B.Checked Then
    '        chkB2C.Checked = False
    '        ChkBoth.Checked = False
    '    End If
    'End Sub

    'Private Sub ChkBoth_CheckStateChanged(sender As Object, e As EventArgs)
    '    If ChkBoth.Checked Then
    '        ChkB2B.Checked = False
    '        chkB2C.Checked = False
    '    End If
    'End Sub

    'Private Sub txtItem__My_Click(sender As Object, e As EventArgs)
    '    Dim qry As String = " select Item_Code,Item_Desc from TSPL_ITEM_MASTER order by Item_Code "
    '    txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("ItemMulSel", qry, "Item_Code", "Item_Code", txtItem.arrValueMember, txtItem.arrDispalyMember)
    'End Sub

    'Private Sub rbtnDetail_Click(sender As Object, e As EventArgs)
    '    txtItem.Visible = True
    '    MyLabel4.Visible = True
    'End Sub

    'Private Sub rbtnSummary_Click(sender As Object, e As EventArgs)
    '    txtItem.Visible = False
    '    MyLabel4.Visible = False
    'End Sub

    'Private Sub rbtnSupplydate_CheckStateChanged(sender As Object, e As EventArgs)
    '    If rbtnSupplydate.IsChecked Then
    '        RadGroupBox3.Visible = True
    '    Else
    '        RadGroupBox3.Visible = False
    '    End If
    'End Sub

    'Private Sub ReStoreGridLayout()
    '    Try
    '        If clsCommon.myLen(PageSetupReport_ID) > 0 Then
    '            Dim obj As clsGridLayout = New clsGridLayout()
    '            obj = CType(obj.GetData(PageSetupReport_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
    '            If Not obj Is Nothing AndAlso obj.GridColumns >= gvData.ColumnCount Then
    '                Dim ii As Integer
    '                For ii = 0 To gvData.Columns.Count - 1 Step ii + 1
    '                    gvData.Columns(ii).IsVisible = False
    '                    gvData.Columns(ii).VisibleInColumnChooser = True
    '                Next
    '                gvData.LoadLayout(obj.GridLayout)
    '                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
    '            End If
    '        End If
    '    Catch err As Exception
    '        MessageBox.Show(err.Message)
    '    End Try
    'End Sub

    Private Sub rmSaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        If clsCommon.myLen(PageSetupReport_ID) > 0 Then
            gvData.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = PageSetupReport_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gvData.SaveLayout(obj.GridLayout)
            obj.GridColumns = gvData.ColumnCount
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

    Private Sub txtMultiCustomer__My_Click_1(sender As Object, e As EventArgs) Handles txtMultiCustomer._My_Click
        Dim qry As String = " select cust_code as [Code], Customer_Name as [Name] from tspl_customer_master "
        txtMultiCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", txtMultiCustomer.arrValueMember, txtMultiCustomer.arrDispalyMember)
    End Sub

    Private Sub TxtRoute__My_Click_1(sender As Object, e As EventArgs) Handles TxtRoute._My_Click
        Dim qry As String = "Select TSPL_ROUTE_MASTER.Route_No AS Code,TSPL_ROUTE_MASTER.Route_Desc as Name from TSPL_ROUTE_MASTER  where 1=1  "
        TxtRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("RouteMulSel", qry, "Code", "Name", TxtRoute.arrValueMember, TxtRoute.arrDispalyMember)
    End Sub

    Private Sub txtItem__My_Click_1(sender As Object, e As EventArgs) Handles txtItem._My_Click
        Dim qry As String = " select Item_Code,Item_Desc from TSPL_ITEM_MASTER where Item_Type='F' order by Item_Code "
        txtItem.arrValueMember = clsCommon.ShowMultipleSelectForm("ItemMulSel", qry, "Item_Code", "Item_Code", txtItem.arrValueMember, txtItem.arrDispalyMember)
    End Sub

    Private Sub rbtnDetail_Click_1(sender As Object, e As EventArgs) Handles rbtnDetail.Click
        txtItem.Visible = True
        MyLabel4.Visible = True
    End Sub

    Private Sub rbtnSummary_Click_1(sender As Object, e As EventArgs) Handles rbtnSummary.Click
        txtItem.arrValueMember = Nothing
        txtItem.Visible = False
        MyLabel4.Visible = False
    End Sub

    Private Sub rbtnSupplydate_CheckStateChanged_1(sender As Object, e As EventArgs) Handles rbtnSupplydate.CheckStateChanged
        If rbtnSupplydate.IsChecked Then
            RadGroupBox3.Visible = True
        Else
            RadGroupBox3.Visible = False
        End If
    End Sub

    Private Sub ChkB2B_CheckedChanged(sender As Object, e As EventArgs) Handles ChkB2B.CheckedChanged
        If ChkB2B.Checked Then
            chkB2C.Checked = False
            ChkBoth.Checked = False
        End If
    End Sub

    Private Sub chkB2C_CheckedChanged(sender As Object, e As EventArgs) Handles chkB2C.CheckedChanged
        If chkB2C.Checked Then
            ChkB2B.Checked = False
            ChkBoth.Checked = False
        End If
    End Sub

    Private Sub ChkBoth_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBoth.CheckedChanged
        If ChkBoth.Checked Then
            ChkB2B.Checked = False
            chkB2C.Checked = False
        End If
    End Sub

    Private Sub chkBPL_CheckStateChanged(sender As Object, e As EventArgs) Handles chkBPL.CheckStateChanged
        If chkBPL.Checked Then
            rbtnDetail.IsChecked = True
            ChkBoth.Checked = True
            RadGroupBox2.Enabled = False
            RadGroupBox6.Visible = False
            RadGroupBox1.Enabled = False
            chkDCSSale.Visible = False
            chkOnlyDCSSale.Visible = False
            TxtRoute.Visible = False
            txtItem.Visible = False
            MyLabel10.Visible = False
            MyLabel4.Visible = False
            chkAPS.Checked = False
        ElseIf chkAPS.Checked Then
            RadGroupBox2.Enabled = True
            RadGroupBox6.Visible = False
            TxtRoute.Visible = False
            txtItem.Visible = False
        Else
            rbtnDetail.IsChecked = True
            ChkBoth.Checked = True
            RadGroupBox2.Enabled = True
            RadGroupBox6.Visible = True
            RadGroupBox1.Enabled = True
            chkDCSSale.Visible = True
            chkOnlyDCSSale.Visible = True
            TxtRoute.Visible = True
            txtItem.Visible = False
            MyLabel10.Visible = True
            MyLabel4.Visible = False
        End If
    End Sub

    Private Sub chkAPS_CheckedChanged(sender As Object, e As EventArgs) Handles chkAPS.CheckedChanged
        If chkAPS.Checked Then
            rbtnDetail.IsChecked = True
            ChkBoth.Checked = True
            RadGroupBox2.Enabled = True
            RadGroupBox6.Visible = False
            RadGroupBox1.Enabled = False
            chkDCSSale.Visible = False
            chkOnlyDCSSale.Visible = False
            TxtRoute.Visible = False
            txtItem.Visible = False
            MyLabel10.Visible = False
            MyLabel4.Visible = False
            chkBPL.Checked = False
        ElseIf chkBPL.Checked Then
            RadGroupBox2.Enabled = False
            RadGroupBox6.Visible = False
            TxtRoute.Visible = False
            txtItem.Visible = False
        Else
            rbtnDetail.IsChecked = True
            ChkBoth.Checked = True
            RadGroupBox2.Enabled = True
            RadGroupBox6.Visible = True
            RadGroupBox1.Enabled = True
            chkDCSSale.Visible = True
            chkOnlyDCSSale.Visible = True
            TxtRoute.Visible = True
            txtItem.Visible = False
            MyLabel10.Visible = True
            MyLabel4.Visible = False
        End If
    End Sub

    'Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
    '    If clsCommon.myLen(PageSetupReport_ID) < 0 Then
    '        gvData.MasterTemplate.FilterDescriptors.Clear()
    '        Dim obj As New clsGridLayout()
    '        obj.ReportID = PageSetupReport_ID
    '        obj.UserID = objCommonVar.CurrentUserCode
    '        obj.GridLayout = New MemoryStream()
    '        gvData.SaveLayout(obj.GridLayout)
    '        obj.GridColumns = gvData.ColumnCount
    '        obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
    '        If obj.SaveData() Then
    '            common.clsCommon.MyMessageBoxShow("Layout saved successfully", "Information")
    '        End If
    '        obj.GridLayout.Close()
    '        obj.GridLayout.Dispose()
    '    End If
    'End Sub
End Class
