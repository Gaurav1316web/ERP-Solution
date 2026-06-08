Imports System.IO
Imports common
Public Class frmAccountSalesReport
    Private Sub frmAccountSalesReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            txtFromDate.Value = clsCommon.GETSERVERDATE()
            txtToDate.Value = txtFromDate.Value
            cboUOMType.SelectedIndex = 0
            HideAndShow()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtLocation__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocation._MYValidating
        Try
            Dim whrcls As String = ""
            whrcls = " Location_Type='Physical' "
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                whrcls += " and Location_Code in (" & objCommonVar.strCurrUserLocations & ") "
            End If
            Dim qry As String = " select Location_Code as [Code],Location_Desc as [Description],TSPL_Location_MASTER.Loc_Short_Name as [Short Name],Add1,Add2,Add3,Add4,City_Code as [City Code],State,Pin_Code as [Pin Code],Country,Hoadd1 ,Hoadd2,Telphone,Email,Location_Type as [Location Type],Loc_Status as [Location Status],Status_Date as [Status Date],Excisable,Loc_Segment_Code as [Location Segment Code],Seg.Description as [Location Segment Description],Type,Purchase_Tax_Group as [Purchase Tax Group],Sales_Tax_Group as [Sales Tax Group],Ecc_Number as [ECC Number],Registration_Number as [Registration Number],Commissionerate as [Commission Rate],Range_Code as [Range Code],Range_Name as [Range Name],Range_Address as [Range Address],Division_Code as [Division Code],Division_Name as [Division Name],Division_Address as [Division Address],tspl_location_master.Created_By as [Created By],tspl_location_master.Created_Date as [Created Date],tspl_location_master.Modify_By as [Modify By],tspl_location_master.Modify_Date as [Modify Date],tspl_location_master.Comp_code as [Company Code],TIN_No as [TIN No],TAN_No as [TAN No],TCAN_No as [TCAN No],Service_Tax_Reg_No as [Service Tax Registration No],DutyPaid as [Duty Paid],Purchase_Tax_GroupIS as [Purchase Tax Group Inter State],Sales_Tax_GroupIS as [Sales Tax Group Inter State],Stock_Transfer_Filled_Ac as [Stock Transfer Filled Account],Stock_Transfer_Empty_Ac as [Stock Transfer Empty Account],GIT_Location as [GIT Location],GIT_Type as [GIT Type],Rejected_Type as [Rejected Type],Rejected_Location as [Rejected Location],CSA_Type as [CSA Type],Cust_Code as [Cust Code],CST_No as [CST No],Phone1,Phone2,stock_transfer_ac as [Stock Tranfer A/C],Loss_Ac as [Loss A/C] ,Is_Consumption_Location as [Is Consumption Location],Is_Section as [Is Section],Section_Code as [Section Code],Is_Sub_Location as [Is Sub Location],Main_Location_Code as [Main Location Code],IsSubLocationWise as [Is Sub Location Wise] from TSPL_Location_MASTER  left join TSPL_GL_SEGMENT_CODE as Seg on TSPL_Location_MASTER.Loc_Segment_Code=Seg.Segment_Code "
            txtLocation.Value = clsCommon.ShowSelectForm("@Location", qry, "Code", whrcls, txtLocation.Value, "code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Try
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Reset()
        BlankGrid()
        EnableDisable(True)
        RadPageView1.SelectedPage = RadPageViewPage1
    End Sub

    Sub EnableDisable(ByVal isVal As Boolean)
        txtFromDate.Enabled = isVal
        txtToDate.Enabled = isVal
        rgbDate.Enabled = isVal
        txtLocation.Enabled = isVal
        MyLabel4.Enabled = isVal
        MyLabel1.Enabled = isVal
        cboUOMType.Enabled = isVal
        RadGroupBox1.Enabled = isVal
        txtCustomer.Enabled = isVal
    End Sub

    Sub BlankGrid()
        gvData.DataSource = Nothing
        gvData.Columns.Clear()
        gvData.Rows.Clear()
        gvData.AutoExpandGroups = True
        gvData.ShowGroupPanel = False
        gvData.ShowRowHeaderColumn = False
        gvData.AllowAddNewRow = False
        gvData.AllowDeleteRow = False
        gvData.EnableFiltering = True
        gvData.ShowFilteringRow = True
        gvData.Refresh()
    End Sub

    Function ReturnBaseQry() As String
        Dim BaseQry As String = "select max(Document_Date) as Document_Date,Document_Code,max(Customer_Name) as Customer_Name,sum(Gross_Amount) as Gross_Amount	,sum(Total_Amt) as Total_Amt,sum(Transporter_Commission_TotalAmt) as Transporter_Commission_TotalAmt,max(Is_Taxable) as Is_Taxable,max(Is_FreshItem) as Is_FreshItem,max(Is_Ambient) as Is_Ambient,Item_Code,max(Item_Desc) as Item_Desc,
max(UOM) as UOM ,sum(Qty) as Qty,sum(Item_Net_Amt) as Item_Net_Amt,max(HSN_Code) as HSN_Code,max(GST_Registered) as GST_Registered,max(GSTNO) as GSTNO,max(Address) as Address,max([Pin Code]) as [Pin Code],max(Remarks) as Remarks,max(Tax1) as Tax1,sum(TAX1_Amt) as TAX1_Amt,max(TAX2) as TAX2,sum(TAX2_Amt) as TAX2_Amt,max(TAX3) as TAX3,sum(TAX3_Amt) as TAX3_Amt,max(TAX4) as TAX4,sum(TAX4_Amt) as TAX4_Amt,max(TAX5) as TAX5,sum(TAX5_Amt) as TAX5_Amt,max(TAX6) as TAX6,sum(TAX6_Amt) as TAX6_Amt	,max(TAX7) as TAX7,sum(TAX7_Amt) as TAX7_Amt,max(TAX8) as TAX8	,sum(TAX8_Amt) as TAX8_Amt	,max(TAX9) as TAX9,sum(TAX9_Amt) as TAX9_Amt,max(TAX10) as TAX10,sum(TAX10_Amt) as TAX10_Amt,sum(CGST) as CGST,sum(SGST) as SGST,sum(IGST) as IGST,sum(KKF) as KKF ,sum(MANDI_TAX) as MANDI_TAX  from (
SELECT CONVERT(VARCHAR(10), H.Document_Date, 103) AS Document_Date,H.Document_Code,C.Customer_Name,H.Gross_Amount,H.Total_Amt,H.Transporter_Commission_TotalAmt,H.Is_Taxable,I.Is_FreshItem,I.Is_Ambient,I.Item_Code,I.Item_Desc,
CFDiv.UOM_Code as UOM,case when isnull(CFDiv.Conversion_Factor,0)=0 then 0 else ((d.Qty * isnull(CFMult.Conversion_Factor,0))/CFDiv.Conversion_Factor) end as Qty,
(D.Item_Net_Amt-D.Total_Tax_Amt) As Item_Net_Amt,I.HSN_Code,C.GST_Registered,C.GSTNO,CONCAT(C.Add1, C.Add2, C.Add3) AS Address,ISNULL(C.PIN_Code, C.PIN_NO) AS [Pin Code],H.Remarks,
D.Tax1,D.TAX1_Amt,D.TAX2,D.TAX2_Amt,D.TAX3,D.TAX3_Amt,D.TAX4,D.TAX4_Amt,D.TAX5,D.TAX5_Amt,D.TAX6,D.TAX6_Amt,D.TAX7,D.TAX7_Amt,D.TAX8,D.TAX8_Amt,D.TAX9,D.TAX9_Amt,D.TAX10,D.TAX10_Amt,
Case When TAX1.TYPE='CGST' Then D.TAX1_Amt
     When TAX2.TYPE='CGST' Then D.TAX2_Amt
	 When TAX3.TYPE='CGST' Then D.TAX3_Amt
	 When TAX4.TYPE='CGST' Then D.TAX4_Amt
	 When TAX5.TYPE='CGST' Then D.TAX5_Amt
	 When TAX6.TYPE='CGST' Then D.TAX6_Amt
	 When TAX7.TYPE='CGST' Then D.TAX7_Amt
	 When TAX8.TYPE='CGST' Then D.TAX8_Amt
	 When TAX9.TYPE='CGST' Then D.TAX9_Amt	
	 When TAX10.TYPE='CGST' Then D.TAX10_Amt
	 Else 0 End As [CGST],
Case When TAX1.TYPE='SGST' Then D.TAX1_Amt
     When TAX2.TYPE='SGST' Then D.TAX2_Amt
	 When TAX3.TYPE='SGST' Then D.TAX3_Amt
	 When TAX4.TYPE='SGST' Then D.TAX4_Amt
	 When TAX5.TYPE='SGST' Then D.TAX5_Amt
	 When TAX6.TYPE='SGST' Then D.TAX6_Amt
	 When TAX7.TYPE='SGST' Then D.TAX7_Amt
	 When TAX8.TYPE='SGST' Then D.TAX8_Amt
	 When TAX9.TYPE='SGST' Then D.TAX9_Amt	
	 When TAX10.TYPE='SGST' Then D.TAX10_Amt
	 Else 0 End As [SGST],
Case When TAX1.TYPE='IGST' Then D.TAX1_Amt
     When TAX2.TYPE='IGST' Then D.TAX2_Amt
	 When TAX3.TYPE='IGST' Then D.TAX3_Amt
	 When TAX4.TYPE='IGST' Then D.TAX4_Amt
	 When TAX5.TYPE='IGST' Then D.TAX5_Amt
	 When TAX6.TYPE='IGST' Then D.TAX6_Amt
	 When TAX7.TYPE='IGST' Then D.TAX7_Amt
	 When TAX8.TYPE='IGST' Then D.TAX8_Amt
	 When TAX9.TYPE='IGST' Then D.TAX9_Amt	
	 When TAX10.TYPE='IGST' Then D.TAX10_Amt
	 Else 0 End As [IGST],
Case When TAX1.TYPE='K' Then D.TAX1_Amt
     When TAX2.TYPE='K' Then D.TAX2_Amt
	 When TAX3.TYPE='K' Then D.TAX3_Amt
	 When TAX4.TYPE='K' Then D.TAX4_Amt
	 When TAX5.TYPE='K' Then D.TAX5_Amt
	 When TAX6.TYPE='K' Then D.TAX6_Amt
	 When TAX7.TYPE='K' Then D.TAX7_Amt
	 When TAX8.TYPE='K' Then D.TAX8_Amt
	 When TAX9.TYPE='K' Then D.TAX9_Amt	
	 When TAX10.TYPE='K' Then D.TAX10_Amt
	 Else 0 End As [KKF],
Case When TAX1.TYPE='M' Then D.TAX1_Amt
     When TAX2.TYPE='M' Then D.TAX2_Amt
	 When TAX3.TYPE='M' Then D.TAX3_Amt
	 When TAX4.TYPE='M' Then D.TAX4_Amt
	 When TAX5.TYPE='M' Then D.TAX5_Amt
	 When TAX6.TYPE='M' Then D.TAX6_Amt
	 When TAX7.TYPE='M' Then D.TAX7_Amt
	 When TAX8.TYPE='M' Then D.TAX8_Amt
	 When TAX9.TYPE='M' Then D.TAX9_Amt	
	 When TAX10.TYPE='M' Then D.TAX10_Amt
	 Else 0 End As [MANDI_TAX]
FROM TSPL_SD_SALE_INVOICE_DETAIL D
LEFT JOIN TSPL_SD_SALE_INVOICE_HEAD H ON H.Document_Code = D.Document_Code 
LEFT JOIN TSPL_CUSTOMER_MASTER C ON C.Cust_Code = H.Customer_Code
LEFT JOIN TSPL_ITEM_MASTER I ON I.Item_Code = D.Item_Code 
Left Join TSPL_TAX_MASTER As TAX1 On TAX1.Tax_Code = D.TAX1
Left Join TSPL_TAX_MASTER As TAX2 On TAX2.Tax_Code = D.TAX2
Left Join TSPL_TAX_MASTER As TAX3 On TAX3.Tax_Code = D.TAX3
Left Join TSPL_TAX_MASTER As TAX4 On TAX4.Tax_Code = D.TAX4
Left Join TSPL_TAX_MASTER As TAX5 On TAX5.Tax_Code = D.TAX5
Left Join TSPL_TAX_MASTER As TAX6 On TAX6.Tax_Code = D.TAX6
Left Join TSPL_TAX_MASTER As TAX7 On TAX7.Tax_Code = D.TAX7
Left Join TSPL_TAX_MASTER As TAX8 On TAX8.Tax_Code = D.TAX8
Left Join TSPL_TAX_MASTER As TAX9 On TAX9.Tax_Code = D.TAX9
Left Join TSPL_TAX_MASTER As TAX10 On Tax10.Tax_Code = D.TAX10 
left outer join TSPL_ITEM_UOM_DETAIL as CFMult on CFMult.Item_Code=d.Item_Code and CFMult.UOM_Code=D.Unit_code
left outer join TSPL_ITEM_UOM_DETAIL as CFDiv on CFDiv.Item_Code=d.Item_Code "

        If clsCommon.CompairString(clsCommon.myCstr(cboUOMType.SelectedValue), "Billing UOM") = CompairStringResult.Equal Then
            BaseQry &= " and CFDiv.Billing_UOM = 1 "
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboUOMType.SelectedValue), "Default UOM") = CompairStringResult.Equal Then
            BaseQry &= " and CFDiv.Default_UOM = 1 "
        Else
            BaseQry &= " and CFDiv.Report_UOM = 1 "
        End If
        BaseQry &= " Where CONVERT(date,H.Document_Date,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And CONVERT(date,H.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
        If clsCommon.myLen(txtLocation.Value) > 0 Then
            BaseQry &= " and (H.Bill_To_Location='" & clsCommon.myCstr(txtLocation.Value) & "' or H.Sub_Location_code='" & clsCommon.myCstr(txtLocation.Value) & "') "
        End If
        If txtCustomer.arrValueMember IsNot Nothing AndAlso txtCustomer.arrValueMember.Count > 0 Then
            BaseQry &= " and H.Customer_Code in (" & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & ")"
        End If
        BaseQry &= " and H.Status=1 ) xx group by Document_Code,item_Code "
        Return BaseQry
    End Function

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            Dim Qry As String = Nothing
            If rbtnSaleVoucher.Checked Then
                Qry = "; With BaseQry As (" & ReturnBaseQry() & ")"
                Qry &= ",
--Header Row (Customer only)
HeaderRows AS
(SELECT DISTINCT Document_Date,CASE WHEN Is_Taxable = 1 THEN 'TAX INVOICE' ELSE 'BILL OF SUPPLY' END AS [Sale Vch. Type Name],Document_Code,Customer_Name,Gross_Amount,'Dr' AS [Dr/Cr],NULL AS Item_Code,NULL AS Item_Desc,0 as Qty,NULL AS UOM,NULL AS Item_Net_Amt,NULL AS HSN_Code,GSTNO,Address,[Pin Code],Remarks,CASE  WHEN GST_Registered = 1 THEN 'Regular' ELSE 'Unregisterd'END AS [GST Type],1 AS SortOrder FROM BaseQry),
--Detail Rows (Item only)
DetailRows AS
(SELECT Document_Date,CASE WHEN Is_Taxable = 1 THEN 'TAX INVOICE' ELSE 'BILL OF SUPPLY' END AS [Sale Vch. Type Name],Document_Code,NULL AS Customer_Name,NULL AS Gross_Amount, NULL AS [Dr/Cr],Item_Code, Item_Desc,Qty,UOM,Item_Net_Amt,HSN_Code,NULL AS GSTNO,NULL AS Address,NULL AS [Pin Code],NULL AS Remarks,NULL AS [GST Type],2 AS SortOrder FROM BaseQry
),
----Tax Rows (Tax only)
TaxDetailRows AS
(SELECT B.Document_Date, CASE WHEN MAX(B.Is_Taxable) = 1 THEN 'TAX INVOICE' ELSE 'BILL OF SUPPLY' END AS [Sale Vch. Type Name],B.Document_Code,NULL AS Customer_Name, NULL AS Gross_Amount, NULL AS [Dr/Cr],NULL AS Item_Code,T.TaxName AS Item_Desc,0 as Qty,NULL AS UOM,SUM(T.TaxAmount) AS Item_Net_Amt,NULL AS HSN_Code, NULL AS GSTNO, NULL AS Address, NULL AS [Pin Code], NULL AS Remarks,NULL AS [GST Type],3 AS SortOrder FROM BaseQry B CROSS APPLY (VALUES ('CGST', B.CGST),('SGST', B.SGST),('IGST', B.IGST),('KKF', B.KKF),('MANDI TAX', B.MANDI_TAX)) T (TaxName, TaxAmount) WHERE T.TaxAmount <> 0 GROUP BY B.Document_Code,B.Document_Date,T.TaxName
),
--Additional Charges
AdditionalChargesRows AS
(SELECT B.Document_Date, CASE WHEN MAX(B.Is_Taxable) = 1 THEN 'TAX INVOICE' ELSE 'BILL OF SUPPLY' END  AS [Sale Vch. Type Name],B.Document_Code,
Concat('Milk Transportation Chgs',' ('+Max(Customer_Name)+')') AS Customer_Name, Max(Transporter_Commission_TotalAmt) AS Total_Amt, 'Dr' AS [Dr/Cr],NULL AS Item_Code,NULL AS Item_Desc,0 as Qty,NULL AS UOM,
Null AS Item_Net_Amt,NULL AS HSN_Code, NULL AS GSTNO, NULL AS Address, NULL AS [Pin Code], NULL AS Remarks,NULL AS [GST Type],4 AS SortOrder 
FROM BaseQry B 
WHERE Transporter_Commission_TotalAmt <> 0 GROUP BY B.Document_Code,B.Document_Date
)
--Final Output
SELECT Document_Date AS [Sale Vch. Date],[Sale Vch. Type Name],Document_Code AS [Vch. Number],Customer_Name AS [Ledger Name],Gross_Amount AS [Ledger Amt.],[Dr/Cr],Item_Code as Item, Item_Desc AS [Product/Ledger Name],Qty,UOM,Item_Net_Amt AS [Product Value],HSN_Code AS [HSN/SAC],GSTNO AS [Buyer/Supplier - GSTIN/UIN],Address,[Pin Code],Remarks AS [Vch. Narration],[GST Type] FROM
(
    SELECT * FROM HeaderRows
    UNION ALL
    SELECT * FROM DetailRows
    UNION ALL
    SELECT * FROM TaxDetailRows
    Union All
	Select * from AdditionalChargesRows
) A
ORDER BY Document_Date,Document_Code,[Sale Vch. Type Name],SortOrder "

            ElseIf rbtnHSNWise.Checked Then
                Dim rpt As New rptHSNWiseSaleReport()
                Dim BaseQry As String
                Dim strOutTaxColumnSum As String = ""
                If cboUOMType.SelectedIndex = 0 Then
                    BaseQry = rpt.ReturnQry(True, txtFromDate.Value, txtToDate.Value, txtLocation.Value, "Report UOM", strOutTaxColumnSum)
                ElseIf cboUOMType.SelectedIndex = 1 Then
                    BaseQry = rpt.ReturnQry(True, txtFromDate.Value, txtToDate.Value, txtLocation.Value, "Default UOM", strOutTaxColumnSum)
                Else
                    BaseQry = rpt.ReturnQry(True, txtFromDate.Value, txtToDate.Value, txtLocation.Value, "Billing UOM", strOutTaxColumnSum)
                End If
                rpt = Nothing
                Qry = " ;WITH FinalData AS ("
                Qry &= "Select [Supply Type],Max(HSN_Code) As [HSN],Item_Code As [Item Code],Max(Short_Description) As [Description],Max(UOM) As [UQC],Sum(Total_Qty) As [Total Quantity],Sum(Item_Net_Amt) As [Total Value],Max(Tax_Rate) As Rate,Sum(MandiAmt) As [Mandi Amount],Sum(kkfAmt) As [KKF Amount],Sum(MandiAmt)+Sum(kkfAmt)+Sum(Taxable_Amt) As [Taxable Value]" + strOutTaxColumnSum + ",'' As [Cess Amt] from  (" & BaseQry & ")finalQry Group By [Supply Type],Item_Code "
                Qry &= "),
DataWithRowNo AS
(SELECT  '' AS [Supply Type],[HSN], [Description],[UQC], 
Convert(Decimal(18,2),(Case When [Total Quantity]<0 Then ([Total Quantity])*-1 Else ([Total Quantity]) End)) As [Total Quantity],
Convert(Decimal(18,2),(Case When [Total Value]<0 Then [Total Value]*-1 Else [Total Value] End)) As [Total Value],
Convert(Decimal(18,2),(FinalData.Rate)) As Rate,
Convert(Decimal(18,2),(Case When [Mandi Amount]<0 Then [Mandi Amount]*-1 Else [Mandi Amount] End)) As [Mandi Amount],
Convert(Decimal(18,2),(Case When [KKF Amount]<0 Then [KKF Amount]*-1 Else [KKF Amount] End)) As [KKF Amount],
Convert(Decimal(18,2),(Case When [Taxable Value]<0 Then [Taxable Value]*-1 Else [Taxable Value] End)) As [Taxable Value],
Convert(Decimal(18,2),(Case When [IGST]<0 Then [IGST]*-1 Else [IGST] End)) As [IGST],
Convert(Decimal(18,2),(Case When [CGST]<0 Then [CGST]*-1 Else [CGST] End)) As [CGST], 
Convert(Decimal(18,2),(Case When [S/UGST] <0 Then [S/UGST]*-1 Else [S/UGST] End)) As [S/UGST],[Cess Amt],[Supply Type] AS GroupType, ROW_NUMBER() OVER (PARTITION BY [Supply Type] ORDER BY IsNull(TSPL_ITEM_MASTER.IsTaxable,0) ASC,[Description]) AS SortNo,1 AS RowType FROM FinalData
Left Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=FinalData.[Item Code]
),
HeaderRows AS
(SELECT [Supply Type], NULL AS [HSN],NULL AS [Description],NULL AS [UQC], NULL AS [Total Quantity], NULL AS [Total Value],NULL AS Rate,Null As [Mandi Amount],Null As [KKF Amount], NULL AS [Taxable Value], NULL AS [IGST], NULL AS [CGST], NULL AS [S/UGST],NULL AS [Cess Amt],[Supply Type] AS GroupType,0 AS SortNo, 0 AS RowType FROM FinalData GROUP BY [Supply Type] )
SELECT 
    [Supply Type],[HSN],[Description],[UQC],[Total Quantity],[Total Value],Rate,[Mandi Amount],[KKF Amount], [Taxable Value],[IGST],[CGST],[S/UGST],[Cess Amt]
FROM
(
    SELECT * FROM HeaderRows
    UNION ALL
    SELECT * FROM DataWithRowNo
) A ORDER BY GroupType, RowType, SortNo;"
            Else
                Dim rpt As New rptSaleInvoiceStatusReport()
                Dim BaseQry As String = rpt.BaseQryLoadDataInvoiceCount(txtFromDate.Value, txtToDate.Value, txtLocation.Value)
                rpt = Nothing
                Qry = "Select Max(Invoice_Tax_Type) As [Sale Voucher Type],Min(First_Invoice) As [Sr.No. From],Max(Last_Invoice) As [Sr. No. To],Sum(Total_Invoice) As [Total Number],
(Sum(Total_CancelInvoice)+SUM(Total_DeleteInvoice)) As [Cancelled] from (" & BaseQry & ")final Group By Prefix " ' ,Invoice_Tax_Type"
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                BlankGrid()
                gvData.DataSource = dt
                gvData.AutoExpandGroups = True
                gvData.ShowGroupPanel = False
                gvData.ShowRowHeaderColumn = False
                gvData.AllowAddNewRow = False
                gvData.AllowDeleteRow = False
                gvData.EnableFiltering = True
                gvData.ShowFilteringRow = True
                gvData.BestFitColumns()
                gvData.ReadOnly = True
                RadPageView1.SelectedPage = RadPageViewPage2
                EnableDisable(False)
                ReStoreGridLayout()
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmenuExport_Click(sender As Object, e As EventArgs) Handles rmenuExport.Click
        Try
            If gvData.Rows.Count > 0 Then
                ExportToExcel(EnumExportTo.Excel)
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ExportToExcel(ByVal exporter As EnumExportTo)
        Try
            Dim arrHeader As List(Of String) = New List(Of String)()
            Dim strtemp As String = "Date Range : " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") & " To " & clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")
            arrHeader.Add(strtemp)
            strtemp = Nothing
            If rbtnSaleVoucher.Checked Then
                strtemp = "Sales Voucher Report"
            ElseIf rbtnHSNWise.Checked Then
                strtemp = "Sales HSN Wise Report"
            Else
                strtemp = "Sales Invoice Count"
            End If
            arrHeader.Add(strtemp)
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            If exporter = EnumExportTo.Excel Then
                clsCommon.MyExportToExcelGrid(strtemp, gvData, arrHeader, Me.Text)
            Else
                clsCommon.MyExportToPDF(strtemp, gvData, arrHeader, strtemp, PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub rmenuPDF_Click(sender As Object, e As EventArgs) Handles rmenuPDF.Click
        Try
            If gvData.Rows.Count > 0 Then
                ExportToExcel(EnumExportTo.PDF)
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnSaleVoucher_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnSaleVoucher.CheckedChanged
        HideAndShow()
    End Sub

    Private Sub rbtnHSNWise_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnHSNWise.CheckedChanged
        HideAndShow()
    End Sub

    Private Sub rbtnInvoiceCount_CheckedChanged(sender As Object, e As EventArgs) Handles rbtnInvoiceCount.CheckedChanged
        HideAndShow()
    End Sub

    Sub HideAndShow()
        If rbtnHSNWise.Checked Then
            MyLabel4.Visible = True
            cboUOMType.Visible = True
            MyLabel2.Visible = False
            txtCustomer.Visible = False
        ElseIf rbtnSaleVoucher.Checked Then
            MyLabel4.Visible = True
            cboUOMType.Visible = True
            MyLabel2.Visible = True
            txtCustomer.Visible = True
        Else
            MyLabel4.Visible = False
            cboUOMType.Visible = False
            MyLabel2.Visible = False
            txtCustomer.Visible = False
        End If
    End Sub

    Private Sub txtCustomer__My_Click(sender As Object, e As EventArgs) Handles txtCustomer._My_Click
        Dim strQry As String = " select Cust_Code as [code],Customer_Name as [Name] from TSPL_CUSTOMER_MASTER"
        txtCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("TransTypeMu23", strQry, "Code", "Name", txtCustomer.arrValueMember, txtCustomer.arrDispalyMember)
    End Sub

    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        If clsCommon.myLen(MyBase.Form_ID) > 0 Then
            gvData.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = GetReportID()
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gvData.SaveLayout(obj.GridLayout)
            obj.GridColumns = gvData.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(GetReportID(), objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(GetReportID(), "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gvData.ColumnCount Then
                    Dim ii As Integer = 0
                    For ii = 0 To gvData.Columns.Count - 1 Step ii + 1
                        gvData.Columns(ii).IsVisible = False
                        gvData.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gvData.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
                obj = Nothing
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub
    Function GetReportID() As String

        If rbtnSaleVoucher.Checked Then
            Return MyBase.Form_ID + "S"
        ElseIf rbtnHSNWise.Checked Then
            Return MyBase.Form_ID + "H"
        Else
            Return MyBase.Form_ID + "C"
        End If
    End Function
End Class