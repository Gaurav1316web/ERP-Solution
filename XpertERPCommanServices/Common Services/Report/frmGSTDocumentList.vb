Imports System.Text
Imports common
Public Class frmGSTDocumentList
    Dim sbQry As StringBuilder = Nothing
    Dim sbFinal As StringBuilder = Nothing
    Dim isDoubleClick As Boolean = False

    Private Sub frmGSTDocumentList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            txtFromDate.Value = clsCommon.GETSERVERDATE()
            txtToDate.Value = txtFromDate.Value
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Reset()
        Try
            ddlType.SelectedIndex = 0
            ddlInvoiceType.SelectedIndex = 0
            ddlTransaction.SelectedIndex = 0
            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gvDetails.DataSource = Nothing
            gvDetails.Rows.Clear()
            gvDetails.Columns.Clear()
            RadPageView1.SelectedPage = RadPageViewPage1
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Try
            Reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnGO_Click(sender As Object, e As EventArgs) Handles btnGO.Click
        Try
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(ReturnFinalQry())
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv.MasterTemplate.SummaryRowsBottom.Clear()
                gv.DataSource = Nothing
                gv.Rows.Clear()
                gv.Columns.Clear()
                gv.DataSource = dt
                gv.EnableFiltering = True
                gv.EnableSorting = True
                gv.ShowFilteringRow = True
                gv.BestFitColumns()
                RadPageView1.SelectedPage = RadPageViewPage2
                GvFormatGrid()
                SummaryRow()
                view()
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub GvFormatGrid()
        Try
            gv.AllowAddNewRow = False
            gv.TableElement.TableHeaderHeight = 40
            gv.MasterTemplate.ShowRowHeaderColumn = False
            gv.EnableFiltering = True
            For ii As Integer = 0 To gv.Columns.Count - 1
                gv.Columns(ii).ReadOnly = True
            Next
            gv.Columns("Trans_Type").HeaderText = "Transaction"
            gv.Columns("Type").HeaderText = "Type"
            gv.Columns("DocFrom").HeaderText = "From"
            gv.Columns("DocTo").HeaderText = "To"

            gv.Columns("DocCount").HeaderText = "Total Document Count"
            gv.Columns("B2BInPortal").HeaderText = "B2B"
            gv.Columns("B2BOutPortal").HeaderText = "B2B"
            gv.Columns("B2C").HeaderText = "B2C"
            gv.Columns("Cancel").HeaderText = "Cancel"
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Sub view()
        Try
            If gv.Rows.Count > 0 Then
                Dim view As New ColumnGroupsViewDefinition()
                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Trans_Type").Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Type").Name)


                view.ColumnGroups.Add(New GridViewColumnGroup("Doc No"))
                view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("DocFrom").Name)
                view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("DocTo").Name)


                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("DocCount").Name)


                view.ColumnGroups.Add(New GridViewColumnGroup("GST Portal"))
                view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv.Columns("B2BInPortal").Name)
                view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv.Columns("B2BOutPortal").Name)

                view.ColumnGroups.Add(New GridViewColumnGroup(""))
                view.ColumnGroups(4).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(gv.Columns("B2C").Name)
                view.ColumnGroups(4).Rows(0).ColumnNames.Add(gv.Columns("Cancel").Name)

                gv.ViewDefinition = view
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Private Sub gv_CellDoubleClick(sender As Object, e As GridViewCellEventArgs) Handles gv.CellDoubleClick
        Try
            isDoubleClick = True
            If gv IsNot Nothing AndAlso gv.Rows.Count > 0 AndAlso gv.CurrentRow.Index > -1 Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(ReturnFinalQry())
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    gvDetails.MasterTemplate.SummaryRowsBottom.Clear()
                    gvDetails.DataSource = Nothing
                    gvDetails.Rows.Clear()
                    gvDetails.Columns.Clear()
                    gvDetails.DataSource = dt
                    gvDetails.EnableFiltering = True
                    gvDetails.EnableSorting = True
                    gvDetails.ShowFilteringRow = True
                    gvDetails.BestFitColumns()
                    gvDetails.BestFitColumns()
                    SummaryRow()
                    RadPageView1.SelectedPage = RadPageViewPage3
                Else
                    clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
                End If
            End If
            isDoubleClick = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub SummaryRow()
        Try
            Dim summaryRowItem As New GridViewSummaryRowItem()
            If isDoubleClick Then
                For ii As Integer = 23 To gvDetails.Columns.Count - 1
                    If clsCommon.CompairString(gvDetails.Columns(ii).Name, "HSN Code") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(gvDetails.Columns(ii).Name, "EwayBillNo") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(gvDetails.Columns(ii).Name, "EwayBillDate") <> CompairStringResult.Equal Then
                        Dim Item As New GridViewSummaryItem(gvDetails.Columns(ii).Name, "{0:n2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(Item)
                    End If
                Next
                gvDetails.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                gvDetails.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            Else
                For ii As Integer = 4 To gv.Columns.Count - 1
                    Dim Item As New GridViewSummaryItem(gv.Columns(ii).Name, "{0:n2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(Item)
                Next
                gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Function ReturnFinalQry() As String
        sbFinal = New StringBuilder()
        sbFinal.Append(ReturnSaleInvoiceBaseQry() & Environment.NewLine & "Union All" & Environment.NewLine & ReturnSaleReturnBaseQry() & " Union All" & Environment.NewLine & ReturnCustomerInvoiceBaseQry() & " Union All " & Environment.NewLine & ReturnVendorInvoiceBaseQry())

        Dim sbfinalQry As New StringBuilder()
        If Not isDoubleClick Then
            sbfinalQry.Append("Select Trans_Type,Type,Min(Document_Code) As DocFrom,MAX(Document_Code) As DocTo,COUNT(Distinct Document_Code) As DocCount,Count(Distinct B2BInPortal) As B2BInPortal,Count(Distinct B2BOutPortal)B2BOutPortal,Count(Distinct B2C)B2C,
Count(Distinct Cancel)Cancel  from (Select Document_Code,Document_Date,Trans_Type,Case When Is_taxable=1 Then 'Taxable' Else 'Non-Taxable' End As Type,
Case When ISNULL(IRN_No,'')<>'' And EInvoice_Type In ('B2B','BB') Then Document_Code End As B2BInPortal,
Case When ISNULL(IRN_No,'')='' And EInvoice_Type In ('B2B','BB') Then Document_Code  End As B2BOutPortal,
Case When EInvoice_Type In ('B2C','BC') Then Document_Code End As B2C,
Case When IsNull(Cancel_By,'')<>'' Then Document_Code  End As Cancel
from 
(")
            sbfinalQry.Append(sbFinal)
            sbfinalQry.Append(")xyz )fff Group By Trans_Type,Type Order By Trans_Type,Type")
        Else
            sbfinalQry.Append("Select Trans_Type As [Transaction],Convert(varchar(10),Supply_Date,103) As [Supply Date],Shift_Type As [Shift Type],Bill_To_Location As [Location],Sub_Location_code As [Sub Location],CompGSTNO As [GST No],CompState As [State Code],Cust_Code As [Customer Code],Customer_Name As [Customer Name],State As [Party State],
GSTNO As [GRecipient Gst No],EInvoice_Type As [E-Invoice Type],Case When Isnull(IRN_No,'')<>'' Then 'Yes' When ISNULL(Cancel_By,'')<>'' Then 'Cancel' Else 'No' End As [GST Portal Status],Ack_No As [Ack No],Ack_Date As [Ack Date],IRN_No As [IRN No],Document_Code As [Invoice No],Document_Date As [Invoice Date],Case When Is_Taxable=1 Then 'Taxable' Else 'Non-Taxable' End As [Invoice Type],Route_No As [Route No],Item_Code As [Item Code],Item_Desc As [Item Name],Unit_code As [UOM],Qty,Amount As [Item Amount],HSN_Code As [HSN Code],EWayBillNo,EWayBillDate,
Case When TAX1='KKF' Then TAX1_Rate
     When TAX2='KKF' Then TAX2_Rate
	 When TAX3='KKF' Then TAX3_Rate
	 When TAX4='KKF' Then TAX4_Rate
	 When TAX5='KKF' Then TAX5_Rate
	 When TAX6='KKF' Then TAX6_Rate
	 When TAX7='KKF' Then TAX7_Rate
	 When TAX8='KKF' Then TAX8_Rate
	 When TAX9='KKF' Then TAX9_Rate
	 When TAX10='KKF' Then TAX10_Rate Else 0 End As [KKF %],

Case When TAX1='KKF' Then TAX1_Amt
     When TAX2='KKF' Then TAX2_Amt
	 When TAX3='KKF' Then TAX3_Amt
	 When TAX4='KKF' Then TAX4_Amt
	 When TAX5='KKF' Then TAX5_Amt
	 When TAX6='KKF' Then TAX6_Amt
	 When TAX7='KKF' Then TAX7_Amt
	 When TAX8='KKF' Then TAX8_Amt
	 When TAX9='KKF' Then TAX9_Amt
	 When TAX10='KKF' Then TAX10_Amt Else 0 End As [KKF Amt],

Case When TAX1='MNDTAX' Then TAX1_Rate
     When TAX2='MNDTAX' Then TAX2_Rate
	 When TAX3='MNDTAX' Then TAX3_Rate
	 When TAX4='MNDTAX' Then TAX4_Rate
	 When TAX5='MNDTAX' Then TAX5_Rate
	 When TAX6='MNDTAX' Then TAX6_Rate
	 When TAX7='MNDTAX' Then TAX7_Rate
	 When TAX8='MNDTAX' Then TAX8_Rate
	 When TAX9='MNDTAX' Then TAX9_Rate
	 When TAX10='MNDTAX' Then TAX10_Rate Else 0 End As [Mandi Tax %],

Case When TAX1='MNDTAX' Then TAX1_Amt
     When TAX2='MNDTAX' Then TAX2_Amt
	 When TAX3='MNDTAX' Then TAX3_Amt
	 When TAX4='MNDTAX' Then TAX4_Amt
	 When TAX5='MNDTAX' Then TAX5_Amt
	 When TAX6='MNDTAX' Then TAX6_Amt
	 When TAX7='MNDTAX' Then TAX7_Amt
	 When TAX8='MNDTAX' Then TAX8_Amt
	 When TAX9='MNDTAX' Then TAX9_Amt
	 When TAX10='MNDTAX' Then TAX10_Amt Else 0 End As [Mandi Tax Amt],

Case When TAX1='CGST' Then TAX1_Rate
     When TAX2='CGST' Then TAX2_Rate
	 When TAX3='CGST' Then TAX3_Rate
	 When TAX4='CGST' Then TAX4_Rate
	 When TAX5='CGST' Then TAX5_Rate
	 When TAX6='CGST' Then TAX6_Rate
	 When TAX7='CGST' Then TAX7_Rate
	 When TAX8='CGST' Then TAX8_Rate
	 When TAX9='CGST' Then TAX9_Rate
	 When TAX10='CGST' Then TAX10_Rate Else 0 End As [CGST %],

Case When TAX1='CGST' Then TAX1_Amt
     When TAX2='CGST' Then TAX2_Amt
	 When TAX3='CGST' Then TAX3_Amt
	 When TAX4='CGST' Then TAX4_Amt
	 When TAX5='CGST' Then TAX5_Amt
	 When TAX6='CGST' Then TAX6_Amt
	 When TAX7='CGST' Then TAX7_Amt
	 When TAX8='CGST' Then TAX8_Amt
	 When TAX9='CGST' Then TAX9_Amt
	 When TAX10='CGST' Then TAX10_Amt Else 0 End As [CGST Amt],

Case When TAX1='SGST' Then TAX1_Rate
     When TAX2='SGST' Then TAX2_Rate
	 When TAX3='SGST' Then TAX3_Rate
	 When TAX4='SGST' Then TAX4_Rate
	 When TAX5='SGST' Then TAX5_Rate
	 When TAX6='SGST' Then TAX6_Rate
	 When TAX7='SGST' Then TAX7_Rate
	 When TAX8='SGST' Then TAX8_Rate
	 When TAX9='SGST' Then TAX9_Rate
	 When TAX10='SGST' Then TAX10_Rate Else 0 End As [SGST %],

Case When TAX1='SGST' Then TAX1_Amt
     When TAX2='SGST' Then TAX2_Amt
	 When TAX3='SGST' Then TAX3_Amt
	 When TAX4='SGST' Then TAX4_Amt
	 When TAX5='SGST' Then TAX5_Amt
	 When TAX6='SGST' Then TAX6_Amt
	 When TAX7='SGST' Then TAX7_Amt
	 When TAX8='SGST' Then TAX8_Amt
	 When TAX9='SGST' Then TAX9_Amt
	 When TAX10='SGST' Then TAX10_Amt Else 0 End As [SGST Amt],


Case When TAX1='IGST' Then TAX1_Rate
     When TAX2='IGST' Then TAX2_Rate
	 When TAX3='IGST' Then TAX3_Rate
	 When TAX4='IGST' Then TAX4_Rate
	 When TAX5='IGST' Then TAX5_Rate
	 When TAX6='IGST' Then TAX6_Rate
	 When TAX7='IGST' Then TAX7_Rate
	 When TAX8='IGST' Then TAX8_Rate
	 When TAX9='IGST' Then TAX9_Rate
	 When TAX10='IGST' Then TAX10_Rate Else 0 End As [IGST %],

Case When TAX1='IGST' Then TAX1_Amt
     When TAX2='IGST' Then TAX2_Amt
	 When TAX3='IGST' Then TAX3_Amt
	 When TAX4='IGST' Then TAX4_Amt
	 When TAX5='IGST' Then TAX5_Amt
	 When TAX6='IGST' Then TAX6_Amt
	 When TAX7='IGST' Then TAX7_Amt
	 When TAX8='IGST' Then TAX8_Amt
	 When TAX9='IGST' Then TAX9_Amt
	 When TAX10='IGST' Then TAX10_Amt Else 0 End As [IGST Amt],

Case When TAX1='TCS' Then TAX1_Rate
     When TAX2='TCS' Then TAX2_Rate
	 When TAX3='TCS' Then TAX3_Rate
	 When TAX4='TCS' Then TAX4_Rate
	 When TAX5='TCS' Then TAX5_Rate
	 When TAX6='TCS' Then TAX6_Rate
	 When TAX7='TCS' Then TAX7_Rate
	 When TAX8='TCS' Then TAX8_Rate
	 When TAX9='TCS' Then TAX9_Rate
	 When TAX10='TCS' Then TAX10_Rate Else 0 End As [TCS %],

Case When TAX1='TCS' Then TAX1_Amt
     When TAX2='TCS' Then TAX2_Amt
	 When TAX3='TCS' Then TAX3_Amt
	 When TAX4='TCS' Then TAX4_Amt
	 When TAX5='TCS' Then TAX5_Amt
	 When TAX6='TCS' Then TAX6_Amt
	 When TAX7='TCS' Then TAX7_Amt
	 When TAX8='TCS' Then TAX8_Amt
	 When TAX9='TCS' Then TAX9_Amt
	 When TAX10='TCS' Then TAX10_Amt Else 0 End As [TCS Amt],Total_Tax As [Total Tax Amount],Total_Amount As [Total Amount]

from  (")
            sbfinalQry.Append(sbFinal)
            sbfinalQry.Append(")xyz")
        End If

        Return clsCommon.myCstr(sbfinalQry)
    End Function

    Function ReturnSaleInvoiceBaseQry() As String
        sbQry = Nothing
        sbQry = New StringBuilder()
        Dim TSPL_Sale_Invoice_Head_Table As String = Nothing
        Dim TSPL_Sale_Invoice_Detail_Table As String = Nothing
        Dim TSPL_Shipment_Head_Table As String = Nothing
        For i As Integer = 0 To 1
            If i <> 0 Then
                TSPL_Sale_Invoice_Head_Table = "TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data As TSPL_SD_SALE_INVOICE_HEAD"
                TSPL_Sale_Invoice_Detail_Table = "TSPL_SD_SALE_INVOICE_DETAIL_Cancel_Data As TSPL_SD_SALE_INVOICE_DETAIL"
                TSPL_Shipment_Head_Table = "TSPL_SD_Shipment_Head_Cancel_Data As TSPL_SD_Shipment_Head"
                sbQry.Append(Environment.NewLine & " Union All " & Environment.NewLine)
            Else
                TSPL_Sale_Invoice_Head_Table = "TSPL_SD_SALE_INVOICE_HEAD"
                TSPL_Sale_Invoice_Detail_Table = "TSPL_SD_SALE_INVOICE_DETAIL"
                TSPL_Shipment_Head_Table = "TSPL_SD_Shipment_Head"
            End If
            sbQry.Append("select " & clsCommon.myCstr(IIf(i = 0, "'' As Cancel_By", "TSPL_SD_SALE_INVOICE_HEAD.Cancel_By")) & ", TSPL_SD_SALE_INVOICE_HEAD.Trans_Type,TSPL_SD_Shipment_Head.Supply_Date,TSPL_SD_Shipment_Head.Shift_Type,TSPL_SD_SALE_INVOICE_HEAD.Bill_To_Location,
TSPL_SD_SALE_INVOICE_HEAD.Sub_Location_code ,'' AS CompGSTNO,'' As CompState,
TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.State,TSPL_CUSTOMER_MASTER.GSTNO,
'' As GSTPortalStatus,
TSPL_SD_SALE_INVOICE_HEAD.Ack_Date,TSPL_SD_SALE_INVOICE_HEAD.Ack_No,TSPL_SD_SALE_INVOICE_HEAD.IRN_No,TSPL_SD_SALE_INVOICE_HEAD.Document_Code,
Convert(Varchar(10),TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) As Document_Date,TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type,TSPL_SD_SALE_INVOICE_HEAD.Invoice_Type,'' As Zone_Code,
TSPL_SD_SALE_INVOICE_HEAD.Route_No,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_SD_SALE_INVOICE_DETAIL.Unit_code,
TSPL_SD_SALE_INVOICE_DETAIL.Qty,TSPL_SD_SALE_INVOICE_DETAIL.Amount,TSPL_ITEM_MASTER.HSN_Code,TSPL_SD_SALE_INVOICE_HEAD.EWayBillNo,TSPL_SD_SALE_INVOICE_HEAD.EWayBillDate,

TSPL_SD_SALE_INVOICE_DETAIL.TAX1,
TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX1_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX2,
TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX2_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX3,
TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX3_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX4,
TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX4_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX5,
TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX5_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX6,
TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX6_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX7,
TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX7_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX8,
TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX8_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX9,
TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX9_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.TAX10,
TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Rate,
TSPL_SD_SALE_INVOICE_DETAIL.TAX10_Amt,

TSPL_SD_SALE_INVOICE_DETAIL.Total_Tax_Amt As Total_Tax,
TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt As Total_Amount,TSPL_SD_SALE_INVOICE_HEAD.Is_taxable 


from " & TSPL_Sale_Invoice_Detail_Table & "   
Left Outer Join " & TSPL_Sale_Invoice_Head_Table & " On TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE 
Left Outer Join " & TSPL_Shipment_Head_Table & " On TSPL_SD_Shipment_Head.document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No 
Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code 
Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code Where TSPL_SD_SALE_INVOICE_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And TSPL_SD_SALE_INVOICE_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' ")
            If ddlInvoiceType.SelectedItem.Text = "B2B" Then
                sbQry.Append(" And TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type='BB' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "B2C" Then
                sbQry.Append(" And TSPL_SD_SALE_INVOICE_HEAD.EInvoice_Type='BC' ")
            Else
                If i <> 0 Then
                    sbQry.Append(" And IsNull(TSPL_SD_SALE_INVOICE_HEAD.Cancel_By,'')<>'' ")
                End If
            End If

            If ddlType.SelectedItem.Text = "Taxable" Then
                sbQry.Append(" And IsNull(TSPL_SD_SALE_INVOICE_HEAD.Is_taxable,0)='1' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "Non-Taxable" Then
                sbQry.Append(" And IsNull(TSPL_SD_SALE_INVOICE_HEAD.Is_taxable,0)='0' ")
            End If

        Next
        Return clsCommon.myCstr(sbQry)
    End Function

    Function ReturnSaleReturnBaseQry() As String
        sbQry = Nothing
        sbQry = New StringBuilder()
        Dim TSPL_SD_SALE_RETURN_HEAD_Table As String = Nothing
        Dim TSPL_SD_SALE_RETURN_Detail_Table As String = Nothing
        Dim TSPL_SD_SALE_INVOICE_HEAD_Table As String = Nothing
        Dim TSPL_Shipment_Head_Table As String = Nothing
        For i As Integer = 0 To 1
            If i <> 0 Then
                TSPL_SD_SALE_RETURN_HEAD_Table = "TSPL_SD_SALE_RETURN_HEAD_Cancel_Data As TSPL_SD_SALE_RETURN_HEAD"
                TSPL_SD_SALE_RETURN_Detail_Table = "TSPL_SD_SALE_RETURN_DETAIL_Cancel_Data As TSPL_SD_SALE_RETURN_DETAIL"
                TSPL_Shipment_Head_Table = "TSPL_SD_Shipment_Head_Cancel_Data As TSPL_SD_Shipment_Head"
                TSPL_SD_SALE_INVOICE_HEAD_Table = "TSPL_SD_SALE_INVOICE_HEAD_Cancel_Data As TSPL_SD_SALE_INVOICE_HEAD"
                sbQry.Append(Environment.NewLine & " Union All " & Environment.NewLine)
            Else
                TSPL_SD_SALE_RETURN_HEAD_Table = "TSPL_SD_SALE_RETURN_HEAD"
                TSPL_SD_SALE_RETURN_Detail_Table = "TSPL_SD_SALE_RETURN_DETAIL"
                TSPL_SD_SALE_INVOICE_HEAD_Table = "TSPL_SD_SALE_INVOICE_HEAD"
                TSPL_Shipment_Head_Table = "TSPL_SD_Shipment_Head"
            End If
            sbQry.Append("select " & clsCommon.myCstr(IIf(i = 0, "'' As Cancel_By", "TSPL_SD_SALE_RETURN_HEAD.Cancel_By")) & ",TSPL_SD_SALE_RETURN_HEAD.Trans_Type,TSPL_SD_Shipment_Head.Supply_Date,TSPL_SD_Shipment_Head.Shift_Type,TSPL_SD_SALE_RETURN_HEAD.Bill_To_Location,
TSPL_SD_SALE_RETURN_HEAD.Sub_Location_code ,'' AS CompGSTNO,'' As CompState,
TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.State,TSPL_CUSTOMER_MASTER.GSTNO,
'' As GSTPortalStatus,
TSPL_SD_SALE_RETURN_HEAD.Ack_Date,TSPL_SD_SALE_RETURN_HEAD.Ack_No,TSPL_SD_SALE_RETURN_HEAD.IRN_No,TSPL_SD_SALE_RETURN_HEAD.Document_Code,
Convert(Varchar(10),TSPL_SD_SALE_RETURN_HEAD.Document_Date,103) As Document_Date,TSPL_SD_SALE_RETURN_HEAD.EInvoice_Type,TSPL_SD_SALE_RETURN_HEAD.Invoice_Type,'' As Zone_Code,
TSPL_SD_SALE_RETURN_HEAD.Route_No,TSPL_SD_SALE_RETURN_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_SD_SALE_RETURN_DETAIL.Unit_code,
TSPL_SD_SALE_RETURN_DETAIL.Qty,TSPL_SD_SALE_RETURN_DETAIL.Amount,TSPL_ITEM_MASTER.HSN_Code,'' As WayBillNo,'' As EWayBillDate,

TSPL_SD_SALE_RETURN_DETAIL.TAX1,
TSPL_SD_SALE_RETURN_DETAIL.TAX1_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX1_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX2,
TSPL_SD_SALE_RETURN_DETAIL.TAX2_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX2_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX3,
TSPL_SD_SALE_RETURN_DETAIL.TAX3_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX3_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX4,
TSPL_SD_SALE_RETURN_DETAIL.TAX4_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX4_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX5,
TSPL_SD_SALE_RETURN_DETAIL.TAX5_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX5_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX6,
TSPL_SD_SALE_RETURN_DETAIL.TAX6_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX6_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX7,
TSPL_SD_SALE_RETURN_DETAIL.TAX7_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX7_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX8,
TSPL_SD_SALE_RETURN_DETAIL.TAX8_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX8_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX9,
TSPL_SD_SALE_RETURN_DETAIL.TAX9_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX9_Amt,

TSPL_SD_SALE_RETURN_DETAIL.TAX10,
TSPL_SD_SALE_RETURN_DETAIL.TAX10_Rate,
TSPL_SD_SALE_RETURN_DETAIL.TAX10_Amt,

TSPL_SD_SALE_RETURN_DETAIL.Total_Tax_Amt As Total_Tax,
TSPL_SD_SALE_RETURN_DETAIL.Item_Net_Amt As Total_Amount,TSPL_SD_SALE_RETURN_HEAD.Is_taxable


from " & TSPL_SD_SALE_RETURN_Detail_Table & "
Left Outer Join " & TSPL_SD_SALE_RETURN_HEAD_Table & " On TSPL_SD_SALE_RETURN_HEAD.document_code=TSPL_SD_SALE_RETURN_DETAIL.DOCUMENT_CODE
Left Outer Join " & TSPL_SD_SALE_INVOICE_HEAD_Table & " On TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_RETURN_HEAD.Against_Invoice_No
Left Outer Join " & TSPL_Shipment_Head_Table & " On TSPL_SD_Shipment_Head.document_Code=TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_RETURN_DETAIL.Item_Code
Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code where TSPL_SD_SALE_RETURN_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And TSPL_SD_SALE_RETURN_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' ")
            If ddlInvoiceType.SelectedItem.Text = "B2B" Then
                sbQry.Append(" And TSPL_SD_SALE_RETURN_HEAD.EInvoice_Type='BB' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "B2C" Then
                sbQry.Append(" And TSPL_SD_SALE_RETURN_HEAD.EInvoice_Type='BC' ")
            Else
                If i <> 0 Then
                    sbQry.Append(" And IsNull(TSPL_SD_SALE_RETURN_HEAD.Cancel_By,'')<>'' ")
                End If
            End If

            If ddlType.SelectedItem.Text = "Taxable" Then
                sbQry.Append(" And IsNull(TSPL_SD_SALE_RETURN_HEAD.Is_taxable,0)='1' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "Non-Taxable" Then
                sbQry.Append(" And IsNull(TSPL_SD_SALE_RETURN_HEAD.Is_taxable,0)='0' ")
            End If

            If ddlType.SelectedItem.Text = "Taxable" Then
                sbQry.Append(" And IsNull(TSPL_SD_SALE_RETURN_HEAD.Is_taxable,0)='1' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "Non-Taxable" Then
                sbQry.Append(" And IsNull(TSPL_SD_SALE_RETURN_HEAD.Is_taxable,0)='0' ")
            End If
        Next
        Return clsCommon.myCstr(sbQry)
    End Function

    Function ReturnCustomerInvoiceBaseQry() As String
        sbQry = Nothing
        sbQry = New StringBuilder()
        Dim TSPL_Customer_Invoice_Head_Table As String = Nothing
        Dim TSPL_Customer_Invoice_Detail_Table As String = Nothing
        For i As Integer = 0 To 1
            If i <> 0 Then
                sbQry.Append(Environment.NewLine & " Union All " & Environment.NewLine)
                TSPL_Customer_Invoice_Head_Table = "TSPL_Customer_Invoice_Head_Cancel_Data As TSPL_Customer_Invoice_Head"
                TSPL_Customer_Invoice_Detail_Table = "TSPL_Customer_Invoice_Detail_Cancel_Data As TSPL_Customer_Invoice_Detail"
            Else
                TSPL_Customer_Invoice_Head_Table = "TSPL_Customer_Invoice_Head"
                TSPL_Customer_Invoice_Detail_Table = "TSPL_Customer_Invoice_Detail"
            End If

            sbQry.Append("select " & clsCommon.myCstr(IIf(i = 0, "'' As Cancel_By", "TSPL_Customer_INVOICE_HEAD.Cancel_By")) & ",TSPL_Customer_INVOICE_HEAD.Trans_Type,'' As Supply_Date,'' As Shift_Type,TSPL_Customer_INVOICE_HEAD.Loc_Code,
'' As Sub_Location_code ,'' AS CompGSTNO,'' As CompState,
TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_CUSTOMER_MASTER.State,TSPL_CUSTOMER_MASTER.GSTNO,
'' As GSTPortalStatus,
TSPL_Customer_INVOICE_HEAD.Ack_Date,TSPL_Customer_INVOICE_HEAD.Ack_No,TSPL_Customer_INVOICE_HEAD.IRN_No,TSPL_Customer_INVOICE_HEAD.Document_No,
Convert(Varchar(10),TSPL_Customer_INVOICE_HEAD.Document_Date,103) As Document_Date,TSPL_Customer_INVOICE_HEAD.EInvoice_Type,'' As Invoice_Type,'' As Zone_Code,
TSPL_Customer_INVOICE_HEAD.Route_No,'' As Item_Code,'' As Item_Desc,'' As Unit_code,
0 As Qty,TSPL_Customer_INVOICE_Detail.Amount,'' As HSN_Code,'' As EWayBillNo,'' As EWayBillDate,

TSPL_Customer_INVOICE_Detail.TAX1,
TSPL_Customer_INVOICE_Detail.TAX1_Rate,
TSPL_Customer_INVOICE_Detail.TAX1_Amt,

TSPL_Customer_INVOICE_Detail.TAX2,
TSPL_Customer_INVOICE_Detail.TAX2_Rate,
TSPL_Customer_INVOICE_Detail.TAX2_Amt,

TSPL_Customer_INVOICE_Detail.TAX3,
TSPL_Customer_INVOICE_Detail.TAX3_Rate,
TSPL_Customer_INVOICE_Detail.TAX3_Amt,

TSPL_Customer_INVOICE_Detail.TAX4,
TSPL_Customer_INVOICE_Detail.TAX4_Rate,
TSPL_Customer_INVOICE_Detail.TAX4_Amt,

TSPL_Customer_INVOICE_Detail.TAX5,
TSPL_Customer_INVOICE_Detail.TAX5_Rate,
TSPL_Customer_INVOICE_Detail.TAX5_Amt,

TSPL_Customer_INVOICE_Detail.TAX6,
TSPL_Customer_INVOICE_Detail.TAX6_Rate,
TSPL_Customer_INVOICE_Detail.TAX6_Amt,

TSPL_Customer_INVOICE_Detail.TAX7,
TSPL_Customer_INVOICE_Detail.TAX7_Rate,
TSPL_Customer_INVOICE_Detail.TAX7_Amt,

TSPL_Customer_INVOICE_Detail.TAX8,
TSPL_Customer_INVOICE_Detail.TAX8_Rate,
TSPL_Customer_INVOICE_Detail.TAX8_Amt,

TSPL_Customer_INVOICE_Detail.TAX9,
TSPL_Customer_INVOICE_Detail.TAX9_Rate,
TSPL_Customer_INVOICE_Detail.TAX9_Amt,

TSPL_Customer_INVOICE_Detail.TAX10,
TSPL_Customer_INVOICE_Detail.TAX10_Rate,
TSPL_Customer_INVOICE_Detail.TAX10_Amt,

TSPL_Customer_INVOICE_Detail.Total_Tax,
TSPL_Customer_INVOICE_Detail.Total_Amount,0 As Is_taxable


from " & TSPL_Customer_Invoice_Detail_Table & " 
Left Outer Join " & TSPL_Customer_Invoice_Head_Table & " On TSPL_Customer_INVOICE_HEAD.Document_No=TSPL_Customer_INVOICE_Detail.Document_No
--Left Outer Join TSPL_SD_Shipment_Head On TSPL_SD_Shipment_Head.document_Code=TSPL_Customer_INVOICE_HEAD.Against_Shipment_No
--Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_Customer_INVOICE_Detail.Item_Code
Left Outer Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_Customer_INVOICE_HEAD.Customer_Code Where IsNull(TSPL_Customer_INVOICE_HEAD.Is_Einvoice,'')<>'' And  TSPL_Customer_INVOICE_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And TSPL_Customer_INVOICE_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' ")
            If ddlInvoiceType.SelectedItem.Text = "B2B" Then
                sbQry.Append(" And TSPL_Customer_INVOICE_HEAD.EInvoice_Type='BB' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "B2C" Then
                sbQry.Append(" And TSPL_Customer_INVOICE_HEAD.EInvoice_Type='BC' ")
            Else
                If i <> 0 Then
                    sbQry.Append(" And IsNull(TSPL_Customer_INVOICE_HEAD.Cancel_By,'')<>'' ")
                End If
            End If
        Next
        Return clsCommon.myCstr(sbQry)
    End Function


    Function ReturnVendorInvoiceBaseQry() As String
        sbQry = Nothing
        sbQry = New StringBuilder()
        Dim TSPL_Vendor_Invoice_Head_Table As String = Nothing
        Dim TSPL_Vendor_Invoice_Detail_Table As String = Nothing
        For i As Integer = 0 To 1
            If i <> 0 Then
                sbQry.Append(Environment.NewLine & " Union All " & Environment.NewLine)
                TSPL_Vendor_Invoice_Head_Table = "TSPL_VENDOR_INVOICE_HEAD_Cancel_Data As TSPL_VENDOR_INVOICE_HEAD"
                TSPL_Vendor_Invoice_Detail_Table = "TSPL_VENDOR_INVOICE_Detail_Cancel_Data As TSPL_VENDOR_INVOICE_Detail"
            Else
                TSPL_Vendor_Invoice_Head_Table = "TSPL_VENDOR_INVOICE_HEAD"
                TSPL_Vendor_Invoice_Detail_Table = "TSPL_VENDOR_INVOICE_Detail"
            End If
            sbQry.Append("select " & clsCommon.myCstr(IIf(i = 0, "'' As Cancel_By", "TSPL_VENDOR_INVOICE_HEAD.Cancel_By")) & ",'Vendor Service Charge' As Trans_Type,'' As Supply_Date,'' As Shift_Type,TSPL_VENDOR_INVOICE_HEAD.Loc_Code,
'' As Sub_Location_code ,'' AS CompGSTNO,'' As CompState,
TSPL_VENDOR_MASTER.Vendor_Code As Cust_Cude,TSPL_VENDOR_MASTER.Vendor_Name As Customer_Name,TSPL_VENDOR_MASTER.State,TSPL_VENDOR_MASTER.GSTFinalNo As GSTNO,
'' As GSTPortalStatus,
TSPL_VENDOR_INVOICE_HEAD.Ack_Date,TSPL_VENDOR_INVOICE_HEAD.Ack_No,TSPL_VENDOR_INVOICE_HEAD.IRN_No,TSPL_VENDOR_INVOICE_HEAD.Document_No,
Convert(Varchar(10),TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date,103) As Document_Date,TSPL_VENDOR_INVOICE_HEAD.EInvoice_Type,TSPL_VENDOR_INVOICE_HEAD.Invoice_Type,'' As Zone_Code,
'' As Route_No,'' As Item_Code,'' As Item_Desc,'' As Unit_code,
0 As Qty,TSPL_VENDOR_INVOICE_Detail.Amount,'' As HSN_Code,'' As EWayBillNo,'' As EWayBillDate,

TSPL_VENDOR_INVOICE_Detail.TAX1,
TSPL_VENDOR_INVOICE_Detail.TAX1_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX1_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX2,
TSPL_VENDOR_INVOICE_Detail.TAX2_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX2_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX3,
TSPL_VENDOR_INVOICE_Detail.TAX3_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX3_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX4,
TSPL_VENDOR_INVOICE_Detail.TAX4_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX4_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX5,
TSPL_VENDOR_INVOICE_Detail.TAX5_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX5_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX6,
TSPL_VENDOR_INVOICE_Detail.TAX6_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX6_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX7,
TSPL_VENDOR_INVOICE_Detail.TAX7_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX7_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX8,
TSPL_VENDOR_INVOICE_Detail.TAX8_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX8_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX9,
TSPL_VENDOR_INVOICE_Detail.TAX9_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX9_Amt,

TSPL_VENDOR_INVOICE_Detail.TAX10,
TSPL_VENDOR_INVOICE_Detail.TAX10_Rate,
TSPL_VENDOR_INVOICE_Detail.TAX10_Amt,

TSPL_VENDOR_INVOICE_Detail.Total_Tax,
TSPL_VENDOR_INVOICE_Detail.Total_Amount,0 As Is_taxable

from " & TSPL_Vendor_Invoice_Detail_Table & " 
Left Outer Join " & TSPL_Vendor_Invoice_Head_Table & " On TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_VENDOR_INVOICE_Detail.Document_No
--Left Outer Join TSPL_SD_Shipment_Head On TSPL_SD_Shipment_Head.document_Code=TSPL_VENDOR_INVOICE_HEAD.Against_Shipment_No
--Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_VENDOR_INVOICE_Detail.Item_Code
Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code Where IsNull(TSPL_VENDOR_INVOICE_HEAD.IsEinvoice,'')<>'' And TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' ")
            If ddlInvoiceType.SelectedItem.Text = "B2B" Then
                sbQry.Append(" And TSPL_VENDOR_INVOICE_HEAD.EInvoice_Type='BB' ")
            ElseIf ddlInvoiceType.SelectedItem.Text = "B2C" Then
                sbQry.Append(" And TSPL_VENDOR_INVOICE_HEAD.EInvoice_Type='BC' ")
            Else
                If i <> 0 Then
                    sbQry.Append(" And IsNull(TSPL_VENDOR_INVOICE_HEAD.Cancel_By,'')<>'' ")
                End If
            End If
        Next
        Return clsCommon.myCstr(sbQry)
    End Function

    Private Sub QExpExcel_Click(sender As Object, e As EventArgs) Handles QExpExcel.Click
        Try
            If gv IsNot Nothing AndAlso gv.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Date Range: " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") & " To " & clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy"))
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_CODE='" & clsUserMgtCode.frmGSTDocumentList & "'"))
                transportSql.QuickExportToExcel(gv, "", Me.Text, , arrHeader)
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found to export !", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


End Class