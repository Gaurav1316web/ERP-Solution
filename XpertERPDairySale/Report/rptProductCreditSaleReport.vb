Imports System.Data.SqlClient
Imports System.IO
Imports common
Imports System.Globalization
Imports System.Text.RegularExpressions


Public Class rptProductCreditSaleReport
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
        Reset()
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

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            LoadData()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
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
            Dim itemNames1 As String = Nothing
            Dim itemNames2 As String = Nothing
            Dim itemNames3 As String = Nothing
            Dim itemNames4 As String = Nothing

            Dim ItemQry As String = " Select Distinct TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,Sku_Seq
from  TSPL_SD_SHIPMENT_HEAD
left outer join TSPL_SD_SHIPMENT_BOOKING_DETAIL ON TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE=TSPL_SD_SHIPMENT_HEAD.Document_Code
left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code
where Convert( Date, TSPL_SD_SHIPMENT_HEAD.Document_Date,103) >=convert(date,'" + clsCommon.GetPrintDate(txtfDate.Value, "dd/MMM/yyyy") + "',103) 
and convert(date,TSPL_SD_SHIPMENT_HEAD.Document_Date,103) <= Convert(Date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "',103) order by Sku_Seq"
            Dim dtitemName As DataTable = clsDBFuncationality.GetDataTable(ItemQry)
            If dtitemName.Rows.Count > 0 Then
                For i As Integer = 0 To dtitemName.Rows.Count - 1
                    If i = 0 Then
                        itemNames1 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames2 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + "]"
                        itemNames4 += " Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " Qty],Sum([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + "]) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " Amt]"
                    Else
                        itemNames1 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames2 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + "]"
                        itemNames4 += ", Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " Qty],Sum([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + "]) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " Amt]"
                    End If
                Next
            End If

            If dtitemName.Rows.Count > 0 Then


                qry = " Select Customer_Code,max(Credit_Customer)Credit_Customer,case when max(Credit_Customer)='Y' then max(Customer_Name) +'(In Distributor)' else  max(Customer_Name) +'(To Agent)' end as Customer_Name," & itemNames4 & "
 from (SELECT Customer_Code,Credit_Customer,Customer_Name, " & itemNames1 & ", " & itemNames2 & "
	  FROM
(
    SELECT TSPL_SD_SHIPMENT_HEAD.Customer_Code,max(TSPL_CUSTOMER_MASTER.Credit_Customer)Credit_Customer,Max(TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name,
        TSPL_DEMAND_BOOKING_DETAIL.Item_Code,max(Item_desc)Item_Desc,SUM(TSPL_DEMAND_BOOKING_DETAIL.Qty) AS Qty, "
                If chkShowAmt.Checked Then
                    qry += " Sum(ItemNetAmount) as Item_Net_Amt "
                Else
                    qry += " CASE WHEN max(TSPL_CUSTOMER_MASTER.Credit_Customer) = 'Y' THEN 0
                            ELSE Sum(ItemNetAmount) END AS Item_Net_Amt "
                End If
                qry += " 
    FROM TSPL_SD_SHIPMENT_HEAD 
    LEFT JOIN TSPL_SD_SHIPMENT_BOOKING_DETAIL ON TSPL_SD_SHIPMENT_BOOKING_DETAIL.DOCUMENT_CODE = TSPL_SD_SHIPMENT_HEAD.Document_Code
    LEFT JOIN TSPL_DEMAND_BOOKING_DETAIL ON TSPL_DEMAND_BOOKING_DETAIL.TR_Code = TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code
    LEFT JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SHIPMENT_HEAD.Customer_Code
	left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SHIPMENT_BOOKING_DETAIL.Item_Code

    WHERE CONVERT(DATE, TSPL_SD_SHIPMENT_HEAD.Document_Date, 103) >= CONVERT(DATE, '" + strtxtfDate + "', 103) and
                    CONVERT(DATE, TSPL_SD_SHIPMENT_HEAD.Document_Date, 103) <= CONVERT(DATE, '" + strToDate + "', 103) 
    GROUP BY TSPL_SD_SHIPMENT_HEAD.Customer_Code,TSPL_DEMAND_BOOKING_DETAIL.Item_Code
) X
PIVOT
( SUM(Qty) FOR Item_Code IN (" & itemNames1 & ") ) P
	Pivot 
	(Sum(Item_Net_Amt) for item_Desc In (" & itemNames2 & ") ) Q
) YY Group by Customer_Code "
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            dt = clsDBFuncationality.GetDataTable(qry)
            gvdata.DataSource = Nothing
            gvdata.Rows.Clear()
            gvdata.Columns.Clear()
            gvdata.GroupDescriptors.Clear()

            If dt.Rows.Count > 0 Then
                Try
                    gvdata.DataSource = dt
                    gvdata.GroupDescriptors.Clear()
                    gvdata.EnableFiltering = True
                    gvdata.MasterTemplate.SummaryRowsBottom.Clear()
                    SetGridFormation()
                    EnableDisableCntrl(False)
                    RadPageView1.SelectedPage = RadPageViewPage3


                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
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

        gvdata.Columns("Customer_Code").IsVisible = False
        gvdata.Columns("Customer_Code").VisibleInColumnChooser = True
        gvdata.Columns("Credit_Customer").IsVisible = False
        gvdata.Columns("Credit_Customer").VisibleInColumnChooser = True
        gvdata.Columns("Customer_Code").HeaderText = "Particulars"

    End Sub
    Private Sub rmenuExport_Click(sender As Object, e As EventArgs) Handles rmenuExport.Click
        If gvdata.Rows.Count > 0 Then
            ExporttoExcel(EnumExportTo.Excel)
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
                clsCommon.MyExportToExcelGrid("Sale Invoice Status Report", gvdata, arrHeader, Me.Text)
            Else
                clsCommon.MyExportToPDF("Sale Invoice Status Report", gvdata, arrHeader, "Sale Invoice Status Report", PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub

    Private Sub rmenuPDF_Click(sender As Object, e As EventArgs) Handles rmenuPDF.Click
        If gvdata.Rows.Count > 0 Then
            ExportToExcel(EnumExportTo.PDF)
        Else
            RadMessageBox.Show("No Data Found to Display", Me.Text)
        End If
    End Sub
End Class