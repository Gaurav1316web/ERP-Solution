Imports System.IO
Imports common
Imports System.Text
Imports common.UserControls

Public Class rptDailySTAFGSStatement
    Inherits FrmMainTranScreen

#Region "Variables"
    Dim atchqry As String = ""
#End Region
    Private Sub rptDailySTAFGSStatement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fromDate.Value = clsCommon.GETSERVERDATE()
        ToDate.Value = clsCommon.GETSERVERDATE()
    End Sub

    Private Sub txtParty__My_Click(sender As Object, e As EventArgs) Handles txtParty._My_Click
        Dim qry As String = " Select Cust_Code as code,Customer_Name as Name from TSPL_CUSTOMER_MASTER "
        txtParty.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "Code", "Name", txtParty.arrValueMember, txtParty.arrDispalyMember)
    End Sub

    Private Sub txtSubLocation__My_Click(sender As Object, e As EventArgs) Handles txtSubLocation._My_Click
        Try
            Dim qry As String = " select location_code as Code,Location_Desc as Description from tspl_location_master where Is_Sub_Location='Y'  "
            txtSubLocation.arrValueMember = clsCommon.ShowMultipleSelectForm("sublocn", qry, "Code", "Description", txtSubLocation.arrValueMember, txtSubLocation.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Reset()
        Try
            funreset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Sub funreset()
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()
        RadPageView1.SelectedPage = RadPageViewPage1
        RadGroupBox2.Enabled = True
        RadGroupBox3.Enabled = True
        RadPageView1.SelectedPage = RadPageViewPage1
        txtParty.arrValueMember = Nothing
        txtSubLocation.arrValueMember = Nothing
        EnableDisableCntrl(True)
    End Sub

    Sub EnableDisableCntrl(ByVal val As Boolean)
        RadGroupBox3.Enabled = val
        RadGroupBox2.Enabled = val
        txtParty.Enabled = val
        txtSubLocation.Enabled = val

    End Sub

    Sub FormatGrid()
        Gv1.AutoExpandGroups = False
        Gv1.ShowGroupPanel = False
        Gv1.ShowRowHeaderColumn = False
        Gv1.AllowAddNewRow = False
        Gv1.AllowDeleteRow = False
        Gv1.EnableFiltering = True
        Gv1.ShowFilteringRow = True
        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).BestFit()
        Next
        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim intCount As Integer = 0
        'If rbtnSale.IsChecked Then
        Gv1.Columns("Customer_Code").HeaderText = "Customer Code"
        Gv1.Columns("Customer_Code").IsVisible = False
        Gv1.Columns("Customer_Code").VisibleInColumnChooser = True


        'End If

        Gv1.ShowGroupPanel = True
        Gv1.MasterTemplate.AutoExpandGroups = True
        Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            atchqry = GetAttachQry()

            If clsCommon.myLen(atchqry) > 0 Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(atchqry)

                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Gv1.DataSource = Nothing
                    Gv1.GroupDescriptors.Clear()
                    Gv1.SummaryRowsBottom.Clear()
                    Gv1.DataSource = dt
                    'gv1.Columns("TransType").IsVisible = False
                    'gv1.Columns("PROD_ENTRY_CODE").IsVisible = False
                    RadPageView1.SelectedPage = RadPageViewPage2
                    'FormatGrid()
                    'View()
                    FormatGrid()

                    Gv1.BestFitColumns()

                    EnableDisableCntrl(False)
                    'ReStoreGridLayout()
                Else
                    clsCommon.MyMessageBoxShow(Me, "No data found to display.", "Item Stock Report")
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "No data found to display.", "Item Stock Report")
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        'atchqry = GetAttachQry()

        'If rbtnBoth.IsChecked Then

        'End If
    End Sub

    Private Function GetAttachQry() As String
        Dim Qry As String = ""
        Dim itemNames1 As String = Nothing
        Dim itemNames2 As String = Nothing
        Dim itemNames3 As String = Nothing
        Dim itemNames4 As String = Nothing
        Dim itemSumBagName As String = Nothing
        Dim itemPivBagName As String = Nothing
        If rbtnSale.IsChecked Then

            Dim ItemQry As String = "  Select Distinct TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,Sku_Seq,TSPL_ITEM_UOM_DETAIL.UOM_Code
from TSPL_SD_SALE_INVOICE_DETAIL
left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE
left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left outer join TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM=1
where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >=convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103) order by Sku_Seq"
            Dim dtitemName As DataTable = clsDBFuncationality.GetDataTable(ItemQry)
            If dtitemName.Rows.Count > 0 Then
                For i As Integer = 0 To dtitemName.Rows.Count - 1
                    If i = 0 Then
                        itemNames1 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames4 += " Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " (" + clsCommon.myCstr(dtitemName.Rows(i)("UOM_Code")) + ")]"
                    Else
                        itemNames1 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames4 += ", Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " (" + clsCommon.myCstr(dtitemName.Rows(i)("UOM_Code")) + ")]"
                    End If
                Next
            End If

            If dtitemName.Rows.Count > 0 Then

                Qry = " Select Customer_Code,max(Customer_Name)[Party Name], " & itemNames4 & " from (select max(Document_Code)Document_Code,Customer_Code,max(Customer_Name)Customer_Name,Item_Code,sum(Qty)Qty,sum(ReportQty)ReportQty,sum(ReportQty1)ReportQty1 
                    from (Select TSPL_SD_SALE_INVOICE_HEAD.Document_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_UOM_DETAIL.UOM_Code,TSPL_SD_SALE_INVOICE_DETAIL.Qty,
                    ISNULL(CAST(TSPL_SD_SALE_INVOICE_DETAIL.Qty / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS INT), 0) AS ReportQty,
                    TSPL_SD_SALE_INVOICE_DETAIL.Qty/TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS ReportQty1,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name
                    from TSPL_SD_SALE_INVOICE_DETAIL
                    left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE
                    left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
                    left outer join TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                    LEFT OUTER JOIN TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code =TSPL_SD_SALE_INVOICE_DETAIL.Item_Code AND Report_UOM=1 
                    where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >=Convert(Date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
                    and TSPL_SD_SALE_INVOICE_HEAD.Document_Date<=Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103) "
                If txtParty.arrValueMember IsNot Nothing AndAlso txtParty.arrValueMember.Count > 0 Then
                    Qry += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code In (" + clsCommon.GetMulcallString(txtParty.arrValueMember) + ") "
                    ' qry += " and TSPL_ITEM_PRICE_PLAN_detail.Item_Code in ('" + clsCommon.GetMulcallString(txtItem.arrValueMember) + "')  "
                End If

                Qry += " ) xx GROUP BY Customer_Code,Item_Code)YY
                    PIVOT (SUM(reportQty) FOR item_code IN (" & itemNames1 & ") )AS Tab2  group by Customer_Code "
            End If
        ElseIf rbtnTransfer.IsChecked Then

            Dim ItemQry As String = "  Select Distinct TSPL_TRANSFER_ORDER_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,Sku_Seq,TSPL_ITEM_UOM_DETAIL.UOM_Code
from TSPL_TRANSFER_ORDER_DETAIL
left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code
left outer join TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM=1
where Convert( Date, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >=convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <= Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103) order by Sku_Seq"
            Dim dtitemName As DataTable = clsDBFuncationality.GetDataTable(ItemQry)
            If dtitemName.Rows.Count > 0 Then
                For i As Integer = 0 To dtitemName.Rows.Count - 1
                    If i = 0 Then
                        itemNames1 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames4 += " Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " (" + clsCommon.myCstr(dtitemName.Rows(i)("UOM_Code")) + ")]"
                    Else
                        itemNames1 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames4 += ", Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " (" + clsCommon.myCstr(dtitemName.Rows(i)("UOM_Code")) + ")]"
                    End If
                Next
            End If

            If dtitemName.Rows.Count > 0 Then

                Qry = "  Select To_Location as Customer_Code,max(Customer_Name)[Party Name], " & itemNames4 & " from (select max(Document_No)Document_Code,To_Location,
                    max(Location_Desc)Customer_Name,Item_Code,sum(Qty)Qty,sum(ReportQty)ReportQty,sum(ReportQty1)ReportQty1 
                    from (Select TSPL_TRANSFER_ORDER_HEAD.Document_No,TSPL_TRANSFER_ORDER_HEAD.Document_Date,TSPL_TRANSFER_ORDER_DETAIL.Item_Code,
                    TSPL_ITEM_UOM_DETAIL.UOM_Code,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,
                    ISNULL(CAST(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS INT), 0) AS ReportQty,
                    TSPL_TRANSFER_ORDER_DETAIL.Out_Qty/TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS ReportQty1,TSPL_TRANSFER_ORDER_HEAD.To_Location,
                    TSPL_LOCATION_MASTER.Location_Desc
                    from TSPL_TRANSFER_ORDER_DETAIL
                    left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No
                    left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code
                    left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                    LEFT OUTER JOIN TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code =TSPL_TRANSFER_ORDER_DETAIL.Item_Code AND Report_UOM=1
                    where Convert( Date, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >=Convert(Date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
                    and Convert( Date, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <=Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103)"
                If txtSubLocation.arrValueMember IsNot Nothing AndAlso txtSubLocation.arrValueMember.Count > 0 Then
                    Qry += " and TSPL_TRANSFER_ORDER_HEAD.To_Location In (" + clsCommon.GetMulcallString(txtSubLocation.arrValueMember) + ") "
                    ' qry += " and TSPL_ITEM_PRICE_PLAN_detail.Item_Code in ('" + clsCommon.GetMulcallString(txtItem.arrValueMember) + "')  "
                End If

                Qry += " ) xx GROUP BY To_Location,Item_Code)YY
                    PIVOT (SUM(reportQty) FOR item_code IN (" & itemNames1 & ") )AS Tab2  group by To_Location "
            End If
        Else

            Dim ItemQry As String = "   select distinct Item_Code,Item_Desc,Sku_Seq,UOM_Code from ( Select  TSPL_TRANSFER_ORDER_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,Sku_Seq,TSPL_ITEM_UOM_DETAIL.UOM_Code
from TSPL_TRANSFER_ORDER_DETAIL
left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code
left outer join TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM=1
where Convert( Date, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >=convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
and convert(date,TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) <= Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103) 

Union all

 Select  TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,Sku_Seq,TSPL_ITEM_UOM_DETAIL.UOM_Code
from TSPL_SD_SALE_INVOICE_DETAIL
left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE
left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
left outer join TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code AND TSPL_ITEM_UOM_DETAIL.Report_UOM=1
where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >=convert(date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103)
)XX ORDER BY Sku_Seq "

            Dim dtitemName As DataTable = clsDBFuncationality.GetDataTable(ItemQry)
            If dtitemName.Rows.Count > 0 Then
                For i As Integer = 0 To dtitemName.Rows.Count - 1
                    If i = 0 Then
                        itemNames1 += "[" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames4 += " Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " (" + clsCommon.myCstr(dtitemName.Rows(i)("UOM_Code")) + ")]"
                    Else
                        itemNames1 += ", [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "] "
                        itemNames4 += ", Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Code")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Item_Desc")) + " (" + clsCommon.myCstr(dtitemName.Rows(i)("UOM_Code")) + ")]"
                    End If
                Next
            End If

            If dtitemName.Rows.Count > 0 Then

                Qry = " Select Customer_Code,max(Customer_Name)[Party Name], " & itemNames4 & " from 
                    (select max(Document_Code)Document_Code,Customer_Code,max(Location_Desc)Customer_Name,Item_Code,sum(Qty)Qty,sum(ReportQty)ReportQty,sum(ReportQty1)ReportQty1 
                    from (Select TSPL_TRANSFER_ORDER_HEAD.Document_No as Document_Code,TSPL_TRANSFER_ORDER_HEAD.Document_Date,TSPL_TRANSFER_ORDER_DETAIL.Item_Code,
                    TSPL_ITEM_UOM_DETAIL.UOM_Code,TSPL_TRANSFER_ORDER_DETAIL.Out_Qty as Qty,
                    ISNULL(CAST(TSPL_TRANSFER_ORDER_DETAIL.Out_Qty / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS INT), 0) AS ReportQty,
                    TSPL_TRANSFER_ORDER_DETAIL.Out_Qty/TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS ReportQty1,TSPL_TRANSFER_ORDER_HEAD.To_Location as Customer_Code,
                    TSPL_LOCATION_MASTER.Location_Desc
                    from TSPL_TRANSFER_ORDER_DETAIL
                    left outer join TSPL_TRANSFER_ORDER_HEAD on TSPL_TRANSFER_ORDER_HEAD.Document_No=TSPL_TRANSFER_ORDER_DETAIL.Document_No
                    left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_TRANSFER_ORDER_DETAIL.Item_Code
                    left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_TRANSFER_ORDER_HEAD.To_Location
                    LEFT OUTER JOIN TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code =TSPL_TRANSFER_ORDER_DETAIL.Item_Code AND Report_UOM=1
					where Convert( Date, TSPL_TRANSFER_ORDER_HEAD.Document_Date,103) >=Convert(Date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
                    and TSPL_TRANSFER_ORDER_HEAD.Document_Date<=Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103)"
            If txtSubLocation.arrValueMember IsNot Nothing AndAlso txtSubLocation.arrValueMember.Count > 0 Then
                Qry += " and TSPL_TRANSFER_ORDER_HEAD.To_Location In (" + clsCommon.GetMulcallString(txtSubLocation.arrValueMember) + ") "
                ' qry += " and TSPL_ITEM_PRICE_PLAN_detail.Item_Code in ('" + clsCommon.GetMulcallString(txtItem.arrValueMember) + "')  "
            End If

            Qry += " UNION ALL

					Select TSPL_SD_SALE_INVOICE_HEAD.Document_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,TSPL_SD_SALE_INVOICE_DETAIL.Item_Code,TSPL_ITEM_UOM_DETAIL.UOM_Code,TSPL_SD_SALE_INVOICE_DETAIL.Qty,
                    ISNULL(CAST(TSPL_SD_SALE_INVOICE_DETAIL.Qty / TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS INT), 0) AS ReportQty,
                    TSPL_SD_SALE_INVOICE_DETAIL.Qty/TSPL_ITEM_UOM_DETAIL.Conversion_Factor AS ReportQty1,TSPL_SD_SALE_INVOICE_HEAD.Customer_Code,TSPL_CUSTOMER_MASTER.Customer_Name
                    from TSPL_SD_SALE_INVOICE_DETAIL
                    left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code=TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE
                    left outer join TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code
                    left outer join TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
                    LEFT OUTER JOIN TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code =TSPL_SD_SALE_INVOICE_DETAIL.Item_Code AND Report_UOM=1
					where Convert( Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >=Convert(Date,'" + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + "',103) 
                    and TSPL_SD_SALE_INVOICE_HEAD.Document_Date<=Convert(Date,'" + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + "',103) "
            If txtParty.arrValueMember IsNot Nothing AndAlso txtParty.arrValueMember.Count > 0 Then
                Qry += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code In (" + clsCommon.GetMulcallString(txtParty.arrValueMember) + ") "
                ' qry += " and TSPL_ITEM_PRICE_PLAN_detail.Item_Code in ('" + clsCommon.GetMulcallString(txtItem.arrValueMember) + "')  "
            End If
                Qry += " ) xx GROUP BY Customer_Code,Item_Code)YY
                    PIVOT (SUM(reportQty) FOR item_code IN (" & itemNames1 & ") )AS Tab2  group by Customer_Code "
            End If
        End If
        Return Qry
    End Function

    Private Sub btnreset_Click(sender As Object, e As EventArgs) Handles btnreset.Click
        Try
            funreset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Export(EnumExportTo.Excel)
    End Sub

    Private Sub Export(ByVal exporter As EnumExportTo)
        Try
            If Gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptDailySTAFGSStatement & "'"))

                transportSql.QuickExportToExcel(Gv1, "", Me.Text, , arrHeader)
            Else
                clsCommon.MyMessageBoxShow(Me, "No data found to export", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnTransfer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnTransfer.ToggleStateChanged
        If rbtnTransfer.IsChecked Then
            txtParty.Visible = False
            lblItem.Visible = False
        Else
            txtParty.Visible = True
            lblItem.Visible = True
        End If
    End Sub

    Private Sub rbtnSale_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnSale.ToggleStateChanged
        If rbtnSale.IsChecked Then
            txtSubLocation.Visible = False
            lblCode.Visible = False
        Else
            txtSubLocation.Visible = True
            lblCode.Visible = True
        End If
    End Sub

    Private Sub rbtnBoth_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnBoth.ToggleStateChanged
        If rbtnBoth.IsChecked Then
            txtSubLocation.Visible = True
            lblCode.Visible = True
            txtParty.Visible = True
            lblItem.Visible = True
        End If
    End Sub
End Class