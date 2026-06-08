
Imports common
Imports System.IO
Imports Telerik.Pivot.Core
Imports Telerik.WinControls.Export
Imports Telerik.WinControls.UI
Imports Telerik.WinControls.UI.Export
Imports Telerik.Windows.Controls

Public Class rptMilkSupplySaleReport
    Inherits FrmMainTranScreen

#Region "Variables"
    Private spreadExporter As PivotGridSpreadExport
    Private pdfExporter As PivotGridPdfExport
    Private radPrintDocument1 As RadPrintDocument
    'Private SpreadExportRenderer As ISpreadExportRenderer
#End Region
    Private Sub rptMilkSupplySaleReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        funreset()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
    End Sub

    Private Sub txtRoute__My_Click(sender As Object, e As EventArgs) Handles txtRoute._My_Click
        Try
            Dim qry As String = "select distinct TSPL_SD_SALE_INVOICE_HEAD.Route_No as  [ROUTE NO],TSPL_ROUTE_MASTER.Route_Desc as [ROUTE NAME]  FROM TSPL_SD_SALE_INVOICE_DETAIL left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code 
           left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_SD_SALE_INVOICE_HEAD.Route_No
            where  TSPL_SD_SALE_INVOICE_HEAD.Status = 1 AND convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >=Convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) 
            and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "',103)  and TSPL_ITEM_MASTER.Is_FreshItem = 1 "
            If txtCustomer.arrValueMember IsNot Nothing Then
                qry += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") "
            End If
            If rbtnMorning.IsChecked Then
                qry += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'AM' "
            ElseIf rbtnEvening.IsChecked Then
                qry += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'PM' "
            End If
            txtRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("MilkSupplyRoute", qry, "ROUTE NO", "ROUTE NAME", txtRoute.arrValueMember, txtRoute.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtCustomer__My_Click(sender As Object, e As EventArgs) Handles txtCustomer._My_Click
        Try
            Dim qry As String = "select distinct TSPL_SD_SALE_INVOICE_HEAD.Customer_Code as Code ,tspl_customer_master.Customer_Name as Name FROM TSPL_SD_SALE_INVOICE_DETAIL
            left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code  left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE
            left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No inner join tspl_customer_master on tspl_customer_master.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code 
            where  TSPL_SD_SALE_INVOICE_HEAD.Status = 1 AND convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >=Convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "',103)  and TSPL_ITEM_MASTER.Is_FreshItem = 1  "

            If txtRoute.arrValueMember IsNot Nothing Then
                qry += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No  in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ") "
            End If

            If rbtnMorning.IsChecked Then
                qry += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'AM' "
            ElseIf rbtnEvening.IsChecked Then
                qry += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'PM' "
            End If
            txtCustomer.arrValueMember = clsCommon.ShowMultipleSelectForm("MilkSupplyCustomer", qry, "Code", "Name", txtCustomer.arrValueMember, txtCustomer.arrDispalyMember)
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
        chkMargin()
        gv.DataSource = Nothing
        PvtGrid.DataSource = Nothing
        txtRoute.arrValueMember = Nothing
        txtCustomer.arrValueMember = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
        rbtnDetail.IsChecked = True
        rbtnCustomer.IsChecked = True
        rbtnMorning.IsChecked = True
    End Sub

    Private Sub EnableDisableControls(ByVal val As Boolean)
        RadGroupBox1.Enabled = val
    End Sub

    Private Sub LoadData()
        Try
            Dim BaseQry As String = ""
            Dim dt As DataTable = Nothing
            If chkDisplayMargin.IsChecked Then
                Dim strItem As String = Nothing
                Dim strPivotItem As String = Nothing
                Dim strSumItem As String = Nothing
                BaseQry = "Select Item_Code,Max(Item_Desc)Item_Desc,Max(Short_Description)Short_Description,Sku_Seq from(" + ReturnQry() + ")Item Group By Item_Code,Sku_Seq Order By Sku_Seq "
                dt = clsDBFuncationality.GetDataTable(BaseQry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For Each rows In dt.Rows
                        If clsCommon.myLen(strItem) > 0 AndAlso clsCommon.CompairString(strItem, clsCommon.myCstr(rows("Short_Description"))) Then
                            strItem += ",[" + clsCommon.myCstr(rows("Short_Description")) + "]"
                            strPivotItem += ",[" + clsCommon.myCstr(rows("Short_Description")) + "]"
                            strSumItem += ",Sum(IsNull([" + clsCommon.myCstr(rows("Short_Description")) + "],0))[" + clsCommon.myCstr(rows("Short_Description")) + "]"
                        Else
                            strItem = "[" + clsCommon.myCstr(rows("Short_Description")) + "]"
                            strPivotItem = " [" + clsCommon.myCstr(rows("Short_Description")) + "]"
                            strSumItem = " Sum(IsNull([" + clsCommon.myCstr(rows("Short_Description")) + "],0))[" + clsCommon.myCstr(rows("Short_Description")) + "]"
                        End If
                    Next
                Else
                    Throw New Exception("No Data Found !")
                End If

                BaseQry = "SELECT "
                If rbtnSummary.IsChecked Then
                    BaseQry += "[Party Code],Max([Party Name])[Party Name],Max(Margin_Rate) As [Margin Rate],Max(Route_No) As [Area],Max(Zone_Code) AS [Zone]," + strSumItem + ",Sum([Basic Amount])[Basic Amount],Sum(Margin_Amt) As [Margin Amt],Sum(Total_Tax_Amt) As [Total Tax],Sum([Total Amount])[Net Amount],Sum(Crate)[Total Crate]
FROM (SELECT Convert(varchar,Supply_Date,103)Supply_Date,[Party Code],[Party Name],Route_No,Route_Desc,Zone_Code,Zone_Desc,Item_Code,Short_Description,Convert(Decimal(18,2),CinCFQty)CinCFQty,[Basic Amount],Convert(Decimal(18,3),Margin_Rate)Margin_Rate,Margin_Amt,Total_Tax_Amt,[Total Amount],Crate FROM (" + ReturnQry() + ") AS BaseQry) AS SourceTable PIVOT (SUM(CinCFQty) FOR Short_Description IN (" + strItem + ")) AS PivotTable"
                Else
                    BaseQry += " Supply_Date As [Date],[Party Code],Max([Party Name])[Party Name],Max(Margin_Rate) As [Margin Rate],Max(Route_No) As [Area],Max(Zone_Code) AS [Zone]," + strSumItem + ",Sum([Basic Amount])[Basic Amount],Sum(Margin_Amt) As [Margin Amt],Sum(Total_Tax_Amt) As [Total Tax],Sum([Total Amount])[Net Amount],Sum(Crate)[Total Crate]
FROM (SELECT Convert(varchar,Supply_Date,103)Supply_Date,[Party Code],[Party Name],Route_No,Route_Desc,Zone_Code,Zone_Desc,Item_Code,Short_Description,Convert(Decimal(18,2),CinCFQty)CinCFQty,[Basic Amount],Convert(Decimal(18,3),Margin_Rate)Margin_Rate,Margin_Amt,Total_Tax_Amt,[Total Amount],Crate FROM (" + ReturnQry() + ") AS BaseQry) AS SourceTable PIVOT (SUM(CinCFQty) FOR Short_Description IN (" + strItem + ")) AS PivotTable"
                End If
                If rbtnDetail.IsChecked Then
                    BaseQry += " Group By Supply_Date,[Party Code],Route_No,Zone_Code"
                    BaseQry += " Order By Supply_Date,[Party Code],Route_No,Zone_Code"
                End If
                If rbtnSummary.IsChecked Then
                    BaseQry += " Group By [Party Code],Route_No Order By [Party Code],Route_No"
                End If

                dt = clsDBFuncationality.GetDataTable(BaseQry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    gv.DataSource = Nothing
                    gv.Rows.Clear()
                    gv.Columns.Clear()
                    gv.GroupDescriptors.Clear()
                    gv.MasterView.Refresh()
                    gv.GroupDescriptors.Clear()
                    gv.MasterTemplate.SummaryRowsBottom.Clear()
                    gv.AllowAddNewRow = False
                    gv.EnableFiltering = True
                    gv.DataSource = dt
                    gv.ReadOnly = True
                    gv.ShowGroupPanel = False
                    ReStoreGridLayout()
                    gv.MasterTemplate.AutoExpandGroups = False

                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    Dim item As GridViewSummaryItem

                    item = New GridViewSummaryItem("Margin Rate", "{0:n2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(item)
                    For i As Integer = 5 To gv.Columns.Count - 1
                        item = New GridViewSummaryItem(clsCommon.myCstr(gv.Columns(i).Name), "{0:n2}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(item)
                    Next
                    gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                    gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                    RadPageView1.SelectedPage = RadPageViewPage3
                    gv.BestFitColumns()
                Else
                    Throw New Exception("No Data Found !")
                End If
            Else
                BaseQry = ReturnQry()
                Me.PvtGrid.RowGroupDescriptions.Clear()
                Me.PvtGrid.AggregateDescriptions.Clear()
                Me.PvtGrid.ColumnGroupDescriptions.Clear()
                dt = clsDBFuncationality.GetDataTable(BaseQry)
                If dt.Rows.Count > 0 Then
                    PvtGrid.DataSource = dt
                    Me.PvtGrid.PivotGridElement.AggregateDescriptorsArea.Visibility = ElementVisibility.Collapsed
                    Me.PvtGrid.PivotGridElement.ColumnDescriptorsArea.Visibility = ElementVisibility.Hidden
                    Me.PvtGrid.PivotGridElement.Margin = New Padding(0, -30, 0, 0)
                    Me.PvtGrid.PivotGridElement.VScrollBar.Margin = New Padding(0, 30, 0, 0)

                    Dim PropertyGroupDescription1 As PropertyGroupDescription = New PropertyGroupDescription()
                    Dim GroupNameComparer1 As GroupNameComparer = New GroupNameComparer()
                    Dim PropertyGroupDescription2 As PropertyGroupDescription = New PropertyGroupDescription()
                    Dim GroupNameComparer2 As GroupNameComparer = New GroupNameComparer()
                    Dim PropertyGroupDescription3 As PropertyGroupDescription = New PropertyGroupDescription()
                    Dim GroupNameComparer3 As GroupNameComparer = New GroupNameComparer()
                    Dim PropertyGroupDescription4 As PropertyGroupDescription = New PropertyGroupDescription()
                    Dim GroupNameComparer4 As GroupNameComparer = New GroupNameComparer()
                    Dim PropertyGroupDescription5 As PropertyGroupDescription = New PropertyGroupDescription()
                    Dim GroupNameComparer5 As GroupNameComparer = New GroupNameComparer()
                    Dim PropertyGroupDescription6 As PropertyGroupDescription = New PropertyGroupDescription()
                    Dim GroupNameComparer6 As GroupNameComparer = New GroupNameComparer()

                    Dim PropertyAggregateDescription1 As PropertyAggregateDescription = New PropertyAggregateDescription()
                    Dim SumAggregateFunction1 As SumAggregateFunction = New SumAggregateFunction()
                    Dim PropertyAggregateDescription2 As PropertyAggregateDescription = New PropertyAggregateDescription()
                    Dim SumAggregateFunction2 As SumAggregateFunction = New SumAggregateFunction()


                    PropertyAggregateDescription1.AggregateFunction = SumAggregateFunction1
                    PropertyAggregateDescription1.CustomName = "Qty"
                    PropertyAggregateDescription1.IgnoreNullValues = False
                    PropertyAggregateDescription1.PropertyName = "Report_UOM_Qty"
                    PropertyAggregateDescription1.StringFormat = Nothing
                    PropertyAggregateDescription1.StringFormatSelector = Nothing
                    PropertyAggregateDescription1.TotalFormat = Nothing
                    PropertyAggregateDescription2.AggregateFunction = SumAggregateFunction2
                    PropertyAggregateDescription2.CustomName = "Amount"
                    PropertyAggregateDescription2.IgnoreNullValues = False
                    PropertyAggregateDescription2.PropertyName = "Amount"
                    PropertyAggregateDescription2.StringFormat = Nothing
                    PropertyAggregateDescription2.StringFormatSelector = Nothing
                    PropertyAggregateDescription2.TotalFormat = Nothing

                    Me.PvtGrid.AggregateDescriptions.Add(PropertyAggregateDescription1)
                    Me.PvtGrid.AggregateDescriptions.Add(PropertyAggregateDescription2)
                    Me.PvtGrid.AllowFieldsDragDrop = True
                    Me.PvtGrid.ColumnGrandTotalsPosition = Telerik.WinControls.UI.TotalsPos.Last

                    PropertyGroupDescription2.AutoShowSubTotals = False
                    PropertyGroupDescription2.CustomName = "Sku_Seq"
                    PropertyGroupDescription2.GroupComparer = GroupNameComparer2
                    PropertyGroupDescription2.GroupFilter = Nothing
                    PropertyGroupDescription2.PropertyName = "Sku_Seq"
                    PropertyGroupDescription2.ShowGroupsWithNoData = False
                    PropertyGroupDescription2.SortOrder = SortOrder.Ascending
                    Me.PvtGrid.ColumnGroupDescriptions.Add(PropertyGroupDescription2)

                    PropertyGroupDescription1.AutoShowSubTotals = True
                    PropertyGroupDescription1.CustomName = "Item Name"
                    PropertyGroupDescription1.GroupComparer = GroupNameComparer1
                    PropertyGroupDescription1.GroupFilter = Nothing
                    PropertyGroupDescription1.PropertyName = "Short_Description"
                    PropertyGroupDescription1.ShowGroupsWithNoData = False
                    PropertyGroupDescription1.SortOrder = SortOrder.None
                    Me.PvtGrid.ColumnGroupDescriptions.Add(PropertyGroupDescription1)

                    Me.PvtGrid.ColumnsSubTotalsPosition = Telerik.WinControls.UI.TotalsPos.Last
                    Me.PvtGrid.Location = New System.Drawing.Point(0, 0)

                    'PropertyGroupDescription2.AutoShowSubTotals = True
                    'PropertyGroupDescription2.GroupComparer = GroupNameComparer2
                    'PropertyGroupDescription2.GroupFilter = Nothing
                    'PropertyGroupDescription2.ShowGroupsWithNoData = False
                    'PropertyGroupDescription2.SortOrder =  SortOrder.Ascending
                    'PropertyGroupDescription2.CustomName = "Customer"
                    'PropertyGroupDescription2.PropertyName = "Cust_Code"

                    PropertyGroupDescription3.AutoShowSubTotals = True
                    PropertyGroupDescription3.GroupComparer = GroupNameComparer3
                    PropertyGroupDescription3.GroupFilter = Nothing
                    PropertyGroupDescription3.ShowGroupsWithNoData = False
                    PropertyGroupDescription3.SortOrder = SortOrder.Ascending
                    PropertyGroupDescription3.CustomName = "Customer Name"
                    PropertyGroupDescription3.PropertyName = "Customer_Name"

                    PropertyGroupDescription4.AutoShowSubTotals = False
                    PropertyGroupDescription4.GroupComparer = GroupNameComparer4
                    PropertyGroupDescription4.GroupFilter = Nothing
                    PropertyGroupDescription4.ShowGroupsWithNoData = False
                    PropertyGroupDescription4.SortOrder = SortOrder.None
                    PropertyGroupDescription4.CustomName = "Supply Date"
                    PropertyGroupDescription4.PropertyName = "Supply_Date"

                    PropertyGroupDescription5.AutoShowSubTotals = True
                    PropertyGroupDescription5.GroupComparer = GroupNameComparer5
                    PropertyGroupDescription5.GroupFilter = Nothing
                    PropertyGroupDescription5.ShowGroupsWithNoData = False
                    PropertyGroupDescription5.SortOrder = SortOrder.Ascending
                    PropertyGroupDescription5.CustomName = "Route"
                    PropertyGroupDescription5.PropertyName = "Route_No"

                    PropertyGroupDescription6.AutoShowSubTotals = False
                    PropertyGroupDescription6.GroupComparer = GroupNameComparer6
                    PropertyGroupDescription6.GroupFilter = Nothing
                    PropertyGroupDescription6.ShowGroupsWithNoData = False
                    PropertyGroupDescription6.SortOrder = SortOrder.Descending
                    PropertyGroupDescription6.CustomName = "Shift"
                    PropertyGroupDescription6.PropertyName = "Shift_Type"

                    'Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription2)
                    Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription3)

                    If rbtnCustomer.IsChecked Then
                        If rbtnDetail.IsChecked Then
                            Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription4)
                            Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription6)
                        End If
                    ElseIf rbtnCustRoute.IsChecked Then
                        If rbtnDetail.IsChecked Then
                            Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription5)
                            Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription4)
                            Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription6)
                        ElseIf rbtnSummary.IsChecked Then
                            Me.PvtGrid.RowGroupDescriptions.Add(PropertyGroupDescription5)
                        End If
                    End If

                    Me.PvtGrid.RowsSubTotalsPosition = Telerik.WinControls.UI.TotalsPos.Last
                    Me.PvtGrid.PivotGridElement.BestFitHelper.BestFitColumns()
                    Me.PvtGrid.ColumnWidth = 90
                    FindAndRestoreGridLayout(Me)
                    RadPageView1.SelectedPage = RadPageViewPage2
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)

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
    Private Function ReturnQry() As String
        If rbtnDetail.IsChecked Then
            If rbtnCustRoute.IsChecked Then
                If txtCustomer.arrValueMember IsNot Nothing Then
                    If txtCustomer.arrValueMember.Count > 1 Then
                        clsCommon.MyMessageBoxShow(Me, "You cannot select more than one Customer at a time with Detail with Customer-Route Wise option", Me.Text)
                        Return ""
                        Exit Function
                    End If
                End If
            End If
        End If

        Dim strQry As String = Nothing
        Dim BaseQry As String = Nothing
        Dim whrcls As String = ""

        If Not chkDisplayMargin.IsChecked Then

            whrcls = " And Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "', 103)  and   convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= CONVERT(DATE, '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "', 103)"

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                If rbtnMorning.IsChecked Then
                    If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                        whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='PM' then 3 else 2 end  )"
                    Else
                        whrcls += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' "
                    End If
                ElseIf rbtnEvening.IsChecked Then
                    If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                        whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 3 else 2 end  )"
                    Else
                        whrcls += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'PM' "
                    End If
                ElseIf rbtnBothShift.IsChecked Then
                    If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                        If txtFromDate.Value.Day <> 1 Then
                            whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 3 else 2 end  )"
                        End If
                    Else
                        If txtFromDate.Value.Day <> 1 Then
                            whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 3 else 2 end  )"
                        End If
                        If (txtToDate.Value.Day <> DateTime.DaysInMonth(txtToDate.Value.Year, txtToDate.Value.Month)) Then
                            whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='PM' then 3 else 2 end  )"
                        End If
                    End If
                End If
            Else
                If rbtnMorning.IsChecked Then
                    whrcls += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' "
                ElseIf rbtnEvening.IsChecked Then
                    whrcls += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'PM' "
                End If
            End If

        End If

        whrcls += " and TSPL_ITEM_MASTER.Is_FreshItem = 1  "

        If txtRoute.arrValueMember IsNot Nothing Then
            whrcls += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ")"
        End If

        If txtCustomer.arrValueMember IsNot Nothing Then
            whrcls += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") "
        End If

        If chkDisplayMargin.IsChecked Then
            BaseQry = "TSPL_SD_SHIPMENT_HEAD.Shift_Type,TSPL_CUSTOMER_MASTER.Cust_Code As [Party Code],TSPL_CUSTOMER_MASTER.Customer_Name As [Party Name],TSPL_SD_SHIPMENT_HEAD.Route_No,TSPL_SD_SHIPMENT_HEAD.Route_Desc,TSPL_CUSTOMER_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description As Zone_Desc,TSPL_ITEM_MASTER.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,TSPL_SD_SALE_INVOICE_Detail.Qty,TSPL_SD_SALE_INVOICE_Detail.Unit_code,((TSPL_SD_SALE_INVOICE_Detail.Qty * ISNULL(TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 1)) / I.Conversion_Factor) AS CinCFQty,I.UOM_Code,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,I.Conversion_Factor As CinCF,TSPL_SD_SALE_INVOICE_Detail.Amt_Less_Discount As [Basic Amount],TSPL_SD_SALE_INVOICE_Detail.Disc_Amt,TSPL_SD_SALE_INVOICE_Detail.Total_Tax_Amt,TSPL_SD_SALE_INVOICE_Detail.Item_Net_Amt As [Total Amount],TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_PKID,TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate,TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,TSPL_SD_SHIPMENT_DETAIL.Crate "
        Else
            BaseQry = " Select  '" & objCommonVar.CurrentUserCode & "' as UserName,tspl_company_master.State,tspl_company_master.City_Code,tspl_company_master.Circle_No,tspl_company_master.Phone1,tspl_company_master.Phone2,   tspl_company_master.Comp_Name,tspl_company_master.Add1,tspl_company_master.Add2,tspl_company_master.Pincode,tspl_company_master.Fax,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + "' as FromDate,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy") + "' as ToDate, TSPL_ITEM_MASTER.Item_Code ,Sku_Seq, TSPL_CUSTOMER_MASTER.Cust_Code ,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_SD_SALE_INVOICE_HEAD.Route_No, TSPL_SD_SALE_INVOICE_HEAD.Route_Desc,"
            If rbtnDetail.IsChecked Then
                BaseQry += " Case When isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'AM' THEN 'Morning' else 'Evening' END AS Shift_Type,convert(varchar, Supply_Date,103) as Supply_Date,"
            End If
            BaseQry += " TSPL_ITEM_MASTER.Item_Desc, TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt As Amount, TSPL_ITEM_MASTER.Short_Description + ' (' +  isnull( I.UOM_Code,'') + ' )' as Short_Description,isnull((TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Report_UOM_Qty "
        End If
        BaseQry += " From TSPL_SD_SALE_INVOICE_DETAIL 
Left OUTER Join TSPL_SD_SALE_INVOICE_HEAD On TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE 
left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No 
Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_Detail.Item_Code
Left Outer Join TSPL_SD_SHIPMENT_DETAIL On TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE =TSPL_SD_SHIPMENT_HEAD.Document_Code And TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_SD_SALE_INVOICE_DETAIL.Unit_code	
LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON TSPL_SD_SALE_INVOICE_DETAIL.Item_Code = I.item_code
Left OUTER Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
Left Outer Join TSPL_ZONE_MASTER On TSPL_ZONE_MASTER.Zone_Code=TSPL_CUSTOMER_MASTER.Zone_Code
left outer join tspl_company_master on 2 = 2  "
        BaseQry += " where 2 = 2   and TSPL_SD_SALE_INVOICE_HEAD.Status = 1 " & whrcls & " "
        If Not chkDisplayMargin.IsChecked Then
            BaseQry += " order by convert(date,Supply_Date, 103), TSPL_SD_SHIPMENT_HEAD.Shift_Type, Sku_Seq "
        End If

        If chkDisplayMargin.IsChecked Then
            strQry = "Select Supply_Date,[Party Code],MAX([Party Name])[Party Name],Route_No,MAX(Route_Desc)Route_Desc,Zone_Code,MAX(Zone_Desc)Zone_Desc,Item_Code,
            MAX(Item_Desc)Item_Desc,MAX(Short_Description)Short_Description,MAX(Sku_Seq)Sku_Seq,SUM(Qty)Qty,MAX(Unit_code)Unit_code,Sum(CinCFQty)CinCFQty,MAX(UOM_Code)Report_UOM,SUM([Basic Amount])[Basic Amount],
            SUM(Disc_Amt)Disc_Amt,Max(Distributor_Commission_Rate) As Margin_Rate,SUM(Distributor_Commission_Amt) As Margin_Amt,
            SUM(Total_Tax_Amt)Total_Tax_Amt,SUM([Total Amount])[Total Amount],SUM(Crate)Crate from ("
            strQry += " Select dateadd(day,1,TSPL_SD_SHIPMENT_HEAD.Supply_Date)Supply_Date,IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate,0) As Margin_Rate,IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0) As Margin_Amt," + BaseQry + ""
            strQry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 1  
And TSPL_SD_SHIPMENT_HEAD.Shift_Type='PM' 
 and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)>=dateadd(day,-1,convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103)) 
 AND convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)<=dateadd(day,-1,convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "',103)) "
            strQry += " Union All "
            strQry += " Select TSPL_SD_SHIPMENT_HEAD.Supply_Date,IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Rate,0) As Margin_Rate,IsNull(TSPL_SD_SHIPMENT_DETAIL.Distributor_Commission_Amt,0) As Margin_Amt," + BaseQry + ""
            strQry += " and TSPL_SD_SALE_INVOICE_HEAD.Status = 1  And TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' 
 and convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) 
 AND convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103)<=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "',103)"
            strQry += " )BaseQry Group By Supply_Date,[Party Code],Route_No,Zone_Code,Item_Code"

            BaseQry = Nothing
            BaseQry = strQry
        End If


        Return BaseQry
    End Function
    Public Function ReturnRouteCustWiseQry() As String
        Dim strQry As String = Nothing
        Dim BaseQry As String = Nothing
        Dim whrcls As String = ""

        whrcls = " And Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "', 103)  and   convert(date,TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= CONVERT(DATE, '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "', 103)"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            If rbtnMorning.IsChecked Then
                If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                    whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='PM' then 3 else 2 end  )"
                Else
                    whrcls += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' "
                End If
            ElseIf rbtnEvening.IsChecked Then
                If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                    whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 3 else 2 end  )"
                Else
                    whrcls += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'PM' "
                End If
            ElseIf rbtnBothShift.IsChecked Then
                If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                    If txtFromDate.Value.Day <> 1 Then
                        whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 3 else 2 end  )"
                    End If
                Else
                    If txtFromDate.Value.Day <> 1 Then
                        whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='AM' then 3 else 2 end  )"
                    End If
                    If (txtToDate.Value.Day <> DateTime.DaysInMonth(txtToDate.Value.Year, txtToDate.Value.Month)) Then
                        whrcls += " and 2=( case when Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SHIPMENT_HEAD.Supply_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_SD_SHIPMENT_HEAD.Shift_Type='PM' then 3 else 2 end  )"
                    End If
                End If
            End If
        Else
            If rbtnMorning.IsChecked Then
                whrcls += " And TSPL_SD_SHIPMENT_HEAD.Shift_Type = 'AM' "
            ElseIf rbtnEvening.IsChecked Then
                whrcls += " and TSPL_SD_SHIPMENT_HEAD.Shift_Type  = 'PM' "
            End If
        End If
        whrcls += " and TSPL_ITEM_MASTER.Is_FreshItem = 1  "

        If txtRoute.arrValueMember IsNot Nothing Then
            whrcls += " and TSPL_SD_SALE_INVOICE_HEAD.Route_No in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ")"
        End If

        If txtCustomer.arrValueMember IsNot Nothing Then
            whrcls += " and TSPL_SD_SALE_INVOICE_HEAD.Customer_Code in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") "
        End If
        BaseQry = " Select  tspl_company_master.State,tspl_company_master.City_Code,tspl_company_master.Circle_No,tspl_company_master.Phone1,tspl_company_master.Phone2, TSPL_CUSTOMER_MASTER.Customer_Name as Booth,  tspl_company_master.Comp_Name,tspl_company_master.Add1,tspl_company_master.Add2,tspl_company_master.Pincode,tspl_company_master.Fax,'" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + "' as FromDate,'" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy") + "' as ToDate, TSPL_ITEM_MASTER.Item_Code ,Sku_Seq, TSPL_SD_SALE_INVOICE_HEAD.Customer_Code ,TSPL_CUSTOMER_MASTER.Customer_Name, TSPL_SD_SALE_INVOICE_HEAD.Route_No, TSPL_SD_SALE_INVOICE_HEAD.Route_Desc,"
        BaseQry += " TSPL_ITEM_MASTER.IsTaxable, Case When isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'AM' THEN 'Morning' else 'Evening' END AS Shift_Type,convert(varchar, Supply_Date,103) as Supply_Date,"
        BaseQry += " TSPL_ITEM_MASTER.Item_Desc,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,Display_Seq, TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt As Amount, TSPL_ITEM_MASTER.Short_Description ,isnull((TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1)) /(I.Conversion_Factor),0) As Report_UOM_Qty "
        BaseQry += " From TSPL_SD_SALE_INVOICE_DETAIL 
Left OUTER Join TSPL_SD_SALE_INVOICE_HEAD On TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE 
left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No 
Left Outer Join TSPL_ITEM_MASTER On TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_Detail.Item_Code
Left Outer Join TSPL_SD_SHIPMENT_DETAIL On TSPL_SD_SHIPMENT_DETAIL.DOCUMENT_CODE =TSPL_SD_SHIPMENT_HEAD.Document_Code And TSPL_SD_SHIPMENT_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_SD_SALE_INVOICE_DETAIL.Unit_code	
LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON TSPL_SD_SALE_INVOICE_DETAIL.Item_Code = I.item_code
Left OUTER Join TSPL_CUSTOMER_MASTER On TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_SD_SALE_INVOICE_HEAD.Customer_Code
Left Outer Join TSPL_ZONE_MASTER On TSPL_ZONE_MASTER.Zone_Code=TSPL_CUSTOMER_MASTER.Zone_Code
left outer join tspl_company_master on 2 = 2  "
        BaseQry += " where 2 = 2   and TSPL_SD_SALE_INVOICE_HEAD.Status = 1 " & whrcls & " "
        Return BaseQry
    End Function
    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        Dim s As String = "default.xml"
        Dim dialog As New SaveFileDialog()
        dialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*"
        dialog.Title = "Select a xml file"
        If dialog.ShowDialog() = DialogResult.OK Then
            s = dialog.FileName
        End If
        Me.PvtGrid.SaveLayout(s)
    End Sub

    'Private Sub RadButtonLoadLayout_Click(sender As Object, e As EventArgs) Handles RadButtonLoadLayout.Click
    '    Dim s As String = "default.xml"
    '    Dim dialog As New OpenFileDialog()
    '    dialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*"
    '    dialog.Title = "Select a xml file"
    '    If dialog.ShowDialog() = DialogResult.OK Then
    '        s = dialog.FileName
    '    End If
    '    Me.RadPivotGrid1.LoadLayout(s)
    'End Sub   

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Try
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim filePath As String
            sfd.FileName = Me.Text
            sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
            If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                filePath = sfd.FileName
            Else
                Exit Sub
            End If
            If chkDisplayMargin.IsChecked Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Margin Report")
                arrHeader.Add("Date: " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " to " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy") + "  ")
                clsCommon.MyExportToExcelGrid("Margin Report", gv, arrHeader, "Margin Report")
            Else
                Dim Export As New PivotExportToExcelML(Me.PvtGrid)
                Export.RunExport(filePath)
                common.clsCommon.MyMessageBoxShow("Exported Successfully.")
                Process.Start(filePath)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        Try
            Dim arrHeader As String = ""
            arrHeader = "Company : " & objCommonVar.CurrentCompanyName + Environment.NewLine
            arrHeader += "Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptMilkSupplySaleReport & "'")

            Dim doc As New RadPrintDocument()
            Dim style As PivotGridPrintStyle = New PivotGridPrintStyle()

            style.HeadersBackColor = Color.White
            style.GrandTotalsBackColor = Color.White
            PvtGrid.PrintStyle = style
            style.GridLinesColor = Color.Black
            style.CellBackColor = Color.White
            style.SubTotalsBackColor = Color.White
            style.SubTotalCellsFont = New Font("Segoe UI", 8, FontStyle.Bold)
            style.DataCellsFont = New Font("Segoe UI", 7.5, FontStyle.Regular)
            style.HeaderCellsFont = New Font("Segoe UI", 7.5, FontStyle.Bold)
            doc.Margins.Top = 50
            doc.Margins.Bottom = 50
            doc.Margins.Left = 20
            doc.Margins.Right = 100
            doc.LeftHeader = "[Logo]"
            doc.Margins.Left = 50
            doc.Margins.Right = 100
            doc.HeaderHeight = 80
            doc.Landscape = True
            doc.PaperSize = New System.Drawing.Printing.PaperSize("A3", 1169, 1654)
            doc.LeftFooter = "User:" + objCommonVar.CurrentUser + " " + " Print Date:" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd-MMM-yyyy hh:mm:ss tt")
            doc.HeaderFont = New Font("Segoe UI", 9, FontStyle.Bold)
            doc.RightFooter = "Page [Page #] Of [Total Pages]"
            doc.MiddleHeader = arrHeader
            arrHeader = " Date : " & clsCommon.myCDate(txtFromDate.Value) + "  To " + clsCommon.myCDate(txtToDate.Value) + Environment.NewLine
            If txtRoute.arrValueMember IsNot Nothing Then
                arrHeader += "Route No : " & clsCommon.GetMulcallString(txtRoute.arrValueMember) & "  Route Name :" & clsCommon.GetMulcallString(txtRoute.arrDispalyMember) & "" + Environment.NewLine
            End If

            If txtCustomer.arrValueMember IsNot Nothing Then
                arrHeader += "Customer Code : " & clsCommon.GetMulcallString(txtCustomer.arrValueMember) & "   Customer Name :" & clsCommon.GetMulcallString(txtCustomer.arrDispalyMember) & "" + Environment.NewLine
            End If

            If rbtnSummary.IsChecked = True Then
                arrHeader += "Report Type : " & "Summary" + Environment.NewLine
            ElseIf rbtnDetail.IsChecked = True Then
                arrHeader += "Report Type : " & "Details"
            End If

            doc.RightHeader = arrHeader
            doc.AssociatedObject = PvtGrid

            Dim dialog As New RadPrintPreviewDialog
            dialog.Document = doc
            dialog.ToolMenu.Visible = True
            dialog.SetZoom(1)
            dialog.WindowState = FormWindowState.Maximized
            dialog.ShowDialog(Me)
            doc = Nothing

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

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim BaseQry As String = ""
            If chkRouteWise.Checked OrElse chkRouteShiftCustWise.Checked Then
                BaseQry = ReturnRouteCustWiseQry()
            Else
                BaseQry = ReturnQry()
            End If

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(BaseQry)
            If chkRouteWise.Checked OrElse chkRouteShiftCustWise.Checked Then
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    clsCommon.MyMessageBoxShow("No Data Found to Print")
                    Exit Sub
                ElseIf dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    Dim dtItems As DataTable = clsDBFuncationality.GetDataTable("select Item_Code,max(Sku_Seq) as Sku_Seq,max(Short_Description) as Short_Description from (" + BaseQry + ") x group by Item_Code order by Sku_Seq")
                    Dim FinalQuery As String = " With CTERawData as ( " + BaseQry + "  )" + Environment.NewLine + Environment.NewLine
                    Dim shift As String = ""
                    If rbtnMorning.IsChecked Then
                        shift = "Morning"
                    ElseIf rbtnEvening.IsChecked Then
                        shift = "Evening"
                    End If
                    For ii As Integer = 1 To dtItems.Rows.Count Step 10
                        If ii > 1 Then
                            FinalQuery += Environment.NewLine + " Union all " + Environment.NewLine
                        End If
                        FinalQuery += " select '" & objCommonVar.CurrentUser & "' as UserName," + clsCommon.myCstr(ii) + " as Grp ,'" + shift + "' as FromShift, ROW_NUMBER() over (order by max(Booth)) As SNo, '" + clsCommon.GetPrintDate(clsCommon.myCDate(txtFromDate.Value), "dd/MM/yyyy") + "' AS FromDate,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtToDate.Value), "dd/MM/yyyy") + "' AS ToDate, max(Comp_Name) as Comp_Name,max(Add1) as Add1,max(Booth) as Booth,"

                        If chkRouteWise.Checked Then
                            FinalQuery += " Route_No,max(Route_Desc)Route_Desc,max(Customer_Code)Customer_Code "
                        ElseIf chkRouteShiftCustWise.Checked Then
                            FinalQuery += " Route_No,max(Route_Desc)Route_Desc,Customer_Code "
                        End If
                        FinalQuery += ",max(Document_Date) as Document_Date,"
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
                            End If
                            FinalQuery += " '" & strICODE & "' as Item_" & jj & ",
        '" & strIShortDesc & "' as Item_Short_Description_" & jj & ",
        (SUM(CASE WHEN Item_Code='" & strICODE & "' THEN Report_UOM_Qty ELSE NULL END)) as ItemQtyCrate_" & jj & ",
                           (SUM(CASE WHEN Item_Code= '" & strICODE & "' THEN Amount * CASE WHEN IsTaxable = 0 THEN 1 ELSE 0 END ELSE NULL END )) as ItemAmt_" & jj & ","
                        Next
                        If ii > 1 Then
                            FinalQuery += " null as Amount,null as TotalQty "
                        Else
                            FinalQuery += " sum(Amount*case when IsTaxable=0 then 1 else 0 end) as Amount ,sum(Report_UOM_Qty)TotalQty"
                        End If
                        FinalQuery += ",max(Display_Seq) as Display_Seq from ( select xx.*	from CTERawData xx ) x  "
                        If chkRouteWise.Checked Then
                            FinalQuery += " group by route_no  "
                        ElseIf chkRouteShiftCustWise.Checked Then
                            FinalQuery += " group by route_no,Customer_Code  "
                        End If
                    Next
                    If chkRouteWise.Checked Then
                        FinalQuery += " order by grp,Route_Desc "
                    ElseIf chkRouteShiftCustWise.Checked Then
                        FinalQuery += " order by grp,Route_Desc,Booth "
                    End If
                    dt = clsDBFuncationality.GetDataTable(FinalQuery + " ")
                End If
            End If
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If chkRouteShiftCustWise.Checked Then
                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptMilkSupplySaleShiftRouteCustWiseSummary", "Milk Supply Sale Shift Route Customer Wise Summary")
                ElseIf chkRouteWise.Checked Then
                    frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, Nothing, "crptMilkSupplySaleShiftRouteWiseSummary", "Milk Supply Sale Shift Route Wise Summary")
                Else
                        If rbtnDetail.IsChecked Then
                            If rbtnCustRoute.IsChecked Then
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustRouteWiseDetail", "")
                            ElseIf rbtnCustomer.IsChecked Then
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustomerWiseDetail", "")
                            End If
                        ElseIf rbtnSummary.IsChecked Then
                            If rbtnCustRoute.IsChecked Then
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustRouteWiseSummary", "")
                            ElseIf rbtnCustomer.IsChecked Then
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustomerWiseSummary", "")
                            End If
                        End If
                    End If
                    frmCRV = Nothing
                End If

            Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub chkMargin()
        Try
            If chkDisplayMargin.IsChecked Then
                RadGroupBox3.Visible = False
                RadGroupBox5.Visible = False
                btnPrint.Enabled = False
            Else
                RadGroupBox3.Visible = True
                RadGroupBox5.Visible = True
                btnPrint.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkDisplayMargin_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkDisplayMargin.ToggleStateChanged
        chkMargin()
    End Sub
End Class
