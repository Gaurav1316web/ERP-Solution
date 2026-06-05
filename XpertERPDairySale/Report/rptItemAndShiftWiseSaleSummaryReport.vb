
Imports common
Imports System.IO
Imports Telerik.Pivot.Core
Imports Telerik.WinControls.Export
Imports Telerik.WinControls.UI
Imports Telerik.WinControls.UI.Export
Imports Telerik.Windows.Controls

Public Class rptItemAndShiftWiseSaleSummaryReport
    Inherits FrmMainTranScreen
    Dim ShowTodayDemandAsCurrentandUpcoming As Boolean = False
    Private Sub rptItemAndShiftWiseSaleSummaryReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        funreset()
        ShowTodayDemandAsCurrentandUpcoming = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AndroidDemandBooking, clsFixedParameterCode.ShowTodayDemandAsCurrentandUpcoming, Nothing))
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
    End Sub

    Private Sub txtRoute__My_Click(sender As Object, e As EventArgs) Handles txtRoute._My_Click
        Try
            Dim qry As String = "select distinct TSPL_SD_SALE_INVOICE_HEAD.Route_No as  [ROUTE NO],TSPL_ROUTE_MASTER.Route_Desc as [ROUTE NAME]  FROM TSPL_SD_SALE_INVOICE_DETAIL left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code 
           left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_SD_SALE_INVOICE_HEAD.Route_No
            where  TSPL_SD_SALE_INVOICE_HEAD.Status = 1 AND convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >=Convert(date,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "',103) 
            and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(date,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "',103)   "

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

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        funreset()
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        GetReportGridID()
        EnableDisableControls(False)
        LoadData()
    End Sub
    Sub funreset()
        EnableDisableControls(True)
        gv1.DataSource = Nothing
        chkDCSSale.Checked = False
        RadPageView1.SelectedPage = RadPageViewPage1
        rbtnDemand.IsChecked = True
        rbtnMorning.IsChecked = True
    End Sub

    Private Sub EnableDisableControls(ByVal val As Boolean)
        RadGroupBox1.Enabled = val
    End Sub
    Sub GetReportGridID()
        Dim VarID As String = ""
        If rbtnDemand.IsChecked Then
            VarID += "_DE"
        ElseIf rbtnDispatch.IsChecked Then
            VarID += "_DI"
        End If
        If rbtnMorning.IsChecked Then
            VarID += "_MS"
        ElseIf rbtnEvening.IsChecked Then
            VarID += "_ES"
        ElseIf rbtnBothShift.IsChecked Then
            VarID += "_BS"
        End If
        gv1.VarID = VarID
    End Sub
    Private Sub LoadData()
        Try
            Dim qry As String = ""
            qry = ReturnQry()
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
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
                RadPageView1.SelectedPage = RadPageViewPage2
                gv1.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If

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
            gv1.Columns(ii).FormatString = "{0:n2}"
        Next
        gv1.ShowGroupPanel = False
        If rbtnMorning.IsChecked Then
            gv1.Columns("Morning_Qty").IsVisible = True
            gv1.Columns("Morning_Amt").IsVisible = True
            gv1.Columns("Evening_Qty").IsVisible = False
            gv1.Columns("Evening_Amt").IsVisible = False
        ElseIf rbtnEvening.IsChecked Then
            gv1.Columns("Morning_Qty").IsVisible = False
            gv1.Columns("Morning_Amt").IsVisible = False
            gv1.Columns("Evening_Qty").IsVisible = True
            gv1.Columns("Evening_Amt").IsVisible = True
        End If

        gv1.Columns("Short_Description").HeaderText = "Item"
        gv1.Columns("UOM_Code").HeaderText = "UOM"
        gv1.Columns("Evening_Qty").HeaderText = "Evening Quantity"
        gv1.Columns("Morning_Qty").HeaderText = "Morning Quantity"
        gv1.Columns("Evening_Amt").HeaderText = "Evening Amount"
        gv1.Columns("Morning_Amt").HeaderText = "Morning Amount"
        gv1.Columns("Total_Qty").HeaderText = "Total Quantity"
        gv1.Columns("Total_Amt").HeaderText = "Total Amount"

        'Dim summaryRowItem As New GridViewSummaryRowItem()

        'For ii As Integer = 2 To gv1.Columns.Count - 1
        '    summaryRowItem.Add(New GridViewSummaryItem(gv1.Columns(ii).Name, "{0:F2}", GridAggregateFunction.Sum))
        'Next
        'gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        'gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
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

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Try
            If gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptItemAndShiftWiseSaleSummaryReport & "'"))
                arrHeader.Add("Date : " & clsCommon.myCDate(txtFromDate.Value) + "  To " + clsCommon.myCDate(txtToDate.Value))
                If rbtnDemand.IsChecked Then
                    arrHeader.Add("Transaction Type : " & "Demand")
                ElseIf rbtnDispatch.IsChecked Then
                    arrHeader.Add("Transaction Type : " & "Dispatch")
                End If
                If txtRoute.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Route Code : " & clsCommon.GetMulcallString(txtRoute.arrValueMember) & "   Route Name :" & clsCommon.GetMulcallString(txtRoute.arrDispalyMember) & "")
                End If
                clsCommon.MyExportToExcelGrid(Me.Text, gv1, arrHeader, Me.Text)
            Else
                clsCommon.MyMessageBoxShow(Me, "No data found to export", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub btnPDF_Click(sender As Object, e As EventArgs) Handles btnPDF.Click
        Try
            If gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Date : " & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "  To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy"))
                Dim ReportHeading As String = ""
                If rbtnDemand.IsChecked Then
                    arrHeader.Add("Transaction Type : " & "Demand")
                ElseIf rbtnDispatch.IsChecked Then
                    arrHeader.Add("Transaction Type : " & "Dispatch")
                End If
                If txtRoute.arrValueMember IsNot Nothing Then
                    arrHeader.Add("Route Code : " & clsCommon.GetMulcallString(txtRoute.arrValueMember) & "   Route Name :" & clsCommon.GetMulcallString(txtRoute.arrDispalyMember) & "")
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

    'Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
    '    Dim BaseQry As String = ""
    '    BaseQry = ReturnQry()
    '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(BaseQry)
    '    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
    '        Dim frmCRV As New frmCrystalReportViewer()
    '        If rbtnDemand.IsChecked Then
    '            If rbtnCustRoute.IsChecked Then
    '                frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustRouteWiseDetail", "")
    '            ElseIf rbtnCustomer.IsChecked Then
    '                frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustomerWiseDetail", "")
    '            End If
    '        ElseIf rbtnDispatch.IsChecked Then
    '            If rbtnCustRoute.IsChecked Then
    '                frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustRouteWiseSummary", "")
    '            ElseIf rbtnCustomer.IsChecked Then
    '                frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "crptMilkSupplySaleCustomerWiseSummary", "")
    '            End If
    '        End If
    '    End If
    'End Sub
    Function ReturnQry()
        Dim qry As String = ""
        Try
            Dim BaseQry As String = ""
            Dim BaseQry2 As String = ""
            Dim whrcls As String = ""
            Dim whrShift As String = ""
            Dim whrDCSSaleShift As String = ""
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                If rbtnDemand.IsChecked Then
                    whrcls = " And convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "'   and convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
                ElseIf rbtnDispatch.IsChecked Then
                    whrcls = " And convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "'   and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
                End If
            Else

                If ShowTodayDemandAsCurrentandUpcoming Then
                    If rbtnDemand.IsChecked Then
                        whrcls = " And convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "'   and convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value.AddDays(1), "dd/MMM/yyyy") & "' "
                    ElseIf rbtnDispatch.IsChecked Then
                        whrcls = " And convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "'   and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value.AddDays(1), "dd/MMM/yyyy") & "' "
                    End If
                Else
                    If rbtnDemand.IsChecked Then
                        whrcls = " And convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "'   and convert(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
                    ElseIf rbtnDispatch.IsChecked Then
                        whrcls = " And convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= '" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "'   and convert(date,TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
                    End If
                End If
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                If rbtnMorning.IsChecked Then
                    If rbtnDemand.IsChecked Then
                        whrcls += " And TSPL_DEMAND_BOOKING_MASTER.ShiftType = 'Morning' "
                    ElseIf rbtnDispatch.IsChecked Then
                        whrShift += " And Shift_Type = 'AM'  "
                        whrDCSSaleShift += " And Shift_Type = 'AM'  "
                    End If
                ElseIf rbtnEvening.IsChecked Then
                    If rbtnDemand.IsChecked Then
                        whrcls += " and TSPL_DEMAND_BOOKING_MASTER.ShiftType  = 'Evening' "
                    ElseIf rbtnDispatch.IsChecked Then
                        whrDCSSaleShift += " And Shift_Type = 'PM'  "
                        whrShift += " And Shift_Type = 'PM'  "
                    End If
                End If
            Else
                If rbtnMorning.IsChecked Then
                    If rbtnDemand.IsChecked Then
                        If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                            whrcls += " and 2=( case when Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_DEMAND_BOOKING_MASTER.ShiftType='Evening' then 3 else 2 end  )"
                        Else
                            whrcls += " And TSPL_DEMAND_BOOKING_MASTER.ShiftType = 'Morning' "
                        End If
                    ElseIf rbtnDispatch.IsChecked Then
                        If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                            whrDCSSaleShift = " and 2=( case when Convert(Date, Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='PM' then 3 else 2 end  )"
                            whrShift = " and 2=( case when Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='PM' then 3 else 2 end  )"
                        Else
                            whrShift += " And Shift_Type = 'AM'  "
                            whrDCSSaleShift += " And Shift_Type = 'AM'  "
                        End If
                    End If
                ElseIf rbtnEvening.IsChecked Then
                    If rbtnDemand.IsChecked Then
                        If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                            whrcls += " and 2=( case when Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_DEMAND_BOOKING_MASTER.ShiftType='Morning' then 3 else 2 end  )"
                        Else
                            whrcls += " and TSPL_DEMAND_BOOKING_MASTER.ShiftType  = 'Evening' "
                        End If
                    ElseIf rbtnDispatch.IsChecked Then
                        If clsCommon.GetPrintDate(txtFromDate.Value, "dd-MMM-yyyy") = clsCommon.GetPrintDate(txtToDate.Value, "dd-MMM-yyyy") Then
                            whrDCSSaleShift += " and 2=( case when Convert(Date, Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='AM' then 3 else 2 end  )"
                            whrShift += " and 2=( case when Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='AM' then 3 else 2 end  )"
                        Else
                            whrDCSSaleShift += " And Shift_Type = 'PM'  "
                            whrShift += " And Shift_Type = 'PM'  "
                        End If
                    End If

                ElseIf rbtnBothShift.IsChecked Then
                    If rbtnDemand.IsChecked Then
                        If ShowTodayDemandAsCurrentandUpcoming Then
                            whrcls += " and 2=( case when Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_DEMAND_BOOKING_MASTER.ShiftType='Morning' then 3 else 2 end  )"
                            whrcls += " and 2=( case when Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value.AddDays(1)), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value.AddDays(1)), "dd/MMM/yyyy hh:mm tt") + "',103) and TSPL_DEMAND_BOOKING_MASTER.ShiftType='Evening' then 3 else 2 end  )"
                        End If
                    ElseIf rbtnDispatch.IsChecked Then
                        If ShowTodayDemandAsCurrentandUpcoming Then
                            whrShift += " and 2=( case when Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='AM' then 3 else 2 end  )"
                            whrShift += " and 2=( case when Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value.AddDays(1)), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, TSPL_SD_SALE_INVOICE_HEAD.Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value.AddDays(1)), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='PM' then 3 else 2 end  )"

                            whrDCSSaleShift += " and 2=( case when Convert(Date, Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='AM' then 3 else 2 end  )"
                            whrDCSSaleShift += " and 2=( case when Convert(Date, Document_Date,103) >= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value.AddDays(1)), "dd/MMM/yyyy hh:mm tt") + "',103) and Convert(Date, Document_Date,103) <= Convert(Date, '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value.AddDays(1)), "dd/MMM/yyyy hh:mm tt") + "',103) and Shift_Type='PM' then 3 else 2 end  )"
                        End If
                    End If
                End If
            End If


            If txtRoute.arrValueMember IsNot Nothing Then
                If rbtnDemand.IsChecked Then
                    whrcls += "  And TSPL_DEMAND_BOOKING_MASTER.Route_No In (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ")"
                ElseIf rbtnDispatch.IsChecked Then
                    If chkDCSSale.Checked Then
                        whrcls += " and isnull(TSPL_SD_SALE_INVOICE_HEAD.Route_No,'') in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ",'')"
                    Else
                        whrcls += " and isnull(TSPL_SD_SALE_INVOICE_HEAD.Route_No,'') in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ")"
                    End If
                End If
            End If
            Dim whrDCSSaleData As String = ""
            If rbtnDispatch.IsChecked Then
                If chkDCSSale.Checked Then
                    whrDCSSaleData = " and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type = 'MCC' "
                Else
                End If
            End If
            If rbtnDispatch.IsChecked Then
                If chkDCSSale.Checked Then
                    qry = " Select max(Short_Description)Short_Description,max(UOM_Code)UOM_Code,sum(Evening_Qty)Evening_Qty,sum(Morning_Qty)Morning_Qty,sum(Evening_Amt)Evening_Amt,sum(Morning_Amt)Morning_Amt,sum(Total_Qty)Total_Qty,sum(Total_Amt)Total_Amt from ( "
                Else
                    qry = " Select Short_Description,UOM_Code,Evening_Qty,Morning_Qty,Evening_Amt,Morning_Amt,Total_Qty,Total_Amt from ( "
                End If
            ElseIf rbtnDemand.IsChecked Then
                qry = " Select Short_Description,UOM_Code,Evening_Qty,Morning_Qty,Evening_Amt,Morning_Amt,Total_Qty,Total_Amt from ( "
            End If
            If rbtnDemand.IsChecked Then
                BaseQry = " max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,isnull(sum(case when isnull(TSPL_DEMAND_BOOKING_MASTER.ShiftType,'') = 'Evening' then TSPL_DEMAND_BOOKING_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor) else 0 end),0) as Evening_Qty,isnull(sum(case when isnull(TSPL_DEMAND_BOOKING_MASTER.ShiftType,'') = 'Morning' then TSPL_DEMAND_BOOKING_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor) else 0 end),0) as Morning_Qty,isnull(sum(case when isnull(TSPL_DEMAND_BOOKING_MASTER.ShiftType,'') = 'Evening' then TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount  else 0 end),0) as Evening_Amt,isnull(sum(case when isnull(TSPL_DEMAND_BOOKING_MASTER.ShiftType,'') = 'Morning' then TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount  else 0 end),0) as Morning_Amt,isnull(sum(TSPL_DEMAND_BOOKING_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor)),0) as Total_Qty,isnull(sum(TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount),0) as Total_Amt
            FROM TSPL_DEMAND_BOOKING_DETAIL left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code	LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON TSPL_DEMAND_BOOKING_DETAIL.Item_Code = I.item_code
            where  TSPL_DEMAND_BOOKING_MASTER.Posted = 1 " & whrcls & "	"
            ElseIf rbtnDispatch.IsChecked Then
                Dim sqlqry As String = " max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'PM' then TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor) else 0 end),0) as Evening_Qty,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'AM' then TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor) else 0 end),0) as Morning_Qty,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'PM' then TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt  else 0 end),0) as Evening_Amt,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'AM' then TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt  else 0 end),0) as Morning_Amt "
                BaseQry2 = " 'PM' as Shift_Type, max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'PM' then TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor) else 0 end),0) as Evening_Qty,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'AM' then TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor) else 0 end),0) as Morning_Qty,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'PM' then TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt  else 0 end),0) as Evening_Amt,isnull(sum(case when isnull(TSPL_SD_SHIPMENT_HEAD.Shift_Type,'') = 'AM' then TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt  else 0 end),0) as Morning_Amt "

                BaseQry = " ,isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.Qty * isnull((TSPL_ITEM_UOM_DETAIL.Conversion_Factor),1) /(I.Conversion_Factor)),0) as Total_Qty,isnull(sum(TSPL_SD_SALE_INVOICE_DETAIL.Item_Net_Amt),0) as Total_Amt
            FROM TSPL_SD_SALE_INVOICE_DETAIL left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_SD_SALE_INVOICE_DETAIL.Item_Code  left outer join TSPL_SD_SALE_INVOICE_HEAD on TSPL_SD_SALE_INVOICE_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_DETAIL.DOCUMENT_CODE left outer join TSPL_SD_SHIPMENT_HEAD on TSPL_SD_SHIPMENT_HEAD.Document_Code = TSPL_SD_SALE_INVOICE_HEAD.Against_Shipment_No
            LEFT JOIN  TSPL_ITEM_UOM_DETAIL ON TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_SD_SALE_INVOICE_DETAIL.Item_Code and TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_SD_SALE_INVOICE_DETAIL.Unit_code	LEFT JOIN  ( select item_code,uom_code,conversion_factor,UOM_Description from  TSPL_ITEM_UOM_DETAIL where Report_UOM = 1 ) as  I ON TSPL_SD_SALE_INVOICE_DETAIL.Item_Code = I.item_code
            where  TSPL_SD_SALE_INVOICE_HEAD.Status = 1 " & whrcls & " 	"
                BaseQry2 += BaseQry + whrDCSSaleData
                BaseQry = "" + sqlqry + BaseQry + whrShift
            End If
            If rbtnDispatch.IsChecked Then
                qry += "  SELECT 1 as SNO,max(TSPL_ITEM_MASTER.Short_Description)Short_Description, max(I.UOM_Code)UOM_Code, " & BaseQry & "  and TSPL_ITEM_MASTER.Is_FreshItem=1 and IsTaxable = 0 and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' group by TSPL_ITEM_MASTER.Item_Code  "
                If chkDCSSale.Checked Then
                    qry += "" & Environment.NewLine & " union all SELECT sno,max(Short_Description)Short_Description,max(UOM_Code)UOM_Code, max(Sku_Seq)Sku_Seq,sum(Evening_Qty) as Evening_Qty,0 as Morning_Qty,sum(Evening_Amt) as Evening_Amt,0 as Morning_Amt  ,sum(Total_Qty) as Total_Qty,isnull(sum(Total_Amt),0) as Total_Amt FROM ( SELECT 1 as SNO,max(TSPL_ITEM_MASTER.Short_Description)Short_Description, max(I.UOM_Code)UOM_Code,TSPL_ITEM_MASTER.Item_Code, TSPL_SD_SALE_INVOICE_HEAD.Document_Date," & BaseQry2 & "  and TSPL_ITEM_MASTER.Is_FreshItem=1 and IsTaxable = 0  group by TSPL_ITEM_MASTER.Item_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date)xxx where 2=2  " + whrDCSSaleShift + "  group by Item_Code,sno "
                End If

                qry += "" & Environment.NewLine & " union all " & Environment.NewLine & "  SELECT 2 as SNO, 'Milk Total' as Short_Description,'' as UOM_Code, " & BaseQry & "  and TSPL_ITEM_MASTER.Is_FreshItem=1 and IsTaxable = 0  and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' "
                If chkDCSSale.Checked Then
                    qry += "" & Environment.NewLine & " union all " & Environment.NewLine & " SELECT max(sno)sno,max(Short_Description)Short_Description,max(UOM_Code)UOM_Code, max(Sku_Seq)Sku_Seq,sum(Evening_Qty) as Evening_Qty,0 as Morning_Qty,sum(Evening_Amt) as Evening_Amt,0 as Morning_Amt  ,sum(Total_Qty) as Total_Qty,isnull(sum(Total_Amt),0) as Total_Amt FROM (  SELECT 2 as SNO, 'Milk Total' as Short_Description,'' as UOM_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date, " & BaseQry2 & "  and TSPL_ITEM_MASTER.Is_FreshItem=1 and IsTaxable = 0 group by TSPL_SD_SALE_INVOICE_HEAD.Document_Date)xxx where 2=2 " + whrDCSSaleShift + "  "
                End If

                qry += "" & Environment.NewLine & " union all " & Environment.NewLine & "  SELECT 3 as SNO, max(TSPL_ITEM_MASTER.Short_Description)Short_Description, max(I.UOM_Code)UOM_Code, " & BaseQry & " and (TSPL_ITEM_MASTER.Is_Ambient=1 or IsTaxable = 1 ) and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' group by TSPL_ITEM_MASTER.Item_Code  "
                If chkDCSSale.Checked Then
                    qry += "" & Environment.NewLine & " union all " & Environment.NewLine & " SELECT sno,max(Short_Description)Short_Description,max(UOM_Code)UOM_Code, max(Sku_Seq)Sku_Seq,sum(Evening_Qty) as Evening_Qty,0 as Morning_Qty,sum(Evening_Amt) as Evening_Amt,0 as Morning_Amt  ,sum(Total_Qty) as Total_Qty,isnull(sum(Total_Amt),0) as Total_Amt FROM ( SELECT 3 as SNO, max(TSPL_ITEM_MASTER.Short_Description)Short_Description, max(I.UOM_Code)UOM_Code, TSPL_ITEM_MASTER.Item_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date," & BaseQry2 & " and IsTaxable = 1  group by TSPL_ITEM_MASTER.Item_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date )xxx where 2=2  " + whrDCSSaleShift + " group by Item_Code,sno "
                End If
                qry += "" & Environment.NewLine & " union all " & Environment.NewLine & "  SELECT 4 as SNO, 'Product Total' as Short_Description,'' as UOM_Code, " & BaseQry & "  and (TSPL_ITEM_MASTER.Is_Ambient=1 or IsTaxable = 1 ) and TSPL_SD_SALE_INVOICE_HEAD.Trans_Type <> 'MCC' "
                If chkDCSSale.Checked Then
                    qry += "" & Environment.NewLine & " union all " & Environment.NewLine & " SELECT max(sno)sno,max(Short_Description)Short_Description,max(UOM_Code)UOM_Code, max(Sku_Seq)Sku_Seq,sum(Evening_Qty) as Evening_Qty,0 as Morning_Qty,sum(Evening_Amt) as Evening_Amt,0 as Morning_Amt  ,sum(Total_Qty) as Total_Qty,isnull(sum(Total_Amt),0) as Total_Amt FROM ( SELECT 4 as SNO, 'Product Total' as Short_Description,'' as UOM_Code,TSPL_SD_SALE_INVOICE_HEAD.Document_Date, " & BaseQry2 & "  and (TSPL_ITEM_MASTER.Is_Ambient=1 or IsTaxable = 1 ) group by TSPL_SD_SALE_INVOICE_HEAD.Document_Date)xxx where 2=2  " + whrDCSSaleShift + " "
                End If
                If chkDCSSale.Checked Then
                    qry += " ) xxxx where SNO is not null and total_qty >0	group by sno,Short_Description		order by SNO, max(Sku_Seq ) "
                Else
                    qry += " ) xx	order by SNO, Sku_Seq "
                End If
            ElseIf rbtnDemand.IsChecked Then
                qry += "  SELECT 1 as SNO,max(TSPL_ITEM_MASTER.Short_Description)Short_Description, max(I.UOM_Code)UOM_Code, " & BaseQry & "  and TSPL_ITEM_MASTER.Is_FreshItem=1 and IsTaxable = 0 group by TSPL_ITEM_MASTER.Item_Code  " & Environment.NewLine & " union all " & Environment.NewLine & "  SELECT 2 as SNO, 'Milk Total' as Short_Description,'' as UOM_Code, " & BaseQry & "  and TSPL_ITEM_MASTER.Is_FreshItem=1 and IsTaxable = 0  
             " & Environment.NewLine & " union all " & Environment.NewLine & "  SELECT 3 as SNO, max(TSPL_ITEM_MASTER.Short_Description)Short_Description, max(I.UOM_Code)UOM_Code, " & BaseQry & " and (TSPL_ITEM_MASTER.Is_Ambient=1 or IsTaxable = 1 ) group by TSPL_ITEM_MASTER.Item_Code  " & Environment.NewLine & " union all " & Environment.NewLine & "  SELECT 4 as SNO, 'Product Total' as Short_Description,'' as UOM_Code, " & BaseQry & "  and (TSPL_ITEM_MASTER.Is_Ambient=1 or IsTaxable = 1 ) "
                qry += " ) xx	order by SNO, Sku_Seq "
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return qry
    End Function

    Private Sub rbtnDemand_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnDemand.ToggleStateChanged
        If rbtnDemand.IsChecked Then
            chkDCSSale.Visible = False
        ElseIf rbtnDispatch.IsChecked Then
            chkDCSSale.Visible = True
        End If
    End Sub
End Class
