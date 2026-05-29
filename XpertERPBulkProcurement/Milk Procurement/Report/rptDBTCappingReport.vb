Imports common
Imports System.IO
Public Class rptDBTCappingReport
    Inherits FrmMainTranScreen

#Region "Variables"
    Const ReportID As String = "DBTCappingReport"
    Dim SettMPIncentiveEntryCycleWiseButNEFTMonthly As Boolean = False
    Dim SettMPIncentiveEntryApplyMonthly As Boolean = False
#End Region
    Private Sub rptDBTCappingReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SettMPIncentiveEntryCycleWiseButNEFTMonthly = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryCycleWiseButNEFTMonthly, clsFixedParameterCode.MPIncentiveEntryCycleWiseButNEFTMonthly, Nothing))
        SettMPIncentiveEntryApplyMonthly = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, Nothing))
        funreset()
    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        funreset()
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        GetReportID()
        Print(False)
    End Sub

    Sub GetReportID()
        Dim VarID As String = ""
        If rbtnAll.Checked Then
            VarID += "_A"
        ElseIf rbtnCappingRequired.Checked Then
            VarID += "_CR"
        ElseIf rbtnCappingApproved.Checked Then
            VarID += "_CA"
        End If
        gv1.VarID = VarID
    End Sub
    Sub funreset()
        EnableDisableControls(True)
        gv1.DataSource = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
        rbtnAll.Checked = True
        txtfromDate.Value = clsCommon.GETSERVERDATE()
        txtToDate.Value = clsCommon.GETSERVERDATE()
    End Sub
    Private Sub EnableDisableControls(ByVal val As Boolean)
        RadGroupBox3.Enabled = val
    End Sub

    Private Sub Print(ByVal print As Boolean)
        Try
            Dim requiredCapping As String = ""
            Dim cappingApproved As String = ""
            If rbtnCappingRequired.Checked Then
                requiredCapping = " and isnull(TSPL_DBT_CAPING_DETAIL.Capping_Status,0)=0 "
            End If
            If rbtnCappingApproved.Checked Then
                cappingApproved = " and TSPL_DBT_CAPING_DETAIL.Capping_Increase_By is not null"
            End If
            Dim qry As String = "select TSPL_DBT_CAPING.Document_Code as [Capping Check No] ,convert(varchar,TSPL_DBT_CAPING.Document_Date,103) as [Check Date],TSPL_DBT_CAPING.Reco_Code as [Reco No],convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103) as [Reco From Date],convert(varchar,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To,103) as [Reco To Date],TSPL_DBT_CAPING_DETAIL.DCS_Code as [DCS Code],TSPL_VLC_MASTER_HEAD.VLC_Name as [DCS Name], TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [DCS Uploader Code],TSPL_DBT_CAPING_DETAIL.MP_Code as [Farmar Code],TSPL_MP_MASTER.MP_Name as [Farmar Name],TSPL_MP_MASTER.MP_Code_VLC_Uploader as [Farmar Uploader Code],TSPL_DBT_CAPING_DETAIL.Qty as [DBT Quantity], TSPL_DBT_CAPING_DETAIL.Capping_Qty as [DBT Capping qty],TSPL_DBT_CAPING_DETAIL.Capping_Increase_By as [Capping By],TSPL_DBT_CAPING_DETAIL.Capping_Increase_Date as [Capping on],TSPL_MP_MASTER.THIRD_PARTY_CODE as [Third Party Code],TSPL_DBT_CAPING_DETAIL.Capping_Increase_Remarks as [Capping Remarks]
              from " & Environment.NewLine & " TSPL_DBT_CAPING_DETAIL
               left outer join TSPL_DBT_CAPING on TSPL_DBT_CAPING.Document_Code = TSPL_DBT_CAPING_DETAIL.Document_Code
               left outer join TSPL_DCS_MP_INCENTIVE_RECO_HEAD on TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code = TSPL_DBT_CAPING.Reco_Code
               left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DBT_CAPING_DETAIL.DCS_Code
               left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_DBT_CAPING_DETAIL.MP_Code
               where 2=2 and CONVERT(date,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date,103)=convert(date,'" & txtfromDate.Text & "',103) and CONVERT(date,TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To,103)= Convert(date,'" & txtToDate.Text & "',103) " + requiredCapping + " " + cappingApproved + ""
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
        gv1.TableElement.TableHeaderHeight = 30
        gv1.MasterTemplate.ShowRowHeaderColumn = True
        For ii As Integer = 0 To gv1.Columns.Count - 1
            gv1.Columns(ii).ReadOnly = True
            gv1.Columns(ii).BestFit()
        Next
        gv1.ShowGroupPanel = False
        '' gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(ReportID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(ReportID, "", objCommonVar.CurrentUserCode), clsGridLayout)
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
    Function SetToDate() As Boolean
        Try
            Dim PaymentCycleType As String = ""
            Dim PaymentCycleValue As Integer = 0
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(" select Payment_Cycle,PC_TYPE,PC_VALUE from ( select TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE as Payment_Cycle,
 TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE  from TSPL_PAYMENT_CYCLE_MASTER  where 
 TSPL_PAYMENT_CYCLE_MASTER.IsDefault=1 ) xx group by Payment_Cycle,PC_TYPE,PC_VALUE")
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("No Payment Cycle found on current MCC/Location")
            End If
            If dt.Rows.Count > 1 Then
                Throw New Exception("Selected MCC's Payment Cycle Should be Same")
            End If
            If SettMPIncentiveEntryApplyMonthly OrElse SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
                txtfromDate.Value = New Date(txtfromDate.Value.Year, txtfromDate.Value.Month, 1)
                txtToDate.Value = txtfromDate.Value.AddMonths(1).AddDays(-1)
            Else
                PaymentCycleType = clsCommon.myCstr(dt.Rows(0)("PC_TYPE"))
                PaymentCycleValue = clsCommon.myCdbl(dt.Rows(0)("PC_VALUE"))
                Dim dtCurr As DateTime = clsCommon.GETSERVERDATE()
                If clsCommon.CompairString(PaymentCycleType, "Day") = CompairStringResult.Equal Then
                    If txtfromDate.Value.Day Mod PaymentCycleValue <> 1 And (Not PaymentCycleValue = 1) Then
                        txtfromDate.Value = New Date(dtCurr.Year, dtCurr.Month, 1)
                        txtToDate.Value = txtfromDate.Value
                        Throw New Exception("Date can only be first day of month or at interval of " & PaymentCycleValue & " Day, Because MCC has payment Cycle of " & PaymentCycleValue & " Day ")
                    End If
                    txtToDate.Value = txtfromDate.Value.AddDays(PaymentCycleValue - 1)

                    If txtfromDate.Value.Month <> txtToDate.Value.Month Then
                        txtToDate.Value = New Date(txtfromDate.Value.Year, txtfromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
                    End If
                    Dim dtNxtPay As DateTime = txtToDate.Value.AddDays(Math.Ceiling(PaymentCycleValue / 2.0))
                    If txtfromDate.Value.Month <> dtNxtPay.Month Then
                        txtToDate.Value = New Date(txtfromDate.Value.Year, txtfromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
                    End If
                ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
                    If clsCommon.myCdbl(clsCommon.GetPrintDate(txtfromDate.Value, "dd")) <> 1 Then

                        txtfromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                        txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                        Throw New Exception("Date can only be first day of month, Because MCC has payment Cycle of Month Type")
                    End If
                    txtToDate.Value = DateAdd(DateInterval.Month, PaymentCycleValue, txtfromDate.Value)
                ElseIf clsCommon.CompairString(PaymentCycleType, "Year") = CompairStringResult.Equal Then
                    If clsCommon.myCdbl(clsCommon.GetPrintDate(txtfromDate.Value, "dd")) <> 1 Then
                        txtfromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                        txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                        Throw New Exception("Date can only be first day of month, Because MCC has payment Cycle of Year Type")
                    End If
                    txtToDate.Value = DateAdd(DateInterval.Year, PaymentCycleValue, txtfromDate.Value)
                ElseIf clsCommon.CompairString(PaymentCycleType, "Week") = CompairStringResult.Equal Then
                    Dim today As Date = txtfromDate.Value
                    Dim dayDiff As Integer = today.DayOfWeek - IIf(PaymentCycleValue = 1, DayOfWeek.Sunday, IIf(PaymentCycleValue = 2, DayOfWeek.Monday, IIf(PaymentCycleValue = 3, DayOfWeek.Tuesday, IIf(PaymentCycleValue = 4, DayOfWeek.Wednesday, IIf(PaymentCycleValue = 5, DayOfWeek.Thursday, IIf(PaymentCycleValue = 6, DayOfWeek.Friday, DayOfWeek.Saturday))))))
                    txtfromDate.Value = today.AddDays(-dayDiff)
                    txtToDate.Value = txtfromDate.Value.AddDays(6)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
        Return True
    End Function
    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", "Information", Me.Text)
    End Sub
    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Try
            If gv1.Rows.Count > 0 Then
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptDBTCappingReport & "'"))
                arrHeader.Add("Date : " & clsCommon.myCDate(txtfromDate.Value) + "  To " + clsCommon.myCDate(txtToDate.Value))
                transportSql.applyExportTemplate(gv1, PageSetupReport_ID)
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
                arrHeader.Add("Date : " & clsCommon.myCDate(txtfromDate.Value) + "  To " + clsCommon.myCDate(txtToDate.Value))

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

    Private Sub btnPrint_Click(sender As Object, e As EventArgs)
        Print(True)
    End Sub

    Private Sub txtfromDate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtfromDate.Validating
        SetToDate()
    End Sub
End Class