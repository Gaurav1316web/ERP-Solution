Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports CrystalDecisions.Windows.Forms
Imports common
Imports System.IO
Imports System.Globalization
Public Class RptMPWiseMilkCollectionAtPoolingPoint3

    Inherits FrmMainTranScreen
    Dim Slot1 As DateTime = Nothing
    Dim Slot2 As DateTime = Nothing
    Dim Slot3 As DateTime = Nothing
    Dim Slot4 As DateTime = Nothing
    Dim dt As DataTable
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim btnReferesh As Boolean = False
    Dim isPrint As Boolean = False
    Dim tmpValLoad As Boolean = True
    Dim arrLoc As String = Nothing
    Dim DisplayAverageFatSNFMPWise As Boolean = False
    Dim ItemStructureMandatoryOnWeightConversion As Boolean = False
    Dim Month1 As Date
    Dim Month2 As Date
    Private Sub LOCATIONRIGTHS()
        Try
            Dim obj As New clsMCCCodes()
            obj = clsMCCCodes.GetData()
            If obj.arrLocCodes IsNot Nothing AndAlso clsCommon.myLen(obj.arrLocCodes) > 0 Then
                arrLoc = obj.arrLocCodes
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.rptMPWiseMilkCollection)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        RadSplitButton1.Visible = MyBase.isExport
        btnprint.Visible = MyBase.isPrintFlag
    End Sub



    Sub LoadShiftFrom()
        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Code")
        dt.Columns.Add("Shift")

        Dim dr As DataRow = dt.NewRow
        dr("Code") = "M"
        dr("Shift") = "Morning"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "E"
        dr("Shift") = "Evening"
        dt.Rows.Add(dr)

        txtFromShift.DataSource = dt
        txtFromShift.ValueMember = "Code"
    End Sub
    Sub LoadShiftTo()
        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Code")
        dt.Columns.Add("Shift")

        Dim dr As DataRow = dt.NewRow
        dr("Code") = "M"
        dr("Shift") = "Morning"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "E"
        dr("Shift") = "Evening"
        dt.Rows.Add(dr)

        txtToShift.DataSource = dt
        txtToShift.ValueMember = "Code"
    End Sub

    Sub LoadEntrySource()
        Dim qry As String = Nothing
        qry = "select xx.ttype as Code from (select 'All' as ttype union all select 'Manual' as ttype union all select Distinct Entry_Source as ttype from TSPL_VLC_DATA_UPLOADER where Entry_Source is not null)xx"
        cmbEntrySource.DataSource = clsDBFuncationality.GetDataTable(qry)
        cmbEntrySource.ValueMember = "Code"
        cmbEntrySource.DisplayMember = "Code"
    End Sub

    Sub LoadReport()
        Dim dt As DataTable = New DataTable
        dt.Columns.Add("Code")
        dt.Columns.Add("Name")

        Dim dr As DataRow = dt.NewRow
        'dr("Code") = "Detail"
        'dr("Name") = "Detail"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow
        'dr("Code") = "Summary"
        'dr("Name") = "Summary"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow
        'dr("Code") = "MP Polling"
        'dr("Name") = "MP Polling"
        'dt.Rows.Add(dr)

        If Not objCommonVar.RCDFCFP Then
            dr = dt.NewRow
            dr("Code") = "Farmer Collection Detail"
            dr("Name") = "Farmer Collection Detail"
            dt.Rows.Add(dr)
        End If

        dr = dt.NewRow
        dr("Code") = "Farmer Wise Collection at DCS"
        dr("Name") = "Farmer Wise Collection at DCS"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "Farmer Collection Entry Status"
        dr("Name") = "Farmer Collection Entry Status"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "Farmer Wise Milk Collection"
        dr("Name") = "Farmer Wise Milk Collection"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "DCS Farmer Collection Details"
        dr("Name") = "DCS Farmer Collection Details"
        dt.Rows.Add(dr)


        dr = dt.NewRow
        dr("Code") = "DCS Collection v/s Farmer Collection Day & Shift Wise"
        dr("Name") = "DCS Collection v/s Farmer Collection Day & Shift Wise"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "DCS Collection v/s Farmer Collection Summary"
        dr("Name") = "DCS Collection v/s Farmer Collection Summary"
        dt.Rows.Add(dr)

        dr = dt.NewRow
        dr("Code") = "DBT Reco  v/s Farmer Collection"
        dr("Name") = "DBT Reco  v/s Farmer Collection"
        dt.Rows.Add(dr)

        cboReportType.DataSource = dt
        cboReportType.ValueMember = "Code"
        cboReportType.DisplayMember = "Name"
        cboReportType.SelectedValue = "Farmer Wise Collection at DCS"
    End Sub



    Sub FormatGrid()
        Dim summaryItem As New GridViewSummaryItem()
        gv.TableElement.TableHeaderHeight = 25
        gv.MasterTemplate.ShowRowHeaderColumn = True
        For ii As Integer = 0 To gv.Columns.Count - 1
            gv.Columns(ii).ReadOnly = True
            gv.Columns(ii).IsVisible = True
        Next

        gv.Columns("Account NO").IsVisible = False
        gv.Columns("Bank Branch").IsVisible = False
        gv.Columns("Bank Name").IsVisible = False
        gv.Columns("fromDate").IsVisible = False
        gv.Columns("Todate").IsVisible = False
        gv.Columns("companyADD").IsVisible = False
        gv.Columns("CompName").IsVisible = False
        gv.Columns("CompCode").IsVisible = False
        gv.Columns("HShift").IsVisible = False
        gv.Columns("ShiftMor").IsVisible = False
        gv.Columns("ShiftEve").IsVisible = False
        gv.Columns("Unit").IsVisible = False
        gv.Columns("Doc_No").IsVisible = False

        ''richa ERO/14/12/18-000443 18 Dec,2018
        gv.Columns("Buffalo_Fat_avg_per_tol").IsVisible = False
        gv.Columns("Buffalo_SNF_avg_per_tol").IsVisible = False
        gv.Columns("Cow_Fat_avg_per_tol").IsVisible = False
        gv.Columns("Cow_SNF_avg_per_tol").IsVisible = False
        gv.Columns("ChangeColorYesNo").IsVisible = False
        gv.Columns("Buff_fat%_avg").IsVisible = False
        gv.Columns("Buff_SNF%_avg").IsVisible = False
        gv.Columns("Cow_fat%_avg").IsVisible = False
        gv.Columns("Cow_SNF%_avg").IsVisible = False
        gv.Columns("TOLERANCE").IsVisible = False

        gv.Columns("Buffalo_Fat_avg_per_tol").FormatString = "{0:F2}"
        gv.Columns("Buffalo_SNF_avg_per_tol").FormatString = "{0:F2}"
        gv.Columns("Cow_Fat_avg_per_tol").FormatString = "{0:F2}"
        gv.Columns("Cow_SNF_avg_per_tol").FormatString = "{0:F2}"
        gv.Columns("Buff_fat%_avg").FormatString = "{0:F2}"
        gv.Columns("Buff_SNF%_avg").FormatString = "{0:F2}"
        gv.Columns("Cow_fat%_avg").FormatString = "{0:F2}"
        gv.Columns("Cow_SNF%_avg").FormatString = "{0:F2}"

        gv.Columns("Buffalo Milk Qty (KG)").IsVisible = True
        gv.Columns("Buffalo Milk Qty (KG)").Width = 100
        gv.Columns("Buffalo Milk Qty (KG)").HeaderText = "Buffalo Milk Qty"
        gv.Columns("Buffalo Milk Qty (KG)").FormatString = "{0:F3}"

        gv.Columns("Buffalo FAT(%)").IsVisible = True
        gv.Columns("Buffalo FAT(%)").Width = 100
        gv.Columns("Buffalo FAT(%)").HeaderText = "Buffalo FAT(%)"
        gv.Columns("Buffalo FAT(%)").FormatString = "{0:F2}"

        gv.Columns("Buffalo SNF(%)").IsVisible = True
        gv.Columns("Buffalo SNF(%)").Width = 100
        gv.Columns("Buffalo SNF(%)").HeaderText = "Buffalo SNF(%)"
        gv.Columns("Buffalo SNF(%)").FormatString = "{0:F2}"

        gv.Columns("Buffalo FAT (KG)").IsVisible = True
        gv.Columns("Buffalo FAT (KG)").Width = 100
        gv.Columns("Buffalo FAT (KG)").HeaderText = "Total Buffalo FAT"
        gv.Columns("Buffalo FAT (KG)").FormatString = "{0:F3}"

        gv.Columns("Buffalo SNF (KG)").IsVisible = True
        gv.Columns("Buffalo SNF (KG)").Width = 100
        gv.Columns("Buffalo SNF (KG)").HeaderText = "Total Buffalo SNF"
        gv.Columns("Buffalo SNF (KG)").FormatString = "{0:F3}"

        gv.Columns("Buffalo Amount").IsVisible = True
        gv.Columns("Buffalo Amount").Width = 100
        gv.Columns("Buffalo Amount").HeaderText = "Buffalo Amount"
        gv.Columns("Buffalo Amount").FormatString = "{0:F3}"

        gv.Columns("Cow Milk Qty (KG)").IsVisible = True
        gv.Columns("Cow Milk Qty (KG)").Width = 100
        gv.Columns("Cow Milk Qty (KG)").HeaderText = "Cow Milk Qty"
        gv.Columns("Cow Milk Qty (KG)").FormatString = "{0:F3}"

        gv.Columns("Cow FAT(%)").IsVisible = True
        gv.Columns("Cow FAT(%)").Width = 100
        gv.Columns("Cow FAT(%)").HeaderText = "Cow FAT(%)"
        gv.Columns("Cow FAT(%)").FormatString = "{0:F2}"

        gv.Columns("Cow SNF(%)").IsVisible = True
        gv.Columns("Cow SNF(%)").Width = 100
        gv.Columns("Cow SNF(%)").HeaderText = "Cow SNF(%)"
        gv.Columns("Cow SNF(%)").FormatString = "{0:F2}"

        gv.Columns("Cow FAT (KG)").IsVisible = True
        gv.Columns("Cow FAT (KG)").Width = 100
        gv.Columns("Cow FAT (KG)").HeaderText = "Total Cow FAT"
        gv.Columns("Cow FAT (KG)").FormatString = "{0:F3}"

        gv.Columns("Cow SNF (KG)").IsVisible = True
        gv.Columns("Cow SNF (KG)").Width = 100
        gv.Columns("Cow SNF (KG)").HeaderText = "Total Cow SNF"
        gv.Columns("Cow SNF (KG)").FormatString = "{0:F3}"

        gv.Columns("Cow Amount").IsVisible = True
        gv.Columns("Cow Amount").Width = 100
        gv.Columns("Cow Amount").HeaderText = "Cow Amount"
        gv.Columns("Cow Amount").FormatString = "{0:F3}"

        gv.Columns("Total").IsVisible = True
        gv.Columns("Total").Width = 100
        gv.Columns("Total").HeaderText = "Total"
        gv.Columns("Total").FormatString = "{0:F3}"

        Dim summaryRowItem As New GridViewSummaryRowItem()

        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Detail") = CompairStringResult.Equal Then
            gv.Columns("Total Milk Qty (KG)").IsVisible = True
            gv.Columns("Total Milk Qty (KG)").Width = 100
            gv.Columns("Total Milk Qty (KG)").HeaderText = "Total Milk Qty (KG)"
            gv.Columns("Total Milk Qty (KG)").FormatString = "{0:F3}"

            gv.Columns("Total FAT(%)").IsVisible = True
            gv.Columns("Total FAT(%)").Width = 100
            gv.Columns("Total FAT(%)").HeaderText = "Total FAT(%)"
            gv.Columns("Total FAT(%)").FormatString = "{0:F2}"

            gv.Columns("Total SNF(%)").IsVisible = True
            gv.Columns("Total SNF(%)").Width = 100
            gv.Columns("Total SNF(%)").HeaderText = "Total SNF(%)"
            gv.Columns("Total SNF(%)").FormatString = "{0:F2}"

            gv.Columns("Total FAT (KG)").IsVisible = True
            gv.Columns("Total FAT (KG)").Width = 100
            gv.Columns("Total FAT (KG)").HeaderText = "Total FAT (KG)"
            gv.Columns("Total FAT (KG)").FormatString = "{0:F3}"

            gv.Columns("Total SNF (KG)").IsVisible = True
            gv.Columns("Total SNF (KG)").Width = 100
            gv.Columns("Total SNF (KG)").HeaderText = "Total SNF (KG)"
            gv.Columns("Total SNF (KG)").FormatString = "{0:F3}"

            gv.Columns("Total Amount").IsVisible = True
            gv.Columns("Total Amount").Width = 100
            gv.Columns("Total Amount").HeaderText = "Total Amount"
            gv.Columns("Total Amount").FormatString = "{0:F3}"

            Dim item111 As New GridViewSummaryItem("Total Milk Qty (KG)", "{0:F3}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item111)

            Dim item112 As New GridViewSummaryItem()
            item112.FormatString = "{0:F2}"
            item112.Name = "Total FAT(%)"
            item112.AggregateExpression = "sum([Total FAT (KG)])*100/sum([Total Milk Qty (KG)])"
            summaryRowItem.Add(item112)

            Dim Item113 As New GridViewSummaryItem()
            Item113.FormatString = "{0:F2}"
            Item113.Name = "Total SNF(%)"
            Item113.AggregateExpression = "sum([Total SNF (KG)])*100/sum([Total Milk Qty (KG)])"
            summaryRowItem.Add(Item113)

            Dim item114 As New GridViewSummaryItem("Total FAT (KG)", "{0:F3}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item114)

            Dim item115 As New GridViewSummaryItem("Total SNF (KG)", "{0:F3}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item115)

            Dim item116 As New GridViewSummaryItem("Total Amount", "{0:F3}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item116)

        End If



        'Dim intCount As Integer = 0

        Dim item6 As New GridViewSummaryItem("Cow Milk Qty (KG)", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item6)
        Dim item7 As New GridViewSummaryItem("Cow FAT (KG)", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item7)

        Dim item8 As New GridViewSummaryItem("Cow SNF (KG)", "{0:F3}", GridAggregateFunction.Sum)

        summaryRowItem.Add(item8)
        Dim summaryItem3 As New GridViewSummaryItem()
        summaryItem3.FormatString = "{0:F2}"
        summaryItem3.Name = "Cow SNF(%)"
        summaryItem3.AggregateExpression = "sum([Cow SNF (KG)])*100/sum([Cow Milk Qty (KG)])"
        summaryRowItem.Add(summaryItem3)

        Dim summaryItem4 As New GridViewSummaryItem()
        summaryItem4.FormatString = "{0:F2}"
        summaryItem4.Name = "Cow FAT(%)"
        summaryItem4.AggregateExpression = "sum([Cow FAT (KG)])*100/sum([Cow Milk Qty (KG)])"
        summaryRowItem.Add(summaryItem4)

        Dim item9 As New GridViewSummaryItem("Buffalo Milk Qty (KG)", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item9)
        Dim item10 As New GridViewSummaryItem("Buffalo FAT (KG)", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item10)
        Dim item11 As New GridViewSummaryItem("Buffalo SNF (KG)", "{0:F3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item11)
        Dim summaryItem5 As New GridViewSummaryItem()
        summaryItem5.FormatString = "{0:F2}"
        summaryItem5.Name = "Buffalo FAT(%)"
        summaryItem5.AggregateExpression = "sum([Buffalo FAT (KG)])*100/sum([Buffalo Milk Qty (KG)])"
        summaryRowItem.Add(summaryItem5)

        Dim summaryItem6 As New GridViewSummaryItem()
        summaryItem6.FormatString = "{0:F2}"
        summaryItem6.Name = "Buffalo SNF(%)"
        summaryItem6.AggregateExpression = "sum([Buffalo SNF (KG)])*100/sum([Buffalo Milk Qty (KG)])"
        summaryRowItem.Add(summaryItem6)

        Dim item12 As New GridViewSummaryItem("Cow Amount", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item12)
        Dim item13 As New GridViewSummaryItem("Buffalo Amount", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item13)

        Dim item14 As New GridViewSummaryItem("Total", "{0:F2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(item14)

        gv.ShowGroupPanel = False
        gv.MasterTemplate.AutoExpandGroups = True
        gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub
    Sub Reset()
        gv.DataSource = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
        gv.DataSource = Nothing
        cboUnit.Text = "Kg"
        txtReportType.Text = ""
        Dim viewBlank As New TableViewDefinition()
        gv.ViewDefinition = viewBlank
        EnableDisableControl(True)
        CheckReportTypeforDateRange()
        EnableDisableCntrl(True)
        TxtZone.arrValueMember = Nothing

        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DBT Reco  v/s Farmer Collection") = CompairStringResult.Equal Then
            txtFromDate.Value = clsCommon.GETSERVERDATE()
            txtToDate.Value = clsCommon.GETSERVERDATE()


        End If

    End Sub

    Private Sub EnableDisableControl(ByVal val As Boolean)
        txtFromDate.Enabled = val
        txtFromShift.Enabled = val
        txtToDate.Enabled = val
        txtToShift.Enabled = val
        cboReportType.Enabled = val
        cmbEntrySource.Enabled = val
        cboUnit.Enabled = val
        txtMCC.Enabled = val
        txtVLC.Enabled = val
        txtReportType.Enabled = val

        txtMinDays.Enabled = val
    End Sub

    Private Sub DateWise()
        Try

            Dim qry As String
            qry = GetBaseQuery()

            If chkSupervisor.Checked = True Then
                Dim strSource As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select STUFF((Select ', ['+Code+']' from (select xx.ttype as Code from (select 'Manual' as ttype union all select Distinct Entry_Source as ttype from TSPL_VLC_DATA_UPLOADER where Entry_Source is not null)xx) XXX For XML Path('')),1,1,'') "))

                Dim SourceName As String = ""
                If clsCommon.myLen(strSource) > 0 Then
                    SourceName = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select STUFF((Select ', sum(isnull (['+Code+'],0)) as ' +'['+ Code+']' from (select xx.ttype as Code from (select 'Manual' as ttype union all select Distinct Entry_Source as ttype from TSPL_VLC_DATA_UPLOADER where Entry_Source is not null)xx) XXX For XML Path('')),1,1,'')   "))
                    SourceName = " ," + SourceName
                End If

                qry += "  select [VLC Uploader],[VLC Code],max([VLC Name]) as [VLC Name],[MCC Code],max([MCC Name]) as [MCC Name]
                         ,[Route Code],max([Route Name]) as [Route Name],max([Supervisor Name]) as [Supervisor Name]
                         ,convert(nvarchar, [Date],103) as [Date]
                         " + SourceName + " from 
                         (select CTEMP.[VLC Uploader],CTEMP.[VLC Code],max(CTEMP.[VLC Name]) as [VLC Name],CTEMP.[MCC Code],max(CTEMP.[MCC Name]) as [MCC Name]
                         ,CTEMP.[Route Code],max(CTEMP.[Route Name]) as [Route Name],max(isnull(tspl_employee_master.Emp_Name,'')) as [Supervisor Name]
                         ,convert(date, Doc_Date,103) as [Date],tttype,count(1) as [cc] 
                         from CTEMP
                         left join TSPL_VLC_Supervisor_Tagging on TSPL_VLC_Supervisor_Tagging.vlc_code=CTEMP.[VLC Code]
                         and  TSPL_VLC_Supervisor_Tagging.mcc_code=CTEMP.[MCC Code]
                         left join tspl_employee_master on tspl_employee_master.emp_code=TSPL_VLC_Supervisor_Tagging.Supervisor_code
                          WHERE 1=1 
                          group by tttype,
                         convert(date, Doc_Date,103) ,CTEMP.[VLC Uploader] ,CTEMP.[VLC Code],CTEMP.[MCC Code],TSPL_VLC_Supervisor_Tagging.Supervisor_code,CTEMP.[Route Code] 
                          )tt pivot (  sum(cc) for tttype in (" + strSource + ") ) as zpivot
                          group by [Date],[VLC Uploader],[VLC Code],[MCC Code],[Route Code]
                          order by [Date] ,[VLC Uploader] "
            Else
                qry += " select doc_date as [Date],count(1) as [Count] from CTEMP WHERE 1=1 group by doc_date order by Doc_Date  "
            End If

            dt = clsDBFuncationality.GetDataTable(qry)

            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gv.DataSource = dt
            gv.GroupDescriptors.Clear()
            gv.MasterTemplate.SummaryRowsBottom.Clear()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            txtReportType.Text = "DateWise"

            gv.TableElement.TableHeaderHeight = 25
            gv.MasterTemplate.ShowRowHeaderColumn = True
            For ii As Integer = 0 To gv.Columns.Count - 1
                gv.Columns(ii).ReadOnly = True
                gv.Columns(ii).IsVisible = True
            Next

            Dim summaryRowItem As New GridViewSummaryRowItem()
            If chkSupervisor.Checked = True Then
                Dim dtDeduction As DataTable = clsDBFuncationality.GetDataTable("select xx.ttype as Code from (select 'Manual' as ttype union all select Distinct Entry_Source as ttype from TSPL_VLC_DATA_UPLOADER where Entry_Source is not null)xx")

                If dtDeduction IsNot Nothing AndAlso dtDeduction.Rows.Count > 0 Then
                    For i As Integer = 0 To dtDeduction.Rows.Count - 1
                        Dim aa = clsCommon.myCstr(dtDeduction.Rows(i).Item(0))
                        Dim item111 As New GridViewSummaryItem(aa, "{0:F0}", GridAggregateFunction.Sum)
                        summaryRowItem.Add(item111)
                        gv.Columns(i).FormatString = "{0:F0}"
                    Next
                End If
            Else
                Dim item14 As New GridViewSummaryItem("Count", "{0:F0}", GridAggregateFunction.Sum)
                summaryRowItem.Add(item14)
            End If

            gv.ShowGroupPanel = False
            gv.MasterTemplate.AutoExpandGroups = True
            gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            gv.BestFitColumns()
            RadPageView1.SelectedPage = RadPageViewPage2
            ReStoreGridLayout()
            EnableDisableControl(False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub LoadData()
        Try
            If objCommonVar.RCDFCFP Then
                dt = clsMilkUnion.UnionDBName()
            Else
                Dim arrUnion As New ArrayList()
                arrUnion.Add(objCommonVar.CurrDatabase)
                dt = clsMilkUnion.UnionDBName1(arrUnion)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    dt = clsMilkUnion.GetUnionDBandLocName(objCommonVar.CurrDatabase)
                End If
            End If
            If clsCommon.GetDateWithStartTime(txtFromDate.Value) > clsCommon.GetDateWithStartTime(txtToDate.Value) Then
                txtFromDate.Focus()
                clsCommon.MyMessageBoxShow(Me, "From date can not be greater then to Date", Me.Text)
                Exit Sub
            End If
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal Then
                DateWise()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "MP Polling") = CompairStringResult.Equal Then
                MPPollingFarmerCount()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Detail") = CompairStringResult.Equal Then
                Detail()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Detail") = CompairStringResult.Equal Then
                FarmerCollectionDetails()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Collection at DCS") = CompairStringResult.Equal Then
                FarmerWiseCollectionAtDCS()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Entry Status") = CompairStringResult.Equal Then
                FarmerCollectionEntryStatus()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Milk Collection") = CompairStringResult.Equal Then
                FarmerWiseMilkCollection()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Farmer Collection Details") = CompairStringResult.Equal Then
                DCSFarmerCollectionDetails()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Day & Shift Wise") = CompairStringResult.Equal Then
                DCSFarmerCollectionDayShiftWiseNew()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
                DCSFarmerCollectionDayShiftWiseNew()
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DBT Reco  v/s Farmer Collection") = CompairStringResult.Equal Then
                DbtReco()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub DbtReco()
        Try
            If txtFromDate.Value.Day <> 1 Then
                 clsCommon.MyMessageBoxShow(Me, "Please select the 1st day of the month.", Me.Text)
                txtFromDate.Focus()
                Exit Sub
            End If
          Dim dtS As DateTime = txtToDate.Value

             If dtS.Day <> DateTime.DaysInMonth(dtS.Year, dtS.Month) Then
               ClsCommon.MyMessageBoxShow(Me, "Please select the last day of the month.", Me.Text)
                txtToDate.Focus()
                Exit Sub
            End If
            Dim dt As DataTable = Nothing
            Dim WHR As String = Nothing
            Dim Baseqry As String = Nothing
            Dim Qry As String = Nothing
            Qry = "SELECT max(CONVERT(varchar, XXX.DOC_DATE,103))DOC_DATE,max(XXX.SHIFT)SHIFT, XXX.VLC_CODE,MAX(XXX.MCC_CODE)MCC_CODE,MAX(XXX.DOC_CODE)DOC_CODE,max(XXX.DCS_Code) AS[DCS_Code],max(XXX.DCS_Uploader_code) as [DCS_Uploader_code],max(XXX.DCS_NAME) as [DCS_NAME],max(XXX.ROUTE_CODE) as [ROUTE_CODE],max(xxx.Zone_Code)Zone_Code,max(xxx.Description)[Zone_Name],
cast(sum(XXX.RECOQTY) as DECIMAL(18,2)) as [RECOQTY],sum(XXX.RECOFAT_KG)RECOFAT_KG,sum(XXX.RECOSNF_KG)RECOSNF_KG,
sum(CAST((ISNULL(XXX.RECOFAT_KG,0) * 100 /NULLIF(XXX.RECOQTY,0)) AS DECIMAL(18,2))) AS RECOFatAVG,
sum( CAST((ISNULL(XXX.RECOSNF_KG,0) * 100 / NULLIF(XXX.RECOQTY,0)) AS DECIMAL(18,2))) AS RECOSNFAVG
,  COUNT(DISTINCT NULLIF(LTRIM(RTRIM(xxx.MP_CODE)), '')) AS Farmer_Count,cast(SUM(XXX.FARMERQTY) as DECIMAL(18,2)) AS FARMERQTY,cast(SUM(XXX.FARMERFAT_KG) as decimal(18,2)) AS FARMERFAT_KG	,cast(SUM(snf_KG) as decimal (18,2))AS FARMERSNF_KG	,
cast(SUM(XXX.FARMERFatPer) as decimal(18,2)) AS FARMERFatPer,	cast(SUM(XXX.FARMERSNFPer) as decimal(18,2)) AS  FARMERSNFPer	
,ISNULL(CAST(SUM(XXX.FARMERFAT_KG) * 100.0 / NULLIF(SUM(FARMERQTY), 0) AS DECIMAL(10,2)), 0.00) AS FARMERFATAVG,
ISNULL( CAST(SUM(XXX.FARMERSNF_KG) * 100.0 / NULLIF(SUM(FARMERQTY), 0) AS DECIMAL(10,2))
, 0.00) AS FARMERSNFAVG 
--,(sum(XXX.RECOQTY)-SUM(XXX.FARMERQTY) ) AS Difference1
,CAST((SUM(XXX.RECOQTY) - SUM(XXX.FARMERQTY)) AS DECIMAL(18,2)) AS Difference,
CASE 
    WHEN (SUM(XXX.RECOQTY) - SUM(XXX.FARMERQTY)) < 0.09
    THEN 'MATCH' 
    ELSE 'NOT MATCH' 
END AS STATUS

FROM  "
            Qry += " (
           				 Select '' as MP_CODE, TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description, fORMAT(CONVERT(date, TSPL_DCS_MP_INCENTIVE_RECO_head.Document_Date, 103), 'MMMM yyyy') AS Month , TSPL_VLC_MASTER_HEAD.VLC_CODE, (TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.MCC_CODE)MCC_CODE, (TSPL_DCS_MP_INCENTIVE_RECO_head.document_code)DOC_CODE,CONVERT(varchar, TSPL_DCS_MP_INCENTIVE_RECO_head.Document_Date,103)DOC_DATE,'' AS SHIFT,(TSPL_VLC_MASTER_HEAD.VSP_CODE) AS[DCS_Code],(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) as [DCS_Uploader_code],(TSPL_VLC_MASTER_HEAD.VLC_Name) as [DCS_NAME],(TSPL_BULK_ROUTE_MASTER.ROUTE_NO) as [ROUTE_CODE],
           				 (TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.qty) as [RECOQTY],(TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.fat)RECOFAT_KG,(TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.snf)RECOSNF_KG,
           				 (TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code) AS Farmer_Count
           				 ,0 AS FARMERQTY,0 AS FARMERFAT_KG,0 AS FARMERSNF_KG,0 AS FARMERFatPer,0 AS FARMERSNFPer,0 AS FAT_KG,0 AS SNF_KG
           				 from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL
           				  left outer join TSPL_DCS_MP_INCENTIVE_RECO_head on TSPL_DCS_MP_INCENTIVE_RECO_head.Document_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code
           				 left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code
				         LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_MASTER  ON TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_CODE
				         LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_DETAIL ON TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code=TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code
				         left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code 
	   left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
                          WHERE 2=2 and     "
            Qry += "   convert(date,TSPL_DCS_MP_INCENTIVE_RECO_head.Reco_Date,103)>=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) and
convert(date,TSPL_DCS_MP_INCENTIVE_RECO_head.Reco_Date,103) <=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
            If clsCommon.CompairString(txtFromShift.Text, "e") = CompairStringResult.Equal Then
                Qry += " and 2=( case when tspl_dcs_mp_incentive_reco_head.Reco_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and tspl_dcs_mp_incentive_reco_head.Reco_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and shift='m' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "m") = CompairStringResult.Equal Then
                Qry += " and 2=( case when tspl_dcs_mp_incentive_reco_head.Reco_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and tspl_dcs_mp_incentive_reco_head.Reco_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and shift='e' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            Qry += "UNION ALL
SELECT '' as MP_CODE,TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description, fORMAT(CONVERT(date, TSPL_DCS_MP_INCENTIVE_RECO_head.Document_Date, 103), 'MMMM yyyy') AS Month , TSPL_VLC_MASTER_HEAD.VLC_CODE, (TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.MCC_CODE)MCC_CODE, 
(TSPL_DCS_MP_INCENTIVE_RECO_head.document_code)DOC_CODE,CONVERT(varchar, TSPL_DCS_MP_INCENTIVE_RECO_head.Document_Date,103)DOC_DATE,'' AS SHIFT,(TSPL_VLC_MASTER_HEAD.VSP_CODE) AS[DCS_Code],
(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) as [DCS_Uploader_code],(TSPL_VLC_MASTER_HEAD.VLC_Name) as [DCS_NAME],(TSPL_BULK_ROUTE_MASTER.ROUTE_NO) as [ROUTE_CODE],
        				 (TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.qty) as [RECOQTY],(TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.fat)RECOFAT_KG,(TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.snf)RECOSNF_KG,
           				 (TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code) AS Farmer_Count
           				 ,0 AS FARMERQTY,0 AS FARMERFAT_KG,0 AS FARMERSNF_KG,0 AS FARMERFatPer,0 AS FARMERSNFPer,0 AS FAT_KG,0 AS SNF_KG
						 
						 FROM TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID 
						 LEFT OUTER JOIN TSPL_DCS_MP_INCENTIVE_RECO_head ON TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.Document_Code=TSPL_DCS_MP_INCENTIVE_RECO_head.Document_Code
						 left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.VLC_Code
						 left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code 
LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_MASTER  ON TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID.VLC_CODE
				      LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_DETAIL ON TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code=TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code
	   left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
WHERE 2=2 and    "

            Qry += "   convert(date,TSPL_DCS_MP_INCENTIVE_RECO_head.Reco_Date,103)>=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) and
convert(date,TSPL_DCS_MP_INCENTIVE_RECO_head.Reco_Date,103) <=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
            If clsCommon.CompairString(txtFromShift.Text, "e") = CompairStringResult.Equal Then
                Qry += " and 2=( case when tspl_dcs_mp_incentive_reco_head.Reco_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and tspl_dcs_mp_incentive_reco_head.Reco_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and shift='m' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "m") = CompairStringResult.Equal Then
                Qry += " and 2=( case when tspl_dcs_mp_incentive_reco_head.Reco_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and tspl_dcs_mp_incentive_reco_head.Reco_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and shift='e' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            Qry += "  UNION ALL
                      select '' as MP_CODE, TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description, fORMAT(CONVERT(date, TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date, 103), 'MMMM yyyy') AS Month ,  TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code AS VLC_CODE, (TSPL_VLC_MASTER_HEAD.MCC) AS MCC_CODE,(TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code) AS DOC_CODE, convert(varchar,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) AS DOC_DATE,TSPL_VLC_DATA_UPLOADER_MASTER.shift AS SHIFT , (VSP_CODE) AS DCS_CODE,
                     ( TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) AS DCS_Uploader_code, (TSPL_VLC_MASTER_HEAD.VLC_Name) AS DCS_NAME,(TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code) AS ROUTE_CODE,
                      0 AS RECOQTY,0 AS RECOFAT_KG,0 AS RECOSNF_KG   ,(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code)Farmer_Count,
                      (TSPL_VLC_DATA_UPLOADER_DETAIL.QTY) AS FARMERQTY, (TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FARMERFAT_KG,
                      (TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FARMERSNF_KG
                     , (TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer )FARMERFatPer, (TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer)FARMERSNFPer,0 as  FAT_KG,0 as SNF_KG
                      from TSPL_VLC_DATA_UPLOADER_DETAIL
                      left outer join TSPL_VLC_DATA_UPLOADER_MASTER on TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
                      LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code 
	   left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
                      where 2=2 and    "
            Qry += "   convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)>=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',103) and convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) <=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            Qry += " UNION ALL
                 select TSPL_VLC_DATA_UPLOADER.MP_CODE, TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description,fORMAT(CONVERT(date, TSPL_VLC_DATA_UPLOADER.Doc_Date, 103), 'MMMM yyyy') AS Month  ,TSPL_VLC_MASTER_HEAD.VLC_CODE	,(TSPL_VLC_MASTER_HEAD.MCC)MCC,	(TSPL_VLC_DATA_UPLOADER.DOC_NO) AS DOC_CODE	,CONVERT(VARCHAR,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)DOC_DATE,	SHIFT,	(VSP_Code)DCS_Code,	(VLC_Code_VLC_Uploader)DCS_Uploader_code,	(VLC_Name)DCS_NAME	,(ROUTE_CODE)ROUTE_CODE	, 0 AS RECOQTY	,0 AS RECOFAT_KG	,0 AS RECOSNF_KG	,	
                       (TSPl_MP_MAster.MP_Code)Farmer_Count	,(TSPL_VLC_DATA_UPLOADER.QTY) AS FARMERQTY	 ,(fat_KG) AS FARMERFAT_KG	,(snf_KG) AS FARMERSNF_KG	,(fat) AS FARMERFatPer,	(snf) AS  FARMERSNFPer	,
                      (TSPL_VLC_DATA_UPLOADER.fat_KG) FAT_KG ,(TSPL_VLC_DATA_UPLOADER.snf_KG) SNF_KG
                      from TSPL_VLC_DATA_UPLOADER
                      Left Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE
	   left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
                      Left Join TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
                  where 2=2 and  "
            Qry += "   convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)>=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "',104) and convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103) <=CONVERT(DATE,'" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "',103) "
            If clsCommon.CompairString(txtFromShift.Text, "e") = CompairStringResult.Equal Then
                Qry += " and 2=( case when tspl_vlc_data_uploader.doc_date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER.doc_date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and shift='m' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "m") = CompairStringResult.Equal Then
                Qry += " and 2=( case when tspl_vlc_data_uploader.doc_date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER.doc_date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/mmm/yyyy hh:mm tt") + "' and shift='e' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            Qry += ")  XXX GROUP BY  XXX.VLC_CODE  "
            dt = clsDBFuncationality.GetDataTable(Qry)
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Baseqry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv.DataSource = Nothing
                gv.Rows.Clear()
                gv.Columns.Clear()
                gv.GroupDescriptors.Clear()
                gv.MasterTemplate.SummaryRowsBottom.Clear()
                gv.MasterView.Refresh()
                gv.DataSource = dt
                For ii As Integer = 0 To gv.Columns.Count - 1
                    gv.Columns(ii).ReadOnly = True
                Next
                RadPageView1.SelectedPage = RadPageViewPage2
                gv.EnableFiltering = True
                SetGridFormationDBT()
                gv.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found To Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub DCSFarmerCollectionDayShiftWiseNew()
        Try
            Dim dt As DataTable = Nothing
            Dim WHR As String = Nothing
            Dim Baseqry As String = Nothing
            Dim Qry As String = Nothing
            Qry = "SELECT XXX.VLC_CODE,MAX(XXX.MCC_CODE)MCC_CODE,

  MAX(XXX.DOC_CODE)DOC_CODE,CONVERT(varchar, XXX.DOC_DATE,103)DOC_DATE,XXX.SHIFT,max(XXX.DCS_Code) AS[DCS_Code],max(XXX.DCS_Uploader_code) as [DCS_Uploader_code],max(XXX.DCS_NAME) as [DCS_NAME],max(XXX.ROUTE_CODE) as [ROUTE_CODE],max(xxx.Zone_Code)Zone_Code,max(xxx.Description)[Zone_Name],CONVERT(decimal(18,2),sum(XXX.SRNQTY)) as [SRNQTY],CONVERT(decimal(18,2),sum(xxx.ACC_Qty_LTR))[SRNLTR],CONVERT(decimal(18,2),sum(XXX.SRNFAT_KG))SRNFAT_KG,CONVERT(decimal(18,2),sum(XXX.SRNSNF_KG))SRNSNF_KG,
sum(CAST((ISNULL(XXX.SRNFAT_KG,0) * 100 /NULLIF(XXX.SRNQTY,0)) AS DECIMAL(18,2))) AS SRNFatAVG,
sum( CAST((ISNULL(XXX.SRNSNF_KG,0) * 100 / NULLIF(XXX.SRNQTY,0)) AS DECIMAL(18,2))) AS SRNSNFAVG
,COUNT(XXX.Farmer_Count) AS Farmer_Count

,CONVERT(decimal(18,2),SUM(XXX.FARMERQTY)) AS FARMERQTY	
  ,CONVERT(decimal(18,2),SUM(XXX.FARMERFAT_KG)) AS FARMERFAT_KG	,CONVERT(decimal(18,2),SUM(snf_KG)) AS FARMERSNF_KG	,SUM(XXX.FARMERFatPer) AS FARMERFatPer,	SUM(XXX.FARMERSNFPer) AS  FARMERSNFPer	
,ISNULL(
    CAST(SUM(XXX.FARMERFAT_KG) * 100.0 / NULLIF(SUM(FARMERQTY), 0) AS DECIMAL(10,2))
, 0.00) AS FARMERFATAVG,

ISNULL(
    CAST(SUM(XXX.FARMERSNF_KG) * 100.0 / NULLIF(SUM(FARMERQTY), 0) AS DECIMAL(10,2))
, 0.00) AS FARMERSNFAVG
,CAST((SUM(XXX.SRNQTY) - SUM(XXX.FARMERQTY)) AS DECIMAL(18,2)) AS Difference,
CASE 
    WHEN (SUM(XXX.SRNQTY) - SUM(XXX.FARMERQTY)) = 0 
    THEN 'MATCH' 
    ELSE 'NOT MATCH' 
END AS STATUS

FROM  "
            Qry += " (

Select TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description,  TSPL_MILK_SRN_DETAIL.ACC_Qty_LTR, TSPL_MILK_SRN_head.VLC_CODE, (TSPL_MILK_SRN_head.MCC_CODE)MCC_CODE, (TSPL_MILK_SRN_head.DOC_CODE)DOC_CODE,CONVERT(varchar, TSPL_MILK_SRN_head.DOC_DATE,103)DOC_DATE,TSPL_MILK_SRN_head.SHIFT,(TSPL_MILK_SRN_head.VSP_CODE) AS[DCS_Code],(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) as [DCS_Uploader_code],(TSPL_VLC_MASTER_HEAD.VLC_Name) as [DCS_NAME],(TSPL_MILK_SRN_head.ROUTE_CODE) as [ROUTE_CODE],(TSPL_MILK_SRN_DETAIL.qty) as [SRNQTY],(TSPL_MILK_SRN_DETAIL.fat_kg)SRNFAT_KG,(TSPL_MILK_SRN_DETAIL.SNF_KG)SRNSNF_KG,
(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code) AS Farmer_Count
,0 AS FARMERQTY,0 AS FARMERFAT_KG,0 AS FARMERSNF_KG,0 AS FARMERFatPer,0 AS FARMERSNFPer

,TSPL_MILK_SRN_DETAIL.FAT_KG,TSPL_MILK_SRN_DETAIL.SNF_KG
                  from TSPL_MILK_SRN_DETAIL
                   LEFT OUTER JOIN TSPL_MILK_SRN_head ON TSPL_MILK_SRN_head.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE
				   LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_MASTER  ON TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_MILK_SRN_head.VLC_CODE
				   LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_DETAIL ON TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code=TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code
                  LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_MILK_SRN_head.VSP_CODE 
	  left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
                  WHERE 2=2 and     "
            Qry += "   convert(date,TSPL_MILK_SRN_head.DOC_DATE,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' and convert(date,TSPL_MILK_SRN_head.DOC_DATE,103) <='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            ' convert(date,TSPL_MILK_SRN_head.DOC_DATE,103)>='01/Feb/2026' and convert(date,TSPL_MILK_SRN_head.DOC_DATE,103) <='01/Feb/2026' 
            'AND VLC_Code_VLC_Uploader='5102'

            Qry += "  UNION ALL
 select TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description, 0 as ACC_Qty_LTR , TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code AS VLC_CODE, (TSPL_VLC_MASTER_HEAD.MCC) AS MCC_CODE,(TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code) AS DOC_CODE, convert(varchar,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) AS DOC_DATE,TSPL_VLC_DATA_UPLOADER_MASTER.shift AS SHIFT , (VSP_CODE) AS DCS_CODE,
( TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) AS DCS_Uploader_code,
 (TSPL_VLC_MASTER_HEAD.VLC_Name) AS DCS_NAME,(TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code) AS ROUTE_CODE,
 0 AS SRNQTY,0 AS SRNFAT_KG,0 AS SRNSNF_KG
 ,(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code)Farmer_Count,
 (TSPL_VLC_DATA_UPLOADER_DETAIL.QTY) AS FARMERQTY,
(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FARMERFAT_KG,
 (TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FARMERSNF_KG
, (TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer )FARMERFatPer,
(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer)FARMERSNFPer,
0 as  FAT_KG,0 as SNF_KG
 from TSPL_VLC_DATA_UPLOADER_DETAIL
 left outer join TSPL_VLC_DATA_UPLOADER_MASTER on TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
 LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code 
	  left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
 where 2=2 and    "
            Qry += "   convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' and convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) <='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            'convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)>='01/Feb/2026' and convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) <='01/Feb/2026'
            'AND VLC_Code_VLC_Uploader='5102'

            Qry += "   UNION ALL
   select TSPL_ZONE_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description, 0 as ACC_Qty_LTR ,  TSPL_VLC_MASTER_HEAD.VLC_CODE	,(TSPL_VLC_MASTER_HEAD.MCC)MCC,	(TSPL_VLC_DATA_UPLOADER.DOC_NO) AS DOC_CODE	,CONVERT(VARCHAR,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)DOC_DATE,	SHIFT,	(VSP_Code)DCS_Code,	(VLC_Code_VLC_Uploader)DCS_Uploader_code,	(VLC_Name)DCS_NAME	,(ROUTE_CODE)ROUTE_CODE	, 0 AS SRNQTY	,0 AS SRNFAT_KG	,0 AS SRNSNF_KG	,	
      (TSPl_MP_MAster.MP_Code)Farmer_Count	,(TSPL_VLC_DATA_UPLOADER.QTY) AS FARMERQTY	
   ,(fat_KG) AS FARMERFAT_KG	,(snf_KG) AS FARMERSNF_KG	,(fat) AS FARMERFatPer,	(snf) AS  FARMERSNFPer	,
   (TSPL_VLC_DATA_UPLOADER.fat_KG) FAT_KG
,(TSPL_VLC_DATA_UPLOADER.snf_KG) SNF_KG
 from TSPL_VLC_DATA_UPLOADER
 Left Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE
Left Join TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
	  left  join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
LEFT  JOIN TSPL_ZONE_MASTER ON TSPL_ZONE_MASTER.Zone_Code =TSPL_VENDOR_MASTER.Zone_Code
where 2=2 and  "
            Qry += "   convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' and convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103) <='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_VLC_DATA_UPLOADER.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Qry += " and 2=( case when TSPL_VLC_DATA_UPLOADER.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Qry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                Qry += " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
            End If
            '         convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)>='01/Feb/2026' and convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103) <='01/Feb/2026' 
            'AND VLC_Code_VLC_Uploader='5102'



            Qry += ")  XXX GROUP BY XXX.DOC_DATE,XXX.shift, XXX.VLC_CODE  "



            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
                Baseqry = " SELECT 
 XXXX.VLC_CODE,(convert(varchar,DOC_DATE,103))DOC_DATE,MAX(XXXX.DCS_Code)DCS_Code,CAST(MAX([DCS_Uploader_code]) AS INT) AS DCS_Uploader_code,MAX([DCS_NAME])[DCS_NAME],max(xxxx.Zone_Code)Zone_Code,max(xxxx.Zone_Name)[Zone_Name],SUM(XXXX.[SRNQTY])[SRNQTY],sum([SRNLTR])[SRNLTR],SUM(XXXX.SRNFAT_KG)SRNFAT_KG,SUM(XXXX.SRNSNF_KG)SRNSNF_KG,MAX(XXXX.SRNFatAVG)SRNFatAVG,MAX(XXXX.SRNSNFAVG)SRNSNFAVG,
    
SUM(Farmer_Count) AS Farmer_Count,SUM(XXXX.FARMERQTY)FARMERQTY,SUM(XXXX.FARMERFAT_KG)FARMERFAT_KG,SUM(XXXX.FARMERSNF_KG)FARMERSNF_KG,SUM(XXXX.FARMERFatPer)FARMERFatPer,SUM(XXXX.FARMERSNFPer)FARMERSNFPer,sum( XXXX.FARMERFatAVG)FARMERFatAVG,SUM( XXXX.FARMERSNFAVG )FARMERSNFAVG from(" + Qry + ")XXXX GROUP BY VLC_CODE,DOC_DATE"
            End If



            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
                dt = clsDBFuncationality.GetDataTable(Baseqry)
            Else
                dt = clsDBFuncationality.GetDataTable(Qry)
            End If
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Baseqry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv.DataSource = Nothing
                gv.Rows.Clear()
                gv.Columns.Clear()
                gv.GroupDescriptors.Clear()
                gv.MasterTemplate.SummaryRowsBottom.Clear()
                gv.MasterView.Refresh()
                gv.DataSource = dt
                For ii As Integer = 0 To gv.Columns.Count - 1
                    gv.Columns(ii).ReadOnly = True
                Next
                RadPageView1.SelectedPage = RadPageViewPage2
                gv.EnableFiltering = True
                'SetGridFormat()
                SetGridFormationOFGV1Collection()

                ' View()
                'View1()
                ' EnableDisableCntrl(False)
                gv.BestFitColumns()
                '   EnableDisableCntrl(False)
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found To Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub DCSFarmerCollectionDayShiftWiseSummary()
        Try
            Dim dt As DataTable = Nothing
            Dim WHR As String = Nothing
            Dim Baseqry As String = Nothing
            Dim Qry As String = Nothing
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                WHR = " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                WHR += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                WHR += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            Baseqry = " SELECT *
FROM (Select  TSPL_MILK_SRN_head.VLC_CODE, MAX(TSPL_MILK_SRN_head.MCC_CODE)MCC_CODE, MAX(TSPL_MILK_SRN_head.DOC_CODE)DOC_CODE,CONVERT(varchar, TSPL_MILK_SRN_head.DOC_DATE,103)DOC_DATE,TSPL_MILK_SRN_head.SHIFT,max(TSPL_MILK_SRN_head.VSP_CODE) AS[DCS_Code],max(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) as [DCS_Uploader_code],max(TSPL_VLC_MASTER_HEAD.VLC_Name) as [DCS_NAME],max(TSPL_MILK_SRN_head.ROUTE_CODE) as [ROUTE_CODE],sum(TSPL_MILK_SRN_DETAIL.qty) as [SRNQTY],sum(TSPL_MILK_SRN_DETAIL.fat_kg)SRNFAT_KG,sum(TSPL_MILK_SRN_DETAIL.SNF_KG)SRNSNF_KG,
sum(CAST((ISNULL(TSPL_MILK_SRN_DETAIL.FAT_KG,0) * 100 /NULLIF(TSPL_MILK_SRN_DETAIL.qty,0)) AS DECIMAL(18,2))) AS SRNFatAVG,
sum( CAST((ISNULL(TSPL_MILK_SRN_DETAIL.SNF_KG,0) * 100 / NULLIF(TSPL_MILK_SRN_DETAIL.qty,0)) AS DECIMAL(18,2))) AS SRNSNFAVG
,COUNT(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code) AS Farmer_Count
,0 AS FARMERQTY,0 AS FARMERFAT_KG,0 AS FARMERSNF_KG,0 AS FARMERFatPer,0 AS FARMERSNFPer,0 AS FARMERFatAVG, 0 AS FARMERSNFAVG
                  from TSPL_MILK_SRN_DETAIL
                   LEFT OUTER JOIN TSPL_MILK_SRN_head ON TSPL_MILK_SRN_head.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE
				   LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_MASTER  ON TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_MILK_SRN_head.VLC_CODE
				   LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_DETAIL ON TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code=TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code
                  LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_MILK_SRN_head.VSP_CODE 
                  WHERE 2=2 and  "
            Baseqry += "   convert(date,TSPL_MILK_SRN_head.DOC_DATE,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' and convert(date,TSPL_MILK_SRN_head.DOC_DATE,103) <='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Baseqry += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Baseqry += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If


            Baseqry += " group by TSPL_MILK_SRN_head.DOC_DATE,TSPL_MILK_SRN_head.shift,"
            Baseqry += " TSPL_MILK_SRN_head.VLC_CODE
 UNION ALL
select TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code AS VLC_CODE, MAX(TSPL_VLC_MASTER_HEAD.MCC) AS MCC_CODE,MAX(TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code) AS DOC_CODE, convert(varchar,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) AS DOC_DATE,TSPL_VLC_DATA_UPLOADER_MASTER.shift AS SHIFT , MAX(VSP_CODE) AS DCS_CODE,
MAX( TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) AS DCS_Uploader_code,
 MAX(TSPL_VLC_MASTER_HEAD.VLC_Name) AS DCS_NAME,MAX(TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code) AS ROUTE_CODE,
 0 AS SRNQTY,0 AS SRNFAT_KG,0 AS SRNSNF_KG,0 AS SRNFatAVG,0 AS SRNSNFAVG
 ,COUNT(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code)Farmer_Count,
 SUM(TSPL_VLC_DATA_UPLOADER_DETAIL.QTY) AS FARMERQTY,
sum(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FARMERFAT_KG,
 sum(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FARMERSNF_KG
, SUM(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer )FARMERFatPer,
SUM(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer)FARMERSNFPer, 
CAST(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer * TSPL_VLC_DATA_UPLOADER_DETAIL.QTY / 100, 0)) * 100 / NULLIF(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.QTY, 0)), 0) AS DECIMAL(18,2)) AS FARMERFatAVG,
CAST(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer * TSPL_VLC_DATA_UPLOADER_DETAIL.QTY / 100, 0)) * 100 / NULLIF(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.QTY, 0)), 0)
AS DECIMAL(18,2)) AS FARMERSNFAVG
 from TSPL_VLC_DATA_UPLOADER_DETAIL
 left outer join TSPL_VLC_DATA_UPLOADER_MASTER on TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
 LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code 
 where 2=2 and  "
            Baseqry += "   convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' and convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) <='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Baseqry += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Baseqry += " and 2=( case when TSPL_MILK_SRN_head.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            Baseqry += "GROUP BY TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,TSPL_VLC_DATA_UPLOADER_MASTER.shift,TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code"


            Baseqry += " UNION ALL
   select   TSPL_VLC_MASTER_HEAD.VLC_CODE	,MAX(TSPL_VLC_MASTER_HEAD.MCC)MCC,	MAX(TSPL_VLC_DATA_UPLOADER.DOC_NO) AS DOC_CODE	,CONVERT(VARCHAR,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)DOC_DATE,	SHIFT,	MAX(VSP_Code)DCS_Code,	MAX(VLC_Code_VLC_Uploader)DCS_Uploader_code,	MAX(VLC_Name)DCS_NAME	,MAX(ROUTE_CODE)ROUTE_CODE	, 0 AS SRNQTY	,0 AS SRNFAT_KG	,0 AS SRNSNF_KG	,0 AS SRNFatAVG,0 AS 	SRNSNFAVG,	COUNT(TSPl_MP_MAster.MP_Code)Farmer_Count	,SUM(TSPL_VLC_DATA_UPLOADER.QTY) AS FARMERQTY	
   ,SUM(fat_KG) AS FARMERFAT_KG	,SUM(snf_KG) AS FARMERSNF_KG	,SUM(fat) AS FARMERFatPer,	SUM(snf) AS  FARMERSNFPer	,
   SUM(TSPL_VLC_DATA_UPLOADER.fat_KG) 
/ NULLIF(SUM(TSPL_VLC_DATA_UPLOADER.qty), 0) AS FARMERFATAVG,
SUM(TSPL_VLC_DATA_UPLOADER.snf_KG) 
/ NULLIF(SUM(TSPL_VLC_DATA_UPLOADER.qty), 0) AS FARMERSNFAVG
 from TSPL_VLC_DATA_UPLOADER
 Left Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE
Left Join TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code

where 2=2 and "
            Baseqry += "   convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103)>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' and convert(date,TSPL_VLC_DATA_UPLOADER.DOC_DATE,103) <='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
            If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                Baseqry += " and 2=( case when TSPL_VLC_DATA_UPLOADER.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
            End If
            If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                Baseqry += " and 2=( case when TSPL_VLC_DATA_UPLOADER.DOC_DATE >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_MILK_SRN_head.DOC_DATE <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
            End If
            If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                Baseqry += " and TSPL_VLC_MASTER_HEAD.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
            End If
            Baseqry += "  GROUP BY DOC_DATE,shift,TSPL_VLC_MASTER_HEAD.VLC_Code"











            Baseqry += ")XXX "
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
                Qry = " SELECT 
 XXXX.VLC_CODE,MAX(convert(varchar,DOC_DATE,103))DOC_DATE,MAX(SHIFT)SHIFT,MAX(XXXX.DCS_Code)DCS_Code,CAST(MAX([DCS_Uploader_code]) AS INT) AS DCS_Uploader_code,MAX([DCS_NAME])[DCS_NAME],SUM(XXXX.[SRNQTY])[SRNQTY],sum(xxx.ACC_Qty_LTR)[SRNLTR],SUM(XXXX.SRNFAT_KG)SRNFAT_KG,SUM(XXXX.SRNSNF_KG)SRNSNF_KG,MAX(XXXX.SRNFatAVG)SRNFatAVG,MAX(XXXX.SRNSNFAVG)SRNSNFAVG,

COUNT(Farmer_Count) AS Farmer_Count,SUM(XXXX.FARMERQTY)FARMERQTY,SUM(XXXX.FARMERFAT_KG)FARMERFAT_KG,SUM(XXXX.FARMERSNF_KG)FARMERSNF_KG,SUM(XXXX.FARMERFatPer)FARMERFatPer,SUM(XXXX.FARMERSNFPer)FARMERSNFPer,sum( XXXX.FARMERFatAVG)FARMERFatAVG,SUM( XXXX.FARMERSNFAVG )FARMERSNFAVG from(" + Baseqry + ")XXXX GROUP BY VLC_CODE"
            End If



            '            Baseqry = "SELECT *
            'FROM (
            'Select  TSPL_MILK_SRN_head.VLC_CODE,COUNT(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code) AS Farmer_Count, MAX(TSPL_MILK_SRN_head.MCC_CODE)MCC_CODE, MAX(TSPL_MILK_SRN_head.DOC_CODE)DOC_CODE,CONVERT(date, TSPL_MILK_SRN_head.DOC_DATE,103)DOC_DATE,TSPL_MILK_SRN_head.SHIFT,max(TSPL_MILK_SRN_head.VSP_CODE) AS[DCS_Code],max(TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) as [DCS_Uploader_code],max(TSPL_VLC_MASTER_HEAD.VLC_Name) as [DCS_NAME],max(TSPL_MILK_SRN_head.ROUTE_CODE) as [ROUTE_CODE],sum(TSPL_MILK_SRN_DETAIL.qty) as [QTY],sum(TSPL_MILK_SRN_DETAIL.fat_kg)FAT_KG,sum(TSPL_MILK_SRN_DETAIL.SNF_KG)SNF_KG,
            '0 AS FatPer,0 AS SNFPer,
            ' sum(CAST((ISNULL(TSPL_MILK_SRN_DETAIL.FAT_KG,0) * 100 /NULLIF(TSPL_MILK_SRN_DETAIL.qty,0)) AS DECIMAL(18,2))) AS FatAVG,
            'sum( CAST((ISNULL(TSPL_MILK_SRN_DETAIL.SNF_KG,0) * 100 / NULLIF(TSPL_MILK_SRN_DETAIL.qty,0)) AS DECIMAL(18,2))) AS SNFAVG
            '                   from TSPL_MILK_SRN_DETAIL
            '                   LEFT OUTER JOIN TSPL_MILK_SRN_head ON TSPL_MILK_SRN_head.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE
            '				   LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_MASTER  ON TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code=TSPL_MILK_SRN_head.VLC_CODE
            '				   LEFT OUTER JOIN TSPL_VLC_DATA_UPLOADER_DETAIL ON TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code=TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code
            '                   LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VSP_Code=TSPL_MILK_SRN_head.VSP_CODE 
            '                   WHERE 2=2 and "
            '            Baseqry += "   convert(date,TSPL_MILK_SRN_head.DOC_DATE,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_MILK_SRN_head.DOC_DATE,103) <=convert(date,'" + txtToDate.Value + "' ,103)"

            '            Baseqry += " group by TSPL_MILK_SRN_head.DOC_DATE,TSPL_MILK_SRN_head.shift,
            '  TSPL_MILK_SRN_head.VLC_CODE
            '  UNION ALL
            ' select TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code AS VLC_CODE,COUNT(TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code)Farmer_Count, MAX(TSPL_VLC_MASTER_HEAD.MCC) AS MCC_CODE,MAX(TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code) AS DOC_CODE, TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date AS DOC_DATE,TSPL_VLC_DATA_UPLOADER_MASTER.shift AS SHIFT , MAX(VSP_CODE) AS DCS_CODE,
            'MAX( TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader) AS DCS_Uploader_code,
            ' MAX(TSPL_VLC_MASTER_HEAD.VLC_Name) AS DCS_NAME,MAX(TSPL_VLC_DATA_UPLOADER_MASTER.Route_Code) AS ROUTE_CODE,SUM(TSPL_VLC_DATA_UPLOADER_DETAIL.QTY) AS QTY,
            'sum(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)FAT_KG,
            ' sum(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer*TSPL_VLC_DATA_UPLOADER_DETAIL.QTY/100)SNF_KG
            ', SUM(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer )FatPer,
            'SUM(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer)SNFPer, 
            'CAST(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer * TSPL_VLC_DATA_UPLOADER_DETAIL.QTY / 100, 0)) * 100 / NULLIF(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.QTY, 0)), 0) AS DECIMAL(18,2)) AS FatAVG,
            'CAST(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer * TSPL_VLC_DATA_UPLOADER_DETAIL.QTY / 100, 0)) * 100 / NULLIF(SUM(ISNULL(TSPL_VLC_DATA_UPLOADER_DETAIL.QTY, 0)), 0)
            'AS DECIMAL(18,2)) AS SNFAVG
            ' from TSPL_VLC_DATA_UPLOADER_DETAIL
            ' left outer join TSPL_VLC_DATA_UPLOADER_MASTER on TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
            ' LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code 
            ' where 2=2 and "
            '            Baseqry += "   convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)>=convert(date,'" + txtFromDate.Value + "',103) and convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) <=convert(date,'" + txtToDate.Value + "' ,103)"

            '            Baseqry += " GROUP BY TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,TSPL_VLC_DATA_UPLOADER_MASTER.shift,TSPL_VLC_DATA_UPLOADER_MASTER.VLC_Code
            '		)XXX "
            '            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
            '                Qry = " SELECT 
            'VLC_CODE,COUNT(XXXX.Farmer_Count)Farmer_Count,MAX(XXXX.DOC_CODE)DOC_CODE,MAX(CONVERT(DATE,XXXX.DOC_DATE,103))DOC_DATE,MAX(XXXX.SHIFT)SHIFT,MAX(XXXX.DCS_Code)DCS_Code,MAX(XXXX.DCS_Uploader_code)DCS_Uploader_code,MAX(XXXX.DCS_NAME)DCS_NAME,MAX(XXXX.ROUTE_CODE)ROUTE_CODE,SUM(XXXX.QTY)QTY,SUM(XXXX.SNF_KG)SNF_KG,SUM(XXXX.FAT_KG)FAT_KG,

            'SUM(XXXX.FatPer)FatPer,SUM(XXXX.SNFPer)SNFPer,MAX(XXXX.FatAVG)FatAVG,MAX(XXXX.SNFAVG)SNFAVG
            'FROM (" + Baseqry + ")XXXX GROUP BY VLC_CODE"
            ' End If
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
                dt = clsDBFuncationality.GetDataTable(Qry)
            Else
                dt = clsDBFuncationality.GetDataTable(Baseqry)
            End If
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Baseqry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv.DataSource = Nothing
                gv.Rows.Clear()
                gv.Columns.Clear()
                gv.GroupDescriptors.Clear()
                gv.MasterTemplate.SummaryRowsBottom.Clear()
                gv.MasterView.Refresh()
                gv.DataSource = dt
                For ii As Integer = 0 To gv.Columns.Count - 1
                    gv.Columns(ii).ReadOnly = True
                Next
                RadPageView1.SelectedPage = RadPageViewPage2
                gv.EnableFiltering = True
                'SetGridFormat()
                SetGridFormationOFGV1Collection()
                ' View()
                'View1()
                ' EnableDisableCntrl(False)
                gv.BestFitColumns()
                '   EnableDisableCntrl(False)
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found To Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub View1()
        If gv.Rows.Count > 0 Then
            Dim view As New ColumnGroupsViewDefinition()
            view.ColumnGroups.Add(New GridViewColumnGroup(" "))
            view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            'If rdbDetails.Checked = True Then

            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("DocumentDate").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Description").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Temp").Name)

            view.ColumnGroups.Add(New GridViewColumnGroup("Farmer data"))
            view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("Entered_QTY").Name)

            gv.ViewDefinition = view
        End If
    End Sub
    Sub EnableDisableCntrl(ByVal val As Boolean)
        cmbEntrySource.Enabled = False
        cboUnit.Enabled = False
        txtMCC.Enabled = False
        txtRoute.Enabled = False

    End Sub
    Sub SetGridFormationDBT()
        gv.TableElement.TableHeaderHeight = 40
        gv.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To gv.Columns.Count - 1
            gv.Columns(ii).ReadOnly = True
            gv.Columns(ii).IsVisible = True
            'If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "dcs collection v/s farmer collection day & shift wise") = CompairStringResult.Equal Then
            '    gv.Columns("mcc_code").IsVisible = False
            '    gv.Columns("doc_code").IsVisible = False
            '    gv.Columns("route_code").IsVisible = False
            'End If
            gv.Columns("VLC_CODE").IsVisible = False
            gv.Columns("MCC_CODE").IsVisible = True
            gv.Columns("ROUTE_CODE").IsVisible = False
            gv.Columns("SHIFT").IsVisible = False
            gv.Columns("MCC_CODE").IsVisible = False
            gv.Columns("MCC_CODE").HeaderText = "MCC CODE"

            gv.Columns("Zone_Code").HeaderText = "Zone Code"
            gv.Columns("Zone_Name").HeaderText = "Zone Name"


            gv.Columns("DOC_DATE").IsVisible = False
            gv.Columns("Farmer_Count").HeaderText = "Farmer Count"

            gv.Columns("DOC_DATE").HeaderText = "Document Date"
            gv.Columns("DCS_Code").IsVisible = True
            gv.Columns("DOC_Code").IsVisible = False

            gv.Columns("DCS_Code").HeaderText = "DCS Code"
            gv.Columns("DCS_Uploader_code").IsVisible = True
            gv.Columns("DCS_Uploader_code").HeaderText = "DCS Uploader code"
            gv.Columns("DCS_NAME").IsVisible = True
            gv.Columns("DCS_NAME").HeaderText = "DCS Name"
            'gv.Columns("MP_CODE").IsVisible = True

            gv.Columns("RECOQTY").IsVisible = True
            gv.Columns("RECOQTY").HeaderText = "RECO Qty"
            gv.Columns("RECOFAT_KG").IsVisible = False
            gv.Columns("RECOSNF_KG").IsVisible = False
            gv.Columns("RECOFatAVG").IsVisible = True
            gv.Columns("RECOFatAVG").HeaderText = "RECO Fat Avg"
            gv.Columns("RECOSNFAVG").IsVisible = True
            gv.Columns("RECOSNFAVG").HeaderText = "RECO Snf Avg"
            gv.Columns("FARMERQTY").IsVisible = True
            gv.Columns("FARMERQTY").HeaderText = "Farmer Qty"
            gv.Columns("FARMERFAT_KG").IsVisible = False
            gv.Columns("FARMERSNF_KG").IsVisible = False
            gv.Columns("FARMERSNFPer").IsVisible = False
            gv.Columns("FARMERFatPer").IsVisible = False
            gv.Columns("FARMERFatAVG").IsVisible = True
            gv.Columns("FARMERFatAVG").HeaderText = "Farmer Fat Avg"
            gv.Columns("FARMERSNFAVG").IsVisible = True
            gv.Columns("FARMERSNFAVG").HeaderText = "Farmer Snf Avg"
            'gv1.Columns("VLC_CODE").HeaderText = "Document No."
            'gv1.Columns("Document_No").HeaderText = "Document No."

        Next
        Dim summaryRowItemB As New GridViewSummaryRowItem()
        Dim SRNQTY As New GridViewSummaryItem("SRNQTY", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(SRNQTY)
        'Dim SRNFatAVG As New GridViewSummaryItem("SRNFatAVG", "{0:n2}", GridAggregateFunction.Sum)
        'summaryRowItemB.Add(SRNFatAVG)
        'Dim SRNSNFAVG As New GridViewSummaryItem("SRNSNFAVG", "{0:n2}", GridAggregateFunction.Sum)
        'summaryRowItemB.Add(SRNSNFAVG)
        Dim RECOQTY As New GridViewSummaryItem("RECOQTY", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(RECOQTY)


        Dim FARMERQTY As New GridViewSummaryItem("FARMERQTY", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(FARMERQTY)
        'Dim FARMERFatAVG As New GridViewSummaryItem("FARMERFatAVG", "{0:n2}", GridAggregateFunction.Sum)
        'summaryRowItemB.Add(FARMERFatAVG)
        'Dim FARMERSNFAVG As New GridViewSummaryItem("FARMERSNFAVG", "{0:n2}", GridAggregateFunction.Sum)
        'summaryRowItemB.Add(FARMERSNFAVG)
        gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItemB)
        gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        gv.AutoSizeRows = True
        gv.BestFitColumns()
        gv.MasterTemplate.AutoExpandGroups = True
    End Sub
    Sub SetGridFormationOFGV1Collection()
        gv.TableElement.TableHeaderHeight = 40
        gv.MasterTemplate.ShowRowHeaderColumn = False
        Dim summaryRowItem As New GridViewSummaryRowItem()
        For ii As Integer = 0 To gv.Columns.Count - 1
            gv.Columns(ii).ReadOnly = True
            gv.Columns(ii).IsVisible = True
            gv.Columns("VLC_CODE").IsVisible = False
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Day & Shift Wise") = CompairStringResult.Equal Then
                gv.Columns("MCC_CODE").IsVisible = False
                gv.Columns("DOC_CODE").IsVisible = False
                gv.Columns("ROUTE_CODE").IsVisible = False
            End If
            gv.Columns("DOC_DATE").IsVisible = True
            gv.Columns("Farmer_Count").HeaderText = "Farmer Count"

            gv.Columns("DOC_DATE").HeaderText = "Document Date"
            '  gv.Columns("SHIFT").IsVisible = True
            gv.Columns("DCS_Code").IsVisible = True
            gv.Columns("DCS_Code").HeaderText = "DCS Code"
            gv.Columns("DCS_Uploader_code").IsVisible = True
            gv.Columns("DCS_Uploader_code").HeaderText = "DCS Uploader code"
            gv.Columns("DCS_NAME").IsVisible = True
            gv.Columns("DCS_NAME").HeaderText = "DCS Name"
            gv.Columns("Zone_Code").HeaderText = "Zone Code"
            gv.Columns("Zone_Name").HeaderText = "Zone Name"

            gv.Columns("SRNQTY").IsVisible = False
            gv.Columns("SRNQTY").VisibleInColumnChooser = True
            gv.Columns("SRNQTY").HeaderText = "SRN Qty"

            gv.Columns("SRNLTR").IsVisible = True
            gv.Columns("SRNLTR").VisibleInColumnChooser = True
            gv.Columns("SRNLTR").HeaderText = "SRN Qty LTR"



            gv.Columns("SRNFAT_KG").IsVisible = False
            gv.Columns("SRNSNF_KG").IsVisible = False
            gv.Columns("SRNFatAVG").IsVisible = True
            gv.Columns("SRNFatAVG").HeaderText = "SRN Fat Avg"
            gv.Columns("SRNSNFAVG").IsVisible = True
            gv.Columns("SRNSNFAVG").HeaderText = "SRN Snf Avg"
            gv.Columns("FARMERQTY").IsVisible = True
            gv.Columns("FARMERQTY").HeaderText = "Farmer Qty"
            gv.Columns("FARMERFAT_KG").IsVisible = False
            gv.Columns("FARMERSNF_KG").IsVisible = False
            gv.Columns("FARMERSNFPer").IsVisible = False
            gv.Columns("FARMERFatPer").IsVisible = False
            gv.Columns("FARMERFatAVG").IsVisible = True
            gv.Columns("FARMERFatAVG").HeaderText = "Farmer Fat Avg"
            gv.Columns("FARMERSNFAVG").IsVisible = True
            gv.Columns("FARMERSNFAVG").HeaderText = "Farmer Snf Avg"

            'gv.Columns("Difference").HeaderText = "DIFFERENCE"

            'gv1.Columns("VLC_CODE").HeaderText = "Document No."
            'gv1.Columns("Document_No").HeaderText = "Document No."

        Next
        Dim summaryRowItemB As New GridViewSummaryRowItem()
        Dim SRNQTY As New GridViewSummaryItem("SRNQTY", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(SRNQTY)

        Dim SRNLTR As New GridViewSummaryItem("SRNLTR", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(SRNLTR)

        Dim FARMERQTY As New GridViewSummaryItem("FARMERQTY", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(FARMERQTY)

        gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItemB)
        gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        gv.AutoSizeRows = True
        gv.BestFitColumns()
        gv.MasterTemplate.AutoExpandGroups = True
    End Sub
    Private Sub DCSFarmerCollectionSummary()
        Try
            Dim Qry As String = Nothing


            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                gv.DataSource = Nothing
                gv.Rows.Clear()
                gv.Columns.Clear()
                gv.GroupDescriptors.Clear()
                gv.MasterTemplate.SummaryRowsBottom.Clear()
                gv.MasterView.Refresh()
                gv.DataSource = dt
                For ii As Integer = 0 To gv.Columns.Count - 1
                    gv.Columns(ii).ReadOnly = True
                Next
                RadPageView1.SelectedPage = RadPageViewPage2
                gv.EnableFiltering = True
                'SetGridFormat()
                ' SetGridFormationOFGV1Collection()
                ' View()
                ' EnableDisableCntrl(False)
                gv.BestFitColumns()

            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found To Display", Me.Text)
                Exit Sub
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub FarmerWiseMilkCollection()
        Try
            Dim Qry As String = "Select ROW_NUMBER() Over (Order By (Select 1)) As [S.No.],finalQry.* "
            If isPrint Then
                Qry &= " , TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.State,TSPL_STATE_MASTER.STATE_NAME,'" & objCommonVar.CurrentUser & "' As PrintBy "
            End If
            Qry &= " from(Select [Union],MAX(Doc_Date)Doc_Date,
CONVERT(decimal(18,2),Sum(MorningQty))MorningQty,
CONVERT(decimal(18,2),Sum(case When (MorningFATKG)>0 Then ((MorningFATKG/MorningQty)*100) Else 0 End)) As MorningFAT,
CONVERT(decimal(18,2),Sum(Case When (MorningSNFKG)>0 Then ((MorningSNFKG/MorningQty)*100) Else 0 End)) As MorningSNF,
CONVERT(decimal(18,2),Sum(MorningAmount)) As MorningAmount,
CONVERT(decimal(18,2),Sum(EveningQty))EveningQty,
CONVERT(decimal(18,2),Sum(Case When (EveningFATKG)>0 Then ((EveningFATKG/EveningQty)*100) Else 0 End)) As EveningFAT,
CONVERT(decimal(18,2),Sum(Case When (EveningSNFKG)>0 Then ((EveningSNFKG/EveningQty)*100) Else 0 End)) As EveningSNF,
CONVERT(decimal(18,2),Sum(EveningAmount)) As EveningAmount,CONVERT(decimal(18,2),Sum(MorningQty + EveningQty)) As TotalQuantity,max(ZONE_CODE)[Zone_Code],max(Description) as [Zone_Name]
from 
(Select [Union],Max(Convert(Varchar(10),Doc_Date,103))Doc_Date,
Sum(Case When shift='M' Then Qty Else 0 End) As MorningQty,
Sum(Case When shift='M' Then fat_KG Else 0 End) As MorningFATKG,
Sum(Case When shift='M' Then snf_KG Else 0 End) As MorningSNFKG,
Sum(Case When shift='M' Then Amount Else 0 End) As MorningAmount,
SUM(Case When shift='E' Then Qty Else 0 End) As EveningQty,  
Sum(Case When shift='E' Then fat_KG Else 0 End) As EveningFATKG,
Sum(Case When shift='E' Then snf_KG Else 0 End) As EveningSNFKG,
Sum(Case When shift='E' Then Amount Else 0 End) As EveningAmount ,max(ZONE_CODE)ZONE_CODE,max(Description) as [Description] "
            Qry &= " from(" & ReturnFarmerBaseQry() & ")final Group By [Union],shift)BAseQry group by [Union]  "
            Qry &= ")finalQry "
            If isPrint Then
                Qry &= " Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State "
            End If
            Qry &= " Order By [Union]"
            dt = Nothing
            dt = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If isPrint Then
                    Dim frm As New frmCrystalReportViewer()
                    frm.funreport(Me.Form_ID, CrystalReportFolder.CommonForUnionAndCattlefeed, dt, "crptFarmerWiseMilkCollection", "Farmer Wise Milk Collection")
                    frm = Nothing
                Else
                    gv.DataSource = Nothing
                    gv.Rows.Clear()
                    gv.Columns.Clear()
                    gv.DataSource = dt
                    gv.GroupDescriptors.Clear()
                    gv.MasterTemplate.SummaryRowsBottom.Clear()
                    gv.TableElement.TableHeaderHeight = 25
                    gv.MasterTemplate.ShowRowHeaderColumn = True
                    For ii As Integer = 0 To gv.Columns.Count - 1
                        gv.Columns(ii).ReadOnly = True
                        gv.Columns(ii).IsVisible = True
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.BestFitColumns()
                    setGridFormat()
                    RadPageView1.SelectedPage = RadPageViewPage2
                    ReStoreGridLayout()
                    EnableDisableControl(False)
                    View()
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    summaryRowItem.Add(New GridViewSummaryItem("MorningQty", "{0:N2}", GridAggregateFunction.Sum))
                    summaryRowItem.Add(New GridViewSummaryItem("MorningAmount", "{0:N2}", GridAggregateFunction.Sum))
                    summaryRowItem.Add(New GridViewSummaryItem("EveningQty", "{0:N2}", GridAggregateFunction.Sum))
                    summaryRowItem.Add(New GridViewSummaryItem("EveningAmount", "{0:N2}", GridAggregateFunction.Sum))
                    summaryRowItem.Add(New GridViewSummaryItem("TotalQuantity", "{0:N2}", GridAggregateFunction.Sum))
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                    gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not found !", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub DCSFarmerCollectionDetails()
        Try
            Dim Qry As String = "Select ROW_NUMBER() Over (Order By (Select 1)) As [S.No.],finalQry.*,CONVERT(decimal(18,2),(MorningQty + EveningQty)) As TotalQuantity,CONVERT(decimal(18,2),(MorningFATKG + EveningFATKG)) as TotalFATKG,CONVERT(decimal(18,2),(MorningSNFKG + EveningSNFKG)) as TotalSNFKG "
            If isPrint Then
                Qry &= " , '" + clsCommon.GetPrintDate(clsCommon.myCDate(txtFromDate.Value), "dd/MM/yyyy") + "' AS FromDate,'" + clsCommon.GetPrintDate(clsCommon.myCDate(txtToDate.Value), "dd/MM/yyyy") + "' AS ToDate, TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.State,TSPL_STATE_MASTER.STATE_NAME,'" & objCommonVar.CurrentUser & "' As PrintBy "
            End If
            Qry &= " from ( Select [Union],MAX(Doc_Date) as Date, max(TotalDCS)TotalDCS,sum(MorningDCSCount)MorningDCSCount,SUM(MorningMPCount)MorningMPCount,CONVERT(decimal(18,2),sum(MorningQty/1000))MorningQty,CONVERT(decimal(18,2),SUM(MorningFATKG/1000))MorningFATKG,CONVERT(decimal(18,2),sum(MorningSNFKG/1000))MorningSNFKG,CONVERT(decimal(18,2),sum(MorningAmount/1000))MorningAmount,sum(EveningDCSCount)EveningDCSCount,sum(EveningMPCount)EveningMPCount,CONVERT(decimal(18,2),sum(EveningQty/1000))EveningQty,CONVERT(decimal(18,2),sum(EveningFATKG/1000))EveningFATKG,CONVERT(decimal(18,2),sum(EveningSNFKG/1000))EveningSNFKG,CONVERT(decimal(18,2),sum(EveningAmount/1000))EveningAmount  ,max(ZONE_CODE)ZONE_CODE,max(ZONE_NAME)ZONE_NAME from ( 
        select   [Union],Max(Convert(Varchar(10),Doc_Date,103)) as Doc_Date,Max([DCSCount]) As TotalDCS,
Case When Max(shift)='M'  Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End As MorningDCSCount, (Case When shift='M' Then Count(Distinct MP_Code) Else 0 End) As MorningMPCount,Sum(Case When shift='M' Then Qty Else 0 End) As MorningQty,
Sum(Case When shift='M' Then fat_KG Else 0 End) As MorningFATKG,
Sum(Case When shift='M' Then snf_KG Else 0 End) As MorningSNFKG,
Sum(Case When shift='M' Then Amount Else 0 End) As MorningAmount,
(Case When max(shift)='E'  THEN COUNT(Distinct vlc_code_vlc_Uploader) ELSE 0 end) as EveningDCSCount,
(Case When shift='E' Then Count(Distinct MP_Code) Else 0 End) As EveningMPCount,
SUM(Case When shift='E' Then Qty Else 0 End) As EveningQty,  
Sum(Case When shift='E' Then fat_KG Else 0 End) As EveningFATKG,
Sum(Case When shift='E' Then snf_KG Else 0 End) As EveningSNFKG,
Sum(Case When shift='E' Then Amount Else 0 End) As EveningAmount, max(ZONE_CODE)ZONE_CODE,max(Description)ZONE_NAME  "
            Qry &= " from(" & ReturnFarmerBaseQry() & ")final Group By [Union],Doc_Date,shift)BAseQry group by [Union],Doc_Date  "
            Qry &= ")finalQry "
            If isPrint Then
                Qry &= " Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State"
            End If
            Qry &= " Order By [Union] "
            dt = Nothing
            dt = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If isPrint Then
                    Dim frm As New frmCrystalReportViewer()
                    Try
                        frm.funreport(Me.Form_ID, CrystalReportFolder.CommonForUnionAndCattlefeed, dt, "crptDCSFarmerCollectionDetails", "Farmer Collection Entry Status")
                        frm = Nothing
                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
                Else
                    gv.DataSource = Nothing
                    gv.Rows.Clear()
                    gv.Columns.Clear()
                    gv.DataSource = dt
                    gv.GroupDescriptors.Clear()
                    gv.MasterTemplate.SummaryRowsBottom.Clear()
                    gv.TableElement.TableHeaderHeight = 25
                    gv.MasterTemplate.ShowRowHeaderColumn = True
                    For ii As Integer = 0 To gv.Columns.Count - 1
                        gv.Columns(ii).ReadOnly = True
                        gv.Columns(ii).IsVisible = True
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.BestFitColumns()
                    setGridFormat()
                    RadPageView1.SelectedPage = RadPageViewPage2
                    ReStoreGridLayout()
                    EnableDisableControl(False)
                    View()
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    For i As Integer = 3 To gv.Columns.Count - 1
                        summaryRowItem.Add(New GridViewSummaryItem(gv.Columns(i).Name, "{0:N2}", GridAggregateFunction.Sum))
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                    gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not found !", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Function ReturnDCSFarmerCollectionBaseQry() As String
        Dim i As Integer = 0
        Dim qry As String = ""
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each strUnion As DataRow In dt.Rows
                Dim dcsCount As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Count(Distinct TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code) from " & clsCommon.myCstr(strUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_DETAIL
Left Join " & clsCommon.myCstr(strUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_HEAD On TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
Where TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtToDate.Value.AddMonths(-6), "dd/MMM/yyyy") & "' And TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'"))
                If dcsCount <= 0 Then
                    dcsCount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select COUNT(1) from " & clsCommon.myCstr(strUnion("Database_Name")) & ".dbo.TSPL_VLC_MASTER_HEAD where Active=1"))
                End If
                If i <> 0 Then
                    qry &= " Union All "
                End If
                qry &= " Select * from ("
                qry &= "select '" & strUnion("Location_Name") & "' As [Union],Cast('" & clsCommon.myCstr(dcsCount) & "' As Int) As DCSCount,Doc_No,Convert(Varchar(10),Doc_Date,103)Doc_Date,File_Date,shift,
TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader,
TSPL_VLC_MASTER_HEAD.VLC_Name,
TSPl_MP_MAster.MP_CODE,TSPl_MP_MAster.MP_Name,
TSPl_MP_MAster.MP_Code_VLC_Uploader As MP_Uploader_Code,qty,fat,snf,Rate,Amount,fat_KG,snf_KG,Entry_Source As tttype from " & strUnion("Database_Name") & ".dbo.TSPL_VLC_DATA_UPLOADER
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE
Left Join " & strUnion("Database_Name") & ".dbo.TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
where Doc_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And Doc_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
                If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when Doc_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and Doc_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
                End If
                If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when Doc_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and Doc_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
                End If
                qry &= Environment.NewLine & " Union All " & Environment.NewLine
                qry &= " select '" & strUnion("Location_Name") & "' As [Union],Cast('" & clsCommon.myCstr(dcsCount) & "' As Int) As DCSCount,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code As Doc_No,
Convert(Varchar(10),TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) As Doc_Date,'' As File_Date,TSPL_VLC_DATA_UPLOADER_MASTER.Shift,
TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPl_MP_MAster.MP_CODE,TSPl_MP_MAster.MP_Name,
TSPl_MP_MAster.MP_Code_VLC_Uploader As MP_Uploader_Code,
TSPL_VLC_DATA_UPLOADER_DETAIL.qty,TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer,TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer,TSPL_VLC_DATA_UPLOADER_DETAIL.Rate,
TSPL_VLC_DATA_UPLOADER_DETAIL.Amount,
((TSPL_VLC_DATA_UPLOADER_DETAIL.qty*TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer)/100) As FAT_KG,((TSPL_VLC_DATA_UPLOADER_DETAIL.qty*TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer)/100) As SNF_KG,
TSPL_VLC_DATA_UPLOADER_MASTER.Dock_Collection_Milk_Type As  tttype
from " & strUnion("Database_Name") & ".dbo.TSPL_VLC_DATA_UPLOADER_DETAIL
Left Outer Join " & strUnion("Database_Name") & ".dbo.TSPL_VLC_DATA_UPLOADER_MASTER On TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_CODE
Left Join " & strUnion("Database_Name") & ".dbo.TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
Where 1=1 "
                If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
                End If
                If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
                End If

                qry &= Environment.NewLine & " Union All " & Environment.NewLine
                qry &= "select '" & strUnion("Location_Name") & "' As [Union],Cast('" & clsCommon.myCstr(dcsCount) & "' As Int) As DCSCount,'' As Doc_No,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") & "' As Doc_Date,'' As File_Date,'' As shift,
'' As Vlc_Code_VLC_Uploader,'' As VLC_Name,'' As MP_CODE,'' As MP_Name,'' As MP_Uploader_Code,0 As qty,0 As fat,0 As snf,0 As Rate,0 As Amount,0 As fat_KG,0 As snf_KG,'' As  tttype ) As " & strUnion("Database_Name")


                i += 1
            Next
        Else
            Throw New Exception("Database name not found !")
        End If
        Return qry
    End Function
    Private Sub FarmerCollectionEntryStatus()
        Try
            Dim Qry As String = "Select ROW_NUMBER() Over (Order By (Select 1)) As [S.No.],finalQry.* "
            If isPrint Then
                Qry &= " , TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.State,TSPL_STATE_MASTER.STATE_NAME,'" & objCommonVar.CurrentUser & "' As PrintBy "
            End If
            Qry &= " from( Select [Union],MAX(Doc_Date)Doc_Date,SUM(MorningQty)MorningQty,SUM(EveningQty)EveningQty, MAX(ZONE_CODE)ZONE_CODE, MAX(ZONE_NAME)ZONE_NAME  from (Select
[Union],Max(Convert(Varchar(10),Doc_Date,103))Doc_Date,(Case When shift='M' Then Count(Distinct MP_Code) Else 0 End) As MorningQty,
(Case When shift='E' Then Count(Distinct MP_Code) Else 0 End) As EveningQty, MAX(ZONE_CODE)ZONE_CODE, MAX(Description)ZONE_NAME "
            Qry &= " from(" & ReturnFarmerBaseQry() & ")final Group By [Union],shift)BAseQry group by [Union]  "
            Qry &= ")finalQry "
            If isPrint Then
                Qry &= " Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State"
            End If
            Qry &= " Order By [Union] "
            dt = Nothing
            dt = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If isPrint Then
                    Dim frm As New frmCrystalReportViewer()
                    frm.funreport(Me.Form_ID, CrystalReportFolder.CommonForUnionAndCattlefeed, dt, "crptFarmerCollectionEntryStatus", "Farmer Collection Entry Status")
                    frm = Nothing
                Else
                    gv.DataSource = Nothing
                    gv.Rows.Clear()
                    gv.Columns.Clear()
                    gv.DataSource = dt
                    gv.GroupDescriptors.Clear()
                    gv.MasterTemplate.SummaryRowsBottom.Clear()
                    gv.TableElement.TableHeaderHeight = 25
                    gv.MasterTemplate.ShowRowHeaderColumn = True
                    For ii As Integer = 0 To gv.Columns.Count - 1
                        gv.Columns(ii).ReadOnly = True
                        gv.Columns(ii).IsVisible = True
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.BestFitColumns()
                    setGridFormat()
                    RadPageView1.SelectedPage = RadPageViewPage2
                    ReStoreGridLayout()
                    EnableDisableControl(False)
                    View()
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    For i As Integer = 2 To gv.Columns.Count - 1
                        summaryRowItem.Add(New GridViewSummaryItem(gv.Columns(i).Name, "{0:N2}", GridAggregateFunction.Sum))
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                    gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not found !", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Function ReturnFarmerBaseQry() As String
        Dim i As Integer = 0
        Dim qry As String = ""
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each strUnion As DataRow In dt.Rows
                Dim dcsCount As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Count(Distinct TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code) from " & clsCommon.myCstr(strUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_DETAIL
Left Join " & clsCommon.myCstr(strUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_HEAD On TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
Where TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtToDate.Value.AddMonths(-6), "dd/MMM/yyyy") & "' And TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'"))
                If dcsCount <= 0 Then
                    dcsCount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select COUNT(1) from " & clsCommon.myCstr(strUnion("Database_Name")) & ".dbo.TSPL_VLC_MASTER_HEAD where Active=1"))
                End If
                If i <> 0 Then
                    qry &= " Union All "
                End If
                qry &= " Select * from ("
                qry &= "select '" & strUnion("Location_Name") & "' As [Union],Cast('" & clsCommon.myCstr(dcsCount) & "' As Int) As DCSCount,Doc_No,Convert(Varchar(10),Doc_Date,103)Doc_Date,File_Date,shift,
TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader,
TSPL_VLC_MASTER_HEAD.VLC_Name,
TSPl_MP_MAster.MP_CODE,TSPl_MP_MAster.MP_Name,
TSPl_MP_MAster.MP_Code_VLC_Uploader As MP_Uploader_Code,qty,fat,snf,Rate,Amount,fat_KG,snf_KG,Entry_Source As tttype ,TSPL_ZONE_MASTER.ZONE_CODE,TSPL_ZONE_MASTER.Description from " & strUnion("Database_Name") & ".dbo.TSPL_VLC_DATA_UPLOADER
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_ZONE_MASTER On TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code



Left Join " & strUnion("Database_Name") & ".dbo.TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
where Doc_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And Doc_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "' "
                If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when Doc_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and Doc_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
                End If
                If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when Doc_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and Doc_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
                End If
                If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                    qry &= " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
                End If
                qry &= Environment.NewLine & " Union All " & Environment.NewLine
                qry &= " select '" & strUnion("Location_Name") & "' As [Union],Cast('" & clsCommon.myCstr(dcsCount) & "' As Int) As DCSCount,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code As Doc_No,
Convert(Varchar(10),TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) As Doc_Date,'' As File_Date,TSPL_VLC_DATA_UPLOADER_MASTER.Shift,
TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPl_MP_MAster.MP_CODE,TSPl_MP_MAster.MP_Name,
TSPl_MP_MAster.MP_Code_VLC_Uploader As MP_Uploader_Code,
TSPL_VLC_DATA_UPLOADER_DETAIL.qty,TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer,TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer,TSPL_VLC_DATA_UPLOADER_DETAIL.Rate,
TSPL_VLC_DATA_UPLOADER_DETAIL.Amount,
((TSPL_VLC_DATA_UPLOADER_DETAIL.qty*TSPL_VLC_DATA_UPLOADER_DETAIL.FatPer)/100) As FAT_KG,((TSPL_VLC_DATA_UPLOADER_DETAIL.qty*TSPL_VLC_DATA_UPLOADER_DETAIL.SNFPer)/100) As SNF_KG,
TSPL_VLC_DATA_UPLOADER_MASTER.Dock_Collection_Milk_Type As  tttype,TSPL_ZONE_MASTER.ZONE_CODE,TSPL_ZONE_MASTER.Description
from " & strUnion("Database_Name") & ".dbo.TSPL_VLC_DATA_UPLOADER_DETAIL
Left Outer Join " & strUnion("Database_Name") & ".dbo.TSPL_VLC_DATA_UPLOADER_MASTER On TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_CODE
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code
Left Join " & strUnion("Database_Name") & ".dbo.TSPL_ZONE_MASTER On TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
Left Join " & strUnion("Database_Name") & ".dbo.TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
Where 1=1 and convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) >='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)<= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'  "
                If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
                End If
                If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                    qry &= " and 2=( case when TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
                End If
                If TxtZone.arrValueMember IsNot Nothing AndAlso TxtZone.arrValueMember.Count > 0 Then
                    qry &= " AND TSPL_ZONE_MASTER.Zone_Code  IN (" + clsCommon.GetMulcallString(TxtZone.arrValueMember) + ") "
                End If
                qry &= Environment.NewLine & " Union All " & Environment.NewLine
                qry &= "select '" & strUnion("Location_Name") & "' As [Union],Cast('" & clsCommon.myCstr(dcsCount) & "' As Int) As DCSCount,'' As Doc_No,'" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") & "' As Doc_Date,'' As File_Date,'' As shift,
'' As Vlc_Code_VLC_Uploader,'' As VLC_Name,'' As MP_CODE,'' As MP_Name,'' As MP_Uploader_Code,0 As qty,0 As fat,0 As snf,0 As Rate,0 As Amount,0 As fat_KG,0 As snf_KG,'' As  tttype, '' AS ZONE_CODE, '' AS Description ) As " & strUnion("Database_Name")


                i += 1
            Next
        Else
            Throw New Exception("Database name not found !")
        End If
        Return qry
    End Function

    Private Sub FarmerWiseCollectionAtDCS()
        Try
            Dim Qry As String = "Select Row_Number() Over (Order By (Select 1)) As [S.No.],finalQry.* "
            If isPrint Then
                Qry &= " , TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.Add2,TSPL_COMPANY_MASTER.Add3,TSPL_COMPANY_MASTER.State,TSPL_STATE_MASTER.STATE_NAME,'" & objCommonVar.CurrentUser & "' As PrintBy "
            End If
            Qry &= " from(Select [Union],Convert(Varchar(10),Max([Date]),103)[Date],Max([DCS])[DCS],Sum([MorningManual])[MorningManual],Sum([MorningAuto])[MorningAuto],Min([MorningPending])[MorningPending],Sum([EveningManual])[EveningManual],Sum([EveningAuto])[EveningAuto],Min([EveningPending])[EveningPending],MAX(ZONE_CODE)[Zone_Code],MAX(Description)[Zone_Name] from (Select [Union],Max([DCSCount]) As [DCS],Max(Doc_Date) As [Date], Case When Max(shift)='M' And Max(tttype) IN ('MOBILE APP','MANUAL') Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End As [MorningManual], 
Case When (shift)='M' And Max(tttype)='REIL' Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End As [MorningAuto],Max([DCSCount])-(Case When (shift)='M' And Max(tttype) In ('REIL','MOBILE APP','MANUAL') Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End) As [MorningPending],
 Case When (shift)='E' And Max(tttype) IN ('MOBILE APP','MANUAL') Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End As [EveningManual], 
Case When (shift)='E' And Max(tttype)='REIL' Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End As [EveningAuto],
Max([DCSCount])-(Case When Max(shift)='E' And Max(tttype) In ('REIL','MOBILE APP','MANUAL') Then COUNT(Distinct vlc_code_vlc_Uploader) Else 0 End) As [EveningPending],MAX(ZONE_CODE)ZONE_CODE,MAX(Description)Description "

            Qry &= "from(" & ReturnFarmerBaseQry() & ")final Group By [Union],Shift)BaseQry Group By [Union] "
            Qry &= ")finalQry "
            If isPrint Then
                Qry &= " Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State"
            End If
            Qry &= " Order By [Union] "
            dt = Nothing
            dt = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If isPrint Then
                    Dim frm As New frmCrystalReportViewer()
                    frm.funreport(Me.Form_ID, CrystalReportFolder.CommonForUnionAndCattlefeed, dt, "crptFarmerWiseCollectionAtDCS", "Farmer Wise Collection at DCS")
                    frm = Nothing
                Else
                    gv.DataSource = Nothing
                    gv.Rows.Clear()
                    gv.Columns.Clear()
                    gv.DataSource = dt
                    gv.GroupDescriptors.Clear()
                    gv.MasterTemplate.SummaryRowsBottom.Clear()
                    gv.TableElement.TableHeaderHeight = 25
                    gv.MasterTemplate.ShowRowHeaderColumn = True
                    For ii As Integer = 0 To gv.Columns.Count - 1
                        gv.Columns(ii).ReadOnly = True
                        gv.Columns(ii).IsVisible = True
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.BestFitColumns()
                    setGridFormat()
                    RadPageView1.SelectedPage = RadPageViewPage2
                    ReStoreGridLayout()
                    EnableDisableControl(False)
                    View()
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    For i As Integer = 2 To gv.Columns.Count - 1
                        summaryRowItem.Add(New GridViewSummaryItem(gv.Columns(i).Name, "{0:N2}", GridAggregateFunction.Sum))
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                    gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub setGridFormat()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Collection at DCS") = CompairStringResult.Equal Then
                gv.Columns("MorningManual").HeaderText = "Manual"
                gv.Columns("MorningAuto").HeaderText = "Auto/iOT"
                gv.Columns("MorningPending").HeaderText = "Pending"
                gv.Columns("EveningManual").HeaderText = "Manual"
                gv.Columns("EveningAuto").HeaderText = "Auto/iOT"
                gv.Columns("EveningPending").HeaderText = "Pending"
                gv.Columns("EveningAuto").HeaderText = "Auto/iOT"
                gv.Columns("EveningPending").HeaderText = "Pending"
                gv.Columns("Zone_Code").HeaderText = "Zone Code"
                gv.Columns("Zone_Name").HeaderText = "Zone Name"


            End If

            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Entry Status") = CompairStringResult.Equal Then
                gv.Columns("Doc_Date").HeaderText = "Date"
                gv.Columns("MorningQty").HeaderText = "Morning"
                gv.Columns("EveningQty").HeaderText = "Evening"
                gv.Columns("ZONE_CODE").HeaderText = "Zone Code"
                gv.Columns("ZONE_NAME").HeaderText = "Zone Name"

            End If

            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Milk Collection") = CompairStringResult.Equal Then
                gv.Columns("MorningQty").HeaderText = "Quantity"
                gv.Columns("MorningFat").HeaderText = "FAT %"
                gv.Columns("MorningSnf").HeaderText = "SNF %"
                gv.Columns("MorningAmount").HeaderText = "Amount"
                gv.Columns("EveningQty").HeaderText = "Quantity"
                gv.Columns("EveningFat").HeaderText = "FAT %"
                gv.Columns("EveningSnf").HeaderText = "SNF %"
                gv.Columns("EveningAmount").HeaderText = "Amount"
                gv.Columns("TotalQuantity").HeaderText = "Total Quantity"
                gv.Columns("Zone_Code").HeaderText = "Zone Code"
                gv.Columns("Zone_Name").HeaderText = "Zone Name"


            End If
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Farmer Collection Details") = CompairStringResult.Equal Then
                gv.Columns("Union").HeaderText = "Union Name"
                gv.Columns("TotalDCS").HeaderText = "Total DCS"
                gv.Columns("MorningDCSCount").HeaderText = "DCS Count"
                gv.Columns("MorningMPCount").HeaderText = "Farmer Count"
                gv.Columns("MorningQty").HeaderText = "Quantity"
                gv.Columns("MorningFATKG").HeaderText = "FAT KG"
                gv.Columns("MorningSNFKG").HeaderText = "SNF KG"
                gv.Columns("MorningAmount").HeaderText = "Amount"
                gv.Columns("EveningDCSCount").HeaderText = "DCS Count"
                gv.Columns("EveningMPCount").HeaderText = "Farmer Count"
                gv.Columns("EveningQty").HeaderText = "Quantity"
                gv.Columns("EveningFATKG").HeaderText = "FAT KG"
                gv.Columns("EveningSNFKG").HeaderText = "SNF KG"
                gv.Columns("EveningAmount").HeaderText = "Amount"
                gv.Columns("TotalQuantity").HeaderText = "Total Quantity"
                gv.Columns("TotalFATKG").HeaderText = "Total FAT KG"
                gv.Columns("TotalSNFKG").HeaderText = "Total SNF KG"
                gv.Columns("TotalSNFKG").HeaderText = "Total SNF KG"
                gv.Columns("ZONE_CODE").HeaderText = "Zone Code"

                gv.Columns("ZONE_NAME").HeaderText = "Zone Name"

            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub View()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Collection at DCS") = CompairStringResult.Equal Then
                If gv.Rows.Count > 0 Then
                    Dim view As New ColumnGroupsViewDefinition()
                    view.ColumnGroups.Add(New GridViewColumnGroup(""))
                    view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("S.No.").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Union").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("DCS").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Zone_Name").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("ZONE_CODE").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Morning"))
                    view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningManual").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningAuto").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningPending").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Evening"))
                    view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningManual").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningAuto").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningPending").Name)
                    gv.ViewDefinition = view
                End If
            End If
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Milk Collection") = CompairStringResult.Equal Then
                If gv.Rows.Count > 0 Then
                    Dim view As New ColumnGroupsViewDefinition()
                    view.ColumnGroups.Add(New GridViewColumnGroup(""))
                    view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("S.No.").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Union").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Zone_Code").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Zone_Name").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Morning"))
                    view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningQty").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningFat").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningSnf").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningAmount").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Evening"))
                    view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningQty").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningFat").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningSnf").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningAmount").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup(""))
                    view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv.Columns("TotalQuantity").Name)
                    gv.ViewDefinition = view
                End If
            End If
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Farmer Collection Details") = CompairStringResult.Equal Then
                If gv.Rows.Count > 0 Then
                    Dim view As New ColumnGroupsViewDefinition()
                    view.ColumnGroups.Add(New GridViewColumnGroup(""))
                    view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("S.No.").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Union").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("Date").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("TotalDCS").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("ZONE_CODE").Name)
                    view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv.Columns("ZONE_NAME").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Morning Entry Status"))
                    view.ColumnGroups(1).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningDCSCount").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningMPCount").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningQty").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningFATKG").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningSNFKG").Name)
                    view.ColumnGroups(1).Rows(0).ColumnNames.Add(gv.Columns("MorningAmount").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Evening Entry Status"))
                    view.ColumnGroups(2).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningDCSCount").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningMPCount").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningQty").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningFATKG").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningSNFKG").Name)
                    view.ColumnGroups(2).Rows(0).ColumnNames.Add(gv.Columns("EveningAmount").Name)

                    view.ColumnGroups.Add(New GridViewColumnGroup("Grand Total"))
                    view.ColumnGroups(3).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv.Columns("TotalQuantity").Name)
                    view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv.Columns("TotalFATKG").Name)
                    view.ColumnGroups(3).Rows(0).ColumnNames.Add(gv.Columns("TotalSNFKG").Name)
                    gv.ViewDefinition = view
                End If
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Private Sub FarmerCollectionDetails()
        Try
            Dim qry As String = " Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,[Union],Convert(Varchar(10),Doc_Date,103) As [Date],
 Vlc_Code_VLC_Uploader As [DCS],VLC_Name As [DCS Name],MP_Uploader_Code As [Farmer Code],MP_Name As [Farmer Name],shift As [Shift],qty As [Qty],fat As [Fat %],snf As [Snf %],Convert(Decimal(18,2),Case When qty>0 Then Amount/qty Else 0 End) As [Rate],Amount,tttype As [Source]  ,ZONE_CODE as [Zone Code],Description as [Zone Name] "
            If isPrint Then
                qry &= " ,'" & objCommonVar.CurrentUser & "' As PrintBy,TSPL_COMPANY_MASTER.Logo_Img, TSPL_COMPANY_MASTER.Logo_Img2, TSPL_COMPANY_MASTER.Comp_Name, TSPL_COMPANY_MASTER.Add1, TSPL_COMPANY_MASTER.Add2, TSPL_COMPANY_MASTER.Add3, TSPL_COMPANY_MASTER.State, TSPL_STATE_MASTER.STATE_NAME"
            End If
            qry &= " from("
            qry &= ReturnFarmerBaseQry()
            qry &= ")final"
            If isPrint Then
                qry &= " Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'
Left Outer Join TSPL_STATE_MASTER On TSPL_STATE_MASTER.STATE_CODE=TSPL_COMPANY_MASTER.State"
            End If
            dt = Nothing
            dt = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If isPrint Then
                    Dim frm As New frmCrystalReportViewer()
                    frm.funreport(Me.Form_ID, CrystalReportFolder.CommonForUnionAndCattlefeed, dt, "crptFarmerCollectionReportForUnion", "Farmer Collection Details")
                    frm = Nothing
                Else
                    gv.DataSource = Nothing
                    gv.Rows.Clear()
                    gv.Columns.Clear()
                    gv.DataSource = dt
                    gv.GroupDescriptors.Clear()
                    gv.MasterTemplate.SummaryRowsBottom.Clear()
                    gv.TableElement.TableHeaderHeight = 25
                    gv.MasterTemplate.ShowRowHeaderColumn = True
                    For ii As Integer = 0 To gv.Columns.Count - 1
                        gv.Columns(ii).ReadOnly = True
                        gv.Columns(ii).IsVisible = True
                    Next
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.BestFitColumns()
                    RadPageView1.SelectedPage = RadPageViewPage2
                    ReStoreGridLayout()
                    EnableDisableControl(False)
                    Dim summaryRowItem As New GridViewSummaryRowItem()
                    Dim item1 As New GridViewSummaryItem("Qty", "{0:N2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(item1)

                    Dim item2 As New GridViewSummaryItem("Amount", "{0:N2}", GridAggregateFunction.Sum)
                    summaryRowItem.Add(item2)
                    gv.ShowGroupPanel = False
                    gv.MasterTemplate.AutoExpandGroups = True
                    gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
                    gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Data not found !", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function GetQryUploader() As String
        Dim i As Integer = 0
        Dim QryUploader As String = ""
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each drUnion As DataRow In dt.Rows
                Dim dcsCount As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Count(Distinct TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code) from " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_DETAIL
Left Join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_HEAD On TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
Where TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtToDate.Value.AddMonths(-6), "dd/MMM/yyyy") & "' And TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'"))
                If dcsCount <= 0 Then
                    dcsCount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select COUNT(1) from " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_MASTER_HEAD where Active=1"))
                End If
                If i <> 0 Then
                    QryUploader &= " Union All "
                End If
                QryUploader &= " select  '" & clsCommon.myCstr(drUnion("Location_Name")) & "' As [Union],'" & clsCommon.myCstr(dcsCount) & "' As [DCSCount],'" & clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, Nothing)) & "'  as item_code, TSPL_VLC_DATA_UPLOADER.Doc_No , TSPL_Mp_MASTER.Mcc_Code_VLC_Uploader,TSPL_VLC_DATA_UPLOADER.MCC_Code,TSPL_VLC_DATA_UPLOADER.VLC_CODE as VLC_uploader_CODE,TSPL_Mp_MASTER.MCC_NAME ,TSPL_MP_MASTER.VLC_Code ,TSPL_MP_MASTER.VLC_Name,TSPL_Mp_MASTER.MP_Code ,TSPL_VLC_DATA_UPLOADER.MP_CODE   as MP_Code_uploader ,TSPL_MP_MASTER.MP_Name ,TSPL_MP_MASTER.AccountNO ,TSPL_MP_MASTER.BankBranch ,TSPL_MP_MASTER. BankName,TSPL_VLC_DATA_UPLOADER.Doc_Date ,TSPL_VLC_DATA_UPLOADER.shift,TSPL_VLC_DATA_UPLOADER.qty as qty,TSPL_VLC_DATA_UPLOADER.fat, TSPL_VLC_DATA_UPLOADER.snf ,TSPL_VLC_DATA_UPLOADER.Amount ,TSPL_Mp_Master.UOM_Code,TSPL_MP_MASTER.Route_Code,RT.Route_Name,TSPL_MP_MASTER.TOLERANCE,TSPL_VLC_DATA_UPLOADER.Entry_Source as tttype   from  (select TSPL_MP_MASTER.TOLERANCE,TSPL_Mp_MASTER.MP_Code ," & Environment.NewLine &
            " TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader ,VLC_Code_VLC_Uploader,MP_Code_VLC_Uploader,Mp_Name,TSPL_MP_MASTER.AccountNO,TSPL_MP_MASTER.BankBranch,TSPL_MP_MASTER.BankName,TSPL_VLC_MASTER_HEAD.MCC, MCC_NAME,UOM_Code,TSPL_VLC_MASTER_HEAD.VLC_Code ,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.Route_Code  " & Environment.NewLine &
            " from " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_MASTER_HEAD   left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code =TSPL_VLC_MASTER_HEAD.MCC " & Environment.NewLine &
            " Left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_Mcc_UOM_DETAIL on TSPL_Mcc_UOM_DETAIL.MCC_CODE =TSPL_MCC_MASTER.MCC_Code and Stocking_Unit ='Y' Left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MP_MASTER on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_MASTER.VLC_Code  left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_BANK_BRANCH_MASTER on TSPL_BANK_BRANCH_MASTER.BRANCH_CODE =TSPL_MP_MASTER.BankBranch  left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE =TSPL_MP_MASTER.BankName) TSPL_MP_MASTER left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_DATA_UPLOADER " & Environment.NewLine &
            " on TSPL_MP_MASTER.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPL_MP_MASTER.VLC_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE and TSPL_MP_MASTER.MCC =TSPL_VLC_DATA_UPLOADER.MCC_Code  " & Environment.NewLine &
            " left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MCC_ROUTE_MASTER RT on TSPL_MP_MASTER.Route_Code=RT.Route_Code where 2=2"

                If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                    QryUploader += " and 2=( case when File_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and File_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='M' then 3 else 2 end  )"
                End If
                If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                    QryUploader += " and 2=( case when File_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and File_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and shift='E' then 3 else 2 end  )"
                End If
                If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                    QryUploader += " and TSPL_VLC_DATA_UPLOADER.MCC_Code  IN (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ") "
                End If
                If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                    QryUploader += " and TSPL_MP_MASTER.VLC_CODE  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
                End If
                If Not txtRoute.arrValueMember Is Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
                    QryUploader += " and TSPL_MP_MASTER.Route_Code  IN (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ") "
                End If
                QryUploader += "  and convert(date,File_Date,103)>=convert(date,('" + txtFromDate.Value + "'),103) and convert(date,File_Date,103) <=convert(date,('" + txtToDate.Value + "'),103) "

                If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(cmbEntrySource.Text, "All") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(cmbEntrySource.Text, "Manual") <> CompairStringResult.Equal Then
                        QryUploader += " and TSPL_VLC_DATA_UPLOADER.Entry_Source='" + clsCommon.myCstr(cmbEntrySource.Text) + "'"
                    End If
                End If
                i += 1
            Next
        End If
        Return QryUploader
    End Function

    Public Function GetQryManual() As String
        Dim i As Integer = 0
        Dim QryManual As String = ""
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For Each drUnion As DataRow In dt.Rows
                Dim dcsCount As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Count(Distinct TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code) from " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_DETAIL
Left Join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_MP_INCENTIVE_ENTRY_HEAD On TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
Where TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date>='" & clsCommon.GetPrintDate(txtToDate.Value.AddMonths(-6), "dd/MMM/yyyy") & "' And TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'"))
                If dcsCount <= 0 Then
                    dcsCount = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" Select COUNT(1) from " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_MASTER_HEAD where Active=1"))
                End If
                If i <> 0 Then
                    QryManual &= " Union All "
                End If
                QryManual &= "select  '" & clsCommon.myCstr(drUnion("Location_Name")) & "' As [Union],'" & clsCommon.myCstr(dcsCount) & "' As [DCSCount],'" & clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, Nothing)) & "'  as item_code,VDUM.document_code as Doc_No,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader, VLCM.MCC as MCC_Code,VLCM.vlc_code_vlc_uploader AS VLC_uploader_CODE," & Environment.NewLine &
        " tspl_mcc_master.mcc_name,VDUM.VLC_Code,VLCM.vlc_Name as VLC_Name,TSPL_MP_MASTER.mp_code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as MP_Code_uploader,TSPL_MP_MASTER.mp_name,TSPL_MP_MASTER.AccountNO," & Environment.NewLine &
        " TSPL_MP_MASTER.BankBranch,TSPL_MP_MASTER.BankName,VDUM.Document_Date as Doc_Date,(case when VDUM.Shift='MORNING' THEN 'M' ELSE 'E' END) AS Shift," & Environment.NewLine &
        " VDUD.Qty,VDUD.fatper as fat,VDUD.snfper as snf,VDUD.Amount as Amount,VDUD.Unit_Code as UOM_Code,VLCM.Route_Code, " & Environment.NewLine &
        "  tspl_mcc_route_master.route_name ,TSPL_MP_MASTER.TOLERANCE,'Manual' as tttype from " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_DATA_UPLOADER_DETAIL VDUD  inner join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_DATA_UPLOADER_MASTER VDUM on VDUD.Document_Code=VDUM.Document_Code  " & Environment.NewLine &
        " left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.TSPL_VLC_MASTER_HEAD VLCM on VDUM.VLC_Code=VLCM.VLC_Code  left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.tspl_mcc_master on tspl_mcc_master.mcc_code=VLCM.MCC " & Environment.NewLine &
        " left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.tspl_mp_master on tspl_mp_master.mp_code=VDUD.farmer_code left join " & clsCommon.myCstr(drUnion("Database_Name")) & ".dbo.tspl_mcc_route_master on tspl_mcc_route_master.Route_Code=VLCM.Route_Code " & Environment.NewLine &
        " where 2 = 2  and  convert(date, VDUM.Document_Date,103) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and  convert(date, VDUM.Document_Date,103) <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "'"

                If clsCommon.CompairString(txtFromShift.Text, "E") = CompairStringResult.Equal Then
                    QryManual += " and 2=( case when convert(date, VDUM.Document_Date,103) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and convert(date, VDUM.Document_Date,103) <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtFromDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and SHIFT='MORNING' then 3 else 2 end  )"
                End If
                If clsCommon.CompairString(txtToShift.Text, "M") = CompairStringResult.Equal Then
                    QryManual += " and 2=( case when convert(date, VDUM.Document_Date,103) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and convert(date, VDUM.Document_Date,103) <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(txtToDate.Value), "dd/MMM/yyyy hh:mm tt") + "' and SHIFT='EVENING' then 3 else 2 end  )"
                End If
                If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                    QryManual += " and VLCM.MCC  IN (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ") "
                End If
                If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                    QryManual += " and VLCM.VLC_Code IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
                End If
                If Not txtRoute.arrValueMember Is Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
                    QryManual += " and VLCM.Route_Code  IN (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ") "
                End If
                i += 1
            Next
        End If
        Return QryManual
    End Function


    Private Function GetBaseQuery()
        Dim qry As String = ""
        Dim companyADD, CompName, CompCode As String
        Try
            qry = " select   TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end + case when LEN(TSPL_COMPANY_MASTER.City_Code)>0 then ', '+TSPL_COMPANY_MASTER.City_Code else ' ' end + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end  as comp_address from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
            companyADD = dt1.Rows(0).Item("comp_address")

            qry = " select   TSPL_COMPANY_MASTER.Comp_Name  from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
            CompName = dt2.Rows(0).Item("Comp_Name")


            qry = " select   TSPL_COMPANY_MASTER.comp_code  from TSPL_COMPANY_MASTER where  Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "
            Dim dt5 As DataTable = clsDBFuncationality.GetDataTable(qry)
            CompCode = dt5.Rows(0).Item("Comp_Code")
            Dim Shh As String = String.Empty
            Dim fromDate As Date = clsCommon.GetPrintDate(txtFromDate.Value)

            Dim Todate As Date = clsCommon.GetPrintDate(txtToDate.Value)
            Dim DiffDate As Integer = clsCommon.myCDate(Todate).Day - clsCommon.myCDate(fromDate).Day
            If DiffDate > 0 Then
                Shh = "Both"
            ElseIf DiffDate = 0 And txtFromShift.SelectedValue <> txtToShift.SelectedValue Then
                Shh = "Both"
            Else
                Shh = IIf(txtFromShift.SelectedValue = "M", "Morning", "Evening")
            End If
            ''richa ERO/14/12/18-000443 18 Dec,2018

            Dim QryUploader As String = GetQryUploader()
            Dim QryManual As String = GetQryManual()


            ' Change By Prabhakar (Live issue in EROD) : replace "<" to "<="  for [Cow FAT(%)] and [Cow SNF(%)]    
            qry = "with CTEMP AS  (" & Environment.NewLine &
            "select ROW_NUMBER() OVER(ORDER BY Doc_Date) as SNo,'" + txtFromShift.Text + "' as ShiftMor,'" + txtToShift.Text + "' as ShiftEve,'" + cboUnit.Text + "' as Unit,'" & Shh & "'  as HShift ,'" & fromDate & "'  as fromDate ,'" & Todate & "'  as Todate ,'" & companyADD & "'  as companyADD, '" & CompName & "'  as CompName,'" & CompCode & "'  as CompCode," & Environment.NewLine &
            " (Mcc_Code_VLC_Uploader) as [MCC Uploader], MCC_Code as [MCC Code] ,(MCC_NAME) as [MCC Name],(VLC_uploader_CODE) as [VLC Uploader],VLC_CODE as [VLC Code],(Route_Code) as [Route Code],(Route_Name) as [Route Name],(VLC_Name )as [VLC Name],MP_Code_uploader as [MP Uploader],MP_Code as [MP Code] ,(MP_Name) as [MP Name]," & Environment.NewLine &
            " (AccountNO) as [Account NO],(BankBranch) as [Bank Branch],(BankName) as [Bank Name] , shift,TOLERANCE,tttype, " & Environment.NewLine &
            " isnull([Buffalo Milk Qty (KG)],0) as [Buffalo Milk Qty (KG)] ," & Environment.NewLine &
            " isnull( [Buffalo FAT(%)],0) as [Buffalo FAT(%)],  " & Environment.NewLine &
            "  isnull([Buffalo SNF(%)],0) as [Buffalo SNF(%)] , " & Environment.NewLine &
            " case when isnull(TOLERANCE,0 )<>0 then  cast((((((sum([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end))*TOLERANCE )/100)+(((sum([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Buffalo FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end)) as decimal(18,2)) else 0 end as [Buff_fat%_avg], " & Environment.NewLine &
            " case when isnull(TOLERANCE,0 )<>0 then  cast((((((sum([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end))*TOLERANCE )/100)+(((sum([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Buffalo SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end)) as decimal(18,2)) else 0 end as [Buff_SNF%_avg], " & Environment.NewLine &
            "  isnull([Buffalo FAT (KG)],0) as [Buffalo FAT (KG)]," & Environment.NewLine &
            " isnull ([Buffalo SNF (KG)],0) as [Buffalo SNF (KG)] ," & Environment.NewLine &
            "  isnull([Buffalo Amount],0 ) as [Buffalo Amount], " & Environment.NewLine &
            " isnull([Cow Milk Qty (KG)],0) as [Cow Milk Qty (KG)], " & Environment.NewLine &
            " isnull(  [Cow FAT(%)],0) as [Cow FAT(%)], " & Environment.NewLine &
            " isnull( [Cow SNF(%)],0) as [Cow SNF(%)] , " & Environment.NewLine &
            " case when isnull(TOLERANCE,0 )<>0 then  cast((((((sum([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end))*TOLERANCE )/100)+(((sum([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Cow FAT(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end)) as decimal(18,2)) else 0 end as [Cow_fat%_avg], " & Environment.NewLine &
            "  case when isnull(TOLERANCE,0 )<>0 then  cast((((((sum([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end))*TOLERANCE )/100)+(((sum([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )) -(last_value([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding )))/(case when (count([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))=1 then 1 else (count([Cow SNF(%)]) over (partition by mp_code order by Doc_Date asc rows  3 preceding ))-1 end)) as decimal(18,2)) else 0 end as [Cow_SNF%_avg], " & Environment.NewLine &
            "  isnull([cow FAT (KG)],0) as [Cow FAT (KG)]," & Environment.NewLine &
            " isnull([cow SNF (KG)],0) as [Cow SNF (KG)] , " & Environment.NewLine &
            " isnull([Cow Amount],0) as [Cow Amount],Doc_No,Doc_date,isnull([Cow Amount],0)+isnull([Buffalo Amount],0 ) as Total " & Environment.NewLine &
            "from (" & Environment.NewLine &
            "select qq.*,Case When qq.FAT <= 5 Then  qq.qty    Else 0 End [Cow Milk Qty (KG)] ,Case When qq.FAT <= 5 Then qq.FAT Else 0 End [Cow FAT(%)], " & Environment.NewLine &
            " Case When qq.FAT <= 5 Then qq.SNF Else 0 End [Cow SNF(%)],Case When qq.FAT  <= 5 Then qq.NewFAT_KG  Else 0 End [Cow FAT (KG)]," & Environment.NewLine &
            "   Case When qq.FAT <= 5 Then qq.NewSNF_KG Else 0 End [Cow SNF (KG)], Case When qq.FAT <= 5 Then qq.Amount Else 0 End [Cow Amount]," & Environment.NewLine &
            "  Case When qq.FAT > 5 Then qq.qty  Else 0 End [Buffalo Milk Qty (KG)],Case When qq.FAT > 5 Then qq.FAT Else 0 End [Buffalo FAT(%)],  " & Environment.NewLine &
            " Case When qq.FAT > 5 Then qq.SNF Else 0 End [Buffalo SNF(%)], Case When qq.FAT > 5 Then qq.NewFAT_KG Else 0 End [Buffalo FAT (KG)],  " & Environment.NewLine &
            " Case When qq.FAT > 5 Then qq.NewSNF_KG Else 0 End [Buffalo SNF (KG)],Case When qq.FAT > 5 Then qq.Amount  Else 0 End [Buffalo Amount] from ( " & Environment.NewLine &
            " select pp.tttype,pp.TOLERANCE,pp.doc_no, pp.Mcc_Code_VLC_Uploader ,pp.VLC_uploader_CODE ,pp.MCC_Code,pp.VLC_CODE,pp.VLC_Name,pp.MCC_NAME, pp.MP_CODE   as MP_Code,pp.MP_Code_uploader  ," & Environment.NewLine &
            " pp.MP_Name ,pp.AccountNO ,pp.BankBranch ,pp. BankName,pp.Doc_Date ,pp.shift,pp.NewQty  as qty,pp.fat, pp.snf,pp.NewFAT_KG ,pp.NewSNF_KG ,pp.NewAmount as Amount ,pp.UOM_Code,Route_Code,Route_Name  " & Environment.NewLine &
            "  from ( select xx.*,xx.qty*CF as NewQty,xx.Amount as NewAmount,fat_KG *Cf as NewFAT_KG,snf_KG *cf as NewSNF_KG  from (" & Environment.NewLine &
            " select *,(qty *fat) /100 as fat_KG,(qty *snf) /100 as snf_KG from ("
            '& Environment.NewLine &
            '" " & QryUploader & " "

            If Not (clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal) OrElse ((clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal) AndAlso clsCommon.CompairString(cmbEntrySource.Text, "All") = CompairStringResult.Equal) Then
                qry += QryUploader & Environment.NewLine & " union all " & Environment.NewLine &
                " " & QryManual & " "
            ElseIf ((clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal) AndAlso clsCommon.CompairString(cmbEntrySource.Text, "Manual") = CompairStringResult.Equal) Then
                qry += QryManual & Environment.NewLine & ""
            Else
                qry += QryUploader & Environment.NewLine & ""
            End If

            qry += " ) as pp )xx" & Environment.NewLine &
              " LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code =xx.Item_Code "
            ''richa agarwal 28 May,2019  TEC/28/03/19-000462 add item structure on setting based
            If ItemStructureMandatoryOnWeightConversion = True Then
                qry += " left outer join (Select Distinct yyy.* From (  Select Container_UOM as FromUOM, Contained_UOM as TOUOM, Container_Qty*Contained_Qty as CF,Structure_code  from TSPL_WEIGHT_CONVERSION where Contained_UOM='KG' UNION All Select Contained_UOM as FromUOM, Container_UOM as TOUOM, Container_Qty/ nullif (Contained_Qty,0) as CF,Structure_code from TSPL_WEIGHT_CONVERSION where Container_UOM='LTR' UNION All   Select Contained_UOM as FromUOM, Contained_UOM as TOUOM, 1 as CF,Structure_code  from TSPL_WEIGHT_CONVERSION UNION All Select Container_UOM as FromUOM, Container_UOM as TOUOM, 1 as CF,Structure_code  from TSPL_WEIGHT_CONVERSION  ) yyy) zzz on zzz.FromUOM =xx.UOM_Code   and lower(zzz.TOUOM)='" + cboUnit.Text + "' where 2 = 2 AND TSPL_ITEM_MASTER.Structure_Code =zzz.Structure_code ) "
            Else
                qry += " left outer join (Select Distinct yyy.* From (  Select Container_UOM as FromUOM, Contained_UOM as TOUOM, Container_Qty*Contained_Qty as CF from TSPL_WEIGHT_CONVERSION where Contained_UOM='KG' UNION All Select Contained_UOM as FromUOM, Container_UOM as TOUOM, Container_Qty/ nullif (Contained_Qty,0) as CF from TSPL_WEIGHT_CONVERSION where Container_UOM='LTR' UNION All   Select Contained_UOM as FromUOM, Contained_UOM as TOUOM, 1 as CF from TSPL_WEIGHT_CONVERSION UNION All Select Container_UOM as FromUOM, Container_UOM as TOUOM, 1 as CF from TSPL_WEIGHT_CONVERSION  ) yyy) zzz on zzz.FromUOM =xx.UOM_Code and lower(zzz.TOUOM)='" + cboUnit.Text + "' ) "
            End If

            qry += " as pp  ) as qq   ) as final) "

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return qry
    End Function

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(PageSetupReport_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(PageSetupReport_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
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
            clsCommon.MyMessageBoxShow(Me, err.Message, Me.Text)
        End Try
    End Sub

    Sub print(ByVal exporter As EnumExportTo)
        Try
            If gv.Rows.Count > 0 Then

                Dim arrHeader As List(Of String) = New List(Of String)()
                Dim CompName As String = clsDBFuncationality.getSingleValue("Select Comp_Name from TSPL_COMPANY_MASTER Where Comp_Code='" + objCommonVar.CurrentCompanyCode + "'")
                arrHeader.Add(CompName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptMPWiseMilkCollectionATPoolingPoint & "'"))
                arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy")) + " ")
                If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                    arrHeader.Add(("MCC : " + clsCommon.GetMulcallStringWithComma(txtMCC.arrDispalyMember) + " "))
                End If
                If txtRoute.arrValueMember IsNot Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
                    arrHeader.Add(("Route : " + clsCommon.GetMulcallStringWithComma(txtRoute.arrDispalyMember) + " "))
                End If
                If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
                    arrHeader.Add(("VLC : " + clsCommon.GetMulcallStringWithComma(txtVLC.arrDispalyMember) + " "))
                End If

                If exporter = EnumExportTo.Excel Then
                    transportSql.applyExportTemplate(gv, PageSetupReport_ID)
                    transportSql.QuickExportToExcel(gv, "", Me.Text, , arrHeader)
                Else
                    transportSql.applyExportTemplate(gv, PageSetupReport_ID)
                    clsCommon.MyExportToPDF(Me.Text, gv, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
                End If
            Else
                common.clsCommon.MyMessageBoxShow(Me, "No Data Found to Export.", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub


    Private Sub rmSaveLayout_Click(sender As Object, e As EventArgs) Handles rmSaveLayout.Click
        'If Not (clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal) Then
        If clsCommon.myLen(PageSetupReport_ID) > 0 Then
            gv.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = PageSetupReport_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv.SaveLayout(obj.GridLayout)
            obj.GridColumns = gv.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
        'End If
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(PageSetupReport_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub



    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        btnReferesh = True
        PageSetupReport_ID = MyBase.Form_ID
        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "S"
            If chkSupervisor.Checked = True Then
                PageSetupReport_ID = PageSetupReport_ID + "WSUP"
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "MP Polling") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "P"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Detail") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "D"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Detail") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "F"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Collection at DCS") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "C"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Entry Status") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "E"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Milk Collection") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "M"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Farmer Collection Details") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "DF"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Summary") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "DS"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DCS Collection v/s Farmer Collection Day & Shift Wise") = CompairStringResult.Equal Then
            PageSetupReport_ID = PageSetupReport_ID + "DD"
        End If
        TemplateGridview = gv
        GetReportID()
        LoadData()
    End Sub

    Sub GetReportID()
        Dim VarID As String = ""
        If clsCommon.CompairString(cboReportType.SelectedItem.Value, "Detail") = CompairStringResult.Equal Then
            VarID += "_D"
        ElseIf clsCommon.CompairString(cboReportType.SelectedItem.Value, "Summary") = CompairStringResult.Equal Then
            VarID += "_S"
        ElseIf clsCommon.CompairString(cboReportType.SelectedItem.Value, "MP Polling") = CompairStringResult.Equal Then
            VarID += "_MP"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Detail") = CompairStringResult.Equal Then
            VarID += "_F"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Collection at DCS") = CompairStringResult.Equal Then
            VarID += "_C"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Entry Status") = CompairStringResult.Equal Then
            VarID += "_E"
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Entry Status") = CompairStringResult.Equal Then
            VarID += "_MC"
        End If
        gv.VarID = VarID
    End Sub

    Private Sub BtnReset_Click(sender As Object, e As EventArgs) Handles BtnReset.Click
        Reset()
    End Sub


    Private Sub btnprint_Click(sender As Object, e As EventArgs) Handles btnprint.Click
        btnReferesh = False
        isPrint = True
        LoadData()
        isPrint = False
    End Sub

    Private Sub btnclose_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub

    Private Sub RptMPWiseMilkCollectionAtPoolingPoint3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ItemStructureMandatoryOnWeightConversion = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.ItemStructureMandatoryOnWeightConversion, clsFixedParameterCode.ItemStructureMandatoryOnWeightConversion, Nothing)) = 1, True, False))
        LOCATIONRIGTHS()
        SetUserMgmtNew()
        DisplayAverageFatSNFMPWise = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.DisplayAverageFatSNFMPWise, clsFixedParameterCode.DisplayAverageFatSNFMPWise, Nothing)) = "1", True, False))
        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnGo, "Press Alt+R Refresh ")
        ButtonToolTip.SetToolTip(BtnReset, "Press Alt+N Adding New")
        RadPageView1.SelectedPage = RadPageViewPage1
        txtToDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        LoadReport()
        LoadShiftFrom()
        LoadShiftTo()
        LoadEntrySource()
        Reset()
    End Sub

    Private Sub RptMPWiseMilkCollectionAtPoolingPoint3_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            LoadData()
        ElseIf e.Alt And e.KeyCode = Keys.C Then
            Me.Close()
        ElseIf e.Alt And e.KeyCode = Keys.R Then
            Reset()
        End If
    End Sub


    Private Sub gv_RowFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.RowFormattingEventArgs)
        Try
            If Not (clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal) Then
                If DisplayAverageFatSNFMPWise = True Then
                    If clsCommon.CompairString(e.RowElement.RowInfo.Cells("ChangeColorYesNo").Value, "Y") = CompairStringResult.Equal Then
                        e.RowElement.DrawFill = True
                        e.RowElement.GradientStyle = GradientStyles.Solid
                        e.RowElement.ForeColor = Color.Black
                        e.RowElement.BackColor = Color.LightGreen
                    Else
                        e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
                        e.RowElement.ResetValue(LightVisualElement.GradientStyleProperty, ValueResetFlags.Local)
                        e.RowElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local)
                        e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
                    End If
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        print(EnumExportTo.Excel)
    End Sub

    Private Sub rmiPDF_Click(sender As Object, e As EventArgs) Handles rmiPDF.Click
        print(EnumExportTo.PDF)
    End Sub

    Private Sub chkSummary_CheckedChanged(sender As Object, e As EventArgs)
        Reset()
    End Sub



    Private Sub VLCWise()
        Try

            Dim qry As String
            qry = GetBaseQuery()

            qry += " select doc_date as [Date],[VLC Uploader],[VLC Name],[MCC Code],[MCC Name],count(1) as [Count] from CTEMP where  convert(date, doc_date,103)=convert(date, '" + clsCommon.myCstr(gv.CurrentRow.Cells("Date").Value) + "',103) group by doc_date,[VLC Uploader],[VLC Name],[MCC Code],[MCC Name] order by doc_date,[MCC Code],[VLC Uploader]  "
            dt = clsDBFuncationality.GetDataTable(qry)

            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gv.DataSource = dt
            gv.GroupDescriptors.Clear()
            gv.MasterTemplate.SummaryRowsBottom.Clear()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            txtReportType.Text = "VLCWise"
            gv.TableElement.TableHeaderHeight = 25
            gv.MasterTemplate.ShowRowHeaderColumn = True
            For ii As Integer = 0 To gv.Columns.Count - 1
                gv.Columns(ii).ReadOnly = True
                gv.Columns(ii).IsVisible = True
            Next
            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item14 As New GridViewSummaryItem("Count", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item14)
            gv.ShowGroupPanel = False
            gv.MasterTemplate.AutoExpandGroups = True
            gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            gv.BestFitColumns()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub ShiftWise()
        Try

            Dim qry As String
            qry = GetBaseQuery()

            qry += " select ROW_NUMBER() OVER(ORDER BY [VLC Uploader]) as SoNo,CTEMP.*,0.00 as Buffalo_Fat_avg_per_tol,0.00 as Buffalo_SNF_avg_per_tol,0.00 AS Cow_Fat_avg_per_tol,0.00 AS Cow_SNF_avg_per_tol, '' as ChangeColorYesNo from CTEMP where convert(date, doc_date,103)=convert(date, '" + clsCommon.myCstr(gv.CurrentRow.Cells("Date").Value) + "',103) and [VLC Uploader]='" + clsCommon.myCstr(gv.CurrentRow.Cells("VLC Uploader").Value) + "' and [MCC Code]='" + clsCommon.myCstr(gv.CurrentRow.Cells("MCC Code").Value) + "' order by CTEMP.doc_date,CTEMP.[MCC Code],CTEMP.[VLC Uploader]  "
            dt = clsDBFuncationality.GetDataTable(qry)

            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gv.DataSource = dt
            gv.GroupDescriptors.Clear()
            gv.MasterTemplate.SummaryRowsBottom.Clear()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            txtReportType.Text = "ShiftWise"
            FormatGrid()
            gv.Columns("SNo").IsVisible = False
            gv.BestFitColumns()

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv_CellDoubleClick(sender As Object, e As GridViewCellEventArgs)
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal Then
                If gv.CurrentRow.Index >= 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(txtReportType.Text), "DateWise") = CompairStringResult.Equal Then
                        VLCWise()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(txtReportType.Text), "VLCWise") = CompairStringResult.Equal Then
                        ShiftWise()
                    End If
                End If
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "MP Polling") = CompairStringResult.Equal Then
                If gv.CurrentRow.Index >= 0 Then
                    If clsCommon.CompairString(clsCommon.myCstr(txtReportType.Text), "MP Polling Count") = CompairStringResult.Equal Then
                        MPDetail()
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(txtReportType.Text), "MP Detail") = CompairStringResult.Equal Then
                        Detail()
                    End If
                End If
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub txtMCC__My_Click(sender As Object, e As EventArgs) Handles txtMCC._My_Click
        Dim qry As String = "select MCC_Code as [MCC Code],MCC_NAME as [MCC Name],TSPL_MCC_MASTER.plant_code as [Plant Code],tspl_location_master.location_desc as [Plant Name] from TSPL_MCC_MASTER left join tspl_location_master on tspl_location_master.location_code=TSPL_MCC_MASTER.plant_code"
        txtMCC.arrValueMember = clsCommon.ShowMultipleSelectForm("Pooling@M", qry, "MCC Code", "MCC Name", txtMCC.arrValueMember, txtMCC.arrDispalyMember)
        txtVLC.arrValueMember = Nothing

    End Sub
    Private Sub txtRoute__My_Click(sender As Object, e As EventArgs) Handles txtRoute._My_Click
        Dim qry As String = " SELECT  DISTINCT RT.Route_Code as Code,RT.Route_Name as Name FROM TSPL_MCC_ROUTE_MASTER RT LEFT JOIN TSPL_VLC_MASTER_HEAD VLC ON RT.Route_Code=VLC.Route_Code where 2=2 "
        If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
            qry += " and VLC.VLC_Code  IN (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
        End If
        txtRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("Pooling@R", qry, "Code", "Name", txtRoute.arrValueMember, txtRoute.arrDispalyMember)
    End Sub
    Private Sub txtVLC__My_Click(sender As Object, e As EventArgs) Handles txtVLC._My_Click
        Dim qry As String = "Select VLC.VLC_Code_vlc_Uploader as [Code],VLC.VLC_Code as [DCS Code],VLC.VLC_Name as [DCS Name],VLC.MCC as [MCC Code],VLC.Route_Code as [Route Code],RM.Route_Name " &
                  " from TSPL_VLC_MASTER_HEAD    VLC left join TSPL_MCC_ROUTE_MASTER RM on vlc.Route_Code=RM.Route_Code where 2=2 "
        If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
            qry += " and VLC.MCC in (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ")"
        End If
        txtVLC.arrValueMember = clsCommon.ShowMultipleSelectForm("Pooling@V", qry, "DCS Code", "DCS Name", txtVLC.arrValueMember, txtVLC.arrDispalyMember)

    End Sub

    Private Sub cboReportType_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles cboReportType.Validating
        lblConvRate.Visible = False
        txtMinDays.Visible = False
        chkSupervisor.Visible = False
        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "MP Polling") = CompairStringResult.Equal Then
            lblConvRate.Visible = True
            txtMinDays.Visible = True
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Summary") = CompairStringResult.Equal Then
            chkSupervisor.Visible = True
        End If
    End Sub



    Private Sub MPPollingFarmerCount()
        Try
            Dim ts As TimeSpan = txtToDate.Value.Subtract(txtFromDate.Value)
            If ts.Days < txtMinDays.Value Then
                Throw New Exception("Min Days Should be less then Date Range Days")
            End If
            Dim qry As String
            qry = GetBaseQuery()

            qry += " select [MCC Code],max([MCC Name]) as [MCC Name],[VLC Code],max([VLC Uploader]) as [VLC Uploader],max([VLC Name]) as [VLC Name],count([MP Code]) as NoOFMP from (
select [MCC Code],max([MCC Name]) as [MCC Name],[VLC Code],max([VLC Uploader]) as [VLC Uploader],max([VLC Name]) as [VLC Name],[MP Code],max([MP Name]) as [MP Name],COUNT(Doc_Date) as NoOFDays from(
select [MCC Code],max([MCC Name]) as [MCC Name],[VLC Code],max([VLC Uploader]) as [VLC Uploader],max([VLC Name]) as [VLC Name],[MP Code],max([MP Name]) as [MP Name],Doc_Date from (
 select [MCC Code],[MCC Name],[VLC Code],[VLC Uploader],[VLC Name],[MP Code],[MP Name],[MP Uploader],convert(date,Doc_Date,103) as Doc_Date ,shift,Doc_No,Total from CTEMP
)xx group by [MCC Code],[VLC Code],[MP Code],Doc_Date
)xxx group by [MCC Code],[VLC Code],[MP Code] having count(Doc_Date)>" + clsCommon.myCstr(txtMinDays.Value) + "
)xxxx group by [MCC Code],[VLC Code]  "
            dt = clsDBFuncationality.GetDataTable(qry)

            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gv.DataSource = dt
            gv.GroupDescriptors.Clear()
            gv.MasterTemplate.SummaryRowsBottom.Clear()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            txtReportType.Text = "MP Polling Count"

            gv.TableElement.TableHeaderHeight = 25
            gv.MasterTemplate.ShowRowHeaderColumn = True
            For ii As Integer = 0 To gv.Columns.Count - 1
                gv.Columns(ii).ReadOnly = True
                gv.Columns(ii).IsVisible = True
            Next

            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item14 As New GridViewSummaryItem("NoOFMP", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item14)
            gv.ShowGroupPanel = False
            gv.MasterTemplate.AutoExpandGroups = True
            gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            gv.BestFitColumns()
            RadPageView1.SelectedPage = RadPageViewPage2
            EnableDisableControl(False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub MPDetail()
        Try
            Dim ts As TimeSpan = txtToDate.Value.Subtract(txtFromDate.Value)
            If ts.Days < txtMinDays.Value Then
                Throw New Exception("Min Days Should be less then Date Range Days")
            End If
            Dim qry As String
            qry = GetBaseQuery()

            qry += "select [MCC Code],max([MCC Name]) as [MCC Name],[VLC Code],max([VLC Uploader]) as [VLC Uploader],max([VLC Name]) as [VLC Name],max([MP Uploader]) as [MP Uploader],[MP Code],max([MP Name]) as [MP Name],COUNT(Doc_Date) as NoOFDays from(
select [MCC Code],max([MCC Name]) as [MCC Name],[VLC Code],max([VLC Uploader]) as [VLC Uploader],max([VLC Name]) as [VLC Name],[MP Code],max([MP Uploader]) as [MP Uploader],max([MP Name]) as [MP Name],Doc_Date from (
 select [MCC Code],[MCC Name],[VLC Code],[VLC Uploader],[VLC Name],[MP Code],[MP Name],[MP Uploader],convert(date,Doc_Date,103) as Doc_Date ,shift,Doc_No,Total from CTEMP
where [VLC Code]='" + clsCommon.myCstr(gv.CurrentRow.Cells("VLC Code").Value) + "'
)xx group by [MCC Code],[VLC Code],[MP Code],Doc_Date
)xxx group by [MCC Code],[VLC Code],[MP Code] having count(Doc_Date)>" + clsCommon.myCstr(txtMinDays.Value) + ""
            dt = clsDBFuncationality.GetDataTable(qry)

            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gv.DataSource = dt
            gv.GroupDescriptors.Clear()
            gv.MasterTemplate.SummaryRowsBottom.Clear()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If
            txtReportType.Text = "MP Detail"

            gv.TableElement.TableHeaderHeight = 25
            gv.MasterTemplate.ShowRowHeaderColumn = True
            For ii As Integer = 0 To gv.Columns.Count - 1
                gv.Columns(ii).ReadOnly = True
                gv.Columns(ii).IsVisible = True
            Next
            gv.Columns("MP Code").IsVisible = False
            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item14 As New GridViewSummaryItem("NoOFDays", "{0:F0}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item14)
            gv.ShowGroupPanel = False
            gv.MasterTemplate.AutoExpandGroups = True
            gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
            gv.BestFitColumns()
            RadPageView1.SelectedPage = RadPageViewPage2
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub Detail()
        Try
            Dim qry As String
            qry = GetBaseQuery()

            qry += " select CTEMP.*
                     ,TSPL_VLC_Supervisor_Tagging.supervisor_code as [Supervisor Code]
                     ,tspl_employee_master.Emp_Name as [Supervisor Name]
                     ,0.00 as Buffalo_Fat_avg_per_tol,0.00 as Buffalo_SNF_avg_per_tol,0.00 AS Cow_Fat_avg_per_tol,0.00 AS Cow_SNF_avg_per_tol, '' as ChangeColorYesNo
                     ,CTEMP.[Cow Milk Qty (KG)] +CTEMP.[Buffalo Milk Qty (KG)] as [Total Milk Qty (KG)]
                      ,CASE WHEN (CTEMP.[Cow Milk Qty (KG)] +CTEMP.[Buffalo Milk Qty (KG)])>0 THEN 
                      ((CTEMP.[Cow FAT (KG)]+CTEMP.[Buffalo FAT (KG)])*100)/(CTEMP.[Cow Milk Qty (KG)] +CTEMP.[Buffalo Milk Qty (KG)])
                       ELSE 0 END as [Total FAT(%)]
                    ,CASE WHEN (CTEMP.[Cow Milk Qty (KG)] +CTEMP.[Buffalo Milk Qty (KG)])>0 THEN 
                      ((CTEMP.[Cow SNF (KG)]+CTEMP.[Buffalo SNF (KG)])*100)/(CTEMP.[Cow Milk Qty (KG)] +CTEMP.[Buffalo Milk Qty (KG)])
                      ELSE 0 END as [Total SNF(%)]
                      ,CTEMP.[Cow FAT (KG)]+CTEMP.[Buffalo FAT (KG)] as [Total FAT (KG)]
                      ,CTEMP.[Cow SNF (KG)]+CTEMP.[Buffalo SNF (KG)] as [Total SNF (KG)]
                      ,CTEMP.[Cow Amount]+CTEMP.[Buffalo Amount] as [Total Amount]
                      from CTEMP 					  
                      left join TSPL_VLC_Supervisor_Tagging on TSPL_VLC_Supervisor_Tagging.vlc_code=CTEMP.[VLC Code]
					  and  TSPL_VLC_Supervisor_Tagging.mcc_code=CTEMP.[MCC Code]
					  left join tspl_employee_master on tspl_employee_master.emp_code=TSPL_VLC_Supervisor_Tagging.Supervisor_code "
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "MP Polling") = CompairStringResult.Equal Then
                qry += " where [MP Code]='" + clsCommon.myCstr(gv.CurrentRow.Cells("MP Code").Value) + "'"
            End If
            qry += "order by CTEMP .Doc_Date,CTEMP.SNo  "

            dt = clsDBFuncationality.GetDataTable(qry)
            If DisplayAverageFatSNFMPWise = True Then
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    dt.Rows(0)("Buffalo_Fat_avg_per_tol") = clsCommon.myCdbl(dt.Rows(0)("Buff_fat%_avg"))
                    dt.Rows(0)("Buffalo_SNF_avg_per_tol") = clsCommon.myCdbl(dt.Rows(0)("Buff_SNF%_avg"))
                    dt.Rows(0)("Cow_Fat_avg_per_tol") = clsCommon.myCdbl(dt.Rows(0)("Cow_fat%_avg"))
                    dt.Rows(0)("Cow_SNF_avg_per_tol") = clsCommon.myCdbl(dt.Rows(0)("Cow_SNF%_avg"))

                    For i As Integer = 1 To dt.Rows.Count - 1
                        If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(i)("MP Code")), clsCommon.myCstr(dt.Rows(i - 1)("MP Code"))) = CompairStringResult.Equal Then
                            dt.Rows(i)("Buffalo_Fat_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i - 1)("Buff_fat%_avg"))
                            dt.Rows(i)("Buffalo_SNF_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i - 1)("Buff_SNF%_avg"))
                            dt.Rows(i)("Cow_Fat_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i - 1)("Cow_fat%_avg"))
                            dt.Rows(i)("Cow_SNF_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i - 1)("Cow_SNF%_avg"))
                        Else
                            dt.Rows(i)("Buffalo_Fat_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i)("Buff_fat%_avg"))
                            dt.Rows(i)("Buffalo_SNF_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i)("Buff_SNF%_avg"))
                            dt.Rows(i)("Cow_Fat_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i)("Cow_fat%_avg"))
                            dt.Rows(i)("Cow_SNF_avg_per_tol") = clsCommon.myCdbl(dt.Rows(i)("Cow_SNF%_avg"))
                        End If
                        If clsCommon.myCdbl(dt.Rows(i)("TOLERANCE")) > 0 AndAlso ((clsCommon.myCdbl(dt.Rows(i)("Buffalo FAT(%)")) > clsCommon.myCdbl(dt.Rows(i)("Buffalo_Fat_avg_per_tol"))) Or (clsCommon.myCdbl(dt.Rows(i)("Buffalo SNF(%)")) > clsCommon.myCdbl(dt.Rows(i)("Buffalo_SNF_avg_per_tol"))) Or (clsCommon.myCdbl(dt.Rows(i)("Cow FAT(%)")) > clsCommon.myCdbl(dt.Rows(i)("Cow_Fat_avg_per_tol"))) Or (clsCommon.myCdbl(dt.Rows(i)("Cow SNF(%)")) > clsCommon.myCdbl(dt.Rows(i)("Cow_SNF_avg_per_tol")))) Then
                            dt.Rows(i)("ChangeColorYesNo") = "Y"
                        Else
                            dt.Rows(i)("ChangeColorYesNo") = "N"
                        End If
                    Next
                End If
            End If


            gv.DataSource = Nothing
            gv.Rows.Clear()
            gv.Columns.Clear()
            gv.DataSource = dt
            gv.GroupDescriptors.Clear()
            gv.MasterTemplate.SummaryRowsBottom.Clear()
            FormatGrid()
            If btnReferesh = False Then
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, clsERPFuncationality.CompanyAddresShowinHeader(), "crptMPWiseMilkCollectionAtPoolingPoint", "MP Wise Milk Collection Pooling Point", "")
                frmCRV = Nothing
            End If
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub
            End If

            gv.BestFitColumns()
            RadPageView1.SelectedPage = RadPageViewPage2
            ReStoreGridLayout()
            EnableDisableControl(False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkSupervisor_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkSupervisor.ToggleStateChanged
        If chkSupervisor.Checked = True Then
            cmbEntrySource.Text = "All"
            cmbEntrySource.Enabled = False
        Else
            cmbEntrySource.Enabled = True
        End If
    End Sub

    Private Sub cboReportType_SelectedValueChanged(sender As Object, e As EventArgs) Handles cboReportType.SelectedValueChanged
        CheckReportTypeforDateRange()


    End Sub

    Sub CheckReportTypeforDateRange()
        Try
            If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Collection at DCS") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Collection Entry Status") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "Farmer Wise Milk Collection") = CompairStringResult.Equal Then
                txtToDate.Value = txtFromDate.Value
                txtToDate.Enabled = False
                txtToShift.SelectedValue = "E"
                txtToShift.Enabled = False
            ElseIf clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DBT Reco  v/s Farmer Collection") = CompairStringResult.Equal Then

                'Selected date
                Dim selectedDate As DateTime = txtFromDate.Value

                'First day of month
                Dim slot1 As New DateTime(selectedDate.Year, selectedDate.Month, 1)

                'Last day of same month
                Dim slot2 As New DateTime(selectedDate.Year,
                                selectedDate.Month,
                                Date.DaysInMonth(selectedDate.Year,
                                                 selectedDate.Month))

                'Set values
                txtFromDate.Value = slot1
                txtToDate.Value = slot2

            Else
                txtToDate.Enabled = True
                txtToShift.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtFromDate_Validated(sender As Object, e As EventArgs) Handles txtFromDate.Validated
        CheckReportTypeforDateRange()
    End Sub

    Private Sub gv_CellFormatting(sender As Object, e As CellFormattingEventArgs)
        If e.CellElement.Value IsNot Nothing AndAlso IsNumeric(e.CellElement.Value) AndAlso clsCommon.CompairString(e.Column.Name, "S.No.") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(e.Column.Name, "DCS") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(e.Column.Name, "Farmer Code") <> CompairStringResult.Equal Then
            e.CellElement.Text = Convert.ToDecimal(e.CellElement.Value).ToString("N2", New CultureInfo("en-IN"))
        End If
    End Sub

    'Private Sub txtFromDate_ValueChanged(sender As Object, e As EventArgs) Handles txtFromDate.ValueChanged
    '    Try
    '        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DBT Reco  v/s Farmer Collection") = CompairStringResult.Equal Then

    '            Dim SM As Integer = txtFromDate.Value.Month
    '            Dim SY As Integer = txtFromDate.Value.Year

    '            Dim CD As New DateTime(SY, SM, 1)
    '            Slot1 = clsCommon.GetPrintDate(CD, "dd/MMM/yyyy")
    '            Month()
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    Sub Month()
        If clsCommon.myLen(txtFromDate.Value) > 0 Then
            Dim SM As Integer = txtFromDate.Value.Month
            Dim SY As Integer = txtFromDate.Value.Year

            Dim CD As New DateTime(SY, SM, 1)
            Slot2 = clsCommon.GetPrintDate(CD.AddMonths(1).AddDays(-1), "dd/MMM/yyyy")

            'txtTODate.Value = txtFromDate.Value.AddMonths(2)
            '' Month1 = clsCommon.GetPrintDate(txtFromDate.Value, "MM-yyyy")
            ''Month2 = clsCommon.GetPrintDate(txtFromDate.Value.AddMonths(1), "MM-yyyy")
            ''
        End If
    End Sub



    Private Sub txtFromDate_ValueChanged(sender As Object, e As EventArgs) Handles txtFromDate.ValueChanged
        'If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DBT Reco  v/s Farmer Collection") = CompairStringResult.Equal Then

        '    Dim SM As Integer = txtFromDate.Value.Month
        '    Dim SY As Integer = txtFromDate.Value.Year

        '    Dim CD As New DateTime(SY, SM, 1)
        '    Slot1 = clsCommon.GetPrintDate(CD, "dd/MMM/yyyy")
        '    Slot2 = clsCommon.GetPrintDate(CD.AddMonths(1).AddDays(-1), "dd/MMM/yyyy")
        '    txtToDate.Value = Slot2
        '    '   Month()
        'End If
    End Sub

    Private Sub TxtZone__My_Click(sender As Object, e As EventArgs) Handles TxtZone._My_Click

        Dim strQry As String = Nothing
        strQry = "select Zone_Code as Code ,Description as Name from TSPL_ZONE_MASTER where 1=1"
        'strQry += " and TSPL_ZONE_MASTER. Zone_Code in (Select TSPL_CUSTOMER_MASTER.Zone_Code from TSPL_CUSTOMER_MASTER where TSPL_CUSTOMER_MASTER.Cust_Code in (" + clsCommon.GetMulcallString(txtCustomer.arrValueMember) + ") )"
        TxtZone.arrValueMember = clsCommon.ShowMultipleSelectForm("ZoneMulSelp", strQry, "Code", "Name", TxtZone.arrValueMember, TxtZone.arrDispalyMember)
    End Sub

    'Private Sub txtToDate_ValueChanged(sender As Object, e As EventArgs) Handles txtToDate.ValueChanged
    '    Try
    '        If clsCommon.CompairString(clsCommon.myCstr(cboReportType.SelectedValue), "DBT Reco  v/s Farmer Collection") = CompairStringResult.Equal Then

    '            Dim SM As Integer = txtToDate.Value.Month
    '            Dim SY As Integer = txtToDate.Value.Year

    '            Dim CD As New DateTime(SY, SM, 1)
    '            Slot3 = clsCommon.GetPrintDate(CD, "dd/MMM/yyyy")
    '            Month1()
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '    End Try
    'End Sub
    'Sub Month1()
    '    If clsCommon.myLen(txtFromDate.Value) > 0 Then
    '        Dim SM As Integer = txtToDate.Value.Month
    '        Dim SY As Integer = txtToDate.Value.Year

    '        Dim CD As New DateTime(SY, SM, 1)
    '        Slot4 = clsCommon.GetPrintDate(CD.AddMonths(1).AddDays(-1), "dd/MMM/yyyy")

    '        ''txtTODate.Value = txtFromDate.Value.AddMonths(2)
    '        ' Month1 = clsCommon.GetPrintDate(txtFromDate.Value, "MM-yyyy")
    '        'Month2 = clsCommon.GetPrintDate(txtFromDate.Value.AddMonths(1), "MM-yyyy")
    '        ' Month3 = clsCommon.GetPrintDate(txtFromDate.Value.AddMonths(2), "MM-yyyy")

    '    End If
    'End Sub
End Class
