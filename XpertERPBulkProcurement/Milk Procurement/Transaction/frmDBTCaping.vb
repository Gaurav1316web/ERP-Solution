Imports System.Data.SqlClient
Imports common
Imports System.IO

Public Class frmDBTCaping
    Inherits FrmMainTranScreen

#Region "Variables"
    Public Const colPKID As String = "colPKID"
    Public Const colSlNo As String = "colSlNo"
    Public Const colCapingStatus As String = "colCapingStatus"
    Public Const colBMCCode As String = "colBMCCode"
    Public Const colBMCUploaderNo As String = "colBMCUploaderNo"
    Public Const colBMCName As String = "colBMCName"
    Public Const colVLCCode As String = "colVLCCode"
    Public Const colVLCUploaderCode As String = "colVLCUploaderCode"
    Public Const colVLCName As String = "colVLCName"
    Public Const colMPCode As String = "colMPCode"
    Public Const colMPUploaderCode As String = "colMPUploaderCode"
    Public Const colMPName As String = "colMPName"
    Public Const colQty As String = "colQty"
    Public Const colCapingQty As String = "colCapingQty"
    Public Const colCapingIncreaseBy As String = "colCapingIncreaseBy"
    Public Const colCapingIncreaseDate As String = "colCapingIncreaseDate"
    Public Const colCapingIncreaseRemarks As String = "colCapingIncreaseRemarks"
    Public Const colThirdPartyCode As String = "colThirdPartyCode"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim IsinsideLoadData As Boolean = False
    Dim qry As String
    Dim isCappingIncreaseScreen As Boolean = False
#End Region

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub FrmVLCDataUploaderManual_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If clsCommon.CompairString(MyBase.Form_ID, clsUserMgtCode.DBTCappingIncrease) = CompairStringResult.Equal Then
            isCappingIncreaseScreen = True
            RadLabel10.Visible = True '
        Else
            RadLabel10.Visible = False
        End If

        SetUserMgmtNew()
        Reset()

        ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Transaction")
        ButtonToolTip.SetToolTip(btndelete, "Press Alt+D Delete Transaction")
        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnReset, "Press Alt+N New Transaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        If clsCommon.myLen(Me.Tag) > 0 Then
            LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
        End If
        Me.Focus()
        txtdate.Focus()
    End Sub

    Private Sub FrmVLCDataUploaderManual_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            Reset()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag AndAlso btnsave.Enabled Then
            SaveData(False)
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btndelete.Enabled Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnclose.Enabled Then
            CloseForm()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnclose.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.SIR
            frm.strCode = clsFixedParameterCode.SIReversAndCreate
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                btnReverse.Visible = True
            End If
        End If
    End Sub

    Sub loadBlankGrid()
        gvItem.DataSource = Nothing
        gvItem.Rows.Clear()
        gvItem.Columns.Clear()

        Dim farmercode As New GridViewTextBoxColumn()

        Dim lineNo As New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "PKID."
        lineNo.Name = colPKID
        lineNo.IsVisible = False
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        lineNo = New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "SNo."
        lineNo.Name = colSlNo
        lineNo.Width = 60
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        Dim repoIsSurTax1 As New GridViewCheckBoxColumn()
        repoIsSurTax1.HeaderText = "Status"
        repoIsSurTax1.Name = colCapingStatus
        repoIsSurTax1.ReadOnly = True
        repoIsSurTax1.IsVisible = True
        repoIsSurTax1.Width = 50
        repoIsSurTax1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax1) '30

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "BMC Code"
        farmercode.Name = colBMCCode
        farmercode.ReadOnly = True
        farmercode.Width = 100
        farmercode.IsVisible = False
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        Dim farmername As New GridViewTextBoxColumn()
        farmername.FormatString = ""
        farmername.HeaderText = "BMC"
        farmername.Name = colBMCUploaderNo
        farmername.Width = 100
        farmername.ReadOnly = True
        farmername.IsVisible = True
        farmername.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmername)

        farmername = New GridViewTextBoxColumn()
        farmername.FormatString = ""
        farmername.HeaderText = "BMC Name"
        farmername.Name = colBMCName
        farmername.Width = 150
        farmername.ReadOnly = True
        farmername.IsVisible = True
        farmername.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmername)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "DCS Code"
        farmercode.Name = colVLCCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "DCS"
        farmercode.Name = colVLCUploaderCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 80
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "DCS Name"
        farmercode.Name = colVLCName
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Farmer Code"
        farmercode.Name = colMPCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Farmer"
        farmercode.Name = colMPUploaderCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 80
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Farmer Name"
        farmercode.Name = colMPName
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)


        Dim Qty As New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "DBT Qty"
        Qty.Name = colQty
        Qty.Width = 100
        Qty.ReadOnly = True
        Qty.FormatString = "{0:n2}"
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "Capping Qty"
        Qty.Name = colCapingQty
        Qty.Width = 100
        Qty.ReadOnly = True
        Qty.FormatString = "{0:n2}"
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Capping By"
        farmercode.Name = colCapingIncreaseBy
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Capping on"
        farmercode.Name = colCapingIncreaseDate
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Capping Remarks"
        farmercode.Name = colCapingIncreaseRemarks
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)


        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Third Party Code"
        farmercode.Name = colThirdPartyCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        gvItem.AllowAddNewRow = False
        gvItem.AllowDeleteRow = False
        gvItem.AllowRowReorder = False
        gvItem.ShowGroupPanel = False
        gvItem.EnableFiltering = True
        gvItem.ShowFilteringRow = True
        gvItem.EnableSorting = False
        gvItem.EnableGrouping = False
        gvItem.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvItem.GridBehavior = New MyBehavior()

        ReStoreGridLayout()

        gvItem.MasterTemplate.SummaryRowsBottom.Clear()
        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim Smitem As New GridViewSummaryItem(colQty, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colCapingQty, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        gvItem.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)

        gvItem.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            '  If rbtnNEFT.IsChecked Then
            If clsCommon.myLen(Me.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Me.Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gvItem.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gvItem.Columns.Count - 1 Step ii + 1
                        gvItem.Columns(ii).IsVisible = False
                        gvItem.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gvItem.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        If isCappingIncreaseScreen Then
            MyBase.isModifyFlag = False
            MyBase.isDeleteFlag = False
            MyBase.isPostFlag = False
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btndelete.Visible = MyBase.isDeleteFlag
        btnPost.Visible = MyBase.isPostFlag
    End Sub

    Sub Reset()
        loadBlankGrid()
        Dim dt As Date = clsCommon.GETSERVERDATE()

        txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dt) & "/" & DatePart(DateInterval.Year, dt)
        txtToDate.Value = txtFromDate.Value.AddMonths(1).AddDays(-1)


        txtDocumentNo.Value = ""
        txtdate.Value = dt
        txtReco.Value = ""
        txtDocumentNo.MyReadOnly = False
        btnsave.Text = "Save"
        btndelete.Enabled = False
        btnsave.Enabled = True
        btnPost.Enabled = False
        txtdate.Focus()
        EnableInputDataField()
        isNewEntry = True
        IsinsideLoadData = False
        lblPending.Status = ERPTransactionStatus.Pending


    End Sub

    Private Function AllowToSave(ByVal isByPost As Boolean) As Boolean
        Xtra.TransactionValidity(txtdate.Value)
        If AllowFutureDateTransaction(txtdate.Value, Nothing) = False Then
            txtdate.Focus()
            Return False
        End If

        Return True
    End Function

    Sub SaveData(ByVal isByPost As Boolean)
        Try
            If AllowToSave(isByPost) Then
                Dim obj As New clsDBTCaping()
                obj.Document_Code = txtDocumentNo.Value
                obj.Document_Date = txtdate.Value
                obj.Reco_Code = txtReco.Value
                SetDate()
                Dim objTr As New clsDBTCapingDetail
                obj.arr = New List(Of clsDBTCapingDetail)
                For Each grow As GridViewRowInfo In gvItem.Rows
                    If clsCommon.myLen(grow.Cells(colQty).Value) > 0 Then
                        objTr = New clsDBTCapingDetail()
                        objTr.Capping_Status = clsCommon.myCBool(grow.Cells(colCapingStatus).Value)
                        objTr.BMC_Code = clsCommon.myCstr(grow.Cells(colBMCCode).Value)
                        objTr.DCS_Code = clsCommon.myCstr(grow.Cells(colVLCCode).Value)
                        objTr.MP_Code = clsCommon.myCstr(grow.Cells(colMPCode).Value)
                        objTr.Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                        objTr.Capping_Qty = clsCommon.myCDecimal(grow.Cells(colCapingQty).Value)
                        obj.arr.Add(objTr)
                    End If
                Next
                If obj.arr.Count <= 0 Then
                    Throw New Exception("No data found to save")
                End If
                clsDBTCaping.SaveData(obj, isNewEntry)
                clsCommon.MyMessageBoxShow(Me, "Data saved successfully", Me.Text)
                LoadData(obj.Document_Code, NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub DeleteData()
        Try
            If (deleteConfirm()) Then
                If (clsDBTCaping.DeleteData(txtDocumentNo.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    Reset()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        IsinsideLoadData = True
        Reset()
        Dim Whrcls As String = ""
        If isCappingIncreaseScreen Then
            Whrcls += " and TSPL_DBT_CAPING.Status=1 "
        End If
        Dim obj As clsDBTCaping = clsDBTCaping.GetData(strCode, NavTyep, Nothing, Whrcls)
        If obj IsNot Nothing Then
            DisableInputDataField()
            isNewEntry = False
            txtDocumentNo.Value = obj.Document_Code
            txtdate.Value = obj.Document_Date
            lblPending.Status = obj.Status
            txtReco.Value = obj.Reco_Code
            SetDate()
            Whrcls = ""
            If isCappingIncreaseScreen Then
                Whrcls += " and isnull(TSPL_DBT_CAPING_DETAIL.Capping_Status,0)=0 "
            End If
            SetFieldName(clsDBTCapingDetail.getData(obj.Document_Code, Whrcls))
            txtDocumentNo.MyReadOnly = True
            If obj.Status = ERPTransactionStatus.Approved Then
                btnsave.Enabled = False
                btndelete.Enabled = False
                btnPost.Enabled = False
            Else
                btnsave.Text = "Update"
                btnsave.Enabled = True
                btndelete.Enabled = True
                btnPost.Enabled = True
            End If
        End If
        IsinsideLoadData = False
    End Sub

    Private Sub txtDocumentNo__MYNavigator(ByVal sender As Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocumentNo._MYNavigator
        Try
            Dim Whrcls As String = ""
            If isCappingIncreaseScreen Then
                Whrcls += " and TSPL_DBT_CAPING.Status=1 "
            End If
            Dim qry As String = "select count(*) from TSPL_DBT_CAPING where Document_Code='" + txtDocumentNo.Value + "'   " + Whrcls
            Dim check As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry))
            If check > 0 Then
                txtDocumentNo.MyReadOnly = True
            ElseIf check <= 0 Then
                txtDocumentNo.MyReadOnly = False
            End If

            LoadData(txtDocumentNo.Value, NavType)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDocumentNo__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtDocumentNo._MYValidating
        Try
            Dim Whrcls As String = ""
            If isCappingIncreaseScreen Then
                Whrcls += " and TSPL_DBT_CAPING.Status=1 "
            End If

            txtDocumentNo.Value = clsDBTCaping.getFinder(Whrcls, txtDocumentNo.Value, isButtonClicked)
            LoadData(txtDocumentNo.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub DisableInputDataField()
        txtdate.Enabled = False
        'txtFromDate.Enabled = False
        txtReco.Enabled = False
    End Sub

    Sub EnableInputDataField()
        txtdate.Enabled = True
        'txtFromDate.Enabled = True
        txtReco.Enabled = True
    End Sub

    Private Sub btnclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnclose.Click
        CloseForm()
    End Sub

    Private Sub btndelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btndelete.Click
        DeleteData()
    End Sub

    Private Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub btnsave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnsave.Click
        SaveData(False)
    End Sub

    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        PostData()
    End Sub

    Private Sub PostData()
        Try
            Dim qry As String = ""
            Dim msg As String = ""
            Dim dt As DataTable = Nothing
            If (myMessages.postConfirm()) Then
                clsDBTCaping.PostData(txtDocumentNo.Value)
                clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                LoadData(txtDocumentNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub

    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        Try
            If clsCommon.myLen(Me.Form_ID) > 0 Then
                gvItem.MasterTemplate.FilterDescriptors.Clear()
                Dim obj As New clsGridLayout()
                obj.ReportID = Me.Form_ID
                obj.UserID = objCommonVar.CurrentUserCode
                obj.GridLayout = New MemoryStream()
                gvItem.SaveLayout(obj.GridLayout)
                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                obj.GridColumns = gvItem.ColumnCount
                If obj.SaveData() Then
                    common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
                End If
                obj.GridLayout.Close()
                obj.GridLayout.Dispose()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        Try
            clsGridLayout.DeleteData(Me.Form_ID, objCommonVar.CurrentUserCode)
            common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtReco__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtReco._MYValidating
        Dim qry As String = "select TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code,convert(date, TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Date,103) as Document_Date, 
TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date as FromDate, TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Reco_Date_To as ToDate
from TSPL_DCS_MP_INCENTIVE_RECO_HEAD"
        Dim whrcls As String = " Status=1 and isnull(DBT_Capping_Apply,0)=1 and not exists(select 1 from TSPL_DBT_CAPING where TSPL_DBT_CAPING.Reco_Code=TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code and TSPL_DBT_CAPING.Document_Code not in ('" + txtDocumentNo.Value + "'))"
        txtReco.Value = clsCommon.ShowSelectForm("DBTCAP@F", qry, "Document_Code", whrcls, txtReco.Value, "Document_Code", isButtonClicked)
        If clsCommon.myLen(txtReco.Value) > 0 Then
            qry = clsDBFuncationality.getSingleValue("select Document_Code from TSPL_DBT_CAPING where Reco_Code='" + txtReco.Value + "'")
            If clsCommon.myLen(qry) > 0 Then
                LoadData(qry, NavigatorType.Current)
            Else
                SetDate()
                FillRecoData()
            End If
        End If
    End Sub

    Private Sub FillRecoData()
        Dim ii As Integer = 0
        Try
            loadBlankGrid()
            Dim qry As String = "select 0 as PK_Id,ROW_NUMBER() over (order by TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_MP_MASTER.MP_Code_VLC_Uploader ) as SNo
,xxx.MCC_Code as BMC_Code,TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader as BMC_Uploader_Code,TSPL_MCC_MASTER.MCC_NAME as BMC_Name
,xxx.VLC_Code as DCS_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as DCS_Uploader_Code,TSPL_VLC_MASTER_HEAD.VLC_Name as DCS_Name
,xxx.MP_Code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as MP_Uploader_Code,TSPL_MP_MASTER.MP_Name
,xxx.Qty,(TSPL_MP_MASTER.DBT_Capping_Qty*" + clsCommon.myCstr(txtToDate.Value.Day) + ") as Capping_Qty,case when xxx.Qty>(TSPL_MP_MASTER.DBT_Capping_Qty*" + clsCommon.myCstr(txtToDate.Value.Day) + ") then 0 else 1 end as Capping_Status,null as Capping_Increase_By,null as Capping_Increase_Date,0 as Capping_Increase_Remarks  from (
select MP_Code,max(MCC_Code) as MCC_Code,MAX(VLC_Code) as VLC_Code,sum(Qty) as Qty from (
select TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty,TSPL_MP_INCENTIVE_ENTRY_DETAIL.UOM 
From  TSPL_MP_INCENTIVE_ENTRY_DETAIL
inner join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
inner join TSPL_DCS_MP_INCENTIVE_RECO_DETAIL on TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Year=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Year and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Month=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Month and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_No=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_No and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code 
inner join TSPL_DCS_MP_INCENTIVE_RECO_HEAD on TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code=TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code
where TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code='" + txtReco.Value + "' and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.DBT_Capping_Apply=1 and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Status=1  
) xx group by MP_Code
)xxx
left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=xxx.MP_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=xxx.VLC_Code
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=xxx.MCC_Code
order by TSPL_MCC_MASTER.Mcc_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_MP_MASTER.MP_Code_VLC_Uploader"
            SetFieldName(clsDBFuncationality.GetDataTable(qry))



        Catch ex As Exception
            loadBlankGrid()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub SetFieldName(ByVal dt As DataTable)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            gvItem.DataSource = Nothing
            gvItem.AutoGenerateColumns = False
            gvItem.DataSource = dt
            gvItem.Columns(colPKID).FieldName = "PK_Id"
            gvItem.Columns(colSlNo).FieldName = "SNo"
            gvItem.Columns(colCapingStatus).FieldName = "Capping_Status"
            gvItem.Columns(colBMCCode).FieldName = "BMC_Code"
            gvItem.Columns(colBMCUploaderNo).FieldName = "BMC_Uploader_Code"
            gvItem.Columns(colBMCName).FieldName = "BMC_Name"
            gvItem.Columns(colVLCCode).FieldName = "DCS_Code"
            gvItem.Columns(colVLCUploaderCode).FieldName = "DCS_Uploader_Code"
            gvItem.Columns(colVLCName).FieldName = "DCS_Name"
            gvItem.Columns(colMPCode).FieldName = "MP_Code"
            gvItem.Columns(colMPUploaderCode).FieldName = "MP_Uploader_Code"
            gvItem.Columns(colMPName).FieldName = "MP_Name"
            gvItem.Columns(colQty).FieldName = "Qty"
            gvItem.Columns(colCapingQty).FieldName = "Capping_Qty"
            gvItem.Columns(colCapingIncreaseBy).FieldName = "Capping_Increase_By"
            gvItem.Columns(colCapingIncreaseDate).FieldName = "Capping_Increase_Date"
            gvItem.Columns(colCapingIncreaseRemarks).FieldName = "Capping_Increase_Remarks"
            gvItem.Columns(colThirdPartyCode).FieldName = "THIRD_PARTY_CODE"

        Else
            Throw New Exception("Data Not Found")
        End If
    End Sub

    Sub SetDate()
        If clsCommon.myLen(txtReco.Value) > 0 Then
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Reco_Date,Reco_Date_To from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code='" + txtReco.Value + "'")
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                txtFromDate.Value = clsCommon.myCDate(dt.Rows(0)("Reco_Date"))
                txtToDate.Value = clsCommon.myCDate(dt.Rows(0)("Reco_Date_To"))
            End If
        End If
    End Sub



    Private Sub gvItem_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gvItem.DoubleClick
        Try
            If isCappingIncreaseScreen Then
                If gvItem.CurrentRow.Index >= 0 Then
                    Dim frm As New frmFreeTxtNumBox
                    frm.Text = "Remarks for Increate the capping"
                    frm.NumValue = clsCommon.myRoundOFF(clsCommon.myCDecimal(gvItem.CurrentRow.Cells(colQty).Value) / txtToDate.Value.Day, 0, 0)
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 AndAlso frm.NumValue <= 0 Then
                        Exit Sub
                    Else
                        clsDBTCapingDetail.CappingIncrease(clsCommon.myCDecimal(gvItem.CurrentRow.Cells(colPKID).Value), frm.strRmks, frm.NumValue, txtToDate.Value.Day)
                        clsCommon.MyMessageBoxShow(Me, "Capping increased of selected farmer", Me.Text)
                        LoadData(txtDocumentNo.Value, NavigatorType.Current)
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocumentNo.Value, "Document_Code", "TSPL_DBT_CAPING", "TSPL_DBT_CAPING_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub btnReverse_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                If clsDBTCaping.ReverseAndUnpost(txtDocumentNo.Value) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully reversed and unposted", Me.Text)
                    LoadData(txtDocumentNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class

