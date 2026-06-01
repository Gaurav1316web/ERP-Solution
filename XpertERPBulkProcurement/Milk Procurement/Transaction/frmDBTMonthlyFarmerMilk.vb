Imports System.Data.SqlClient
Imports common
Imports System.IO

Public Class frmDBTMonthlyFarmerMilk
    Inherits FrmMainTranScreen
#Region "Variables"
    Public Const colVLCCode As String = "VLC_Code"
    Public Const colVLCUploaderCode As String = "Vlc_Code_VLC_Uploader"
    Public Const colVLCName As String = "VLC_Name"
    Public Const colMPCode As String = "MP_Code"
    Public Const colMPUploaderCode As String = "MP_Uploader_Code"
    Public Const colMPName As String = "MP_Name"
    Public Const colQty As String = "Qty"
    Public Const colCollectionQty As String = "Collection_Qty"
    Public Const colTHIRDPARTYCODE As String = "THIRD_PARTY_CODE"

    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim IsinsideLoadData As Boolean = False
#End Region
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub FrmVLCDataUploaderManual_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetUserMgmtNew()
        Reset()

        ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Transaction")
        ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D Delete Transaction")
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
            SaveData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btnDelete.Enabled Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnclose.Enabled Then
            CloseForm()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnclose.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                      "========Table Name=========" + Environment.NewLine +
                      "TSPL_DBT_MONTHLY_FARMER_MILK" + Environment.NewLine +
                      "TSPL_DBT_MONTHLY_FARMER_MILK_DETAIL" + Environment.NewLine)
        End If
    End Sub
    Private Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub
    Sub loadBlankGrid()
        Try
            gvItem.DataSource = Nothing
            gvItem.Rows.Clear()
            gvItem.Columns.Clear()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.FrmVLCDataUploaderManual)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btndelete.Visible = MyBase.isDeleteFlag
        btnPost.Visible = MyBase.isPostFlag
    End Sub
    Sub Reset()
        loadBlankGrid()
        Dim dt As Date = clsCommon.GETSERVERDATE()
        txtdate.Value = dt
        txtFromDate.Value = dt
        txtToDate.Value = dt
        txtDocumentNo.Value = ""
        txtDBTReco.Value = ""
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
        txtDBTReco.Enabled = True
    End Sub

    Private Function AllowToSave() As Boolean
        Xtra.TransactionValidity(txtdate.Value)
        If AllowFutureDateTransaction(txtdate.Value, Nothing) = False Then
            txtdate.Focus()
            Return False
        End If
        If clsCommon.myLen(txtDBTReco.Value) <= 0 Then
            txtDBTReco.Focus()
            Throw New Exception("Please select " + txtDBTReco.MyLinkLable1.Text)
        End If
        If clsCommon.GetDateWithStartTime(txtFromDate.Value) = clsCommon.GetDateWithStartTime(txtToDate.Value) Then
            txtFromDate.Focus()
            Throw New Exception("Invalid from and to Date")
        End If
        Return True
    End Function
    Sub SaveData()
        Try
            If AllowToSave() Then
                Dim obj As New clsDBTMonthlyFarmerMilk()
                obj.Document_Code = txtDocumentNo.Value
                obj.Document_Date = txtdate.Value
                obj.DBT_Reco_Code = txtDBTReco.Value
                obj.From_Date = txtFromDate.Value
                obj.To_Date = txtToDate.Value
                Dim objTr As New clsDBTMonthlyFarmerMilkDetail
                obj.arr = New List(Of clsDBTMonthlyFarmerMilkDetail)
                For Each grow As GridViewRowInfo In gvItem.Rows
                    If clsCommon.myCdbl(grow.Cells(clsMPIncetiveEntryColumns.colQty).Value) > 0 Then
                        objTr = New clsDBTMonthlyFarmerMilkDetail()
                        objTr.VLC_Code = clsCommon.myCstr(grow.Cells(colVLCCode).Value)
                        objTr.MP_Code = clsCommon.myCstr(grow.Cells(colMPCode).Value)
                        objTr.Collection_Qty = clsCommon.myCdbl(grow.Cells(colCollectionQty).Value)
                        objTr.Qty = clsCommon.myCdbl(grow.Cells(colQty).Value)
                        obj.arr.Add(objTr)
                    End If
                Next
                If obj.arr.Count <= 0 Then
                    Throw New Exception("Please Enter Qty of At lease one Farmer")
                End If
                If (clsDBTMonthlyFarmerMilk.SaveData(obj, isNewEntry)) Then
                    clsCommon.MyMessageBoxShow(Me, "Data saved successfully", Me.Text)
                    LoadData(obj.Document_Code, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            If clsCommon.myLen(ex.Message) > 200 Then
                clsERPFuncationality.OpenNotepadFile(ex.Message, Me.Text)
            Else
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End If
        End Try
    End Sub
    Private Sub DeleteData()
        Try
            If (deleteConfirm()) Then
                If (clsDBTMonthlyFarmerMilk.DeleteData(txtDocumentNo.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    Reset()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Reset()
        IsinsideLoadData = True
        Dim obj As clsDBTMonthlyFarmerMilk = clsDBTMonthlyFarmerMilk.GetData(strCode, NavTyep, Nothing, False)
        If obj IsNot Nothing Then

            isNewEntry = False
            txtDocumentNo.Value = obj.Document_Code
            txtdate.Value = obj.Document_Date
            txtDBTReco.Value = obj.DBT_Reco_Code
            txtFromDate.Value = obj.From_Date
            txtToDate.Value = obj.To_Date
            lblPending.Status = obj.Status
            gvItem.DataSource = Nothing

            gvItem.DataSource = clsDBFuncationality.GetDataTable(clsDBTMonthlyFarmerMilkDetail.GetQry(obj.Document_Code))
            FormatGrid()
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
            DisableInputDataField()
        End If
        IsinsideLoadData = False
    End Sub

    Private Sub FormatGrid()
        Try

            gvItem.Columns(colVLCCode).HeaderText = "DCS Code"
            gvItem.Columns(colVLCCode).ReadOnly = True
            gvItem.Columns(colVLCCode).IsVisible = False
            gvItem.Columns(colVLCCode).TextAlignment = System.Drawing.ContentAlignment.MiddleLeft


            gvItem.Columns(colVLCUploaderCode).HeaderText = "DCS"
            gvItem.Columns(colVLCUploaderCode).ReadOnly = True
            gvItem.Columns(colVLCUploaderCode).IsVisible = True
            gvItem.Columns(colVLCUploaderCode).Width = 100
            gvItem.Columns(colVLCUploaderCode).TextAlignment = System.Drawing.ContentAlignment.MiddleLeft

            gvItem.Columns(colVLCName).HeaderText = "DCS Name"
            gvItem.Columns(colVLCName).ReadOnly = True
            gvItem.Columns(colVLCName).IsVisible = True
            gvItem.Columns(colVLCName).TextAlignment = System.Drawing.ContentAlignment.MiddleLeft

            gvItem.Columns(colMPCode).HeaderText = "Farmer Code"
            gvItem.Columns(colMPCode).ReadOnly = True
            gvItem.Columns(colMPCode).IsVisible = False
            gvItem.Columns(colMPCode).TextAlignment = System.Drawing.ContentAlignment.MiddleLeft

            gvItem.Columns(colMPUploaderCode).HeaderText = "Farmer"
            gvItem.Columns(colMPUploaderCode).Width = 100
            gvItem.Columns(colMPUploaderCode).ReadOnly = True
            gvItem.Columns(colMPUploaderCode).TextAlignment = System.Drawing.ContentAlignment.MiddleLeft

            gvItem.Columns(colMPName).HeaderText = "Farmer Name"
            gvItem.Columns(colMPName).Width = 200
            gvItem.Columns(colMPName).ReadOnly = True
            gvItem.Columns(colMPName).TextAlignment = System.Drawing.ContentAlignment.MiddleLeft

            gvItem.Columns(colCollectionQty).HeaderText = "Collection Quantity"
            gvItem.Columns(colCollectionQty).Width = 120
            gvItem.Columns(colCollectionQty).ReadOnly = False
            gvItem.Columns(colCollectionQty).TextAlignment = System.Drawing.ContentAlignment.MiddleRight

            gvItem.Columns(colQty).HeaderText = "Quantity"
            gvItem.Columns(colQty).Width = 120
            gvItem.Columns(colQty).ReadOnly = False
            gvItem.Columns(colQty).TextAlignment = System.Drawing.ContentAlignment.MiddleRight

            gvItem.Columns(colTHIRDPARTYCODE).HeaderText = "Third Party Code"
            gvItem.Columns(colTHIRDPARTYCODE).Width = 120
            gvItem.Columns(colTHIRDPARTYCODE).ReadOnly = True

            gvItem.Columns(colTHIRDPARTYCODE).IsVisible = False
            gvItem.Columns(colTHIRDPARTYCODE).VisibleInColumnChooser = True

            gvItem.Columns(colTHIRDPARTYCODE).TextAlignment = System.Drawing.ContentAlignment.MiddleRight

            gvItem.MasterTemplate.SummaryRowsBottom.Clear()
            Dim summaryRowItem As New GridViewSummaryRowItem()
            summaryRowItem.Add(New GridViewSummaryItem(colCollectionQty, "{0:n2}", GridAggregateFunction.Sum))
            gvItem.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            summaryRowItem.Add(New GridViewSummaryItem(colQty, "{0:n2}", GridAggregateFunction.Sum))
            gvItem.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gvItem.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom

            gvItem.BestFitColumns()

            ReStoreGridLayout()

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
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDocumentNo__MYNavigator(ByVal sender As Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocumentNo._MYNavigator
        Try
            Dim qry As String = "select count(*) from TSPL_MP_INCENTIVE_ENTRY_HEAD where Document_Code='" + txtDocumentNo.Value + "' "
            Dim check As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
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
            txtDocumentNo.Value = clsDBTMonthlyFarmerMilk.GetFinder("", txtDocumentNo.Value, isButtonClicked)
            LoadData(txtDocumentNo.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub DisableInputDataField()
        txtdate.Enabled = False
        txtDBTReco.Enabled = False
        txtFromDate.Enabled = False
        txtToDate.Enabled = False
    End Sub
    Sub EnableInputDataField()
        txtdate.Enabled = True
        txtDBTReco.Enabled = True
        txtFromDate.Enabled = True
        txtToDate.Enabled = True
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
        SaveData()
    End Sub


    Sub GridFocus(Optional ByVal gvRow As GridViewRowInfo = Nothing)
        If gvItem.Rows.Count > 0 Then
            gvItem.Focus()
            gvItem.CurrentColumn = gvItem.Columns(clsMPIncetiveEntryColumns.colMPUploaderCode)
            If Not gvRow Is Nothing Then
                gvItem.CurrentRow = gvItem.Rows(gvRow.Index + 1)
            Else
                If gvItem.CurrentRow Is Nothing Then
                    gvItem.CurrentRow = gvItem.Rows(gvItem.Rows.Count - 1)
                End If
            End If


            gvItem.PerformLayout()
            gvItem.BeginEdit()
            gvItem.EndEdit()
        Else
            gvItem.Rows.AddNew()
            gvItem.CurrentColumn = gvItem.Columns(clsMPIncetiveEntryColumns.colMPUploaderCode)
            gvItem.CurrentRow = gvItem.Rows(gvItem.Rows.Count - 1)
            gvItem.PerformLayout()
            gvItem.BeginEdit()
        End If
    End Sub
    Private Sub gvItem_KeyDown(sender As Object, e As KeyEventArgs) Handles gvItem.KeyDown
        If True Then
            Dim x As Integer = 10
        End If
        'If e.KeyCode = Keys.Enter Then
        '    If gvItem.CurrentCell Is Nothing Then
        '        Exit Sub
        '    End If
        '    If (gvItem.CurrentCell.ColumnInfo.Name = clsMPIncetiveEntryColumns.colAmount) Then
        '        GridFocus(gvItem.CurrentRow)
        '    ElseIf (gvItem.CurrentCell.ColumnInfo.Name = clsMPIncetiveEntryColumns.colMPUploaderCode) Then
        '        gvItem.EndEdit()
        '        gvItem.CurrentColumn = gvItem.Columns(clsMPIncetiveEntryColumns.colQty)
        '        gvItem.BeginEdit()
        '    ElseIf (gvItem.CurrentCell.Colum   nInfo.Name = clsMPIncetiveEntryColumns.colMPUploaderCode Or gvItem.CurrentCell.ColumnInfo.Name = clsMPIncetiveEntryColumns.colMPName Or gvItem.CurrentCell.ColumnInfo.Name = clsMPIncetiveEntryColumns.colUOM Or gvItem.CurrentCell.ColumnInfo.Name = clsMPIncetiveEntryColumns.colQty) Then
        '        gvItem.EndEdit()
        '        gvItem.CurrentColumn = gvItem.Columns(gvItem.CurrentCell.ColumnInfo.Index + 1)
        '        gvItem.BeginEdit()
        '    End If
        'End If
    End Sub
    Private Sub txtMCC__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDBTReco._MYValidating
        txtDBTReco.Value = clsMPDCSInsentiveReco.getFinder(" not exists( select 1 from TSPL_DBT_MONTHLY_FARMER_MILK where TSPL_DBT_MONTHLY_FARMER_MILK.DBT_Reco_Code=TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code  and TSPL_DBT_MONTHLY_FARMER_MILK.Document_Code not in ('" + txtDocumentNo.Value + "'))", txtDBTReco.Value, isButtonClicked)
        If clsCommon.myLen(txtDBTReco.Value) > 0 Then
            SetRecoDate()
            fillMPS()
        End If
    End Sub

    Private Sub SetRecoDate()
        Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Reco_Date,Reco_Date_To from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code='" + txtDBTReco.Value + "' ")
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            txtFromDate.Value = clsCommon.myCDate(dt.Rows(0)("Reco_Date"))
            txtToDate.Value = clsCommon.myCDate(dt.Rows(0)("Reco_Date_To"))
        End If
    End Sub

    Sub fillMPS()
        If clsCommon.myLen(txtDBTReco.Value) <= 0 Then
            txtDBTReco.Focus()
            Throw New Exception("Please select " + txtDBTReco.MyLinkLable1.Text)
        End If
        loadBlankGrid()
        'Dim Mcc_Uom As String = clsDBFuncationality.getSingleValue("select Uom_Code from TSPL_Mcc_UOM_DETAIL where Stocking_Unit='Y' and  MCC_CODE='" & txtDBTReco.Value & "'")
        Dim Qry As String = "Select VLC_Code,max(Vlc_Code_VLC_Uploader) As Vlc_Code_VLC_Uploader,max(VLC_Name) As VLC_Name,MP_Code,max(MP_Uploader_Code) As MP_Uploader_Code,max(MP_Name) As MP_Name,sum(Collection_Qty) As Collection_Qty,sum(qty) As Qty,MAX(THIRD_PARTY_CODE)THIRD_PARTY_CODE from (
select Doc_Date,VLC_Code,max(Vlc_Code_VLC_Uploader) As Vlc_Code_VLC_Uploader,max(VLC_Name) As VLC_Name,MP_Code,max(MP_Uploader_Code) As MP_Uploader_Code,max(MP_Name) As MP_Name,sum(qty) as Collection_Qty ,(case when max(DBT_Capping_Qty)>0 and sum(qty)>max(DBT_Capping_Qty) then  max(DBT_Capping_Qty) else sum(qty) end)  As Qty,MAX(THIRD_PARTY_CODE)THIRD_PARTY_CODE from (
 select Doc_No,Convert(Varchar(10),Doc_Date,103)Doc_Date,shift,TSPL_VLC_MASTER_HEAD.VLC_Code, TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,
TSPl_MP_MAster.MP_CODE,TSPl_MP_MAster.MP_Name, TSPl_MP_MAster.MP_Code_VLC_Uploader As MP_Uploader_Code,TSPL_VLC_DATA_UPLOADER.qty ,TSPl_MP_MAster.THIRD_PARTY_CODE,isnull(TSPl_MP_MAster.DBT_Capping_Qty,0) as DBT_Capping_Qty
from TSPL_VLC_DATA_UPLOADER
Left Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.VLC_CODE
Left Join TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER.MP_CODE and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
 inner join (select VLC_Code,Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL union all select VLC_Code,Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID) as TSPL_DCS_MP_INCENTIVE_RECO_DETAIL on TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
where TSPL_VLC_DATA_UPLOADER.Doc_Date>='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And TSPL_VLC_DATA_UPLOADER.Doc_Date<='" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'  and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code='" & txtDBTReco.Value & "'
 Union All 
 select  TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code As Doc_No,
Convert(Varchar(10),TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) As Doc_Date, TSPL_VLC_DATA_UPLOADER_MASTER.Shift,TSPL_VLC_DATA_UPLOADER_MASTER.VLC_CODE,
TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPl_MP_MAster.MP_CODE,TSPl_MP_MAster.MP_Name,
TSPl_MP_MAster.MP_Code_VLC_Uploader As MP_Uploader_Code,TSPL_VLC_DATA_UPLOADER_DETAIL.qty ,TSPl_MP_MAster.THIRD_PARTY_CODE,isnull(TSPl_MP_MAster.DBT_Capping_Qty,0) as DBT_Capping_Qty
from TSPL_VLC_DATA_UPLOADER_DETAIL
Left Outer Join TSPL_VLC_DATA_UPLOADER_MASTER On TSPL_VLC_DATA_UPLOADER_MASTER.Document_Code=TSPL_VLC_DATA_UPLOADER_DETAIL.Document_Code
Left Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.Vlc_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_CODE
Left Join TSPl_MP_MAster On TSPl_MP_MAster.MP_Code_VLC_Uploader=TSPL_VLC_DATA_UPLOADER_DETAIL.Farmer_Code and TSPl_MP_MAster.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code
 inner join (select VLC_Code,Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL union all select VLC_Code,Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID) as TSPL_DCS_MP_INCENTIVE_RECO_DETAIL on TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code=TSPL_VLC_DATA_UPLOADER_MASTER.VLC_CODE
 where  convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103) >='" & clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") & "' And convert(date,TSPL_VLC_DATA_UPLOADER_MASTER.Document_Date,103)<= '" & clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") & "'  and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code='" & txtDBTReco.Value & "' 
) x group by VLC_Code,MP_CODE,Doc_Date
) xx where MP_CODE is not null group by VLC_Code,MP_CODE  order by Vlc_Code_VLC_Uploader,MP_Uploader_Code"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            gvItem.DataSource = dt
            FormatGrid()
        End If
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
                clsDBTMonthlyFarmerMilk.PostData(txtDocumentNo.Value)
                clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                LoadData(txtDocumentNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub gvItem_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles gvItem.CellValueChanged
        'Try
        '    If (Not IsinsideLoadData) Then
        '        If Not isCellValueChangedOpen Then
        '            isCellValueChangedOpen = True
        '            'If e.Column Is gvItem.Columns(clsMPIncetiveEntryColumns.colQty) OrElse e.Column Is gvItem.Columns(clsMPIncetiveEntryColumns.colPashuAaharQty) OrElse e.Column Is gvItem.Columns(clsMPIncetiveEntryColumns.colMineralMixtureQty) OrElse e.Column Is gvItem.Columns(clsMPIncetiveEntryColumns.colSailejQty) OrElse e.Column Is gvItem.Columns(clsMPIncetiveEntryColumns.colRahatKampekatFeedQty) Then
        '            '    UpdateRow(gvItem.CurrentRow.Index)
        '            'End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        'Finally
        '    isCellValueChangedOpen = False
        'End Try
    End Sub

    Private Sub RadMenuItem1_Click(sender As Object, e As EventArgs) Handles RadMenuItem1.Click
        If clsCommon.myLen(MyBase.Form_ID) > 0 Then
            gvItem.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = MyBase.Form_ID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gvItem.SaveLayout(obj.GridLayout)
            obj.GridColumns = gvItem.ColumnCount
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            ''stuti regarding memory leakage
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub


    Private Sub RadMenuItem2_Click(sender As Object, e As EventArgs) Handles RadMenuItem2.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
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






End Class

