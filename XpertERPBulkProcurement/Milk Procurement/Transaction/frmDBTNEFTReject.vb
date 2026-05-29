Imports common

Public Class frmDBTNEFTReject
    Inherits FrmMainTranScreen
#Region "Variables"
    Public Const colSlNo As String = "SLNO"
    Public Const colAgainstDBT As String = "colAgainstDBT"
    Public Const colRemAccNo As String = "colRemAccNo"
    Public Const colRemName As String = "colRemName"
    Public Const colSociety As String = "colSociety"
    Public Const colMPUploaderCode As String = "colMPUploaderCode"
    Public Const colAmount As String = "colAmount"
    Public Const colMPIFSCCode As String = "colMPIFSCCode"
    Public Const colMPAccountNo As String = "colMPAccountNo"
    Public Const colMPName As String = "colMPName"
    Public Const colTransaction As String = "colTransaction"
    Public Const colSelect As String = "colSelect"
    Public Const colRemarks As String = "colRemarks"

    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim IsinsideLoadData As Boolean = False
    Dim Qry As String
    Dim isFlag As Boolean = False

    Dim strPortNo As String
#End Region
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub FrmVLCDataUploaderManual_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        strPortNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  [TSPL_APP_LOCATION].PD_Account_Prefix as PortNo  from TSPL_MASTER.dbo.TSPL_APP_LOCATION WHERE  [TSPL_APP_LOCATION].DataBase_Name='" + objCommonVar.CurrDatabase + "' and [TSPL_APP_LOCATION].Apply_PD_Account = 1 order by [TSPL_APP_LOCATION].Location_Name "))
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
            SaveData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btndelete.Enabled Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnclose.Enabled Then
            CloseForm()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnPost.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            Dim frm As New frmPWDHighSecrity(Nothing)
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                btnReverse.Visible = True
            End If
        End If
    End Sub
    Private Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub

    Private Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.FrmVLCDataUploaderManual)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btndelete.Visible = MyBase.isDeleteFlag
        btnPost.Visible = MyBase.isPostFlag
        btn_Import.Visible = MyBase.isExport
        If MyBase.isExport = True Then
            btn_Import.Enabled = True
        Else
            btn_Import.Enabled = False
        End If
    End Sub
    Sub Reset()
        'loadBlankGrid()
        gvItem.DataSource = Nothing
        Dim dt As Date = clsCommon.GETSERVERDATE()

        txtDocumentNo.Value = ""
        txtdate.Value = dt
        txtDocumentNo.MyReadOnly = False
        btnsave.Text = "Save"

        btndelete.Enabled = False
        btnsave.Enabled = True
        btnPost.Enabled = False
        txtdate.Focus()
        isNewEntry = True
        IsinsideLoadData = False
        lblPending.Status = ERPTransactionStatus.Pending
        txtRemarks.Text = ""
        txtDPTNEFT.Value = ""
        txtPDABillNo.Value = ""
    End Sub

    Private Function AllowToSave() As Boolean
        Xtra.TransactionValidity(txtdate.Value)
        If clsCommon.myLen(txtDPTNEFT.Value) <= 0 Then
            Throw New Exception("Please select DPT NEFT No")
        End If
        Return True
    End Function
    Sub SaveData()
        Try
            If AllowToSave() Then
                Dim obj As New clsDBTNEFTReject()
                obj.Document_Code = txtDocumentNo.Value
                obj.Document_Date = txtdate.Value
                obj.Remarks = txtRemarks.Text
                obj.Against_DBT_NEFT = txtDPTNEFT.Value
                obj.Against_PDA_BillNo = txtPDABillNo.Value
                obj.arr = New List(Of clsDBTNEFTRejectDetail)
                obj.arrSuccess = New List(Of clsDBTNEFTRejectSucess)
                For Each grow As GridViewRowInfo In gvItem.Rows
                    If clsCommon.myCBool(grow.Cells(colSelect).Value) Then
                        If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Remarks").Value), "SUCCESS") = CompairStringResult.Equal Then
                            Dim objTr As New clsDBTNEFTRejectSucess
                            objTr = New clsDBTNEFTRejectSucess()
                            objTr.Against_DBT_NEFT_TR = clsCommon.myCstr(grow.Cells("AgainstMPIncetive").Value)
                            objTr.Remarks = clsCommon.myCstr(grow.Cells("Remarks").Value)
                            obj.arrSuccess.Add(objTr)
                        Else
                            Dim objTr As New clsDBTNEFTRejectDetail
                            objTr = New clsDBTNEFTRejectDetail()
                            objTr.Against_DBT_NEFT_TR = clsCommon.myCstr(grow.Cells("AgainstMPIncetive").Value)
                            objTr.Remarks = clsCommon.myCstr(grow.Cells("Remarks").Value)
                            obj.arr.Add(objTr)
                        End If
                    End If
                Next
                If ((obj.arr Is Nothing OrElse obj.arr.Count <= 0) AndAlso (obj.arrSuccess Is Nothing OrElse obj.arrSuccess.Count <= 0)) Then
                    Throw New Exception("Please select at least one rejected/Success Account")
                End If
                If (clsDBTNEFTReject.SaveData(obj, isNewEntry)) Then
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
            'clsLockMPPaymentCycle.LockMPTransaction(txtMCC.Value, txtdate.Value)
            If (deleteConfirm()) Then
                If (clsDBTNEFTReject.DeleteData(txtDocumentNo.Value)) Then
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
        Dim obj As clsDBTNEFTReject = clsDBTNEFTReject.GetData(strCode, NavTyep)
        If obj IsNot Nothing Then
            isNewEntry = False
            txtDocumentNo.Value = obj.Document_Code
            txtdate.Value = obj.Document_Date
            lblPending.Status = obj.Status
            txtRemarks.Text = obj.Remarks
            txtDPTNEFT.Value = obj.Against_DBT_NEFT
            txtPDABillNo.Value = obj.Against_PDA_BillNo
            If clsCommon.myLen(obj.Document_Code) > 0 Then
                Dim qry As String = "Select CAST(1 as bit) as colSelect,ROW_NUMBER() OVER (ORDER BY (SELECT 1)) AS SNo,x.* from (
select  TSPL_DBT_NEFT_REJECT_DETAIL.Against_DBT_NEFT_TR AS AgainstMPIncetive,TSPL_DBT_NEFT_DETAIL.Rem_Account_No as [REMITTER ACCOUNT NO.],TSPL_DBT_NEFT_DETAIL.Rem_Name as [REMITTER NAME],TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code as [Society],TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code as [MP Uploader Code],TSPL_DBT_NEFT_DETAIL.Amount as AMOUNT,TSPL_DBT_NEFT_DETAIL.MP_IFSC_No as [IFSC CODE],TSPL_DBT_NEFT_DETAIL.MP_Account_No as [BENEFICERY ACCOUNT  NO.],TSPL_DBT_NEFT_DETAIL.MP_Name as [BENEFICERY NAME],TSPL_DBT_NEFT_DETAIL.[Transaction] as [TRANSACTION DESC CREDIT],TSPL_DBT_NEFT_REJECT_DETAIL.Remarks,TSPL_MP_MASTER.THIRD_PARTY_CODE 
from TSPL_DBT_NEFT_REJECT_DETAIL 
left outer join TSPL_DBT_NEFT_DETAIL on TSPL_DBT_NEFT_DETAIL.PK_Id=TSPL_DBT_NEFT_REJECT_DETAIL.Against_DBT_NEFT_TR
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code

where TSPL_DBT_NEFT_REJECT_DETAIL.Document_Code='" & txtDocumentNo.Value & "'
union all
select TSPL_DBT_NEFT_REJECT_SUCESS.Against_DBT_NEFT_TR AS AgainstMPIncetive ,TSPL_DBT_NEFT_DETAIL.Rem_Account_No as [REMITTER ACCOUNT NO.],TSPL_DBT_NEFT_DETAIL.Rem_Name as [REMITTER NAME],TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code as [Society],TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code as [MP Uploader Code],TSPL_DBT_NEFT_DETAIL.Amount as AMOUNT,TSPL_DBT_NEFT_DETAIL.MP_IFSC_No as [IFSC CODE],TSPL_DBT_NEFT_DETAIL.MP_Account_No as [BENEFICERY ACCOUNT  NO.],TSPL_DBT_NEFT_DETAIL.MP_Name as [BENEFICERY NAME],TSPL_DBT_NEFT_DETAIL.[Transaction] as [TRANSACTION DESC CREDIT],TSPL_DBT_NEFT_REJECT_SUCESS.Remarks,TSPL_MP_MASTER.THIRD_PARTY_CODE 
from TSPL_DBT_NEFT_REJECT_SUCESS 
left outer join TSPL_DBT_NEFT_DETAIL on TSPL_DBT_NEFT_DETAIL.PK_Id=TSPL_DBT_NEFT_REJECT_SUCESS.Against_DBT_NEFT_TR
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code

where TSPL_DBT_NEFT_REJECT_SUCESS.Document_Code='" & txtDocumentNo.Value & "'
)x order by AgainstMPIncetive"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    gvItem.DataSource = dt
                    FormatGrid()
                End If
            End If
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
            txtDocumentNo.Value = clsDBTNEFTReject.getFinder("", txtDocumentNo.Value, isButtonClicked)
            LoadData(txtDocumentNo.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
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
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
            myMessages.blankValue(Me, "Invoice not found to Print", Me.Text)
        Else
            'funPrint(txtDocumentNo.Value)
        End If
    End Sub

    Sub GridFocus(Optional ByVal gvRow As GridViewRowInfo = Nothing)
        If gvItem.Rows.Count > 0 Then
            gvItem.Focus()
            gvItem.CurrentColumn = gvItem.Columns(colMPUploaderCode)
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
            gvItem.CurrentColumn = gvItem.Columns(colMPUploaderCode)
            gvItem.CurrentRow = gvItem.Rows(gvItem.Rows.Count - 1)
            gvItem.PerformLayout()
            gvItem.BeginEdit()
        End If
    End Sub






    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        PostData()
    End Sub

    Sub PostData()
        Try
            Dim qry As String = ""
            Dim msg As String = ""
            Dim dt As DataTable = Nothing

            If (myMessages.postConfirm()) Then
                clsDBTNEFTReject.PostData(txtDocumentNo.Value)
                clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                LoadData(txtDocumentNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtPDABillNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtPDABillNo._MYValidating
        Try
            If clsCommon.myLen(strPortNo) > 0 Then
                Dim qry As String = "Select xx.PDABillNo,TSPL_DBT_NEFT.Document_Code as Code,Convert(varchar,TSPL_DBT_NEFT.Document_Date,103) as Date
,TSPL_DBT_NEFT.Remarks as [Remarks],Convert(varchar,TSPL_DBT_NEFT.From_Date,103) as [From Date],Convert(varchar,TSPL_DBT_NEFT.To_Date,103) as [To Date],TSPL_DBT_NEFT.Zone_code as Zone
,case when isnull(TSPL_DBT_NEFT.Status,0)=0 then 'Pending' else 'Approved' end as Status,case when isnull(TSPL_DBT_NEFT.RCDF_Status,0)=0 then 'Pending' else 'Approved' end as [RCDF Status]
from TSPL_DBT_NEFT 
left outer join (
select Document_Code,PDABillNo from (
select TSPL_DBT_NEFT.Document_Code, '" + strPortNo + "'+ FORMAT(TSPL_DBT_NEFT.UKID, '0000')+CAST(TSPL_DBT_NEFT_DETAIL.Lot_No as varchar) as PDABillNo from TSPL_DBT_NEFT 
left outer join TSPL_DBT_NEFT_DETAIL on TSPL_DBT_NEFT_DETAIL.Document_Code=TSPL_DBT_NEFT.Document_Code
) xx group by Document_Code,PDABillNo 
)xx on xx.Document_Code=TSPL_DBT_NEFT.Document_Code
where isnull(TSPL_DBT_NEFT.Status,0)=1 "
                gvItem.DataSource = Nothing
                Dim dr As DataRow = clsCommon.ShowSelectFormForRow("DPTNeft#S", qry, "PDABillNo", txtPDABillNo.Value)
                If dr IsNot Nothing Then
                    txtPDABillNo.Value = clsCommon.myCstr(dr("PDABillNo"))
                    txtDPTNEFT.Value = clsCommon.myCstr(dr("Code"))
                    LoadDBTData(txtPDABillNo.Value, txtDPTNEFT.Value)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDPTNEFT__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDPTNEFT._MYValidating
        Try
            gvItem.DataSource = Nothing
            txtDPTNEFT.Value = clsDBTNEFT.getFinder(" isnull(TSPL_DBT_NEFT.Status,0)=1 ", txtDPTNEFT.Value, isButtonClicked)
            LoadDBTData("", txtDPTNEFT.Value)
            txtPDABillNo.Value = ""
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub LoadDBTData(strRefNo As String, strDBTNo As String)
        If clsCommon.myLen(strDBTNo) > 0 Then
            Dim obj As clsDBTNEFT = clsDBTNEFT.GetData(strDBTNo, NavigatorType.Current)
            If obj.arr IsNot Nothing AndAlso obj.arr.Count > 0 Then
                Qry = "Select "
                If clsCommon.myLen(strRefNo) > 0 Then
                    Qry+=" CAST(1 as bit) as colSelect,"
                Else
                    Qry+="CAST(0 as bit) as colSelect,"
                End If
                Qry += "ROW_NUMBER() over (order by max(xx.SNo) )  as SNo,xx.PK_Id as AgainstMPIncetive,max(TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code) as  MP_Code,max(xx.Rem_Account_No) AS 
[REMITTER ACCOUNT NO.],max(xx.Rem_Name) AS [REMITTER NAME],max(xx.VLC_Uploader_Code) AS [Society],max(xx.MP_Uploader_Code) AS [MP Uploader Code],max(xx.Amount) AS AMOUNT,max(xx.MP_IFSC_No) AS [IFSC CODE],max(xx.MP_Account_No) AS [BENEFICERY ACCOUNT  NO.],max(xx.MP_Name) AS [BENEFICERY NAME],max(xx.[Transaction]) AS [TRANSACTION DESC CREDIT],'' AS Remarks from (
select TSPL_DBT_NEFT_DETAIL.SNo,TSPL_DBT_NEFT_DETAIL.PK_Id,TSPL_DBT_NEFT_DETAIL.Rem_Account_No,TSPL_DBT_NEFT_DETAIL.Rem_Name ,TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code  ,TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code ,TSPL_DBT_NEFT_DETAIL.Amount,TSPL_DBT_NEFT_DETAIL.MP_IFSC_No ,TSPL_DBT_NEFT_DETAIL.MP_Account_No ,TSPL_DBT_NEFT_DETAIL.MP_Name,TSPL_DBT_NEFT_DETAIL.[Transaction],1 as RI ,TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR  
from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_DBT_NEFT  on TSPL_DBT_NEFT.Document_Code=TSPL_DBT_NEFT_DETAIL.Document_Code
where TSPL_DBT_NEFT_DETAIL.Document_Code='" + strDBTNo + "'"
                If clsCommon.myLen(strRefNo) > 0 Then
                    If clsCommon.myLen(strRefNo) <> 8 Then
                        Throw New Exception("Invalid Document [" + strRefNo + "]")
                    End If

                    Dim strUKID As String = strRefNo.Substring(2, 4)
                    Dim strLotNo As String = strRefNo.Substring(6, 2)
                    Qry += " and TSPL_DBT_NEFT.UKID=" + strUKID + " and TSPL_DBT_NEFT_DETAIL.Lot_No=" + strLotNo + ""
                End If
                Qry += " union all
select 0 as  SNo,TSPL_DBT_NEFT_REJECT_DETAIL.Against_DBT_NEFT_TR as PK_Id,'' as Rem_Account_No,'' as Rem_Name ,'' as VLC_Uploader_Code,'' as MP_Uploader_Code ,0 as Amount,'' as MP_IFSC_No,'' as MP_Account_No ,'' as MP_Name,'' as [Transaction],-1 as RI ,0 as Against_MP_Incentive_TR 
from TSPL_DBT_NEFT_REJECT_DETAIL 
left outer join TSPL_DBT_NEFT_REJECT on TSPL_DBT_NEFT_REJECT.Document_Code= TSPL_DBT_NEFT_REJECT_DETAIL.Document_Code
where TSPL_DBT_NEFT_REJECT.Against_DBT_NEFT='" + strDBTNo + "' and TSPL_DBT_NEFT_REJECT.Document_Code not in ('" + txtDocumentNo.Value + "')
)xx 
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=xx.Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
group by xx.PK_Id having sum(RI)>0"
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    gvItem.Rows.Clear()
                    gvItem.Refresh()
                    gvItem.MasterTemplate.SummaryRowsBottom.Clear()
                    gvItem.DataSource = dt
                    FormatGrid()
                End If
            End If
        End If
    End Sub

    Private Sub FormatGrid()
        Try
            gvItem.MasterTemplate.SummaryRowsBottom.Clear()
            For ii As Integer = 1 To gvItem.Columns.Count - 1
                gvItem.Columns(ii).ReadOnly = True
                gvItem.Columns(ii).FormatString = ""
                gvItem.Columns(ii).BestFit()
            Next
            gvItem.Columns("colSelect").HeaderText = "Verified"
            gvItem.Columns("SNo").HeaderText = "SNo"
            gvItem.Columns("AgainstMPIncetive").HeaderText = "DBT TR No"
            gvItem.Columns("MP_Code").HeaderText = "Farmer Code"
            gvItem.Columns("REMITTER ACCOUNT NO.").HeaderText = "REMITTER ACCOUNT NO."
            gvItem.Columns("REMITTER NAME").HeaderText = "REMITTER NAME"
            gvItem.Columns("Society").HeaderText = "Society"
            gvItem.Columns("MP Uploader Code").HeaderText = "MP Code"
            gvItem.Columns("AMOUNT").HeaderText = "AMOUNT"
            gvItem.Columns("AMOUNT").FormatString = "{0:n2}"

            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item1 As New GridViewSummaryItem("AMOUNT", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item1)
            'gvItem.ShowGroupPanel = False
            'gvItem.MasterTemplate.AutoExpandGroups = True
            gvItem.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            'gvItem.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom

            gvItem.Columns("AMOUNT").TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvItem.Columns("IFSC CODE").HeaderText = "IFSC CODE"
            gvItem.Columns("BENEFICERY ACCOUNT  NO.").HeaderText = "BENEFICERY ACCOUNT  NO."
            gvItem.Columns("BENEFICERY NAME").HeaderText = "BENEFICERY NAME"
            gvItem.Columns("TRANSACTION DESC CREDIT").HeaderText = "TRANSACTION DESC CREDIT"
            gvItem.Columns("Remarks").HeaderText = "Status"
            gvItem.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"

            gvItem.AllowAddNewRow = False
            gvItem.AllowDeleteRow = True
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

    Private Function GenerateTransposedTable(ByVal inputTable As DataTable) As DataTable
        Dim outputTable As DataTable = New DataTable()
        Try

            outputTable.Columns.Add(New DataColumn("Code", System.Type.GetType("System.String")))

            For rCount As Integer = 0 To inputTable.Columns.Count - 1
                Dim newRow As DataRow = outputTable.NewRow()
                newRow(0) = inputTable.Columns(rCount).ColumnName.ToString()
                outputTable.Rows.Add(newRow)
            Next

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return outputTable
    End Function
    Private Sub btn_Import_Click(sender As Object, e As EventArgs) Handles btn_Import.Click
        Dim gv As New UserControls.MyRadGridView
        Try
            Me.Controls.Add(gv)

            transportSql.LoadDocument(gv, "")
            Dim dtt As DataTable = TryCast(gv.DataSource, DataTable)

            Dim dt As DataTable = GenerateTransposedTable(dtt)
            Dim strFarmerCodeColumnName As String = ""
            Dim strStatusColumnName As String = ""

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim frmFC As New FrmFreeComboBox
                frmFC.ComboSource = dt
                frmFC.ComboValueMember = "Code"
                frmFC.ComboDisplayMember = "Code"
                frmFC.LabelCaption = "Farmer Code"
                frmFC.ShowDialog()
                If clsCommon.myLen(frmFC.strRetValue) > 0 Then
                    strFarmerCodeColumnName = clsCommon.myCstr(frmFC.strRetValue)
                Else
                    clsCommon.MyMessageBoxShow(Me, "Select PK ID Column", Me.Text)
                    Return
                End If

                frmFC.LabelCaption = "Status"
                frmFC.ShowDialog()
                If clsCommon.myLen(frmFC.strRetValue) > 0 Then
                    strStatusColumnName = clsCommon.myCstr(frmFC.strRetValue)
                Else
                    clsCommon.MyMessageBoxShow(Me, "Select Status Column", Me.Text)
                    Return
                End If
            End If

            For jj As Integer = 0 To dtt.Rows.Count - 1
                For Each grow As GridViewRowInfo In gvItem.Rows
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("MP_Code").Value), clsCommon.myCstr(dtt.Rows(jj).Item(strFarmerCodeColumnName))) = CompairStringResult.Equal Then
                        grow.Cells(colSelect).Value = True
                        If Not (clsCommon.CompairString(clsCommon.myCstr(dtt.Rows(jj).Item(strStatusColumnName)), "FAIL") = CompairStringResult.Equal OrElse
                                clsCommon.CompairString(clsCommon.myCstr(dtt.Rows(jj).Item(strStatusColumnName)), "REJECT") = CompairStringResult.Equal OrElse
                                clsCommon.CompairString(clsCommon.myCstr(dtt.Rows(jj).Item(strStatusColumnName)), "SUCCESS") = CompairStringResult.Equal) Then
                            Throw New Exception("Status Should be FAIL/REJECT/SUCCESS")
                        End If
                        grow.Cells("Remarks").Value = clsCommon.myCstr(dtt.Rows(jj).Item(strStatusColumnName))
                    End If
                Next
            Next

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            Me.Controls.Remove(gv)
        End Try
    End Sub

    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        Try
            If gvItem.Rows.Count <= 0 Then

                Dim arrHeader As List(Of String) = New List(Of String)()
                Dim qry As String = " select  '' as [S.no],'' as [SocietyID],'' as [FarmerName],'' as [Farmer account number],'' as [PaymentAmount],'' as [Farmer Code],'' as [Status] "
                transportSql.ExporttoExcel(qry, Me)

            Else
                Dim arrHeader As List(Of String) = New List(Of String)()
                arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtdate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtdate.Value, "dd/MM/yyyy")) + " ")
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptMobileAppMilkCollection & "'"))
                transportSql.applyExportTemplate(gvItem, PageSetupReport_ID)
                clsCommon.MyExportToExcelGrid(Me.Text, gvItem, arrHeader, Me.Text)

                'clsCommon.MyExportToExcelGrid("code", qry, arrHeader, Me.Text)
            End If
            common.clsCommon.MyMessageBoxShow(Me, "Exported Successfully.", Me.Text)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)

        End Try
    End Sub

    Private Sub rmiPDF_Click(sender As Object, e As EventArgs) Handles rmiPDF.Click
        Try
            Dim arrHeader As List(Of String) = New List(Of String)()
            arrHeader.Add(("Date Range: " + clsCommon.GetPrintDate(txtdate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtdate.Value, "dd/MM/yyyy")) + " ")
            arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
            arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptMobileAppMilkCollection & "'"))

            transportSql.applyExportTemplate(gvItem, PageSetupReport_ID)
            clsCommon.MyExportToPDF(Me.Text, gvItem, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)

        End Try

    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocumentNo.Value, "Document_Code", "TSPL_DBT_NEFT_REJECT", "TSPL_DBT_NEFT_REJECT_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub btnReverse_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
                If clsCommon.MyMessageBoxShow(Me, "Unpost the current transaction" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    clsDBTNEFTReject.ReverseAndUnpost(txtDocumentNo.Value)
                    clsCommon.MyMessageBoxShow(Me, "Tansaction unposted succesffuly", Me.Text)
                    LoadData(txtDocumentNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


End Class

