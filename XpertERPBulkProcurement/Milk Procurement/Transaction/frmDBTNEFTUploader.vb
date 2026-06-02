Imports System.Data.SqlClient
Imports common
Imports System.IO
Imports common.UserControls

Public Class frmDBTNEFTUploader
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim IsinsideLoadData As Boolean = False
    Dim Qry As String
    Dim isFlag As Boolean = False
    Dim arrLoc As String = Nothing
    Dim SettMPIncentiveEntryApplyMonthly As Boolean = False
    Dim SettMPIncentiveEntryCycleWiseButNEFTMonthly As Boolean = False
    Dim SettMPIncentiveEntryIncentiveRate As Decimal = 0
    Dim SettDCSMPIncetiveReco As Boolean = False
    Dim SettApplyZoneOnDBT As Boolean = False
    Dim dtPerforma As DataTable
    Dim settMaxRowExport As Integer = 0
    Dim SettMCCOneDBTOneDoc As String = ""
    Dim DBTRevisePayment As Boolean = False

    Const PK_Id As String = "PK_Id"
    Const SNo As String = "SNO"
    Const UnionId As String = "UnionId"
    Const UnionName As String = "UnionName"
    Const SocietyID As String = "SocietyID"
    Const SocietyName As String = "SocietyName"
    Const FarmerID As String = "FarmerID"
    Const FarmerName As String = "FarmerName"
    Const FarmerContactNumber As String = "FarmerContactNumber"
    Const FarmerEmailID As String = "FarmerEmailID"
    Const FarmerAccountNumber As String = "FarmerAccountNumber"
    Const FarmerIFSCCode As String = "FarmerIFSCCode"
    'Const FarmerIFCSCode As String = "FarmerIFCSCode"
    Const FarmerBankName As String = "FarmerBankName"
    Const FarmerBankBranch As String = "FarmerBankBranch"
    Const AMOUNT As String = "AMOUNT"
    Const SocietySubName As String = "SocietySubName"
    Const AddInfo1 As String = "AddInfo1"
    Const AddInfo2 As String = "AddInfo2"
    Const AddInfo3 As String = "AddInfo3"
    Const AddInfo4 As String = "AddInfo4"
    Const DBTTRNO As String = "DBTTRNO"
    Const FarmerID2 As String = "FarmerID2"
    Const ZoneName As String = "ZoneName"
#End Region
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub FrmVLCDataUploaderManual_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DBTRevisePayment = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DBTRevisePayment, clsFixedParameterCode.DBTRevisePayment, Nothing)) = 1, True, False)

        If clsCommon.myCBool(DBTRevisePayment) Then
            chkDBTRevisePayment.Visible = True
        End If
        Qry = "select TSPL_BANK_MASTER_NEFT_PERFORMA.* from TSPL_BANK_MASTER_NEFT_PERFORMA 
left outer join TSPL_BANK_MASTER on TSPL_BANK_MASTER.BANK_CODE=TSPL_BANK_MASTER_NEFT_PERFORMA.BANK_CODE
where TSPL_BANK_MASTER.NEFT_DBT_Default=1 order by TRCode"
        dtPerforma = clsDBFuncationality.GetDataTable(Qry)
        If dtPerforma Is Nothing OrElse dtPerforma.Rows.Count <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please First Create DBT NEFT Performa With Default DBT NEFT Bank", Me.Text)
            Me.Close()
        Else
            Dim dtDefault As DataTable = clsDBTNEFTPerforma.GetDefault()
            For Each dr As DataRow In dtDefault.Rows
                Dim flag As Boolean = False
                For Each drPer As DataRow In dtPerforma.Rows
                    If clsCommon.CompairString(clsCommon.myCstr(dr("Code")), clsCommon.myCstr(drPer("NEFT_Col_Code"))) = CompairStringResult.Equal Then
                        flag = True
                        Exit For
                    End If
                Next
                If Not flag Then
                    clsCommon.MyMessageBoxShow(Me, "Please Set [" + clsCommon.myCstr(dr("Code")) + "] in Default DBT NEFT Performa", Me.Text)
                    Me.Close()
                    Exit Sub
                End If
            Next
        End If

        SplitContainer3.Panel2Collapsed = True
        SettApplyZoneOnDBT = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyZoneInDBT, clsFixedParameterCode.ApplyZoneInDBT, Nothing)) > 0)
        txtZone.MendatroryField = SettApplyZoneOnDBT
        SettMCCOneDBTOneDoc = clsFixedParameter.GetData(clsFixedParameterType.AndroidAPP, clsFixedParameterCode.OneDBTOneDoc, Nothing)
        SettMPIncentiveEntryApplyMonthly = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, Nothing))
        SettMPIncentiveEntryCycleWiseButNEFTMonthly = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryCycleWiseButNEFTMonthly, clsFixedParameterCode.MPIncentiveEntryCycleWiseButNEFTMonthly, Nothing))
        SettMPIncentiveEntryIncentiveRate = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryIncentiveRate, clsFixedParameterCode.MPIncentiveEntryIncentiveRate, Nothing))
        SettDCSMPIncetiveReco = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MandatoryDCSMPIncetiveReco, Nothing)) = 1)
        settMaxRowExport = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MaxRowsExcelDBTNEFTUploader, clsFixedParameterCode.MaxRowsExcelDBTNEFTUploader, Nothing))

        UcAttachment1.Form_ID = "A" + MyBase.Form_ID
        UcAttachment1.isDeleteTheAttachment = False
        UcAttachment1.settAutoAttachment = True
        'RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Collapsed

        UcAttachment2.Form_ID = MyBase.Form_ID
        UcAttachment2.MandatoryPDFFile = True
        SetUserMgmtNew()
        Reset()


        ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Transaction")
        ButtonToolTip.SetToolTip(btndelete, "Press Alt+D Delete Transaction")
        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnReset, "Press Alt+N New Transaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        If SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
            RadPageView1.SelectedPage = RadPageViewPage3
        Else
            RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Collapsed
            RadPageView1.SelectedPage = RadPageViewPage1
        End If
        If clsCommon.myLen(Me.Tag) > 0 Then
            LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
        End If
        Me.Focus()
        txtdate.Focus()
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
            btnPrint.Visible = True
        Else
            btnPrint.Visible = False
        End If

        'btnPrint.Visible = True
        BtnBank.Visible = False

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
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnclose.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            Dim frm As New frmPWDHighSecrity(Nothing)
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                btnReverse.Visible = True
                btnClrApproval.Visible = True
                SplitContainer3.Panel2Collapsed = False
            End If
        End If
    End Sub
    Private Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub



    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btndelete.Visible = MyBase.isDeleteFlag
        btnNEFTUploader.Visible = MyBase.isPrintFlag
        If lblPending.Status = ERPTransactionStatus.Pending Then
            btnNEFTUploader.Visible = False
        Else
            btnNEFTUploader.Visible = True
        End If
        btnPost.Visible = MyBase.isPostFlag
        btnReverse.Enabled = MyBase.isReverse
    End Sub
    Sub Reset()
        'loadBlankGrid()
        gvItem.DataSource = Nothing
        gvInvalid.DataSource = Nothing
        gvHold.DataSource = Nothing
        gvFarmer.DataSource = Nothing

        Dim dt As Date = clsCommon.GETSERVERDATE()
        txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dt) & "/" & DatePart(DateInterval.Year, dt)
        If SettMPIncentiveEntryApplyMonthly OrElse SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
            txtToDate.Value = txtFromDate.Value.AddMonths(1).AddDays(-1)
        Else
            txtToDate.Value = "01/" & DatePart(DateInterval.Month, dt) & "/" & DatePart(DateInterval.Year, dt)
        End If
        txtBankLetterDate.Value = dt
        txtDocumentNo.Value = ""
        txtdate.Value = dt
        chkDBTRevisePayment.Checked = False
        txtDocumentNo.MyReadOnly = False
        btnsave.Text = "Save"
        txtMCC.arrValueMember = Nothing
        txtVLC.arrValueMember = Nothing
        btndelete.Enabled = False
        btnsave.Enabled = True
        btnPost.Enabled = False
        txtdate.Focus()
        EnableInputDataField()
        isNewEntry = True
        IsinsideLoadData = False
        lblPending.Status = ERPTransactionStatus.Pending
        txtRemarks.Text = ""
        txtZone.Value = ""
        lblZone.Text = ""
        If clsCommon.myLen(SettMCCOneDBTOneDoc) > 0 Then
            Dim arr As New ArrayList()
            arr.Add(SettMCCOneDBTOneDoc)
            txtMCC.arrValueMember = arr
        End If
        gvItem.MasterTemplate.SummaryRowsBottom.Clear()
        UcAttachment1.BlankAllControls()
        UcAttachment2.BlankAllControls()
        UcAttachment2.isDeleteTheAttachment = True
        BtnBank.Enabled = False
    End Sub

    Private Function AllowToSave() As Boolean
        'Xtra.TransactionValidity(txtdate.Value)
        clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocumentNo.Value))
        If clsCommon.GetDateWithStartTime(txtBankLetterDate.Value) < clsCommon.GetDateWithStartTime(txtdate.Value) Then
            txtBankLetterDate.Focus()
            Throw New Exception("Bank Date Not Valid always greater then Document Date ")
        End If
        Return True
    End Function
    Sub SaveData()
        Try
            If AllowToSave() Then
                Dim obj As New clsDBTNEFT()
                obj.Document_Code = txtDocumentNo.Value
                obj.Document_Date = txtdate.Value
                obj.DBT_Revise_Payment = IIf(chkDBTRevisePayment.Checked = True, 1, 0)
                obj.From_Date = txtFromDate.Value
                obj.To_Date = txtToDate.Value
                obj.Remarks = txtRemarks.Text
                obj.Zone_Code = txtZone.Value
                obj.Bank_Letter_Date = txtBankLetterDate.Value
                If gvItem.Rows.Count > 0 Then
                    obj.arr = New List(Of clsDBTNEFTDetail)
                    For Each grow As GridViewRowInfo In gvItem.Rows
                        Dim objTr As New clsDBTNEFTDetail
                        objTr.SNo = obj.arr.Count + 1
                        objTr.Against_MP_Incentive_TR = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colAgainstMPIncetive).Value)

                        objTr.VLC_Uploader_Code = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colSociety).Value)
                        objTr.MP_Uploader_Code = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPUploaderCode).Value)
                        objTr.Amount = clsCommon.myCdbl(grow.Cells(clsDBTNEFTPerforma.colAmount).Value)
                        objTr.MP_IFSC_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPIFSCCode).Value)
                        objTr.MP_Account_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPAccountNo).Value)
                        objTr.MP_Bank = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPBank).Value)
                        objTr.MP_Mobile_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPMobileNo).Value)
                        objTr.MP_Name = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPName).Value)
                        'objTr.Transaction = clsCommon.myCstr(grow.Cells(colTransaction).Value)
                        obj.arr.Add(objTr)
                    Next
                End If
                If gvInvalid.Rows.Count > 0 Then
                    obj.arrInvalid = New List(Of clsDBTNEFTDetailInvalid)
                    For Each grow As GridViewRowInfo In gvInvalid.Rows
                        Dim objTr As New clsDBTNEFTDetailInvalid
                        objTr.SNo = obj.arrInvalid.Count + 1
                        objTr.Against_MP_Incentive_TR = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colAgainstMPIncetive).Value)
                        objTr.VLC_Uploader_Code = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colSociety).Value)
                        objTr.MP_Uploader_Code = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPUploaderCode).Value)
                        objTr.Amount = clsCommon.myCdbl(grow.Cells(clsDBTNEFTPerforma.colAmount).Value)
                        objTr.MP_IFSC_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPIFSCCode).Value)
                        objTr.MP_Account_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPAccountNo).Value)
                        objTr.MP_Bank = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPBank).Value)
                        objTr.MP_Mobile_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPMobileNo).Value)
                        objTr.MP_Name = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPName).Value)
                        obj.arrInvalid.Add(objTr)
                    Next
                End If
                If gvHold.Rows.Count > 0 Then
                    obj.arrHold = New List(Of clsDBTNEFTDetailHold)
                    For Each grow As GridViewRowInfo In gvHold.Rows
                        Dim objTr As New clsDBTNEFTDetailHold
                        objTr.SNo = obj.arrHold.Count + 1
                        objTr.Against_MP_Incentive_TR = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colAgainstMPIncetive).Value)
                        objTr.VLC_Uploader_Code = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colSociety).Value)
                        objTr.MP_Uploader_Code = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPUploaderCode).Value)
                        objTr.Amount = clsCommon.myCdbl(grow.Cells(clsDBTNEFTPerforma.colAmount).Value)
                        objTr.MP_IFSC_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPIFSCCode).Value)
                        objTr.MP_Account_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPAccountNo).Value)
                        objTr.MP_Bank = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPBank).Value)
                        objTr.MP_Mobile_No = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPMobileNo).Value)
                        objTr.MP_Name = clsCommon.myCstr(grow.Cells(clsDBTNEFTPerforma.colMPName).Value)
                        obj.arrHold.Add(objTr)
                    Next
                End If
                Dim objApproval As New clsApply_Approval()
                If (clsDBTNEFT.SaveData(obj, isNewEntry)) Then
                    Dim qry As String = "select sum(Amount) from TSPL_DBT_NEFT_DETAIL where Document_Code='" + txtDocumentNo.Value + "' "
                    Dim TotAmt As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry))
                    clsApply_Approval.CheckApprovalRequired(MyBase.Form_ID, obj.Document_Code, txtdate.Text, "", clsCommon.myCstr(txtRemarks.Text), clsCommon.myCdbl(TotAmt), 0, "", objApproval)
                    qry = "select 1 from TSPL_APPROVAL_LEVEL_TRANSACTION_DETAIL where TRANS_Code='" + MyBase.Form_ID + "' and Document_Code='" + obj.Document_Code + "'"
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        UcAttachment1.BlankAllControls()
                        Dim objP As New clsDBTNEFT()
                        Dim strCycle As String = ""

                        Dim Filename As String = objP.funPrintBankLetter(obj.Document_Code, True)
                        Dim SafeFileName As String = "BankLetter.pdf"
                        UcAttachment1.AddAttachment(Filename, SafeFileName)

                        gvItem.FilterDescriptors.Clear()
                        For Each col As Telerik.WinControls.UI.GridViewDataColumn In gvItem.Columns
                            col.FilterDescriptor = Nothing
                            col.ExcelExportType = Telerik.WinControls.UI.Export.DisplayFormatType.None
                        Next

                        Filename = clsCommon.MyExportToExcelGridPath("NEFT Uploader", gvItem, Nothing, Me.Text, False, "", "")
                        SafeFileName = "NEFTDetail.xls"
                        UcAttachment1.AddAttachment(Filename, SafeFileName)

                        UcAttachment1.SaveData(obj.Document_Code, True, Nothing)
                    End If

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
            clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocumentNo.Value))
            If (deleteConfirm()) Then
                If (clsDBTNEFT.DeleteData(txtDocumentNo.Value)) Then
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
        Dim obj As clsDBTNEFT = clsDBTNEFT.GetData(strCode, NavTyep)
        If obj IsNot Nothing Then
            DisableInputDataField()
            'BtnBank.Enabled = True
            isNewEntry = False
            txtDocumentNo.Value = obj.Document_Code
            txtdate.Value = obj.Document_Date
            txtFromDate.Value = obj.From_Date
            txtToDate.Value = obj.To_Date
            lblPending.Status = obj.Status
            txtRemarks.Text = obj.Remarks
            txtZone.Value = obj.Zone_Code
            chkDBTRevisePayment.Checked = IIf(obj.DBT_Revise_Payment = 1, True, False)
            lblZone.Text = ClsZoneMaster.GetName(obj.Zone_Code)
            txtMCC.arrValueMember = obj.arrMCC
            txtVLC.arrValueMember = obj.arrVLC
            txtBankLetterDate.Value = obj.Bank_Letter_Date
            If obj.arr IsNot Nothing AndAlso obj.arr.Count > 0 Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False))
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    gvItem.DataSource = dt
                    FormatGrid(gvItem)
                End If
                If SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
                    Dim dtFarmer As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", True))
                    If dtFarmer IsNot Nothing AndAlso dtFarmer.Rows.Count > 0 Then
                        gvFarmer.DataSource = dtFarmer
                        FormatGrid(gvFarmer)
                    End If
                End If
            End If

            If obj.arrInvalid IsNot Nothing AndAlso obj.arrInvalid.Count > 0 Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL_INVALID", False))
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    gvInvalid.DataSource = dt
                    FormatGrid(gvInvalid)
                End If
            End If

            If True Then
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL_HOLD", False))
                gvHold.DataSource = dt
                FormatGrid(gvHold)
            End If


            txtDocumentNo.MyReadOnly = True
            If obj.Status = ERPTransactionStatus.Approved Then
                btnsave.Enabled = False
                btndelete.Enabled = False
                btnPost.Enabled = False
                btnNEFTUploader.Visible = True
                UcAttachment2.isDeleteTheAttachment = False
                BtnBank.Enabled = True
            Else
                btnsave.Text = "Update"
                btnsave.Enabled = True
                btndelete.Enabled = True
                btnPost.Enabled = True
                BtnBank.Enabled = False
                UcAttachment2.isDeleteTheAttachment = True
            End If
            UcAttachment1.LoadData(obj.Document_Code)
            UcAttachment2.LoadData(obj.Document_Code)
        End If
        IsinsideLoadData = False
    End Sub


    Private Sub FormatGrid(ByVal gv As RadGridView)
        Try
            gv.MasterTemplate.SummaryRowsBottom.Clear()
            For ii As Integer = 0 To gv.Columns.Count - 1
                gv.Columns(ii).ReadOnly = True
                gv.Columns(ii).FormatString = ""
                gv.Columns(ii).BestFit()
            Next
            If gv.Columns.Contains("PK_Id") Then
                gv.Columns("PK_Id").IsVisible = False
            End If

            For ii As Integer = 0 To dtPerforma.Rows.Count - 1
                If clsCommon.CompairString(gv.Name, "gvFarmer") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Name")), clsDBTNEFTPerforma.colAgainstMPIncetive) = CompairStringResult.Equal Then
                        Continue For
                    End If
                End If
                gv.Columns(clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code"))).HeaderText = clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Name"))
                gv.Columns(clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code"))).IsVisible = Not clsCommon.myCBool(dtPerforma.Rows(ii)("NEFT_Col_Hide"))
            Next


            gv.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            gv.Columns("THIRD_PARTY_CODE").ReadOnly = True

            gv.Columns("THIRD_PARTY_CODE").IsVisible = False
            gv.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True

            'gvFarmer.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            'gvFarmer.Columns("THIRD_PARTY_CODE").IsVisible = False
            'gvFarmer.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True



            gv.Columns(clsDBTNEFTPerforma.colAmount).FormatString = "{0:n2}"
            gv.Columns(clsDBTNEFTPerforma.colAmount).TextAlignment = System.Drawing.ContentAlignment.MiddleRight

            gv.BestFitColumns()
            gv.AllowAddNewRow = False
            gv.AllowDeleteRow = True
            gv.AllowRowReorder = False
            gv.ShowGroupPanel = False
            gv.EnableFiltering = True
            gv.ShowFilteringRow = True
            gv.EnableSorting = False
            gv.EnableGrouping = False
            gv.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
            gv.GridBehavior = New MyBehavior()

            'Add Summary row for Amount and S.No'
            Dim summaryRowItem As New GridViewSummaryRowItem()
            gv.Columns("SNo").FormatString = ""
            Dim item1 As New GridViewSummaryItem("SNo", "", GridAggregateFunction.Count)
            summaryRowItem.Add(item1)

            gv.Columns("AMOUNT").FormatString = "{0:n2}"
            Dim item11 As New GridViewSummaryItem("AMOUNT", "{0:n2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item11)

            gv.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            gv.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
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
            txtDocumentNo.Value = clsDBTNEFT.getFinder("", txtDocumentNo.Value, isButtonClicked)
            LoadData(txtDocumentNo.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub DisableInputDataField()
        txtFromDate.Enabled = False
        txtToDate.Enabled = False
    End Sub
    Sub EnableInputDataField()
        txtFromDate.Enabled = True
        txtToDate.Enabled = True
    End Sub

    'Private Sub btnclose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnclose.Click
    '    CloseForm()
    'End Sub

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
            gvItem.CurrentColumn = gvItem.Columns(clsDBTNEFTPerforma.colMPUploaderCode)
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
            gvItem.CurrentColumn = gvItem.Columns(clsDBTNEFTPerforma.colMPUploaderCode)
            gvItem.CurrentRow = gvItem.Rows(gvItem.Rows.Count - 1)
            gvItem.PerformLayout()
            gvItem.BeginEdit()
        End If
    End Sub


    Private Sub txtFromDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFromDate.Validating
        SetToDate()
    End Sub
    Private Sub txtFromDate_Leave(sender As Object, e As EventArgs) Handles txtFromDate.Leave
        SetToDate()
    End Sub
    Sub SetToDate()
        Try
            If Not IsinsideLoadData Then
                Dim PaymentCycleType As String = ""
                Dim PaymentCycleValue As Integer = 0
                Dim arrMCC As ArrayList = txtMCC.arrValueMember
                If arrMCC Is Nothing OrElse arrMCC.Count <= 0 Then
                    Dim str As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 MCC_Code from TSPL_MCC_MASTER"))
                    arrMCC = New ArrayList()
                    arrMCC.Add(str)
                End If
                If SettMPIncentiveEntryApplyMonthly OrElse SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
                    txtFromDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1)
                    txtToDate.Value = txtFromDate.Value.AddMonths(1).AddDays(-1)
                Else
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(" select Payment_Cycle,PC_TYPE,PC_VALUE from ( select TSPL_MCC_MASTER.Payment_Cycle,TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE  from TSPL_MCC_MASTER left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle   where TSPL_MCC_MASTER.MCC_Code  in (" + clsCommon.GetMulcallString(arrMCC) + ") ) xx group by Payment_Cycle,PC_TYPE,PC_VALUE")
                    If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                        Throw New Exception("No Payment Cycle found on current MCC/Location")
                    End If
                    If dt.Rows.Count > 1 Then
                        Throw New Exception("Selected MCC's Payment Cycle Should be Same")
                    End If
                    PaymentCycleType = clsCommon.myCstr(dt.Rows(0)("PC_TYPE"))
                    PaymentCycleValue = clsCommon.myCdbl(dt.Rows(0)("PC_VALUE"))
                    Dim dtCurr As DateTime = clsCommon.GETSERVERDATE()
                    If clsCommon.CompairString(PaymentCycleType, "Day") = CompairStringResult.Equal Then
                        If txtFromDate.Value.Day Mod PaymentCycleValue <> 1 And (Not PaymentCycleValue = 1) Then
                            clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month or at interval of " & PaymentCycleValue & " Day, Because MCC has payment Cycle of " & PaymentCycleValue & " Day ", Me.Text)
                            txtFromDate.Value = New Date(dtCurr.Year, dtCurr.Month, 1)
                            txtToDate.Value = txtFromDate.Value
                            Exit Sub
                        End If
                        txtToDate.Value = txtFromDate.Value.AddDays(PaymentCycleValue - 1)

                        If txtFromDate.Value.Month <> txtToDate.Value.Month Then
                            txtToDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
                        End If
                        Dim dtNxtPay As DateTime = txtToDate.Value.AddDays(Math.Ceiling(PaymentCycleValue / 2.0))
                        If txtFromDate.Value.Month <> dtNxtPay.Month Then
                            txtToDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
                        End If
                    ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(clsCommon.GetPrintDate(txtFromDate.Value, "dd")) <> 1 Then
                            clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month, Because MCC has payment Cycle of Month Type", Me.Text)
                            txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                            txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                            Exit Sub
                        End If
                        txtToDate.Value = DateAdd(DateInterval.Month, PaymentCycleValue, txtFromDate.Value)
                    ElseIf clsCommon.CompairString(PaymentCycleType, "Year") = CompairStringResult.Equal Then
                        If clsCommon.myCdbl(clsCommon.GetPrintDate(txtFromDate.Value, "dd")) <> 1 Then
                            clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month, Because MCC has payment Cycle of Year Type", Me.Text)
                            txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                            txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                            Exit Sub
                        End If
                        txtToDate.Value = DateAdd(DateInterval.Year, PaymentCycleValue, txtFromDate.Value)
                    ElseIf clsCommon.CompairString(PaymentCycleType, "Week") = CompairStringResult.Equal Then
                        Dim today As Date = txtFromDate.Value
                        Dim dayDiff As Integer = today.DayOfWeek - IIf(PaymentCycleValue = 1, DayOfWeek.Sunday, IIf(PaymentCycleValue = 2, DayOfWeek.Monday, IIf(PaymentCycleValue = 3, DayOfWeek.Tuesday, IIf(PaymentCycleValue = 4, DayOfWeek.Wednesday, IIf(PaymentCycleValue = 5, DayOfWeek.Thursday, IIf(PaymentCycleValue = 6, DayOfWeek.Friday, DayOfWeek.Saturday))))))
                        txtFromDate.Value = today.AddDays(-dayDiff)
                        txtToDate.Value = txtFromDate.Value.AddDays(6)
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
    Private Sub mtxtVLC__My_Click(sender As Object, e As EventArgs) Handles txtVLC._My_Click
        If txtMCC.arrValueMember Is Nothing OrElse txtMCC.arrValueMember.Count <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please Select MCC First.", Me.Text)
            txtMCC.Focus()
            Exit Sub
        End If
        If SettApplyZoneOnDBT Then
            If clsCommon.myLen(txtZone.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select " + txtZone.MyLinkLable2.Text, Me.Text)
                txtZone.Focus()
                Exit Sub
            End If
        End If
        Dim qry As String = "Select VLC.VLC_Code_vlc_Uploader as [Code],VLC.VLC_Code as [DCS Code],VLC.VLC_Name as [DCS Name],VLC.MCC as [MCC Code],VLC.Route_Code as [Route Code],RM.Route_Name ,TSPL_VENDOR_MASTER.Zone_Code as ZoneCode, TSPL_ZONE_MASTER.Description as ZoneName
from TSPL_VLC_MASTER_HEAD    VLC left join TSPL_MCC_ROUTE_MASTER RM on vlc.Route_Code=RM.Route_Code 
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=VLC.VSP_Code
left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
where 2=2 "
        If clsCommon.myLen(SettMCCOneDBTOneDoc) <= 0 Then
            qry += " and VLC.MCC in (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ")"
        End If
        If SettApplyZoneOnDBT Then
            qry += "and TSPL_VENDOR_MASTER.Zone_Code='" + txtZone.Value + "' "
        End If
        txtVLC.arrValueMember = clsCommon.ShowMultipleSelectForm("dbtneftv", qry, "DCS Code", "DCS Name", txtVLC.arrValueMember, txtVLC.arrDispalyMember)


    End Sub
    Sub fillMPS()
        Try
            'If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
            If True Then
                Dim flag As Boolean = True
                Dim dtCheck As DataTable = clsDBFuncationality.GetDataTable(GetMpQry(True, False, True))
                If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                    If common.clsCommon.MyMessageBoxShow(Me, "Error in " & dtCheck.Rows.Count & " Records.Do you want to open it", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes Then
                        Dim frm As New FrmFreeGrid
                        frm.ReportID = "DBTNeftError"
                        frm.Text = "Record Could not Loaded"
                        frm.dt = dtCheck
                        frm.ShowDialog()
                    End If
                    If clsCommon.MyMessageBoxShow(Me, "Do you want to continue? Your input will be marked as invalid for above data", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.No Then
                        flag = False
                    End If
                    gvItem.DataSource = Nothing
                    gvFarmer.DataSource = Nothing
                    gvInvalid.DataSource = Nothing
                    gvHold.DataSource = Nothing
                End If

                If flag Then
                    gvItem.DataSource = Nothing
                    gvFarmer.DataSource = Nothing
                    gvInvalid.DataSource = Nothing
                    gvHold.DataSource = Nothing
                    'loadBlankGrid()
                    gvItem.DataSource = Nothing
                    Dim isDataFound As Boolean = False
                    Dim dtValid As DataTable = clsDBFuncationality.GetDataTable(GetMpQry(True, False, False))
                    If dtValid IsNot Nothing AndAlso dtValid.Rows.Count > 0 Then
                        isDataFound = True
                        gvItem.DataSource = dtValid
                        FormatGrid(gvItem)
                    End If

                    gvFarmer.DataSource = Nothing
                    If SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
                        Dim dtFarmerWiseValid As DataTable = clsDBFuncationality.GetDataTable(GetMpQry(True, True, False))
                        If dtFarmerWiseValid IsNot Nothing AndAlso dtFarmerWiseValid.Rows.Count > 0 Then
                            isDataFound = True
                            gvFarmer.DataSource = dtFarmerWiseValid
                            FormatGrid(gvFarmer)
                        End If
                    End If

                    gvInvalid.DataSource = Nothing
                    Dim dtInValid As DataTable = clsDBFuncationality.GetDataTable(GetMpQry(False, False, False))
                    If dtInValid IsNot Nothing AndAlso dtInValid.Rows.Count > 0 Then
                        isDataFound = True
                        gvInvalid.DataSource = dtInValid
                        FormatGrid(gvInvalid)
                    End If

                    'gvHold.DataSource = Nothing
                    'Dim dtInValid As DataTable = clsDBFuncationality.GetDataTable(GetMpQry(False, False))
                    'If dtInValid IsNot Nothing AndAlso dtInValid.Rows.Count > 0 Then
                    '    isDataFound = True
                    '    gvHold.DataSource = dtInValid
                    '    FormatGrid(gvHold)
                    'End If

                    If Not isDataFound Then
                        clsCommon.MyMessageBoxShow(Me, "No Data found", Me.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function GetQry(ByVal TableName As String, ByVal GrpByFarmer As Boolean) As String
        Return GetQry(TableName, GrpByFarmer, Nothing, Nothing)
    End Function
    Function GetQry(ByVal TableName As String, ByVal GrpByFarmer As Boolean, ByVal lstMPCode As List(Of String), ByVal lstAccCode As List(Of String)) As String
        Dim qry As String = "Select  " + TableName + ".PK_Id, " + TableName + ".SNo as [" + clsDBTNEFTPerforma.colSlNo + "]," + TableName + ".Against_MP_Incentive_TR AS [" + clsDBTNEFTPerforma.colAgainstMPIncetive + "]," + TableName + ".VLC_Uploader_Code AS [" + clsDBTNEFTPerforma.colSociety + "]
                ," + TableName + ".MP_Uploader_Code AS [" + clsDBTNEFTPerforma.colMPUploaderCode + "]"

        'If clsCommon.CompairString(objCommonVar.CurrDatabase, "BKN") = CompairStringResult.Equal Then
        '    qry += " ,CASE  WHEN (TSPL_DBT_NEFT_DETAIL.Amount) % 1 >= 0.5    THEN CEILING(TSPL_DBT_NEFT_DETAIL.Amount)  Else FLOOR(TSPL_DBT_NEFT_DETAIL.Amount) End As Amount "
        'Else
        '    qry += " ," + TableName + ".Amount AS [" + clsDBTNEFTPerforma.colAmount + "]"
        'End If
        qry += " ," + TableName + ".Amount AS [" + clsDBTNEFTPerforma.colAmount + "]"
        qry += "," + TableName + ".MP_IFSC_No AS [" + clsDBTNEFTPerforma.colMPIFSCCode + "]
," + TableName + ".MP_Bank AS [" + clsDBTNEFTPerforma.colMPBank + "]," + TableName + ".MP_Mobile_No AS [" + clsDBTNEFTPerforma.colMPMobileNo + "]
," + TableName + ".MP_Account_No AS [" + clsDBTNEFTPerforma.colMPAccountNo + "]," + TableName + ".MP_Name AS [" + clsDBTNEFTPerforma.colMPName + "]
,TSPL_VLC_MASTER_HEAD.VLC_Name as [" + clsDBTNEFTPerforma.colSocietyName + "],TSPL_ZONE_MASTER.Description as [" + clsDBTNEFTPerforma.colZoneName + "],TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code as [" + clsDBTNEFTPerforma.colFarmerCode + "],tspl_mp_master.THIRD_PARTY_CODE as [" + clsDBTNEFTPerforma.colTHIRDPARTYCODE + "]
from " + TableName + " 
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=" + TableName + ".Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VLC_MASTER_HEAD.VSP_Code
left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
left outer join tspl_mp_master on tspl_mp_master.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code

where " + TableName + ".Document_Code='" & txtDocumentNo.Value & "'"
        If lstMPCode IsNot Nothing AndAlso lstMPCode.Count > 0 Then
            qry += " and TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code in (" & clsCommon.GetMulcallString(lstMPCode) & ")"
        End If
        If lstAccCode IsNot Nothing AndAlso lstAccCode.Count > 0 Then
            qry += " and TSPL_DBT_NEFT_DETAIL.MP_Account_No in (" & clsCommon.GetMulcallString(lstAccCode) & ")"
        End If
        If GrpByFarmer Then
            qry = "select min(PK_Id) as PK_Id,ROW_NUMBER() OVER (ORDER BY [" + clsDBTNEFTPerforma.colFarmerCode + "]) AS [" + clsDBTNEFTPerforma.colSlNo + "],[" + clsDBTNEFTPerforma.colFarmerCode + "],max([" + clsDBTNEFTPerforma.colSociety + "]) as [" + clsDBTNEFTPerforma.colSociety + "],max([" + clsDBTNEFTPerforma.colMPUploaderCode + "]) as [" + clsDBTNEFTPerforma.colMPUploaderCode + "],sum([" + clsDBTNEFTPerforma.colAmount + "]) as [" + clsDBTNEFTPerforma.colAmount + "],max([" + clsDBTNEFTPerforma.colMPIFSCCode + "]) as [" + clsDBTNEFTPerforma.colMPIFSCCode + "],max([" + clsDBTNEFTPerforma.colMPAccountNo + "]) as [" + clsDBTNEFTPerforma.colMPAccountNo + "],max([" + clsDBTNEFTPerforma.colMPBank + "]) as [" + clsDBTNEFTPerforma.colMPBank + "],max([" + clsDBTNEFTPerforma.colMPMobileNo + "]) as [" + clsDBTNEFTPerforma.colMPMobileNo + "],max([" + clsDBTNEFTPerforma.colMPName + "]) as [" + clsDBTNEFTPerforma.colMPName + "],max([" + clsDBTNEFTPerforma.colSocietyName + "]) as [" + clsDBTNEFTPerforma.colSocietyName + "],max([" + clsDBTNEFTPerforma.colZoneName + "]) as [" + clsDBTNEFTPerforma.colZoneName + "] from (" + qry + ")xx group by  [" + clsDBTNEFTPerforma.colFarmerCode + "]"
        End If
        Dim strMain As String = "Select PK_Id,"
        For ii As Integer = 0 To dtPerforma.Rows.Count - 1
            If GrpByFarmer Then
                If clsCommon.CompairString(clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")), clsDBTNEFTPerforma.colAgainstMPIncetive) = CompairStringResult.Equal Then
                    Continue For
                End If
            End If
            If clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")).Contains("CR") Then
                strMain += "'" + clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Value")) + "' as [" + clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")) + "]"
                If ii <> dtPerforma.Rows.Count - 1 Then
                    strMain += ","
                End If
            Else
                strMain += "[" + clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")) + "] "
                If ii <> dtPerforma.Rows.Count - 1 Then
                    strMain += ","
                End If
            End If
        Next
        strMain += " ,[Third_Party_Code] from (" + qry + ")xxx order by [" + clsDBTNEFTPerforma.colSlNo + "]"

        Return strMain
    End Function

    Private Function GetMpQry(ByVal IsPickValid As Boolean, ByVal GrpByFarmer As Boolean, ByVal isCheckOnly As Boolean) As String

        Dim BaseQry As String = "select TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id,TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code as Doc_No,convert(varchar, TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date,103) as Date_Range,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code,tspl_MCC_Master.MCC_Name,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MP_MASTER.MP_Code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as VLC_CODE_Uploader,TSPL_MP_MASTER.PayeeName as Payee_Joint_Name,TSPL_MP_MASTER.BankName as Bank_Code,TSPL_MP_MASTER.BankName as Bank_Code_Desc,case when len(isnull(TSPL_MP_MASTER.Telphone,''))=10 then TSPL_MP_MASTER.Telphone else '' end as Telphone,TSPL_MP_MASTER.AccountNO as Payee_Joint_Account_No,TSPL_MP_MASTER.IFCICode as Payee_Joint_IFSC_Code,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Amount_Actual as Payable_Amount 
,TSPL_ZONE_MASTER.Description  as ZoneName,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid   
from TSPL_MP_INCENTIVE_ENTRY_DETAIL "
        If SettDCSMPIncetiveReco Then
            BaseQry += " left outer join TSPL_DCS_MP_INCENTIVE_RECO_DETAIL on TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Year=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Year and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Month=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Month and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_No=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_No and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code 
left outer join TSPL_DCS_MP_INCENTIVE_RECO_HEAD on TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code =TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code
left outer join TSPL_DBT_CAPING on TSPL_DBT_CAPING.Reco_Code=TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Document_Code
left outer join TSPL_DBT_CAPING_DETAIL on TSPL_DBT_CAPING_DETAIL.Document_Code=TSPL_DBT_CAPING.Document_Code and  TSPL_DBT_CAPING_DETAIL.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code"
        End If

        BaseQry += " left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
    left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
    left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
    left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VLC_MASTER_HEAD.VSP_Code
    left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
    left outer join tspl_MCC_Master on tspl_MCC_Master.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
    where 2=2 "
        If SettDCSMPIncetiveReco Then
            BaseQry += " and isnull(TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Status,0)=1 and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.PK_Id is not null 
and 2=(case when ISNULL(TSPL_DCS_MP_INCENTIVE_RECO_HEAD.DBT_Capping_Apply,0)=1 then ((case when ISNULL(TSPL_DBT_CAPING_DETAIL.Capping_Status,0)=1 then 2 else 3 end)) else 2 end)"
        End If
        If Not isCheckOnly Then
            If IsPickValid Then
                BaseQry += " and TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid=0"
            Else
                BaseQry += " and TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid>0"
            End If
        End If
        If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
            BaseQry += " and TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code in (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ") "
        End If
        BaseQry += " and  TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date >='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' "
        If txtVLC.arrValueMember IsNot Nothing AndAlso txtVLC.arrValueMember.Count > 0 Then
            BaseQry += "  and TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code in (" + clsCommon.GetMulcallString(txtVLC.arrValueMember) + ") "
        End If
        BaseQry += " and  TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date <='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' 
        and not exists(select 1 from TSPL_DBT_NEFT_DETAIL where TSPL_DBT_NEFT_DETAIL.Document_Code <> '" + txtDocumentNo.Value + "' and TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR=TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id
        and not exists(select 1 from TSPL_DBT_NEFT_REJECT_DETAIL inner join TSPL_DBT_NEFT_REJECT on TSPL_DBT_NEFT_REJECT.Document_Code=TSPL_DBT_NEFT_REJECT_DETAIL.Document_Code where TSPL_DBT_NEFT_REJECT_DETAIL.Against_DBT_NEFT_TR=TSPL_DBT_NEFT_DETAIL.PK_Id and TSPL_DBT_NEFT_REJECT.Status=1))"
        Dim strMain As String = ""
        If isCheckOnly Then
            Dim tempQty As String = "With CTE as ( " + BaseQry + ")
update TSPL_MP_INCENTIVE_ENTRY_DETAIL set  Mark_Invalid = 0 where PK_Id in (select PK_Id from CTE);"
            clsDBFuncationality.ExecuteNonQuery(tempQty)

            tempQty = "With CTE as ( " + BaseQry + ")
UPDATE TSPL_MP_INCENTIVE_ENTRY_DETAIL SET Mark_Invalid = 1 FROM ( 
SELECT  PK_Id, COUNT(*) OVER (PARTITION BY Payee_Joint_Account_No) AS cnt FROM CTE 
) c INNER JOIN TSPL_MP_INCENTIVE_ENTRY_DETAIL  ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id = c.PK_Id 
WHERE c.cnt > 1 AND TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid= 0;"
            clsDBFuncationality.ExecuteNonQuery(tempQty)

            tempQty = "With CTE as ( " + BaseQry + ")
SELECT  PK_Id, Payee_Joint_Name, dbo.RemoveExtraSpaces( UPPER(dbo.RemoveSpecialCharactersWithNumber(Payee_Joint_Name)) ) AS CleanName INTO #TempCTE FROM CTE;

UPDATE TSPL_MP_INCENTIVE_ENTRY_DETAIL SET Mark_Invalid = 2 FROM TSPL_MP_INCENTIVE_ENTRY_DETAIL INNER JOIN #TempCTE c  ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id = c.PK_Id
WHERE c.CleanName <> c.Payee_Joint_Name AND TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid = 0;"
            clsDBFuncationality.ExecuteNonQuery(tempQty)

            tempQty = "With CTE as ( " + BaseQry + ")
SELECT  PK_Id, Payee_Joint_Account_No, dbo.GetNumbersOnly(Payee_Joint_Account_No) AS CleanAccNo 
INTO #TempCTE FROM CTE;

UPDATE TSPL_MP_INCENTIVE_ENTRY_DETAIL SET Mark_Invalid = 3 FROM TSPL_MP_INCENTIVE_ENTRY_DETAIL INNER JOIN #TempCTE c  ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id = c.PK_Id 
WHERE TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid = 0 AND c.CleanAccNo <> c.Payee_Joint_Account_No;"
            clsDBFuncationality.ExecuteNonQuery(tempQty)

            tempQty = "With CTE as ( " + BaseQry + ")
UPDATE TSPL_MP_INCENTIVE_ENTRY_DETAIL SET Mark_Invalid = 4 FROM TSPL_MP_INCENTIVE_ENTRY_DETAIL JOIN CTE c  ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id = c.PK_Id 
WHERE TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid = 0 AND NOT EXISTS ( SELECT 1 FROM TSPL_MASTER.dbo.TSPL_IFSC i WHERE i.IFSC = c.Payee_Joint_IFSC_Code );"
            clsDBFuncationality.ExecuteNonQuery(tempQty)

            tempQty = "With CTE as ( " + BaseQry + ")
SELECT  PK_Id, Payee_Joint_IFSC_Code, LEN(ISNULL(Payee_Joint_Account_No, '')) AS AccLen
INTO #TempCTE FROM CTE;

UPDATE TSPL_MP_INCENTIVE_ENTRY_DETAIL SET Mark_Invalid = 5 FROM TSPL_MP_INCENTIVE_ENTRY_DETAIL
JOIN #TempCTE c  ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id = c.PK_Id 
LEFT JOIN TSPL_MASTER.dbo.TSPL_IFSC i  ON i.IFSC = c.Payee_Joint_IFSC_Code
LEFT JOIN TSPL_MASTER.dbo.IFSC_BANK_LENGTH bl  ON bl.BANK = i.BANK
WHERE TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mark_Invalid = 0 AND ( bl.BANK IS NULL OR c.AccLen < bl.MinLength OR c.AccLen > bl.MaxLength );"
            clsDBFuncationality.ExecuteNonQuery(tempQty)

            strMain = "With CTE as ( " + BaseQry + ")
select (case when Mark_Invalid=1 then 'Repeated Account No' 
when Mark_Invalid=2 then  'Special Character in Payee Name' 
when Mark_Invalid=3 then 'Special Character in Account No' 
when Mark_Invalid=4 then 'Wrong IFSC No' 
when Mark_Invalid=5 then 'Wrong Account No' end  ) as ErrorCode,
(case when Mark_Invalid=1 then Payee_Joint_Account_No
when Mark_Invalid=2 then  Payee_Joint_Name
when Mark_Invalid=3 then Payee_Joint_Account_No
when Mark_Invalid=4 then Payee_Joint_IFSC_Code
when Mark_Invalid=5 then Payee_Joint_Account_No end) as ErrorValue,VLC_CODE_Uploader as MPUploaderCode, MP_Code as MPCode, Payee_Joint_Name as MPPayeeName, VLC_Code_VLC_Uploader as DCSUploaderCode 
from  CTE   where Mark_Invalid>0 "

            'strMain = "With CTE as ( " + BaseQry + ")
            'select 'Repeated Account No' as ErrorCode,Payee_Joint_Account_No as ErrorValue,STRING_AGG(VLC_CODE_Uploader,',') as MPUploaderCode,STRING_AGG(MP_Code,',') as MPCode,STRING_AGG(Payee_Joint_Name,',') as MPPayeeName,STRING_AGG(VLC_Code_VLC_Uploader,',') as DCSUploaderCode  from  CTE  group by  Payee_Joint_Account_No having sum(1)>1
            'union all
            'select 'Special Character'as ErrorCode,Payee_Joint_Name as ErrorValue,VLC_CODE_Uploader as MPUploaderCode, MP_Code as MPCode, Payee_Joint_Name as MPPayeeName, VLC_Code_VLC_Uploader as DCSUploaderCode  from CTE where dbo.RemoveExtraSpaces(UPPER(dbo.RemoveSpecialCharactersWithNumber(Payee_Joint_Name))) <> Payee_Joint_Name
            'union all
            'select 'Wrong Account No' as ErrorCode,Payee_Joint_Account_No as ErrorValue,VLC_CODE_Uploader as MPUploaderCode, MP_Code as MPCode, Payee_Joint_Name as MPPayeeName, VLC_Code_VLC_Uploader as DCSUploaderCode from CTE where len(isnull(Payee_Joint_Account_No,''))<=8  or dbo.GetNumbersOnly(Payee_Joint_Account_No) <> Payee_Joint_Account_No
            'union all
            'select 'Wrong IFSC No' as ErrorCode,Payee_Joint_Account_No as ErrorValue,VLC_CODE_Uploader as MPUploaderCode, MP_Code as MPCode, Payee_Joint_Name as MPPayeeName, VLC_Code_VLC_Uploader as DCSUploaderCode from CTE where len(isnull(Payee_Joint_IFSC_Code,''))<>11 ;"

        Else
            Qry = "select   ROW_NUMBER() OVER (ORDER BY Bank_Code,MCC_Code,VLC_Code_VLC_Uploader) AS [" + clsDBTNEFTPerforma.colSlNo + "],MP_Code as [" + clsDBTNEFTPerforma.colFarmerCode + "],PK_Id as [" + clsDBTNEFTPerforma.colAgainstMPIncetive + "],VLC_Code_VLC_Uploader as [" + clsDBTNEFTPerforma.colSociety + "],VLC_CODE_Uploader as [" + clsDBTNEFTPerforma.colMPUploaderCode + "],Payable_Amount as [" + clsDBTNEFTPerforma.colAmount + "],Payee_Joint_IFSC_Code as [" + clsDBTNEFTPerforma.colMPIFSCCode + "],Payee_Joint_Account_No as [" + clsDBTNEFTPerforma.colMPAccountNo + "],Bank_Code as [" + clsDBTNEFTPerforma.colMPBank + "],Telphone as [" + clsDBTNEFTPerforma.colMPMobileNo + "],Payee_Joint_Name as [" + clsDBTNEFTPerforma.colMPName + "],Bank_Code,MCC_Code,VLC_Name as [" + clsDBTNEFTPerforma.colSocietyName + "],ZoneName as [" + clsDBTNEFTPerforma.colZoneName + "] from (" + BaseQry + ")xxx "
            If GrpByFarmer Then
                Qry = "select ROW_NUMBER() OVER (ORDER BY max(Bank_Code),max(MCC_Code),max([" + clsDBTNEFTPerforma.colMPUploaderCode + "])) AS [" + clsDBTNEFTPerforma.colSlNo + "],[" + clsDBTNEFTPerforma.colFarmerCode + "],max([" + clsDBTNEFTPerforma.colSociety + "]) as [" + clsDBTNEFTPerforma.colSociety + "],max([" + clsDBTNEFTPerforma.colMPUploaderCode + "]) as [" + clsDBTNEFTPerforma.colMPUploaderCode + "],sum([" + clsDBTNEFTPerforma.colAmount + "]) as [" + clsDBTNEFTPerforma.colAmount + "],max([" + clsDBTNEFTPerforma.colMPIFSCCode + "]) as [" + clsDBTNEFTPerforma.colMPIFSCCode + "],max([" + clsDBTNEFTPerforma.colMPAccountNo + "]) as [" + clsDBTNEFTPerforma.colMPAccountNo + "],max([" + clsDBTNEFTPerforma.colMPBank + "]) as [" + clsDBTNEFTPerforma.colMPBank + "],max([" + clsDBTNEFTPerforma.colMPMobileNo + "]) as [" + clsDBTNEFTPerforma.colMPMobileNo + "],max([" + clsDBTNEFTPerforma.colMPName + "]) as [" + clsDBTNEFTPerforma.colMPName + "],max(Bank_Code) as Bank_Code,max(MCC_Code) as MCC_Code,max([" + clsDBTNEFTPerforma.colSocietyName + "]) as [" + clsDBTNEFTPerforma.colSocietyName + "],max([" + clsDBTNEFTPerforma.colZoneName + "]) as [" + clsDBTNEFTPerforma.colZoneName + "] from  (
select MP_Code as [" + clsDBTNEFTPerforma.colFarmerCode + "],VLC_Code_VLC_Uploader as [" + clsDBTNEFTPerforma.colSociety + "],VLC_CODE_Uploader as [" + clsDBTNEFTPerforma.colMPUploaderCode + "],Payable_Amount as [" + clsDBTNEFTPerforma.colAmount + "],Payee_Joint_IFSC_Code as [" + clsDBTNEFTPerforma.colMPIFSCCode + "],Payee_Joint_Account_No as [" + clsDBTNEFTPerforma.colMPAccountNo + "],Bank_Code as [" + clsDBTNEFTPerforma.colMPBank + "],Telphone as [" + clsDBTNEFTPerforma.colMPMobileNo + "],Payee_Joint_Name as [" + clsDBTNEFTPerforma.colMPName + "],Bank_Code,MCC_Code,VLC_Name as [" + clsDBTNEFTPerforma.colSocietyName + "],ZoneName as [" + clsDBTNEFTPerforma.colZoneName + "] from (" + BaseQry + ")xx 
)xxx  group by  [" + clsDBTNEFTPerforma.colFarmerCode + "]"
            End If
            strMain = "Select "
            For ii As Integer = 0 To dtPerforma.Rows.Count - 1
                If GrpByFarmer Then
                    If clsCommon.CompairString(clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")), clsDBTNEFTPerforma.colAgainstMPIncetive) = CompairStringResult.Equal Then
                        Continue For
                    End If
                End If
                If clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")).Contains("CR") Then
                    strMain += "'" + clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Value")) + "' as [" + clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")) + "]"
                    If ii <> dtPerforma.Rows.Count - 1 Then
                        strMain += ","
                    End If
                Else
                    strMain += "[" + clsCommon.myCstr(dtPerforma.Rows(ii)("NEFT_Col_Code")) + "] "
                    If ii <> dtPerforma.Rows.Count - 1 Then
                        strMain += ","
                    End If
                End If
            Next
            strMain += " from (" + Qry + ")xxx " + Environment.NewLine
            strMain += "order by Bank_Code,MCC_Code,[" + clsDBTNEFTPerforma.colSociety + "],[" + clsDBTNEFTPerforma.colMPUploaderCode + "]"
        End If
        Return strMain
    End Function

    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        PostData()
    End Sub

    Private Sub PostData()
        Try
            Dim msg As String = ""
            Dim qry As String = "With CTE as (
select TSPL_DBT_NEFT_DETAIL.MP_Account_No,TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code,TSPL_DBT_NEFT_DETAIL.MP_Name ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code
from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
where TSPL_DBT_NEFT_DETAIL.Document_Code='" & txtDocumentNo.Value & "' )
select 'Repeated Account No' as ErrorCode,MP_Account_No as ErrorValue,STRING_AGG(MP_Uploader_Code,',') as MPUploaderCode,STRING_AGG(MP_Code,',') as MPCode,STRING_AGG(MP_Name,',') as MPName,STRING_AGG(VLC_Uploader_Code,',') as DCSUploaderCode from  CTE  group by  MP_Account_No having sum(1)>1
union all
select 'Special Character'as ErrorCode,MP_Name as ErrorValue,MP_Uploader_Code as MPUploaderCode,MP_Code as MPCode,MP_Name as MPName, VLC_Uploader_Code as DCSUploaderCode from CTE   where dbo.RemoveExtraSpaces(UPPER(dbo.RemoveSpecialCharactersWithNumber(MP_Name))) <> MP_Name
union all
select 'Wrong Account No' as ErrorCode,MP_Account_No as ErrorValue,MP_Uploader_Code as MPUploaderCode,MP_Code as MPCode,MP_Name as MPName, VLC_Uploader_Code as DCSUploaderCode from CTE where MP_Account_No like '%.%' ;"
            Dim dtCheck As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dtCheck IsNot Nothing AndAlso dtCheck.Rows.Count > 0 Then
                If common.clsCommon.MyMessageBoxShow(Me, "Error in " & dtCheck.Rows.Count & " Records.Do you want to open it", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes Then
                    Dim frm As New FrmFreeGrid
                    frm.ReportID = "DBTNeftError"
                    frm.Text = "Record Could not Loaded"
                    frm.dt = dtCheck
                    frm.ShowDialog()
                End If
            Else
                clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, txtDocumentNo.Value)
                If (myMessages.postConfirm()) Then
                    UcAttachment2.AllowToSave()
                    UcAttachment2.SaveData(txtDocumentNo.Value)

                    Dim ExcellPath As String = ExportGridPath()
                    clsDBTNEFT.PostData(txtDocumentNo.Value, ExcellPath)
                    clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                    LoadData(txtDocumentNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ExportGrid(ByVal exporter As EnumExportTo)
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                txtDocumentNo.Focus()
                Throw New Exception("Please select Document No")
            End If
            Dim strHeading As String = clsCommon.myCstr("NEFT Uploader")
            Dim arrHeader As List(Of String) = New List(Of String)()
            arrHeader.Add(" ")
            If exporter = EnumExportTo.Excel Then
                If SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
                    transportSql.exportdata(settMaxRowExport, gvFarmer, "", "", 0, 0)
                Else
                    transportSql.exportdata(settMaxRowExport, gvItem, "", "", 0, 0)
                End If
            Else
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID)
                If SettMPIncentiveEntryCycleWiseButNEFTMonthly Then
                    clsCommon.MyExportToPDF(strHeading, gvFarmer, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
                Else
                    clsCommon.MyExportToPDF(strHeading, gvItem, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Function ExportGridPath() As String
        Dim excelPath As String = ""
        Try
            Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.DBTNEFTUploader + "'")
            If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then
                Dim qry As String = "select Email from TSPL_BANK_MASTER where NEFT_DBT_Default = 1 "
                Dim bankEmail As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                If clsCommon.myLen(bankEmail) > 0 Then
                    Dim strHeading As String = clsCommon.myCstr("NEFT Uploader")
                    Dim arrHeader As List(Of String) = New List(Of String)()
                    arrHeader.Add(" ")
                    excelPath = clsCommon.MyExportToExcelGridPath(strHeading, gvItem, Nothing, Me.Text, False, "", "")
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return excelPath
    End Function


    Private Sub txtMCC__My_Click(sender As Object, e As EventArgs) Handles txtMCC._My_Click
        Dim qry As String = "select MCC_Code as [MCC Code],MCC_NAME as [MCC Name],TSPL_MCC_MASTER.plant_code as [Plant Code],tspl_location_master.location_desc as [Plant Name] from TSPL_MCC_MASTER left join tspl_location_master on tspl_location_master.location_code=TSPL_MCC_MASTER.plant_code where  2=2 "
        If clsCommon.myLen(SettMCCOneDBTOneDoc) > 0 Then
            qry += " and MCC_Code='" + SettMCCOneDBTOneDoc + "'"
        End If
        txtMCC.arrValueMember = clsCommon.ShowMultipleSelectForm("MBPSMCCr3", qry, "MCC Code", "MCC Name", txtMCC.arrValueMember, txtMCC.arrDispalyMember)
        txtVLC.arrValueMember = Nothing
        'loadBlankGrid()
        gvItem.DataSource = Nothing
        gvInvalid.DataSource = Nothing
        gvHold.DataSource = Nothing
        gvFarmer.DataSource = Nothing
        SetToDate()

    End Sub

    Private Sub RadMenuItem1_Click(sender As Object, e As EventArgs) Handles RadMenuItem1.Click
        ExportGrid(EnumExportTo.Excel)
    End Sub

    Private Sub RadMenuItem2_Click(sender As Object, e As EventArgs) Handles RadMenuItem2.Click
        ExportGrid(EnumExportTo.PDF)
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                txtDocumentNo.Focus()
                Throw New Exception("Please select Document No")
            End If
            Dim strHeading As String = clsCommon.myCstr("Invalid Data")
            Dim arrHeader As List(Of String) = New List(Of String)()
            arrHeader.Add(" ")
            clsCommon.MyExportToExcelGrid(strHeading, gvInvalid, Nothing, Me.Text)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub txtZone__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtZone._MYValidating
        Dim whr As String = "2=3"
        If SettApplyZoneOnDBT Then
            whr = "2=2"
        End If
        txtZone.Value = ClsZoneMaster.getFinder(whr, txtZone.Value, isButtonClicked)
        lblZone.Text = ClsZoneMaster.GetName(txtZone.Value)
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
            myMessages.blankValue(Me, "Doument not found to Print", Me.Text)
        Else
            funPrint(txtDocumentNo.Value)
        End If
    End Sub
    Public Sub funPrint(ByVal strDocNo As String)
        Try
            Try
                If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                    txtDocumentNo.Focus()
                    Throw New Exception("Please select Document No")
                End If
                Dim qry As String = "Select TSPL_COMPANY_MASTER.Comp_Name,TSPL_DBT_NEFT_DETAIL.MP_Bank as Bank_Code,TSPL_DBT_NEFT_DETAIL.MP_Bank as Bank_Code_Desc,TSPL_DBT_NEFT_DETAIL.MP_Mobile_No,CONVERT(varchar, TSPL_DBT_NEFT.From_Date,103)+' - '+CONVERT(varchar, TSPL_DBT_NEFT.To_Date,103)  as Date_Range,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code as MCC_Code,TSPL_MCC_MASTER.MCC_NAME,
ROW_NUMBER() OVER(PArtition by TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Bank ORDER BY (TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Bank )) As SNo
,TSPL_DBT_NEFT_DETAIL.MP_IFSC_No as Payee_Joint_IFSC_Code ,TSPL_DBT_NEFT_DETAIL.MP_Account_No as Payee_Joint_Account_No,TSPL_DBT_NEFT_DETAIL.MP_Name as Payee_Joint_Name,TSPL_DBT_NEFT_DETAIL.Amount as Payable_Amount
from TSPL_DBT_NEFT_DETAIL 
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code=TSPL_DBT_NEFT_DETAIL.Document_Code
Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR   
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VLC_MASTER_HEAD.VSP_Code
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
where TSPL_DBT_NEFT_DETAIL.Document_Code='" + txtDocumentNo.Value + "' order by TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Bank "

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptDBTNEFTUploader", "Bank Wise NEFT Uploader")
                frmCRV = Nothing
            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnPrintBankLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintBankLetter.Click
        Dim objP As New clsDBTNEFT()
        objP.funPrintBankLetter(txtDocumentNo.Value, False)
    End Sub

    Private Sub btnReverse_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
                If clsCommon.MyMessageBoxShow(Me, "Unpost the current transaction" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    clsDBTNEFT.ReverseAndUnpost(txtDocumentNo.Value)
                    clsCommon.MyMessageBoxShow(Me, "Tansaction unposted succesffuly", Me.Text)
                    LoadData(txtDocumentNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub BtnBank_Click(sender As Object, e As EventArgs) Handles BtnBank.Click
        Try
            Dim obj As New clsDBTNEFT()
            Dim RCDF_Status As Decimal
            obj.Bank_Letter_Date = txtBankLetterDate.Value
            obj.Document_Code = txtDocumentNo.Value

            RCDF_Status = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(" select isnull(RCDF_Status,0) FROM TSPL_DBT_NEFT WHERE Document_Code='" + txtDocumentNo.Value + "'"))
            If RCDF_Status = 1 Then
                clsCommon.MyMessageBoxShow(Me, "DBT NEFT updated by RCDF.", Me.Text)
            Else

                If obj.Bank_Letter_Date > clsCommon.GETSERVERDATE Then
                    txtBankLetterDate.Focus()

                    Throw New Exception("Bank Letter Date should be less than Server Date")
                End If

                If (clsDBTNEFT.SaveBankLetter(obj)) Then
                    clsCommon.MyMessageBoxShow(Me, "Bank Letter Date Save Successfully.", Me.Text)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub btnClrApproval_Click(sender As Object, e As EventArgs) Handles btnClrApproval.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
                If clsCommon.MyMessageBoxShow(Me, "Remove Approval Level" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    clsDBTNEFT.ClearApprovalLevel(txtDocumentNo.Value)
                    clsCommon.MyMessageBoxShow(Me, "Approval Level Remove succesffuly", Me.Text)
                    LoadData(txtDocumentNo.Value, NavigatorType.Current)
                End If
            End If
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
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocumentNo.Value, "Document_Code", "TSPL_DBT_NEFT", "TSPL_DBT_NEFT_DETAIL")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub gvItem_CellDoubleClick(sender As Object, e As GridViewCellEventArgs) Handles gvItem.CellDoubleClick
        Try
            If gvItem.CurrentRow.Index >= 0 Then
                HoldUnloadRow(gvItem, gvHold)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub gvHold_CellDoubleClick(sender As Object, e As GridViewCellEventArgs)Handles gvHold.CellDoubleClick
        Try
            If gvHold.CurrentRow.Index >= 0 Then
                HoldUnloadRow(gvHold, gvItem)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub HoldUnloadRow(gvFrom As MyRadGridView, gvTo As MyRadGridView)
        Try
            gvTo.Rows.AddNew()
            For ii As Integer = 0 To gvFrom.Columns.Count - 1
                gvTo.Rows(gvTo.Rows.Count - 1).Cells(gvFrom.Columns(ii).Name).Value = gvFrom.CurrentRow.Cells(gvFrom.Columns(ii).Name).Value
            Next
            gvFrom.CurrentRow.Delete()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub RadMenuItem5_Click(sender As Object, e As EventArgs) Handles RadMenuItem5.Click
        Try
            clsCommon.MyExportToExcelGrid(Me.Text, gvItem, Nothing, Me.Text)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadMenuItem4_Click(sender As Object, e As EventArgs) Handles RadMenuItem4.Click
        Try
            clsCommon.MyExportToExcelGrid(Me.Text, gvHold, Nothing, Me.Text)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub RadMenuItem8_Click(sender As Object, e As EventArgs) Handles RadMenuItem8.Click


        'Try
        '    Dim gvImport As New UserControls.MyRadGridView
        '    Me.Controls.Add(gvImport)
        '    Dim currentdate As Date = Date.Today
        '    If transportSql.importExcel(gvImport, "SNo", "UnionId", "UnionName", "SocietyID", "SocietyName", "FarmerID", "FarmerName", "FarmerContactNumber", "FarmerEmailID", "FarmerAccountNumber", "FarmerIFCSCode", "FarmerBankName", "FarmerBankBranch", "AMOUNT", "SocietySubName", "AddInfo1", "AddInfo2", "AddInfo3", "AddInfo4") Then
        '        Dim arr As New List(Of String)
        '        Dim strDCSCode As String = ""
        '        Dim dtError As New DataTable
        '        dtError.Columns.Add("RowNo", GetType(Integer))
        '        dtError.Columns.Add("Error", GetType(String))
        '        Try
        '            LoadBlankGrid()
        '            Dim qry As String = "Valid Row [" + clsCommon.myCstr(gvImport.Rows.Count) + "] Do You want to Proceed"
        '            clsCommon.ProgressBarPercentShow()
        '            For ii As Integer = 0 To gvImport.Rows.Count - 1
        '                clsCommon.ProgressBarPercentUpdate((gvImport.Rows(ii).Index + 1) * 100 / (gvImport.Rows.Count + 1), "Importing  : " & (gvImport.Rows(ii).Index + 1) & "/" & gvImport.Rows.Count & "")
        '                gvItem.Rows.AddNew()
        '                For ic As Integer = 0 To gvImport.Columns.Count - 1
        '                    'gvHold.Rows(ii).Cells(gvImport.Columns(ic).Name).Value = gvImport.Rows(ii).Cells(gvImport.Columns(ic).Name).Value
        '                    gvItem.Rows(ii).Cells(ic).Value = gvImport.Rows(ii).Cells(ic).Value
        '                Next
        '                gvItem.Rows(ii).Delete()
        '            Next

        '            clsCommon.ProgressBarPercentHide()
        '            clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
        '            'End If
        '        Catch ex As Exception
        '            clsCommon.ProgressBarPercentHide()
        '            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        '        End Try

        '    End If
        'Catch ex As Exception
        '    clsCommon.ProgressBarPercentHide()
        '    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        'End Try

        Try
            Dim gvImport As New UserControls.MyRadGridView
            Me.Controls.Add(gvImport)
            Dim currentdate As Date = Date.Today
            If transportSql.importExcel(gvImport, "SNo", "UnionId", "UnionName", "SocietyID", "SocietyName", "FarmerID", "FarmerName", "FarmerContactNumber", "FarmerEmailID", "FarmerAccountNumber", "FarmerIFCSCode", "FarmerBankName", "FarmerBankBranch", "AMOUNT", "SocietySubName", "AddInfo1", "AddInfo2", "AddInfo3", "AddInfo4") Then
                Dim arr As New List(Of String)
                Dim strDCSCode As String = ""
                Dim dtError As New DataTable
                dtError.Columns.Add("RowNo", GetType(Integer))
                dtError.Columns.Add("Error", GetType(String))
                Try
                    Dim qry As String = "Valid Row [" + clsCommon.myCstr(gvImport.Rows.Count) + "] Do You want to Proceed"
                    clsCommon.ProgressBarPercentShow()
                    For ii As Integer = 0 To gvImport.Rows.Count - 1
                        clsCommon.ProgressBarPercentUpdate((gvImport.Rows(ii).Index + 1) * 100 / (gvImport.Rows.Count + 1), "Importing  : " & (gvImport.Rows(ii).Index + 1) & "/" & gvImport.Rows.Count & "")
                        gvItem.Rows.AddNew()
                        gvItem.Rows(ii).Cells(SNo).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(SNo).Value)
                        If ii < gvItem.Rows.Count AndAlso ii < gvImport.Rows.Count AndAlso UnionId < gvItem.Columns.Count AndAlso UnionId < gvImport.Columns.Count Then gvItem.Rows(ii).Cells(UnionId).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(UnionId).Value)
                        ''gvItem.Rows(ii).Cells(UnionId).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(UnionId).Value)
                        gvItem.Rows(ii).Cells(UnionName).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(UnionName).Value)
                        gvItem.Rows(ii).Cells(SocietyID).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(SocietyID).Value)
                        gvItem.Rows(ii).Cells(SocietyName).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(SocietyName).Value)
                        gvItem.Rows(ii).Cells(FarmerID).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerID).Value)
                        gvItem.Rows(ii).Cells(FarmerName).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerName).Value)
                        gvItem.Rows(ii).Cells(FarmerContactNumber).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerContactNumber).Value)
                        gvItem.Rows(ii).Cells(FarmerEmailID).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerEmailID).Value)
                        gvItem.Rows(ii).Cells(FarmerAccountNumber).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerAccountNumber).Value)
                        gvItem.Rows(ii).Cells(FarmerIFSCCode).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerIFSCCode).Value)
                        gvItem.Rows(ii).Cells(FarmerBankName).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerBankName).Value)
                        gvItem.Rows(ii).Cells(FarmerBankBranch).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerBankBranch).Value)
                        gvItem.Rows(ii).Cells(AMOUNT).Value = clsCommon.myCDecimal(gvImport.Rows(ii).Cells(AMOUNT).Value)
                        gvItem.Rows(ii).Cells(SocietySubName).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(SocietySubName).Value)
                        gvItem.Rows(ii).Cells(AddInfo1).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(AddInfo1).Value)
                        gvItem.Rows(ii).Cells(AddInfo2).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(AddInfo2).Value)
                        gvItem.Rows(ii).Cells(AddInfo3).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(AddInfo3).Value)
                        gvItem.Rows(ii).Cells(AddInfo4).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(AddInfo4).Value)
                        gvItem.Rows(ii).Cells(DBTTRNO).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(DBTTRNO).Value)
                        gvItem.Rows(ii).Cells(FarmerID2).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(FarmerID2).Value)
                        gvItem.Rows(ii).Cells(ZoneName).Value = clsCommon.myCstr(gvImport.Rows(ii).Cells(ZoneName).Value)
                    Next

                    clsCommon.ProgressBarPercentHide()
                    clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                    'End If
                Catch ex As Exception
                    clsCommon.ProgressBarPercentHide()
                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                End Try

            End If
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub RadMenuItem9_Click(sender As Object, e As EventArgs) Handles RadMenuItem9.Click

        Try
            Dim gvImport As New UserControls.MyRadGridView
            Me.Controls.Add(gvImport)
            Dim currentdate As Date = Date.Today
            If transportSql.importExcel(gvImport, "SNo", "UnionId", "UnionName", "SocietyID", "SocietyName", "FarmerID", "FarmerName", "FarmerContactNumber", "FarmerEmailID", "FarmerAccountNumber", "FarmerIFCSCode", "FarmerBankName", "FarmerBankBranch", "AMOUNT", "SocietySubName", "AddInfo1", "AddInfo2", "AddInfo3", "AddInfo4") Then
                Dim arr As New List(Of String)
                Dim strDCSCode As String = ""
                Dim dtError As New DataTable
                dtError.Columns.Add("RowNo", GetType(Integer))
                dtError.Columns.Add("Error", GetType(String))
                Try
                    LoadBlankGrid()
                    Dim qry As String = "Valid Row [" + clsCommon.myCstr(gvImport.Rows.Count) + "] Do You want to Proceed"
                    clsCommon.ProgressBarPercentShow()
                    For ii As Integer = 0 To gvImport.Rows.Count - 1

                        'If clsCommon.myLen(gvImport.Rows(ii).Cells("FarmerID").Value) > 0 Then

                        clsCommon.ProgressBarPercentUpdate((gvImport.Rows(ii).Index + 1) * 100 / (gvImport.Rows.Count + 1), "Importing  : " & (gvImport.Rows(ii).Index + 1) & "/" & gvImport.Rows.Count & "")

                        'For i As Integer = 0 To gvItem.Rows.Count - 1
                        gvHold.Rows.AddNew()
                        For ic As Integer = 0 To gvImport.Columns.Count - 1
                            'gvHold.Rows(ii).Cells(gvImport.Columns(ic).Name).Value = gvImport.Rows(ii).Cells(gvImport.Columns(ic).Name).Value
                            gvHold.Rows(ii).Cells(ic).Value = gvImport.Rows(ii).Cells(ic).Value

                        Next
                        gvItem.Rows(ii).Delete()
                    Next
                    clsCommon.ProgressBarPercentHide()
                    clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                    'End If
                Catch ex As Exception
                    clsCommon.ProgressBarPercentHide()
                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                End Try

            End If
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

        'trans = clsDBFuncationality.GetTransactin()
        'Try
        '    For i As Integer = 0 To gvItem.Rows.Count - 1
        '        gvHold.Rows.AddNew()
        '        For ii As Integer = 0 To gvItem.Columns.Count - 1
        '            gvHold.Rows(i).Cells(gvItem.Columns(ii).Name).Value = gvItem.Rows(i).Cells(gvItem.Columns(ii).Name).Value
        '        Next
        '        gvItem.Rows(i).Delete()
        '    Next

        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        'End Try
    End Sub



    Private Sub LoadBlankGrid()
        gvHold.DataSource = Nothing
        gvHold.Rows.Clear()
        gvHold.Columns.Clear()

        Dim repoSNO As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSNO.FormatString = ""
        repoSNO.HeaderText = "SNo"
        repoSNO.Name = SNo
        repoSNO.Width = 40
        repoSNO.ReadOnly = True
        repoSNO.IsVisible = True
        gvHold.MasterTemplate.Columns.Add(repoSNO)

        Dim repoUnionId As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnionId.FormatString = ""
        repoUnionId.HeaderText = "UnionId"
        repoUnionId.Name = UnionId
        repoUnionId.Width = 70
        repoUnionId.ReadOnly = True
        repoUnionId.IsVisible = True
        gvHold.MasterTemplate.Columns.Add(repoUnionId)

        Dim repoUnionName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnionName.FormatString = ""
        repoUnionName.HeaderText = "UnionName"
        repoUnionName.Name = UnionName
        repoUnionName.Width = 100
        repoUnionName.ReadOnly = True
        repoUnionName.IsVisible = False
        gvHold.MasterTemplate.Columns.Add(repoUnionName)

        Dim repoSocietyID As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSocietyID.FormatString = ""
        repoSocietyID.HeaderText = "SocietyID"
        repoSocietyID.Name = SocietyID
        repoSocietyID.Width = 150
        repoSocietyID.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoSocietyID)

        Dim repoSocietyName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSocietyName.FormatString = ""
        repoSocietyName.HeaderText = "SocietyName"
        repoSocietyName.Name = SocietyName
        repoSocietyName.Width = 150
        repoSocietyName.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoSocietyName)

        Dim repoFarmerID As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerID.FormatString = ""
        repoFarmerID.HeaderText = "FarmerID"
        repoFarmerID.Name = FarmerID
        repoFarmerID.Width = 150
        repoFarmerID.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerID)

        Dim repoFarmerName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerName.FormatString = ""
        repoFarmerName.HeaderText = "FarmerName"
        repoFarmerName.Name = FarmerName
        repoFarmerName.Width = 150
        repoFarmerName.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerName)

        Dim repoFarmerContactNumber As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerContactNumber.FormatString = ""
        repoFarmerContactNumber.HeaderText = "FarmerContactNumber"
        repoFarmerContactNumber.Name = FarmerContactNumber
        repoFarmerContactNumber.Width = 150
        repoFarmerContactNumber.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerContactNumber)

        Dim repoFarmerEmailID As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerEmailID.FormatString = ""
        repoFarmerEmailID.HeaderText = "FarmerEmailID"
        repoFarmerEmailID.Name = FarmerEmailID
        repoFarmerEmailID.Width = 150
        repoFarmerEmailID.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerEmailID)

        Dim repoFarmerAccountNumber As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerAccountNumber.FormatString = ""
        repoFarmerAccountNumber.HeaderText = "FarmerAccountNumber"
        repoFarmerAccountNumber.Name = FarmerAccountNumber
        repoFarmerAccountNumber.Width = 150
        repoFarmerAccountNumber.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerAccountNumber)

        Dim repoFarmerIFSCCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerIFSCCode.FormatString = ""
        repoFarmerIFSCCode.HeaderText = "FarmerIFSCCode"
        repoFarmerIFSCCode.Name = FarmerIFSCCode
        repoFarmerIFSCCode.Width = 150
        repoFarmerIFSCCode.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerIFSCCode)

        Dim repoFarmerBankName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerBankName.FormatString = ""
        repoFarmerBankName.HeaderText = "FarmerBankName"
        repoFarmerBankName.Name = FarmerBankName
        repoFarmerBankName.Width = 150
        repoFarmerBankName.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerBankName)

        Dim repoFarmerBankBranch As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoFarmerBankBranch.FormatString = ""
        repoFarmerBankBranch.HeaderText = "FarmerBankBranch"
        repoFarmerBankBranch.Name = FarmerBankBranch
        repoFarmerBankBranch.Width = 150
        repoFarmerBankBranch.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoFarmerBankBranch)

        Dim repoAMOUNT As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAMOUNT.FormatString = ""
        repoAMOUNT.HeaderText = "AMOUNT"
        repoAMOUNT.Name = AMOUNT
        repoAMOUNT.Width = 150
        repoAMOUNT.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoAMOUNT)

        Dim repoSocietySubName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSocietySubName.FormatString = ""
        repoSocietySubName.HeaderText = "SocietySubName"
        repoSocietySubName.Name = SocietySubName
        repoSocietySubName.Width = 150
        repoSocietySubName.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoSocietySubName)

        Dim repoAddInfo1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAddInfo1.FormatString = ""
        repoAddInfo1.HeaderText = "AddInfo1"
        repoAddInfo1.Name = AddInfo1
        repoAddInfo1.Width = 150
        repoAddInfo1.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoAddInfo1)

        Dim repoAddInfo2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAddInfo2.FormatString = ""
        repoAddInfo2.HeaderText = "AddInfo2"
        repoAddInfo2.Name = AddInfo2
        repoAddInfo2.Width = 150
        repoAddInfo2.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoAddInfo2)

        Dim repoAddInfo3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAddInfo3.FormatString = ""
        repoAddInfo3.HeaderText = "AddInfo3"
        repoAddInfo3.Name = AddInfo3
        repoAddInfo3.Width = 150
        repoAddInfo3.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoAddInfo3)

        Dim repoAddInfo4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAddInfo4.FormatString = ""
        repoAddInfo4.HeaderText = "AddInfo4"
        repoAddInfo4.Name = AddInfo4
        repoAddInfo4.Width = 150
        repoAddInfo4.ReadOnly = True
        gvHold.MasterTemplate.Columns.Add(repoAddInfo4)

        gvHold.AllowDeleteRow = True
        gvHold.AllowAddNewRow = False
        gvHold.ShowGroupPanel = False
        gvHold.AllowColumnReorder = False
        gvHold.AllowRowReorder = False
        gvHold.EnableSorting = False
        gvHold.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvHold.MasterTemplate.ShowRowHeaderColumn = False
        gvHold.TableElement.TableHeaderHeight = 40
        gvHold.AutoSizeRows = False
        gvHold.Rows.AddNew()
        gvHold.BestFitColumns()
        'ReStoreGridLayoutgv1()
    End Sub

    Private Sub RadButton2_Click(sender As Object, e As EventArgs) Handles RadButton2.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                txtDocumentNo.Focus()
                Throw New Exception("Please select Document No")
            End If
            Dim strHeading As String = clsCommon.myCstr("Hold Data")
            Dim arrHeader As List(Of String) = New List(Of String)()
            arrHeader.Add(" ")
            clsCommon.MyExportToExcelGrid(strHeading, gvHold, Nothing, Me.Text)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub RadMenuItem6_Click(sender As Object, e As EventArgs) Handles RadMenuItem6.Click
        Dim lstMPUploaderCode As New List(Of String)
        Dim gvImport As New UserControls.MyRadGridView
        Me.Controls.Add(gvImport)

        Dim currentDate As Date = Date.Today
        If transportSql.importExcel(gvImport, "FarmerID") Then

            If gvImport.Columns.Count = 1 AndAlso gvImport.Columns.Contains("FarmerID") = True Then

            Else
                clsCommon.ProgressBarPercentHide()
                clsCommon.MyMessageBoxShow(Me, "Invalid Excel Sheet.", Me.Text)
                Exit Sub
            End If

            For Each row As GridViewRowInfo In gvImport.Rows
                lstMPUploaderCode.Add(clsCommon.myCstr(row.Cells("FarmerID").Value))
            Next
        End If
        Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False, lstMPUploaderCode, Nothing))
        If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
            Dim source As DataTable = TryCast(gvInvalid.DataSource, DataTable)
            If source IsNot Nothing Then
                If dt1 Is Nothing OrElse dt1.Columns.Count = 0 Then
                    dt1 = source.Copy()
                End If
                For Each dr As DataRow In source.Rows
                    dt1.ImportRow(dr)
                Next
            End If
            gvInvalid.DataSource = Nothing
            gvInvalid.Rows.Clear()
            gvInvalid.Refresh()
            gvInvalid.DataSource = dt1
            FormatGrid(gvInvalid)
        End If
        'If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
        '    gvInvalid.DataSource = dt1
        '    FormatGrid(gvInvalid)
        'End If

        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False))
        If dt1 IsNot Nothing AndAlso dt2 IsNot Nothing Then
            For i As Integer = dt2.Rows.Count - 1 To 0 Step -1
                Dim farmerId2 As String = clsCommon.myCstr(dt2.Rows(i)("MP Uploader Code"))
                Dim found() As DataRow = dt1.Select("[MP Uploader Code] = '" & farmerId2.Replace("'", "''") & "'")
                If found.Length > 0 Then
                    dt2.Rows.RemoveAt(i)
                End If
            Next
        End If
        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
            gvItem.DataSource = Nothing
            gvItem.Rows.Clear()
            gvItem.Refresh()
            gvItem.DataSource = dt2
            FormatGrid(gvItem)
        End If

        lstMPUploaderCode = Nothing
    End Sub




    'Private Sub RadMenuItem7_Click(sender As Object, e As EventArgs) Handles RadMenuItem7.Click
    '    Dim lstMPAccountNo As New List(Of String)
    '    Dim gvImport As New UserControls.MyRadGridView
    '    Me.Controls.Add(gvImport)

    '    Dim currentDate As Date = Date.Today
    '    If transportSql.importExcel(gvImport, "FarmerAccountNumber") Then
    '        For Each row As GridViewRowInfo In gvImport.Rows
    '            lstMPAccountNo.Add(clsCommon.myCstr(row.Cells("FarmerAccountNumber").Value))
    '        Next
    '    End If
    '    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False, Nothing, lstMPAccountNo))
    '    If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
    '        Dim source As DataTable = TryCast(gvInvalid.DataSource, DataTable)
    '        If source IsNot Nothing Then
    '            If dt1 Is Nothing OrElse dt1.Columns.Count = 0 Then
    '                dt1 = source.Copy()
    '            End If
    '            For Each dr As DataRow In source.Rows
    '                dt1.ImportRow(dr)
    '            Next
    '        End If
    '        gvInvalid.DataSource = Nothing
    '        gvInvalid.Rows.Clear()
    '        gvInvalid.Refresh()
    '        gvInvalid.DataSource = dt1
    '        FormatGrid(gvInvalid)
    '    End If
    '    Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False))
    '    If dt1 IsNot Nothing AndAlso dt2 IsNot Nothing Then
    '        For i As Integer = dt2.Rows.Count - 1 To 0 Step -1
    '            Dim farmerId2 As String = clsCommon.myCstr(dt2.Rows(i)("BENEFICERY ACCOUNT  NO."))
    '            Dim found() As DataRow = dt1.Select("[BENEFICERY ACCOUNT  NO.] = '" & farmerId2.Replace("'", "''") & "'")
    '            If found.Length > 0 Then
    '                dt2.Rows.RemoveAt(i)
    '            End If
    '        Next
    '    End If
    '    If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
    '        gvItem.DataSource = Nothing
    '        gvItem.Rows.Clear()
    '        gvItem.Refresh()
    '        gvItem.DataSource = dt2
    '        FormatGrid(gvItem)
    '    End If
    '    lstMPAccountNo = Nothing
    'End Sub


    Private Sub RadMenuItem7_Click(sender As Object, e As EventArgs) Handles RadMenuItem7.Click

        Try


            Dim lstMPAccountNo As New List(Of String)
            Dim gvImport As New UserControls.MyRadGridView
            Me.Controls.Add(gvImport)

            Dim currentDate As Date = Date.Today


            If transportSql.importExcel(gvImport, "FarmerAccountNumber") Then

                If gvImport.Columns.Count = 1 AndAlso gvImport.Columns.Contains("FarmerAccountNumber") = True Then

                Else
                    clsCommon.ProgressBarPercentHide()
                    clsCommon.MyMessageBoxShow(Me, "Invalid Excel Sheet.", Me.Text)
                    Exit Sub
                End If

                For Each row As GridViewRowInfo In gvImport.Rows
                    lstMPAccountNo.Add(clsCommon.myCstr(row.Cells("FarmerAccountNumber").Value))
                Next
            End If

            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False, Nothing, lstMPAccountNo))
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                Dim source As DataTable = TryCast(gvInvalid.DataSource, DataTable)
                If source IsNot Nothing Then
                    If dt1 Is Nothing OrElse dt1.Columns.Count = 0 Then
                        dt1 = source.Copy()
                    End If
                    For Each dr As DataRow In source.Rows
                        dt1.ImportRow(dr)
                    Next
                End If
                gvInvalid.DataSource = Nothing
                gvInvalid.Rows.Clear()
                gvInvalid.Refresh()
                gvInvalid.DataSource = dt1
                FormatGrid(gvInvalid)
            End If
            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(GetQry("TSPL_DBT_NEFT_DETAIL", False))
            If dt1 IsNot Nothing AndAlso dt2 IsNot Nothing Then
                For i As Integer = dt2.Rows.Count - 1 To 0 Step -1
                    Dim farmerId2 As String = clsCommon.myCstr(dt2.Rows(i)("BENEFICERY ACCOUNT  NO."))
                    Dim found() As DataRow = dt1.Select("[BENEFICERY ACCOUNT  NO.] = '" & farmerId2.Replace("'", "''") & "'")
                    If found.Length > 0 Then
                        dt2.Rows.RemoveAt(i)
                    End If
                Next
            End If
            If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                gvItem.DataSource = Nothing
                gvItem.Rows.Clear()
                gvItem.Refresh()
                gvItem.DataSource = dt2
                FormatGrid(gvItem)
            End If
            lstMPAccountNo = Nothing
        Catch ex As Exception
            clsCommon.ProgressBarPercentHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub




    Private Sub RadMenuItem10_Click(sender As Object, e As EventArgs) Handles RadMenuItem10.Click
        Try
            clsCommon.MyExportToExcelGrid(Nothing, gvItem, Nothing, Me.Text)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadMenuItem11_Click(sender As Object, e As EventArgs) Handles RadMenuItem11.Click
        Try
            clsCommon.MyExportToExcelGrid(Nothing, gvItem, Nothing, Me.Text)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadButton3_Click(sender As Object, e As EventArgs) Handles RadButton3.Click
        fillMPS()
    End Sub

    Private Sub RadButton4_Click(sender As Object, e As EventArgs) Handles RadButton4.Click
        PrintChangeAC(txtDocumentNo.Value)
    End Sub

    Public Sub PrintChangeAC(ByVal strDocNo As String)
        Try
            Try
                Dim qry As String = "WITH CTE_ALL_DBT_DATA AS(SELECT
                                TSPL_DBT_NEFT_DETAIL.Document_Code,
                                TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code AS Farmer_Uploder_Number,
                                TSPL_DBT_NEFT_DETAIL.MP_Name AS Farmer_Name,
                                TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader AS DCS_Uploder_Code,
                                TSPL_VLC_MASTER_HEAD.VLC_Name AS DCS_Name,
                                TSPL_DBT_NEFT_DETAIL.MP_Account_No,
                                TSPL_DBT_NEFT_DETAIL.MP_IFSC_No,
                                TSPL_DBT_NEFT_DETAIL.MP_Bank,
                                TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
                            FROM TSPL_DBT_NEFT_DETAIL
                            LEFT JOIN TSPL_MP_INCENTIVE_ENTRY_DETAIL
                                ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id =
                                   TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
                            LEFT JOIN TSPL_VLC_MASTER_HEAD
                                ON TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code =
                                   TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader
                            UNION ALL
                            SELECT
                                TSPL_DBT_NEFT_DETAIL_INVALID.Document_Code,
                                TSPL_DBT_NEFT_DETAIL_INVALID.MP_Uploader_Code,
                                TSPL_DBT_NEFT_DETAIL_INVALID.MP_Name,
                                TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,
                                TSPL_VLC_MASTER_HEAD.VLC_Name,
                                TSPL_DBT_NEFT_DETAIL_INVALID.MP_Account_No,
                                TSPL_DBT_NEFT_DETAIL_INVALID.MP_IFSC_No,
                                TSPL_DBT_NEFT_DETAIL_INVALID.MP_Bank,
                                TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
                            FROM TSPL_DBT_NEFT_DETAIL_INVALID
                            LEFT JOIN TSPL_MP_INCENTIVE_ENTRY_DETAIL
                                ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id =
                                   TSPL_DBT_NEFT_DETAIL_INVALID.Against_MP_Incentive_TR
                            LEFT JOIN TSPL_VLC_MASTER_HEAD
                                ON TSPL_DBT_NEFT_DETAIL_INVALID.VLC_Uploader_Code =
                                   TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader
                            UNION ALL
                            SELECT
                                TSPL_DBT_NEFT_DETAIL_HOLD.Document_Code,
                                TSPL_DBT_NEFT_DETAIL_HOLD.MP_Uploader_Code,
                                TSPL_DBT_NEFT_DETAIL_HOLD.MP_Name,
                                TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,
                                TSPL_VLC_MASTER_HEAD.VLC_Name,
                                TSPL_DBT_NEFT_DETAIL_HOLD.MP_Account_No,
                                TSPL_DBT_NEFT_DETAIL_HOLD.MP_IFSC_No,
                                TSPL_DBT_NEFT_DETAIL_HOLD.MP_Bank,
                                TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
                            FROM TSPL_DBT_NEFT_DETAIL_HOLD
                            LEFT JOIN TSPL_MP_INCENTIVE_ENTRY_DETAIL
                                ON TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id =
                                   TSPL_DBT_NEFT_DETAIL_HOLD.Against_MP_Incentive_TR
                            LEFT JOIN TSPL_VLC_MASTER_HEAD
                                ON TSPL_DBT_NEFT_DETAIL_HOLD.VLC_Uploader_Code =
                                   TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader)
                        SELECT
                            NEW_DATA.Farmer_Uploder_Number,
                            NEW_DATA.Farmer_Name,
                            NEW_DATA.DCS_Uploder_Code,
                            NEW_DATA.DCS_Name,

                            NEW_DATA.MP_Account_No AS New_Account_Number,
                            NEW_DATA.MP_IFSC_No    AS New_IFSC_Code,
                            NEW_DATA.MP_Bank       AS New_Bank_Name,

                            OLD_DATA.MP_Account_No AS Old_Account_Number,
                            OLD_DATA.MP_IFSC_No    AS Old_IFSC_Code,
                            OLD_DATA.MP_Bank       AS Old_Bank_Name

                        FROM CTE_ALL_DBT_DATA AS NEW_DATA
                        LEFT JOIN CTE_ALL_DBT_DATA AS OLD_DATA
                            ON NEW_DATA.MP_Code = OLD_DATA.MP_Code
                           AND OLD_DATA.Document_Code = '" + txtDocumentNo.Value + "'
                        WHERE NEW_DATA.Document_Code = '" + txtDocumentNo.Value + "' "

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptFarmerChangeACNumber", "Farmer Change AC Number")
                frmCRV = Nothing
            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub



End Class

