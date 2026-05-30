Imports common
Imports System.Data.SqlClient
Imports System
Imports System.Text

Public Class frmNIRQC
    Inherits FrmMainTranScreen
#Region "Variable"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Dim Qry As String
    Dim NoOfDaysForShowFOSSDocument As Integer
#End Region
    Private Sub SetUserMgmtNew()
        'MyBase.SetUserMgmt(clsUserMgtCode.capexmaster)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnPrint.Visible = MyBase.isPrintFlag
        btnPost.Visible = MyBase.isPostFlag
        CancelBtn.Visible = MyBase.isCancel_Flag

        'If MyBase.isReverse Then
        '    RadButton1.Enabled = True
        'Else
        '    RadButton1.Enabled = False
        'End If
        RadButton1.Visible = False
    End Sub
    Private Sub FrmCapexMaster_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            NoOfDaysForShowFOSSDocument = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.NoOfDaysForShowFOSSDocument, clsFixedParameterCode.NoOfDaysForShowFOSSDocument, Nothing))
            SetUserMgmtNew()
            LoadQCStatus()
            AddNew()
            txtDate.Value = clsCommon.GETSERVERDATE()
            ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update ")
            ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D  for Delete ")
            ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
            ButtonToolTip.SetToolTip(btnAddNew, "Press Alt+N Adding New ")
            ButtonToolTip.SetToolTip(btnPost, "Press Alt+P for Post ")
            ButtonToolTip.SetToolTip(CancelBtn, "Press Alt+L for Cancel ")
            If clsCommon.myLen(Me.Tag) > 0 Then
                LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
            End If
            CancelBtn.Enabled = False
            chkManual.Visible = clsCommon.myCBool(clsDBFuncationality.getSingleValue("select isnull(Show_Manual_Collection,0) from TSPL_USER_MASTER where user_code = '" & objCommonVar.CurrentUserCode & "'") = 1)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadQCStatus()
        Dim dt As New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = Nothing

        dr = dt.NewRow()
        dr("Code") = ""
        dr("Name") = "Select"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "1"
        dr("Name") = "OK"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "2"
        dr("Name") = "Not OK"
        dt.Rows.Add(dr)



        cboVisualQCStatus.DataSource = dt
        cboVisualQCStatus.ValueMember = "Code"
        cboVisualQCStatus.DisplayMember = "Name"
    End Sub
    Sub AddNew()
        RadButton1.Visible = False
        isNewEntry = True
        txtCode.Value = Nothing
        txtCode.Focus()
        btnSave.Text = "Save"
        btnSave.Enabled = True
        btnPrint.Enabled = True
        btnPrint.Visible = True
        btnDelete.Enabled = False
        LBListMoisture.Text = ""
        LBListSilica.Text = ""
        LBListFat.Text = ""
        LBListProtein.Text = ""
        LBListFiber.Text = ""
        txtSampleNo.Value = ""
        'txtDate.Value = clsCommon.GETSERVERDATE()
        BlankAllControls()
        txtAccept.Visible = False
    End Sub
    Sub BlankAllControls()
        txtCode.Value = ""
        txtRemarks.Text = ""
        cboVisualQCStatus.SelectedValue = ""
        txtMRNNo.Value = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtRemarks.Text = ""
        UsLock1.Status = ERPTransactionStatus.Pending
        BlankMRNFields()
    End Sub
    Sub BlankMRNFields()
        lblGRNNo.Text = ""
        lblGRNDate.Text = ""
        lblWeightmentNo.Text = ""
        lblWeightmentDate.Text = ""
        lblRAL.Text = ""
        lblVendorCode.Text = ""
        lblVendorName.Text = ""
        lblBillToLocationCode.Text = ""
        lblBillToLocationName.Text = ""
        lblItem.Text = ""
        lblItemName.Text = ""
        lblVehicleNo.Text = ""
    End Sub
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        AddNew()
        Dim obj As New clsNIRQC()
        obj = clsNIRQC.GetData(strCode, NavTyep, Nothing)
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
            isNewEntry = False
            txtCode.Value = obj.Document_No
            txtDate.Value = obj.Document_Date
            txtMRNNo.Value = obj.MRN_No
            ' txtSampleNo.Value = obj.REF_PK_ID
            txtRemarks.Text = obj.QC_Remarks
            txtSampleNo.Value = obj.Against_Foss_PK_ID

            Dim qry1 As String = "SELECT  Moisture,Silica_DM,Fat_DM,Protein_DM,Fiber_DM  FROM TSPL_NIR_QC_FOSS WHERE PK_Id= '" + txtSampleNo.Value + "'"
            Dim dts As DataTable = clsDBFuncationality.GetDataTable(qry1)
            If dts.Rows.Count > 0 Then

                LBListMoisture.Text = dts.Rows(0)("Moisture").ToString()
                LBListSilica.Text = dts.Rows(0)("Silica_DM").ToString()
                LBListFat.Text = dts.Rows(0)("Fat_DM").ToString()
                LBListProtein.Text = dts.Rows(0)("Protein_DM").ToString()
                LBListFiber.Text = dts.Rows(0)("Fiber_DM").ToString()

                LBListMoisture.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
                LBListSilica.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
                LBListFat.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
                LBListProtein.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
                LBListFiber.TextAlignment = System.Drawing.ContentAlignment.MiddleRight


            End If
            cboVisualQCStatus.SelectedValue = clsCommon.myCstr(obj.QC_Status)
            UsLock1.Status = obj.Status
            txtAccept.Text = obj.FOSS_Status
            chkManual.Checked = IIf(obj.Is_Manual = 1, True, False)
            txtAccept.Visible = True
            If clsCommon.myLen(obj.FOSS_Status) <= 0 Then
                txtAccept.BackColor = Color.Gray
            ElseIf clsCommon.CompairString(obj.FOSS_Status, "Accepted") = CompairStringResult.Equal Then
                txtAccept.BackColor = Color.Green
            ElseIf clsCommon.CompairString(obj.FOSS_Status, "Under Deviation") = CompairStringResult.Equal Then
                txtAccept.BackColor = Color.Yellow
                txtAccept.Enabled = True
            ElseIf clsCommon.CompairString(obj.FOSS_Status, "Rejected") = CompairStringResult.Equal Then
                txtAccept.BackColor = Color.Red
            End If
            If obj.Status = ERPTransactionStatus.Approved Then
                btnSave.Enabled = False
                btnPost.Enabled = False
                btnDelete.Enabled = False
                CancelBtn.Enabled = True
                'CancelBtn.Visible = True
            Else
                btnSave.Enabled = True
                btnPost.Enabled = True
                btnDelete.Enabled = True
            End If
            LoadMRNData()
        End If
        CancelBtn.Enabled = True
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Save()
    End Sub
    Public Sub Save()
        Try
            If AllowToSave() Then
                Dim obj As New clsNIRQC()
                obj.Document_No = txtCode.Value
                obj.Document_Date = txtDate.Value
                obj.MRN_No = txtMRNNo.Value
                obj.QC_Remarks = txtRemarks.Text
                obj.QC_Status = clsCommon.myCDecimal(cboVisualQCStatus.SelectedValue)
                obj.Form_ID = Me.Form_ID
                obj.FOSS_Status = txtAccept.Text
                obj.Is_Manual = IIf(chkManual.Checked = True, 1, 0)
                'Dim sampleNo As String = txtSampleNo.Value.ToString().Trim()
                '           Dim cnt As Integer = Convert.ToInt32(clsDBFuncationality.getSingleValue("SELECT COUNT(*)  FROM TSPL_NIR_QC  INNER JOIN TSPL_NIR_QC_FOSS  ON TSPL_NIR_QC.REF_PK_ID = TSPL_NIR_QC_FOSS.PK_Id
                'WHERE TSPL_NIR_QC_FOSS.Sample_Number = '" & sampleNo & "'"))
                '           If cnt > 0 Then
                '               clsCommon.MyMessageBoxShow(Me, "Sample already used!", Me.Text)

                '               Exit Sub
                '           End If

                obj.Against_Foss_PK_ID = txtSampleNo.Value

                'If clsCommon.CompairString(clsCommon.myCstr(cboVisualQCStatus.SelectedItem), "Not Ok") <> CompairStringResult.Equal Then
                'Dim dt As DataTable = clsDBFuncationality.GetDataTable(ReturnMRNDataQry())
                'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                '    clsApply_Approval.CheckApprovalRequired(clsCommon.myCstr(dt.Rows(0)("Bill_To_Location")), Nothing, Me.Form_ID, txtCode.Value, txtDate.Value, Nothing, txtRemarks.Text, clsCommon.myCdbl(dt.Rows(0)("MRN_Total_Amt")), 0, Nothing, Nothing, 0, False)
                'End If
                'End If
                If (obj.SaveData(obj, isNewEntry)) Then
                    LoadData(obj.Document_No, NavigatorType.Current)
                    clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
    Function AllowToSave() As Boolean
        Xtra.TransactionValidity(txtDate.Value)
        If clsCommon.myCDate(txtDate.Value).Date() > clsCommon.GETSERVERDATE().Date() Then
            txtDate.Focus()
            Throw New Exception("Cannot allow future date -  " & clsCommon.myCDate(txtDate.Value).Date())
        End If
        If clsCommon.myLen(cboVisualQCStatus.SelectedValue) <= 0 Then
            cboVisualQCStatus.Focus()
            Throw New Exception("Please select " + cboVisualQCStatus.MyLinkLable1.Text)
        End If
        If clsCommon.myLen(txtMRNNo.Value) <= 0 Then
            txtMRNNo.Focus()
            Throw New Exception("Please select " + txtMRNNo.MyLinkLable1.Text)
        End If
        If clsCommon.myCDate(txtDate.Value, "dd/MM/yyyy") < clsCommon.myCDate(lblWeightmentDate.Text, "dd/MM/yyyy") Then
            Throw New Exception(" NIR QC date can't be less then Weighment date !")
        End If
        If chkManual.Checked = False Then
            If clsCommon.myLen(txtSampleNo.Value) <= 0 Then
                Throw New Exception("Please select " & txtSampleNo.MyLinkLable1.Text)
            End If
        End If

        Return True
    End Function
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        funDelete()
    End Sub
    Sub funDelete()
        Try
            clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, txtCode.Value)
            If (myMessages.deleteConfirm()) Then
                If (clsNIRQC.DeleteData(txtCode.Value)) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If
        Catch ex As Exception
            myMessages.myExceptions(ex)
        End Try
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub FrmCapexMaster_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
            AddNew()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag AndAlso btnSave.Enabled Then
            Save()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btnDelete.Enabled Then
            funDelete()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso MyBase.isPostFlag AndAlso btnPost.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso MyBase.isPostFlag AndAlso btnPost.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.L AndAlso MyBase.isCancel_Flag AndAlso CancelBtn.Enabled Then
            CancelNIRQCData()
        ElseIf e.Alt And e.KeyCode = Keys.C Then
            Me.Close()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then

                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    RadButton1.Visible = True
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
                'MessageBox.Show("You are not authorized to perform this action.", "Unauthorized Access", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
    End Sub
    Private Sub RadMenuItem4_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    Private Sub txtCode__MYNavigator(sender As Object, e As EventArgs, NavType As NavigatorType) Handles txtCode._MYNavigator
        Try
            LoadData(txtCode.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtCode._MYValidating

        txtCode.Value = clsNIRQC.getFinder("", txtCode.Value, isButtonClicked)
        If txtCode.Value <> "" Then
            LoadData(txtCode.Value, NavigatorType.Current)
        Else
            AddNew()
        End If
    End Sub
    Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
        AddNew()
    End Sub
    Private Sub txtMRNNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtMRNNo._MYValidating
        Dim qry As String = "select 
TSPL_MRN_DETAIL.MRN_No,TSPL_MRN_HEAD.MRN_Date,
TSPL_MRN_HEAD.Against_GRN,TSPL_GRN_HEAD.GRN_Date,TSPL_GRN_HEAD.VehicleNo ,TSPL_PO_WEIGHTMENT_HEAD.Weighment_Code,TSPL_PO_WEIGHTMENT_HEAD.Weighment_Date,TSPL_PURCHASE_ORDER_HEAD.RefTendorNo,TSPL_MRN_HEAD.Vendor_Code,TSPL_MRN_HEAD.Vendor_Name,TSPL_MRN_HEAD.Bill_To_Location,TSPL_LOCATION_MASTER.Location_Desc,TSPL_MRN_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc , case when TSPL_GRN_HEAD.VisualQCStatus=1 then 'Ok' when TSPL_GRN_HEAD.VisualQCStatus='2' then 'Not Ok' when TSPL_GRN_HEAD.VisualQCStatus='3' then 'Partial Ok'  when TSPL_GRN_HEAD.VisualQCStatus='4' then 'On Hold' else 'Pending' end as [Visual QC Status], case when TSPL_GRN_HEAD.VisualQCStatusSecond=1 then 'Ok' when TSPL_GRN_HEAD.VisualQCStatusSecond='2' then 'Not Ok' when TSPL_GRN_HEAD.VisualQCStatusSecond='3' then 'Partial Ok'  when TSPL_GRN_HEAD.VisualQCStatusSecond='4' then 'On Hold' else ' ' end as [Visual QC Second Status]
from TSPL_MRN_DETAIL
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_MRN_DETAIL.Item_Code 
left outer join TSPL_MRN_HEAD  on TSPL_MRN_HEAD.MRN_No=TSPL_MRN_DETAIL.MRN_No
left outer join TSPL_GRN_HEAD on TSPL_GRN_HEAD.GRN_No=TSPL_MRN_HEAD.Against_GRN
left outer join TSPL_PO_WEIGHTMENT_HEAD on TSPL_PO_WEIGHTMENT_HEAD.Against_GRN_No=TSPL_MRN_HEAD.Against_GRN
left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_GRN_HEAD.Against_PO
left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_MRN_HEAD.Bill_To_Location"
        Dim whrcls As String = "TSPL_MRN_HEAD.Status=1 and TSPL_GRN_HEAD.IsSkipPurchaseQC=0 and TSPL_MRN_HEAD.NIR_QC=1 and TSPL_ITEM_MASTER.NIR_QC=1
and not exists(select 1 from TSPL_NIR_QC where TSPL_NIR_QC.MRN_No=TSPL_MRN_DETAIL.MRN_No and TSPL_NIR_QC.Document_No not in ('" + txtCode.Value + "'))  "
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            whrcls += " and TSPL_MRN_HEAD.Bill_To_Location in (" + objCommonVar.strCurrUserLocations + ")"
        End If
        txtMRNNo.Value = clsCommon.ShowSelectForm("NICQCMRNFnd", qry, "MRN_No", whrcls, txtMRNNo.Value, "", isButtonClicked)
        LoadMRNData()
    End Sub



    Private Sub LoadMRNData()
        BlankMRNFields()
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(clsNIRQC.ReturnMRNDataQry(txtMRNNo.Value))
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim QC1 As Integer = clsCommon.myCdbl(dt.Rows(0)("VisualQCStatus"))
            Dim QC2 As Integer = clsCommon.myCdbl(dt.Rows(0)("VisualQCStatusSecond"))
            If QC1 = 2 OrElse QC2 = 2 Then
                clsCommon.MyMessageBoxShow(Me, "Visual QC NotOk so NIR QC not allowed ", Me.Text)
                Exit Sub
            End If
            lblGRNNo.Text = clsCommon.myCstr(dt.Rows(0)("Against_GRN"))
            If dt.Rows(0)("GRN_Date") IsNot DBNull.Value Then
                lblGRNDate.Text = clsCommon.myCstr(dt.Rows(0)("GRN_Date"))
            End If
            lblWeightmentNo.Text = clsCommon.myCstr(dt.Rows(0)("Weighment_Code"))
            If dt.Rows(0)("Weighment_Date") IsNot DBNull.Value Then
                lblWeightmentDate.Text = clsCommon.myCstr(dt.Rows(0)("Weighment_Date"))
            End If
            lblRAL.Text = clsCommon.myCstr(dt.Rows(0)("RefTendorNo"))
            lblVendorCode.Text = clsCommon.myCstr(dt.Rows(0)("Vendor_Code"))
            lblVendorName.Text = clsCommon.myCstr(dt.Rows(0)("Vendor_Name"))
            lblBillToLocationCode.Text = clsCommon.myCstr(dt.Rows(0)("Bill_To_Location"))
            lblBillToLocationName.Text = clsCommon.myCstr(dt.Rows(0)("Location_Desc"))
            lblItem.Text = clsCommon.myCstr(dt.Rows(0)("Item_Code"))
            lblItemName.Text = clsCommon.myCstr(dt.Rows(0)("Item_Desc"))
            lblVehicleNo.Text = clsCommon.myCstr(dt.Rows(0)("VehicleNo"))
        End If
    End Sub
    'Public Function ChkApproval() As Boolean
    '    Try
    '        Dim Qry As String = clsApprovalScreen.GetDataQry(Module_Code, Form_ID, lblBillToLocationCode.Text, Nothing)
    '        Dim chkCount As Integer = clsDBFuncationality.getSingleValue("Select Count(*) from(" & Qry & ")xyz")
    '        If chkCount > 0 Then
    '            Throw New Exception("Need to send for approval !")
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        Throw New Exception(ex.Message)
    '    End Try
    'End Function
    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        PostData()
    End Sub
    Sub PostData()
        Try
            clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, txtCode.Value)
            If clsCommon.CompairString(cboVisualQCStatus.SelectedValue, "2") = CompairStringResult.Equal Then
                If clsCommon.MyMessageBoxShow("NIR QC Status : " & clsCommon.myCstr(cboVisualQCStatus.SelectedItem) & Environment.NewLine & "Do you want to cancel the NIRQC with GRN ?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    If clsNIRQC.CancelData(Me.Form_ID, clsCommon.myCstr(txtCode.Value), lblGRNNo.Text, True) Then
                        clsCommon.MyMessageBoxShow(Me, "Data Cancel Successfully ", Me.Text)
                        AddNew()
                    End If
                End If
            Else
                If myMessages.postConfirm() AndAlso clsNIRQC.postdata(txtCode.Value) Then
                    clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                    LoadData(txtCode.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub



    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        Try
            If clsCommon.myLen(txtCode.Value) > 0 Then
                If clsCommon.MyMessageBoxShow("Unpost the current transaction" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    clsNIRQC.ReverseAndUnpost(txtCode.Value)
                    clsCommon.MyMessageBoxShow(Me, "Tansaction unposted succesffuly", Me.Text)
                    LoadData(txtCode.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim sqlqry As String = "Select TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1,Add2,TSPL_COMPANY_MASTER.ADD2, '" + objCommonVar.CurrentUserCode + "' as UserName ,  ROW_NUMBER() OVER (ORDER BY TSPL_NIR_QC.Document_No) AS SrNo, Case When TSPL_MRN_HEAD.Status=1 Then 'OK' else 'Not OK' end as QC_Status,
TSPL_MRN_HEAD.Against_GRN,TSPL_GRN_HEAD.GRN_Date,TSPL_PO_WEIGHTMENT_HEAD.Weighment_Code,TSPL_PO_WEIGHTMENT_HEAD.Weighment_Date,TSPL_PURCHASE_ORDER_HEAD.RefTendorNo,TSPL_MRN_HEAD.Vendor_Code,TSPL_MRN_HEAD.Vendor_Name,TSPL_MRN_HEAD.Bill_To_Location,TSPL_LOCATION_MASTER.Location_Desc,TSPL_MRN_DETAIL.Item_Code,TSPL_ITEM_MASTER.Item_Desc,TSPL_MRN_HEAD.VehicleNo,TSPL_NIR_QC.Document_No,TSPL_NIR_QC.QC_Remarks,TSPL_NIR_QC.MRN_No,TSPL_NIR_QC.Document_Date
From TSPL_MRN_DETAIL
Left join TSPL_NIR_QC on TSPL_NIR_QC.MRN_No=TSPL_MRN_DETAIL.MRN_No
Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_MRN_DETAIL.Item_Code SS
Left outer join TSPL_MRN_HEAD  on TSPL_MRN_HEAD.MRN_No=TSPL_MRN_DETAIL.MRN_No
Left outer join TSPL_GRN_HEAD on TSPL_GRN_HEAD.GRN_No=TSPL_MRN_HEAD.Against_GRN
Left outer join TSPL_PO_WEIGHTMENT_HEAD on TSPL_PO_WEIGHTMENT_HEAD.Against_GRN_No=TSPL_MRN_HEAD.Against_GRN
Left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_GRN_HEAD.Against_PO
Left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_MRN_HEAD.Bill_To_Location
LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code1='" & objCommonVar.CurrComp_Code1 & "'

where TSPL_MRN_DETAIL.MRN_No ='" + txtMRNNo.Value + "' and TSPL_MRN_HEAD.Status=1 and TSPL_MRN_HEAD.NIR_QC=1 and TSPL_ITEM_MASTER.NIR_QC=1"

            Dim dtItem As DataTable
            dtItem = clsDBFuncationality.GetDataTable(sqlqry)
            If dtItem.Rows.Count > 0 Then
                Dim crysFrm As New frmCrystalReportViewer()
                crysFrm.funreport(MyBase.Form_ID, CrystalReportFolder.PurchaseOrder, dtItem, "NIRQc", "NIRQc Details")

            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub CancelNIRQCData()
        Try
            If clsCommon.myLen(txtCode.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Select Document Code", Me.Text)
                Exit Sub
            End If
            If clsCommon.myLen(txtCode.Value) > 0 Then
                If clsCommon.MyMessageBoxShow("Do you want to cancel the NIRQC?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If
                '    If clsPurchaseOrderHead.CheckPOUsedInSRNorGRN(clsCommon.myCstr(txtDocNo.Value), Nothing) Then
                '        Throw New Exception("PO can not be cancelled because it is used in SRN/GRN.")

                '    End If

                '    If (myMessages.CancelConfirms(Me)) Then
                '                Dim Qry As String = "select distinct TSPL_SRN_DETAIL.SRN_No,TSPL_SRN_HEAD.Status from TSPL_SRN_DETAIL 
                'left outer join TSPL_SRN_HEAD on TSPL_SRN_HEAD.SRN_No=TSPL_SRN_DETAIL.SRN_No where TSPL_SRN_DETAIL.MRN_Id ='" + txtMRNNo.Value + "'"
                Dim frm1 As New FrmPWD(Nothing)
                frm1.strType = "PO Cancel"
                frm1.strCode = "PO Cancel"
                frm1.ShowDialog()

                If frm1.isPasswordCorrect Then
                    Dim iscancel As Boolean = False
                    If clsNIRQC.CheckNIRQCUsedInSRN(clsCommon.myCstr(txtMRNNo.Value), Nothing) Then
                        Throw New Exception("NIRQC can not be cancelled because it is used in SRN.")
                        'Else
                        '    clsPurchaseOrderHead.ReverseAndUnpost(txtDocNo.Value, MyBase.Form_ID)
                    End If
                    'If common.clsCommon.MyMessageBoxShow("Do you want to cancel the NIRQC?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Dim Reason As String = ""
                    clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtCode.Value))
                    If clsCancelLog.CheckForReasonOnDelete() Then
                        '' REASON FOR DELETE 
                        Dim frm As New FrmFreeTxtBox1
                        frm.Text = "Remarks for Cancel"
                        frm.ShowDialog()
                        If clsCommon.myLen(frm.strRmks) <= 0 Then
                            Exit Sub
                        Else
                            Reason = frm.strRmks
                        End If
                    End If
                    If clsNIRQC.CancelData(Me.Form_ID, clsCommon.myCstr(txtCode.Value), IIf(clsCommon.CompairString(clsCommon.myCstr(cboVisualQCStatus.SelectedItem), "Not Ok") = CompairStringResult.Equal, True, False), False) Then

                        'If clsNIRQC.CancelData(Me.Form_ID, clsCommon.myCstr(txtCode.Value)) Then
                        ' saveCancelLog(Reason, "Cancel", Nothing)
                        clsCommon.MyMessageBoxShow(Me, "Data Cancel Successfully ", Me.Text)
                        AddNew()
                    End If
                    'End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub CancelBtn_Click(sender As Object, e As EventArgs) Handles CancelBtn.Click
        CancelNIRQCData()
    End Sub

    Private Sub txtSampleNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtSampleNo._MYValidating
        Try
            Dim InstrumentalId As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select NIR_QC_Instrumental_ID from TSPL_LOCATION_MASTER where Location_Code='" & lblBillToLocationCode.Text & "'"))
            Dim ProductId As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select NIR_QC_Product_ID from TSPL_ITEM_MASTER where Item_Code='" + lblItem.Text + "'"))

            Dim qry1 As String = " select NIR_QC_Instrumental_ID from TSPL_LOCATION_MASTER where Location_Code='" + lblBillToLocationCode.Text + "'"


            Dim QRY11 As String = "SELECT TSPL_NIR_QC_FOSS.PK_Id as SampleID, TSPL_NIR_QC_FOSS.Sample_Number as [Sample Number] ,TSPL_NIR_QC_FOSS.Vehicle_No as [Vehicle No],format( TSPL_NIR_QC_FOSS.Analysis_Time,'dd/MM/yyyy hh:mm tt') as [Analysis Time],TSPL_NIR_QC_FOSS.Sample_Type AS [Sample Type],TSPL_NIR_QC_FOSS.Sample_Comment AS [Comment],TSPL_NIR_QC_FOSS.Product_Code AS [Product Code],TSPL_NIR_QC_FOSS.Instrument_Serial_Number as [Instrument No.],Moisture,Silica_DM as [Silica],Fat_DM AS [Fat],Protein_DM as [Protein],Fiber_DM AS [Fiber]
                FROM TSPL_NIR_QC_FOSS "

            Dim Whr As String = " TSPL_NIR_QC_FOSS.Instrument_Serial_Number='" + InstrumentalId + "'
and TSPL_NIR_QC_FOSS.Product_Code 	='" + ProductId + "'
and convert(date,TSPL_NIR_QC_FOSS.Analysis_Time,103) >= CONVERT(DATE,'" + clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") + "',103)
and convert(date,TSPL_NIR_QC_FOSS.Analysis_Time,103) <= CONVERT(DATE,'" + clsCommon.GetPrintDate(txtDate.Value.AddDays(NoOfDaysForShowFOSSDocument), "dd/MMM/yyyy") + "',103)
and not exists ( Select 1 FROM TSPL_NIR_QC WHERE TSPL_NIR_QC.Against_Foss_PK_ID = TSPL_NIR_QC_FOSS.PK_Id and  TSPL_NIR_QC.Document_No not in ('" + txtCode.Value + "') ) "

            txtSampleNo.Value = clsCommon.ShowSelectForm("Location@Plant@Master", QRY11, "SampleID", Whr, txtSampleNo.Value, "SampleID", isButtonClicked)
            Dim dt As New DataTable
            Dim qry12 As String = "SELECT  Moisture,Silica_DM,Fat_DM,Protein_DM,Fiber_DM  FROM TSPL_NIR_QC_FOSS WHERE PK_Id= '" + txtSampleNo.Value + "'"
            Dim dts As DataTable = clsDBFuncationality.GetDataTable(qry12)
            If dts IsNot Nothing AndAlso dts.Rows.Count > 0 Then
                LBListMoisture.Text = dts.Rows(0)("Moisture").ToString()
                LBListSilica.Text = dts.Rows(0)("Silica_DM").ToString()
                LBListFat.Text = dts.Rows(0)("Fat_DM").ToString()
                LBListProtein.Text = dts.Rows(0)("Protein_DM").ToString()
                LBListFiber.Text = dts.Rows(0)("Fiber_DM").ToString()
                UpdateStausNewInput()
            Else
                LBListMoisture.Text = ""
                LBListSilica.Text = ""
                LBListFat.Text = ""
                LBListProtein.Text = ""
                LBListFiber.Text = ""
                txtAccept.Text = ""
                txtAccept.BackColor = Color.Gray
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub UpdateStausNewInput()
        Dim NetResult As Boolean
        Dim TempInputData As Decimal = 0
        Dim InputDataDeductionPer As Decimal = 0
        Try
            Dim qry As String = "select isnull(TSPL_QC_LOG_SHEET_MASTER.NIRQC_Para_type,'')as NIRQC_Para_type,(TSPL_PARAMETER_RANGE_MASTER_QC.lower_range) as lower_range,(TSPL_PARAMETER_RANGE_MASTER_QC.upper_range) as upper_range,(TSPL_PARAMETER_RANGE_MASTER_QC.status) as status,(TSPL_PARAMETER_RANGE_MASTER_QC.value1) as value1,(TSPL_PARAMETER_RANGE_MASTER_QC.qc_status) as qc_status,(TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_Per) as Deduction_Per
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_lower_range
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_upper_range
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_Ratio
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_lower_range2
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_upper_range2
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_Ratio2
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_lower_range3
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_upper_range3
                ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_Ratio3
                 ,TSPL_PARAMETER_RANGE_MASTER_QC.Deduction_Method
                 from TSPL_ITEM_MASTER_PURCHASE_QC_PARAMETER left outer join TSPL_PARAMETER_RANGE_MASTER_QC on TSPL_PARAMETER_RANGE_MASTER_QC.qc_param_code=TSPL_ITEM_MASTER_PURCHASE_QC_PARAMETER.QC_code and TSPL_PARAMETER_RANGE_MASTER_QC.trans_id='standard' left outer join TSPL_QC_LOG_SHEET_MASTER on TSPL_QC_LOG_SHEET_MASTER.code=TSPL_ITEM_MASTER_PURCHASE_QC_PARAMETER.QC_code 
                left outer join tspl_item_master on tspl_item_master.item_code=TSPL_ITEM_MASTER_PURCHASE_QC_PARAMETER.item_code where TSPL_ITEM_MASTER_PURCHASE_QC_PARAMETER.item_code = '" + lblItem.Text + "' and NIRQC_Para_type in ('Moisture','Silica','Fat','Protein','Fiber')  "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            Dim TempDedPercentage As Decimal = 0
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If clsCommon.CompairString(dr("NIRQC_Para_type"), "Moisture") = CompairStringResult.Equal Then
                        TempInputData = clsCommon.myCDecimal(LBListMoisture.Text)
                    ElseIf clsCommon.CompairString(dr("NIRQC_Para_type"), "Silica") = CompairStringResult.Equal Then
                        TempInputData = clsCommon.myCDecimal(LBListSilica.Text)
                    ElseIf clsCommon.CompairString(dr("NIRQC_Para_type"), "Fat") = CompairStringResult.Equal Then
                        TempInputData = clsCommon.myCDecimal(LBListFat.Text)
                    ElseIf clsCommon.CompairString(dr("NIRQC_Para_type"), "Protein") = CompairStringResult.Equal Then
                        TempInputData = clsCommon.myCDecimal(LBListProtein.Text)
                    ElseIf clsCommon.CompairString(dr("NIRQC_Para_type"), "Fiber") = CompairStringResult.Equal Then
                        TempInputData = clsCommon.myCDecimal(LBListFiber.Text)
                    End If
                    If clsCommon.myCDecimal(TempInputData) > 0 Then
                        If TempInputData >= clsCommon.myCDecimal(dr("lower_range")) And TempInputData <= clsCommon.myCDecimal(dr("upper_range")) Then
                            NetResult = True
                        ElseIf TempInputData >= clsCommon.myCDecimal(dr("Deduction_lower_range")) AndAlso TempInputData <= clsCommon.myCDecimal(dr("Deduction_upper_range")) Then
                            NetResult = True
                            If clsCommon.myCDecimal(dr("Deduction_Method")) = 0 Then
                                If clsCommon.myCDecimal(dr("Deduction_lower_range")) > clsCommon.myCDecimal(dr("upper_range")) Then
                                    InputDataDeductionPer = System.Math.Round((TempInputData - clsCommon.myCDecimal(dt.Rows(0)("upper_range"))) * clsCommon.myCDecimal(dt.Rows(0)("Deduction_Ratio")), 2, MidpointRounding.AwayFromZero)
                                ElseIf clsCommon.myCDecimal(dt.Rows(0)("Deduction_lower_range")) < clsCommon.myCDecimal(dt.Rows(0)("lower_range")) Then
                                    InputDataDeductionPer = System.Math.Round((clsCommon.myCDecimal(dt.Rows(0)("lower_range")) - TempInputData) * clsCommon.myCDecimal(dt.Rows(0)("Deduction_Ratio")), 2, MidpointRounding.AwayFromZero)
                                Else
                                    InputDataDeductionPer = System.Math.Round((clsCommon.myCDecimal(dt.Rows(0)("lower_range")) - TempInputData) * clsCommon.myCDecimal(dt.Rows(0)("Deduction_Ratio")), 2, MidpointRounding.AwayFromZero)
                                End If
                            Else
                                InputDataDeductionPer = System.Math.Round(clsCommon.myCDecimal(dt.Rows(0)("Deduction_Ratio")), 2, MidpointRounding.AwayFromZero)
                            End If
                            If clsCommon.myCdbl(InputDataDeductionPer) > 0 Then
                                Exit For
                            End If
                        ElseIf TempInputData >= clsCommon.myCDecimal(dr("Deduction_lower_range2")) AndAlso TempInputData <= clsCommon.myCDecimal(dr("Deduction_upper_range2")) Then
                            NetResult = True
                            If clsCommon.myCDecimal(dr("Deduction_Method")) = 0 Then
                                If clsCommon.myCDecimal(dr("Deduction_lower_range2")) > clsCommon.myCDecimal(dr("upper_range")) Then
                                    TempDedPercentage = ((clsCommon.myCDecimal(dr("Deduction_upper_range")) - clsCommon.myCDecimal(dr("upper_range"))) * clsCommon.myCDecimal(dr("Deduction_Ratio")))
                                    InputDataDeductionPer = System.Math.Round(TempDedPercentage + ((TempInputData - clsCommon.myCDecimal(dr("Deduction_upper_range"))) * clsCommon.myCDecimal(dr("Deduction_Ratio2"))), 2, MidpointRounding.AwayFromZero)
                                ElseIf clsCommon.myCDecimal(dr("Deduction_lower_range2")) < clsCommon.myCDecimal(dr("lower_range")) Then
                                    TempDedPercentage = ((clsCommon.myCDecimal(dr("lower_range")) - clsCommon.myCDecimal(dr("Deduction_lower_range"))) * clsCommon.myCDecimal(dr("Deduction_Ratio")))
                                    InputDataDeductionPer = System.Math.Round(TempDedPercentage + ((clsCommon.myCDecimal(dr("Deduction_lower_range")) - TempInputData) * clsCommon.myCDecimal(dr("Deduction_Ratio2"))), 2, MidpointRounding.AwayFromZero)
                                Else
                                    InputDataDeductionPer = System.Math.Round((clsCommon.myCDecimal(dr("lower_range")) - TempInputData) * clsCommon.myCDecimal(dr("Deduction_Ratio2")), 2, MidpointRounding.AwayFromZero)
                                End If
                            Else
                                InputDataDeductionPer = System.Math.Round(clsCommon.myCDecimal(dr("Deduction_Ratio2")), 2, MidpointRounding.AwayFromZero)
                            End If
                            If clsCommon.myCdbl(InputDataDeductionPer) > 0 Then
                                Exit For
                            End If
                        ElseIf TempInputData >= clsCommon.myCDecimal(dr("Deduction_lower_range3")) AndAlso TempInputData <= clsCommon.myCDecimal(dr("Deduction_upper_range3")) Then
                            NetResult = True
                            If clsCommon.myCDecimal(dr("Deduction_Method")) = 0 Then
                                If clsCommon.myCDecimal(dr("Deduction_lower_range3")) > clsCommon.myCDecimal(dr("upper_range")) Then
                                    TempDedPercentage = ((clsCommon.myCDecimal(dr("Deduction_upper_range")) - clsCommon.myCDecimal(dr("upper_range"))) * clsCommon.myCDecimal(dr("Deduction_Ratio")))
                                    TempDedPercentage += ((clsCommon.myCDecimal(dr("Deduction_upper_range2")) - clsCommon.myCDecimal(dr("Deduction_upper_range"))) * clsCommon.myCDecimal(dr("Deduction_Ratio2")))
                                    InputDataDeductionPer = System.Math.Round(TempDedPercentage + ((TempInputData - clsCommon.myCDecimal(dr("Deduction_upper_range2"))) * clsCommon.myCDecimal(dr("Deduction_Ratio3"))), 2, MidpointRounding.AwayFromZero)
                                ElseIf clsCommon.myCDecimal(dr("Deduction_lower_range3")) < clsCommon.myCDecimal(dr("lower_range")) Then
                                    TempDedPercentage = ((clsCommon.myCDecimal(dr("lower_range")) - clsCommon.myCDecimal(dr("Deduction_lower_range"))) * clsCommon.myCDecimal(dr("Deduction_Ratio")))
                                    TempDedPercentage += ((clsCommon.myCDecimal(dr("Deduction_lower_range")) - clsCommon.myCDecimal(dr("Deduction_lower_range2"))) * clsCommon.myCDecimal(dr("Deduction_Ratio2")))
                                    InputDataDeductionPer = System.Math.Round(TempDedPercentage + ((clsCommon.myCDecimal(dr("Deduction_lower_range2")) - TempInputData) * clsCommon.myCDecimal(dr("Deduction_Ratio3"))), 2, MidpointRounding.AwayFromZero)
                                Else
                                    InputDataDeductionPer = System.Math.Round((clsCommon.myCDecimal(dr("lower_range")) - TempInputData) * clsCommon.myCDecimal(dr("Deduction_Ratio3")), 2, MidpointRounding.AwayFromZero)
                                End If
                            Else
                                InputDataDeductionPer = System.Math.Round(clsCommon.myCDecimal(dr("Deduction_Ratio3")), 2, MidpointRounding.AwayFromZero)
                            End If
                            If clsCommon.myCdbl(InputDataDeductionPer) > 0 Then
                                Exit For
                            End If
                        Else
                            InputDataDeductionPer = 0
                            NetResult = False
                            Exit For
                        End If
                    End If
                Next
            End If
            Dim strTempResult As Boolean = True
                    Dim strTempUD As Boolean = False

            If Not NetResult Then
                strTempResult = False
            End If
            If clsCommon.myCdbl(InputDataDeductionPer) > 0 Then
                strTempUD = True
            End If

            If strTempResult Then
                If strTempUD = True Then
                    txtAccept.Text = "Under Deviation"
                    txtAccept.BackColor = Color.Yellow
                ElseIf strTempUD = False Then
                    txtAccept.Text = "Accepted"
                    txtAccept.BackColor = Color.Green
                End If
            ElseIf Not strTempResult Then
                txtAccept.Text = "Rejected"
                txtAccept.BackColor = Color.Red
            End If
            txtAccept.Visible = True

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
End Class
