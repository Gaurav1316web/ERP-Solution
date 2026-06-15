Imports System.Data.SqlClient
Imports System.IO
Imports common
Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class frmTransporterPaymentProcess
    Inherits FrmMainTranScreen

#Region "Variables"
    Private isNewEntry As Boolean = True
    Private isInsideLoadData As Boolean = False
    Private isCellValueChangedOpen As Boolean = False
    Const colDate As String = "colDate"
    Const colTransporterBillNo As String = "colTransporterBillNo"
    Const ColTankerNo As String = "ColTankerNo"
    Const colTransporterCode As String = "colTransporterCode"
    Const colTransporterName As String = "colTransporterName"
    Const ColBankCode As String = "ColBankCode"
    Const ColBankName As String = "ColBankName"
    Const ColBankIFSC As String = "ColBankIFSC"
    Const ColAmount As String = "ColAmount"
    Const ColType As String = "ColType"
    Const ColKM As String = "ColKM"
    Const ColStation4 As String = "ColStation4"
    Const ColIceBox As String = "ColIceBox"
    Const ColGPSKM As String = "ColGPSKM"
    Dim TotalAmount As Decimal = 0
    Dim TotalDiesel As Decimal = 0
    Dim TotalQuantity As Decimal = 0
    Dim TotalBMCQuantity As Decimal = 0
    Dim Total_Toll_Tax As Decimal = 0
    Dim Total_Ice_Charge As Decimal = 0
    Dim Total_BMC_TOTAL As Decimal = 0
    Dim Total_fat_snf_shortage As Decimal = 0
    Dim Total_Amount As Decimal = 0
    Public EnableOnPrivateChkbox As Boolean = False
    Public tripValue As String = ""


#End Region

    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
    End Sub

    Private Sub frmTransporterPaymentProcess_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Try
        '    'clsDBFuncationality.ExecuteNonQuery("CREATE UNIQUE INDEX UQ_MCC_Document_Code ON TSPL_BMC_TRANSPORTER_BILL_DETAIL (MCC_Document_Code) WHERE MCC_Document_Code IS NOT NULL")
        '    clsDBFuncationality.ExecuteNonQuery("ALTER TABLE TSPL_BMC_TRANSPORTER_BILL_DETAIL ALTER COLUMN MCC_Document_Code VARCHAR(30) NULL;")
        '    clsDBFuncationality.ExecuteNonQuery("ALTER TABLE TSPL_BMC_TRANSPORTER_BILL_DETAIL ADD CONSTRAINT UQ_MCC_Document_Code UNIQUE (MCC_Document_Code);")
        '    clsDBFuncationality.ExecuteNonQuery("ALTER TABLE TSPL_BMC_TRANSPORTER_BILL_DETAIL ADD CONSTRAINT FK_MCC_Document_Code FOREIGN KEY (MCC_Document_Code) REFERENCES TSPL_MILK_COLLECTION_MCC(Document_No);")

        'Catch ex As Exception
        'End Try

        'If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
        '    qry = "CREATE UNIQUE INDEX Unique_Mupliple_Day ON TSPL_BMC_TRANSPORTER_BILL_DETAIL (MCC_Document_Code) WHERE MCC_Document_Code IS NOT NULL;"
        '    clsDBFuncationality.ExecuteNonQuery(qry)
        'End If

        Dim coll As New Dictionary(Of String, String)()
        coll = New Dictionary(Of String, String)()
        coll.Add("Document_Code", "varchar(30) NOT NULL PRIMARY KEY")
        coll.Add("Document_Date", "datetime NOT NULL")
        coll.Add("From_Date", "datetime NOT NULL")
        coll.Add("To_Date", "datetime NOT NULL")
        coll.Add("Status", "integer null")
        coll.Add("Type", "varchar(20) NULL")
        coll.Add("Comment", "varchar(100) NULL")
        coll.Add("Remarks", "varchar(100) NULL")
        coll.Add("Created_By", "varchar(12)  NOT NULL")
        coll.Add("Created_Date", "datetime  NOT NULL")
        coll.Add("Modify_By", "varchar(12)  NOT NULL")
        coll.Add("Modify_Date", "datetime NOT NULL")
        coll.Add("Posted_By", "varchar(12) NULL")
        coll.Add("Posted_Date", "datetime null")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", coll, Nothing, True, False, "", "Document_Code", "Document_Date", True)

        coll = New Dictionary(Of String, String)()
        coll.Add("PK_ID", "integer NOT NULL identity NOT FOR REPLICATION PRIMARY KEY")
        coll.Add("Document_Code", "varchar(30) NOT NULL References TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD(Document_Code)")
        coll.Add("Transporter_Bill_No", "varchar(30)  NULL References TSPL_BMC_TRANSPORTER_BILL_HEAD(Document_Code)")
        coll.Add("Transporter_Bill_Date", "datetime NULL")
        coll.Add("Tanker_No", "varchar(20) NULL ")
        coll.Add("KM", "decimal (18,2) NULL")
        coll.Add("Type", "varchar(20) NULL")
        coll.Add("Transporter_Code", "varchar(30) NULL")
        coll.Add("Bank_Code", "varchar(50) NULL")
        coll.Add("Bank_Name", "varchar(50) NULL ")
        coll.Add("IFSC_Code", "varchar(50) NULL")
        coll.Add("Amount", "decimal (18,2) NULL")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_TRANSPORTER_PAYMENT_PROCESS_DETAIL", coll, Nothing, True, False, "TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD", "Document_Code", "", True)

        EnableOnPrivateChkbox = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableOnPrivateChkbox, clsFixedParameterCode.EnableOnPrivateChkbox, Nothing)) > 0)

        SetUserMgmtNew()
        RadPageView1.SelectedPage = RadPageViewPage1
        LoadBlankGrid()
        AddNew()
        RadGroupBox2.Enabled = True
        ReStoreGridLayout()
        'LoadHeadData()
    End Sub

    Sub LoadBlankGrid()

        Dim qry As String = String.Empty
        gv.Rows.Clear()
        gv.Columns.Clear()

        Dim repoDocumentNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDocumentNo.FormatString = ""
        repoDocumentNo.HeaderText = "TransporterBill No"
        repoDocumentNo.Name = colTransporterBillNo
        repoDocumentNo.Width = 150
        repoDocumentNo.IsVisible = False
        gv.MasterTemplate.Columns.Add(repoDocumentNo)

        Dim repoDate As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoDate.Format = DateTimePickerFormat.Custom
        'repoDate.CustomFormat = "dd-MM-yyyy"
        repoDate.CustomFormat = "dd/MMM/yyyy"
        repoDate.HeaderText = "Bill Date"
        repoDate.WrapText = True
        repoDate.FormatString = "{0:d}"
        repoDate.Name = colDate
        repoDate.ReadOnly = True
        repoDate.IsVisible = True
        repoDate.Width = 150
        gv.MasterTemplate.Columns.Add(repoDate)

        Dim repoTanker As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTanker.FormatString = ""
        repoTanker.HeaderText = "Tanker No"
        repoTanker.Name = ColTankerNo
        repoTanker.Width = 150
        repoTanker.IsVisible = True
        repoTanker.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoTanker)

        Dim repoKM As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoKM = New GridViewDecimalColumn()
        repoKM.FormatString = ""
        repoKM.HeaderText = "KM"
        repoKM.WrapText = True
        repoKM.Name = ColKM
        repoKM.Width = 150
        repoKM.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoKM.IsVisible = True
        repoKM.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoKM)

        Dim repoType As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoType.FormatString = ""
        repoType.HeaderText = "Type"
        repoType.Name = ColType
        repoType.Width = 150
        repoType.IsVisible = True
        repoType.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoType)

        Dim repoTransporter As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTransporter.FormatString = ""
        repoTransporter.HeaderText = "Transporter Code"
        repoTransporter.Name = colTransporterCode
        repoTransporter.Width = 150
        repoTransporter.IsVisible = True
        repoTransporter.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoTransporter)

        Dim repoTransporterName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTransporterName.FormatString = ""
        repoTransporterName.HeaderText = "Transporter Name"
        repoTransporterName.Name = colTransporterName
        repoTransporterName.Width = 150
        repoTransporterName.IsVisible = True
        repoTransporterName.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoTransporterName)

        Dim repoBankCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBankCode.FormatString = ""
        repoBankCode.HeaderText = "Bank Code"
        repoBankCode.Name = ColBankCode
        repoBankCode.Width = 150
        repoBankCode.IsVisible = True
        repoBankCode.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoBankCode)

        Dim repoBankName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBankName.FormatString = ""
        repoBankName.HeaderText = "Bank Name"
        repoBankName.Name = ColBankName
        repoBankName.Width = 150
        repoBankName.IsVisible = True
        repoBankName.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoBankName)

        Dim repoBankIFSC As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBankIFSC.FormatString = ""
        repoBankIFSC.HeaderText = "Bank IFSC"
        repoBankIFSC.Name = ColBankIFSC
        repoBankIFSC.Width = 150
        repoBankIFSC.IsVisible = True
        repoBankIFSC.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoBankIFSC)

        Dim repoAmount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmount = New GridViewDecimalColumn()
        repoAmount.FormatString = ""
        repoAmount.HeaderText = "Amount"
        repoAmount.WrapText = True
        repoAmount.Name = ColAmount
        repoAmount.Width = 150
        repoAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmount.IsVisible = True
        repoAmount.ReadOnly = True
        gv.MasterTemplate.Columns.Add(repoAmount)

        gv.AllowDeleteRow = True
        gv.AllowAddNewRow = False
        gv.ShowGroupPanel = False
        gv.AllowColumnReorder = False
        gv.AllowRowReorder = False
        gv.EnableSorting = False

        gv.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv.MasterTemplate.ShowRowHeaderColumn = False
        gv.TableElement.TableHeaderHeight = 40

        gv.SummaryRowsBottom.Clear()
        gv.ShowColumnHeaders = True

    End Sub

    Sub AddNew()
        BlankAllControls()
        LoadBlankGrid()
        ReStoreGridLayout()
        RadGroupBox2.Enabled = True
        btnSave.Text = "Save"
        rbtnBoth.IsChecked = True
    End Sub

    Sub BlankAllControls()
        fndDocNo.Value = ""
        txtTanker.arrValueMember = Nothing
        rbtnPrivate.IsChecked = False
        rbtnBMC.IsChecked = False
        rbtnBoth.IsChecked = False
        dtpDate.Value = clsCommon.GETSERVERDATE()
        dtpFromDate.Value = clsCommon.GETSERVERDATE()
        dtpToDate.Value = clsCommon.GETSERVERDATE()
        txtComment.Text = ""
        txtRemarks.Text = ""

        isNewEntry = True
        btnGo.Enabled = True
        gv.Rows.Clear()

        gv.SummaryRowsBottom.Clear()
        lblPrePending.Status = ERPTransactionStatus.Pending
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(MyBase.Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
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

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            'If txtTanker.arrValueMember IsNot Nothing AndAlso txtTanker.arrValueMember.Count > 0 Then
            LoadGridData()
            Dim TankerDetail As String = "Select Tanker_No 
                              From TSPL_BMC_TRANSPORTER_BILL_HEAD
                              Where Convert(Date,From_Date,103) >= Convert(Date,'" & clsCommon.GetPrintDate(dtpFromDate.Value) & "',103)
                              And Convert(Date,To_Date,103) <= Convert(Date,'" & clsCommon.GetPrintDate(dtpToDate.Value) & "',103)
                              And Status = 1"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(TankerDetail)
            Dim arrUserType As New ArrayList
            'Dim TankerNos As String = ""

            If dt.Rows.Count > 0 Then

                For Each dr As DataRow In dt.Rows
                    'arrUserType &= dr("Tanker_No").ToString().Trim() & ","
                    arrUserType.Add(dr("Tanker_No").ToString().Trim())
                Next

                'TankerNos = TankerNos.TrimEnd(","c)

            End If
            txtTanker.arrValueMember = arrUserType

            btnGo.Enabled = False

            'Dim arrUserType As New ArrayList

            'For i As Integer = 0 To obj.Arr.Count - 1
            '    arrUserType.Add(obj.Arr(i).Login_Type)
            'Next
            'txtUserType.arrValueMember = arrUserType
            'Else
            '    clsCommon.MyMessageBoxShow(Me, "Please Select TankerNo", Me.Text)
            'End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadGridData()
        Try
            Dim qry As String = " Select TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Code,TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Date,TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No,
                                KM,Case when Is_Private=1 then'Private' else 'BMC' END AS Typess,TSPL_VENDOR_MASTER.Vendor_Code,TSPL_VENDOR_MASTER.Vendor_Name,TSPL_VENDOR_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Bank_Name,TSPL_VENDOR_MASTER.IFSC_Code,
                                TSPL_BMC_TRANSPORTER_BILL_DETAIL.Amount from TSPL_BMC_TRANSPORTER_BILL_DETAIL 
                                LEFT OUTER JOIN TSPL_BMC_TRANSPORTER_BILL_HEAD ON TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code = TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Code
                                left outer join TSPL_TANKER_MASTER ON TSPL_TANKER_MASTER.Tanker_No = TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No
                                left outer join TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=TSPL_TANKER_MASTER.Tanker_Transporter_Code
                                where 2 = 2 and convert(date,TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(dtpFromDate.Value) + "',103) 
                                and convert(date,TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(dtpToDate.Value) + "' ,103) 
                                and TSPL_BMC_TRANSPORTER_BILL_HEAD.status=1 "
            If txtTanker.arrValueMember IsNot Nothing AndAlso txtTanker.arrValueMember.Count > 0 Then
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No in (" + clsCommon.GetMulcallString(txtTanker.arrValueMember) + ")"
            End If
            If rbtnPrivate.IsChecked Then
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Is_Private=1 "
            ElseIf rbtnBMC.IsChecked Then
                qry += " and TSPL_BMC_TRANSPORTER_BILL_HEAD.Is_Private=0 "
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            If dt.Rows.Count > 0 Then
                For ii = 0 To dt.Rows.Count - 1
                    gv.Rows.AddNew()
                    gv.CurrentRow.Cells(colTransporterBillNo).Value = clsCommon.myCstr(dt.Rows(ii)("Document_Code"))
                    gv.CurrentRow.Cells(colDate).Value = clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(ii)("Document_Date"), "dd/MMM/yyyy"))
                    gv.CurrentRow.Cells(ColTankerNo).Value = clsCommon.myCstr(dt.Rows(ii)("Tanker_No"))
                    gv.CurrentRow.Cells(ColType).Value = clsCommon.myCstr(dt.Rows(ii)("Typess"))
                    gv.CurrentRow.Cells(ColKM).Value = clsCommon.myCdbl(dt.Rows(ii)("KM"))
                    gv.CurrentRow.Cells(colTransporterCode).Value = clsCommon.myCstr(dt.Rows(ii)("Vendor_Code"))
                    gv.CurrentRow.Cells(colTransporterName).Value = clsCommon.myCstr(dt.Rows(ii)("Vendor_Name"))
                    gv.CurrentRow.Cells(ColBankCode).Value = clsCommon.myCstr(dt.Rows(ii)("Bank_Code"))
                    gv.CurrentRow.Cells(ColBankName).Value = clsCommon.myCstr(dt.Rows(ii)("Bank_Name"))
                    gv.CurrentRow.Cells(ColBankIFSC).Value = clsCommon.myCstr(dt.Rows(ii)("IFSC_Code"))
                    gv.CurrentRow.Cells(ColAmount).Value = clsCommon.myCdbl(dt.Rows(ii)("Amount"))

                Next
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveData(False)
    End Sub

    Sub SaveData(ByVal isPost As Boolean)
        Try
            If (AllowToSave()) Then
                Dim obj As New ClsTransporterPaymentProcess()
                obj.Document_Code = fndDocNo.Value
                obj.Document_Date = dtpDate.Value
                obj.From_Date = dtpFromDate.Value
                obj.To_Date = dtpToDate.Value
                obj.Remarks = txtRemarks.Text
                obj.Comment = txtComment.Text
                If rbtnPrivate.IsChecked Then
                    obj.Type = "Private"
                ElseIf rbtnBMC.IsChecked Then
                    obj.Type = "BMC"
                Else
                    obj.Type = "BOTH"
                End If

                obj.Arr = New List(Of ClsTransporterPaymentProcessDetail)
                For Each grow As GridViewRowInfo In gv.Rows
                    Dim objTr As New ClsTransporterPaymentProcessDetail()

                    objTr.Transporter_Bill_No = clsCommon.myCstr(grow.Cells(colTransporterBillNo).Value)
                    objTr.Transporter_Bill_Date = clsCommon.myCDate(grow.Cells(colDate).Value)
                    objTr.Tanker_No = clsCommon.myCstr(grow.Cells(ColTankerNo).Value)
                    objTr.KM = clsCommon.myCdbl(grow.Cells(ColKM).Value)
                    objTr.Transporter_Code = clsCommon.myCstr(grow.Cells(colTransporterCode).Value)
                    objTr.Type = clsCommon.myCstr(grow.Cells(ColType).Value)
                    objTr.Bank_Code = clsCommon.myCstr(grow.Cells(ColBankCode).Value)
                    objTr.Bank_Name = clsCommon.myCstr(grow.Cells(ColBankName).Value)
                    objTr.IFSC_Code = clsCommon.myCstr(grow.Cells(ColBankIFSC).Value)
                    objTr.Amount = clsCommon.myCdbl(grow.Cells(ColAmount).Value)

                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colTransporterBillNo).Value)) > 0 OrElse clsCommon.myLen(clsCommon.myCstr(grow.Cells(colDate).Value)) > 0 Then
                        obj.Arr.Add(objTr)

                    End If
                Next

                If (obj.SaveData(obj, isNewEntry)) Then
                    If Not isPost Then
                        common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                    End If
                    LoadData(obj.Document_Code, NavigatorType.Current)
                    btnGo.Enabled = False
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Function AllowToSave() As Boolean
        Try
            Xtra.TransactionValidity(dtpDate.Value)
            If txtTanker.arrValueMember IsNot Nothing AndAlso txtTanker.arrValueMember.Count < 0 Then
                Throw New Exception("Please Select TankerNo")
            End If
            Return True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
    End Function

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        DeleteData()
    End Sub

    Sub DeleteData()
        Try
            Dim Reason As String = ""
            If (myMessages.deleteConfirm()) Then
                If clsCancelLog.CheckForReasonOnDelete() Then
                    '' REASON FOR DELETE 
                    Dim frm As New FrmFreeTxtBox1
                    frm.Text = "Remarks for Delete"
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 Then
                        Exit Sub
                    Else
                        Reason = frm.strRmks
                    End If
                End If
                If (ClsTransporterPaymentProcess.DeleteData(fndDocNo.Value)) Then
                    saveCancelLog(Reason, "Delete", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function saveCancelLog(ByVal Reason As String, ByVal Activity_Type As String, Optional ByVal trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim obj As New clsCancelLog
        obj.Program_Code = Form_ID
        obj.DOCUMENT_NO = clsCommon.myCstr(Me.fndDocNo.Value)
        obj.REASON = Reason
        obj.ACTIVITY_TYPE = Activity_Type
        Return clsCancelLog.SaveData(obj, True, trans)
    End Function

    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        PostData()
    End Sub

    Sub PostData()
        Try
            Dim msg As String = ""
            Dim qry As String = ""
            Dim dt As DataTable = Nothing
            If (myMessages.postConfirm()) Then
                'If Not AllowToSave() Then
                '    Exit Sub
                'End If
                'SaveData(True)
                If (ClsTransporterPaymentProcess.PostData(fndDocNo.Value)) Then
                    msg = "Successfully Posted"
                End If
                common.clsCommon.MyMessageBoxShow(Me, msg, Me.Text)
                LoadData(fndDocNo.Value, NavigatorType.Current)
                btnGo.Enabled = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        CloseForm()
    End Sub

    Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub

    Sub LoadData(ByVal strDocumentNo As String, NavType As common.NavigatorType)
        Try
            btnSave.Enabled = True
            btnPost.Enabled = True
            btnDelete.Enabled = True
            btnSave.Text = "Update"
            fndDocNo.MyReadOnly = True
            isInsideLoadData = True
            BlankAllControls()
            LoadBlankGrid()
            isNewEntry = False

            Dim obj As New ClsTransporterPaymentProcess()
            obj = ClsTransporterPaymentProcess.GetData(strDocumentNo, NavType, True, Nothing)

            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
                isNewEntry = False
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                End If
            End If

            lblPrePending.Status = obj.Status
            fndDocNo.Value = obj.Document_Code
            dtpDate.Value = obj.Document_Date
            dtpFromDate.Value = obj.From_Date
            dtpToDate.Value = obj.To_Date
            txtComment.Text = obj.Comment
            txtRemarks.Text = obj.Remarks
            If obj.Type = "Private" Then
                rbtnPrivate.IsChecked = True
            ElseIf obj.Type = "BMC" Then
                rbtnBMC.IsChecked = True
            Else
                rbtnBoth.IsChecked = True
            End If

            Dim TankerDetail As String = "Select Tanker_No From TSPL_BMC_TRANSPORTER_BILL_HEAD
                              Where Convert(Date,From_Date,103) >= Convert(Date,'" & obj.From_Date & "',103)
                              And Convert(Date,To_Date,103) <= Convert(Date,'" & obj.To_Date & "',103) And Status = 1"

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(TankerDetail)
            Dim arrUserType As New ArrayList
            'Dim TankerNos As String = ""

            If dt.Rows.Count > 0 Then

                For Each dr As DataRow In dt.Rows
                    'arrUserType &= dr("Tanker_No").ToString().Trim() & ","
                    arrUserType.Add(dr("Tanker_No").ToString().Trim())
                Next

                'TankerNos = TankerNos.TrimEnd(","c)

            End If
            txtTanker.arrValueMember = arrUserType
            If obj.Arr IsNot Nothing Then
                Dim rowCount As Integer = 0
                Dim i As Integer = 0
                For Each objrow As ClsTransporterPaymentProcessDetail In obj.Arr
                    gv.Rows.AddNew()
                    gv.Rows(gv.Rows.Count - 1).Cells(colTransporterBillNo).Value = objrow.Transporter_Bill_No
                    gv.Rows(gv.Rows.Count - 1).Cells(colDate).Value = objrow.Transporter_Bill_Date
                    gv.Rows(gv.Rows.Count - 1).Cells(ColTankerNo).Value = objrow.Tanker_No
                    gv.Rows(gv.Rows.Count - 1).Cells(ColKM).Value = objrow.KM
                    gv.Rows(gv.Rows.Count - 1).Cells(ColType).Value = objrow.Type
                    gv.Rows(gv.Rows.Count - 1).Cells(colTransporterCode).Value = objrow.Transporter_Code
                    Dim qry1 As String = clsDBFuncationality.getSingleValue(" select Vendor_Name from TSPL_VENDOR_MASTER where Vendor_Code='" + objrow.Transporter_Code + "'")
                    gv.Rows(gv.Rows.Count - 1).Cells(colTransporterCode).Value = clsCommon.myCstr(qry1)
                    gv.Rows(gv.Rows.Count - 1).Cells(ColBankCode).Value = objrow.Bank_Code
                    gv.Rows(gv.Rows.Count - 1).Cells(ColBankName).Value = objrow.Bank_Name
                    gv.Rows(gv.Rows.Count - 1).Cells(ColBankIFSC).Value = objrow.IFSC_Code
                    gv.Rows(gv.Rows.Count - 1).Cells(ColAmount).Value = objrow.Amount
                Next
            End If


        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtTanker__My_Click(sender As Object, e As EventArgs) Handles txtTanker._My_Click
        Dim qry As String = " select Tanker_No as TankerNo,Tanker_Name as TankerName,Price_KM from   TSPL_TANKER_MASTER "
        txtTanker.arrValueMember = clsCommon.ShowMultipleSelectForm("CustMulSel", qry, "TankerNo", "TankerName", txtTanker.arrValueMember, txtTanker.arrDispalyMember)

    End Sub

    Private Sub fndDocNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndDocNo._MYValidating
        Try

            Dim qry As String = "select TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_Code ,convert(varchar,TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.Document_date,103) as Document_date,case when TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD.status =1  then 'Approved' else 'Pending' end as Status   
                             from TSPL_TRANSPORTER_PAYMENT_PROCESS_HEAD "
            fndDocNo.Value = clsCommon.ShowSelectForm("fmGroup_Code", qry, "Document_Code", "", fndDocNo.Value, "", isButtonClicked)
            If clsCommon.myLen(fndDocNo.Value) > 0 Then
                LoadData(fndDocNo.Value, NavigatorType.Current)
                btnGo.Enabled = False
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndDocNo__MYNavigator(sender As Object, e As EventArgs, NavType As NavigatorType) Handles fndDocNo._MYNavigator
        Try

            LoadData(fndDocNo.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        AddNew()
    End Sub
End Class