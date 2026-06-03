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
    Const ColBankIFSC As String = "ColBankName"
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

        EnableOnPrivateChkbox = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableOnPrivateChkbox, clsFixedParameterCode.EnableOnPrivateChkbox, Nothing)) > 0)

        SetUserMgmtNew()
        RadPageView1.SelectedPage = RadPageViewPage1
        LoadBlankGrid()
        AddNew()
        RadGroupBox2.Enabled = False
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

    End Sub

    Sub AddNew()
        BlankAllControls()
        LoadBlankGrid()
        ReStoreGridLayout()
        RadGroupBox2.Enabled = False
        btnSave.Text = "Save"
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
            LoadGridData()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadGridData()
        Try

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class