Imports System.Data.SqlClient
Imports System.IO
Imports common
Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class BMC_Transporter_Bill
    Inherits FrmMainTranScreen

#Region "Variables"
    Private isNewEntry As Boolean = True
    Private isInsideLoadData As Boolean = False
    Private isCellValueChangedOpen As Boolean = False
    Const colDate As String = "colDate"
    Const colDocumentNo As String = "colDocumentNo"
    Const ColCategory As String = "ColCategory"
    Const ColTrip As String = "ColTrip"
    Const ColKM As String = "ColKM"
    Const ColQuantity As String = "ColQuantity"
    Const ColAmount As String = "ColAmount"
    Const ColDiesel As String = "ColDiesel"
    Const ColStation As String = "ColStation"
    Const ColStation2 As String = "ColStation2"
    Const ColStation3 As String = "ColStation3"
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



    Private Sub txtTankerNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTankerNo._MYValidating
        Try
            Dim qry As String = "select Tanker_No as TankerNo,Tanker_Name as TankerName,Price_KM from   TSPL_TANKER_MASTER "
            txtTankerNo.Value = clsCommon.ShowSelectForm("RoutMasFND", qry, "TankerNo", "", txtTankerNo.Value, "", isButtonClicked)
            lblTankerDesc.Text = clsDBFuncationality.getSingleValue(" select Tanker_Name from TSPL_TANKER_MASTER where Tanker_No='" + txtTankerNo.Value + "'")

            Dim qry1 As String = clsDBFuncationality.getSingleValue(" select Tanker_Transporter_Code from TSPL_TANKER_MASTER where Tanker_No='" + txtTankerNo.Value + "'")
            txtTransporter.Text = clsDBFuncationality.getSingleValue(" select Vendor_Name from TSPL_VENDOR_MASTER where Vendor_Code='" + qry1 + "'")
            TxtBarelCap.Text = clsDBFuncationality.getSingleValue(" select Storage_Capacity from TSPL_TANKER_MASTER where Tanker_No='" + txtTankerNo.Value + "'")

            Dim KMQRY As String = " select Tanker_No as TankerNo,Tanker_Name as TankerName,Price_KM from   TSPL_TANKER_MASTER where Tanker_No ='" + txtTankerNo.Value + "'  "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(KMQRY)
            If dt.Rows.Count > 0 Then
                Dim kmrate As Double
                kmrate = clsCommon.myCdbl(dt.Rows(0).Item("Price_KM"))
                If kmrate <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Please Set KMRate for This Tanker", Me.Text)
                    txtTankerNo.Value = ""
                    lblTankerDesc.Text = ""
                    txtTransporter.Text = ""
                    TxtBarelCap.Text = ""
                    Exit Sub
                Else
                    TxtKMRate.Text = kmrate
                End If

            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub BMC_Transporter_Bill_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        RadGroupBox3.Enabled = False
        ReStoreGridLayout()
        LoadHeadData()

    End Sub

    Sub LoadHeadData()
        Try
            TxtKMRate.Text = clsDBFuncationality.getSingleValue(" Select top 1 KM_Rate from TSPL_BMC_TRANSPORTER_BILL_HEAD order by Document_Date desc")
            TxtDieselMinus.Text = clsDBFuncationality.getSingleValue(" Select top 1 Diesel_Rate_Minus from TSPL_BMC_TRANSPORTER_BILL_HEAD order by Document_Date desc")
            txtDieselplus.Text = clsDBFuncationality.getSingleValue(" Select top 1 Diesel_Rate_Plus from TSPL_BMC_TRANSPORTER_BILL_HEAD order by Document_Date desc")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadBlankGrid()
        Dim qry As String = String.Empty
        gv1.Rows.Clear()
        gv1.Columns.Clear()

        Dim repoDate As GridViewDateTimeColumn = New GridViewDateTimeColumn()
        repoDate.Format = DateTimePickerFormat.Custom
        'repoDate.CustomFormat = "dd-MM-yyyy"
        repoDate.CustomFormat = "dd/MMM/yyyy"
        repoDate.HeaderText = "Document Date"
        repoDate.WrapText = True
        repoDate.FormatString = "{0:d}"
        repoDate.Name = colDate
        If chkPrivate.Checked Then
            repoDate.ReadOnly = False
        Else
            repoDate.ReadOnly = True
        End If
        'repoDate.ReadOnly = True
        repoDate.IsVisible = True
        repoDate.Width = 150
        gv1.MasterTemplate.Columns.Add(repoDate)

        Dim repoDocumentNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDocumentNo.FormatString = ""
        repoDocumentNo.HeaderText = "Document No"
        repoDocumentNo.Name = colDocumentNo
        repoDocumentNo.Width = 150
        repoDocumentNo.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoDocumentNo)

        Dim repoCategory As GridViewComboBoxColumn = New GridViewComboBoxColumn()
        repoCategory.FormatString = ""
        repoCategory.HeaderText = "Category"
        repoCategory.Name = ColCategory
        repoCategory.DataSource = GetItemType()
        repoCategory.ValueMember = "Code"
        repoCategory.DisplayMember = "Code"
        'PaymentMode.DisplayMember = "RTGS"
        repoCategory.Width = 100
        If chkPrivate.Checked Then
            repoCategory.ReadOnly = False
        Else
            repoCategory.ReadOnly = True
        End If
        'repoCategory.ReadOnly = False
        gv1.MasterTemplate.Columns.Add(repoCategory)

        'Dim repoCategory As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        'repoCategory.FormatString = ""
        'repoCategory.HeaderText = "Category"
        'repoCategory.Name = ColCategory
        'repoCategory.Width = 150
        'repoCategory.IsVisible = True
        'If chkPrivate.Checked Then
        '    repoCategory.ReadOnly = False
        'Else
        '    repoCategory.ReadOnly = True
        'End If
        ''repoCategory.ReadOnly = True
        'gv1.MasterTemplate.Columns.Add(repoCategory)

        Dim repoStation As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoStation.FormatString = ""
        repoStation.HeaderText = "Station"
        repoStation.Name = ColStation
        repoStation.Width = 150
        repoStation.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoStation)

        Dim repoStation2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoStation2.FormatString = ""
        repoStation2.HeaderText = "Station 2"
        repoStation2.Name = ColStation2
        repoStation2.Width = 150
        repoStation2.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoStation2)

        Dim repoStation3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoStation3.FormatString = ""
        repoStation3.HeaderText = "Station 3"
        repoStation3.Name = ColStation3
        repoStation3.Width = 150
        repoStation3.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoStation3)

        Dim repoStation4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoStation4.FormatString = ""
        repoStation4.HeaderText = "Station 4"
        repoStation4.Name = ColStation4
        repoStation4.Width = 150
        repoStation4.IsVisible = True
        gv1.MasterTemplate.Columns.Add(repoStation4)

        Dim repoTrip As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTrip.FormatString = ""
        repoTrip.HeaderText = "Trip"
        repoTrip.Name = ColTrip
        repoTrip.Width = 150
        repoTrip.IsVisible = True
        If chkPrivate.Checked Then
            repoTrip.ReadOnly = False
        Else
            repoTrip.ReadOnly = True
        End If
        'repoTrip.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoTrip)

        'Dim IceBox As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        Dim IceBox As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        IceBox.FormatString = ""
        IceBox.HeaderText = "Ice Box"
        IceBox.Name = ColIceBox
        IceBox.Width = 150
        IceBox.IsVisible = True
        If chkPrivate.Checked Then
            IceBox.ReadOnly = False
        Else
            IceBox.ReadOnly = True
        End If
        gv1.MasterTemplate.Columns.Add(IceBox)

        Dim repoKM As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoKM = New GridViewDecimalColumn()
        repoKM.FormatString = ""
        repoKM.HeaderText = "KM"
        repoKM.WrapText = True
        repoKM.Name = ColKM
        repoKM.Width = 150
        repoKM.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoKM.VisibleInColumnChooser = False
        repoKM.ReadOnly = False
        gv1.MasterTemplate.Columns.Add(repoKM)

        Dim repoGPSKM As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoGPSKM = New GridViewDecimalColumn()
        repoGPSKM.FormatString = ""
        repoGPSKM.HeaderText = "GPSKM"
        repoGPSKM.WrapText = True
        repoGPSKM.Name = ColGPSKM
        repoGPSKM.Width = 150
        repoGPSKM.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoGPSKM.VisibleInColumnChooser = False
        repoGPSKM.ReadOnly = False
        gv1.MasterTemplate.Columns.Add(repoGPSKM)

        Dim repoQuantity As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoQuantity = New GridViewDecimalColumn()
        repoQuantity.FormatString = ""
        repoQuantity.HeaderText = "Quantity(KG)"
        repoQuantity.WrapText = True
        repoQuantity.Name = ColQuantity
        repoQuantity.Width = 150
        repoQuantity.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoQuantity.VisibleInColumnChooser = False
        If chkPrivate.Checked Then
            repoQuantity.ReadOnly = False
        Else
            repoQuantity.ReadOnly = True
        End If
        'repoQuantity.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoQuantity)

        Dim repoAmount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmount = New GridViewDecimalColumn()
        repoAmount.FormatString = ""
        repoAmount.HeaderText = "Amount"
        repoAmount.WrapText = True
        repoAmount.Name = ColAmount
        repoAmount.Width = 150
        repoAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmount.VisibleInColumnChooser = False
        repoAmount.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmount)

        Dim repoDiesel As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDiesel = New GridViewDecimalColumn()
        repoDiesel.FormatString = ""
        repoDiesel.HeaderText = "Diesel(RD)"
        repoDiesel.WrapText = True
        repoDiesel.Name = ColDiesel
        repoDiesel.Width = 150
        repoDiesel.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDiesel.VisibleInColumnChooser = False
        repoDiesel.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoDiesel)

        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False

        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40




        gv1.SummaryRowsBottom.Clear()
        gv1.ShowColumnHeaders = True
        Dim summaryRow As New GridViewSummaryRowItem()
        Dim item As GridViewSummaryItem = New GridViewSummaryItem(ColTrip, "{0:n2}", GridAggregateFunction.Sum)
        summaryRow.Add(item)
        Dim item1 As GridViewSummaryItem = New GridViewSummaryItem(ColGPSKM, "{0:n2}", GridAggregateFunction.Sum)
        summaryRow.Add(item1)
        Dim item2 As GridViewSummaryItem = New GridViewSummaryItem(ColKM, "{0:n2}", GridAggregateFunction.Sum)
        summaryRow.Add(item2)
        Dim item3 As GridViewSummaryItem = New GridViewSummaryItem(ColQuantity, "{0:n2}", GridAggregateFunction.Sum)
        summaryRow.Add(item3)
        Dim item4 As GridViewSummaryItem = New GridViewSummaryItem(ColAmount, "{0:n2}", GridAggregateFunction.Sum)
        summaryRow.Add(item4)
        Dim item5 As GridViewSummaryItem = New GridViewSummaryItem(ColDiesel, "{0:n2}", GridAggregateFunction.Sum)
        summaryRow.Add(item5)

        'summaryRow.Add(New GridViewSummaryItem(colDocumentNo, "Grand Total", GridAggregateFunction.Sum))
        'summaryRow.Add(New GridViewSummaryItem(ColTrip, "{0:n2}", GridAggregateFunction.Sum))
        'summaryRow.Add(New GridViewSummaryItem(ColGPSKM, "{0:n2}", GridAggregateFunction.Sum))
        'summaryRow.Add(New GridViewSummaryItem(ColKM, "{0:n2}", GridAggregateFunction.Sum))
        'summaryRow.Add(New GridViewSummaryItem(ColQuantity, "{0:n2}", GridAggregateFunction.Sum))
        'summaryRow.Add(New GridViewSummaryItem(ColAmount, "{0:n2}", GridAggregateFunction.Sum))
        'summaryRow.Add(New GridViewSummaryItem(ColDiesel, "{0:n2}", GridAggregateFunction.Sum))
        gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRow)
        gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom



    End Sub

    Private Function GetItemType() As DataTable
        'Dim dt As New DataTable()
        'dt.Columns.Add("Code", GetType(String))

        'Dim dr As DataRow = dt.NewRow()
        'dr("Code") = "BMC"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "RMG"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "NMG"
        'dt.Rows.Add(dr)

        'dr = dt.NewRow()
        'dr("Code") = "MCC"
        'dt.Rows.Add(dr)

        'Return dt
        Try
            Dim Qry As String = "Select Document_Code as Code from TSPL_BMC_Category_Master"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            Return dt
            'cboTransaction.DataSource = dt
            'cboTransaction.DisplayMember = "Name"
            'cboTransaction.ValueMember = "Code"
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function
    Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
        AddNew()
        LoadHeadData()
        chkPrivate.Checked = False
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Sub AddNew()
        BlankAllControls()
        LoadBlankGrid()
        ReStoreGridLayout()
        RadGroupBox3.Enabled = False
        btnSave.Text = "Save"
    End Sub

    Sub BlankAllControls()
        txtDocNo.Value = ""
        txtTankerNo.Value = ""
        lblTankerDesc.Text = ""
        txtTransporter.Text = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        TxtDieselMinus.Text = ""
        txtDieselplus.Text = ""
        TxtTollTax.Text = ""
        TxtIceCharge.Text = ""
        TxtKMRate.Text = ""
        TxtRateprorata.Text = ""
        TxtTDR.Text = ""
        TxtTankerprorata.Text = ""
        TxtBarelCap.Text = ""
        txtFatShortage.Text = ""
        txtFatShortageNMG.Text = ""
        TxtSnfShortage.Text = ""
        TXTSNFShortageNMG.Text = ""
        TxtFatRate.Text = ""
        TxtFatRateNMG.Text = ""
        TxtSnfRate.Text = ""
        TxtSnfRateNMG.Text = ""
        TxtTotalAddition.Text = ""
        TxtTotalDeduction.Text = ""
        TxtGrossAmount.Text = ""
        TxtTotalFatSnfShortage.Text = ""
        TxtTotalIceCharge.Text = ""
        TxtTotalTollTax.Text = ""
        TxtTotalAmount.Text = ""
        TxtBMCProrataamt.Text = ""
        TxtBMCDiesel.Text = ""
        TxtBMCTotal.Text = ""
        txtComment.Text = ""
        txtRemarks.Text = ""
        NLabourCharge.Text = ""
        NESIEmployee.Text = ""
        NESIEmployer.Text = ""
        NTotalESI.Text = ""
        NTotalPF.Text = ""
        NAdmincharge.Text = ""
        NEDLICharge.Text = ""
        NempPF.Text = ""
        TotalEPFAmt.Text = ""
        TotalESIAmt.Text = ""
        NDailyWage.Text = ""
        NCOESI_PER.Text = ""
        NEMPESI_PER.Text = ""
        NEMPEPF_PER.Text = ""
        NACCOEPF_PER.Text = ""


        isNewEntry = True
        btnGo.Enabled = True
        gv1.Rows.Clear()

        gv1.SummaryRowsBottom.Clear()
        UsLock1.Status = ERPTransactionStatus.Pending
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        LoadBlankGrid()
        loadGridData()
        FatSnfShortageDetail()
        FatSnfRate()
        FatSnfCalculation()
        PFESIDEDUCTION()
        'IceCharge(Nothing)
        'AddSummaryRowToGrid()

    End Sub

    Sub PFESIDEDUCTION()
        Try
            Dim Rowcount As Double = 0
            Dim Labourcharge As Double = 0
            Dim EmployeePF As Double = 0
            Dim EmployerPF As Double = 0
            Dim AdminCharge As Double = 0
            Dim EDLICHARGE As Double = 0
            Dim TOTALPF As Double = 0
            Dim ESIEmployee As Double = 0
            Dim ESIEmployer As Double = 0
            Dim TotalESI As Double = 0
            Dim ESIEmployer1 As Double = 0
            Dim ESIEmployee1 As Double = 0
            Dim employeepf1 As Double = 0
            Dim adminc As Double = 0

            Dim qry As String = " Select Labour_Charge,PFRULE_CODE,ESIRULE_CODE from tspl_tanker_master where tanker_no = '" + clsCommon.myCstr(txtTankerNo.Value) + "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            Dim LabourC As Double = clsCommon.myCdbl(dt.Rows(0)("Labour_Charge"))
            Dim ESI As String = clsCommon.myCstr(dt.Rows(0)("ESIRULE_CODE"))
            Dim EPF As String = clsCommon.myCstr(dt.Rows(0)("PFRULE_CODE"))

            Rowcount = gv1.Rows.Count - 1

            Labourcharge = Rowcount * LabourC

            Dim esiqry As String = " Select COESI_PER,EMPESI_PER from TSPL_ESI_RULE_MASTER where ESIRULE_CODE = '" + clsCommon.myCstr(ESI) + "'"
            Dim dtesi As DataTable = clsDBFuncationality.GetDataTable(esiqry)
            If dtesi.Rows.Count > 0 Then
                ESIEmployer1 = clsCommon.myCdbl(dtesi.Rows(0)("COESI_PER"))  ''3.25
                ESIEmployee1 = clsCommon.myCdbl(dtesi.Rows(0)("EMPESI_PER"))  ''0.75
                ESIEmployee = (Labourcharge * ESIEmployee1) / 100
                ESIEmployer = (Labourcharge * ESIEmployer1) / 100

            Else
                ESIEmployee = 0
                ESIEmployer = 0
            End If

            Dim epfqry As String = " select EMPEPF_PER,ACCOEPF_PER from TSPL_PF_RULE_MASTER where PFRULE_CODE = '" + clsCommon.myCstr(EPF) + "' "
            Dim dtpf As DataTable = clsDBFuncationality.GetDataTable(epfqry)
            If dtpf.Rows.Count > 0 Then
                employeepf1 = clsCommon.myCdbl(dtpf.Rows(0)("EMPEPF_PER"))  ''12
                adminc = clsCommon.myCdbl(dtpf.Rows(0)("ACCOEPF_PER"))  ''0.5
                EmployeePF = (Labourcharge * employeepf1) / 100
                EmployerPF = (Labourcharge * employeepf1) / 100
                AdminCharge = (Labourcharge * adminc) / 100
                EDLICHARGE = (Labourcharge * adminc) / 100
                MyLabel37.Text = "Employee PF(CPF)" + employeepf1
            Else
                EmployeePF = 0
                AdminCharge = 0
                EDLICHARGE = 0
            End If

            TOTALPF = clsCommon.myCdbl(EmployeePF + EmployerPF + AdminCharge + EDLICHARGE)
            TotalESI = clsCommon.myCdbl(ESIEmployee + ESIEmployer)

            NDailyWage.Text = clsCommon.myCdbl(LabourC)
            NLabourCharge.Text = clsCommon.myCdbl(Labourcharge)
            NempPF.Text = clsCommon.myCdbl(EmployeePF)
            NemployerPF.Text = clsCommon.myCdbl(EmployerPF)
            NAdmincharge.Text = clsCommon.myCdbl(AdminCharge)
            NEDLICharge.Text = clsCommon.myCdbl(EDLICHARGE)
            NTotalPF.Text = clsCommon.myCdbl(TOTALPF)

            NESIEmployee.Text = clsCommon.myCdbl(ESIEmployee)
            NESIEmployer.Text = clsCommon.myCdbl(ESIEmployer)
            NTotalESI.Text = clsCommon.myCdbl(TotalESI)

            TotalEPFAmt.Text = clsCommon.myCdbl(TOTALPF)
            TotalESIAmt.Text = clsCommon.myCdbl(TotalESI)

            NCOESI_PER.Text = clsCommon.myCdbl(ESIEmployer1)
            NEMPESI_PER.Text = clsCommon.myCdbl(ESIEmployee1)
            NEMPEPF_PER.Text = clsCommon.myCdbl(employeepf1)
            NACCOEPF_PER.Text = clsCommon.myCdbl(adminc)

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub FatSnfCalculation()
        Dim FatShortage As Double = clsCommon.myCdbl(txtFatShortage.Text)
        Dim FatShortageNMG As Double = clsCommon.myCdbl(txtFatShortageNMG.Text)
        Dim SnfShortage As Double = clsCommon.myCdbl(TxtSnfShortage.Text)
        Dim SnfShortageNMG As Double = clsCommon.myCdbl(TXTSNFShortageNMG.Text)
        Dim fatrateNMG As Double = clsCommon.myCdbl(TxtFatRateNMG.Text)
        Dim fatrate As Double = clsCommon.myCdbl(TxtFatRate.Text)
        Dim snfRate As Double = clsCommon.myCdbl(TxtSnfRate.Text)
        Dim snfRateNMG As Double = clsCommon.myCdbl(TxtSnfRateNMG.Text)
        Dim FatAmt As Double = clsCommon.myCdbl((FatShortage * fatrate) + (FatShortageNMG * fatrateNMG))
        Dim SnfAmt As Double = clsCommon.myCdbl((SnfShortage * snfRate) + (SnfShortageNMG * snfRateNMG))
        TxtTotalFatSnfShortage.Text = clsCommon.myCdbl(FatAmt + SnfAmt)
    End Sub

    Sub IceChargeFun(ByVal IntRowNo As Integer)
        Try
            If Not chkPrivate.Checked Then
                Dim IceChargeAmt As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Ice_Charge from TSPL_TANKER_MASTER where Tanker_No = '" + clsCommon.myCstr(txtTankerNo.Value) + "' "))
                TxtIceCharge.Text = clsCommon.myCdbl(IceChargeAmt)

                Dim Iceboxqry As String = "Select Count(*)Ice_Box from TSPL_BMC_DCS_SAMPLE_RECEIVING where Tanker_No = '" + clsCommon.myCstr(txtTankerNo.Value) + "' 
            and CONVERT(DATE, Document_Date, 103) >= CONVERT(DATE, '" + clsCommon.GetPrintDate(txtFromDate.Value) + "', 103)
            and CONVERT(DATE, Document_Date, 103) <= CONVERT(DATE, '" + clsCommon.GetPrintDate(txtToDate.Value) + "', 103) and Trip = '" + tripValue + "' and Is_IceBox=1   "
                Dim dataicebox As DataTable = clsDBFuncationality.GetDataTable(Iceboxqry)

                Dim rowCount As Integer = clsCommon.myCstr(dataicebox.Rows(0)("Ice_Box"))
                Dim GrossAmt As Double = 0
                Dim totalAmt As Double = 0
                Dim Icecharge As Double = clsCommon.myCdbl(TxtIceCharge.Text)
                If rowCount > 0 Then
                    TxtTotalIceCharge.Text = clsCommon.myCdbl(Icecharge * rowCount)
                    GrossAmt = clsCommon.myCdbl(TxtGrossAmount.Text)
                    totalAmt = clsCommon.myCdbl(TxtTotalAmount.Text)
                    TxtTotalAmount.Text = totalAmt + clsCommon.myCdbl(TxtTotalIceCharge.Text)
                    TxtGrossAmount.Text = GrossAmt + clsCommon.myCdbl(TxtTotalIceCharge.Text)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
    Sub FatSnfRate()
        Try
            Dim qry As String = ""
            qry = " SELECT TOP 1  Loss_FAT_Rate,Loss_SNF_Rate FROM TSPL_OWN_BMC_GAIN_LOSS_RATE WHERE convert(date,Start_Date,103) <= convert(date,'" + txtDate.Value + "',103) and End_Date is null 
                    ORDER BY Start_Date DESC "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            Dim Fat_Rate As Decimal = 0
            Dim Snf_Rate As Decimal = 0
            Dim Fat_Rate_NMG As Decimal = 0
            Dim Snf_Rate_NMG As Decimal = 0
            For Each row As DataRow In dt.Rows
                Fat_Rate += clsCommon.myCdbl(row("Loss_FAT_Rate"))
                Snf_Rate += clsCommon.myCdbl(row("Loss_SNF_Rate"))
            Next
            TxtFatRate.Text = Fat_Rate
            TxtSnfRate.Text = Snf_Rate
            TxtFatRateNMG.Text = Fat_Rate_NMG
            TxtSnfRateNMG.Text = Snf_Rate_NMG


        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub FatSnfShortageDetail()
        Try
            Dim qry As String = ""
            qry = " select (sum(Original_FATKg) * 0.25/100) + sum(FATKG) as ADJFATKG,(sum(Original_SNFKg) * 0.25/100) + sum(SNFKG) as ADJSNFKG from 
                (Select * from 
                (Select * FROM 
                ( select MAX(convert(Varchar,TSPL_MILK_COLLECTION_MCC.Document_Date,103)) AS DocumentDate,
                (TSPL_MILK_COLLECTION_MCC.Tanker_No)Tanker_No,(TSPL_MILK_COLLECTION_MCC.Route_Code)Route_Code,
                sum(TSPL_MILK_COLLECTION_MCC_DETAIL.Qty) AS Original_Qty,
                sum(TSPL_MILK_COLLECTION_MCC_DETAIL.FATKg)Original_FATKg,
                sum(TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKg)Original_SNFKg,
 (max(TSPL_MILK_COLLECTION_MCC.Entered_FATKg)-sum(TSPL_MILK_COLLECTION_MCC_DETAIL.FATKg)) AS FATKG,
 (max(TSPL_MILK_COLLECTION_MCC.Entered_SNFKg)-sum(TSPL_MILK_COLLECTION_MCC_DETAIL.SNFKg)) AS SNFKG
from TSPL_MILK_COLLECTION_MCC 
LEFT OUTER JOIN TSPL_MILK_COLLECTION_MCC_DETAIL ON TSPL_MILK_COLLECTION_MCC_DETAIL.Document_No=TSPL_MILK_COLLECTION_MCC.Document_No
left outer join TSPL_TANKER_MASTER on TSPL_TANKER_MASTER.Tanker_No=TSPL_MILK_COLLECTION_MCC.Tanker_No
where 2 = 2 and convert(date,Document_Date,103)>=convert(date,'" + clsCommon.GetPrintDate(txtFromDate.Value) + "',103) 
and convert(date,Document_Date,103) <=convert(date,'" + clsCommon.GetPrintDate(txtToDate.Value) + "' ,103) 
and TSPL_MILK_COLLECTION_MCC.Tanker_No in ('" + clsCommon.myCstr(txtTankerNo.Value) + "') GROUP BY Route_Code,TSPL_MILK_COLLECTION_MCC.Tanker_No,Document_Date)ZZZ)xxxx)zzzz group by zzzz.Route_Code,zzzz.Tanker_No"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            Dim negativeFatKg As Decimal = 0
            Dim negativeSnfKg As Decimal = 0
            Dim Fat_KG_NMG As Decimal = 0
            Dim Snf_KG_NMG As Decimal = 0
            For Each row As DataRow In dt.Rows
                If clsCommon.myCdbl(row("ADJFATKG")) < 0 Then
                    negativeFatKg += clsCommon.myCdbl(row("ADJFATKG"))
                End If
                If clsCommon.myCdbl(row("ADJSNFKG")) < 0 Then
                    negativeSnfKg += clsCommon.myCdbl(row("ADJSNFKG"))
                End If
            Next
            txtFatShortage.Text = negativeFatKg
            TxtSnfShortage.Text = negativeSnfKg
            txtFatShortageNMG.Text = Fat_KG_NMG
            TXTSNFShortageNMG.Text = Snf_KG_NMG
            TxtTotalFatSnfShortage.Text = negativeFatKg + negativeSnfKg

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub loadGridData()
        Try
            Dim Slot1 As String = ""
            Dim Slot2 As String = ""
            Dim tanker_No As String = ""

            Slot1 = clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy")
            Slot2 = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
            tanker_No = clsCommon.myCstr(txtTankerNo.Value)

            Dim qr12y As String = " Select * from TSPL_BMC_TRANSPORTER_BILL_HEAD where CONVERT(DATE, TSPL_BMC_TRANSPORTER_BILL_HEAD.From_Date, 103) >= CONVERT(DATE, '" + Slot1 + "', 103) and
                                  CONVERT(DATE, TSPL_BMC_TRANSPORTER_BILL_HEAD.To_Date, 103) <= CONVERT(DATE, '" + Slot2 + "', 103) and Tanker_No = '" + tanker_No + "' "
            Dim dt12 As DataTable = clsDBFuncationality.GetDataTable(qr12y)

            If dt12.Rows.Count > 0 Then
                clsCommon.MyMessageBoxShow(Me, "Document is already created for this date range and Tanker", Me.Text)
                Exit Sub
            End If
            Dim qry As String = ""
            Dim qry1 As String = ""
            Dim qry2 As String = ""

            If clsCommon.myLen(txtDieselplus.Text) = 0 AndAlso clsCommon.myLen(TxtDieselMinus.Text) = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Enter Diesel Amount", Me.Text)
                txtDieselplus.Focus()
                Exit Sub
            End If
            If clsCommon.myLen(TxtKMRate.Text) = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Enter KMRate Amount", Me.Text)
                TxtKMRate.Focus()
                Exit Sub
            End If

            qry = " WITH RankedLocations AS (SELECT 'BMC' as Category,TSPL_MILK_COLLECTION_MCC.Document_No,CAST(TSPL_MILK_COLLECTION_MCC.Document_Date AS DATE) AS Document_Date,Route_Code,
                    Trip_No,TSPL_TANKER_MASTER.Storage_Capacity,TSPL_BULK_ROUTE_MASTER.Distance,TSPL_BULK_route_master_location.Location_Code,TSPL_LOCATION_MASTER.Location_Desc,ROW_NUMBER() OVER (
                    PARTITION BY TSPL_MILK_COLLECTION_MCC.Document_No ORDER BY TSPL_BULK_route_master_location.Location_Code) AS LocRank FROM TSPL_MILK_COLLECTION_MCC  
                    LEFT JOIN TSPL_TANKER_MASTER ON TSPL_TANKER_MASTER.Tanker_No = TSPL_MILK_COLLECTION_MCC.Tanker_No
                    LEFT JOIN TSPL_BULK_ROUTE_MASTER ON TSPL_BULK_ROUTE_MASTER.ROUTE_NO = TSPL_MILK_COLLECTION_MCC.Route_Code
                    LEFT JOIN TSPL_BULK_route_master_location ON TSPL_BULK_route_master_location.BULK_ROUTE_No = TSPL_MILK_COLLECTION_MCC.Route_Code
                    left outer join TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_BULK_route_master_location.Location_Code
                    WHERE CONVERT(DATE, TSPL_MILK_COLLECTION_MCC.Document_Date, 103) >= CONVERT(DATE, '" + clsCommon.GetPrintDate(txtFromDate.Value) + "', 103) and
                    CONVERT(DATE, TSPL_MILK_COLLECTION_MCC.Document_Date, 103) <= CONVERT(DATE, '" + clsCommon.GetPrintDate(txtToDate.Value) + "', 103)
                    AND TSPL_MILK_COLLECTION_MCC.Tanker_No = '" + clsCommon.myCstr(txtTankerNo.Value) + "')
                    SELECT Document_No,Document_Date,Route_Code,Trip_No,Storage_Capacity,Distance,MAX(CASE WHEN LocRank = 1 THEN Location_Desc END) AS Station_1,
                    MAX(CASE WHEN LocRank = 2 THEN Location_Desc END) AS Station_2,MAX(CASE WHEN LocRank = 3 THEN Location_Desc END) AS Station_3,MAX(CASE WHEN LocRank = 4 THEN Location_Desc END) AS Station_4
                    FROM RankedLocations GROUP BY Document_No,Document_Date,Route_Code,Trip_No,Storage_Capacity,Distance  "

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            qry1 = " Select max(Distance)Distance from TSPL_BULK_ROUTE_MASTER where Tanker_No='" + clsCommon.myCstr(txtTankerNo.Value) + "' "
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry1)
            qry2 = " Select Storage_Capacity from TSPL_TANKER_MASTER where Tanker_No='" + clsCommon.myCstr(txtTankerNo.Value) + "' "
            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry2)
            Dim ii As Integer = 0
            If dt.Rows.Count > 0 Then
                For ii = 0 To dt.Rows.Count - 1
                    gv1.Rows.AddNew()
                    gv1.CurrentRow.Cells(colDocumentNo).Value = clsCommon.myCstr(dt.Rows(ii)("Document_No"))
                    gv1.CurrentRow.Cells(colDate).Value = clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(ii)("Document_Date"), "dd/MMM/yyyy"))
                    Dim sTARTdate As DateTime = clsCommon.GetPrintDate(clsCommon.myCDate(dt.Rows(ii)("Document_Date"), "dd/MMM/yyyy"))
                    Dim docDate As String = Format(CDate(gv1.CurrentRow.Cells(colDate).Value), "dd/MMM/yyyy")
                    gv1.CurrentRow.Cells(ColCategory).Value = clsCommon.myCstr(dt.Rows(ii)("Category"))
                    gv1.CurrentRow.Cells(ColStation).Value = "Jaipur>>" & clsCommon.myCstr(dt.Rows(0)("Station_1"))
                    gv1.CurrentRow.Cells(ColStation2).Value = clsCommon.myCstr(dt.Rows(0)("Station_2"))
                    gv1.CurrentRow.Cells(ColStation3).Value = clsCommon.myCstr(dt.Rows(0)("Station_3"))
                    gv1.CurrentRow.Cells(ColStation4).Value = clsCommon.myCstr(dt.Rows(0)("Station_4"))
                    gv1.CurrentRow.Cells(ColTrip).Value = clsCommon.myCstr(dt.Rows(ii)("Trip_No"))
                    gv1.CurrentRow.Cells(ColKM).Value = clsCommon.myCdbl(dt.Rows(ii)("Distance"))
                    gv1.CurrentRow.Cells(ColQuantity).Value = clsCommon.myCdbl(dt.Rows(ii)("Storage_Capacity"))
                    gv1.CurrentRow.Cells(ColAmount).Value = clsCommon.myCdbl(clsCommon.myCdbl(dt.Rows(ii)("Distance"))) * clsCommon.myCdbl(TxtKMRate.Text)
                    gv1.CurrentRow.Cells(ColIceBox).Value = clsCommon.myCBool(clsDBFuncationality.getSingleValue("Select isnull(Is_IceBox,0) as Ice_Box from TSPL_BMC_DCS_SAMPLE_RECEIVING where Tanker_No= '" + txtTankerNo.Value + "' and Trip = '" + gv1.CurrentRow.Cells(ColTrip).Value + "' and Document_Date= '" + docDate + "'  "))
                    If clsCommon.myLen(txtDieselplus.Text) > 0 Then
                        gv1.CurrentRow.Cells(ColDiesel).Value = clsCommon.myCdbl(clsCommon.myCdbl(dt.Rows(ii)("Distance")) * clsCommon.myCdbl(txtDieselplus.Text))
                    Else
                        gv1.CurrentRow.Cells(ColDiesel).Value = -1 * clsCommon.myCdbl(clsCommon.myCdbl(dt.Rows(ii)("Distance")) * clsCommon.myCdbl(TxtDieselMinus.Text))
                    End If
                Next
                tripValue = gv1.CurrentRow.Cells(ColTrip).Value.ToString()
                IceChargeFun(ii)
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub




    Private Sub txtToDate_ValueChanged(sender As Object, e As EventArgs) Handles txtToDate.ValueChanged
        RadGroupBox3.Enabled = True
        If chkPrivate.Checked Then
            txtFatShortage.ReadOnly = False
            TxtSnfShortage.ReadOnly = False
            FatSnfRate()
        End If
    End Sub

    Private Sub gv1_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    UpdateCurrentRow(gv1.CurrentRow.Index, Nothing)
                    UpdateAllTotals(Nothing)
                End If

            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub UpdateAllTotals(ByVal trans As SqlTransaction)
        Try
            Dim BMCProrataAmt As Double = 0
            Dim BMCDiesel As Double = 0
            Dim BMCTotalAmt As Double = 0
            Dim TollTaxAmt As Double = 0
            Dim IceCharge As Double = 0
            Dim FatSnfShortage As Double = 0
            Dim AmtBfrGross As Double = 0
            Dim iceQty As Integer = 0
            Dim TotalAddition As Double = 0
            Dim TotalDeduction As Double = 0
            Dim TotalEPF As Double = 0
            Dim TotalESI As Double = 0
            For ii As Integer = 0 To gv1.Rows.Count - 1
                BMCProrataAmt = BMCProrataAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(ColAmount).Value)
                BMCDiesel = BMCDiesel + clsCommon.myCdbl(gv1.Rows(ii).Cells(ColDiesel).Value)
                iceQty += IIf(clsCommon.myCBool(gv1.Rows(ii).Cells(ColIceBox).Value), 1, 0)
            Next
            If chkPrivate.Checked Then
                FatSnfRate()
                TxtTotalIceCharge.Text = iceQty * clsCommon.myCDecimal(TxtIceCharge.Text)
            End If
            FatSnfCalculation()
            TxtBMCProrataamt.Text = clsCommon.myCdbl(BMCProrataAmt)
            TxtBMCDiesel.Text = clsCommon.myCdbl(BMCDiesel)
            TxtBMCTotal.Text = clsCommon.myCdbl(BMCProrataAmt + BMCDiesel)

            BMCTotalAmt = clsCommon.myCdbl(TxtBMCTotal.Text)
            TollTaxAmt = clsCommon.myCdbl(TxtTollTax.Text)
            IceCharge = clsCommon.myCdbl(TxtTotalIceCharge.Text)
            TxtTotalTollTax.Text = clsCommon.myCdbl(TollTaxAmt)
            TxtTotalAmount.Text = clsCommon.myCdbl(BMCTotalAmt + TollTaxAmt + clsCommon.myCDecimal(TxtTotalIceCharge.Text))
            AmtBfrGross = clsCommon.myCdbl(TxtTotalAmount.Text)
            FatSnfShortage = clsCommon.myCdbl(TxtTotalFatSnfShortage.Text)
            TotalAddition = clsCommon.myCdbl(TxtTotalAddition.Text)
            TotalDeduction = clsCommon.myCdbl(TxtTotalDeduction.Text)
            TotalEPF = clsCommon.myCdbl(NTotalPF.Text)
            TotalESI = clsCommon.myCdbl(NTotalESI.Text)
            TxtGrossAmount.Text = clsCommon.myCdbl(AmtBfrGross + TotalAddition - FatSnfShortage - TotalDeduction - TotalEPF - TotalESI)


        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer, ByVal trans As SqlTransaction)
        Try
            Dim dblKM As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColKM).Value)
            Dim dblGPSKM As Double = clsCommon.myCdbl(gv1.Rows(IntRowNo).Cells(ColGPSKM).Value)
            Dim KMRate As Double = clsCommon.myCdbl(TxtKMRate.Text)
            Dim DieselPlus As Double = clsCommon.myCdbl(txtDieselplus.Text)
            Dim DieselMinus As Double = clsCommon.myCdbl(TxtDieselMinus.Text)
            Dim dblBasicAmt As Double = 0
            Dim dblBasicDiesel As Double = 0

            Dim barel As Double = clsCommon.myCdbl(TxtBarelCap.Text)
            Dim tankerrate As Double = clsCommon.myCdbl(TxtTankerprorata.Text)
            Dim newkmrate As Double = 0
            If barel > 0 AndAlso tankerrate > 0 Then
                newkmrate = clsCommon.myCdbl(tankerrate / barel * KMRate)
            End If

            If newkmrate > 0 Then
                If dblGPSKM > 0 Then
                    dblBasicAmt = newkmrate * dblGPSKM
                Else
                    dblBasicAmt = newkmrate * dblKM
                End If
            Else
                If dblGPSKM > 0 Then
                    dblBasicAmt = KMRate * dblGPSKM
                Else
                    dblBasicAmt = KMRate * dblKM
                End If
            End If

            If DieselPlus > 0 Then
                If dblGPSKM > 0 Then
                    dblBasicDiesel += DieselPlus * dblGPSKM
                Else
                    dblBasicDiesel += DieselPlus * dblKM
                End If
            End If
            If DieselMinus > 0 Then
                If dblGPSKM > 0 Then
                    dblBasicDiesel -= DieselMinus * dblGPSKM
                Else
                    dblBasicDiesel -= DieselMinus * dblKM
                End If
            End If
            gv1.Rows(IntRowNo).Cells(ColAmount).Value = clsCommon.myCdbl(dblBasicAmt)
            gv1.Rows(IntRowNo).Cells(ColDiesel).Value = clsCommon.myCdbl(dblBasicDiesel)
            Dim docDate As String = Format(CDate(gv1.Rows(IntRowNo).Cells(colDate).Value), "dd/MMM/yyyy")
            If Not chkPrivate.Checked Then
                gv1.Rows(IntRowNo).Cells(ColIceBox).Value = clsCommon.myCBool(clsDBFuncationality.getSingleValue("Select isnull(Is_IceBox,0) from TSPL_BMC_DCS_SAMPLE_RECEIVING where Tanker_No= '" + txtTankerNo.Value + "' and Trip = '" + gv1.Rows(IntRowNo).Cells(ColTrip).Value + "' and Document_Date= '" + docDate + "'  "))
                IceChargeFun(IntRowNo)
            End If

            PFESIDEDUCTION()

            isCellValueChangedOpen = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        SaveData(False)
    End Sub
    Function AllowToSave() As Boolean
        Try
            Xtra.TransactionValidity(txtDate.Value)
            If txtDieselplus.Text < 0 Then
                Throw New Exception("Diesel value should be +ve")
            End If
            If TxtDieselMinus.Text < 0 Then
                Throw New Exception("Diesel value should be +ve")
            End If
            'Dim TxtfromDate As String = txtfromdate.value
            'Dim Slot1 As String = ""
            'Dim Slot2 As String = ""
            'Dim tanker_No As String = ""

            'Slot1 = clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy")
            'Slot2 = clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy")
            'tanker_No = clsCommon.myCstr(txtTankerNo.Value)

            'Dim qry As String = " Select * from TSPL_BMC_TRANSPORTER_BILL_HEAD where CONVERT(DATE, TSPL_BMC_TRANSPORTER_BILL_HEAD.From_Date, 103) >= CONVERT(DATE, '" + Slot1 + "', 103) and
            '                      CONVERT(DATE, TSPL_BMC_TRANSPORTER_BILL_HEAD.To_Date, 103) <= CONVERT(DATE, '" + Slot2 + "', 103) and Tanker_No = '" + tanker_No + "' "
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            'If dt.Rows.Count > 0 Then
            '    clsCommon.MyMessageBoxShow(Me, "Document is already created for this date range and Tanker", Me.Text)
            '    Return False
            'End If
            For ii As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colDate).Value)) > 0 Then
                    If clsCommon.GetDateWithStartTime(clsCommon.myCDate(gv1.Rows(ii).Cells(colDate).Value)) < clsCommon.GetDateWithStartTime(txtFromDate.Value) OrElse
                        clsCommon.GetDateWithStartTime(clsCommon.myCDate(gv1.Rows(ii).Cells(colDate).Value)) > clsCommon.GetDateWithStartTime(txtToDate.Value) Then
                        Throw New Exception("Invalid document date at " & (ii + 1))
                    End If
                End If
                UpdateCurrentRow(ii, Nothing)
            Next
            UpdateAllTotals(Nothing)
            Return True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
    End Function
    Sub SaveData(ByVal isPost As Boolean)
        Try
            If (AllowToSave()) Then
                Dim obj As New clsBMCTransporterBill()
                obj.Document_Code = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.From_Date = txtFromDate.Value
                obj.To_Date = txtToDate.Value
                obj.Tanker_No = txtTankerNo.Value
                obj.Toll_Tax = clsCommon.myCdbl(TxtTotalTollTax.Text)
                obj.Ice_Charge = clsCommon.myCdbl(TxtTotalIceCharge.Text)
                'obj.Fat_Shortage = clsCommon.myCdbl(txtFatShortage.Text)
                'obj.Snf_Shortage = clsCommon.myCdbl(TxtSnfShortage.Text)
                'obj.Fat_Snf_Shortage = clsCommon.myCdbl(TxtTotalFatSnfShortage.Text)
                'obj.Fat_Rate = clsCommon.myCdbl(TxtFatRate.Text)
                'obj.Snf_Rate = clsCommon.myCdbl(TxtSnfRate.Text)
                obj.Fat_Shortage = clsCommon.myCdbl(txtFatShortage.Text)
                obj.Fat_Shortage_NMG = clsCommon.myCdbl(txtFatShortageNMG.Text)
                obj.Snf_Shortage = clsCommon.myCdbl(TxtSnfShortage.Text)
                obj.Snf_Shortage_NMG = clsCommon.myCdbl(TXTSNFShortageNMG.Text)
                obj.Fat_Snf_Shortage = clsCommon.myCdbl(TxtTotalFatSnfShortage.Text)
                obj.Fat_Rate = clsCommon.myCdbl(TxtFatRate.Text)
                obj.Fat_Rate_NMG = clsCommon.myCdbl(TxtFatRateNMG.Text)
                obj.Snf_Rate = clsCommon.myCdbl(TxtSnfRate.Text)
                obj.Snf_Rate_NMG = clsCommon.myCdbl(TxtSnfRateNMG.Text)
                obj.Tanker_Capacity = clsCommon.myCdbl(TxtBarelCap.Text)
                obj.KM_Rate = clsCommon.myCdbl(TxtKMRate.Text)
                obj.Total_Amount = clsCommon.myCdbl(TxtTotalAmount.Text)
                obj.Gross_Amount = clsCommon.myCdbl(TxtGrossAmount.Text)
                obj.Total_Addition = clsCommon.myCdbl(TxtTotalAddition.Text)
                obj.Total_Deduction = clsCommon.myCdbl(TxtTotalDeduction.Text)
                obj.Diesel_Rate_Plus = clsCommon.myCdbl(txtDieselplus.Text)
                obj.Diesel_Rate_Minus = clsCommon.myCdbl(TxtDieselMinus.Text)
                obj.Total_Diesel = clsCommon.myCdbl(TxtBMCDiesel.Text)
                obj.Prorata_Amt = clsCommon.myCdbl(TxtBMCProrataamt.Text)
                obj.Total_Before_Calc = clsCommon.myCdbl(TxtBMCTotal.Text)
                obj.Is_Private = chkPrivate.Checked
                obj.Comment = txtComment.Text
                obj.Remarks = txtRemarks.Text
                obj.Total_PF_Amount = clsCommon.myCdbl(TotalEPFAmt.Text)
                obj.Total_ESI_Amount = clsCommon.myCdbl(TotalESIAmt.Text)
                obj.Labour_Charge = clsCommon.myCdbl(NDailyWage.Text)
                obj.Labour_Amt = clsCommon.myCdbl(NLabourCharge.Text)
                obj.Employee_PF = clsCommon.myCdbl(NempPF.Text)
                obj.Employer_PF = clsCommon.myCdbl(NemployerPF.Text)
                obj.Admin_Charge = clsCommon.myCdbl(NAdmincharge.Text)
                obj.EDLI_Charge = clsCommon.myCdbl(NEDLICharge.Text)
                obj.ESI_Employee = clsCommon.myCdbl(NESIEmployee.Text)
                obj.ESI_Employer = clsCommon.myCdbl(NESIEmployer.Text)
                obj.COESI_PER = clsCommon.myCdbl(NCOESI_PER.Text)
                obj.EMPESI_PER = clsCommon.myCdbl(NEMPESI_PER.Text)
                obj.EMPEPF_PER = clsCommon.myCdbl(NEMPEPF_PER.Text)
                obj.ACCOEPF_PER = clsCommon.myCdbl(NACCOEPF_PER.Text)
                Dim totalTrip As Double = 0
                Dim totalGPSKM As Double = 0
                Dim totalKM As Double = 0
                Dim totalQty As Double = 0
                Dim totalAmt As Double = 0
                Dim totalDiesel As Double = 0


                obj.Arr = New List(Of clsBMCTransporterBillDetail)
                For Each grow As GridViewRowInfo In gv1.Rows
                    Dim objTr As New clsBMCTransporterBillDetail()
                    objTr.MCC_Document_Code = clsCommon.myCstr(grow.Cells(colDocumentNo).Value)
                    objTr.Category = clsCommon.myCstr(grow.Cells(ColCategory).Value)
                    objTr.Station_1 = clsCommon.myCstr(grow.Cells(ColStation).Value)
                    objTr.Station_2 = clsCommon.myCstr(grow.Cells(ColStation2).Value)
                    objTr.Station_3 = clsCommon.myCstr(grow.Cells(ColStation3).Value)
                    objTr.Station_4 = clsCommon.myCstr(grow.Cells(ColStation4).Value)
                    objTr.Trip = clsCommon.myCdbl(grow.Cells(ColTrip).Value)
                    objTr.GPS_KM = clsCommon.myCdbl(grow.Cells(ColGPSKM).Value)
                    objTr.KM = clsCommon.myCdbl(grow.Cells(ColKM).Value)
                    objTr.Quantity_KG = clsCommon.myCdbl(grow.Cells(ColQuantity).Value)
                    objTr.Amount = clsCommon.myCdbl(grow.Cells(ColAmount).Value)
                    objTr.Diesel_RD = clsCommon.myCdbl(grow.Cells(ColDiesel).Value)
                    objTr.BMC_Date = clsCommon.myCDate(grow.Cells(colDate).Value)
                    objTr.Ice_Box = IIf(clsCommon.myCBool(grow.Cells(ColIceBox).Value), "Y", "N")


                    totalTrip += objTr.Trip
                    totalGPSKM += objTr.GPS_KM
                    totalKM += objTr.KM
                    totalQty += objTr.Quantity_KG
                    totalAmt += objTr.Amount
                    totalDiesel += objTr.Diesel_RD


                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colDocumentNo).Value)) > 0 OrElse clsCommon.myLen(clsCommon.myCstr(grow.Cells(colDate).Value)) > 0 Then
                        obj.Arr.Add(objTr)

                    End If
                Next
                'AddSummaryRow(totalTrip, totalGPSKM, totalKM, totalQty, totalAmt, totalDiesel)


                If (obj.SaveData(obj, isNewEntry)) Then
                    If Not isPost Then
                        common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                    End If
                    LoadData(obj.Document_Code, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    'Private Sub AddSummaryRow(totalTrip As Double, totalGPSKM As Double, totalKM As Double, totalQty As Double, totalAmt As Double, totalDiesel As Double)
    '    Try
    '        ' Remove existing summary rows
    '        gv1.SummaryRowsBottom.Clear()

    '        ' Enable footer and summary
    '        gv1.ShowColumnHeaders = True

    '        Dim summaryRow As New GridViewSummaryRowItem()

    '        summaryRow.Add(New GridViewSummaryItem(colDocumentNo, "Grand Total", GridAggregateFunction.Sum))
    '        summaryRow.Add(New GridViewSummaryItem(ColTrip, totalTrip.ToString("N2"), GridAggregateFunction.Sum))
    '        summaryRow.Add(New GridViewSummaryItem(ColGPSKM, totalGPSKM.ToString("N2"), GridAggregateFunction.Sum))
    '        summaryRow.Add(New GridViewSummaryItem(ColKM, totalKM.ToString("N2"), GridAggregateFunction.Sum))
    '        summaryRow.Add(New GridViewSummaryItem(ColQuantity, totalQty.ToString("N2"), GridAggregateFunction.Sum))
    '        summaryRow.Add(New GridViewSummaryItem(ColAmount, totalAmt.ToString("N2"), GridAggregateFunction.Sum))
    '        summaryRow.Add(New GridViewSummaryItem(ColDiesel, totalDiesel.ToString("N2"), GridAggregateFunction.Sum))

    '        gv1.SummaryRowsBottom.Add(summaryRow)
    '    Catch ex As Exception
    '        MessageBox.Show("Error adding summary row: " & ex.Message)
    '    End Try
    'End Sub




    Sub LoadData(ByVal strDocumentNo As String, NavType As common.NavigatorType)
        Try
            btnSave.Enabled = True
            btnPost.Enabled = True
            btnDelete.Enabled = True
            btnSave.Text = "Update"
            txtDocNo.MyReadOnly = True
            isInsideLoadData = True
            BlankAllControls()
            LoadBlankGrid()
            isNewEntry = False

            Dim obj As New clsBMCTransporterBill()
            obj = clsBMCTransporterBill.GetData(strDocumentNo, NavType, True, Nothing)

            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_Code) > 0) Then
                isNewEntry = False
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                End If

                UsLock1.Status = obj.Status
                txtDocNo.Value = obj.Document_Code
                txtDate.Value = obj.Document_Date
                txtFromDate.Value = obj.From_Date
                txtToDate.Value = obj.To_Date
                txtTankerNo.Value = obj.Tanker_No
                TxtTotalTollTax.Text = obj.Toll_Tax
                TxtTotalIceCharge.Text = obj.Ice_Charge
                'txtFatShortage.Text = obj.Fat_Shortage
                'TxtSnfShortage.Text = obj.Snf_Shortage
                'TxtTotalFatSnfShortage.Text = obj.Fat_Snf_Shortage
                'TxtFatRate.Text = obj.Fat_Rate
                'TxtSnfRate.Text = obj.Snf_Rate
                txtFatShortage.Text = obj.Fat_Shortage
                txtFatShortageNMG.Text = obj.Fat_Shortage_NMG
                TxtSnfShortage.Text = obj.Snf_Shortage
                TXTSNFShortageNMG.Text = obj.Snf_Shortage_NMG
                TxtTotalFatSnfShortage.Text = obj.Fat_Snf_Shortage
                TxtFatRate.Text = obj.Fat_Rate
                TxtFatRateNMG.Text = obj.Fat_Rate_NMG
                TxtSnfRate.Text = obj.Snf_Rate
                TxtSnfRateNMG.Text = obj.Snf_Rate_NMG

                TxtBarelCap.Text = obj.Tanker_Capacity
                TxtKMRate.Text = obj.KM_Rate
                TxtTotalAmount.Text = obj.Total_Amount
                TxtGrossAmount.Text = obj.Gross_Amount
                TxtTotalDeduction.Text = obj.Total_Deduction
                TxtTotalAddition.Text = obj.Total_Addition
                txtDieselplus.Text = obj.Diesel_Rate_Plus
                TxtDieselMinus.Text = obj.Diesel_Rate_Minus
                TxtBMCDiesel.Text = obj.Total_Diesel
                TxtBMCTotal.Text = obj.Total_Before_Calc
                TxtBMCProrataamt.Text = obj.Prorata_Amt
                Dim qry1 As String = clsDBFuncationality.getSingleValue(" select Tanker_Transporter_Code from TSPL_TANKER_MASTER where Tanker_No='" + txtTankerNo.Value + "'")
                txtTransporter.Text = clsDBFuncationality.getSingleValue(" select Vendor_Name from TSPL_VENDOR_MASTER where Vendor_Code='" + qry1 + "'")
                Dim fromDate As Date = obj.From_Date
                Dim toDate As Date = obj.To_Date
                Dim totalDays As Integer = DateDiff(DateInterval.Day, fromDate, toDate) + 1
                TxtTollTax.Text = obj.Toll_Tax

                chkPrivate.Checked = obj.Is_Private
                txtComment.Text = obj.Comment
                txtRemarks.Text = obj.Remarks

                NDailyWage.Text = obj.Labour_Charge
                NLabourCharge.Text = obj.Labour_Amt
                NempPF.Text = obj.Employee_PF
                NemployerPF.Text = obj.Employer_PF
                NAdmincharge.Text = obj.Admin_Charge
                NEDLICharge.Text = obj.EDLI_Charge
                NTotalPF.Text = obj.Total_PF_Amount
                NESIEmployee.Text = obj.ESI_Employee
                NESIEmployer.Text = obj.ESI_Employer
                NTotalESI.Text = obj.Total_ESI_Amount
                TotalEPFAmt.Text = obj.Total_PF_Amount
                TotalESIAmt.Text = obj.Total_ESI_Amount
                NCOESI_PER.Text = obj.COESI_PER
                NEMPESI_PER.Text = obj.EMPESI_PER
                NACCOEPF_PER.Text = obj.ACCOEPF_PER
                NEMPEPF_PER.Text = obj.EMPEPF_PER

                If obj.Arr IsNot Nothing Then
                    '               Dim qry As String = " SELECT TSPL_MILK_COLLECTION_MCC.Document_No,Cast(TSPL_MILK_COLLECTION_MCC.Document_Date as Date)Document_Date,Route_Code,Trip_No,TSPL_TANKER_MASTER.Storage_Capacity,TSPL_BULK_ROUTE_MASTER.Distance
                    '                                       FROM TSPL_MILK_COLLECTION_MCC  
                    '               left outer join TSPL_TANKER_MASTER ON TSPL_TANKER_MASTER.Tanker_No=TSPL_MILK_COLLECTION_MCC.Tanker_No
                    '               left outer join TSPL_BULK_ROUTE_MASTER ON TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_COLLECTION_MCC.Route_Code
                    'where convert(date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) >= convert(date,'" + txtFromDate.Value + "' ,103)
                    '               and  convert(date,TSPL_MILK_COLLECTION_MCC.Document_Date,103) <= convert(date,'" + txtToDate.Value + "' ,103)
                    '               and TSPL_MILK_COLLECTION_MCC.Tanker_No = '" & txtTankerNo.Value & "' "
                    '               Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    Dim rowCount As Integer = 0
                    Dim i As Integer = 0
                    For Each objrow As clsBMCTransporterBillDetail In obj.Arr
                        gv1.Rows.AddNew()
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDocumentNo).Value = objrow.MCC_Document_Code
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCategory).Value = objrow.Category
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColGPSKM).Value = objrow.GPS_KM
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColKM).Value = objrow.KM
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColAmount).Value = objrow.Amount
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColDiesel).Value = objrow.Diesel_RD
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColStation).Value = objrow.Station_1
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColStation2).Value = objrow.Station_2
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColStation3).Value = objrow.Station_3
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColStation4).Value = objrow.Station_4
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColTrip).Value = objrow.Trip
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColQuantity).Value = objrow.Quantity_KG
                        gv1.Rows(gv1.Rows.Count - 1).Cells(ColCategory).Value = objrow.Category
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colDate).Value = objrow.BMC_Date
                        If clsCommon.CompairString(objrow.Ice_Box, "Y") = CompairStringResult.Equal Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(ColIceBox).Value = True
                            rowCount += 1
                        End If
                        i += 1
                        'AddSummaryRowToGrid()
                    Next
                    If Not chkPrivate.Checked Then
                        rowCount = i
                    End If

                    Dim Icecharges As Double = clsCommon.myCdbl(TxtTotalIceCharge.Text)
                    If rowCount > 0 Then
                        TxtIceCharge.Text = clsCommon.myCdbl(Icecharges / rowCount)
                    End If
                End If
                'LoadPFESIDeduction()
                'PFESIDEDUCTION()
            End If
            ReStoreGridLayout()

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
        End Try
    End Sub

    Sub LoadPFESIDeduction()
        Try

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub AddSummaryRowToGrid()
        Try
            ' Enable footer row
            gv1.ShowColumnHeaders = True

            ' Clear existing summary rows
            gv1.SummaryRowsBottom.Clear()

            ' Calculate totals
            Dim totalTrip As Double = 0
            Dim totalGPSKM As Double = 0
            Dim totalKM As Double = 0
            Dim totalQty As Double = 0
            Dim totalAmt As Double = 0
            Dim totalDiesel As Double = 0

            For Each row As GridViewRowInfo In gv1.Rows
                totalTrip += clsCommon.myCdbl(row.Cells(ColTrip).Value)
                totalGPSKM += clsCommon.myCdbl(row.Cells(ColGPSKM).Value)
                totalKM += clsCommon.myCdbl(row.Cells(ColKM).Value)
                totalQty += clsCommon.myCdbl(row.Cells(ColQuantity).Value)
                totalAmt += clsCommon.myCdbl(row.Cells(ColAmount).Value)
                totalDiesel += clsCommon.myCdbl(row.Cells(ColDiesel).Value)
            Next

            ' Add summary row
            'Dim summaryRow As New GridViewSummaryRowItem()
            'summaryRow.Add(New GridViewSummaryItem(ColCategory, "Grand Total", GridAggregateFunction.Sum))
            'summaryRow.Add(New GridViewSummaryItem(ColTrip, totalTrip.ToString("N2"), GridAggregateFunction.Sum))
            'summaryRow.Add(New GridViewSummaryItem(ColGPSKM, totalGPSKM.ToString("N2"), GridAggregateFunction.Sum))
            'summaryRow.Add(New GridViewSummaryItem(ColKM, totalKM.ToString("N2"), GridAggregateFunction.Sum))
            'summaryRow.Add(New GridViewSummaryItem(ColQuantity, totalQty.ToString("N2"), GridAggregateFunction.Sum))
            'summaryRow.Add(New GridViewSummaryItem(ColAmount, totalAmt.ToString("N2"), GridAggregateFunction.Sum))
            'summaryRow.Add(New GridViewSummaryItem(ColDiesel, totalDiesel.ToString("N2"), GridAggregateFunction.Sum))

            'gv1.SummaryRowsBottom.Add(summaryRow)
        Catch ex As Exception
            MessageBox.Show("Error while adding summary row: " & ex.Message)
        End Try
    End Sub

    Function saveCancelLog(ByVal Reason As String, ByVal Activity_Type As String, Optional ByVal trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim obj As New clsCancelLog
        obj.Program_Code = Form_ID
        obj.DOCUMENT_NO = clsCommon.myCstr(Me.txtDocNo.Value)
        obj.REASON = Reason
        obj.ACTIVITY_TYPE = Activity_Type
        Return clsCancelLog.SaveData(obj, True, trans)
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
                If (clsBMCTransporterBill.DeleteData(txtDocNo.Value)) Then
                    saveCancelLog(Reason, "Delete", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

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
                If (clsBMCTransporterBill.PostData(txtDocNo.Value)) Then
                    msg = "Successfully Posted"
                End If
                common.clsCommon.MyMessageBoxShow(Me, msg, Me.Text)
                LoadData(txtDocNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDocNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDocNo._MYValidating
        Try

            Dim qry As String = "select TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code ,convert(varchar,TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_date,103) as Document_date,TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No,case when TSPL_BMC_TRANSPORTER_BILL_HEAD.status =1  then 'Approved' else 'Pending' end as Status   
                             from TSPL_BMC_TRANSPORTER_BILL_HEAD "
            txtDocNo.Value = clsCommon.ShowSelectForm("fmGroup_Code", qry, "Document_Code", "", txtDocNo.Value, "", isButtonClicked)
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                LoadData(txtDocNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub txtDocNo__MYNavigator(sender As Object, e As EventArgs, NavType As NavigatorType) Handles txtDocNo._MYNavigator
        Try

            LoadData(txtDocNo.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiImport_Click(sender As Object, e As EventArgs) Handles rmiImport.Click
        Dim ReportID As String = MyBase.Form_ID
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
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
            End If
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
            ''---------------
        End If
    End Sub

    Private Sub rmiExport_Click(sender As Object, e As EventArgs) Handles rmiExport.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(MyBase.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(MyBase.Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
                        gv1.Columns(ii).IsVisible = False
                        gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub

    Private Sub btnReverse_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select a document before reversing.", Me.Text)
                Exit Sub
            Else
                If clsCommon.MyMessageBoxShow(Me, "Do you want to Reverse and unpost the current Document" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    '' REASON FOR DELETE 
                    Dim Reason As String = ""
                    Dim frm As New FrmFreeTxtBox1
                    frm.Text = "Remarks for Reverse"
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 Then
                        Exit Sub
                    Else
                        Reason = frm.strRmks
                    End If

                    clsBMCTransporterBill.ReverseAndUnpost(txtDocNo.Value)
                    clsCommon.MyMessageBoxShow(Me, "Task done Successfully", Me.Text)
                    LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub



    Private Sub TxtDieselMinus_Validating(sender As Object, e As EventArgs) Handles TxtDieselMinus.Validating
        For ii As Integer = 0 To gv1.Rows.Count - 1
            UpdateCurrentRow(ii, Nothing)
        Next
        UpdateAllTotals(Nothing)


        'Dim rate As Decimal

        '' Validate and parse the rate
        'If Decimal.TryParse(TxtDieselMinus.Text, rate) Then
        '    For Each row As GridViewRowInfo In gv1.Rows
        '        If TypeOf row Is Telerik.WinControls.UI.GridViewDataRowInfo Then
        '            'If Not row.IsNewRow Then
        '            Dim gpskm As Decimal = 0
        '            Dim km As Decimal = 0

        '            ' Safely try to read GPSKM and KM
        '            Decimal.TryParse(row.Cells("ColGPSKM").Value, gpskm)
        '            Decimal.TryParse(row.Cells("ColKM").Value, km)

        '            ' Use GPSKM if it's greater than 0, otherwise use KM
        '            Dim usedKM As Decimal = If(gpskm > 0, gpskm, km)

        '            ' Calculate Amount = usedKM * rate
        '            row.Cells("ColDiesel").Value = usedKM * rate
        '        End If
        '    Next
        'End If
    End Sub

    Private Sub BMC_Transporter_Bill_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Try
            If e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
                If MyBase.isReverse Then
                    Dim frm As New FrmPWD(Nothing)
                    frm.strType = clsFixedParameterType.SIR
                    frm.strCode = clsFixedParameterCode.SIReversAndCreate
                    frm.ShowDialog()
                    If frm.isPasswordCorrect Then
                        btnReverse.Visible = True
                    End If
                Else
                    clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
                End If
            ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
                CloseForm()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim User_Name As String = objCommonVar.CurrentUser
            Dim FinalQuery As String = ""
            FinalQuery = "   SELECT '" & User_Name & "' as UserName,ROW_NUMBER() OVER (PARTITION BY TSPL_MILK_COLLECTION_MCC.Document_No ORDER BY TSPL_MILK_COLLECTION_MCC.Document_No) AS SrNo
,TSPL_BMC_TRANSPORTER_BILL_DETAIL.Category,
 TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Code,TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code,
 TSPL_MILK_COLLECTION_MCC.Document_No,Cast(TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Date as Date)Document_Date,TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Date as HeaderDate,Route_Code,TSPL_BMC_TRANSPORTER_BILL_DETAIL.Trip as Trip_No,TSPL_TANKER_MASTER.Storage_Capacity,TSPL_BULK_ROUTE_MASTER.Distance
, TSPL_BMC_TRANSPORTER_BILL_DETAIL.STATION_1, TSPL_BMC_TRANSPORTER_BILL_DETAIL.STATION_2,TSPL_BMC_TRANSPORTER_BILL_DETAIL.STATION_3,TSPL_BMC_TRANSPORTER_BILL_DETAIL.STATION_4, TSPL_BMC_TRANSPORTER_BILL_DETAIL.GPS_KM,TSPL_BMC_TRANSPORTER_BILL_DETAIL.KM     ,TSPL_BMC_TRANSPORTER_BILL_DETAIL.AMOUNT,TSPL_BMC_TRANSPORTER_BILL_DETAIL.DIESEL_RD
,TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_no,TSPL_BMC_TRANSPORTER_BILL_HEAD.toll_tax,
TSPL_BMC_TRANSPORTER_BILL_HEAD.ice_charge,TSPL_BMC_TRANSPORTER_BILL_HEAD.Fat_Shortage,TSPL_BMC_TRANSPORTER_BILL_HEAD.Snf_Shortage,TSPL_BMC_TRANSPORTER_BILL_HEAD.Fat_Snf_Shortage,TSPL_BMC_TRANSPORTER_BILL_HEAD.Fat_Rate,TSPL_BMC_TRANSPORTER_BILL_HEAD.Snf_Rate
,TSPL_TANKER_MASTER.Description as Descriptions,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Comp_Code,TSPL_COMPANY_MASTER.Add1,TSPL_COMPANY_MASTER.State,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.City_Code
,convert(date,TSPL_BMC_TRANSPORTER_BILL_HEAD.From_Date,103)From_Date,

convert(date,TSPL_BMC_TRANSPORTER_BILL_HEAD.To_Date,103)To_Date
,TSPL_BMC_TRANSPORTER_BILL_HEAD.KM_Rate,TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_Capacity
,	Diesel_Rate_Plus,	Diesel_Rate_Minus	,Total_Diesel,Total_Amount,	Gross_Amount
,TSPL_BMC_TRANSPORTER_BILL_HEAD.comment,TSPL_BMC_TRANSPORTER_BILL_HEAD.Remarks
,TSPL_BMC_TRANSPORTER_BILL_HEAD.Fat_Shortage_NMG,TSPL_BMC_TRANSPORTER_BILL_HEAD.Snf_Shortage_NMG,TSPL_BMC_TRANSPORTER_BILL_HEAD.Fat_Rate_NMG,TSPL_BMC_TRANSPORTER_BILL_HEAD.Snf_Rate_NMG
FROM TSPL_BMC_TRANSPORTER_BILL_DETAIL 
LEFT OUTER JOIN TSPL_BMC_TRANSPORTER_BILL_HEAD ON TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code=TSPL_BMC_TRANSPORTER_BILL_DETAIL.Document_Code 
LEFT OUTER JOIN TSPL_MILK_COLLECTION_MCC  on TSPL_MILK_COLLECTION_MCC.Document_No=TSPL_BMC_TRANSPORTER_BILL_DETAIL.MCC_Document_Code
left outer join TSPL_TANKER_MASTER ON TSPL_TANKER_MASTER.Tanker_No=TSPL_BMC_TRANSPORTER_BILL_HEAD.Tanker_No
--left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_TANKER_MASTER.Tanker_Transporter_Code
 left outer join TSPL_BULK_ROUTE_MASTER ON TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_COLLECTION_MCC.Route_Code
 left outer join TSPL_COMPANY_MASTER on 2=2
where TSPL_BMC_TRANSPORTER_BILL_HEAD.Document_Code='" & txtDocNo.Value & "'  "

            Dim dt As DataTable = clsDBFuncationality.GetDataTable(FinalQuery)
            If dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBmcTranspoterBill_1", "BMC Transpoter Bill") ''report for both (RCDF And RCDFCF)

                frmCRV = Nothing
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnExport_Click_1(sender As Object, e As EventArgs) Handles btnExport.Click
        '=======Update By preeti Gupta Against Ticket No[BM00000008831]
        Dim sql As String = Nothing
        sql = "select '' as Date,'BMC' AS Category,'' AS [Station 1],'' AS [Station 2],'' AS [Station 3],''	AS [Station 4],0 AS [Trip],'' as [Ice Box],0 AS [GPS KM],0 AS [KM],0 AS [Quantity KG],0 AS [Amount]	,0 AS [Diesel RD]  "
        ListImpExpColumnsMandatory = New List(Of String)({"Document_Code"})
        ListImpExpColumnsSuperMandatory = New List(Of String)({"Document_Code"})
        transportSql.ExporttoExcel(sql, "", "", Me, ListImpExpColumnsMandatory, ListImpExpColumnsSuperMandatory, MyBase.Form_ID)
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        funImport()
    End Sub
    Public Sub funImport()
        Dim gvimport As New UserControls.MyRadGridView
        Me.Controls.Add(gvimport)
        Dim currentdate As Date = Date.Today
        LoadBlankGrid()
        gv1.Rows.AddNew()
        Dim qry1 As String = Nothing
        Dim dt As DataTable = Nothing
        If clsCommon.myLen(txtDocNo.Value) > 0 Then

            If transportSql.importExcel(gvimport, "Date", "Category", "Station 1", "Station 2", "Station 3", "Station 4", "Trip", "Ice Box", "GPS KM", "KM", "Quantity KG", "Amount", "Diesel RD") Then

                Try
                    clsCommon.ProgressBarPercentShow()
                    For ii As Integer = 0 To gvimport.Rows.Count - 1
                        clsCommon.ProgressBarPercentUpdate((gvimport.Rows(ii).Index + 1) * 100 / (gvimport.Rows.Count + 1), "Importing  : " & (gvimport.Rows(ii).Index + 1) & "/" & gvimport.Rows.Count & "")
                        Try

                            gv1.Rows(ii).Cells(colDate).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Date").Value)
                            gv1.Rows(ii).Cells(ColCategory).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Category").Value)
                            gv1.Rows(ii).Cells(ColStation).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 1").Value)
                            gv1.Rows(ii).Cells(ColStation2).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 2").Value)
                            gv1.Rows(ii).Cells(ColStation3).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 3").Value)
                            gv1.Rows(ii).Cells(ColStation4).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 4").Value)
                            gv1.Rows(ii).Cells(ColTrip).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Trip").Value)
                            gv1.Rows(ii).Cells(ColIceBox).Value = (clsCommon.CompairString(clsCommon.myCstr(gvimport.Rows(ii).Cells("Ice Box").Value), "Y") = CompairStringResult.Equal)
                            gv1.Rows(ii).Cells(ColGPSKM).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("GPS KM").Value)
                            gv1.Rows(ii).Cells(ColKM).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("KM").Value)
                            gv1.Rows(ii).Cells(ColQuantity).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Quantity KG").Value)
                            gv1.Rows(ii).Cells(ColDiesel).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Diesel RD").Value)
                            gv1.Rows(ii).Cells(ColAmount).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Amount").Value)
                            If gvimport.Rows.Count > gv1.Rows.Count Then
                                gv1.Rows.AddNew()
                            End If
                        Catch ex As Exception
                            gv1.Rows.RemoveAt(ii)
                            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                        End Try
                    Next

                    clsCommon.ProgressBarPercentHide()
                    common.clsCommon.MyMessageBoxShow(Me, "Data imported successfully", Me.Text, MessageBoxButtons.OK)
                Catch ex As Exception
                    clsCommon.ProgressBarPercentHide()
                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                End Try
            End If
            Me.Controls.Remove(gvimport)
        Else
            If clsCommon.myLen(txtFromDate.Value) = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select a 'From Date'.", Me.Text)
                Exit Sub
            End If
            If clsCommon.myLen(txtToDate.Value) = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select a 'To Date'.", Me.Text)
                Exit Sub
            End If
            ' Validate Tanker Number
            If clsCommon.myLen(txtTankerNo.Value) = 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please enter a 'Tanker Number'.", Me.Text)
                Exit Sub
            End If

            ' Optional: Validate that From Date is not greater than To Date
            If CDate(txtFromDate.Value) > CDate(txtToDate.Value) Then
                clsCommon.MyMessageBoxShow(Me, "'From Date' cannot be greater than 'To Date'.", Me.Text)
                Exit Sub
            End If
            If transportSql.importExcel(gvimport, "Date", "Category", "Station 1", "Station 2", "Station 3", "Station 4", "Trip", "Ice Box", "GPS KM", "KM", "Quantity KG", "Amount", "Diesel RD") Then

                Try
                    clsCommon.ProgressBarPercentShow()

                    For ii As Integer = 0 To gvimport.Rows.Count - 1
                        clsCommon.ProgressBarPercentUpdate((gvimport.Rows(ii).Index + 1) * 100 / (gvimport.Rows.Count + 1), "Importing  : " & (gvimport.Rows(ii).Index + 1) & "/" & gvimport.Rows.Count & "")
                        Try
                            gv1.Rows(ii).Cells(colDate).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Date").Value)
                            gv1.Rows(ii).Cells(ColCategory).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Category").Value)
                            gv1.Rows(ii).Cells(ColStation).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 1").Value)
                            gv1.Rows(ii).Cells(ColStation2).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 2").Value)
                            gv1.Rows(ii).Cells(ColStation3).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 3").Value)
                            gv1.Rows(ii).Cells(ColStation4).Value = clsCommon.myCstr(gvimport.Rows(ii).Cells("Station 4").Value)
                            gv1.Rows(ii).Cells(ColTrip).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Trip").Value)
                            gv1.Rows(ii).Cells(ColGPSKM).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("GPS KM").Value)
                            gv1.Rows(ii).Cells(ColIceBox).Value = (clsCommon.CompairString(clsCommon.myCstr(gvimport.Rows(ii).Cells("Ice Box").Value), "Y") = CompairStringResult.Equal)
                            gv1.Rows(ii).Cells(ColKM).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("KM").Value)
                            gv1.Rows(ii).Cells(ColQuantity).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Quantity KG").Value)
                            gv1.Rows(ii).Cells(ColDiesel).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Diesel RD").Value)
                            gv1.Rows(ii).Cells(ColAmount).Value = clsCommon.myCdbl(gvimport.Rows(ii).Cells("Amount").Value)
                            If clsCommon.myLen(txtDocNo.Value) = 0 Then
                                If gv1.Rows.Count = gvimport.Rows.Count Then
                                Else
                                    gv1.Rows.AddNew()

                                End If
                            End If
                        Catch ex As Exception
                            gv1.Rows.RemoveAt(ii)
                            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                        End Try
                    Next
                    common.clsCommon.MyMessageBoxShow(Me, "Data imported successfully", Me.Text, MessageBoxButtons.OK)
                    clsCommon.ProgressBarPercentHide()

                Catch ex As Exception
                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                End Try
            End If
            Me.Controls.Remove(gvimport)
        End If

    End Sub

    Private Sub chkPrivate_CheckStateChanged(sender As Object, e As EventArgs) Handles chkPrivate.CheckStateChanged
        Try
            If chkPrivate.Checked Then
                btnGo.Enabled = False
                RadGroupBox3.Enabled = True
                LoadBlankGrid()
                gv1.Rows.AddNew()
                txtFatShortage.ReadOnly = False
                TxtSnfShortage.ReadOnly = False
            Else
                btnGo.Enabled = True
                txtFatShortage.ReadOnly = True
                TxtSnfShortage.ReadOnly = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub gv1_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gv1.CurrentColumnChanged
        If gv1.RowCount > 0 Then
            Dim intCurrRow As Integer = gv1.CurrentRow.Index
            If intCurrRow = gv1.Rows.Count - 1 Then
                If chkPrivate.Checked Then
                    gv1.Rows.AddNew()
                    gv1.CurrentRow = gv1.Rows(intCurrRow)
                End If
            End If
        End If
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        CloseForm()
    End Sub

    Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub

    Sub updateAll()
        For ii As Integer = 0 To gv1.Rows.Count - 1
            UpdateCurrentRow(ii, Nothing)
        Next
        UpdateAllTotals(Nothing)
    End Sub
    Private Sub txtDieselplus_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtDieselplus.Validating
        updateAll()
    End Sub

    Private Sub txtFatShortage_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtFatShortage.Validating
        updateAll()
    End Sub

    Private Sub TxtSnfShortage_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtSnfShortage.Validating
        updateAll()
    End Sub

    Private Sub TxtIceCharge_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtIceCharge.Validating
        updateAll()
    End Sub

    Private Sub TxtKMRate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtKMRate.Validating
        updateAll()
    End Sub

    Private Sub TxtDieselMinus_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtDieselMinus.Validating

    End Sub
End Class