'' work done agaist ticket no. BHA/01/11/18-000659
Imports common
Imports System.Data.SqlClient
Imports System

Public Class FrmARInvoiceEntry
    Inherits FrmMainTranScreen

#Region "Variables"
    Public strAPInvoice As String = Nothing
    Private blnInvoice As Boolean = False
    Private isCellValueChangedOpen As Boolean = False
    Private isCellValueChangedTaxOpen As Boolean = False
    Dim iStxtTaxGroup_TxtChangedComplete As Boolean = True
    Private isNewEntry As Boolean = False
    Private isInsideLoadData As Boolean = False
    Const colLineNo As String = "LNO"
    Const colACCode As String = "NAME"
    Const colACName As String = "QTY"
    Const colAddChgCode As String = "COLACCODE"
    Const colAddChgName As String = "COLACNAME"
    Const colSACCode As String = "COLSACCODE"
    Const colSACName As String = "COLSACNAME"
    Const colGLType As String = "colGLType"
    Const colAmt As String = "AMT"
    Dim GSTStatus As Boolean = False
    Const colDisPer As String = "DISPER"
    Const colDisAmt As String = "DISAMT"
    Const colAmtAfterDis As String = "AMTAFTERDIS"
    Const colTax1 As String = "COLTAX1"
    Const colTaxRate1 As String = "COLTAXRATE1"
    Const colTaxAmt1 As String = "COLTAXAMT1"
    Const colTax2 As String = "COLTAX2"
    Const colTaxRate2 As String = "COLTAXRATE2"
    Const colTaxAmt2 As String = "COLTAXAMT2"
    Const colTax3 As String = "COLTAX3"
    Const colTaxRate3 As String = "COLTAXRATE3"
    Const colTaxAmt3 As String = "COLTAXAMT3"
    Const colTax4 As String = "COLTAX4"
    Const colTaxRate4 As String = "COLTAXRATE4"
    Const colTaxAmt4 As String = "COLTAXAMT4"
    Const colTax5 As String = "COLTAX5"
    Const colTaxRate5 As String = "COLTAXRATE5"
    Const colTaxAmt5 As String = "COLTAXAMT5"
    Const colTax6 As String = "COLTAX6"
    Const colTaxRate6 As String = "COLTAXRATE6"
    Const colTaxAmt6 As String = "COLTAXAMT6"
    Const colTax7 As String = "COLTAX7"
    Const colTaxRate7 As String = "COLTAXRATE7"
    Const colTaxAmt7 As String = "COLTAXAMT7"
    Const colTax8 As String = "COLTAX8"
    Const colTaxRate8 As String = "COLTAXRATE8"
    Const colTaxAmt8 As String = "COLTAXAMT8"
    Const colTax9 As String = "COLTAX9"
    Const colTaxRate9 As String = "COLTAXRATE9"
    Const colTaxAmt9 As String = "COLTAXAMT9"
    Const colTax10 As String = "COLTAX10"
    Const colTaxRate10 As String = "COLTAXRATE10"
    Const colTaxAmt10 As String = "COLTAXAMT10"
    Const colTotTaxAmt As String = "TAXAMT"
    Const colAmtAfterTax As String = "AMTAFTERTAX"
    Const colTTaxAutCode As String = "TAXAUTCODE"
    Const colTTaxAutName As String = "TAXAUTNAME"
    Const colTIsTaxable As String = "ISTAXABLE"
    Const colTTaxRate As String = "TAXRATE"
    Const colTIsSurTax As String = "ISSURTAX"
    Const colTSurTaxCode As String = "SURTAXCODE"
    Const colTBaseAmt As String = "TAXBASEAMT"
    Const colTTaxAmt As String = "TAXAMT"
    Const colAChgCode As String = "COLACCODE"
    Const colAChgName As String = "COLACNAME"
    Const colAChgAmount As String = "COLACAMOUNT"
    Const colTaxBaseAmt1 As String = "COLTAXBASEAMT1"
    Const colTaxBaseAmt2 As String = "COLTAXBASEAMT2"
    Const colTaxBaseAmt3 As String = "COLTAXBASEAMT3"
    Const colTaxBaseAmt4 As String = "COLTAXBASEAMT4"
    Const colTaxBaseAmt5 As String = "COLTAXBASEAMT5"
    Const colTaxBaseAmt6 As String = "COLTAXBASEAMT6"
    Const colTaxBaseAmt7 As String = "COLTAXBASEAMT7"
    Const colTaxBaseAmt8 As String = "COLTAXBASEAMT8"
    Const colTaxBaseAmt9 As String = "COLTAXBASEAMT9"
    Const colTaxBaseAmt10 As String = "COLTAXBASEAMT10"

    Const colHierarchyCode As String = "colHierarchyCode"
    Const colCostCenterCode As String = "colCostCenterCode"

    Const colHierarchyLevelNumber As String = "colHierarchyLevelNumber"
    Const colHierarchyLevel1 As String = "colHierarchyLevel1"
    Const colHierarchyLevel2 As String = "colHierarchyLevel2"
    Const colHierarchyLevel3 As String = "colHierarchyLevel3"
    Const colHierarchyLevel4 As String = "colHierarchyLevel4"

    Dim ButtonToolTip As ToolTip = New ToolTip()
    Public clicked As Boolean = False

    Dim SettingCostCenter As Boolean
    Dim SettingCostCenterlevel As Boolean
    Dim ERPStartDate As Date

    Dim repoAdChagCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
    Dim repoAdChagName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
    Dim FlagDocumentIsTaxable As Integer = 0
    Dim EInvoiceType As String = ""
    Dim CostCenterAndHirerachyCodeUpdateAfterPost As Boolean = False
#End Region

    Private Sub FrmAPInvoiceEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SettingCostCenter = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowHierarchyAndCostCenterInARInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInARInvoiceEntry, Nothing)) = 1)
        SettingCostCenterlevel = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableHirerachyCostCentre, clsFixedParameterCode.EnableHirerachyCostCentre, Nothing)) = 1)
        CostCenterAndHirerachyCodeUpdateAfterPost = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CostCenterAndHirerachyCodeUpdateAfterPost, clsFixedParameterCode.CostCenterAndHirerachyCodeUpdateAfterPost, Nothing)) = 1, True, False)
        Try
            ERPStartDate = clsCommon.myCDate(objCommonVar.ERPStartDate)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, "Invalid ERP Start Date", Me.Text)
            Me.Close()
        End Try
        If objCommonVar.ApplyLocationWisePrefix Then
            pnlLocation.Visible = True
        Else
            pnlLocation.Visible = False
        End If
        SetUserMgmtNew()
        SetMailRight()
        ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D Delete Trasnaction")
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnAddNew, "Press Alt+N Adding New Trasnaction")
        ButtonToolTip.SetToolTip(btnPrint, "Press Ctrl+P for Print ")
        LoadReturnType()
        RadPageView1.SelectedPage = RadPageViewPage1
        txtDocNo.MyReadOnly = True
        LoadBlankGridGL()
        LoadBlankGridTax()
        AddNew()
        loadSecurityDepositType()
        LoadInvoiceType()

        If Not objCommonVar.IsDemoERP Then
            btnDrillDown.Visible = False
        End If
        ValidateLength()
        ApplyReadOption()
        txtlocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code in (select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' ) "))
        '' Anubhooti 28-Nov-2014 BM00000004810
        If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
            LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
        Else
            LblLocDesp.Text = ""
        End If
        txtDueDate.Enabled = False
        ''

        If clsCommon.myLen(Me.Tag) > 0 Then
            LoadData(clsCommon.myCstr(Me.Tag))
        End If
        ''For Custom Fields
        RadPageView1.Pages("pvpCustomFields").Item.Visibility = MyBase.customFieldTabProperty
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.Report_ID = MyBase.Form_ID
            UcCustomFields1.LoadCustomControls()
        End If
        ''End of For Custom Fields

        '' MultiCurrency
        SetMultiCurrencyVisibility()
        '' End of MultiCurrency

        ''For Attachment
        If objCommonVar.IsDemoERP Then
            UcAttachment1.Form_ID = MyBase.Form_ID
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Visible
        Else
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Collapsed
        End If

        If chkEinvoice.Checked = True Then
            RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Visible
        Else
            RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed

        End If
        ''End of For Attachment
        If Not objCommonVar.IsDemoERP Then
            pnlPCJ.Visible = False
            fndProject.Visible = False
            lblProject.Visible = False
        End If
        '' Anubhooti 05-Dec-2014 (InvoiceNo. finder should visible in case of D/C)
        If ((clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal)) Then
            txtRefDocNo.Visible = True
            RadLabel15.Visible = True
            txtRefDocNo.Value = ""
        Else
            txtRefDocNo.Visible = False
            RadLabel15.Visible = False
            txtRefDocNo.Value = ""
        End If
        If clsCommon.myLen(strAPInvoice) > 0 Then
            LoadData(strAPInvoice)
        End If
    End Sub

    Private Sub SetUserMgmtNew()
        MyBase.SetUserMgmt(clsUserMgtCode.mbtnARInvoiceEntry)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")

        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnCancel.Visible = MyBase.isCancel_Flag_After_Posting
        'RadMenu1.Visible = MyBase.isExport
        btnReverse.Visible = False
        btnAdminCancel.Visible = False
        'If MyBase.isReverse Then
        '    btnReverse.Enabled = True
        'Else
        '    btnReverse.Enabled = False
        'End If
        If MyBase.isExport = True Then
            RadMenuItem2.Enabled = True
            RadMenuItem3.Enabled = True
        Else
            RadMenuItem2.Enabled = False
            RadMenuItem3.Enabled = False
        End If
    End Sub

    Private Sub fndProject__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndProject._MYValidating
        Dim qry As String = "select PROJECT_CODE as Code,SPECIFICATION,PROJECT_STATUS as Status from TSPL_PJC_PROJECT"
        fndProject.Value = clsCommon.ShowSelectForm("Project Code", qry, "Code", "", fndProject.Value, "", isButtonClicked)
        lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")
    End Sub


    Sub SetMultiCurrencyVisibility()
        Dim strq As String = ""
        '=======shivani
        Dim Currency As Integer = clsModuleCurrencyMapping.GetmulticurrencyDecimalPlaces()
        txtConversionRate.DecimalPlaces = clsCommon.myCdbl(Currency)
        '================
        If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
            pnlCurrConv.Visible = True

            If clsCommon.myLen(Me.TxtCustomer.Value) > 0 Then
                strq = "select currency_code from TSPL_CUSTOMER_MASTER where cust_code='" & clsCommon.myCstr(Me.TxtCustomer.Value) & "'"
                Me.txtCurrencyCode.Value = clsDBFuncationality.getSingleValue(strq).ToString
                ShowCurrencyDetail()
            End If
            ShowCurrencyDetail()
        Else
            pnlCurrConv.Visible = False
        End If

    End Sub

    Sub ShowCurrencyDetail()
        ' Dim strq As String
        Dim dt As DataTable
        If clsCommon.myLen(clsCommon.myCstr(Me.TxtCustomer.Value)) = 0 Then
            Exit Sub
        End If

        If clsCommon.myLen(txtCurrencyCode.Value) > 0 Then
            dt = clsModuleCurrencyMapping.GetLatestCurConvRateDT(Me.txtDate.Value, txtCurrencyCode.Value)
            If dt.Rows.Count = 0 Then
                If Me.txtCurrencyCode.Value = objCommonVar.BaseCurrencyCode Then
                    Me.txtConversionRate.Text = 1
                    Me.txtApplicableFrom.Text = ""
                Else
                    clsCommon.MyMessageBoxShow("Conversion rate not entered for currency '" & Me.txtCurrencyCode.Value & "'")
                    Exit Sub
                End If
            Else
                Me.txtConversionRate.Text = dt.Rows(0).Item("Rate")
                Me.txtApplicableFrom.Text = clsCommon.GetPrintDate(dt.Rows(0).Item("FROM_DATE"), "dd/MMM/yyyy")
            End If
        Else
            Me.txtConversionRate.Text = 1
            Me.txtApplicableFrom.Text = ""
        End If

    End Sub

    Private Sub ValidateLength()
        txtDocNo.MyMaxLength = 30
        txtDesc.MaxLength = 250
        txtOrderNo.MaxLength = 30
    End Sub

    Private Sub ApplyReadOption()
        txtDocNo.MyReadOnly = False
    End Sub

    Private Sub txtChangeVendorNo()
        If Not isInsideLoadData Then
            Dim qry As String = "select  TSPL_CUSTOMER_MASTER.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name ,TSPL_CUSTOMER_MASTER.Terms_Code,TSPL_TERMS_MASTER.Terms_Desc ,TSPL_CUSTOMER_MASTER.Cust_Account ,TSPL_CUSTOMER_MASTER.Tax_Group,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc from TSPL_CUSTOMER_MASTER left outer join  TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code=TSPL_CUSTOMER_MASTER.Terms_Code left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_CUSTOMER_MASTER.Tax_Group and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' where Cust_Code ='" + TxtCustomer.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                lblVendorName.Text = clsCommon.myCstr(dt.Rows(0)("Customer_Name"))
                TxtCustomer.Value = clsCommon.myCstr(dt.Rows(0)("Cust_Code"))
                txtTermCode.Value = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
                lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Terms_Desc"))
                txtACSet.Value = clsCommon.myCstr(dt.Rows(0)("Cust_Account"))
                txtTaxGroup.Value = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
                lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
            Else
                TxtCustomer.Value = ""
                lblVendorName.Text = ""
                txtTermCode.Value = ""
                lblTermName.Text = ""
                txtACSet.Value = ""
                txtTaxGroup.Value = ""
                lblTaxGrpName.Text = ""
            End If
            txtTaxGroup_TxtChanged()
        End If
    End Sub

    Private Sub txtTaxGroup_TxtChanged()
        If Not isInsideLoadData AndAlso iStxtTaxGroup_TxtChangedComplete Then
            iStxtTaxGroup_TxtChangedComplete = False
            LoadBlankGridTax()
            Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,Tax_Code,Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='S') AS TaxRate,Taxable from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + txtTaxGroup.Value + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='S' order by Trans_Code"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                If (dt.Rows.Count > 10) Then
                    MessageBox.Show("Can't Handle More than 10 Tax Types in a Group")
                    Return
                End If
                Dim ii As Integer = 0
                txtTaxGroup.Value = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Code"))
                lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
                For Each dr As DataRow In dt.Rows
                    gv2.Rows.AddNew()
                    gv2.Rows(ii).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dr("Tax_Code"))
                    gv2.Rows(ii).Cells(colTTaxAutName).Value = clsCommon.myCstr(dr("Tax_Code_Desc"))
                    If rbtnTaxCalAutomatic.IsChecked Then
                        gv2.Rows(ii).Cells(colTTaxRate).Value = clsCommon.myCdbl(dr("TaxRate"))
                        gv2.Rows(ii).Cells(colTIsSurTax).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        gv2.Rows(ii).Cells(colTIsTaxable).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        gv2.Rows(ii).Cells(colTSurTaxCode).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                    ElseIf rbtnTaxCalManual.IsChecked Then
                        gv2.Rows(ii).Cells(colTTaxRate).Value = Nothing
                        gv2.Rows(ii).Cells(colTIsSurTax).Value = Nothing
                        gv2.Rows(ii).Cells(colTIsTaxable).Value = Nothing
                        gv2.Rows(ii).Cells(colTSurTaxCode).Value = Nothing
                    End If
                    ii = ii + 1
                Next
            Else
                lblTaxGrpName.Text = ""
            End If

            For ii As Integer = 0 To gv1.Rows.Count - 1
                UpdateCurrentRowWihtRowNo(ii)
            Next
            UpdateAllTotals()
            iStxtTaxGroup_TxtChangedComplete = True
        End If
    End Sub

    Private Sub txtTermCode_TxtChanged()
        If Not isInsideLoadData Then
            Dim qry As String = "select Terms_Code,Terms_Desc,No_Days from TSPL_TERMS_MASTER where Terms_Code='" + clsCommon.myCstr(txtTermCode.Value) + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                txtTermCode.Value = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
                lblTermName.Text = clsCommon.myCstr(dt.Rows(0)("Terms_Desc"))
                txtDueDate.Value = txtDate.Value.AddDays(clsCommon.myCdbl(dt.Rows(0)("No_Days")))
            Else
                lblTermName.Text = ""
                txtDueDate.Value = clsCommon.GETSERVERDATE()
            End If
        End If
    End Sub

    Sub LoadInvoiceType()
        blnInvoice = True
        Dim dt As New DataTable
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        Dim dr As DataRow = dt.NewRow()
        dr("Code") = "I"
        dr("Name") = "Invoice"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "D"
        dr("Name") = "Debit Note"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "C"
        dr("Name") = "Credit Note"
        dt.Rows.Add(dr)

        cboDocType.DataSource = dt
        cboDocType.ValueMember = "Code"
        cboDocType.DisplayMember = "Name"

        txtRefDocNo.Visible = True
        RadLabel15.Visible = True
        blnInvoice = False
    End Sub

    Private Sub loadSecurityDepositType()
        ddlSecDepositType.DataSource = clsCustomerInvoiceHead.GetSecurityDepositType()
        ddlSecDepositType.DisplayMember = "Code"
        ddlSecDepositType.ValueMember = "Value"
        ddlSecDepositType.Enabled = False
        chkSecurityDposit.Enabled = False
    End Sub

    Sub BlankAllControls()
        EInvoiceType = ""
        FlagDocumentIsTaxable = 0
        chkEinvoice.Checked = False
        txtServiceVisitCode.Text = ""
        txtVCGNo.Text = ""
        txtSaleReturnNo.Text = ""
        txtSaleInvoice.Text = ""
        txtSaleInvoice.ReadOnly = True
        txtVCGNo.ReadOnly = True
        txtServiceVisitCode.ReadOnly = True
        txtSaleReturnNo.ReadOnly = True

        UsLock1.Status = ERPTransactionStatus.Pending
        txtDocNo.Value = ""
        txtDesc.Text = ""
        chkOnHold.Checked = False
        TxtCustomer.Value = ""
        lblVendorName.Text = ""
        txtACSet.Value = ""
        txtDate.Value = clsCommon.GETSERVERDATE()
        cboDocType.SelectedIndex = 0
        txtOrderNo.Text = ""
        txtTaxGroup.Value = ""
        lblTaxGrpName.Text = ""
        txtTermCode.Value = ""
        lblTermName.Text = ""
        txtDueDate.Value = txtDate.Value
        lblAmtWithDiscount.Text = ""
        lblDiscountAmt.Text = ""
        lblAmtAfterDiscount.Text = ""
        lblTaxAmt.Text = ""
        lblTotRAmt.Text = ""
        lblTotRAmt1.Text = ""
        lblTotalSubsidyAmt.Text = ""
        txtRefDocNo.Value = ""
        rbtnTaxCalAutomatic.IsChecked = True
        txtRefDocNo.Visible = False
        RadLabel15.Visible = False
        txtDate.Enabled = True
        chkAgainstServiceInvoice.Checked = False
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.BlankAllControls()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()
        lblProject.Text = ""

        fndProject.Value = ""
        fndProject.Enabled = True
        lblProject.Enabled = True

        chkAgainstSecurityReceiptNo.Checked = False
        txtSercurityReceiptNo.Value = ""
        lblRoundoffAmount.Text = ""
        txtDataAndTimeSelection.Value = clsCommon.GETSERVERDATE()
        txtTapalNo.Text = ""
        txtDataAndTimeSelection.Checked = False
    End Sub

    Private Sub TxtVendorNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles TxtCustomer._MYValidating
        Try
            If clsCommon.myLen(txtlocation.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please first select Location", Me.Text)
                txtlocation.Focus()
                Return
            End If
            Dim Qry As String = clsERPFuncationality.glCustomerQuery
            Dim strwherecls As String = "Status ='N' AND OnHold='N'"
            If objCommonVar.ApplyLocationFilterBasedOnPermission = True Then
                If clsCommon.myLen(objCommonVar.strCurrUserCustomers) > 0 Then
                    strwherecls += " and m.Cust_Code in (" + objCommonVar.strCurrUserCustomers + ") "
                End If
            End If
            TxtCustomer.Value = clsCommon.ShowSelectForm("CustomerCodeSelector", Qry, "Code", strwherecls, TxtCustomer.Value, "Code", isButtonClicked)
            txtChangeVendorNo()
            '' Anubhooti 28-Nov-2014 BM00000004810 (Due Date Cal. acc. to terms code)
            txtTermCode_TxtChanged()
            ''
            LoadSecurityAccount(TxtCustomer.Value)
            SetMultiCurrencyVisibility()

            ''richa 05 feb,2019  TEC/05/02/19-000411 not check for opening in case of Credit or debit note 
            'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, Nothing)
            Dim JEWithOPening As Boolean = False
            If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                If clsCommon.myCDate(clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                    JEWithOPening = True
                End If
            End If
            If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, Nothing)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "C") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "D") = CompairStringResult.Equal) And JEWithOPening = True Then
                LoadBlankGridGL()
                gv1.Rows.AddNew()
            End If

        Catch ex As Exception
            Me.txtCurrencyCode.Value = ""
            Me.txtConversionRate.Text = 1
            Me.txtApplicableFrom.Text = ""
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub LoadSecurityAccount(ByVal strCustCode As String)
        If Not isInsideLoadData Then
            Try
                gv1.Columns(colACCode).ReadOnly = False
                If clsCommon.myLen(strCustCode) > 0 Then
                    If chkSecurityDposit.Checked Then
                        gv1.Rows.Clear()
                        gv1.Rows.AddNew()
                        isInsideLoadData = True
                        If clsCommon.myLen(ddlSecDepositType.SelectedValue) > 0 Then
                            Dim dt As DataTable = clsDBFuncationality.GetDataTable("Select XXX.Account_Code, TSPL_GL_ACCOUNTS.Description as Account_desc from (" &
                    " Select Case When '" & ddlSecDepositType.SelectedValue & "'='S' then SECURITY_ACCOUNT When '" + ddlSecDepositType.SelectedValue + "'='C' Then CREATE_SECURITY_ACCOUNT When '" + ddlSecDepositType.SelectedValue + "'='R' Then BANK_GUARANTEE When '" + ddlSecDepositType.SelectedValue + "'='O' Then ACCOUNT1 End as Account_Code from TSPL_CUSTOMER_ACCOUNT_SET Where Cust_Account=(Select Cust_Account from TSPL_CUSTOMER_MASTER WHERE Cust_Code='" & strCustCode & "')" &
                    " ) XXX LEFT OUTER JOIN TSPL_GL_ACCOUNTS ON TSPL_GL_ACCOUNTS.Account_Code=XXX.Account_Code")
                            If dt.Rows.Count > 0 Then

                                Dim strACWithLocation As String = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dt.Rows(0)("Account_Code")), txtlocation.Value, True, Nothing)
                                gv1.CurrentRow.Cells(colACCode).Value = strACWithLocation
                                Dim strAccountDis As String = clsDBFuncationality.getSingleValue("select TSPL_GL_ACCOUNTS.Description  from TSPL_GL_ACCOUNTS where Account_Code ='" + strACWithLocation + "'")
                                gv1.CurrentRow.Cells(colACName).Value = strAccountDis
                                gv1.Columns(colACCode).ReadOnly = True
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
            Finally
                isInsideLoadData = False
            End Try
        End If
    End Sub

    Private Sub TxtVendorNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtCustomer.KeyPress
        If (e.KeyChar = Chr(39)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtACSet__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtACSet._MYValidating
        Dim Qry As String = "select Cust_Account as Code ,Cust_Acct_Desc as Description from TSPL_CUSTOMER_ACCOUNT_SET"
        txtACSet.Value = clsCommon.ShowSelectForm("Account_Filter", Qry, "Code", "", txtACSet.Value, "Code", isButtonClicked)
    End Sub

    Private Sub txtACSet_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtACSet.KeyPress
        If (e.KeyChar = Chr(39)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtTaxGroup__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTaxGroup._MYValidating
        'Dim Qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
        'txtTaxGroup.Value = clsCommon.ShowSelectForm("TaxGroupSelectorSale", Qry, "Code", " Tax_Group_Type='S'", txtTaxGroup.Value, "Code", isButtonClicked)
        ' txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroup(txtlocation.Value, TxtCustomer.Value, "S", txtTaxGroup.Value, isButtonClicked)
        txtTaxGroup.Value = clsLocationWiseTax.FinderForTaxGroupLocationSegment(txtlocation.Value, TxtCustomer.Value, "S", txtTaxGroup.Value, isButtonClicked)
        txtTaxGroup_TxtChanged()
    End Sub

    Private Sub txtTaxGroup_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTaxGroup.KeyPress
        If (e.KeyChar = Chr(39)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtTermCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtTermCode._MYValidating
        Dim Qry As String = "select Terms_Code as Code,Terms_Desc as Description,No_Days as [No Of Days] from TSPL_TERMS_MASTER"
        txtTermCode.Value = clsCommon.ShowSelectForm("fmTermCode", Qry, "Code", "", txtTermCode.Value, "Code", isButtonClicked)
        txtTermCode_TxtChanged()
    End Sub

    Private Sub txtTermCode_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTermCode.KeyPress
        If (e.KeyChar = Chr(39)) Then
            e.Handled = True
        End If
    End Sub

    Sub LoadBlankGridGL()
        gv1.Rows.Clear()
        gv1.Columns.Clear()

        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLineNo)

        repoAdChagCode = New GridViewTextBoxColumn()
        repoAdChagCode.FormatString = ""
        repoAdChagCode.HeaderText = "Additional Charges"
        repoAdChagCode.Name = colAddChgCode
        repoAdChagCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoAdChagCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoAdChagCode.Width = 150
        repoAdChagCode.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAdChagCode)

        repoAdChagName = New GridViewTextBoxColumn()
        repoAdChagName.FormatString = ""
        repoAdChagName.HeaderText = "Additional Charges Description"
        repoAdChagName.Name = colAddChgName
        repoAdChagName.Width = 150
        repoAdChagName.ReadOnly = True
        repoAdChagName.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAdChagName)

        repoAdChagCode = New GridViewTextBoxColumn()
        repoAdChagCode.FormatString = ""
        repoAdChagCode.HeaderText = "SAC Code"
        repoAdChagCode.Name = colSACCode
        repoAdChagCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoAdChagCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoAdChagCode.Width = 100
        repoAdChagCode.ReadOnly = True
        repoAdChagCode.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAdChagCode)

        repoAdChagName = New GridViewTextBoxColumn()
        repoAdChagName.FormatString = ""
        repoAdChagName.HeaderText = "SAC Description"
        repoAdChagName.Name = colSACName
        repoAdChagName.Width = 150
        repoAdChagName.ReadOnly = True
        repoAdChagName.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoAdChagName)


        Dim repoAcCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoAcCode.FormatString = ""
        repoAcCode.HeaderText = "GL Account"
        repoAcCode.Name = colACCode
        repoAcCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoAcCode.TextImageRelation = TextImageRelation.TextBeforeImage

        repoAcCode.Width = 150
        gv1.MasterTemplate.Columns.Add(repoAcCode)

        Dim repoACName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoACName.FormatString = ""
        repoACName.HeaderText = "Account Description"
        repoACName.Name = colACName
        repoACName.Width = 150
        repoACName.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoACName)

        Dim repoGLType As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoGLType.FormatString = ""
        repoGLType.HeaderText = "GL Type"
        repoGLType.Name = colGLType
        repoGLType.Width = 100
        repoGLType.ReadOnly = True
        repoGLType.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoGLType)

        Dim repoAmt As GridViewCalculatorColumn = New GridViewCalculatorColumn()
        'repoAmt = New GridViewCalculatorColumn()
        repoAmt.FormatString = ""
        repoAmt.HeaderText = "Amount"
        repoAmt.Name = colAmt
        repoAmt.Width = 80
        'repoAmt.Minimum = 0
        repoAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAmt)

        Dim repoDisPer As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisPer = New GridViewDecimalColumn()
        repoDisPer.FormatString = ""
        repoDisPer.HeaderText = "Discount %"
        repoDisPer.Minimum = 0
        repoDisPer.Maximum = 100
        repoDisPer.Name = colDisPer
        repoDisPer.Width = 100
        repoDisPer.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoDisPer)

        Dim repoDisAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisAmt = New GridViewDecimalColumn()
        repoDisAmt.FormatString = ""
        repoDisAmt.HeaderText = "Discount Amount"
        repoDisAmt.Name = colDisAmt
        repoDisAmt.Width = 100
        repoDisAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoDisAmt.VisibleInColumnChooser = False
        repoDisAmt.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoDisAmt)

        Dim repoAmtAfterDis As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterDis = New GridViewDecimalColumn()
        repoAmtAfterDis.FormatString = ""
        repoAmtAfterDis.HeaderText = "Amount After Discount"
        repoAmtAfterDis.Name = colAmtAfterDis
        repoAmtAfterDis.Width = 100
        repoAmtAfterDis.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterDis.VisibleInColumnChooser = False
        repoAmtAfterDis.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterDis)


        Dim repoTax1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax1.FormatString = ""
        repoTax1.HeaderText = "Tax 1"
        repoTax1.Name = colTax1
        repoTax1.ReadOnly = True
        repoTax1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax1)

        Dim repoTaxBaseAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt1.FormatString = ""
        repoTaxBaseAmt1.HeaderText = "Tax Base Amount 1"
        repoTaxBaseAmt1.Name = colTaxBaseAmt1
        repoTaxBaseAmt1.ReadOnly = True
        repoTaxBaseAmt1.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt1)

        Dim repoTaxRate1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate1 = New GridViewDecimalColumn()
        repoTaxRate1.FormatString = ""
        repoTaxRate1.HeaderText = "Tax Rate 1"
        repoTaxRate1.Name = colTaxRate1
        repoTaxRate1.IsVisible = False
        repoTaxRate1.ReadOnly = True
        repoTaxRate1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate1)

        Dim repoTaxAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt1 = New GridViewDecimalColumn()
        repoTaxAmt1.FormatString = ""
        repoTaxAmt1.HeaderText = "Tax Amt 1"
        repoTaxAmt1.Name = colTaxAmt1
        repoTaxAmt1.IsVisible = False
        repoTaxAmt1.ReadOnly = True
        repoTaxAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt1)

        Dim repoTax2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax2.FormatString = ""
        repoTax2.HeaderText = "Tax 2"
        repoTax2.Name = colTax2
        repoTax2.ReadOnly = True
        repoTax2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax2)

        Dim repoTaxBaseAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt2.FormatString = ""
        repoTaxBaseAmt2.HeaderText = "Tax Base Amount 2"
        repoTaxBaseAmt2.Name = colTaxBaseAmt2
        repoTaxBaseAmt2.ReadOnly = True
        repoTaxBaseAmt2.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt2)

        Dim repoTaxRate2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate2 = New GridViewDecimalColumn()
        repoTaxRate2.FormatString = ""
        repoTaxRate2.HeaderText = "Tax Rate 2"
        repoTaxRate2.Name = colTaxRate2
        repoTaxRate2.IsVisible = False
        repoTaxRate2.ReadOnly = True
        repoTaxRate2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate2)

        Dim repoTaxAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt2 = New GridViewDecimalColumn()
        repoTaxAmt2.FormatString = ""
        repoTaxAmt2.HeaderText = "Tax Amt 2"
        repoTaxAmt2.Name = colTaxAmt2
        repoTaxAmt2.IsVisible = False
        repoTaxAmt2.ReadOnly = True
        repoTaxAmt2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt2)

        Dim repoTax3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax3.FormatString = ""
        repoTax3.HeaderText = "Tax 3"
        repoTax3.Name = colTax3
        repoTax3.ReadOnly = True
        repoTax3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax3)

        Dim repoTaxBaseAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt3.FormatString = ""
        repoTaxBaseAmt3.HeaderText = "Tax Base Amount 3"
        repoTaxBaseAmt3.Name = colTaxBaseAmt3
        repoTaxBaseAmt3.ReadOnly = True
        repoTaxBaseAmt3.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt3)

        Dim repoTaxRate3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate3 = New GridViewDecimalColumn()
        repoTaxRate3.FormatString = ""
        repoTaxRate3.HeaderText = "Tax Rate 3"
        repoTaxRate3.Name = colTaxRate3
        repoTaxRate3.IsVisible = False
        repoTaxRate3.ReadOnly = True
        repoTaxRate3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate3)

        Dim repoTaxAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt3 = New GridViewDecimalColumn()
        repoTaxAmt3.FormatString = ""
        repoTaxAmt3.HeaderText = "Tax Amt 3"
        repoTaxAmt3.Name = colTaxAmt3
        repoTaxAmt3.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt3)

        Dim repoTax4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax4.FormatString = ""
        repoTax4.HeaderText = "Tax 4"
        repoTax4.Name = colTax4
        repoTax4.ReadOnly = True
        repoTax4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax4)

        Dim repoTaxBaseAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt4.FormatString = ""
        repoTaxBaseAmt4.HeaderText = "Tax Base Amount 4"
        repoTaxBaseAmt4.Name = colTaxBaseAmt4
        repoTaxBaseAmt4.ReadOnly = True
        repoTaxBaseAmt4.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt4)

        Dim repoTaxRate4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate4 = New GridViewDecimalColumn()
        repoTaxRate4.FormatString = ""
        repoTaxRate4.HeaderText = "Tax Rate 4"
        repoTaxRate4.Name = colTaxRate4
        repoTaxRate4.IsVisible = False
        repoTaxRate4.ReadOnly = True
        repoTaxRate4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate4)

        Dim repoTaxAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt4 = New GridViewDecimalColumn()
        repoTaxAmt4.FormatString = ""
        repoTaxAmt4.HeaderText = "Tax Amt 4"
        repoTaxAmt4.Name = colTaxAmt4
        repoTaxAmt4.IsVisible = False
        repoTaxAmt4.ReadOnly = True
        repoTaxAmt4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt4)

        Dim repoTax5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax5.FormatString = ""
        repoTax5.HeaderText = "Tax 5"
        repoTax5.Name = colTax5
        repoTax5.ReadOnly = True
        repoTax5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax5)


        Dim repoTaxBaseAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt5.FormatString = ""
        repoTaxBaseAmt5.HeaderText = "Tax Base Amount 5"
        repoTaxBaseAmt5.Name = colTaxBaseAmt5
        repoTaxBaseAmt5.ReadOnly = True
        repoTaxBaseAmt5.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt5)

        Dim repoTaxRate5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate5 = New GridViewDecimalColumn()
        repoTaxRate5.FormatString = ""
        repoTaxRate5.HeaderText = "Tax Rate 5"
        repoTaxRate5.Name = colTaxRate5
        repoTaxRate5.IsVisible = False
        repoTaxRate5.ReadOnly = True
        repoTaxRate5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate5)

        Dim repoTaxAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt5 = New GridViewDecimalColumn()
        repoTaxAmt5.FormatString = ""
        repoTaxAmt5.HeaderText = "Tax Amt 5"
        repoTaxAmt5.Name = colTaxAmt5
        repoTaxAmt5.IsVisible = False
        repoTaxAmt5.ReadOnly = True
        repoTaxAmt5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt5)


        Dim repoTax6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax6.FormatString = ""
        repoTax6.HeaderText = "Tax 6"
        repoTax6.Name = colTax6
        repoTax6.ReadOnly = True
        repoTax6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax6)


        Dim repoTaxBaseAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt6.FormatString = ""
        repoTaxBaseAmt6.HeaderText = "Tax Base Amount 6"
        repoTaxBaseAmt6.Name = colTaxBaseAmt6
        repoTaxBaseAmt6.ReadOnly = True
        repoTaxBaseAmt6.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt6)


        Dim repoTaxRate6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate6 = New GridViewDecimalColumn()
        repoTaxRate6.FormatString = ""
        repoTaxRate6.HeaderText = "Tax Rate 6"
        repoTaxRate6.Name = colTaxRate6
        repoTaxRate6.IsVisible = False
        repoTaxRate6.ReadOnly = True
        repoTaxRate6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate6)

        Dim repoTaxAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt6 = New GridViewDecimalColumn()
        repoTaxAmt6.FormatString = ""
        repoTaxAmt6.HeaderText = "Tax Amt 6"
        repoTaxAmt6.Name = colTaxAmt6
        repoTaxAmt6.IsVisible = False
        repoTaxAmt6.ReadOnly = True
        repoTaxAmt6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt6)

        Dim repoTax7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax7.FormatString = ""
        repoTax7.HeaderText = "Tax 7"
        repoTax7.Name = colTax7
        repoTax7.ReadOnly = True
        repoTax7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax7)


        Dim repoTaxBaseAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt7.FormatString = ""
        repoTaxBaseAmt7.HeaderText = "Tax Base Amount 7"
        repoTaxBaseAmt7.Name = colTaxBaseAmt7
        repoTaxBaseAmt7.ReadOnly = True
        repoTaxBaseAmt7.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt7)


        Dim repoTaxRate7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate7 = New GridViewDecimalColumn()
        repoTaxRate7.FormatString = ""
        repoTaxRate7.HeaderText = "Tax Rate 7"
        repoTaxRate7.Name = colTaxRate7
        repoTaxRate7.IsVisible = False
        repoTaxRate7.ReadOnly = True
        repoTaxRate7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate7)

        Dim repoTaxAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt7 = New GridViewDecimalColumn()
        repoTaxAmt7.FormatString = ""
        repoTaxAmt7.HeaderText = "Tax Amt 7"
        repoTaxAmt7.Name = colTaxAmt7
        repoTaxAmt7.IsVisible = False
        repoTaxAmt7.ReadOnly = True
        repoTaxAmt7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt7)

        Dim repoTax8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax8.FormatString = ""
        repoTax8.HeaderText = "Tax 8"
        repoTax8.Name = colTax8
        repoTax8.ReadOnly = True
        repoTax8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax8)

        Dim repoTaxBaseAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt8.FormatString = ""
        repoTaxBaseAmt8.HeaderText = "Tax Base Amount 8"
        repoTaxBaseAmt8.Name = colTaxBaseAmt8
        repoTaxBaseAmt8.ReadOnly = True
        repoTaxBaseAmt8.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt8)


        Dim repoTaxRate8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate8 = New GridViewDecimalColumn()
        repoTaxRate8.FormatString = ""
        repoTaxRate8.HeaderText = "Tax Rate 8"
        repoTaxRate8.Name = colTaxRate8
        repoTaxRate8.IsVisible = False
        repoTaxRate8.ReadOnly = True
        repoTaxRate8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate8)

        Dim repoTaxAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt8 = New GridViewDecimalColumn()
        repoTaxAmt8.FormatString = ""
        repoTaxAmt8.HeaderText = "Tax Amt 8"
        repoTaxAmt8.Name = colTaxAmt8
        repoTaxAmt8.IsVisible = False
        repoTaxAmt8.ReadOnly = True
        repoTaxAmt8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt8)

        Dim repoTax9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax9.FormatString = ""
        repoTax9.HeaderText = "Tax 9"
        repoTax9.Name = colTax9
        repoTax9.ReadOnly = True
        repoTax9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax9)

        Dim repoTaxBaseAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt9.FormatString = ""
        repoTaxBaseAmt9.HeaderText = "Tax Base Amount 9"
        repoTaxBaseAmt9.Name = colTaxBaseAmt9
        repoTaxBaseAmt9.ReadOnly = True
        repoTaxBaseAmt9.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt9)

        Dim repoTaxRate9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate9 = New GridViewDecimalColumn()
        repoTaxRate9.FormatString = ""
        repoTaxRate9.HeaderText = "Tax Rate 9"
        repoTaxRate9.Name = colTaxRate9
        repoTaxRate9.IsVisible = False
        repoTaxRate9.ReadOnly = True
        repoTaxRate9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate9)

        Dim repoTaxAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt9 = New GridViewDecimalColumn()
        repoTaxAmt9.FormatString = ""
        repoTaxAmt9.HeaderText = "Tax Amt 9"
        repoTaxAmt9.Name = colTaxAmt9
        repoTaxAmt9.IsVisible = False
        repoTaxAmt9.ReadOnly = True
        repoTaxAmt9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt9)


        Dim repoTax10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax10.FormatString = ""
        repoTax10.HeaderText = "Tax 10"
        repoTax10.Name = colTax10
        repoTax10.ReadOnly = True
        repoTax10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTax10)

        Dim repoTaxBaseAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt10.FormatString = ""
        repoTaxBaseAmt10.HeaderText = "Tax Base Amount 10"
        repoTaxBaseAmt10.Name = colTaxBaseAmt10
        repoTaxBaseAmt10.ReadOnly = True
        repoTaxBaseAmt10.IsVisible = False
        gv1.MasterTemplate.Columns.Add(repoTaxBaseAmt10)

        Dim repoTaxRate10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate10 = New GridViewDecimalColumn()
        repoTaxRate10.FormatString = ""
        repoTaxRate10.HeaderText = "Tax Rate 10"
        repoTaxRate10.Name = colTaxRate10
        repoTaxRate10.IsVisible = False
        repoTaxRate10.ReadOnly = True
        repoTaxRate10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxRate10)

        Dim repoTaxAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt10 = New GridViewDecimalColumn()
        repoTaxAmt10.FormatString = ""
        repoTaxAmt10.HeaderText = "Tax Amt 10"
        repoTaxAmt10.Name = colTaxAmt10
        repoTaxAmt10.IsVisible = False
        repoTaxAmt10.ReadOnly = True
        repoTaxAmt10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTaxAmt10)


        Dim repoTotTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotTaxAmt = New GridViewDecimalColumn()
        repoTotTaxAmt.FormatString = ""
        repoTotTaxAmt.HeaderText = "Tax Amount"
        repoTotTaxAmt.Name = colTotTaxAmt
        repoTotTaxAmt.Width = 100
        repoTotTaxAmt.ReadOnly = True
        repoTotTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTotTaxAmt)

        Dim repoAmtAfterTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmtAfterTax = New GridViewDecimalColumn()
        repoAmtAfterTax.FormatString = ""
        repoAmtAfterTax.HeaderText = "Included Tax Amount"
        repoAmtAfterTax.Name = colAmtAfterTax
        repoAmtAfterTax.Width = 120
        repoAmtAfterTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterTax.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterTax)

        ' '==========added by Richa Agarwal

        Dim repoHierarchyCode3 As GridViewTextBoxColumn

        repoHierarchyCode3 = New GridViewTextBoxColumn()
        repoHierarchyCode3.FormatString = ""
        repoHierarchyCode3.HeaderText = "Hierarchy Level"
        repoHierarchyCode3.Name = colHierarchyCode
        repoHierarchyCode3.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoHierarchyCode3.TextImageRelation = TextImageRelation.TextBeforeImage
        repoHierarchyCode3.Width = 150
        repoHierarchyCode3.IsVisible = (SettingCostCenter Or SettingCostCenterlevel)
        gv1.MasterTemplate.Columns.Add(repoHierarchyCode3)

        repoHierarchyCode3 = New GridViewTextBoxColumn()
        repoHierarchyCode3.FormatString = ""
        repoHierarchyCode3.HeaderText = "Cost Center"
        repoHierarchyCode3.Name = colCostCenterCode
        repoHierarchyCode3.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoHierarchyCode3.TextImageRelation = TextImageRelation.TextBeforeImage
        repoHierarchyCode3.Width = 150
        repoHierarchyCode3.IsVisible = IIf(SettingCostCenterlevel, False, IIf(SettingCostCenter, True, False))
        gv1.MasterTemplate.Columns.Add(repoHierarchyCode3)


        repoAmtAfterTax = New GridViewDecimalColumn()
        repoAmtAfterTax.FormatString = ""
        repoAmtAfterTax.HeaderText = "Hierarchy Level Number"
        repoAmtAfterTax.Name = colHierarchyLevelNumber
        repoAmtAfterTax.IsVisible = SettingCostCenterlevel
        repoAmtAfterTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmtAfterTax.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoAmtAfterTax)

        repoHierarchyCode3 = New GridViewTextBoxColumn()
        repoHierarchyCode3.FormatString = ""
        repoHierarchyCode3.HeaderText = "Hierarchy Level 4"
        repoHierarchyCode3.Name = colHierarchyLevel4
        repoHierarchyCode3.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoHierarchyCode3.TextImageRelation = TextImageRelation.TextBeforeImage
        repoHierarchyCode3.Width = 150
        repoHierarchyCode3.IsVisible = SettingCostCenterlevel
        repoHierarchyCode3.Tag = 4
        gv1.MasterTemplate.Columns.Add(repoHierarchyCode3)

        repoHierarchyCode3 = New GridViewTextBoxColumn()
        repoHierarchyCode3.FormatString = ""
        repoHierarchyCode3.HeaderText = "Hierarchy Level 3"
        repoHierarchyCode3.Name = colHierarchyLevel3
        repoHierarchyCode3.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoHierarchyCode3.TextImageRelation = TextImageRelation.TextBeforeImage
        repoHierarchyCode3.Width = 150
        repoHierarchyCode3.IsVisible = SettingCostCenterlevel
        repoHierarchyCode3.Tag = 3
        gv1.MasterTemplate.Columns.Add(repoHierarchyCode3)

        repoHierarchyCode3 = New GridViewTextBoxColumn()
        repoHierarchyCode3.FormatString = ""
        repoHierarchyCode3.HeaderText = "Hierarchy Level 2"
        repoHierarchyCode3.Name = colHierarchyLevel2
        repoHierarchyCode3.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoHierarchyCode3.TextImageRelation = TextImageRelation.TextBeforeImage
        repoHierarchyCode3.Width = 150
        repoHierarchyCode3.IsVisible = SettingCostCenterlevel
        repoHierarchyCode3.Tag = 2
        gv1.MasterTemplate.Columns.Add(repoHierarchyCode3)

        repoHierarchyCode3 = New GridViewTextBoxColumn()
        repoHierarchyCode3.FormatString = ""
        repoHierarchyCode3.HeaderText = "Hierarchy Level 1"
        repoHierarchyCode3.Name = colHierarchyLevel1
        repoHierarchyCode3.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoHierarchyCode3.TextImageRelation = TextImageRelation.TextBeforeImage
        repoHierarchyCode3.Width = 150
        repoHierarchyCode3.IsVisible = SettingCostCenterlevel
        repoHierarchyCode3.Tag = 1
        gv1.MasterTemplate.Columns.Add(repoHierarchyCode3)
        RefreshHeaderText()

        clsCustomFieldGrid.LoadBlankGrid(gv1, MyBase.ArrDetailFields)

        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
    End Sub

    Sub RefreshHeaderText()
        If SettingCostCenterlevel Then
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Description from TSPL_HIRERACHY_LEVEL_MASTER  order by level")
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If dt.Rows.Count > 0 Then
                    gv1.Columns(colHierarchyLevel1).HeaderText = clsCommon.myCstr(dt.Rows(0)("Description"))
                End If
                If dt.Rows.Count > 1 Then
                    gv1.Columns(colHierarchyLevel2).HeaderText = clsCommon.myCstr(dt.Rows(1)("Description"))
                End If
                If dt.Rows.Count > 2 Then
                    gv1.Columns(colHierarchyLevel3).HeaderText = clsCommon.myCstr(dt.Rows(2)("Description"))
                End If
                If dt.Rows.Count > 3 Then
                    gv1.Columns(colHierarchyLevel4).HeaderText = clsCommon.myCstr(dt.Rows(3)("Description"))
                End If
            End If
        End If
    End Sub

    Public Shared Function GetReturnType() As DataTable
        Dim dt As DataTable = New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))

        Dim dr As DataRow = dt.NewRow()
        dr("Code") = ""
        dr("Name") = "Select"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "D"
        dr("Name") = "Damaged Goods"
        dt.Rows.Add(dr)


        dr = dt.NewRow()
        dr("Code") = "P"
        dr("Name") = "Price Only"
        dt.Rows.Add(dr)


        dr = dt.NewRow()
        dr("Code") = "I"
        dr("Name") = "Inventory Type"
        dt.Rows.Add(dr)

        Return dt
    End Function

    Sub LoadReturnType()
        ddlReturnType.DataSource = GetReturnType()
        ddlReturnType.ValueMember = "Code"
        ddlReturnType.DisplayMember = "Name"
        ddlReturnType.SelectedValue = "I"
    End Sub

    Sub LoadBlankGridTax()
        gv2.Rows.Clear()
        gv2.Columns.Clear()

        Dim repoTaxAuthCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthCode.FormatString = ""
        repoTaxAuthCode.HeaderText = "Tax Authority Code"
        repoTaxAuthCode.Name = colTTaxAutCode
        repoTaxAuthCode.Width = 150
        repoTaxAuthCode.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthCode)

        Dim repoTaxAuthName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthName.FormatString = ""
        repoTaxAuthName.HeaderText = "Tax Authority"
        repoTaxAuthName.Name = colTTaxAutName
        repoTaxAuthName.Width = 200
        repoTaxAuthName.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthName)

        Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate.FormatString = ""
        repoTaxRate.HeaderText = "Tax Rate"
        repoTaxRate.Name = colTTaxRate
        repoTaxRate.Width = 100
        repoTaxRate.ReadOnly = True
        repoTaxRate.IsVisible = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, Nothing)) = 1, True, False)
        repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxRate)

        Dim repoIsSurTax As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax.HeaderText = "Is Surtax"
        repoIsSurTax.Name = colTIsSurTax
        repoIsSurTax.Width = 80
        repoIsSurTax.ReadOnly = True
        repoIsSurTax.IsVisible = False
        repoIsSurTax.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv2.MasterTemplate.Columns.Add(repoIsSurTax)

        Dim repoSurTaxCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode.FormatString = ""
        repoSurTaxCode.HeaderText = "Surtax"
        repoSurTaxCode.Name = colTSurTaxCode
        repoSurTaxCode.Width = 100
        repoSurTaxCode.ReadOnly = True
        repoSurTaxCode.IsVisible = False
        gv2.MasterTemplate.Columns.Add(repoSurTaxCode)

        Dim repoIsTaxable As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable.HeaderText = "Is Taxable"
        repoIsTaxable.Name = colTIsTaxable
        repoIsTaxable.Width = 80
        repoIsTaxable.ReadOnly = True
        repoIsTaxable.IsVisible = False
        repoIsTaxable.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv2.MasterTemplate.Columns.Add(repoIsTaxable)

        Dim repoTaxBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt.FormatString = ""
        repoTaxBaseAmt.HeaderText = "Base Amount"
        repoTaxBaseAmt.Name = colTBaseAmt
        repoTaxBaseAmt.Width = 100
        repoTaxBaseAmt.ReadOnly = True
        repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxBaseAmt)


        Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt.FormatString = ""
        repoTaxAmt.HeaderText = "Tax Amount"
        repoTaxAmt.Name = colTTaxAmt
        repoTaxAmt.Minimum = 0
        repoTaxAmt.Width = 100
        repoTaxAmt.ReadOnly = False
        repoTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxAmt)


        gv2.AllowAddNewRow = False
        gv2.ShowGroupPanel = False
        gv2.AllowColumnReorder = False
        gv2.AllowRowReorder = False
        gv2.EnableSorting = False
        gv2.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv2.MasterTemplate.ShowRowHeaderColumn = False

    End Sub

    Private Sub gv1_CellDoubleClick(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellDoubleClick
        'If clsCommon.myLen(gv1.CurrentRow.Cells(colACName).Value) > 0 Then
        '    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.glAccount, gv1.CurrentRow.Cells(colACName).Value)
        'End If
        Try
            If e.Column Is gv1.Columns(colTotTaxAmt) AndAlso rbtnTaxCalAutomatic.IsChecked Then
                Dim frm As New FrmPOItemTaxDetails()
                frm.strLineNo = clsCommon.myCstr(gv1.CurrentRow.Cells(colLineNo).Value)
                frm.strItemCode = clsCommon.myCstr(gv1.CurrentRow.Cells(colACCode).Value)
                frm.strItemName = clsCommon.myCstr(gv1.CurrentRow.Cells(colACName).Value)
                frm.dblTotTax = clsCommon.myCdbl(gv1.CurrentRow.Cells(colTotTaxAmt).Value)
                frm.dblAmtAfterDis = clsCommon.myCdbl(gv1.CurrentRow.Cells(colAmtAfterDis).Value)
                frm.strTaxType = "P"
                frm.IslocationSegment = True
                If clsCommon.myLen(frm.strItemCode) > 0 Then
                    frm.ArrIn = New List(Of clsTempItemTaxDetails)
                    For ii As Integer = 1 To 10
                        Dim strii As String = clsCommon.myCstr(ii)
                        Dim obj As New clsTempItemTaxDetails()
                        obj.AuthorityCode = clsCommon.myCstr(gv1.CurrentRow.Cells("COLTAX" + strii).Value)
                        If clsCommon.myLen(obj.AuthorityCode) > 0 Then
                            obj.AuthorityName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Tax_Code_Desc from TSPL_TAX_MASTER WHERE Tax_Code='" + obj.AuthorityCode + "'"))
                            obj.Rate = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value)
                            obj.BaseAmt = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value)
                            obj.TaxAmt = clsCommon.myCdbl(gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value)
                            'obj.isSurTax = clsCommon.myCBool(gv1.CurrentRow.Cells("ISSURTAX" + strii).Value)
                            'obj.SurTax = clsCommon.myCstr(gv1.CurrentRow.Cells("SURTAXCODE" + strii).Value)
                            'obj.IsTaxable = clsCommon.myCBool(gv1.CurrentRow.Cells("ISTAXABLE" + strii).Value)
                            frm.ArrIn.Add(obj)
                        End If
                    Next

                    frm.ShowDialog()
                    If frm.ArrOut IsNot Nothing AndAlso frm.ArrOut.Count > 0 Then
                        BlankTaxDetailsCurrentRowWihtRowNo(gv1.CurrentRow.Index)
                        For ii As Integer = 0 To frm.ArrOut.Count - 1
                            Dim strii As String = clsCommon.myCstr(ii + 1)
                            gv1.CurrentRow.Cells("COLTAX" + strii).Value = frm.ArrOut(ii).AuthorityCode
                            gv1.CurrentRow.Cells("COLTAXRATE" + strii).Value = frm.ArrOut(ii).Rate
                            gv1.CurrentRow.Cells("COLTAXBASEAMT" + strii).Value = frm.ArrOut(ii).BaseAmt
                            gv1.CurrentRow.Cells("COLTAXAMT" + strii).Value = frm.ArrOut(ii).TaxAmt
                            'gv1.CurrentRow.Cells("ISSURTAX" + strii).Value = frm.ArrOut(ii).isSurTax
                            'gv1.CurrentRow.Cells("SURTAXCODE" + strii).Value = frm.ArrOut(ii).SurTax
                            'gv1.CurrentRow.Cells("ISTAXABLE" + strii).Value = frm.ArrOut(ii).IsTaxable
                        Next

                        gv1.CurrentRow.Cells(colTotTaxAmt).Value = frm.dblTotTax
                        gv1.CurrentRow.Cells(colAmtAfterTax).Value = frm.dblTotTax + frm.dblAmtAfterDis

                        UpdateAllTotals()
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub gv1_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column.FieldName.StartsWith("_CFLD_") Then
                        clsCustomFieldGrid.getFinderForCustomFieldGrid(gv1, e.Column.Name.ToString, MyBase.Form_ID)
                    End If
                    If (clsCommon.CompairString(e.Column.Name, colAmt) = CompairStringResult.Equal) OrElse (clsCommon.CompairString(e.Column.Name, colDisPer) = CompairStringResult.Equal) Then
                        If rbtnTaxCalAutomatic.IsChecked Then
                            UpdateCurrentRowWihtRowNo(gv1.CurrentRow.Index)
                        ElseIf rbtnTaxCalManual.IsChecked Then
                            For ii As Integer = 0 To gv1.Rows.Count - 1
                                UpdateCurrentRowWihtRowNo(ii)
                            Next
                        End If
                        UpdateAllTotals()
                    ElseIf (clsCommon.CompairString(e.Column.Name, colAddChgCode) = CompairStringResult.Equal) Then
                        OpenAdditionCharges(False)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colACCode) = CompairStringResult.Equal) Then
                        OpenGLAccount(False)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colHierarchyCode) = CompairStringResult.Equal) Then
                        OpenHierarchyCode(False)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colCostCenterCode) = CompairStringResult.Equal) Then
                        OpenCostCenterCode(False)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colHierarchyLevel4) = CompairStringResult.Equal) Then
                        OpenHirerchyALLlevel(4)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colHierarchyLevel3) = CompairStringResult.Equal) Then
                        OpenHirerchyALLlevel(3)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colHierarchyLevel2) = CompairStringResult.Equal) Then
                        OpenHirerchyALLlevel(2)
                    ElseIf (clsCommon.CompairString(e.Column.Name, colHierarchyLevel1) = CompairStringResult.Equal) Then
                        OpenHirerchyALLlevel(1)
                    End If
                    setGridFocus()
                    isCellValueChangedOpen = False
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    ''richa agarwal 11/10/2019 
    Private Sub OpenAdditionCharges(ByVal isButtonClick As Boolean)
        If txtlocation.Value = "" Then
            common.clsCommon.MyMessageBoxShow(Me, "Please first select Location", Me.Text)
            txtlocation.Focus()
            Return
        End If
        Try
            Dim obj As clsAdditionalCharge = clsAdditionalCharge.GetFinder(clsCommon.myCstr(gv1.CurrentRow.Cells(colAddChgCode).Value), isButtonClick, False, False, txtDate.Value)
            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Code) > 0 Then
                gv1.CurrentRow.Cells(colAddChgCode).Value = obj.Code
                gv1.CurrentRow.Cells(colAddChgName).Value = obj.desc
                gv1.CurrentRow.Cells(colSACCode).Value = obj.SACCode
                gv1.CurrentRow.Cells(colSACName).Value = obj.SAC_Description
                gv1.CurrentRow.Cells(colACCode).Value = obj.Account_Code

                gv1.CurrentRow.Cells(colACCode).Value = clsCommon.myCstr(clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Account_Code, clsCommon.myCstr(txtlocation.Value), True, Nothing))
                gv1.CurrentRow.Cells(colACName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_GL_ACCOUNTS where Account_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colACCode).Value) + "'"))
            Else
                gv1.CurrentRow.Cells(colAddChgCode).Value = ""
                gv1.CurrentRow.Cells(colAddChgName).Value = ""
                gv1.CurrentRow.Cells(colSACCode).Value = ""
                gv1.CurrentRow.Cells(colSACName).Value = ""
            End If
        Catch ex As Exception
            gv1.CurrentRow.Cells(colAddChgCode).Value = ""
            gv1.CurrentRow.Cells(colAddChgName).Value = ""
            gv1.CurrentRow.Cells(colACCode).Value = ""
            gv1.CurrentRow.Cells(colACName).Value = ""
            gv1.CurrentRow.Cells(colSACCode).Value = ""
            gv1.CurrentRow.Cells(colSACName).Value = ""
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Private Sub OpenHirerchyALLlevel(ByVal lvl As Integer)
        If SettingCostCenterlevel Then
            If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyCode).Value)) > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") <> CompairStringResult.Equal Then
                    Dim dr As DataRow = OpenHirerchylevelFinder(lvl)
                    If dr IsNot Nothing Then
                        Select Case lvl
                            Case 1
                                gv1.CurrentRow.Cells(colHierarchyLevel1).Value = clsCommon.myCstr(dr(0))
                            Case 2
                                gv1.CurrentRow.Cells(colHierarchyLevel2).Value = clsCommon.myCstr(dr(0))
                                gv1.CurrentRow.Cells(colHierarchyLevel1).Value = clsCommon.myCstr(dr(1))
                            Case 3
                                gv1.CurrentRow.Cells(colHierarchyLevel3).Value = clsCommon.myCstr(dr(0))
                                gv1.CurrentRow.Cells(colHierarchyLevel2).Value = clsCommon.myCstr(dr(1))
                                gv1.CurrentRow.Cells(colHierarchyLevel1).Value = clsCommon.myCstr(dr(2))
                            Case 4
                                gv1.CurrentRow.Cells(colHierarchyLevel4).Value = clsCommon.myCstr(dr(0))
                                gv1.CurrentRow.Cells(colHierarchyLevel3).Value = clsCommon.myCstr(dr(1))
                                gv1.CurrentRow.Cells(colHierarchyLevel2).Value = clsCommon.myCstr(dr(2))
                                gv1.CurrentRow.Cells(colHierarchyLevel1).Value = clsCommon.myCstr(dr(3))
                        End Select
                    End If
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Please select hirerachy level first.", Me.Text)
            End If
        End If
    End Sub

    Function OpenHirerchylevelFinder(ByVal lvl As Integer) As DataRow
        Dim qry As String = " select "
        Select Case lvl
            Case 1
                qry += " TSPL_COST_CENTRE_HIRERACHY_DETAIL.COST_CENTRE_HIRERACHY_CODE as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel1).HeaderText) + "],TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name as [Description] "
            Case 2
                qry += " COST_CENTRE_HIRERACHY_CODE as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel2).HeaderText) + "],TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL_CODE1 as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel1).HeaderText) + "],TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name as [Description] "
            Case 3
                qry += " COST_CENTRE_HIRERACHY_CODE as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel3).HeaderText) + "],TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL_CODE2 as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel2).HeaderText) + "],TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL_CODE1 as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel1).HeaderText) + "],TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name as [Description] "
            Case 4
                qry += " COST_CENTRE_HIRERACHY_CODE as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel4).HeaderText) + "],TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL_CODE3 as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel3).HeaderText) + "],TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL_CODE2 as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel2).HeaderText) + "],TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL_CODE1 as [" + clsCommon.myCstr(gv1.Columns(colHierarchyLevel1).HeaderText) + "],TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name as [Description] "
            Case Else
                Throw New Exception("Wrong level ")
        End Select
        qry += " from TSPL_COST_CENTRE_HIRERACHY_DETAIL left outer join TSPL_COST_CENTRE_FINANCIAL on TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code=TSPL_COST_CENTRE_HIRERACHY_DETAIL.COST_CENTRE_HIRERACHY_CODE where TSPL_COST_CENTRE_HIRERACHY_DETAIL.HIRERACHY_LEVEL=" + clsCommon.myCstr(lvl)

        Return clsCommon.ShowSelectFormForRow("DDAdde" + clsCommon.myCstr(lvl), qry)
    End Function

    Private Sub OpenHierarchyCode(ByVal isButtonClick As Boolean)
        If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") <> CompairStringResult.Equal Then
            Dim qry As String = " select TSPL_HIRERACHY_LEVEL_MASTER.HIRERACHY_CODE as [Code] ,TSPL_HIRERACHY_LEVEL_MASTER.Description as [Description],TSPL_HIRERACHY_LEVEL_MASTER.Level as [Level] ,TSPL_HIRERACHY_LEVEL_MASTER.Created_By as [Created By] ,CONVERT(VARCHAR,TSPL_HIRERACHY_LEVEL_MASTER.Created_Date,103) as [Created Date] ,TSPL_HIRERACHY_LEVEL_MASTER.Modified_By as [Modified By] ,CONVERT(VARCHAR,TSPL_HIRERACHY_LEVEL_MASTER.Modified_Date,103) as [Modified Date]  From TSPL_HIRERACHY_LEVEL_MASTER  "
            gv1.CurrentRow.Cells(colHierarchyCode).Value = clsCommon.ShowSelectForm("Hierarchy", qry, "Code", "", clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyCode).Value), "Code", isButtonClick)
            gv1.CurrentRow.Cells(colHierarchyLevelNumber).Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select ISNULL(Level,0) AS Level from TSPL_HIRERACHY_LEVEL_MASTER Where HIRERACHY_CODE='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyCode).Value) + "' "))

            gv1.CurrentRow.Cells(colCostCenterCode).Value = ""
            gv1.CurrentRow.Cells(colHierarchyLevel1).Value = ""
            gv1.CurrentRow.Cells(colHierarchyLevel2).Value = ""
            gv1.CurrentRow.Cells(colHierarchyLevel3).Value = ""
            gv1.CurrentRow.Cells(colHierarchyLevel4).Value = ""
        End If
    End Sub

    Private Sub OpenCostCenterCode(ByVal isButtonClick As Boolean)
        If SettingCostCenter Then
            If clsCommon.myLen(clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyCode).Value)) > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") <> CompairStringResult.Equal Then
                    Dim qry As String = "select TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code as [Code] ,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name,ISNULL(TSPL_COST_CENTRE_FINANCIAL.Hirerachy_Level_Code,'') AS [Hirerachy Level Code],ISNULL(TSPL_COST_CENTRE_FINANCIAL.Cost_Centre_Fin_Level_Code,'') AS [Cost Centre Fin Level Code],ISNULL(TSPL_COST_CENTRE_FINANCIAL.Hirerachy_Level,'') AS [Hirerachy Level] ,TSPL_COST_CENTRE_FINANCIAL.Created_By as [Created By] ,Convert(varchar,TSPL_COST_CENTRE_FINANCIAL.Created_Date,103) as [Created Date] ,TSPL_COST_CENTRE_FINANCIAL.Modified_By as [Modified By] ,Convert(varchar,TSPL_COST_CENTRE_FINANCIAL.Modified_Date,103) as [Modified Date]  From TSPL_COST_CENTRE_FINANCIAL "
                    gv1.CurrentRow.Cells(colCostCenterCode).Value = clsCommon.ShowSelectForm("Hierarchy", qry, "Code", " Hirerachy_Level = '" + clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyLevelNumber).Value) + "'", clsCommon.myCstr(gv1.CurrentRow.Cells(colCostCenterCode).Value), "Code", isButtonClick)
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "Please select hirerachy level first.", Me.Text)
            End If
        End If
    End Sub


    Private Sub OpenHirerchylevelName(ByVal strHierarchyCode As String)
        Dim HirerchylevelName As String = String.Empty
        Try
            Dim DBLevel As String = String.Empty
            DBLevel = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select ISNULL(Level,'') AS Level from TSPL_HIRERACHY_LEVEL_MASTER Where HIRERACHY_CODE='" + strHierarchyCode + "' "))
            gv1.Columns(colCostCenterCode).HeaderText = clsCommon.myCstr(clsDBFuncationality.getSingleValue("sELECT Description FROM TSPL_HIRERACHY_LEVEL_MASTER WHERE Hirerachy_Code ='" + strHierarchyCode + "' "))
            If clsCommon.CompairString(DBLevel, "4") = CompairStringResult.Equal Then
                gv1.Columns(colHierarchyLevel3).HeaderText = clsCommon.myCstr(clsDBFuncationality.getSingleValue("sELECT Description FROM TSPL_HIRERACHY_LEVEL_MASTER WHERE Level ='3' "))
                gv1.Columns(colHierarchyLevel4).HeaderText = clsCommon.myCstr(clsDBFuncationality.getSingleValue("sELECT Description FROM TSPL_HIRERACHY_LEVEL_MASTER WHERE Level ='2' "))
            ElseIf clsCommon.CompairString(DBLevel, "3") = CompairStringResult.Equal Then
                gv1.Columns(colHierarchyLevel3).HeaderText = clsCommon.myCstr(clsDBFuncationality.getSingleValue("sELECT Description FROM TSPL_HIRERACHY_LEVEL_MASTER WHERE Level ='2' "))
                gv1.Columns(colHierarchyLevel4).HeaderText = "Hierarchy Level 4"
            Else
                gv1.Columns(colHierarchyLevel3).HeaderText = "Hierarchy Level 3"
                gv1.Columns(colHierarchyLevel4).HeaderText = "Hierarchy Level 4"
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, "Please select hirerachy level first.", Me.Text)
        End Try
    End Sub
    ''-------------
    Private Sub setGridFocus()
        Dim intCurrRow As Integer = gv1.CurrentRow.Index
        gv1.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
        If intCurrRow = gv1.Rows.Count - 1 Then
            gv1.Rows.AddNew()
            gv1.CurrentRow = gv1.Rows(intCurrRow)
        End If
        If clsCommon.myLen(gv1.CurrentRow.Cells(colACCode).Value) > 0 Then
            If gv1.CurrentColumn Is gv1.Columns(colACCode) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colAmt)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colAmt) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow)
                gv1.CurrentColumn = gv1.Columns(colDisPer)
            ElseIf gv1.CurrentColumn Is gv1.Columns(colDisPer) Then
                gv1.CurrentRow = gv1.Rows(intCurrRow + 1)
                gv1.CurrentColumn = gv1.Columns(colACCode)
            End If
        End If

    End Sub

    Private Sub OpenGLAccount(ByVal isButtonClick As Boolean)
        Dim qry As String
        If txtlocation.Value = "" Then
            common.clsCommon.MyMessageBoxShow(Me, "Please first select Location", Me.Text)
            txtlocation.Focus()
            Return
        End If
        Dim whrcls As String
        Dim arr As New ArrayList()
        arr = clsERPFuncationality.glaccountquery(objCommonVar.CurrentUserCode)
        qry = arr.Item(0)
        whrcls = arr.Item(1)
        If whrcls = "" Then
            whrcls = "2=2 AND ControlAccount<>'Y'"
        Else
            whrcls = "(" + whrcls + ")"
        End If

        whrcls += "  and  TSPL_GL_ACCOUNTS.Account_Seg_Code7='" + txtlocation.Value + "' AND ControlAccount<>'Y'"
        If chkAgainstSecurityReceiptNo.Checked OrElse clsCommon.GetDateWithStartTime(txtDate.Value) < clsCommon.GetDateWithStartTime(ERPStartDate) Then
            whrcls = clsCommon.ReplaceString(whrcls, "TSPL_GL_ACCOUNTS.ControlAccount='N'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "ControlAccount ='N'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "( ControlAccount ='N')", " ")

            whrcls = clsCommon.ReplaceString(whrcls, "TSPL_GL_ACCOUNTS.ControlAccount<>'Y'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "ControlAccount<>'Y'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "( ControlAccount<>'Y')", " ")
        End If

        ''richa 05 feb,2019  TEC/05/02/19-000411 not check for opening in case of Credit or debit note 
        'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, Nothing)
        Dim JEWithOPening As Boolean = False
        If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
            Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
            If clsCommon.myCDate(clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                JEWithOPening = True
            End If
        End If
        Dim strCustomerOpeningAccount As String = String.Empty
        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, Nothing)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "C") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "D") = CompairStringResult.Equal) And JEWithOPening = True Then
            If clsCommon.myLen(TxtCustomer.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please first select Customer", Me.Text)
                TxtCustomer.Focus()
                Return
            End If
            strCustomerOpeningAccount = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select left(Customer_Opening_Clearing_AC,len(Customer_Opening_Clearing_AC)-4)  from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account in (select Cust_Account from TSPL_CUSTOMER_MASTER where cust_code='" & TxtCustomer.Value & "')"))
            If clsCommon.myLen(strCustomerOpeningAccount) = 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please set Customer Opening Clearing Account for customer - " + TxtCustomer.Value, Me.Text)
                Return
            End If
            whrcls += " and Account_Seg_Code1='" & strCustomerOpeningAccount & "'"
        End If

        'richa 17 SEp,2019 TEC/03/07/19-000927
        Dim strqry As String = " Select Account_Code,Description from (" & qry & " where " & whrcls & Environment.NewLine &
            " UNION All " & Environment.NewLine &
            " select Account_Code , Description  from TSPL_GL_ACCOUNTS " & Environment.NewLine &
" left outer join (select TSPL_GL_SEGMENT_CODE.Account_Code as AccCode from TSPL_GL_SEGMENT_CODE where TSPL_GL_SEGMENT_CODE.Seg_No='7' " & Environment.NewLine &
" and len(isnull(TSPL_GL_SEGMENT_CODE.Account_Code,''))>0 ) as segTable  on segTable.AccCode=TSPL_GL_ACCOUNTS.Account_Code " & Environment.NewLine &
    " inner join TSPL_GL_STRUCTURE on TSPL_GL_ACCOUNTS .Str_Code=TSPL_GL_STRUCTURE.Str_Code where ( 2=2  and TSPL_GL_ACCOUNTS.Status='Y' and ( segTable.AccCode is null  ))" & Environment.NewLine &
    " and 1<>(isnull(Seg_No1,0) +isnull(Seg_No2,0) +isnull(Seg_No3,0) +isnull(Seg_No4,0) +isnull(Seg_No5,0) +isnull(Seg_No6,0) +isnull(Seg_No7,0) +isnull(Seg_No8,0) +isnull(Seg_No9,0) +isnull(Seg_No10,0) ) " & Environment.NewLine &
    " and TSPL_GL_ACCOUNTS.Account_Code in (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForAR =1) and  TSPL_GL_ACCOUNTS.Account_Seg_Code7='" + txtlocation.Value + "' "

        If clsCommon.myLen(strCustomerOpeningAccount) > 0 Then
            strqry += " and Account_Seg_Code1='" & strCustomerOpeningAccount & "'"
        End If
        strqry += " ) Final "
        gv1.CurrentRow.Cells(colACCode).Value = clsCommon.ShowSelectForm("TaxRateChange", strqry, "Account_Code", "", clsCommon.myCstr(gv1.CurrentRow.Cells(colACCode).Value), "Account_Code", isButtonClick)
        'gv1.CurrentRow.Cells(colACCode).Value = clsCommon.ShowSelectForm("TaxRateChange", qry, "Account_Code", whrcls, clsCommon.myCstr(gv1.CurrentRow.Cells(colACCode).Value), "Account_Code", isButtonClick)
        gv1.CurrentRow.Cells(colACName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_GL_ACCOUNTS where Account_Code='" + clsCommon.myCstr(gv1.CurrentRow.Cells(colACCode).Value) + "'"))
        If clsCommon.myLen(gv1.CurrentRow.Cells(colACCode).Value) > 0 Then
            gv1.CurrentRow.Cells(colGLType).Value = clsPaymentHeader.CheckGLAccountType(clsCommon.myCstr(gv1.CurrentRow.Cells(colACCode).Value), Nothing)
        Else
            gv1.CurrentRow.Cells(colGLType).Value = ""
        End If
        txtlocation.Enabled = False
    End Sub

    Private Sub gv1_UserAddedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserAddedRow
        For i As Integer = 0 To gv1.Rows.Count - 1
            gv1.Rows(0).Cells(0).Value = 1
            If i <> 0 Then
                gv1.Rows(i).Cells(colLineNo).Value = i + 1
            End If
        Next
    End Sub

    Private Sub BlankTaxDetailsCurrentRow()
        For ii As Integer = 1 To 10
            Dim strII As String = clsCommon.myCstr(ii)
            gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = Nothing
            gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
            gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
        Next
    End Sub

    Private Function GetCurrentRowTotalTaxAmt() As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
        Next
        Return dblTotTax
    End Function

    Private Function GetCurrentRowSurTaxAmt(ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                Return clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
            End If
        Next
        Return 0
    End Function

    Private Function GetCurrentRowOtherTaxAmt(ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Double
        Dim dblRet As Double = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    dblRet = dblRet + clsCommon.myCdbl(gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            Next
        Next
        Return dblRet
    End Function

    Private Sub UpdateCurrentRowWihtRowNo(ByVal intRowNo As Integer)
        Dim arrTaxableAuth As New List(Of String)
        Dim dblAmt As Double = clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(colAmt).Value)
        Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(colDisPer).Value)
        Dim dblDisAmt As Double = (dblAmt * dblDisPer) / 100
        Dim dblAmtAfterDis As Double = dblAmt - dblDisAmt
        BlankTaxDetailsCurrentRowWihtRowNo(intRowNo)
        Dim ii As String = "1"
        For Each gr As GridViewRowInfo In gv2.Rows
            If rbtnTaxCalAutomatic.IsChecked Then
                Dim strTaxCode As String = clsCommon.myCstr(gr.Cells(colTTaxAutCode).Value)
                Dim dblTaxRate As Double = clsCommon.myCdbl(gr.Cells(colTTaxRate).Value)
                Dim IsSurTax As Boolean = clsCommon.myCBool(gr.Cells(colTIsSurTax).Value)
                Dim strSurTaxCode As String = clsCommon.myCstr(gr.Cells(colTSurTaxCode).Value)
                Dim IsTaxable As Boolean = clsCommon.myCBool(gr.Cells(colTIsTaxable).Value)
                Dim dblTaxAmt As Double = 0
                Dim dblBaseAmt As Double = 0
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + ii)).Value = strTaxCode
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + ii)).Value = dblTaxRate
                If IsSurTax Then
                    Dim dblSurTaxAmt As Double = GetCurrentRowSurTaxAmtWihtRowNo(intRowNo, ii, strSurTaxCode)
                    dblBaseAmt = dblSurTaxAmt
                    dblTaxAmt = (dblSurTaxAmt * dblTaxRate) / 100
                Else
                    Dim dblOtherTaxAmt As Double = 0
                    ''If IsTaxable Then
                    dblOtherTaxAmt = GetCurrentRowOtherTaxAmtWihtRowNo(intRowNo, ii, arrTaxableAuth)
                    ''End If
                    dblBaseAmt = (dblAmtAfterDis)
                    dblTaxAmt = (dblAmtAfterDis * dblTaxRate) / 100
                End If
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + ii)).Value = Math.Round(dblBaseAmt, 2)
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + ii)).Value = Math.Round(dblTaxAmt, IIf(objCommonVar.IsRoundOffTaxToZeroDecimal, 0, 2))
                If IsTaxable AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper()) Then
                    arrTaxableAuth.Add(strTaxCode.ToUpper())
                End If
            Else
                Dim dblTaxAmt As Double = clsCommon.myCdbl(gr.Cells(colTTaxAmt).Value)
                Dim dblCurrRowAmt As Double = clsCommon.myCdbl(gv1.Rows(clsCommon.myCdbl(intRowNo)).Cells(colAmt).Value)
                Dim dblTotAmt As Double = 0
                For jj As Integer = 0 To gv1.Rows.Count - 1
                    dblTotAmt += clsCommon.myCdbl(gv1.Rows(jj).Cells(colAmt).Value)
                Next
                Dim dblCurrCalTax As Double = 0
                If dblTotAmt <> 0 Then
                    dblCurrCalTax = Math.Round(clsCommon.myCdbl(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 2, MidpointRounding.ToEven)
                End If
                gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + ii)).Value = dblCurrCalTax
            End If
            ii = clsCommon.myCstr(clsCommon.myCdbl(ii) + 1)
        Next
        Dim dblTotTaxAmt As Double = GetCurrentRowTotalTaxAmtWihtRowNo(intRowNo)
        Dim dblAmtAfterTax As Double = dblAmtAfterDis + dblTotTaxAmt
        gv1.Rows(intRowNo).Cells(colDisAmt).Value = dblDisAmt
        gv1.Rows(intRowNo).Cells(colAmtAfterDis).Value = dblAmtAfterDis
        gv1.Rows(intRowNo).Cells(colTotTaxAmt).Value = dblTotTaxAmt
        gv1.Rows(intRowNo).Cells(colAmtAfterTax).Value = dblAmtAfterTax
    End Sub

    Private Function GetCurrentRowOtherTaxAmtWihtRowNo(ByVal intRowNo As Integer, ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Double
        Dim dblRet As Double = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    dblRet = dblRet + clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            Next
        Next
        Return dblRet
    End Function

    Private Sub BlankTaxDetailsCurrentRowWihtRowNo(ByVal intRowNo As Integer)
        For ii As Integer = 1 To 10
            Dim strII As String = clsCommon.myCstr(ii)
            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = Nothing
            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
            gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
        Next
    End Sub

    Private Function GetCurrentRowSurTaxAmtWihtRowNo(ByVal intRowNo As Integer, ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                Return clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
            End If
        Next
        Return 0
    End Function

    Private Function GetCurrentRowTotalTaxAmtWihtRowNo(ByVal intRowNo As Integer) As Double
        Dim dblTotTax As Double = 0
        For ii As Integer = 1 To 10
            Dim strii As String = clsCommon.myCstr(ii)
            dblTotTax = dblTotTax + clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
        Next
        Return dblTotTax
    End Function

    Private Sub UpdateAllTotals()
        Dim dblTotAmt As Double = 0
        Dim dblTotDisAmt As Double = 0
        Dim dblAmtAfterDis As Double = 0
        Dim dblTaxAmt1 As Double = 0
        Dim dblTaxAmt2 As Double = 0
        Dim dblTaxAmt3 As Double = 0
        Dim dblTaxAmt4 As Double = 0
        Dim dblTaxAmt5 As Double = 0
        Dim dblTaxAmt6 As Double = 0
        Dim dblTaxAmt7 As Double = 0
        Dim dblTaxAmt8 As Double = 0
        Dim dblTaxAmt9 As Double = 0
        Dim dblTaxAmt10 As Double = 0
        Dim dblTaxRate1 As Double = 0
        Dim dblTaxRate2 As Double = 0
        Dim dblTaxRate3 As Double = 0
        Dim dblTaxRate4 As Double = 0
        Dim dblTaxRate5 As Double = 0
        Dim dblTaxRate6 As Double = 0
        Dim dblTaxRate7 As Double = 0
        Dim dblTaxRate8 As Double = 0
        Dim dblTaxRate9 As Double = 0
        Dim dblTaxRate10 As Double = 0
        Dim dblTaxTotAmt As Double = 0
        Dim dblNetAmt As Double = 0
        Dim dblTaxBaseAmt1 As Double = 0
        Dim dblTaxBaseAmt2 As Double = 0
        Dim dblTaxBaseAmt3 As Double = 0
        Dim dblTaxBaseAmt4 As Double = 0
        Dim dblTaxBaseAmt5 As Double = 0
        Dim dblTaxBaseAmt6 As Double = 0
        Dim dblTaxBaseAmt7 As Double = 0
        Dim dblTaxBaseAmt8 As Double = 0
        Dim dblTaxBaseAmt9 As Double = 0
        Dim dblTaxBaseAmt10 As Double = 0

        For ii As Integer = 0 To gv1.Rows.Count - 1
            If (clsCommon.myLen(gv1.Rows(ii).Cells(colACCode).Value) > 0) Then
                dblTotAmt = dblTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmt).Value)
                dblTotDisAmt = dblTotDisAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colDisAmt).Value)
                dblAmtAfterDis = dblAmtAfterDis + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmtAfterDis).Value)
                dblTaxAmt1 = dblTaxAmt1 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt1).Value)
                dblTaxAmt2 = dblTaxAmt2 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt2).Value)
                dblTaxAmt3 = dblTaxAmt3 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt3).Value)
                dblTaxAmt4 = dblTaxAmt4 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt4).Value)
                dblTaxAmt5 = dblTaxAmt5 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt5).Value)
                dblTaxAmt6 = dblTaxAmt6 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt6).Value)
                dblTaxAmt7 = dblTaxAmt7 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt7).Value)
                dblTaxAmt8 = dblTaxAmt8 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt8).Value)
                dblTaxAmt9 = dblTaxAmt9 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt9).Value)
                dblTaxAmt10 = dblTaxAmt10 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxAmt10).Value)
                dblTaxTotAmt = dblTaxTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTotTaxAmt).Value)
                dblTaxRate1 = dblTaxRate1 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate1).Value)
                dblTaxRate2 = dblTaxRate2 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate2).Value)
                dblTaxRate3 = dblTaxRate3 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate3).Value)
                dblTaxRate4 = dblTaxRate4 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate4).Value)
                dblTaxRate5 = dblTaxRate5 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate5).Value)
                dblTaxRate6 = dblTaxRate6 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate6).Value)
                dblTaxRate7 = dblTaxRate7 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate7).Value)
                dblTaxRate8 = dblTaxRate8 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate8).Value)
                dblTaxRate9 = dblTaxRate9 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate9).Value)
                dblTaxRate10 = dblTaxRate10 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxRate10).Value)
                dblTaxTotAmt = dblTaxTotAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTotTaxAmt).Value)
                dblNetAmt = dblNetAmt + clsCommon.myCdbl(gv1.Rows(ii).Cells(colAmtAfterTax).Value)

                dblTaxBaseAmt1 = dblTaxBaseAmt1 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt1).Value)
                dblTaxBaseAmt2 = dblTaxBaseAmt2 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt2).Value)
                dblTaxBaseAmt3 = dblTaxBaseAmt3 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt3).Value)
                dblTaxBaseAmt4 = dblTaxBaseAmt4 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt4).Value)
                dblTaxBaseAmt5 = dblTaxBaseAmt5 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt5).Value)
                dblTaxBaseAmt6 = dblTaxBaseAmt6 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt6).Value)
                dblTaxBaseAmt7 = dblTaxBaseAmt7 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt7).Value)
                dblTaxBaseAmt8 = dblTaxBaseAmt8 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt8).Value)
                dblTaxBaseAmt9 = dblTaxBaseAmt9 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt9).Value)
                dblTaxBaseAmt10 = dblTaxBaseAmt10 + clsCommon.myCdbl(gv1.Rows(ii).Cells(colTaxBaseAmt10).Value)

            End If
        Next
        If rbtnTaxCalAutomatic.IsChecked Then
            For ii As Integer = 1 To gv2.Rows.Count
                Select Case (ii)
                    Case 1
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt1
                        If dblTaxAmt1 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt1 * 100) / (dblTaxBaseAmt1), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 2
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt2
                        If dblTaxAmt2 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt2 * 100) / (dblTaxBaseAmt2), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 3
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt3
                        If dblTaxAmt3 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt3 * 100) / (dblTaxBaseAmt3), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 4
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt4
                        If dblTaxAmt4 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt4 * 100) / (dblTaxBaseAmt4), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 5
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt5
                        If dblTaxAmt5 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt5 * 100) / (dblTaxBaseAmt5), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 6
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt6
                        If dblTaxAmt6 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt6 * 100) / (dblTaxBaseAmt6), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 7
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt7
                        If dblTaxAmt7 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt7 * 100) / (dblTaxBaseAmt7), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 8
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt8
                        If dblTaxAmt7 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt8 * 100) / (dblTaxBaseAmt8), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 9
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt9
                        If dblTaxAmt7 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt9 * 100) / (dblTaxBaseAmt9), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                    Case 10
                        gv2.Rows(ii - 1).Cells(colTTaxAmt).Value = dblTaxAmt10
                        If dblTaxAmt7 > 0 Then
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = clsCommon.myRoundOFF((dblTaxAmt10 * 100) / (dblTaxBaseAmt10), 1, 4)
                        Else
                            gv2.Rows(ii - 1).Cells(colTTaxRate).Value = 0
                        End If
                End Select
                Dim dblBaseAmt As Double = 0
                If clsCommon.myCBool(gv2.Rows(ii - 1).Cells(colTIsSurTax).Value) Then
                    dblBaseAmt = GetBaseTaxAmount(ii - 1, clsCommon.myCstr(gv2.Rows(ii - 1).Cells(colTSurTaxCode).Value))
                Else
                    Dim dblOtherTaxAmt = 0
                    'If clsCommon.myCBool(gv2.Rows(ii - 1).Cells(colTIsTaxable).Value) Then
                    dblOtherTaxAmt = GetBaseOtherTaxableAmount(ii - 1)
                    ''End If

                    dblBaseAmt = dblAmtAfterDis
                End If
                gv2.Rows(ii - 1).Cells(colTBaseAmt).Value = dblBaseAmt
            Next
        End If

        Dim dblACAmount As Double = 0
        For ii As Integer = 0 To gvAC.Rows.Count - 1
            If (clsCommon.myLen(gvAC.Rows(ii).Cells(colAChgCode).Value) > 0) Then
                dblACAmount = dblACAmount + clsCommon.myCdbl(gvAC.Rows(ii).Cells(colAChgAmount).Value)
            End If
        Next

        lblAmtWithDiscount.Text = clsCommon.myFormat(dblTotAmt)
        lblDiscountAmt.Text = clsCommon.myFormat(dblTotDisAmt)
        lblAmtAfterDiscount.Text = clsCommon.myFormat(dblAmtAfterDis)
        lblTaxAmt.Text = clsCommon.myFormat(dblTaxTotAmt)

        lblAddCharges.Text = clsCommon.myFormat(dblACAmount)
        lblAddCharges1.Text = clsCommon.myFormat(dblACAmount)
        dblNetAmt = dblNetAmt + dblACAmount

        lblTotRAmt.Text = clsCommon.myFormat(dblNetAmt)
        lblTotRAmt1.Text = clsCommon.myFormat(dblNetAmt)
    End Sub

    Private Function GetBaseOtherTaxableAmount(ByVal intEndCol As Integer) As Double
        Dim dblRetVal As Double = 0
        For ii As Integer = 0 To intEndCol - 1
            If clsCommon.myCBool(gv2.Rows(ii).Cells(colTIsTaxable).Value) Then
                dblRetVal = dblRetVal + clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
            End If
        Next
        Return dblRetVal
    End Function

    Private Function GetBaseTaxAmount(ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double

        For ii As Integer = 0 To intEndCol
            If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAutCode).Value)) = CompairStringResult.Equal Then
                Return clsCommon.myCdbl(clsCommon.myCstr(gv2.Rows(ii).Cells(colTTaxAmt).Value))
            End If
        Next
        Return 0
    End Function

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        AddNew()

    End Sub

    Sub AddNew()
        butCostCenterAndHirerachy_Update_AfterPost.Visible = False
        txtlocation.Enabled = True
        txtlocation.Value = ""
        LblLocDesp.Text = ""
        txtLocationPrefix.Value = ""
        txtLocationPrefixName.Text = ""
        EInvoiceIRNNo.Text = ""
        EInvoiceQrCode.Text = ""
        EinvoiceBtnUpdate.Enabled = True
        BlankAllControls()
        LoadBlankGridGL()
        LoadBlankGridTax()
        LoadBlankGridAC()
        cboDocType.Enabled = True
        isNewEntry = True
        btnSave.Text = "Save"
        btnSave.Enabled = True
        btnPost.Enabled = True
        btnDelete.Enabled = True
        btnPrintServiceInvoice.Visible = False
        gv1.Rows.AddNew()
        gvAC.Rows.AddNew()
        ApplyReadOption()
        ValidateLength()
        txtTapalNo.Text = ""

        '' Anubhooti 05-Dec-2014 (InvoiceNo. finder should visible in case of D/C)
        If ((clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal)) Then
            txtRefDocNo.Visible = True
            RadLabel15.Visible = True
            txtRefDocNo.Value = ""
        Else
            txtRefDocNo.Visible = False
            RadLabel15.Visible = False
            txtRefDocNo.Value = ""
        End If
        ''For Custom Fields
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.SetDefaultValues()
        End If
        ''End of For Custom Fields
        UcAttachment1.BlankAllControls()
        txtlocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code in (select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' ) "))
        '' Anubhooti 28-Nov-2014 BM00000004810
        If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
            LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
        Else
            LblLocDesp.Text = ""
        End If

    End Sub

    Function AllowToSave() As Boolean
        Try
            btnSave.Focus()
            Xtra.TransactionValidity(txtDate.Value)
            '========================Added by Preeti Gupta[01/02/2017]==============
            If AllowFutureDateTransaction(txtDate.Value, Nothing) = False Then
                txtDate.Focus()
                Return False
            End If
            If objCommonVar.ApplyLocationWisePrefix Then
                If clsCommon.myLen(txtLocationPrefix.Value) <= 0 Then
                    txtLocationPrefix.Focus()
                    Throw New Exception("Please select Location")
                End If
            End If
            '===============================================================================
            If btnSave.Text = "Update" Then
                Dim strchk As String = "select Status from TSPL_Customer_Invoice_Head where Document_No='" + txtDocNo.Value + "'"
                Dim chkpost As String = clsDBFuncationality.getSingleValue(strchk)
                If chkpost = "1" Then
                    clsCommon.MyMessageBoxShow(Me, "Transaction already posted", Me.Text)
                    Return False
                End If
            End If
            If chkSecurityDposit.Checked Then
                If clsCommon.myLen(ddlSecDepositType.SelectedValue) <= 0 Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please select security deposit type.", Me.Text)
                    ddlSecDepositType.Focus()
                    Return False
                End If
            End If
            UpdateAllTotals()
            If clsCommon.myLen(txtlocation.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please first select Location", Me.Text)
                txtlocation.Focus()
                Return False
            End If
            '' Added BY abhishek as on 1 Dec 2012 For Lock Transaction Location Wise
            Dim LocSeg As String = txtlocation.Value
            Dim qry1 As String = " select Description  from TSPL_GL_SEGMENT_CODE  where Seg_No ='7' and  Segment_code ='" + LocSeg + "'"
            Dim LocSegName As String = clsDBFuncationality.getSingleValue(qry1)

            If clsCommon.myLen(TxtCustomer.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer", Me.Text)
                TxtCustomer.Focus()
                Return False
            End If
            If clsCommon.myLen(txtACSet.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer Account Set", Me.Text)
                txtACSet.Focus()
                Return False
            End If

            If clsCommon.myLen(txtTaxGroup.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Tax Group", Me.Text)
                txtTaxGroup.Focus()
                Return False
            End If
            '' Anubhooti 03-Dec-2014 (Terms Code is manadatory)
            If clsCommon.myLen(txtTermCode.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Terms Code", Me.Text)
                Me.RadPageView1.SelectedPage = RadPageViewPage2
                txtTermCode.Focus()
                Return False
            End If
            Dim DrCr As Integer = 0
            If clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal Or clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal Then
                DrCr = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("SELECT COUNT(*) AS ROW FROM TSPL_CUSTOMER_INVOICE_HEAD WHERE DOCUMENT_NO='" + clsCommon.myCstr(txtRefDocNo.Value) + "' AND Document_Type IN ('C','D')"))
                If clsCommon.myCdbl(DrCr) > 0 Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please select ref doc type.", Me.Text)
                    txtRefDocNo.Focus()
                    Return False
                End If
            End If
            ''
            If clsCommon.myLen(txtRefDocNo.Value) > 0 Then
                Dim qry As String
                If objCommonVar.IsDemoERP Then
                    qry = "select 1 from TSPL_Customer_Invoice_Head where Status=1 and Customer_Code='" + TxtCustomer.Value + "' and Document_No='" + txtRefDocNo.Value + "'"
                Else
                    qry = "select 1 from TSPL_Customer_Invoice_Head where Status=1 and Customer_Code='" + TxtCustomer.Value + "' and Document_No='" + txtRefDocNo.Value + "'"
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    common.clsCommon.MyMessageBoxShow("Customer is not valid for invoice No " + txtRefDocNo.Value, Me.Text)
                    Return False
                End If
            End If
            For ii As Integer = 0 To gvAC.Rows.Count - 1
                If clsCommon.myLen(gvAC.Rows(ii).Cells(colAChgCode).Value) > 0 Then
                    For jj As Integer = 0 To gvAC.Rows.Count - 1
                        If ii = jj Then
                            Continue For
                        End If
                        If clsCommon.CompairString(clsCommon.myCstr(gvAC.Rows(ii).Cells(colAChgCode).Value), clsCommon.myCstr(gvAC.Rows(jj).Cells(colAChgCode).Value)) = CompairStringResult.Equal Then
                            common.clsCommon.MyMessageBoxShow("Additional Charges: " + clsCommon.myCstr(gvAC.Rows(ii).Cells(colAChgCode).Value) + "Repeated at Row No " + clsCommon.myCstr(ii + 1) + " and " + clsCommon.myCstr(jj + 1) + "")
                            Return False
                        End If
                    Next
                End If

                If SettingCostCenter AndAlso clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colACCode).Value)) > 0 Then
                    Dim grouptype As String = clsPaymentHeader.CheckGLAccountType(clsCommon.myCstr(gv1.Rows(ii).Cells(colACCode).Value), Nothing)
                    If Not clsCommon.CompairString(grouptype, "Balance Sheet") = CompairStringResult.Equal Then
                        If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colHierarchyCode).Value)) <= 0 Then
                            common.clsCommon.MyMessageBoxShow(Me, " Please select Hierarchy ", Me.Text)
                            Return False
                        End If
                        If SettingCostCenterlevel Then
                            Dim lvl As Integer = clsCommon.myCdbl(gv1.Rows(ii).Cells(colHierarchyLevelNumber).Value)
                            If lvl > 0 Then
                                gv1.Rows(ii).Cells(colCostCenterCode).Value = clsCommon.myCstr(gv1.Rows(ii).Cells("colHierarchyLevel" + clsCommon.myCstr(lvl)).Value)
                                For pp As Integer = lvl To 1 Step -1
                                    If clsCommon.myLen(gv1.Rows(ii).Cells("colHierarchyLevel" + clsCommon.myCstr(pp)).Value) <= 0 Then
                                        common.clsCommon.MyMessageBoxShow("Please select " & gv1.Columns("colHierarchyLevel" + clsCommon.myCstr(pp)).HeaderText)
                                        Return False
                                    End If
                                Next
                            End If
                        End If
                        If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(ii).Cells(colCostCenterCode).Value)) <= 0 Then
                            common.clsCommon.MyMessageBoxShow(Me, " Please select Cost Center ", Me.Text)
                            Return False
                        End If
                    End If
                End If
            Next

            GSTStatus = clsERPFuncationality.GetGSTStatus(txtDate.Value)
            If GSTStatus = True Then
                If ((clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal)) Then
                    Dim strcreditNoteSetting As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowCreditNoteWithoutReference, clsFixedParameterCode.AllowCreditNoteWithoutReference, Nothing))
                    If clsCommon.CompairString(strcreditNoteSetting, "1") = CompairStringResult.Equal Then
                        ' Dim strtaxgroup As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select distinct TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code from TSPL_TAX_GROUP_MASTER left outer join TSPL_LOCATION_WISE_TAX_MASTER on TSPL_LOCATION_WISE_TAX_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_MASTER.Tax_Group_Code where Location_Code='" & clsCommon.myCstr(txtlocation.Value) & "' and TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted=1 and Tax_Type ='S'  and  TSPL_TAX_GROUP_MASTER.Tax_Group_Code='" & clsCommon.myCstr(txtTaxGroup.Value) & "'"))
                        Dim strtaxgroup As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select distinct Tax_Group_Code from TSPL_TAX_GROUP_MASTER  where TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted=1 and Tax_Group_Type='S' and TSPL_TAX_GROUP_MASTER.Tax_Group_Code='" & clsCommon.myCstr(txtTaxGroup.Value) & "'"))
                        If clsCommon.myLen(strtaxgroup) <= 0 Then
                            If clsCommon.myLen(clsCommon.myCstr(txtRefDocNo.Value)) <= 0 Then
                                common.clsCommon.MyMessageBoxShow(Me, " Please select Invoice No ", Me.Text)
                                Return False
                            End If
                        End If
                    Else
                        ''richa 05 feb,2019  TEC/05/02/19-000411 not check for opening in case of Credit or debit note 
                        'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, Nothing)
                        Dim JEWithOPening As Boolean = False
                        If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                            Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                            If clsCommon.myCDate(clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                                JEWithOPening = True
                            End If
                        End If
                        If JEWithOPening = False Then
                            If clsCommon.myLen(clsCommon.myCstr(txtRefDocNo.Value)) <= 0 AndAlso Not chkAgainstSecurityReceiptNo.Checked Then
                                common.clsCommon.MyMessageBoxShow(Me, " Please select Invoice No ", Me.Text)
                                Return False
                            End If
                        End If

                    End If
                End If
            End If
            If chkAgainstSecurityReceiptNo.Checked Then
                If clsCommon.myLen(txtSercurityReceiptNo.Value) <= 0 Then
                    txtSercurityReceiptNo.Focus()
                    Throw New Exception("Please select Security receipt No")
                End If
                If clsCommon.myLen(txtRefDocNo.Value) > 0 Then
                    txtRefDocNo.Focus()
                    Throw New Exception("Please do not enter " + txtRefDocNo.MyLinkLable1.Text)
                End If
            End If
            Dim arrAccountCode As New List(Of String)
            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strACode As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colACCode).Value)
                If clsCommon.myLen(strACode) > 0 Then
                    If Not arrAccountCode.Contains(strACode) Then
                        arrAccountCode.Add(strACode)
                    End If
                End If
            Next
            ''richa UDL/13/03/19-000279 add chkSecurityDposit.Checked = False condition in below condition
            If chkAgainstServiceInvoice.Checked = False Then
                If arrAccountCode IsNot Nothing AndAlso arrAccountCode.Count > 0 Then
                    If (Not chkAgainstSecurityReceiptNo.Checked AndAlso chkSecurityDposit.Checked = False) AndAlso clsCommon.GetDateWithStartTime(txtDate.Value) >= clsCommon.GetDateWithStartTime(ERPStartDate) Then
                        'richa 17 SEp,2019 TEC/03/07/19-000927
                        Dim qry As String = "select Account_Code from TSPL_GL_ACCOUNTS where Account_Code in (" + clsCommon.GetMulcallString(arrAccountCode) + ") and ControlAccount<>'N' AND TSPL_GL_ACCOUNTS.Account_Code NOT IN  (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForAR  =1)"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            Throw New Exception("Can not select control Account -" + clsCommon.myCstr(dt.Rows(0)("Account_Code")))
                        End If
                    End If
                End If

            End If
            For irow As Integer = 0 To gv1.Rows.Count - 1
                Dim strTaxCode As String = ""
                Dim strTaxRate As Double = 0
                Dim strInnerTaxCode As String = ""
                Dim strInnerTaxRate As Double = 0
                For ii As Integer = 0 To gv1.Columns.Count - 1
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(irow).Cells(ii).Value), "CGST") = CompairStringResult.Equal Then
                        strTaxCode = clsCommon.myCstr(gv1.Rows(irow).Cells(ii).Value)
                        strTaxRate = clsCommon.myCdbl(gv1.Rows(irow).Cells(ii + 2).Value)
                    End If
                Next
                For jj As Integer = 0 To gv1.Columns.Count - 1
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(irow).Cells(jj).Value), "SGST") = CompairStringResult.Equal Then
                        strInnerTaxCode = clsCommon.myCstr(gv1.Rows(irow).Cells(jj).Value)
                        strInnerTaxRate = clsCommon.myCdbl(gv1.Rows(irow).Cells(jj + 2).Value)
                    End If
                Next

                If clsCommon.CompairString(strTaxCode, "CGST") = CompairStringResult.Equal AndAlso clsCommon.CompairString(strInnerTaxCode, "SGST") = CompairStringResult.Equal Then
                    If strTaxRate <> strInnerTaxRate Then
                        clsCommon.MyMessageBoxShow(Me, "Tax rate mismatch [" & strTaxCode & "-" & strTaxRate & "] and [ " & strInnerTaxCode & "-" & strInnerTaxRate & " ] at Line no[ " & clsCommon.myCstr(irow + 1) & " ]", Me.Text)
                        Return False
                    End If
                End If

            Next

            UcCustomFields1.AllowToSave()
            UcAttachment1.AllowToSave()
            clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleReceivable, clsUserMgtCode.mbtnARInvoiceEntry, txtlocation.Value, txtDate.Value, Nothing)

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
        Return True
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveData(clicked)

    End Sub

    Sub SaveData(ByVal clicked As Boolean)
        Try
            If (AllowToSave()) Then
                clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value))
                Dim obj As New clsCustomerInvoiceHead()

                obj.Document_No = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Customer_Code = TxtCustomer.Value
                obj.Customer_Name = lblVendorName.Text
                '------added by usha (31-10-2012)-----
                obj.loc_code = txtlocation.Value
                obj.Location_Code_Prefix = txtLocationPrefix.Value
                '------end-----
                obj.Account_Set = txtACSet.Value
                obj.Document_Type = clsCommon.myCstr(cboDocType.SelectedValue)
                obj.RefDocNo = txtRefDocNo.Value
                obj.Order_No = txtOrderNo.Text
                obj.Total_Tax = clsCommon.myCdbl(lblTaxAmt.Text)
                obj.On_Hold = chkOnHold.Checked
                obj.Description = txtDesc.Text
                obj.SecurityDeposit = chkSecurityDposit.Checked
                obj.SecurityDepositType = clsCommon.myCstr(ddlSecDepositType.SelectedValue)
                obj.AgainstServiceInvoice = IIf(chkAgainstServiceInvoice.Checked = True, "Y", "N")
                obj.Tax_Group = txtTaxGroup.Value
                obj.PROJECT_ID = fndProject.Value
                obj.TapalNo = clsCommon.myCstr(txtTapalNo.Text)
                obj.Is_Einvoice = IIf(chkEinvoice.Checked, 1, 0)
                If txtDataAndTimeSelection.Checked Then
                    obj.DateAndTime = txtDataAndTimeSelection.Value
                End If
                '------------------By vipin (12-04-2012)-----------------
                If (gv2.Rows.Count > 0) Then
                    obj.TAX1 = clsCommon.myCstr(gv2.Rows(0).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX1, Nothing) Then
                        'obj.TAX1_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX1)
                        obj.TAX1_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX1), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX1_Rate = clsCommon.myCdbl(gv2.Rows(0).Cells(colTTaxRate).Value)
                    obj.Tax1_BAmount = clsCommon.myCdbl(gv2.Rows(0).Cells(colTBaseAmt).Value)
                    obj.TAX1_Amt = clsCommon.myCdbl(gv2.Rows(0).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 1) Then
                    obj.TAX2 = clsCommon.myCstr(gv2.Rows(1).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX2, Nothing) Then
                        'obj.TAX2_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX2)
                        obj.TAX2_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX2), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX2_Rate = clsCommon.myCdbl(gv2.Rows(1).Cells(colTTaxRate).Value)
                    obj.Tax2_BAmount = clsCommon.myCdbl(gv2.Rows(1).Cells(colTBaseAmt).Value)
                    obj.TAX2_Amt = clsCommon.myCdbl(gv2.Rows(1).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 2) Then
                    obj.TAX3 = clsCommon.myCstr(gv2.Rows(2).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX3, Nothing) Then
                        'obj.TAX3_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX3)
                        obj.TAX3_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX3), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX3_Rate = clsCommon.myCdbl(gv2.Rows(2).Cells(colTTaxRate).Value)
                    obj.Tax3_BAmount = clsCommon.myCdbl(gv2.Rows(2).Cells(colTBaseAmt).Value)
                    obj.TAX3_Amt = clsCommon.myCdbl(gv2.Rows(2).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 3) Then
                    obj.TAX4 = clsCommon.myCstr(gv2.Rows(3).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX4, Nothing) Then
                        'obj.TAX4_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX4)
                        obj.TAX4_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX4), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX4_Rate = clsCommon.myCdbl(gv2.Rows(3).Cells(colTTaxRate).Value)
                    obj.Tax4_BAmount = clsCommon.myCdbl(gv2.Rows(3).Cells(colTBaseAmt).Value)
                    obj.TAX4_Amt = clsCommon.myCdbl(gv2.Rows(3).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 4) Then
                    obj.TAX5 = clsCommon.myCstr(gv2.Rows(4).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX5, Nothing) Then
                        'obj.TAX5_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX5)
                        obj.TAX5_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX5), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX5_Rate = clsCommon.myCdbl(gv2.Rows(4).Cells(colTTaxRate).Value)
                    obj.Tax5_BAmount = clsCommon.myCdbl(gv2.Rows(4).Cells(colTBaseAmt).Value)
                    obj.TAX5_Amt = clsCommon.myCdbl(gv2.Rows(4).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 5) Then
                    obj.TAX6 = clsCommon.myCstr(gv2.Rows(5).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX6, Nothing) Then
                        'obj.TAX6_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX6)
                        obj.TAX6_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX6), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX6_Rate = clsCommon.myCdbl(gv2.Rows(5).Cells(colTTaxRate).Value)
                    obj.Tax6_BAmount = clsCommon.myCdbl(gv2.Rows(5).Cells(colTBaseAmt).Value)
                    obj.TAX6_Amt = clsCommon.myCdbl(gv2.Rows(5).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 6) Then
                    obj.TAX7 = clsCommon.myCstr(gv2.Rows(6).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX7, Nothing) Then
                        'obj.TAX7_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX7)
                        obj.TAX7_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX7), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX7_Rate = clsCommon.myCdbl(gv2.Rows(6).Cells(colTTaxRate).Value)
                    obj.Tax7_BAmount = clsCommon.myCdbl(gv2.Rows(6).Cells(colTBaseAmt).Value)
                    obj.TAX7_Amt = clsCommon.myCdbl(gv2.Rows(6).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 7) Then
                    obj.TAX8 = clsCommon.myCstr(gv2.Rows(7).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX8, Nothing) Then
                        'obj.TAX8_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX8)
                        obj.TAX8_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX8), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX8_Rate = clsCommon.myCdbl(gv2.Rows(7).Cells(colTTaxRate).Value)
                    obj.Tax8_BAmount = clsCommon.myCdbl(gv2.Rows(7).Cells(colTBaseAmt).Value)
                    obj.TAX8_Amt = clsCommon.myCdbl(gv2.Rows(7).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 8) Then
                    obj.TAX9 = clsCommon.myCstr(gv2.Rows(8).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX9, Nothing) Then
                        'obj.TAX9_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX9)
                        obj.TAX9_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX9), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX9_Rate = clsCommon.myCdbl(gv2.Rows(8).Cells(colTTaxRate).Value)
                    obj.Tax9_BAmount = clsCommon.myCdbl(gv2.Rows(8).Cells(colTBaseAmt).Value)
                    obj.TAX9_Amt = clsCommon.myCdbl(gv2.Rows(8).Cells(colTTaxAmt).Value)
                End If
                If (gv2.Rows.Count > 9) Then
                    obj.TAX10 = clsCommon.myCstr(gv2.Rows(9).Cells(colTTaxAutCode).Value)
                    If clsTaxMaster.ISTaxRecoverableAC(obj.TAX10, Nothing) Then
                        'obj.TAX10_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX10)
                        obj.TAX10_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsTaxMaster.GetTaxPayAC(obj.TAX10), txtlocation.Value, True, Nothing)
                    End If
                    obj.TAX10_Rate = clsCommon.myCdbl(gv2.Rows(9).Cells(colTTaxRate).Value)
                    obj.Tax10_BAmount = clsCommon.myCdbl(gv2.Rows(9).Cells(colTBaseAmt).Value)
                    obj.TAX10_Amt = clsCommon.myCdbl(gv2.Rows(9).Cells(colTTaxAmt).Value)
                End If

                obj.Terms_Code = txtTermCode.Value
                obj.Terms_Description = lblTermName.Text
                obj.Due_Date = txtDueDate.Value
                obj.Discount_Base = clsCommon.myCdbl(lblAmtWithDiscount.Text)
                obj.Discount_Amount = clsCommon.myCdbl(lblDiscountAmt.Text)
                obj.Amount_Less_Discount = clsCommon.myCdbl(lblAmtAfterDiscount.Text)
                obj.Document_Total = clsCommon.myCdbl(lblTotRAmt.Text)
                obj.Balance_Amt = clsCommon.myCdbl(lblTotRAmt.Text)

                If rbtnTaxCalAutomatic.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                ElseIf rbtnTaxCalManual.IsChecked Then
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual
                End If

                Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Receivable_Control_acct,Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + txtACSet.Value + "'")

                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    obj.Customer_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dt.Rows(0)("Receivable_Control_acct")), txtlocation.Value, True, Nothing)
                    If clsCommon.myCdbl(lblDiscountAmt.Text) > 0 Then
                        obj.Discount_GL_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(clsCommon.myCstr(dt.Rows(0)("Receipts_Discount_acct")), txtlocation.Value, True, Nothing)
                    End If
                End If

                If (gvAC.Rows.Count > 0) Then
                    If clsCommon.myLen(gvAC.Rows(0).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code1 = clsCommon.myCstr(gvAC.Rows(0).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name1 = clsCommon.myCstr(gvAC.Rows(0).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt1 = clsCommon.myCdbl(gvAC.Rows(0).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 1) Then
                    If clsCommon.myLen(gvAC.Rows(1).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code2 = clsCommon.myCstr(gvAC.Rows(1).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name2 = clsCommon.myCstr(gvAC.Rows(1).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt2 = clsCommon.myCdbl(gvAC.Rows(1).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 2) Then
                    If clsCommon.myLen(gvAC.Rows(2).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code3 = clsCommon.myCstr(gvAC.Rows(2).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name3 = clsCommon.myCstr(gvAC.Rows(2).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt3 = clsCommon.myCdbl(gvAC.Rows(2).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 3) Then
                    If clsCommon.myLen(gvAC.Rows(3).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code4 = clsCommon.myCstr(gvAC.Rows(3).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name4 = clsCommon.myCstr(gvAC.Rows(3).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt4 = clsCommon.myCdbl(gvAC.Rows(3).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 4) Then
                    If clsCommon.myLen(gvAC.Rows(4).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code5 = clsCommon.myCstr(gvAC.Rows(4).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name5 = clsCommon.myCstr(gvAC.Rows(4).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt5 = clsCommon.myCdbl(gvAC.Rows(4).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 5) Then
                    If clsCommon.myLen(gvAC.Rows(5).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code6 = clsCommon.myCstr(gvAC.Rows(5).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name6 = clsCommon.myCstr(gvAC.Rows(5).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt6 = clsCommon.myCdbl(gvAC.Rows(5).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 6) Then
                    If clsCommon.myLen(gvAC.Rows(6).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code7 = clsCommon.myCstr(gvAC.Rows(6).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name7 = clsCommon.myCstr(gvAC.Rows(6).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt7 = clsCommon.myCdbl(gvAC.Rows(6).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 7) Then
                    If clsCommon.myLen(gvAC.Rows(7).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code8 = clsCommon.myCstr(gvAC.Rows(7).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name8 = clsCommon.myCstr(gvAC.Rows(7).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt8 = clsCommon.myCdbl(gvAC.Rows(7).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 8) Then
                    If clsCommon.myLen(gvAC.Rows(8).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code9 = clsCommon.myCstr(gvAC.Rows(8).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name9 = clsCommon.myCstr(gvAC.Rows(8).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt9 = clsCommon.myCdbl(gvAC.Rows(8).Cells(colAChgAmount).Value)
                    End If
                End If
                If (gvAC.Rows.Count > 9) Then
                    If clsCommon.myLen(gvAC.Rows(9).Cells(colAChgCode).Value) > 0 Then
                        obj.Add_Charge_Code10 = clsCommon.myCstr(gvAC.Rows(9).Cells(colAChgCode).Value)
                        obj.Add_Charge_Name10 = clsCommon.myCstr(gvAC.Rows(9).Cells(colAChgName).Value)
                        obj.Add_Charge_Amt10 = clsCommon.myCdbl(gvAC.Rows(9).Cells(colAChgAmount).Value)
                    End If
                End If
                obj.Total_Add_Charge = clsCommon.myCdbl(lblAddCharges.Text)
                ''richa agarwal 14/10/2014
                obj.RoundOffAmount = clsCommon.myCdbl(lblRoundoffAmount.Text)
                obj.TotalSubsidyAmt = clsCommon.myCdbl(lblTotalSubsidyAmt.Text)
                ''===========================

                obj.Is_Against_Security_Receipt = chkAgainstSecurityReceiptNo.Checked
                obj.Against_Security_Receipt_No = txtSercurityReceiptNo.Value

                obj.Arr = New List(Of clsCustomerInvoiceDetail)
                For Each grow As GridViewRowInfo In gv1.Rows
                    Dim objTr As New clsCustomerInvoiceDetail()
                    objTr.SNo = clsCommon.myCdbl(grow.Cells(colLineNo).Value)
                    If chkAgainstServiceInvoice.Checked = True Then
                        objTr.AddChargeCode = clsCommon.myCstr(grow.Cells(colAddChgCode).Value)
                        objTr.AddChargeDesc = clsCommon.myCstr(grow.Cells(colAddChgName).Value)
                    End If
                    objTr.GL_Account_Code = clsCommon.myCstr(grow.Cells(colACCode).Value)
                    objTr.GL_Account_Desc = clsCommon.myCstr(grow.Cells(colACName).Value)
                    objTr.Amount = clsCommon.myCdbl(grow.Cells(colAmt).Value)
                    objTr.Discount_Per = clsCommon.myCdbl(grow.Cells(colDisPer).Value)
                    objTr.Discount = clsCommon.myCdbl(grow.Cells(colDisAmt).Value)
                    objTr.Amount_less_Discount = clsCommon.myCdbl(grow.Cells(colAmtAfterDis).Value)
                    objTr.TAX1 = clsCommon.myCstr(grow.Cells(colTax1).Value)
                    objTr.TAX1_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate1).Value)
                    objTr.TAX1_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt1).Value)
                    objTr.TAX2 = clsCommon.myCstr(grow.Cells(colTax2).Value)
                    objTr.TAX2_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate2).Value)
                    objTr.TAX2_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt2).Value)
                    objTr.TAX3 = clsCommon.myCstr(grow.Cells(colTax3).Value)
                    objTr.TAX3_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate3).Value)
                    objTr.TAX3_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt3).Value)
                    objTr.TAX4 = clsCommon.myCstr(grow.Cells(colTax4).Value)
                    objTr.TAX4_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate4).Value)
                    objTr.TAX4_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt4).Value)
                    objTr.TAX5 = clsCommon.myCstr(grow.Cells(colTax5).Value)
                    objTr.TAX5_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate5).Value)
                    objTr.TAX5_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt5).Value)
                    objTr.TAX6 = clsCommon.myCstr(grow.Cells(colTax6).Value)
                    objTr.TAX6_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate6).Value)
                    objTr.TAX6_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt6).Value)
                    objTr.TAX7 = clsCommon.myCstr(grow.Cells(colTax7).Value)
                    objTr.TAX7_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate7).Value)
                    objTr.TAX7_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt7).Value)
                    objTr.TAX8 = clsCommon.myCstr(grow.Cells(colTax8).Value)
                    objTr.TAX8_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate8).Value)
                    objTr.TAX8_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt8).Value)
                    objTr.TAX9 = clsCommon.myCstr(grow.Cells(colTax9).Value)
                    objTr.TAX9_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate9).Value)
                    objTr.TAX9_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt9).Value)
                    objTr.TAX10 = clsCommon.myCstr(grow.Cells(colTax10).Value)
                    objTr.TAX10_Rate = clsCommon.myCdbl(grow.Cells(colTaxRate10).Value)
                    objTr.TAX10_Amt = clsCommon.myCdbl(grow.Cells(colTaxAmt10).Value)
                    objTr.Total_Tax = clsCommon.myCdbl(grow.Cells(colTotTaxAmt).Value)
                    objTr.Total_Amount = clsCommon.myCdbl(grow.Cells(colAmtAfterTax).Value)

                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt1).Value)
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt2).Value)
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt3).Value)
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt4).Value)
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt5).Value)
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt6).Value)
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt7).Value)
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt8).Value)
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt9).Value)
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(grow.Cells(colTaxBaseAmt10).Value)

                    objTr.Hirerachy_Code = clsCommon.myCstr(grow.Cells(colHierarchyCode).Value)
                    objTr.Cost_Centre_Code = clsCommon.myCstr(grow.Cells(colCostCenterCode).Value)
                    objTr.Hirerachy_Code1 = clsCommon.myCstr(grow.Cells(colHierarchyLevel1).Value)
                    objTr.Hirerachy_Code2 = clsCommon.myCstr(grow.Cells(colHierarchyLevel2).Value)
                    objTr.Hirerachy_Code3 = clsCommon.myCstr(grow.Cells(colHierarchyLevel3).Value)
                    objTr.Hirerachy_Code4 = clsCommon.myCstr(grow.Cells(colHierarchyLevel4).Value)

                    If (clsCommon.myLen(objTr.GL_Account_Code) > 0) And objTr.Amount <> 0 Then
                        obj.Arr.Add(objTr)
                    End If
                Next
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please Fill atleast one GL Account with Amount.", Me.Text)
                    Return
                End If

                '' CurrencConversion
                If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
                    obj.CURRENCY_CODE = Me.txtCurrencyCode.Value
                    obj.ConvRate = clsCommon.myCdbl(Me.txtConversionRate.Text)
                    If clsCommon.myLen(txtApplicableFrom.Text) > 0 Then
                        obj.ApplicableFrom = Me.txtApplicableFrom.Text
                    Else
                        obj.ApplicableFrom = Nothing
                    End If
                Else
                    obj.CURRENCY_CODE = Nothing
                    obj.ConvRate = 1
                    obj.ApplicableFrom = Nothing
                End If
                '' end CurrencyConversion

                ''For Custom Fields
                obj.Form_ID = MyBase.Form_ID
                obj.arrCustomFields = New List(Of clsCustomFieldValues)
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.GetData(obj.arrCustomFields)
                End If
                If MyBase.ArrDetailFields IsNot Nothing AndAlso MyBase.ArrDetailFields.Count > 0 Then
                    clsCustomFieldGrid.GetData(obj.arrCustomFields, gv1, MyBase.ArrDetailFields, colACCode)
                End If
                ''End of For Custom Fields

                If (obj.SaveData(obj, isNewEntry, "", True)) Then
                    UcAttachment1.SaveData(obj.Document_No)
                    txtDocNo.Value = obj.Document_No
                    If clicked = False Then
                        common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                    End If

                    ''approval work 11/02/2020
                    Dim xNewDesc As String = ""
                    xNewDesc = "Party Name : " + obj.Customer_Name
                    ''=====================capex cond==============

                    xNewDesc = xNewDesc + Environment.NewLine + "Description : " + obj.Description
                    clsApply_Approval.CheckApprovalRequired(MyBase.Form_ID, obj.Document_No, txtDate.Value, clsCommon.myCstr(xNewDesc), clsCommon.myCstr(txtDesc.Text), clsCommon.myCdbl(lblTotRAmt1.Text), 0, "")
                    '================================================================

                    LoadData(obj.Document_No)
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadData(ByVal strDocumentNo As String)
        Try
            txtlocation.Enabled = False
            btnSave.Enabled = True
            btnPost.Enabled = True
            btnDelete.Enabled = True
            cboDocType.Enabled = False
            isInsideLoadData = True
            isNewEntry = False
            btnSave.Text = "Update"
            BlankAllControls()
            LoadBlankGridGL()
            LoadBlankGridTax()
            LoadBlankGridAC()
            Dim obj As New clsCustomerInvoiceHead()
            obj = clsCustomerInvoiceHead.GetData(strDocumentNo)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
                If obj.Status = ERPTransactionStatus.Approved Then
                    btnSave.Enabled = False
                    btnPost.Enabled = False
                    btnDelete.Enabled = False
                    If clsCommon.CompairString(obj.AgainstServiceInvoice, "Y") = CompairStringResult.Equal Then
                        btnCancel.Enabled = True
                    Else
                        btnCancel.Enabled = False
                    End If
                    If CostCenterAndHirerachyCodeUpdateAfterPost = True Then
                        butCostCenterAndHirerachy_Update_AfterPost.Visible = True
                    End If
                Else
                    btnCancel.Enabled = False
                    butCostCenterAndHirerachy_Update_AfterPost.Visible = False
                End If

                Dim qry As String = "select IRN_No,Ack_No,Ack_Date,QR_Code from TSPL_Customer_Invoice_Head where Document_No = '" + obj.Document_No + "'"
                Dim dtInv As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dtInv IsNot Nothing AndAlso dtInv.Rows.Count > 0 Then
                    EInvoiceIRNNo.Text = clsCommon.myCstr(dtInv.Rows(0)("IRN_no"))
                    EinvoiceAckNo.Text = clsCommon.myCstr(dtInv.Rows(0)("Ack_No"))
                    If dtInv.Rows(0)("Ack_Date") IsNot DBNull.Value Then
                        txtAckDate.Value = clsCommon.myCDate(dtInv.Rows(0)("Ack_Date"))
                    End If
                    EInvoiceQrCode.Text = clsCommon.myCstr(dtInv.Rows(0)("QR_Code"))
                    EinvoiceBtnUpdate.Enabled = False
                Else
                    EinvoiceBtnUpdate.Enabled = True
                End If
                UsLock1.Status = obj.Status
                txtDocNo.Value = obj.Document_No
                txtDate.Value = obj.Document_Date
                'txtDate.Enabled = False
                TxtCustomer.Value = obj.Customer_Code
                '------added by usha--
                txtlocation.Value = obj.loc_code
                '' Anubhooti 28-Nov-2014 BM00000004810
                If clsCommon.myLen(clsCommon.myCstr(obj.loc_code)) > 0 Then
                    LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(obj.loc_code) & "'"))
                Else
                    LblLocDesp.Text = ""
                End If
                '-----------
                txtLocationPrefix.Value = obj.Location_Code_Prefix
                If clsCommon.myLen(clsCommon.myCstr(obj.Location_Code_Prefix)) > 0 Then
                    txtLocationPrefixName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Location_Desc  from TSPL_LOCATION_MASTER WHERE  Location_Code='" & txtLocationPrefix.Value & "' "))
                Else
                    txtLocationPrefixName.Text = ""
                End If
                lblVendorName.Text = obj.Customer_Name
                txtACSet.Value = obj.Account_Set
                cboDocType.SelectedValue = obj.Document_Type
                txtRefDocNo.Value = obj.RefDocNo
                If ((clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal)) Then
                    txtSaleInvoice.Text = obj.Against_Sale_No
                    txtServiceVisitCode.Text = obj.Against_Service_Visit_Code
                Else
                    txtServiceVisitCode.Text = ""
                End If
                If ((clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal)) Then
                    txtRefDocNo.Visible = True
                    RadLabel15.Visible = True
                    '' Anubhooti 18-Mar-2015 (Show Against VCGL)
                    txtVCGNo.Text = clsCommon.myCstr(obj.Against_VCGL)
                Else
                    txtRefDocNo.Visible = False
                    RadLabel15.Visible = False
                    txtVCGNo.Text = ""
                End If
                chkSecurityDposit.Checked = obj.SecurityDeposit
                ddlSecDepositType.SelectedValue = clsCommon.myCstr(obj.SecurityDepositType)
                If chkSecurityDposit.Checked AndAlso clsCommon.myLen(ddlSecDepositType.SelectedValue) > 0 Then
                    gv1.Columns(colACCode).ReadOnly = True
                End If
                chkAgainstServiceInvoice.Checked = IIf(obj.AgainstServiceInvoice = "Y", True, False)
                If chkAgainstServiceInvoice.Checked = True Then
                    btnPrintServiceInvoice.Visible = True
                Else
                    btnPrintServiceInvoice.Visible = False
                End If
                If (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                    Dim strReturnNo, strInvoiceNo As String
                    strReturnNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Against_Sale_Return_No from TSPL_Customer_Invoice_Head where Document_No='" & txtDocNo.Value & "'"))
                    strInvoiceNo = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Against_Invoice_No from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & strReturnNo & "'"))
                    txtSaleInvoice.Text = strInvoiceNo
                    txtSaleReturnNo.Text = strReturnNo
                    lblReturnType.Visible = True
                    ddlReturnType.Visible = True
                    ddlReturnType.SelectedValue = obj.Return_Type
                Else
                    lblReturnType.Visible = False
                    ddlReturnType.Visible = False
                End If

                txtOrderNo.Text = obj.Order_No

                chkOnHold.Checked = obj.On_Hold
                chkEinvoice.Checked = IIf(obj.Is_Einvoice = 1, True, False)
                '=================================shivani
                If clsCommon.myLen(obj.DateAndTime) > 0 Then
                    txtDataAndTimeSelection.Value = obj.DateAndTime
                    txtDataAndTimeSelection.Checked = True
                End If
                txtTapalNo.Text = clsCommon.myCstr(obj.TapalNo)
                Dim t1 As New clsVCGLHead()
                t1 = clsVCGLHead.GetData(obj.Against_VCGL)
                If obj.Against_VCGL = "" Then
                    If clsCommon.myCstr(clsCommon.myLen(obj.Description)) > 0 Then
                        txtDesc.Text = obj.Description
                    End If
                Else

                    If t1.Document_Type = "C" And obj.Customer_Code = t1.VC_Code Then
                        txtDesc.Text = t1.Remarks

                    End If
                    For Each Tr1 As clsVCGLDetail In t1.Arr
                        If Tr1.Row_Type = "Customer" And obj.Customer_Code = Tr1.VCGL_Code Then
                            txtDesc.Text = Tr1.Remarks
                        End If
                    Next
                End If
                '=========================================================
                txtTaxGroup.Value = obj.Tax_Group

                Dim objTaxGrpMaster As New clsTaxGroupMaster()
                objTaxGrpMaster = objTaxGrpMaster.GetDataForPurchase(obj.Tax_Group)
                If (objTaxGrpMaster IsNot Nothing) Then
                    lblTaxGrpName.Text = objTaxGrpMaster.Tax_Group_Desc
                End If
                txtTaxGroup.Value = obj.Tax_Group
                txtTermCode.Value = obj.Terms_Code
                lblTermName.Text = obj.Terms_Description
                txtDueDate.Value = obj.Due_Date
                lblAmtWithDiscount.Text = clsCommon.myFormat(obj.Discount_Base)
                lblDiscountAmt.Text = clsCommon.myFormat(obj.Discount_Amount)
                lblAmtAfterDiscount.Text = clsCommon.myFormat(obj.Amount_Less_Discount)
                lblTaxAmt.Text = clsCommon.myFormat(obj.Total_Tax)
                lblTotRAmt.Text = clsCommon.myFormat(obj.Document_Total)
                lblTotRAmt1.Text = clsCommon.myFormat(obj.Document_Total)

                If obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic Then
                    rbtnTaxCalAutomatic.IsChecked = True
                ElseIf obj.Tax_Calculation_Type = EnumTaxCalucationType.Mannual Then
                    rbtnTaxCalManual.IsChecked = True
                End If
                fndProject.Value = obj.PROJECT_ID
                lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")

                FlagDocumentIsTaxable = CheckIsTaxable(obj.Document_No)
                EInvoiceType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_Customer_Invoice_Head", "Document_No", obj.Document_No, Nothing)

                If (clsCommon.myLen(obj.TAX1) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX1
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX1_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax1_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX1_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX1) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX2) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX2
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX2_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax2_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX2_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX2) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX3) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX3
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX3_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax3_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX3_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX3) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX4) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX4
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX4_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax4_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX4_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX4) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX5) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX5
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX5_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax5_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX5_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX5) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX6) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX6
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX6_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax6_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX6_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX6) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX7) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX7
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX7_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax7_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX7_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX7) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX8) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX8
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX8_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax8_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX8_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX8) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX9) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX9
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX9_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax9_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX9_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX9) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If
                If (clsCommon.myLen(obj.TAX10) > 0) Then
                    gv2.Rows.AddNew()
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutCode).Value = obj.TAX10
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxRate).Value = obj.TAX10_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTBaseAmt).Value = obj.Tax10_BAmount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAmt).Value = obj.TAX10_Amt
                    If (objTaxGrpMaster IsNot Nothing AndAlso objTaxGrpMaster.Arr IsNot Nothing AndAlso objTaxGrpMaster.Arr.Count > 0) Then
                        For Each objTaxGrpTr As clsTaxGroupDetail In objTaxGrpMaster.Arr
                            If (clsCommon.CompairString(objTaxGrpTr.Tax_Code, obj.TAX10) = CompairStringResult.Equal) Then
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTTaxAutName).Value = objTaxGrpTr.Tax_Code_Desc
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsSurTax).Value = objTaxGrpTr.Surtax
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTIsTaxable).Value = objTaxGrpTr.Taxable
                                gv2.Rows(gv2.Rows.Count - 1).Cells(colTSurTaxCode).Value = objTaxGrpTr.Surtax_Tax_Code
                                Exit For
                            End If
                        Next
                    End If
                End If


                If (clsCommon.myLen(obj.Add_Charge_Code1) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code1
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name1
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt1
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code2) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code2
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name2
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt2
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code3) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code3
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name3
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt3
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code4) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code4
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name4
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt4
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code5) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code5
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name5
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt5
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code6) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code6
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name6
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt6
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code7) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code7
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name7
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt7
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code8) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code8
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name8
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt8
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code9) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code9
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name9
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt9
                End If
                If (clsCommon.myLen(obj.Add_Charge_Code10) > 0) Then
                    gvAC.Rows.AddNew()
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgCode).Value = obj.Add_Charge_Code10
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgName).Value = obj.Add_Charge_Name10
                    gvAC.Rows(gvAC.Rows.Count - 1).Cells(colAChgAmount).Value = obj.Add_Charge_Amt10
                End If

                lblAddCharges.Text = clsCommon.myFormat(obj.Total_Add_Charge)
                lblAddCharges1.Text = clsCommon.myFormat(obj.Total_Add_Charge)

                ''richa agarwal 14/10/2014
                lblRoundoffAmount.Text = clsCommon.myFormat(obj.RoundOffAmount)
                lblTotalSubsidyAmt.Text = clsCommon.myFormat(obj.TotalSubsidyAmt)
                '========================

                chkAgainstSecurityReceiptNo.Checked = obj.Is_Against_Security_Receipt
                txtSercurityReceiptNo.Value = obj.Against_Security_Receipt_No
                For Each objTr As clsCustomerInvoiceDetail In obj.Arr
                    gv1.Rows.AddNew()
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = objTr.SNo
                    ''richa agarwal 11 Oct,2019
                    If chkAgainstServiceInvoice.Checked = True Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAddChgCode).Value = objTr.AddChargeCode
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colAddChgName).Value = objTr.AddChargeDesc
                        gv1.Columns(colAddChgCode).IsVisible = True
                        gv1.Columns(colAddChgName).IsVisible = True
                        gv1.Columns(colSACCode).IsVisible = True
                        gv1.Columns(colSACName).IsVisible = True
                        If clsCommon.myLen(objTr.AddChargeCode) > 0 Then
                            gv1.CurrentRow.Cells(colSACCode).Value = clsAdditionalCharge.GetSACCode(objTr.AddChargeCode, Nothing)
                            gv1.CurrentRow.Cells(colSACName).Value = ClsSACMaster.GetName(clsCommon.myCstr(gv1.CurrentRow.Cells(colSACCode).Value))
                        End If
                    End If
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colACCode).Value = objTr.GL_Account_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colACName).Value = objTr.GL_Account_Desc
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colGLType).Value = clsPaymentHeader.CheckGLAccountType(objTr.GL_Account_Code, Nothing)
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAmt).Value = objTr.Amount
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisPer).Value = objTr.Discount_Per
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colDisAmt).Value = objTr.Discount
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterDis).Value = objTr.Amount_less_Discount
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax1).Value = objTr.TAX1
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate1).Value = objTr.TAX1_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt1).Value = objTr.TAX1_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax2).Value = objTr.TAX2
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate2).Value = objTr.TAX2_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt2).Value = objTr.TAX2_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax3).Value = objTr.TAX3
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate3).Value = objTr.TAX3_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt3).Value = objTr.TAX3_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax4).Value = objTr.TAX4
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate4).Value = objTr.TAX4_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt4).Value = objTr.TAX4_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax5).Value = objTr.TAX5
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate5).Value = objTr.TAX5_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt5).Value = objTr.TAX5_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax6).Value = objTr.TAX6
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate6).Value = objTr.TAX6_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt6).Value = objTr.TAX6_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax7).Value = objTr.TAX7
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate7).Value = objTr.TAX7_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt7).Value = objTr.TAX7_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax8).Value = objTr.TAX8
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate8).Value = objTr.TAX8_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt8).Value = objTr.TAX8_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax9).Value = objTr.TAX9
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate9).Value = objTr.TAX9_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt9).Value = objTr.TAX9_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTax10).Value = objTr.TAX10
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxRate10).Value = objTr.TAX10_Rate
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxAmt10).Value = objTr.TAX10_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTotTaxAmt).Value = objTr.Total_Tax
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colAmtAfterTax).Value = objTr.Total_Amount

                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt1).Value = objTr.TAX1_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt2).Value = objTr.TAX2_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt3).Value = objTr.TAX3_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt4).Value = objTr.TAX4_Base_Amt

                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt5).Value = objTr.TAX5_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt6).Value = objTr.TAX6_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt7).Value = objTr.TAX7_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt8).Value = objTr.TAX8_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt9).Value = objTr.TAX9_Base_Amt
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTaxBaseAmt10).Value = objTr.TAX10_Base_Amt
                    ''richa


                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHierarchyCode).Value = objTr.Hirerachy_Code
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCostCenterCode).Value = objTr.Cost_Centre_Code
                    If clsCommon.myLen(objTr.Hirerachy_Code) > 0 Then
                        gv1.Rows(gv1.Rows.Count - 1).Cells(colHierarchyLevelNumber).Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Level from TSPL_HIRERACHY_LEVEL_MASTER where Hirerachy_Code='" + objTr.Hirerachy_Code + "'"))
                    End If
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHierarchyLevel1).Value = objTr.Hirerachy_Code1
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHierarchyLevel2).Value = objTr.Hirerachy_Code2
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHierarchyLevel3).Value = objTr.Hirerachy_Code3
                    gv1.Rows(gv1.Rows.Count - 1).Cells(colHierarchyLevel4).Value = objTr.Hirerachy_Code4

                Next
                If obj.Status = ERPTransactionStatus.Pending Then
                    gv1.Rows.AddNew()
                    gvAC.Rows.AddNew()
                End If

                '' MULTICURRENCY
                Me.txtCurrencyCode.Value = obj.CURRENCY_CODE
                Me.txtConversionRate.Text = obj.ConvRate
                Me.txtApplicableFrom.Text = obj.ApplicableFrom.ToString
                '' end  MULTICURRENCY
                '=====================if document go for approval then no post button visible or if document contain related setting
                btnPost.Visible = MyBase.isPostFlag
                If Not clsApply_Approval.Visibility_PostButtonForApproval(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value), clsCommon.myCdbl(lblTotRAmt1.Text), 0, "") Then
                    btnPost.Visible = False
                    If UsLock1.Status = ERPTransactionStatus.Pending Then
                        UsLock1.Status = clsApply_Approval.ApprovalCondCheck_Doc(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value), Nothing)
                    End If
                End If
                '============================================
                If chkAgainstServiceInvoice.Checked = True AndAlso FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                    btnReverse.Enabled = False
                    If obj.Status = ERPTransactionStatus.Approved Then
                        btnCancel.Enabled = True
                    ElseIf obj.Status = ERPTransactionStatus.Pending Then
                        btnCancel.Enabled = False
                    End If
                End If
                ''For Custom Fields
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.LoadData(obj.Document_No)
                End If
                clsCustomFieldGrid.FillDataInGrid(obj.Document_No, MyBase.Form_ID, gv1)
                ''End of For Custom Fields
                UcAttachment1.LoadData(obj.Document_No)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
        End Try
    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        CloseForm()
    End Sub

    Sub CloseForm()
        Me.Close()
    End Sub

    Private Sub gv2_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv2.DoubleClick
        Try
            If rbtnTaxCalAutomatic.IsChecked Then
                Dim qry As String = "select Tax_Rate_Code as [Rate Code],Tax_Rate_Desc as [Rate Description],Tax_Rate as [Rate] from TSPL_TAX_RATES "
                Dim dblNewRate As Double = clsCommon.myCdbl(clsCommon.ShowSelectForm("FndVendorTaxRate", qry, "Rate", "Tax_Code='" + clsCommon.myCstr(gv2.CurrentRow.Cells(colTTaxAutCode).Value) + "' and Tax_Type='S'", "", "", True))
                gv2.CurrentRow.Cells(colTTaxRate).Value = dblNewRate
                For ii As Integer = 0 To gv1.Rows.Count - 1
                    UpdateCurrentRowWihtRowNo(ii)
                Next
                UpdateAllTotals()
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub btnPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPost.Click
        clicked = True
        PostData()
    End Sub

    Sub PostData()
        Try
            Dim msg As String = ""
            Dim qry As String = ""
            Dim dt As DataTable = Nothing

            If (myMessages.postConfirm()) Then
                SaveData(clicked)
                If (clsCustomerInvoiceHead.PostData(MyBase.Form_ID, txtDocNo.Value, txtRefDocNo.Value)) Then
                    msg = "Successfully Posted"
                Else
                    qry = "select No_Of_Level, LEVEL from TSPL_APPROVAL_LEVEL_SCREEN where User_Code='" + objCommonVar.CurrentUserCode + "' and Trans_Code='" + MyBase.Form_ID + "' "
                    dt = clsDBFuncationality.GetDataTable(qry)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        Dim level As String = dt.Rows(0)("LEVEL").ToString()
                        Dim NoOflevel As Integer = clsCommon.myCdbl(dt.Rows(0)("No_Of_Level"))
                        If clsCommon.CompairString(level, "Level1") = CompairStringResult.Equal Then
                            msg = "Level 1 Approval done. "
                            If NoOflevel = 1 Then
                                msg += "Successfully Posted. "
                            Else
                                msg += "Level 2 Approval Required."
                            End If
                        ElseIf clsCommon.CompairString(level, "Level2") = CompairStringResult.Equal Then
                            msg = "Level 2 Approval done. "
                            If NoOflevel = 2 Then
                                msg += "Successfully Posted "
                            Else
                                msg += "Level 3 Approval Required."
                            End If
                        Else
                            msg = "Level 3 Approval done. Successfully Posted"
                        End If
                    End If
                End If
                common.clsCommon.MyMessageBoxShow(Me, msg, Me.Text)
                LoadData(txtDocNo.Value)
            End If

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        clicked = False
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
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
                clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtDocNo.Value))
                If (clsCustomerInvoiceHead.DeleteData(txtDocNo.Value)) Then
                    saveCancelLog(Reason, "Delete", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    AddNew()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
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

    Private Sub gv1_CellEditorInitialized(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellEditorInitialized
        If TypeOf Me.gv1.CurrentColumn Is GridViewMultiComboBoxColumn Then
            Dim editor As RadMultiColumnComboBoxElement = DirectCast(Me.gv1.ActiveEditor, RadMultiColumnComboBoxElement)
            editor.AutoSizeDropDownToBestFit = True
            editor.EditorControl.MasterTemplate.BestFitColumns()
            editor.DropDownStyle = RadDropDownStyle.DropDown
            editor.AutoFilter = True
            If editor.EditorControl.MasterTemplate.FilterDescriptors.Count = 0 Then
                Dim autoFilter As FilterDescriptor = New FilterDescriptor("Name", FilterOperator.StartsWith, "")
                autoFilter.IsFilterEditor = True
                editor.EditorControl.FilterDescriptors.Add(autoFilter)
            End If
        End If
    End Sub

    Private Sub txtDocNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocNo._MYNavigator
        Try
            Dim whrclas As String = ""
            If clsCommon.myLen(objCommonVar.strCurrUserGLAccount) > 0 Then
                whrclas += " and TSPL_Customer_Invoice_Detail.GL_Account_Code in(" + objCommonVar.strCurrUserGLAccount + ")"
            End If
            If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                'If clsCommon.myLen(objCommonVar.strCurrUserCustomers) > 0 Then
                '    whrclas += " and TSPL_Customer_Invoice_Head.Customer_Code in (" + objCommonVar.strCurrUserCustomers + ") "
                'End If
                whrclas += " and TSPL_Customer_Invoice_Head.Loc_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
            End If
            Dim qry As String = "select Document_No from TSPL_Customer_Invoice_Head where Document_No="
            Select Case NavType
                Case NavigatorType.First
                    qry += "(select MIN(TSPL_Customer_Invoice_Head.Document_No) from TSPL_Customer_Invoice_Head left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No=TSPL_Customer_Invoice_Head.Document_No and TSPL_Customer_Invoice_Detail.SNo=1 where 2=2  " + whrclas + ")"
                Case NavigatorType.Last
                    qry += "(select Max(TSPL_Customer_Invoice_Head.Document_No) from TSPL_Customer_Invoice_Head left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No=TSPL_Customer_Invoice_Head.Document_No and TSPL_Customer_Invoice_Detail.SNo=1 where 2=2  " + whrclas + " )"
                Case NavigatorType.Next
                    qry += "(select Min(TSPL_Customer_Invoice_Head.Document_No) from TSPL_Customer_Invoice_Head left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No=TSPL_Customer_Invoice_Head.Document_No and TSPL_Customer_Invoice_Detail.SNo=1 where 2=2  " + whrclas + " and TSPL_Customer_Invoice_Head.Document_No>'" + txtDocNo.Value + "')"
                Case NavigatorType.Previous
                    qry += "(select Max(TSPL_Customer_Invoice_Head.Document_No) from TSPL_Customer_Invoice_Head left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No=TSPL_Customer_Invoice_Head.Document_No and TSPL_Customer_Invoice_Detail.SNo=1 where 2=2  " + whrclas + " and TSPL_Customer_Invoice_Head.Document_No<'" + txtDocNo.Value + "')"
            End Select
            LoadData(clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry)))
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    '============BM00000008196==================
    Private Sub txtDocNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtDocNo._MYValidating
        Dim qry As String = "select TSPL_Customer_Invoice_Head.Document_No as DocumentNo,CONVERT(varchar(10), Document_Date,103)+' '+ CONVERT(varchar(5), Document_Date,114) as Date,TSPL_Customer_Invoice_Head.Customer_Code as [Customer Code],TSPL_Customer_Invoice_Head.Customer_Name as [Customer Name],ISNULL(TSPL_CUSTOMER_MASTER.Alies_Name,'') As [Alies Name],case when TSPL_Customer_Invoice_Head.[status]=1 then 'Posted' else 'Pending'end as [Status],Account_Set as AccountSet,Against_Sale_No as [Against Sale Invoice],AgainstScrap as [Against Scrap No],case when coalesce(Against_Sale_Return_No,'')='' then coalesce(Against_MCC_Material_Sale_Return,'') else coalesce(Against_Sale_Return_No,'') end as [Against Sale Return],ISNULL(TSPL_Customer_Invoice_Head.Against_VCGL,'') AS [Against VCGL],Against_Asset_Disposal as [Against Asset Disposal],AgainstScrapReturn as [Against Scrap Return No],TSPL_Customer_Invoice_Head.Location_Code_Prefix as [Location Code Prefix] from TSPL_Customer_Invoice_Head left outer join TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No=TSPL_Customer_Invoice_Head.Document_No and TSPL_Customer_Invoice_Detail.SNo=1 "
        '' Anubhooti 13-Mar-2015 (Fetch Alies Name On Customer Finder)
        qry += " LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_Customer_Invoice_Head.Customer_Code "
        Dim whrclas As String = "1=1"
        If clsCommon.myLen(objCommonVar.strCurrUserGLAccount) > 0 Then
            whrclas += " and TSPL_Customer_Invoice_Detail.GL_Account_Code in(" + objCommonVar.strCurrUserGLAccount + ")"
        End If
        If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
            'If clsCommon.myLen(objCommonVar.strCurrUserCustomers) > 0 Then
            '    whrclas += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" + objCommonVar.strCurrUserCustomers + ") "
            'End If
            whrclas += " and TSPL_Customer_Invoice_Head.Loc_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
        End If
        LoadData(clsCommon.ShowSelectForm("fmAPInvoiceSelector", qry, "DocumentNo", whrclas, txtDocNo.Value, "Document_Date desc", isButtonClicked, "TSPL_Customer_Invoice_Head.Document_Date"))
    End Sub

    Private Sub txtDocNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDocNo.KeyPress
        If (e.KeyChar = Chr(39)) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        'PrintData()
        funPrint(txtDocNo.Value)
    End Sub

    Sub PrintData()
        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
            common.clsCommon.MyMessageBoxShow(Me, "Document No not found to print", Me.Text)
        End If
        Dim Arr As New ArrayList
        Arr.Add(txtDocNo.Value)
        funPrint(txtDocNo.Value)
        'frmRptAPInvoice.PrintData("", "", True, Arr, False, Nothing, False, Nothing)
    End Sub

    ' Modified By Prabhakar ; Ticket Ref : BM00000009469
    Sub funPrint(ByVal StrCode As String)
        Try
            ' Dim qry As String = "select RefDocNo, tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code,TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo ,TSPL_CUSTOMER_MASTER.GSTNO AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode,TSPL_STATE_MASTER.state_name AS Vendor_State_Name, Loc_Code, locAdd, XXX.Location_Desc,xxx.Tin_No,XXX.PAN,TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end  as CompAdd , " & _
            ' "XXX.Description,XXX.Account_Code,XXX.Account_Desc ,XXX.DrAmt ,XXX.CrAmt ,XXX.Document_No ,XXX.Document_Date,XXX.Status , " & _
            ' "XXX.Document_Type ,XXX.Account_Set ,XXX.DocAmt,XXX.Customer_Code ,XXX.Customer_Name ,XXX.Created_By ,XXX.Modify_By ,XXX.Detail_Line_No , " & _
            ' "XXX .Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2,XXX.Cost_Centre_Code,XXX.Cost_Center_Fin_Name,XXX.Hirerachy_Code,XXX.HIRERACHY_Name  from " & _
            ' "(select distinct Loc_Code,RefDocNo, locAdd,Location_Desc,Tin_No,PAN,(final.Account_Code),final.Account_Desc ,final.DrAmt ,final.CrAmt ,final.Document_No ,final.Document_Date, " & _
            ' "final.Status ,final.Document_Type ,final.Account_Set ,final.DocAmt,final.Customer_Code ,final.Customer_Name ,final.Created_By , " & _
            ' "final.Modify_By ,final.Detail_Line_No ,final .Comp_Code,Description ,final.Cost_Centre_Code,final.Cost_Center_Fin_Name,final.Hirerachy_Code,final.HIRERACHY_Name from " & _
            ' "(select isnull(TSPL_Customer_Invoice_Head.RefDocNo,'') as RefDocNo, TSPL_Customer_Invoice_Head.Loc_Code," & _
            ' " TSPL_LOCATION_MASTER.Add1 + case When TSPL_LOCATION_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_LOCATION_MASTER.Add2, 103) End + Case When TSPL_LOCATION_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add3,103) end + case When TSPL_LOCATION_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.City_Code, 103) end+ Case When TSPL_LOCATION_MASTER.State='' Then '' else ', '+Convert(Varchar, loc_state.STATE_NAME ) end +  Case When TSPL_LOCATION_MASTER.Pin_code='' Then '' Else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.Pin_code, 103)  end  as locAdd, Location_Desc,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.PAN,TSPL_Customer_Invoice_Head.Description,case  when TSPL_JOURNAL_DETAILS.Amount >=0 then TSPL_JOURNAL_DETAILS.Amount else 0 end as DrAmt , " & _
            ' "case  when TSPL_JOURNAL_DETAILS.Amount <0 then TSPL_JOURNAL_DETAILS.Amount*-1 else 0 end as CrAmt, " & _
            ' "TSPL_Customer_Invoice_Head.Document_No, Document_Date , " & _
            ' "case when TSPL_Customer_Invoice_Head.Status=1 then 'Authorized' else 'UnAuthorized' end as Status , " & _
            ' "case when TSPL_Customer_Invoice_Head.Document_Type='I' then 'Invoice' else case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'Debit Note' else " & _
            ' "case when TSPL_Customer_Invoice_Head.Document_Type='C' then 'Credit Note' else '' end end end as Document_Type,Account_Set,Document_Total as DocAmt,  " & _
            ' "Customer_Code ,TSPL_Customer_Invoice_Head.Customer_Name,TSPL_Customer_Invoice_Head.Created_By ,TSPL_Customer_Invoice_Head.Modify_By  , " & _
            ' "TSPL_JOURNAL_DETAILS.Detail_Line_No as Detail_Line_No ,TSPL_JOURNAL_DETAILS.Account_code as Account_Code , " & _
            ' "TSPL_JOURNAL_DETAILS.Account_Desc as Account_Desc ,TSPL_Customer_Invoice_Detail.Amount , " & _
            ' "TSPL_Customer_Invoice_Detail.Discount ,TSPL_Customer_Invoice_Detail.Amount_less_Discount , " & _
            ' "TSPL_Customer_Invoice_Detail .Total_Tax ,TSPL_Customer_Invoice_Detail.Total_Amount  , " & _
            ' "TSPL_Customer_Invoice_Head.Comp_Code ,TSPL_Customer_Invoice_Detail.Cost_Centre_Code,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name,TSPL_Customer_Invoice_Detail.Hirerachy_Code,TSPL_HIRERACHY_LEVEL_MASTER.Description as HIRERACHY_Name   from TSPL_Customer_Invoice_Head left outer join " & _
            ' "TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No =TSPL_Customer_Invoice_Head.Document_No " & _
            ' "left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =TSPL_Customer_Invoice_Head.Document_No " & _
            ' "left outer join TSPL_JOURNAL_DETAILS on TSPL_JOURNAL_DETAILS.Voucher_No = TSPL_JOURNAL_MASTER.Voucher_No   " & _
            ' " and isnull(TSPL_JOURNAL_DETAILS.Cost_Centre_Code,'')=isnull(TSPL_Customer_Invoice_Detail.Cost_Centre_Code,'') " & _
            ' " and isnull(TSPL_JOURNAL_DETAILS.Hirerachy_Code,'')=isnull(TSPL_Customer_Invoice_Detail.Hirerachy_Code,'') " & _
            ' "left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_Customer_Invoice_Head.Customer_Code  " & _
            ' "left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_Customer_Invoice_Head.Loc_Code  " & _
            ' " left outer join TSPL_COST_CENTRE_FINANCIAL on TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code = TSPL_Customer_Invoice_Detail.Cost_Centre_Code " & _
            ' " left outer join TSPL_HIRERACHY_LEVEL_MASTER on TSPL_HIRERACHY_LEVEL_MASTER.Hirerachy_Code = TSPL_Customer_Invoice_Detail.Hirerachy_Code " & _
            '" left outer join TSPL_STATE_MASTER as loc_state on loc_state.STATE_CODE =TSPL_LOCATION_MASTER.State " & _
            ' "where TSPL_Customer_Invoice_Head.Document_No ='" + StrCode + "' )final   )XXX left outer join " & _
            ' "TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = XXX.Comp_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = XXX.Customer_Code left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code = XXX.Loc_Code left outer join tspl_state_master as tspl_state_master_for_location_state on  tspl_state_master_for_location_state.state_code=tspl_location_master.state    order by XXX.Detail_Line_No  "
            ''Sanjay Ticket no-ALF/05/06/19-000105, Created By and Modify By Name instead of Code
            '' Changes by Richa on 28/11/2017 

            '''''''''''''''''Check document is taxable --------------------
            'Dim IsTaxable As Decimal = 0
            'If clsCommon.myCdbl(lblTaxAmt.Text) > 0 Then
            '    If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
            '        Dim IsExempted = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Is_Tax_Exempted from TSPL_TAX_GROUP_MASTER WHERE TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' and Tax_Group_Code='" + txtTaxGroup.Value + "'"))
            '        If IsExempted = 1 Then
            '            IsTaxable = 0
            '        Else
            '            Dim TaxQuery As String = "select (case when isnull(tax1.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax1_amt else 0 end " &
            '            " + case when isnull(tax2.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax2_amt else 0 end " &
            '            " + case when isnull(tax3.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax3_amt else 0 end " &
            '            " + case when isnull(tax4.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax4_amt else 0 end " &
            '            " + case when isnull(tax5.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax5_amt else 0 end " &
            '            " + case when isnull(tax6.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax6_amt else 0 end " &
            '            " + case when isnull(tax7.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax7_amt else 0 end " &
            '            " + case when isnull(tax8.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax8_amt else 0 end " &
            '            " + case when isnull(tax9.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax9_amt else 0 end " &
            '            " + case when isnull(tax10.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax10_amt else 0 end) as aa " &
            '            " from TSPL_Customer_Invoice_Head " &
            '            " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_Customer_Invoice_Head.tax1  " &
            '             " left outer join tspl_tax_master As tax2 On tax2.tax_code =TSPL_Customer_Invoice_Head.tax2  " &
            '             " left outer join tspl_tax_master as tax3 on tax3.Tax_Code =TSPL_Customer_Invoice_Head .TAX3  " &
            '             " left outer join TSPL_TAX_MASTER As tax4 On tax4.Tax_Code =TSPL_Customer_Invoice_Head .tax4 " &
            '             " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code =TSPL_Customer_Invoice_Head .tax5 " &
            '             " left outer join TSPL_TAX_MASTER As tax6 On tax6.Tax_Code =TSPL_Customer_Invoice_Head .TAX6 " &
            '             " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_Customer_Invoice_Head .TAX7 " &
            '             " left outer join TSPL_TAX_MASTER As tax8 On tax8.Tax_Code =TSPL_Customer_Invoice_Head .TAX8 " &
            '             " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_Customer_Invoice_Head .TAX9 " &
            '             " left outer join TSPL_TAX_MASTER As tax10 On tax10.Tax_Code =TSPL_Customer_Invoice_Head .TAX10 " &
            '             " where TSPL_Customer_Invoice_Head.Document_No ='" + StrCode + "'"
            '            Dim TaxAmount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(TaxQuery))
            '            If TaxAmount > 0 Then
            '                IsTaxable = 1
            '            End If
            '        End If
            '    End If
            'End If
            '''''''''''''''''Check document is taxable --------------------
            'Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_Customer_Invoice_Head", "Document_No", StrCode, Nothing)
            Dim IsEInvoiceApply As Integer = 0
            If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                IsEInvoiceApply = 1
            End If
            clsReceiptInvoiceHead.funARInvoicePrint(MyBase.Form_ID, Nothing, txtDate.Value, txtDocNo.Value, IsEInvoiceApply, SettingCostCenterlevel)
            'Dim qry As String = "SELECT cast(TSPL_Customer_Invoice_Head.BarCode_Img as image) As BarCode_Img,isnull (TSPL_Customer_Invoice_Head.IRN_No,'') as IRN_No,isnull (TSPL_Customer_Invoice_Head.Ack_No,'') as Ack_No,case when len(isnull (TSPL_Customer_Invoice_Head.Ack_No,'')) > 0 then convert (varchar, TSPL_Customer_Invoice_Head.Ack_Date,103) else ''  end as Ack_Date, " + clsCommon.myCstr(IsEInvoiceApply) + " as  IsEInvoiceApply," &
            '  " XXX.RefDocNo, tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code,TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo ,TSPL_CUSTOMER_MASTER.GSTNO AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode,TSPL_STATE_MASTER.state_name AS Vendor_State_Name, XXX.Loc_Code, locAdd, XXX.Location_Desc,xxx.Tin_No,XXX.PAN,TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end  as CompAdd , " &
            '"XXX.Description,XXX.Account_Code,XXX.Account_Desc ,XXX.DrAmt ,XXX.CrAmt ,XXX.Document_No ,XXX.Document_Date,XXX.Status , " &
            '"XXX.Document_Type ,XXX.Account_Set ,XXX.DocAmt,XXX.Customer_Code ,XXX.Customer_Name ,XXX.Created_By ,XXX.Modify_By ,XXX.Detail_Line_No , " &
            '"XXX .Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2,XXX.Cost_Centre_Code,XXX.Cost_Center_Fin_Name,XXX.Hirerachy_Code,XXX.HIRERACHY_Name,XXX.Hirerachy_Code3,XXX.Hirerachy_Code4,XXX.TapalNo,XXX.DateAndTime   from " &
            '"(select distinct Loc_Code,RefDocNo, locAdd,Location_Desc,Tin_No,PAN,(final.Account_Code),final.Account_Desc ,final.DrAmt ,final.CrAmt ,final.Document_No ,final.Document_Date, " &
            '"final.Status ,final.Document_Type ,final.Account_Set ,final.DocAmt,final.Customer_Code ,final.Customer_Name ,final.Created_By , " &
            '"final.Modify_By ,final.Detail_Line_No ,final .Comp_Code,Description ,final.Cost_Centre_Code,final.Cost_Center_Fin_Name,final.Hirerachy_Code,final.HIRERACHY_Name,final.Hirerachy_Code3 ,final.Hirerachy_Code4,final.TapalNo,final.DateAndTime  from " &
            '"(select isnull(TSPL_Customer_Invoice_Head.RefDocNo,'') as RefDocNo, TSPL_Customer_Invoice_Head.Loc_Code," &
            '" TSPL_LOCATION_MASTER.Add1 + case When TSPL_LOCATION_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_LOCATION_MASTER.Add2, 103) End + Case When TSPL_LOCATION_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add3,103) end + case When TSPL_LOCATION_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.City_Code, 103) end+ Case When TSPL_LOCATION_MASTER.State='' Then '' else ', '+Convert(Varchar, loc_state.STATE_NAME ) end +  Case When TSPL_LOCATION_MASTER.Pin_code='' Then '' Else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.Pin_code, 103)  end  as locAdd," &
            '" Location_Desc,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.PAN,TSPL_Customer_Invoice_Head.Description,case  when TSPL_JOURNAL_DETAILS.Amount >=0 then TSPL_JOURNAL_DETAILS.Amount else 0 end as DrAmt ," &
            '" case  when TSPL_JOURNAL_DETAILS.Amount <0 then TSPL_JOURNAL_DETAILS.Amount*-1 else 0 end as CrAmt, TSPL_Customer_Invoice_Head.Document_No, Document_Date , case when TSPL_Customer_Invoice_Head.Status=1 then 'Authorized' else 'UnAuthorized' end as Status , " &
            '" case WHEN isnull(AgainstServiceInvoice,'')='Y' then 'Tax Invoice' when TSPL_Customer_Invoice_Head.Document_Type='I' then 'Invoice' else case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'Debit Note' else case when TSPL_Customer_Invoice_Head.Document_Type='C' then 'Credit Note' else '' end end end as Document_Type," &
            '" Account_Set,Document_Total as DocAmt,  Customer_Code ,TSPL_Customer_Invoice_Head.Customer_Name,tspl_user_master.User_Name as Created_By,user_master_modify.User_Name as Modify_By  , TSPL_JOURNAL_DETAILS.Detail_Line_No as Detail_Line_No ," &
            '" TSPL_JOURNAL_DETAILS.Account_code as Account_Code , TSPL_JOURNAL_DETAILS.Account_Desc as Account_Desc ,TSPL_JOURNAL_DETAILS.Amount , 0 as Discount ,TSPL_JOURNAL_DETAILS.Amount as Amount_less_Discount , 0 Total_Tax ,TSPL_JOURNAL_DETAILS.Amount as Total_Amount  ," &
            '" TSPL_Customer_Invoice_Head.Comp_Code ,TSPL_JOURNAL_DETAILS.Cost_Centre_Code,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name,TSPL_JOURNAL_DETAILS.Hirerachy_Code,TSPL_HIRERACHY_LEVEL_MASTER.Description as HIRERACHY_Name,TSPL_JOURNAL_DETAILS.Hirerachy_Code3 ,TSPL_JOURNAL_DETAILS.Hirerachy_Code4,TSPL_Customer_Invoice_Head.TapalNo,TSPL_Customer_Invoice_Head.DateAndTime  from TSPL_Customer_Invoice_Head " &
            '" left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =TSPL_Customer_Invoice_Head.Document_No left outer join TSPL_JOURNAL_DETAILS on TSPL_JOURNAL_DETAILS.Voucher_No = TSPL_JOURNAL_MASTER.Voucher_No   " &
            '" left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_Customer_Invoice_Head.Customer_Code  left outer join TSPL_LOCATION_MASTER on left(TSPL_LOCATION_MASTER.Location_Code,3)=TSPL_Customer_Invoice_Head.Loc_Code  " &
            '" left outer join TSPL_COST_CENTRE_FINANCIAL on TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code = TSPL_JOURNAL_DETAILS.Cost_Centre_Code  left outer join TSPL_HIRERACHY_LEVEL_MASTER on TSPL_HIRERACHY_LEVEL_MASTER.Hirerachy_Code = TSPL_JOURNAL_DETAILS.Hirerachy_Code  left outer join TSPL_STATE_MASTER as loc_state on loc_state.STATE_CODE =TSPL_LOCATION_MASTER.State  " &
            '" left outer join tspl_user_master on tspl_user_master.User_Code=TSPL_Customer_Invoice_Head.Created_By " &
            '" left outer join tspl_user_master as user_master_modify on user_master_modify.User_Code=TSPL_Customer_Invoice_Head.Modify_By " &
            '"where TSPL_Customer_Invoice_Head.Document_No ='" + StrCode + "' and Rejected_Type='N' )final   )XXX left outer join " &
            '"TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = XXX.Comp_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = XXX.Customer_Code left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code = XXX.Loc_Code left outer join tspl_state_master as tspl_state_master_for_location_state on  tspl_state_master_for_location_state.state_code=tspl_location_master.state  left outer join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Document_No=XXX.Document_No  order by XXX.Detail_Line_No  "
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            'Dim frmCRV As New frmCrystalReportViewer()
            'If SettingCostCenterlevel Then
            '    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, "crptAPInvc_Hierarchy", "AR INVOICE", clsCommon.myCDate(dt.Rows(0)("Document_Date")))
            'Else
            '    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, "crptAPInvc", "AR INVOICE", clsCommon.myCDate(dt.Rows(0)("Document_Date")))
            'End If
            'frmCRV = Nothing
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub FrmAPInvoiceEntry_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing Then
            If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentCell.ColumnInfo.Name), colACCode) = CompairStringResult.Equal Then
                OpenGLAccount(True)
            End If
        ElseIf e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
            AddNew()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag AndAlso btnSave.Enabled Then
            SaveData(clicked)
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso MyBase.isPostFlag AndAlso btnPost.Enabled Then
            PostData()
        ElseIf e.Control AndAlso e.KeyCode = Keys.P AndAlso btnPrint.Enabled Then
            PrintData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btnDelete.Enabled Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
            CloseForm()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverse.Visible = True
                    btnAdminCancel.Visible = True
                End If
                ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                          "========Table Name=========" + Environment.NewLine +
                          "TSPL_Customer_Invoice_Head " + Environment.NewLine +
                          "TSPL_Customer_Invoice_Detail" + Environment.NewLine +
                          "TSPL_REMITTANCE, TSPL_CUSTOM_FIELD_VALUES" + Environment.NewLine +
                          "=========Setting Name======" + Environment.NewLine +
                          "AllowCreditNoteWithoutReference" + Environment.NewLine +
                          "AllowUseApplyDocSeriesForPayment" + Environment.NewLine +
                          "CreateSeperateSeriesforRefDocARinvforCreditdebit" + Environment.NewLine +
                          "ShowHierarchyAndCostCenterInARInvoiceEntry" + Environment.NewLine +
                          "EnableHirerachyCostCentre" + Environment.NewLine +
                          "ShowTaxRateColumnOnTransaction" + Environment.NewLine +
                          "SkipCogsEntry" + Environment.NewLine +
                          "DefaultRoundOffGLAccount " + Environment.NewLine +
                          "=========Function======" + Environment.NewLine +
                          "GL entry ")
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
                'MessageBox.Show("You are not authorized to perform this action.", "Unauthorized Access", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If

    End Sub

    Private Sub MyComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs)
        txtRefDocNo.Value = ""
    End Sub

    Private Sub gv1_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv1.CellFormatting

        Try
            If e.Column.Index >= 0 Then
                If e.Column Is gv1.Columns(colHierarchyCode) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") = CompairStringResult.Equal OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colGLType).Value) <= 0 Then
                        gv1.CurrentRow.Cells(colHierarchyCode).ReadOnly = True
                    Else
                        If SettingCostCenter Or SettingCostCenterlevel Then
                            gv1.CurrentRow.Cells(colHierarchyCode).ReadOnly = False
                        Else
                            gv1.CurrentRow.Cells(colHierarchyCode).ReadOnly = True
                        End If
                    End If
                ElseIf e.Column Is gv1.Columns(colCostCenterCode) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") = CompairStringResult.Equal OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colGLType).Value) <= 0 OrElse SettingCostCenterlevel Then
                        gv1.CurrentRow.Cells(colCostCenterCode).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colCostCenterCode).ReadOnly = Not SettingCostCenter
                    End If
                ElseIf e.Column Is gv1.Columns(colHierarchyLevel4) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") = CompairStringResult.Equal OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colGLType).Value) <= 0 OrElse Not SettingCostCenterlevel OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colHierarchyCode).Value) <= 0 Then
                        gv1.CurrentRow.Cells(colHierarchyLevel4).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colHierarchyLevel4).ReadOnly = Not (clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyLevelNumber).Value) = 4)
                    End If
                ElseIf e.Column Is gv1.Columns(colHierarchyLevel3) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") = CompairStringResult.Equal OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colGLType).Value) <= 0 OrElse Not SettingCostCenterlevel OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colHierarchyCode).Value) <= 0 Then
                        gv1.CurrentRow.Cells(colHierarchyLevel3).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colHierarchyLevel3).ReadOnly = Not (clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyLevelNumber).Value) = 3)
                    End If
                ElseIf e.Column Is gv1.Columns(colHierarchyLevel2) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") = CompairStringResult.Equal OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colGLType).Value) <= 0 OrElse Not SettingCostCenterlevel OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colHierarchyCode).Value) <= 0 Then
                        gv1.CurrentRow.Cells(colHierarchyLevel2).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colHierarchyLevel2).ReadOnly = Not (clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyLevelNumber).Value) = 2)
                    End If
                ElseIf e.Column Is gv1.Columns(colHierarchyLevel1) Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colGLType).Value), "Balance Sheet") = CompairStringResult.Equal OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colGLType).Value) <= 0 OrElse Not SettingCostCenterlevel OrElse clsCommon.myLen(gv1.CurrentRow.Cells(colHierarchyCode).Value) <= 0 Then
                        gv1.CurrentRow.Cells(colHierarchyLevel1).ReadOnly = True
                    Else
                        gv1.CurrentRow.Cells(colHierarchyLevel1).ReadOnly = Not (clsCommon.myCstr(gv1.CurrentRow.Cells(colHierarchyLevelNumber).Value) = 1)
                    End If
                End If
            End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        End Try
    End Sub

    Private Sub gv1_UserDeletedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gv1.UserDeletedRow
        UpdateAllTotals()
        For ii As Integer = 1 To gv1.Rows.Count
            gv1.Rows(ii - 1).Cells(colLineNo).Value = ii
        Next
    End Sub

    Private Sub gv1_UserDeletingRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv1.UserDeletingRow
        If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub txtDocNo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocNo.Load

    End Sub

    Private Sub txtRefDocNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtRefDocNo._MYValidating
        Try
            Dim StrInvoiceType As String = String.Empty

            If clsCommon.myLen(TxtCustomer.Value) > 0 Then
                If objCommonVar.IsDemoERP Then
                    If clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal Or clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal Then
                        StrInvoiceType = " AND TSPL_Customer_Invoice_Head.Document_Type not in ('C','D') "
                    End If
                    ''richa 25 Dec,2020 
                    Dim qry As String = String.Empty
                    If clsERPFuncationality.GetEInvoiceStatus(txtDate.Value, Nothing) = True Then
                        StrInvoiceType += " and isnull(TSPL_SD_SALE_INVOICE_HEAD.IRN_No,'')=''"
                        qry = "SELECT TSPL_Customer_Invoice_Head.Document_No as Code,Convert(varchar,TSPL_Customer_Invoice_Head.Document_Date,103) as Date,ISNULL(TSPL_Customer_Invoice_Head.Against_Sale_No,'') as [Sale Invoice No] ,isnull(TSPL_Customer_Invoice_Head.PROJECT_ID,'') AS PROJECT_ID,ISNULL(TSPL_Customer_Invoice_Head.Customer_Code,'') AS Customer_Code FROM TSPL_Customer_Invoice_Head left outer join TSPL_SD_SALE_INVOICE_HEAD on  TSPL_Customer_Invoice_Head.Against_Sale_No = TSPL_SD_SALE_INVOICE_HEAD.Document_Code"
                    Else
                        qry = "SELECT TSPL_Customer_Invoice_Head.Document_No as Code,Convert(varchar,TSPL_Customer_Invoice_Head.Document_Date,103) as Date,ISNULL(TSPL_Customer_Invoice_Head.Against_Sale_No,'') as [Sale Invoice No] , isnull(TSPL_Customer_Invoice_Head.PROJECT_ID,'') AS PROJECT_ID,ISNULL(TSPL_Customer_Invoice_Head.Customer_Code,'') AS Customer_Code FROM TSPL_Customer_Invoice_Head "
                    End If

                    txtRefDocNo.Value = clsCommon.ShowSelectForm("ARINVNO", qry, "Code", " TSPL_Customer_Invoice_Head.Status=1 and TSPL_Customer_Invoice_Head.Customer_Code='" + TxtCustomer.Value + "' " + StrInvoiceType, txtRefDocNo.Value, "", isButtonClicked)
                    'fndProject.Value = clsDBFuncationality.getSingleValue("select isnull(Project_Id,'') from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" + txtRefDocNo.Value + "'")
                    fndProject.Value = clsDBFuncationality.getSingleValue("select isnull(PROJECT_ID,'') from TSPL_Customer_Invoice_Head where Document_No='" + txtRefDocNo.Value + "'")
                    lblProject.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + fndProject.Value + "'")
                    fndProject.Enabled = False
                    lblProject.Enabled = False
                Else
                    'Dim qry As String = "select Sale_Invoice_No as Code,Sale_Invoice_Date as Date,Location from TSPL_SALE_INVOICE_HEAD"
                    Dim qry As String = "SELECT Document_No as Code,Document_Date as Date, isnull(PROJECT_ID,'') AS PROJECT_ID,ISNULL(Customer_Code,'') AS Customer_Code FROM TSPL_Customer_Invoice_Head"
                    'txtRefDocNo.Value = clsCommon.ShowSelectForm("ARINVNO", qry, "Code", "Is_Post='Y' and Cust_Code='" + TxtCustomer.Value + "' ", txtRefDocNo.Value, "", isButtonClicked)
                    txtRefDocNo.Value = clsCommon.ShowSelectForm("ARINVNO", qry, "Code", "Status=1 and Customer_Code='" + TxtCustomer.Value + "' ", txtRefDocNo.Value, "", isButtonClicked)

                End If

            Else
                common.clsCommon.MyMessageBoxShow(Me, "Please select Customer", Me.Text)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadBlankGridAC()
        gvAC.Rows.Clear()
        gvAC.Columns.Clear()

        Dim repoACCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoACCode.FormatString = ""
        repoACCode.HeaderText = "Addition Charges Code"
        repoACCode.Name = colAChgCode
        repoACCode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoACCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoACCode.Width = 150
        repoACCode.ReadOnly = False
        gvAC.MasterTemplate.Columns.Add(repoACCode)

        Dim repoACName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoACName.FormatString = ""
        repoACName.HeaderText = "Addition Charges Description"
        repoACName.Name = colAChgName
        repoACName.Width = 300
        repoACName.ReadOnly = True
        gvAC.MasterTemplate.Columns.Add(repoACName)

        Dim repoACAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoACAmt.FormatString = ""
        repoACAmt.HeaderText = "Amount"
        repoACAmt.Name = colAChgAmount
        repoACAmt.Width = 100
        repoACAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoACAmt.ReadOnly = False
        gvAC.MasterTemplate.Columns.Add(repoACAmt)

        gvAC.AllowAddNewRow = False
        gvAC.ShowGroupPanel = False
        gvAC.AllowColumnReorder = True
        gvAC.AllowRowReorder = False
        gvAC.EnableSorting = False
        gvAC.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvAC.MasterTemplate.ShowRowHeaderColumn = False
        gvAC.TableElement.TableHeaderHeight = 40
    End Sub

    Private Sub gvAC_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gvAC.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column Is gvAC.Columns(colAChgAmount) Then

                        UpdateAllTotals()
                    ElseIf e.Column Is gvAC.Columns(colAChgCode) Then
                        Dim obj As clsAdditionalCharge = clsAdditionalCharge.GetFinder(clsCommon.myCstr(gvAC.CurrentRow.Cells(colAChgCode).Value), False)
                        If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Code) > 0 Then
                            gvAC.CurrentRow.Cells(colAChgCode).Value = obj.Code
                            gvAC.CurrentRow.Cells(colAChgName).Value = obj.desc
                        Else
                            gvAC.CurrentRow.Cells(colAChgCode).Value = ""
                            gvAC.CurrentRow.Cells(colAChgName).Value = ""
                            gvAC.CurrentRow.Cells(colAChgAmount).Value = 0
                        End If
                    End If
                End If
                setGridFocusAC()
                isCellValueChangedOpen = False
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub setGridFocusAC()
        Try
            Dim intCurrRow As Integer = gvAC.CurrentRow.Index
            If intCurrRow = gvAC.Rows.Count - 1 AndAlso gvAC.Rows.Count <= 10 Then
                gvAC.Rows.AddNew()
                gvAC.CurrentRow = gvAC.Rows(intCurrRow)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnDrillDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDrillDown.Click
        If clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "I") = CompairStringResult.Equal Then
            Dim strInvoiceNo As String = String.Empty
            strInvoiceNo = clsDBFuncationality.getSingleValue("select Against_Sale_No from TSPL_Customer_Invoice_Head where Document_No='" & txtDocNo.Value & "'")
            'strDemoInvNo = clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & strInvoiceNo & "'")
            'strDemoInvNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_Customer_Invoice_Head where Against_Sale_No='" & strInvoiceNo & "'")
            Dim StrTrans_type As String = clsDBFuncationality.getSingleValue("select Trans_type from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & strInvoiceNo & "'")

            'If Not strDemoInvNo Is Nothing Then
            'If clsCommon.CompairString(StrTrans_type, "MCC") = CompairStringResult.Equal Then
            '    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmMCCMaterial, strDemoInvNo)
            'Else
            '    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSNServiceInvoice, strDemoInvNo)
            'End If
            '' Anubhooti 06-Apr-2015 (Drill Down Working According To All Types Of Sale)
            Dim strBSType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Trans_Type,'') As Trans_Type FROM TSPL_Customer_Invoice_Head WHERE Document_No='" & txtDocNo.Value & "'"))

            If clsCommon.myLen(StrTrans_type) > 0 Then
                If clsCommon.CompairString(StrTrans_type, "MCC") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmMCCMaterial, strInvoiceNo)
                    'ElseIf clsCommon.CompairString(StrTrans_type, "FS") = CompairStringResult.Equal Then
                    '    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmInvoiceFreshSale, strInvoiceNo)
                ElseIf clsCommon.CompairString(StrTrans_type, "FS") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSNSaleInvoice, strInvoiceNo)
                ElseIf clsCommon.CompairString(StrTrans_type, "PS") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSaleInvoiceProductSale, strInvoiceNo)
                ElseIf clsCommon.CompairString(StrTrans_type, "CSA") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmCSASaleInvoice, strInvoiceNo)
                ElseIf clsCommon.CompairString(StrTrans_type, "EXP") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmEXSalesInvoice, strInvoiceNo)
                ElseIf clsCommon.CompairString(StrTrans_type, "ALL") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSNServiceInvoice, strInvoiceNo)
                End If
            ElseIf clsCommon.myLen(StrTrans_type) <= 0 AndAlso clsCommon.CompairString(strBSType, "BS") = CompairStringResult.Equal Then '' Bulk Sale
                clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.FrmInvoiceBulkSale, strInvoiceNo)
            End If
            ''
            'End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "C") = CompairStringResult.Equal Then
            Dim strReturnNo As String = String.Empty
            strReturnNo = clsDBFuncationality.getSingleValue("select Against_Sale_Return_No from TSPL_Customer_Invoice_Head where Document_No='" & txtDocNo.Value & "'")
            'strDemoRetNo = clsDBFuncationality.getSingleValue("select Document_Code from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & strReturnNo & "'")
            Dim StrSRType = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  ISNULL(Trans_Type,'') AS Trans_Type from TSPL_SD_SALE_RETURN_HEAD where Document_Code='" & strReturnNo & "'"))
            '' Anubhooti 06-Apr-2015 (Drill Down Working According To All Types Of Sale)
            Dim strBSRType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Trans_Type,'') As Trans_Type FROM TSPL_Customer_Invoice_Head WHERE Document_No='" & txtDocNo.Value & "'"))

            If clsCommon.myLen(StrSRType) > 0 Then
                If clsCommon.CompairString(StrSRType, "FS") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.saleReturn, strReturnNo)
                ElseIf clsCommon.CompairString(StrSRType, "PS") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSaleReturnProductSale, strReturnNo)
                ElseIf clsCommon.CompairString(StrSRType, "MCC") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmMCCMaterialSaleReturn, strReturnNo)
                ElseIf clsCommon.CompairString(StrSRType, "ALL") = CompairStringResult.Equal Then
                    clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.frmSNSaleReturn, strReturnNo)
                ElseIf clsCommon.CompairString(StrSRType, "EXP") = CompairStringResult.Equal Then
                    ' Screen not ready
                End If
            ElseIf clsCommon.myLen(StrSRType) <= 0 AndAlso clsCommon.CompairString(strBSRType, "BS") = CompairStringResult.Equal Then '' Bulk Sale Return
                clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.FrmBulkSaleReturn, strReturnNo)
            End If
        ElseIf clsCommon.CompairString(clsCommon.myCstr(cboDocType.SelectedValue), "D") = CompairStringResult.Equal Then '' DEBIT NOTE
            Dim strVCGLNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select ISNULL(Against_VCGL,'') AS Against_VCGL FROM TSPL_Customer_Invoice_Head WHERE Document_No='" & txtDocNo.Value & "'"))
            If clsCommon.myLen(strVCGLNo) > 0 Then
                clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.mbtnVCGLEntry, strVCGLNo)
            End If
        End If

    End Sub

    Private Sub rbtnTaxCalAutomatic_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtnTaxCalAutomatic.ToggleStateChanged, rbtnTaxCalManual.ToggleStateChanged
        If Not isInsideLoadData Then
            If rbtnTaxCalAutomatic.IsChecked Then
                txtTaxGroup_TxtChanged()
            ElseIf rbtnTaxCalManual.IsChecked Then
                For intRowNo As Integer = 0 To gv2.Rows.Count - 1
                    gv2.Rows(intRowNo).Cells(colTTaxRate).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTTaxAmt).Value = Nothing
                    gv2.Rows(intRowNo).Cells(colTBaseAmt).Value = Nothing
                Next
                For intRowNo As Integer = 0 To gv1.Rows.Count - 1
                    For ii As Integer = 1 To 10
                        Dim strII As String = clsCommon.myCstr(ii)
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = Nothing
                        gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strII)).Value = Nothing
                    Next
                Next

            End If
        End If
    End Sub

    Private Sub gv2_CellFormatting(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CellFormattingEventArgs) Handles gv2.CellFormatting
        Try
            If e.Column.Index >= 0 Then
                If (e.Column Is gv2.Columns(colTTaxAmt)) Then
                    gv2.CurrentRow.Cells(colTTaxAmt).ReadOnly = rbtnTaxCalAutomatic.IsChecked
                End If

                Dim cell As GridDataCellElement = TryCast(e.CellElement, GridDataCellElement)
                cell.GradientStyle = GradientStyles.Solid
                cell.BackColor = Color.FromArgb(243, 181, 51)
            End If
        Catch ex As Exception
            'common.clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        End Try
    End Sub

    Private Sub gv2_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv2.CellValueChanged
        Try
            If (Not isInsideLoadData) Then
                If Not isCellValueChangedTaxOpen Then
                    isCellValueChangedTaxOpen = True
                    If (e.Column Is (gv2.Columns(colTTaxAmt)) AndAlso rbtnTaxCalManual.IsChecked) Then
                        For ii As Integer = 0 To gv1.Rows.Count - 1
                            UpdateCurrentRowWihtRowNo(ii)
                        Next
                        UpdateAllTotals()
                    End If
                    isCellValueChangedTaxOpen = False
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub cboDocType_SelectedIndexChanged(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs) Handles cboDocType.SelectedIndexChanged
        If blnInvoice = False Then
            If ((clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal)) Then
                txtRefDocNo.Visible = True
                RadLabel15.Visible = True
                txtRefDocNo.Value = ""
                chkSecurityDposit.Enabled = True
            Else
                txtRefDocNo.Visible = False
                RadLabel15.Visible = False
                txtRefDocNo.Value = ""
                chkSecurityDposit.Enabled = False
                chkSecurityDposit.Checked = False
            End If
        End If
    End Sub

    Private Sub ddlSecDepositType_SelectedIndexChanged(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs) Handles ddlSecDepositType.SelectedIndexChanged
        LoadSecurityAccount(TxtCustomer.Value)
    End Sub

    Private Sub cboDocType_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboDocType.Validating
        If ((clsCommon.CompairString(cboDocType.SelectedValue, "D") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(cboDocType.SelectedValue, "C") = CompairStringResult.Equal)) Then
            txtRefDocNo.Visible = True
            RadLabel15.Visible = True
            txtRefDocNo.Value = ""
        Else
            txtRefDocNo.Visible = False
            RadLabel15.Visible = False
            txtRefDocNo.Value = ""
        End If

    End Sub

    Private Sub RadMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem4.Click
        Dim Sql As String = "select '' as [Date], '' as Customer, 0 as Amount,'' as LocSegment"
        transportSql.ExporttoExcel(Sql, Me)
    End Sub

    Private Sub RadMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem5.Click
        Dim type As String = "I"
        funfillimport(type)
    End Sub
    '----------------------For Credit Note---------------------
    Private Sub RadMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem6.Click
        Dim type As String = "C"
        funfillimport(type)
    End Sub
    '----------------------For Debit Note------------------------
    Private Sub RadMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadMenuItem7.Click
        Dim type As String = "D"
        funfillimport(type)
    End Sub

    Sub funfillimport(ByRef type As String)
        Dim BalanceType As String = type
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim Counter As Integer = 0
        Dim qry As String = "select top 1 Account_Code,Description from TSPL_GL_ACCOUNTS where LEN(Account_Code)>6 order by Account_Code"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        Dim strACCode As String = clsCommon.myCstr(dt.Rows(0)("Account_Code"))
        Dim strACName As String = clsCommon.myCstr(dt.Rows(0)("Description"))
        If transportSql.importExcel(gv, "Date", "Customer", "Amount", "LocSegment") Then
            Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
            clsCommon.ProgressBarShow()
            Try
                For Each grow As GridViewRowInfo In gv.Rows
                    Dim strDate As String
                    Dim strCustomer As String = clsCommon.myCstr(grow.Cells("Customer").Value)
                    Dim dblAmt As Double = clsCommon.myCdbl(grow.Cells("Amount").Value)
                    If Not IsDBNull(grow.Cells("Date").Value) Then
                        strDate = clsCommon.GetPrintDate(grow.Cells("Date").Value, "dd/MM/yyyy")
                    End If
                    Dim segment As String = clsCommon.myCstr(grow.Cells("LocSegment").Value)

                    If clsCommon.myLen(strCustomer) <= 0 Then
                        Throw New Exception("Customer not found")
                    End If

                    If clsCommon.myLen(grow.Cells("Date").Value) <= 0 Then
                        Throw New Exception("Transactin Date is not found")
                    End If
                    If dblAmt <= 0 Then
                        Throw New Exception("Amount not found")
                    End If
                    Counter += 1


                    Dim obj As New clsCustomerInvoiceHead()
                    'obj.Document_No = txtDocNo.Value
                    obj.Document_Date = strDate
                    obj.Customer_Code = strCustomer
                    obj.loc_code = segment
                    Dim objCustomer As clsCustomerMaster = clsCustomerMaster.GetData(strCustomer, trans)
                    obj.Customer_Name = objCustomer.Customer_Name
                    obj.Account_Set = objCustomer.Cust_Account
                    obj.Document_Type = BalanceType
                    ''obj.RefDocNo = txtRefDocNo.Value
                    obj.Order_No = "C" ''if it is C then can't be put in GL entry
                    obj.Total_Tax = 0
                    obj.On_Hold = False
                    'obj.Description = txtDesc.Text
                    obj.Tax_Group = "NILL"

                    obj.Due_Date = strDate
                    obj.Discount_Base = dblAmt
                    obj.Discount_Amount = 0
                    obj.Amount_Less_Discount = dblAmt
                    obj.Document_Total = dblAmt
                    obj.Balance_Amt = dblAmt
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic

                    dt = clsDBFuncationality.GetDataTable("select Receivable_Control_acct,Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + obj.Account_Set + "'", trans)

                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        obj.Customer_Control_AC = clsCommon.myCstr(dt.Rows(0)("Receivable_Control_acct"))
                        If obj.Discount_Amount > 0 Then
                            obj.Discount_GL_AC = clsCommon.myCstr(dt.Rows(0)("Receipts_Discount_acct"))
                        End If
                    End If
                    obj.Total_Add_Charge = 0
                    obj.Arr = New List(Of clsCustomerInvoiceDetail)
                    Dim objTr As New clsCustomerInvoiceDetail()
                    objTr.SNo = 1


                    '---------------------------------------------------

                    'For retriving Control Account for Customer

                    ' Dim segment As String = clsCommon.myCstr(grow.Cells("LocSegment").Value)

                    ''richa 05 feb,2019  TEC/05/02/19-000411 
                    'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, trans)
                    Dim JEWithOPening As Boolean = False
                    If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                        Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                        If clsCommon.myCDate(clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                            JEWithOPening = True
                        End If
                    End If
                    Dim qry1 As String = String.Empty
                    If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, trans)), "1") = CompairStringResult.Equal And JEWithOPening = True Then
                        qry1 = "select TSPL_CUSTOMER_ACCOUNT_SET.Customer_Opening_Clearing_AC ,tspl_gl_accounts.Description " &
                 "  from TSPL_CUSTOMER_MASTER " &
                 " left outer join TSPL_CUSTOMER_ACCOUNT_SET  on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account " &
                  " left outer join tspl_gl_accounts on  TSPL_CUSTOMER_ACCOUNT_SET.Receivable_Control_acct=tspl_gl_accounts.Account_Code where TSPL_CUSTOMER_MASTER.Cust_Code='" + strCustomer + "' "
                    Else
                        qry1 = "select TSPL_CUSTOMER_ACCOUNT_SET.Receivable_Control_acct ,tspl_gl_accounts.Description " &
                 "  from TSPL_CUSTOMER_MASTER " &
                 " left outer join TSPL_CUSTOMER_ACCOUNT_SET  on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account " &
                  " left outer join tspl_gl_accounts on  TSPL_CUSTOMER_ACCOUNT_SET.Receivable_Control_acct=tspl_gl_accounts.Account_Code where TSPL_CUSTOMER_MASTER.Cust_Code='" + strCustomer + "' "
                    End If


                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry1, trans)
                    Dim strACName1 As String
                    Dim strACCode1 As String
                    If dt1.Rows.Count > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, trans)), "1") = CompairStringResult.Equal And JEWithOPening = True Then
                            strACCode1 = clsCommon.myCstr(dt1.Rows(0)("Customer_Opening_Clearing_AC"))
                        Else
                            strACCode1 = clsCommon.myCstr(dt1.Rows(0)("Receivable_Control_acct"))
                        End If

                        strACName1 = clsCommon.myCstr(dt1.Rows(0)("Description"))
                    Else
                        Throw New Exception("Account does not exist for Customer '" + strCustomer + "'")
                    End If


                    '-------------For  Segment code overwriting
                    ' Dim acccontrol As String
                    ' Dim acccontdesc As String

                    Dim AccountFinal As String = strACCode1.Substring(0, (clsCommon.myLen(strACCode1) - 4)) + "-" + segment

                    Dim qryAcc As String = "select Account_Code,Description from TSPL_GL_ACCOUNTS where Account_Code='" + AccountFinal + "'"
                    Dim dtfinal As DataTable = clsDBFuncationality.GetDataTable(qryAcc, trans)
                    Dim FinalAcc As String
                    Dim FinalDesc As String
                    If dtfinal.Rows.Count > 0 Then
                        FinalAcc = clsCommon.myCstr(dtfinal.Rows(0)("Account_Code"))
                        FinalDesc = clsCommon.myCstr(dtfinal.Rows(0)("Description"))
                    Else
                        Throw New Exception("Account does not exist for Customer '" + strCustomer + "'")
                    End If

                    objTr.GL_Account_Code = FinalAcc
                    objTr.GL_Account_Desc = FinalDesc
                    '---------------------------------------------------





                    'objTr.GL_Account_Code = strACCode
                    'objTr.GL_Account_Desc = strACName
                    objTr.Amount = dblAmt
                    objTr.Discount_Per = 0
                    objTr.Discount = 0
                    objTr.Amount_less_Discount = dblAmt

                    objTr.Total_Tax = 0
                    objTr.Total_Amount = dblAmt
                    If (clsCommon.myLen(objTr.GL_Account_Code) > 0) Then
                        obj.Arr.Add(objTr)
                    End If
                    ''Next
                    obj.SaveData(obj, True, trans, "")
                Next

                clsCommon.ProgressBarHide()
                trans.Commit()
                common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                clsCommon.ProgressBarHide()
                trans.Rollback()
                common.clsCommon.MyMessageBoxShow(Me, "Error at Rowno " + clsCommon.myCstr(Counter) + Environment.NewLine + ex.Message, Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
            Finally
                clsCommon.ProgressBarHide()
                Me.Controls.Remove(gv)
            End Try
        End If
    End Sub

    Private Sub txtlocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtlocation._MYValidating
        Try
            'Dim qry As String = "select Segment_code as Code ,Description  from TSPL_GL_SEGMENT_CODE "
            '' Anubhooti 16-Sep-2014 BM00000003782 (Filetr AND GIT='N')
            Dim qry As String = "select distinct(Segment_code) as Code ,Description  from TSPL_GL_SEGMENT_CODE left outer join TSPL_LOCATION_MASTER on TSPL_GL_SEGMENT_CODE .Segment_code =TSPL_LOCATION_MASTER .Loc_Segment_Code "
            Dim WhrCls As String = "Seg_No = '7' AND GIT='N'"
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                WhrCls += "  and  TSPL_LOCATION_MASTER.Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
            End If
            txtlocation.Value = clsCommon.ShowSelectForm("ARlocation", qry, "Code", WhrCls, txtlocation.Value, "Code", isButtonClicked)
            '' Anubhooti 28-Nov-2014 BM00000004810
            If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
                LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
            Else
                LblLocDesp.Text = ""
            End If
            ''
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub gv1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gv1.Click

    End Sub

    Private Sub btnReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverse.Click
        ''richa TEC/22/01/19-000406
        Try
            If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
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

                If clsCustomerInvoiceHead.ReverseAndUnpost(txtDocNo.Value) Then
                    saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                    LoadData(txtDocNo.Value)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub txtCurrencyCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtCurrencyCode._MYValidating
        Dim qry As String = "select CURRENCY_CODE AS Code, CURRENCY_NAME as Name from TSPL_CURRENCY_MASTER"
        txtCurrencyCode.Value = clsCommon.ShowSelectForm("CURRENCY_MASTER", qry, "Code", "", txtCurrencyCode.Value, "CURRENCY_CODE", isButtonClicked)
        ShowCurrencyDetail()
    End Sub

    Public Sub SetMailRight()
        'Dim obj As clsCheckMailSetting = clsCheckMailSetting.CheckMailRight()
        If objCommonVar.IsMailSend Then
            btnsend.Enabled = True
        Else
            btnsend.Enabled = False
        End If

    End Sub

    Private Sub btnSendForApproval_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSendForApproval.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select Document No. First", Me.Text)
                txtDocNo.Focus()
                txtDocNo.Select()
                Return
            End If

            If Not (common.clsCommon.MyMessageBoxShow(Me, "Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
                Return
            End If
            LoadData(txtDocNo.Value)

            Dim qry As String = "Select * from TSPL_APPROVAL_LEVEL_SCREEN where 1=1 and Trans_Code='" + MyBase.Form_ID + "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            Dim lstUsers As New List(Of String)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    lstUsers.Add(dr("User_Code").ToString())
                Next
            End If

            If lstUsers.Count = 0 Then
                Throw New Exception("No Receiptent Found")
            End If
            SendSMSandEmail(lstUsers, True)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnsend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsend.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select Document No. First", Me.Text)
                txtDocNo.Focus()
                txtDocNo.Select()
                Return
            End If

            If Not (common.clsCommon.MyMessageBoxShow(Me, "Send E-Mail/SMS Of Respective Sale Order No. " + txtDocNo.Value + "" + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes) Then
                Return
            End If
            LoadData(txtDocNo.Value)
            Dim lstUsers As New List(Of String)
            lstUsers.Add(TxtCustomer.Value)
            SendSMSandEmail(lstUsers, False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub SendSMSandEmail(ByVal lstUsers As List(Of String), ByVal isSendForApproval As Boolean)

        Try
            'Dim obj As clsEmailSMSSettingNew = clsEmailSMSSettingNew.GetData(clsUserMgtCode.mbtnPurchaseInvoice)

            'If obj Is Nothing Then
            '    clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
            '    Return
            'End If
            'If clsCommon.myLen(obj.mailsubjct) <= 0 Then
            '    clsCommon.MyMessageBoxShow("First do email and sms setting", Me.Text)
            '    Return
            'End If

            'Dim strContactPerson As String = ""
            'Dim strSubject As String = obj.mailsubjct.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'strSubject = strSubject.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))

            'Dim strbody As String = obj.mailbody.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
            'strbody = strbody.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
            'strbody = strbody.Replace(clsEmailSMSConstants.VendorNo, TxtCustomer.Value)
            'strbody = strbody.Replace(clsEmailSMSConstants.VendorName, lblVendorName.Text)
            'strbody = strbody.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
            'strbody = strbody.Replace(clsEmailSMSConstants.Form_Code, MyBase.Form_ID)

            ''------------------------code for attchament-------------------------------------
            'Dim strRptPath As String = ""
            ''obj.atchmnt = "N"
            'If obj.atchmnt = "Y" Then

            '    Dim atchqry As String = GetAtchmentPrintQuery()
            '    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(atchqry)
            '    If dt1.Rows.Count > 0 Then
            '        'SetItemWiseTax(dt1, txtDocNo.Value)
            '        Dim frmCRV As New frmCrystalReportViewer()
            '        strRptPath = frmCRV.EmailAttachment(CrystalReportFolder.NewSalesReports, dt1, "crptAPInvc", "AR Invoice")
            '        frmCRV = Nothing
            '    End If
            'End If
            ''---------------------------------------------------------------------------



            'For Each strUser As String In lstUsers
            '    'lstUsers.Add(dr("User_Code").ToString())
            '    Dim lstReceiptents As New List(Of String)
            '    Dim qry As String = ""
            '    Dim emailId As String = ""
            '    If isSendForApproval Then
            '        strContactPerson = strUser
            '        qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" + strUser + "') "
            '        emailId = clsDBFuncationality.getSingleValue(qry)
            '    Else
            '        strContactPerson = clsDBFuncationality.getSingleValue("select upper(substring(Contact_Person_Name,1,1)) + lower(substring(Contact_Person_Name,2,49)) from TSPL_CUSTOMER_MASTER where Cust_Code ='" & strUser & "' ")
            '        emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_CUSTOMER_MASTER where Cust_Code ='" & strUser & "' ")
            '    End If

            '    strbody = strbody.Replace(clsEmailSMSConstants.ContactPerson, strContactPerson)
            '    lstReceiptents.Add(emailId)

            '    Dim body As String = strbody.Replace(clsEmailSMSConstants.UserCode, strUser)

            '    clsMailViaOutlook.SendEmail(strSubject, GridTableBodyElement, lstReceiptents, Nothing, strRptPath)
            'Next
            'clsCommon.MyMessageBoxShow("E-Mail Send Successfully", Me.Text)


            'sanjay
            Dim strPhoneno As String = Nothing

            Dim dtContent As DataTable = clsDBFuncationality.GetDataTable("SELECT SMS_Text,Email_Text,Email_subject from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.mbtnARInvoiceEntry + "'", Nothing)
            Dim objEmailH As New clsEMailHead()
            objEmailH.arrEMail = New List(Of String)()
            Dim objSMSH As New clsSMSHead()
            objSMSH.arrMobilNo = New List(Of String)()
            If dtContent IsNot Nothing AndAlso dtContent.Rows.Count > 0 Then

                If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 Then
                    objEmailH.Email_Subject = clsCommon.myCstr(dtContent.Rows(0)("Email_subject"))
                    objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                    objEmailH.Email_Subject = objEmailH.Email_Subject.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))

                    objEmailH.Email_Text = clsCommon.myCstr(dtContent.Rows(0)("Email_Text"))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Cust_Code, TxtCustomer.Value)
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Cust_Name, lblVendorName.Text)
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.DocAmount, clsCommon.myCstr(lblTotRAmt.Text))
                    objEmailH.Email_Text = objEmailH.Email_Text.Replace(frmEMailAndSMSSetting.Form_Code, MyBase.Form_ID)

                End If


                If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then

                    objSMSH.SMS_Text = clsCommon.myCstr(dtContent.Rows(0)("SMS_Text"))

                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_No, txtDocNo.Value)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Doc_Date, clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy"))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Cust_Code, TxtCustomer.Value)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Cust_Name, lblVendorName.Text)
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.DocAmount, clsCommon.myCstr(lblTotRAmt.Text))
                    objSMSH.SMS_Text = objSMSH.SMS_Text.Replace(frmEMailAndSMSSetting.Form_Code, MyBase.Form_ID)
                End If

            End If

            '------------------------code for attchament-------------------------------------
            If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 Then

                Dim strRptPath As String = ""
                'obj.atchmnt = "N"
                'If obj.atchmnt = "Y" Then

                Dim atchqry As String = GetAtchmentPrintQuery()
                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(atchqry)
                If dt1.Rows.Count > 0 Then
                    'SetItemWiseTax(dt1, txtDocNo.Value)
                    Dim frmCRV As New frmCrystalReportViewer()
                    strRptPath = frmCRV.EmailAttachment(CrystalReportFolder.SalesReport, dt1, "crptAPInvc", "AR Invoice")
                    objEmailH.Attachment_1_Path = strRptPath
                    frmCRV = Nothing
                End If
                'End If
            End If
            '---------------------------------------------------------------------------

            For Each strUser As String In lstUsers
                Dim lstReceiptents As New List(Of String)
                Dim qry As String = ""
                Dim emailId As String = ""
                If isSendForApproval Then
                    qry = "select E_Mail from TSPL_USER_MASTER where User_Code in ('" + strUser + "') "
                    emailId = clsDBFuncationality.getSingleValue(qry)
                    strPhoneno = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(mob_no,'') from TSPL_USER_MASTER where user_code ='" & strUser & "' "))
                Else
                    emailId = clsDBFuncationality.getSingleValue("select Email from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
                    strPhoneno = clsDBFuncationality.getSingleValue("select isnull(Phone1,'') from TSPL_customer_MASTER where cust_code ='" & strUser & "' ")
                End If

                lstReceiptents.Add(emailId)
                objSMSH.arrMobilNo.Add(clsCommon.myCstr(strPhoneno))
                objEmailH.arrEMail.Add(clsCommon.myCstr(emailId))
            Next

            If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 Then
                objEmailH.SaveData(clsUserMgtCode.mbtnARInvoiceEntry, objEmailH, Nothing)
                objEmailH = Nothing
            End If
            If clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                objSMSH.SaveData(clsUserMgtCode.mbtnARInvoiceEntry, objSMSH, Nothing)
                objSMSH = Nothing
            End If
            If clsCommon.myLen(dtContent.Rows(0)("Email_Text")) > 0 Or clsCommon.myLen(dtContent.Rows(0)("SMS_Text")) > 0 Then
                clsCommon.MyMessageBoxShow(Me, "E-Mail/SMS Send Successfully", Me.Text)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        'Try
        '    Dim client As New System.Net.WebClient()

        '    If clsCommon.myLen(obj.smsbody) <= 0 Then
        '        Return
        '    End If

        '    'strMes = "Dear  " & strCustomer & "  your order No " & txtDocNo.Value & "  dated  " & clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy") & "has been booked with amount" & lblTotRAmt.Text

        '    strMes = obj.smsbody
        '    strMes = strMes.Replace("'", " ").Replace("`", "/")

        '    If strMes.Contains(clsEmailSMSConstants.SaleOrderNo) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.SaleOrderNo, txtDocNo.Value)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.SaleOrderDate) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.SaleOrderDate, clsCommon.GetPrintDate(txtDate.Text, "dd/MMM/yyyy"))
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.VendorNo) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.VendorNo, txtVendorNo.Value)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.VendorName) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.VendorName, lblVendorName.Text)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.ContactPerson) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.ContactPerson, strContactperson)
        '    End If
        '    If strMes.Contains(clsEmailSMSConstants.TotalAmount) Then
        '        strMes = strMes.Replace(clsEmailSMSConstants.TotalAmount, lblTotRAmt.Text)
        '    End If


        '    strphone = clsDBFuncationality.getSingleValue("select Phone1 from TSPL_customer_MASTER where cust_code ='" & txtVendorNo.Value & "' ")

        '    Dim baseurl As String = "http://bulksms.mysmsmantra.com:8080/WebSMS/SMSAPI.jsp?username=tecxpert&password=1818948263&sendername=vipin&mobileno=91" + strphone + "&message=" + strMes
        '    Dim data As Stream = client.OpenRead(baseurl)
        '    Dim reader As StreamReader = New StreamReader(data)
        '    Dim s As String = reader.ReadToEnd()
        '    data.Close()
        '    reader.Close()

        '    clsCommon.MyMessageBoxShow("SMS Send Successfully", Me.Text)
        'Catch ex As Exception
        '    Throw New Exception(ex.Message)
        'End Try
    End Sub

    Private Function GetAtchmentPrintQuery()
        Dim qry As String = "select  Location_Desc,xxx.Tin_No,PAN,TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end  as CompAdd , " &
            "XXX.Description,XXX.Account_Code,XXX.Account_Desc ,XXX.DrAmt ,XXX.CrAmt ,XXX.Document_No ,XXX.Document_Date,XXX.Status , " &
            "XXX.Document_Type ,XXX.Account_Set ,XXX.DocAmt,XXX.Customer_Code ,XXX.Customer_Name ,XXX.Created_By ,XXX.Modify_By ,XXX.Detail_Line_No , " &
            "XXX .Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2  from " &
            "(select distinct Location_Desc,Tin_No,PAN,(final.Account_Code),final.Account_Desc ,final.DrAmt ,final.CrAmt ,final.Document_No ,final.Document_Date, " &
            "final.Status ,final.Document_Type ,final.Account_Set ,final.DocAmt,final.Customer_Code ,final.Customer_Name ,final.Created_By , " &
            "final.Modify_By ,final.Detail_Line_No ,final .Comp_Code,Description  from " &
            "(select Location_Desc,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.PAN,TSPL_Customer_Invoice_Head.Description,case  when TSPL_JOURNAL_DETAILS.Amount >=0 then TSPL_JOURNAL_DETAILS.Amount else 0 end as DrAmt , " &
            "case  when TSPL_JOURNAL_DETAILS.Amount <0 then TSPL_JOURNAL_DETAILS.Amount*-1 else 0 end as CrAmt, " &
            "TSPL_Customer_Invoice_Head.Document_No, Document_Date , " &
            "case when TSPL_Customer_Invoice_Head.Status=1 then 'Authorized' else 'UnAuthorized' end as Status , " &
            "case when Document_Type='I' then 'Invoice' else case when Document_Type='D' then 'Debit' else " &
            "case when Document_Type='C' then 'Credit' else '' end end end as Document_Type,Account_Set,Document_Total as DocAmt,  " &
            "Customer_Code ,TSPL_Customer_Invoice_Head.Customer_Name,TSPL_Customer_Invoice_Head.Created_By ,TSPL_Customer_Invoice_Head.Modify_By  , " &
            "TSPL_JOURNAL_DETAILS.Detail_Line_No as Detail_Line_No ,TSPL_JOURNAL_DETAILS.Account_code as Account_Code , " &
            "TSPL_JOURNAL_DETAILS.Account_Desc as Account_Desc ,TSPL_Customer_Invoice_Detail.Amount , " &
            "TSPL_Customer_Invoice_Detail.Discount ,TSPL_Customer_Invoice_Detail.Amount_less_Discount , " &
            "TSPL_Customer_Invoice_Detail .Total_Tax ,TSPL_Customer_Invoice_Detail.Total_Amount  , " &
            "TSPL_Customer_Invoice_Head.Comp_Code    from TSPL_Customer_Invoice_Head left outer join " &
            "TSPL_Customer_Invoice_Detail on TSPL_Customer_Invoice_Detail.Document_No =TSPL_Customer_Invoice_Head.Document_No " &
            "left outer join TSPL_JOURNAL_MASTER on TSPL_JOURNAL_MASTER.Source_Doc_No =TSPL_Customer_Invoice_Head.Document_No " &
            "left outer join TSPL_JOURNAL_DETAILS on TSPL_JOURNAL_DETAILS.Journal_No = TSPL_JOURNAL_MASTER.Journal_No   " &
            "left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_Customer_Invoice_Head.Customer_Code  " &
            "left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_Customer_Invoice_Head.Loc_Code  " &
            "where TSPL_Customer_Invoice_Head.Document_No ='" + txtDocNo.Value + "' )final   )XXX left outer join " &
            "TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = XXX.Comp_Code order by XXX.Detail_Line_No  "

        Return qry

    End Function

    Private Sub funPrint()
        Throw New NotImplementedException
    End Sub

    Private Sub chkSecurityDposit_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkSecurityDposit.ToggleStateChanged
        ddlSecDepositType.SelectedValue = ""
        ddlSecDepositType.Enabled = chkSecurityDposit.Checked
    End Sub

    Private Sub mnuExportARInvoiceTrans_Click(sender As Object, e As EventArgs) Handles mnuExportARInvoiceTrans.Click
        Dim Sql As String = "select 'dd-MMM-yyyy' as [Date], '' as Customer,'' as LocSegment,'Exempted' as TaxGroup,'' as GLAccount1,0 as GLAmount1,'' as GLAccount2,0 as GLAmount2,'' as GLAccount3,0 as GLAmount3,'' as GLAccount4,0 as GLAmount4,'' as GLAccount5,0 as GLAmount5,'' as Remarks,'N' AS [Service Invoice],'' AS [Additional Charge Code1],'' AS [Additional Charge Code2],'' AS [Additional Charge Code3],'' AS [Additional Charge Code4],'' AS [Additional Charge Code5]"
        If SettingCostCenter Then
            Sql += ",'' AS [Hirerachy Level]"
            If SettingCostCenterlevel Then
                Sql += ",'' AS [" + gv1.Columns(colHierarchyLevel4).HeaderText + "],'' AS [" + gv1.Columns(colHierarchyLevel3).HeaderText + "],'' AS [" + gv1.Columns(colHierarchyLevel2).HeaderText + "],'' AS [" + gv1.Columns(colHierarchyLevel1).HeaderText + "],'N' AS [Service Invoice],'' AS [Additional Charge Code1],'' AS [Additional Charge Code2],'' AS [Additional Charge Code3],'' AS [Additional Charge Code4],'' AS [Additional Charge Code5] "
            Else
                Sql += ",'' AS [Cost Centre],'N' AS [Service Invoice],'' AS [Additional Charge Code1],'' AS [Additional Charge Code2],'' AS [Additional Charge Code3],'' AS [Additional Charge Code4],'' AS [Additional Charge Code5]"
            End If
        End If
        transportSql.ExporttoExcel(Sql, Me)
    End Sub

    Private Sub mnuExportCN_Click(sender As Object, e As EventArgs) Handles mnuExportCN.Click
        ExportAR_CN_DN()
    End Sub

    Private Sub mnuExportDN_Click(sender As Object, e As EventArgs) Handles mnuExportDN.Click
        ExportAR_CN_DN()
    End Sub

    Sub ExportAR_CN_DN()
        Try
            Dim strQry As String
            strQry = "Select '' as [Location], '' as [Customer], '' as [Document Date], '' as [Account Code], 0 as [Amount], 0 as [Discount %],'N' as [Security Deposit],'' as [Security Deposit Type],'' as Remarks"
            If SettingCostCenter Then
                strQry += ",'' AS [Hirerachy Level]"
                If SettingCostCenterlevel Then
                    strQry += ",'' AS [" + gv1.Columns(colHierarchyLevel4).HeaderText + "],'' AS [" + gv1.Columns(colHierarchyLevel3).HeaderText + "],'' AS [" + gv1.Columns(colHierarchyLevel2).HeaderText + "],'' AS [" + gv1.Columns(colHierarchyLevel1).HeaderText + "] "
                Else
                    strQry += ",'' AS [Cost Centre]"
                End If
            End If
            strQry += ",'N' as [Service Invoice],'' as [Additional Charge Code]"
            transportSql.ExporttoExcel(strQry, Me)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub mnuImportDN_Click(sender As Object, e As EventArgs) Handles mnuImportDN.Click
        cboDocType.SelectedValue = "D"
        If clsCommon.MyMessageBoxShow(Me, "You are going to import " & IIf(clsCommon.myCstr(cboDocType.SelectedValue) = "I", "Invoice", IIf(clsCommon.myCstr(cboDocType.SelectedValue) = "D", "Debit Note", "Credit Note")) & Environment.NewLine & "want to continue ?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If
        ImporAR_CN_DN("D")
    End Sub

    Private Sub mnuImportCR_Click(sender As Object, e As EventArgs) Handles mnuImportCR.Click
        cboDocType.SelectedValue = "C"
        If clsCommon.MyMessageBoxShow(Me, "You are going to import " & IIf(clsCommon.myCstr(cboDocType.SelectedValue) = "I", "Invoice", IIf(clsCommon.myCstr(cboDocType.SelectedValue) = "D", "Debit Note", "Credit Note")) & Environment.NewLine & "want to continue ?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If
        ImporAR_CN_DN("C")

    End Sub

    Sub ImporAR_CN_DN(ByVal CN_DN As String)
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim Counter As String = ""
        Dim qry As String
        Dim IsLevel As Boolean = False
        Dim flag As Boolean = False
        If SettingCostCenter Then
            If SettingCostCenterlevel Then
                flag = transportSql.importExcel(gv, "Location", "Customer", "Document Date", "Account Code", "Amount", "Discount %", "Security Deposit", "Security Deposit Type", "Remarks", "Hirerachy Level", gv1.Columns(colHierarchyLevel4).HeaderText, gv1.Columns(colHierarchyLevel3).HeaderText, gv1.Columns(colHierarchyLevel2).HeaderText, gv1.Columns(colHierarchyLevel1).HeaderText, "Service Invoice", "Additional Charge Code")
            Else
                flag = transportSql.importExcel(gv, "Location", "Customer", "Document Date", "Account Code", "Amount", "Discount %", "Security Deposit", "Security Deposit Type", "Remarks", "Hirerachy Level", "Cost Centre", "Service Invoice", "Additional Charge Code")
            End If
        Else
            flag = transportSql.importExcel(gv, "Location", "Customer", "Document Date", "Account Code", "Amount", "Discount %", "Security Deposit", "Security Deposit Type", "Remarks", "Service Invoice", "Additional Charge Code")
        End If

        If flag Then
            Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
            clsCommon.ProgressBarShow()
            Try
                Dim obj As clsCustomerInvoiceHead
                Dim dtTemp As DataTable
                For Each grow As GridViewRowInfo In gv.Rows
                    Counter = clsCommon.myCstr(grow.Index + 2) + " :"
                    obj = New clsCustomerInvoiceHead()
                    obj.Customer_Code = clsCommon.myCstr(grow.Cells("Customer").Value)
                    If clsCommon.myLen(obj.Customer_Code) > 0 Then
                        obj.Customer_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Code from TSPL_Customer_MASTER WHERE Cust_Code='" & obj.Customer_Code & "'", trans))
                        If clsCommon.myLen(obj.Customer_Code) <= 0 Then
                            Throw New Exception("Customer does not exist.")
                        End If
                    Else
                        Throw New Exception("Please enter Customer Code")
                    End If
                    Dim objCust As clsCustomerMaster = clsCustomerMaster.GetData(obj.Customer_Code, trans)
                    obj.Account_Set = objCust.Cust_Account
                    obj.Customer_Name = objCust.Customer_Name
                    obj.Terms_Code = objCust.Terms_Code
                    If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code, trans) = True Then
                        obj.CURRENCY_CODE = objCust.CURRENCY_CODE
                    End If
                    If clsCommon.myLen(obj.CURRENCY_CODE) > 0 Then
                        dtTemp = clsModuleCurrencyMapping.GetLatestCurConvRateDT(Me.txtDate.Value, txtCurrencyCode.Value)
                        If dtTemp.Rows.Count > 0 Then
                            obj.ConvRate = clsCommon.myCdbl(dtTemp.Rows(0).Item("Rate"))
                        Else
                            obj.ConvRate = 1
                        End If
                    End If
                    obj.loc_code = clsCommon.myCstr(grow.Cells("Location").Value)
                    If clsCommon.myLen(obj.loc_code) > 0 Then
                        obj.loc_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct(Segment_code) as Code from TSPL_GL_SEGMENT_CODE left outer join TSPL_LOCATION_MASTER on TSPL_GL_SEGMENT_CODE .Segment_code =TSPL_LOCATION_MASTER .Loc_Segment_Code WHERE TSPL_GL_SEGMENT_CODE.Segment_code='" + obj.loc_code + "' AND Seg_No = '7' AND GIT='N'", trans))
                        If clsCommon.myLen(obj.loc_code) <= 0 Then
                            Throw New Exception("Location Segment does not exist.")
                        End If
                    Else
                        Throw New Exception("Please enter Location Segment Code")
                    End If

                    obj.Document_Date = clsCommon.myCstr(grow.Cells("Document Date").Value)
                    'obj.Vendor_Invoice_No = "" '"Opening Balance " + clsCommon.myCstr(Counter)
                    'obj.Vendor_Invoice_Date = obj.Invoice_Entry_Date
                    obj.Due_Date = obj.Document_Date
                    obj.Terms_Code = objCust.Terms_Code
                    obj.Document_Type = clsCommon.myCstr(cboDocType.SelectedValue)
                    obj.On_Hold = False
                    obj.Discount_Base = clsCommon.myCdbl(grow.Cells("Amount").Value)
                    If obj.Discount_Base <= 0 Then
                        Throw New Exception("Please enter amount.")
                    End If
                    obj.Discount_Percentage = clsCommon.myCdbl(grow.Cells("Discount %").Value)

                    dtTemp = clsDBFuncationality.GetDataTable("select Receivable_Control_acct,Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Account=TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account WHERE TSPL_CUSTOMER_MASTER.Cust_Code='" + obj.Customer_Code + "'", trans)
                    If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                        obj.Customer_Control_AC = clsCommon.myCstr(dtTemp.Rows(0)("Receivable_Control_acct"))
                        obj.Discount_GL_AC = clsCommon.myCstr(dtTemp.Rows(0)("Receipts_Discount_acct"))
                    End If
                    If obj.Discount_Percentage > 0 Then
                        obj.Discount_Amount = (obj.Discount_Base * obj.Discount_Percentage) / 100
                    Else
                        obj.Discount_Amount = 0
                    End If
                    obj.Document_Type = CN_DN '"D"
                    obj.Amount_Less_Discount = obj.Discount_Base - obj.Discount_Amount
                    obj.Document_Total = obj.Amount_Less_Discount
                    obj.Balance_Amt = obj.Amount_Less_Discount
                    'obj.Total_Landed_Amt = 0
                    'obj.is_For_Provision = 0
                    'obj.isDeduction = IIf(clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Deduction (Yes/No)").Value), "Yes") = CompairStringResult.Equal, 1, 0)
                    'If obj.isDeduction = 1 Then
                    '    obj.Security = IIf(clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Security Deduction (Yes/No)").Value), "Yes") = CompairStringResult.Equal, 1, 0)
                    'Else
                    '    obj.Security = 0
                    'End If
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Security Deposit").Value), "Y") = CompairStringResult.Equal Then
                        obj.SecurityDeposit = True
                    Else
                        obj.SecurityDeposit = False
                    End If

                    obj.SecurityDepositType = clsCommon.myCstr(grow.Cells("Security Deposit Type").Value)
                    'obj.PO_Number = ""
                    obj.Description = clsCommon.myCstr(grow.Cells("Remarks").Value) '"Opening Balance as On " + obj.Invoice_Entry_Date + " for Vendor " + obj.Vendor_Code + ""
                    obj.Tax_Group = "Exempted"
                    obj.PROJECT_ID = ""

                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic

                    obj.Total_Add_Charge = 0
                    'obj.Empty_Amount = 0
                    'obj.Is_ProRated = "N"
                    ''richa agarwal 14 October,2019 ERO/11/10/19-001050
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Service Invoice").Value), "Y") = CompairStringResult.Equal Then
                        obj.AgainstServiceInvoice = "Y"
                    Else
                        obj.AgainstServiceInvoice = "N"
                    End If
                    '-------------------------Detail Level Starts From Here----------------
                    obj.Arr = New List(Of clsCustomerInvoiceDetail)
                    Dim objTr As New clsCustomerInvoiceDetail()
                    '"Deduction Code", "Account Code", "Amount", "Discount %", "Discount Amount", "Remarks"
                    objTr.SNo = 1
                    objTr.GL_Account_Code = clsCommon.myCstr(grow.Cells("Account Code").Value)
                    If clsCommon.myLen(objTr.GL_Account_Code) > 0 Then
                        dtTemp = clsDBFuncationality.GetDataTable("select TSPL_GL_ACCOUNTS.Account_Code, TSPL_GL_ACCOUNTS.Description as Account_Desc from TSPL_GL_ACCOUNTS WHERE TSPL_GL_ACCOUNTS.Account_Code='" & objTr.GL_Account_Code & "'", trans)
                        If dtTemp.Rows.Count <= 0 Then
                            Throw New Exception("Account does not exist.")
                        Else
                            objTr.GL_Account_Code = clsCommon.myCstr(dtTemp.Rows(0)("Account_Code"))
                            objTr.GL_Account_Desc = clsCommon.myCstr(dtTemp.Rows(0)("Account_Desc"))
                        End If
                        If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Service Invoice").Value), "N") = CompairStringResult.Equal Then
                            If clsCommon.GetDateWithStartTime(clsCommon.myCDate(obj.Document_Date)) < clsCommon.GetDateWithStartTime(ERPStartDate) Then
                                'richa 17 SEp,2019 TEC/03/07/19-000927
                                dtTemp = clsDBFuncationality.GetDataTable("select TSPL_GL_ACCOUNTS.Account_Code, TSPL_GL_ACCOUNTS.Description as Account_Desc from TSPL_GL_ACCOUNTS WHERE TSPL_GL_ACCOUNTS.Account_Code='" & objTr.GL_Account_Code & "' and ( ControlAccount='N' or TSPL_GL_ACCOUNTS.Account_Code IN (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForAR  =1)) ", trans)
                                If dtTemp.Rows.Count <= 0 Then
                                    Throw New Exception("Account should be type of non Control Account.")
                                Else
                                    objTr.GL_Account_Code = clsCommon.myCstr(dtTemp.Rows(0)("Account_Code"))
                                    objTr.GL_Account_Desc = clsCommon.myCstr(dtTemp.Rows(0)("Account_Desc"))
                                End If
                            End If
                        End If
                    Else
                        Throw New Exception("Please enter Account code")
                    End If
                    Dim strGLType As String = clsPaymentHeader.CheckGLAccountType(objTr.GL_Account_Code, trans)
                    If SettingCostCenter AndAlso Not clsCommon.CompairString(strGLType, "Balance Sheet") = CompairStringResult.Equal Then
                        objTr.Hirerachy_Code = clsCommon.myCstr(grow.Cells("Hirerachy Level").Value)
                        If clsCommon.myLen(objTr.Hirerachy_Code) <= 0 Then
                            Throw New Exception("Hierarchy code not found")
                        End If
                        objTr.Hirerachy_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Hirerachy_Code As Row from TSPL_HIRERACHY_LEVEL_MASTER where Hirerachy_Code='" + objTr.Hirerachy_Code + "'", trans))
                        If clsCommon.myLen(objTr.Hirerachy_Code) <= 0 Then
                            Throw New Exception("Invalid Hierarchy code-" + clsCommon.myCstr(grow.Cells("Hirerachy Level").Value))
                        End If

                        If SettingCostCenterlevel Then
                            objTr.Hirerachy_Code4 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel4).HeaderText).Value)
                            objTr.Hirerachy_Code3 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel3).HeaderText).Value)
                            objTr.Hirerachy_Code2 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel2).HeaderText).Value)
                            objTr.Hirerachy_Code1 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel1).HeaderText).Value)
                            If clsCommon.myLen(objTr.Hirerachy_Code4) > 0 Then
                                objTr.Hirerachy_Code4 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code4 + "'", trans))
                            End If
                            If clsCommon.myLen(objTr.Hirerachy_Code3) > 0 Then
                                objTr.Hirerachy_Code3 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code3 + "'", trans))
                            End If
                            If clsCommon.myLen(objTr.Hirerachy_Code2) > 0 Then
                                objTr.Hirerachy_Code2 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code2 + "'", trans))
                            End If
                            If clsCommon.myLen(objTr.Hirerachy_Code1) > 0 Then
                                objTr.Hirerachy_Code1 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code1 + "'", trans))
                            End If
                            Dim lvl As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select ISNULL(Level,0) AS Level from TSPL_HIRERACHY_LEVEL_MASTER Where HIRERACHY_CODE='" + objTr.Hirerachy_Code + "' ", trans))
                            If lvl > 3 Then
                                If clsCommon.myLen(objTr.Hirerachy_Code4) <= 0 Then
                                    Throw New Exception("Hierarchy Level 4 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                If clsCommon.myLen(objTr.Hirerachy_Code3) <= 0 Then
                                    Throw New Exception("Hierarchy Level 3 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                If clsCommon.myLen(objTr.Hirerachy_Code2) <= 0 Then
                                    Throw New Exception("Hierarchy Level 2 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                    Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                qry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=4 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code4 + "' and HIRERACHY_LEVEL_CODE3='" + objTr.Hirerachy_Code3 + "' and HIRERACHY_LEVEL_CODE2='" + objTr.Hirerachy_Code2 + "' and HIRERACHY_LEVEL_CODE1='" + objTr.Hirerachy_Code1 + "'"
                                dtTemp = clsDBFuncationality.GetDataTable(qry, trans)
                                If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                    Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code4 + "][" + objTr.Hirerachy_Code3 + "][" + objTr.Hirerachy_Code2 + "][" + objTr.Hirerachy_Code1 + "] ")
                                End If
                                objTr.Cost_Centre_Code = objTr.Hirerachy_Code4
                            ElseIf lvl > 2 Then
                                If clsCommon.myLen(objTr.Hirerachy_Code3) <= 0 Then
                                    Throw New Exception("Hierarchy Level 3 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                If clsCommon.myLen(objTr.Hirerachy_Code2) <= 0 Then
                                    Throw New Exception("Hierarchy Level 2 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                    Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                qry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=3 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code3 + "' and HIRERACHY_LEVEL_CODE2='" + objTr.Hirerachy_Code2 + "' and HIRERACHY_LEVEL_CODE1='" + objTr.Hirerachy_Code1 + "'"
                                dtTemp = clsDBFuncationality.GetDataTable(qry, trans)
                                If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                    Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code3 + "][" + objTr.Hirerachy_Code2 + "][" + objTr.Hirerachy_Code1 + "] ")
                                End If
                                objTr.Cost_Centre_Code = objTr.Hirerachy_Code3
                            ElseIf lvl > 1 Then
                                If clsCommon.myLen(objTr.Hirerachy_Code2) <= 0 Then
                                    Throw New Exception("Hierarchy Level 2 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                    Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                qry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=2 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code2 + "' and HIRERACHY_LEVEL_CODE1='" + objTr.Hirerachy_Code1 + "'"
                                dtTemp = clsDBFuncationality.GetDataTable(qry, trans)
                                If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                    Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code2 + "][" + objTr.Hirerachy_Code1 + "] ")
                                End If
                                objTr.Cost_Centre_Code = objTr.Hirerachy_Code2
                            ElseIf lvl > 0 Then
                                If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                    Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                End If
                                qry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=1 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code1 + "'"
                                dtTemp = clsDBFuncationality.GetDataTable(qry, trans)
                                If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                    Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code1 + "]")
                                End If
                                objTr.Cost_Centre_Code = objTr.Hirerachy_Code1
                            End If
                        Else
                            objTr.Cost_Centre_Code = clsCommon.myCstr(grow.Cells("Cost Centre").Value)
                            If clsCommon.myLen(objTr.Cost_Centre_Code) <= 0 Then
                                Throw New Exception("Cost center not found")
                            End If
                            objTr.Cost_Centre_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Cost_Centre_Code + "'", trans))
                            If clsCommon.myLen(objTr.Hirerachy_Code) <= 0 Then
                                Throw New Exception("Invalid Cost center-" + clsCommon.myCstr(grow.Cells("Cost Centre").Value))
                            End If
                        End If
                        If clsCommon.myLen(objTr.Cost_Centre_Code) <= 0 Then
                            Throw New Exception(" Please select Cost Center of GL Account -" + objTr.GL_Account_Code)
                        End If
                    End If


                    ''richa agarwal 14 October,2019 ERO/11/10/19-001050
                    If clsCommon.CompairString(obj.AgainstServiceInvoice, "Y") = CompairStringResult.Equal Then
                        objTr.AddChargeCode = clsCommon.myCstr(grow.Cells("Additional Charge Code").Value)
                        If clsCommon.myLen(objTr.AddChargeCode) <= 0 Then
                            Throw New Exception("Additional Charge Code not found")
                        End If
                        objTr.AddChargeCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Code ,Description,abatement,specification from TSPL_Additional_Charges where TSPL_ADDITIONAL_CHARGES.Service_Type ='Y' and  TSPL_Additional_Charges.Code ='" & objTr.AddChargeCode & "'", trans))
                        If clsCommon.myLen(objTr.AddChargeCode) <= 0 Then
                            Throw New Exception("Invalid Additional Charge Code-" + clsCommon.myCstr(grow.Cells("Additional Charge Code").Value))
                        End If

                        objTr.AddChargeDesc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_Additional_Charges where TSPL_ADDITIONAL_CHARGES.Service_Type ='Y' and  TSPL_Additional_Charges.Code ='" & objTr.AddChargeCode & "'", trans))
                    End If

                    objTr.Amount = obj.Discount_Base
                    objTr.Discount_Per = obj.Discount_Percentage
                    objTr.Discount = obj.Discount_Amount
                    objTr.Amount_less_Discount = obj.Discount_Base - obj.Discount_Amount
                    'objTr.Landed_Amount = 0
                    objTr.Total_Tax = 0
                    objTr.Total_Amount = obj.Amount_Less_Discount
                    'objTr.AddChargeCode = ""
                    'objTr.AddChargeDesc = ""
                    'objTr.is_Unclaimed_Tax = False
                    objTr.Remarks = "" 'clsCommon.myCstr(grow.Cells("Remarks").Value)
                    'objTr.Invoice_Type = ""
                    'objTr.Invoice_No = ""
                    '---------------------------------Detail Level Ends Here-----------------------------
                    If (clsCommon.myLen(objTr.GL_Account_Code) > 0) And objTr.Amount <> 0 Then
                        obj.Arr.Add(objTr)
                    End If
                    obj.SaveData(obj, True, trans, Me.Form_ID)
                Next
                clsCommon.ProgressBarHide()
                trans.Commit()
                common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                trans.Rollback()
                clsCommon.ProgressBarHide()
                common.clsCommon.MyMessageBoxShow(Me, "Error at Rowno " + Counter + Environment.NewLine + ex.Message, Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
            Finally
                clsCommon.ProgressBarHide()
                Me.Controls.Remove(gv)
            End Try
        End If
    End Sub

    Private Sub mnuExportOpeningCredit_Click(sender As Object, e As EventArgs) Handles mnuExportOpeningCredit.Click
        Try
            Dim strQry As String
            strQry = "select '' as [Date], '' as Customer, 0 as Amount,'' as LocSegment"
            transportSql.ExporttoExcel(strQry, Me)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub mnuExortOpeningDebit_Click(sender As Object, e As EventArgs) Handles mnuExortOpeningDebit.Click
        Try
            Dim strQry As String
            strQry = "select '' as [Date], '' as Customer, 0 as Amount,'' as LocSegment"
            transportSql.ExporttoExcel(strQry, Me)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub mnuARInvoiceImport_Click(sender As Object, e As EventArgs) Handles mnuARInvoiceImport.Click
        ARInvoiceImport()
    End Sub

    Sub ARInvoiceImport()
        cboDocType.SelectedValue = "I"
        If clsCommon.MyMessageBoxShow(Me, "You are going to import " & IIf(clsCommon.myCstr(cboDocType.SelectedValue) = "I", "Invoice", IIf(clsCommon.myCstr(cboDocType.SelectedValue) = "D", "Debit Note", "Credit Note")) & Environment.NewLine & "want to continue ?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If
        Dim strQry As String
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim Counter As String = ""

        Dim flag As Boolean = False
        If SettingCostCenter Then
            If SettingCostCenterlevel Then
                flag = transportSql.importExcel(gv, "Date", "Customer", "LocSegment", "TaxGroup", "GLAccount1", "GLAmount1", "GLAccount2", "GLAmount2", "GLAccount3", "GLAmount3", "GLAccount4", "GLAmount4", "GLAccount5", "GLAmount5", "Remarks", gv1.Columns(colHierarchyLevel4).HeaderText, gv1.Columns(colHierarchyLevel3).HeaderText, gv1.Columns(colHierarchyLevel2).HeaderText, gv1.Columns(colHierarchyLevel1).HeaderText, "Service Invoice", "Additional Charge Code1", "Additional Charge Code2", "Additional Charge Code3", "Additional Charge Code4", "Additional Charge Code5")
            Else
                flag = transportSql.importExcel(gv, "Date", "Customer", "LocSegment", "TaxGroup", "GLAccount1", "GLAmount1", "GLAccount2", "GLAmount2", "GLAccount3", "GLAmount3", "GLAccount4", "GLAmount4", "GLAccount5", "GLAmount5", "Remarks", "Hirerachy Level", "Cost Centre", "Service Invoice", "Additional Charge Code1", "Additional Charge Code2", "Additional Charge Code3", "Additional Charge Code4", "Additional Charge Code5")
            End If
        Else
            flag = transportSql.importExcel(gv, "Date", "Customer", "LocSegment", "TaxGroup", "GLAccount1", "GLAmount1", "GLAccount2", "GLAmount2", "GLAccount3", "GLAmount3", "GLAccount4", "GLAmount4", "GLAccount5", "GLAmount5", "Remarks", "Service Invoice", "Additional Charge Code1", "Additional Charge Code2", "Additional Charge Code3", "Additional Charge Code4", "Additional Charge Code5")
        End If
        If flag Then
            Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
            clsCommon.ProgressBarShow()
            Try
                Dim obj As clsCustomerInvoiceHead
                Dim dtTemp As DataTable
                For Each grow As GridViewRowInfo In gv.Rows
                    Counter = clsCommon.myCstr(grow.Index + 2) + " :"
                    obj = New clsCustomerInvoiceHead()
                    obj.Customer_Code = clsCommon.myCstr(grow.Cells("Customer").Value)
                    If clsCommon.myLen(obj.Customer_Code) > 0 Then
                        obj.Customer_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Code from TSPL_CUSTOMER_MASTER WHERE Cust_Code='" & obj.Customer_Code & "'", trans))
                        If clsCommon.myLen(obj.Customer_Code) <= 0 Then
                            Throw New Exception("Customer does not exist.")
                        End If
                    Else
                        Throw New Exception("Please enter Customer Code")
                    End If
                    Dim objVendor As clsCustomerMaster = clsCustomerMaster.GetData(obj.Customer_Code, trans)
                    obj.Account_Set = objVendor.Cust_Account
                    obj.Customer_Name = objVendor.Customer_Name
                    obj.Terms_Code = objVendor.Terms_Code
                    If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code, trans) = True Then
                        obj.CURRENCY_CODE = objVendor.CURRENCY_CODE
                    End If
                    If clsCommon.myLen(obj.CURRENCY_CODE) > 0 Then
                        dtTemp = clsModuleCurrencyMapping.GetLatestCurConvRateDT(Me.txtDate.Value, txtCurrencyCode.Value)
                        If dtTemp.Rows.Count > 0 Then
                            obj.ConvRate = clsCommon.myCdbl(dtTemp.Rows(0).Item("Rate"))
                        Else
                            obj.ConvRate = 1
                        End If
                    End If
                    obj.loc_code = clsCommon.myCstr(grow.Cells("LocSegment").Value)
                    If clsCommon.myLen(obj.loc_code) > 0 Then
                        obj.loc_code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct(Segment_code) as Code from TSPL_GL_SEGMENT_CODE left outer join TSPL_LOCATION_MASTER on TSPL_GL_SEGMENT_CODE .Segment_code =TSPL_LOCATION_MASTER .Loc_Segment_Code WHERE TSPL_GL_SEGMENT_CODE.Segment_code='" + obj.loc_code + "' AND Seg_No = '7' AND GIT='N'", trans))
                        If clsCommon.myLen(obj.loc_code) <= 0 Then
                            Throw New Exception("Location Segment does not exist.")
                        End If
                    Else
                        Throw New Exception("Please enter Location Segment Code")
                    End If


                    obj.Document_Date = clsCommon.myCstr(grow.Cells("Date").Value)
                    obj.RefDocNo = "" 'clsCommon.myCstr(grow.Cells("ARInvoiceNo").Value)
                    'obj.Vendor_Invoice_Date = obj.Invoice_Entry_Date
                    obj.Due_Date = obj.Document_Date
                    obj.Terms_Code = objVendor.Terms_Code
                    obj.Document_Type = clsCommon.myCstr(cboDocType.SelectedValue)
                    obj.On_Hold = False
                    obj.Discount_Base = clsCommon.myCdbl(grow.Cells("GLAmount1").Value) + clsCommon.myCdbl(grow.Cells("GLAmount2").Value) + clsCommon.myCdbl(grow.Cells("GLAmount3").Value) + clsCommon.myCdbl(grow.Cells("GLAmount4").Value) + clsCommon.myCdbl(grow.Cells("GLAmount5").Value)
                    If obj.Discount_Base <= 0 Then
                        Throw New Exception("Please enter amount.")
                    End If
                    'obj.Discount_Percentage = clsCommon.myCdbl(grow.Cells("Discount %").Value)

                    dtTemp = clsDBFuncationality.GetDataTable("select Receivable_Control_acct,Receipts_Discount_acct from TSPL_CUSTOMER_ACCOUNT_SET LEFT OUTER JOIN TSPL_CUSTOMER_MASTER ON TSPL_CUSTOMER_MASTER.Cust_Account=TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account WHERE TSPL_CUSTOMER_MASTER.Cust_Code='" + obj.Customer_Code + "'", trans)
                    If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                        obj.Customer_Control_AC = clsCommon.myCstr(dtTemp.Rows(0)("Receivable_Control_acct"))
                        obj.Discount_GL_AC = clsCommon.myCstr(dtTemp.Rows(0)("Receipts_Discount_acct"))
                    End If
                    If obj.Discount_Percentage > 0 Then
                        obj.Discount_Amount = (obj.Discount_Base * obj.Discount_Percentage) / 100
                    Else
                        obj.Discount_Amount = 0
                    End If
                    obj.Document_Type = "I"
                    obj.Amount_Less_Discount = obj.Discount_Base - obj.Discount_Amount
                    obj.Document_Total = obj.Amount_Less_Discount
                    obj.Balance_Amt = obj.Amount_Less_Discount
                    obj.Description = clsCommon.myCstr(grow.Cells("Remarks").Value)
                    obj.Tax_Group = "Exempted"
                    obj.PROJECT_ID = ""
                    obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic
                    obj.Total_Add_Charge = 0
                    ''richa agarwal 14 October,2019 ERO/11/10/19-001050
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Service Invoice").Value), "Y") = CompairStringResult.Equal Then
                        obj.AgainstServiceInvoice = "Y"
                    Else
                        obj.AgainstServiceInvoice = "N"
                    End If
                    obj.Arr = New List(Of clsCustomerInvoiceDetail)
                    Dim objTr As New clsCustomerInvoiceDetail()
                    For CounterForAccount As Integer = 1 To 5
                        If clsCommon.myLen(grow.Cells("GLAccount" + clsCommon.myCstr(1)).Value) > 0 Then
                            objTr = New clsCustomerInvoiceDetail
                            objTr.SNo = obj.Arr.Count + 1
                            objTr.GL_Account_Code = clsCommon.myCstr(grow.Cells("GLAccount" + clsCommon.myCstr(CounterForAccount)).Value)
                            If clsCommon.myLen(objTr.GL_Account_Code) > 0 Then
                                dtTemp = clsDBFuncationality.GetDataTable("select TSPL_GL_ACCOUNTS.Account_Code, TSPL_GL_ACCOUNTS.Description as Account_Desc from TSPL_GL_ACCOUNTS WHERE TSPL_GL_ACCOUNTS.Account_Code='" & objTr.GL_Account_Code & "'", trans)
                                If dtTemp.Rows.Count <= 0 Then
                                    Throw New Exception("Account does not exist.")
                                Else
                                    objTr.GL_Account_Code = clsCommon.myCstr(dtTemp.Rows(0)("Account_Code"))
                                    objTr.GL_Account_Desc = clsCommon.myCstr(dtTemp.Rows(0)("Account_Desc"))
                                End If
                                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Service Invoice").Value), "N") = CompairStringResult.Equal Then
                                    If clsCommon.GetDateWithStartTime(obj.Document_Date) >= clsCommon.GetDateWithStartTime(ERPStartDate) Then
                                        'richa 17 SEp,2019 TEC/03/07/19-000927
                                        dtTemp = clsDBFuncationality.GetDataTable("select TSPL_GL_ACCOUNTS.Account_Code, TSPL_GL_ACCOUNTS.Description as Account_Desc from TSPL_GL_ACCOUNTS WHERE TSPL_GL_ACCOUNTS.Account_Code='" & objTr.GL_Account_Code & "' and ( ControlAccount='N' or TSPL_GL_ACCOUNTS.Account_Code IN  (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForAR  =1)) ", trans)
                                        If dtTemp.Rows.Count <= 0 Then
                                            Throw New Exception("Account should be type of non Control Account.")
                                        Else
                                            objTr.GL_Account_Code = clsCommon.myCstr(dtTemp.Rows(0)("Account_Code"))
                                            objTr.GL_Account_Desc = clsCommon.myCstr(dtTemp.Rows(0)("Account_Desc"))
                                        End If
                                    End If
                                End If

                                ''richa agarwal 14 October,2019 ERO/11/10/19-001050
                                If clsCommon.CompairString(obj.AgainstServiceInvoice, "Y") = CompairStringResult.Equal Then
                                    objTr.AddChargeCode = clsCommon.myCstr(grow.Cells("Additional Charge Code" + clsCommon.myCstr(CounterForAccount)).Value)
                                    If clsCommon.myLen(objTr.AddChargeCode) <= 0 Then
                                        Throw New Exception("Additional Charge Code" + clsCommon.myCstr(CounterForAccount) + " not found")
                                    End If
                                    objTr.AddChargeCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Code ,Description,abatement,specification from TSPL_Additional_Charges where TSPL_ADDITIONAL_CHARGES.Service_Type ='Y' and  TSPL_Additional_Charges.Code ='" & objTr.AddChargeCode & "'", trans))
                                    If clsCommon.myLen(objTr.AddChargeCode) <= 0 Then
                                        Throw New Exception("Invalid Additional Charge Code" + clsCommon.myCstr(CounterForAccount) + "-" + clsCommon.myCstr(grow.Cells("Additional Charge Code" + clsCommon.myCstr(CounterForAccount)).Value))
                                    End If

                                    objTr.AddChargeDesc = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_Additional_Charges where TSPL_ADDITIONAL_CHARGES.Service_Type ='Y' and  TSPL_Additional_Charges.Code ='" & objTr.AddChargeCode & "'", trans))
                                End If

                                Dim strGLType As String = clsPaymentHeader.CheckGLAccountType(objTr.GL_Account_Code, trans)
                                If SettingCostCenter AndAlso Not clsCommon.CompairString(strGLType, "Balance Sheet") = CompairStringResult.Equal Then
                                    objTr.Hirerachy_Code = clsCommon.myCstr(grow.Cells("Hirerachy Level").Value)
                                    If clsCommon.myLen(objTr.Hirerachy_Code) <= 0 Then
                                        Throw New Exception("Hierarchy code not found")
                                    End If
                                    objTr.Hirerachy_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Hirerachy_Code As Row from TSPL_HIRERACHY_LEVEL_MASTER where Hirerachy_Code='" + objTr.Hirerachy_Code + "'", trans))
                                    If clsCommon.myLen(objTr.Hirerachy_Code) <= 0 Then
                                        Throw New Exception("Invalid Hierarchy code-" + clsCommon.myCstr(grow.Cells("Hirerachy Level").Value))
                                    End If

                                    If SettingCostCenterlevel Then
                                        objTr.Hirerachy_Code4 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel4).HeaderText).Value)
                                        objTr.Hirerachy_Code3 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel3).HeaderText).Value)
                                        objTr.Hirerachy_Code2 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel2).HeaderText).Value)
                                        objTr.Hirerachy_Code1 = clsCommon.myCstr(grow.Cells(gv1.Columns(colHierarchyLevel1).HeaderText).Value)
                                        If clsCommon.myLen(objTr.Hirerachy_Code4) > 0 Then
                                            objTr.Hirerachy_Code4 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code4 + "'", trans))
                                        End If
                                        If clsCommon.myLen(objTr.Hirerachy_Code3) > 0 Then
                                            objTr.Hirerachy_Code3 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code3 + "'", trans))
                                        End If
                                        If clsCommon.myLen(objTr.Hirerachy_Code2) > 0 Then
                                            objTr.Hirerachy_Code2 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code2 + "'", trans))
                                        End If
                                        If clsCommon.myLen(objTr.Hirerachy_Code1) > 0 Then
                                            objTr.Hirerachy_Code1 = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Hirerachy_Code1 + "'", trans))
                                        End If
                                        Dim lvl As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select ISNULL(Level,0) AS Level from TSPL_HIRERACHY_LEVEL_MASTER Where HIRERACHY_CODE='" + objTr.Hirerachy_Code + "' ", trans))
                                        If lvl > 3 Then
                                            If clsCommon.myLen(objTr.Hirerachy_Code4) <= 0 Then
                                                Throw New Exception("Hierarchy Level 4 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            If clsCommon.myLen(objTr.Hirerachy_Code3) <= 0 Then
                                                Throw New Exception("Hierarchy Level 3 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            If clsCommon.myLen(objTr.Hirerachy_Code2) <= 0 Then
                                                Throw New Exception("Hierarchy Level 2 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                                Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            strQry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=4 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code4 + "' and HIRERACHY_LEVEL_CODE3='" + objTr.Hirerachy_Code3 + "' and HIRERACHY_LEVEL_CODE2='" + objTr.Hirerachy_Code2 + "' and HIRERACHY_LEVEL_CODE1='" + objTr.Hirerachy_Code1 + "'"
                                            dtTemp = clsDBFuncationality.GetDataTable(strQry, trans)
                                            If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                                Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code4 + "][" + objTr.Hirerachy_Code3 + "][" + objTr.Hirerachy_Code2 + "][" + objTr.Hirerachy_Code1 + "] ")
                                            End If
                                            objTr.Cost_Centre_Code = objTr.Hirerachy_Code4
                                        ElseIf lvl > 2 Then
                                            If clsCommon.myLen(objTr.Hirerachy_Code3) <= 0 Then
                                                Throw New Exception("Hierarchy Level 3 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            If clsCommon.myLen(objTr.Hirerachy_Code2) <= 0 Then
                                                Throw New Exception("Hierarchy Level 2 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                                Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            strQry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=3 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code3 + "' and HIRERACHY_LEVEL_CODE2='" + objTr.Hirerachy_Code2 + "' and HIRERACHY_LEVEL_CODE1='" + objTr.Hirerachy_Code1 + "'"
                                            dtTemp = clsDBFuncationality.GetDataTable(strQry, trans)
                                            If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                                Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code3 + "][" + objTr.Hirerachy_Code2 + "][" + objTr.Hirerachy_Code1 + "] ")
                                            End If
                                            objTr.Cost_Centre_Code = objTr.Hirerachy_Code3
                                        ElseIf lvl > 1 Then
                                            If clsCommon.myLen(objTr.Hirerachy_Code2) <= 0 Then
                                                Throw New Exception("Hierarchy Level 2 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                                Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            strQry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=2 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code2 + "' and HIRERACHY_LEVEL_CODE1='" + objTr.Hirerachy_Code1 + "'"
                                            dtTemp = clsDBFuncationality.GetDataTable(strQry, trans)
                                            If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                                Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code2 + "][" + objTr.Hirerachy_Code1 + "] ")
                                            End If
                                            objTr.Cost_Centre_Code = objTr.Hirerachy_Code2
                                        ElseIf lvl > 0 Then
                                            If clsCommon.myLen(objTr.Hirerachy_Code1) <= 0 Then
                                                Throw New Exception("Hierarchy Level 1 Code not found for GL Account -" + objTr.GL_Account_Code)
                                            End If
                                            strQry = "select 1 from TSPL_COST_CENTRE_HIRERACHY_DETAIL where HIRERACHY_LEVEL=1 and COST_CENTRE_HIRERACHY_CODE='" + objTr.Hirerachy_Code1 + "'"
                                            dtTemp = clsDBFuncationality.GetDataTable(strQry, trans)
                                            If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                                                Throw New Exception("Not a Valid Combition of Hierachy level [" + objTr.Hirerachy_Code1 + "]")
                                            End If
                                            objTr.Cost_Centre_Code = objTr.Hirerachy_Code1
                                        End If
                                    Else
                                        objTr.Cost_Centre_Code = clsCommon.myCstr(grow.Cells("Cost Centre").Value)
                                        If clsCommon.myLen(objTr.Cost_Centre_Code) <= 0 Then
                                            Throw New Exception("Cost center not found")
                                        End If
                                        objTr.Cost_Centre_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Code from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + objTr.Cost_Centre_Code + "'", trans))
                                        If clsCommon.myLen(objTr.Hirerachy_Code) <= 0 Then
                                            Throw New Exception("Invalid Cost center-" + clsCommon.myCstr(grow.Cells("Cost Centre").Value))
                                        End If
                                    End If
                                    If clsCommon.myLen(objTr.Cost_Centre_Code) <= 0 Then
                                        Throw New Exception("Please select Cost Center of GL Account -" + objTr.GL_Account_Code)
                                    End If
                                End If

                                objTr.Amount = clsCommon.myCdbl(grow.Cells("GLAmount" + clsCommon.myCstr(CounterForAccount)).Value)
                                objTr.Discount_Per = obj.Discount_Percentage
                                objTr.Discount = obj.Discount_Amount
                                objTr.Amount_less_Discount = objTr.Amount - obj.Discount_Amount
                                objTr.Total_Tax = 0
                                objTr.Total_Amount = obj.Amount_Less_Discount
                                objTr.Remarks = ""
                                If (clsCommon.myLen(objTr.GL_Account_Code) > 0) And objTr.Amount <> 0 Then
                                    obj.Arr.Add(objTr)
                                End If
                            End If
                        End If
                    Next
                    If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                        obj.SaveData(obj, True, trans, Me.Form_ID)
                    End If
                Next
                clsCommon.ProgressBarHide()
                trans.Commit()
                common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                trans.Rollback()
                clsCommon.ProgressBarHide()
                common.clsCommon.MyMessageBoxShow(Me, "Error at Rowno " + Counter + Environment.NewLine + ex.Message, Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error)
            Finally
                clsCommon.ProgressBarHide()
                Me.Controls.Remove(gv)
            End Try
        End If
    End Sub

    Private Sub txtSercurityReceiptNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtSercurityReceiptNo._MYValidating
        Try
            Dim Qry As String = clsRcptEntryHeader.GetAgainstSercurityQry(TxtCustomer.Value, txtDocNo.Value, "")
            Qry = "select * from ( " + Qry + " )xxx"
            txtSercurityReceiptNo.Value = clsCommon.ShowSelectForm("ARInvSecReceipt", Qry, "ReceiptNo", " Amount<>0 ", txtSercurityReceiptNo.Value, "", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkAgainstSecurityReceiptNo_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkAgainstSecurityReceiptNo.ToggleStateChanged
        PickAgainstSecurity()
    End Sub

    Sub PickAgainstSecurity()
        Panel2.Visible = chkAgainstSecurityReceiptNo.Checked
        If chkAgainstSecurityReceiptNo.Checked Then
            cboDocType.SelectedValue = "C"
        End If

        If chkAgainstSecurityReceiptNo.Checked AndAlso clsCommon.myLen(txtlocation.Value) > 0 AndAlso clsCommon.myLen(TxtCustomer.Value) > 0 AndAlso isNewEntry Then
            If gv1.Rows.Count > 0 Then
                If clsCommon.myLen(gv1.Rows(0).Cells(colACCode).Value) <= 0 Then
                    Try
                        Dim qry As String = "select TSPL_CUSTOMER_ACCOUNT_SET.SECURITY_ACCOUNT from TSPL_CUSTOMER_MASTER left outer join TSPL_CUSTOMER_ACCOUNT_SET on TSPL_CUSTOMER_ACCOUNT_SET.Cust_Account=TSPL_CUSTOMER_MASTER.Cust_Account where TSPL_CUSTOMER_MASTER.Cust_Code='" + TxtCustomer.Value + "'"
                        Dim strAC As String = clsDBFuncationality.getSingleValue(qry)
                        strAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strAC, txtlocation.Value, True, Nothing)
                        gv1.CurrentRow = gv1.Rows(0)
                        gv1.CurrentColumn = gv1.Columns(colACCode)
                        gv1.Rows(0).Cells(colACCode).Value = strAC
                        gv1.CurrentRow = gv1.Rows(0)
                        gv1.CurrentColumn = gv1.Columns(colAmt)
                        'gv1.Rows(0).Cells(colAmt).Value = clsCommon.myCdbl(txtProvAmt.Text)
                    Catch ex As Exception
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub chkAgainstServiceInvoice_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkAgainstServiceInvoice.ToggleStateChanged
        If chkAgainstServiceInvoice.Checked = True Then
            gv1.Columns(colAddChgCode).IsVisible = True
            gv1.Columns(colAddChgName).IsVisible = True
            gv1.Columns(colSACCode).IsVisible = True
            gv1.Columns(colSACName).IsVisible = True
        Else
            LoadBlankGridGL()
            gv1.Rows.AddNew()
            gv1.Columns(colAddChgCode).IsVisible = False
            gv1.Columns(colAddChgName).IsVisible = False
            gv1.Columns(colSACCode).IsVisible = False
            gv1.Columns(colSACName).IsVisible = False
        End If
    End Sub
    ' Ticket No : ERO/11/10/19-001051,ERO/23/09/19-001037 By Prabhakar for Service invoice print
    Private Sub btnPrintServiceInvoice_Click(sender As Object, e As EventArgs) Handles btnPrintServiceInvoice.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please select document first!", Me.Text)
                Return
            End If

            If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
                funPrint(txtDocNo.Value)
                Return
            End If


            'Dim ECustomerType = clsERPFuncationality.GetCustomerEInvoiceTypeFromTransationTable("TSPL_Customer_Invoice_Head", "Document_No", txtDocNo.Value, Nothing)
            Dim IsEInvoiceApply As Integer = 0
            If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal Then
                IsEInvoiceApply = 1
            End If
            clsReceiptInvoiceHead.funPrintServiceInvoicePrint(MyBase.Form_ID, Nothing, txtDate.Value, txtDocNo.Value, IsEInvoiceApply)

            'Dim qry As String = " Select * from ( 
            'Dim qry As String = "   Select  tspl_company_master.Access_Officer as FSSAI,cast(TSPL_Customer_Invoice_Head.BarCode_Img as image) As BarCode_Img,isnull (TSPL_Customer_Invoice_Head.IRN_No,'') as IRN_No,isnull (TSPL_Customer_Invoice_Head.Ack_No,'') as Ack_No,case when len(isnull (TSPL_Customer_Invoice_Head.Ack_No,'')) > 0 then convert (varchar, TSPL_Customer_Invoice_Head.Ack_Date,103) else ''  end as Ack_Date, " + clsCommon.myCstr(IsEInvoiceApply) + " as  IsEInvoiceApply,'1' as CopyType, TSPL_COMPANY_MASTER.TIN_NO , TSPL_COMPANY_MASTER.CST_LST ,TSPL_COMPANY_MASTER.Pan_No,tspl_location_master.PAN_NO as LocationPAN," &
            '                        " TSPL_COMPANY_MASTER.Add1 + case When TSPL_COMPANY_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_COMPANY_MASTER.Add2, 103) End + Case When TSPL_COMPANY_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_COMPANY_MASTER.Add3,103) end + case When TSPL_COMPANY_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.City_Code, 103) end+ Case When TSPL_COMPANY_MASTER.State='' Then '' else ', '+Convert(Varchar, TSPL_COMPANY_MASTER.State) end +  Case When TSPL_COMPANY_MASTER.Pincode='' Then '' Else ', '+ Convert(Varchar,TSPL_COMPANY_MASTER.Pincode, 103)  end  as CompAdd ," &
            '                        " " &
            '                        " tspl_state_master_For_Comp.GST_STATE_Code as Comp_GST_STATE_CODE , tspl_state_master_For_Comp.State_Name as Comp_State_Name,tspl_state_master_For_Comp.STATE_CODE as Comp_State_Code,TSPL_COMPANY_MASTER.Email as Comp_Email, TSPL_COMPANY_MASTER.CINNO as Comp_CINNO, TSPL_COMPANY_MASTER.GSTReg_No as Comp_GSTReg_No," &
            '                        " TSPL_COMPANY_MASTER.Add1 as Comp_Add1, TSPL_COMPANY_MASTER.Add2 as Comp_Add2, TSPL_COMPANY_MASTER.add3 as Comp_Add3," &
            '                        " " &
            '                        " XXX.Customer_Code ,XXX.Customer_Name ," &
            '                        " TSPL_CUSTOMER_MASTER.Add1 as Cust_Add1, TSPL_CUSTOMER_MASTER.Add2 as Cust_Add2, TSPL_CUSTOMER_MASTER .Add3 as Cust_Add3,TSPL_CUSTOMER_MASTER.GSTNO AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode,TSPL_STATE_MASTER.state_name AS Vendor_State_Name,xxx.Tin_No as Customer_Tin_No,XXX.PAN as Customer_PAN_NO," &
            '                        " " &
            '                        " XXX.Loc_Code,locAdd,XXX.Loc_Add1,XXX.Loc_Add2,XXX.Loc_Add3,XXX.Loc_City_Code,XXX.Loc_City_Name,XXX.Loc_State_Name,XXX.Loc_Pin_code,XXX.Location_Desc," &
            '                        " TSPL_LOCATION_MASTER.Email as Loc_Email,tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code,TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo," &
            '                        " " &
            '                        "  XXX.RefDocNo    , XXX.Description  ,XXX.Document_No ,Convert (varchar,XXX.Document_Date,103) as Document_Date ,XXX.Status , XXX.AddChargeCode ,XXX.AddChargeDesc,XXX.SAC_Code,XXX.SNo,XXX.Amount_less_Discount,XXX.Document_Type ,XXX.Account_Set ,XXX.DocAmt,XXX.Created_By ,XXX.Modify_By , XXX .Comp_Code,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Logo_Img ,TSPL_COMPANY_MASTER.Logo_Img2,XXX.Taxdesc1,XXX.Taxdesc2,XXX.Taxdesc3,XXX.Taxdesc4,XXX.Taxdesc5,XXX.Taxdesc6,XXX.Taxdesc7,XXX.Taxdesc8,XXX.Taxdesc9,XXX.Taxdesc10,XXX.TAX1_Amt,XXX.TAX2_Amt,XXX.TAX3_Amt,XXX.TAX4_Amt,XXX.TAX5_Amt,XXX.TAX6_Amt,XXX.TAX7_Amt,XXX.TAX8_Amt,XXX.TAX9_Amt,XXX.TAX10_Amt,XXX.Tax1_Rate,XXX.Tax2_Rate,XXX.Tax3_Rate,XXX.Tax4_Rate,XXX.Tax5_Rate,XXX.Tax6_Rate,XXX.Tax7_Rate,XXX.Tax8_Rate,XXX.Tax9_Rate,XXX.Tax10_Rate,TSPL_COMPANY_MASTER.Bank_Name,TSPL_COMPANY_MASTER.BankAccountNo,TSPL_COMPANY_MASTER.BankIFSCCode,TSPL_COMPANY_MASTER.BankBranchAddress,TSPL_LOCATION_MASTER.ACType,TSPL_LOCATION_MASTER.accountholdername,TSPL_CUSTOMER_MASTER.PIN_NO  ,TSPL_CUSTOMER_MASTER.PIN_Code,TSPL_Customer_Invoice_Head.Balance_Amt,TSPL_Customer_Invoice_Head.Discount_Amount,TSPL_Customer_Invoice_Head.Total_Tax  from (" &
            '                        "  " &
            '                        "  select distinct Loc_Code,RefDocNo, locAdd,Final.Loc_Add1,final.Loc_Add2,Final.Loc_Add3,final.Loc_City_Code,Final.Loc_City_Name,Final.Loc_State_Name,Final.Loc_Pin_code,Location_Desc,Tin_No,PAN ,final.Document_No ,final.Document_Date, final.Status ,final.Document_Type ,final.Account_Set ,final.DocAmt,final.Customer_Code ,final.Customer_Name ,final.Created_By , final.Modify_By  ,final .Comp_Code,Description ,final.AddChargeCode ,final.AddChargeDesc,final.SAC_Code,final.SNo,final.Amount_less_Discount," &
            '                        "  final.Taxdesc1,final.Taxdesc2,final.Taxdesc3,final.Taxdesc4,final.Taxdesc5,final.Taxdesc6,final.Taxdesc7,final.Taxdesc8,final.Taxdesc9,final.Taxdesc10,final.TAX1_Amt,final.TAX2_Amt,final.TAX3_Amt,final.TAX4_Amt,final.TAX5_Amt,final.TAX6_Amt,final.TAX7_Amt,final.TAX8_Amt,final.TAX9_Amt,final.TAX10_Amt,final.Tax1_Rate,final.Tax2_Rate,final.Tax3_Rate,final.Tax4_Rate,final.Tax5_Rate,final.Tax6_Rate,final.Tax7_Rate,final.Tax8_Rate,final.Tax9_Rate,final.Tax10_Rate" &
            '                        "   from (" &
            '                        "  " &
            '                        "  select isnull(TSPL_Customer_Invoice_Head.RefDocNo,'') as RefDocNo, TSPL_Customer_Invoice_Head.Loc_Code, TSPL_LOCATION_MASTER.Add1 + case When TSPL_LOCATION_MASTER.Add2='' Then '' else ', '+ Convert(Varchar(50),TSPL_LOCATION_MASTER.Add2, 103) End + Case When TSPL_LOCATION_MASTER.Add3='' Then '' Else ', '+ COnvert( Varchar,TSPL_LOCATION_MASTER.Add3,103) end + case When TSPL_LOCATION_MASTER.City_Code ='' then '' else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.City_Code, 103) end+ Case When TSPL_LOCATION_MASTER.State='' Then '' else ', '+Convert(Varchar, loc_state.STATE_NAME ) end +  Case When TSPL_LOCATION_MASTER.Pin_code='' Then '' Else ', '+ Convert(Varchar,TSPL_LOCATION_MASTER.Pin_code, 103)  end  as locAdd, " &
            '                        "  TSPL_LOCATION_MASTER.Add1 as Loc_Add1, TSPL_LOCATION_MASTER.Add2 as Loc_Add2, TSPL_LOCATION_MASTER.Add3 as Loc_Add3, TSPL_LOCATION_MASTER.City_Code as Loc_City_Code,TSPL_CITY_MASTER.City_Name as Loc_City_Name , loc_state.STATE_NAME as Loc_State_Name,TSPL_LOCATION_MASTER.Pin_code as Loc_Pin_code ," &
            '                        "  Location_Desc,TSPL_CUSTOMER_MASTER.Tin_No,TSPL_CUSTOMER_MASTER.PAN,TSPL_Customer_Invoice_Head.Description,TSPL_Customer_Invoice_Detail.AddChargeCode,TSPL_Customer_Invoice_Detail.AddChargeDesc,TSPL_Additional_Charges.SAC_Code,TSPL_Customer_Invoice_Detail.SNo,TSPL_Customer_Invoice_Detail.Amount_less_Discount," &
            '                        "  " &
            '                        "   TSPL_Customer_Invoice_Head.Document_No, Document_Date , case when TSPL_Customer_Invoice_Head.Status=1 then 'Authorized' else 'UnAuthorized' end as Status ,  case when TSPL_Customer_Invoice_Head.Document_Type='I' then 'Tax Invoice' else case when TSPL_Customer_Invoice_Head.Document_Type='D' then 'Debit Note' else case when TSPL_Customer_Invoice_Head.Document_Type='C' then 'Credit Note' else '' end end end as Document_Type, Account_Set,Document_Total as DocAmt,  Customer_Code ,TSPL_Customer_Invoice_Head.Customer_Name,tspl_user_master.User_Name as Created_By,user_master_modify.User_Name as Modify_By  ," &
            '                        " " &
            '                        "   TSPL_Customer_Invoice_Head.Comp_Code ," &
            '                        " " &
            '                        "    isnull(Tspl_Tax1.Tax_Code,'') as Taxdesc1,isnull(Tspl_Tax2.Tax_Code,'') as Taxdesc2,isnull(Tspl_Tax3.Tax_Code,'') as Taxdesc3,isnull(Tspl_Tax4.Tax_Code,'') as Taxdesc4,isnull(Tspl_Tax5.Tax_Code,'') as Taxdesc5,isnull(Tspl_Tax6.Tax_Code,'') as Taxdesc6,isnull(Tspl_Tax7.Tax_Code,'') as Taxdesc7,isnull(Tspl_Tax8.Tax_Code,'') as Taxdesc8,isnull(Tspl_Tax8.Tax_Code,'') as Taxdesc9,isnull(Tspl_Tax10.Tax_Code,'') as Taxdesc10, TSPL_Customer_Invoice_Detail.Tax1_Rate,TSPL_Customer_Invoice_Detail.Tax2_Rate,TSPL_Customer_Invoice_Detail.Tax3_Rate,TSPL_Customer_Invoice_Detail.Tax4_Rate,TSPL_Customer_Invoice_Detail.Tax5_Rate,TSPL_Customer_Invoice_Detail.Tax6_Rate,TSPL_Customer_Invoice_Detail.Tax7_Rate,TSPL_Customer_Invoice_Detail.Tax8_Rate,TSPL_Customer_Invoice_Detail.Tax9_Rate,TSPL_Customer_Invoice_Detail.Tax10_Rate," &
            '                        "    TSPL_Customer_Invoice_Detail.TAX1_Amt,TSPL_Customer_Invoice_Detail.TAX2_Amt,TSPL_Customer_Invoice_Detail.TAX3_Amt,TSPL_Customer_Invoice_Detail.TAX4_Amt,TSPL_Customer_Invoice_Detail.TAX5_Amt,TSPL_Customer_Invoice_Detail.TAX6_Amt,TSPL_Customer_Invoice_Detail.TAX7_Amt,TSPL_Customer_Invoice_Detail.TAX8_Amt,TSPL_Customer_Invoice_Detail.TAX9_Amt,TSPL_Customer_Invoice_Detail.TAX10_Amt,TSPL_Customer_Invoice_Detail.Total_Tax" &
            '                        "  " &
            '                        "  from TSPL_Customer_Invoice_Head  " &
            '                        " left outer join TSPL_Customer_Invoice_Detail on  TSPL_Customer_Invoice_Detail.Document_No = TSPL_Customer_Invoice_Head.Document_No" &
            '                        "  " &
            '                        "     left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_Customer_Invoice_Head.Customer_Code "

            'If objCommonVar.RCDFCFP Then
            '    qry += "  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Loc_Segment_Code=TSPL_Customer_Invoice_Head.Loc_Code   "
            'Else
            '    qry += "  left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_Customer_Invoice_Head.Loc_Code   "
            'End If

            'qry += " 	left outer join TSPL_STATE_MASTER as loc_state on loc_state.STATE_CODE =TSPL_LOCATION_MASTER.State   left outer join tspl_user_master on tspl_user_master.User_Code=TSPL_Customer_Invoice_Head.Created_By  left outer join tspl_user_master as user_master_modify on user_master_modify.User_Code=TSPL_Customer_Invoice_Head.Modify_By " &
            '                        " left outer join TSPL_Additional_Charges on 	TSPL_Additional_Charges.Code = TSPL_Customer_Invoice_Detail.AddChargeCode" &
            '                        " " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax1 on Tspl_Tax1.Tax_Code =TSPL_Customer_Invoice_Head.TAX1 " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax2 on Tspl_Tax2.Tax_Code =TSPL_Customer_Invoice_Head.TAX2  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax3 on Tspl_Tax3.Tax_Code =TSPL_Customer_Invoice_Head.TAX3  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax4 on Tspl_Tax4.Tax_Code =TSPL_Customer_Invoice_Head.TAX4  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax5 on Tspl_Tax5.Tax_Code =TSPL_Customer_Invoice_Head.TAX5  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax6 on Tspl_Tax6.Tax_Code =TSPL_Customer_Invoice_Head.TAX6  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax7 on Tspl_Tax7.Tax_Code =TSPL_Customer_Invoice_Head.TAX7  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax8 on Tspl_Tax8.Tax_Code =TSPL_Customer_Invoice_Head.TAX8  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax9 on Tspl_Tax9.Tax_Code =TSPL_Customer_Invoice_Head.TAX9  " &
            '                        "  left outer join TSPL_TAX_MASTER Tspl_Tax10 on Tspl_Tax10.Tax_Code =TSPL_Customer_Invoice_Head.TAX10 " &
            '                        "  left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code = TSPL_LOCATION_MASTER.City_Code " &
            '                        " 	where TSPL_Customer_Invoice_Head.Document_No ='" & txtDocNo.Value & "'" &
            '                        " " &
            '                        " 	 )final   )XXX left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = XXX.Comp_Code  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = XXX.Customer_Code left outer join TSPL_STATE_MASTER on TSPL_CUSTOMER_MASTER.State = TSPL_STATE_MASTER.STATE_CODE "
            'If objCommonVar.RCDFCFP Then
            '    qry += " left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Loc_Segment_Code = XXX.Loc_Code "
            'Else
            '    qry += " left outer join TSPL_LOCATION_MASTER  on TSPL_LOCATION_MASTER.Location_Code = XXX.Loc_Code "
            'End If


            'qry += "  left outer join tspl_state_master as tspl_state_master_for_location_state on  tspl_state_master_for_location_state.state_code=tspl_location_master.state    " &
            '                        " " &
            '                        "  left outer join tspl_state_master as tspl_state_master_For_Comp on tspl_state_master_For_Comp.state_code = TSPL_COMPANY_MASTER.State" &
            '                        "  left outer join TSPL_Customer_Invoice_Head on TSPL_Customer_Invoice_Head.Document_No=XXX.DOCUMENT_NO "

            ''&
            ''                        "  ) XXXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE' as CopyType1) YYY ON YYY.COL1=XXXX.CopyType ORDER BY YYY.COL2 ,XXXX.SNO  " &
            ''                        " " &
            ''                        " "
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            'qry = "  select TSPL_Additional_Charges.SAC_Code, sum (TSPL_Customer_Invoice_Detail.Amount_Less_Discount) as Amount_Less_Discount  from TSPL_Customer_Invoice_Detail " &
            '      "  left outer join TSPL_Additional_Charges on 	TSPL_Additional_Charges.Code = TSPL_Customer_Invoice_Detail.AddChargeCode " &
            '      " where  Document_No ='" & txtDocNo.Value & "' group by TSPL_Additional_Charges.SAC_Code "
            'Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
            'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '    Dim frmCRV As New frmCrystalReportViewer()
            '    'frmCRV.funsubreportWithdt(CrystalReportFolder.SalesReport, dt, dt2, "crptAPServiceInvc", "rptAPInvBySAC.rpt", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptAPInvBySAC.rpt", "Address.rpt")
            '    If objCommonVar.RCDFCFP Then
            '        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, dt2, "crptAPServiceInvcnew", "rptAPInvBySAC.rpt", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptAPInvBySAC.rpt", "Address.rpt")
            '    Else
            '        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, dt2, "crptAPServiceInvcnewForAll", "rptAPInvBySAC.rpt", clsCommon.myCDate(dt.Rows(0)("Document_Date")), "rptAPInvBySAC.rpt", "Address.rpt")
            '    End If
            '    frmCRV = Nothing
            'Else
            '    clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
            'End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Function CheckIsTaxable(ByVal StrCode As String) As Integer
        '''''''''''''''''Check document is taxable --------------------
        Dim IsTaxable As Decimal = 0
        If clsCommon.myCdbl(lblTaxAmt.Text) > 0 Then
            If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                Dim IsExempted = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select Is_Tax_Exempted from TSPL_TAX_GROUP_MASTER WHERE TSPL_TAX_GROUP_MASTER.Tax_Group_Type='S' and Tax_Group_Code='" + txtTaxGroup.Value + "'"))
                If IsExempted = 1 Then
                    IsTaxable = 0
                Else
                    Dim TaxQuery As String = "select (case when isnull(tax1.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax1_amt else 0 end " &
                        " + case when isnull(tax2.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax2_amt else 0 end " &
                        " + case when isnull(tax3.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax3_amt else 0 end " &
                        " + case when isnull(tax4.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax4_amt else 0 end " &
                        " + case when isnull(tax5.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax5_amt else 0 end " &
                        " + case when isnull(tax6.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax6_amt else 0 end " &
                        " + case when isnull(tax7.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax7_amt else 0 end " &
                        " + case when isnull(tax8.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax8_amt else 0 end " &
                        " + case when isnull(tax9.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax9_amt else 0 end " &
                        " + case when isnull(tax10.Is_TCS,'N')='N' then TSPL_Customer_Invoice_Head.tax10_amt else 0 end) as aa " &
                        " from TSPL_Customer_Invoice_Head " &
                        " left outer join TSPL_TAX_MASTER as tax1 on tax1.tax_code =TSPL_Customer_Invoice_Head.tax1  " &
                         " left outer join tspl_tax_master As tax2 On tax2.tax_code =TSPL_Customer_Invoice_Head.tax2  " &
                         " left outer join tspl_tax_master as tax3 on tax3.Tax_Code =TSPL_Customer_Invoice_Head .TAX3  " &
                         " left outer join TSPL_TAX_MASTER As tax4 On tax4.Tax_Code =TSPL_Customer_Invoice_Head .tax4 " &
                         " left outer join TSPL_TAX_MASTER as tax5 on tax5.Tax_Code =TSPL_Customer_Invoice_Head .tax5 " &
                         " left outer join TSPL_TAX_MASTER As tax6 On tax6.Tax_Code =TSPL_Customer_Invoice_Head .TAX6 " &
                         " left outer join TSPL_TAX_MASTER as tax7 on tax7.Tax_Code =TSPL_Customer_Invoice_Head .TAX7 " &
                         " left outer join TSPL_TAX_MASTER As tax8 On tax8.Tax_Code =TSPL_Customer_Invoice_Head .TAX8 " &
                         " left outer join TSPL_TAX_MASTER as tax9 on tax9.Tax_Code =TSPL_Customer_Invoice_Head .TAX9 " &
                         " left outer join TSPL_TAX_MASTER As tax10 On tax10.Tax_Code =TSPL_Customer_Invoice_Head .TAX10 " &
                         " where TSPL_Customer_Invoice_Head.Document_No ='" + StrCode + "'"
                    Dim TaxAmount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(TaxQuery))
                    If TaxAmount > 0 Then
                        IsTaxable = 1
                    End If
                End If
            End If
        End If
        Return IsTaxable
    End Function


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        CancelData(False)
    End Sub

    Function CancelData(ByVal IsCancelByAdmin As Boolean) As Boolean
        Try
            Dim isEinvoiceCancelled As Boolean = False
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Code is empty")
            End If

            If chkAgainstServiceInvoice.Checked = False Then
                Throw New Exception("You can not cancelled this document because AR Document is not against Service Invoice.")
            End If

            If clsCommon.MyMessageBoxShow("Are you sure to Cancel the Record?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then

                Return False
            End If
            Dim strIrnNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IRN_No from TSPL_Customer_Invoice_Head where Document_No='" & txtDocNo.Value & "'"))
            If IsCancelByAdmin Then
                If clsCommon.myLen(strIrnNo) > 0 Then
                    If clsCommon.MyMessageBoxShow(Me, "Is E-invoice Cancelled on GST Portal?", "", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                        isEinvoiceCancelled = True
                    End If
                End If
            End If
            Dim strReceiptCount As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select receipt_no from TSPL_RECEIPT_DETAIL where Document_No='" & txtDocNo.Value & "'"))
            If clsCommon.myLen(strReceiptCount) > 0 Then
                Throw New Exception("You cannot cancelled this document because receiving (" + clsCommon.myCstr(strReceiptCount) + ") has been done against its AR Invoice.")
            End If

            If FlagDocumentIsTaxable = 1 AndAlso clsERPFuncationality.GetEInvoiceStatus(txtDate.Value) = True AndAlso clsCommon.CompairString(EInvoiceType, "BB") = CompairStringResult.Equal AndAlso Not isEinvoiceCancelled Then
                Dim EInvoiceCancelTimeValid As Int64 = 0
                EInvoiceCancelTimeValid = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(" Select  isnull (DATEDIFF(hour,Posting_Date,GETDATE()),0) as PostedHours from TSPL_Customer_Invoice_Head where Document_No = '" + txtDocNo.Value + "'"))
                If EInvoiceCancelTimeValid >= 24 Then
                    Throw New Exception("AR Document can not be cancelled.It has been more than 24 hours.")
                End If
            End If

            clsCustomerInvoiceHead.CancelData(txtDocNo.Value, IsCancelByAdmin)
            clsCommon.MyMessageBoxShow(Me, "Successfully Cancelled", Me.Text)
            AddNew()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return True
    End Function

    Private Sub butCostCenterAndHirerachy_Update_AfterPost_Click(sender As Object, e As EventArgs) Handles butCostCenterAndHirerachy_Update_AfterPost.Click
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim strJEStatus As String = clsDBFuncationality.getSingleValue("select Authorized from TSPL_JOURNAL_MASTER where Source_Doc_No ='" + txtDocNo.Value + "' ", trans)
            If clsCommon.CompairString(strJEStatus, "A") = CompairStringResult.Equal Then
                clsDBFuncationality.ExecuteNonQuery("ALTER TABLE TSPL_JOURNAL_DETAILS DISABLE TRIGGER TRG_JD_FiscaYearEndNoUpdateNoDelete", trans)
            End If
            For Each grow As GridViewRowInfo In gv1.Rows
                Dim coll As New Hashtable()
                If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colACCode).Value)) > 0 Then

                    Dim strGLAccountCode As String = clsCommon.myCstr(grow.Cells(colACCode).Value)
                    clsCommon.AddColumnsForChange(coll, "Hirerachy_Code", clsCommon.myCstr(grow.Cells(colHierarchyCode).Value), True)
                    clsCommon.AddColumnsForChange(coll, "Cost_Centre_Code", clsCommon.myCstr(grow.Cells(colCostCenterCode).Value), True)
                    clsCommon.AddColumnsForChange(coll, "Hirerachy_Code1", clsCommon.myCstr(grow.Cells(colHierarchyLevel1).Value), True)
                    clsCommon.AddColumnsForChange(coll, "Hirerachy_Code2", clsCommon.myCstr(grow.Cells(colHierarchyLevel2).Value), True)
                    clsCommon.AddColumnsForChange(coll, "Hirerachy_Code3", clsCommon.myCstr(grow.Cells(colHierarchyLevel3).Value), True)
                    clsCommon.AddColumnsForChange(coll, "Hirerachy_Code4", clsCommon.myCstr(grow.Cells(colHierarchyLevel4).Value), True)

                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_Customer_Invoice_Detail", OMInsertOrUpdate.Update, "Document_No='" + txtDocNo.Value + "' and GL_Account_Code = '" + strGLAccountCode + "'", trans)
                    Dim strVoucherNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Voucher_No  from TSPL_JOURNAL_MASTER where Source_Doc_No = '" + txtDocNo.Value + "' ", trans))
                    Dim qry As String = "update TSPL_JOURNAL_DETAILS SET Hirerachy_Code='" + clsCommon.myCstr(grow.Cells(colHierarchyCode).Value) + "',Cost_Centre_Code='" + clsCommon.myCstr(grow.Cells(colCostCenterCode).Value) + "',Hirerachy_Code3= " + IIf(clsCommon.myLen(clsCommon.myCstr(grow.Cells(colHierarchyLevel3).Value)) > 0, " '" & clsCommon.myCstr(grow.Cells(colHierarchyLevel3).Value) & "' ", "NULL") + ",Hirerachy_Code4=" + IIf(clsCommon.myLen(clsCommon.myCstr(grow.Cells(colHierarchyLevel4).Value)) > 0, " '" & clsCommon.myCstr(grow.Cells(colHierarchyLevel4).Value) & "' ", "NULL") + " WHERE Voucher_No='" + strVoucherNo + "' and Account_code='" + strGLAccountCode + "' "
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If
            Next
            If clsCommon.CompairString(strJEStatus, "A") = CompairStringResult.Equal Then
                clsDBFuncationality.ExecuteNonQuery("ALTER TABLE TSPL_JOURNAL_DETAILS ENABLE TRIGGER TRG_JD_FiscaYearEndNoUpdateNoDelete", trans)
            End If
            trans.Commit()
            common.clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
        Catch ex As Exception
            trans.Rollback()
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Public Sub Import()
        Try
            Dim gv As New UserControls.MyRadGridView
            Me.Controls.Add(gv)
            Dim obj As New List(Of clsOPInvoiceForTCS)
            Dim currentdate As Date = Date.Today

            If transportSql.importExcel(gv, "FINANCIAL YEAR CODE", "CUSTOMER CODE", "SALE AMT") Then

                'Dim trans As SqlTransaction = Nothing
                Dim linno As Integer = 0
                Dim TempNewRecord As Boolean = False
                Try
                    'trans = clsDBFuncationality.GetTransactin()
                    clsCommon.ProgressBarShow()
                    For Each grow As GridViewRowInfo In gv.Rows
                        Dim Arr As New clsOPInvoiceForTCS()
                        linno += 1
                        If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("FINANCIAL YEAR CODE").Value))) Then
                            Continue For
                        Else
                            Arr.FINANCIAL_YEAR_CODE = clsCommon.myCstr(grow.Cells("FINANCIAL YEAR CODE").Value)

                        End If
                        If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("CUSTOMER CODE").Value))) Then
                            Continue For
                        Else
                            Dim str As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select cust_Code from TSPL_Customer_Master where cust_Code='" + clsCommon.myCstr(grow.Cells("CUSTOMER CODE").Value) + "'"))
                            If clsCommon.CompairString(str, clsCommon.myCstr(grow.Cells("CUSTOMER CODE").Value)) = CompairStringResult.Equal Then
                                If (clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from TSPL_OP_invoice_for_TCS where Customer_Code='" + clsCommon.myCstr(grow.Cells("CUSTOMER CODE").Value) + "'"))) > 0 Then
                                    Continue For

                                Else
                                    Arr.CUSTOMER_CODE = clsCommon.myCstr(grow.Cells("CUSTOMER CODE").Value)

                                End If
                            Else
                                Continue For
                            End If
                        End If

                        If (String.IsNullOrEmpty(clsCommon.myCstr(grow.Cells("SALE AMT").Value))) Then
                            Continue For
                        Else
                            Arr.SALE_AMT = clsCommon.myCDecimal(grow.Cells("SALE AMT").Value)
                        End If
                        obj.Add(Arr)
                    Next

                    Dim duplicatesRoute As New List(Of clsOPInvoiceForTCS)
                    duplicatesRoute = obj.GroupBy(Function(x) x.CUSTOMER_CODE).Where(Function(group) group.Count() > 1).SelectMany(Function(group) group).ToList
                    Dim strDRoute As String = String.Empty
                    For Each duplicate As clsOPInvoiceForTCS In duplicatesRoute
                        strDRoute += "[" + duplicate.CUSTOMER_CODE + "] "
                    Next

                    clsCommon.ProgressBarHide()
                    If clsCommon.MyMessageBoxShow(Me, "Total Correct Document [" + clsCommon.myCstr(obj.Count) + "] out of [" + clsCommon.myCstr(linno) + "] Are You Sure.", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                        Dim isNewEntry As Boolean = True
                        If obj IsNot Nothing AndAlso obj.Count > 0 Then
                            For Each items As clsOPInvoiceForTCS In obj
                                Try
                                    Dim ObjTCS As New clsOPInvoiceForTCS()
                                    ObjTCS.FINANCIAL_YEAR_CODE = items.FINANCIAL_YEAR_CODE
                                    ObjTCS.CUSTOMER_CODE = items.CUSTOMER_CODE
                                    ObjTCS.SALE_AMT = items.SALE_AMT

                                    ObjTCS.SaveData(ObjTCS, isNewEntry)



                                Catch ex As Exception
                                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                                End Try
                            Next
                            common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                        Else
                            common.clsCommon.MyMessageBoxShow(Me, "Data Not Found", Me.Text, MessageBoxButtons.OK)

                        End If

                    Else
                        common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Failed", Me.Text, MessageBoxButtons.OK)
                    End If

                    clsCommon.ProgressBarHide()

                Catch ex As Exception
                    clsCommon.ProgressBarHide()
                    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                End Try
            Else
                clsCommon.MyMessageBoxShow(Me, "Excel Sheet is not in expected format", Me.Text)

            End If

            'clsCommon.ProgressBarHide()
            Me.Controls.Remove(gv)
        Catch ex As Exception
            'clsCommon.ProgressBarHide()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)

        End Try
    End Sub
    Public Sub Export()
        Try
            Dim str As String = "select FINANCIAL_YEAR_CODE as [FINANCIAL YEAR CODE],CUSTOMER_CODE as [CUSTOMER CODE],SALE_AMT as [SALE AMT] from TSPL_OP_INVOICE_FOR_TCS"
            Dim whrCls As String = ""
            ListImpExpColumnsMandatory = New List(Of String)({"FINANCIAL YEAR CODE", "CUSTOMER CODE", "SALE AMT"})
            transportSql.ExporttoExcel(str, whrCls, Me)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiOPInvoiceForTCS_Click(sender As Object, e As EventArgs) Handles rmiOPInvoiceForTCS.Click
        Import()
    End Sub

    Private Sub rmiExportOPInvoiceForTCS_Click(sender As Object, e As EventArgs) Handles rmiExportOPInvoiceForTCS.Click
        Export()
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtDocNo.Value, "Document_No", "TSPL_Customer_Invoice_Head", "TSPL_Customer_Invoice_Detail")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub txtLocationPrefix__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocationPrefix._MYValidating
        Try
            Dim WhrCls As String = ""
            Dim qry As String = "Select Location_Code as Code,Location_Desc as Description from TSPL_LOCATION_MASTER "
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                WhrCls = "  TSPL_LOCATION_MASTER.Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
            End If
            txtLocationPrefix.Value = clsCommon.ShowSelectForm("LocationFndr", qry, "Code", WhrCls, txtLocationPrefix.Value, "Code", isButtonClicked)
            If clsCommon.myLen(clsCommon.myCstr(txtLocationPrefix.Value)) > 0 Then
                txtLocationPrefixName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Location_Desc  from TSPL_LOCATION_MASTER WHERE  Location_Code='" & txtLocationPrefix.Value & "' "))
            Else
                txtLocationPrefixName.Text = ""
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub EinvoiceBtnUpdate_Click(sender As Object, e As EventArgs) Handles EinvoiceBtnUpdate.Click
        UpdateEInvoice()
    End Sub

    Sub UpdateEInvoice()
        Try

            Dim obj As New clsCustomerInvoiceHead
            obj.Document_No = clsCommon.myCstr(txtDocNo.Value)
            obj.irn_no = EInvoiceIRNNo.Text
            obj.Ack_No = EinvoiceAckNo.Text
            obj.Ack_Date = txtAckDate.Value
            obj.QR_Code = EInvoiceQrCode.Text
            clsCustomerInvoiceHead.UpdateEInvoiceAfterPosting(obj, txtDocNo.Value, Nothing)
            clsCommon.MyMessageBoxShow("E-Invoice Updated Successfully", Me.Text)

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkEinvoice_CheckStateChanged(sender As Object, e As EventArgs) Handles chkEinvoice.CheckStateChanged
        Try
            If chkEinvoice.Checked = True Then
                RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Visible
            Else
                RadPageView1.Pages("RadPageViewPage5").Item.Visibility = ElementVisibility.Collapsed
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnAdminCancel_Click(sender As Object, e As EventArgs) Handles btnAdminCancel.Click
        CancelData(True)
    End Sub
End Class

