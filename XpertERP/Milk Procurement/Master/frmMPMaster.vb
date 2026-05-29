'-----------------BM00000003414 for vlc finder
Imports common
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.IO
Imports XpertERPCommanServices
Imports XpertERPHRandPayroll
Imports System.Text

' Create By Pankaj Jha
' Date 30.05.14
' This Screen Reads Data For Milk Producers

Public Class FrmMPMaster
    Inherits FrmMainTranScreen
    Public Const colSlNO As String = "COLSLNO"
    Public Const COLTypeOfAnimal As String = "COLTypeOfAnimal"
    Public Const colDesc As String = "COLDESC"
    Public Const colCountOfAnimal As String = "colCountOfAnimal"
    Public Const colStatus As String = "colStatus"
    Dim EnableBankFromMaster As Boolean
    Dim Frm_Open As FrmMainTranScreen
    Dim SettBankIFSCCodeValidateByService As Boolean = False
    Dim SettJanAadharNoMandatory As Boolean = False
#Region "Variables"
    Dim userCode As String = Nothing
    Dim compCode As String = Nothing
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim obj As clsMpMaster = Nothing
    Public errorControl As clsErrorControl = New clsErrorControl()
    Dim arrLoc As String = Nothing
    Dim File_Name As String = ""
    Dim SettNoOFIncentiveForImportExport As Integer
    Dim UseVLCUploaderCodeMPUploaderCodeInMCCProcurement As Boolean = False
    Dim IncentiveAccNoMandatoryInMPMaster As Boolean = False
#End Region
    Sub LoadAccount_Type()
        cmbAccount_Type.DataSource = Nothing
        Dim dt As New DataTable()
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))

        Dim dr As DataRow = Nothing

        dr = dt.NewRow()
        dr("Code") = "Sav"
        dr("Name") = "Saving"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Cur"
        dr("Name") = "Current"
        dt.Rows.Add(dr)


        cmbAccount_Type.DataSource = dt
        cmbAccount_Type.DisplayMember = "Name"
        cmbAccount_Type.ValueMember = "Code"
        cmbAccount_Type.SelectedValue = 0
    End Sub


    Private Sub MCCLOCATIONFINDER()
        Try
            Dim obj As New clsMCCCodes()
            obj = clsMCCCodes.GetData()

            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.Default_LocCode) > 0 Then
                arrLoc = obj.arrLocCodes
            Else
                fndVLCode.Enabled = False
                Throw New Exception("Please Set Default Location Of LogIn User")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    '#Region "User Defined Functions and Subroutines"
    Public Sub New(ByVal user As String, ByVal company As String)
        InitializeComponent()
        userCode = user
        compCode = company
    End Sub
    Public Sub New()
        InitializeComponent()
    End Sub

    Function isDuplicateMPCode(ByVal isUpdate As Boolean) As Boolean
        Dim qry As String = "select COUNT(*) from TSPL_MP_MASTER where MP_Code_vlc_uploader='" & txtMPCodeVlcUploader.Text & "' and mp_code<>'" & fndMPCode.Value & "' and vlc_Code='" & fndVLCode.Value & "'"
        Dim rvalue As Boolean = False
        Dim cnt As Integer = 0
        cnt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
        If cnt >= 1 Then
            rvalue = True
            'ElseIf (Not isUpdate) And cnt >= 1 Then
        Else
            rvalue = False
        End If
        Return rvalue
    End Function
    Sub reset()
        Try
            LoadAccount_Type()
            'txtMPCodeVlcUploader.ReadOnly = False
            txtNoOfTotalDependentMember.ReadOnly = True
            txtNoOfBreadableMilkAnimal.TextAlign = HorizontalAlignment.Right
            txtMilkAvlblForSale.TextAlign = HorizontalAlignment.Right
            txtMilkProduction.TextAlign = HorizontalAlignment.Right
            txtHomeConsumption.TextAlign = HorizontalAlignment.Right
            txtMPCodeVlcUploader.Text = ""
            fndMPCode.Value = ""
            fndVLCode.Value = ""
            txtVLCName.Text = ""
            txtMPName.Text = ""
            txtMPNameHindi.Text = ""
            fndVillegeCode.Value = ""
            txtVillegeName.Text = ""
            txtFatherName.Text = ""
            txtAdd1.Text = ""
            txtAdd2.Text = ""
            txtZila.Text = ""
            txtTehsil.Text = ""
            txtCityName.Text = ""
            fndCountryCode.Value = ""
            txtCountryName.Text = ""
            fndCityCode.Value = ""
            fndStateCode.Value = ""
            txtStateName.Text = ""
            fndCountryCode.Value = "INDIA"
            txtCountryName.Text = "INDIA"
            txtPinCode.Text = ""
            txtTelePhone.Text = ""
            txtFAX.Text = ""
            txtJanAadharNo.Text = ""
            TxtThirdPartyCode.Text = ""
            MylblThirdPartySource.Text = ""
            txtEmail.Text = ""
            txtEducation.Text = ""
            dtpDOB.Value = clsCommon.GETSERVERDATE()
            txtLandHolding.Text = ""
            txtNoofAnimal.Text = ""
            txtNoOfChildrenMember.Text = "0"
            txtNoOfAdultMember.Text = "0"
            txtNoOfTotalDependentMember.Text = "0"
            loadBlankDgvBuffaloes()
            'txtNoofCows.Text = ""
            'loadBlankDgvCows()
            txtNoOfBreadableMilkAnimal.Text = ""
            txtMilkProduction.Text = ""
            txtHomeConsumption.Text = ""
            txtMilkAvlblForSale.Text = ""
            UcAttachment1.Form_ID = Me.Form_ID
            RadPageView1.Visible = True
            RadPageView1.Pages("RadPageViewPage6").Item.Visibility = ElementVisibility.Visible
            RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Hidden
            RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Collapsed
            UcAttachment1.BlankAllControls()
            txtPayeeName.Text = ""
            txtBankName.Text = ""
            txtBankCode.Value = ""
            txtBankCityName.Text = ""
            fndBankCity.Value = ""
            fndBankState.Value = ""
            txtBankStateName.Text = ""
            txtBankBranch.Value = ""
            txtBankBranchName.Text = ""
            txtIFCICode.Text = ""
            '===========================
            txtNameOfBank_M.Text = ""
            txtBankBranch_M.Text = ""
            txtBankState_M.Text = ""
            txtBankCity_M.Text = ""
            txtBankIFSC_M.Text = ""
            '============================

            txtBankIFSC_M.Enabled = True
            txtAccountNo.Enabled = True

            txtAccountNo.Text = ""
            txtTolerance.Text = ""
            txtIncentiveMult.arrValueMember = Nothing
            btnSave.Text = "&Save"
            btnDelete.Enabled = False
            RadPageView1.SelectedPage = RadPageViewPage1
            fndMPCode.Focus()
            fndMPCode.MyReadOnly = False
            txtNoofAnimal.Visible = False
            lblNoofBuffaloes.Visible = False
            btnBuffaloesGo.Visible = False
            mnuBuffaloesDetailsExport.Visibility = ElementVisibility.Collapsed
            mnuCowDetailsExport.Visibility = ElementVisibility.Collapsed
            mnuBuffaloesDetailsImport.Visibility = ElementVisibility.Collapsed
            mnuCowDetailsImport.Visibility = ElementVisibility.Collapsed
            MCCLOCATIONFINDER()
            txtBankName.Enabled = False
            fndBankState.Enabled = False
            txtBankStateName.Enabled = False
            fndBankCity.Enabled = False
            txtBankCityName.Enabled = False
            'txtBankBranch.Enabled = False
            txtBankBranchName.Enabled = False
            txtIFCICode.Enabled = False
            ddlGender.Text = ""
            CboMaritalStatus.Text = ""
            chkIsVSP.Checked = False
            chkInActive.Checked = False
            ddlTypeOfFormer.Text = ""
            UcCamControl1.PicBox.Image = Nothing
            txtVLCName.Text = ""
            lblMCCName.Text = ""
            txtVlCUploader.Text = ""
            txtVSPName.Text = ""
            txtMCC.Value = ""
            txtIncentiveGLAccount.Value = ""
            lblIncentiveGLAccount.Text = ""
            txtCastCategory.Value = ""
            lblCastCategory.Text = ""
            txtDistrict.Value = ""
            lblDistrict.Text = ""
            txtBlockCode.Value = ""
            lblBlockCode.Text = ""
            txtZone.Value = ""
            lblZone.Text = ""
            txtVidhanSabha.Value = ""
            lblVidhanSabha.Text = ""
            txtPanchayatSamiti.Value = ""
            lblPanchayatSamiti.Text = ""
            txtGrampanchayat.Value = ""
            lblGrampanchayat.Text = ""
            txtRevenueVillage.Value = ""
            lblRevenueVillage.Text = ""

            If EnableBankFromMaster = True Then
                pnlBankDetailsManual.Visible = False
            Else
                pnlBankDetailsManual.Visible = True
            End If
            If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                UcCustomFields1.BlankAllControls()
                UcCustomFields1.SetDefaultValues()
            End If
            RadGroupBox2.Visible = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Function loadAnimals() As DataTable
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(" select 'Cow' as value union all select 'Buffalo' as value union all select 'Camel' as value union all select 'Goat' as value ")
        Return dt
    End Function

    Function LoadAnimalStatus() As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("CODE", GetType(String))
        dt.Columns.Add("NAME", GetType(String))

        Dim dr As DataRow
        dr = dt.NewRow()
        dr("CODE") = " "
        dr("NAME") = "None"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("CODE") = "Dry"
        dr("NAME") = "Dry"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("CODE") = "Milk"
        dr("NAME") = "Milk"
        dt.Rows.Add(dr)
        Return dt
    End Function
    Sub loadBlankDgvBuffaloes()
        Try
            dgvNoofBuffaloes.Rows.Clear()
            dgvNoofBuffaloes.Columns.Clear()
            Dim repoComboColumn As GridViewComboBoxColumn
            repoComboColumn = New GridViewComboBoxColumn()
            repoComboColumn.Name = COLTypeOfAnimal
            repoComboColumn.Width = 120
            repoComboColumn.HeaderText = "Type Of Animal"
            repoComboColumn.DataSource = loadAnimals()
            repoComboColumn.DisplayMember = "Value"
            repoComboColumn.ValueMember = "Value"

            Dim repoDeciCol As GridViewDecimalColumn
            repoDeciCol = New GridViewDecimalColumn()
            repoDeciCol.Name = colCountOfAnimal
            repoDeciCol.Width = 120
            repoDeciCol.DecimalPlaces = 0
            repoDeciCol.EnableExpressionEditor = True
            repoDeciCol.Minimum = 0
            repoDeciCol.HeaderText = "Count Of Animal"
            'repoDeciCol.DataSource = loadAnimals()
            'repoDeciCol.DisplayMember = "Value"
            'repoDeciCol.ValueMember = "Value"

            dgvNoofBuffaloes.Columns.Add(colSlNO, "SL. NO")
            dgvNoofBuffaloes.MasterTemplate.Columns.Add(repoComboColumn)
            dgvNoofBuffaloes.MasterTemplate.Columns.Add(repoDeciCol)
            dgvNoofBuffaloes.Columns.Add(colDesc, "Bread")
            dgvNoofBuffaloes.Columns(colSlNO).Width = 100
            dgvNoofBuffaloes.Columns(colDesc).Width = 400


            repoComboColumn = New GridViewComboBoxColumn()
            repoComboColumn.Name = colStatus
            repoComboColumn.Width = 120
            repoComboColumn.HeaderText = "Status"
            repoComboColumn.DataSource = LoadAnimalStatus()
            repoComboColumn.ValueMember = "CODE"
            repoComboColumn.DisplayMember = "NAME"
            dgvNoofBuffaloes.MasterTemplate.Columns.Add(repoComboColumn)


            dgvNoofBuffaloes.AllowAddNewRow = True
            dgvNoofBuffaloes.AllowEditRow = True
            dgvNoofBuffaloes.AllowDeleteRow = True
            dgvNoofBuffaloes.AllowRowResize = False
            dgvNoofBuffaloes.AllowRowReorder = False
            dgvNoofBuffaloes.AllowColumnResize = True
            dgvNoofBuffaloes.AllowColumnChooser = False
            dgvNoofBuffaloes.AllowAutoSizeColumns = False
            dgvNoofBuffaloes.ShowGroupPanel = False
            dgvNoofBuffaloes.AddNewRowPosition = SystemRowPosition.Bottom
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try


    End Sub
    Sub UpdateSLNo()
        For i As Integer = 0 To dgvNoofBuffaloes.Rows.Count - 1
            dgvNoofBuffaloes.Rows(i).Cells(colSlNO).Value = (i + 1)
        Next
    End Sub
    'Sub loadBlankDgvCows()
    '    Try
    '        dgvNoofCows.Rows.Clear()
    '        dgvNoofCows.Columns.Clear()
    '        dgvNoofCows.Columns.Add(colSlNO, "SL. NO")
    '        dgvNoofCows.Columns.Add(colDesc, "Bread Of Cow")
    '        dgvNoofCows.Columns(colSlNO).Width = 100
    '        dgvNoofCows.Columns(colDesc).Width = 400
    '        dgvNoofBuffaloes.Rows.AddNew()
    '        dgvNoofCows.AllowAddNewRow = True
    '        dgvNoofCows.AllowEditRow = True
    '        dgvNoofCows.AllowDeleteRow = True
    '        dgvNoofCows.AllowRowResize = False
    '        dgvNoofCows.AllowRowReorder = False
    '        dgvNoofCows.AllowColumnResize = True
    '        dgvNoofCows.AllowColumnChooser = False
    '        dgvNoofCows.AllowAutoSizeColumns = False
    '        dgvNoofCows.ShowGroupPanel = False
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
    '    End Try
    'End Sub
    Private Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.frmMPMaster)
        If Not (MyBase.isReadFlag) Then
            common.clsCommon.MyMessageBoxShow(Me, "Permission Denied", Me.Text)
            Me.Close()
        End If
        btnSave.Visible = MyBase.isModifyFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        'RadMenu1.Visible = MyBase.isExport
        If MyBase.isExport = True Then
            mnuExport.Enabled = True
            mnuImport.Enabled = True
        Else
            mnuExport.Enabled = False
            mnuImport.Enabled = False
        End If
    End Sub
    Function allowToSave() As Boolean
        Try
            If clsCommon.myLen(txtFAX.Text) > 0 Then
                If clsCommon.myLen(txtFAX.Text) <> 12 Then
                    clsCommon.MyMessageBoxShow(Me, "Invalid Aadhar No.Please Enter Aadhar No 12 Digit", Me.Text)
                    errorControl.SetError(txtFAX, "Invalid Aadhar No.Please Enter Aadhar No 12 Digit")
                    Return False
                Else

                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*) from tspl_mp_master where fax='" & clsCommon.myCstr(txtFAX.Text) & "' and MP_Code<>'" & clsCommon.myCstr(fndMPCode.Value) & "'")) > 0 Then
                        common.clsCommon.MyMessageBoxShow(Me, "Same Aadhar No is exist with another MP so please change Aadhar No because Aadhar No is unique.")
                        Return False
                    End If

                    errorControl.ResetError(txtFAX)
                End If
            Else
                errorControl.ResetError(txtFAX)
            End If
            'If clsCommon.MySpecialChars(txtAccountNo.Text) Then
            '    common.clsCommon.MyMessageBoxShow("Account No Should No have And Special Characters")
            '    Return False
            'End If
            If clsCommon.myLen(txtJanAadharNo.Text) > 0 Then
                If clsCommon.myLen(txtJanAadharNo.Text) <> 10 Then
                    clsCommon.MyMessageBoxShow(Me, "Invalid Jan Aadhar No.Please Enter 10 Digit Jan Aadhar No ", Me.Text)
                    errorControl.SetError(txtJanAadharNo, "Invalid Jan Aadhar No.Please Enter 10 Digit Jan Aadhar No")
                    Return False
                    'Else
                    '    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*) from tspl_mp_master where Jan_Aadhar_No='" & clsCommon.myCstr(txtJanAadharNo.Text) & "' and MP_Code<>'" & clsCommon.myCstr(fndMPCode.Value) & "'")) > 0 Then
                    '        common.clsCommon.MyMessageBoxShow("Same Jan Aadhar No is exist with another MP so please change Aadhar No because Jan Aadhar No Should be unique.")
                    '        Return False
                    '    End If
                    '    errorControl.ResetError(txtJanAadharNo)
                End If
            Else
                If SettJanAadharNoMandatory Then
                    clsCommon.MyMessageBoxShow(Me, "Please Fill Jan Aadhar No", Me.Text)
                    errorControl.ResetError(txtJanAadharNo)
                    Return False
                End If
            End If

            If IncentiveAccNoMandatoryInMPMaster = True Then
                If clsCommon.myLen(txtIncentiveGLAccount.Value) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, " Incentive GL Account Must Not be Blank (Under Bank Detail for Payments)", Me.Text)
                    RadPageView1.SelectedPage = RadPageViewPage5
                    txtIncentiveGLAccount.Focus()
                    errorControl.SetError(txtIncentiveGLAccount, "Incentive GL Account Must Not be Blank ")
                    Return False
                Else
                    errorControl.SetError(txtIncentiveGLAccount, "")
                End If
            End If

            If clsCommon.myLen(fndVLCode.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, " VLC Code Must Not be Blank (Under General Tab)", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage1
                fndVLCode.Focus()
                errorControl.SetError(fndVLCode, "VLC Code Must Not be Blank ")
                Return False
            Else
                errorControl.SetError(fndVLCode, "")
            End If
            If clsCommon.myLen(txtMPName.Text) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, " MP Name Must Not be Blank (Under General Tab)")
                RadPageView1.SelectedPage = RadPageViewPage1
                txtMPName.Focus()
                errorControl.SetError(txtMPName, "MP Name Must Not be Blank ")
                Return False
            Else
                errorControl.SetError(txtMPName, "")
            End If

            'If clsCommon.myLen(txtMPCodeVlcUploader.Text) > 0 Then
            '    Dim mpcode As String = clsDBFuncationality.getSingleValue("Select MP_Uploader_Code
            '    from TSPL_DBT_NEFT_DETAIL 
            '    Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR   
            '    left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
            '    where  MP_Uploader_Code= '" + txtMPCodeVlcUploader.Text + "'")
            '    If clsCommon.myLen(mpcode) > 0 Then
            '        clsCommon.MyMessageBoxShow(Me, "Duplicate MP Code ")
            '        Return False
            '    End If
            'End If
            If clsCommon.myLen(txtMPCodeVlcUploader.Text) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, " MP Code For VLC Uploader  Must Not be Blank (Under General Tab)", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage1
                txtMPCodeVlcUploader.Focus()
                errorControl.SetError(txtMPCodeVlcUploader, "MP Code For VLC Uploader  Must Not be Blank")
                Return False
            Else
                errorControl.SetError(txtMPCodeVlcUploader, "")
            End If

            If isDuplicateMPCode(IIf(clsCommon.CompairString(btnSave.Text, "Save") = CompairStringResult.Equal, False, True)) Then

                clsCommon.MyMessageBoxShow(Me, " Duplicate MP Code for VLC Uploader", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage1
                txtMPCodeVlcUploader.Focus()
                errorControl.SetError(txtMPCodeVlcUploader, "Duplicate MP Code for VLC Uploader")
                Return False
            Else
                errorControl.SetError(txtMPCodeVlcUploader, "")
            End If

            If clsCommon.myLen(txtAdd1.Text) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, " Please Enter Address Line 1, It is Manadatory (Under General Tab)", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage1
                txtAdd1.Focus()
                errorControl.SetError(txtAdd1, "Please Enter Address Line 1, It is Manadatory ")
                Return False
            Else
                errorControl.SetError(txtAdd1, "")
            End If

            If clsCommon.myLen(TxtThirdPartyCode.Text) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, " Third Party Code Must Not be Blank (Under General Tab)", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage1
                TxtThirdPartyCode.Focus()
                errorControl.SetError(TxtThirdPartyCode, "Third Party Code Must Not be Blank ")
                Return False
            Else
                errorControl.SetError(TxtThirdPartyCode, "")
            End If

            'If clsCommon.myLen(fndCountryCode.Value) <= 0 Then
            '    clsCommon.MyMessageBoxShow(" Please Enter Country Name, It is Manadatory (Under General Tab)")
            '    RadPageView1.SelectedPage = RadPageViewPage1
            '    fndCountryCode.Focus()
            '    errorControl.SetError(fndCountryCode, "Please Enter Country Name, It is Manadatory")
            '    Return False
            'Else
            '    errorControl.SetError(fndCountryCode, "")
            'End If
            'If clsCommon.myLen(fndStateCode.Value) <= 0 Then
            '    clsCommon.MyMessageBoxShow(" Please Enter State Name, It is Manadatory (Under General Tab)")
            '    RadPageView1.SelectedPage = RadPageViewPage1
            '    fndStateCode.Focus()
            '    errorControl.SetError(fndStateCode, "Please Enter State Name, It is Manadatory")
            '    Return False
            'Else
            '    errorControl.SetError(fndStateCode, "")
            'End If
            'If clsCommon.myLen(fndCityCode.Value) <= 0 Then
            '    clsCommon.MyMessageBoxShow(" Please Enter City Name, It is Manadatory (Under General Tab)")
            '    RadPageView1.SelectedPage = RadPageViewPage1
            '    fndCityCode.Focus()
            '    errorControl.SetError(fndCityCode, "Please Enter City Name, It is Manadatory")
            '    Return False
            'Else
            '    errorControl.SetError(fndCityCode, "")
            'End If
            'If clsCommon.myLen(txtPinCode.Text) <= 0 Then
            '    clsCommon.MyMessageBoxShow(" Please Enter Pincode, It is Manadatory (Under General Tab)")
            '    RadPageView1.SelectedPage = RadPageViewPage1
            '    txtPinCode.Focus()
            '    errorControl.SetError(txtPinCode, "Please Enter Pincode, It is Manadatory")
            '    Return False
            'Else
            '    errorControl.SetError(txtPinCode, "")
            'End If
            Dim check As Match = Regex.Match(txtEmail.Text, "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
            If check.Success = False AndAlso txtEmail.Text <> "" Then
                clsCommon.MyMessageBoxShow(Me, " Please Enter Valid Email, It is in Invalid Format (Under General Tab)", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage1
                txtEmail.Focus()
                errorControl.SetError(txtEmail, "Please Enter Valid Email, It is in Invalid Format")
                Return False
            Else
                errorControl.SetError(txtEmail, "")
            End If
            'If clsCommon.myCdbl(txtNoofAnimal.Text) < 0 Then
            '    clsCommon.MyMessageBoxShow("No of Animal Can Not be Negative")
            '    RadPageView1.SelectedPage = RadPageViewPage2
            '    txtNoofAnimal.Focus()
            '    Return False
            'End If
            Dim i As Integer = 0
            'If clsCommon.myCdbl(txtNoofAnimal.Text) > dgvNoofBuffaloes.Rows.Count Then
            '    clsCommon.MyMessageBoxShow("Please click on go button under Animal details tab, and fill Animal Details")
            '    RadPageView1.SelectedPage = RadPageViewPage2
            '    btnBuffaloesGo.Focus()
            '    Return False
            'End If
            'If clsCommon.myCdbl(txtNoofCows.Text) > dgvNoofCows.Rows.Count Then
            '    clsCommon.MyMessageBoxShow("Please click on go button under Cow details tab, and fill Cow Details")
            '    RadPageView1.SelectedPage = RadPageViewPage3
            '    btnNoofCowsgo.Focus()
            '    Return False
            'End If

            'If dgvNoofBuffaloes.Rows.Count > 0 Then
            For i = 0 To dgvNoofBuffaloes.Rows.Count - 1
                If clsCommon.myLen(clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(colDesc).Value)) <= 0 AndAlso clsCommon.myLen(clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(COLTypeOfAnimal).Value)) > 0 Then
                    clsCommon.MyMessageBoxShow(Me, " Please Enter Bread Of Animal For Row No " & (i + 1) & ", It is Manadatory (Under Animal Details  Tab)", Me.Text)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    dgvNoofBuffaloes.Rows(i).Cells(colDesc).ErrorText = "Please Enter Bread Of Animal For Row No " & (i + 1) & ", It is Manadatory  "
                    Return False
                Else
                    dgvNoofBuffaloes.Rows(i).Cells(colDesc).ErrorText = ""
                End If
            Next

            For i = 0 To dgvNoofBuffaloes.Rows.Count - 1
                If clsCommon.myLen(clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(colDesc).Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(COLTypeOfAnimal).Value)) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, " Please Select Type Of Animal For Row No " & (i + 1) & ", It is Manadatory (Under Animal Details  Tab)", Me.Text)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    dgvNoofBuffaloes.Rows(i).Cells(colDesc).ErrorText = "Please Type Of Animal For Row No " & (i + 1) & ", It is Manadatory  "
                    Return False
                Else
                    dgvNoofBuffaloes.Rows(i).Cells(colDesc).ErrorText = ""
                End If
            Next
            'End If
            For i = 0 To dgvNoofBuffaloes.Rows.Count - 1
                If clsCommon.myLen(clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(colDesc).Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(COLTypeOfAnimal).Value)) > 0 AndAlso clsCommon.myCdbl(dgvNoofBuffaloes.Rows(i).Cells(colCountOfAnimal).Value) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, " Please Enter Count Of  '" & clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(COLTypeOfAnimal).Value) & "' For Row No " & (i + 1) & ", It is Manadatory (Under Animal Details  Tab)", Me.Text)
                    RadPageView1.SelectedPage = RadPageViewPage2
                    dgvNoofBuffaloes.Rows(i).Cells(colDesc).ErrorText = "Please Enter Count Of  '" & clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(COLTypeOfAnimal).Value) & "' For Row No " & (i + 1) & ", It is Manadatory (Under Animal Details  Tab)"
                    Return False
                Else
                    dgvNoofBuffaloes.Rows(i).Cells(colDesc).ErrorText = ""
                End If
            Next

            Dim TotCount As Integer = 0
            'If dgvNoofBuffaloes.Rows.Count > 0 Then
            For i = 0 To dgvNoofBuffaloes.Rows.Count - 1
                TotCount = TotCount + clsCommon.myCdbl(dgvNoofBuffaloes.Rows(i).Cells(colCountOfAnimal).Value)
            Next
            If TotCount <> clsCommon.myCdbl(txtNoOfBreadableMilkAnimal.Text) Then
                clsCommon.MyMessageBoxShow(Me, " Total of Animals must be equals to total milch animal count", Me.Text)
                RadPageView1.SelectedPage = RadPageViewPage2
                'dgvNoofBuffaloes.Rows(0).Cells(colDesc).ErrorText = "Total of Animals must be equals to total milch animal count"
                Return False
            End If
            'End If


            'If dgvNoofCows.Rows.Count > 0 Then
            '    For i = 0 To dgvNoofCows.Rows.Count - 1
            '        If clsCommon.myLen(clsCommon.myCstr(dgvNoofCows.Rows(i).Cells(colDesc).Value)) <= 0 Then
            '            clsCommon.MyMessageBoxShow(" Please Enter Bread Of Cow For Row No " & (i + 1) & ", It is Manadatory (Under Cow Details  Tab)")
            '            RadPageView1.SelectedPage = RadPageViewPage3
            '            dgvNoofCows.Rows(i).Cells(colDesc).ErrorText = "Please Enter Bread Of Cow For Row No " & (i + 1) & ", It is Manadatory  "
            '            Return False
            '        Else
            '            dgvNoofCows.Rows(i).Cells(colDesc).ErrorText = ""
            '        End If
            '    Next
            'End If
            If String.IsNullOrEmpty(txtTolerance.Text) = False Then
                If clsCommon.myCdbl(txtTolerance.Text) > 100 Then
                    clsCommon.MyMessageBoxShow(Me, "Tolerance value should be less then 100.", Me.Text)
                    txtTolerance.Text = ""
                    txtTolerance.Focus()
                    Return False
                End If
            End If

            If clsCommon.myLen(txtBankIFSC_M.Text) > 0 Then
                If SettBankIFSCCodeValidateByService Then
                    Dim arrFilter As New Dictionary(Of String, String)
                    arrFilter.Add("IFSC", txtBankIFSC_M.Text)
                    arrFilter.Add("OutPutType", "1")
                    Dim dt As DataTable = Xtra.GetDataFromAPI(EnumAPI.BankIFSC, "GetIFSC", arrFilter)
                    If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                        Throw New Exception("Invalid IFSC Code")
                    End If
                    If clsCommon.CompairString(dt.Rows(0)("Result"), "true") = CompairStringResult.Equal Then
                        txtNameOfBank_M.Text = clsCommon.myCstr(dt.Rows(0)("BANK"))
                        txtBankBranch_M.Text = clsCommon.myCstr(dt.Rows(0)("BRANCH"))
                        txtBankState_M.Text = clsCommon.myCstr(dt.Rows(0)("STATE"))
                        txtBankCity_M.Text = clsCommon.myCstr(dt.Rows(0)("CITY"))
                        txtBankIFSC_M.Text = clsCommon.myCstr(dt.Rows(0)("IFSC"))
                    Else
                        Throw New Exception(dt.Rows(0)("Response"))
                    End If
                End If
            End If

            If clsCommon.myLen(txtTelePhone.Text) > 0 AndAlso Not IsValidPhoneNumber(txtTelePhone.Text) Then
                txtTelePhone.Focus()
                Throw New Exception("Please enter a valid 10-digit phone number")
            End If

            UcCustomFields1.AllowToSave()
            Return True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Function

    Function IsValidPhoneNumber(phone As String) As Boolean
        ' Regex pattern for US phone numbers (you can adjust it as per the requirement)
        Dim pattern As String = "^\(?([0-9]{3})\)?[-.●]?([0-9]{3})[-.●]?([0-9]{4})$"
        Dim regex As New Regex(pattern)
        Return regex.IsMatch(phone)
    End Function

#Region "Event Routines"
    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click, btnClose.Click
        Try
            Me.Close()
            GC.Collect()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnReset.Click
        Try
            reset()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub FrmMPMaster_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim coll As New Dictionary(Of String, String)()
            coll.Add("THIRD_PARTY_CODE", "varchar(30) NULL")
            coll.Add("Third_Party_Source", "char(4) NULL")
            clsCommonFunctionality.CreateOrAlterTable(False, "TSPL_MP_MASTER", coll, Nothing, True)
            Try
                clsDBFuncationality.ExecuteNonQuery("CREATE UNIQUE INDEX Unique_THIRD_PARTY_CODE ON TSPL_MP_MASTER (THIRD_PARTY_CODE) WHERE THIRD_PARTY_CODE IS NOT NULL;")

            Catch ex As Exception
            End Try


        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try


        SetUserMgmtNew()

        ' EnableBankFromMaster = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.EnableBankFromMaster & "'")) = 0, False, True)
        SettBankIFSCCodeValidateByService = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.BankIFSCCodeValidateByService, clsFixedParameterCode.BankIFSCCodeValidateByService, Nothing)) > 1) ''Means 2 ERP or 3 Service And ERP

        SettJanAadharNoMandatory = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.JanAadharNoMandatory, clsFixedParameterCode.JanAadharNoMandatory, Nothing)) > 0)
        IncentiveAccNoMandatoryInMPMaster = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.IncentiveAccNoMandatoryInMPMaster, clsFixedParameterCode.IncentiveAccNoMandatoryInMPMaster, Nothing)) > 0)
        reset()
        ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update ")
        ButtonToolTip.SetToolTip(btnDelete, "Press Alt+D To Delete ")
        ButtonToolTip.SetToolTip(btnClose, "Press Alt+C To Close the Window")
        ButtonToolTip.SetToolTip(rbtnReset, "Press Alt+N For New")
        RadPageView1.Pages("pvpCustomFields").Item.Visibility = MyBase.customFieldTabProperty
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.Report_ID = MyBase.Form_ID
            UcCustomFields1.LoadCustomControls()
        End If
        SettNoOFIncentiveForImportExport = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.NoOFIncentiveForMPImportExport, clsFixedParameterCode.NoOFIncentiveForMPImportExport, Nothing))
        UseVLCUploaderCodeMPUploaderCodeInMCCProcurement = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.UseVLCUploaderCodeMPUploaderCodeInMCCProcurement, clsFixedParameterCode.UseVLCUploaderCodeMPUploaderCodeInMCCProcurement, Nothing)) = 1, True, False)
        btnUnverifiedJanAdhaar.Enabled = False
    End Sub

    Private Sub txtPinCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPinCode.KeyPress
        e.Handled = Not clsNumberValidate.IntValidate(e.KeyChar)
    End Sub

#End Region

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If allowToSave() Then
                save()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub save()
        If clsCommon.CompairString(btnSave.Text, "&Update") = CompairStringResult.Equal Then
            If MyBase.isUpdateFlag = False Then
                clsCommon.MyMessageBoxShow(Me, "Don't have permission to update MP Master.", Me.Text)
                Return
            End If
        End If

        If MyBase.isModifyonPasswordFlag Then
            If clsPasswordCheckForMasters.CheckMasterPwd(clsUserMgtCode.frmMPMaster, clsCommon.myCstr(objCommonVar.CurrentCompanyCode)) Then
            Else
                Return
            End If
        End If
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            obj = New clsMpMaster()
            If clsCommon.CompairString(btnSave.Text, "&Save") = CompairStringResult.Equal Then
                obj.isNewEntry = True
            Else
                obj.isNewEntry = False
            End If
            Dim dt As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
            If obj.isNewEntry Then
                obj.MP_Code = clsERPFuncationality.GetNextCode(trans, dt, clsDocType.MPMaster, "", "")
                If clsCommon.myLen(obj.MP_Code) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Error In Document Code Genertion", Me.Text)
                    Exit Sub
                End If
            Else
                obj.MP_Code = clsCommon.myCstr(fndMPCode.Value)
            End If
            fndMPCode.Value = obj.MP_Code
            obj.MP_CODE_VLC_UPLOADER = clsCommon.myCstr(txtMPCodeVlcUploader.Text)
            'txtMPCodeVlcUploader.ReadOnly = True
            obj.MP_Name = clsCommon.myCstr(txtMPName.Text)
            obj.MP_Name_Hindi = clsCommon.myCstr(txtMPNameHindi.Text)
            obj.DISTRICT_Code = clsCommon.myCstr(txtDistrict.Value)
            obj.BLOCK_CODE = clsCommon.myCstr(txtBlockCode.Value)
            obj.Zone_Code = clsCommon.myCstr(txtZone.Value)
            obj.REVENUE_VILLAGE_CODE = clsCommon.myCstr(txtRevenueVillage.Value)
            obj.GRAMPANCHAYAT_CODE = clsCommon.myCstr(txtGrampanchayat.Value)
            obj.PANCHAYAT_SAMITI_CODE = clsCommon.myCstr(txtPanchayatSamiti.Value)
            obj.VIDHAN_SABHA_CODE = clsCommon.myCstr(txtVidhanSabha.Value)
            obj.CAST_CATEGORY_CODE = clsCommon.myCstr(txtCastCategory.Value)
            obj.Villege_Code = clsCommon.myCstr(fndVillegeCode.Value)
            obj.Father_Name = clsCommon.myCstr(txtFatherName.Text)
            obj.Add1 = clsCommon.myCstr(txtAdd1.Text)
            obj.Add2 = clsCommon.myCstr(txtAdd2.Text)
            obj.Zila = clsCommon.myCstr(txtZila.Text)
            obj.Tehsil = clsCommon.myCstr(txtTehsil.Text)
            obj.City_code = clsCommon.myCstr(fndCityCode.Value)
            obj.State_Code = clsCommon.myCstr(fndStateCode.Value)
            obj.Country_code = clsCommon.myCstr(fndCountryCode.Value)
            obj.Pin_code = clsCommon.myCstr(txtPinCode.Text)
            obj.Telphone = clsCommon.myCstr(txtTelePhone.Text)
            obj.Email = clsCommon.myCstr(txtEmail.Text)
            obj.Fax = clsCommon.myCstr(txtFAX.Text)
            obj.Jan_Aadhar_No = txtJanAadharNo.Text
            obj.THIRD_PARTY_CODE = TxtThirdPartyCode.Text
            obj.Third_Party_Source = MylblThirdPartySource.Text
            obj.DOB = clsCommon.myCDate(dtpDOB.Value, "dd/MMM/yyyy")
            obj.Education = clsCommon.myCstr(txtEducation.Text)
            obj.Land_Holding = clsCommon.myCdbl(txtLandHolding.Text)
            'obj.No_Of_Buffaloes = clsCommon.myCdbl(txtNoofBuffaloes.Text)
            'obj.No_Of_Cows = clsCommon.myCdbl(txtNoofCows.Text)
            '         obj.No_Of_Animal = clsCommon.myCdbl(txtNoofAnimal.Text)
            obj.No_Of_Children_member = clsCommon.myCdbl(txtNoOfChildrenMember.Text)
            obj.No_Of_Adult_member = clsCommon.myCdbl(txtNoOfAdultMember.Text)
            obj.No_Of_Total_Dependent_member = clsCommon.myCdbl(txtNoOfTotalDependentMember.Text)
            obj.No_Of_breedable_milk_animal = clsCommon.myCdbl(txtNoOfBreadableMilkAnimal.Text)
            obj.Milk_production = clsCommon.myCdbl(txtMilkProduction.Text)
            obj.Milk_Home_consumption = clsCommon.myCdbl(txtHomeConsumption.Text)
            obj.Milk_For_sale = clsCommon.myCdbl(txtMilkAvlblForSale.Text)
            obj.MCC_Code = clsCommon.myCstr(fndVLCode.Value)
            obj.Cust_Account = fndCustAccSet.Value
            obj.Acct_Set_Code = fndVendorAccSet.Value
            ' Ticket : ERO/11/12/18-000430 By prabhakar Add TOLERANCE feild
            If String.IsNullOrEmpty(txtTolerance.Text) = False Then
                obj.TOLERANCE = txtTolerance.Text
            End If
            obj.Incentive_Account = txtIncentiveGLAccount.Value
            Dim i As Integer = 0
            'Dim objbuff As New clsBuffaloesDetails
            'obj.arrBuffaloesDetail = New List(Of clsBuffaloesDetails)
            'For i = 0 To dgvNoofBuffaloes.Rows.Count - 1
            '    objbuff = New clsBuffaloesDetails
            '    objbuff.Prog_Code = Me.Form_ID
            '    objbuff.Trans_Code = obj.MP_Code
            '    objbuff.Line_No = clsCommon.myCdbl(dgvNoofBuffaloes.Rows(i).Cells(colSlNO).Value)
            '    objbuff.Bread_of_Buffalo = clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(colDesc).Value)
            '    obj.arrBuffaloesDetail.Add(objbuff)
            'Next
            'Dim objcows As New clsCowDetails
            'obj.arrCowDetail = New List(Of clsCowDetails)
            'For i = 0 To dgvNoofCows.Rows.Count - 1
            '    objcows = New clsCowDetails
            '    objcows.Prog_Code = Me.Form_ID
            '    objcows.Trans_Code = obj.MP_Code
            '    objcows.Line_No = clsCommon.myCdbl(dgvNoofCows.Rows(i).Cells(colSlNO).Value)
            '    objcows.Bread_of_cow = clsCommon.myCstr(dgvNoofCows.Rows(i).Cells(colDesc).Value)
            '    obj.arrCowDetail.Add(objcows)
            'Next

            Dim objAnimal As New clsAnimalDetails
            obj.arrAnimalDetail = New List(Of clsAnimalDetails)
            For i = 0 To dgvNoofBuffaloes.Rows.Count - 1
                objAnimal = New clsAnimalDetails
                objAnimal.Prog_Code = Me.Form_ID
                objAnimal.Trans_Code = obj.MP_Code
                objAnimal.Line_No = clsCommon.myCdbl(dgvNoofBuffaloes.Rows(i).Cells(colSlNO).Value)
                objAnimal.Bread_of_Animal = clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(colDesc).Value)
                objAnimal.Type_Of_Animal = clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(COLTypeOfAnimal).Value)
                objAnimal.Count_Of_animal = clsCommon.myCdbl(dgvNoofBuffaloes.Rows(i).Cells(colCountOfAnimal).Value)
                objAnimal.Animal_Staus = clsCommon.myCstr(dgvNoofBuffaloes.Rows(i).Cells(colStatus).Value)
                obj.arrAnimalDetail.Add(objAnimal)
            Next
            obj.PayeeName = clsCommon.myCstr(txtPayeeName.Text)
            Dim IFSCText As String
            If EnableBankFromMaster = True Then
                obj.BankName = clsCommon.myCstr(txtBankCode.Value)
                obj.BankBranch = clsCommon.myCstr(txtBankBranch.Value)
                obj.BankCityCode = clsCommon.myCstr(fndBankCity.Value)
                obj.BankStateCode = clsCommon.myCstr(fndBankState.Value)
                IFSCText = txtIFCICode.Text
                obj.IFCICode = clsCommon.myCstr(txtIFCICode.Text)
            Else
                obj.BankName = clsCommon.myCstr(txtNameOfBank_M.Text)
                obj.BankBranch = clsCommon.myCstr(txtBankBranch_M.Text)
                obj.BankCityCode = clsCommon.myCstr(txtBankCity_M.Text)
                obj.BankStateCode = clsCommon.myCstr(txtBankState_M.Text)
                IFSCText = txtBankIFSC_M.Text
            End If
            IFSCText = IFSCText.ToUpper()
            obj.IFCICode = clsCommon.myCstr(IFSCText)

            obj.AccountNO = clsCommon.myCstr(txtAccountNo.Text)
            obj.Account_Type = clsCommon.myCstr(cmbAccount_Type.Text)
            obj.Modified_By = objCommonVar.CurrentUserCode
            obj.Modified_Date = clsCommon.GetPrintDate(dt, "dd/MMM/yyyy")
            obj.Comp_Code = objCommonVar.CurrentCompanyCode
            If obj.isNewEntry Then
                obj.Created_By = objCommonVar.CurrentUserCode
                obj.Created_Date = clsCommon.GetPrintDate(dt, "dd/MMM/yyyy")
            End If
            ''For Custom Fields
            obj.Form_Id = MyBase.Form_ID
            obj.arrCustomFields = New List(Of clsCustomFieldValues)
            If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                UcCustomFields1.GetData(obj.arrCustomFields)
            End If
            'If MyBase.ArrDetailFields IsNot Nothing AndAlso MyBase.ArrDetailFields.Count > 0 Then
            '    clsCustomFieldGrid.GetData(obj.arrCustomFields, gv1, MyBase.ArrDetailFields, colICode)
            'End If
            ''End of For Custom Fields

            '===================================
            obj.IsVSP = clsCommon.myCdbl(chkIsVSP.Checked)
            obj.InActive = clsCommon.myCdbl(chkInActive.Checked)
            obj.TypeOfFormer = clsCommon.myCstr(ddlTypeOfFormer.Text)
            obj.Gender = clsCommon.myCstr(ddlGender.Text)
            obj.MaritalStatus = clsCommon.myCstr(CboMaritalStatus.Text)
            'Dim converter As New ImageConverter
            'obj.MpPicture = converter.ConvertTo(UcCamControl1.PicBox.Image, GetType(Byte()))



            '=====================================
            obj.Form_Id = MyBase.Form_ID
            obj.ArrMPIncentiveMapping = txtIncentiveMult.arrValueMember
            If clsMpMaster.SaveData(obj, trans) Then
                trans.Commit()
                UcAttachment1.SaveData(obj.MP_Code)

                If clsCommon.CompairString(btnSave.Text, "&Save") = CompairStringResult.Equal Then '-----------27/06/2014 Monika
                    File_Name = UcCamControl1.PicBox.Tag
                    If clsCommon.myLen(File_Name) > 0 Then
                        Dim bData As Byte()
                        Dim br As BinaryReader = New BinaryReader(System.IO.File.OpenRead(clsCommon.myCstr(File_Name)))
                        bData = br.ReadBytes(br.BaseStream.Length)

                        Dim Str As String = " UPDATE TSPL_MP_MASTER set MP_Picture = @BLOBData where MP_Code='" + obj.MP_Code + "'"
                        Dim cmd As SqlCommand = New SqlCommand(Str, clsDBFuncationality.GetConnnection)
                        Dim prm As New SqlParameter("@BLOBData", bData)
                        cmd.Parameters.Add(prm)
                        cmd.ExecuteNonQuery()
                        br.Close() ' done by stuti reagrding memory leakage
                    End If
                    clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                Else
                    File_Name = UcCamControl1.PicBox.Tag
                    If clsCommon.myLen(File_Name) > 0 Then
                        Dim bData As Byte()
                        Dim br As BinaryReader = New BinaryReader(System.IO.File.OpenRead(clsCommon.myCstr(File_Name)))
                        bData = br.ReadBytes(br.BaseStream.Length)

                        Dim Str As String = " UPDATE TSPL_MP_MASTER set MP_Picture = @BLOBData where MP_Code='" + obj.MP_Code + "'"
                        Dim cmd As SqlCommand = New SqlCommand(Str, clsDBFuncationality.GetConnnection)
                        Dim prm As New SqlParameter("@BLOBData", bData)
                        cmd.Parameters.Add(prm)
                        cmd.ExecuteNonQuery()
                        br.Close() ' done by stuti reagrding memory leakage
                    End If
                    clsCommon.MyMessageBoxShow(Me, "Data Updated Successfully", Me.Text)
                End If

                loadData(obj.MP_Code, NavigatorType.Current)
                btnSave.Text = "&Update"
                fndMPCode.MyReadOnly = True
                btnDelete.Enabled = True
                Exit Sub

            End If
            clsCommon.MyMessageBoxShow(Me, "Data Not Saved ", Me.Text)
            btnSave.Text = "&Save"
            btnDelete.Enabled = False
            fndMPCode.MyReadOnly = False
            trans.Rollback()

        Catch ex As Exception
            trans.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndMPCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles fndMPCode._MYValidating
        Try
            Dim whrcls As String = ""
            fndMPCode.Value = clsMpMaster.getFinder("", fndMPCode.Value, isButtonClicked)
            If clsCommon.myLen(fndMPCode.Value) > 0 Then
                loadData(fndMPCode.Value, NavigatorType.Current)
                Dim isJanAdhaarVerified As Integer = clsDBFuncationality.getSingleValue("select isnull(Jan_Aadhar_No_Verified,0) as Is_Verified from tspl_mp_master  where MP_Code = '" & fndMPCode.Value & "'")
                If isJanAdhaarVerified = 1 Then
                    btnUnverifiedJanAdhaar.Enabled = True
                Else
                    btnUnverifiedJanAdhaar.Enabled = False
                End If
                btnSave.Text = "&Update"
                btnDelete.Enabled = True
                fndMPCode.MyReadOnly = True
                UcAttachment1.LoadData(fndMPCode.Value)
            Else
                reset()
                fndMPCode.MyReadOnly = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub loadData(ByVal strCode As String, ByVal nav As NavigatorType)
        Try
            Dim obj As clsMpMaster = clsMpMaster.loadData(fndMPCode.Value, nav, Me.Form_ID)
            If obj IsNot Nothing Then
                fndMPCode.Value = obj.MP_Code
                fndVLCode.Value = obj.MCC_Code
                txtMPCodeVlcUploader.Text = clsCommon.myCstr(obj.MP_CODE_VLC_UPLOADER)
                'txtMPCodeVlcUploader.ReadOnly = True
                txtVLCName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select vlc_name from TSPL_VLC_MASTER_HEAD where vlc_code='" + fndVLCode.Value + "'"))
                txtMPName.Text = obj.MP_Name
                txtMPNameHindi.Text = obj.MP_Name_Hindi
                txtCastCategory.Value = obj.CAST_CATEGORY_CODE
                lblCastCategory.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select case when  CAST_CATEGORY_CODE = CAST_CATEGORY_NAME and   len (isnull( DESCRIPTION,'')) > 0 then  DESCRIPTION else  CAST_CATEGORY_NAME end Name  from TSPL_CAST_CATEGORY_MASTER where CAST_CATEGORY_CODE = '" + clsCommon.myCstr(obj.CAST_CATEGORY_CODE) + "' "))
                txtDistrict.Value = obj.DISTRICT_Code
                lblDistrict.Text = clsDistrictMaster.GetName(obj.DISTRICT_Code)
                txtBlockCode.Value = obj.BLOCK_CODE
                lblBlockCode.Text = clsBlockMaster.GetName(obj.BLOCK_CODE)
                txtZone.Value = obj.Zone_Code
                lblZone.Text = ClsZoneMaster.GetName(obj.Zone_Code)
                'TxtThirdPartyCode.Text = ClsZoneMaster.GetName(obj.ThirdPartyCode)
                txtRevenueVillage.Value = obj.REVENUE_VILLAGE_CODE
                lblRevenueVillage.Text = clsRevenueVillageMaster.GetName(obj.REVENUE_VILLAGE_CODE)
                txtGrampanchayat.Value = obj.GRAMPANCHAYAT_CODE
                lblGrampanchayat.Text = clsGrampanchayatMaster.GetName(obj.GRAMPANCHAYAT_CODE)
                txtPanchayatSamiti.Value = obj.PANCHAYAT_SAMITI_CODE
                lblPanchayatSamiti.Text = clsPanchayatSamitiMaster.GetName(obj.PANCHAYAT_SAMITI_CODE)
                txtVidhanSabha.Value = obj.VIDHAN_SABHA_CODE
                lblVidhanSabha.Text = clsVidhanSabhaMaster.GetName(obj.VIDHAN_SABHA_CODE)
                txtAdd1.Text = obj.Add1
                txtAdd2.Text = obj.Add2
                txtZila.Text = obj.Zila
                txtTehsil.Text = obj.Tehsil
                fndCityCode.Value = obj.City_code
                'txtCastCategory.Value = obj.CAST_CATEGORY_CODE
                If clsCommon.myLen(fndCityCode.Value) > 0 Then
                    txtCityName.Text = clsDBFuncationality.getSingleValue("select city_name from tspl_city_master where city_code='" & obj.City_code & "'")
                Else
                    txtCityName.Text = ""
                End If
                fndStateCode.Value = obj.State_Code
                If clsCommon.myLen(fndStateCode.Value) > 0 Then
                    txtStateName.Text = clsDBFuncationality.getSingleValue("select state_name from tspl_state_master where state_code='" & obj.State_Code & "'")
                Else
                    txtStateName.Text = ""
                End If
                fndCountryCode.Value = obj.Country_code
                If clsCommon.myLen(fndCountryCode.Value) > 0 Then
                    txtCountryName.Text = clsDBFuncationality.getSingleValue("select country_name from tspl_country_master where country_code='" & obj.Country_code & "'")
                Else
                    txtCountryName.Text = ""
                End If
                txtPinCode.Text = obj.Pin_code
                txtTelePhone.Text = obj.Telphone
                txtEmail.Text = obj.Email
                txtFAX.Text = obj.Fax
                txtJanAadharNo.Text = obj.Jan_Aadhar_No
                TxtThirdPartyCode.Text = obj.THIRD_PARTY_CODE
                MylblThirdPartySource.Text = obj.third_party_source
                dtpDOB.Value = obj.DOB
                txtEducation.Text = obj.Education
                txtLandHolding.Text = obj.Land_Holding
                'txtNoofBuffaloes.Text = obj.No_Of_Buffaloes
                'txtNoofCows.Text = obj.No_Of_Cows
                '            txtNoofAnimal.Text = clsCommon.myCdbl(obj.No_Of_Animal)
                txtNoOfChildrenMember.Text = obj.No_Of_Children_member
                txtNoOfAdultMember.Text = clsCommon.myCdbl(obj.No_Of_Adult_member)
                txtNoOfTotalDependentMember.Text = clsCommon.myCdbl(obj.No_Of_Total_Dependent_member)
                txtNoOfBreadableMilkAnimal.Text = clsCommon.myCdbl(obj.No_Of_breedable_milk_animal)
                txtMilkProduction.Text = clsCommon.myCdbl(obj.Milk_production)
                txtHomeConsumption.Text = clsCommon.myCdbl(obj.Milk_Home_consumption)
                txtMilkAvlblForSale.Text = clsCommon.myCdbl(obj.Milk_For_sale)
                fndCustAccSet.Value = obj.Cust_Account
                fndVendorAccSet.Value = obj.Acct_Set_Code
                txtCustAccSetDesc.Text = obj.Cust_Account_Desc
                txtVendorAccDesc.Text = obj.Vendor_Acc_Code_Desc

                If String.IsNullOrEmpty(obj.TOLERANCE) = False Then
                    txtTolerance.Text = obj.TOLERANCE
                End If
                txtIncentiveMult.arrValueMember = obj.ArrMPIncentiveMapping
                txtIncentiveGLAccount.Value = obj.Incentive_Account
                lblIncentiveGLAccount.Text = clsGLAccount.GetName(obj.Incentive_Account)
                '================ Ticket No : ERO/23/09/19-001035 By Prabhakar =================
                If clsCommon.myLen(fndVLCode.Value) > 0 Then
                    txtVLCName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select vlc_name from TSPL_VLC_MASTER_HEAD where vlc_code='" + fndVLCode.Value + "'"))
                    Dim qry As String = " select TSPL_VLC_MASTER_HEAD.vlc_name,TSPL_VLC_MASTER_HEAD.VSP_Code,tspl_vendor_master.Vendor_Name as VSP_Name, TSPL_VLC_MASTER_HEAD.MCC, TSPL_MCC_MASTER.MCC_NAME, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader from TSPL_VLC_MASTER_HEAD LEFT OUTER JOIN TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_CODE = TSPL_VLC_MASTER_HEAD.MCC LEFT OUTER JOIN tspl_vendor_master ON tspl_vendor_master.Vendor_Code = TSPL_VLC_MASTER_HEAD.VSP_Code  where vlc_code= '" + fndVLCode.Value + "' "
                    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                        txtMCC.Value = clsCommon.myCstr(dt.Rows(0)("MCC"))
                        lblMCCName.Text = clsCommon.myCstr(dt.Rows(0)("MCC_NAME"))
                        txtVlCUploader.Text = clsCommon.myCstr(dt.Rows(0)("VLC_Code_VLC_Uploader"))
                        txtVSPName.Text = clsCommon.myCstr(dt.Rows(0)("VSP_Name"))
                    End If
                Else
                    txtVLCName.Text = ""
                    lblMCCName.Text = ""
                    txtVlCUploader.Text = ""
                    txtVSPName.Text = ""
                End If
                '==================================

                'If obj.arrBuffaloesDetail.Count > 0 Then
                '    loadBlankDgvBuffaloes()
                '    For i As Integer = 0 To obj.arrBuffaloesDetail.Count - 1
                '        dgvNoofBuffaloes.Rows.Add(obj.arrBuffaloesDetail.Item(i).Line_No, obj.arrBuffaloesDetail.Item(i).Bread_of_Buffalo)
                '    Next
                '    dgvNoofBuffaloes.BestFitColumns()
                'End If
                'If obj.arrCowDetail.Count > 0 Then
                '    loadBlankDgvCows()
                '    For i As Integer = 0 To obj.arrCowDetail.Count - 1
                '        dgvNoofCows.Rows.Add(obj.arrCowDetail.Item(i).Line_No, obj.arrCowDetail.Item(i).Bread_of_cow)
                '    Next
                '    dgvNoofCows.BestFitColumns()
                'End If
                loadBlankDgvBuffaloes()
                If obj.arrAnimalDetail IsNot Nothing AndAlso obj.arrAnimalDetail.Count > 0 Then
                    For i As Integer = 0 To obj.arrAnimalDetail.Count - 1
                        dgvNoofBuffaloes.Rows.Add(obj.arrAnimalDetail.Item(i).Line_No, obj.arrAnimalDetail.Item(i).Type_Of_Animal, obj.arrAnimalDetail.Item(i).Count_Of_animal, obj.arrAnimalDetail.Item(i).Bread_of_Animal, obj.arrAnimalDetail.Item(i).Animal_Staus)
                    Next
                    ' dgvNoofBuffaloes.Rows.AddNew()
                    dgvNoofBuffaloes.BestFitColumns()
                End If
                txtPayeeName.Text = obj.PayeeName
                If EnableBankFromMaster = True Then
                    pnlBankDetailsManual.Visible = False
                    txtBankCode.Value = obj.BankName
                    If clsCommon.myLen(txtBankCode.Value) > 0 Then
                        txtBankName.Text = clsDBFuncationality.getSingleValue("select Bank_Name  from tspl_vendor_bank_master where bank_code='" & obj.BankName & "'")
                    Else
                        txtBankName.Text = ""
                    End If
                    txtBankBranch.Value = obj.BankBranch
                    'If clsCommon.myLen(txtBankBranch.Value) > 0 Then
                    '    txtBankBranchName.Text = clsDBFuncationality.getSingleValue("select Branch_Name  from tspl_vendor_bank_master where bank_code='" & obj.BankName & "'")
                    'End If

                    fndBankCity.Value = obj.BankCityCode
                    If clsCommon.myLen(fndBankCity.Value) > 0 Then
                        txtBankCityName.Text = clsDBFuncationality.getSingleValue("select city_name from tspl_city_master where city_code='" & obj.BankCityCode & "'")
                    Else
                        txtBankCityName.Text = ""
                    End If
                    fndBankState.Value = obj.BankStateCode
                    If clsCommon.myLen(fndBankState.Value) > 0 Then
                        txtBankStateName.Text = clsDBFuncationality.getSingleValue("select state_name from tspl_state_master where state_code='" & obj.BankStateCode & "'")
                    Else
                        txtBankStateName.Text = ""
                    End If
                    txtIFCICode.Text = obj.IFCICode
                Else
                    pnlBankDetailsManual.Visible = True
                    txtNameOfBank_M.Text = obj.BankName
                    txtBankBranch_M.Text = obj.BankBranch
                    txtBankCity_M.Text = obj.BankCityCode
                    txtBankState_M.Text = obj.BankStateCode
                    txtBankIFSC_M.Text = obj.IFCICode
                End If
                txtAccountNo.Text = obj.AccountNO
                If clsMpMaster.IsDBTDone(obj.MP_Code, Nothing) Then
                    txtBankIFSC_M.Enabled = False
                    txtAccountNo.Enabled = False
                End If
                cmbAccount_Type.Text = obj.Account_Type
                btnSave.Text = "&Update"
                btnDelete.Enabled = True
                fndMPCode.MyReadOnly = True
                'txtMPCodeVlcUploader.ReadOnly = True
            Else
                btnDelete.Enabled = False
                btnSave.Text = "&Save"
                fndMPCode.MyReadOnly = False
                'txtMPCodeVlcUploader.ReadOnly = False
            End If


            '===================================================================
            chkIsVSP.Checked = clsCommon.myCdbl(obj.IsVSP)
            chkInActive.Checked = clsCommon.myCdbl(obj.InActive)
            ddlTypeOfFormer.Text = clsCommon.myCstr(obj.TypeOfFormer)
            ddlGender.Text = clsCommon.myCstr(obj.Gender)
            CboMaritalStatus.Text = clsCommon.myCstr(obj.MaritalStatus)
            fndVillegeCode.Value = clsCommon.myCstr(obj.Villege_Code)
            If clsCommon.myLen(fndVillegeCode.Value) > 0 Then
                txtVillegeName.Text = clsDBFuncationality.getSingleValue("select TSPL_VILLAGE_MASTER.Village_Name  from TSPL_VILLAGE_MASTER  where Village_Code='" & obj.Villege_Code & "'")
            Else
                txtVillegeName.Text = ""
            End If
            SetDefaultBillDetails()
            'Dim Filename As Byte( UcCamControl1.PicBox.Image)
            Try
                File_Name = UcCamControl1.PicBox.Tag
                '============ Display Image====================
                Dim Filename As Byte() = clsDBFuncationality.getSingleValue("select MP_Picture from TSPL_MP_MASTER where MP_Code='" & fndMPCode.Value & "'")

                Using ms As New IO.MemoryStream(CType(Filename, Byte()))
                    Dim img As Image = Image.FromStream(ms)
                    UcCamControl1.PicBox.Image = img
                End Using
                '=============================================
            Catch ex As Exception
                UcCamControl1.PicBox.Image = Nothing
            End Try
            If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                UcCustomFields1.BlankAllControls()
                UcCustomFields1.LoadData(obj.MP_Code)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndMPCode__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles fndMPCode._MYNavigator
        Try
            Dim qst As String = " select tspl_mp_master.Mp_Code as [Code]   From tspl_mp_master   where 2=2"
            Select Case NavType
                Case NavigatorType.Current
                    qst += " and tspl_mp_master.Mp_Code in ('" + fndMPCode.Value + "')"
                Case NavigatorType.Next
                    qst += " and tspl_mp_master.Mp_Code in (select min(Mp_Code ) from tspl_mp_master where Mp_Code  >'" + fndMPCode.Value + "')"
                Case NavigatorType.First
                    qst += " and tspl_mp_master.Mp_Code in (select MIN(Mp_Code ) from tspl_mp_master)"
                Case NavigatorType.Last
                    qst += " and tspl_mp_master.Mp_Code in (select Max(Mp_Code ) from tspl_mp_master)"
                Case NavigatorType.Previous
                    qst += " and tspl_mp_master.Mp_Code in (select Max(Mp_Code ) from tspl_mp_master where Mp_Code  <'" + fndMPCode.Value + "')"
            End Select
            fndMPCode.Value = clsDBFuncationality.getSingleValue(qst)
            If clsCommon.myLen(fndMPCode.Value) > 0 Then
                'loadData(fndMPCode.Value, NavType) comment by Monika(27/06/2014),above navigator queries already gives the correct fndmpcode,and so no need to put here navtype,it result in no data found.
                loadData(fndMPCode.Value, NavigatorType.Current)
                btnSave.Text = "&Update"
                fndMPCode.MyReadOnly = True
                btnDelete.Enabled = True
                UcAttachment1.LoadData(fndMPCode.Value)
            Else
                btnSave.Text = "&Save"
                btnDelete.Enabled = False
                fndMPCode.MyReadOnly = False
                reset()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub fndCountryCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndCountryCode._MYValidating
        'Try
        '    fndCountryCode.Value = clsCountryMaster.getFinder("", fndCountryCode.Value, isButtonClicked)
        '    If clsCommon.myLen(fndCountryCode.Value) > 0 Then
        '        txtCountryName.Text = clsDBFuncationality.getSingleValue("select Country_name from tspl_country_master where country_code='" & fndCountryCode.Value & "'")
        '        fndStateCode.Value = ""
        '        txtStateName.Text = ""
        '        fndCityCode.Value = ""
        '        txtCityName.Text = ""
        '    Else
        '        fndCountryCode.Value = ""
        '        txtCountryName.Text = ""
        '        fndStateCode.Value = ""
        '        txtStateName.Text = ""
        '        fndCityCode.Value = ""
        '        txtCityName.Text = ""
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        'End Try
    End Sub

    Private Sub fndStateCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndStateCode._MYValidating
        'Try
        '    If clsCommon.myLen(fndCountryCode.Value) <= 0 Then
        '        clsCommon.MyMessageBoxShow("Please Select Country First..")
        '        fndCountryCode.Focus()
        '        Exit Sub
        '    End If
        '    Dim whrcls As String = "country_code='" & clsCommon.myCstr(fndCountryCode.Value) & "'"

        '    fndStateCode.Value = clsStateMaster.getFinder(whrcls, fndStateCode.Value, isButtonClicked)
        '    If clsCommon.myLen(fndStateCode.Value) > 0 Then
        '        txtStateName.Text = clsDBFuncationality.getSingleValue("select state_name from tspl_state_master where state_code='" & fndStateCode.Value & "'")
        '        fndCityCode.Value = ""
        '        txtCityName.Text = ""
        '    Else
        '        fndStateCode.Value = ""
        '        txtStateName.Text = ""
        '        fndCityCode.Value = ""
        '        txtCityName.Text = ""
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        'End Try
    End Sub

    Private Sub fndCityCode__MYOpenMasterForm(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndCityCode._MYOpenMasterForm
        Frm_Open = New frmCityMaster(objCommonVar.CurrentUser, objCommonVar.CurrentCompanyName)
        Frm_Open.SetUserMgmt(clsUserMgtCode.cityMaster)
        Frm_Open.Show()

    End Sub

    Private Sub fndCityCode__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndCityCode._MYValidating
        Try
            If clsCommon.myLen(fndCountryCode.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select Country First..", Me.Text)
                fndCountryCode.Focus()
                Exit Sub
            End If
            'If clsCommon.myLen(fndStateCode.Value) <= 0 Then
            '    clsCommon.MyMessageBoxShow("Please Select state First..")
            '    fndStateCode.Focus()
            '    Exit Sub
            'End If
            Dim whrcls As String = "" '"state_code='" & clsCommon.myCstr(fndStateCode.Value) & "'"
            fndCityCode.Value = clsCityMaster.getFinder(whrcls, fndCityCode.Value, isButtonClicked)
            'If clsCommon.myLen(fndCityCode.Value) > 0 Then
            '    txtCityName.Text = clsDBFuncationality.getSingleValue("select  city_name from tspl_city_master where city_code='" & fndCityCode.Value & "'")
            'End If
            If clsCommon.myLen(fndCityCode.Value) > 0 Then
                'txtcity.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select city_name from TSPL_CITY_MASTER where city_code='" + txtcitycode.Value + "' and state_code='" + txtstatecode.Value + "'"))
                txtCityName.Text = clsCommon.myCstr(clsCityMaster.GetName(fndCityCode.Value))
                fndStateCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select state_code from TSPL_CITY_MASTER where city_code='" + fndCityCode.Value + "' "))
                txtStateName.Text = clsCommon.myCstr(clsStateMaster.GetName(fndStateCode.Value))
                fndCountryCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select country_code from TSPL_state_MASTER where state_code='" + fndStateCode.Value + "' "))
                txtCountryName.Text = clsCommon.myCstr(clsCountryMaster.GetName(fndCountryCode.Value, Nothing))
                fndCountryCode.Enabled = False
                fndStateCode.Enabled = False
            Else
                txtCityName.Text = ""
                fndStateCode.Value = ""
                txtStateName.Text = ""
                fndCountryCode.Value = "INDIA"
                txtCountryName.Text = "INDIA"
                fndCountryCode.Enabled = True
                fndStateCode.Enabled = True
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtBankCode__MYOpenMasterForm(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtBankCode._MYOpenMasterForm
        Frm_Open = New FrmVendorBankMaster()
        Frm_Open.Show()
    End Sub

    Private Sub txtBankCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtBankCode._MYValidating
        Try
            'txtBankCode.Value = clsBankMaster.getFinder("", txtBankCode.Value, isButtonClicked)
            'If clsCommon.myLen(txtBankCode.Value) >= 0 Then
            '    txtBankName.Text = clsDBFuncationality.getSingleValue(" select description from tspl_bank_master where bank_code='" & txtBankCode.Value & "'")
            '    fndBankCity.Value = clsDBFuncationality.getSingleValue(" select city from tspl_bank_master where bank_code='" & txtBankCode.Value & "'")
            '    txtBankCityName.Text = clsDBFuncationality.getSingleValue("select city_name from tspl_city_master where city_code='" & fndBankCity.Value & "'")
            '    fndBankState.Value = clsDBFuncationality.getSingleValue(" select state from tspl_bank_master where bank_code='" & txtBankCode.Value & "'")
            '    txtBankStateName.Text = clsDBFuncationality.getSingleValue("select state_name from tspl_state_master where state_code='" & fndBankState.Value & "'")
            '    txtAccountNo.Text = clsDBFuncationality.getSingleValue("select BANKACCNUMBER from tspl_bank_master where bank_code='" & txtBankCode.Value & "'")
            'End If

            txtBankCode.Value = clsVendorBankMaster.GetFinder("", txtBankCode.Value, isButtonClicked)
            If clsCommon.myLen(txtBankCode.Value) > 0 Then
                Dim obj As clsVendorBankMaster = clsVendorBankMaster.GetData(txtBankCode.Value, NavigatorType.Current)
                If obj Is Nothing Then
                    Exit Sub
                End If
                txtBankName.Text = obj.Bank_Name
                fndBankState.Value = obj.state_code
                txtBankStateName.Text = obj.state_name
                fndBankCity.Value = obj.city_code
                txtBankCityName.Text = obj.city_name
                txtBankBranch.Value = obj.Branch_Code
                txtBankBranchName.Text = obj.Branch_Name
                txtIFCICode.Text = obj.IFSC_Code
            Else
                txtBankName.Text = ""
                fndBankState.Value = ""
                txtBankStateName.Text = ""
                fndBankCity.Value = ""
                txtBankCityName.Text = ""
                txtBankBranch.Value = ""
                txtBankBranchName.Text = ""
                txtIFCICode.Text = ""
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndBankState__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles fndBankState._MYValidating
        'Try
        '    fndBankState.Value = clsStateMaster.getFinder("", fndBankState.Value, isButtonClicked)
        '    If clsCommon.myLen(fndBankState.Value) > 0 Then
        '        txtBankStateName.Text = clsDBFuncationality.getSingleValue("select state_name from tspl_state_master where state_code='" & fndBankState.Value & "'")
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        'End Try
    End Sub

    Private Sub fndBankCity__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles fndBankCity._MYValidating
        'Try
        '    If clsCommon.myLen(fndBankState.Value) <= 0 Then
        '        clsCommon.MyMessageBoxShow("Please Select state of Bank  First..")
        '        fndBankState.Focus()
        '        Exit Sub
        '    End If
        '    Dim whrcls As String = "state_code='" & clsCommon.myCstr(fndBankState.Value) & "'"
        '    fndBankCity.Value = clsCityMaster.getFinder(whrcls, fndBankCity.Value, isButtonClicked)
        '    If clsCommon.myLen(fndBankCity.Value) > 0 Then
        '        txtBankCityName.Text = clsDBFuncationality.getSingleValue("select  city_name from tspl_city_master where city_code='" & fndBankCity.Value & "'")
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
        'End Try
    End Sub

    'Private Sub btnBuffaloesGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBuffaloesGo.Click
    '    Try
    '        If clsCommon.myCdbl(txtNoofAnimal.Text) > 0 Then

    '            Dim i As Integer = 0
    '            If clsCommon.myCdbl(txtNoofAnimal.Text) > dgvNoofBuffaloes.Rows.Count Then
    '                For i = dgvNoofBuffaloes.Rows.Count + 1 To clsCommon.myCdbl(txtNoofAnimal.Text)
    '                    dgvNoofBuffaloes.Rows.AddNew()
    '                    dgvNoofBuffaloes.Rows(i - 1).Cells(colSlNO).Value = i
    '                    dgvNoofBuffaloes.Rows(i - 1).Cells(colSlNO).ReadOnly = True
    '                Next
    '            ElseIf clsCommon.myCdbl(txtNoofAnimal.Text) < dgvNoofBuffaloes.Rows.Count Then
    '                For i = dgvNoofBuffaloes.Rows.Count - 1 To clsCommon.myCdbl(txtNoofAnimal.Text) Step -1
    '                    dgvNoofBuffaloes.Rows.RemoveAt(i)
    '                Next
    '            End If
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
    '    End Try
    'End Sub

    Private Sub dgvNoofBuffaloes_UserAddedRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles dgvNoofBuffaloes.UserAddedRow
        UpdateSLNo()
    End Sub

    'Private Sub btnNoofCowsgo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNoofCowsgo.Click
    '    Try
    '        If clsCommon.myCdbl(txtNoofCows.Text) > 0 Then

    '            Dim i As Integer = 0
    '            If clsCommon.myCdbl(txtNoofCows.Text) > dgvNoofCows.Rows.Count Then
    '                For i = dgvNoofCows.Rows.Count + 1 To clsCommon.myCdbl(txtNoofCows.Text)
    '                    dgvNoofCows.Rows.AddNew()
    '                    dgvNoofCows.Rows(i - 1).Cells(colSlNO).Value = i
    '                    dgvNoofCows.Rows(i - 1).Cells(colSlNO).ReadOnly = True
    '                Next
    '            ElseIf clsCommon.myCdbl(txtNoofBuffaloes.Text) < dgvNoofBuffaloes.Rows.Count Then
    '                For i = dgvNoofCows.Rows.Count - 1 To clsCommon.myCdbl(txtNoofCows.Text) Step -1
    '                    dgvNoofCows.Rows.RemoveAt(i)
    '                Next
    '            End If
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
    '    End Try
    'End Sub

    'Private Sub dgvNoofCows_UserDeletedRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles dgvNoofCows.UserDeletedRow
    '    Try
    '        Dim i As Integer = 0
    '        For i = 0 To dgvNoofCows.Rows.Count - 1
    '            dgvNoofCows.Rows(i).Cells(colSlNO).Value = (i + 1)
    '            dgvNoofCows.Rows(i).Cells(colDesc).ReadOnly = True
    '        Next
    '        txtNoofCows.Text = dgvNoofCows.Rows.Count
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
    '    End Try
    'End Sub

    'Private Sub dgvNoofCows_UserDeletingRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles dgvNoofCows.UserDeletingRow
    '    Try
    '        If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
    '            e.Cancel = True
    '        End If
    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(me,ex.Message,me.text)
    '    End Try
    'End Sub

    Private Sub dgvNoofBuffaloes_UserDeletedRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles dgvNoofBuffaloes.UserDeletedRow
        Try
            UpdateSLNo()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub dgvNoofBuffaloes_UserDeletingRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles dgvNoofBuffaloes.UserDeletingRow
        Try
            If common.clsCommon.MyMessageBoxShow(Me, "Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim trans As SqlTransaction = Nothing
        Try
            trans = clsDBFuncationality.GetTransactin()
            If clsMpMaster.deleteData(fndMPCode.Value, Me.Form_ID, trans) Then
                clsCommon.MyMessageBoxShow(Me, "Deleted Successfully", Me.Text)
                trans.Commit()
                reset()
            Else
                clsCommon.MyMessageBoxShow(Me, "Delete Unsuccessful.", Me.Text)
                trans.Rollback()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            trans.Rollback()
        End Try
    End Sub

    Private Sub fndVLCode__MYOpenMasterForm(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndVLCode._MYOpenMasterForm
        Frm_Open = New FrmVLCMaster()
        Frm_Open.SetUserMgmt(clsUserMgtCode.frmVLCMaster)
        Frm_Open.Show()
    End Sub



    Private Sub fndMCCCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles fndVLCode._MYValidating
        Try
            Dim whrcls As String = ""
            If clsCommon.myLen(arrLoc) > 0 Then
                whrcls = " TSPL_VLC_MASTER_HEAD.mcc in (" + arrLoc + ")"
            End If
            Dim qry As String = "select TSPL_VLC_MASTER_HEAD.vlc_code as [Code],TSPL_VLC_MASTER_HEAD.vlc_name as [VLC Name],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [VLC Uploader Code],TSPL_VLC_MASTER_HEAD.vehical_name as [Vehicle Name],TSPL_VLC_MASTER_HEAD.village_code as [Village Code],tspl_village_master.village_name as [Village Name],TSPL_VLC_MASTER_HEAD.route_code as [Route Code],tspl_mcc_route_master.route_name as [Route Name],TSPL_VLC_MASTER_HEAD.vsp_code as [VSP Code],TSPL_VENDOR_MASTER.Vendor_Name as [VSP Name],TSPL_VLC_MASTER_HEAD.mcc as [MCC Code],TSPL_MCC_MASTER.mcc_name as [MCC Name],TSPL_VLC_MASTER_HEAD.created_by as [Created By],TSPL_VLC_MASTER_HEAD.created_date as [Created Date],TSPL_VLC_MASTER_HEAD.modified_by as [Modified By],TSPL_VLC_MASTER_HEAD.modified_date as [Modified Date] from TSPL_VLC_MASTER_HEAD left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.vsp_code and TSPL_VENDOR_MASTER.Form_Type='VSP' left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.mcc_code=TSPL_VLC_MASTER_HEAD.mcc left outer join tspl_village_master on TSPL_VLC_MASTER_HEAD.village_code=tspl_village_master.village_code left outer join tspl_mcc_route_master on TSPL_VLC_MASTER_HEAD.route_code=tspl_mcc_route_master.route_code"
            fndVLCode.Value = clsCommon.ShowSelectForm("VLCFND1", qry, "Code", whrcls, fndVLCode.Value, "Code", isButtonClicked)
            If clsCommon.myLen(fndVLCode.Value) > 0 Then
                txtVLCName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select vlc_name from TSPL_VLC_MASTER_HEAD where vlc_code='" + fndVLCode.Value + "'"))
                qry = " select TSPL_VLC_MASTER_HEAD.vlc_name,TSPL_VLC_MASTER_HEAD.VSP_Code,tspl_vendor_master.Vendor_Name as VSP_Name, TSPL_VLC_MASTER_HEAD.MCC, TSPL_MCC_MASTER.MCC_NAME, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,tspl_vendor_master.DISTRICT_Code,tspl_vendor_master.BLOCK_CODE,tspl_vendor_master.Zone_Code,tspl_vendor_master.REVENUE_VILLAGE_CODE,tspl_vendor_master.GRAMPANCHAYAT_CODE,
 tspl_vendor_master.PANCHAYAT_SAMITI_CODE,tspl_vendor_master.VIDHAN_SABHA_CODE from TSPL_VLC_MASTER_HEAD LEFT OUTER JOIN TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_CODE = TSPL_VLC_MASTER_HEAD.MCC LEFT OUTER JOIN tspl_vendor_master ON tspl_vendor_master.Vendor_Code = TSPL_VLC_MASTER_HEAD.VSP_Code  where vlc_code= '" + fndVLCode.Value + "' "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    txtMCC.Value = clsCommon.myCstr(dt.Rows(0)("MCC"))
                    lblMCCName.Text = clsCommon.myCstr(dt.Rows(0)("MCC_NAME"))
                    txtVlCUploader.Text = clsCommon.myCstr(dt.Rows(0)("VLC_Code_VLC_Uploader"))
                    txtVSPName.Text = clsCommon.myCstr(dt.Rows(0)("VSP_Name"))
                    txtDistrict.Value = clsCommon.myCstr(dt.Rows(0)("DISTRICT_Code"))
                    txtBlockCode.Value = clsCommon.myCstr(dt.Rows(0)("BLOCK_CODE"))
                    txtZone.Value = clsCommon.myCstr(dt.Rows(0)("ZONE_CODE"))
                    txtRevenueVillage.Value = clsCommon.myCstr(dt.Rows(0)("REVENUE_VILLAGE_CODE"))
                    txtGrampanchayat.Value = clsCommon.myCstr(dt.Rows(0)("GRAMPANCHAYAT_CODE"))
                    txtPanchayatSamiti.Value = clsCommon.myCstr(dt.Rows(0)("PANCHAYAT_SAMITI_CODE"))
                    txtVidhanSabha.Value = clsCommon.myCstr(dt.Rows(0)("VIDHAN_SABHA_CODE"))
                    FillAddressDetails()
                End If
            Else
                txtMCC.Value = ""
                lblMCCName.Text = ""
                txtVlCUploader.Text = ""
                txtVSPName.Text = ""
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndVillegeCode__MYOpenMasterForm(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles fndVillegeCode._MYOpenMasterForm
        Frm_Open = New FrmVillageMaster()
        Frm_Open.SetUserMgmt(clsUserMgtCode.frmVillageMaster)
        Frm_Open.Show()
    End Sub


    Private Sub fndVillegeCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles fndVillegeCode._MYValidating
        Try
            If clsCommon.myLen(fndCountryCode.Value) <= 0 Then
                fndCountryCode.Focus()
                Throw New Exception("Please Select Country Code First")
            End If
            If clsCommon.myLen(fndStateCode.Value) <= 0 Then
                fndStateCode.Focus()
                Throw New Exception("Please Select State Code First")
            End If
            If clsCommon.myLen(fndCityCode.Value) <= 0 Then
                fndCityCode.Focus()
                Throw New Exception("Please Select City Code First")
            End If
            Dim whrcls As String = " country_code='" & fndCountryCode.Value & "' and state_code='" & fndStateCode.Value & "' and city_code='" & fndCityCode.Value & "'"
            Dim str As String = ""
            Dim qry As String = " select TSPL_VILLAGE_MASTER.Village_Code as [Code] ,TSPL_VILLAGE_MASTER.Village_Name as [Village Name] ,TSPL_VILLAGE_MASTER.Add1 as [Address1] ,TSPL_VILLAGE_MASTER.Add2 as [Address2] ,TSPL_VILLAGE_MASTER.City_Code as [City Code] ,TSPL_VILLAGE_MASTER.State_Code as [State Code] ,TSPL_VILLAGE_MASTER.COUNTRY_CODE as [Country Code] ,TSPL_VILLAGE_MASTER.PIN_NO as [Pin No] ,TSPL_VILLAGE_MASTER.Created_By as [Created By] ,TSPL_VILLAGE_MASTER.Created_Date as [Created Date] ,TSPL_VILLAGE_MASTER.Modified_By as [Modified By] ,TSPL_VILLAGE_MASTER.Modified_Date as [Modified Date]  From TSPL_VILLAGE_MASTER "
            fndVillegeCode.Value = clsCommon.ShowSelectForm("VLGFND", qry, "Code", whrcls, fndVillegeCode.Value, "Code", isButtonClicked)

            txtVillegeName.Text = clsDBFuncationality.getSingleValue("select distinct village_name from tspl_village_master where village_code='" & fndVillegeCode.Value & "'")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtBankBranch__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtBankBranch._MYValidating
        'txtBankBranch.Value = clsBankBranchMaster.getFinder("", txtBankBranch.Value, isButtonClicked)
        'If clsCommon.myLen(txtBankBranch.Value) > 0 Then
        '    Dim obj As clsBankBranchMaster = clsBankBranchMaster.GetData(txtBankBranch.Value, NavigatorType.Current)
        '    txtBankBranchName.Text = obj.Branch_Name
        '    txtIFCICode.Text = obj.IFSCCode
        'End If
        If clsCommon.myLen(txtBankCode.Value) <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "Please seclect Bank code First", Me.Text)
            Return
        End If
        txtBankBranch.Value = clsVendorBankMaster.getFinderBranch("TSPL_Vendor_Bank_Branch_Details.Bank_Code= '" + txtBankCode.Value + "'", txtBankBranch.Value, isButtonClicked)
        If clsCommon.myLen(txtBankBranch.Value) > 0 Then
            txtIFCICode.Text = clsDBFuncationality.getSingleValue(" select Bank_IFSC_Code from TSPL_Vendor_Bank_Branch_Details where Branch_Name = '" + txtBankBranch.Value + "' and Bank_Code='" + txtBankCode.Value + "' ")
        End If

    End Sub

    Private Sub mnuMpDetailsExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMpDetailsExport.Click
        Try
            '----------------------BM00000002953(Monika)-------for blank export sheet and new fields------------
            'Dim strIncentiveColumn As String = Nothing
            Dim strIncentiveColumnForBlank As String = Nothing
            Dim str As String = "select count(*) from tspl_mp_master"
            Dim check As Integer = clsDBFuncationality.getSingleValue(str)
            For ii As Integer = 1 To SettNoOFIncentiveForImportExport
                strIncentiveColumnForBlank += ",'' as [Incentive" + clsCommon.myCstr(ii) + "]"
            Next
            'Ticket No : ERO/23/09/19-001035 By Prabhakar
            If check > 0 Then
                'Sanjay Ticket No-ERO/04/11/19-001085
                str = " select TSPL_mp_master.MP_Code as [MP Code] ,TSPL_mp_master.MP_Name as [MP Name] ,TSPL_mp_master.vlc_Code as [VLC Code] , TSPL_VLC_MASTER_HEAD.VLC_NAME as [VLC Name] ,TSPL_mp_master.Father_Name as [Father Name] ,TSPL_mp_master.Add1 as [Address1] ,TSPL_mp_master.Add2 as [Address2] ,TSPL_mp_master.Zila as [Zila] ,TSPL_mp_master.Tehsil as [Tehsil] ,TSPL_mp_master.City_code as [City Code] ,TSPL_mp_master.State_Code as [State Code] ,TSPL_mp_master.Country_code as [Country Code] ,TSPL_mp_master.Pin_code as [Pin Code] ,TSPL_mp_master.Telphone as [Telphone] ,TSPL_mp_master.Email as [Email] ,TSPL_mp_master.Fax as [AadharNo],TSPL_MP_MASTER.Jan_Aadhar_No as [JanAadharNo],convert(varchar,TSPL_mp_master.DOB,103) as [Date of Birth] ,TSPL_mp_master.Education as [Education] ,TSPL_mp_master.Land_Holding as [Land Holding] ,TSPL_mp_master.No_Of_breedable_milk_animal as [No Of Milch Animal] ,TSPL_mp_master.Milk_production as [Total Milk Production] ,TSPL_mp_master.Milk_Home_consumption as [Milk For Self Consumption] ,TSPL_mp_master.Milk_For_sale as [Milk For Sale] ,TSPL_mp_master.PayeeName as [Payee Name] ,TSPL_mp_master.BankName as [Bank Code] ,TSPL_mp_master.BankBranch as [Bank Branch] ,TSPL_mp_master.BankCityCode as [Bank City Code] ,TSPL_mp_master.BankStateCode as [Bank State Code] ,TSPL_mp_master.IFCICode as [IFSC Code] ,convert(varchar,TSPL_mp_master.AccountNO) as [Account No],TSPL_mp_master.MP_Code_VLC_Uploader as [MP Uploader Code],Cust_Account as [Customer Acc Set],Acct_Set_Code as [Vendor Acc Set], TOLERANCE"
                If UseVLCUploaderCodeMPUploaderCodeInMCCProcurement = True Then
                    str += ",TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [VLC Uploader Code] "
                End If
                For j As Integer = 1 To SettNoOFIncentiveForImportExport
                    str += " ,(select INCENTIVE_CODE from (Select ROW_NUMBER () over (order by MP_CODE,INCENTIVE_CODE ) As SNo,MP_CODE,INCENTIVE_CODE  From TSPL_MP_INCENTIVE where MP_CODE=TSPL_mp_master.MP_Code)xxx where xxx.SNo=" & j & ") as [Incentive" & j & "]"
                Next
                str += ",Incentive_Account as  [Incentive Account], TSPL_mp_master.CAST_CATEGORY_CODE,TSPL_mp_master.DISTRICT_Code ,TSPL_mp_master.BLOCK_CODE,TSPL_mp_master.Zone_Code,TSPL_mp_master.REVENUE_VILLAGE_CODE,TSPL_mp_master.GRAMPANCHAYAT_CODE,TSPL_mp_master.PANCHAYAT_SAMITI_CODE,TSPL_mp_master.VIDHAN_SABHA_CODE from TSPL_mp_master left join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_mp_master.VLC_Code "
            Else
                str = "select '' as [MP Code] ,'' as [MP Name] ,'' as [VLC Code] ,'' as [VLC Name],'' as [Father Name] ,'' as [Address1] ,'' as [Address2] ,'' as [Zila] ,'' as [Tehsil] ,'' as [City Code] ,'' as [State Code] ,'' as [Country Code] ,'' as [Pin Code] ,'' as [Telphone] ,'' as [Email] ,'' as [AadharNo],'' as JanAadharNo ,'' as [Date of Birth] ,'' as [Education] ,'' as [Land Holding] ,'' as [No Of Milch Animal] ,'' as [Total Milk Production] ,'' as [Milk For Self Consumption] ,'' as [Milk For Sale] ,'' as [Payee Name] ,'' as [Bank Code] ,'' as [Bank Branch] ,'' as [Bank City Code] ,'' as [Bank State Code] ,'' as [IFSC Code] ,'' as [Account No],'' as [MP Uploader Code],'' as [Customer Acc Set],'' as [Vendor Acc Set] "
                If UseVLCUploaderCodeMPUploaderCodeInMCCProcurement = True Then
                    str += ",'' as [VLC Uploader Code] "
                End If
                str += strIncentiveColumnForBlank + ",'' as  [Incentive Account], '' as CAST_CATEGORY_CODE,'' as DISTRICT_Code,'' as  BLOCK_CODE,'' as Zone_Code,'' as REVENUE_VILLAGE_CODE,'' as GRAMPANCHAYAT_CODE,'' as PANCHAYAT_SAMITI_CODE,'' as VIDHAN_SABHA_CODE"
            End If
            ListImpExpColumnsMandatory = New List(Of String)({"MP Code", "VLC Uploader Code", "VLC Code", "MP Name", "Address1", "JanAadharNo", "MP Uploader Code"})
            ListImpExpColumnsSuperMandatory = New List(Of String)({"MP Code"})
            transportSql.ExporttoExcel(str, "", "", Me, ListImpExpColumnsMandatory, ListImpExpColumnsSuperMandatory, MyBase.Form_ID)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub





    Private Sub mnuMpDetailsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMpDetailsImport.Click
        '=====update by Preeti Gupta[ERO/19/06/18-000354]
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim isSaved As Boolean = True
        Dim i As Integer = 0
        Dim trans As SqlTransaction = Nothing
        'connectSql.OpenConnection()
        Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
        Dim Strs As List(Of String) = New List(Of String)
        Strs.Add("MP Code")
        Strs.Add("MP Name")
        Strs.Add("VLC Code")
        Strs.Add("VLC Name")
        Strs.Add("Father Name")
        Strs.Add("Address1")
        Strs.Add("Address2")
        Strs.Add("Zila")
        Strs.Add("Tehsil")
        Strs.Add("City Code")
        Strs.Add("State Code")
        Strs.Add("Country Code")
        Strs.Add("Pin Code")
        Strs.Add("Telphone")
        Strs.Add("Email")
        Strs.Add("AadharNo")
        Strs.Add("JanAadharNo")
        Strs.Add("Date of Birth")
        Strs.Add("Education")
        Strs.Add("Land Holding")
        Strs.Add("No Of Milch Animal")
        Strs.Add("Total Milk Production")
        Strs.Add("Milk For Self Consumption")
        Strs.Add("Milk For Sale")
        Strs.Add("Payee Name")
        Strs.Add("Bank Code")
        Strs.Add("Bank Branch")
        Strs.Add("Bank City Code")
        Strs.Add("Bank State Code")
        Strs.Add("IFSC Code")
        Strs.Add("Account No")
        Strs.Add("MP Uploader Code")
        Strs.Add("Customer Acc Set")
        Strs.Add("Vendor Acc Set")
        Strs.Add("TOLERANCE")
        If UseVLCUploaderCodeMPUploaderCodeInMCCProcurement = True Then
            Strs.Add("VLC Uploader Code")
        End If

        For ii As Integer = 1 To SettNoOFIncentiveForImportExport
            Strs.Add("Incentive" + clsCommon.myCstr(ii))
        Next
        Strs.Add("Incentive Account")
        Strs.Add("CAST_CATEGORY_CODE")
        Strs.Add("DISTRICT_Code")
        Strs.Add("BLOCK_CODE")
        Strs.Add("Zone_Code")
        Strs.Add("REVENUE_VILLAGE_CODE")
        Strs.Add("GRAMPANCHAYAT_CODE")
        Strs.Add("PANCHAYAT_SAMITI_CODE")
        Strs.Add("VIDHAN_SABHA_CODE")
        If transportSql.importExcel(gv, Strs.ToArray()) Then ', "Village Code"
            trans = clsDBFuncationality.GetTransactin()
            Try
                clsCommon.ProgressBarShow()
                For Each grow As GridViewRowInfo In gv.Rows
                    Dim obj As New clsMpMaster()
                    i = i + 1
                    Dim strData As String = ""
                    If UseVLCUploaderCodeMPUploaderCodeInMCCProcurement = True Then
                        If clsCommon.myLen(clsCommon.myCstr(grow.Cells("VLC Uploader Code").Value)) <= 0 Then
                            Throw New Exception("VLC Uploader Code Can Not Be Left Blank")
                        End If
                        strData = clsDBFuncationality.getSingleValue("select VLC_Code from TSPL_VLC_MASTER_HEAD where VLC_Code_VLC_Uploader='" & clsCommon.myCstr(grow.Cells("VLC Uploader Code").Value) & "'", trans)
                        If clsCommon.myLen(strData) <= 0 Then
                            Throw New Exception("VLC Uploader Code Not Found")
                        End If
                    Else
                        strData = clsCommon.myCstr(grow.Cells("VLC Code").Value)
                        If clsCommon.myLen(strData) <= 0 Then
                            Throw New Exception("VLC Code Can Not Be Left Blank")
                        End If
                        If clsCommon.myLen(strData) > 30 Then
                            Throw New Exception("VLC Code Can Not Be Larger Then 30 Charachter")
                        End If
                    End If

                    obj.MCC_Code = strData
                    strData = clsCommon.myCstr(grow.Cells("MP Code").Value)

                    obj.MP_Code = strData

                    If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & obj.MP_Code & "'", trans) > 0 Then
                        If MyBase.isUpdateFlag = False Then
                            Throw New Exception("Don't have permission to update MP Master.")
                        End If
                    End If

                    strData = clsCommon.myCstr(grow.Cells("MP Name").Value)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("MP Name Can Not Be Left Blank")
                    End If
                    Dim pattern As String = "[^a-zA-Z\s]"
                    Dim regex As New Regex(pattern)

                    If regex.IsMatch(strData) Then
                        Throw New Exception("Special Characters And Numbers Are Not Allowed In MP Name")
                    End If
                    If clsCommon.myLen(strData) > 50 Then
                        Throw New Exception("MP Name Can Not Be Larger Then 50 Charachter")
                    End If
                    obj.MP_Name = strData

                    If clsCommon.myLen(grow.Cells("Father Name").Value) > 0 Then
                        obj.Father_Name = grow.Cells("Father Name").Value
                    End If

                    strData = clsCommon.myCstr(grow.Cells("Address1").Value)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("Address1 Can Not Be Left Blank")
                    End If
                    If clsCommon.myLen(strData) > 50 Then
                        Throw New Exception("Address1 Can Not Be Larger Then 50 Charachter")
                    End If
                    obj.Add1 = strData
                    If clsCommon.myLen(grow.Cells("Address2").Value) > 0 Then
                        obj.Add2 = clsCommon.myCstr(grow.Cells("Address2").Value)
                    End If
                    If clsCommon.myLen(grow.Cells("Zila").Value) > 0 Then
                        obj.Zila = clsCommon.myCstr(grow.Cells("Zila").Value)
                    End If
                    If clsCommon.myLen(grow.Cells("Tehsil").Value) > 0 Then
                        obj.Tehsil = grow.Cells("Tehsil").Value
                    End If
                    strData = clsCommon.myCstr(grow.Cells("Pin Code").Value)
                    obj.Pin_code = strData
                    strData = clsCommon.myCstr(grow.Cells("City Code").Value)

                    obj.City_code = strData

                    strData = clsCommon.myCstr(grow.Cells("State Code").Value)

                    obj.State_Code = strData
                    strData = clsCommon.myCstr(grow.Cells("Country Code").Value)

                    obj.Country_code = strData
                    If clsCommon.myLen(grow.Cells("Telphone").Value) > 0 Then
                        obj.Telphone = grow.Cells("Telphone").Value
                    End If
                    If clsCommon.myLen(grow.Cells("Email").Value) > 0 Then
                        Dim check As Match = Regex.Match(grow.Cells("Email").Value, "\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*")
                        If check.Success Then
                            obj.Email = clsCommon.myCstr(grow.Cells("Email").Value)
                        Else
                            Throw New Exception("Email Is In Invalid Format. It Shoud Be as UserName@Domain Format ")
                        End If
                    End If
                    If clsCommon.myLen(grow.Cells("AadharNo").Value) > 0 Then
                        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*) from tspl_mp_master where fax='" & clsCommon.myCstr(grow.Cells("AadharNo").Value) & "' and MP_Code<>'" & clsCommon.myCstr(obj.MP_Code) & "'", trans)) > 0 Then
                            Throw New Exception("Same Aadhar No is exist with another MP so please change Aadhar No because Aadhar No is unique.")
                        End If
                        obj.Fax = clsCommon.myCstr(grow.Cells("AadharNo").Value)
                    End If

                    If clsCommon.myLen(grow.Cells("JanAadharNo").Value) > 0 Then
                        If clsCommon.myLen(grow.Cells("JanAadharNo").Value) <> 10 Then
                            Throw New Exception("Invalid Jan Aadhar No.Please Enter 10 Digit Jan Aadhar No")
                        End If
                        obj.Jan_Aadhar_No = clsCommon.myCstr(grow.Cells("JanAadharNo").Value)
                    Else
                        If SettJanAadharNoMandatory Then
                            Throw New Exception("Please Fill Jan Aadhar No")
                        End If
                    End If

                    If clsCommon.myLen(grow.Cells("Date Of Birth").Value) > 0 Then
                        If IsDate(grow.Cells("Date Of Birth").Value) Then
                            obj.DOB = clsCommon.myCDate(grow.Cells("Date Of Birth").Value, "dd/MMM/yyyy")
                        Else
                            Throw New Exception("Invalid Date Found For Date Of Birth")
                        End If
                    End If
                    'MP Uploader Code
                    If clsCommon.myLen(grow.Cells("MP Uploader Code").Value) > 0 Then
                        obj.MP_CODE_VLC_UPLOADER = grow.Cells("MP Uploader Code").Value
                    Else
                        Throw New Exception("Please Fill MP Uploader Code")
                    End If

                    Dim qqq As String = "select COUNT(*) from TSPL_MP_MASTER where MP_Code_vlc_uploader='" & obj.MP_CODE_VLC_UPLOADER & "' and mp_code<>'" & obj.MP_Code & "' and vlc_Code='" & obj.MCC_Code & "'"
                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qqq, trans)) >= 1 Then
                        Throw New Exception("Duplicate MP uploader Code")
                    End If

                    If clsCommon.myLen(grow.Cells("Education").Value) > 0 Then
                        obj.Education = grow.Cells("Education").Value
                    End If
                    If clsCommon.myLen(grow.Cells("TOLERANCE").Value) > 0 Then
                        If IsNumeric(grow.Cells("TOLERANCE").Value) = False Then
                            Throw New Exception("TOLERANCE value should be Numeric")
                        End If
                        If clsCommon.myCdbl(grow.Cells("TOLERANCE").Value) > 100 Then
                            Throw New Exception("TOLERANCE value should be less then 100.")
                        End If
                        obj.TOLERANCE = clsCommon.myCdbl(grow.Cells("TOLERANCE").Value)
                    End If
                    obj.Land_Holding = clsCommon.myCdbl(grow.Cells("Land Holding").Value)

                    obj.No_Of_Animal = 0
                    obj.No_Of_breedable_milk_animal = clsCommon.myCdbl(grow.Cells("No Of Milch Animal").Value)
                    obj.Milk_production = clsCommon.myCdbl(grow.Cells("Total Milk Production").Value)
                    obj.Milk_Home_consumption = clsCommon.myCdbl(grow.Cells("Milk For Self Consumption").Value)
                    obj.Milk_For_sale = clsCommon.myCdbl(grow.Cells("Milk For Sale").Value)
                    obj.Milk_For_sale = clsCommon.myCdbl(grow.Cells("Milk For Sale").Value)
                    If regex.IsMatch(clsCommon.myCstr(grow.Cells("Payee Name").Value)) Then
                        Throw New Exception("Special Characters And Numbers Are Not Allowed In Payee Name")
                    End If
                    obj.PayeeName = clsCommon.myCstr(grow.Cells("Payee Name").Value)

                    '' Panch Raj: adding cust acc set ad vendor acc set 
                    obj.Cust_Account = clsCommon.myCstr(grow.Cells("Customer Acc Set").Value)
                    If clsCommon.myLen(obj.Cust_Account) > 0 Then
                        '' check valid acc set
                        qqq = "select Cust_Account from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" & obj.Cust_Account & "' "
                        Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                        If clsCommon.myLen(checkCode) <= 0 Then
                            Throw New Exception("Invalid Customer Account Set- " & obj.Cust_Account & " at line no: " & (grow.Index + 1) & "")
                        End If
                    End If
                    obj.Acct_Set_Code = clsCommon.myCstr(grow.Cells("Vendor Acc Set").Value)
                    If clsCommon.myLen(obj.Acct_Set_Code) > 0 Then
                        '' check valid acc set
                        qqq = "select Acct_Set_Code from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code='" & obj.Acct_Set_Code & "' "
                        Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                        If clsCommon.myLen(checkCode) <= 0 Then
                            Throw New Exception("Invalid Vendor Account Set- " & obj.Acct_Set_Code & " at line no: " & (grow.Index + 1) & "")
                        End If
                    End If
                    obj.Incentive_Account = clsCommon.myCstr(grow.Cells("Incentive Account").Value)
                    If clsCommon.myLen(obj.Incentive_Account) > 0 Then
                        qqq = "select Account_Code  from tspl_gl_accounts where  ControlAccount ='Y' and Account_Code='" + obj.Incentive_Account + "' "
                        Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                        If clsCommon.myLen(checkCode) <= 0 Then
                            Throw New Exception("Invalid Incentive GL Account - " & clsCommon.myCstr(grow.Cells("Incentive Account").Value) & " at line no: " & (grow.Index + 1) & "")
                        End If
                    Else
                        If IncentiveAccNoMandatoryInMPMaster = True Then
                            Throw New Exception("Please Fill Incentive GL Account at line no: " & (grow.Index + 1) & "")
                        End If
                    End If
                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("CAST_CATEGORY_CODE").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (CAST_CATEGORY_CODE) as CAST_CATEGORY_CODE  from TSPL_CAST_CATEGORY_MASTER where CAST_CATEGORY_CODE = '" + clsCommon.myCstr(grow.Cells("CAST_CATEGORY_CODE").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid CAST_CATEGORY_CODE  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.CAST_CATEGORY_CODE = clsCommon.myCstr(grow.Cells("CAST_CATEGORY_CODE").Value)
                        End If
                    End If

                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("DISTRICT_Code").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (code) as code  from TSPL_DISTRICT_MASTER where code = '" + clsCommon.myCstr(grow.Cells("DISTRICT_Code").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid DISTRICT_Code  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.DISTRICT_Code = clsCommon.myCstr(grow.Cells("DISTRICT_Code").Value)
                        End If
                    End If

                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("BLOCK_CODE").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (BLOCK_CODE) as BLOCK_CODE  from TSPL_BLOCK_MASTER where BLOCK_CODE = '" + clsCommon.myCstr(grow.Cells("BLOCK_CODE").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid BLOCK_CODE  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.BLOCK_CODE = clsCommon.myCstr(grow.Cells("BLOCK_CODE").Value)
                        End If
                    End If
                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("Zone_Code").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (Zone_Code) as Zone_Code  from TSPL_ZONE_MASTER where Zone_Code = '" + clsCommon.myCstr(grow.Cells("Zone_Code").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid Zone_Code  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.Zone_Code = clsCommon.myCstr(grow.Cells("Zone_Code").Value)
                        End If
                    End If
                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("REVENUE_VILLAGE_CODE").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (REVENUE_VILLAGE_CODE) as REVENUE_VILLAGE_CODE  from TSPL_REVENUE_VILLAGE_MASTER where REVENUE_VILLAGE_CODE = '" + clsCommon.myCstr(grow.Cells("REVENUE_VILLAGE_CODE").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid REVENUE_VILLAGE_CODE  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.REVENUE_VILLAGE_CODE = clsCommon.myCstr(grow.Cells("REVENUE_VILLAGE_CODE").Value)
                        End If
                    End If

                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("GRAMPANCHAYAT_CODE").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (GRAMPANCHAYAT_CODE) as GRAMPANCHAYAT_CODE  from TSPL_GRAMPANCHAYAT_MASTER where GRAMPANCHAYAT_CODE = '" + clsCommon.myCstr(grow.Cells("GRAMPANCHAYAT_CODE").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid GRAMPANCHAYAT_CODE  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.GRAMPANCHAYAT_CODE = clsCommon.myCstr(grow.Cells("GRAMPANCHAYAT_CODE").Value)
                        End If
                    End If

                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("PANCHAYAT_SAMITI_CODE").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (PANCHAYAT_SAMITI_CODE) as GRAMPANCHAYAT_CODE  from TSPL_PANCHAYAT_SAMITI_MASTER where PANCHAYAT_SAMITI_CODE = '" + clsCommon.myCstr(grow.Cells("PANCHAYAT_SAMITI_CODE").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid PANCHAYAT_SAMITI_CODE  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.PANCHAYAT_SAMITI_CODE = clsCommon.myCstr(grow.Cells("PANCHAYAT_SAMITI_CODE").Value)
                        End If
                    End If

                    If clsCommon.myLen(clsCommon.myCstr(grow.Cells("VIDHAN_SABHA_CODE").Value)) > 0 Then
                        Dim QryCastCatg As String = " select count (VIDHAN_SABHA_CODE) as VIDHAN_SABHA_CODE  from TSPL_VIDHAN_SABHA_MASTER where VIDHAN_SABHA_CODE = '" + clsCommon.myCstr(grow.Cells("VIDHAN_SABHA_CODE").Value) + "' "
                        Dim isCheckCastCategory As Boolean = clsCommon.myCstr(clsDBFuncationality.getSingleValue(QryCastCatg, trans))
                        If isCheckCastCategory = False Then
                            Throw New Exception("Invalid VIDHAN_SABHA_CODE  at line no: " & (grow.Index + 1) & "")
                        Else
                            obj.VIDHAN_SABHA_CODE = clsCommon.myCstr(grow.Cells("VIDHAN_SABHA_CODE").Value)
                        End If
                    End If

                    If SettBankIFSCCodeValidateByService Then
                        If clsCommon.myLen(clsCommon.myCstr(grow.Cells("IFSC Code").Value)) > 0 Then
                            Dim arrFilter As New Dictionary(Of String, String)
                            arrFilter.Add("IFSC", clsCommon.myCstr(grow.Cells("IFSC Code").Value))
                            arrFilter.Add("OutPutType", "1")
                            Dim dt As DataTable = Xtra.GetDataFromAPI(EnumAPI.BankIFSC, "GetIFSC", arrFilter)
                            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                                Throw New Exception("Invalid IFSC Code")
                            End If
                            If clsCommon.CompairString(dt.Rows(0)("Result"), "true") = CompairStringResult.Equal Then
                                obj.BankName = clsCommon.myCstr(dt.Rows(0)("BANK"))
                                obj.BankBranch = clsCommon.myCstr(dt.Rows(0)("BRANCH"))
                                obj.BankStateCode = clsCommon.myCstr(dt.Rows(0)("STATE"))
                                obj.BankCityCode = clsCommon.myCstr(dt.Rows(0)("CITY"))
                                obj.IFCICode = clsCommon.myCstr(dt.Rows(0)("IFSC")).ToUpper()
                                If clsCommon.myCstr(grow.Cells("Account No").Value).Contains("E+") OrElse clsCommon.myCstr(grow.Cells("Account No").Value).Contains("E-") Then
                                    obj.AccountNO = Decimal.Parse(clsCommon.myCstr(grow.Cells("Account No").Value), System.Globalization.NumberStyles.Float)
                                Else
                                    obj.AccountNO = clsCommon.myCstr(grow.Cells("Account No").Value)
                                End If
                            Else
                                Throw New Exception(dt.Rows(0)("Response"))
                            End If
                        End If
                    ElseIf clsCommon.myLen(grow.Cells("Bank Code").Value) > 0 AndAlso EnableBankFromMaster = True Then
                        If clsDBFuncationality.getSingleValue("select count(*) from TSPL_Vendor_Bank_Master where bank_code='" & grow.Cells("Bank Code").Value & "'", trans) > 0 Then
                            Dim bnkDt As DataTable = clsDBFuncationality.GetDataTable("select * from TSPL_Vendor_Bank_Master where bank_code='" & grow.Cells("Bank Code").Value & "'", trans)
                            obj.BankName = clsCommon.myCstr(bnkDt.Rows(0)("BANK_CODE"))
                            obj.BankCityCode = clsCommon.myCstr(bnkDt.Rows(0)("City_Code"))
                            obj.BankStateCode = clsCommon.myCstr(bnkDt.Rows(0)("State_Code"))
                            If clsCommon.myCstr(grow.Cells("Account No").Value).Contains("E+") OrElse clsCommon.myCstr(grow.Cells("Account No").Value).Contains("E-") Then
                                obj.AccountNO = Decimal.Parse(clsCommon.myCstr(grow.Cells("Account No").Value), System.Globalization.NumberStyles.Float)
                            Else
                                obj.AccountNO = clsCommon.myCstr(grow.Cells("Account No").Value)
                            End If

                            obj.BankBranch = clsCommon.myCstr(grow.Cells("Bank Branch").Value)
                            obj.IFCICode = clsCommon.myCstr(grow.Cells("IFSC Code").Value).ToUpper()
                            If clsCommon.myLen(obj.IFCICode) > 0 Then
                                qqq = "select Bank_IFSC_Code from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & grow.Cells("Bank Code").Value & "' and Bank_IFSC_Code = '" & grow.Cells("IFSC Code").Value & "' "
                                Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                If clsCommon.myLen(checkCode) <= 0 Then
                                    Throw New Exception("Invalid IFSC Code for Bank Code - " & grow.Cells("Bank Code").Value & " ")
                                End If

                                If clsCommon.myLen(obj.BankBranch) <= 0 Then
                                    qqq = "select Branch_Name from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & grow.Cells("Bank Code").Value & "' and Bank_IFSC_Code = '" & grow.Cells("IFSC Code").Value & "' "
                                    Dim BranchCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                    If clsCommon.myLen(BranchCode) > 0 Then
                                        obj.BankBranch = BranchCode
                                        grow.Cells("Bank Branch").Value = BranchCode
                                    End If
                                End If

                            End If
                            If clsCommon.myLen(obj.BankBranch) > 0 Then
                                qqq = "select Branch_Name from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & grow.Cells("Bank Code").Value & "' and Branch_Name = '" & grow.Cells("Bank Branch").Value & "' "
                                Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                If clsCommon.myLen(checkCode) <= 0 Then
                                    Throw New Exception("Invalid Bank Branch for Bank Code - " & grow.Cells("Bank Code").Value & " ")
                                End If
                                If clsCommon.myLen(obj.IFCICode) <= 0 Then
                                    qqq = "select Bank_IFSC_Code from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & grow.Cells("Bank Code").Value & "' and Branch_Name = '" & grow.Cells("Bank Branch").Value & "' "
                                    Dim IFSCCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                    If clsCommon.myLen(IFSCCode) > 0 Then
                                        obj.IFCICode = IFSCCode
                                        grow.Cells("IFSC Code").Value = IFSCCode
                                    End If
                                End If

                            End If

                            If clsCommon.myLen(obj.BankBranch) > 0 AndAlso clsCommon.myLen(obj.IFCICode) > 0 Then
                                qqq = "select Branch_Name from TSPL_Vendor_Bank_Branch_Details where Bank_code ='" & grow.Cells("Bank Code").Value & "' and Branch_Name = '" & grow.Cells("Bank Branch").Value & "' and Bank_IFSC_Code = '" & grow.Cells("IFSC Code").Value & "' "
                                Dim checkCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qqq, trans))
                                If clsCommon.myLen(checkCode) <= 0 Then
                                    Throw New Exception("Invalid IFSC Code for Branch - [" & grow.Cells("Bank Branch").Value & "] ,Bank Code - " & grow.Cells("Bank Code").Value & "")
                                End If
                            End If
                        End If
                    Else
                        If EnableBankFromMaster = False Then
                            obj.BankName = clsCommon.myCstr(grow.Cells("Bank Code").Value)
                            obj.BankCityCode = clsCommon.myCstr(grow.Cells("Bank City Code").Value)
                            obj.BankStateCode = clsCommon.myCstr(grow.Cells("Bank State Code").Value)
                            If clsCommon.myCstr(grow.Cells("Account No").Value).Contains("E+") OrElse clsCommon.myCstr(grow.Cells("Account No").Value).Contains("E-") Then
                                obj.AccountNO = Decimal.Parse(clsCommon.myCstr(grow.Cells("Account No").Value), System.Globalization.NumberStyles.Float)
                            Else
                                obj.AccountNO = clsCommon.myCstr(grow.Cells("Account No").Value)
                            End If
                            obj.BankBranch = clsCommon.myCstr(grow.Cells("Bank Branch").Value)
                            obj.IFCICode = clsCommon.myCstr(grow.Cells("IFSC Code").Value)
                        End If
                    End If
                    If clsMpMaster.IsDBTDone(obj.MP_Code, trans) Then
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable("select AccountNO,IFCICode From TSPL_MP_MASTER where MP_Code='" + obj.MP_Code + "'", trans)
                        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                            obj.AccountNO = clsCommon.myCstr(dt.Rows(0)("AccountNO"))
                            obj.IFCICode = clsCommon.myCstr(dt.Rows(0)("IFCICode"))
                        End If
                    End If

                    If clsCommon.myLen(obj.AccountNO) > 0 Then
                        If clsCommon.MySpecialChars(obj.AccountNO, EnumSpecialChactersType.Digit) Then
                            Throw New Exception("Invalid AccountNO [" + clsCommon.myCstr(grow.Cells("Account No").Value) + "]")
                        End If
                    End If

                    obj.ArrMPIncentiveMapping = New ArrayList()
                    For ii As Integer = 1 To SettNoOFIncentiveForImportExport
                        Dim strIncentiveCode As String = clsCommon.myCstr(grow.Cells("Incentive" + clsCommon.myCstr(ii)).Value)
                        If clsCommon.myLen(strIncentiveCode) > 0 Then
                            strIncentiveCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Incentive_Code from TSPL_INCENTIVE_MASTER_HEAD  where Incentive_Code='" + strIncentiveCode + "'", trans))
                            If clsCommon.myLen(strIncentiveCode) <= 0 Then
                                Throw New Exception("Invalid Incentive Code [" + clsCommon.myCstr(grow.Cells("Incentive" + clsCommon.myCstr(ii)).Value) + "]")
                            End If
                            obj.ArrMPIncentiveMapping.Add(strIncentiveCode)
                        End If
                    Next


                    '=====================================================================================

                    If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & obj.MP_Code & "'", trans) = 0 Then
                        Dim dt As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
                        If clsCommon.myLen(obj.MP_Code) <= 0 Then
                            obj.MP_Code = clsERPFuncationality.GetNextCode(trans, dt, clsDocType.MPMaster, "", "")
                        End If

                        If obj.No_Of_breedable_milk_animal > 0 Then
                            Dim objAnimal As clsAnimalDetails
                            obj.arrAnimalDetail = New List(Of clsAnimalDetails)
                            For j As Integer = 0 To obj.No_Of_breedable_milk_animal
                                objAnimal = New clsAnimalDetails
                                objAnimal.Prog_Code = Me.Form_ID
                                objAnimal.Trans_Code = obj.MP_Code
                                objAnimal.Line_No = (j + 1)
                                objAnimal.Bread_of_Animal = "N/A"
                                objAnimal.Type_Of_Animal = "N/A"
                                obj.arrAnimalDetail.Add(objAnimal)
                            Next
                        End If

                        obj.isNewEntry = True
                        obj.Modified_By = objCommonVar.CurrentUserCode
                        obj.Modified_Date = clsCommon.GetPrintDate(strdate, "dd/MMM/yyyy")
                        obj.Comp_Code = objCommonVar.CurrentCompanyCode
                        obj.Created_By = objCommonVar.CurrentUserCode
                        obj.Created_Date = clsCommon.GetPrintDate(strdate, "dd/MMM/yyyy")
                        clsMpMaster.SaveData(obj, trans)
                    Else
                        obj.Modified_By = objCommonVar.CurrentUserCode
                        obj.Modified_Date = clsCommon.GetPrintDate(strdate, "dd/MMM/yyyy")
                        obj.isNewEntry = False

                        isSaved = isSaved AndAlso clsMpMaster.SaveData(obj, trans)
                        If Not isSaved Then
                            Exit For
                        End If
                    End If
                Next
                If isSaved Then
                    trans.Commit()
                    clsCommon.ProgressBarHide()
                    common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!, " + Environment.NewLine + " Only Data Regarding Animal Detail is Not Updated Please Import Respective Sheets ", Me.Text, MessageBoxButtons.OK)
                Else
                    trans.Rollback()
                    clsCommon.ProgressBarHide()
                    clsCommon.MyMessageBoxShow(Me, "Data Not Saved. Try again ", Me.Text)
                End If

            Catch ex As Exception
                trans.Rollback()
                clsCommon.ProgressBarHide()
                clsCommon.MyMessageBoxShow(Me, ex.Message & " At Line No : " & i, Me.Text)
            End Try

        End If
        Me.Controls.Remove(gv)
    End Sub

    'Private Sub mnuBuffaloesDetailsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuBuffaloesDetailsImport.Click

    '    Dim gv As New UserControls.MyRadGridView
    '    Me.Controls.Add(gv)
    '    Dim i As Integer = 0
    '    Dim trans As SqlTransaction
    '    connectSql.OpenConnection()
    '    trans = clsDBFuncationality.GetTransactin()
    '    Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
    '    If transportSql.importExcel(gv, "MP Code", "Line No", "Bread Of Buffalo") Then
    '        Try
    '            clsCommon.ProgressBarShow()
    '            For Each grow As GridViewRowInfo In gv.Rows
    '                Dim obj As New clsBuffaloesDetails()
    '                i = i + 1
    '                Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)
    '                If clsCommon.myLen(strData) <= 0 Then
    '                    Throw New Exception("MP Code Can Not Be Left Blank")
    '                End If
    '                If clsCommon.myLen(strData) > 30 Then
    '                    Throw New Exception("MP Code Can Not Be Larger Then 30 Charachter")
    '                End If
    '                If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & strData & "'", trans) = 0 Then
    '                    Throw New Exception("Invalid MP Code. Code Not Found In Master")
    '                End If
    '                obj.Trans_Code = strData

    '                strData = clsCommon.myCstr(grow.Cells("Line No").Value)
    '                If clsCommon.myLen(strData) <= 0 Then
    '                    Throw New Exception("Line No Can Not Be Left Blank")
    '                End If
    '                obj.Line_No = strData
    '                strData = clsCommon.myCstr(grow.Cells("Bread Of Buffalo").Value)
    '                If clsCommon.myLen(strData) <= 0 Then
    '                    Throw New Exception("Bread Of Buffalo  Can Not Be Left Blank")
    '                End If
    '                obj.Bread_of_Buffalo = strData
    '                obj.Prog_Code = Me.Form_ID
    '                obj.SaveData(obj, trans)
    '            Next

    '            trans.Commit()
    '            clsCommon.ProgressBarHide()
    '            common.clsCommon.MyMessageBoxShow("Data Transfer Completed!, ", Me.Text, MessageBoxButtons.OK)
    '        Catch ex As Exception
    '            trans.Rollback()
    '            clsCommon.ProgressBarHide()
    '            clsCommon.MyMessageBoxShow(ex.Message & " At Line No : " & i)
    '        End Try

    '    End If
    '    Me.Controls.Remove(gv)
    'End Sub



    'Private Sub mnuCowDetailsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCowDetailsImport.Click
    '    Dim gv As New UserControls.MyRadGridView
    '    Me.Controls.Add(gv)
    '    Dim i As Integer = 0
    '    Dim trans As SqlTransaction
    '    connectSql.OpenConnection()
    '    trans = clsDBFuncationality.GetTransactin()
    '    Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
    '    If transportSql.importExcel(gv, "MP Code", "Line No", "Bread Of Cow") Then
    '        Try
    '            clsCommon.ProgressBarShow()
    '            For Each grow As GridViewRowInfo In gv.Rows
    '                Dim obj As New clsCowDetails()
    '                i = i + 1
    '                Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)
    '                If clsCommon.myLen(strData) <= 0 Then
    '                    Throw New Exception("MP Code Can Not Be Left Blank")
    '                End If
    '                If clsCommon.myLen(strData) > 30 Then
    '                    Throw New Exception("MP Code Can Not Be Larger Then 30 Charachter")
    '                End If
    '                If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & strData & "'", trans) = 0 Then
    '                    Throw New Exception("Invalid MP Code. Code Not Found In Master")
    '                End If
    '                obj.Trans_Code = strData

    '                strData = clsCommon.myCstr(grow.Cells("Line No").Value)
    '                If clsCommon.myLen(strData) <= 0 Then
    '                    Throw New Exception("Line No Can Not Be Left Blank")
    '                End If
    '                obj.Line_No = strData
    '                strData = clsCommon.myCstr(grow.Cells("Bread Of Cow").Value)
    '                If clsCommon.myLen(strData) <= 0 Then
    '                    Throw New Exception("Bread Of Cow Can Not Be Left Blank")
    '                End If
    '                obj.Bread_of_cow = strData
    '                obj.Prog_Code = Me.Form_ID
    '                obj.SaveData(obj, trans)
    '            Next

    '            trans.Commit()
    '            clsCommon.ProgressBarHide()
    '            common.clsCommon.MyMessageBoxShow("Data Transfer Completed!, ", Me.Text, MessageBoxButtons.OK)
    '        Catch ex As Exception
    '            trans.Rollback()
    '            clsCommon.ProgressBarHide()
    '            clsCommon.MyMessageBoxShow(ex.Message & " At Line No : " & i)
    '        End Try

    '    End If
    '    Me.Controls.Remove(gv)

    'End Sub

    Private Sub lblMilkAvlblForSale_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMilkAvlblForSale.Click

    End Sub

    Private Sub txtNoOfChildrenMember_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNoOfChildrenMember.TextChanged, txtNoOfAdultMember.TextChanged
        txtNoOfTotalDependentMember.Text = clsCommon.myCdbl(txtNoOfChildrenMember.Text) + clsCommon.myCdbl(txtNoOfAdultMember.Text)
    End Sub



    Private Sub mnuExportAnimalDetails_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportAnimalDetails.Click
        Try
            '----------------------BM00000002953(Monika)-------for blank export sheet------------
            Dim str As String = "select count(*) from tspl_Animal_details"
            Dim check As Integer = clsDBFuncationality.getSingleValue(str)

            If check > 0 Then
                str = "select tspl_Animal_details.Trans_Code as [MP Code] ,tspl_Animal_details.Line_No as [Line No] ,tspl_Animal_details.Type_of_Animal as [Type Of Animal],tspl_Animal_details.Count_of_Animal as [Count Of Animal],tspl_Animal_details.Bread_of_Animal as [Bread Of Animal],tspl_Animal_details.Animal_Staus as [Status(Dry/Milk/None)]  From tspl_Animal_details"
            Else
                str = "select '' as [MP Code] ,'' as [Line No] ,'' as [Type Of Animal],0 as [Count of Animal],'' as [Bread Of Animal],'' as [Status(Dry/Milk/None)]"
            End If
            ListImpExpColumnsMandatory = New List(Of String)({"MP Code", "Line No", "Type Of Animal", "Count Of Animal", "Breed Of Animal"})
            ListImpExpColumnsSuperMandatory = New List(Of String)({"MP Code"})
            transportSql.ExporttoExcel(str, "", "", Me, ListImpExpColumnsMandatory, ListImpExpColumnsSuperMandatory, MyBase.Form_ID + "AnimalDetails")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub


    Private Sub MunTHIRD_PARTY_CODE_Click(sender As Object, e As EventArgs) Handles MunTHIRD_PARTY_CODE.Click
        Try
            Dim str As String = "select count(*) from TSPL_MP_MASTER"
            Dim check As Integer = clsDBFuncationality.getSingleValue(str)

            If check > 0 Then
                str = "Select MP_Code as [MP Code],MP_Name as [MP Name],MP_Code_VLC_Uploader as [Mp uploder Code],VLCH.VLC_Code as [Dcs Code],VLCH.VLC_Name as [Dcs Name],TSPL_MP_MASTER.THIRD_PARTY_CODE as [THIRD PARTY CODE],TSPL_MP_MASTER.Third_Party_Source as [THIRD PARTY SOURCE] from TSPL_MP_MASTER
                        left join TSPL_VLC_MASTER_HEAD VLCH on tspl_mp_master.VLC_Code=VLCH.VLC_Code"
            Else
                str = "select '' as [MP Code] ,'' as [MP Name] ,'' as [Mp uploder Code],0 as [Dcs Code],'' as [Dcs Name],'' as [THIRD PARTY CODE] '' as [THIRD PARTY SOURCE]"
            End If
            ListImpExpColumnsMandatory = New List(Of String)({"MP Code", "MP Name", "Mp uploder Code", "Dcs Code", "Dcs Name", "THIRD PARTY CODE", "THIRD PARTY SOURCE"})
            ListImpExpColumnsSuperMandatory = New List(Of String)({"MP Code"})
            transportSql.ExporttoExcel(str, "", "", Me, ListImpExpColumnsMandatory, ListImpExpColumnsSuperMandatory, MyBase.Form_ID + "THIRD_PARTY_CODE")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub RadMenuItem4_Click(sender As Object, e As EventArgs) Handles RadMenuItem4.Click

        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)

        Dim i As Integer = 0
        Dim totalAnimal As Double = 0
        Dim trans As SqlTransaction

        connectSql.OpenConnection()

        Dim mpCode As String = String.Empty

        trans = clsDBFuncationality.GetTransactin()

        Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")

        If transportSql.importExcel(gv, "MP Code", "MP Name", "Mp uploder Code", "Dcs Code", "Dcs Name", "THIRD PARTY CODE", "Third Party Source") Then

            Try

                clsCommon.ProgressBarShow()

                Dim arrMP As New List(Of String)

                For Each grow As GridViewRowInfo In gv.Rows

                    Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)

                    If clsCommon.myLen(strData) > 0 Then

                        If Not arrMP.Contains(strData) Then
                            arrMP.Add(strData)
                        End If

                    End If

                Next

                For Each grow As GridViewRowInfo In gv.Rows

                    Dim obj As New clsAnimalDetails

                    i = i + 1

                    '========================
                    ' MP CODE
                    '========================

                    Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)

                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("MP Code Can Not Be Left Blank")
                    End If

                    If clsCommon.myLen(strData) > 30 Then
                        Throw New Exception("MP Code Can Not Be Larger Then 30 Character")
                    End If

                    If clsDBFuncationality.getSingleValue(
                        "select count(*) from tspl_mp_master where mp_code='" & strData & "'",
                        trans) = 0 Then

                        Throw New Exception("Invalid MP Code. Code Not Found In Master")

                    End If

                    obj.Trans_Code = strData
                    mpCode = obj.Trans_Code

                    '========================
                    ' MP NAME
                    '========================

                    strData = clsCommon.myCstr(grow.Cells("MP Name").Value)

                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("MP Name Can Not Be Left Blank")
                    End If

                    'obj.Line_No = strData

                    '========================
                    ' MP UPLODER CODE
                    '========================

                    strData = clsCommon.myCstr(grow.Cells("Mp uploder Code").Value)

                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("Mp uploder Code Can Not Be Left Blank")
                    End If

                    '========================
                    ' DCS CODE
                    '========================

                    'Dim dblData As String = clsCommon.myCstr(grow.Cells("Dcs Code").Value)

                    'If dblData <= 0 Then
                    '    Throw New Exception("Dcs Code Can not be <=0")
                    'End If

                    Dim dcsCode As String = clsCommon.myCstr(grow.Cells("Dcs Code").Value)

                    If clsCommon.myLen(dcsCode) <= 0 Then
                        Throw New Exception("Dcs Code Can Not Be Blank")
                    End If

                    '========================
                    ' DCS NAME
                    '========================

                    'Dim dblDcsName As String = clsCommon.myCstr(grow.Cells("Dcs Name").Value)

                    'If dblDcsName <= 0 Then
                    '    Throw New Exception("Dcs Name Can not be <=0")
                    'End If

                    Dim dcsName As String = clsCommon.myCstr(grow.Cells("Dcs Name").Value)

                    If clsCommon.myLen(dcsName) <= 0 Then
                        Throw New Exception("Dcs Name Can Not Be Blank")
                    End If

                    '========================
                    ' THIRD PARTY CODE
                    '========================

                    Dim dblTHIRDPARTYCODE As String =
                        clsCommon.myCstr(grow.Cells("THIRD PARTY CODE").Value)

                    If clsCommon.myLen(dblTHIRDPARTYCODE) <= 0 Then
                        Throw New Exception("THIRD PARTY CODE Can Not Be Blank")
                    End If

                    '========================
                    ' THIRD PARTY CODE
                    '========================

                    Dim THIRDPARTYSOURCE As String =
                        clsCommon.myCstr(grow.Cells("THIRD PARTY SOURCE").Value)

                    If clsCommon.myLen(THIRDPARTYSOURCE) <= 0 Then
                        Throw New Exception("THIRD PARTY SOURCE Can Not Be Blank")
                    End If
                    If clsCommon.CompairString("REIL", THIRDPARTYSOURCE) = CompairStringResult.Equal OrElse clsCommon.CompairString("KTPL", THIRDPARTYSOURCE) = CompairStringResult.Equal OrElse clsCommon.CompairString("STAP", THIRDPARTYSOURCE) = CompairStringResult.Equal OrElse clsCommon.CompairString("EVST", THIRDPARTYSOURCE) = CompairStringResult.Equal Then
                    Else
                        Throw New Exception("INVALID MACHINE NAME ")
                    End If

                    '========================
                    ' SAVE DATA
                    '========================

                    clsAnimalDetails.SaveData(False, obj, trans)

                    ''========================
                    '' UPDATE THIRD PARTY CODE
                    ''========================

                    'clsDBFuncationality.ExecuteNonQuery(
                    '    "UPDATE TSPL_MP_MASTER " &
                    '    "SET THIRD_PARTY_CODE='" & dblTHIRDPARTYCODE & "', " &
                    '    "Third_Party_Source='" & THIRDPARTYSOURCE & "' " &
                    '    "WHERE MP_CODE='" & mpCode & "'",
                    '    trans)


                    '========================
                    ' CHECK DUPLICATE THIRD PARTY CODE
                    '========================

                    Dim duplicateCount As Integer = clsCommon.myCstr(
                        clsDBFuncationality.getSingleValue(
                            "SELECT COUNT(*) 
                             FROM TSPL_MP_MASTER 
                             WHERE THIRD_PARTY_CODE='" & dblTHIRDPARTYCODE & "' 
                             AND MP_CODE<>'" & mpCode & "'",
                            trans))

                    If duplicateCount > 0 Then

                        Throw New Exception(
                            "THIRD PARTY CODE : " & dblTHIRDPARTYCODE &
                            " already exists.")

                    End If

                    '========================
                    ' UPDATE DATA
                    '========================

                    clsDBFuncationality.ExecuteNonQuery("UPDATE TSPL_MP_MASTER " & "SET THIRD_PARTY_CODE='" & dblTHIRDPARTYCODE & "', " & "Third_Party_Source='" & THIRDPARTYSOURCE & "' " &
                        "WHERE MP_CODE='" & mpCode & "'", trans)
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, mpCode, "TSPL_MP_MASTER", "MP_CODE", trans)

                    ''========================
                    '' UPDATE ANIMAL COUNT
                    ''========================

                    'clsDBFuncationality.ExecuteNonQuery(
                    '    "UPDATE tspl_mp_master " &
                    '    "SET No_Of_breedable_milk_animal=" &
                    '    clsCommon.myCdbl(
                    '        clsDBFuncationality.getSingleValue(
                    '            "select sum(THIRD_PARTY_CODE) " &
                    '            "from TSPL_MP_MASTER " &
                    '            "where mp_code='" & mpCode & "'",
                    '            trans)) &
                    '    " WHERE mp_code='" & mpCode & "'",
                    '    trans)

                Next

                trans.Commit()

                clsCommon.ProgressBarHide()

                common.clsCommon.MyMessageBoxShow(
                    Me,
                    "Data Transfer Completed!",
                    Me.Text,
                    MessageBoxButtons.OK)

            Catch ex As Exception

                trans.Rollback()

                clsCommon.ProgressBarHide()

                clsCommon.MyMessageBoxShow(
                    Me,
                    ex.Message & " At Line No : " & i,
                    Me.Text)

            End Try

        End If

        Me.Controls.Remove(gv)

        'Dim gv As New UserControls.MyRadGridView
        'Me.Controls.Add(gv)
        'Dim i As Integer = 0
        'Dim totalAnimal As Double = 0
        'Dim trans As SqlTransaction
        'connectSql.OpenConnection()
        'Dim mpCode As String = String.Empty
        'trans = clsDBFuncationality.GetTransactin()
        'Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
        'If transportSql.importExcel(gv, "MP Code", "MP Name", "Mp uploder Code", "Dcs Code", "Dcs Name", "THIRD PARTY CODE") Then
        '    Try
        '        clsCommon.ProgressBarShow()
        '        Dim arrMP As New List(Of String)
        '        For Each grow As GridViewRowInfo In gv.Rows
        '            Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)
        '            If clsCommon.myLen(strData) > 0 Then
        '                If Not arrMP.Contains(strData) Then
        '                    arrMP.Add(strData)
        '                End If
        '            End If
        '        Next
        '        'If arrMP IsNot Nothing AndAlso arrMP.Count > 0 Then
        '        '    Dim qry As String = "delete from tspl_mp_master where prog_code='" & Me.Form_ID & "' and trans_code in (" + clsCommon.GetMulcallString(arrMP) + ")"
        '        '    clsDBFuncationality.ExecuteNonQuery(qry, trans)
        '        'End If

        '        For Each grow As GridViewRowInfo In gv.Rows
        '            'totalAnimal = 0
        '            Dim obj As New clsAnimalDetails
        '            i = i + 1
        '            Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)
        '            If clsCommon.myLen(strData) <= 0 Then
        '                Throw New Exception("MP Code Can Not Be Left Blank")
        '            End If
        '            If clsCommon.myLen(strData) > 30 Then
        '                Throw New Exception("MP Code Can Not Be Larger Then 30 Charachter")
        '            End If

        '            If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & strData & "'", trans) = 0 Then
        '                Throw New Exception("Invalid MP Code. Code Not Found In Master")
        '            End If
        '            obj.Trans_Code = strData
        '            mpCode = obj.Trans_Code
        '            strData = clsCommon.myCstr(grow.Cells("MP Name").Value)
        '            If clsCommon.myLen(strData) <= 0 Then
        '                Throw New Exception("Line No Can Not Be Left Blank")
        '            End If
        '            obj.Line_No = strData

        '            strData = clsCommon.myCstr(grow.Cells("Mp uploder Code").Value)
        '            If clsCommon.myLen(strData) <= 0 Then
        '                Throw New Exception("Type Of Animal  Can Not Be Left Blank")
        '            End If

        '            Dim dblData As String = clsCommon.myCstr(grow.Cells("Dcs Code").Value)
        '            If dblData <= 0 Then
        '                Throw New Exception("Count Of Animal  Can not be <=0")
        '            End If

        '            Dim dblDcsName As String = clsCommon.myCstr(grow.Cells("Dcs Name").Value)
        '            If dblDcsName <= 0 Then
        '                Throw New Exception("Count Of Animal  Can not be <=0")
        '            End If

        '            Dim dblTHIRDPARTYCODE As Double = clsCommon.myCdbl(grow.Cells("THIRD PARTY CODE").Value)
        '            If dblTHIRDPARTYCODE <= 0 Then
        '                Throw New Exception("Count Of Animal  Can not be <=0")
        '            End If
        '            clsAnimalDetails.SaveData(False, obj, trans)
        '            clsDBFuncationality.ExecuteNonQuery("Update tspl_mp_master set No_Of_breedable_milk_animal=" & clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(count_of_animal) from tspl_Animal_Details where trans_code='" & mpCode & "'", trans)) & " where mp_code='" & mpCode & "'", trans)
        '        Next


        '        trans.Commit()
        '        clsCommon.ProgressBarHide()
        '        common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!, ", Me.Text, MessageBoxButtons.OK)
        '    Catch ex As Exception
        '        trans.Rollback()
        '        clsCommon.ProgressBarHide()
        '        clsCommon.MyMessageBoxShow(Me, ex.Message & " At Line No : " & i, Me.Text)
        '    End Try

        'End If
        'Me.Controls.Remove(gv)
    End Sub






    Private Sub mnuImportAnimalDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuImportAnimalDetails.Click
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim i As Integer = 0
        Dim totalAnimal As Double = 0
        Dim trans As SqlTransaction
        connectSql.OpenConnection()
        Dim mpCode As String = String.Empty
        trans = clsDBFuncationality.GetTransactin()
        Dim strdate As Date = clsCommon.GETSERVERDATE(trans, "dd/MMM/yyyy")
        If transportSql.importExcel(gv, "MP Code", "Line No", "Type of Animal", "Count of Animal", "Bread Of Animal", "Status(Dry/Milk/None)") Then
            Try
                clsCommon.ProgressBarShow()
                Dim arrMP As New List(Of String)
                For Each grow As GridViewRowInfo In gv.Rows
                    Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)
                    If clsCommon.myLen(strData) > 0 Then
                        If Not arrMP.Contains(strData) Then
                            arrMP.Add(strData)
                        End If
                    End If
                Next
                If arrMP IsNot Nothing AndAlso arrMP.Count > 0 Then
                    Dim qry As String = "delete from tspl_Animal_Details where prog_code='" & Me.Form_ID & "' and trans_code in (" + clsCommon.GetMulcallString(arrMP) + ")"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If

                For Each grow As GridViewRowInfo In gv.Rows
                    'totalAnimal = 0
                    Dim obj As New clsAnimalDetails
                    i = i + 1
                    Dim strData As String = clsCommon.myCstr(grow.Cells("MP Code").Value)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("MP Code Can Not Be Left Blank")
                    End If
                    If clsCommon.myLen(strData) > 30 Then
                        Throw New Exception("MP Code Can Not Be Larger Then 30 Charachter")
                    End If

                    If clsDBFuncationality.getSingleValue("select count(*) from tspl_mp_master where mp_code='" & strData & "'", trans) = 0 Then
                        Throw New Exception("Invalid MP Code. Code Not Found In Master")
                    End If
                    obj.Trans_Code = strData
                    mpCode = obj.Trans_Code
                    strData = clsCommon.myCstr(grow.Cells("Line No").Value)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("Line No Can Not Be Left Blank")
                    End If
                    obj.Line_No = strData

                    strData = clsCommon.myCstr(grow.Cells("Type Of Animal").Value)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("Type Of Animal  Can Not Be Left Blank")
                    End If

                    Dim dblData As Double = clsCommon.myCdbl(grow.Cells("Count Of Animal").Value)
                    If dblData <= 0 Then
                        Throw New Exception("Count Of Animal  Can not be <=0")
                    End If
                    obj.Count_Of_animal = dblData
                    totalAnimal = totalAnimal + dblData
                    If clsCommon.CompairString(strData, "Buffalo") = CompairStringResult.Equal OrElse clsCommon.CompairString(strData, "Cow") = CompairStringResult.Equal OrElse clsCommon.CompairString(strData, "Camel") = CompairStringResult.Equal OrElse clsCommon.CompairString(strData, "Goat") = CompairStringResult.Equal Then
                    Else
                        Throw New Exception("Type of Animal Can be Either , Cow/Camel/Buffalo/Goat")
                    End If

                    obj.Type_Of_Animal = strData

                    obj.Animal_Staus = clsCommon.myCstr(grow.Cells("Status(Dry/Milk/None)").Value)
                    If clsCommon.CompairString(obj.Animal_Staus, "Dry") = CompairStringResult.Equal Then
                        obj.Animal_Staus = "Dry"
                    ElseIf clsCommon.CompairString(obj.Animal_Staus, "Milk") = CompairStringResult.Equal Then
                        obj.Animal_Staus = "Milk"
                    Else
                        obj.Animal_Staus = ""
                    End If


                    strData = clsCommon.myCstr(grow.Cells("Breed Of Animal").Value)
                    If clsCommon.myLen(strData) <= 0 Then
                        Throw New Exception("Breed Of Animal  Can Not Be Left Blank")
                    End If
                    obj.Bread_of_Animal = strData

                    obj.Prog_Code = Me.Form_ID

                    clsAnimalDetails.SaveData(False, obj, trans)
                    clsDBFuncationality.ExecuteNonQuery("Update tspl_mp_master set No_Of_breedable_milk_animal=" & clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(count_of_animal) from tspl_Animal_Details where trans_code='" & mpCode & "'", trans)) & " where mp_code='" & mpCode & "'", trans)
                Next


                trans.Commit()
                clsCommon.ProgressBarHide()
                common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!, ", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                trans.Rollback()
                clsCommon.ProgressBarHide()
                clsCommon.MyMessageBoxShow(Me, ex.Message & " At Line No : " & i, Me.Text)
            End Try

        End If
        Me.Controls.Remove(gv)

    End Sub

    Private Sub txtPayeeName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPayeeName.TextChanged

    End Sub

    Private Sub txtMPName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMPName.TextChanged
        txtPayeeName.Text = clsCommon.myCstr(txtMPName.Text)
    End Sub

    Private Sub mnuExportMPCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExportMPCode.Click
        Dim dt As System.Data.DataTable = clsDBFuncationality.GetDataTable("select MP_Code from tspl_mp_Master")
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim sfd As SaveFileDialog = New SaveFileDialog()
            Dim filePath As String
            sfd.FileName = Me.Text
            Dim strData As String = "MP Code" & Environment.NewLine
            sfd.Filter = "Text files (*.txt) |*.txt"
            If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                filePath = sfd.FileName
                clsCommon.ProgressBarShow()
                For i As Integer = 0 To dt.Rows.Count - 1
                    strData = strData & dt.Rows(i)("Mp_Code").ToString & Environment.NewLine
                Next
                'Dim fl As System.IO.File
                ' fl.WriteAllText(sfd.FileName, strData)
                File.WriteAllText(sfd.FileName, strData)
                clsCommon.ProgressBarHide()
                clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed", Me.Text)
            End If
        Else
            clsCommon.MyMessageBoxShow(Me, "No Data Found to export", Me.Text)
        End If

    End Sub

    Sub SetDefaultBillDetails()
        Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Deafault_MP_Grp_Code,Deafault_MP_Payment_Code,Deafault_MP_Payment_Cycle,Deafault_MP_Terms_Code " _
         & " from tspl_MP_Master inner join TSPL_VLC_MASTER_HEAD on tspl_MP_Master.VLC_Code=TSPL_VLC_MASTER_HEAD.VLC_Code inner join tspl_Mcc_Master on MCC=tspl_Mcc_Master.MCC_Code where MP_Code='" & clsCommon.myCstr(fndMPCode.Value) & "'")
        If dt.Rows.Count > 0 Then
            FndMpTermsCode.Value = clsCommon.myCstr(dt.Rows(0).Item("Deafault_MP_Terms_Code"))
            FndMPPaymentCycle.Value = clsCommon.myCstr(dt.Rows(0).Item("Deafault_MP_Payment_Cycle"))
            FndMPPaymentCode.Value = clsCommon.myCstr(dt.Rows(0).Item("Deafault_MP_Payment_Code"))
            FndMPGrpCode.Value = clsCommon.myCstr(dt.Rows(0).Item("Deafault_MP_Grp_Code"))

            TxtMPGrpCode.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select group_desc from Tspl_vendor_group where ven_group_code='" + clsCommon.myCstr(FndMPGrpCode.Value) + "'"))
            txtMPtermcodedes.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Terms_Desc from TSPL_TERMS_MASTER where Terms_Code='" + clsCommon.myCstr(FndMpTermsCode.Value) + "'"))
            TxtMPPaymentCode.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select payment_desc from tspl_payment_code where payment_code='" + clsCommon.myCstr(FndMPPaymentCode.Value) + "'"))
            If clsCommon.myLen(FndMPPaymentCycle.Value) > 0 Then
                Dim obj As clsPaymentCycleMaster = clsPaymentCycleMaster.GetData(FndMPPaymentCycle.Value, NavigatorType.Current)
                TxtMPPaymentCycle.Text = obj.Description
            End If
        End If
    End Sub

    Private Sub fndCustAccSet__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndCustAccSet._MYValidating
        Try
            Dim whrcls As String = ""
            fndCustAccSet.Value = clsCustomerAccountSet.getFinder(whrcls, fndCustAccSet.Value, isButtonClicked)

            If clsCommon.myLen(fndCustAccSet.Value) > 0 Then
                txtCustAccSetDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Acct_Desc from TSPL_CUSTOMER_ACCOUNT_SET where Cust_Account='" + fndCustAccSet.Value + "' "))
            Else
                txtCustAccSetDesc.Text = ""
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub fndVendorAccSet__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndVendorAccSet._MYValidating
        Try
            Dim whrcls As String = ""
            fndVendorAccSet.Value = clsVendorAccountSet.getFinder(whrcls, fndVendorAccSet.Value, isButtonClicked)

            If clsCommon.myLen(fndVendorAccSet.Value) > 0 Then
                txtVendorAccDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Acct_Set_Desc from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code='" + fndVendorAccSet.Value + "' "))
            Else
                txtVendorAccDesc.Text = ""
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub txtIncentiveMult__My_Click(sender As Object, e As EventArgs) Handles txtIncentiveMult._My_Click
        Dim qry As String = " select INCENTIVE_CODE as Code,DESCRIPTION as Name,INCENTIVE_DATE as Date,INCENTIVE_TYPE as IncentiveType,SCHEME_FOR as [Scheme For],Calc_Type as [Calculation Type],Rate_Type as [Rate Type],Qty_Type as [Quantity Type] from TSPL_INCENTIVE_MASTER_HEAD "
        '' get already selected data
        Dim qrySel As String = "select MP_CODE,INCENTIVE_CODE from TSPL_MP_INCENTIVE where MP_CODE='" & fndMPCode.Value & "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qrySel)
        Dim arr As New ArrayList
        For Each dr As DataRow In dt.Rows
            arr.Add(clsCommon.myCstr(dr.Item("INCENTIVE_CODE")))
        Next
        If txtIncentiveMult.arrValueMember IsNot Nothing AndAlso txtIncentiveMult.arrValueMember.Count <= 0 Then
            txtIncentiveMult.arrValueMember = arr
        End If
        txtIncentiveMult.arrValueMember = clsCommon.ShowMultipleSelectForm(False, "IncenMulSelForMP", qry, "Code", "Name", txtIncentiveMult.arrValueMember, txtIncentiveMult.arrDispalyMember)
    End Sub

    Private Sub btnIncentiveClear_Click(sender As Object, e As EventArgs)
        txtIncentiveMult.arrValueMember = Nothing
    End Sub

    Private Sub btnIncentiveClear_Click_1(sender As Object, e As EventArgs) Handles btnIncentiveClear.Click
        txtIncentiveMult.arrValueMember = Nothing
    End Sub

    Private Sub txtIncentiveGLAccount__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtIncentiveGLAccount._MYValidating
        Dim query As String = "select Account_Code as [Account],Description from tspl_gl_accounts"
        txtIncentiveGLAccount.Value = clsCommon.ShowSelectForm("CogsInterBranchfd", query, "Account", " ControlAccount ='Y' ", txtIncentiveGLAccount.Value, "Account", isButtonClicked)
        lblIncentiveGLAccount.Text = clsGLAccount.GetName(txtIncentiveGLAccount.Value)
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(fndMPCode.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Select MP Code", Me.Text)
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowHistoryData(fndMPCode.Value, "MP_Code", "tspl_mp_master")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub lblAccountNo_DoubleClick(sender As Object, e As EventArgs) Handles lblAccountNo.DoubleClick
        Dim frm As New FrmPWD(Nothing)
        frm.strType = clsFixedParameterType.ImportorExport
        frm.strCode = clsFixedParameterCode.FarmerMaster
        frm.ShowDialog()
        If frm.isPasswordCorrect Then
            txtBankIFSC_M.Enabled = True
            txtAccountNo.Enabled = True
            RadGroupBox2.Visible = True
        End If
    End Sub

    Private Sub RadMenuItem2_Click(sender As Object, e As EventArgs) Handles RadMenuItem2.Click
        Dim frm As New frmTempleteImportMP()
        frm.Show()
    End Sub

    Private Sub txtMPNameHindi_Leave(sender As Object, e As EventArgs) Handles txtMPNameHindi.Leave
        clsMccMaster.ToEnglishInput()
    End Sub

    Private Sub txtMPNameHindi_Enter(sender As Object, e As EventArgs) Handles txtMPNameHindi.Enter
        clsMccMaster.ToHindiInput()
    End Sub

    Private Sub txtCastCategory__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtCastCategory._MYValidating
        Try
            Dim qry As String = "select CAST_CATEGORY_CODE as 'Code' , CAST_CATEGORY_NAME as 'Name' , DESCRIPTION from TSPL_CAST_CATEGORY_MASTER"
            txtCastCategory.Value = clsCommon.ShowSelectForm("Mp@Filter@CastCategory", qry, "Code", " ", txtCastCategory.Value, "Code", isButtonClicked)
            lblCastCategory.Text = clsDBFuncationality.getSingleValue("select case when  CAST_CATEGORY_CODE = CAST_CATEGORY_NAME and   len (isnull( DESCRIPTION,'')) > 0 then  DESCRIPTION else  CAST_CATEGORY_NAME end Name  from TSPL_CAST_CATEGORY_MASTER where CAST_CATEGORY_CODE = '" + txtCastCategory.Value + "' ")
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtBlockCode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtBlockCode._MYValidating
        Try
            Dim qry As String = " select BLOCK_CODE as Code, BLOCK_NAME as Name from TSPL_BLOCK_MASTER  "

            txtBlockCode.Value = clsCommon.ShowSelectForm("DCS@Block@Finder", qry, "Code", "", txtBlockCode.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtZone__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtZone._MYValidating
        Try
            Dim qry As String = " select TSPL_ZONE_MASTER.Zone_Code as Code, TSPL_ZONE_MASTER.Description as Name ,TSPL_ZONE_MASTER.City_Code as [City Code],TSPL_CITY_MASTER.City_Name as [City Name]  from TSPL_ZONE_MASTER left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code = TSPL_ZONE_MASTER.City_Code "

            txtZone.Value = clsCommon.ShowSelectForm("DCS@Zone@Finder", qry, "Code", "", txtZone.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtRevenueVillage__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtRevenueVillage._MYValidating
        Try
            Dim qry As String = " select REVENUE_VILLAGE_CODE as Code, REVENUE_VILLAGE_NAME as Name from TSPL_REVENUE_VILLAGE_MASTER  "

            txtRevenueVillage.Value = clsCommon.ShowSelectForm("DCS@RevenueVillage@Finder", qry, "Code", "", txtRevenueVillage.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtGrampanchayat__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtGrampanchayat._MYValidating
        Try
            Dim qry As String = " select GRAMPANCHAYAT_CODE as Code, GRAMPANCHAYAT_NAME as Name from TSPL_GRAMPANCHAYAT_MASTER  "

            txtGrampanchayat.Value = clsCommon.ShowSelectForm("DCS@Grampanchayat@Finder", qry, "Code", "", txtGrampanchayat.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtPanchayatSamiti__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtPanchayatSamiti._MYValidating
        Try
            Dim qry As String = " select PANCHAYAT_SAMITI_CODE as Code, PANCHAYAT_SAMITI_NAME as Name from TSPL_PANCHAYAT_SAMITI_MASTER  "

            txtPanchayatSamiti.Value = clsCommon.ShowSelectForm("DCS@PanchayatSamiti@Finder", qry, "Code", "", txtPanchayatSamiti.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtVidhanSabha__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVidhanSabha._MYValidating
        Try
            Dim qry As String = " select VIDHAN_SABHA_CODE as Code, VIDHAN_SABHA_NAME as Name from TSPL_VIDHAN_SABHA_MASTER  "

            txtVidhanSabha.Value = clsCommon.ShowSelectForm("DCS@VidhanSabha@Finder", qry, "Code", "", txtVidhanSabha.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtDistrict__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDistrict._MYValidating
        Try
            Dim qry As String = "select TSPL_DISTRICT_MASTER.Code as Code,TSPL_DISTRICT_MASTER.Name as DistrictName,TSPL_State_MASTER.STATE_CODE as [State Code] ,TSPL_State_MASTER.STATE_NAME as [State] " &
            " from TSPL_DISTRICT_MASTER " &
            " left outer join TSPL_State_MASTER  on TSPL_State_MASTER.STATE_CODE=TSPL_DISTRICT_MASTER.State_Code " &
            " left outer join TSPL_State_MASTER_detail on TSPL_State_MASTER.state_code=TSPL_State_MASTER_detail.state_code "

            txtDistrict.Value = clsCommon.ShowSelectForm("MP@District@Finder", qry, "Code", "", txtDistrict.Value, "", isButtonClicked)
            FillAddressDetails()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Sub FillAddressDetails()
        Try
            If clsCommon.myLen(txtDistrict.Value) > 0 Then
                lblDistrict.Text = clsDistrictMaster.GetName(txtDistrict.Value)
            End If
            If clsCommon.myLen(txtBlockCode.Value) > 0 Then
                lblBlockCode.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select BLOCK_NAME from TSPL_BLOCK_MASTER where BLOCK_CODE = '" + txtBlockCode.Value + "' "))
            End If
            If clsCommon.myLen(txtZone.Value) > 0 Then
                lblZone.Text = clsDBFuncationality.getSingleValue(" select Description from TSPL_ZONE_MASTER where Zone_Code = '" + txtZone.Value + "' ")
            End If
            If clsCommon.myLen(txtRevenueVillage.Value) > 0 Then
                lblRevenueVillage.Text = clsRevenueVillageMaster.GetName(txtRevenueVillage.Value)
            End If
            If clsCommon.myLen(txtGrampanchayat.Value) > 0 Then
                lblGrampanchayat.Text = clsGrampanchayatMaster.GetName(txtGrampanchayat.Value)
            End If
            If clsCommon.myLen(txtPanchayatSamiti.Value) > 0 Then
                lblPanchayatSamiti.Text = clsPanchayatSamitiMaster.GetName(txtPanchayatSamiti.Value)
            End If
            If clsCommon.myLen(txtVidhanSabha.Value) > 0 Then
                lblVidhanSabha.Text = clsVidhanSabhaMaster.GetName(txtVidhanSabha.Value)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub


    Private Sub txtDistrict__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtDistrict._MYOpenMasterForm
        Try
            Frm_Open = New frmDistrictMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.DistrictMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtBlockCode__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtBlockCode._MYOpenMasterForm
        Try
            Frm_Open = New frmBlockMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.frmBlockMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtZone__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtZone._MYOpenMasterForm
        Try
            Frm_Open = New FrmZoneMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.FrmZoneMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtRevenueVillage__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtRevenueVillage._MYOpenMasterForm
        Try
            Frm_Open = New frmRevenueVillageMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.frmRevenueVillageMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtGrampanchayat__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtGrampanchayat._MYOpenMasterForm
        Try
            Frm_Open = New frmGrampanchayatMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.frmGrampanchayatMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtPanchayatSamiti__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtPanchayatSamiti._MYOpenMasterForm
        Try
            Frm_Open = New frmPanchayatSamitiMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.frmPanchayatSamitiMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtVidhanSabha__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVidhanSabha._MYOpenMasterForm
        Try
            Frm_Open = New frmVidhanSabhaMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.frmVidhanSabhaMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub txtCastCategory__MYOpenMasterForm(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtCastCategory._MYOpenMasterForm
        Try
            Frm_Open = New frmCastCategoryMaster()
            Frm_Open.SetUserMgmt(clsUserMgtCode.frmCastCategoryMaster)
            Frm_Open.Show()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try
    End Sub

    Private Sub btnUnverifiedJanAdhaar_Click(sender As Object, e As EventArgs) Handles btnUnverifiedJanAdhaar.Click
        Dim frm As New FrmPWD(Nothing)
        frm.strType = clsFixedParameterType.SIRC
        frm.strCode = clsFixedParameterCode.UpdatePassword
        frm.ShowDialog()
        If frm.isPasswordCorrect Then
            ShowRemarks()
        End If
        btnUnverifiedJanAdhaar.Enabled = False
    End Sub
    Private Sub ShowRemarks()
        Try
            Dim Reason As String = ""
            Dim frm As New FrmFreeTxtBox1
            frm.Text = "Remarks for Verified"
            frm.ShowDialog()
            If clsCommon.myLen(frm.strRmks) <= 0 Then
                Exit Sub
            Else
                Reason = frm.strRmks
            End If
            Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
            Try
                saveCancelLog(Reason, "Verified", trans)
                Dim Verified As String = "update tspl_mp_master set Jan_Aadhar_No_Verified = 0 where MP_Code = '" & fndMPCode.Value & "'"
                clsDBFuncationality.ExecuteNonQuery(Verified, trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, fndMPCode.Value, "tspl_mp_master", "MP_Code", trans)
                trans.Commit()
                clsCommon.MyMessageBoxShow(Me, "This is Unverified Successfully", Me.Text)
            Catch ex As Exception
                trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Function saveCancelLog(ByVal Reason As String, ByVal Activity_Type As String, Optional ByVal trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim obj As New clsCancelLog
        obj.Program_Code = Form_ID
        obj.DOCUMENT_NO = clsCommon.myCstr(Me.fndMPCode.Value)
        obj.REASON = Reason
        obj.ACTIVITY_TYPE = Activity_Type
        Return clsCancelLog.SaveData(obj, True, trans)
    End Function

    Private Sub txtMPName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMPName.KeyPress
        If Char.IsControl(e.KeyChar) = False And Char.IsSeparator(e.KeyChar) = False And Char.IsLetter(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPayeeName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPayeeName.KeyPress

        If Char.IsControl(e.KeyChar) = False And Char.IsSeparator(e.KeyChar) = False And Char.IsLetter(e.KeyChar) = False Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtMPCodeVlcUploader_Click(sender As Object, e As EventArgs) Handles txtMPCodeVlcUploader.Click
        If clsCommon.myLen(txtMPCodeVlcUploader.Text) > 0 OrElse clsCommon.myLen(txtMPCodeVlcUploader.Text) <= 0 Then
            Dim mpcode As String = clsDBFuncationality.getSingleValue("Select MP_Uploader_Code
                from TSPL_DBT_NEFT_DETAIL 
                Left Outer Join TSPL_MP_INCENTIVE_ENTRY_DETAIL On TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR   
                left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
                where  MP_Uploader_Code= '" + txtMPCodeVlcUploader.Text + "' ")
            If clsCommon.myLen(mpcode) > 0 Then
                txtMPCodeVlcUploader.ReadOnly = True
            Else
                txtMPCodeVlcUploader.ReadOnly = False
            End If
        End If
    End Sub

    ' Restrict input to numbers only
    Private Sub txtPhone_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTelePhone.KeyPress
        ' Allow only digits and control characters (e.g., backspace)
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True  ' Ignore the keypress if it's not a digit or control character
        End If
    End Sub

    Private Sub btnImportBankDetails_Click(sender As Object, e As EventArgs) Handles btnImportBankDetails.Click
        Try
            Dim gv As New UserControls.MyRadGridView
            Me.Controls.Add(gv)
            If transportSql.importExcel(gv, "MP Code", "Account Number", "IFSC Code") Then
                Dim ArrAccountNo As New List(Of String)
                Dim ii As Integer = 0
                Dim Arr As New List(Of clsMpMaster)
                Dim dtError As New DataTable
                dtError.Columns.Add("RowNo", GetType(Integer))
                dtError.Columns.Add("Error", GetType(String))
                Try

                    If gv IsNot Nothing AndAlso gv.Rows.Count > 0 Then
                        clsCommon.ProgressBarPercentShow()
                        For Each grow As GridViewRowInfo In gv.Rows
                            Try
                                ii += 1
                                clsCommon.ProgressBarPercentUpdate(ii, gv.Rows.Count, "Validating Data...")

                                If clsCommon.myLen(clsCommon.myCstr(grow.Cells("MP Code").Value)) <= 0 Then
                                    Throw New Exception("MP Code can't be blank !")
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(grow.Cells("IFSC Code").Value)) <= 0 Then
                                    Throw New Exception("IFSC Code can't be blank !")
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(grow.Cells("Account Number").Value)) <= 0 Then
                                    Throw New Exception("Account Number can't be blank !")
                                End If

                                Dim obj As New clsMpMaster
                                obj.MP_Code = clsCommon.myCstr(grow.Cells("MP Code").Value)
                                obj.IFCICode = clsCommon.myCstr(grow.Cells("IFSC Code").Value)
                                Dim dt As DataTable = clsDBFuncationality.GetDataTable("select BANK,BRANCH,STATE,CITY from TSPL_MASTER.dbo.TSPL_IFSC where IFSC='" + obj.IFCICode + "'")
                                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                                    Throw New Exception("Invalid IFSC Code [" + obj.IFCICode + "]")
                                Else
                                    obj.BankName = clsCommon.myCstr(dt.Rows(0)("BANK"))
                                    obj.BankBranch = clsCommon.myCstr(dt.Rows(0)("BRANCH"))
                                    obj.BankStateCode = clsCommon.myCstr(dt.Rows(0)("STATE"))
                                    obj.BankCityCode = clsCommon.myCstr(dt.Rows(0)("CITY"))
                                End If
                                obj.AccountNO = clsCommon.myCstr(grow.Cells("Account Number").Value)
                                Dim chkCount As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select COUNT(*) from TSPL_MP_MASTER where MP_Code='" + clsCommon.myCstr(obj.MP_Code) + "'"))
                                If chkCount > 0 Then
                                Else
                                    Throw New Exception("MP Code (" + clsCommon.myCstr(obj.MP_Code) + ") is not exist !")
                                End If

                                'Dim qry As String = "select 1 from TSPL_MP_MASTER where MP_Code not in ('" & obj.MP_Code & " ') and AccountNO='" & obj.AccountNO & "'"
                                'dt = clsDBFuncationality.GetDataTable(qry)
                                'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                '    Throw New Exception("MP Code (" + clsCommon.myCstr(obj.MP_Code) + ") Duplicate account no [" + obj.AccountNO + "]")
                                'End If

                                Dim dtACNo As DataTable = Nothing
                                Dim qry As String
                                qry = "select MP_Code,Active,AccountNO from TSPL_MP_MASTER where MP_Code  in ('" + obj.MP_Code + "') and Active=1 " ' and AccountNO='" + obj.AccountNO + "' "
                                dtACNo = clsDBFuncationality.GetDataTable(qry)
                                If dtACNo Is Nothing OrElse dtACNo.Rows.Count <= 0 Then
                                    qry = "select MP_Code,AccountNO,Active from TSPL_MP_MASTER where MP_Code not in ('" & obj.MP_Code & "') and AccountNO='" & obj.AccountNO & "' "
                                    dtACNo = Nothing
                                    dtACNo = clsDBFuncationality.GetDataTable(qry)
                                    If dtACNo IsNot Nothing AndAlso dtACNo.Rows.Count > 0 Then
                                        Dim str As New StringBuilder()
                                        For Each strActive As DataRow In dtACNo.Rows
                                            If clsCommon.myCDecimal(strActive("Active")) = 0 Then
                                                If clsCommon.myLen(clsCommon.myCstr(str)) > 0 Then
                                                    str.Append(Environment.NewLine)
                                                End If
                                                str.Append("Account number [" + obj.AccountNO + "] is already in use for Farmer ['" & clsCommon.myCstr(strActive("MP_Code")) & "'].")
                                            End If
                                        Next
                                        If str IsNot Nothing AndAlso clsCommon.myLen(clsCommon.myCstr(str)) > 0 Then
                                            Throw New Exception(clsCommon.myCstr(str))
                                        End If
                                        str = Nothing
                                    End If
                                    ' Else
                                    'Throw New Exception("Farmer (" & clsCommon.myCstr(obj.MP_Code) & ") is inactive. You cannot update this record.")
                                End If
                                If ArrAccountNo.Contains(obj.AccountNO) Then
                                    Throw New Exception("Repeated Account number [" & obj.AccountNO & "] for Farmer [" & obj.MP_Code & "].")
                                End If
                                ArrAccountNo.Add(obj.AccountNO)
                                Arr.Add(obj)
                            Catch ex As Exception
                                Dim dr As DataRow = dtError.NewRow()
                                dr("RowNo") = ii
                                dr("Error") = ex.Message
                                dtError.Rows.Add(dr)
                            End Try
                        Next
                        clsCommon.ProgressBarPercentHide()
                    End If

                    Try
                        If dtError.Rows.Count > 0 Then
                            Dim ff As New FrmFreeGrid
                            ff.ReportID = "MilkShiftUploader"
                            ff.Text = "MP Master Errors"
                            ff.dt = dtError
                            ff.ShowDialog()
                        ElseIf Arr IsNot Nothing AndAlso Arr.Count > 0 Then
                            Dim qry As String = "Valid Row [" + clsCommon.myCstr(Arr.Count) + "] Do You want to Proceed"
                            If clsCommon.MyMessageBoxShow(Me, qry, Me.Text, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                clsCommon.ProgressBarPercentShow()
                                ii = 0
                                Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                                Try
                                    'Dim qry1 As String = "select  MP_Code from TSPL_MP_MASTER where MP_Code ='" + obj.MP_Code + "'AND active = 0  "
                                    '  Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry1, trans)
                                    'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                    '    'qry1 = "GRN is used in following MRN"
                                    '    For Each dr As DataRow In dt.Rows
                                    '        qry1 += Environment.NewLine + clsCommon.myCstr(dr("Active"))
                                    '    Next
                                    '    'qry1 += Environment.NewLine + "Can't cancel it"
                                    '    'clsCommon.MyMessageBoxShow(qry1)
                                    '    ' Exit Sub
                                    'End If

                                    For Each obj As clsMpMaster In Arr
                                        ii += 1
                                        clsCommon.ProgressBarPercentUpdate(ii, Arr.Count, "Saving Details..." & clsCommon.myCstr(ii) & "/" & clsCommon.myCstr(Arr.Count) & "")



                                        Dim UpdateQry As String = "Update TSPL_MP_MASTER Set BankName='" + obj.BankName + "',IFCICode='" + obj.IFCICode + "',AccountNO='" + obj.AccountNO + "' Where MP_Code='" + obj.MP_Code + "'"
                                        clsDBFuncationality.ExecuteNonQuery(UpdateQry, trans)
                                        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.MP_Code, "tspl_mp_master", "MP_Code", trans)
                                    Next
                                    trans.Commit()
                                Catch ex As Exception
                                    trans.Rollback()
                                    Throw New Exception(ex.Message)
                                Finally
                                    clsCommon.ProgressBarPercentHide()
                                End Try
                                clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                            End If
                        Else
                            Throw New Exception("No Valid Rows Found to Save")
                        End If
                    Catch ex As Exception
                        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                    End Try
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Public Shared Function ImportBankDetail(gv As RadGridView) As Boolean

        Return True
    End Function

    Private Sub btnExportBlankSheet_Click(sender As Object, e As EventArgs) Handles btnExportBlankSheet.Click
        Try
            Dim Qry As String = "Select '' As 'MP Code','' As 'Account Number','' As 'IFSC Code'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt.Rows.Count > 0 Then
                transportSql.ExporttoExcel(dt, Me)
                dt = Nothing
            Else
                Throw New Exception("No data found")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub mnuExportMPName_Click(sender As Object, e As EventArgs) Handles mnuExportMPName.Click
        Try


            Dim ExportSheet As String = "FillDataSheet"
            Dim MultiMPCodeName As ArrayList = Nothing
            Dim Qry As String = Nothing
            Qry = "select MP_Code as 'MP Code',MP_Name as 'MP Name',PayeeName as 'Payee Name' from tspl_mp_Master"
            MultiMPCodeName = clsCommon.ShowMultipleSelectForm("DCSMulSelect", Qry, "MP Code", "MP Name", MultiMPCodeName, MultiMPCodeName)

            If clsCommon.myLen(MultiMPCodeName) > 0 Then
                clsMpMaster.ExportDataTable(MultiMPCodeName, Me, ExportSheet)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message.ToString(), Me.Text)
        End Try

    End Sub
    Public Sub funImportMPName()
        Try
            Dim gv As New UserControls.MyRadGridView
            Me.Controls.Add(gv)
            If transportSql.importExcel(gv, "MP Code", "MP Name", "Payee Name") Then
                Dim ii As Integer = 0
                Dim Arr As New List(Of clsMpMaster)
                Dim dtError As New DataTable
                dtError.Columns.Add("RowNo", GetType(Integer))
                dtError.Columns.Add("Error", GetType(String))
                Try

                    If gv IsNot Nothing AndAlso gv.Rows.Count > 0 Then
                        clsCommon.ProgressBarPercentShow()
                        For Each grow As GridViewRowInfo In gv.Rows
                            Try
                                ii += 1
                                clsCommon.ProgressBarPercentUpdate(ii, gv.Rows.Count, "Validating Data...")

                                If clsCommon.myLen(clsCommon.myCstr(grow.Cells("MP Code").Value)) <= 0 Then
                                    Throw New Exception("MP Code can't be blank !")
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(grow.Cells("MP Name").Value)) <= 0 Then
                                    Throw New Exception("MP Name can't be blank !")
                                End If
                                If clsCommon.myLen(clsCommon.myCstr(grow.Cells("Payee Name").Value)) <= 0 Then
                                    Throw New Exception("Payee Name can't be blank !")
                                End If

                                Dim obj As New clsMpMaster
                                obj.MP_Code = clsCommon.myCstr(grow.Cells("MP Code").Value)
                                obj.MP_Name = clsCommon.myCstr(grow.Cells("MP Name").Value)
                                obj.PayeeName = clsCommon.myCstr(grow.Cells("Payee Name").Value)

                                'Dim dt As DataTable = clsDBFuncationality.GetDataTable("select BANK,BRANCH,STATE,CITY from TSPL_MASTER.dbo.TSPL_IFSC where IFSC='" + obj.IFCICode + "'")
                                'If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                                '    Throw New Exception("Invalid IFSC Code [" + obj.IFCICode + "]")
                                'Else
                                '    obj.BankName = clsCommon.myCstr(dt.Rows(0)("BANK"))
                                '    obj.BankBranch = clsCommon.myCstr(dt.Rows(0)("BRANCH"))
                                '    obj.BankStateCode = clsCommon.myCstr(dt.Rows(0)("STATE"))
                                '    obj.BankCityCode = clsCommon.myCstr(dt.Rows(0)("CITY"))
                                'End If
                                ' obj.AccountNO = clsCommon.myCstr(grow.Cells("Account Number").Value)
                                'Dim chkCount As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select COUNT(*) from TSPL_MP_MASTER where MP_Code='" + clsCommon.myCstr(obj.MP_Code) + "'"))
                                'If chkCount > 0 Then
                                'Else
                                '    Throw New Exception("MP Code (" + clsCommon.myCstr(obj.MP_Code) + ") is not exist !")
                                'End If
                                Arr.Add(obj)
                            Catch ex As Exception
                                Dim dr As DataRow = dtError.NewRow()
                                dr("RowNo") = ii
                                dr("Error") = ex.Message
                                dtError.Rows.Add(dr)
                            End Try
                        Next
                        clsCommon.ProgressBarPercentHide()
                    End If

                    Try
                        If dtError.Rows.Count > 0 Then
                            Dim ff As New FrmFreeGrid
                            ff.ReportID = "MilkShiftUploader"
                            ff.Text = "MP Master Errors"
                            ff.dt = dtError
                            ff.ShowDialog()
                        ElseIf Arr IsNot Nothing AndAlso Arr.Count > 0 Then
                            Dim qry As String = "Valid Row [" + clsCommon.myCstr(Arr.Count) + "] Do You want to Proceed"
                            If clsCommon.MyMessageBoxShow(Me, qry, Me.Text, MessageBoxButtons.YesNo) = DialogResult.Yes Then
                                clsCommon.ProgressBarPercentShow()
                                ii = 0
                                Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                                Try
                                    For Each obj As clsMpMaster In Arr
                                        ii += 1
                                        clsCommon.ProgressBarPercentUpdate(ii, Arr.Count, "Saving Details..." & clsCommon.myCstr(ii) & "/" & clsCommon.myCstr(Arr.Count) & "")

                                        Dim UpdateQry As String = "Update TSPL_MP_MASTER Set mp_name='" + obj.MP_Name + "',PayeeName='" + obj.PayeeName + "' Where MP_Code='" + obj.MP_Code + "'"
                                        clsDBFuncationality.ExecuteNonQuery(UpdateQry, trans)
                                        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.MP_Code, "tspl_mp_master", "MP_Code", trans)
                                    Next
                                    trans.Commit()
                                Catch ex As Exception
                                    trans.Rollback()
                                    Throw New Exception(ex.Message)
                                Finally
                                    clsCommon.ProgressBarPercentHide()
                                End Try
                                clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                            End If
                        Else
                            Throw New Exception("No Valid Rows Found to Save")
                        End If
                    Catch ex As Exception
                        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                    End Try
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
    Private Sub RadMnuMPName_Click(sender As Object, e As EventArgs) Handles RadMnuMPName.Click
        ' funImportMPName()
    End Sub

    Private Sub RadMenuItem3_Click(sender As Object, e As EventArgs) Handles RadMenuItem3.Click
        funImportMPName()
    End Sub
End Class
