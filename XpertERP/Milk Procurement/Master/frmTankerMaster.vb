'-------Created By Monika 26/05/2014
'--------------------BM00000003316
Imports common
Imports System.Data.SqlClient
Public Class FrmTankerMaster
    Inherits FrmMainTranScreen

#Region "Variables"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim isImportClicked As Boolean = False
    Dim Errorcontrol As clsErrorControl = New clsErrorControl()
    Dim isNewEntry As Boolean = True
    Public Const colSlabUpto As String = "colSlabUpto"
    Public Const colSlabRate As String = "colSlabRate"
    Dim Frm_Open As FrmMainTranScreen

    Const colChamberNo As String = "colChamberNo"
    Const colChamberDescription As String = "colChamberDescription"

    Dim IsChamberWiseTanker As Boolean = False
    Dim AllowSameTankerNoInMaster As Boolean = False

#End Region

    Private Sub FrmTankerMaster_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Reset()
        SetUserMgmtNew()
        RadPageView1.SelectedPage = RadPageViewPage1
        ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update ")
        ButtonToolTip.SetToolTip(btndelete, "Press Alt+D  for Delete ")
        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")

        IsChamberWiseTanker = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.IsChamberWiseTanker, clsFixedParameterCode.IsChamberWiseTanker, Nothing)) = 1, True, False)
        AllowSameTankerNoInMaster = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowSameTankerNoforPrimarySecondaryTransporter, clsFixedParameterCode.AllowSameTankerNoforPrimarySecondaryTransporter, Nothing)) = 1, True, False)
        If Not IsChamberWiseTanker Then
            RadPageView1.Pages("RadPageViewPage3").Item.Visibility = ElementVisibility.Collapsed
        End If
    End Sub

    Function isDuplicateVechiele(ByVal VNO As String) As Boolean
        Dim rValue As Boolean = False
        Dim qry As String = " select count(xx.V_No)  from (select Tanker_No  as V_No from TSPL_TANKER_MASTER union all  select Vehicle_Code as V_No  from TSPL_Primary_Vehicle_Master )xx  where xx.V_No ='" & VNO & "'"
    Dim intCnt As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
        If intCnt > 0 Then
            rValue = True

        Else
            rValue = False
        End If
        Return rValue
    End Function

    Sub ResetCheckBox()
        If rbtnrental.Checked Then
            rbtnratekm.Checked = False
            rbtnrateltr.Checked = False
            rbtndiesel.Checked = False
        ElseIf rbtnratekm.Checked Then
            rbtnrateltr.Checked = False
            rbtndiesel.Checked = False
            rbtnrental.Checked = False
        ElseIf rbtndiesel.Checked Then
            rbtnrateltr.Checked = False
            rbtnrental.Checked = False
            rbtnratekm.Checked = False
        ElseIf rbtnrateltr.Checked Then
            rbtndiesel.Checked = False
            rbtnrental.Checked = False
            rbtnratekm.Checked = False
        End If
    End Sub

    Sub LoadBlankGridChamber()
        Try
            gvChamber.Rows.Clear()
            gvChamber.Columns.Clear()

            Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
            repoLineNo.FormatString = ""
            repoLineNo.HeaderText = "Chamber No"
            repoLineNo.Name = colChamberNo
            repoLineNo.Width = 100
            repoLineNo.ReadOnly = True
            repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvChamber.MasterTemplate.Columns.Add(repoLineNo) '0

            Dim reposerialno As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            reposerialno.FormatString = ""
            reposerialno.HeaderText = "Description"
            reposerialno.Name = colChamberDescription
            reposerialno.ReadOnly = False
            reposerialno.Width = 100
            gvChamber.MasterTemplate.Columns.Add(reposerialno)

            gvChamber.AllowAddNewRow = False
            gvChamber.AddNewRowPosition = SystemRowPosition.Bottom
            gvChamber.AllowEditRow = True
            gvChamber.AllowDeleteRow = True
            gvChamber.AllowRowResize = False
            gvChamber.AllowRowReorder = False
            gvChamber.AllowColumnResize = True
            gvChamber.AllowColumnChooser = False
            gvChamber.AllowAutoSizeColumns = False
            gvChamber.ShowGroupPanel = False
            gvChamber.ShowFilteringRow = False
            gvChamber.AllowColumnReorder = False
            gvChamber.EnableSorting = False
            gvChamber.MasterTemplate.ShowRowHeaderColumn = False
            gvChamber.TableElement.TableHeaderHeight = 40
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub Reset()
        txtTankerNo.Text = ""
        txtname.Text = ""
        fndNo.Value = ""
        txttank_transcode.Value = ""
        TxtESIRule.Value = ""
        TxtEPFRule.Value = ""
        txtstorage.Value = 0
        TxtIceCharge.Value = 0
        NLabourCharge.Value = 0
        txtyear.Text = ""
        rbtninner_yes.IsChecked = False
        rbtinner_no.IsChecked = False
        rbtnouter_no.IsChecked = False
        rbtnouter_yes.IsChecked = False
        isImportClicked = False
        fndNo.MyReadOnly = False
        txtTankerNo.ReadOnly = False
        txtdesc.Text = ""
        txtchrg.Text = ""
        txtavgkm.Text = ""
        txtdiesel.Text = ""
        txtrental_day.Text = ""
        txtrental_month.Text = ""
        txtrental_week.Text = ""
        txt_km.Text = ""
        txt_ltr.Text = ""
        txt_ltr_kg.Text = ""
        rbtndiesel.Checked = False
        rbtnrateltr.Checked = False
        rbtnrental.Checked = False
        rbtnratekm.Checked = False
        txtRentalAmt.Text = ""
        cmbRentalType.Enabled = False
        cmbRentalType.SelectedIndex = 0
        txtProvMinQty.Value = 0
        UcAttachment1.Form_ID = Me.Form_ID
        UcAttachment1.BlankAllControls()
        ''richa Against Ticket No. BM00000003713 on 03/09/2014
        'ddlStorageCapacityDescription.Items.Clear()
        transportSql.FillComboBox(" select 'KG'  as Code union all select 'Liter' as Code ", ddlStorageCapacityDescription, "Code", "Code")
        ddlStorageCapacityDescription.Text = "KG"
        ''----------------------------------------------------
        chkMonthlyRentPlusDiesel.Checked = False
        txtMRPDRent.Value = Nothing
        txtMRPDAverage.Value = Nothing
        txtMRPDDieselRate.Value = Nothing

        isNewEntry = True
        btnsave.Text = "Save"
        btndelete.Enabled = False
        cmbLtrKG.SelectedIndex = 0
        txt_ltr.Text = ""
        rbtKmrange.Checked = False
        txtChamborNo.Value = 0
        LoadBlankGridChamber()
        loadBlankgv()
        gv.Rows.AddNew()
        chkInactive.Checked = False
        ChkPrivate.Checked = False
    End Sub
    Function isBlankGV() As Boolean
        Dim rValue As Boolean = True
        If gv.Rows.Count = 0 Then
            rValue = False
            Return rValue
        End If
        If clsCommon.myCdbl(gv.Rows(0).Cells(colSlabUpto).Value) <= 0 Or clsCommon.myCdbl(gv.Rows(0).Cells(colSlabRate).Value) <= 0 Then
            rValue = False
            Return rValue
        End If
    End Function
    Function isValidateGridValue() As Boolean
        Dim rValue As Boolean = True
        If gv.Rows.Count > 1 Then
            For i As Integer = 1 To gv.Rows.Count - 1
                If clsCommon.myCdbl(gv.Rows(i).Cells(colSlabUpto).Value) <= clsCommon.myCdbl(gv.Rows(i - 1).Cells(colSlabUpto).Value) AndAlso clsCommon.myCdbl(gv.Rows(i).Cells(colSlabUpto).Value) > 0 AndAlso clsCommon.myLen(gv.Rows(i).Cells(colSlabUpto).Value) > 0 Then
                    rValue = False
                End If
            Next
        End If
        Return rValue
    End Function
    Sub loadBlankgv()
        Try
            gv.Rows.Clear()
            gv.Columns.Clear()
            Dim repoDeciCol As GridViewDecimalColumn
            repoDeciCol = New GridViewDecimalColumn()
            repoDeciCol.Name = colSlabUpto
            repoDeciCol.Width = 120
            repoDeciCol.DecimalPlaces = 0
            repoDeciCol.Minimum = 1
            repoDeciCol.HeaderText = "Slab Upto"
            gv.MasterTemplate.Columns.Add(repoDeciCol)

            repoDeciCol = New GridViewDecimalColumn()
            repoDeciCol.Name = colSlabRate
            repoDeciCol.Width = 300
            repoDeciCol.DecimalPlaces = 2
            repoDeciCol.Minimum = 1
            repoDeciCol.HeaderText = "Slab Rate"
            gv.MasterTemplate.Columns.Add(repoDeciCol)

            gv.AllowAddNewRow = True
            gv.AllowEditRow = True
            gv.AllowDeleteRow = False
            gv.AllowRowResize = False
            gv.AllowRowReorder = False
            gv.AllowColumnResize = True
            gv.AllowColumnChooser = False
            gv.AllowAutoSizeColumns = False
            gv.ShowGroupPanel = False
            gv.AddNewRowPosition = SystemRowPosition.Bottom
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.frmTankerMaster)
        If Not (MyBase.isReadFlag) Then
            common.clsCommon.MyMessageBoxShow(Me, "Permission Denied", Me.Text)
            Me.Close()
            Exit Sub
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btndelete.Visible = MyBase.isDeleteFlag
        If Not (MyBase.isReadFlag) Then
            btnexport.Visibility = ElementVisibility.Collapsed
        End If
        If Not (MyBase.isModifyFlag) Then
            btnimport.Visibility = ElementVisibility.Collapsed
        End If
    End Sub

    Private Sub btnnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnnew.Click
        Reset()
    End Sub

    Function AllowToSave() As Boolean
        Try
            If clsCommon.myLen(fndNo.Value) <= 0 Then
                fndNo.Focus()
                fndNo.Select()
                Errorcontrol.SetError(fndNo, "Please Fill Tanker No.")
                Throw New Exception("Please Fill Tanker No.")
            Else
                Errorcontrol.ResetError(fndNo)
            End If

            If AllowSameTankerNoInMaster = False Then
                If isDuplicateVechiele(fndNo.Value) AndAlso isNewEntry Then
                    fndNo.Focus()
                    fndNo.Select()
                    Errorcontrol.SetError(fndNo, "Duplicate Tanker No.")
                    Throw New Exception("Duplicate Tanker No.")
                Else
                    Errorcontrol.ResetError(fndNo)
                End If
            End If
           

            'If clsCommon.myLen(txtdesc.Text) <= 0 Then
            '    txtdesc.Focus()
            '    txtdesc.Select()
            '    Errorcontrol.SetError(txtdesc, "Please fill description")
            '    Throw New Exception("Please fill description")
            'Else
            '    Errorcontrol.ResetError(txtdesc)
            'End If

            If clsCommon.myLen(txttank_transcode.Value) <= 0 Then
                txttank_transcode.Focus()
                txttank_transcode.Select()
                Errorcontrol.SetError(txtname, "Please Select Tanker Transporter No.")
                Throw New Exception("Please Select Tanker Transporter No.")
            Else
                Errorcontrol.ResetError(txtname)
            End If
            ''richa Against Ticket No. BM00000003713 on 03/09/2014

            If clsCommon.myCdbl(txtstorage.Value) <= 0 Then
                txtstorage.Focus()
                txtstorage.Select()
                Errorcontrol.SetError(txtstorage, "Please Enter storage capacity")
                Throw New Exception("Please Enter storage capacity")
            Else
                Errorcontrol.ResetError(txtstorage)
            End If

            If ChkPrivate.Checked = False Then
                If clsCommon.myCdbl(TxtIceCharge.Value) <= 0 Then
                    TxtIceCharge.Focus()
                    TxtIceCharge.Select()
                    Errorcontrol.SetError(TxtIceCharge, "Please Enter Ice Charge")
                    Throw New Exception("Please Enter Ice Charge")
                Else
                    Errorcontrol.ResetError(TxtIceCharge)
                End If
            End If


            If clsCommon.myLen(ddlStorageCapacityDescription.Text) = 0 Then
                ddlStorageCapacityDescription.Focus()
                ddlStorageCapacityDescription.Select()
                Errorcontrol.SetError(ddlStorageCapacityDescription, "Please Select Storage Capacity Description")
                Throw New Exception("Please Enter Storage Capacity Description")
            Else
                Errorcontrol.ResetError(ddlStorageCapacityDescription)
            End If
            ''----------------------------------------------------
            If clsCommon.myLen(txtyear.Text) <= 0 Then
                txtyear.Focus()
                txtyear.Select()
                Errorcontrol.SetError(txtyear, "Please Fill Year of Manufacturing")
                Throw New Exception("Please Fill Year of Manufacturing")
            Else
                Errorcontrol.ResetError(txtyear)
            End If

            If rbtinner_no.IsChecked = False AndAlso rbtninner_yes.IsChecked = False Then
                rbtinner_no.IsChecked = True
            End If

            If rbtnouter_no.IsChecked = False AndAlso rbtnouter_yes.IsChecked = False Then
                rbtnouter_no.IsChecked = True
            End If

            If isImportClicked = False Then
                If Not rbtndiesel.Checked AndAlso Not rbtnratekm.Checked AndAlso Not rbtnrateltr.Checked AndAlso Not rbtnrental.Checked AndAlso Not rbtKmrange.Checked AndAlso Not chkMonthlyRentPlusDiesel.Checked Then
                    RadGroupBox4.Focus()
                    RadGroupBox4.Select()
                    Errorcontrol.SetError(RadGroupBox4, "Please Select Any One Rate Form Basic of Freight Payments")
                    Throw New Exception("Please Select Any One Rate Form Basic of Freight Payments")
                Else
                    Errorcontrol.ResetError(RadGroupBox4)
                End If

                If rbtndiesel.Checked AndAlso (clsCommon.myCdbl(txtchrg.Text) <= 0 Or clsCommon.myCdbl(txtavgkm.Text) <= 0 Or clsCommon.myCdbl(txtdiesel.Text) <= 0) Then
                    If clsCommon.myCdbl(txtchrg.Text) <= 0 Then
                        txtchrg.Focus()
                        txtchrg.Select()
                        Errorcontrol.SetError(txtchrg, "Please Fill Charges Per Day")
                    ElseIf clsCommon.myCdbl(txtavgkm.Text) <= 0 Then
                        txtavgkm.Focus()
                        txtavgkm.Select()
                        Errorcontrol.SetError(txtchrg, "Please Fill Average KM per Ltr.")
                    ElseIf clsCommon.myCdbl(txtdiesel.Text) <= 0 Then
                        txtdiesel.Focus()
                        txtdiesel.Select()
                        Errorcontrol.SetError(txtchrg, "Please Fill Rate of Diesel")
                    End If
                    Throw New Exception("Please Fill Rate Per Day + Diesel(Charges per Shift/Average KM per Ltr./Rate of Diesel)")
                Else
                    Errorcontrol.ResetError(txtchrg)
                    Errorcontrol.ResetError(txtavgkm)
                    Errorcontrol.ResetError(txtdiesel)
                End If

                If rbtnratekm.Checked AndAlso clsCommon.myCdbl(txt_km.Text) <= 0 Then
                    txt_km.Focus()
                    txt_km.Select()
                    Errorcontrol.SetError(txt_km, "Please Fill Rate per K.M")
                    Throw New Exception("Please Fill Rate per K.M")
                Else
                    Errorcontrol.ResetError(txt_km)
                End If

                'If rbtnrateltr.Checked AndAlso (clsCommon.myCdbl(txt_ltr.Text) <= 0 Or clsCommon.myCdbl(txt_ltr_kg.Text) <= 0) Then
                '    If clsCommon.myCdbl(txt_ltr.Text) <= 0 Then
                '        txt_ltr.Focus()
                '        txt_ltr.Select()
                '        Errorcontrol.SetError(txt_ltr, "Please Fill Rate per Ltr.")
                '    ElseIf clsCommon.myCdbl(txt_ltr_kg.Text) <= 0 Then
                '        txt_ltr_kg.Focus()
                '        txt_ltr_kg.Select()
                '        Errorcontrol.SetError(txt_ltr_kg, "Please Fill Ltr/KG")
                '    End If
                '    Throw New Exception("Please Fill Rate per Ltr./Kg And Ltr./Kg")
                'Else
                '    Errorcontrol.ResetError(txt_ltr)
                '    Errorcontrol.ResetError(txt_ltr_kg)
                'End If

                If rbtnrateltr.Checked AndAlso (clsCommon.myLen(txt_ltr.Text) <= 0 Or clsCommon.myCdbl(txt_ltr.Text) <= 0 Or clsCommon.myLen(cmbLtrKG.Text) <= 0) Then
                    If clsCommon.myCdbl(txt_ltr.Text) <= 0 Then
                        txt_ltr.Focus()
                        txt_ltr.Select()
                        Errorcontrol.SetError(txt_ltr, "Please Fill Rate per Ltr/KG")
                    ElseIf clsCommon.myLen(cmbLtrKG.Text) <= 0 Then
                        cmbLtrKG.Focus()
                        Errorcontrol.SetError(cmbLtrKG, "Please select  Ltr/KG")
                    End If
                    Throw New Exception("Please Fill Rate per Ltr./Kg And  select Ltr./Kg")
                Else
                    Errorcontrol.ResetError(txt_ltr)
                    Errorcontrol.ResetError(cmbLtrKG)
                End If

                If rbtnrental.Checked AndAlso clsCommon.myLen(cmbRentalType.Text) <= 0 Then
                    'If clsCommon.myCdbl(txtrental_day.Text) <= 0 Then
                    'AndAlso (clsCommon.myCdbl(txtrental_day.Text) <= 0 AndAlso clsCommon.myCdbl(txtrental_month.Text) <= 0 AndAlso clsCommon.myCdbl(txtrental_week.Text) <= 0)                    '    txtrental_day.Focus()

                    '    txtrental_day.Select()
                    '    Errorcontrol.SetError(txtrental_day, "Please Fill Rental per Day")
                    'ElseIf clsCommon.myCdbl(txtrental_week.Text) <= 0 Then
                    '    txtrental_week.Focus()
                    '    txtrental_week.Select()
                    '    Errorcontrol.SetError(txtrental_week, "Please Fill Rental per Week")
                    'ElseIf clsCommon.myCdbl(txtrental_month.Text) <= 0 Then
                    'txtrental_day.Focus()
                    'txtrental_day.Select()
                    Errorcontrol.SetError(cmbRentalType, "Please Fill any one from rental per Day/ Month/ Year")
                    'End If
                    Throw New Exception("Please Select Any rental type as Day/Month/Year")
                Else
                    Errorcontrol.ResetError(cmbRentalType)
                    'Errorcontrol.ResetError(txtrental_month)
                    'Errorcontrol.ResetError(txtrental_week)
                End If

                If rbtnrental.Checked AndAlso clsCommon.myLen(cmbRentalType.Text) > 0 AndAlso clsCommon.myCdbl(txtRentalAmt.Text) <= 0 Then
                    'If clsCommon.myCdbl(txtrental_day.Text) <= 0 Then
                    'AndAlso (clsCommon.myCdbl(txtrental_day.Text) <= 0 AndAlso clsCommon.myCdbl(txtrental_month.Text) <= 0 AndAlso clsCommon.myCdbl(txtrental_week.Text) <= 0)                    '    txtrental_day.Focus()

                    '    txtrental_day.Select()
                    '    Errorcontrol.SetError(txtrental_day, "Please Fill Rental per Day")
                    'ElseIf clsCommon.myCdbl(txtrental_week.Text) <= 0 Then
                    '    txtrental_week.Focus()
                    '    txtrental_week.Select()
                    '    Errorcontrol.SetError(txtrental_week, "Please Fill Rental per Week")
                    'ElseIf clsCommon.myCdbl(txtrental_month.Text) <= 0 Then
                    'txtrental_day.Focus()
                    'txtrental_day.Select()
                    Errorcontrol.SetError(txtRentalAmt, "Please Fill Rental amount")
                    'End If
                    Throw New Exception("Please Fill Rental amount")
                Else
                    Errorcontrol.ResetError(txtRentalAmt)
                    'Errorcontrol.ResetError(txtrental_month)
                    'Errorcontrol.ResetError(txtrental_week)
                End If
            End If

            If rbtKmrange.Checked AndAlso isBlankGV() Then
                Throw New Exception("Please Fill Slab Value")
            End If
            If rbtKmrange.Checked AndAlso (Not isBlankGV()) AndAlso (Not isValidateGridValue()) Then
                Throw New Exception("Slab Range Value Should be in increasing Order")
            End If
            If IsChamberWiseTanker Then
                If txtChamborNo.Value <= 0 OrElse gvChamber.Rows.Count <= 0 Then
                    RadPageView1.SelectedPage = RadPageViewPage3
                    Throw New Exception("Please enter Chamber Details")
                End If
                If txtChamborNo.Value > 10 OrElse gvChamber.Rows.Count > 10 Then
                    RadPageView1.SelectedPage = RadPageViewPage3
                    Throw New Exception("Chamber can't be more than 10")
                End If
                If txtChamborNo.Value <> gvChamber.Rows.Count Then
                    RadPageView1.SelectedPage = RadPageViewPage3
                    Throw New Exception("Please fill full chamber details.No of chamber-" + clsCommon.myCstr(txtChamborNo.Value) + " but chember description-" + clsCommon.myCstr(gvChamber.Rows.Count))
                End If
                If txtChamborNo.Value > 0 Then
                    For ii As Integer = 0 To gvChamber.RowCount - 1
                        If clsCommon.myLen(gvChamber.Rows(ii).Cells(colChamberDescription).Value) <= 0 Then
                            RadPageView1.SelectedPage = RadPageViewPage3
                            Throw New Exception("Please provide chamber Description at row no " + clsCommon.myCstr(ii + 1))
                        End If
                    Next
                End If
            End If
            Return True
        Catch ex As Exception
            If isImportClicked Then
                Throw New Exception(ex.Message)
            Else
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End If
        End Try
    End Function

    Sub SaveData()
        Try

            If MyBase.isModifyonPasswordFlag Then
                If clsPasswordCheckForMasters.CheckMasterPwd(clsUserMgtCode.frmTankerMaster, clsCommon.myCstr(objCommonVar.CurrentCompanyCode)) Then
                Else
                    Return
                End If
            End If

            Dim obj As New clsfrmTankerMaster()
            obj.code = clsCommon.myCstr(txttank_transcode.Value)
            obj.ESIRule = clsCommon.myCstr(TxtESIRule.Value)
            obj.EPFRule = clsCommon.myCstr(TxtEPFRule.Value)
            obj.desc = clsCommon.myCstr(txtname.Text)
            obj.tankerno = clsCommon.myCstr(fndNo.Value)
            obj.storagecap = clsCommon.myCdbl(txtstorage.Value)
            If ChkPrivate.Checked = True Then
                obj.IceCharge = 0
            Else
                obj.IceCharge = clsCommon.myCdbl(TxtIceCharge.Value)
            End If

            obj.year = clsCommon.myCstr(txtyear.Text)
            obj.tanker_name = clsCommon.myCstr(txtdesc.Text.Replace("'", "`"))
            ''richa Against Ticket No. BM00000003713 on 03/09/2014
            obj.StorageCapacityDesc = clsCommon.myCstr(ddlStorageCapacityDescription.Text)
            ''----------------------------------------------------
            If rbtinner_no.IsChecked Then
                obj.inner = "NO"
            ElseIf rbtninner_yes.IsChecked Then
                obj.inner = "YES"
            End If
            If rbtnouter_no.IsChecked Then
                obj.outer = "NO"
            ElseIf rbtnouter_yes.IsChecked Then
                obj.outer = "YES"
            End If
            obj.shift_chrg = clsCommon.myCdbl(txtchrg.Text)
            obj.avg_km_rate = clsCommon.myCdbl(txtavgkm.Text)
            obj.diesel_rate = clsCommon.myCdbl(txtdiesel.Text)
            obj.rate_km = clsCommon.myCdbl(txt_km.Text)
            obj.Rate_Type = clsCommon.myCstr(cmbLtrKG.Text)
            obj.Price_Ltr_KG = clsCommon.myCdbl(txt_ltr.Text)
            obj.Labour_Charge = clsCommon.myCdbl(NLabourCharge.Value)
            obj.RentalType = clsCommon.myCstr(cmbRentalType.Text)
            obj.RentalAmount = clsCommon.myCdbl(txtRentalAmt.Text)

            obj.Provision_Min_Qty = txtProvMinQty.Value

            If rbtKmrange.Checked Then
                obj.Status = "KM_Range"
            ElseIf rbtnratekm.Checked Then
                obj.Status = "Rate/K.M"
            ElseIf rbtndiesel.Checked Then
                obj.Status = "Day/Diesel"
            ElseIf rbtnrental.Checked Then
                obj.Status = "Rental"
            ElseIf rbtnrateltr.Checked Then
                obj.Status = "Rate/Ltr"
            ElseIf chkMonthlyRentPlusDiesel.Checked Then
                obj.Status = "Rental/Diesel"
                obj.RentalAmount = txtMRPDRent.Value
                obj.avg_km_rate = txtMRPDAverage.Value
                obj.diesel_rate = txtMRPDDieselRate.Value
            End If
            If IsChamberWiseTanker Then
                obj.Total_Chamber = txtChamborNo.Value
                If txtChamborNo.Value > 0 Then
                    obj.arrChamber = New List(Of clsTankerChamberDetail)
                    For ii As Integer = 0 To gvChamber.Rows.Count - 1
                        Dim objtr As New clsTankerChamberDetail
                        objtr.Chamber_No = clsCommon.myCdbl(gvChamber.Rows(ii).Cells(colChamberNo).Value)
                        objtr.Chamber_Description = clsCommon.myCstr(gvChamber.Rows(ii).Cells(colChamberDescription).Value)
                        obj.arrChamber.Add(objtr)
                    Next
                End If
            End If
            obj.Inactive = chkInactive.Checked
            obj.PrivateTanker = ChkPrivate.Checked
            Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
            Try
                If clsfrmTankerMaster.SaveData(obj.tankerno, isNewEntry, obj, trans) Then
                    Dim objSlab As clsSlabRangeDetail = Nothing
                    Dim arrObjSlab As List(Of clsSlabRangeDetail) = Nothing
                    clsSlabRangeDetail.deleteData(Me.Form_ID, obj.tankerno, trans)
                    If rbtKmrange.Checked Then
                        arrObjSlab = New List(Of clsSlabRangeDetail)
                        For i As Integer = 0 To gv.Rows.Count - 1
                            If clsCommon.myCdbl(gv.Rows(i).Cells(colSlabUpto).Value) > 0 AndAlso clsCommon.myCdbl(gv.Rows(i).Cells(colSlabRate).Value) > 0 Then
                                objSlab = New clsSlabRangeDetail()
                                objSlab.Form_ID = Me.Form_ID
                                objSlab.Trans_ID = obj.tankerno
                                objSlab.Slab_Upto = clsCommon.myCdbl(gv.Rows(i).Cells(colSlabUpto).Value)
                                objSlab.Slab_Rate = clsCommon.myCdbl(gv.Rows(i).Cells(colSlabRate).Value)
                                arrObjSlab.Add(objSlab)
                            End If
                        Next
                    End If

                    If arrObjSlab IsNot Nothing AndAlso arrObjSlab.Count > 0 Then
                        clsSlabRangeDetail.SaveData(arrObjSlab, trans)
                    End If
                End If
                trans.Commit()
                If isImportClicked = False Then
                    If clsCommon.CompairString(btnsave.Text, "Save") = CompairStringResult.Equal Then
                        clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                    Else
                        clsCommon.MyMessageBoxShow(Me, "Data Updated Successfully", Me.Text)
                    End If
                    btnsave.Text = "Update"
                    btndelete.Enabled = True
                    fndNo.MyReadOnly = True
                    txtTankerNo.ReadOnly = True
                    UcAttachment1.SaveData(fndNo.Value)
                End If
            Catch ex As Exception
                trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        Catch ex As Exception
            btnsave.Text = "Save"
            btndelete.Enabled = False
            If isImportClicked Then
                Throw New Exception(ex.Message)
            Else
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End If

        End Try
    End Sub

    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        If AllowToSave() Then SaveData()
    End Sub

    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(fndNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please Select Tanker No. For Deletion", Me.Text)
                fndNo.Focus()
                fndNo.Select()
                trans.Commit()
                Return
            End If

            Dim qry As String = "select count(*) from TSPL_TANKER_MASTER where Tanker_No='" + fndNo.Value + "'"
            Dim check As Integer = clsDBFuncationality.getSingleValue(qry, trans)

            If check <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found For Deletion", Me.Text)
                trans.Commit()
                Return
            End If

            If Not clsCommon.MyMessageBoxShow("Are You Sure,Want To Delete Tanker Master Of Tanker No. " + fndNo.Value + " And Tanker Transporter No. " + txttank_transcode.Value + "?", "Attention", MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                trans.Commit()
                Return
            End If
            clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, fndNo.Value, "TSPL_TANKER_MASTER", "Tanker_No", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, fndNo.Value, "TSPL_TANKER_MASTER", "Tanker_No", trans)

            clsDBFuncationality.ExecuteNonQuery("delete from TSPL_TANKER_CHAMBER_DETAIL where Tanker_No='" + fndNo.Value + "'", trans)

            qry = "Delete from TSPL_TANKER_MASTER where Tanker_No='" + fndNo.Value + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
            clsSlabRangeDetail.deleteData(Me.Form_ID, fndNo.Value, trans)

            trans.Commit()

            clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully", Me.Text)
            Reset()
        Catch ex As Exception
            trans.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Me.Close()
    End Sub

    Private Sub btnexport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexport.Click
        Dim qry As String = "select count(*) from tspl_tanker_master"
        Dim check As Integer = clsDBFuncationality.getSingleValue(qry)

        If check > 0 Then
            qry = "select TSPL_TANKER_MASTER.Tanker_No,TSPL_TANKER_MASTER.Tanker_Name,TSPL_VENDOR_MASTER.Vendor_Code as [Tanker Transporter Code],TSPL_VENDOR_MASTER.Vendor_Name as [Tanker Transporter Name],TSPL_TANKER_MASTER.Storage_Capacity,TSPL_TANKER_MASTER.StorageCapacityDesc,TSPL_TANKER_MASTER.Year  as [Year of Manufacturing],TSPL_TANKER_MASTER.Inner_SS,TSPL_TANKER_MASTER.Outer_SS,Shift_Charges,Avg_KM_Ltr,Diesel_Rate,Price_KM,Rate_Type,Price_Ltr_Kg,Rental_Type,Rental_Amount,Provision_Min_Qty "
            If IsChamberWiseTanker Then
                qry += " ,TSPL_TANKER_MASTER.Total_Chamber,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='1') as Chamber_1 " + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='2') as Chamber_2" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='3') as Chamber_3" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='4') as Chamber_4" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='5') as Chamber_5" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='6') as Chamber_6" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='7') as Chamber_7" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='8') as Chamber_8" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='9') as Chamber_9" + Environment.NewLine +
                " ,(select Chamber_Description from TSPL_TANKER_CHAMBER_DETAIL where TSPL_TANKER_CHAMBER_DETAIL.Tanker_No=TSPL_TANKER_MASTER.Tanker_No and Chamber_No='10') as Chamber_10 "
            End If


            qry += " from TSPL_VENDOR_MASTER right outer join TSPL_TANKER_MASTER on TSPL_TANKER_MASTER.Tanker_Transporter_Code=TSPL_VENDOR_MASTER.Vendor_Code and TSPL_VENDOR_MASTER.form_type='TTM'"
        Else
            qry = "select '' as Tanker_No,'' as Tanker_Name,'' as [Tanker Transporter Code],'' as [Tanker Transporter Name],0 as Storage_Capacity,'' as StorageCapacityDesc,'' as [Year of Manufacturing],'' as Inner_SS,'' as Outer_SS,0 as Shift_Charges,0 as Avg_KM_Ltr,0 as Diesel_Rate,0 as Price_KM,'' as Rate_Type,0 as Price_Ltr_Kg,'' as Rental_Type,0 as Rental_Amount,0 as Provision_Min_Qty  "
            If IsChamberWiseTanker Then
                qry += " ,0 as Total_Chamber,'' as Chamber_1,'' as Chamber_2,'' as Chamber_3,'' as Chamber_4,'' as Chamber_5,'' as Chamber_6,'' as Chamber_7,'' as Chamber_8,'' as Chamber_9,'' as Chamber_10"
            End If
        End If
        ListImpExpColumnsMandatory = New List(Of String)({"Tanker_No", "Tanker Transporter Code", "Storage_Capacity", "StorageCapacityDesc", "Year of Manufacturing", "Shift_Charges", "Price_KM", "Rate_Type"})
        ListImpExpColumnsSuperMandatory = New List(Of String)({"Tanker_No"})
        transportSql.ExporttoExcel(qry, "", "", Me, ListImpExpColumnsMandatory, ListImpExpColumnsSuperMandatory, MyBase.Form_ID)
    End Sub

    Private Sub btnimport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnimport.Click
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim currentdate As Date = Date.Today
        Dim result As Boolean = False
        If IsChamberWiseTanker Then
            result = transportSql.importExcel(gv, "Tanker_No", "Tanker_Name", "Tanker Transporter Code", "Tanker Transporter Name", "Storage_Capacity", "StorageCapacityDesc", "Year of Manufacturing", "Inner_SS", "Outer_SS", "Shift_Charges", "Avg_KM_Ltr", "Diesel_Rate", "Price_KM", "Rate_Type", "Price_Ltr_Kg", "Rental_Type", "Rental_Amount", "Provision_Min_Qty", "Total_Chamber", "Chamber_1", "Chamber_2", "Chamber_3", "Chamber_4", "Chamber_5", "Chamber_6", "Chamber_7", "Chamber_8", "Chamber_9", "Chamber_10")
        Else
            result = transportSql.importExcel(gv, "Tanker_No", "Tanker_Name", "Tanker Transporter Code", "Tanker Transporter Name", "Storage_Capacity", "StorageCapacityDesc", "Year of Manufacturing", "Inner_SS", "Outer_SS", "Shift_Charges", "Avg_KM_Ltr", "Diesel_Rate", "Price_KM", "Rate_Type", "Price_Ltr_Kg", "Rental_Type", "Rental_Amount", "Provision_Min_Qty")
        End If
        If result Then
            Try
                clsCommon.ProgressBarShow()
                Dim counter As Integer = 1
                For Each grow As GridViewRowInfo In gv.Rows
                    isImportClicked = True
                    txttank_transcode.Value = clsCommon.myCstr(grow.Cells("Tanker Transporter Code").Value)
                    txtname.Text = clsCommon.myCstr(grow.Cells("Tanker Transporter Name").Value)

                    Dim qry As String = "select count(*) from tspl_vendor_master where vendor_code='" + txttank_transcode.Value + "' and form_type='TTM'"
                    Dim check As Integer = clsDBFuncationality.getSingleValue(qry)

                    If check <= 0 Then
                        Throw New Exception("Please First Make The Tanker Transporter Of No. " + txttank_transcode.Value + " See At Line No. " + clsCommon.myCstr(counter) + "")
                    End If

                    fndNo.Value = clsCommon.myCstr(grow.Cells("Tanker_No").Value)
                    txtdesc.Text = clsCommon.myCstr(grow.Cells("Tanker_Name").Value)

                    If clsCommon.myLen(txtdesc.Text) > 150 Then
                        Throw New Exception("Tanker Name Should Not Exceed Max.150 Characters,See At Line No. " + clsCommon.myCstr(counter) + "")
                    End If
                    ''richa Against Ticket No. BM00000003713 on 03/09/2014
                    If clsCommon.myCdbl(grow.Cells("Storage_Capacity").Value) = 0 Then
                        Throw New Exception("Storage Capacity cannot be left blank or 0,See At Line No. " + clsCommon.myCstr(counter) + "")
                    End If
                    '-------------------------
                    txtProvMinQty.Value = clsCommon.myCdbl(grow.Cells("Provision_Min_Qty").Value)
                    txtstorage.Value = clsCommon.myCdbl(grow.Cells("Storage_Capacity").Value)
                    TxtIceCharge.Value = clsCommon.myCdbl(grow.Cells("Ice_Charge").Value)
                    txtyear.Text = clsCommon.myCstr(grow.Cells("Year of Manufacturing").Value)
                    ''richa Against Ticket No. BM00000003713 on 03/09/2014
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("StorageCapacityDesc").Value).ToUpper(), "LITRE") = CompairStringResult.Equal Then
                        ddlStorageCapacityDescription.Text = "Litre"
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(grow.Cells("StorageCapacityDesc").Value).ToUpper(), "KG") = CompairStringResult.Equal Then
                        ddlStorageCapacityDescription.Text = "KG"
                    ElseIf clsCommon.myLen(grow.Cells("StorageCapacityDesc").Value) <= 0 Then
                        Throw New Exception("Storage Capacity Desc cannot be left blank At Line No. " + clsCommon.myCstr(counter) + "")
                    Else
                        Throw New Exception("Storage Capacity Desc should be Litre/ KG ,See At Line No. " + clsCommon.myCstr(counter) + "")
                    End If
                    ''----------------------------------------------------


                    If clsCommon.myLen(txtyear.Text) <= 0 Then
                        Throw New Exception("Please fill year of manufacturing See At Line No. " + clsCommon.myCstr(counter) + "")
                    End If

                    Try
                        If clsCommon.myLen(fndNo.Value) > 0 Then
                            Dim index As Integer = 0
                            index = clsCommon.myCstr(fndNo.Value).IndexOf(" ")

                            If index > 0 AndAlso index < clsCommon.myLen(fndNo.Value) Then
                                Throw New Exception("No White Space Allowed In Tanker No. See At Line No. " + clsCommon.myCstr(counter) + "")
                            End If
                        End If
                    Catch exx As Exception
                        Throw New Exception(exx.Message)
                    End Try
                    Try
                        If clsCommon.myLen(txtstorage.Value) > 0 Then
                            Convert.ToDecimal(txtstorage.Value)
                        End If
                    Catch exx As Exception
                        Throw New Exception("Storage Capacity Should Be In Numeric See At Line No. " + clsCommon.myCstr(counter) + "")
                    End Try
                    Try
                        If clsCommon.myLen(TxtIceCharge.Value) > 0 Then
                            Convert.ToDecimal(TxtIceCharge.Value)
                        End If
                    Catch exx As Exception
                        Throw New Exception("Ice Charge Should Be In Numeric See At Line No. " + clsCommon.myCstr(counter) + "")
                    End Try

                    rbtinner_no.IsChecked = False
                    rbtninner_yes.IsChecked = False
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Inner_SS").Value), "NO") = CompairStringResult.Equal Then
                        rbtinner_no.IsChecked = True
                        rbtninner_yes.IsChecked = False
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Inner_SS").Value), "YES") = CompairStringResult.Equal Then
                        rbtinner_no.IsChecked = False
                        rbtninner_yes.IsChecked = True
                    End If
                    rbtnouter_no.IsChecked = False
                    rbtnouter_yes.IsChecked = False
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Outer_SS").Value), "NO") = CompairStringResult.Equal Then
                        rbtnouter_no.IsChecked = True
                        rbtnouter_yes.IsChecked = False
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(grow.Cells("Outer_SS").Value), "YES") = CompairStringResult.Equal Then
                        rbtnouter_no.IsChecked = False
                        rbtnouter_yes.IsChecked = True
                    End If
                    txtchrg.Text = clsCommon.myCdbl(grow.Cells("Shift_Charges").Value)
                    txtavgkm.Text = clsCommon.myCdbl(grow.Cells("Avg_KM_Ltr").Value)
                    txtdiesel.Text = clsCommon.myCdbl(grow.Cells("Diesel_Rate").Value)
                    rbtndiesel.Checked = False
                    If clsCommon.myCdbl(txtchrg.Text) > 0 Or clsCommon.myCdbl(txtavgkm.Text) > 0 Or clsCommon.myCdbl(txtdiesel.Text) > 0 Then
                        rbtndiesel.Checked = True
                    End If
                    cmbRentalType.Text = clsCommon.myCstr(grow.Cells("Rental_Type").Value)
                    txtRentalAmt.Text = clsCommon.myCdbl(grow.Cells("Rental_Amount").Value)
                    rbtnrental.Checked = False
                    If clsCommon.myLen(cmbRentalType.Text) > 0 Or clsCommon.myCdbl(txtRentalAmt.Text) > 0 Then
                        rbtnrental.Checked = True
                    End If

                    txt_km.Text = clsCommon.myCdbl(grow.Cells("Price_KM").Value)
                    rbtnratekm.Checked = False
                    If clsCommon.myCdbl(txt_km.Text) > 0 Then
                        rbtnratekm.Checked = True
                    End If

                    cmbLtrKG.Text = clsCommon.myCstr(grow.Cells("Rate_Type").Value)
                    txt_ltr.Text = clsCommon.myCdbl(grow.Cells("Price_Ltr_KG").Value)
                    NLabourCharge.Value = clsCommon.myCdbl(grow.Cells("Labour_Charge").Value)
                    rbtnrateltr.Checked = False
                    If clsCommon.myLen(cmbLtrKG.Text) > 0 Or clsCommon.myCdbl(txt_ltr.Text) > 0 Then
                        rbtnrateltr.Checked = True
                    End If


                    chkMonthlyRentPlusDiesel.Checked = False
                    If clsCommon.myCdbl(grow.Cells("Rental_Amount").Value) > 0 AndAlso clsCommon.myCdbl(grow.Cells("Diesel_Rate").Value) > 0 AndAlso clsCommon.myCdbl(grow.Cells("Avg_KM_Ltr").Value) > 0 Then
                        chkMonthlyRentPlusDiesel.Checked = True
                        txtMRPDRent.Value = clsCommon.myCdbl(grow.Cells("Rental_Amount").Value)
                        txtMRPDDieselRate.Value = clsCommon.myCdbl(grow.Cells("Diesel_Rate").Value)
                        txtMRPDAverage.Value = clsCommon.myCdbl(grow.Cells("Avg_KM_Ltr").Value)

                        rbtndiesel.Checked = False
                        rbtnrental.Checked = False

                    End If
                    rbtKmrange.Checked = False
                    If clsCommon.myCdbl(grow.Cells("Shift_Charges").Value) <= 0 AndAlso clsCommon.myCdbl(grow.Cells("Avg_KM_Ltr").Value) <= 0 AndAlso clsCommon.myCdbl(grow.Cells("Diesel_Rate").Value) <= 0 AndAlso clsCommon.myCdbl(grow.Cells("Price_Ltr_Kg").Value) <= 0 AndAlso clsCommon.myCdbl(grow.Cells("Rental_Amount").Value) <= 0 AndAlso clsCommon.myCdbl(grow.Cells("Price_KM").Value) <= 0 Then
                        rbtKmrange.Checked = True
                    End If


                    If Not rbtndiesel.Checked AndAlso Not rbtnratekm.Checked AndAlso Not rbtnrateltr.Checked AndAlso Not rbtnrental.Checked AndAlso Not chkMonthlyRentPlusDiesel.Checked AndAlso Not rbtKmrange.Checked Then
                        Throw New Exception("Please Select Fill Any One Rate Form Basic of Freight Payments At Line No. " + clsCommon.myCstr(counter) + "")
                    End If

                    If rbtndiesel.Checked AndAlso (clsCommon.myCdbl(txtchrg.Text) <= 0 Or clsCommon.myCdbl(txtavgkm.Text) <= 0 Or clsCommon.myCdbl(txtdiesel.Text) <= 0) Then
                        Throw New Exception("Please Fill Rate Per Day + Diesel(Charges per Day/Average KM per Ltr./Rate of Diesel) At Line No. " + clsCommon.myCstr(counter) + "")
                    End If

                    If rbtnratekm.Checked AndAlso clsCommon.myCdbl(txt_km.Text) <= 0 Then
                        Throw New Exception("Please Fill Rate per K.M At Line No. " + clsCommon.myCstr(counter) + "")
                    End If

                    If rbtnrateltr.Checked AndAlso (clsCommon.myLen(cmbLtrKG.Text) <= 0 Or clsCommon.myCdbl(txt_ltr.Text) <= 0) Then
                        Throw New Exception("Please Fill Rate Type : Ltr./Kg And  Price Ltr/Kg At Line No. " + clsCommon.myCstr(counter) + "")
                    End If


                    LoadBlankGridChamber()
                    If IsChamberWiseTanker Then
                        txtChamborNo.Value = clsCommon.myCdbl(grow.Cells("Total_Chamber").Value)
                        If txtChamborNo.Value > 0 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_1").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 1 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_2").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 2 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_3").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 3 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_4").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 4 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_5").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 5 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_6").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 6 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_7").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 7 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_8").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 8 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_9").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                        If txtChamborNo.Value > 9 Then
                            gvChamber.Rows.AddNew()
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberNo).Value = gvChamber.RowCount
                            gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value = clsCommon.myCstr(grow.Cells("Chamber_10").Value)
                            If clsCommon.myLen(gvChamber.Rows(gvChamber.RowCount - 1).Cells(colChamberDescription).Value) <= 0 Then
                                Throw New Exception("Please provide chamber description " + clsCommon.myCstr(gvChamber.RowCount - 1) + " For tanker " + fndNo.Value)
                            End If
                        End If
                    End If
                    qry = "select count(*) from TSPL_TANKER_MASTER where Tanker_No='" + fndNo.Value + "'"
                    check = clsDBFuncationality.getSingleValue(qry)
                    If check > 0 Then
                        isNewEntry = False
                    Else
                        isNewEntry = True
                    End If

                    If AllowToSave() Then SaveData()

                    counter += 1
                Next

                isImportClicked = False
                clsCommon.ProgressBarHide()
                Reset()
                common.clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                isImportClicked = False
                Reset()
                clsCommon.ProgressBarHide()
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try
        End If
        Me.Controls.Remove(gv)
    End Sub

    Private Sub FrmTankerMaster_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            btnnew.PerformClick()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag AndAlso btnsave.Enabled Then
            btnsave.PerformClick()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btndelete.Enabled Then
            btndelete.PerformClick()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C Then
            Me.Close()
        End If
    End Sub

    

    Sub LoadData(ByVal strCode As String, ByVal navType As NavigatorType)
        Try
            Dim obj As clsfrmTankerMaster = clsfrmTankerMaster.GetData(strCode, navType)
            isNewEntry = True
            Reset()
            If obj IsNot Nothing AndAlso clsCommon.myLen(obj.tankerno) > 0 Then
                isNewEntry = False
                txtTankerNo.Text = obj.tankerno
                fndNo.Value = obj.tankerno
                txtdesc.Text = obj.tanker_name
                txtname.Text = obj.desc
                txttank_transcode.Value = obj.code
                TxtESIRule.Value = obj.ESIRule
                TxtEPFRule.Value = obj.EPFRule
                txtstorage.Value = obj.storagecap
                TxtIceCharge.Value = obj.IceCharge
                NLabourCharge.Value = obj.Labour_Charge
                ''richa Against Ticket No. BM00000003713 on 03/09/2014
                ddlStorageCapacityDescription.Text = obj.StorageCapacityDesc
                ''----------------------------------------------------
                txtyear.Text = obj.year
                If clsCommon.CompairString(obj.inner, "NO") = CompairStringResult.Equal Then
                    rbtinner_no.IsChecked = True
                    rbtninner_yes.IsChecked = False
                ElseIf clsCommon.CompairString(obj.inner, "YES") = CompairStringResult.Equal Then
                    rbtinner_no.IsChecked = False
                    rbtninner_yes.IsChecked = True
                End If

                If clsCommon.CompairString(obj.outer, "NO") = CompairStringResult.Equal Then
                    rbtnouter_no.IsChecked = True
                    rbtnouter_yes.IsChecked = False
                ElseIf clsCommon.CompairString(obj.outer, "YES") = CompairStringResult.Equal Then
                    rbtnouter_no.IsChecked = False
                    rbtnouter_yes.IsChecked = True
                End If
                txtProvMinQty.Value = obj.Provision_Min_Qty

                txtchrg.Text = obj.shift_chrg
                txtavgkm.Text = obj.avg_km_rate
                txtdiesel.Text = obj.diesel_rate
                cmbRentalType.Text = clsCommon.myCstr(obj.RentalType)
                txtRentalAmt.Text = clsCommon.myCdbl(obj.RentalAmount)
                txt_km.Text = obj.rate_km
                cmbLtrKG.Text = obj.Rate_Type
                txt_ltr.Text = obj.Price_Ltr_KG

                If clsCommon.CompairString(obj.Status, "KM_Range") = CompairStringResult.Equal Then
                    rbtKmrange.Checked = True
                ElseIf clsCommon.CompairString(obj.Status, "Rate/K.M") = CompairStringResult.Equal Then
                    rbtnratekm.Checked = True
                ElseIf clsCommon.CompairString(obj.Status, "Day/Diesel") = CompairStringResult.Equal Then
                    rbtndiesel.Checked = True
                ElseIf clsCommon.CompairString(obj.Status, "Rental") = CompairStringResult.Equal Then
                    rbtnrental.Checked = True
                ElseIf clsCommon.CompairString(obj.Status, "Rate/Ltr") = CompairStringResult.Equal Then
                    rbtnrateltr.Checked = True
                ElseIf clsCommon.CompairString(obj.Status, "Rental/Diesel") = CompairStringResult.Equal Then
                    chkMonthlyRentPlusDiesel.Checked = True
                    txtMRPDRent.Value = obj.RentalAmount
                    txtMRPDAverage.Value = obj.avg_km_rate
                    txtMRPDDieselRate.Value = obj.diesel_rate

                    txtdiesel.Text = ""
                    txtavgkm.Text = ""
                    txtRentalAmt.Text = ""
                End If
                chkInactive.Checked = obj.Inactive
                ChkPrivate.Checked = obj.PrivateTanker

                Dim kmRange As Boolean = False
                Dim arrObjSlab As List(Of clsSlabRangeDetail) = clsSlabRangeDetail.getData(Me.Form_ID, obj.tankerno)
                If arrObjSlab IsNot Nothing AndAlso arrObjSlab.Count > 0 Then
                    gv.Rows.Clear()
                    For i As Integer = 0 To arrObjSlab.Count - 1
                        gv.Rows.AddNew()
                        gv.Rows(i).Cells(colSlabUpto).Value = arrObjSlab.Item(i).Slab_Upto
                        gv.Rows(i).Cells(colSlabRate).Value = arrObjSlab.Item(i).Slab_Rate
                        kmRange = True
                    Next
                End If
                txtChamborNo.Value = obj.Total_Chamber
                If obj.arrChamber IsNot Nothing AndAlso obj.arrChamber.Count > 0 Then
                    gvChamber.Rows.Clear()
                    For i As Integer = 0 To obj.arrChamber.Count - 1
                        gvChamber.Rows.AddNew()
                        gvChamber.Rows(i).Cells(colChamberNo).Value = obj.arrChamber(i).Chamber_No
                        gvChamber.Rows(i).Cells(colChamberDescription).Value = obj.arrChamber(i).Chamber_Description
                    Next
                End If
                rbtKmrange.Checked = kmRange
                UcAttachment1.LoadData(fndNo.Value)
            Else
                Reset()
            End If

            If clsCommon.myLen(fndNo.Value) > 0 Then
                btnsave.Text = "Update"
                btndelete.Enabled = True
                fndNo.MyReadOnly = True
                txtTankerNo.ReadOnly = True
                UcAttachment1.SaveData(fndNo.Value)
            Else
                btnsave.Text = "Save"
                btndelete.Enabled = False
                txtTankerNo.ReadOnly = False
                fndNo.MyReadOnly = False
            End If
        Catch ex As Exception
            isNewEntry = True
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub fndNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles fndNo._MYNavigator
        LoadData(fndNo.Value, NavType)
    End Sub

    Private Sub fndNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles fndNo._MYValidating
        Dim qry As String = "select count(*) from tspl_tanker_master where tanker_no='" + fndNo.Value + "'"
        Dim check As Integer = clsDBFuncationality.getSingleValue(qry)

        If check > 0 Then
            fndNo.MyReadOnly = True
            txtTankerNo.ReadOnly = True
        Else
            fndNo.MyReadOnly = False
            txtTankerNo.ReadOnly = False
        End If

        If fndNo.MyReadOnly OrElse isButtonClicked Then
            txtTankerNo.Text = clsfrmTankerMaster.GetFinder(" tspl_vendor_master.Form_type='TTM'", txtTankerNo.Text, isButtonClicked)

            LoadData(fndNo.Value, NavigatorType.Current)
        End If
    End Sub

    Private Sub txtstorage_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If isImportClicked Then
            Return
        End If

        Try
            If clsCommon.myLen(txtstorage.Value) > 0 Then
                Convert.ToDecimal(txtstorage.Value)
            End If
            Errorcontrol.ResetError(txtstorage)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, "Storage Capacity Should Be In Numeric", Me.Text)
            txtstorage.Value = 0
            txtstorage.Focus()
            txtstorage.Select()
            Errorcontrol.SetError(txtstorage, "Storage Capacity Should Be In Numeric")
        End Try
    End Sub

    Private Sub txttank_transcode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txttank_transcode._MYValidating
        ' Ticket No : BHA/21/06/18-000073 By Prabhakar
        txttank_transcode.Value = clsVendorMaster.getFinder(" TSPL_VENDOR_MASTER.Form_Type='TTM' and TSPL_VENDOR_MASTER.isBulkProcurement=0 ", txttank_transcode.Value, isButtonClicked)

        If clsCommon.myLen(txttank_transcode.Value) > 0 Then
            txtname.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select vendor_name  from tspl_vendor_master where form_type='TTM' and vendor_code='" & txttank_transcode.Value & "'"))
        Else
            txtname.Text = ""
        End If
    End Sub

    Private Sub fndNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles fndNo.KeyPress
        'If (Asc(e.KeyChar) >= 48 AndAlso Asc(e.KeyChar) <= 57) OrElse (Asc(e.KeyChar) >= 65 AndAlso Asc(e.KeyChar) <= 90) OrElse (Asc(e.KeyChar) >= 97 AndAlso Asc(e.KeyChar) <= 122) OrElse Asc(e.KeyChar) = Keys.Back OrElse Asc(e.KeyChar) = Keys.Delete Then
        'Else
        '    e.Handled = True
        'End If
    End Sub

    Private Sub fndNo_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles fndNo.Validating
        If isImportClicked Then
            Return
        End If

        Try
            If clsCommon.myLen(fndNo.Value) > 0 Then
                Dim index As Integer = 0
                index = clsCommon.myCstr(fndNo.Value).IndexOf(" ")

                If index > 0 AndAlso index < clsCommon.myLen(fndNo.Value) Then
                    fndNo.Focus()
                    fndNo.Select()
                    Errorcontrol.SetError(fndNo, "No White Space Allowed In Tanker No.")
                    Throw New Exception("No White Space Allowed In Tanker No.")
                Else
                    Errorcontrol.ResetError(fndNo)
                End If
            End If
        Catch ex As Exception
            fndNo.Focus()
            fndNo.Select()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtndiesel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtndiesel.Click
        rbtndiesel.IsChecked = False
        rbtnratekm.IsChecked = False
        rbtnrateltr.IsChecked = False
        rbtnrental.IsChecked = False
        rbtKmrange.IsChecked = False
        chkMonthlyRentPlusDiesel.IsChecked = False
    End Sub

    Private Sub rbtndiesel_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtndiesel.ToggleStateChanged
        If isImportClicked = True Then
            Return
        End If
        'rbtndiesel.Checked = Not rbtndiesel.Checked
        'ResetCheckBox()
        If rbtndiesel.IsChecked = True Then
            txtchrg.Enabled = True
            txtavgkm.Enabled = True
            txtdiesel.Enabled = True
            txtchrg.MendatroryField = True
            txtavgkm.MendatroryField = True
            txtdiesel.MendatroryField = True
        Else
            txtchrg.Text = Nothing
            txtavgkm.Text = Nothing
            txtdiesel.Text = Nothing
            txtchrg.Enabled = False
            txtavgkm.Enabled = False
            txtdiesel.Enabled = False
            txtchrg.MendatroryField = False
            txtavgkm.MendatroryField = False
            txtdiesel.MendatroryField = False
        End If
    End Sub

    Private Sub rbtnrental_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnrental.Click
        rbtndiesel.IsChecked = False
        rbtnratekm.IsChecked = False
        rbtnrateltr.IsChecked = False
        rbtnrental.IsChecked = False
        rbtKmrange.IsChecked = False
        chkMonthlyRentPlusDiesel.IsChecked = False
    End Sub

    Private Sub rbtnrental_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtnrental.ToggleStateChanged
        'rbtnrental.Checked = Not rbtnrental.Checked
        'ResetCheckBox()
        If isImportClicked = True Then
            Return
        End If
        If rbtnrental.Checked = True Then
            cmbRentalType.Enabled = True
            'cmbRentalType.SelectedIndex = 0
            cmbRentalType.MendatroryField = True
            txtRentalAmt.Enabled = True
            '            txtRentalAmt.Text = "0"
            txtRentalAmt.MendatroryField = True
            'txtrental_day.Enabled = True
            'txtrental_month.Enabled = True
            'txtrental_week.Enabled = True
            'txtrental_day.MendatroryField = True
            'txtrental_month.MendatroryField = True
            'txtrental_week.MendatroryField = True
        Else
            cmbRentalType.Enabled = False
            cmbRentalType.Text = ""
            cmbRentalType.MendatroryField = False
            txtRentalAmt.Enabled = False
            txtRentalAmt.Text = Nothing
            txtRentalAmt.MendatroryField = False

            'txtrental_day.Text = Nothing
            'txtrental_month.Text = Nothing
            'txtrental_week.Text = Nothing
            'txtrental_day.Enabled = False
            'txtrental_month.Enabled = False
            'txtrental_week.Enabled = False
            'txtrental_day.MendatroryField = False
            'txtrental_month.MendatroryField = False
            'txtrental_week.MendatroryField = False
        End If
    End Sub

    Private Sub rbtnrateltr_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnrateltr.Click
        rbtndiesel.IsChecked = False
        rbtnratekm.IsChecked = False
        rbtnrateltr.IsChecked = False
        rbtnrental.IsChecked = False
        rbtKmrange.IsChecked = False
        chkMonthlyRentPlusDiesel.IsChecked = False
    End Sub

    Private Sub rbtnrateltr_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtnrateltr.ToggleStateChanged
        'rbtnrateltr.Checked = Not rbtnrateltr.Checked
        'ResetCheckBox()
        If isImportClicked = True Then
            Return
        End If
        'If rbtnrateltr.Checked = True Then
        '    txt_ltr.Enabled = True
        '    txt_ltr_kg.Enabled = True
        '    txt_ltr.MendatroryField = True
        '    txt_ltr_kg.MendatroryField = True
        'Else
        '    txt_ltr.Text = Nothing
        '    txt_ltr_kg.Text = Nothing
        '    txt_ltr.Enabled = False
        '    txt_ltr_kg.Enabled = False
        '    txt_ltr.MendatroryField = False
        '    txt_ltr_kg.MendatroryField = False
        'End If
        If rbtnrateltr.Checked = True Then
            cmbLtrKG.Enabled = True
            'cmbRentalType.SelectedIndex = 0
            cmbLtrKG.MendatroryField = True
            txt_ltr.Enabled = True
            '            txtRentalAmt.Text = "0"
            txt_ltr.MendatroryField = True
            'txtrental_day.Enabled = True
            'txtrental_month.Enabled = True
            'txtrental_week.Enabled = True
            'txtrental_day.MendatroryField = True
            'txtrental_month.MendatroryField = True
            'txtrental_week.MendatroryField = True
        Else
            cmbLtrKG.Enabled = False
            cmbLtrKG.Text = ""
            cmbLtrKG.MendatroryField = False
            txt_ltr.Enabled = False
            txt_ltr.Text = Nothing
            txt_ltr.MendatroryField = False

            'txtrental_day.Text = Nothing
            'txtrental_month.Text = Nothing
            'txtrental_week.Text = Nothing
            'txtrental_day.Enabled = False
            'txtrental_month.Enabled = False
            'txtrental_week.Enabled = False
            'txtrental_day.MendatroryField = False
            'txtrental_month.MendatroryField = False
            'txtrental_week.MendatroryField = False
        End If
    End Sub

    Private Sub rbtnratekm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtnratekm.Click
        rbtndiesel.IsChecked = False
        rbtnratekm.IsChecked = False
        rbtnrateltr.IsChecked = False
        rbtnrental.IsChecked = False
        rbtKmrange.IsChecked = False
        chkMonthlyRentPlusDiesel.IsChecked = False
    End Sub

    Private Sub rbtnratekm_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtnratekm.ToggleStateChanged
        'rbtnratekm.Checked = Not rbtnratekm.Checked
        'ResetCheckBox()
        If isImportClicked = True Then
            Return
        End If
        If rbtnratekm.Checked = True Then
            txt_km.Enabled = True
            txt_km.MendatroryField = True
        Else
            txt_km.Enabled = False
            txt_km.MendatroryField = False
            txt_km.Text = Nothing
        End If
    End Sub

    Private Sub txtTankerNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTankerNo.KeyPress
        If (Asc(e.KeyChar) >= 48 AndAlso Asc(e.KeyChar) <= 57) OrElse (Asc(e.KeyChar) >= 65 AndAlso Asc(e.KeyChar) <= 90) OrElse (Asc(e.KeyChar) >= 97 AndAlso Asc(e.KeyChar) <= 122) OrElse Asc(e.KeyChar) = Keys.Back OrElse Asc(e.KeyChar) = Keys.Delete Then
        Else
            e.Handled = True
        End If
    End Sub


    Private Sub txtTankerNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTankerNo.TextChanged
        fndNo.Value = txtTankerNo.Text
    End Sub

    Private Sub rbtKmrange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbtKmrange.Click
        rbtndiesel.IsChecked = False
        rbtnratekm.IsChecked = False
        rbtnrateltr.IsChecked = False
        rbtnrental.IsChecked = False
        rbtKmrange.IsChecked = False
        chkMonthlyRentPlusDiesel.IsChecked = False
    End Sub

    Private Sub rbtKmrange_ToggleStateChanged(ByVal sender As Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles rbtKmrange.ToggleStateChanged
        If rbtKmrange.Checked = True Then
            'Txt1to10Km.Enabled = True
            'Txt10to20Km.Enabled = True
            'TxtAbove20Km.Enabled = True
            'Txt1to10Km.MendatroryField = True
            'Txt10to20Km.MendatroryField = True
            'TxtAbove20Km.MendatroryField = True
            gv.Enabled = True
        Else
            'Txt1to10Km.Text = ""
            'Txt10to20Km.Text = ""
            'TxtAbove20Km.Text = ""
            'Txt1to10Km.Enabled = False
            'Txt10to20Km.Enabled = False
            'TxtAbove20Km.Enabled = False
            'Txt1to10Km.MendatroryField = False
            'Txt10to20Km.MendatroryField = False
            'TxtAbove20Km.MendatroryField = False
            gv.Enabled = False
            gv.Rows.Clear()
        End If
    End Sub
    Private Sub gv_UserDeletingRow(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gv.UserDeletingRow
        If Not myMessages.deleteConfirm() Then
            e.Cancel = True
        End If
    End Sub

    Private Sub txttank_transcode__MYOpenMasterForm(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txttank_transcode._MYOpenMasterForm
        Frm_Open = New frmTankerTransporterMaster(objCommonVar.CurrentUser, objCommonVar.CurrentCompanyName)
        Frm_Open.SetUserMgmt(clsUserMgtCode.frmSilageTankerTransporterMaster)
        Frm_Open.Show()
    End Sub

    Private Sub btnGO_Click(sender As Object, e As EventArgs) Handles btnGO.Click
        Try
            If clsCommon.myCdbl(txtChamborNo.Text) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Please enter no chmber", Me.Text)
                txtChamborNo.Focus()
                Exit Sub
            End If
            If clsCommon.myCdbl(txtChamborNo.Text) > 10 Then
                clsCommon.MyMessageBoxShow(Me, "Please enter no chmber less than or equl to 10", Me.Text)
                txtChamborNo.Focus()
                Exit Sub
            End If
            Dim i As Integer = 0
            If clsCommon.myCdbl(txtChamborNo.Value) > gvChamber.Rows.Count Then
                For i = gvChamber.Rows.Count + 1 To clsCommon.myCdbl(txtChamborNo.Value)
                    gvChamber.Rows.AddNew()
                    gvChamber.Rows(i - 1).Cells(colChamberNo).Value = i
                Next
            ElseIf clsCommon.myCdbl(txtChamborNo.Value) < gvChamber.Rows.Count Then
                For i = gvChamber.Rows.Count - 1 To clsCommon.myCdbl(txtChamborNo.Value) Step -1
                    gvChamber.Rows.RemoveAt(i)
                Next
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkMonthlyRentPlusDiesel_Click(sender As Object, e As EventArgs) Handles chkMonthlyRentPlusDiesel.Click
        rbtndiesel.IsChecked = False
        rbtnratekm.IsChecked = False
        rbtnrateltr.IsChecked = False
        rbtnrental.IsChecked = False
        rbtKmrange.IsChecked = False
        chkMonthlyRentPlusDiesel.IsChecked = False
    End Sub

    Private Sub chkMonthlyRentPlusDiesel_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkMonthlyRentPlusDiesel.ToggleStateChanged
        If isImportClicked = True Then
            Return
        End If
        If chkMonthlyRentPlusDiesel.IsChecked = True Then
            txtMRPDRent.Enabled = True
            txtMRPDAverage.Enabled = True
            txtMRPDDieselRate.Enabled = True
            txtMRPDRent.MendatroryField = True
            txtMRPDAverage.MendatroryField = True
            txtMRPDDieselRate.MendatroryField = True
        Else
            txtMRPDRent.Text = Nothing
            txtMRPDAverage.Text = Nothing
            txtMRPDDieselRate.Text = Nothing
            txtMRPDRent.Enabled = False
            txtMRPDAverage.Enabled = False
            txtMRPDDieselRate.Enabled = False
            txtMRPDRent.MendatroryField = False
            txtMRPDAverage.MendatroryField = False
            txtMRPDDieselRate.MendatroryField = False
        End If
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(fndNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowHistoryData(fndNo.Value, "Tanker_No", "TSPL_TANKER_MASTER")
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub TxtEPFRule__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles TxtEPFRule._MYValidating
        Try
            Dim qry As String = "  select TSPL_PF_RULE_MASTER.PFRULE_CODE as [Code] ,TSPL_PF_RULE_MASTER.APPLICABLE_FROM as [Applicable From] ,
                               TSPL_PF_RULE_MASTER.COEPF_PER as [Coepf Per] ,TSPL_PF_RULE_MASTER.COEPF_ROUNDOFF_YPE as [Coepf Roundoff Ype] ,
                               TSPL_PF_RULE_MASTER.COEPS_PER as [Coeps Per] ,TSPL_PF_RULE_MASTER.EPS_MAX as [Eps Max] ,TSPL_PF_RULE_MASTER.EMPEPF_PER as [Empepf Per] 
                               From TSPL_PF_RULE_MASTER  "
            TxtEPFRule.Value = clsCommon.ShowSelectForm("RoutMasFND", qry, "Code", "", TxtEPFRule.Value, "", isButtonClicked)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        'txttank_transcode.Value = clsVendorMaster.getFinder(" TSPL_VENDOR_MASTER.Form_Type='TTM' and TSPL_VENDOR_MASTER.isBulkProcurement=0 ", txttank_transcode.Value, isButtonClicked)

        'If clsCommon.myLen(txttank_transcode.Value) > 0 Then
        '    txtname.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select vendor_name  from tspl_vendor_master where form_type='TTM' and vendor_code='" & txttank_transcode.Value & "'"))
        'Else
        '    txtname.Text = ""
        'End If
    End Sub

    Private Sub TxtESIRule__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles TxtESIRule._MYValidating
        Try
            Dim qry As String = "  select TSPL_ESI_RULE_MASTER.ESIRULE_CODE as [Code] ,TSPL_ESI_RULE_MASTER.APPLICABLE_FROM as [Applicable From] ,
                                   TSPL_ESI_RULE_MASTER.COESI_PER as [Coesi Per] ,TSPL_ESI_RULE_MASTER.EMPESI_PER as [Empesi Per] ,
                                   TSPL_ESI_RULE_MASTER.COESI_ROUNDOFF_YPE as [Coesi Roundoff Ype] ,TSPL_ESI_RULE_MASTER.TOTALEARNING_MAX as [Totalearning Max]  
                                   From TSPL_ESI_RULE_MASTER  "
            TxtESIRule.Value = clsCommon.ShowSelectForm("RoutMasFND", qry, "Code", "", TxtESIRule.Value, "", isButtonClicked)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
End Class
