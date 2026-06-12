' '' '' ''Created By richa
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports common
Imports common.UserControls
Imports XpertERPEngine
Public Class frmDemandBooking
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim isExportTruckSheet As Boolean = False
    Dim PrintOnlyPostedDocument As Boolean = False
    Dim isIndent As Boolean = False
    Dim ApplyItemUOMOnDemand As Boolean = False
    Dim AllowRouteWiseDemandEntryInDecimal As Boolean = False
    'Dim GVTruckSheet As MyRadGridView
    Dim CheckCustomeroutStandingAmt As Boolean = False
    Dim gvFullMode As Boolean = False
    Dim SeprateMorningEveningSequence As Boolean = False
    Dim SetDefaultShiftTime As String = ""
    Dim AmountToCheckCustomerOutstandingForTCSTax As Double = 0
    Public Shared LockUnlock As Integer = 0
    Dim EnableLocation As Boolean = False
    Dim ApplyDepartmentRoute As Boolean = False
    Dim ApplyItemCapacityLimit As Boolean = False
    Dim isDepartmentRoute As Boolean = False
    Dim isDepartmentRouteSetting As Boolean = False
    Dim EnableProductSaleForJPR As Boolean = False
    Dim DisableRouteandVehicle As Boolean = False
    Dim AllowMultipleUOMForProduct As Boolean = False
    Dim EnableResetDemand As Boolean = False
    Dim ConvertPouchtoCrate As Boolean = False
    Dim DontCreateForPouch As Boolean = False
    Dim SeparateDemandMilkandProduct As Boolean = False
    Dim LockedByUserName As String = ""
    Dim LockedByUserCode As String = ""
    Dim dr As DataRow
    Dim SingleUserParticularDairyBookingEdit As Boolean = False
    Dim UseCutOffTimeonRouteForERP As Boolean = False
    Dim checkCreditLimit As Boolean = False
    Dim FlagPaste As Boolean = False
    Dim ChangeVehicleOnDairySaleBooking As Boolean = False
    Dim FlagChangeVehical As Boolean = False
    Dim ChangedRouteNo As String = ""
    Dim ChangedVehicalNo As String = ""
    'Dim ChangedVehicleCustCode As String = ""
    Dim blnPageLoad As Boolean = False
    Dim intChangeColumn As Integer = 0
    Public StrDocNo As String = ""
    Public arrBookingItem As List(Of clsBookingTemp) = Nothing
    Dim blnSaveTotalQTy As Boolean = False
    Dim DOmsg As String = ""
    Private isNewEntry As Boolean = False
    Private DOCreated As Boolean = False
    Dim AllowWo_Outstanding As Boolean
    Dim CheckOutstandingOnbooking As Integer = 0
    Dim DairyTaxableOrNonTaxable As Integer = 0
    Dim ShowItemLocationWiseonBooking As Integer = 0
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Dim isInsideLoadData As Boolean = False
    Dim isLoadData As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Const colItemExist As String = "colItemExist"
    Const colIsItemUpdate As String = "colIsItemUpdate"
    Const colBookingCreatedFor3Days As String = "colBookingCreatedFor3Days"
    Const colLineNo As String = "COLLNO"
    Const colTripNo As String = "colTripNo"
    Const colbtncol As String = "colbtncol"
    Const colCustCode As String = "colCustCode"
    Const colCreated_By As String = "colCreated_By"
    Const colSourceBy As String = "colSourceBy"
    Const colREF_PK_ID As String = "colREF_PK_ID"
    Const colCustName As String = "colCustName"
    Const colShiftName As String = "colShiftName"
    Const colItemCode As String = "colItemCode"
    Const colCrate As String = "colCrate"
    Const colLitre As String = "colLitre"
    Const colPCount As String = "colPCount"
    Const colPCrate As String = "colPCrate"
    Const colPAmt As String = "colPAmt"
    Const colMAmt As String = "colMAmt"
    Const colAmt As String = "COLAMT"
    'Const colTaxGroup As String = "colTaxGroup"
    ''Const colIsKKF As String = "colIsKKF"
    ''Const colIsMNDTax As String = "colIsMNDTax"
    'Const colTax As String = "colTax"
    'Const colTax_Base_Amt As String = "colTax_Base_Amt"
    'Const colTax_Rate As String = "colTax_Rate"
    'Const colTax_Amt As String = "colTax_Amt"
    Const colLocCode As String = "colLocCode"
    Const colLocName As String = "colLocName"
    Const ReportID As String = "DairyBookingGrid"
    Const colQty As String = "colQty"
    Const colICode As String = "colICode"
    Const colIDesc As String = "colIDesc"
    Const colIHSN As String = "colIHSN"
    Const colSDesc As String = "colSDesc"
    Const colUnit As String = "colUnit"
    Const colTotalQty As String = "colTotalQty"
    Const colISeqNo As String = "colISeqNo"
    Const colIGroup As String = "colIGroup"
    Dim strSql As String
    Dim dblCustOutstandingAmt As Double = 0
    Dim CalculateTaxRatefromItemwsieTaxOnSale As Boolean = False
    Dim DoNotConsiderCustomerCreditLimit As Boolean = False
    Dim ImportProcess As Boolean = False
    Dim SettSeprateDemandForMorningEveningShift As Boolean = False
    Dim UpdateDemandBeforePost As Boolean = False
    'Dim lstCustItem As List(Of clsDemandCustItem)
    Dim OneTimeCheck As Boolean = False
#End Region
    Private Sub FrmBookingEntry_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            GVTruckSheet.Visible = False
            isExportTruckSheet = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ExportTruckSheet, clsFixedParameterCode.ExportTruckSheet, Nothing)) = 1, True, False)
            PrintOnlyPostedDocument = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PrintOnlyPostedDocument, clsFixedParameterCode.PrintOnlyPostedDocument, Nothing)) = 1, True, False)
            gv1.EnterKeyMode = RadGridViewEnterKeyMode.EnterMovesToNextRow
            blnPageLoad = True
            EnableLocation = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableLocation, clsFixedParameterCode.EnableLocation, Nothing)) = 1, True, False)
            EnableResetDemand = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableResetDemand, clsFixedParameterCode.EnableResetDemand, Nothing)) = 1, True, False)
            SettSeprateDemandForMorningEveningShift = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SeprateDemandForMorningEveningShift, clsFixedParameterCode.SeprateDemandForMorningEveningShift, Nothing)) = 1)
            ChangeVehicleOnDairySaleBooking = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ChangeVehicleOnDairySaleBooking, clsFixedParameterCode.ChangeVehicleOnDairySaleBooking, Nothing)) = 0, False, True)
            ShowItemLocationWiseonBooking = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowItemLocationWiseonDairyBooking, clsFixedParameterCode.ShowItemLocationWiseonDairyBooking, Nothing))
            CheckOutstandingOnbooking = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CheckOutstandingCreditLimitOnBooking, clsFixedParameterCode.CheckOutstandingCreditLimitOnBooking, Nothing))
            AllowWo_Outstanding = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Description from TSPL_FIXED_PARAMETER where Code='" & clsFixedParameterCode.AllowDispatchOutstandingFS & "'")) = 0, False, True)
            CalculateTaxRatefromItemwsieTaxOnSale = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CalculateTaxRatefromItemwsieTaxOnSale, clsFixedParameterCode.CalculateTaxRatefromItemwsieTaxOnSale, Nothing)) = 1, True, False)
            SingleUserParticularDairyBookingEdit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SingleUserParticularDairyBookingEdit, clsFixedParameterCode.SingleUserParticularDairyBookingEdit, Nothing)) = 1, True, False)
            DoNotConsiderCustomerCreditLimit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DoNotConsiderCustomerCreditLimit, clsFixedParameterCode.DoNotConsiderCustomerCreditLimit, Nothing)) = 1, True, False)
            UseCutOffTimeonRouteForERP = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.UseCutOffTimeonRouteForERP, clsFixedParameterCode.UseCutOffTimeonRouteForERP, Nothing)) = 1, True, False)
            checkCreditLimit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CheckCreditLimit, clsFixedParameterCode.CheckCreditLimit, Nothing)) = 1, True, False)
            SeparateDemandMilkandProduct = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SeparateDemandMilkandProduct, clsFixedParameterCode.SeparateDemandMilkandProduct, Nothing)) = 1, True, False)
            ConvertPouchtoCrate = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ConvertPouchtoCrate, clsFixedParameterCode.ConvertPouchtoCrate, Nothing)) = 1, True, False)
            DisableRouteandVehicle = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DisableRouteandVehicle, clsFixedParameterCode.DisableRouteandVehicle, Nothing)) = 1, True, False)
            UpdateDemandBeforePost = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.UpdateDemandBeforePost, clsFixedParameterCode.UpdateDemandBeforePost, Nothing)) = 1, True, False)
            AmountToCheckCustomerOutstandingForTCSTax = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AmountToCheckCustomerOutstandingForTCSTax, clsFixedParameterCode.AmountToCheckCustomerOutstandingForTCSTax, Nothing))
            DontCreateForPouch = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DontCreateForPouch, clsFixedParameterCode.DontCreateForPouch, Nothing)) = 1, True, False)
            AllowMultipleUOMForProduct = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowMultipleUOMForProduct, clsFixedParameterCode.AllowMultipleUOMForProduct, Nothing)) = 1, True, False)
            SetDefaultShiftTime = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SetDefaultShiftTime, clsFixedParameterCode.SetDefaultShiftTime, Nothing))
            ApplyDepartmentRoute = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyDepartmentRoute, clsFixedParameterCode.ApplyDepartmentRoute, Nothing)) = 1, True, False)
            ApplyItemCapacityLimit = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyItemCapacityLimit, clsFixedParameterCode.ApplyItemCapacityLimit, Nothing)) = 1, True, False)
            SeprateMorningEveningSequence = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.SeprateMorningEveningSequence, clsFixedParameterCode.SeprateMorningEveningSequence, Nothing)) = 1, True, False)
            ApplyItemUOMOnDemand = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyItemUOMOnDemand, clsFixedParameterCode.ApplyItemUOMOnDemand, Nothing)) = 1, True, False)
            AllowRouteWiseDemandEntryInDecimal = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowRouteWiseDemandEntryInDecimal, clsFixedParameterCode.AllowRouteWiseDemandEntryInDecimal, Nothing)) = 1, True, False)
            EnableProductSaleForJPR = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableProductSaleForJPR, clsFixedParameterCode.EnableProductSaleForJPR, Nothing)) = 1, True, False)
            CheckCustomeroutStandingAmt = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CheckCustomeroutStandingAmt, clsFixedParameterCode.CheckCustomeroutStandingAmt, Nothing)) = 1, True, False)

            CrateHisTable()
            AddNew()
            SetUserMgmtNew()
            If clsCommon.myLen(StrDocNo) > 0 Then
                LoadData(StrDocNo, NavigatorType.Current)
            End If
            If clsCommon.myLen(Me.Tag) > 0 Then
                LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
            End If
            txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "' "))
            If clsCommon.myLen(txtLocation.Value) > 0 Then
                lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
            End If
            blnPageLoad = False
            'LoadBlankGrid()
            If SettSeprateDemandForMorningEveningShift Then
                rbtnMorningEveningBoth.Enabled = False
                rbtnMorningEveningBoth.Visible = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub frmBookingDairyMultipleCustomer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            'sanjay GKD/20/06/18-000150
            If SingleUserParticularDairyBookingEdit AndAlso clsCommon.myLen(txtDocNo.Value) > 0 AndAlso LockUnlock = 1 AndAlso LockedByUserCode = objCommonVar.CurrentUserCode Then
                'If LockUnlock = 1 AndAlso LockedByUserCode = objCommonVar.CurrentUserCode Then
                Dim qry As String = ""
                qry = "update tspl_booking_matser set User_Lock_For_Edit=0,lockedby_usercode=''  where LockedBy_UserCode='" & objCommonVar.CurrentUserCode & "'"
                clsDBFuncationality.ExecuteNonQuery(qry)
                'End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub FrmBookingEntry_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Try
            If e.KeyCode = Keys.F2 AndAlso gv1.CurrentCell IsNot Nothing Then
                'setGridFocus()
            ElseIf e.Alt AndAlso e.KeyCode = Keys.N AndAlso btnAddNew.Enabled Then
                AddNew()
            ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso btnSave.Enabled AndAlso MyBase.isModifyFlag Then
                If clsCommon.myLen(txtDocNo.Value) > 0 Then
                    Dim RouteNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Route_No from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + txtDocNo.Value + "'"))
                    If clsCommon.CompairString(txtRouteNo.Value, RouteNo) = CompairStringResult.Equal Then
                        SaveData(False)
                    Else
                        clsCommon.MyMessageBoxShow(Me, "You can't change route", Me.Text)
                        txtRouteNo.Value = RouteNo
                    End If
                Else
                    SaveData(False)
                End If

            ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso btnDelete.Enabled AndAlso MyBase.isDeleteFlag Then
                DeleteData()
            ElseIf e.Alt AndAlso e.Control AndAlso e.Shift AndAlso e.KeyCode = Keys.F10 Then

                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
                frm.strCode = "ShuffleDemand"
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    gbShuffleDemand.Visible = True
                End If
            ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnClose.Enabled Then
                CloseForm()
            ElseIf e.KeyCode = Keys.Enter Then
                setGridFocus()
            ElseIf e.KeyCode = Keys.PageDown Then
                setPagedown()
            ElseIf e.KeyCode = Keys.Home Then
                setGridFocusHome()
            ElseIf e.KeyCode = Keys.End Then
                setGridFocusEnd()
            ElseIf e.Alt AndAlso e.Shift AndAlso e.Control AndAlso e.KeyCode = Keys.F12 Then
                If MyBase.isReverse Then
                    Dim frm As New FrmPWD(Nothing)
                    frm.strType = clsFixedParameterType.SIR
                    frm.strCode = clsFixedParameterCode.SIReversAndCreate
                    frm.ShowDialog()
                    If frm.isPasswordCorrect Then
                        btnreverse.Visible = True
                    End If
                    ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                      "TSPL_DEMAND_BOOKING_MASTER " + Environment.NewLine +
                                      "TSPL_DEMAND_BOOKING_DETAIL " + Environment.NewLine +
                    "TSPL_BOOKING_MATSER " + Environment.NewLine +
                                      "TSPL_BOOKING_DETAIL " + Environment.NewLine +
                                      "TSPL_GATEPASS_MASTER_DAIRYSALE (For Gate Pass Document) " + Environment.NewLine +
                                      "TSPL_GATEPASS_DETAIL_DAIRYSALE (For Gate Pass Document) " + Environment.NewLine +
                                      "Press Alt+F for Create DO/Post DO Trasnaction" + Environment.NewLine +
                                      "TSPL_DELIVERY_NOTE_MASTER_FRESHSALE " + Environment.NewLine +
                                      "TSPL_DELIVERY_NOTE_DETAIL_FRESHSALE " + Environment.NewLine +
                                      "TSPL_TRANSACTION_APPROVAL (For Approving Pending Document) ")
                Else
                    clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
                    'MessageBox.Show("You are not authorized to perform this action.", "Unauthorized Access", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub CloseForm()
        Me.Close()
    End Sub
    Public Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.frmbookingdairy)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
            Me.Close()
            Exit Sub
        End If
        btnSave.Visible = MyBase.isModifyFlag
        'btnPost.Visible = MyBase.isPostFlag
        btnDelete.Visible = MyBase.isDeleteFlag
        btnPost.Visible = MyBase.isPostFlag
        btnPrint.Visible = MyBase.isPrintFlag
        btnReverseAndUnpost.Visible = MyBase.isReverse
        'RadMenu1.Visible = MyBase.isExport
        If MyBase.isExport Then
            btnExport.Enabled = True
            btnImport.Enabled = True
        Else
            btnExport.Enabled = False
            btnImport.Enabled = False
        End If
        btnreverse.Visible = False
        'If MyBase.isReverse Then
        '    btnreverse.Enabled = True
        'Else
        '    btnreverse.Enabled = False
        'End If
    End Sub
    Function FillMorningEvening() As DataTable
        Dim qry As String = String.Empty
        If rbtnMorning.IsChecked Then
            qry = "select 'Morning' as value union all select 'Evening' as value "
        ElseIf rbtnEvening.IsChecked Then
            qry = "select 'Evening' as value "
        Else
            qry = "select 'Morning' as value union all select 'Evening' as value "
        End If
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        Return dt
    End Function
    Sub LoadBlankGrid()
        Dim qry As String = String.Empty
        gv1.Rows.Clear()
        gv1.Columns.Clear()
        If Not SettSeprateDemandForMorningEveningShift Then
            rbtnMorningEveningBoth.IsChecked = True
        End If
        'rdbnFreshAmbientBoth.IsChecked = True
        gv1.DataSource = Nothing
        'gv1.ViewDefinition = New TableViewDefinition
        gv1.Rows.AddNew()
        'Dim TAX_PAID As New GridViewComboBoxColumn
        'Dim BOOK_RATE_UOM As New GridViewTextBoxColumn
        Dim RepobtnCol As GridViewCommandColumn = New GridViewCommandColumn()
        RepobtnCol.HeaderText = "Action "
        RepobtnCol.Name = colbtncol
        RepobtnCol.ReadOnly = False
        RepobtnCol.Width = 150
        RepobtnCol.DefaultText = "Reset ..."
        If EnableResetDemand Then
            RepobtnCol.IsVisible = True
        Else
            RepobtnCol.IsVisible = False
        End If
        RepobtnCol.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gv1.MasterTemplate.Columns.Add(RepobtnCol)
        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = colLineNo
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.IsPinned = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLineNo)
        Dim repoTripNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTripNo = New GridViewDecimalColumn()
        repoTripNo.FormatString = ""
        repoTripNo.HeaderText = "Trip No"
        repoTripNo.Name = colTripNo
        repoTripNo.Width = 50
        repoTripNo.ReadOnly = False
        repoTripNo.IsPinned = True
        repoTripNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoTripNo)
        Dim repoCustCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCustCode.FormatString = ""
        repoCustCode.HeaderText = "Code"
        repoCustCode.Name = colCustCode
        repoCustCode.HeaderImage = My.Resources.search4
        repoCustCode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoCustCode.Width = 50
        repoCustCode.IsVisible = True
        repoCustCode.IsPinned = True
        repoCustCode.ReadOnly = True
        gv1.MasterTemplate.Columns.Add(repoCustCode)
        Dim repoCName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCName.FormatString = ""
        repoCName.HeaderText = "Name"
        repoCName.Name = colCustName
        repoCName.Width = 150
        repoCName.ReadOnly = True
        repoCName.IsVisible = True
        repoCName.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoCName)
        Dim repoShiftName As GridViewComboBoxColumn = New GridViewComboBoxColumn()
        repoShiftName.FormatString = ""
        repoShiftName.HeaderText = "Shift"
        repoShiftName.Name = colShiftName
        repoShiftName.Width = 100
        repoShiftName.IsVisible = True
        repoShiftName.DataSource = FillMorningEvening()
        repoShiftName.ReadOnly = True
        repoShiftName.DisplayMember = "Value"
        repoShiftName.ValueMember = "Value"
        repoShiftName.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoShiftName)
        Dim repoItemValue As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoItemValue.FormatString = ""
        repoItemValue.HeaderText = "Is Item Value Exist"
        repoItemValue.Name = colItemExist
        repoItemValue.Width = 150
        repoItemValue.ReadOnly = True
        repoItemValue.IsVisible = False
        repoItemValue.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoItemValue)
        Dim repoItemUpdate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoItemUpdate.FormatString = ""
        repoItemUpdate.HeaderText = "Is Item Update After Save"
        repoItemUpdate.Name = colIsItemUpdate
        repoItemUpdate.Width = 150
        repoItemUpdate.ReadOnly = True
        repoItemUpdate.IsVisible = False
        repoItemUpdate.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoItemUpdate)
        Dim repoBookingCreatedFor3Days As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoBookingCreatedFor3Days.FormatString = ""
        repoBookingCreatedFor3Days.HeaderText = "Booking Created for last 3 days"
        repoBookingCreatedFor3Days.Name = colBookingCreatedFor3Days
        repoBookingCreatedFor3Days.Width = 150
        repoBookingCreatedFor3Days.ReadOnly = True
        repoBookingCreatedFor3Days.IsVisible = False
        repoBookingCreatedFor3Days.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoBookingCreatedFor3Days)
        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        qry = "select * from (select 'Fresh' as FreshAmbient,tspl_item_master.Item_Code,tspl_item_master.Short_Description ,tspl_item_master.Item_Desc , TSPL_ITEM_UOM_DETAIL.UOM_Code,tspl_item_master.Short_Description +' - '+ TSPL_ITEM_UOM_DETAIL.UOM_Code as ItemDescNew,1 as RowNo,tspl_item_master.Sku_Seq,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,TSPL_ITEM_UOM_DETAIL.Stocking_Unit   from tspl_item_master 
    left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL .item_code=tspl_item_master.Item_Code 
    where  tspl_item_master.Is_FreshItem =1 AND TSPL_ITEM_MASTER.Is_Milk_Pouch =1 and isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0  and Item_Type ='F' and tspl_item_master.Active=1 and tspl_item_master.Is_DisplayDemand=1
    and TSPL_ITEM_UOM_DETAIL.Uom_code in ('Crate','Pouch','LTR')
    union
    select 'Fresh' as FreshAmbient,tspl_item_master.Item_Code ,tspl_item_master.Short_Description,tspl_item_master.Item_Desc , TSPL_ITEM_UOM_DETAIL.UOM_Code,tspl_item_master.Short_Description +' - '+ TSPL_ITEM_UOM_DETAIL.UOM_Code as ItemDescNew,1 as RowNo,tspl_item_master.Sku_Seq,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,TSPL_ITEM_UOM_DETAIL.Stocking_Unit  from tspl_item_master 
    left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL .item_code=tspl_item_master.Item_Code 
    where  tspl_item_master.Is_FreshItem =1 AND TSPL_ITEM_MASTER.Is_Milk_Pouch =1 and isnull(TSPL_ITEM_MASTER.CAN,0)=0  and isnull(TSPL_ITEM_MASTER.CRATE,0)=0  and Item_Type ='F' and tspl_item_master.Active=1 and tspl_item_master.Is_DisplayDemand=1
    and TSPL_ITEM_UOM_DETAIL.Uom_code in ('Crate','Pouch','LTR')"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal Then
            qry += " union all
    Select 'Ambient' as FreshAmbient,tspl_item_master.Item_Code ,tspl_item_master.Short_Description,tspl_item_master.Item_Desc , TSPL_ITEM_UOM_DETAIL.UOM_Code,tspl_item_master.Short_Description +' - '+ TSPL_ITEM_UOM_DETAIL.UOM_Code as ItemDescNew,2 as RowNo,tspl_item_master.Sku_Seq,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,TSPL_ITEM_UOM_DETAIL.Stocking_Unit   from tspl_item_master 
    Left outer join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL .item_code=tspl_item_master.Item_Code 
    where  tspl_item_master.Is_Ambient = 1 And isnull(TSPL_ITEM_MASTER.CAN, 0) = 0 And isnull(TSPL_ITEM_MASTER.CRATE, 0) = 0 And Item_Type ='F' and tspl_item_master.Active=1 and tspl_item_master.Is_DisplayDemand=1
And TSPL_ITEM_UOM_DETAIL.Default_UOM = 1"


        End If
        qry += " )z order by RowNo,Sku_Seq,Item_Code"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim i As Integer = 1
            Dim obj As ItemValueClass = New ItemValueClass()
            For Each dr As DataRow In dt.Rows
                repoIName = New GridViewTextBoxColumn()
                repoIName.FormatString = ""
                repoIName.HeaderText = clsCommon.myCstr(dr("UOM_Code"))
                obj = New ItemValueClass()
                obj.itemCode = clsCommon.myCstr(dr("Item_Code"))
                obj.itemDesc = clsCommon.myCstr(dr("Item_Desc"))
                obj.Unit_code = clsCommon.myCstr(dr("UOM_Code"))
                obj.IsFreshAmbient = clsCommon.myCstr(dr("FreshAmbient"))
                obj.ShortDesc = clsCommon.myCstr(dr("Short_Description"))
                obj.Conversion_Factor = clsCommon.myCdbl(dr("Conversion_Factor"))
                obj.Stocking_Unit = clsCommon.myCstr(dr("Stocking_Unit"))
                repoIName.Tag = obj
                repoIName.Name = colItemCode + clsCommon.myCstr(i)
                repoIName.Width = 150
                repoIName.IsVisible = True
                i = i + 1
                gv1.MasterTemplate.Columns.Add(repoIName)
            Next
        End If
        Dim repoCrate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoCrate = New GridViewDecimalColumn()
        repoCrate.FormatString = ""
        repoCrate.HeaderText = "Crate"
        repoCrate.Name = colCrate
        repoCrate.Width = 80
        repoCrate.Minimum = 0
        repoCrate.ReadOnly = True
        repoCrate.IsVisible = True
        repoCrate.IsPinned = True
        repoCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoCrate)
        Dim repoLtr As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLtr = New GridViewDecimalColumn()
        repoLtr.FormatString = ""
        repoLtr.HeaderText = "Litre"
        repoLtr.Name = colLitre
        repoLtr.Width = 80
        repoLtr.Minimum = 0
        repoLtr.ReadOnly = True
        repoLtr.IsVisible = True
        repoLtr.IsPinned = True
        repoCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoLtr)
        Dim repoMAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoMAmt = New GridViewDecimalColumn()
        repoMAmt.FormatString = ""
        repoMAmt.HeaderText = "MAmt"
        repoMAmt.Name = colMAmt
        repoMAmt.Width = 80
        repoMAmt.Minimum = 0
        repoMAmt.ReadOnly = True
        repoMAmt.IsVisible = True
        repoMAmt.IsPinned = True
        repoCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoMAmt)
        Dim repoPCrate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPCrate = New GridViewDecimalColumn()
        repoPCrate.FormatString = ""
        repoPCrate.HeaderText = "PCrate"
        repoPCrate.Name = colPCrate
        repoPCrate.Width = 80
        repoPCrate.Minimum = 0
        repoPCrate.ReadOnly = True
        If AllowMultipleUOMForProduct Then
            repoPCrate.IsVisible = True
        Else
            repoPCrate.IsVisible = False
        End If

        repoPCrate.IsPinned = True
        repoCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoPCrate)
        Dim repoPCount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPCount = New GridViewDecimalColumn()
        repoPCount.FormatString = ""
        repoPCount.HeaderText = "PCount"
        repoPCount.Name = colPCount
        repoPCount.Width = 80
        repoPCount.Minimum = 0
        repoPCount.ReadOnly = True
        repoPCount.IsVisible = True
        repoPCount.IsPinned = True
        repoCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoPCount)
        Dim repoPAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoPAmt = New GridViewDecimalColumn()
        repoPAmt.FormatString = ""
        repoPAmt.HeaderText = "PAmt"
        repoPAmt.Name = colPAmt
        repoPAmt.Width = 80
        repoPAmt.Minimum = 0
        repoPAmt.ReadOnly = True
        repoPAmt.IsVisible = True
        repoPAmt.IsPinned = True
        repoCrate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoPAmt)
        Dim repoAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmt = New GridViewDecimalColumn()
        repoAmt.FormatString = ""
        repoAmt.HeaderText = "Total Amt"
        repoAmt.Name = colAmt
        repoAmt.Width = 80
        repoAmt.Minimum = 0
        repoAmt.ReadOnly = True
        repoAmt.IsVisible = True
        repoAmt.IsPinned = True
        repoAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv1.MasterTemplate.Columns.Add(repoAmt)
        Dim repoCreatedby As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoCreatedby.FormatString = ""
        repoCreatedby.HeaderText = "Created By"
        repoCreatedby.Name = colCreated_By
        repoCreatedby.Width = 150
        repoCreatedby.ReadOnly = True
        repoCreatedby.IsVisible = False
        repoCreatedby.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoCreatedby)
        Dim repoREF_PKID As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoREF_PKID.FormatString = ""
        repoREF_PKID.HeaderText = "Ref PK ID"
        repoREF_PKID.Name = colREF_PK_ID
        repoREF_PKID.Width = 150
        repoREF_PKID.ReadOnly = True
        repoREF_PKID.IsVisible = False
        repoREF_PKID.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoREF_PKID)
        Dim repoSourceby As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSourceby.FormatString = ""
        repoSourceby.HeaderText = "Source By"
        repoSourceby.Name = colSourceBy
        repoSourceby.Width = 150
        repoSourceby.ReadOnly = True
        repoSourceby.IsVisible = False
        repoSourceby.IsPinned = True
        gv1.MasterTemplate.Columns.Add(repoSourceby)

        gv1.AllowDeleteRow = True
        gv1.AllowAddNewRow = False
        gv1.ShowGroupPanel = False
        gv1.AllowColumnReorder = False
        gv1.AllowRowReorder = False
        gv1.EnableSorting = False
        gv1.EnableFiltering = False
        gv1.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv1.MasterTemplate.ShowRowHeaderColumn = False
        gv1.TableElement.TableHeaderHeight = 40
        View()
    End Sub
    Private Sub setGridFocus()
        Dim intCurrRow As Integer = gv1.CurrentRow.Index
        If intCurrRow = gv1.Rows.Count - 1 Then
            intCurrRow = 0
            gv1.CurrentRow = gv1.Rows(intCurrRow)
        End If
        'If gv1.CurrentColumn Is gv1.Columns(gv1.Columns.Count - 11) Then
        'End If
    End Sub
    Private Sub setGridFocusHome()
        Dim intCurrRow As Integer = gv1.CurrentRow.Index
        If gv1.Rows.Count > 0 Then
            'gv1.CurrentColumn = gv1.Columns(7)
            gv1.Rows(intCurrRow).Cells(9).IsSelected = True
            gv1.Rows(intCurrRow).IsCurrent = True
            gv1.Columns(9).IsCurrent = True
        End If
    End Sub
    Private Sub setGridFocusEnd()
        Dim intCurrRow As Integer = gv1.CurrentRow.Index
        If gv1.Rows.Count > 0 Then
            gv1.Rows(intCurrRow).Cells(gv1.Columns.Count - 11).IsSelected = True
            gv1.Rows(intCurrRow).IsCurrent = True
            gv1.Columns(gv1.Columns.Count - 11).IsCurrent = True
        End If
    End Sub
    Private Sub setPagedown()
        Try
            Dim scrollDelta As Integer = gv1.TableElement.ViewElement.ScrollableRows.Size.Height + CInt(gv1.TableElement.ViewElement.ScrollableRows.ScrollOffset.Height)
            Dim newVScrollValue As Integer = gv1.TableElement.VScrollBar.Value + scrollDelta
            If newVScrollValue < gv1.TableElement.VScrollBar.Maximum - gv1.TableElement.VScrollBar.LargeChange Then
                gv1.TableElement.VScrollBar.Value = newVScrollValue
            Else
                gv1.TableElement.VScrollBar.Value = gv1.TableElement.VScrollBar.Maximum - gv1.TableElement.VScrollBar.LargeChange
            End If
            Dim navigator As IGridNavigator = gv1.GridViewElement.Navigator
            navigator.BeginSelection(New GridNavigationContext(GridNavigationInputType.Keyboard, MouseButtons.None, Keys.None))
            navigator.SelectRow(Me.GetLastScrollableRow(gv1.TableElement))
            navigator.EndSelection()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Function GetLastScrollableRow(ByVal tableElement As GridTableElement) As GridViewRowInfo
        Dim rows As ScrollableRowsContainerElement = tableElement.ViewElement.ScrollableRows
        Dim traverser As GridTraverser = CType((CType(tableElement.RowScroller, IEnumerable)).GetEnumerator(), GridTraverser)
        For i As Integer = 0 To rows.Children.Count - 1
            If rows.Children(i).BoundingRectangle.Bottom > rows.Size.Height Then
                Exit For
            End If
            If Not traverser.MoveNext() Then
                Exit For
            End If
        Next
        Return traverser.Current
    End Function
    Sub View()
        Try
            If gv1.Rows.Count > 0 Then
                Dim view As New ColumnGroupsViewDefinition()
                view.ColumnGroups.Add(New GridViewColumnGroup("Booth"))
                view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colbtncol).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colLineNo).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colTripNo).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colCustCode).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colCustName).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colShiftName).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colItemExist).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colIsItemUpdate).Name)
                view.ColumnGroups(0).Rows(0).ColumnNames.Add(gv1.Columns(colBookingCreatedFor3Days).Name)
                view.ColumnGroups(0).IsPinned = True
                Dim TempColGroupCount As Integer = 1
                'Dim obj As ItemValueClass = New ItemValueClass()
                'Dim i As Integer = 1
                For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                    Dim obj1 As ItemValueClass = TryCast(gv1.Columns(dblcolumns).Tag, ItemValueClass)
                    If obj1 IsNot Nothing Then
                        If clsCommon.CompairString(obj1.IsFreshAmbient, "Fresh") = CompairStringResult.Equal Then
                            view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(obj1.ShortDesc)))
                            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            dblcolumns = dblcolumns + 1
                            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            dblcolumns = dblcolumns + 1
                            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            TempColGroupCount += 1
                        Else
                            view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(obj1.ShortDesc)))
                            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(dblcolumns).Name)
                            TempColGroupCount += 1
                        End If
                    End If
                Next
                view.ColumnGroups.Add(New GridViewColumnGroup("Total"))
                view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colCrate).Name)
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") <> CompairStringResult.Equal Then
                    'view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colCrate).Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colLitre).Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colMAmt).Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colPCrate).Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colPCount).Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colPAmt).Name)
                End If
                'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
                '    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colCrate).Name)
                'End If
                ''17/12 add crate for visibility for GNG 
                'view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colCrate).Name)
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(gv1.Columns(colAmt).Name)
                view.ColumnGroups(TempColGroupCount).IsPinned = True
                view.ColumnGroups(TempColGroupCount).PinPosition = PinnedColumnPosition.Right
                'MergeHorizontally(gv1, 0, gv1.Rows.Count - 1)
                gv1.ViewDefinition = view
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        'detect up arrow key
        Try
            If keyData = Keys.Up AndAlso gv1.CurrentRow.IsSelected Then
                gv1.CurrentRow = gv1.Rows(gv1.CurrentRow.Index - 1)
            ElseIf keyData = Keys.Down Then
                gv1.CurrentRow = gv1.Rows(gv1.CurrentRow.Index + 1)
            End If
            Return MyBase.ProcessCmdKey(msg, keyData)
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub ReStoreGridLayout()
        Try
            If clsCommon.myLen(ReportID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(ReportID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If obj IsNot Nothing AndAlso obj.GridColumns >= gv1.ColumnCount Then
                    Dim ii As Integer = 0
                    For ii = 0 To gv1.Columns.Count - 1 Step ii + 1
                        gv1.Columns(ii).IsVisible = False
                        gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
                obj = Nothing
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub
    Sub AddNew()
        Try
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            '    btnPrintChallan.Visible = True
            'Else
            '    btnPrintChallan.Visible = False
            'End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                btnSplitPrint.Visible = True
            Else
                btnSplitPrint.Visible = False
            End If
            'lstCustItem = New List(Of clsDemandCustItem)
            blnSaveTotalQTy = False
            isNewEntry = True
            btnSave.Text = "Save"
            txtDocNo.Value = ""
            btnSave.Enabled = True
            btnDelete.Enabled = True
            btnPost.Enabled = True
            btn_TruckSheet.Enabled = False
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                SplitButtonTruckSheet.Enabled = True
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                    If rdbnFreshAmbientBoth.IsChecked Then
                        SplitButtonTruckSheet.Enabled = True
                    Else
                        SplitButtonTruckSheet.Enabled = False
                    End If
                End If

            Else
                SplitButtonTruckSheet.Enabled = False
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                lblDemandUniqueID.Visible = True
                txtDemandUniqueID.Visible = True
                txtDemandUniqueID.Text = ""
            Else
                lblDemandUniqueID.Visible = False
                txtDemandUniqueID.Visible = False
                txtDemandUniqueID.Text = ""
            End If
            btn_TSCancel.Enabled = True
            btn_Gatepass.Enabled = False
            btnPrint.Enabled = True
            btn_TSCancel.Enabled = False
            btn_GPCancel.Enabled = False
            chkNoCrateIssue.Enabled = True
            chkNoCrateIssue.Checked = False
            txtcustomersearch.Text = ""
            txtDate.Value = clsCommon.GETSERVERDATE()
            txtShuffleDate.Value = txtDate.Value
            cmbShift.Text = "Evening"
            UsLock1.Status = ERPTransactionStatus.Pending
            chkIndividualCustomer.Checked = False
            chkIndividualCustomer.Enabled = True
            lblTotalCrate.Text = "0"
            lblTotalLitre.Text = "0"
            lblDocumentAmt.Text = "0"
            TxtCity.Value = ""
            lblTransporterName.Text = ""
            lblCityName.Text = ""
            txtRouteNo.Value = ""
            txtRouteNo.Enabled = True
            lblRouteDesc.Text = ""
            txtVehicleNo.Value = ""
            txtVehicleNo.Enabled = True
            lblVehicleNo.Text = ""
            txtTripNo.Text = ""
            txtCustomerNo.Value = ""
            lblCustomerName.Text = ""
            txtLocation.Value = Nothing
            txtCustomerNo.Enabled = False
            lblLocation.Text = ""
            txtPCount.Text = "0"
            txtPAmt.Text = "0"
            txtDocAmt.Text = "0"
            chkMorningPosted.Checked = False
            chkEveningPosted.Checked = False
            RadGroupBox1.Enabled = True
            txtDate.Enabled = True
            If EnableLocation Then
                txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Route_Master where Route_No='" + txtRouteNo.Value + "' "))
            Else
                txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
            End If
            If clsCommon.myLen(txtLocation.Value) > 0 Then
                lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtLocation.Value + "'"))
            End If
            If Not SettSeprateDemandForMorningEveningShift Then
                rbtnMorningEveningBoth.IsChecked = True
            End If
            RadGroupBox3.Enabled = True
            'rdbnFreshAmbientBoth.IsChecked = True
            chkMorningGatepassTruckSheetGenerated.Checked = False
            chkEveningGatepassTruckSheetGenerated.Checked = False
            FlagChangeVehical = False
            LoadBlankGrid()
            FlagPaste = False
            RefreshFormName()
            If SetDefaultShiftTime.Length > 0 Then
                Dim CurrDateTime As DateTime = clsCommon.GETSERVERDATE
                Dim EndTime As DateTime = clsCommon.GetPrintDate(SetDefaultShiftTime, "dd/MMM/yyyy hh:mm tt")
                If CurrDateTime.TimeOfDay < EndTime.TimeOfDay Then
                    txtDate.Value = clsCommon.GetPrintDate(CurrDateTime)
                    rbtnEvening.IsChecked = True

                Else
                    txtDate.Value = clsCommon.GetPrintDate(CurrDateTime.AddDays(1))
                    rbtnMorning.IsChecked = True
                End If
            End If
            gbShuffleDemand.Visible = False

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub gv1_CellValueChanged(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gv1.CellValueChanged
        Try
            If Not isInsideLoadData Then
                If Not isCellValueChangedOpen Then
                    isCellValueChangedOpen = True
                    If e.Column.Index >= 9 AndAlso e.Column.Name <> colCrate AndAlso e.Column.Name <> colAmt AndAlso e.Column.Name <> colLitre AndAlso e.Column.Name <> colMAmt AndAlso e.Column.Name <> colPCount AndAlso e.Column.Name <> colPCount AndAlso Not isLoadData Then
                        'If isLoadData = False AndAlso (clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) > 0) Then
                        'If Not isLoadData Then
                        ''UpdateItemQtyAfterSave(gv1.CurrentRow.Index, gv1.CurrentColumn.Index)
                        ' UpdateAllTotals(False)
                        UpdateRowTotal()
                        'HideUnhideRowsAndColumnsOFGrid()
                        'End If
                    End If
                    isCellValueChangedOpen = False
                End If
                isInsideLoadData = False
            End If
        Catch ex As Exception
            isInsideLoadData = False
            isCellValueChangedOpen = False
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            'Finally
            '    isInsideLoadData = False
            '    isCellValueChangedOpen = False
        End Try
    End Sub
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        Try
            txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" & objCommonVar.CurrentUserCode & "' "))
            lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
            AddNew()
            gv1.DataSource = Nothing
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Function AllowToSave(ByVal trans As SqlTransaction) As Boolean
        Try
            Xtra.TransactionValidity(txtDate.Value)
            If clsCommon.myLen(txtRouteNo.Value) <= 0 Then
                Throw New Exception("Please select Route")
            End If
            If clsCommon.myLen(TxtCity.Value) <= 0 Then
                Throw New Exception("Please select City")
            End If
            If clsCommon.myLen(txtLocation.Value) <= 0 Then
                Throw New Exception("Please select Location")
            End If
            If clsCommon.myLen(txtVehicleNo.Value) <= 0 Then
                Throw New Exception("Please select Vehicle")
            End If
            If SettSeprateDemandForMorningEveningShift AndAlso Not (rbtnMorning.IsChecked OrElse rbtnEvening.IsChecked) Then
                'If Not (rbtnMorning.IsChecked OrElse rbtnEvening.IsChecked) Then
                Throw New Exception("Shift should be Morning/Evening")
                'End If
            End If
            isInsideLoadData = True
            'UpdateAllTotals(False)
            isInsideLoadData = False
            Dim dblQuantityCount As Double = 0
            Dim dblQuantityMORNINGCount As Double = 0
            Dim dblQuantityEveningCount As Double = 0
            For ii As Integer = 0 To gv1.Rows.Count - 1
                Dim strItemValueExist As String = ""
                If clsCommon.myLen(gv1.Rows(ii).Cells(colItemExist)) > 0 Then
                    strItemValueExist = clsCommon.myCstr(gv1.Rows(ii).Cells(colItemExist).Value)
                End If
                If clsCommon.CompairString(strItemValueExist, "Yes") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colShiftName).Value), "Morning") = CompairStringResult.Equal Then
                        dblQuantityMORNINGCount = dblQuantityMORNINGCount + 1
                    End If
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colShiftName).Value), "Evening") = CompairStringResult.Equal Then
                        dblQuantityEveningCount = dblQuantityEveningCount + 1
                    End If
                    dblQuantityCount = dblQuantityCount + 1
                End If
            Next
            If dblQuantityCount <= 0 Then
                Throw New Exception("Please enter quantity for atleast one customer")
            End If
            ''If  Then
            Dim Morningvalue As Double = 0
            Dim eveningvalue As Double = 0
            If Not chkIndividualCustomer.Checked AndAlso clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) <= 0 AndAlso UseCutOffTimeonRouteForERP Then
                Dim strMorningShiftTime As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  MorningCutOff_Time from TSPL_ROUTE_MASTER where route_no='" & txtRouteNo.Value & "'"))
                Dim strEveningShiftTime As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  EveningCutOff_Time from TSPL_ROUTE_MASTER where route_no='" & txtRouteNo.Value & "'"))
                If dblQuantityMORNINGCount > 0 Then
                    If clsCommon.myLen(strMorningShiftTime) <= 0 Then
                        Throw New Exception("Please enter Morning Cut Off time for Route " & lblRouteDesc.Text & "")
                    End If
                    If clsCommon.myLen(strMorningShiftTime) > 0 Then
                        Morningvalue = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select datediff(minute,cast(MorningCutOff_Time as time),cast('" + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "hh:mm tt") + "' as time)) as aa from TSPL_ROUTE_MASTER where route_no='" & txtRouteNo.Value & "'"))
                        If Morningvalue > 0 Then
                            Throw New Exception("You cannot create Demand Booking due to Morning cut off is over")
                        End If
                    End If
                End If
                If dblQuantityEveningCount > 0 Then
                    If clsCommon.myLen(strEveningShiftTime) <= 0 Then
                        Throw New Exception("Please enter Evening Cut Off time for Route " & lblRouteDesc.Text & "")
                    End If
                    If clsCommon.myLen(strEveningShiftTime) > 0 Then
                        eveningvalue = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select datediff(minute,cast(EveningCutOff_Time as time),cast('" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "hh:mm tt") & "' as time)) as aa from TSPL_ROUTE_MASTER where route_no='" & txtRouteNo.Value & "'"))
                        If eveningvalue > 0 Then
                            Throw New Exception("You cannot create Demand Booking due to Evening cut off is over")
                        End If
                    End If
                End If
            End If
            '' to check cut off time on update 
            If Not chkIndividualCustomer.Checked AndAlso clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) > 0 AndAlso UseCutOffTimeonRouteForERP Then
                Morningvalue = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select datediff(minute,cast(MorningCutOff_Time as time),cast('" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "hh:mm tt") & "' as time)) as aa from TSPL_ROUTE_MASTER where route_no='" & txtRouteNo.Value & "'"))
                eveningvalue = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select datediff(minute,cast(EveningCutOff_Time as time),cast('" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "hh:mm tt") & "' as time)) as aa from TSPL_ROUTE_MASTER where route_no='" & txtRouteNo.Value & "'"))
                If Morningvalue > 0 OrElse eveningvalue > 0 Then
                    For ii As Integer = 0 To gv1.Rows.Count - 1
                        Dim strItemValueExist As String = clsCommon.myCstr(gv1.Rows(ii).Cells(colItemExist).Value)
                        If clsCommon.CompairString(strItemValueExist, "Yes") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colShiftName).Value), "Morning") = CompairStringResult.Equal Then
                                If Morningvalue > 0 Then
                                    Dim strDemandBooKingNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No from tspl_demand_booking_detail where document_no='" & clsCommon.myCstr(txtDocNo.Value) & "' and ShiftType='Morning' and Cust_code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colCustCode).Value) & "'"))
                                    If clsCommon.myLen(clsCommon.myCstr(strDemandBooKingNo)) <= 0 Then
                                        Throw New Exception("You cannot create Demand Booking for customer " & clsCommon.myCstr(gv1.Rows(ii).Cells(colCustName).Value) & " due to Morning cut off is over")
                                    End If
                                End If
                            End If
                            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(ii).Cells(colShiftName).Value), "Evening") = CompairStringResult.Equal AndAlso eveningvalue > 0 Then
                                'If eveningvalue > 0 Then
                                Dim strDemandBooKingNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No from tspl_demand_booking_detail where document_no='" & clsCommon.myCstr(txtDocNo.Value) & "' and ShiftType='Evening' and Cust_code='" & clsCommon.myCstr(gv1.Rows(ii).Cells(colCustCode).Value) & "'"))
                                If clsCommon.myLen(clsCommon.myCstr(strDemandBooKingNo)) <= 0 Then
                                    Throw New Exception("You cannot create Demand Booking for customer " & clsCommon.myCstr(gv1.Rows(ii).Cells(colCustName).Value) & " due to Evening cut off is over")
                                End If
                                'End If
                            End If
                        End If
                    Next
                End If
            End If
            '' to check gatepass or truck sheet generated
            If clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) > 0 Then
                Dim strDocNoForGatePassOrTrucksheetGeneratedMorning As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & txtDocNo.Value & "' and (IsGatePassGenerated='Y' or IsTruckSheetGenerated ='Y') and ShiftType='Evening' "))
                Dim strDocNoForGatePassOrTrucksheetGeneratedEvening As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & txtDocNo.Value & "' and (IsGatePassGenerated='Y' or IsTruckSheetGenerated ='Y') and ShiftType='Morning' "))
                If clsCommon.myLen(clsCommon.myCstr(strDocNoForGatePassOrTrucksheetGeneratedMorning)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(strDocNoForGatePassOrTrucksheetGeneratedEvening)) > 0 Then
                    Throw New Exception("Demand cannot be updated because its Gate Pass/Trucksheet has generated for both shifts.")
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                Dim RouteNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Route_No from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + txtDocNo.Value + "'"))
                If clsCommon.CompairString(txtRouteNo.Value, RouteNo) = CompairStringResult.Equal Then
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal AndAlso Not chkIndividualCustomer.Checked Then
                        'If Not chkIndividualCustomer.Checked Then
                        If rbtnEvening.IsChecked Then
                            Dim dutofTime As DateTime = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select EveningCutOff_Time from TSPL_ROUTE_MASTER where Route_No='" + txtRouteNo.Value + "'"))
                            Dim CurrDateTime = clsCommon.GETSERVERDATE()
                            If CurrDateTime.TimeOfDay < dutofTime.TimeOfDay Then
                                Throw New Exception("Update allow after Cutoff time [" & clsCommon.myCstr(dutofTime.TimeOfDay) & "]")
                            End If
                        Else
                            Dim dutofTime As DateTime = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select MorningCutOff_Time from TSPL_ROUTE_MASTER where Route_No='" + txtRouteNo.Value + "'"))
                            Dim CurrDateTime = clsCommon.GETSERVERDATE()
                            If CurrDateTime.TimeOfDay < dutofTime.TimeOfDay Then
                                Throw New Exception("Update allow after Cutoff time [" & clsCommon.myCstr(dutofTime.TimeOfDay) & "]")
                            End If
                        End If
                        'End If

                    End If
                    SaveData(False)
                Else
                    Throw New Exception("You can't change route")
                    txtRouteNo.Value = RouteNo
                End If
            Else
                SaveData(False)
            End If
            ' SaveData(0, False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Function SaveData(ByVal isQuickDemand As Boolean) As Boolean
        Try
            Dim qry As String = ""

            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                qry = "select IsPosting,IsUpdating,Posted,Curr_User from TSPL_DEMAND_BOOKING_MASTER where Document_No='" & txtDocNo.Value & "'"
                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 AndAlso clsCommon.myCdbl(dt1.Rows(0)("IsPosting")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("IsUpdating")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("Posted")) = 1 Then
                    'If clsCommon.myCdbl(dt1.Rows(0)("IsPosting")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("IsUpdating")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("Posted")) = 1 Then
                    Throw New Exception("Document in used by [" & clsCommon.myCstr(dt1.Rows(0)("Curr_User")) & "]")
                    'End If
                End If
                qry = "Update TSPL_DEMAND_BOOKING_MASTER set IsUpdating=1,Curr_User='" & objCommonVar.CurrentUser & "' where Document_No='" & txtDocNo.Value & "' "
                clsDBFuncationality.ExecuteNonQuery(qry)
            End If
            Dim lstCustItem As New List(Of clsDemandCustItem)
            lstCustItem = UpdateAllTotals(False)
            'UpdateAllTotals(False)


            blnSaveTotalQTy = True
            'BookingStatus = 0
            'Dim strPriceCode As String = String.Empty
            Dim LineNo As Integer = 1
            If (AllowToSave(Nothing)) Then
                If clsCommon.myLen(txtDocNo.Value) > 0 Then
                    ResetDemandOnSave(txtDocNo.Value)
                End If
                Dim obj As New clsDemandBookingSale()
                obj.Document_No = txtDocNo.Value
                obj.Document_Date = txtDate.Value
                obj.Location_Code = txtLocation.Value
                obj.Route_No = txtRouteNo.Value
                obj.City_Code = TxtCity.Value
                obj.TripNo = txtTripNo.Text
                If rbtnEvening.IsChecked Then
                    obj.ShiftType = "Evening"
                ElseIf rbtnMorning.IsChecked Then
                    obj.ShiftType = "Morning"
                Else
                    obj.ShiftType = "Both"
                End If
                If rbtn_Fresh.IsChecked Then
                    obj.ItemType = "Fresh"
                ElseIf rbtn_Ambient.IsChecked Then
                    obj.ItemType = "Ambient"
                Else
                    obj.ItemType = "Both"
                End If
                If chkIndividualCustomer.Checked Then
                    obj.IsIndividualCustomer = 1
                Else
                    obj.IsIndividualCustomer = 0
                End If
                obj.TotalQtyInCrates = clsCommon.myCdbl(lblTotalCrate.Text)
                obj.TotalQtyInLtr = clsCommon.myCdbl(lblTotalLitre.Text)
                obj.DocumentAmount = clsCommon.myCdbl(lblDocumentAmt.Text)
                obj.NoCrateIssue = IIf(chkNoCrateIssue.Checked, 1, 0)
                obj.Arr = New List(Of clsDemandBookingSaleDetail)
                ''richa 4 Aug,2021 optimization related
                'Dim intLine As Integer = 0
                For dblrows As Integer = 0 To gv1.Rows.Count - 1
                    If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colItemExist).Value), "Yes") = CompairStringResult.Equal AndAlso clsCommon.myLen(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value)) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colShiftName).Value), "Morning") = CompairStringResult.Equal AndAlso chkMorningPosted.Checked Then
                            Continue For
                        End If
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colShiftName).Value), "Evening") = CompairStringResult.Equal AndAlso chkEveningPosted.Checked Then
                            Continue For
                        End If
                        Dim k As Integer = 1
                        Dim isCustRouteNotChanged As Boolean = True
                        For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                            Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode & clsCommon.myCstr(k)).Tag, ItemValueClass)
                            k = k + 1
                            If obj1 IsNot Nothing AndAlso (clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0) AndAlso isCustRouteNotChanged Then
                                'If (clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0) AndAlso isCustRouteNotChanged Then
                                Dim objTr As New clsDemandBookingSaleDetail()
                                objTr.Line_No = LineNo
                                objTr.Trip_No = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colTripNo).Value)
                                objTr.Cust_Code = clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value)
                                objTr.Created_By = clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCreated_By).Value)
                                objTr.Source_By = clsCommon.myCstr(gv1.Rows(dblrows).Cells(colSourceBy).Value)
                                objTr.REF_PK_ID = clsCommon.myCstr(gv1.Rows(dblrows).Cells(colREF_PK_ID).Value)
                                objTr.ShiftType = clsCommon.myCstr(gv1.Rows(dblrows).Cells(colShiftName).Value)
                                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colIsItemUpdate).Value), "Yes") = CompairStringResult.Equal Then
                                    objTr.IsItemUpdate = 1
                                Else
                                    objTr.IsItemUpdate = 0
                                End If
                                If clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 Then
                                    If clsCommon.CompairString(objTr.Cust_Code, clsDBFuncationality.getSingleValue("Select Cust_Code from TSPL_CUSTOMER_MASTER where cust_code='" + clsCommon.myCstr(objTr.Cust_Code) + "' and Status='N'")) = CompairStringResult.Equal Then
                                        objTr.Qty = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                    Else
                                        objTr.Qty = 0
                                    End If
                                    'End If
                                    For Each obj2 As clsDemandCustItem In lstCustItem
                                        If clsCommon.CompairString(objTr.Cust_Code, obj2.Cust_Code) = CompairStringResult.Equal AndAlso (clsCommon.CompairString(obj1.itemCode, obj2.ItemCode) = CompairStringResult.Equal AndAlso clsCommon.CompairString(obj1.Unit_code, obj2.UnitCode) = CompairStringResult.Equal) Then
                                            objTr.Item_Code = clsCommon.myCstr(obj1.itemCode)
                                            objTr.Unit_code = clsCommon.myCstr(obj1.Unit_code)
                                            objTr.Rate = clsCommon.myCdbl(obj2.ItemRate)
                                            objTr.ItemNetAmount = clsCommon.myCDecimal(obj2.ItemTotAmt)
                                            objTr.TAX_Group = clsCommon.myCstr(obj2.TAX_Group)
                                            objTr.TAX1 = clsCommon.myCstr(obj2.TAX1)
                                            objTr.TAX1_Rate = obj2.TAX1_Rate
                                            objTr.TAX1_Amt = obj2.TAX1_Amt

                                            objTr.TAX1_Base_Amt = clsCommon.myCdbl(obj2.TAX1_Base_Amt)
                                            objTr.TAX2 = clsCommon.myCstr(obj2.TAX2)
                                            objTr.TAX2_Rate = clsCommon.myCdbl(obj2.TAX2_Rate)
                                            objTr.TAX2_Amt = clsCommon.myCdbl(obj2.TAX2_Amt)

                                            objTr.TAX2_Base_Amt = clsCommon.myCdbl(obj2.TAX2_Base_Amt)
                                            objTr.TAX3 = clsCommon.myCstr(obj2.TAX3)
                                            objTr.TAX3_Rate = clsCommon.myCdbl(obj2.TAX3_Rate)
                                            objTr.TAX3_Amt = clsCommon.myCdbl(obj2.TAX3_Amt)

                                            objTr.TAX3_Base_Amt = clsCommon.myCdbl(obj2.TAX3_Base_Amt)
                                            objTr.TAX4 = clsCommon.myCstr(obj2.TAX4)
                                            objTr.TAX4_Rate = clsCommon.myCdbl(obj2.TAX4_Rate)
                                            objTr.TAX4_Amt = clsCommon.myCdbl(obj2.TAX4_Amt)

                                            objTr.TAX4_Base_Amt = clsCommon.myCdbl(obj2.TAX4_Base_Amt)
                                            objTr.TAX5 = clsCommon.myCstr(obj2.TAX5)
                                            objTr.TAX5_Rate = clsCommon.myCdbl(obj2.TAX5_Rate)
                                            objTr.TAX5_Amt = clsCommon.myCdbl(obj2.TAX5_Amt)
                                            objTr.TAX5_Base_Amt = clsCommon.myCdbl(obj2.TAX5_Base_Amt)
                                            objTr.TAX6 = clsCommon.myCstr(obj2.TAX6)
                                            objTr.TAX6_Rate = clsCommon.myCdbl(obj2.TAX6_Rate)
                                            objTr.TAX6_Amt = clsCommon.myCdbl(obj2.TAX6_Amt)
                                            objTr.TAX6_Base_Amt = clsCommon.myCdbl(obj2.TAX6_Base_Amt)
                                            objTr.TAX7 = clsCommon.myCstr(obj2.TAX7)
                                            objTr.TAX7_Rate = clsCommon.myCdbl(obj2.TAX7_Rate)
                                            objTr.TAX7_Amt = clsCommon.myCdbl(obj2.TAX7_Amt)
                                            objTr.TAX7_Base_Amt = clsCommon.myCdbl(obj2.TAX7_Base_Amt)
                                            objTr.TAX8 = clsCommon.myCstr(obj2.TAX8)
                                            objTr.TAX8_Rate = clsCommon.myCdbl(obj2.TAX8_Rate)
                                            objTr.TAX8_Amt = clsCommon.myCdbl(obj2.TAX8_Amt)
                                            objTr.TAX8_Base_Amt = clsCommon.myCdbl(obj2.TAX8_Base_Amt)
                                            objTr.TAX9 = clsCommon.myCstr(obj2.TAX9)
                                            objTr.TAX9_Rate = clsCommon.myCdbl(obj2.TAX9_Rate)
                                            objTr.TAX9_Amt = clsCommon.myCdbl(obj2.TAX9_Amt)
                                            objTr.TAX9_Base_Amt = clsCommon.myCdbl(obj2.TAX9_Base_Amt)
                                            objTr.TAX10 = clsCommon.myCstr(obj2.TAX10)
                                            objTr.TAX10_Rate = clsCommon.myCdbl(obj2.TAX10_Rate)
                                            objTr.TAX10_Amt = clsCommon.myCdbl(obj2.TAX10_Amt)
                                            objTr.TAX10_Base_Amt = clsCommon.myCdbl(obj2.TAX10_Base_Amt)
                                            If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                                                If clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code), "Crate") = CompairStringResult.Equal Then
                                                    objTr.TotalCrates_ItemWise = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                                Else
                                                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                                                    If ItemCrateType = 1 Then
                                                        objTr.TotalCrates_ItemWise = clsCommon.myCdbl(obj2.FreshItem_QtyInCrate)
                                                    End If
                                                End If
                                                objTr.TotalLtr_ItemWise = clsCommon.myCdbl(obj2.FreshItem_QtyInLitres)
                                            End If

                                        End If
                                    Next

                                    objTr.Vehicle_Code = clsCommon.myCstr(txtVehicleNo.Value)

                                    qry = "select Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_ROUTE_MASTER.Zone_Code,TSPL_CUSTOMER_MASTER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER left outer join " &
                           "TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join TSPL_VEHICLE_MASTER on " &
                           "TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id where Cust_Code='" & clsCommon.myCstr(objTr.Cust_Code) & "'"
                                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
                                    If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                                        objTr.Price_Code = clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))
                                        If clsCommon.myLen(clsCommon.myCstr(dt1.Rows(0)("Route_No"))) <= 0 Then
                                            Throw New Exception("Plz tag Route for this Customer " & clsCommon.myCstr(objTr.Cust_Code) & "")
                                        End If
                                        'If clsCommon.myLen(clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))) <= 0 Then
                                        '    Throw New Exception("Plz tag Vehicle for this route " & clsCommon.myCstr(dt1.Rows(0)("Route_No")) & "")
                                        'End If
                                        If clsCommon.myLen(clsCommon.myCstr(dt1.Rows(0)("price_CodeNon"))) <= 0 Then
                                            Throw New Exception("Plz tag price_CodeNon for this Customer " & clsCommon.myCstr(objTr.Cust_Code) & "")
                                        End If
                                    End If
                                    If (clsCommon.myLen(objTr.Cust_Code) > 0) AndAlso (clsCommon.myLen(objTr.Item_Code) > 0) Then
                                        obj.Arr.Add(objTr)
                                    End If
                                    LineNo = LineNo + 1
                                End If
                                'End If
                            End If
                        Next
                    End If
                Next
                If (obj.SaveData(obj, isNewEntry)) Then
                    If Not isQuickDemand Then
                        clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
                        LoadData(obj.Document_No, NavigatorType.Current)
                    End If
                    Return True
                End If
            Else
                Return False
            End If
            blnSaveTotalQTy = True
        Catch ex As Exception
            blnSaveTotalQTy = True
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        Finally
            Dim qry As String = "Update TSPL_DEMAND_BOOKING_MASTER set IsUpdating=0,Curr_User= null where Document_No='" & txtDocNo.Value & "' "
            clsDBFuncationality.ExecuteNonQuery(qry)
        End Try
        Return False
    End Function
    Sub GatePass_TruckSheet_Button()
        Try
            If rbtnMorningEveningBoth.IsChecked Then
                btn_TruckSheet.Enabled = False
                SplitButtonTruckSheet.Enabled = False
                btn_Gatepass.Enabled = False
                btn_TSCancel.Enabled = False
                btn_GPCancel.Enabled = False
                btnPrint.Enabled = False
            Else
                Dim strDocNoForGatePass As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & txtDocNo.Value & "' and ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "' and IsGatePassGenerated='Y'"))
                Dim strDocNoForTrucksheet As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & txtDocNo.Value & "' and ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "' and  IsTruckSheetGenerated ='Y'"))
                If clsCommon.myLen(strDocNoForGatePass) > 0 Then
                    btn_Gatepass.Enabled = False
                    btn_GPCancel.Enabled = True
                Else
                    btn_Gatepass.Enabled = True
                    btn_GPCancel.Enabled = False
                End If
                If clsCommon.myLen(strDocNoForTrucksheet) > 0 Then
                    btn_TruckSheet.Enabled = False
                    SplitButtonTruckSheet.Enabled = False
                    btn_TSCancel.Enabled = True
                Else
                    btn_TruckSheet.Enabled = True
                    SplitButtonTruckSheet.Enabled = True
                    btn_TSCancel.Enabled = False
                End If
                btnPrint.Enabled = True
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        Try
            Dim isLoadData As Boolean = True
            'Dim dblTotalDocAmt As Decimal = 0
            'Dim qry As String = ""
            Dim obj As New clsDemandBookingSale
            'Dim intRow As Integer
            obj = clsDemandBookingSale.GetData(strCode, NavTyep, Nothing)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
                AddNew()
                btnSave.Enabled = True
                btnPost.Enabled = True
                btnDelete.Enabled = True
                isInsideLoadData = True
                txtDate.Enabled = False
                chkNoCrateIssue.Enabled = False
                isNewEntry = False
                btn_TSCancel.Enabled = True
                btn_Gatepass.Enabled = True
                btnSave.Text = "Update"
                txtDocNo.Value = obj.Document_No
                txtDate.Value = obj.Document_Date
                txtLocation.Value = obj.Location_Code
                TxtCity.Value = obj.City_Code
                txtRouteNo.Value = obj.Route_No
                txtTripNo.Text = obj.TripNo
                chkNoCrateIssue.Checked = IIf(obj.NoCrateIssue = 1, True, False)
                If obj.IsIndividualCustomer = 1 Then
                    chkIndividualCustomer.Checked = True
                    txtCustomerNo.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT top 1 cust_code  FROM TSPL_DEMAND_BOOKING_DETAIL where document_no='" & txtDocNo.Value & "' "))
                    lblCustomerName.Text = clsCommon.myCstr(clsCustomerMaster.GetName(txtCustomerNo.Value, Nothing))
                Else
                    chkIndividualCustomer.Checked = False
                End If
                chkIndividualCustomer.Enabled = False
                If obj.Posted = 1 Then
                    btnSave.Enabled = False
                    btnDelete.Enabled = False
                    btnPost.Enabled = False
                    UsLock1.Status = ERPTransactionStatus.Approved
                ElseIf obj.Posted = 2 Then
                    btnSave.Enabled = False
                    btnDelete.Enabled = False
                    btnPost.Enabled = False
                    UsLock1.Status = ERPTransactionStatus.Cancel
                Else
                    UsLock1.Status = ERPTransactionStatus.Pending
                End If
                checkPrintSetting()
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    txtDemandUniqueID.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Demand_UniqueID from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + txtDocNo.Value + "'"))
                End If
                lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtLocation.Value + "'"))
                lblRouteDesc.Text = clsRouteMaster.GetName(txtRouteNo.Value, Nothing)
                lblCityName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select City_name from tspl_city_master where city_code=  '" & TxtCity.Value & "' "))
                lblTotalCrate.Text = obj.TotalQtyInCrates
                lblTotalLitre.Text = obj.TotalQtyInLtr
                lblDocumentAmt.Text = obj.DocumentAmount
                If clsCommon.CompairString(obj.ShiftType, "Morning") = CompairStringResult.Equal Then
                    rbtnMorning.IsChecked = True
                ElseIf clsCommon.CompairString(obj.ShiftType, "Evening") = CompairStringResult.Equal Then
                    rbtnEvening.IsChecked = True
                Else
                    rbtnMorningEveningBoth.IsChecked = True
                End If
                If SettSeprateDemandForMorningEveningShift Then
                    RadGroupBox3.Enabled = False
                End If
                If clsCommon.CompairString(obj.ItemType, "Fresh") = CompairStringResult.Equal Then
                    rbtn_Fresh.IsChecked = True
                    If SeparateDemandMilkandProduct Then
                        RadGroupBox1.Enabled = False
                        'LoadBlankGrid()
                        HideUnhideRowsAndColumnsOFGrid()
                    End If
                ElseIf clsCommon.CompairString(obj.ItemType, "Ambient") = CompairStringResult.Equal Then
                    rbtn_Ambient.IsChecked = True
                    If SeparateDemandMilkandProduct Then
                        RadGroupBox1.Enabled = False
                        'LoadBlankGrid()
                        HideUnhideRowsAndColumnsOFGrid()
                    End If
                Else
                    rdbnFreshAmbientBoth.IsChecked = True
                    If SeparateDemandMilkandProduct Then
                        RadGroupBox1.Enabled = False
                        'LoadBlankGrid()
                        HideUnhideRowsAndColumnsOFGrid()
                    End If
                End If
                If Not SeparateDemandMilkandProduct Then
                    'LoadBlankGrid()
                End If
                setCustomerDetail(TxtCity.Value, txtRouteNo.Value, True)
                chkMorningPosted.Checked = (obj.Posted_Morning = 1)
                chkEveningPosted.Checked = (obj.Posted_Evening = 1)
                Dim dblMorningCount As Integer = 0
                Dim dblEveningCount As Integer = 0
                isLoadData = True
                If obj.Arr IsNot Nothing AndAlso obj.Arr.Count > 0 Then
                    For Each objTr As clsDemandBookingSaleDetail In obj.Arr
                        For dblrows As Integer = 0 To gv1.Rows.Count - 1
                            If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value), objTr.Cust_Code) = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colShiftName).Value), objTr.ShiftType) = CompairStringResult.Equal Then
                                gv1.Rows(dblrows).Cells(colCreated_By).Value = objTr.Created_By
                                gv1.Rows(dblrows).Cells(colSourceBy).Value = objTr.Source_By
                                gv1.Rows(dblrows).Cells(colREF_PK_ID).Value = objTr.REF_PK_ID
                                Dim k As Integer = 1
                                For columns = 9 To gv1.Columns.Count - 11
                                    Dim obj1 As ItemValueClass = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                                    k = k + 1
                                    If clsCommon.CompairString(objTr.Item_Code, clsCommon.myCstr(obj1.itemCode)) = CompairStringResult.Equal AndAlso clsCommon.CompairString(objTr.Unit_code, clsCommon.myCstr(obj1.Unit_code)) = CompairStringResult.Equal AndAlso clsCommon.CompairString(objTr.ShiftType, "Morning") = CompairStringResult.Equal Then
                                        gv1.Rows(dblrows).Cells(columns).Value = objTr.Qty
                                        gv1.Rows(dblrows).Cells(colTripNo).Value = objTr.Trip_No
                                        dblMorningCount = 1
                                    ElseIf clsCommon.CompairString(objTr.Item_Code, clsCommon.myCstr(obj1.itemCode)) = CompairStringResult.Equal AndAlso clsCommon.CompairString(objTr.Unit_code, clsCommon.myCstr(obj1.Unit_code)) = CompairStringResult.Equal AndAlso clsCommon.CompairString(objTr.ShiftType, "Evening") = CompairStringResult.Equal Then
                                        gv1.Rows(dblrows).Cells(columns).Value = objTr.Qty
                                        gv1.Rows(dblrows).Cells(colTripNo).Value = objTr.Trip_No
                                        dblEveningCount = 1
                                    End If
                                Next


                            End If
                        Next
                        'strCustCode = objTr.Cust_Code
                        If isLoadData Then
                            txtVehicleNo.Value = objTr.Vehicle_Code
                            lblVehicleNo.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Description  from TSPL_VEHICLE_MASTER where Vehicle_Id = '" & Convert.ToString(txtVehicleNo.Value) & "'"))
                            isLoadData = False
                        End If
                        If clsCommon.CompairString(objTr.ShiftType, "Evening") = CompairStringResult.Equal AndAlso clsCommon.CompairString(objTr.IsGatePassGenerated, "Y") = CompairStringResult.Equal OrElse clsCommon.CompairString(objTr.IsTruckSheetGenerated, "Y") = CompairStringResult.Equal Then
                            'If clsCommon.CompairString(objTr.IsGatePassGenerated, "Y") = CompairStringResult.Equal OrElse clsCommon.CompairString(objTr.IsTruckSheetGenerated, "Y") = CompairStringResult.Equal Then
                            chkEveningGatepassTruckSheetGenerated.Checked = True
                            'End If
                        End If
                        If clsCommon.CompairString(objTr.ShiftType, "Morning") = CompairStringResult.Equal AndAlso clsCommon.CompairString(objTr.IsGatePassGenerated, "Y") = CompairStringResult.Equal OrElse clsCommon.CompairString(objTr.IsTruckSheetGenerated, "Y") = CompairStringResult.Equal Then
                            'If clsCommon.CompairString(objTr.IsGatePassGenerated, "Y") = CompairStringResult.Equal OrElse clsCommon.CompairString(objTr.IsTruckSheetGenerated, "Y") = CompairStringResult.Equal Then
                            chkMorningGatepassTruckSheetGenerated.Checked = True
                            'End If
                        End If
                    Next
                End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal AndAlso Not chkIndividualCustomer.Checked Then
                    Dim dtTransporter As DataTable = clsDBFuncationality.GetDataTable(ReturnDistributorRouteTaggingQry())
                    If dtTransporter IsNot Nothing AndAlso dtTransporter.Rows.Count > 0 Then
                        lblTransporterName.Text = dtTransporter.Rows(0)("Customer_Name")
                    End If
                    dtTransporter = Nothing
                Else
                    lblTransporterName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select transporter_name from tspl_transport_master where Transport_Id =(select Transport_Id from tspl_vehicle_master where vehicle_id= '" & Convert.ToString(txtVehicleNo.Value) & "')"))
                End If
                'HideUnhideRowsOFGrid()
                isLoadData = False
                UpdateAllTotals(True)

                If Not SettSeprateDemandForMorningEveningShift Then
                    HideUnhideRowsAndColumnsOFGrid()
                End If
            End If
            SetRouteColumns()
            RefreshFormName()
            If DisableRouteandVehicle Then
                txtRouteNo.Enabled = False
                txtVehicleNo.Enabled = False
            Else
                txtRouteNo.Enabled = True
                txtVehicleNo.Enabled = True
            End If
            isFreshAmbientBoth()
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            isInsideLoadData = False
            isLoadData = False
        End Try
    End Sub
    Private Sub txtDocNo__MYNavigator(ByVal sender As Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtDocNo._MYNavigator
        Try
            Dim qry As String = "select count(*) from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + txtDocNo.Value + "' and comp_code='" + objCommonVar.CurrentCompanyCode + "'  "
            Dim check As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry))
            If check > 0 Then
                txtDocNo.MyReadOnly = True
            Else
                txtDocNo.MyReadOnly = False
            End If
            LoadData(txtDocNo.Value, NavType)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtDocNo__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtDocNo._MYValidating
        Try
            Dim qry As String = "select TSPL_DEMAND_BOOKING_MASTER.Document_No as DocumentNo,convert(varchar(12),TSPL_DEMAND_BOOKING_MASTER.Document_date,103) as DocumentDate,TSPL_DEMAND_BOOKING_MASTER.ShiftType,TSPL_DEMAND_BOOKING_MASTER.Route_No as [Route No],TSPL_DEMAND_BOOKING_MASTER.Location_Code as [Location Code],TSPL_DEMAND_BOOKING_MASTER.City_Code as [City Code],TripNo AS [Trip No],TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer as [Individual Cust],case when Posted=1 then 'posted' else 'Unposted' end as Posted from TSPL_DEMAND_BOOKING_MASTER "
            'Dim whrClas As String = " TSPL_DEMAND_BOOKING_MASTER.comp_code='" + objCommonVar.CurrentCompanyCode + "' "
            Reset()
            LoadData(clsCommon.ShowSelectForm("FSBook1DocNo", qry, "DocumentNo", "", txtDocNo.Value, "Document_date DESC", isButtonClicked, " TSPL_DEMAND_BOOKING_MASTER.Document_date "), NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        CloseData()
    End Sub
    Sub CloseData()
        clsERPFuncationality.closeForm(Me)
    End Sub
    Private Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            DeleteData()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
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
                If (clsDemandBookingSale.DeleteData(txtDocNo.Value)) Then
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
    Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                LoadData(txtDocNo.Value, NavigatorType.Current)
                Export()
            Else
                Throw New Exception("Document not found!")
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub Export()
        If gv1.Rows.Count > 0 Then
            ExportToExcel()
        Else
            common.clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
        End If
    End Sub
    Private Sub ExportToExcel()
        Try
            Dim strCreatedBy As String = clsDBFuncationality.getSingleValue("Select Created_By from TSPL_BOOKING_MATSER where Document_No='" + txtDocNo.Value + "'")
            Dim arrHeader As List(Of String) = New List(Of String)()
            Dim strtemp As String = "Booking No : " & txtDocNo.Value
            arrHeader.Add(strtemp)
            strtemp = "Booking Date : " & clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy")
            arrHeader.Add(strtemp)
            strtemp = "Created By : " & strCreatedBy
            arrHeader.Add(strtemp)
            strtemp = "Transaction Type : " & IIf(rbtn_Fresh.IsChecked, "Fresh", "Ambient")
            arrHeader.Add(strtemp)
            clsCommon.MyExportToExcelGrid("Booking Entry", gv1, arrHeader, Me.Text)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub
    Private Shared Function CreateNotificationContentEMP(ByVal Booking_Id As String, ByVal Booking_Date As DateTime, ByVal Ex_Factory_Date As DateTime, ByVal trans As SqlTransaction) As Boolean
        Dim strNotifiContent As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_Text from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'", trans))
        Dim strNotifi_DetalContent As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_Detail_Text from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'", trans))
        Dim strNotifiCaption As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_Caption from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'", trans))
        Dim strNotificationOn As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Notification_On from TSPL_ES_Content where Form_ID='" & clsUserMgtCode.frmbookingdairy & "'", trans))
        'Dim strNotification_Condition_Date As DateTime = Nothing
        'Dim strdate As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT Condition_Date from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmbookingdairy + "'", trans))
        'If clsCommon.myLen(strdate) > 0 Then
        '    strNotification_Condition_Date = clsCommon.myCDate(clsDBFuncationality.getSingleValue("SELECT Condition_Date from TSPL_ES_Content where Form_ID='" + clsUserMgtCode.frmbookingdairy + "'", trans))
        'End If
        If clsCommon.myLen(strNotifiContent) > 0 Then
            Dim objNotification As New clsNotificationHead()
            objNotification.Notification_Text = strNotifiContent
            objNotification.Notification_Caption = strNotifiCaption
            objNotification.Notification_On = strNotificationOn
            objNotification.Notification_Detail_Text = strNotifi_DetalContent
            objNotification.Notification_ConditionDate = clsCommon.myCDate(Ex_Factory_Date)
            objNotification.Notification_Text = objNotification.Notification_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_No, clsCommon.myCstr(Booking_Id))
            objNotification.Notification_Text = objNotification.Notification_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Doc_Date, clsCommon.myCstr(Booking_Date))
            objNotification.Notification_Text = objNotification.Notification_Text.Replace(XpertERPEngine.frmEMailAndSMSSetting.Ex_Factory_Date, clsCommon.myCstr(clsCommon.myCDate(Ex_Factory_Date)))
            objNotification.SaveData(clsUserMgtCode.frmbookingdairy, objNotification, trans)
            objNotification = Nothing
            Return True
        End If
        Return False
    End Function
    Private Sub btnLayout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLayout.Click
        clsGridLayout.DeleteData(ReportID, objCommonVar.CurrentUserCode)
    End Sub
    Private Sub btnSaveLayout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveLayout.Click
        If clsCommon.myLen(ReportID) > 0 Then
            gv1.MasterTemplate.FilterDescriptors.Clear()
            Dim obj As New clsGridLayout()
            obj.ReportID = ReportID
            obj.UserID = objCommonVar.CurrentUserCode
            obj.GridLayout = New MemoryStream()
            gv1.SaveLayout(obj.GridLayout)
            obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
            obj.GridColumns = gv1.ColumnCount
            If obj.SaveData() Then
                common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", "Information", Me.Text)
            End If
            ''stuti regarding memory leakage
            obj.GridLayout.Close()
            obj.GridLayout.Dispose()
        End If
    End Sub
    Private Sub gv1_Pasting(sender As Object, e As GridViewClipboardEventArgs) Handles gv1.Pasting
        FlagPaste = True
    End Sub
    Private Sub txtLocation__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtLocation._MYValidating
        Dim qry As String = " select Location_Code as [Code],Location_Desc as [Description],Loc_Short_Name as [Short Name],Add1,Add2,Add3,Add4,City_Code as [City Code],State,Pin_Code as [Pin Code],Country,Telphone,Email,Location_Type as [Location Type],Loc_Status as [Location Status],Status_Date as [Status Date],Excisable,Loc_Segment_Code as [Location Segment Code],Type,Purchase_Tax_Group as [Purchase Tax Group],Sales_Tax_Group as [Sales Tax Group],Ecc_Number as [ECC Number],Registration_Number as [Registration Number],Commissionerate as [Commission Rate],Range_Code as [Range Code],Range_Name as [Range Name],Range_Address as [Range Address],Division_Code as [Division Code],Division_Name as [Division Name],Division_Address as [Division Address],Created_By as [Created By],Created_Date as [Created Date],Modify_By as [Modify By],Modify_Date as [Modify Date],Comp_code as [Company Code],TIN_No as [TIN No],TAN_No as [TAN No],TCAN_No as [TCAN No],Service_Tax_Reg_No as [Service Tax Registration No],DutyPaid as [Duty Paid],Purchase_Tax_GroupIS as [Purchase Tax Group Inter State],Sales_Tax_GroupIS as [Sales Tax Group Inter State],Stock_Transfer_Filled_Ac as [Stock Transfer Filled Account],Stock_Transfer_Empty_Ac as [Stock Transfer Empty Account],GIT_Location as [GIT Location],GIT_Type as [GIT Type],Rejected_Type as [Rejected Type],Rejected_Location as [Rejected Location],CSA_Type as [CSA Type],Cust_Code as [Cust Code],MCC_Type as [MCC Type],CST_No as [CST No],Phone1,Phone2  from TSPL_Location_MASTER"
        Dim WhrCls As String = " Loc_Status='N' and Location_Type='Physical' and Is_Section='N' and Is_Sub_Location='N' and CSA_Type <>'Y' and DutyPaid <>'Y' and Rejected_Type <>'Y' and GIT_Type<>'Y'"
        If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
            WhrCls += "  and  Location_Code in (" & objCommonVar.strCurrUserLocations & ")"
        End If
        txtLocation.Value = clsCommon.ShowSelectForm("MulDS-BOLocFndr", qry, "Code", WhrCls, txtLocation.Value, "Code", isButtonClicked)
        lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" & txtLocation.Value & "'"))
    End Sub
    Public Sub funPrint(ByVal strDocNo As String)
        Try
            Dim Qry As String = GetQuery()
            Qry += " where TSPL_BOOKING_MATSER.Document_No ='" & strDocNo & "'"
            'Qry = "Select * from (" & Qry & ") XXX LEFT OUTER JOIN (Select '1' as COL1, 1 as COL2,  'ORIGINAL COPY' as CopyType1 UNION Select '1' as COL1, 2 as COL2,  'DUPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 3 as COL2,  'TRIPLICATE COPY' as CopyType1 UNION Select '1' as COL1, 4 as COL2,  'QUADRUPLICATE COPY' as CopyType1) YYY ON YYY.COL1=XXX.CopyType ORDER BY YYY.COL2"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDSBoking", "Performa Invoice", "rptCompanyAddress.rpt")
                frmCRV = Nothing
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Function GetQuery() As String
        Dim Qry As String = " select TSPL_COMPANY_MASTER.CINNo as CompCinNo,TSPL_COMPANY_MASTER.logo_Img,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP,TSPL_COMPANY_MASTER.CE_Division as GP_Division,TSPL_COMPANY_MASTER.ServiceTax_Reg_No +case when len(TSPL_COMPANY_MASTER.Ecc_No)>0 then ''+TSPL_COMPANY_MASTER.Ecc_No else '' end as GP_ECC_No,TSPL_COMPANY_MASTER.CE_Range as GP_CE_Range,TSPL_COMPANY_MASTER.Comp_Name AS CompName ," &
                            " TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ' , '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ' , '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when LEN(TSPL_COMPANY_MASTER.tin_no)>0 then ' , TIN No '+TSPL_COMPANY_MASTER.tin_no else ' ' end as Comp_Add_GP," &
                            " TSPL_LOCATION_MASTER.TIN_No as Loc_Tin_No,TSPL_LOCATION_MASTER.add1 +case when len(TSPL_LOCATION_MASTER.add2)>0 then ', '+TSPL_LOCATION_MASTER.add2 else '' end +case when LEN(isnull(TSPL_LOCATION_MASTER.Add3,''))>0 then ', '+isnull(TSPL_LOCATION_MASTER.Add3,'') else ' ' end  + case when len(TSPL_STATE_MASTER.STATE_NAME   )>0 then ', '+ TSPL_STATE_MASTER.STATE_NAME  else ' ' end    as Location_Address_GP, " &
                            " TSPL_BOOKING_DETAIL.item_code,tspl_item_master.item_desc,TSPL_BOOKING_DETAIL.booking_Qty,TSPL_BOOKING_DETAIL.unit_Code,TSPL_BOOKING_DETAIL.item_Rate,TSPL_BOOKING_DETAIL.documentAmount,TSPL_BOOKING_DETAIL.vehicle_code,tspl_vehicle_master.Number as Vechicle_Name,TSPL_BOOKING_MATSER.Document_No,TSPL_BOOKING_MATSER.Document_Date,TSPL_BOOKING_DETAIL.cust_code,TSPL_CUSTOMER_MASTER.customer_name " &
                            " from TSPL_BOOKING_DETAIL  left join TSPL_BOOKING_MATSER on TSPL_BOOKING_MATSER.Document_No =TSPL_BOOKING_DETAIL.Document_No " &
                            " left outer join TSPL_COMPANY_MASTER on  tspl_company_Master.Comp_Code = TSPL_BOOKING_MATSER.comp_code " &
                            " LEFT OUTER JOIN TSPL_CITY_MASTER  AS TSPL_CITY_MASTER_fOR_Comp ON TSPL_CITY_MASTER_fOR_Comp.City_Code =TSPL_COMPANY_MASTER.City_Code " &
                            " LEFT OUTER JOIN TSPL_STATE_MASTER AS TSPL_STATE_MASTER_For_Comp  ON TSPL_STATE_MASTER_For_Comp.STATE_CODE  =TSPL_COMPANY_MASTER.State " &
                            " left join tspl_location_master on TSPL_LOCATION_MASTER .Location_Code =TSPL_BOOKING_MATSER.location_code " &
                            " left join TSPL_STATE_MASTER on TSPL_STATE_MASTER .STATE_CODE =tspl_location_master.State " &
                            " left join tspl_item_master on tspl_item_master.item_code=TSPL_BOOKING_DETAIL.item_code " &
                            " left join tspl_vehicle_master on tspl_vehicle_master.vehicle_id=TSPL_BOOKING_DETAIL.vehicle_code" &
                            " left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER .cust_code=TSPL_BOOKING_DETAIL.cust_code "
        Return Qry
    End Function
    Private Sub ddlTaxType_SelectedIndexChanged(sender As Object, e As UI.Data.PositionChangedEventArgs)
        If Not isInsideLoadData Then
            'AddNew()
            'ddlTaxType.Text = "Select"
            txtLocation.Value = Nothing
            lblLocation.Text = ""
            gv1.DataSource = Nothing
        End If
    End Sub
    Private Sub btnreverse_Click(sender As Object, e As EventArgs) Handles btnreverse.Click
        Reverse()
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        'Dim NextDayDocNo As String = ""

        'Try
        '    If clsCommon.myLen(txtDocNo.Value) > 0 Then
        '        Dim isDispatch As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select COUNT(*) from TSPL_SD_SHIPMENT_BOOKING_DETAIL where Booking_TR_Code in(select TR_Code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + txtDocNo.Value + "')"))
        '        If isDispatch >= 1 Then
        '            Throw New Exception("Dispatch already Created!")
        '        End If
        '    Else
        '        Throw New Exception("Please Select Document")

        '    End If
        '    If objCommonVar.ApplyBoothRouteMapping Then
        '        NextDayDocNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 103 )='" + clsCommon.GetPrintDate(txtDate.Value.AddDays(1)) + "') and location_code='" + clsCommon.myCstr(txtLocation.Value) + "' and ShiftType='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' and IsIndividualCustomer=0 ")
        '        If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document " + IIf(clsCommon.myLen(NextDayDocNo) > 0, "and Delete Next Day Document [" + NextDayDocNo + "]", "") + " " + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
        '            Dim Reason As String = ""
        '            Dim qry As String = ""
        '            Dim frm As New FrmFreeTxtBox1
        '            frm.Text = "Remarks for Reverse"
        '            frm.ShowDialog()
        '            If clsCommon.myLen(frm.strRmks) <= 0 Then
        '                Exit Sub
        '            Else
        '                Reason = frm.strRmks
        '            End If
        '            If clsDemandBookingSale.ReverseMultipleDOC(txtDocNo.Value, txtRouteNo.Value, clsCommon.GetPrintDate(txtDate.Value.AddDays(1)), txtLocation.Value, IIf(rbtnMorning.IsChecked, "Morning", "Evening")) Then
        '                saveCancelLog(Reason, "Reverse And Recreate", Nothing)
        '                common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
        '                LoadData(txtDocNo.Value, NavigatorType.Current)
        '            End If
        '        End If

        '    Else
        '        If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, Nothing)) = 1 Or clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, Nothing)) = 1 Then
        '            NextDayDocNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 103 )='" + clsCommon.GetPrintDate(txtDate.Value.AddDays(1)) + "') and location_code='" + clsCommon.myCstr(txtLocation.Value) + "' and ShiftType='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' and IsIndividualCustomer=0 ")
        '        End If
        '        'Dim dt As DataTable = clsDBFuncationality.GetDataTable("select location_code,Document_Date from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + NextDayDocNo + "'", "")
        '        'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '        'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Dairy Sale", "Demand Booking", txtLocation.Value, txtDate.Value, trans)
        '        'End If
        '        If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document " + IIf(clsCommon.myLen(NextDayDocNo) > 0, "and Delete Next Day Document [" + NextDayDocNo + "]", "") + " " + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
        '            '' REASON FOR DELETE 
        '            Dim Reason As String = ""
        '            Dim qry As String = ""
        '            Dim frm As New FrmFreeTxtBox1
        '            frm.Text = "Remarks for Reverse"
        '            frm.ShowDialog()
        '            If clsCommon.myLen(frm.strRmks) <= 0 Then
        '                Exit Sub
        '            Else
        '                Reason = frm.strRmks
        '            End If
        '            If clsCommon.myLen(clsCommon.myCstr(NextDayDocNo)) > 0 Then
        '                qry = "select Posted from TSPL_Demand_BOOKING_MAstER where Document_No='" + NextDayDocNo + "'"
        '                If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) = 1 Then
        '                    Throw New Exception("Please Reverse/Unpost Document No: [ " + NextDayDocNo + " ]")
        '                End If
        '                Dim dt As DataTable = Nothing
        '                '' to check gatepass or truck sheet generated
        '                Dim strDocNoForGatePass As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & NextDayDocNo & "' and IsGatePassGenerated='Y' "))
        '                Dim strDocNoForTrucksheet As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & NextDayDocNo & "' and  IsTruckSheetGenerated ='Y'  "))
        '                If clsCommon.myLen(clsCommon.myCstr(strDocNoForGatePass)) > 0 Then
        '                    Throw New Exception("Demand cannot be reverse because Next Day Demand Gate Pass has generated.")
        '                End If
        '                If clsCommon.myLen(clsCommon.myCstr(strDocNoForTrucksheet)) > 0 Then
        '                    Throw New Exception("Demand cannot be reverse because Next Day Demand Gate Pass has generated.")
        '                End If
        '            End If
        '            If clsCommon.myLen(NextDayDocNo) > 0 Then
        '                If clsDemandBookingSale.DeleteData(NextDayDocNo) Then
        '                    If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
        '                        saveCancelLog(Reason, "Reverse And Recreate", Nothing)
        '                        common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
        '                        LoadData(txtDocNo.Value, NavigatorType.Current)
        '                    End If
        '                End If
        '            Else
        '                If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
        '                    saveCancelLog(Reason, "Reverse And Recreate", Nothing)
        '                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
        '                    LoadData(txtDocNo.Value, NavigatorType.Current)
        '                End If
        '            End If
        '        End If


        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        'End Try
    End Sub
    Private Sub rbtn_Ambient_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtn_Ambient.ToggleStateChanged
        Try
            If Not isInsideLoadData AndAlso rbtn_Ambient.IsChecked Then
                'If rbtn_Ambient.IsChecked Then
                HideUnhideRowsAndColumnsOFGrid()
                'End If
                'HideUnhideRowsAndColumnsOFGrid()
            End If
            'btnPrint.Enabled = True
            lblTotalLitre.Text = ""
            lblTotalCrate.Text = ""
            lblDocumentAmt.Text = ""
            checkPrintSetting()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub txtRouteNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtRouteNo._MYValidating
        Try
            Dim isClicked As Boolean = isButtonClicked
            RouteData(isClicked, False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub RouteData(ByVal isClicked As Boolean, ByVal isQuickDemand As Boolean)
        Try
            Dim qry As String = String.Empty
            Dim ItemType As String = ""
            Dim shiftType As String = ""
            Dim Whrcls As String = ""
            'If SeparateDemandMilkandProduct Then
            '    qry = "Select TSPL_ROUTE_MASTER.Route_No as Code,Route_Desc as Description,Type,Employee_Code as 'Employee Code',Off_Day as 'Off Day' from TSPL_ROUTE_MASTER"
            'Else
            qry = "Select TSPL_ROUTE_MASTER.Route_No as Code,Route_Desc as Description,Type,Employee_Code as 'Employee Code',Off_Day as 'Off Day' from TSPL_ROUTE_MASTER"
            ' End If
            If EnableProductSaleForJPR Then
                Whrcls = " TSPL_ROUTE_MASTER.Item_Type='M' "
            End If

            If Not isQuickDemand Then
                txtRouteNo.Value = clsCommon.ShowSelectForm("DSRouteFinder", qry, "Code", Whrcls, txtRouteNo.Value, "", isClicked)
                lblRouteDesc.Text = clsCommon.myCstr(clsRouteMaster.GetName(txtRouteNo.Value, Nothing))
            End If
            If rbtnMorning.IsChecked Then
                shiftType = "Morning"
            Else
                shiftType = "Evening"
            End If
            qry = "Select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No = '" & txtRouteNo.Value & "' and   CONVERT(date,Document_Date, 103) ='" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' and ShiftType = '" & shiftType & "' and IsIndividualCustomer=0"
            If SeparateDemandMilkandProduct Then
                If rbtn_Fresh.IsChecked Then
                    qry += " and ItemType='Fresh' "
                ElseIf rbtn_Ambient.IsChecked Then
                    qry += " and ItemType='Ambient' "
                ElseIf rdbnFreshAmbientBoth.IsChecked Then
                    qry += " and ItemType='Both' "
                End If
            End If
            If chkIndividualCustomer.Checked Then
                qry += " and IsIndividualCustomer=1 "
            End If
            Dim DocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
            If clsCommon.myLen(DocNo) > 0 Then
                LoadData(DocNo, NavigatorType.Current)
            Else
                txtDate.Enabled = False
                If SeparateDemandMilkandProduct Then
                    If rbtn_Fresh.IsChecked Then
                        ItemType = "Milk"
                    ElseIf rbtn_Ambient.IsChecked Then
                        ItemType = "Product"
                    ElseIf rdbnFreshAmbientBoth.IsChecked Then
                        ItemType = "Both"
                    End If
                    qry = " select  top 1 x.ItemType 
                from(
                select TSPL_DISTRIBUTOR_ROUTE.Code as Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks,max(TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code) as cust_code,TSPL_DISTRIBUTOR_ROUTE.ItemType
                from TSPL_DISTRIBUTOR_ROUTE
                left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
                where  TSPL_DISTRIBUTOR_ROUTE.Status=1 and IS_Transpoter=0 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" & txtRouteNo.Value & "' and TSPL_DISTRIBUTOR_ROUTE.ItemType='" & ItemType & "'
                 Group by TSPL_DISTRIBUTOR_ROUTE.Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks,TSPL_DISTRIBUTOR_ROUTE.ItemType
                ) X "
                    Dim DRTItem As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                    If clsCommon.myLen(DRTItem) > 0 Then
                        If clsCommon.CompairString(DRTItem, "Milk") = CompairStringResult.Equal Then
                            rbtn_Fresh.IsChecked = True
                            RadGroupBox1.Enabled = False
                            LoadBlankGrid()
                            HideUnhideRowsAndColumnsOFGrid()
                        ElseIf clsCommon.CompairString(DRTItem, "Product") = CompairStringResult.Equal Then
                            rbtn_Ambient.IsChecked = True
                            RadGroupBox1.Enabled = False
                            LoadBlankGrid()
                            HideUnhideRowsAndColumnsOFGrid()
                        ElseIf clsCommon.CompairString(DRTItem, "Both") = CompairStringResult.Equal Then
                            rdbnFreshAmbientBoth.IsChecked = True
                            RadGroupBox1.Enabled = False
                            LoadBlankGrid()
                            HideUnhideRowsAndColumnsOFGrid()
                        End If
                    Else
                        Throw New Exception("Distributor Route Not Tagged!")
                    End If
                End If
                If EnableLocation Then
                    txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Route_Master where Route_No='" + txtRouteNo.Value + "' "))
                    txtLocation.Enabled = False
                Else
                    txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
                End If
                If clsCommon.myLen(txtLocation.Value) > 0 Then
                    lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtLocation.Value + "'"))
                End If
                If clsCommon.myLen(clsCommon.myCstr(txtRouteNo.Value)) > 0 Then
                    setRouteVehicleCityDetail()
                End If
            End If
            SetRouteColumns()
            RefreshFormName()
            If DisableRouteandVehicle Then
                txtRouteNo.Enabled = False
                txtVehicleNo.Enabled = False
            Else
                txtRouteNo.Enabled = True
                txtVehicleNo.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub SetRouteColumns()
        Try
            If rbtn_Fresh.IsChecked OrElse rdbnFreshAmbientBoth.IsChecked Then
                isDepartmentRoute = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select department_Route from TSPL_ROUTE_MASTER where Route_No='" & txtRouteNo.Value & "'")) = 1, True, False)
                isDepartmentRouteSetting = False
                If isDepartmentRoute AndAlso ApplyDepartmentRoute Then
                    'If ApplyDepartmentRoute Then
                    isDepartmentRouteSetting = True
                    'End If
                End If
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("select isnull(Entry_UOM,0) as Entry_UOM from TSPL_ROUTE_MASTER where Route_No='" & txtRouteNo.Value & "' ")
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                        Dim obj1 As ItemValueClass = TryCast(gv1.Columns(dblcolumns).Tag, ItemValueClass)
                        If obj1 IsNot Nothing AndAlso clsCommon.CompairString(obj1.IsFreshAmbient, "Fresh") = CompairStringResult.Equal Then
                            'If clsCommon.CompairString(obj1.IsFreshAmbient, "Fresh") = CompairStringResult.Equal Then
                            If clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 0 Then
                                gv1.Columns(dblcolumns).IsVisible = True
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = False
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                            ElseIf clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 1 Then
                                gv1.Columns(dblcolumns).IsVisible = True
                                gv1.Columns(dblcolumns).Width = 100
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = False
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = False
                            ElseIf clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 2 Then
                                gv1.Columns(dblcolumns).IsVisible = False
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                                gv1.Columns(dblcolumns).Width = 100
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = False
                            ElseIf clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 3 Then
                                gv1.Columns(dblcolumns).IsVisible = True
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                                gv1.Columns(dblcolumns).Width = 100
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = False
                            ElseIf clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 4 Then
                                gv1.Columns(dblcolumns).IsVisible = False
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                                gv1.Columns(dblcolumns).Width = 100
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                            ElseIf clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 5 Then
                                gv1.Columns(dblcolumns).IsVisible = False
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = False
                                gv1.Columns(dblcolumns).Width = 100
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                            ElseIf clsCommon.myCDecimal(dt.Rows(0)("Entry_UOM")) = 6 Then
                                gv1.Columns(dblcolumns).IsVisible = True
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True
                                gv1.Columns(dblcolumns).Width = 100
                                dblcolumns += 1
                                gv1.Columns(dblcolumns).IsVisible = True

                            End If
                        End If
                        'End If
                    Next
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Sub RefreshFormName()
        Me.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when len(isnull(Re_Name,''))>0 then Re_Name else Program_Name end as FormName from tspl_Program_master where program_code='" + clsUserMgtCode.frmDemandBooking + "'"))
        If clsCommon.myLen(txtRouteNo.Value) > 0 Then
            Me.Text += " - " & txtRouteNo.Value & ""
        End If
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            If clsCommon.myLen(clsCommon.myCstr(txtRouteNo.Value)) <= 0 Then
                Throw New Exception("Please select Route First")
            End If
            If clsCommon.myLen(clsCommon.myCstr(TxtCity.Value)) <= 0 Then
                Throw New Exception("Please select City First")
            End If
            setCustomerDetail(TxtCity.Value, txtRouteNo.Value, False)
            '            Dim strQry As String = "select Document_No,Document_Date,Location_Code from tspl_demand_booking_master where convert(date,document_date,103)='05-Jun-2026' and ShiftType='Morning'"
            '            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry, trans)
            '            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '                For Each dr As DataRow In dt.Rows
            '                    strQry = "select TR_Code,Cust_Code,Item_Code,Unit_code,Price_code,ShiftType,Item_Rate,ItemNetAmount,* from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + clsCommon.myCstr(dr("Document_No")) + "'"
            '                    Dim dtd As DataTable = clsDBFuncationality.GetDataTable(strQry, trans)
            '                    If dtd IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            '                        For Each drs As DataRow In dtd.Rows
            '                            Dim PricePlan As String = "  Select Is_With_Tax, RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, XXXE.Tax_group,XXXE.TAX1_Rate, " &
            '                    " XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
            '                    "  XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, " &
            '                    " XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, " &
            '                    " XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7, " &
            '                    " XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.TAX1_Amt, " &
            '                    " XXXE.TAX2_Amt,XXXE.TAX3_Amt,XXXE.TAX4_Amt, XXXE.TAX5_Amt,XXXE.TAX6_Amt,XXXE.TAX7_Amt,XXXE.TAX8_Amt,XXXE.TAX9_Amt,XXXE.TAX10_Amt,XXXE.Against_Plan_TR_Code  from ( " &
            '                    "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
            '                    "Start_Date Desc) as RowNo,Is_With_Tax, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
            '                    "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, TSPL_ITEM_PRICE_MASTER.Tax_group,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,  " &
            '                    "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
            '                    " TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
            '                    " TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
            '                    " TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7, " &
            '                    " TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.TAX1_Amt , TSPL_ITEM_PRICE_MASTER.TAX2_Amt ,TSPL_ITEM_PRICE_MASTER.TAX3_Amt ,TSPL_ITEM_PRICE_MASTER.TAX4_Amt,TSPL_ITEM_PRICE_MASTER.TAX5_Amt,TSPL_ITEM_PRICE_MASTER.TAX6_Amt,TSPL_ITEM_PRICE_MASTER.TAX7_Amt,   TSPL_ITEM_PRICE_MASTER.TAX8_Amt,TSPL_ITEM_PRICE_MASTER.TAX9_Amt,TSPL_ITEM_PRICE_MASTER.TAX10_Amt,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
            '                    "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
            '                    "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  2=( case when CONVERT(date,Start_Date,103)='" + clsCommon.GetPrintDate(dr("Document_Date")) + "' and Shift_Type='Morning' then 2 else ( case when CONVERT(date,Start_Date,103)<='" + clsCommon.GetPrintDate(clsCommon.myCDate(dr("Document_Date")).AddDays(-1)) + "' then 2 else 3 end)  end)  and (End_Date >= '" & clsCommon.GetPrintDate(clsCommon.myCDate(dr("Document_Date")), "dd/MMM/yyyy") & "'  or End_date is null)  and  " &
            '                    "TSPL_ITEM_PRICE_MASTER.Price_Code='" & clsCommon.myCstr(drs("Price_code")) & "' and UOM='" & clsCommon.myCstr(drs("Unit_code")) & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & clsCommon.myCstr(drs("Item_Code")) & "' AND Location_Code='" & clsCommon.myCstr(dr("Location_Code")) & "'  " &
            '                    ") XXXE WHERE RowNo=1  "
            '                            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(PricePlan, trans)
            '                            Dim dblRate As Double = 0
            '                            If dt1.Rows.Count > 0 Then

            '                                dblRate = clsCommon.myCdbl(dt1.Rows(0).Item("Item_Basic_Price"))
            '                            Else
            '                                Throw New Exception("price not found!")
            '                            End If
            '                            Dim updateItem As String = "update TSPL_DEMAND_BOOKING_DETAIL set Item_Rate='" + clsCommon.myCstr(dblRate) + "',
            'TAX1_Base_Amt=Qty*Item_Rate,TAX1_Amt=((Qty*Item_Rate)*(TAX1_Rate/100)),
            'TAX2_Base_Amt=Qty*Item_Rate,TAX2_Amt=((Qty*Item_Rate)*(TAX2_Rate/100)),
            'TAX3_Base_Amt=Qty*Item_Rate,TAX3_Amt=((Qty*Item_Rate)*(TAX3_Rate/100)),
            'TAX4_Base_Amt=Qty*Item_Rate,TAX4_Amt=((Qty*Item_Rate)*(TAX4_Rate/100)),
            'TAX5_Base_Amt=Qty*Item_Rate,TAX5_Amt=((Qty*Item_Rate)*(TAX5_Rate/100)),
            'TAX6_Base_Amt=Qty*Item_Rate,TAX6_Amt=((Qty*Item_Rate)*(TAX6_Rate/100)),
            'TAX7_Base_Amt=Qty*Item_Rate,TAX7_Amt=((Qty*Item_Rate)*(TAX7_Rate/100)),
            'TAX8_Base_Amt=Qty*Item_Rate,TAX8_Amt=((Qty*Item_Rate)*(TAX8_Rate/100)),
            'TAX9_Base_Amt=Qty*Item_Rate,TAX9_Amt=((Qty*Item_Rate)*(TAX9_Rate/100)),
            'TAX10_Base_Amt=Qty*Item_Rate,TAX10_Amt=((Qty*Item_Rate)*(TAX10_Rate/100)),
            'ItemNetAmount=(Qty*Item_Rate)+TAX1_Amt+TAX2_Amt+TAX3_Amt+TAX4_Amt+TAX5_Amt
            'where TR_Code='" + clsCommon.myCstr(drs("TR_Code")) + "'"
            '                            clsDBFuncationality.ExecuteNonQuery(updateItem, trans)
            '                            'Dim strQ1 As String = "select TR_Code,Cust_Code,Item_Code,Unit_code,Price_code,ShiftType,Item_Rate,ItemNetAmount from TSPL_DEMAND_BOOKING_DETAIL where TR_Code='" + clsCommon.myCstr(drs("TR_Code")) + "'"
            '                            'Dim dtrs As DataTable = clsDBFuncationality.GetDataTable(strQ1, trans)
            '                            'If dtrs IsNot Nothing AndAlso dtrs.Rows.Count > 0 Then

            '                            'End If
            '                        Next
            '                    End If
            '                    Dim ItemNetAmount As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select sum(ItemNetAmount) as ItemNetAmount from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + clsCommon.myCstr(dr("Document_No")) + "'", trans))
            '                    clsDBFuncationality.ExecuteNonQuery("update TSPL_DEMAND_BOOKING_MASTER set DocumentAmount='" + clsCommon.myCstr(ItemNetAmount) + "' where Document_No='" + clsCommon.myCstr(dr("Document_No")) + "'", trans)
            '                Next
            '                trans.Commit()
            '            Else
            '                trans.Rollback()
            '            End If

        Catch ex As Exception
            'trans.Rollback()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub setCustomerDetail(ByVal strCityCode As String, ByVal strtRouteCode As String, ByVal isLoadData As Boolean)
        Try
            Dim MainQry As String = ""
            Dim qry As String = ""
            Dim isPosted As Boolean = False
            If objCommonVar.ApplyBoothRouteMapping Then
                qry = "select top 1 TSPL_Booth_Route_Mapping_Head.Document_No from TSPL_Booth_Route_Mapping_Head
left join TSPL_Booth_Route_Mapping_Detail on TSPL_Booth_Route_Mapping_Detail.Document_No=TSPL_Booth_Route_Mapping_Head.Document_No
where CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)<='" + clsCommon.GetPrintDate(txtDate.Value) + "' and TSPL_Booth_Route_Mapping_Head.Route_No='" + strtRouteCode + "' 
and isnull(TSPL_Booth_Route_Mapping_Head.Posted,0)=1 and Item_Type='Milk' and 2=( case when CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)<='" + IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "' then 2 else 3 end)  end) order by Document_No desc"
                Dim BRM_DocNO As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                qry = " select TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Document_No, TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code as cust_code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Serial_No from TSPL_BOOTH_ROUTE_MAPPING_DETAIL 
left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code 
where Document_No='" & BRM_DocNO & "' and TSPL_CUSTOMER_MASTER.Status='N' "
                If chkIndividualCustomer.Checked Then
                    qry += " and TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code ='" & txtCustomerNo.Value & "' "
                End If
                If isLoadData Then
                    qry += " union select max(TSPL_DEMAND_BOOKING_DETAIL.Document_No) as Document_No,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,max(TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name,max(TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Serial_No) as Serial_No from TSPL_DEMAND_BOOKING_MASTER 
                left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
                left join TSPL_BOOTH_ROUTE_MAPPING_DETAIL on TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
                left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
                where TSPL_DEMAND_BOOKING_MASTER.Document_No='" + txtDocNo.Value + "' and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code not in (select TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code as cust_code from TSPL_BOOTH_ROUTE_MAPPING_DETAIL 
left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code 
where Document_No='" & BRM_DocNO & "' and TSPL_CUSTOMER_MASTER.Status='N')
                 group by TSPL_DEMAND_BOOKING_DETAIL.Cust_Code  " '--,TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Serial_No
                End If
                'MainQry = "Select xx.* from (" + qry + " ) xx   order by xx.Serial_No"
                MainQry = qry & "  order by TSPL_Booth_Route_Mapping_Detail.Serial_No "
                isPosted = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Posted from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + txtDocNo.Value + "'  and Posted=1")) = 1, True, False)
                If isPosted Then
                    MainQry = "select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,max(TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name,max(TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Serial_No) as Serial_No from TSPL_DEMAND_BOOKING_MASTER 
                left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
                left join TSPL_BOOTH_ROUTE_MAPPING_DETAIL on TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code and TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Document_No='" + BRM_DocNO + "' 
                left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
                where TSPL_DEMAND_BOOKING_MASTER.Document_No='" + txtDocNo.Value + "'
                 group by TSPL_DEMAND_BOOKING_DETAIL.Cust_Code order by Serial_No "
                End If
            Else
                qry = "select cust_code,Customer_name,"
                If SeprateMorningEveningSequence Then
                    If rbtnEvening.IsChecked Then
                        qry += " display_seqE "
                    ElseIf rbtnMorning.IsChecked Then
                        qry += " display_seqM "
                    End If
                Else
                    qry += " display_seq "
                End If
                qry += " From TSPL_CUSTOMER_MASTER Where route_no ='" & strtRouteCode & "'  and  TSPL_CUSTOMER_MASTER.Status='N' "
                If chkIndividualCustomer.Checked Then
                    qry += " and TSPL_CUSTOMER_MASTER.Cust_Code ='" & txtCustomerNo.Value & "'"
                End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    qry += "  and IsDistributor='N'"
                End If

                ' qry += " order by isnull(TSPL_CUSTOMER_MASTER.display_seq,0)  "
                If isLoadData Then
                    qry += "union 
select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,max(TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name,"
                    If SeprateMorningEveningSequence Then
                        If rbtnEvening.IsChecked Then
                            qry += " TSPL_CUSTOMER_MASTER.display_seqE "
                        ElseIf rbtnMorning.IsChecked Then
                            qry += " TSPL_CUSTOMER_MASTER.display_seqM "
                        End If
                    Else
                        qry += " TSPL_CUSTOMER_MASTER.display_seq "
                    End If
                    qry += "from TSPL_DEMAND_BOOKING_MASTER 
left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left join TSPL_CUSTOMER_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
where TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code is not null
group by TSPL_DEMAND_BOOKING_DETAIL.Cust_Code"
                    If SeprateMorningEveningSequence Then
                        If rbtnEvening.IsChecked Then
                            qry += " ,TSPL_CUSTOMER_MASTER.display_seqE "
                        ElseIf rbtnMorning.IsChecked Then
                            qry += " ,TSPL_CUSTOMER_MASTER.display_seqM "
                        End If
                    Else
                        qry += " ,TSPL_CUSTOMER_MASTER.display_seq "
                    End If
                    MainQry += "select X.* from (" & qry & " )X "
                    If SeprateMorningEveningSequence Then
                        If rbtnEvening.IsChecked Then
                            MainQry += " order by isnull(X.display_seqE,0) "
                        ElseIf rbtnMorning.IsChecked Then
                            MainQry += " order by isnull(X.display_seqM,0) "
                        End If
                    Else
                        MainQry += " order by isnull(X.display_seq,0) "
                    End If
                Else
                    MainQry = qry '+ " order by isnull(TSPL_CUSTOMER_MASTER.display_seq,0) "
                    If SeprateMorningEveningSequence Then
                        If rbtnEvening.IsChecked Then
                            MainQry += " order by isnull(TSPL_CUSTOMER_MASTER.display_seqE,0) "
                        ElseIf rbtnMorning.IsChecked Then
                            MainQry += " order by isnull(TSPL_CUSTOMER_MASTER.display_seqM,0) "
                        End If
                    Else
                        MainQry += " order by isnull(TSPL_CUSTOMER_MASTER.display_seq,0) "
                    End If
                End If
                If Not SeparateDemandMilkandProduct Then
                    LoadBlankGrid()
                End If
            End If

            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(MainQry)
            If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                Dim i As Integer = 1
                For Each dr As DataRow In dt1.Rows
                    If objCommonVar.ApplyBoothRouteMapping Then
                        If Not isPosted Then
                            Dim strBoothQry As String = "select top 1 * from 
TSPL_BOOTH_ROUTE_MAPPING_HEAD
left join TSPL_BOOTH_ROUTE_MAPPING_DETAIL on TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Document_No=TSPL_BOOTH_ROUTE_MAPPING_HEAD.Document_No
 where TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code='" & clsCommon.myCstr(dr("cust_code")) & "' and CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)<='" & clsCommon.GetPrintDate(txtDate.Value) & "'
and isnull(TSPL_Booth_Route_Mapping_Head.Posted,0)=1 and Item_Type='Milk' and 2=( case when CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value) & "' and Shift_Type='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "' then 2 else ( case when CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)<='" & IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) & "' then 2 else 3 end)  end) order by PK_ID desc "
                            Dim dtBooth As DataTable = clsDBFuncationality.GetDataTable(strBoothQry)
                            If clsCommon.CompairString(dtBooth.Rows(0)("Document_No"), clsCommon.myCstr(dr("Document_No"))) = CompairStringResult.Equal Then
                                Dim flagE As Boolean = True
                                Dim flagM As Boolean = True
                                If SettSeprateDemandForMorningEveningShift Then
                                    flagE = rbtnEvening.IsChecked
                                    flagM = rbtnMorning.IsChecked
                                    RadGroupBox3.Enabled = False
                                End If
                                If flagE Then
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colbtncol).Value = "Reset "
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTripNo).Value = 1
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).Value = "Evening"
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).ReadOnly = True
                                    i = i + 1
                                    gv1.Rows.AddNew()
                                End If
                                If flagM Then
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colbtncol).Value = "Reset "
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colTripNo).Value = 1
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).Value = "Morning"
                                    gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).ReadOnly = True
                                    i = i + 1
                                    gv1.Rows.AddNew()
                                End If
                            End If
                        Else
                            Dim flagE As Boolean = True
                            Dim flagM As Boolean = True
                            If SettSeprateDemandForMorningEveningShift Then
                                flagE = rbtnEvening.IsChecked
                                flagM = rbtnMorning.IsChecked
                                RadGroupBox3.Enabled = False
                            End If
                            If flagE Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colbtncol).Value = "Reset "
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTripNo).Value = 1
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).Value = "Evening"
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).ReadOnly = True
                                i = i + 1
                                gv1.Rows.AddNew()
                            End If
                            If flagM Then
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colbtncol).Value = "Reset "
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colTripNo).Value = 1
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).Value = "Morning"
                                gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).ReadOnly = True
                                i = i + 1
                                gv1.Rows.AddNew()
                            End If
                        End If
                    Else
                        Dim flagE As Boolean = True
                        Dim flagM As Boolean = True
                        If SettSeprateDemandForMorningEveningShift Then
                            flagE = rbtnEvening.IsChecked
                            flagM = rbtnMorning.IsChecked
                            RadGroupBox3.Enabled = False
                        End If
                        If flagE Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colbtncol).Value = "Reset "
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTripNo).Value = 1
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).Value = "Evening"
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).ReadOnly = True
                            i = i + 1
                            gv1.Rows.AddNew()
                        End If
                        If flagM Then
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colbtncol).Value = "Reset "
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colLineNo).Value = i
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colTripNo).Value = 1
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colCustCode).Value = clsCommon.myCstr(dr("cust_code"))
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colCustName).Value = clsCommon.myCstr(dr("Customer_name"))
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).Value = "Morning"
                            gv1.Rows(gv1.Rows.Count - 1).Cells(colShiftName).ReadOnly = True
                            i = i + 1
                            gv1.Rows.AddNew()
                        End If
                    End If



                Next
                '                For n As Integer = 0 To gv1.Rows.Count - 1
                '                    Try
                '                        If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(n).Cells(colCustCode).Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(gv1.Rows(n).Cells(colShiftName).Value)) > 0 Then
                '                            Dim strqry As String = "Select count(ShiftType) from (select ShiftType,convert(date,Document_Date ,103) as Document_Date from (select tspl_demand_booking_master.Document_Date , tspl_demand_booking_detail.ShiftType  from tspl_demand_booking_detail
                'left outer join tspl_demand_booking_master on tspl_demand_booking_detail.document_no=tspl_demand_booking_master.document_no
                'where convert(date,tspl_demand_booking_master.document_date,103) in ('" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)) & "','" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-2)) & "','" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-3)) & "')
                'and tspl_demand_booking_detail.Cust_Code='" & clsCommon.myCstr(gv1.Rows(n).Cells(colCustCode).Value) & "' and TSPL_DEMAND_BOOKING_DETAIL .ShiftType ='" & clsCommon.myCstr(gv1.Rows(n).Cells(colShiftName).Value) & "')Final 
                'group by ShiftType ,convert(date,Document_Date ,103))FinalQry"
                '                            gv1.Rows(n).Cells(colBookingCreatedFor3Days).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strqry))
                '                            Try
                '                                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(n).Cells(colBookingCreatedFor3Days).Value), "3") = CompairStringResult.Equal Then
                '                                    gv1.Rows(n).Cells(colLineNo).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colLineNo).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colLineNo).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colCustCode).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colCustCode).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colCustCode).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colCustName).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colCustName).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colCustName).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colShiftName).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colShiftName).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colShiftName).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colCrate).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colCrate).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colCrate).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colAmt).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colAmt).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colAmt).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colLitre).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colLitre).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colLitre).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colMAmt).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colMAmt).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colMAmt).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colPCount).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colPCount).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colPCount).Style.BackColor = Color.LightGreen
                '                                    gv1.Rows(n).Cells(colPAmt).Style.DrawFill = True
                '                                    gv1.Rows(n).Cells(colPAmt).Style.CustomizeFill = True
                '                                    gv1.Rows(n).Cells(colPAmt).Style.BackColor = Color.LightGreen
                '                                End If
                '                            Catch ex As Exception
                '                                Throw New Exception(ex.Message)
                '                            End Try
                '                            Try
                '                                For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                '                                    Dim obj1 As ItemValueClass = TryCast(gv1.Columns(dblcolumns).Tag, ItemValueClass)
                '                                    If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code).ToUpper, "POUCH") = CompairStringResult.Equal Then
                '                                        gv1.Rows(n).Cells(dblcolumns).Style.DrawFill = True
                '                                        gv1.Rows(n).Cells(dblcolumns).Style.CustomizeFill = True
                '                                        gv1.Rows(n).Cells(dblcolumns).Style.BackColor = Color.LightGreen
                '                                    End If
                '                                Next
                '                            Catch ex As Exception
                '                                Throw New Exception(ex.Message)
                '                            End Try
                '                        End If
                '                    Catch ex As Exception
                '                        Throw New Exception(ex.Message)
                '                    End Try
                '                Next
            Else
                Throw New Exception("No Customers found for Selected Route and City")
            End If
            gv1.Columns(colLineNo).IsPinned = True
            gv1.Columns(colCustCode).IsPinned = True
            gv1.Columns(colCustName).IsPinned = True
            gv1.Columns(colShiftName).IsPinned = True
            gv1.Columns(colItemExist).IsPinned = True
            gv1.Columns(colIsItemUpdate).IsPinned = True
            gv1.Columns(colBookingCreatedFor3Days).IsPinned = True
            gv1.Columns(colAmt).IsPinned = True
            gv1.Columns(colCrate).IsPinned = True
            gv1.Columns(colLitre).IsPinned = True
            gv1.Columns(colMAmt).IsPinned = True
            gv1.Columns(colPCount).IsPinned = True
            gv1.Columns(colPAmt).IsPinned = True
            gv1.Columns(colLineNo).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colCustCode).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colCustName).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colShiftName).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colItemExist).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colIsItemUpdate).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colBookingCreatedFor3Days).PinPosition = PinnedColumnPosition.Left
            gv1.Columns(colAmt).PinPosition = PinnedColumnPosition.Right
            gv1.Columns(colCrate).PinPosition = PinnedColumnPosition.Right
            gv1.Columns(colLitre).PinPosition = PinnedColumnPosition.Right
            gv1.Columns(colMAmt).PinPosition = PinnedColumnPosition.Right
            gv1.Columns(colPCount).PinPosition = PinnedColumnPosition.Right
            gv1.Columns(colPAmt).PinPosition = PinnedColumnPosition.Right
            gv1.ShowGroupPanel = False
            gv1.MasterTemplate.AutoExpandGroups = True
            gv1.BestFitColumns()
            'MergeHorizontally(gv1, 1, gv1.Rows.Count - 1)
            If gv1.Rows.Count > 0 Then
                gv1.CurrentRow = gv1.Rows(0)
                gv1.CurrentColumn = gv1.Columns(7)
            End If
            'MergeVertically(gv1, New Integer() {1, 2})
            'MergeVertically(gv1, {1, 2})
            'View()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub rbtnMorning_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnMorning.ToggleStateChanged
        Try
            If Not isInsideLoadData Then
                HideUnhideRowsAndColumnsOFGrid()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub HideUnhideRowsAndColumnsOFGrid()
        Try
            isLoadData = True
            Dim dblTotalCount As Decimal = 0
            Dim dblTotalPAmt As Decimal = 0
            Dim dblTotalCeart As Decimal = 0
            Dim dblTotalMAmt As Decimal = 0
            Dim dblTotalLiter As Decimal = 0
            For dblrows As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.myLen(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value)) > 0 Then
                    If rbtnMorning.IsChecked Then
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colShiftName).Value), "Evening") = CompairStringResult.Equal Then
                            gv1.Rows(dblrows).IsVisible = False
                        Else
                            If rbtn_Fresh.IsChecked Then
                                dblTotalCeart += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                                dblTotalLiter += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                                dblTotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
                            ElseIf rbtn_Ambient.IsChecked Then
                                dblTotalCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                                dblTotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                            Else
                                dblTotalCeart += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                                dblTotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
                                dblTotalLiter += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                                dblTotalCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                                dblTotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                            End If
                            gv1.Rows(dblrows).IsVisible = True
                        End If
                    ElseIf rbtnEvening.IsChecked Then
                        If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colShiftName).Value), "Morning") = CompairStringResult.Equal Then
                            gv1.Rows(dblrows).IsVisible = False
                        Else
                            If rbtn_Fresh.IsChecked Then
                                dblTotalCeart += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                                dblTotalLiter += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                                dblTotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
                            ElseIf rbtn_Ambient.IsChecked Then
                                dblTotalCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                                dblTotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                            Else
                                dblTotalCeart += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                                dblTotalLiter += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                                dblTotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
                                dblTotalCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                                dblTotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                            End If
                            gv1.Rows(dblrows).IsVisible = True
                            'gv1.Rows(dblrows).IsVisible = True
                        End If
                    Else
                        If rbtn_Fresh.IsChecked Then
                            dblTotalCeart += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                            dblTotalLiter += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                            dblTotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
                        ElseIf rbtn_Ambient.IsChecked Then
                            dblTotalCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                            dblTotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                        Else
                            dblTotalCeart += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                            dblTotalLiter += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                            dblTotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
                            dblTotalCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                            dblTotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                        End If
                        gv1.Rows(dblrows).IsVisible = True
                    End If
                End If
            Next
            Dim k As Integer = 1
            For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                Dim obj1 As ItemValueClass = TryCast(gv1.Columns(dblcolumns).Tag, ItemValueClass)
                If obj1 IsNot Nothing Then
                    If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 Then
                        If rbtn_Fresh.IsChecked Then
                            If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                                gv1.Columns(dblcolumns).IsVisible = True
                                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                                    gv1.Columns(colPAmt).IsVisible = False
                                    gv1.Columns(colPCount).IsVisible = False
                                    gv1.Columns(colMAmt).IsVisible = False
                                    gv1.Columns(colCrate).IsVisible = True
                                    gv1.Columns(colLitre).IsVisible = False
                                Else
                                    gv1.Columns(colMAmt).IsVisible = True
                                    gv1.Columns(colCrate).IsVisible = True
                                    gv1.Columns(colLitre).IsVisible = True
                                End If

                            Else
                                'gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                gv1.Columns(dblcolumns).IsVisible = False
                                gv1.Columns(colPAmt).IsVisible = False
                                gv1.Columns(colPCount).IsVisible = False
                            End If
                        ElseIf rbtn_Ambient.IsChecked Then
                            If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Ambient") = CompairStringResult.Equal Then
                                gv1.Columns(dblcolumns).IsVisible = True
                                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                                    gv1.Columns(colPAmt).IsVisible = False
                                    gv1.Columns(colPCount).IsVisible = False
                                    gv1.Columns(colMAmt).IsVisible = False
                                    gv1.Columns(colCrate).IsVisible = True
                                    gv1.Columns(colLitre).IsVisible = False
                                Else

                                    gv1.Columns(colPAmt).IsVisible = True
                                    gv1.Columns(colPCount).IsVisible = True
                                End If

                            Else
                                'gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                gv1.Columns(dblcolumns).IsVisible = False
                                gv1.Columns(colMAmt).IsVisible = False
                                gv1.Columns(colCrate).IsVisible = False
                                gv1.Columns(colLitre).IsVisible = False
                            End If
                        ElseIf rdbnFreshAmbientBoth.IsChecked Then
                            gv1.Columns(dblcolumns).IsVisible = True
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                                gv1.Columns(colPAmt).IsVisible = False
                                gv1.Columns(colPCount).IsVisible = False
                                gv1.Columns(colMAmt).IsVisible = False
                                gv1.Columns(colCrate).IsVisible = True
                                gv1.Columns(colLitre).IsVisible = False
                            Else

                                gv1.Columns(colPAmt).IsVisible = True
                                gv1.Columns(colPCount).IsVisible = True
                                gv1.Columns(colMAmt).IsVisible = True
                                gv1.Columns(colCrate).IsVisible = True
                                gv1.Columns(colLitre).IsVisible = True
                            End If

                        End If


                    End If
                End If
                k = k + 1
            Next
            If clsCommon.myLen(txtDocNo.Value) > 0 AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") <> CompairStringResult.Equal Then
                'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") <> CompairStringResult.Equal Then
                GatePass_TruckSheet_Button()
                'End If
            End If
            'MergeVertically(gv1, New Integer() {1, 2})
            isLoadData = False
            'UpdateAllTotals()
            lblTotalCrate.Text = dblTotalCeart
            lblTotalLitre.Text = dblTotalLiter
            lblDocumentAmt.Text = dblTotalMAmt
            txtPCount.Text = dblTotalCount
            txtPAmt.Text = dblTotalPAmt
            txtDocAmt.Text = dblTotalMAmt + dblTotalPAmt
            SetRouteColumns()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            isLoadData = False
        End Try
    End Sub
    Private Sub UpdateItemQtyAfterSave(ByVal dblrows As Integer, ByVal dblcolumns As Integer)
        Try
            Dim strItemUpdateAfterSave As String = String.Empty
            'For dblrows As Integer = 0 To gv1.Rows.Count - 1
            Dim k As Integer = 1
            ' For dblrows As Integer = 0 To gv1.Rows.Count - 1
            Dim TotalCrate As Double = 0
            Dim dblTotalDocAmt As Decimal = 0
            Dim dblTotalCrateRowWise As Double = 0
            Dim strPriceCode As String = String.Empty
            Dim strItemValueExist As String = String.Empty
            Dim TotalLitre As Decimal = 0
            Dim dblTotalLitreRowWise As Double = 0
            k = 1
            strItemValueExist = "No"
            strItemUpdateAfterSave = "No"
            Dim obj1 As ItemValueClass = TryCast(gv1.Columns(dblcolumns).Tag, ItemValueClass)
            k = k + 1
            If obj1 IsNot Nothing AndAlso clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                'If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 AndAlso clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                'If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                If clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code), "Crate") = CompairStringResult.Equal Then
                    'TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colItemCode + clsCommon.myCstr(dblcolumns)).Value)
                    TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                    obj1.FreshItem_QtyInCrates = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                Else
                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                    If ItemCrateType = 1 Then
                        ' Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                        Dim CrateConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                        Dim ItemConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                        If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                            Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                            If DispatchQty >= CrateConvFactor Then
                                dblTotalCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                            Else
                                dblTotalCrateRowWise = 0
                            End If
                        Else
                            ' clsCommon.MyMessageBoxShow("Please fill conversion factor for this unit at line no." & i + 1 & "")
                        End If
                        obj1.FreshItem_QtyInCrates = dblTotalCrateRowWise
                        TotalCrate = TotalCrate + dblTotalCrateRowWise
                    End If
                End If
                ''to convert into litre
                'Dim IsStockingUnit_Ltr As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                Dim CrateConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='Ltr' "))
                Dim ItemConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                If CrateConvFactor_Ltr > 0 AndAlso ItemConvFactor_Ltr > 0 Then
                    Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor_Ltr
                    If DispatchQty >= CrateConvFactor_Ltr Then
                        dblTotalLitreRowWise = Math.Floor(DispatchQty / CrateConvFactor_Ltr)
                    Else
                        dblTotalLitreRowWise = 0
                    End If
                Else
                    ' clsCommon.MyMessageBoxShow("Please fill conversion factor for this unit at line no." & i + 1 & "")
                End If
                obj1.FreshItem_QtyInLitres = dblTotalLitreRowWise
                TotalLitre = TotalLitre + dblTotalLitreRowWise
                ''---------end of litre conversion
                'End If
                strItemValueExist = "Yes"
                Dim dt As New DataTable()
                Dim dblRate As Double = 0
                Dim qry As String = "select tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER where Cust_Code='" & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value) & "'"
                strPriceCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                If clsCommon.myLen(strPriceCode) <= 0 Then
                    Throw New InvalidOperationException("price_CodeNon not found for Customer " & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustName).Value) & "")
                End If
                qry = " Select Is_With_Tax, RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price,XXXE.TAX1_Rate, " &
                " XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
                "  XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, " &
                " XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, " &
                " XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7, " &
                " XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.TAX1_Amt, " &
                " XXXE.TAX2_Amt,XXXE.TAX3_Amt,XXXE.TAX4_Amt,XXXE.Against_Plan_TR_Code  from ( " &
                "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                "Start_Date Desc) as RowNo,Is_With_Tax, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
                "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,  " &
                "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
                " TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
                " TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
                " TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7, " &
                " TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.TAX1_Amt , TSPL_ITEM_PRICE_MASTER.TAX2_Amt ,TSPL_ITEM_PRICE_MASTER.TAX3_Amt ,TSPL_ITEM_PRICE_MASTER.TAX4_Amt,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  2=( case when CONVERT(date,Start_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,Start_Date,103)<='" + IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "' then 2 else 3 end)  end) and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null)  and  " &
                "TSPL_ITEM_PRICE_MASTER.Price_Code='" & strPriceCode & "' and UOM='" & obj1.Unit_code & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & obj1.itemCode & "' AND Location_Code='" & clsCommon.myCstr(txtLocation.Value) & "'  " &
                ") XXXE WHERE RowNo=1  "
                dt = clsDBFuncationality.GetDataTable(qry)
                If dt.Rows.Count > 0 Then
                    dblRate = clsCommon.myCdbl(dt.Rows(0).Item("Item_Selling_Price"))
                    If dblRate = 0 Then
                        Throw New InvalidOperationException("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(obj1.ShortDesc) & Environment.NewLine)
                    End If
                    dblTotalDocAmt = dblTotalDocAmt + Math.Round(clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * clsCommon.myCdbl(dblRate), 2)
                Else
                    Throw New InvalidOperationException("Please create Price for Item " & obj1.ShortDesc & " and customer " & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustName).Value) & " ")
                End If
                dt = Nothing
                'End If
            End If
            ' Next
            gv1.Rows(dblrows).Cells(colCrate).Value = Math.Round(clsCommon.myCdbl(TotalCrate), 2)
            gv1.Rows(dblrows).Cells(colLitre).Value = Math.Round(clsCommon.myCdbl(TotalLitre), 2)
            gv1.Rows(dblrows).Cells(colAmt).Value = clsCommon.myCdbl(dblTotalDocAmt)
            'gv1.Rows(dblrows).Cells(colMAmt).Value = clsCommon.myCdbl(dblTotalDocAmt)
            'gv1.Rows(dblrows).Cells(colPCount).Value = clsCommon.myCdbl(dblTotalDocAmt)
            'gv1.Rows(dblrows).Cells(colPAmt).Value = clsCommon.myCdbl(dblTotalDocAmt)
            gv1.Rows(dblrows).Cells(colItemExist).Value = strItemValueExist
            ' Next
        Catch ex As InvalidOperationException
            Throw New InvalidOperationException(ex.Message)
        End Try
    End Sub
    Private Function UpdateAllTotals(ByVal isLoad As Boolean) As List(Of clsDemandCustItem)
        Dim lstCustItem As New List(Of clsDemandCustItem)
        Try

            isInsideLoadData = True
            Dim TotalCrate As Double = 0
            Dim TotalLitre As Double = 0
            Dim TotalPCount As Double = 0
            Dim TotalPCrate As Double = 0
            Dim TotalPAmt As Double = 0
            Dim TotalMAmt As Double = 0
            Dim dblTotalDocAmtRowWise As Decimal = 0
            Dim dblTotalCrateRowWise As Double = 0
            Dim dblTotalPCrateRowWise As Double = 0
            Dim dblTotalLitreRowWise As Double = 0
            Dim strPriceCode As String = String.Empty
            Dim strItemValueExist As String = String.Empty
            Dim strItemUpdateAfterSave As String = String.Empty
            Dim dblDocTotalCrate As Double = 0
            Dim dblDocTotalLitre As Double = 0
            Dim dblDocTotalAmt As Double = 0
            Dim dblTotalPCount As Double = 0
            Dim dblTotalPCrate As Double = 0
            'Dim dblToalPouchCount As Double = 0
            Dim dblTotalPAmt As Double = 0
            Dim dblTotalMAmt As Double = 0
            'Dim colTotalQty As Double = 0
            For dblrows As Integer = 0 To gv1.Rows.Count - 2
                Dim k As Integer = 1
                TotalCrate = 0
                dblTotalPCrate = 0
                dblTotalCrateRowWise = 0
                dblTotalPCrateRowWise = 0
                TotalLitre = 0
                dblTotalLitreRowWise = 0
                dblTotalDocAmtRowWise = 0
                dblTotalPCount = 0
                dblTotalPAmt = 0
                dblTotalMAmt = 0
                strItemValueExist = "No"
                strItemUpdateAfterSave = "No"
                For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                    Dim obj1 As ItemValueClass = Nothing
                    Try
                        obj1 = TryCast(gv1.Columns(colItemCode & clsCommon.myCstr(k)).Tag, ItemValueClass)
                    Catch ex As InvalidOperationException
                        Throw New InvalidOperationException(ex.Message)
                    End Try
                    k = k + 1
                    If obj1 IsNot Nothing AndAlso clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 Then
                        'If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code), "Crate") = CompairStringResult.Equal Then
                                TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                obj1.FreshItem_QtyInCrates = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                'gv1.Rows(gv1.Rows.Count - 1).Cells(dblcolumns).Value = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                            Else
                                Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                                If ItemCrateType = 1 Then
                                    'Dim IsStockingUnit As String = obj1.Stocking_Unit 'clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                    Dim CrateConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                    Dim ItemConvFactor As Double = obj1.Conversion_Factor 'clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal AndAlso CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                        'If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                        Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                                        If ConvertPouchtoCrate Then
                                            If DispatchQty > (CrateConvFactor / 2) Then
                                                dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                            Else
                                                dblTotalCrateRowWise = 1
                                            End If
                                        Else
                                            If DispatchQty > (CrateConvFactor / 2) Then
                                                dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                            Else
                                                dblTotalCrateRowWise = 0
                                            End If
                                        End If
                                        'End If
                                    End If
                                    TotalCrate = TotalCrate + dblTotalCrateRowWise
                                    obj1.FreshItem_QtyInCrates = dblTotalCrateRowWise
                                End If
                            End If
                            ''to convert into litre
                            Dim CrateConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='Ltr' "))
                            Dim ItemConvFactor_Ltr As Double = obj1.Conversion_Factor 'clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                            If CrateConvFactor_Ltr > 0 AndAlso ItemConvFactor_Ltr > 0 Then
                                Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor_Ltr
                                dblTotalLitreRowWise = (DispatchQty / CrateConvFactor_Ltr)
                            End If
                            TotalLitre = TotalLitre + dblTotalLitreRowWise
                            obj1.FreshItem_QtyInLitres = dblTotalLitreRowWise
                            ''---------end of litre conversion
                        Else
                            If AllowMultipleUOMForProduct Then
                                If clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code), "Crate") = CompairStringResult.Equal Then
                                    TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                Else
                                    Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                                    If ItemCrateType = 1 Then
                                        'Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                        Dim CrateConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                        Dim ItemConvFactor As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                        Dim ItempouchCF As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))

                                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal AndAlso CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                            'If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                            Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                                            'If ConvertPouchtoCrate Then
                                            If DispatchQty > (CrateConvFactor / 2) Then
                                                dblTotalPCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                            Else
                                                dblTotalPCrateRowWise = 0
                                            End If
                                            'ElseIf DontCreateForPouch Then
                                            '    dblTotalCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                            'Else
                                            '    If DispatchQty > (CrateConvFactor / 2) Then
                                            '        dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                            '    Else
                                            '        dblTotalCrateRowWise = 0
                                            '    End If
                                            'End If
                                            'End If
                                        End If

                                        If isDepartmentRoute AndAlso ApplyDepartmentRoute Then
                                            'If ApplyDepartmentRoute Then
                                            Dim cellValue As String = clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value)

                                            If ItemConvFactor = ItempouchCF Then
                                                If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Decimal values are not allowed for [" & clsCommon.myCstr(obj1.itemCode) & "]")
                                                End If
                                            ElseIf ItempouchCF = 6 Then
                                                If clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) Mod 6 <> 0 Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Should be in multiple of 6")
                                                End If
                                            Else
                                                If clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) Mod 0.5 <> 0 Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Should be in multiple of 0.5")
                                                End If
                                            End If
                                            If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                                Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor

                                                If DispatchQty >= (CrateConvFactor) Then
                                                    dblTotalCrateRowWise = clsCommon.myRoundOFF(DispatchQty / CrateConvFactor, 0, 9)
                                                Else
                                                    dblTotalCrateRowWise = 0
                                                End If

                                            End If
                                            'End If
                                        End If
                                        dblTotalPCrate = dblTotalPCrate + dblTotalPCrateRowWise
                                        'obj1.FreshItem_QtyInCrates = dblTotalCrateRowWise
                                    End If
                                End If
                            End If
                            dblTotalPCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                        End If
                        strItemValueExist = "Yes"
                        Dim dt As New DataTable()
                        Dim dblRate As Double = 0
                        Dim qry As String = "select tspl_customer_master.price_CodeNon from TSPL_CUSTOMER_MASTER where Cust_Code='" & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value) & "'"
                        strPriceCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                        If clsCommon.myLen(strPriceCode) <= 0 Then
                            Throw New InvalidOperationException("price_CodeNon not found for Customer " & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustName).Value) & "")
                        End If
                        If Not isLoad Then
                            qry = " Select Is_With_Tax, RowNo, Item_Price_ID, XXXE.Item_Code, UOM, Start_Date, Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, XXXE.Tax_group,XXXE.TAX1_Rate, " &
                    " XXXE.TAX2_Rate,XXXE.TAX3_Rate,XXXE.TAX4_Rate,XXXE.TAX5_Rate, " &
                    "  XXXE.TAX6_Rate,XXXE.TAX7_Rate,XXXE.TAX8_Rate,XXXE.TAX9_Rate, " &
                    " XXXE.TAX10_Rate,XXXE.TAX1 ,XXXE.TAX2,XXXE.TAX3, " &
                    " XXXE.TAX4,XXXE.TAX5,XXXE.TAX6,XXXE.TAX7, " &
                    " XXXE.TAX8,XXXE.TAX9,XXXE.TAX10,XXXE.TAX1_Amt, " &
                    " XXXE.TAX2_Amt,XXXE.TAX3_Amt,XXXE.TAX4_Amt, XXXE.TAX5_Amt,XXXE.TAX6_Amt,XXXE.TAX7_Amt,XXXE.TAX8_Amt,XXXE.TAX9_Amt,XXXE.TAX10_Amt,XXXE.Against_Plan_TR_Code  from ( " &
                    "Select ROW_NUMBER() OVER (Partition By TSPL_ITEM_PRICE_MASTER.Item_Code ORDER BY TSPL_ITEM_PRICE_MASTER.Item_Code,  " &
                    "Start_Date Desc) as RowNo,Is_With_Tax, Item_Price_ID, TSPL_ITEM_PRICE_MASTER.Item_Code, UOM, Start_Date,  " &
                    "Item_Basic_Price,Item_Basic_Net,Price_Code,Item_Selling_Price, TSPL_ITEM_PRICE_MASTER.Tax_group,TSPL_ITEM_PRICE_MASTER.TAX1_Rate,  " &
                    "TSPL_ITEM_PRICE_MASTER.TAX2_Rate,TSPL_ITEM_PRICE_MASTER.TAX3_Rate,TSPL_ITEM_PRICE_MASTER.TAX4_Rate,TSPL_ITEM_PRICE_MASTER.TAX5_Rate,  " &
                    " TSPL_ITEM_PRICE_MASTER.TAX6_Rate, TSPL_ITEM_PRICE_MASTER.TAX7_Rate, TSPL_ITEM_PRICE_MASTER.TAX8_Rate, TSPL_ITEM_PRICE_MASTER.TAX9_Rate, " &
                    " TSPL_ITEM_PRICE_MASTER.TAX10_Rate, TSPL_ITEM_PRICE_MASTER.TAX1, TSPL_ITEM_PRICE_MASTER.TAX2, TSPL_ITEM_PRICE_MASTER.TAX3, " &
                    " TSPL_ITEM_PRICE_MASTER.TAX4, TSPL_ITEM_PRICE_MASTER.TAX5, TSPL_ITEM_PRICE_MASTER.TAX6, TSPL_ITEM_PRICE_MASTER.TAX7, " &
                    " TSPL_ITEM_PRICE_MASTER.TAX8,TSPL_ITEM_PRICE_MASTER.TAX9,TSPL_ITEM_PRICE_MASTER.TAX10,TSPL_ITEM_PRICE_MASTER.TAX1_Amt , TSPL_ITEM_PRICE_MASTER.TAX2_Amt ,TSPL_ITEM_PRICE_MASTER.TAX3_Amt ,TSPL_ITEM_PRICE_MASTER.TAX4_Amt,TSPL_ITEM_PRICE_MASTER.TAX5_Amt,TSPL_ITEM_PRICE_MASTER.TAX6_Amt,TSPL_ITEM_PRICE_MASTER.TAX7_Amt,   TSPL_ITEM_PRICE_MASTER.TAX8_Amt,TSPL_ITEM_PRICE_MASTER.TAX9_Amt,TSPL_ITEM_PRICE_MASTER.TAX10_Amt,TSPL_ITEM_PRICE_MASTER.Against_Plan_TR_Code from TSPL_ITEM_PRICE_MASTER  left  outer join  " &
                    "TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_PRICE_MASTER.Item_Code=TSPL_ITEM_UOM_DETAIL.Item_Code and  " &
                    "TSPL_ITEM_PRICE_MASTER.UOM=TSPL_ITEM_UOM_DETAIL.UOM_Code   where  2=( case when CONVERT(date,Start_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,Start_Date,103)<='" + IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "' then 2 else 3 end)  end)  and (End_Date >= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "'  or End_date is null)  and  " &
                    "TSPL_ITEM_PRICE_MASTER.Price_Code='" & strPriceCode & "' and UOM='" & obj1.Unit_code & "' and TSPL_ITEM_PRICE_MASTER.item_code='" & obj1.itemCode & "' AND Location_Code='" & clsCommon.myCstr(txtLocation.Value) & "'  " &
                    ") XXXE WHERE RowNo=1  "
                        Else
                            qry = "select item_Rate as Item_Basic_Price,tax_group,TAX1,TAX2,TAX3,TAX4,TAX5,TAX6,TAX7,TAX8,TAX9,TAX10,TAX1_Rate,TAX2_Rate,TAX3_Rate,TAX4_Rate, TAX5_Rate,TAX6_Rate,TAX7_Rate,TAX8_Rate,TAX9_Rate,TAX10_Rate,TAX1_Amt,TAX2_Amt,TAX3_Amt,TAX4_Amt,TAX5_Amt,TAX6_Amt,TAX7_Amt,TAX8_Amt,TAX9_Amt,TAX10_Amt from TSPL_DEMAND_BOOKING_DETAIL where  Document_No='" & txtDocNo.Value & "' and Item_Code='" & obj1.itemCode & "' and Unit_code='" & obj1.Unit_code & "' and Cust_Code='" & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value) & "'"
                        End If
                        dt = clsDBFuncationality.GetDataTable(qry)
                        If dt.Rows.Count > 0 Then
                            Dim objCustItem As clsDemandCustItem = New clsDemandCustItem()

                            dblRate = clsCommon.myCdbl(dt.Rows(0).Item("Item_Basic_Price"))
                            If dblRate = 0 Then
                                Throw New InvalidOperationException("Please Fill Selling Price for Location " & txtLocation.Value & "  for item " & clsCommon.myCstr(obj1.ShortDesc) & Environment.NewLine)
                            End If
                            If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                                'obj1.ItemRate = dblRate
                                dblTotalMAmt = dblTotalMAmt + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * clsCommon.myCdbl(dblRate)

                            Else
                                'obj1.ItemRate = dblRate
                                dblTotalPAmt = dblTotalPAmt + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * clsCommon.myCdbl(dblRate)
                            End If
                            objCustItem.ItemRate = dblRate
                            'obj1.ItemTotAmt = Math.Round(clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * clsCommon.myCdbl(dblRate), 2)
                            objCustItem.Cust_Code = clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value)
                            objCustItem.FreshItem_QtyInLitres = obj1.FreshItem_QtyInLitres
                            objCustItem.FreshItem_QtyInCrate = obj1.FreshItem_QtyInCrates
                            objCustItem.ItemCode = obj1.itemCode
                            objCustItem.UnitCode = obj1.Unit_code
                            objCustItem.ItemTotAmt = Math.Round(clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * clsCommon.myCdbl(dblRate), 2)
                            objCustItem.TAX_Group = clsCommon.myCstr(dt.Rows(0).Item("TAX_Group"))
                            dblTotalDocAmtRowWise = dblTotalDocAmtRowWise + objCustItem.ItemTotAmt

                            objCustItem.TAX1 = clsCommon.myCstr(dt.Rows(0).Item("TAX1"))
                            If clsCommon.CompairString(objCustItem.TAX1, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX1_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX1_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX1_Rate"))
                            End If
                            objCustItem.TAX1_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX1_Rate / 100), 2)
                            objCustItem.TAX1_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX2 = clsCommon.myCstr(dt.Rows(0).Item("TAX2"))
                            If clsCommon.CompairString(objCustItem.TAX2, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX2_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX2_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX2_Rate"))
                            End If
                            objCustItem.TAX2_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX2_Rate / 100), 2)
                            objCustItem.TAX2_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX3 = clsCommon.myCstr(dt.Rows(0).Item("TAX3"))
                            If clsCommon.CompairString(objCustItem.TAX3, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX3_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX3_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX3_Rate"))
                            End If
                            objCustItem.TAX3_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX3_Rate / 100), 2)
                            objCustItem.TAX3_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX4 = clsCommon.myCstr(dt.Rows(0).Item("TAX4"))
                            If clsCommon.CompairString(objCustItem.TAX4, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX4_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX4_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX4_Rate"))
                            End If
                            objCustItem.TAX4_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX4_Rate / 100), 2)
                            objCustItem.TAX4_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX5 = clsCommon.myCstr(dt.Rows(0).Item("TAX5"))
                            If clsCommon.CompairString(objCustItem.TAX5, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX5_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX5_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX5_Rate"))
                            End If
                            objCustItem.TAX5_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX5_Rate / 100), 2)
                            objCustItem.TAX5_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX6 = clsCommon.myCstr(dt.Rows(0).Item("TAX6"))
                            If clsCommon.CompairString(objCustItem.TAX6, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX6_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX6_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX6_Rate"))
                            End If
                            objCustItem.TAX6_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX6_Rate / 100), 2)
                            objCustItem.TAX6_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX7 = clsCommon.myCstr(dt.Rows(0).Item("TAX7"))
                            If clsCommon.CompairString(objCustItem.TAX7, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX7_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX7_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX7_Rate"))
                            End If
                            objCustItem.TAX7_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX7_Rate / 100), 2)
                            objCustItem.TAX7_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX8 = clsCommon.myCstr(dt.Rows(0).Item("TAX8"))
                            If clsCommon.CompairString(objCustItem.TAX8, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX8_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX8_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX8_Rate"))
                            End If
                            objCustItem.TAX8_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX8_Rate / 100), 2)
                            objCustItem.TAX8_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX9 = clsCommon.myCstr(dt.Rows(0).Item("TAX9"))
                            If clsCommon.CompairString(objCustItem.TAX9, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX9_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX9_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX9_Rate"))
                            End If
                            objCustItem.TAX9_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX9_Rate / 100), 2)
                            objCustItem.TAX9_Base_Amt = objCustItem.ItemTotAmt
                            objCustItem.TAX10 = clsCommon.myCstr(dt.Rows(0).Item("TAX10"))
                            If clsCommon.CompairString(objCustItem.TAX10, "TCS") = CompairStringResult.Equal Then
                                objCustItem.TAX10_Rate = CalculateTCS(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value))
                            Else
                                objCustItem.TAX10_Rate = clsCommon.myCdbl(dt.Rows(0).Item("TAX10_Rate"))
                            End If
                            objCustItem.TAX10_Amt = Math.Round(objCustItem.ItemTotAmt * (objCustItem.TAX10_Rate / 100), 2)
                            objCustItem.TAX10_Base_Amt = objCustItem.ItemTotAmt
                            lstCustItem.Add(objCustItem)
                        Else
                            gv1.Rows(dblrows).Cells(dblcolumns).Value = 0
                            Throw New InvalidOperationException("Please create Price for Item " & obj1.ShortDesc & " and customer " & clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustName).Value) & " ")
                        End If
                        'End If
                    End If
                Next

                gv1.Rows(dblrows).Cells(colCrate).Value = Math.Round(clsCommon.myCdbl(TotalCrate), 2)
                gv1.Rows(dblrows).Cells(colLitre).Value = Math.Round(clsCommon.myCdbl(TotalLitre), 2)
                If isDepartmentRoute AndAlso ApplyDepartmentRoute Then
                    'If ApplyDepartmentRoute Then
                    gv1.Rows(dblrows).Cells(colCrate).Value = IIf(clsCommon.myRoundOFF(clsCommon.myCdbl(TotalLitre / 12), 0, 9) >= 1, clsCommon.myRoundOFF(clsCommon.myCdbl(TotalLitre / 12), 0, 9) - 1, 0)
                    'End If
                End If
                gv1.Rows(dblrows).Cells(colMAmt).Value = clsCommon.myRoundOFF(clsCommon.myCdbl(dblTotalMAmt), 2, 5)
                gv1.Rows(dblrows).Cells(colPCount).Value = clsCommon.myCdbl(dblTotalPCount)
                If AllowMultipleUOMForProduct Then
                    gv1.Rows(dblrows).Cells(colPCrate).Value = clsCommon.myCdbl(dblTotalPCrate)
                End If
                gv1.Rows(dblrows).Cells(colPAmt).Value = clsCommon.myRoundOFF(clsCommon.myCdbl(dblTotalPAmt), 2, 5)
                gv1.Rows(dblrows).Cells(colAmt).Value = clsCommon.myRoundOFF(clsCommon.myCdbl(dblTotalDocAmtRowWise), 2, 5)
                If clsCommon.myLen(gv1.Rows(dblrows).Cells(colItemExist)) > 0 Then
                    gv1.Rows(dblrows).Cells(colItemExist).Value = strItemValueExist
                End If
                dblDocTotalAmt = dblDocTotalAmt + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colAmt).Value)
                dblDocTotalLitre = dblDocTotalLitre + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                dblDocTotalCrate = dblDocTotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                TotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                TotalPCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                TotalPCrate += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCrate).Value)
                TotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
            Next
            lblTotalCrate.Text = clsCommon.myCdbl(dblDocTotalCrate)
            lblTotalLitre.Text = clsCommon.myCdbl(dblDocTotalLitre)
            txtDocAmt.Text = clsCommon.myCdbl(dblDocTotalAmt)
            lblDocumentAmt.Text = clsCommon.myCdbl(TotalMAmt)
            txtPCount.Text = clsCommon.myCdbl(TotalPCount)
            txtPAmt.Text = clsCommon.myCdbl(TotalPAmt)
            txtTotalPCrate.Text = clsCommon.myCdbl(TotalPCrate)
            UpdateColumnTotal()
            gv1.Rows(gv1.Rows.Count - 1).IsPinned = True
            gv1.Rows(gv1.Rows.Count - 1).PinPosition = PinnedRowPosition.Bottom
            isInsideLoadData = False
        Catch ex As InvalidOperationException
            Throw New InvalidOperationException(ex.Message)
        End Try
        Return lstCustItem
    End Function

    Function CalculateTCS(ByVal CustCode As String) As Double
        Dim TCSBaseAmount As Double = 0
        Dim TCSTaxRate As Double = 0

        Dim balanceAmt As Double = 0
        Dim OPInvoice_Sale_Amt As Double = 0
        Dim CurrFinYR As String = String.Empty
        Dim FinancialYear As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT CASE WHEN DatePart(Month, '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "') >= 4 THEN DatePart(Year, '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "') + 1 ELSE DatePart(Year, '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "') END AS Fiscal_Year"))
        Dim strStartDate As Date = "01/Apr/" & clsCommon.myCstr(clsCommon.myCdbl(FinancialYear - 1))
        Dim strEndDate As Date = "31/Mar/" & FinancialYear
        CurrFinYR = clsCommon.myCstr(clsCommon.myCdbl(FinancialYear - 1)) & "-" & FinancialYear
        TCSBaseAmount = 0
        Dim strqry As String = "select sum(ItemNetAmount) from TSPL_DEMAND_BOOKING_MASTER left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No where TSPL_DEMAND_BOOKING_DETAIL.Cust_Code='" & clsCommon.myCstr(CustCode) & "' and convert(date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)>='" & clsCommon.GetPrintDate(strStartDate) & "' and convert(date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)<='" & clsCommon.GetPrintDate(strEndDate) & "' "
        balanceAmt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strqry))
        OPInvoice_Sale_Amt = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Sale_amt from TSPL_OP_INVOICE_FOR_TCS where Customer_Code='" & clsCommon.myCstr(CustCode) & "' and Financial_Year_Code='" & clsCommon.myCstr(CurrFinYR) & "'"))
        TCSBaseAmount = OPInvoice_Sale_Amt + balanceAmt
        If AmountToCheckCustomerOutstandingForTCSTax > 0 Then
            If TCSBaseAmount > AmountToCheckCustomerOutstandingForTCSTax Then

                If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing)) = 1 Then
                    Dim Is_ITR_Filled_And_TCSAmountGreater50K As Boolean = IIf(clsCommon.myCstr(clsDBFuncationality.getSingleValue(" SELECT CASE WHEN ISNULL(IsTCSGreaterthan50K,0)=1 AND ISNULL(IsITRfilledinLast2Years,0)=1 THEN 1 ELSE 0 END FROM TSPL_CUSTOMER_MASTER WHERE Cust_Code='" & clsCommon.myCstr(CustCode) & "'")) = 1, True, False)
                    If Is_ITR_Filled_And_TCSAmountGreater50K Then
                        TCSTaxRate = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                    Else
                        TCSTaxRate = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                    End If
                Else
                    Dim panno As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(pan,'')+isnull(Additional3 ,'') as PanNoAdhar from tspl_customer_master where cust_code='" & clsCommon.myCstr(CustCode) & "'"))
                    If clsCommon.myLen(panno) > 0 Then
                        TCSTaxRate = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithPanNo, clsFixedParameterCode.TCSRateforCustomerWithPanNo, Nothing))
                    Else
                        TCSTaxRate = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TCSRateforCustomerWithoutPanNo, clsFixedParameterCode.TCSRateforCustomerWithoutPanNo, Nothing))
                    End If
                End If
            Else
                TCSTaxRate = 0
            End If
        Else
            TCSTaxRate = 0
        End If
        Return TCSTaxRate
    End Function
    Private Sub UpdateRowTotal()
        Try
            Dim TotalCrate As Double = 0
            Dim TotalLitre As Double = 0
            Dim TotalPCount As Double = 0
            Dim TotalPAmt As Double = 0
            'Dim TotalPCrate As Double = 0
            Dim TotalMAmt As Double = 0
            Dim dblTotalDocAmtRowWise As Decimal = 0
            Dim dblTotalCrateRowWise As Double = 0
            Dim dblTotalPCrateRowWise As Double = 0
            Dim dblTotalLitreRowWise As Double = 0
            'Dim strPriceCode As String = String.Empty
            Dim strItemValueExist As String = String.Empty
            Dim strItemUpdateAfterSave As String = String.Empty
            Dim dblDocTotalCrate As Double = 0
            Dim dblDocTotalLitre As Double = 0
            Dim dblDocTotalAmt As Double = 0
            Dim dblTotalPCount As Double = 0
            Dim dblTotalPCrate As Double = 0
            'Dim dblToalPouchCount As Double = 0
            Dim dblTotalPAmt As Double = 0
            Dim dblTotalMAmt As Double = 0
            'Dim colTotalQty As Double = 0
            Dim status As Boolean = False
            For dblrows As Integer = 0 To gv1.Rows.Count - 2
                Dim k As Integer = 1
                TotalCrate = 0
                dblTotalPCrate = 0
                dblTotalCrateRowWise = 0
                dblTotalPCrateRowWise = 0
                TotalLitre = 0
                dblTotalLitreRowWise = 0
                dblTotalDocAmtRowWise = 0
                dblTotalPCount = 0
                dblTotalPAmt = 0
                dblTotalMAmt = 0
                strItemValueExist = "No"
                strItemUpdateAfterSave = "No"
                For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                    Dim obj1 As ItemValueClass = Nothing
                    Try
                        obj1 = TryCast(gv1.Columns(colItemCode & clsCommon.myCstr(k)).Tag, ItemValueClass)
                    Catch ex As InvalidOperationException
                        Throw New InvalidOperationException(ex.Message)
                    End Try
                    k = k + 1
                    If obj1 IsNot Nothing AndAlso clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 Then
                        'If clsCommon.myLen(clsCommon.myCstr(obj1.itemCode)) > 0 AndAlso clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) > 0 Then
                        If clsCommon.CompairString(clsCommon.myCstr(obj1.IsFreshAmbient), "Fresh") = CompairStringResult.Equal Then
                            If clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code), "Crate") = CompairStringResult.Equal Then
                                Dim cellValue As String = clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value)

                                Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                'Dim ItempouchCF As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))
                                'Dim cellValue As String = clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                Dim DispatchQty As Decimal = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                                If AllowRouteWiseDemandEntryInDecimal Then

                                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_Route_Master where Route_No='" & txtRouteNo.Value & "' and  AllowEntryInDecimal=1 ")) = 0 Then
                                        If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                            Throw New InvalidOperationException("Decimal values are not allowed.")
                                        End If
                                    Else
                                        If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                            Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                            If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                            End If
                                        End If

                                    End If

                                Else
                                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_ITEM_MASTER where Item_Code='" & obj1.itemCode & "' and  AllowEntryInDecimal=1")) = 0 Then
                                        If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                            Throw New InvalidOperationException("Decimal values are not allowed.")
                                        End If
                                    Else
                                        If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                            Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                            If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                            End If
                                        End If

                                    End If
                                End If


                                If ApplyItemCapacityLimit Then
                                    status = CheckItemCapacityLimit(clsCommon.myCstr(obj1.itemCode), clsCommon.myCstr(obj1.Unit_code), cellValue, dblrows, dblcolumns)
                                    If Not status Then
                                        gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                    End If
                                End If
                                TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                obj1.FreshItem_QtyInCrates = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                'gv1.Rows(gv1.Rows.Count - 1).Cells(dblcolumns).Value = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                            Else
                                Dim ItemCrateType As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                                If ItemCrateType = 1 Then
                                    'Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                    Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                    Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                    Dim ItempouchCF As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))
                                    Dim cellValue As String = clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal AndAlso CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                        'If CrateConvFactor > 0 And ItemConvFactor > 0 Then
                                        Dim DispatchQty As Decimal = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                                        If AllowRouteWiseDemandEntryInDecimal Then
                                            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_Route_Master where Route_No='" & txtRouteNo.Value & "' and  AllowEntryInDecimal=1 ")) = 0 Then
                                                If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Decimal values are not allowed.")
                                                End If
                                            Else
                                                Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                                If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                                End If
                                            End If
                                        Else
                                            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_ITEM_MASTER where Item_Code='" & obj1.itemCode & "' and  AllowEntryInDecimal=1")) = 0 Then
                                                If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Decimal values are not allowed.")
                                                End If
                                            Else
                                                Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                                If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                    Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                                End If
                                            End If
                                        End If

                                        If ConvertPouchtoCrate Then
                                            If DispatchQty > (CrateConvFactor / 2) Then
                                                dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                            Else
                                                dblTotalCrateRowWise = 1
                                            End If
                                        ElseIf DontCreateForPouch Then
                                            dblTotalCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                        Else
                                            If DispatchQty > (CrateConvFactor / 2) Then
                                                dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                            Else
                                                dblTotalCrateRowWise = 0
                                            End If
                                        End If
                                        'End If


                                    End If
                                    If ApplyItemCapacityLimit Then
                                        status = CheckItemCapacityLimit(clsCommon.myCstr(obj1.itemCode), clsCommon.myCstr(obj1.Unit_code), cellValue, dblrows, dblcolumns)
                                        If Not status Then
                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                        End If
                                    End If
                                    If isDepartmentRoute AndAlso ApplyDepartmentRoute Then
                                        'If ApplyDepartmentRoute Then


                                        If ItemConvFactor = ItempouchCF Then
                                            If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Decimal values are not allowed for [" & clsCommon.myCstr(obj1.itemCode) & "]")
                                            End If
                                        ElseIf ItempouchCF = 6 Then
                                            If clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) Mod 6 <> 0 Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                            End If
                                        Else
                                            If clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) Mod 0.5 <> 0 Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                            End If
                                        End If
                                        If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                            Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor

                                            If DispatchQty >= (CrateConvFactor) Then
                                                dblTotalCrateRowWise = clsCommon.myRoundOFF(DispatchQty / CrateConvFactor, 0, 9)
                                            Else
                                                dblTotalCrateRowWise = 0
                                            End If

                                        End If
                                        'End If
                                    End If

                                    TotalCrate = TotalCrate + dblTotalCrateRowWise
                                    obj1.FreshItem_QtyInCrates = dblTotalCrateRowWise
                                End If
                            End If
                            ''to convert into litre
                            Dim CrateConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code='Ltr' "))
                            Dim ItemConvFactor_Ltr As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                            If CrateConvFactor_Ltr > 0 AndAlso ItemConvFactor_Ltr > 0 Then
                                Dim DispatchQty As Double = clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor_Ltr
                                dblTotalLitreRowWise = (DispatchQty / CrateConvFactor_Ltr)
                            End If
                            TotalLitre = TotalLitre + dblTotalLitreRowWise
                            obj1.FreshItem_QtyInLitres = dblTotalLitreRowWise
                            ''---------end of litre conversion
                        Else
                            Dim ItemCrateType As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select IS_CrateType  from TSPL_ITEM_MASTER Where Item_Code  ='" & clsCommon.myCstr(obj1.itemCode) & "'"))
                            If ItemCrateType = 1 Then
                                If AllowMultipleUOMForProduct Then
                                    If clsCommon.CompairString(clsCommon.myCstr(obj1.Unit_code), "Crate") = CompairStringResult.Equal Then
                                        TotalCrate = TotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                                    Else

                                        If ItemCrateType = 1 Then
                                            'Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                            Dim CrateConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and tspl_unit_master.Crate_Type ='Y' "))
                                            Dim ItemConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                            Dim cellValue As String = clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value)

                                            'If Not clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                                            If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                                Dim DispatchQty As Decimal = clsCommon.myCDecimal(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                                                'If ConvertPouchtoCrate Then
                                                If DispatchQty > (CrateConvFactor / 2) Then
                                                    dblTotalPCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                                Else
                                                    dblTotalPCrateRowWise = 0
                                                End If
                                                If AllowRouteWiseDemandEntryInDecimal Then
                                                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_Route_Master where Route_No='" & txtRouteNo.Value & "' and  AllowEntryInDecimal=1 ")) = 0 Then
                                                        If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                            Throw New InvalidOperationException("Decimal values are not allowed.")
                                                        End If
                                                    Else
                                                        Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                                        If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                            Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                                        End If
                                                    End If
                                                Else
                                                    If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_ITEM_MASTER where Item_Code='" & obj1.itemCode & "' and  AllowEntryInDecimal=1")) = 0 Then
                                                        If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                            Throw New InvalidOperationException("Decimal values are not allowed.")
                                                        End If
                                                    Else
                                                        Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                                        If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                            gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                            Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                                        End If
                                                    End If
                                                End If

                                                'ElseIf DontCreateForPouch Then
                                                '    dblTotalCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                                'Else
                                                '    If DispatchQty > (CrateConvFactor / 2) Then
                                                '        dblTotalCrateRowWise = Math.Ceiling(DispatchQty / CrateConvFactor)
                                                '    Else
                                                '        dblTotalCrateRowWise = 0
                                                '    End If
                                                'End If
                                            End If
                                            'End If
                                            If ApplyItemCapacityLimit Then
                                                status = CheckItemCapacityLimit(clsCommon.myCstr(obj1.itemCode), clsCommon.myCstr(obj1.Unit_code), clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value), dblrows, dblcolumns)
                                                If Not status Then
                                                    gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                End If
                                            End If
                                            dblTotalPCrate = dblTotalPCrate + dblTotalPCrateRowWise
                                            'obj1.FreshItem_QtyInCrates = dblTotalCrateRowWise
                                        End If
                                    End If

                                End If
                            Else
                                'Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(obj1.Unit_code) & "'"))
                                Dim CrateConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' "))
                                Dim ItemConvFactor As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(obj1.Unit_code) & "' "))
                                Dim cellValue As String = clsCommon.myCstr(gv1.Rows(dblrows).Cells(dblcolumns).Value)

                                'If Not clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                                If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                                    Dim DispatchQty As Decimal = clsCommon.myCDecimal(gv1.Rows(dblrows).Cells(dblcolumns).Value) * ItemConvFactor
                                    'If ConvertPouchtoCrate Then
                                    If DispatchQty > (CrateConvFactor / 2) Then
                                        dblTotalPCrateRowWise = Math.Floor(DispatchQty / CrateConvFactor)
                                    Else
                                        dblTotalPCrateRowWise = 0
                                    End If
                                    If AllowRouteWiseDemandEntryInDecimal Then
                                        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_Route_Master where Route_No='" & txtRouteNo.Value & "' and  AllowEntryInDecimal=1 ")) = 0 Then
                                            If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Decimal values are not allowed.")
                                            End If
                                        Else
                                            Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                            If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                            End If
                                        End If
                                    Else
                                        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(AllowEntryInDecimal) from TSPL_ITEM_MASTER where Item_Code='" & obj1.itemCode & "' and  AllowEntryInDecimal=1")) = 0 Then
                                            If cellValue.Contains(".") OrElse cellValue.Contains(",") Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Decimal values are not allowed.")
                                            End If
                                        Else
                                            Dim ItemCFofDecimalUom As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(obj1.itemCode) & "' and TSPL_ITEM_UOM_DETAIL.Decimal_UOM ='1'"))

                                            If DispatchQty Mod ItemCFofDecimalUom <> 0 Then
                                                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                                                Throw New InvalidOperationException("Please Enter Valid Qty for " & obj1.ShortDesc)
                                            End If
                                        End If

                                    End If


                                End If
                            End If

                            dblTotalPCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(dblcolumns).Value)
                        End If
                        'End If
                    End If
                Next

                gv1.Rows(dblrows).Cells(colCrate).Value = Math.Round(clsCommon.myCdbl(TotalCrate), 2)
                gv1.Rows(dblrows).Cells(colLitre).Value = Math.Round(clsCommon.myCdbl(TotalLitre), 2)
                gv1.Rows(dblrows).Cells(colPCount).Value = clsCommon.myCdbl(dblTotalPCount)
                If AllowMultipleUOMForProduct Then
                    gv1.Rows(dblrows).Cells(colPCrate).Value = clsCommon.myCdbl(dblTotalPCrate)

                End If
                If clsCommon.myLen(gv1.Rows(dblrows).Cells(colItemExist)) > 0 Then
                    gv1.Rows(dblrows).Cells(colItemExist).Value = strItemValueExist
                End If
                'If isDepartmentRoute Then
                '    If ApplyDepartmentRoute Then
                '        gv1.Rows(dblrows).Cells(colCrate).Value = IIf(clsCommon.myRoundOFF(clsCommon.myCdbl(TotalLitre / 12), 0, 9) >= 1, clsCommon.myRoundOFF(clsCommon.myCdbl(TotalLitre / 12), 0, 9) - 1, 0)
                '    End If
                'End If
                dblDocTotalAmt = dblDocTotalAmt + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colAmt).Value)
                dblDocTotalLitre = dblDocTotalLitre + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colLitre).Value)
                dblDocTotalCrate = dblDocTotalCrate + clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colCrate).Value)
                TotalPAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPAmt).Value)
                TotalPCount += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colPCount).Value)
                TotalMAmt += clsCommon.myCdbl(gv1.Rows(dblrows).Cells(colMAmt).Value)
            Next
            UpdateColumnTotal()
            gv1.Rows(gv1.Rows.Count - 1).IsPinned = True
            gv1.Rows(gv1.Rows.Count - 1).PinPosition = PinnedRowPosition.Bottom
        Catch ex As InvalidOperationException
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Function CheckItemCapacityLimit(ByVal strItemCode As String, ByVal strUOM As String, ByVal qty As String, ByVal dblrows As Integer, ByVal dblcolumns As Integer) As Boolean
        Dim status As Boolean = False
        Try
            Dim strQry As String = "select TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Item_Code,TSPL_ITEM_CAPACITY_LIMIT_DETAIL.UOM,TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Qty  from TSPL_ITEM_CAPACITY_LIMIT_HEAD
left join TSPL_ITEM_CAPACITY_LIMIT_DETAIL on TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Document_No=TSPL_ITEM_CAPACITY_LIMIT_head.Document_No
                where TSPL_ITEM_CAPACITY_LIMIT_HEAD.Posted=1 and TSPL_ITEM_CAPACITY_LIMIT_HEAD.Document_No in(
select top 1 TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Document_No from TSPL_ITEM_CAPACITY_LIMIT_HEAD
left join TSPL_ITEM_CAPACITY_LIMIT_DETAIL on TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Document_No=TSPL_ITEM_CAPACITY_LIMIT_head.Document_No
where TSPL_ITEM_CAPACITY_LIMIT_head.From_Date<='" & clsCommon.GetPrintDate(txtDate.Value) & "' and TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Item_Code='" & strItemCode & "'  and 2=(Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date is null then 2 else (Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date>='" & clsCommon.GetPrintDate(txtDate.Value) & "' then 2 else 3 end) end) order by From_Date desc)
 and TSPL_ITEM_CAPACITY_LIMIT_head.From_Date<='" & clsCommon.GetPrintDate(txtDate.Value) & "' and TSPL_ITEM_CAPACITY_LIMIT_DETAIL.Item_Code='" & strItemCode & "' and 2=(Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date is null then 2 else (Case when TSPL_ITEM_CAPACITY_LIMIT_head.To_Date>='" & clsCommon.GetPrintDate(txtDate.Value) & "' then 2 else 3 end) end)"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(strQry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                'Dim IsStockingUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Stocking_Unit from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code  ='" & clsCommon.myCstr(dt.Rows(0)("UOM")) & "'"))
                Dim CrateConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and tspl_unit_master.Crate_Type ='Y' "))
                Dim ItemConvFactor As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(strUOM) & "' "))
                Dim ItemLimitCF As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='" & clsCommon.myCstr(dt.Rows(0)("UOM")) & "' "))
                'Dim ItemLimitCFPouch As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Conversion_Factor  from TSPL_ITEM_UOM_DETAIL Left Outer Join tspl_unit_master on tspl_unit_master.Unit_Code = TSPL_ITEM_UOM_DETAIL.UOM_Code Where TSPL_ITEM_UOM_DETAIL.Item_Code ='" & clsCommon.myCstr(dt.Rows(0)("Item_Code")) & "' and TSPL_ITEM_UOM_DETAIL.UOM_Code ='Pouch' "))
                'Dim cellValue As String = clsCommon.myCstr(qty)
                If CrateConvFactor > 0 AndAlso ItemConvFactor > 0 Then
                    Dim DispatchQty As Decimal = (clsCommon.myCdbl(qty) * ItemConvFactor) / ItemLimitCF
                    If DispatchQty > clsCommon.myCdbl(dt.Rows(0)("Qty")) Then
                        gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                        Throw New InvalidOperationException("The maximum allowed quantity for this item is [" & clsCommon.myCstr(clsCommon.myCDecimal(dt.Rows(0)("Qty")) / (ItemConvFactor / ItemLimitCF)) & " ] ." & Environment.NewLine & " Please reduce the quantity To [" & clsCommon.myCstr(clsCommon.myCDecimal(dt.Rows(0)("Qty")) / (ItemConvFactor / ItemLimitCF)) & "] Or less To proceed.")
                    Else
                        status = True
                    End If
                End If
            Else
                gv1.Rows(dblrows).Cells(dblcolumns).Value = ""
                Throw New InvalidOperationException("Capacity Limit Not found For [ " & strItemCode & " ]")
            End If

        Catch ex As InvalidOperationException
            Throw New InvalidOperationException(ex.Message)
        End Try
        Return status
    End Function
    Private Sub UpdateColumnTotal()
        Try
            Dim TotalQty As Double = 0
            'For dbrows1 As Integer = 0 To gv1.Rows.Count - 1
            For dblcolumns As Integer = 9 To gv1.Columns.Count - 1
                TotalQty = 0
                For dbrows As Integer = 0 To gv1.Rows.Count - 2
                    TotalQty += clsCommon.myCdbl(gv1.Rows(dbrows).Cells(dblcolumns).Value)
                Next
                gv1.Rows(gv1.Rows.Count - 1).Cells(dblcolumns).Value = TotalQty
            Next
            'Next
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnUpdateCrateAndAmt_Click(sender As Object, e As EventArgs) Handles btnUpdateCrateAndAmt.Click
        Try
            UpdateAllTotals(False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnAssessment_Click(sender As Object, e As EventArgs) Handles btnAssessment.Click
        Dim frm As New frmAssessmentGrid
        frm.IDate = txtDate.Value
        frm.ShowDialog()
    End Sub
    Private Sub txtVehicleNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVehicleNo._MYValidating
        Dim whrcls As String = ""
        Dim qry As String = "Select vehicle_id, Description, route_no as 'Route No',route_desc as 'Route Name'  from TSPL_VEHICLE_MASTER left join tspl_route_master on tspl_route_master.vehicle_code=TSPL_VEHICLE_MASTER.vehicle_id "
        If clsCommon.myLen(clsCommon.myCstr(txtRouteNo.Value)) > 0 Then
            whrcls = " tspl_route_master.route_no ='" & txtRouteNo.Value & "' "
        End If
        txtVehicleNo.Value = clsCommon.ShowSelectForm("DBookingVehicle", qry, "vehicle_id", whrcls, txtVehicleNo.Value, "vehicle_id", isButtonClicked)
        lblVehicleNo.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Description  from TSPL_VEHICLE_MASTER where Vehicle_Id = '" + Convert.ToString(txtVehicleNo.Value) + "'"))
    End Sub
    Private Sub chkIndividualCustomer_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkIndividualCustomer.ToggleStateChanged
        If chkIndividualCustomer.Checked Then
            txtCustomerNo.Enabled = True
        Else
            txtCustomerNo.Enabled = False
        End If
    End Sub
    Private Sub txtCustomerNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtCustomerNo._MYValidating
        Try
            Dim qry As String = "select Cust_Code as Code,TSPL_CUSTOMER_MASTER.Customer_Name as [Customer Name],TSPL_CUSTOMER_MASTER.Alies_Name as [Short Name],TSPL_CUSTOMER_MASTER.Route_No"
            qry += " ,TSPL_CUSTOMER_MASTER.Zone_Code as [Zone Code],TSPL_ZONE_MASTER.Description as [Zone Name],isnull(TSPL_CUSTOMER_MASTER.Customer_Category,'') as [Customer Category],tspl_customer_master.Cust_Group_Code as [Customer Group Code] "
            qry += " from TSPL_CUSTOMER_MASTER "
            qry += " left outer join TSPL_CITY_MASTER on TSPL_CITY_MASTER.City_Code=TSPL_CUSTOMER_MASTER.City_Code"
            qry += " left join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_CUSTOMER_MASTER.Zone_Code"
            Dim WhrCls As String = ""
            If clsCommon.myLen(txtRouteNo.Value) > 0 Then
                WhrCls = "  TSPL_CUSTOMER_MASTER.Status='N' and TSPL_CUSTOMER_MASTER.Route_No='" & txtRouteNo.Value & "' "
            Else
                WhrCls = "  TSPL_CUSTOMER_MASTER.Status='N' "
            End If
            '-------richa 17/12/2019 show customer according to custoer permission Ticket No. ---------
            Dim strwherecls As String = ""
            strwherecls = Xtra.CustomerPermission()
            If clsCommon.myLen(strwherecls) > 0 Then
                WhrCls += " and TSPL_CUSTOMER_MASTER.Cust_Code in (" + strwherecls + ")"
            End If
            txtCustomerNo.Value = clsCommon.ShowSelectForm("CustomerFnder1", qry, "Code", WhrCls, txtCustomerNo.Value, "Code", isButtonClicked)
            lblCustomerName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Customer_Name from TSPL_CUSTOMER_MASTER where Cust_Code='" + txtCustomerNo.Value + "'"))
            If clsCommon.myLen(txtCustomerNo.Value) > 0 Then
                setRouteVehicleCityDetail()
            End If
            If ApplyItemUOMOnDemand Then
                SetRouteColumns()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function ReturnDistributorRouteTaggingQry() As String
        Dim qry As String = "  select top 1 TSPL_ROUTE_MASTER.City_Code,Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_CUSTOMER_MASTER.Zone_Code,
 TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon,tspl_transport_master.Transporter_Name 
 from TSPL_DISTRIBUTOR_ROUTE_CUSTOMER 
 left join TSPL_DISTRIBUTOR_ROUTE on TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code=TSPL_DISTRIBUTOR_ROUTE.Code
 LEFT outer join tspl_customer_master On tspl_customer_master.Cust_Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code
 left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No =TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No
 left outer join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id 
 left outer join tspl_transport_master on tspl_transport_master.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id 
 where TSPL_DISTRIBUTOR_ROUTE.Status=1 and TSPL_DISTRIBUTOR_ROUTE.ItemType='M' and TSPL_DISTRIBUTOR_ROUTE.Start_Date<='" & clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy") & "' "
        If chkIndividualCustomer.Checked Then
            qry += " And TSPL_CUSTOMER_MASTER.Cust_Code ='" & txtCustomerNo.Value & "'  "
        ElseIf clsCommon.myLen(txtVehicleNo.Value) > 0 Then
            qry &= " and TSPL_VEHICLE_MASTER.vehicle_id='" & txtVehicleNo.Value & "'"
        Else
            qry += " and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" & txtRouteNo.Value & "' "
        End If
        qry &= " and 2=(Case when TSPL_DISTRIBUTOR_ROUTE.End_Date is null then 2 else (Case when TSPL_DISTRIBUTOR_ROUTE.End_Date>='" & clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy") & "' then 2 else 3 end) end) 
 order by TSPL_DISTRIBUTOR_ROUTE.Start_Date desc "
        Return qry
    End Function

    Sub setRouteVehicleCityDetail()
        Try
            Dim qry As String = ""
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal AndAlso Not chkIndividualCustomer.Checked Then
                qry = ReturnDistributorRouteTaggingQry()
            Else
                If objCommonVar.ApplyBoothRouteMapping Then
                    qry = "select top 1 TSPL_ROUTE_MASTER.City_Code,Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_CUSTOMER_MASTER.Zone_Code,
TSPL_BOOTH_ROUTE_MAPPING_HEAD.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon,tspl_transport_master.Transporter_Name
from TSPL_BOOTH_ROUTE_MAPPING_HEAD
left join TSPL_BOOTH_ROUTE_MAPPING_DETAIL on TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Document_No=TSPL_BOOTH_ROUTE_MAPPING_HEAD.Document_No
LEFT join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_BOOTH_ROUTE_MAPPING_DETAIL.Booth_Code
left outer join TSPL_ROUTE_MASTER on TSPL_BOOTH_ROUTE_MAPPING_HEAD.Route_No=TSPL_ROUTE_MASTER.Route_No
left outer join TSPL_VEHICLE_MASTER on  TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id 
left outer join tspl_transport_master on tspl_transport_master.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id "
                    If chkIndividualCustomer.Checked Then
                        qry += " where TSPL_CUSTOMER_MASTER.Cust_Code ='" & txtCustomerNo.Value & "'  "
                    Else
                        qry += " where TSPL_ROUTE_MASTER.Route_No ='" & txtRouteNo.Value & "' "
                    End If
                    qry += " and isnull(TSPL_Booth_Route_Mapping_Head.Posted,0)=1 and Item_Type='Milk' and 2=( case when CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "' and Shift_Type='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' then 2 else ( case when CONVERT(date,TSPL_Booth_Route_Mapping_Head.Supply_Date,103)<='" + IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "' then 2 else 3 end)  end) order by TSPL_BOOTH_ROUTE_MAPPING_HEAD.Document_No desc"
                Else
                    qry = "select TSPL_ROUTE_MASTER.City_Code,Customer_Name,vehicle_code,TSPL_VEHICLE_MASTER.Vehicle_No,TSPL_CUSTOMER_MASTER.Zone_Code,TSPL_CUSTOMER_MASTER.Route_No,Number,TSPL_ROUTE_MASTER.Route_Desc,tspl_customer_master.price_CodeNon,tspl_transport_master.Transporter_Name from TSPL_CUSTOMER_MASTER left outer join " &
                   "TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No left outer join TSPL_VEHICLE_MASTER on " &
                   "TSPL_ROUTE_MASTER.vehicle_code=TSPL_VEHICLE_MASTER.Vehicle_Id left outer join tspl_transport_master on tspl_transport_master.Transport_Id=TSPL_VEHICLE_MASTER.Transport_Id "
                    If chkIndividualCustomer.Checked Then
                        qry += " where TSPL_CUSTOMER_MASTER.Cust_Code ='" & txtCustomerNo.Value & "'"
                    Else
                        qry += " where TSPL_ROUTE_MASTER.Route_No ='" & txtRouteNo.Value & "'"
                    End If
                End If
            End If
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0) Then
                txtVehicleNo.Value = clsCommon.myCstr(dt1.Rows(0)("vehicle_code"))
                lblVehicleNo.Text = clsCommon.myCstr(dt1.Rows(0)("Number"))
                txtRouteNo.Value = clsCommon.myCstr(dt1.Rows(0)("Route_No"))
                lblRouteDesc.Text = clsCommon.myCstr(dt1.Rows(0)("Route_Desc"))
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal AndAlso Not chkIndividualCustomer.Checked Then
                    lblTransporterName.Text = clsCommon.myCstr(dt1.Rows(0)("Customer_Name"))
                Else
                    lblTransporterName.Text = clsCommon.myCstr(dt1.Rows(0)("Transporter_Name"))
                End If
                TxtCity.Value = clsCommon.myCstr(dt1.Rows(0)("City_Code"))
                lblCityName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select City_name from tspl_city_master where city_code=  '" & TxtCity.Value & "' "))
                If chkIndividualCustomer.Checked Then
                    If clsCommon.myLen(clsCommon.myCstr(txtRouteNo.Value)) <= 0 Then
                        Throw New Exception("Please Map Route with Customer " & lblCustomerName.Text & "")
                    End If
                End If
                If clsCommon.myLen(clsCommon.myCstr(TxtCity.Value)) <= 0 Then
                    Throw New Exception("Please Map city with Route " & lblRouteDesc.Text & "")
                End If
                If clsCommon.myLen(clsCommon.myCstr(txtVehicleNo.Value)) <= 0 Then
                    Throw New Exception("Please Map Vehicle with Route " & lblRouteDesc.Text & "")
                End If
                If EnableLocation Then
                    txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Route_Master where Route_No='" + txtRouteNo.Value + "' "))
                    txtLocation.Enabled = False
                Else
                    txtLocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' "))
                End If
                If clsCommon.myLen(txtLocation.Value) > 0 Then
                    lblLocation.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Desc from TSPL_LOCATION_MASTER where Location_Code='" + txtLocation.Value + "'"))
                End If
                If clsCommon.myLen(clsCommon.myCstr(txtRouteNo.Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(TxtCity.Value)) > 0 Then
                    setCustomerDetail(TxtCity.Value, txtRouteNo.Value, False)
                End If
            End If
            RefreshFormName()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    '    Private Sub gv1_ViewRowFormatting(sender As Object, e As RowFormattingEventArgs) Handles gv1.ViewRowFormatting
    '        If TypeOf e.RowElement Is GridTableHeaderRowElement Then
    '            If gv1.Rows.Count > 0 Then
    '                If clsCommon.myLen(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colCustCode).Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colShiftName).Value)) > 0 Then
    '                    Dim strqry As String = "Select count(ShiftType) from (select ShiftType,convert(date,Document_Date ,103) as Document_Date from (select tspl_demand_booking_master.Document_Date , tspl_demand_booking_detail.ShiftType  from tspl_demand_booking_detail
    'left outer join tspl_demand_booking_master on tspl_demand_booking_detail.document_no=tspl_demand_booking_master.document_no
    'where convert(date,tspl_demand_booking_master.document_date,103) in ('" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)) & "','" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-2)) & "','" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-3)) & "')
    'and tspl_demand_booking_detail.Cust_Code='" & clsCommon.myCstr(e.RowElement.RowInfo.Cells(colCustCode).Value) & "' and TSPL_DEMAND_BOOKING_DETAIL .ShiftType ='" & clsCommon.myCstr(e.RowElement.RowInfo.Cells(colShiftName).Value) & "')Final 
    'group by ShiftType ,convert(date,Document_Date ,103))FinalQry"
    '                    Dim bookingcount As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strqry))
    '                    If bookingcount = 3 Then
    '                        e.RowElement.DrawFill = True
    '                        e.RowElement.GradientStyle = GradientStyles.Solid
    '                        e.RowElement.BackColor = Color.LightGreen
    '                    Else
    '                        e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
    '                        e.RowElement.ResetValue(LightVisualElement.GradientStyleProperty, ValueResetFlags.Local)
    '                        'e.RowElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local)
    '                        e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
    '                    End If
    '                End If
    '            End If
    '        Else
    '            e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
    '            e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
    '            e.RowElement.ResetValue(LightVisualElement.NumberOfColorsProperty, ValueResetFlags.Local)
    '            e.RowElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local)
    '        End If
    '    End Sub
    '    Private Sub gv1_RowFormatting(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.RowFormattingEventArgs) Handles gv1.RowFormatting
    '        'Dim font_Renamed As Font
    '        'font_Renamed = New Font(e.RowElement.Font, FontStyle.Bold)
    '        If gv1.Rows.Count > 0 Then
    '            If clsCommon.myLen(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colCustCode).Value)) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colShiftName).Value)) > 0 Then
    '                '                Dim strqry As String = "Select count(ShiftType) from (select ShiftType,convert(date,Document_Date ,103) as Document_Date from (select tspl_demand_booking_master.Document_Date , tspl_demand_booking_detail.ShiftType  from tspl_demand_booking_detail
    '                'left outer join tspl_demand_booking_master on tspl_demand_booking_detail.document_no=tspl_demand_booking_master.document_no
    '                'where convert(date,tspl_demand_booking_master.document_date,103) in ('" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)) & "','" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-2)) & "','" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-3)) & "')
    '                'and tspl_demand_booking_detail.Cust_Code='" & clsCommon.myCstr(e.RowElement.RowInfo.Cells(colCustCode).Value) & "' and TSPL_DEMAND_BOOKING_DETAIL .ShiftType ='" & clsCommon.myCstr(e.RowElement.RowInfo.Cells(colShiftName).Value) & "')Final 
    '                'group by ShiftType ,convert(date,Document_Date ,103))FinalQry"
    '                '                Dim bookingcount As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strqry))
    '                If clsCommon.myLen(e.RowElement.RowInfo.Cells(colBookingCreatedFor3Days)) > 0 Then
    '                    If clsCommon.CompairString(clsCommon.myCstr(e.RowElement.RowInfo.Cells(colBookingCreatedFor3Days).Value), "3") = CompairStringResult.Equal Then
    '                        'End If
    '                        'If bookingcount = 3 Then
    '                        e.RowElement.DrawFill = True
    '                        e.RowElement.GradientStyle = GradientStyles.Solid
    '                        e.RowElement.BackColor = Color.LightGreen
    '                        'e.RowElement.Font = font_Renamed
    '                    Else
    '                        e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
    '                        e.RowElement.ResetValue(LightVisualElement.GradientStyleProperty, ValueResetFlags.Local)
    '                        'e.RowElement.ResetValue(LightVisualElement.ForeColorProperty, ValueResetFlags.Local)
    '                        e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
    '                        ' e.RowElement.ResetValue(LightVisualElement.FontProperty, font_Renamed)
    '                        'e.RowElement.BackColor = Color.Black
    '                    End If
    '                End If
    '            End If
    '        End If
    '        If e.RowElement.RowInfo.IsCurrent Then
    '            e.RowElement.DrawFill = True
    '            e.RowElement.BackColor = Color.LightGreen
    '        Else
    '            e.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
    '            e.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
    '        End If
    '    End Sub
    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        '' Setting UpdateDemandBeforePost for Save Demand before post 
        If UpdateDemandBeforePost Then
            SaveData(True)
        End If
        PostData()

    End Sub
    Private Function LoadOutstanding(ByVal strCustCode As String, ByVal trans As SqlTransaction) As Double
        Dim qry As String = "Select  case when (( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt)) )>=0 then -abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) else abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) end  as BalAmt From ( " &
                    "Select MAX(TSPL_CUSTOMER_MASTER.Cust_Group_Code) as Cust_Group_Code, ACode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName, '' as CurrencyCode,  " &
                    "null as ConvRate, SUM(DrAmt* Final.ConvRate)-SUM(CrAmt) as OpngBal, 0 as DrAmt, 0 as CrAmt, 0 as [Sales], 0 as CollectionRefund, 0 as DrNote,  " &
                    "0 as CrNote, MAX(tspl_customer_master.Cust_Category_Code) as Cust_Category_Code,MAX(CUST_CATEGORY_DESC) as Cust_Category_Desc,  " &
                    "MAX(tspl_customer_master.Cust_Type_Code) As Cust_Type_Code,MAX(Cust_Type_Desc) As Cust_Type_Desc from   " &
                    "(" & clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, "ConvRate", "'" + strCustCode + "'", True, clsCommon.GetPrintDate(txtDate.Value.AddDays(1), "dd/MMM/yyyy"), "", False, False, True, trans, False, txtDocNo.Value) & "   " &
                    " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code LEFT OUTER JOIN TSPL_CUSTOMER_GROUP_MASTER ON TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code " &
                    "Left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No =Final.DocNo  LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=Final.Bank_Code " &
                    "where  CONVERT(DATE,final.DocDate,103) <= '" & clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") & "' AND LEN(ACode)>0 and ACode in ('" & strCustCode & "')   AND TSPL_CUSTOMER_MASTER.Status='N' GROUP BY ACode " &
                    ") XXX GROUP BY ACode ORDER BY ACode"
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
    End Function
    Sub PostData()
        Dim msg As String = Nothing
        Dim qry As String = Nothing
        Dim dt As DataTable = Nothing
        Dim desc As String = Nothing
        Dim IsPost As Boolean = False
        Try
            Dim StrQry As String = ""
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                StrQry = "select sum(ItemNetAmount) as ItemNetAmount, TSPL_DEMAND_BOOKING_DETAIL.Cust_Code from TSPL_DEMAND_BOOKING_DETAIL
left join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
where Document_No='" & txtDocNo.Value & "' and TSPL_CUSTOMER_MASTER.Cash_Customer='Y'
group by TSPL_DEMAND_BOOKING_DETAIL.Cust_Code"
                Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(StrQry)
                If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                    For Each dr As DataRow In dt2.Rows
                        If CheckCustomeroutStandingAmt Then
                            Dim custOutStanding As Double = clsCommon.myCdbl(LoadOutstanding(clsCommon.myCstr(dr("Cust_Code")), Nothing))
                            custOutStanding = Math.Abs(custOutStanding)
                            Dim TotalDocAmt As Double = clsCommon.myCdbl(dr("ItemNetAmount"))
                            If TotalDocAmt > custOutStanding Then
                                Throw New Exception("Insufficient Balance. [ Customer Code : " + clsCommon.myCstr(dr("Cust_Code")) + ", Current Outstanding Bal: " + clsCommon.myCstr(custOutStanding) + " ]")
                            End If
                        End If
                    Next
                End If

                StrQry = "select IsPosting,IsUpdating,Posted,Curr_User from TSPL_DEMAND_BOOKING_MASTER where Document_No='" & txtDocNo.Value & "'"
                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(StrQry)
                If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 AndAlso (clsCommon.myCdbl(dt1.Rows(0)("IsPosting")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("IsUpdating")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("Posted")) = 1) Then
                    'If clsCommon.myCdbl(dt1.Rows(0)("IsPosting")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("IsUpdating")) = 1 OrElse clsCommon.myCdbl(dt1.Rows(0)("Posted")) = 1 Then
                    Throw New Exception("Document in used by [" & clsCommon.myCstr(dt1.Rows(0)("Curr_User")) & "]")
                    'End If
                End If
                StrQry = "Update TSPL_DEMAND_BOOKING_MASTER set IsPosting=1,Curr_User='" & objCommonVar.CurrentUser & "' where Document_No='" & txtDocNo.Value & "' "
                clsDBFuncationality.ExecuteNonQuery(StrQry)
            Else
                Throw New Exception("Please Select Document")
            End If

                'Dim custCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cust_Code from TSPL_CUSTOMER_MASTER where route_no='" + txtRouteNo.Value + "' and IsDistributor='Y'"))
                StrQry = "select  top 1 x.Cust_Code 
from(
select TSPL_DISTRIBUTOR_ROUTE.Code as Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks,max(TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code) as cust_code
from TSPL_DISTRIBUTOR_ROUTE
left join TSPL_DISTRIBUTOR_ROUTE_CUSTOMER on TSPL_DISTRIBUTOR_ROUTE.Code=TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Code
where  TSPL_DISTRIBUTOR_ROUTE.Status=1 and IS_Transpoter=0 and TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No='" & txtRouteNo.Value & "'
 Group by TSPL_DISTRIBUTOR_ROUTE.Code,TSPL_DISTRIBUTOR_ROUTE.Start_Date,TSPL_DISTRIBUTOR_ROUTE.Remarks
) X"
            Dim custCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(StrQry))
            If (myMessages.postConfirm()) Then
                If checkCreditLimit Then
                    If clsCommon.CompairString(clsDBFuncationality.getSingleValue("select Credit_Customer from TSPL_CUSTOMER_MASTER where Cust_Code='" + custCode + "'"), "N") = CompairStringResult.Equal Then
                        Dim OsBal As Decimal = GetOutStandingBal(custCode, txtDate.Value)
                        If OsBal > clsCommon.myCdbl(txtDocAmt.Text) Then
                            IsPost = True
                        Else
                            IsPost = False
                            If clsCommon.myLen(custCode) > 0 Then
                                Dim count As Integer = clsDBFuncationality.getSingleValue("select count(*) from TSPL_TRANSACTION_APPROVAL where Program_Code='" + clsUserMgtCode.frmDemandBooking + "' and Document_No='" + txtDocNo.Value + "' and Cust_Code='" + custCode + "'")
                                If count = 0 Then
                                    qry = "insert into TSPL_TRANSACTION_APPROVAL(Screen_Name,Program_Code,Document_No,Doc_Date,approval_type,Approve,Created_By,Created_Date,Modified_By,Modified_Date,Comp_Code,Cust_Code,Loc_Code) " &
                                      "values ('Demand Booking','" & clsUserMgtCode.frmDemandBooking & "','" & txtDocNo.Value & "', " &
                                      "'" & clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy") & "','Credit Limit',0,'" & objCommonVar.CurrentUserCode & "', " &
                                      "'" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") & "', " &
                                      "'" & objCommonVar.CurrentUserCode & "','" & clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MMM/yyyy hh:mm tt") & "', " &
                                      "'" & objCommonVar.CurrentCompanyCode & "','" & custCode & "','" & txtLocation.Value & "')"
                                    clsDBFuncationality.ExecuteNonQuery(qry)
                                    common.clsCommon.MyMessageBoxShow(Me, "Insufficient O/S Balance send for Approval", Me.Text)
                                Else
                                    common.clsCommon.MyMessageBoxShow(Me, "Demand already send for Approval", Me.Text)
                                End If
                            Else
                                If clsCommon.MyMessageBoxShow(Me, "Distributor Not Found! Do You Want to Continue?", Me.Text, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                                    IsPost = True
                                Else
                                    IsPost = False
                                End If
                            End If
                        End If
                    Else
                        IsPost = True
                    End If
                Else
                    IsPost = True
                End If
                If IsPost Then
                    If rbtnMorningEveningBoth.IsChecked Then
                        Throw New Exception("Please select morning/evening")
                    End If
                    If (clsDemandBookingSale.PostData(MyBase.Form_ID, txtDocNo.Value, IIf(rbtnMorning.IsChecked, 1, 2))) Then
                        msg = "Successfully posted"
                        common.clsCommon.MyMessageBoxShow(Me, msg, Me.Text)
                        'If Not chkIndividualCustomer.Checked Then
                        '    If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, Nothing)) = 1 Or clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, Nothing)) = 1 Then
                        '        SaveData(1, True)
                        '    End If
                        'End If
                        LoadData(txtDocNo.Value, NavigatorType.Current)
                    End If
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally

            qry = "Update TSPL_DEMAND_BOOKING_MASTER set IsPosting=0,Curr_User= null where Document_No='" & txtDocNo.Value & "'"
            clsDBFuncationality.ExecuteNonQuery(qry)
            msg = Nothing
            qry = Nothing
            dt = Nothing
            desc = Nothing

        End Try
    End Sub
    Private Sub rbtnEvening_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnEvening.ToggleStateChanged
        Try
            If Not isInsideLoadData Then
                HideUnhideRowsAndColumnsOFGrid()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtnMorningEveningBoth_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnMorningEveningBoth.ToggleStateChanged
        Try
            If Not isInsideLoadData Then
                HideUnhideRowsAndColumnsOFGrid()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rbtn_Fresh_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtn_Fresh.ToggleStateChanged
        Try
            If Not isInsideLoadData AndAlso rbtn_Fresh.IsChecked Then
                'If rbtn_Fresh.IsChecked Then
                HideUnhideRowsAndColumnsOFGrid()
                'End If
            End If
            'btnPrint.Enabled = True
            txtPCount.Text = ""
            txtPAmt.Text = ""
            checkPrintSetting()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub checkPrintSetting()
        If PrintOnlyPostedDocument Then
            If clsCommon.CompairString(clsCommon.myCstr(UsLock1.Status), "Approved") = CompairStringResult.Equal Then
                btnPrint.Enabled = True
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                    SplitButtonTruckSheet.Enabled = True
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                        If rdbnFreshAmbientBoth.IsChecked Then
                            SplitButtonTruckSheet.Enabled = True
                        Else
                            SplitButtonTruckSheet.Enabled = False
                        End If
                    End If

                Else
                    SplitButtonTruckSheet.Enabled = False
                End If
            Else
                btnPrint.Enabled = False
                SplitButtonTruckSheet.Enabled = False
            End If
        Else
            btnPrint.Enabled = True
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                SplitButtonTruckSheet.Enabled = True
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                    If rdbnFreshAmbientBoth.IsChecked Then
                        SplitButtonTruckSheet.Enabled = True
                    Else
                        SplitButtonTruckSheet.Enabled = False
                    End If
                End If
            Else
                SplitButtonTruckSheet.Enabled = False
            End If
        End If
    End Sub

    Private Sub rdbnBoth_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rdbnFreshAmbientBoth.ToggleStateChanged
        isFreshAmbientBoth()
    End Sub

    Sub isFreshAmbientBoth()
        Try
            If Not isInsideLoadData Then
                'If rdbnFreshAmbientBoth.IsChecked Then
                HideUnhideRowsAndColumnsOFGrid()
                'End If
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                    checkPrintSetting()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                        If rdbnFreshAmbientBoth.IsChecked Then
                            'SplitButtonTruckSheet.Enabled = True
                            checkPrintSetting()
                        Else
                            SplitButtonTruckSheet.Enabled = False
                        End If
                    End If

                End If
                'HideUnhideRowsAndColumnsOFGrid()
            End If
            'btnPrint.Enabled = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub Btn_TruckSheet_Click(sender As Object, e As EventArgs) Handles btn_TruckSheet.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select document", Me.Text)
                txtDocNo.Focus()
                Exit Sub
            End If
            If rbtnMorningEveningBoth.IsChecked Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Morning/Evening Shift", Me.Text)
                rbtnMorning.Focus()
                Exit Sub
            End If
            'clsDBFuncationality.ExecuteNonQuery("update TSPL_DEMAND_BOOKING_DETAIL set IsTruckSheetGenerated='Y' where document_no='" & txtDocNo.Value & "' and ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'")
            '    Dim Qry As String = "select TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1 as Comp_Add1,TSPL_COMPANY_MASTER.Add2 as Comp_Add2,TSPL_COMPANY_MASTER.Add3 as Comp_Add3,TSPL_COMPANY_MASTER.Pincode as Comp_Pin
            ',TSPL_LOCATION_MASTER.Add1,TSPL_LOCATION_MASTER.Add2,TSPL_LOCATION_MASTER.ADD3,TSPL_LOCATION_MASTER.Pin_Code,TSPL_LOCATION_MASTER.Location_Desc
            ',Main_Final.shiftType,Main_Final.City_Name,Main_Final.Demand_No,Main_Final.Demand_Date,Main_Final.Route_No,Main_Final.Route_Desc
            ',Main_Final.cust_code,Main_Final.Customer_Name_Hindi,Main_Final.Crate_Qty,Main_Final.Milk_Amt,Main_Final.Product_Amt,Main_Final.Total_Amt
            'from (select 
            'max(TSPL_DEMAND_BOOKING_MASTER.shiftType) as shiftType,
            'max(TSPL_city_MASTER.City_Name) as City_Name,
            'max(TSPL_DEMAND_BOOKING_MASTER.Comp_Code) as Comp_Code,
            'max(TSPL_DEMAND_BOOKING_MASTER.location_code) as location_code
            ',TSPL_DEMAND_BOOKING_MASTER.Document_No as Demand_No,max(convert(varchar(15),TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)) as Demand_Date ,isnull(TSPL_DEMAND_BOOKING_MASTER.Route_No,'') as Route_No ,max(isnull(TSPL_ROUTE_MASTER.Route_Desc,'')) as Route_Desc 
            ' ,TSPL_DEMAND_BOOKING_DETAIL.cust_code,max(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi) as Customer_Name_Hindi
            ',sum(TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise) AS Crate_Qty
            ',sum(case when TSPL_ITEM_MASTER.Is_Milk_Pouch=1 then TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount else 0 end) as Milk_Amt
            ',sum(case when TSPL_ITEM_MASTER.Is_Milk_Pouch=0 then TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount else 0 end) as Product_Amt
            ',sum(TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount) AS Total_Amt 
            'from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL
            'on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
            ' left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code 
            'left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
            ' left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
            'left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DEMAND_BOOKING_MASTER.Route_No
            'left outer join TSPL_city_MASTER on TSPL_city_MASTER.city_code = TSPL_DEMAND_BOOKING_MASTER.city_code
            'WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" + txtDocNo.Value + "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'
            ' group by TSPL_DEMAND_BOOKING_DETAIL.cust_code,TSPL_DEMAND_BOOKING_MASTER.Document_No,TSPL_DEMAND_BOOKING_MASTER.Route_No
            ') as Main_Final
            'LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=Main_Final.Comp_Code
            ' LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code =Main_Final.location_code"
            '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            '    If dt.Rows.Count > 0 Then
            '        Dim frmCRV As New frmCrystalReportViewer()
            '        frmCRV.funsubreportWithdt(CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptTruckSheetPrint", "Truck Sheet", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            '        frmCRV = Nothing
            '    End If
            'Dim Qry As String = "select TSPL_DEMAND_BOOKING_MASTER.TripNo,
            '    '" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "' shiftType,
            '    TSPL_city_MASTER.City_Name,
            '    TSPL_DEMAND_BOOKING_MASTER.Comp_Code,
            '    TSPL_DEMAND_BOOKING_MASTER.location_code
            '    ,TSPL_DEMAND_BOOKING_MASTER.Document_No as Demand_No,convert(varchar(15),TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) as Demand_Date 
            '    ,isnull(TSPL_DEMAND_BOOKING_MASTER.Route_No,'') as Route_No ,isnull(TSPL_ROUTE_MASTER.Route_Desc,'') as Route_Desc 
            '    ,TSPL_VENDOR_MASTER.vendor_name as Distributor 
            '    ,TSPL_DEMAND_BOOKING_DETAIL.cust_code,TSPL_CUSTOMER_MASTER.Customer_Name_Hindi
            '    ,TSPL_DEMAND_BOOKING_DETAIL.Qty
            '    ,TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise
            '    ,TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise
            '    ,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount
            '    ,TSPL_DEMAND_BOOKING_DETAIL.unit_code
            '    ,TSPL_ITEM_MASTER.Alies_Name_Hindi as Item_Alies_Name_Hindi 
            '    ,case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Crate' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end as CrateQty
            '    ,case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Pouch' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end as PouchQty
            '    ,N'क्रेट' as Crate,N'क्रेट    थैली' as Pouch
            '    from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL
            '    on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
            '     left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code 
            '    left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
            '     left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
            '    left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DEMAND_BOOKING_MASTER.Route_No
            '    left outer join TSPL_city_MASTER on TSPL_city_MASTER.city_code = TSPL_DEMAND_BOOKING_MASTER.city_code
            '    left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
            '    LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code =TSPL_DEMAND_BOOKING_MASTER.location_code
            '    left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VEHICLE_MASTER.Transport_Id
            '    WHERE TSPL_ITEM_MASTER.Is_Milk_Pouch=1
            '    and TSPL_DEMAND_BOOKING_MASTER.Document_No='" + txtDocNo.Value + "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'"
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            'If dt.Rows.Count > 0 Then
            '    Dim frmCRV As New frmCrystalReportViewer()
            '    frmCRV.funsubreportWithdt(CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptTruckSheetPrintNew", "Truck Sheet", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            '    frmCRV = Nothing
            'End If
            'TruckSheetExcel()
            'TruckPDF()
            btn_TruckSheet.Enabled = False
            btn_TSCancel.Enabled = True
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub TruckSheetExcel(ByVal isExcelPDF As Boolean, ByVal TripNo As String)
        Dim BaseQry As String = Nothing
        Dim doc As New XpertERPEngine.clsMyPrintDocument()
        'GVTruckSheet = New UserControls.MyRadGridView()
        'Me.Controls.Add(GVTruckSheet)

        GVTruckSheet.DataSource = Nothing
        GVTruckSheet.Rows.Clear()
        GVTruckSheet.Columns.Clear()
        GVTruckSheet.GroupDescriptors.Clear()
        GVTruckSheet.Refresh()
        GVTruckSheet.MasterView.Refresh()
        GVTruckSheet.GroupDescriptors.Clear()
        GVTruckSheet.EnableFiltering = True
        GVTruckSheet.MasterTemplate.SummaryRowsBottom.Clear()
        GVTruckSheet.MasterTemplate.AutoExpandGroups = True
        GVTruckSheet.BestFitColumns()
        GVTruckSheet.ReadOnly = True
        GVTruckSheet.AllowAddNewRow = False
        GVTruckSheet.AllowDragToGroup = False
        GVTruckSheet.AllowDeleteRow = False

        Try
            Dim ItemInUse As String = " TSPL_DEMAND_BOOKING_MASTER Left outer join TSPL_DEMAND_BOOKING_DETAIL
                On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
                Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
                Left Outer Join TSPL_UNIT_MASTER On TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
                where TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'
                and TSPL_ITEM_MASTER.Is_Milk_Pouch=1 "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                ItemInUse &= "" & IIf(clsCommon.CompairString(TripNo, "ALL") = CompairStringResult.Equal, "", "and TSPL_DEMAND_BOOKING_DETAIL.Trip_No='" & TripNo & "'") & " "  ' order by sku_seq"
            End If
            Dim ItemInUseProduct As String = " TSPL_DEMAND_BOOKING_MASTER Left outer join TSPL_DEMAND_BOOKING_DETAIL
                On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
                Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
                Left Outer Join TSPL_UNIT_MASTER On TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
                where TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'
                and TSPL_ITEM_MASTER.Is_Milk_Pouch=0  "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                ItemInUseProduct &= " " & IIf(clsCommon.CompairString(TripNo, "ALL") = CompairStringResult.Equal, "", "and TSPL_DEMAND_BOOKING_DETAIL.Trip_No='" & TripNo & "'") & " "
            End If
            ItemInUseProduct &= "  order by sku_seq "
            Dim dtDataExist As DataTable = clsDBFuncationality.GetDataTable(" Select  Item_Code,Alies_Name,Max(Sku_Seq)Sku_Seq,Max(Case When Unit_Desc='Crate' Then Case When LEN(Unit_Desc_Hindi)>0 Then Unit_Desc_Hindi Else Unit_Desc End End) As SizeC,Max(Case When Unit_Desc='Pouch' Then Case When LEN(Unit_Desc_Hindi)>0 Then Unit_Desc_Hindi Else Unit_Desc End Else 'Pouch' End) As SizeP,Max(Case When Unit_Desc='LTR' Then Case When LEN(Unit_Desc_Hindi)>0 Then Unit_Desc_Hindi Else Unit_Desc End Else 'LTR' End) As SizeL  from (select distinct isnull(TSPL_ITEM_MASTER.Alies_Name_Hindi,'')  Alies_Name,TSPL_ITEM_MASTER.Item_Code,sku_seq,SUBSTRING(TSPL_ITEM_MASTER.Alies_Name, LEN(TSPL_ITEM_MASTER.Alies_Name) -  CHARINDEX(' ', REVERSE(TSPL_ITEM_MASTER.Alies_Name))+2,LEN(TSPL_ITEM_MASTER.Alies_Name))+' '+ Case When LEN(TSPL_UNIT_MASTER.Unit_Desc_Hindi)>0 Then TSPL_UNIT_MASTER.Unit_Desc_Hindi Else TSPL_UNIT_MASTER.Unit_Desc End AS Size1,Case When LEN(TSPL_UNIT_MASTER.Unit_Desc_Hindi)>0 Then TSPL_UNIT_MASTER.Unit_Desc_Hindi Else TSPL_UNIT_MASTER.Unit_Desc End AS Size,TSPL_UNIT_MASTER.Unit_Desc_Hindi,TSPL_UNIT_MASTER.Unit_Desc from " & ItemInUse & ")xyz Group By Item_Code,Alies_Name order by sku_seq")
            Dim dtDataExistProduct As DataTable = clsDBFuncationality.GetDataTable("select distinct isnull(TSPL_ITEM_MASTER.Alies_Name_Hindi,'')  Alies_Name,TSPL_ITEM_MASTER.Item_Code,sku_seq,SUBSTRING(TSPL_ITEM_MASTER.Alies_Name, LEN(TSPL_ITEM_MASTER.Alies_Name) -  CHARINDEX(' ', REVERSE(TSPL_ITEM_MASTER.Alies_Name))+2,LEN(TSPL_ITEM_MASTER.Alies_Name))+' '+ Case When LEN(TSPL_UNIT_MASTER.Unit_Desc_Hindi)>0 Then TSPL_UNIT_MASTER.Unit_Desc_Hindi Else TSPL_UNIT_MASTER.Unit_Desc End AS Size1,Case When LEN(TSPL_UNIT_MASTER.Unit_Desc_Hindi)>0 Then TSPL_UNIT_MASTER.Unit_Desc_Hindi Else TSPL_UNIT_MASTER.Unit_Desc End AS Size   from " & ItemInUseProduct)
            If (dtDataExist Is Nothing OrElse dtDataExist.Rows.Count = 0) AndAlso (dtDataExistProduct Is Nothing OrElse dtDataExistProduct.Rows.Count = 0) Then
                'clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
                Throw New InvalidOperationException("No Data Found")
#Disable Warning S3385 ' "Exit" statements should not be used
                Exit Sub
#Enable Warning S3385 ' "Exit" statements should not be used
            Else
                Dim sbAlias As New StringBuilder()
                If dtDataExist IsNot Nothing AndAlso dtDataExist.Rows.Count > 0 Then
                    For Each strAlias As DataRow In dtDataExist.Rows
                        If clsCommon.myLen(clsCommon.myCstr(strAlias("Alies_Name"))) <= 0 Then
                            If clsCommon.myLen(clsCommon.myCstr(sbAlias)) > 0 Then
                                sbAlias.Append(Environment.NewLine)
                            End If
                            sbAlias.Append("Alias Name can't be blank for Item [" & clsCommon.myCstr(strAlias("Item_Code")) & "]")
                        End If
                    Next
                End If

                If dtDataExistProduct IsNot Nothing AndAlso dtDataExistProduct.Rows.Count > 0 Then
                    For Each strAlias As DataRow In dtDataExistProduct.Rows
                        If clsCommon.myLen(clsCommon.myCstr(strAlias("Alies_Name"))) <= 0 Then
                            If clsCommon.myLen(clsCommon.myCstr(sbAlias)) > 0 Then
                                sbAlias.Append(Environment.NewLine)
                            End If
                            sbAlias.Append("Alias Name can't be blank for Item [" & clsCommon.myCstr(strAlias("Item_Code")) & "]")
                        End If
                    Next
                End If
                If clsCommon.myLen(clsCommon.myCstr(sbAlias)) > 0 Then
                    Throw New Exception(clsCommon.myCstr(sbAlias))
                End If
                sbAlias = Nothing
            End If
            Dim sbstrItemC As New StringBuilder()
            Dim sbstrItemP As New StringBuilder()
            Dim sbstrItemL As New StringBuilder()
            Dim sbstrItemA As New StringBuilder()
            Dim sbstrProdQ As New StringBuilder()
            Dim sbstrItemSUM As New StringBuilder()





            If rdbnFreshAmbientBoth.IsChecked Then
#Disable Warning S2259 ' Null pointers should not be dereferenced
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
#Enable Warning S2259 ' Null pointers should not be dereferenced
                    If i = 0 Then
                        sbstrItemC.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemC.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                Next
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    'If clsCommon.CompairString(strItemSUM, "") = CompairStringResult.Equal Then
                    If sbstrItemSUM.Length = 0 Then
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                    'If clsCommon.CompairString(strProdQ, "") = CompairStringResult.Equal Then
                    If sbstrProdQ.Length = 0 Then
                        sbstrProdQ.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    Else
                        sbstrProdQ.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    End If
                Next
            ElseIf rbtn_Fresh.IsChecked Then
#Disable Warning S2259 ' Null pointers should not be dereferenced
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
#Enable Warning S2259 ' Null pointers should not be dereferenced
                    If i = 0 Then
                        sbstrItemC.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemC.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                Next
            ElseIf rbtn_Ambient.IsChecked Then
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    'If clsCommon.CompairString(strItemSUM, "") = CompairStringResult.Equal Then
                    If sbstrItemSUM.Length = 0 Then
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                    'If clsCommon.CompairString(strProdQ, "") = CompairStringResult.Equal Then
                    If sbstrProdQ.Length = 0 Then
                        sbstrProdQ.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    Else
                        sbstrProdQ.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    End If
                Next
            End If
            Dim strItemC As String = clsCommon.myCstr(sbstrItemC)
            Dim strItemP As String = sbstrItemP.ToString()
            Dim strItemL As String = sbstrItemL.ToString()
            Dim strItemA As String = sbstrItemA.ToString()
            Dim strProdQ As String = sbstrProdQ.ToString()
            Dim strItemSUM As String = sbstrItemSUM.ToString()
            Dim Qry As String = "select Row_Number() Over (Order By (Select 1)) As SNo,Cust_Code As Code,max(customer_Name) as Agents, " & strItemSUM & ""
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") <> CompairStringResult.Equal AndAlso Not isExportTruckSheet Then
                Qry += " ,sum(isnull(TotalLtr_CustWise,0)) as [Milk In Ltr],sum(isnull(TotalCrates_ItemWise,0)) as [Crates],sum(isnull(MAmt,0)) as [Milk Amount],sum(isnull(PQty,0)) as [Product Quantity],sum(isnull(PAmt,0)) as [Product Amount] "
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") <> CompairStringResult.Equal AndAlso Not isExportTruckSheet Then
                Qry += ",(sum(isnull(MAmt,0))+sum(isnull(PAmt,0))) as [Total Amount] "
            Else
                If rbtn_Fresh.IsChecked Then
                    Qry += ",(sum(isnull(MAmt,0))) as [Total Amount] "
                ElseIf rbtn_Ambient.IsChecked Then
                    Qry += ",(sum(isnull(PAmt,0))) as [Total Amount] "
                Else
                    Qry += ",(sum(isnull(MAmt,0))+sum(isnull(PAmt,0))) as [Total Amount] "
                End If
            End If

            Qry += " from   
        (Select  max(Display_Seq) as Display_Seq,Cust_Code,max(Customer_Name) as Customer_Name,Item_Code,max(Alies_Name)+'#C' as Alies_Name#C
        ,max(Alies_Name)+'#P' AS Alies_Name#P
		,max(Alies_Name)+'#L' AS Alies_Name#L,max(Alies_Name)+'#A' AS Alies_Name#A
        ,max(Alies_Name)+'#ProdQ' AS Alies_Name#ProdQ
		,max(Unit_Desc) as Unit_Desc,max(Code) as Code
        ,sum(Qty_Crate) as Qty_Crate, sum (Qty_Pouch) as Qty_Pouch
        ,sum(TotalLtr_ItemWise) as TotalLtr_CustWise 
		,sum(TotalLtr_ItemWise) as TotalLtr_ItemWise
		,sum(ItemNetAmount) as TotalAmt_ItemWise
        ,SUM(TotalCrates_ItemWise) AS TotalCrates_ItemWise
        ,sum(ProdQ) as ProdQ
        ,SUM(MAmt) AS MAmt
        ,SUM(PQty) AS PQty
        ,SUM(PAmt) AS PAmt from ("

            BaseQry += " Select '1' as  Code,TSPL_CUSTOMER_MASTER.Display_Seq,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,coalesce(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name 
	, TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name_Hindi as Alies_Name
	,TSPL_ITEM_MASTER.Is_FreshItem,TSPL_ITEM_MASTER.Is_Ambient,"
            BaseQry += " TSPL_UNIT_MASTER.Unit_Code,TSPL_UNIT_MASTER.Unit_Desc, "
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                BaseQry += " Cast(TSPL_DEMAND_BOOKING_DETAIL.Qty As decimal(18,2))Qty "
                BaseQry += ",Cast((CASE WHEN TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Crate' THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As decimal(18,2)) as Qty_Crate
	,Cast((CASE WHEN TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Pouch' THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As decimal(18,2)) as Qty_Pouch 	
    ,Cast((CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As decimal(18,2)) as ProdQ	
    ,Cast((CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As decimal(18,2)) as PQty "
            Else
                BaseQry += " Cast(TSPL_DEMAND_BOOKING_DETAIL.Qty As Int)Qty "
                BaseQry += ",Cast((CASE WHEN TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Crate' THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As Int) as Qty_Crate
	,Cast((CASE WHEN TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Pouch' THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As Int) as Qty_Pouch 	
    ,Cast((CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As Int) as ProdQ	
    ,Cast((CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) As Int) as PQty "
            End If
            BaseQry += ",(CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=1 THEN TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount ELSE 0 END) as MAmt
    ,TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount
    ,TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise
	,(CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount ELSE 0 END) as PAmt,TSPL_ITEM_MASTER.Summary_Seq_No
	 from TSPL_DEMAND_BOOKING_MASTER Left outer join TSPL_DEMAND_BOOKING_DETAIL
     On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
     Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
	left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
	left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
    WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'" & IIf(clsCommon.CompairString(TripNo, "ALL") = CompairStringResult.Equal, "", "and TSPL_DEMAND_BOOKING_DETAIL.Trip_No='" & TripNo & "'") & ""

            If (clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal OrElse isExportTruckSheet) AndAlso isIndent Then
                BaseQry += " Union All "
                BaseQry += " Select '1' as  Code,TSPL_CUSTOMER_MASTER.Display_Seq,TSPL_CUSTOMER_MASTER.Cust_Code,coalesce(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name 
	, TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name_Hindi as Alies_Name
	,TSPL_ITEM_MASTER.Is_FreshItem,TSPL_ITEM_MASTER.Is_Ambient
	,TSPL_UNIT_MASTER.Unit_Code,TSPL_UNIT_MASTER.Unit_Desc,0 As Qty
	,0 as Qty_Crate
	,0 as Qty_Pouch	
    ,0 as ProdQ
    ,0 as PQty   
	,0 as MAmt
    ,0 TotalLtr_ItemWise, 0 As ItemNetAmount,0 As TotalCrates_ItemWise      
	,0 as PAmt,TSPL_ITEM_MASTER.Summary_Seq_No
from TSPL_CUSTOMER_MASTER
Left Outer Join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No  
LEFT Outer Join ( Select Document_No,Max(Document_Date)Document_Date,Route_No,Max(ShiftType)ShiftType from TSPL_DEMAND_BOOKING_MASTER Where CONVERT(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)=Convert(Date,'" & txtDate.Value & "',103) Group By Document_No,Route_No )TSPL_DEMAND_BOOKING_MASTER On TSPL_DEMAND_BOOKING_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  
Left Outer Join TSPL_Demand_Booking_Detail On TSPL_Demand_Booking_Detail.Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No
Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code  
left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code

where 2=2  And TSPL_CUSTOMER_MASTER.Status='N'  and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" & txtRouteNo.Value & "' and TSPL_DEMAND_BOOKING_MASTER.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "' 
and CONVERT(date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= Convert(Date,'" & txtDate.Value & "',103)  and TSPL_CUSTOMER_MASTER.Cust_Code Not In ( Select Cust_Code from TSPL_DEMAND_BOOKING_DETAIL Where Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No) "
            End If



            Qry += "" & BaseQry & " ) XXXFirst Group By
    XXXFirst.Cust_Code,	XXXFirst.Item_Code  ) 
    as s "
            If rdbnFreshAmbientBoth.IsChecked Then
                Qry += "  pivot (  sum(Qty_Crate) for Alies_Name#C in (" & strItemC & " ) ) as zpivot 
    pivot (  sum(Qty_Pouch) for Alies_Name#P in (" & strItemP & " )  ) as zpivot1
    pivot (  sum(TotalLtr_ItemWise) for Alies_Name#L in (" & strItemL & ") ) as zpivotLtr_ItemWise
    pivot (  sum(TotalAmt_ItemWise) for Alies_Name#A in (" & strItemA & ") ) as zpivotAmt_ItemWise "
                If clsCommon.myLen(strProdQ) > 0 Then
                    Qry += " pivot (  sum(ProdQ) for Alies_Name#ProdQ in (" & strProdQ & ") ) as zpivotProdQ "
                End If
            ElseIf rbtn_Fresh.IsChecked Then
                Qry += "  pivot (  sum(Qty_Crate) for Alies_Name#C in (" & strItemC & " ) ) as zpivot 
    pivot (  sum(Qty_Pouch) for Alies_Name#P in (" & strItemP & " )  ) as zpivot1
    pivot (  sum(TotalLtr_ItemWise) for Alies_Name#L in (" & strItemL & ") ) as zpivotLtr_ItemWise
    pivot (  sum(TotalAmt_ItemWise) for Alies_Name#A in (" & strItemA & ") ) as zpivotAmt_ItemWise "
            ElseIf rbtn_Ambient.IsChecked Then
                Qry += " pivot (  sum(TotalAmt_ItemWise) for Alies_Name#A in (" & strItemA & ") ) as zpivotAmt_ItemWise
    pivot (  sum(ProdQ) for Alies_Name#ProdQ in (" & strProdQ & ") ) as zpivotProdQ "
            End If
            Qry += "  group by Cust_Code,Display_Seq order by Display_Seq"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            'Total Row
            Dim newTotalRow As DataRow = dt.NewRow
            newTotalRow("Agents") = "Net Total"
            Dim newTotalLtrRow As DataRow = Nothing
            Dim newTotalAmtRow As DataRow = Nothing
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                newTotalLtrRow = dt.NewRow
                If rdbnFreshAmbientBoth.IsChecked OrElse rbtn_Fresh.IsChecked Then
                    newTotalLtrRow("Agents") = "Total Qty(Ltr)"
                End If
                newTotalAmtRow = dt.NewRow
                newTotalAmtRow("Agents") = "Total Amt"
            End If
            For i As Integer = 3 To dt.Columns.Count - 1
                Dim ColName As String = dt.Columns(i).ColumnName
                newTotalRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColName & "])", ""))
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    If ColName.Contains("#C") Then
                        Dim ColNameLtr As String = ColName.Substring(0, ColName.Length - 2) & "#L"
                        newTotalLtrRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameLtr & "])", ""))
                        Dim ColNameAmt As String = ColName.Substring(0, ColName.Length - 2) & "#A"
                        newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                    ElseIf ColName.Contains("#ProdQ") Then
                        Dim ColNameAmt As String = ColName.Substring(0, ColName.Length - 6) & "#A"
                        newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                    ElseIf ColName.Contains("Milk Amount") Then
                        Dim ColNameAmt As String = ColName
                        newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                    ElseIf ColName.Contains("Product Amount") Then
                        Dim ColNameAmt As String = ColName
                        newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                    End If
                    If ColName.Contains("Total Amount") Then
                        Dim ColNameAmt As String = ColName
                        newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                    End If
                End If
            Next
            dt.Rows.Add(newTotalRow)
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                If rdbnFreshAmbientBoth.IsChecked OrElse rbtn_Fresh.IsChecked Then
                    dt.Rows.Add(newTotalLtrRow)
                End If
                dt.Rows.Add(newTotalAmtRow)
            End If


            Dim dtUOM As DataTable = clsDBFuncationality.GetDataTable("Select Item_code, Max(Unit_Desc)Unit_Desc,Max(Is_FreshItem)Is_FreshItem,Max(Is_Ambient)Is_Ambient from (" & BaseQry & ")xyz group by Item_Code,Unit_Code")
            Dim strFUOM As String = Nothing
            Dim strFUOMPivot As String = Nothing
            Dim strPUOM As String = Nothing
            Dim strPUOMPivot As String = Nothing
            Dim strFSumItem As String = Nothing
            Dim strPSumItem As String = Nothing
            If dtUOM IsNot Nothing AndAlso dtUOM.Rows.Count > 0 Then
                For Each row In dtUOM.Rows
                    If clsCommon.myCdbl(row("Is_FreshItem")) = 1 Then
                        If clsCommon.myLen(strFUOMPivot) > 0 AndAlso Not strFUOMPivot.Contains("[" & clsCommon.myCstr(row("Unit_Desc")) & "]") Then
                            strFUOM += ",IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                            strFUOMPivot += ",[" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                            strFSumItem += ",Sum(IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0)) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                        Else
                            If clsCommon.myLen(strFUOMPivot) <= 0 OrElse Not strFUOMPivot.Contains("[" & clsCommon.myCstr(row("Unit_Desc")) & "]") Then
                                strFUOM = " IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                                strFUOMPivot = " [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                                strFSumItem = " Sum(IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0)) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                            End If
                        End If
                    End If
                    If clsCommon.myCdbl(row("Is_Ambient")) = 1 Then
                        If clsCommon.myLen(strPUOMPivot) > 0 AndAlso Not strPUOMPivot.Contains("[" & clsCommon.myCstr(row("Unit_Desc")) & "]") Then
                            strPUOM += ",IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                            strPUOMPivot += ",[" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                            strPSumItem += ",Sum(IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0)) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                        Else
                            If clsCommon.myLen(strPUOMPivot) <= 0 OrElse Not strPUOMPivot.Contains("[" & clsCommon.myCstr(row("Unit_Desc")) & "]") Then
                                strPUOM = " IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                                strPUOMPivot = " [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                                strPSumItem = " Sum(IsNull([" & clsCommon.myCstr(row("Unit_Desc")) & "],0)) As [" & clsCommon.myCstr(row("Unit_Desc")) & "]"
                            End If
                        End If
                    End If
                Next
            End If

            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                GVTruckSheet.DataSource = dt
            Else
                Throw New InvalidOperationException("Data Not Found !")
            End If
            Dim chkEntryUOM As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue("Select isnull(Entry_UOM,0) as Entry_UOM from TSPL_Route_Master where Route_No='" & clsCommon.myCstr(txtRouteNo.Value) & "'"))
            Dim chkCratePouch As Boolean = False
            Dim chkCrateLTR As Boolean = False
            Dim chkPouchLTR As Boolean = False
            Dim chkCrate As Boolean = False
            Dim chkLTR As Boolean = False
            Dim chkPouch As Boolean = False
            Dim chkALL As Boolean = False
            If chkEntryUOM = 0 Then
                chkCratePouch = True
            ElseIf chkEntryUOM = 1 Then
                chkCrate = True
            ElseIf chkEntryUOM = 2 Then
                chkLTR = True
            ElseIf chkEntryUOM = 3 Then
                chkCrateLTR = True
            ElseIf chkEntryUOM = 4 Then
                chkPouchLTR = True
            ElseIf chkEntryUOM = 5 Then
                chkPouch = True
            ElseIf chkEntryUOM = 6 Then
                chkALL = True
            End If
            If rdbnFreshAmbientBoth.IsChecked Then
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#C").FormatString = "{0:n0}"
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#P").FormatString = "{0:n0}"
                    'GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").TextAlignment = ContentAlignment.MiddleCenter
                    'GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").TextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("SizeC")) '& " क्रेट"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("SizeP")) '& " थैली"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("SizeL")) '& "LTR"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").IsVisible = clsCommon.myCBool(IIf(chkCrate, True, IIf(chkCrateLTR, True, IIf(chkCratePouch, True, IIf(chkCrate, True, IIf(chkALL, True, False))))))
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").IsVisible = clsCommon.myCBool(IIf(chkCratePouch, True, IIf(chkPouch, True, IIf(chkPouchLTR, True, IIf(chkALL, True, False)))))
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").IsVisible = clsCommon.myCBool(IIf(chkLTR, True, IIf(chkCrateLTR, True, IIf(chkPouchLTR, True, IIf(chkALL, True, False))))) 'False 
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) + "#ProdQ").FormatString = "{0:n0}"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").TextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderText = clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Size")) '& " "
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
            ElseIf rbtn_Fresh.IsChecked Then
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#C").FormatString = "{0:n0}"
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#P").FormatString = "{0:n0}"
                    'GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").TextAlignment = ContentAlignment.MiddleCenter
                    'GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").TextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("SizeC")) '& " क्रेट"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("SizeP")) '& " थैली"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("SizeL")) '& "LTR"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").IsVisible = clsCommon.myCBool(IIf(chkCrate, True, IIf(chkCrateLTR, True, IIf(chkCratePouch, True, IIf(chkCrate, True, IIf(chkALL, True, False))))))
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").IsVisible = clsCommon.myCBool(IIf(chkCratePouch, True, IIf(chkPouch, True, IIf(chkPouchLTR, True, IIf(chkALL, True, False)))))
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").IsVisible = clsCommon.myCBool(IIf(chkLTR, True, IIf(chkCrateLTR, True, IIf(chkPouchLTR, True, IIf(chkALL, True, False))))) 'False 
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
            ElseIf rbtn_Ambient.IsChecked Then
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) + "#ProdQ").FormatString = "{0:n0}"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").TextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderText = clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Size")) '&  " "
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") <> CompairStringResult.Equal AndAlso Not isExportTruckSheet Then
                GVTruckSheet.Columns("Milk In Ltr").FormatString = "{0:n2}"
                GVTruckSheet.Columns("Milk In Ltr").TextAlignment = ContentAlignment.MiddleCenter
                GVTruckSheet.Columns("Crates").FormatString = "{0:n0}"
                GVTruckSheet.Columns("Crates").TextAlignment = ContentAlignment.MiddleCenter
                GVTruckSheet.Columns("Milk Amount").FormatString = "{0:n2}"
                GVTruckSheet.Columns("Milk Amount").TextAlignment = ContentAlignment.MiddleRight
                GVTruckSheet.Columns("Product Quantity").FormatString = "{0:n0}"
                GVTruckSheet.Columns("Product Quantity").TextAlignment = ContentAlignment.MiddleCenter
                GVTruckSheet.Columns("Product Amount").FormatString = "{0:n2}"
                GVTruckSheet.Columns("Product Amount").TextAlignment = ContentAlignment.MiddleRight
            End If
            GVTruckSheet.Columns("Total Amount").FormatString = "{0:n2}"
            GVTruckSheet.Columns("Total Amount").TextAlignment = ContentAlignment.MiddleRight

            Dim view As New ColumnGroupsViewDefinition()
            Dim TempColGroupCount As Integer = 0
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("SNo").Name)
            TempColGroupCount += 1
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Code").Name)
            TempColGroupCount += 1
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Agents").Name)
            TempColGroupCount += 1
            If rdbnFreshAmbientBoth.IsChecked Then
                For Each dr As DataRow In dtDataExist.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#C").Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#P").Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#L").Name)
                    TempColGroupCount += 1
                Next
                For Each dr As DataRow In dtDataExistProduct.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#ProdQ").Name)
                    TempColGroupCount += 1
                Next
            ElseIf rbtn_Fresh.IsChecked Then
                For Each dr As DataRow In dtDataExist.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#C").Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#P").Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#L").Name)
                    TempColGroupCount += 1
                Next
            ElseIf rbtn_Ambient.IsChecked Then
                For Each dr As DataRow In dtDataExistProduct.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(dr("Alies_Name") & "#ProdQ").Name)
                    TempColGroupCount += 1
                Next
            End If
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") <> CompairStringResult.Equal AndAlso Not isExportTruckSheet Then
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Milk In Ltr").Name)
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Crates").Name)
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Milk Amount").Name)
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Product Quantity").Name)
                view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Product Amount").Name)
            End If
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Total Amount").Name)
            GVTruckSheet.ViewDefinition = view

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal OrElse isExportTruckSheet Then
                Dim newRow As DataRow
                Dim dtNew As New DataTable()
                Dim strQry As String
                Dim dtFresh As DataTable = Nothing
                Dim dtFreshTotal As DataTable = Nothing
                Dim dtAmbient As DataTable = Nothing
                Dim dtAmbientTotal As DataTable = Nothing
                Dim GVFresh As New UserControls.MyRadGridView
                'Dim GVFreshTotal As New UserControls.MyRadGridView

                If strFUOMPivot IsNot Nothing AndAlso (rbtn_Fresh.IsChecked OrElse rdbnFreshAmbientBoth.IsChecked) Then
                    strQry = "Select Alies_Name As [Product Name]," & strFUOM & ",[Amount] from (Select Item_Code,Max(Alies_Name)Alies_Name,MAX(Unit_Desc)Unit_Desc,Sum(Qty)Qty ,SUM(ItemNetAmount)[Amount],MAX(Summary_Seq_No)Summary_Seq_No  
from (" & BaseQry & ")xyz where Is_FreshItem=1 And Qty>0 group By  Item_code,Unit_Code) AS SourceTable PIVOT ( SUM(Qty) FOR Unit_Desc IN (" & strFUOMPivot & ") ) AS PivotTable Order By Summary_Seq_No"
                    dtFresh = clsDBFuncationality.GetDataTable(strQry)
                    If dtFresh IsNot Nothing AndAlso dtFresh.Rows.Count > 0 Then
                        GVFresh.Refresh()
                        GVFresh.DataSource = dtFresh
                    End If
                    'strQry += " Union All "
                    strQry = "Select 'Total : ' As [Product Name]," & strFSumItem & ",Sum([Amount])[Amount] from (Select Item_Code,Max(Alies_Name)Alies_Name,MAX(Unit_Desc)Unit_Desc,Sum(Qty)Qty ,SUM(ItemNetAmount)[Amount] 
from (" & BaseQry & ")xyz where Is_FreshItem=1 And Qty>0 group By  Item_code,Unit_Code) AS SourceTable PIVOT ( SUM(Qty) FOR Unit_Desc IN (" & strFUOMPivot & ") ) AS PivotTable"
                    dtFreshTotal = clsDBFuncationality.GetDataTable(strQry)
                End If

                If strPUOMPivot IsNot Nothing AndAlso (rbtn_Ambient.IsChecked OrElse rdbnFreshAmbientBoth.IsChecked) Then
                    strQry = "Select Alies_Name As [Product Name]," & strPUOM & ",[Amount] from (Select Item_Code,Max(Alies_Name)Alies_Name,MAX(Unit_Desc)Unit_Desc,Sum(Qty)Qty ,SUM(ItemNetAmount)[Amount] ,MAX(Summary_Seq_No)Summary_Seq_No
from (" & BaseQry & ")xyz where Is_Ambient=1 And Qty>0 group By  Item_code,Unit_Code) AS SourceTable PIVOT ( SUM(Qty) FOR Unit_Desc IN (" & strPUOMPivot & ") ) AS PivotTable Order By Summary_Seq_No "
                    dtAmbient = clsDBFuncationality.GetDataTable(strQry)
                    'strQry += " Union All "
                    strQry = "Select 'Total : ' As [Product Name]," & strPSumItem & ",Sum([Amount])[Amount] from (Select Item_Code,Max(Alies_Name)Alies_Name,MAX(Unit_Desc)Unit_Desc,Sum(Qty)Qty ,SUM(ItemNetAmount)[Amount]
from (" & BaseQry & ")xyz where Is_Ambient=1 And Qty>0 group By  Item_code,Unit_Code) AS SourceTable PIVOT ( SUM(Qty) FOR Unit_Desc IN (" & strPUOMPivot & ") ) AS PivotTable "
                    dtAmbientTotal = clsDBFuncationality.GetDataTable(strQry)
                End If

                ' Access the existing DataTable from RadGridView
                Dim dtOld As DataTable = TryCast(GVTruckSheet.DataSource, DataTable)

                If dtOld IsNot Nothing Then
                    '' Create new columns with String data type
                    'For Each col As DataColumn In dtOld.Columns
                    '    dtNew.Columns.Add(col.ColumnName, GetType(String))
                    'Next

                    '' Copy rows and convert all values to String
                    'For Each row As DataRow In dtOld.Rows
                    '    newRow = dtNew.NewRow()
                    '    For Each col As DataColumn In dtOld.Columns
                    '        newRow(col.ColumnName) = row(col.ColumnName).ToString()
                    '    Next
                    '    dtNew.Rows.Add(newRow)
                    'Next

                    ' Convert all columns to String type
                    For Each col As DataColumn In dtOld.Columns
                        dtNew.Columns.Add(col.ColumnName, GetType(String))
                    Next
                    ' Copy data with replacements
                    Dim chkNetTotal As String = Nothing
                    For Each row As DataRow In dtOld.Rows
                        newRow = dtNew.NewRow()
                        For Each col As DataColumn In dtOld.Columns
                            Dim cellValue As Object = row(col)
                            If clsCommon.CompairString(chkNetTotal, "Net Total") <> CompairStringResult.Equal Then
                                chkNetTotal = clsCommon.myCstr(cellValue)
                            End If

                            ' If numeric, check for 0 and replace
                            If clsCommon.CompairString(chkNetTotal, "Net Total") = CompairStringResult.Equal Then
                                newRow(col.ColumnName) = clsCommon.myCstr(cellValue)
                            Else
                                If IsNumeric(cellValue) AndAlso clsCommon.CompairString(clsCommon.myCstr(col), "Total Amount") <> CompairStringResult.Equal Then
                                    If Convert.ToDouble(cellValue) = 0 Then
                                        newRow(col.ColumnName) = "" ' Replace 0 with "-"
                                    Else
                                        newRow(col.ColumnName) = clsCommon.myCstr(cellValue) ' Convert to string
                                    End If
                                Else
                                    newRow(col.ColumnName) = clsCommon.myCstr(cellValue) ' Convert non-numeric to string
                                End If
                            End If
                        Next
                        dtNew.Rows.Add(newRow)
                    Next
                    ' Bind the converted DataTabl
                End If

                Dim isVisibleCol As Integer = 0
                For ii As Integer = 0 To GVTruckSheet.Columns.Count - 1
                    If GVTruckSheet.Columns(ii).IsVisible Then
                        isVisibleCol += 1
                    End If
                Next


                Dim totFreshAmbCol As Integer = 0
                If dtFresh IsNot Nothing AndAlso dtFresh.Columns.Count > 0 Then
                    totFreshAmbCol += dtFresh.Columns.Count
                End If
                If dtAmbient IsNot Nothing AndAlso dtAmbient.Columns.Count > 0 Then
                    totFreshAmbCol += dtAmbient.Columns.Count
                End If
                'If isVisibleCol <= totFreshAmbCol Then
                Dim colToAdd As Integer = (totFreshAmbCol - isVisibleCol) + 5
                For i As Integer = 1 To colToAdd
                    Dim newCol As New GridViewTextBoxColumn("ExtraCol" & i)
                    newCol.HeaderText = ""
                    newCol.IsVisible = True
                    GVTruckSheet.Columns.Add(newCol)
                    dtNew.Columns.Add("ExtraCol" & i, GetType(String))
                Next
                GVTruckSheet.Refresh()
                'End If
                dtNew.Rows.Add(dtNew.NewRow)

                Dim colk As Integer = 1
                colk = GetNextvisibleColumn(GVTruckSheet, colk)
                Dim headerRow As DataRow = dtNew.NewRow()
                If dtFresh IsNot Nothing Then
                    If dtFresh.Columns.Count > 0 Then
                        For cc As Integer = 0 To dtFresh.Columns.Count - 1
                            headerRow(colk) = clsCommon.myCstr(dtFresh.Columns(cc).ColumnName)
                            colk = GetNextvisibleColumn(GVTruckSheet, colk)
                        Next
                    Else
                        For cc As Integer = 1 To dtFresh.Columns.Count
                            colk = GetNextvisibleColumn(GVTruckSheet, colk)
                        Next
                    End If
                End If
                If dtAmbient IsNot Nothing AndAlso dtAmbient.Columns.Count > 0 Then
                    'If dtAmbient.Columns.Count > 0 Then
                    For cc As Integer = 0 To dtAmbient.Columns.Count - 1
                        headerRow(colk) = clsCommon.myCstr(dtAmbient.Columns(cc).ColumnName)
                        colk = GetNextvisibleColumn(GVTruckSheet, colk)
                    Next
                    'End If
                End If
                dtNew.Rows.Add(headerRow)
                dtNew.AcceptChanges()

                Dim maxLoop As Integer = 0
                If dtAmbient IsNot Nothing AndAlso dtAmbient.Rows.Count > 0 Then
                    maxLoop = dtAmbient.Rows.Count
                End If

                If dtFresh IsNot Nothing AndAlso dtFresh.Rows.Count > 0 AndAlso dtFresh.Rows.Count > maxLoop Then
                    maxLoop = dtFresh.Rows.Count
                End If
                If maxLoop > 0 Then
                    For ii As Integer = 0 To maxLoop - 1
                        Dim kk As Integer = 1
                        kk = GetNextvisibleColumn(GVTruckSheet, kk)
                        Dim dr As DataRow = dtNew.NewRow
                        If dtFresh IsNot Nothing AndAlso dtFresh.Rows.Count > 0 Then
                            If ii < dtFresh.Rows.Count Then
                                For cc As Integer = 0 To dtFresh.Columns.Count - 1
                                    If clsCommon.CompairString(clsCommon.myCstr(dtFresh.Rows(ii)(cc)), "0") <> CompairStringResult.Equal Then
                                        dr(kk) = clsCommon.myCstr(dtFresh.Rows(ii)(cc))
                                    Else
                                        dr(kk) = ""
                                    End If
                                    kk = GetNextvisibleColumn(GVTruckSheet, kk)
                                Next
                            Else
                                For cc As Integer = 1 To dtFresh.Columns.Count
                                    kk = GetNextvisibleColumn(GVTruckSheet, kk)
                                Next
                            End If
                        End If
                        If dtAmbient IsNot Nothing AndAlso dtAmbient.Rows.Count > 0 AndAlso ii < dtAmbient.Rows.Count Then
                            'If ii < dtAmbient.Rows.Count Then
                            For cc As Integer = 0 To dtAmbient.Columns.Count - 1
                                If clsCommon.CompairString(clsCommon.myCstr(dtAmbient.Rows(ii)(cc)), "0") <> CompairStringResult.Equal Then
                                    dr(kk) = clsCommon.myCstr(dtAmbient.Rows(ii)(cc))
                                Else
                                    dr(kk) = ""
                                End If
                                kk = GetNextvisibleColumn(GVTruckSheet, kk)
                            Next
                            'End If
                        End If

                        dtNew.Rows.Add(dr)
                        dtNew.AcceptChanges()
                    Next



                    Dim i As Integer = 0
                    Dim tt As Integer = 1
                    tt = GetNextvisibleColumn(GVTruckSheet, tt)
                    Dim drt As DataRow = dtNew.NewRow
                    If dtFreshTotal IsNot Nothing AndAlso dtFreshTotal.Rows.Count > 0 Then
                        If i < dtFreshTotal.Rows.Count Then
                            For cc As Integer = 0 To dtFreshTotal.Columns.Count - 1
                                drt(tt) = clsCommon.myCstr(dtFreshTotal.Rows(i)(cc))
                                tt = GetNextvisibleColumn(GVTruckSheet, tt)
                            Next
                        Else
                            For cc As Integer = 1 To dtFreshTotal.Columns.Count
                                tt = GetNextvisibleColumn(GVTruckSheet, tt)
                            Next
                        End If
                    End If
                    If dtAmbientTotal IsNot Nothing AndAlso dtAmbientTotal.Rows.Count > 0 AndAlso i < dtAmbient.Rows.Count Then
                        'If i < dtAmbient.Rows.Count Then
                        For cc As Integer = 0 To dtAmbientTotal.Columns.Count - 1
                            drt(tt) = clsCommon.myCstr(dtAmbientTotal.Rows(i)(cc))
                            tt = GetNextvisibleColumn(GVTruckSheet, tt)
                        Next
                        'End If
                    End If
                    dtNew.Rows.Add(drt)
                    dtNew.AcceptChanges()
                    GVTruckSheet.Refresh()
                    GVTruckSheet.DataSource = dtNew
                End If
            Else
                Dim dtOld As DataTable = TryCast(GVTruckSheet.DataSource, DataTable)
                Dim newRow As DataRow
                Dim dtNew As New DataTable()
                If dtOld IsNot Nothing Then
                    ' Convert all columns to String type
                    For Each col As DataColumn In dtOld.Columns
                        dtNew.Columns.Add(col.ColumnName, GetType(String))
                    Next
                    ' Copy data with replacements
                    Dim chkNetTotal As String = Nothing
                    For Each row As DataRow In dtOld.Rows
                        newRow = dtNew.NewRow()
                        For Each col As DataColumn In dtOld.Columns
                            Dim cellValue As Object = row(col)
                            'If clsCommon.CompairString(chkNetTotal, "Net Total") <> CompairStringResult.Equal Then
                            '    chkNetTotal = clsCommon.myCstr(cellValue)
                            'End If

                            ' If numeric, check for 0 and replace
                            If clsCommon.CompairString(chkNetTotal, "Net Total") = CompairStringResult.Equal Then
                                newRow(col.ColumnName) = clsCommon.myCstr(cellValue)
                            Else
                                If IsNumeric(cellValue) AndAlso clsCommon.CompairString(clsCommon.myCstr(col), "Total Amount") <> CompairStringResult.Equal Then
                                    If Convert.ToDouble(cellValue) = 0 Then
                                        newRow(col.ColumnName) = "" ' Replace 0 with "-"
                                    Else
                                        newRow(col.ColumnName) = clsCommon.myCstr(cellValue) ' Convert to string
                                    End If
                                Else
                                    newRow(col.ColumnName) = clsCommon.myCstr(cellValue) ' Convert non-numeric to string
                                End If
                            End If
                        Next
                        dtNew.Rows.Add(newRow)
                    Next
                End If
                dtNew.AcceptChanges()
                GVTruckSheet.Refresh()
                GVTruckSheet.DataSource = dtNew
            End If


            'GVTruckSheet.AutoSizeColumnsMode = Telerik.WinControls.UI.GridViewAutoSizeColumnsMode.None
            'GVTruckSheet.AutoSizeRows = True
            GVTruckSheet.BestFitColumns()

            For Each col As GridViewColumn In GVTruckSheet.Columns
                col.WrapText = True
                col.BestFit()
            Next

            GVTruckSheet.MasterTemplate.Refresh()
            GVTruckSheet.BeginUpdate()
            GVTruckSheet.EndUpdate()
            GVTruckSheet.Refresh()


            Dim arrHeader As List(Of String) = New List(Of String)()
            If isExcelPDF Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal OrElse isExportTruckSheet Then
                    Dim dtComp As DataTable = clsDBFuncationality.GetDataTable("Select ISO_No,ISO_Date from TSPL_COMPANY_MASTER Where Comp_code1='" & objCommonVar.CurrComp_Code1 & "'")
                    arrHeader.Add("Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy")) & "     " & "Shift : " & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "     " & "Trip No : " & clsCommon.myCstr(TripNo))
                    arrHeader.Add("Route : " & lblRouteDesc.Text & "     " & "City : " & lblCityName.Text & "     " & "Distributor : " & lblTransporterName.Text)
                    arrHeader.Add("ISO No : " & clsCommon.myCstr(dtComp.Rows(0)("ISO_No")) & "      ISO Date : " & IIf(clsCommon.myLen(dtComp.Rows(0)("ISO_Date")) > 0, clsCommon.GetPrintDate(dtComp.Rows(0)("ISO_Date"), "dd/MM/yyyy"), ""))
                    arrHeader.Add("")
                    'arrHeader.Add("Shift : " & IIf(rbtnMorning.IsChecked, "Morning", "Evening"))
                    'arrHeader.Add("Trip No : " & clsCommon.myCstr(TripNo))
                    'arrHeader.Add("Route : " & lblRouteDesc.Text)
                    'arrHeader.Add("City : " & lblCityName.Text)
                    'arrHeader.Add("Distributor : " & lblTransporterName.Text)
                Else
                    arrHeader.Add("Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy")) & "   Shift : " & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "   Trip No : " & clsCommon.myCstr(TripNo))
                    arrHeader.Add("Route : " & lblRouteDesc.Text & "    City : " & lblCityName.Text & "   Distributor : " & lblTransporterName.Text)
                End If
                'arrHeader.Add("Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy")))
                'arrHeader.Add("Route : " & lblRouteDesc.Text)
                'arrHeader.Add("City : " & lblCityName.Text)
                'arrHeader.Add("Shift : " & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening"))
                'arrHeader.Add("Distributor : " & lblTransporterName.Text)
                'arrHeader.Add("Trip : " & clsCommon.myCstr(txtTripNo.Text))
                'clsCommon.MyExportToExcelGrid(Nothing, GVTruckSheet, arrHeader, "Truck Sheet")
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                    transportSql.exportdata(True, Nothing, GVTruckSheet, "", "Truck Sheet", 0, GVTruckSheet.Rows.Count, False, arrHeader, False, False, True, False, False, Nothing, True, False)
                Else
                    transportSql.exportdata(True, Nothing, GVTruckSheet, "", "Truck Sheet", 0, GVTruckSheet.Rows.Count, False, arrHeader, False, False, True, False, False, Nothing, True, True)
                End If
            Else
                'doc.HeaderHeight = 60
                'doc.Landscape = True
                'doc.AssociatedObject = GVTruckSheet
                'doc.HeaderFont = New Font("Arial", 8)
                'doc.LeftUpperText = "Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy"))
                'doc.LeftUpperFont = New Font("Arial", 8)
                'doc.MiddleHeader = "City : " & lblCityName.Text
                'doc.RightHeader = "Shift : " & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening")
                'doc.LeftMiddleText = "Route : " & lblRouteDesc.Text
                'doc.LeftMiddleFont = New Font("Arial", 8)
                'doc.LeftLowerText = "Distributor : " & lblTransporterName.Text & " || " & "Trip : " & clsCommon.myCstr(txtTripNo.Text)
                'doc.LeftLowerFont = New Font("Arial", 8)
                doc.Margins.Top = 50
                doc.Margins.Bottom = 50
                doc.Margins.Left = 50
                doc.Margins.Right = 50
                doc.HeaderHeight = 90
                doc.Landscape = True
                doc.AssociatedObject = GVTruckSheet
                'Dim strHeader As String = "Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy"))
                'strHeader += "  Route : " & lblRouteDesc.Text
                'strHeader += "  City : " & lblCityName.Text
                'strHeader += "  Shift : " & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening")
                'strHeader += "  Distributor : " & lblTransporterName.Text
                'strHeader += "  Trip : " & clsCommon.myCstr(txtTripNo.Text)
                Dim strHeader As String = "Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy"))
                Dim strHeader2 As String = "Route : " & lblRouteDesc.Text
                strHeader2 += " City : " & lblCityName.Text
                strHeader2 += " Distributor : " & lblTransporterName.Text
                strHeader2 += " Trip : " & clsCommon.myCstr(txtTripNo.Text)
                strHeader2 += " Shift : " & IIf(rbtnMorning.IsChecked, "Morning", "Evening")
                doc.LeftUpperText = strHeader
                doc.LeftHeader = strHeader2
                doc.LeftUpperFont = New Font("Arial", 16, FontStyle.Bold)
                doc.HeaderFont = New Font("Arial", 16, FontStyle.Bold)
                doc.AssociatedObject = GVTruckSheet

                doc.Print()
                doc = Nothing
                'arrHeader.Add("Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy")))
                'arrHeader.Add("Shift : " & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening"))
                'arrHeader.Add("Trip No : " & clsCommon.myCstr(TripNo))
                'arrHeader.Add("Route : " & lblRouteDesc.Text)
                'arrHeader.Add("City : " & lblCityName.Text)
                'arrHeader.Add("Distributor : " & lblTransporterName.Text)
                'clsCommon.MyExportToPDF(Nothing, GVTruckSheet, arrHeader, Me.Text)
            End If

        Catch ex As InvalidOperationException
            doc = Nothing
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            Me.Controls.Remove(GVTruckSheet)
        End Try
    End Sub

    Public Sub GVTruckSheet_CellFormatting(sender As Object, e As CellFormattingEventArgs)
        Dim foundColumnIndex As Integer = -1 ' Store index of column containing "Cash Amount"

        If GVTruckSheet IsNot Nothing AndAlso GVTruckSheet.Rows.Count > 0 Then
            ' Iterate through rows to find the "Cash Amount" value
            For j As Integer = 0 To GVTruckSheet.Rows.Count - 1
                For ij As Integer = 0 To GVTruckSheet.Columns.Count - 1
                    If clsCommon.CompairString(clsCommon.myCstr(GVTruckSheet.Rows(j).Cells(ij).Value), "Cash Amount") = CompairStringResult.Equal Then
                        foundColumnIndex = ij
                        'GVTruckSheet.Columns(ij).FormatString = "{0:N2}" ' Store the column index where "Cash Amount" was found
#Disable Warning S3385 ' "Exit" statements should not be used
                        Exit For
#Enable Warning S3385 ' "Exit" statements should not be used
                    End If
                Next

                If e.RowIndex >= j AndAlso e.ColumnIndex = foundColumnIndex Then
                    e.CellElement.Text = String.Format("{0:N2}", Convert.ToDecimal(e.CellElement.Value)) ' Apply number format
                End If
                '' If "Cash Amount" was found, apply formatting to columns after it
                'If foundColumnIndex <> -1 Then
                '    'For k As Integer = foundColumnIndex To GVTruckSheet.Columns.Count - 1
                '    GVTruckSheet.Columns(k).FormatString = "{0:N2}" ' Apply formatting to subsequent columns
                '    'Next
                '    'Exit For ' Stop processing after applying formatting
                'End If
            Next
        End If

    End Sub

    Private Function GetNextvisibleColumn(gVTruckSheet As MyRadGridView, kk As Integer) As Integer
        Dim retValu As Integer = -1
        For ii As Integer = kk + 1 To gVTruckSheet.Columns.Count
            If ii <> gVTruckSheet.Columns.Count AndAlso gVTruckSheet.Columns(ii).IsVisible Then
                'If gVTruckSheet.Columns(ii).IsVisible Then
                retValu = ii
#Disable Warning S3385 ' "Exit" statements should not be used
                Exit For
#Enable Warning S3385 ' "Exit      " statements should not be used
                'End If
            End If
        Next
        Return retValu
    End Function

    Private Sub TruckSheetPDF()
        Dim GVTruckSheet As New UserControls.MyRadGridView()
        Me.Controls.Add(GVTruckSheet)
        Dim doc As New XpertERPEngine.clsMyPrintDocument()
        Try
            Dim ItemInUse As String = " TSPL_DEMAND_BOOKING_MASTER Left outer join TSPL_DEMAND_BOOKING_DETAIL
                On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
                Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
                where TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'
                and TSPL_ITEM_MASTER.Is_Milk_Pouch=1 order by sku_seq"
            Dim ItemInUseProduct As String = " TSPL_DEMAND_BOOKING_MASTER Left outer join TSPL_DEMAND_BOOKING_DETAIL
                On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
                Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
                where TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'
                and TSPL_ITEM_MASTER.Is_Milk_Pouch=0 order by sku_seq"
            Dim dtDataExist As DataTable = clsDBFuncationality.GetDataTable("select distinct isnull(TSPL_ITEM_MASTER.Alies_Name_Hindi,'')  Alies_Name,sku_seq,SUBSTRING(TSPL_ITEM_MASTER.Alies_Name, LEN(TSPL_ITEM_MASTER.Alies_Name) -  CHARINDEX(' ', REVERSE(TSPL_ITEM_MASTER.Alies_Name))+2,LEN(TSPL_ITEM_MASTER.Alies_Name)) AS Size from " & ItemInUse)
            Dim dtDataExistProduct As DataTable = clsDBFuncationality.GetDataTable("select distinct isnull(TSPL_ITEM_MASTER.Alies_Name_Hindi,'')  Alies_Name,sku_seq,SUBSTRING(TSPL_ITEM_MASTER.Alies_Name, LEN(TSPL_ITEM_MASTER.Alies_Name) -  CHARINDEX(' ', REVERSE(TSPL_ITEM_MASTER.Alies_Name))+2,LEN(TSPL_ITEM_MASTER.Alies_Name)) AS Size from " & ItemInUseProduct)
            If (dtDataExist Is Nothing OrElse dtDataExist.Rows.Count = 0) AndAlso (dtDataExistProduct Is Nothing OrElse dtDataExistProduct.Rows.Count = 0) Then
                ' clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
                Throw New InvalidOperationException("No Data Found")
#Disable Warning S3385 ' "Exit" statements should not be used
                Exit Sub
#Enable Warning S3385 ' "Exit" statements should not be used
            End If
            Dim sbstrItemC As New StringBuilder()
            Dim sbstrItemP As New StringBuilder()
            Dim sbstrItemL As New StringBuilder()
            Dim sbstrItemA As New StringBuilder()
            Dim sbstrProdQ As New StringBuilder()
            Dim sbstrItemSUM As New StringBuilder()

            'Dim strItemC As String = ""
            'Dim strItemP As String = ""
            'Dim strItemL As String = ""
            'Dim strItemA As String = ""
            'Dim strProdQ As String = ""
            'Dim strItemSUM As String = ""
            If rdbnFreshAmbientBoth.IsChecked Then
#Disable Warning S2259 ' Null pointers should not be dereferenced
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
#Enable Warning S2259 ' Null pointers should not be dereferenced
                    If i = 0 Then
                        sbstrItemC.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemC.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                Next
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    If sbstrItemSUM.Length = 0 Then
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                    If sbstrProdQ.Length = 0 Then
                        sbstrProdQ.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    Else
                        sbstrProdQ.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    End If
                Next
            ElseIf rbtn_Fresh.IsChecked Then
#Disable Warning S2259 ' Null pointers should not be dereferenced
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
#Enable Warning S2259 ' Null pointers should not be dereferenced
                    If i = 0 Then
                        sbstrItemC.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemC.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]")
                        sbstrItemP.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]")
                        sbstrItemL.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]")
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                Next
            ElseIf rbtn_Ambient.IsChecked Then
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    If sbstrItemSUM.Length = 0 Then
                        sbstrItemA.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append("sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    Else
                        sbstrItemA.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                        sbstrItemSUM.Append(",sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]
                    ,sum(IsNull([" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A],0)) as [" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A]")
                    End If
                    If sbstrProdQ.Length = 0 Then
                        sbstrProdQ.Append("[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    Else
                        sbstrProdQ.Append(",[" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ]")
                    End If
                Next
            End If
            Dim strItemC As String = sbstrItemC.ToString()
            Dim strItemP As String = sbstrItemP.ToString()
            Dim strItemL As String = sbstrItemL.ToString()
            Dim strItemA As String = sbstrItemA.ToString()
            Dim strProdQ As String = sbstrProdQ.ToString()
            Dim strItemSUM As String = sbstrItemSUM.ToString()
            Dim Qry As String = "select Cust_Code as AgentCode,max(customer_Name) as Agents
        , " & strItemSUM & "
                ,sum(isnull(TotalLtr_CustWise,0)) as [Milk In Ltr] 
        ,floor(sum(isnull(TotalCrates_ItemWise,0))) as [Crates]
        ,sum(isnull(MAmt,0)) as [Milk Amount]
        ,floor(sum(isnull(PQty,0))) as [Product Quantity]
        ,sum(isnull(PAmt,0)) as [Product Amount]
        from   
        (Select  max(Display_Seq) as Display_Seq,Cust_Code,max(Customer_Name) as Customer_Name,Item_Code,max(Alies_Name)+'#C' as Alies_Name#C
        ,max(Alies_Name)+'#P' AS Alies_Name#P
		,max(Alies_Name)+'#L' AS Alies_Name#L,max(Alies_Name)+'#A' AS Alies_Name#A
        ,max(Alies_Name)+'#ProdQ' AS Alies_Name#ProdQ
		,max(Unit_Desc) as Unit_Desc,max(Code) as Code
        ,floor(sum(Qty_Crate)) as Qty_Crate, floor(sum (Qty_Pouch)) as Qty_Pouch
        ,sum(TotalLtr_ItemWise) as TotalLtr_CustWise 
		,sum(TotalLtr_ItemWise) as TotalLtr_ItemWise
		,sum(ItemNetAmount) as TotalAmt_ItemWise
        ,SUM(TotalCrates_ItemWise) AS TotalCrates_ItemWise
        ,floor(sum(ProdQ)) as ProdQ
        ,SUM(MAmt) AS MAmt
        ,SUM(PQty) AS PQty
        ,SUM(PAmt) AS PAmt
        from (   Select '1' as  Code,TSPL_CUSTOMER_MASTER.Display_Seq,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,coalesce(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name 
	, TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Alies_Name_Hindi as Alies_Name
	,TSPL_UNIT_MASTER.Unit_Desc
	,(CASE WHEN TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Crate' THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE NULL END) as Qty_Crate
	,(CASE WHEN TSPL_DEMAND_BOOKING_DETAIL.Unit_code='Pouch' THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE NULL END) as Qty_Pouch
	,TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount
    ,TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise
    ,(CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) as ProdQ
	,(CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=1 THEN TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount ELSE 0 END) as MAmt
    ,(CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.Qty ELSE 0 END) as PQty
	,(CASE WHEN TSPL_ITEM_MASTER.Is_Milk_Pouch=0 THEN TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount ELSE 0 END) as PAmt
	 from TSPL_DEMAND_BOOKING_MASTER Left outer join TSPL_DEMAND_BOOKING_DETAIL
     On TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
     Left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
	left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
	left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
    WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'  ) XXXFirst Group By
    XXXFirst.Cust_Code,	XXXFirst.Item_Code  ) 
    as s "
            If rdbnFreshAmbientBoth.IsChecked Then
                Qry += "  pivot (  sum(Qty_Crate) for Alies_Name#C in (" & strItemC & " ) ) as zpivot 
    pivot (  sum(Qty_Pouch) for Alies_Name#P in (" & strItemP & " )  ) as zpivot1
    pivot (  sum(TotalLtr_ItemWise) for Alies_Name#L in (" & strItemL & ") ) as zpivotLtr_ItemWise
    pivot (  sum(TotalAmt_ItemWise) for Alies_Name#A in (" & strItemA & ") ) as zpivotAmt_ItemWise "
                If clsCommon.myLen(strProdQ) > 0 Then
                    Qry += " pivot (  sum(ProdQ) for Alies_Name#ProdQ in (" & strProdQ & ") ) as zpivotProdQ "
                End If
            ElseIf rbtn_Fresh.IsChecked Then
                Qry += "  pivot (  sum(Qty_Crate) for Alies_Name#C in (" & strItemC & " ) ) as zpivot 
    pivot (  sum(Qty_Pouch) for Alies_Name#P in (" & strItemP & " )  ) as zpivot1
    pivot (  sum(TotalLtr_ItemWise) for Alies_Name#L in (" & strItemL & ") ) as zpivotLtr_ItemWise
    pivot (  sum(TotalAmt_ItemWise) for Alies_Name#A in (" & strItemA & ") ) as zpivotAmt_ItemWise "
            ElseIf rbtn_Ambient.IsChecked Then
                Qry += " pivot (  sum(TotalAmt_ItemWise) for Alies_Name#A in (" & strItemA & ") ) as zpivotAmt_ItemWise
    pivot (  sum(ProdQ) for Alies_Name#ProdQ in (" & strProdQ & ") ) as zpivotProdQ "
            End If
            Qry += "  group by Cust_Code,Display_Seq order by Display_Seq"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            'Total Row
            Dim newTotalRow As DataRow = dt.NewRow
            newTotalRow("Agents") = "Net Total"
            Dim newTotalLtrRow As DataRow = dt.NewRow
            If rdbnFreshAmbientBoth.IsChecked OrElse rbtn_Fresh.IsChecked Then
                newTotalLtrRow("Agents") = "Total Qty(Ltr)"
            End If
            Dim newTotalAmtRow As DataRow = dt.NewRow
            newTotalAmtRow("Agents") = "Total Amt"
            For i As Integer = 2 To dt.Columns.Count - 1
                Dim ColName As String = dt.Columns(i).ColumnName
                newTotalRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColName & "])", ""))
                If ColName.Contains("#C") Then
                    Dim ColNameLtr As String = ColName.Substring(0, ColName.Length - 2) & "#L"
                    newTotalLtrRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameLtr & "])", ""))
                    Dim ColNameAmt As String = ColName.Substring(0, ColName.Length - 2) & "#A"
                    newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                ElseIf ColName.Contains("#ProdQ") Then
                    Dim ColNameAmt As String = ColName.Substring(0, ColName.Length - 6) & "#A"
                    newTotalAmtRow(ColName) = clsCommon.myCdbl(dt.Compute("sum([" & ColNameAmt & "])", ""))
                End If
            Next
            dt.Rows.Add(newTotalRow)
            If rdbnFreshAmbientBoth.IsChecked OrElse rbtn_Fresh.IsChecked Then
                dt.Rows.Add(newTotalLtrRow)
            End If
            dt.Rows.Add(newTotalAmtRow)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                GVTruckSheet.DataSource = dt
            End If
            GVTruckSheet.Columns("AgentCode").HeaderText = "Code"
            If rdbnFreshAmbientBoth.IsChecked Then
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#C").FormatString = "{0:n2}"
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#P").FormatString = "{0:n2}"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("Size")) & Environment.NewLine & "क्रेट"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderTextAlignment = ContentAlignment.TopLeft
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("Size")) & Environment.NewLine & "थैली"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderTextAlignment = ContentAlignment.TopLeft
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").IsVisible = False
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) + "#ProdQ").FormatString = "{0:n2}"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderText = clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Size")) & Environment.NewLine & " "
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderTextAlignment = ContentAlignment.TopLeft
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
            ElseIf rbtn_Fresh.IsChecked Then
                For i As Integer = 0 To dtDataExist.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#C").FormatString = "{0:n2}"
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) + "#P").FormatString = "{0:n2}"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("Size")) & Environment.NewLine & "क्रेट"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#C").HeaderTextAlignment = ContentAlignment.TopLeft
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderText = clsCommon.myCstr(dtDataExist.Rows(i).Item("Size")) & Environment.NewLine & "थैली"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#P").HeaderTextAlignment = ContentAlignment.TopLeft
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#L").IsVisible = False
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExist.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
            ElseIf rbtn_Ambient.IsChecked Then
                For i As Integer = 0 To dtDataExistProduct.Rows.Count - 1
                    'GVTruckSheet.Columns("" + clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) + "#ProdQ").FormatString = "{0:n2}"
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderText = clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Size")) & Environment.NewLine & " "
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#ProdQ").HeaderTextAlignment = ContentAlignment.TopLeft
                    GVTruckSheet.Columns("" & clsCommon.myCstr(dtDataExistProduct.Rows(i).Item("Alies_Name")) & "#A").IsVisible = False
                Next
            End If
            GVTruckSheet.Columns("Milk In Ltr").FormatString = "{0:n2}"
            'GVTruckSheet.Columns("Crates").FormatString = "{0:n2}"
            GVTruckSheet.Columns("Milk Amount").FormatString = "{0:n2}"
            'GVTruckSheet.Columns("Product Quantity").FormatString = "{0:n2}"
            GVTruckSheet.Columns("Product Amount").FormatString = "{0:n2}"
            GVTruckSheet.Columns("Milk Amount").HeaderText = "Milk Amt"
            GVTruckSheet.Columns("Product Quantity").HeaderText = "Prod Qty"
            GVTruckSheet.Columns("Product Amount").HeaderText = "Prod Amt"
            'GVTruckSheet.Columns("Milk In Ltr").Width = 50
            'GVTruckSheet.Columns("Crates").Width = 25
            'GVTruckSheet.Columns("Milk Amount").Width = 50
            'GVTruckSheet.Columns("Product Quantity").Width = 50
            'GVTruckSheet.Columns("Product Amount").Width = 50
            'GVTruckSheet.MasterTemplate.BestFitColumns()
            If rbtn_Fresh.IsChecked Then
                GVTruckSheet.Columns("Milk Amount").IsVisible = False
                GVTruckSheet.Columns("Product Quantity").IsVisible = False
                GVTruckSheet.Columns("Product Amount").IsVisible = False
            ElseIf rbtn_Ambient.IsChecked Then
                GVTruckSheet.Columns("Milk Amount").IsVisible = False
                GVTruckSheet.Columns("Milk In Ltr").IsVisible = False
                GVTruckSheet.Columns("Crates").IsVisible = False
                GVTruckSheet.Columns("Product Amount").IsVisible = False
            End If
            For I As Int16 = 0 To GVTruckSheet.Columns.Count - 1
                GVTruckSheet.Columns(I).WrapText = True
                'GVTruckSheet.Columns(I).BestFit()
            Next
            GVTruckSheet.Columns("Agents").Width = 100
            GVTruckSheet.Columns("Agents").HeaderTextAlignment = ContentAlignment.TopLeft
            GVTruckSheet.Columns("AgentCode").HeaderTextAlignment = ContentAlignment.TopLeft
            Dim view As New ColumnGroupsViewDefinition()
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(0).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("AgentCode").Name)
            view.ColumnGroups(0).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Agents").Name)
            Dim TempColGroupCount As Integer = 1
            If rdbnFreshAmbientBoth.IsChecked Then
                For Each dr As DataRow In dtDataExist.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    Dim groupRow = New GridViewColumnGroupRow()
                    groupRow.MinHeight = 30
                    'view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(TempColGroupCount).Rows.Add(groupRow)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(clsCommon.myCstr(dr("Alies_Name")) & "#C").Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(clsCommon.myCstr(dr("Alies_Name")) & "#P").Name)
                    TempColGroupCount += 1
                Next
                For Each dr As DataRow In dtDataExistProduct.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    Dim groupRow = New GridViewColumnGroupRow()
                    groupRow.MinHeight = 30
                    'view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    view.ColumnGroups(TempColGroupCount).Rows.Add(groupRow)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(clsCommon.myCstr(dr("Alies_Name")) & "#ProdQ").Name)
                    TempColGroupCount += 1
                Next
            ElseIf rbtn_Fresh.IsChecked Then
                For Each dr As DataRow In dtDataExist.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    'view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    Dim groupRow = New GridViewColumnGroupRow()
                    groupRow.MinHeight = 30
                    view.ColumnGroups(TempColGroupCount).Rows.Add(groupRow)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(clsCommon.myCstr(dr("Alies_Name")) & "#C").Name)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(clsCommon.myCstr(dr("Alies_Name")) & "#P").Name)
                    TempColGroupCount += 1
                Next
            ElseIf rbtn_Ambient.IsChecked Then
                For Each dr As DataRow In dtDataExistProduct.Rows
                    view.ColumnGroups.Add(New GridViewColumnGroup(clsCommon.myCstr(dr("Alies_Name"))))
                    view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
                    'Dim groupRow = New GridViewColumnGroupRow()
                    'groupRow.MinHeight = 30
                    'view.ColumnGroups(TempColGroupCount).Rows.Add(groupRow)
                    view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns(clsCommon.myCstr(dr("Alies_Name")) & "#ProdQ").Name)
                    TempColGroupCount += 1
                Next
            End If
            view.ColumnGroups.Add(New GridViewColumnGroup(""))
            view.ColumnGroups(TempColGroupCount).Rows.Add(New GridViewColumnGroupRow())
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Milk In Ltr").Name)
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Crates").Name)
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Milk Amount").Name)
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Product Quantity").Name)
            view.ColumnGroups(TempColGroupCount).Rows(0).ColumnNames.Add(GVTruckSheet.Columns("Product Amount").Name)
            GVTruckSheet.ViewDefinition = view
            'doc.HeaderHeight = 60
            'doc.Landscape = True
            'doc.AssociatedObject = GVTruckSheet
            'doc.HeaderFont = New Font("Arial", 8)
            'doc.LeftUpperText = "Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy"))
            'doc.LeftUpperFont = New Font("Arial", 8)
            'doc.MiddleHeader = "City : " & lblCityName.Text
            'doc.RightHeader = "Shift : " & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening")
            'doc.LeftMiddleText = "Route : " & lblRouteDesc.Text
            'doc.LeftMiddleFont = New Font("Arial", 8)
            'doc.LeftLowerText = "Distributor : " & lblTransporterName.Text & " || " & "Trip : " & clsCommon.myCstr(txtTripNo.Text)
            'doc.LeftLowerFont = New Font("Arial", 8)
            doc.Margins.Top = 50
            doc.Margins.Bottom = 50
            doc.Margins.Left = 50
            doc.Margins.Right = 50
            doc.HeaderHeight = 90
            doc.Landscape = True
            doc.AssociatedObject = GVTruckSheet
            'Dim strHeader As String = "Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy"))
            'strHeader += "  Route : " & lblRouteDesc.Text
            'strHeader += "  City : " & lblCityName.Text
            'strHeader += "  Shift : " & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening")
            'strHeader += "  Distributor : " & lblTransporterName.Text
            'strHeader += "  Trip : " & clsCommon.myCstr(txtTripNo.Text)
            Dim strHeader As String = "Doc Date : " & clsCommon.myCstr(clsCommon.GetPrintDate(txtDate.Value, "dd-MMM-yyyy"))
            Dim strHeader2 As String = "Route : " & lblRouteDesc.Text
            strHeader2 += " City : " & lblCityName.Text
            strHeader2 += " Distributor : " & lblTransporterName.Text
            strHeader2 += " Trip : " & clsCommon.myCstr(txtTripNo.Text)
            strHeader2 += " Shift : " & IIf(rbtnMorning.IsChecked, "Morning", "Evening")
            doc.LeftUpperText = strHeader
            doc.LeftHeader = strHeader2
            doc.LeftUpperFont = New Font("Arial", 16, FontStyle.Bold)
            doc.HeaderFont = New Font("Arial", 16, FontStyle.Bold)
            doc.AssociatedObject = GVTruckSheet
            doc.Print()
            doc = Nothing
        Catch ex As InvalidOperationException
            doc = Nothing
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            Me.Controls.Remove(GVTruckSheet)
        End Try
    End Sub
    Private Sub Btn_Gatepass_Click(sender As Object, e As EventArgs) Handles btn_Gatepass.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select document", Me.Text)
                txtDocNo.Focus()
                Exit Sub
            End If
            If rbtnMorningEveningBoth.IsChecked Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Morning/Evening Shift", Me.Text)
                rbtnMorning.Focus()
                Exit Sub
            End If
            '' Commneted IsGatePassGenerated update 'Y' 20/08/2024
            '' clsDBFuncationality.ExecuteNonQuery("update TSPL_DEMAND_BOOKING_DETAIL set IsGatePassGenerated='Y' where document_no='" & txtDocNo.Value & "' and ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'")
            '' Comment End
            'Dim Qry As String = "select TSPL_DEMAND_BOOKING_MASTER.TripNo,
            '    TSPL_DEMAND_BOOKING_MASTER.shiftType,
            '    TSPL_city_MASTER.City_Name,
            '    TSPL_DEMAND_BOOKING_MASTER.Comp_Code,
            '    TSPL_DEMAND_BOOKING_MASTER.location_code
            '    ,TSPL_DEMAND_BOOKING_MASTER.Document_No as Demand_No,convert(varchar(15),TSPL_DEMAND_BOOKING_MASTER.Document_Date,103) as Demand_Date 
            '    ,isnull(TSPL_DEMAND_BOOKING_MASTER.Route_No,'') as Route_No ,isnull(TSPL_ROUTE_MASTER.Route_Desc,'') as Route_Desc 
            '     ,TSPL_DEMAND_BOOKING_DETAIL.cust_code,TSPL_CUSTOMER_MASTER.Customer_Name_Hindi
            '    ,TSPL_DEMAND_BOOKING_DETAIL.Qty
            '    ,TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise
            '    ,TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise
            '    ,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount
            '    ,TSPL_DEMAND_BOOKING_DETAIL.unit_code
            '    ,TSPL_ITEM_MASTER.Alies_Name_Hindi as Item_Alies_Name_Hindi 
            '    ,case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Crate' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end as CrateQty
            '    ,case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Pouch' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end as PouchQty
            '    ,N'क्रेट' as Crate,N'क्रेट    थैली' as Pouch
            '    from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL
            '    on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
            '     left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code 
            '    left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
            '     left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
            '    left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DEMAND_BOOKING_MASTER.Route_No
            '    left outer join TSPL_city_MASTER on TSPL_city_MASTER.city_code = TSPL_DEMAND_BOOKING_MASTER.city_code
            '    left outer join TSPL_UNIT_MASTER on TSPL_UNIT_MASTER.Unit_Code=TSPL_DEMAND_BOOKING_DETAIL.Unit_code
            '    LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code =TSPL_DEMAND_BOOKING_MASTER.location_code
            '    WHERE TSPL_ITEM_MASTER.Is_Milk_Pouch=1
            '    and TSPL_DEMAND_BOOKING_MASTER.Document_No='" + txtDocNo.Value + "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'"
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            'If dt.Rows.Count > 0 Then
            '    Dim frmCRV As New frmCrystalReportViewer()
            '    frmCRV.funsubreportWithdt(CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePass", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            '    frmCRV = Nothing
            'End If
            'Dim Qry As String = "select TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1 as Comp_Add1,TSPL_COMPANY_MASTER.Add2 as Comp_Add2,TSPL_COMPANY_MASTER.Add3 as Comp_Add3,TSPL_COMPANY_MASTER.Pincode as Comp_Pin
            '      ,TSPL_LOCATION_MASTER.Add1,TSPL_LOCATION_MASTER.Add2,TSPL_LOCATION_MASTER.ADD3,TSPL_LOCATION_MASTER.Pin_Code,TSPL_LOCATION_MASTER.Location_Desc
            '      ,Main_Final.Distributor,'" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "' shiftType,Main_Final.City_Name,Main_Final.Demand_No,Main_Final.Demand_Date,Main_Final.Route_No,Main_Final.Route_Desc
            '      ,Main_Final.Item_alies_name,Main_Final.Crate_Qty,Main_Final.Pouch_Qty,Main_Final.Loose_Qty,TotalLtr_ItemWise
            '      from (select max(TSPL_VENDOR_MASTER.vendor_name) as Distributor,
            '      max(TSPL_DEMAND_BOOKING_MASTER.shiftType) as shiftType,
            '      max(TSPL_city_MASTER.City_Name) as City_Name,
            '      max(TSPL_DEMAND_BOOKING_MASTER.Comp_Code) as Comp_Code,
            '      max(TSPL_DEMAND_BOOKING_MASTER.location_code) as location_code
            '      ,TSPL_DEMAND_BOOKING_MASTER.Document_No as Demand_No,max(convert(varchar(15),TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)) as Demand_Date ,isnull(TSPL_DEMAND_BOOKING_MASTER.Route_No,'') as Route_No ,max(isnull(TSPL_ROUTE_MASTER.Route_Desc,'')) as Route_Desc 
            '      ,max(TSPL_ITEM_MASTER.alies_name) as Item_alies_name
            '      ,sum(case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Crate' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Crate_Qty
            ',sum(case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Pouch' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Pouch_Qty
            ',sum(case when (TSPL_DEMAND_BOOKING_DETAIL.unit_code<>'Crate' and TSPL_DEMAND_BOOKING_DETAIL.unit_code<>'Pouch') then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Loose_Qty
            ',sum(TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise) AS TotalLtr_ItemWise
            '      from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL
            '      on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
            '       left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code 
            '      left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
            '       left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
            '      left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DEMAND_BOOKING_MASTER.Route_No
            '      left outer join TSPL_city_MASTER on TSPL_city_MASTER.city_code = TSPL_DEMAND_BOOKING_MASTER.city_code
            '      left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VEHICLE_MASTER.Transport_Id
            '      WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" + txtDocNo.Value + "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'
            '       group by TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_DEMAND_BOOKING_MASTER.Document_No,TSPL_DEMAND_BOOKING_MASTER.Route_No
            '      ) as Main_Final
            '      LEFT OUTER JOIN TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=Main_Final.Comp_Code
            '       LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code =Main_Final.location_code"
            PrintGatePass("DB", txtDocNo.Value, IIf(rbtnMorning.IsChecked = True, "Morning", "Evening"), rbtn_Fresh.IsChecked, rbtn_Ambient.IsChecked)
            'Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            'If dt.Rows.Count > 0 Then
            '    Dim frmCRV As New frmCrystalReportViewer()
            '    frmCRV.funsubreportWithdt(CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWise", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            '    frmCRV = Nothing
            'End If
            'btn_Gatepass.Enabled = False
            'btn_GPCancel.Enabled = True
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Function PrintGatePass(ByVal StrFormType As String, ByVal StrDocCode As String, ByVal StrShift As String, ByVal Fresh As Boolean, ByVal Ambient As Boolean) As String
        Return PrintGatePass(StrFormType, StrDocCode, StrShift, Fresh, Ambient, False)
    End Function
    Function PrintGatePass(ByVal StrFormType As String, ByVal StrDocCode As String, ByVal StrShift As String, ByVal Fresh As Boolean, ByVal Ambient As Boolean, ByVal isPdf As Boolean) As String
        'clsDBFuncationality.ExecuteNonQuery("update TSPL_DEMAND_BOOKING_DETAIL set IsGatePassGenerated='Y' where " + IIf(StrFormType = "DB", "TSPL_DEMAND_BOOKING_DETAIL.Document_No", "TSPL_DEMAND_BOOKING_DETAIL.GPCode") + "'='" & StrDocCode & "' and ShiftType='" & StrShift & "'")
        Dim filePath As String = Nothing
        Dim whr As String = ""
        If Fresh Then
            whr = "and TSPL_ITEM_MASTER.IsTaxable=0 "
        ElseIf Ambient Then
            whr = "and TSPL_ITEM_MASTER.IsTaxable=1 "
        Else
            whr = ""
        End If
        Dim mainTrip As String = ""
        Dim Trip As String = ""
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") <> CompairStringResult.Equal Then
            Trip = "     MAX(ISNULL(LEFT((el.files), CASE WHEN LEN((el.files)) > 1 THEN LEN((el.files)) - 1 ELSE 0 END), 'NoFile')) AS Trip,"
            mainTrip = " ,Main_Final.Trip "
        End If
        Dim Qry As String = "select  zone.Phone1 as Dist_Phn, zone.Zone_Code,FSSAI_NO, TSPL_COMPANY_MASTER.Phone1,TSPL_COMPANY_MASTER.Phone2,TSPL_COMPANY_MASTER.GSTINNo,  TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.Add1 as Comp_Add1,TSPL_COMPANY_MASTER.Add2 as Comp_Add2,TSPL_COMPANY_MASTER.Add3 as Comp_Add3,TSPL_COMPANY_MASTER.Pincode as Comp_Pin,TSPL_COMPANY_MASTER.Access_Officer as FSSAI_LIC_NO
                  ,TSPL_LOCATION_MASTER.Add1,TSPL_LOCATION_MASTER.Add2,TSPL_LOCATION_MASTER.ADD3,TSPL_LOCATION_MASTER.Pin_Code,TSPL_LOCATION_MASTER.Location_Desc,'" & objCommonVar.CurrentUser & "' as Currentuser
                  ,Main_Final.Distributor,Main_Final.Distributor_Code,Employee_Name,'" & StrShift & "' shiftType,Main_Final.City_Name,Main_Final.Demand_No,Main_Final.Demand_Date,Main_Final.Route_No,Main_Final.Route_Desc ,Main_Final.Vehicle_Desc
                  ,Main_Final.Item_alies_name,Main_Final.UOM,Main_Final.unit_code_result,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " CASE WHEN (Main_Final.Crate_Qty) - FLOOR(Main_Final.Crate_Qty) > 0.5 THEN CEILING(Main_Final.Crate_Qty) ELSE FLOOR(Main_Final.Crate_Qty) END AS Crate_Qty,"
        Else
            Qry += " Main_Final.Crate_Qty,"
        End If
        Qry += " Main_Final.Pouch_Qty,Main_Final.Loose_Qty,TotalLtr_ItemWise,ItemNetAmount,Main_Final.Production_Remarks" & mainTrip & ""
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
            Qry += " ,Main_Final.HSN_Code ,Main_Final.item_rate  ,Main_Final.TAX1,Main_Final.TAX2,Main_Final.TAX3,Main_Final.TAX4,Main_Final.TAX5,Main_Final.TAX6
				  ,Main_Final.TAX1_Amt,Main_Final.TAX2_Amt,Main_Final.TAX3_Amt,Main_Final.TAX4_Amt,Main_Final.TAX5_Amt,Main_Final.TAX6_Amt ,Main_Final.TAX7_Amt,Main_Final.TAX8_Amt,Main_Final.TAX7,Main_Final.TAX8  "
        End If
        Qry += "       from (select " & Trip & " max(TSPL_VENDOR_MASTER.vendor_name) as Distributor,max(TSPL_VENDOR_MASTER.Vendor_Code) as Distributor_Code,max(TSPL_customer_master.FSSAI_NO)FSSAI_NO ,
                  max(TSPL_DEMAND_BOOKING_MASTER.shiftType) as shiftType,max(TSPL_city_MASTER.City_Name) as City_Name,max(TSPL_DEMAND_BOOKING_MASTER.Comp_Code) as Comp_Code,max(TSPL_DEMAND_BOOKING_MASTER.location_code) as location_code,TSPL_DEMAND_BOOKING_MASTER.Document_No as Demand_No,max(convert(varchar(15),TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)) as Demand_Date ,isnull(TSPL_DEMAND_BOOKING_MASTER.Route_No,'') as Route_No ,max(isnull(TSPL_ROUTE_MASTER.Route_Desc,'')) as Route_Desc,max(TSPL_Route_Master.Employee_Name)Employee_Name
                  ,max(isnull(TSPL_VEHICLE_MASTER.Description,'')) as Vehicle_Desc ,max(TSPL_ITEM_MASTER.alies_name) as Item_alies_name,
                  max(TSPL_ITEM_MASTER.Unit_Code) as UOM,CASE WHEN max(TSPL_ITEM_MASTER.Unit_Code) = 'crate' OR max(TSPL_ITEM_MASTER.Unit_Code) = 'pouch' THEN 'Crate/Pouch' ELSE max(TSPL_ITEM_MASTER.Unit_Code) END AS unit_code_result,"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " Round(sum(TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise)*MAX(CInLTR.Conversion_Factor)/Max(CInCrate.Conversion_Factor),1) As Crate_Qty,"
        Else
            Qry += " sum(case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Crate' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Crate_Qty, "
        End If

        Qry += " sum(case when TSPL_DEMAND_BOOKING_DETAIL.unit_code='Pouch' then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Pouch_Qty
            ,sum(case when (TSPL_DEMAND_BOOKING_DETAIL.unit_code<>'Crate' and TSPL_DEMAND_BOOKING_DETAIL.unit_code<>'Pouch') then TSPL_DEMAND_BOOKING_DETAIL.Qty else 0 end) AS Loose_Qty
            ,sum(TSPL_DEMAND_BOOKING_DETAIL.TotalLtr_ItemWise) AS TotalLtr_ItemWise
                   ,sum(TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount) AS ItemNetAmount
                  ,max(TSPL_DEMAND_BOOKING_DETAIL.Production_Remarks) as Production_Remarks, max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq"

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then

            Qry += "	  ,max(TSPL_ITEM_MASTER.HSN_Code)HSN_Code,  max(TSPL_DEMAND_BOOKING_detail.item_rate)item_rate 	, '' AS TAX1,'' AS TAX2, '' AS TAX3,'' AS TAX4,'' AS TAX5,'' AS TAX6
				 ,'' AS TAX1_Amt,'' AS TAX2_Amt,'' AS TAX3_Amt,'' AS TAX4_Amt,'' AS TAX5_Amt,'' AS TAX6_Amt ,'' AS TAX7,'' AS TAX8, '' AS TAX7_Amt,'' AS TAX8_Amt"
        End If
        Qry += "     from TSPL_DEMAND_BOOKING_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL
                  on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No 
                   left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code =TSPL_DEMAND_BOOKING_DETAIL.Cust_Code 
                  left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
                  LEFT Outer Join (Select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL Where UOM_Code='LTR')CInLTR ON CInLTR.Item_Code=TSPL_ITEM_MASTER.Item_Code
                  LEFT Outer Join (Select Item_Code,UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL Where UOM_Code='CRATE')CInCrate ON CInCrate.Item_Code=TSPL_ITEM_MASTER.Item_Code
                   left outer join TSPL_VEHICLE_MASTER on TSPL_VEHICLE_MASTER.Vehicle_Id  =TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
                  left outer join TSPL_ROUTE_MASTER on TSPL_ROUTE_MASTER.Route_No = TSPL_DEMAND_BOOKING_MASTER.Route_No
                  left outer join TSPL_city_MASTER on TSPL_city_MASTER.city_code = TSPL_DEMAND_BOOKING_MASTER.city_code
                  left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VEHICLE_MASTER.Transport_Id
                  cross apply(select distinct  convert(varchar,TSPL_DEMAND_BOOKING_DETAIL.Trip_No) + ',' as [text()] from TSPL_DEMAND_BOOKING_DETAIL where Document_No= '" & StrDocCode & "' 
                  FOR XML PATH(''))el(files)
                  WHERE " & IIf(StrFormType = "DB", "TSPL_DEMAND_BOOKING_DETAIL.Document_No", "TSPL_DEMAND_BOOKING_DETAIL.GPCode") & " = '" & StrDocCode & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & StrShift & "'
                  " & whr & " group by TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_DEMAND_BOOKING_MASTER.Document_No,TSPL_DEMAND_BOOKING_MASTER.Route_No
                  ) as Main_Final
                  LEFT OUTER JOIN TSPL_COMPANY_MASTER ON 2=2
                   LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code =Main_Final.location_code
                  left join ( select distinct TSPL_CUSTOMER_MASTER.Phone1,  TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Route_No,TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code,TSPL_CUSTOMER_MASTER.Zone_Code from TSPL_DISTRIBUTOR_ROUTE_CUSTOMER 
                  left outer join TSPL_CUSTOMER_MASTER on TSPL_CUSTOMER_MASTER.Cust_Code = TSPL_DISTRIBUTOR_ROUTE_CUSTOMER.Cust_Code ) zone on zone.Route_No = Main_Final.Route_No and zone.Cust_Code = Main_Final.Distributor_Code order by Main_Final.Sku_Seq"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
        If dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWiseUDP", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal Then
                filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWiseCHITTORGARH", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
                filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWiseKTA", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWiseAJM", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            Else
                filePath = frmCRV.funsubreportWithdt(MyBase.Form_ID, isPdf, CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWise", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            End If
            'frmCRV.funsubreportWithdt(CrystalReportFolder.NewSalesReports, dt, clsERPFuncationality.CompanyAddresShowinFooter(), "rptDairySaleGatePassItemWise", "Gate Pass", clsCommon.myCDate(dt.Rows(0)("Demand_Date")), "rptCompanyAddress.rpt")
            frmCRV = Nothing
        End If
        Return filePath
    End Function
    Private Sub Btn_GPCancel_Click(sender As Object, e As EventArgs) Handles btn_GPCancel.Click
        Try
            If clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) > 0 Then
                If rbtnMorningEveningBoth.IsChecked Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please select Morning/Evening Shift", Me.Text)
                    rbtnMorning.Focus()
                    Exit Sub
                End If
                Dim qry As String = ""
                Dim StrGatePassCode As String = ""
                StrGatePassCode = clsDBFuncationality.getSingleValue("select max(GPCode) as GPCode from TSPL_DEMAND_BOOKING_DETAIL where Document_no='" & txtDocNo.Value & "'")
                If clsCommon.myLen(StrGatePassCode) > 0 Then
                    qry = "delete from TSPL_DAIRYSALE_GATEPASS_DETAIL where GPCode='" & StrGatePassCode & "'"
                    clsDBFuncationality.ExecuteNonQuery(qry)
                    qry = "delete from TSPL_DAIRYSALE_GATEPASS_MASTER where GPCode='" & StrGatePassCode & "'"
                    clsDBFuncationality.ExecuteNonQuery(qry)
                End If
                Dim StrQry As String = "update TSPL_DEMAND_BOOKING_DETAIL set IsGatePassGenerated='N',GPCode='' where document_no='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked, "Morning", "Evening") & "'"
                If clsDBFuncationality.ExecuteNonQuery(StrQry) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Gate Pass Cancel Successfully", Me.Text)
                    btn_Gatepass.Enabled = True
                    btn_GPCancel.Enabled = False
                    'LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub Btn_TSCancel_Click(sender As Object, e As EventArgs) Handles btn_TSCancel.Click
        Try
            If clsCommon.myLen(clsCommon.myCstr(txtDocNo.Value)) > 0 Then
                If rbtnMorningEveningBoth.IsChecked Then
                    common.clsCommon.MyMessageBoxShow(Me, "Please select Morning/Evening Shift", Me.Text)
                    rbtnMorning.Focus()
                    Exit Sub
                End If
                Dim StrQry As String = "update TSPL_DEMAND_BOOKING_DETAIL set IsTruckSheetGenerated='N' where document_no='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_DETAIL.ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'"
                If clsDBFuncationality.ExecuteNonQuery(StrQry) Then
                    common.clsCommon.MyMessageBoxShow(Me, "Truck Sheet Cancel Successfully", Me.Text)
                    btn_TruckSheet.Enabled = True
                    SplitButtonTruckSheet.Enabled = True
                    btn_TSCancel.Enabled = False
                    'LoadData(txtDocNo.Value, NavigatorType.Current)
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    'Private Sub gv1_CellFormatting(sender As Object, e As CellFormattingEventArgs) Handles gv1.CellFormatting
    '    Try
    '        If e.Column.Name = colCustCode Then
    '            e.CellElement.Font = New Font("Arial", 10, FontStyle.Bold)
    '        End If
    '        If e.Column.Index >= 9 AndAlso e.Column.Name <> colCrate And e.Column.Name <> colAmt And e.Column.Name <> colLitre And e.Column.Name <> colMAmt And e.Column.Name <> colPCount And e.Column.Name <> colPCount Then
    '            ' If isLoadData = False Then
    '            If (chkEveningGatepassTruckSheetGenerated.Checked OrElse chkEveningPosted.Checked) And clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colShiftName).Value), "Evening ") = CompairStringResult.Equal Then
    '                gv1.CurrentRow.Cells(e.ColumnIndex).ReadOnly = True
    '            End If
    '            If (chkMorningGatepassTruckSheetGenerated.Checked OrElse chkMorningPosted.Checked) And clsCommon.CompairString(clsCommon.myCstr(gv1.CurrentRow.Cells(colShiftName).Value), "Morning ") = CompairStringResult.Equal Then
    '                gv1.CurrentRow.Cells(e.ColumnIndex).ReadOnly = True
    '            End If
    '            e.CellElement.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    '        End If
    '        If e.Column.Index >= 9 AndAlso gv1.Rows.Count > 0 Then
    '            If e.Column.IsCurrent Then
    '                e.CellElement.DrawFill = True
    '                e.CellElement.BackColor = Color.LightGreen
    '            Else
    '                e.CellElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
    '                e.CellElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
    '            End If
    '            If e.CellElement.RowInfo.IsCurrent Then
    '                e.CellElement.RowElement.BackColor = Color.LightGreen
    '                e.CellElement.RowElement.DrawFill = True
    '            Else
    '                e.CellElement.RowElement.ResetValue(LightVisualElement.DrawFillProperty, ValueResetFlags.Local)
    '                e.CellElement.RowElement.ResetValue(LightVisualElement.BackColorProperty, ValueResetFlags.Local)
    '            End If
    '        End If
    '    Catch ex As Exception
    '    End Try
    'End Sub
    'Private Sub MergeHorizontally(radGridView As RadGridView, startColumnIndex As Integer, endColumnIndex As Integer)
    '    For Each item As GridViewRowInfo In radGridView.Rows
    '        For i As Integer = startColumnIndex To endColumnIndex - 1
    '            Dim firstCell As GridViewCellInfo = item.Cells(i)
    '            Dim secondCell As GridViewCellInfo = item.Cells(i + 1)
    '            Dim firstCellText As String = (If(firstCell IsNot Nothing AndAlso firstCell.Value IsNot Nothing, firstCell.Value.ToString(), String.Empty))
    '            Dim secondCellText As String = (If(secondCell IsNot Nothing AndAlso secondCell.Value IsNot Nothing, secondCell.Value.ToString(), String.Empty))
    '            setCellBorders(firstCell, Color.FromArgb(209, 225, 245))
    '            setCellBorders(secondCell, Color.FromArgb(209, 225, 245))
    '            If firstCellText = secondCellText Then
    '                firstCell.Style.BorderRightColor = Color.Transparent
    '                secondCell.Style.BorderLeftColor = Color.Transparent
    '                secondCell.Style.ForeColor = Color.Transparent
    '            Else
    '                secondCell.Style.ForeColor = Color.Black
    '            End If
    '        Next
    '    Next
    'End Sub
    'Private Sub MergeVertically(radGridView As RadGridView, columnIndexes As Integer())
    '    Dim Prev As GridViewRowInfo = Nothing
    '    For Each item As GridViewRowInfo In radGridView.Rows
    '        If Prev IsNot Nothing Then
    '            Dim firstCellText As String = String.Empty
    '            Dim secondCellText As String = String.Empty
    '            For Each i As Integer In columnIndexes
    '                Dim firstCell As GridViewCellInfo = Prev.Cells(i)
    '                Dim secondCell As GridViewCellInfo = item.Cells(i)
    '                firstCellText = (If(firstCell IsNot Nothing AndAlso firstCell.Value IsNot Nothing, firstCell.Value.ToString(), String.Empty))
    '                secondCellText = (If(secondCell IsNot Nothing AndAlso secondCell.Value IsNot Nothing, secondCell.Value.ToString(), String.Empty))
    '                setCellBorders(firstCell, Color.FromArgb(209, 225, 245))
    '                setCellBorders(secondCell, Color.FromArgb(209, 225, 245)) '117, 230, 218
    '                If rbtnMorningEveningBoth.IsChecked = True Then
    '                    If firstCellText = secondCellText Then
    '                        firstCell.Style.BorderBottomColor = Color.Transparent
    '                        secondCell.Style.BorderTopColor = Color.Transparent
    '                        secondCell.Style.ForeColor = Color.Transparent
    '                    Else
    '                        secondCell.Style.ForeColor = Color.Black
    '                        Prev = item
    '                        Exit For
    '                    End If
    '                Else
    '                    If firstCellText = secondCellText Then
    '                        'firstCell.Style.BorderBottomColor = Color.Black
    '                        'secondCell.Style.BorderTopColor = Color.Black
    '                        secondCell.Style.ForeColor = Color.Black
    '                    Else
    '                        secondCell.Style.ForeColor = Color.Black
    '                        Prev = item
    '                        Exit For
    '                    End If
    '                End If
    '            Next
    '        Else
    '            Prev = item
    '        End If
    '    Next
    'End Sub
    'Private Sub setCellBorders(cell As GridViewCellInfo, color As Color)
    '    cell.Style.CustomizeBorder = True
    '    cell.Style.BorderBoxStyle = Telerik.WinControls.BorderBoxStyle.FourBorders
    '    cell.Style.BorderLeftColor = color
    '    cell.Style.BorderRightColor = color
    '    cell.Style.BorderBottomColor = color
    '    If cell.Style.BorderTopColor <> Color.Transparent Then
    '        cell.Style.BorderTopColor = color
    '    End If
    'End Sub
    'Private Sub gv1_SortChanged(sender As Object, e As GridViewCollectionChangedEventArgs) Handles gv1.SortChanged
    '    MergeVertically(gv1, New Integer() {1, 2})
    'End Sub
    'Private Sub gv1_FilterChanged(sender As Object, e As GridViewCollectionChangedEventArgs) Handles gv1.FilterChanged
    '    MergeVertically(gv1, New Integer() {1, 2})
    '    ''MergeHorizontally(gv1, 2, 3)
    'End Sub
    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Try
            For jj As Integer = 0 To gv1.Rows.Count - 1
                'Dim qry As String = ""
                'Dim strCustomerDesc As String = ""
                'Dim strCustomerCode As String = ""
                ' Dim strCustomer = clsCommon.myCstr(gv1.Rows(jj).Cells(colCustCode).Value)
                Dim strCustomer = clsCommon.myCstr(gv1.Rows(jj).Cells(colCustName).Value).ToLower()
                Dim strCustCode = clsCommon.myCstr(gv1.Rows(jj).Cells(colCustCode).Value).ToLower()
                gv1.ClearSelection()
                If strCustomer.Contains(txtcustomersearch.Text.ToLower()) OrElse strCustCode.Contains(txtcustomersearch.Text.ToLower()) Then
                    gv1.Rows(jj).Cells(colCustCode).IsSelected = True
                    gv1.Rows(jj).IsCurrent = True
                    gv1.Columns(colCustCode).IsCurrent = True
                    gv1.PerformLayout()
                    gv1.Focus()
                    gv1.VerticalScroll.Visible = True
                    Exit Sub
                End If
                'If clsCommon.CompairString(strCustomer, txtcustomersearch.Text) = CompairStringResult.Equal Then
                '    gv1.Rows(jj).Cells(colCustCode).IsSelected = True
                '    gv1.Rows(jj).IsCurrent = True
                '    gv1.Columns(colCustCode).IsCurrent = True
                '    gv1.PerformLayout()
                '    gv1.Focus()
                '    gv1.VerticalScroll.Visible = True
                '    Exit Sub
                'End If
            Next
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub rmi_TS_Excel_Click(sender As Object, e As EventArgs) Handles rmi_TS_Excel.Click
        Try
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                isIndent = False
            Else
                isIndent = True
            End If
            exportExcel()
            isIndent = False
        Catch ex As Exception
            isIndent = False
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub exportExcel()
        Try
            Dim TripNO As String = ""
            Dim qry As String = "select distinct CAST(Trip_No AS VARCHAR(10)) as Code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + txtDocNo.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            dr = dt.NewRow()
            dr("Code") = "ALL"
            'dr("Name") = "ALL"
            dt.Rows.Add(dr)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                TripNO = clsCommon.myCstr(dt.Rows(0)("Code"))
                If dt.Rows.Count > 1 Then
                    Dim frmFC As New FrmFreeComboBox
                    frmFC.ComboSource = dt
                    frmFC.ComboValueMember = "Code"
                    frmFC.ComboDisplayMember = "Code"
                    frmFC.LabelCaption = "Trip No"
                    frmFC.ShowDialog()
                    TripNO = frmFC.strRetValue
                End If
                If clsCommon.myLen(TripNO) > 0 Then
                    'If clsCommon.myLen(txtMCC.Value) > 0 Then
                    '    RefreshMCCCollectionDetail(txtMCC.Value, strMilkType, Nothing)
                    'End If
                    'LoadTransactionData(strMilkType)
                End If
            End If
            TruckSheet(EnumExportTo.Excel, TripNO)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub rmi_TS_PDF_Click(sender As Object, e As EventArgs) Handles rmi_TS_PDF.Click
        Try
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                isIndent = False
            Else
                isIndent = True
            End If
            ExportPDF()
            isIndent = False
        Catch ex As Exception
            isIndent = False
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub ExportPDF()
        Try
            Dim TripNO As String = ""
            Dim qry As String = "select distinct CAST(Trip_No AS VARCHAR(10)) as Code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + txtDocNo.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            dr = dt.NewRow()
            dr("Code") = "ALL"
            'dr("Name") = "ALL"
            dt.Rows.Add(dr)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                TripNO = clsCommon.myCstr(dt.Rows(0)("Code"))
                If dt.Rows.Count > 1 Then
                    Dim frmFC As New FrmFreeComboBox
                    frmFC.ComboSource = dt
                    frmFC.ComboValueMember = "Code"
                    frmFC.ComboDisplayMember = "Code"
                    frmFC.LabelCaption = "Trip No"
                    frmFC.ShowDialog()
                    TripNO = frmFC.strRetValue
                End If
                If clsCommon.myLen(TripNO) > 0 Then
                    'If clsCommon.myLen(txtMCC.Value) > 0 Then
                    '    RefreshMCCCollectionDetail(txtMCC.Value, strMilkType, Nothing)
                    'End If
                    'LoadTransactionData(strMilkType)
                End If
            End If
            TruckSheet(EnumExportTo.PDF, TripNO)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub TruckSheet(ByVal exporter As EnumExportTo, ByVal TripNo As String)
        Try
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select document", Me.Text)
                txtDocNo.Focus()
                Exit Sub
            End If
            If rbtnMorningEveningBoth.IsChecked = True Then
                common.clsCommon.MyMessageBoxShow(Me, "Please select Morning/Evening Shift", Me.Text)
                rbtnMorning.Focus()
                Exit Sub
            End If
            'clsDBFuncationality.ExecuteNonQuery("update TSPL_DEMAND_BOOKING_DETAIL set IsTruckSheetGenerated='Y' where document_no='" & txtDocNo.Value & "' and ShiftType='" & IIf(rbtnMorning.IsChecked = True, "Morning", "Evening") & "'")
            If exporter = EnumExportTo.Excel Then
                TruckSheetExcel(True, TripNo)
            End If
            If exporter = EnumExportTo.PDF Then
                TruckSheetExcel(False, TripNo)
                'TruckSheetPDF()
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                SplitButtonTruckSheet.Enabled = True
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                    If rdbnFreshAmbientBoth.IsChecked Then
                        SplitButtonTruckSheet.Enabled = True
                    Else
                        SplitButtonTruckSheet.Enabled = False
                    End If
                End If
            Else
                SplitButtonTruckSheet.Enabled = False
            End If
            btn_TSCancel.Enabled = True
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    '    Private Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
    '        If clsCommon.myLen(txtDocNo.Value) <= 0 Then
    '            myMessages.blankValue(Me, "Booking not found to Print", Me.Text)
    '        End If
    '        If btnSave.Enabled = True Then
    '            SaveData(0, True)
    '        End If
    '        Dim qry As String = Nothing
    '        Dim SubRptQry As String = Nothing
    '        Dim ShiftType As String = ""
    '        Dim shiftAMPMType As String = ""
    '        Dim PreshiftAMPMType As String = ""
    '        Dim Previous_Shift As String = ""
    '        Dim Previous_Date As String
    '        Dim ItemCount As Double = 0
    '        If rbtnEvening.IsChecked = True Then
    '            ShiftType = "Evening"
    '            shiftAMPMType = "PM"
    '            PreshiftAMPMType = "AM"
    '            Previous_Shift = "Morning"
    '            Previous_Date = clsCommon.myCDate(txtDate.Value).AddDays(1)
    '        Else
    '            ShiftType = "Morning"
    '            Previous_Shift = "Evening"
    '            shiftAMPMType = "AM"
    '            PreshiftAMPMType = "PM"
    '            ' Previous_Date = clsDBFuncationality.getSingleValue("select CONVERT(varchar, DATEADD(DAY, -1, convert(Nvarchar, '" & txtDate.Value & "' ,112)),21) as Previous_Date")
    '            Previous_Date = clsCommon.myCDate(txtDate.Value).AddDays(-1)
    '        End If
    '        Dim Comp_Name As String = clsDBFuncationality.getSingleValue("select Comp_Name from TSPL_COMPANY_MASTER where Comp_Code = '" + objCommonVar.CurrentCompanyCode + "'")
    '        Try
    '            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
    '                Throw New Exception("Please select Booking No")
    '            End If
    '            Dim Posted As String = String.Empty
    '            If clsCommon.myLen(txtDocNo.Value) > 0 Then
    '                Dim strItemQry As String = "select count(distinct TSPL_DEMAND_BOOKING_DETAIL.Item_Code) as items from TSPL_DEMAND_BOOKING_DETAIL
    'left join TSPL_ITEM_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
    'where TSPL_DEMAND_BOOKING_DETAIL.Document_No='" + txtDocNo.Value + "' "
    '                If rbtn_Fresh.IsChecked Then
    '                    strItemQry += " and TSPL_ITEM_MASTER.Is_FreshItem=1 "
    '                ElseIf rbtn_Ambient.IsChecked Then
    '                    strItemQry += " and  TSPL_ITEM_MASTER.Is_Ambient=1 "
    '                End If
    '                ItemCount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strItemQry))
    '                Posted = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when posted=0 then 'Pending' else 'Approved' end from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + clsCommon.myCstr(txtDocNo.Value) + "'"))
    '            Else
    '                Throw New Exception("Demand Not Found!")
    '            End If
    '            If rbtn_Fresh.IsChecked Then
    '                qry = " select XXXFinal.*,TSPL_CUSTOMER_MASTER.Display_Seq from( select xx.*
    ' ,case when xx.SNO=1 then ( case when xx.ShiftType='Morning' then (isnull((select sum(ItemNetAmount) as netamt  from TSPL_DEMAND_BOOKING_MASTER left join  TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No where  TSPL_DEMAND_BOOKING_DETAIL.ShiftType = '" + ShiftType + "'  and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= '" + clsCommon.GetPrintDate(txtDate.Value) + "')  and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" + clsCommon.myCstr(txtRouteNo.Value) + "'  and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=xx.Cust_Code ),0 ) + isnull((xx.PrevItemNetAmount),0) ) else 0 end) else 0 end as AmountBE,
    ' case when xx.SNO=1 then xx.Crate_Collect else 0 end as TotalCollectCrate,case when xx.SNO=1 then (case when xx.ShiftType='Morning' then (isnull(tcs.TCSAmount,0)+isnull(prevtcs.pTCSAmount,0)) else 0 end) else 0 end as TotalTCSAmt,'" + Posted + "' as DocStatus
    '  from ( select XXFinal.Cust_Code as Cust_Code, max(XXFinal.ShiftType) as ShiftType, XXFinal.Sku_Seq as Sku_Seq ,max(XXFinal.Document_Date) as Document_Date, max(XXFinal.Short_Description) as Short_Description, sum(XXFinal.Qty) as Qty, max(XXFinal.Unit_code) as Unit_code, sum(XXFinal.Crate) as Crate, max(XXFinal.Pouch) as Pouch, sum(XXFinal.ItemNetAmount) as ItemNetAmount,max(XXFinal.Route_No) as Route_No, max(XXFinal.Route_Desc) as Route_Desc,max(XXFinal.PrevCrate) as Crate_Collect, max(XXFinal.CompanyName) as CompanyName, max(XXFinal.TranspoterName) as TranspoterName, max(XXFinal.DriverName) as DriverName,max(XXFinal.Vehicle_No) as Vehicle_No, max(XXFinal.Item_Rate) as Item_Rate, max(XXFinal.CFForLTR) as CFForLTR, max(XXFinal.Conversion_Factor) as Conversion_Factor, sum(XXFinal.QTYLtr) as QTYLtr,max(XXFinal.PrevItemNetAmount) as PrevItemNetAmount,ROW_NUMBER() over (Partition by Cust_Code order by Cust_Code) as SNO
    'from
    '(
    'select 
    '  TSPL_DEMAND_BOOKING_DETAIL.Cust_Code, 
    '  TSPL_DEMAND_BOOKING_DETAIL.ShiftType, 
    '  TSPL_ITEM_MASTER.Sku_Seq, 
    '  TSPL_DEMAND_BOOKING_MASTER.Document_Date, 
    '  TSPL_ITEM_MASTER.Short_Description, 
    '  TSPL_DEMAND_BOOKING_DETAIL.Qty as Qty, 
    '  0 as PrevQty,
    '  TSPL_DEMAND_BOOKING_DETAIL.Unit_code, 
    '  Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise Else 0 End As Crate, 
    '  Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then 0 Else 0 End As PrevCrate, 
    '  Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then TSPL_DEMAND_BOOKING_DETAIL.Qty Else 0 End As Pouch, 
    '    Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then 0 Else 0 End As PrevPouch,
    '  TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount, 
    '    0 as PrevItemNetAmount,
    '  TSPL_DEMAND_BOOKING_MASTER.Route_No, 
    '  TSPL_ROUTE_MASTER.Route_Desc, 
    '  Isnull(
    '    TSPL_COMPANY_MASTER.Comp_Name, 'Jaipur Zila Dugdh Utpadak Sahakari Sangh Ltd.'
    '  ) as CompanyName,
    '  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
    '  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No, 
    '  TSPL_DEMAND_BOOKING_DETAIL.Item_Rate, 
    '  ITEMDETAIL.CFForLTR, 
    '  TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 
    '  Convert(
    '    decimal(18, 2), 
    '    (
    '      TSPL_DEMAND_BOOKING_DETAIL.Qty * TSPL_ITEM_UOM_DETAIL.Conversion_Factor
    '    )/ ITEMDETAIL.CFForLTR
    '  ) As QTYLtr
    'from 
    '  TSPL_DEMAND_BOOKING_DETAIL 
    '  Left join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No 
    '  Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
    '  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
    '  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
    '  Left Join (
    '    select 
    '      Conversion_factor AS CFForLTR, 
    '      TSPL_ITEM_UOM_DETAIL.Item_code 
    '    from 
    '      TSPL_ITEM_UOM_DETAIL 
    '    where 
    '      UOM_code = 'LTR'
    '  ) as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
    '  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
    '  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
    '  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
    '  Left Join TSPL_COMPANY_MASTER on TSPL_DEMAND_BOOKING_MASTER.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code
    'where 
    '  TSPL_DEMAND_BOOKING_DETAIL.ShiftType = '" + ShiftType + "' 
    '  and (
    '    CONVERT(
    '      date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 
    '      103
    '    )= '" + clsCommon.GetPrintDate(txtDate.Value) + "'
    '  ) 
    '  and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" + clsCommon.myCstr(txtRouteNo.Value) + "' "
    '                If chkIndividualCustomer.Checked Then
    '                    If clsCommon.myLen(txtCustomerNo.Value) > 0 Then
    '                        qry += " and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code='" + txtCustomerNo.Value + "'"
    '                    End If
    '                Else
    '                    qry += " and TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer=0"
    '                End If
    '                qry += "  union all
    '  select 
    '  TSPL_DEMAND_BOOKING_DETAIL.Cust_Code, 
    '  '" + ShiftType + "'  as ShiftType, 
    '  TSPL_ITEM_MASTER.Sku_Seq, 
    '  '" + clsCommon.GetPrintDate(txtDate.Value) + "' as Document_Date, 
    '  TSPL_ITEM_MASTER.Short_Description, 
    '  0 as Qty, 
    '  TabCustWiseCrate.Qty as PrevQty,
    '  TSPL_DEMAND_BOOKING_DETAIL.Unit_code, 
    '  Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then 0 Else 0 End As Crate, 
    '  Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then TabCustWiseCrate.TotalCrates_ItemWise Else 0 End As PrevCrate, 
    '  Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then 0 Else 0 End As Pouch, 
    '    Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then TabCustWiseCrate.qty Else 0 End As PrevPouch,
    '  0 as ItemNetAmount, 
    '    NetAmount as PrevItemNetAmount,
    '  TSPL_DEMAND_BOOKING_MASTER.Route_No, 
    '  TSPL_ROUTE_MASTER.Route_Desc, 
    '  Isnull(
    '    TSPL_COMPANY_MASTER.Comp_Name, 'Jaipur Zila Dugdh Utpadak Sahakari Sangh Ltd.'
    '  ) as CompanyName,
    '  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
    '  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No, 
    '  TSPL_DEMAND_BOOKING_DETAIL.Item_Rate, 
    '  ITEMDETAIL.CFForLTR, 
    '  TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 
    '  0 As QTYLtr
    'from 
    '  (
    '    select 
    '      ROW_NUMBER() over (
    '        PARTITION BY xx.Cust_Code 
    '        order by 
    '          xx.Cust_Code, 
    '          xx.ORDCol desc
    '      ) as SNO, 
    '      xx.Cust_Code, 
    '      xx.ORDCol, 
    '      sum(xx.TotalCrates_ItemWise) as TotalCrates_ItemWise, 
    '      sum(xx.TotalLtr_ItemWise) as TotalLtr, 
    '      sum(xx.ItemNetAmount) as NetAmount, 
    '	  sum(xx.qty) as Qty
    '    from 
    '      (
    '        select 
    '          innBD.Cust_Code, 
    '          convert(
    '            varchar, InnBM.Document_Date, 102
    '          )+ case when innBD.ShiftType = 'Evening' then 'B' else 'A' end as ORDCol, 
    '          innBD.TotalCrates_ItemWise, 
    '          innBD.TotalLtr_ItemWise, 
    '          innBD.ItemNetAmount,innBD.qty
    '        from 
    '          TSPL_DEMAND_BOOKING_MASTER as InnBM 
    '          left outer join TSPL_DEMAND_BOOKING_DETAIL innBD on innBD.Document_No = InnBM.Document_No 
    '        where 
    '          2 = 2  "
    '                If rbtnMorning.IsChecked Then
    '                    qry += " and innBD.ShiftType='" + Previous_Shift + "' and ( CONVERT(date, InnBM.Document_Date, 103)= '" + clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)) + "') "
    '                ElseIf rbtnEvening.IsChecked Then
    '                    qry += " and innBD.ShiftType='" + Previous_Shift + "' and CONVERT(date, InnBM.Document_Date,103)='" + clsCommon.GetPrintDate(txtDate.Value) + "'" ' or CONVERT(date, InnBM.Document_Date,103)<'" + clsCommon.GetPrintDate(txtDate.Value) + "') "
    '                End If
    '                qry += " and innBD.Cust_Code is not null ) xx  
    '    group by 
    '      xx.Cust_Code, 
    '      xx.ORDCol
    '  )  TabCustWiseCrate 
    '    left join TSPL_Demand_Booking_Detail on TabCustWiseCrate.cust_Code=TSPL_Demand_Booking_Detail.cust_Code and TabCustWiseCrate.SNO=1
    '  Left join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No 
    '  Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
    '  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
    '  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
    '  Left Join (
    '    select 
    '      Conversion_factor AS CFForLTR, 
    '      TSPL_ITEM_UOM_DETAIL.Item_code 
    '    from 
    '      TSPL_ITEM_UOM_DETAIL 
    '    where 
    '      UOM_code = 'LTR'
    '  ) as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
    '  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
    '  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
    '  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
    '  Left Join TSPL_COMPANY_MASTER on TSPL_DEMAND_BOOKING_MASTER.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code 
    '  where TSPL_ROUTE_MASTER.Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "'"
    '                If chkIndividualCustomer.Checked Then
    '                    If clsCommon.myLen(txtCustomerNo.Value) > 0 Then
    '                        qry += " and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code='" + txtCustomerNo.Value + "'"
    '                    End If
    '                End If
    '                qry += " )XXFinal
    '  where XXFinal.Cust_Code in (select distinct TSPL_DEMAND_BOOKING_DETAIL.Cust_Code from TSPL_DEMAND_BOOKING_MASTER left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No where TSPL_DEMAND_BOOKING_MASTER.Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + " ' and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code is not null )
    '            Group by XXFinal.Cust_Code,XXFinal.Sku_Seq )xx 
    '   left join (
    ' select 
    ' sum(XYZ.TCSAmount) as TCSAmount,XYZ.Cust_Code,max(XYZ.Against_DemandBooking_No) as Against_DemandBooking_No
    '	  from(
    ' select TSPL_BOOKING_MATSER.TCSAmount,TSPL_BOOKING_MATSER.Against_DemandBooking_No,TSPL_BOOKING_DETAIL.Cust_Code from TSPL_BOOKING_MATSER
    'left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
    'where TSPL_BOOKING_MATSER.GatePass_Type='" + shiftAMPMType + "' and  (CONVERT(date, TSPL_BOOKING_MATSER.Document_Date,103)= '" + clsCommon.GetPrintDate(txtDate.Value) + "') 
    'group by TSPL_BOOKING_DETAIL.Cust_Code,TSPL_BOOKING_MATSER.TCSAmount,TSPL_BOOKING_MATSER.Against_DemandBooking_No) XYZ
    'group by XYZ.Cust_Code
    '  ) as tcs on xx.Cust_Code=tcs.Cust_Code
    'left join (
    '  select 
    ' sum(XYZ.pTCSAmount) as pTCSAmount,XYZ.Cust_Code,max(XYZ.Against_DemandBooking_No) as Against_DemandBooking_No
    '	  from(select (TSPL_BOOKING_MATSER.TCSAmount) as pTCSAmount ,(TSPL_BOOKING_MATSER.Against_DemandBooking_No) Against_DemandBooking_No,TSPL_BOOKING_DETAIL.Cust_Code
    '		  from TSPL_BOOKING_MATSER
    'left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
    'where TSPL_BOOKING_MATSER.GatePass_Type='" + PreshiftAMPMType + "' and  (CONVERT(date, TSPL_BOOKING_MATSER.Document_Date,103)= '" + IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) + "') 
    'group by TSPL_BOOKING_DETAIL.Cust_Code,TSPL_BOOKING_MATSER.TCSAmount,TSPL_BOOKING_MATSER.Against_DemandBooking_No
    ') XYZ
    'group by XYZ.Cust_Code
    ') as prevtcs on xx.Cust_Code=prevtcs.Cust_Code  )XXXFinal
    'left join TSPL_CUSTOMER_MASTER on XXXFinal.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
    'order  by TSPL_CUSTOMER_MASTER.Display_Seq
    '"
    '            ElseIf rbtn_Ambient.IsChecked Then
    '                qry = "select XXXFinal.* from(  select XXFinal.Cust_Code as Cust_Code, max(XXFinal.ShiftType) as ShiftType, XXFinal.Sku_Seq as Sku_Seq ,max(XXFinal.Document_Date) as Document_Date, max(XXFinal.Short_Description)+' '+ max(XXFinal.Unit_code) as Short_Description, sum(XXFinal.Qty) as Qty, max(XXFinal.Unit_code) as Unit_code, sum(XXFinal.ItemNetAmount) as ItemNetAmount,max(XXFinal.Route_No) as Route_No, max(XXFinal.Route_Desc) as Route_Desc, max(XXFinal.CompanyName) as CompanyName, max(XXFinal.TranspoterName) as TranspoterName, max(XXFinal.DriverName) as DriverName,max(XXFinal.Vehicle_No) as Vehicle_No, max(XXFinal.Item_Rate) as Item_Rate, max(XXFinal.CFForKG) as CFForKG, max(XXFinal.Conversion_Factor) as Conversion_Factor,'" + Posted + "' as DocStatus
    'from
    '( select 
    '  TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,  TSPL_DEMAND_BOOKING_DETAIL.ShiftType,  TSPL_ITEM_MASTER.Sku_Seq,TSPL_DEMAND_BOOKING_MASTER.Document_Date,TSPL_ITEM_MASTER.Short_Description,TSPL_DEMAND_BOOKING_DETAIL.Qty as Qty, 
    'TSPL_DEMAND_BOOKING_DETAIL.Unit_code,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,Isnull(TSPL_COMPANY_MASTER.Comp_Name, 'Jaipur Zila Dugdh Utpadak Sahakari Sangh Ltd.') as CompanyName,TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No, 
    '  TSPL_DEMAND_BOOKING_DETAIL.Item_Rate,ITEMDETAIL.CFForKG,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,Convert(decimal(18, 2),(TSPL_DEMAND_BOOKING_DETAIL.Qty * TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ ITEMDETAIL.CFForKG) As QTYKg
    'from 
    '  TSPL_DEMAND_BOOKING_DETAIL 
    '  Left join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No 
    '  Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
    '  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
    '  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
    '  Left Join (
    '    select 
    '      Conversion_factor AS CFForKG, 
    '      TSPL_ITEM_UOM_DETAIL.Item_code 
    '    from 
    '      TSPL_ITEM_UOM_DETAIL 
    '    where 
    '      UOM_code = 'KG'
    '  ) as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
    '  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
    '  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
    '  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
    '  Left Join TSPL_COMPANY_MASTER on TSPL_DEMAND_BOOKING_MASTER.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code
    'where 2=2 "
    '                If rbtnMorning.IsChecked Then
    '                    qry += " and TSPL_DEMAND_BOOKING_DETAIL.ShiftType = 'Morning' "
    '                ElseIf rbtnEvening.IsChecked Then
    '                    qry += " and TSPL_DEMAND_BOOKING_DETAIL.ShiftType = 'Evening' "
    '                End If
    '                qry += " and (CONVERT(      date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= '" + clsCommon.GetPrintDate(txtDate.Value) + "') 
    '  and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" + txtRouteNo.Value + "' and TSPL_ITEM_MASTER.Is_Ambient=1  and TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer=0 )XXFinal
    '  Group by XXFinal.Cust_Code,XXFinal.Sku_Seq ,XXFinal.Short_Description )XXXFinal
    'left join TSPL_CUSTOMER_MASTER on XXXFinal.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
    'order  by TSPL_CUSTOMER_MASTER.Display_Seq "
    '            End If
    '            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '            Dim frmCRV As New frmCrystalReportViewer()
    '            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
    '                If rbtn_Fresh.IsChecked Then
    '                    If ItemCount > 0 AndAlso ItemCount <= 9 Then
    '                        frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingFUDP", "Demand Booking")
    '                    Else
    '                        frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingFUDP9", "Demand Booking")
    '                    End If
    '                ElseIf rbtn_Ambient.IsChecked Then
    '                    If ItemCount > 0 AndAlso ItemCount <= 13 Then
    '                        frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingAUDP", "Demand Booking")
    '                    Else
    '                        frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingAUDP14", "Demand Booking")
    '                    End If
    '                End If
    '            Else
    '                If rbtn_Fresh.IsChecked Then
    '                    frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBooking", "Demand Booking")
    '                ElseIf rbtn_Ambient.IsChecked Then
    '                    frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingAUDP", "Demand Booking")
    '                End If
    '            End If
    '            'frmCRV.funreport(CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBooking", "Demand Booking")
    '            'frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, dt2, "rptDemandBooking", "Demand Booking", "rptSubDemandBooking")
    '            frmCRV = Nothing
    '        Catch ex As Exception
    '            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '        End Try
    '    End Sub

    Private Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Try
            Dim qry As String = Nothing
            'Dim SubRptQry As String = Nothing
            Dim ShiftType As String = ""
            Dim shiftAMPMType As String = ""
            Dim PreshiftAMPMType As String = ""
            Dim Previous_Shift As String = ""
            Dim Previous_Date As String
            Dim ItemCount As Double = 0
            If rbtnEvening.IsChecked Then
                ShiftType = "Evening"
                shiftAMPMType = "PM"
                PreshiftAMPMType = "AM"
                Previous_Shift = "Morning"
                Previous_Date = clsCommon.myCDate(txtDate.Value).AddDays(1)
            Else
                ShiftType = "Morning"
                Previous_Shift = "Evening"
                shiftAMPMType = "AM"
                PreshiftAMPMType = "PM"
                ' Previous_Date = clsDBFuncationality.getSingleValue("select CONVERT(varchar, DATEADD(DAY, -1, convert(Nvarchar, '" & txtDate.Value & "' ,112)),21) as Previous_Date")
                Previous_Date = clsCommon.myCDate(txtDate.Value).AddDays(-1)
            End If
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Booking not found to Print")
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                If clsCommon.myLen(txtRouteNo.Value) <= 0 Then
                    Throw New Exception("Route not found to Print")
                End If
                Dim arrRoute As New ArrayList
                arrRoute.Add(txtRouteNo.Value)
                clsDemandBookingSale.PrintDOSData(arrRoute, ShiftType, txtDate.Value, rbtn_Fresh.IsChecked, rbtn_Ambient.IsChecked, chkIndividualCustomer.Checked, 107, 48, DosPaperSize.A4, PageSetup.Landscap, False, isDepartmentRouteSetting) ''
            Else
                'Dim Comp_Name As String = clsDBFuncationality.getSingleValue("select Comp_Name from TSPL_COMPANY_MASTER where Comp_Code = '" + objCommonVar.CurrentCompanyCode + "'")
                Try
                    If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                        Throw New Exception("Please select Booking No")
                    End If
                    Dim Posted As String = String.Empty
                    If clsCommon.myLen(txtDocNo.Value) > 0 Then
                        Dim strItemQry As String = "select count(distinct TSPL_DEMAND_BOOKING_DETAIL.Item_Code) as items from TSPL_DEMAND_BOOKING_DETAIL
left join TSPL_ITEM_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
where TSPL_DEMAND_BOOKING_DETAIL.Document_No='" & txtDocNo.Value & "' "
                        If rbtn_Fresh.IsChecked Then
                            strItemQry += " and TSPL_ITEM_MASTER.Is_FreshItem=1 "
                        ElseIf rbtn_Ambient.IsChecked Then
                            strItemQry += " and  TSPL_ITEM_MASTER.Is_Ambient=1 "
                        End If
                        ItemCount = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(strItemQry))
                        Posted = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when posted=0 then 'Pending' else 'Approved' end from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + clsCommon.myCstr(txtDocNo.Value) + "'"))
                    Else
                        Throw New Exception("Demand Not Found!")
                    End If
                    If rbtn_Fresh.IsChecked Then
                        qry = " select XXXFinal.*,TSPL_CUSTOMER_MASTER.Display_Seq,TSPL_CUSTOMER_MASTER.Credit_Customer from( 
select xx.* ,case when xx.SNO=1 then ( case when xx.ShiftType='Morning' then (isnull((select sum(ItemNetAmount) as netamt  from TSPL_DEMAND_BOOKING_MASTER left join  TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No where  TSPL_DEMAND_BOOKING_DETAIL.ShiftType = '" + ShiftType + "'  and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= '" + clsCommon.GetPrintDate(txtDate.Value) + "')  and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" + clsCommon.myCstr(txtRouteNo.Value) + "'  and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=xx.Cust_Code ),0 ) + isnull((xx.PrevItemNetAmount),0) ) else 0 end) else 0 end as AmountBE,
case when xx.SNO=1 then xx.Crate_Collect else 0 end as TotalCollectCrate,case when xx.SNO=1 then (case when xx.ShiftType='Morning' then (isnull(tcs.TCSAmount,0)+isnull(prevtcs.pTCSAmount,0)) else 0 end) else 0 end as TotalTCSAmt,'" + Posted + "' as DocStatus from ( 
select XXFinal.Cust_Code as Cust_Code, max(XXFinal.ShiftType) as ShiftType, XXFinal.Sku_Seq as Sku_Seq ,max(XXFinal.Document_Date) as Document_Date, max(XXFinal.Short_Description) as Short_Description, sum(XXFinal.Qty) as Qty, max(XXFinal.Unit_code) as Unit_code, sum(XXFinal.Crate) as Crate, max(XXFinal.Pouch) as Pouch, sum(XXFinal.ItemNetAmount) as ItemNetAmount,max(XXFinal.Route_No) as Route_No, max(XXFinal.Route_Desc) as Route_Desc,max(XXFinal.PrevCrate) as Crate_Collect, max(XXFinal.CompanyName) as CompanyName, max(XXFinal.TranspoterName) as TranspoterName, max(XXFinal.DriverName) as DriverName,max(XXFinal.Vehicle_No) as Vehicle_No, max(XXFinal.Item_Rate) as Item_Rate, max(XXFinal.CFForLTR) as CFForLTR, max(XXFinal.Conversion_Factor) as Conversion_Factor, sum(XXFinal.QTYLtr) as QTYLtr,max(XXFinal.PrevItemNetAmount) as PrevItemNetAmount,ROW_NUMBER() over (Partition by Cust_Code order by Cust_Code) as SNO
from (
select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,TSPL_ITEM_MASTER.Sku_Seq,TSPL_DEMAND_BOOKING_MASTER.Document_Date,
TSPL_ITEM_MASTER.Short_Description,TSPL_DEMAND_BOOKING_DETAIL.Qty as Qty,0 as PrevQty,TSPL_DEMAND_BOOKING_DETAIL.Unit_code, 
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then TSPL_DEMAND_BOOKING_DETAIL.TotalCrates_ItemWise Else 0 End As Crate, 
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then 0 Else 0 End As PrevCrate, 
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then TSPL_DEMAND_BOOKING_DETAIL.Qty Else 0 End As Pouch, 
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then 0 Else 0 End As PrevPouch,
TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount, 0 as PrevItemNetAmount, TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,Isnull(TSPL_COMPANY_MASTER.Comp_Name, 'Jaipur Zila Dugdh Utpadak Sahakari Sangh Ltd.') as CompanyName,
TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName,TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No,TSPL_DEMAND_BOOKING_DETAIL.Item_Rate,ITEMDETAIL.CFForLTR,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,
Convert(decimal(18, 2),(TSPL_DEMAND_BOOKING_DETAIL.Qty * TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ ITEMDETAIL.CFForLTR) As QTYLtr from
TSPL_DEMAND_BOOKING_DETAIL
Left join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No
Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
Left Join ( select  Conversion_factor AS CFForLTR, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'LTR' ) as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
Left Join TSPL_COMPANY_MASTER on TSPL_DEMAND_BOOKING_MASTER.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code
where TSPL_DEMAND_BOOKING_DETAIL.ShiftType = '" & ShiftType & "' and ( CONVERT(date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= '" & clsCommon.GetPrintDate(txtDate.Value) & "') 
and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" & clsCommon.myCstr(txtRouteNo.Value) & "' "
                        If chkIndividualCustomer.Checked Then
                            If clsCommon.myLen(txtCustomerNo.Value) > 0 Then
                                qry += " and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code='" & txtCustomerNo.Value & "'"
                            End If
                        Else
                            qry += " and TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer=0"
                        End If
                        qry += "  union all
select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code, '" & ShiftType & "'  as ShiftType, TSPL_ITEM_MASTER.Sku_Seq, '" & clsCommon.GetPrintDate(txtDate.Value) & "' as Document_Date,
TSPL_ITEM_MASTER.Short_Description,0 as Qty,TabCustWiseCrate.Qty as PrevQty,TSPL_DEMAND_BOOKING_DETAIL.Unit_code,
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then 0 Else 0 End As Crate,
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then TabCustWiseCrate.TotalCrates_ItemWise Else 0 End As PrevCrate,
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then 0 Else 0 End As Pouch,Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then TabCustWiseCrate.qty Else 0 End As PrevPouch,
0 as ItemNetAmount,NetAmount as PrevItemNetAmount,TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc, Isnull(TSPL_COMPANY_MASTER.Comp_Name, 'Jaipur Zila Dugdh Utpadak Sahakari Sangh Ltd.') as CompanyName,
TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName,TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No,TSPL_DEMAND_BOOKING_DETAIL.Item_Rate, 
ITEMDETAIL.CFForLTR,TSPL_ITEM_UOM_DETAIL.Conversion_Factor, 0 As QTYLtr from 
( select  ROW_NUMBER() over ( PARTITION BY xx.Cust_Code  order by  xx.Cust_Code, xx.ORDCol desc) as SNO,xx.Cust_Code,xx.ORDCol, 
sum(xx.TotalCrates_ItemWise) as TotalCrates_ItemWise,sum(xx.TotalLtr_ItemWise) as TotalLtr,sum(xx.ItemNetAmount) as NetAmount,sum(xx.qty) as Qty
from  ( select  innBD.Cust_Code, convert( varchar, InnBM.Document_Date, 102 )+ case when innBD.ShiftType = 'Evening' then 'B' else 'A' end as ORDCol, 
innBD.TotalCrates_ItemWise,innBD.TotalLtr_ItemWise, innBD.ItemNetAmount,innBD.qty from 
TSPL_DEMAND_BOOKING_MASTER as InnBM  
left outer join TSPL_DEMAND_BOOKING_DETAIL innBD on innBD.Document_No = InnBM.Document_No 
where  2 = 2  "
                        If rbtnMorning.IsChecked Then
                            qry += " and innBD.ShiftType='" & Previous_Shift & "' and ( CONVERT(date, InnBM.Document_Date, 103)= '" & clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)) & "') "
                        ElseIf rbtnEvening.IsChecked Then
                            qry += " and innBD.ShiftType='" & Previous_Shift & "' and CONVERT(date, InnBM.Document_Date,103)='" & clsCommon.GetPrintDate(txtDate.Value) & "'" ' or CONVERT(date, InnBM.Document_Date,103)<'" & clsCommon.GetPrintDate(txtDate.Value) & "') "
                        End If
                        qry += " and innBD.Cust_Code is not null ) xx  
group by xx.Cust_Code, xx.ORDCol )  TabCustWiseCrate 
left join TSPL_Demand_Booking_Detail on TabCustWiseCrate.cust_Code=TSPL_Demand_Booking_Detail.cust_Code and TabCustWiseCrate.SNO=1
Left join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No 
Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
Left Join ( select  Conversion_factor AS CFForLTR, TSPL_ITEM_UOM_DETAIL.Item_code  from TSPL_ITEM_UOM_DETAIL where  UOM_code = 'LTR' ) as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
Left Join TSPL_COMPANY_MASTER on TSPL_DEMAND_BOOKING_MASTER.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code 
where TSPL_ROUTE_MASTER.Route_No='" & clsCommon.myCstr(txtRouteNo.Value) & "'"
                        If chkIndividualCustomer.Checked Then
                            If clsCommon.myLen(txtCustomerNo.Value) > 0 Then
                                qry += " and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code='" & txtCustomerNo.Value & "'"
                            End If
                        End If

                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                            qry += " union all 
                      select  TSPL_CUSTOMER_MASTER.Cust_Code,'" & ShiftType & "'  as ShiftType, TSPL_ITEM_MASTER.Sku_Seq,'" & clsCommon.GetPrintDate(txtDate.Value) & "' as Document_Date, TSPL_ITEM_MASTER.Short_Description,0 as Qty,TSPL_DEMAND_BOOKING_DETAIL.Qty as PrevQty,TSPL_DEMAND_BOOKING_DETAIL.Unit_code,  0 As Crate,Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Crate' Then TSPL_Demand_Booking_Detail.TotalCrates_ItemWise Else 0 End As PrevCrate, 0 As Pouch, 
Case When TSPL_DEMAND_BOOKING_DETAIL.Unit_Code = 'Pouch' Then TSPL_Demand_Booking_Detail.Qty Else 0 End As PrevPouch, 0 as ItemNetAmount,TSPL_Demand_Booking_Detail.ItemNetAmount as PrevItemNetAmount, TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,TSPL_COMPANY_MASTER.Comp_Name as CompanyName,TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName,TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No, 
TSPL_DEMAND_BOOKING_DETAIL.Item_Rate,ITEMDETAIL.CFForLTR,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,0 As QTYLtr from  TSPL_CUSTOMER_MASTER  Left Outer Join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No  Left Join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
LEFT Outer Join (Select Document_No,Document_Date,Route_No,Max(ShiftType)ShiftType from TSPL_DEMAND_BOOKING_MASTER Where CONVERT(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='15/Jul/2024' Group By Document_No,Document_Date,Route_No )TSPL_DEMAND_BOOKING_MASTER On TSPL_DEMAND_BOOKING_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No  Left Outer Join TSPL_Demand_Booking_Detail On TSPL_Demand_Booking_Detail.Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No And TSPL_Demand_Booking_Detail.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code   And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  Left Join (select Conversion_factor AS CFForLTR, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'LTR') as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP' 
where 2=2  And TSPL_CUSTOMER_MASTER.Status='N'  and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" & clsCommon.myCstr(txtRouteNo.Value) & "' and TSPL_DEMAND_BOOKING_MASTER.ShiftType='" & ShiftType & "' and CONVERT(date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= '" & clsCommon.GetPrintDate(txtDate.Value) & "'  and TSPL_CUSTOMER_MASTER.Cust_Code Not In (Select Cust_Code from TSPL_DEMAND_BOOKING_DETAIL Where Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No) "
                        End If

                        qry += " )XXFinal
where XXFinal.Cust_Code in (select distinct TSPL_DEMAND_BOOKING_DETAIL.Cust_Code from TSPL_DEMAND_BOOKING_MASTER left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No where TSPL_DEMAND_BOOKING_MASTER.Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code is not null "
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "AJM") = CompairStringResult.Equal Then
                            qry += " union all  select distinct TSPL_CUSTOMER_MASTER.Cust_Code from TSPL_CUSTOMER_MASTER left outer join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_DETAIL.Cust_Code = TSPL_CUSTOMER_MASTER.Cust_Code where TSPL_CUSTOMER_MASTER.Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and TSPL_DEMAND_BOOKING_DETAIL.Cust_Code is  null"
                        End If
                        qry += ") Group by XXFinal.Cust_Code,XXFinal.Sku_Seq )xx 
left join ( select  sum(XYZ.TCSAmount) as TCSAmount,XYZ.Cust_Code,max(XYZ.Against_DemandBooking_No) as Against_DemandBooking_No from (
select TSPL_BOOKING_MATSER.TCSAmount,TSPL_BOOKING_MATSER.Against_DemandBooking_No,TSPL_BOOKING_DETAIL.Cust_Code from TSPL_BOOKING_MATSER
left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
where TSPL_BOOKING_MATSER.GatePass_Type='" & shiftAMPMType & "' and  (CONVERT(date, TSPL_BOOKING_MATSER.Document_Date,103)= '" & clsCommon.GetPrintDate(txtDate.Value) & "') 
group by TSPL_BOOKING_DETAIL.Cust_Code,TSPL_BOOKING_MATSER.TCSAmount,TSPL_BOOKING_MATSER.Against_DemandBooking_No) XYZ
group by XYZ.Cust_Code ) as tcs on xx.Cust_Code=tcs.Cust_Code 
left join (select sum(XYZ.pTCSAmount) as pTCSAmount,XYZ.Cust_Code,max(XYZ.Against_DemandBooking_No) as Against_DemandBooking_No
from(select (TSPL_BOOKING_MATSER.TCSAmount) as pTCSAmount ,(TSPL_BOOKING_MATSER.Against_DemandBooking_No) Against_DemandBooking_No,TSPL_BOOKING_DETAIL.Cust_Code
from TSPL_BOOKING_MATSER
left join TSPL_BOOKING_DETAIL on TSPL_BOOKING_MATSER.Document_No=TSPL_BOOKING_DETAIL.Document_No
where TSPL_BOOKING_MATSER.GatePass_Type='" & PreshiftAMPMType & "' and  (CONVERT(date, TSPL_BOOKING_MATSER.Document_Date,103)= '" & IIf(rbtnMorning.IsChecked, clsCommon.GetPrintDate(txtDate.Value.AddDays(-1)), clsCommon.GetPrintDate(txtDate.Value)) & "') 
group by TSPL_BOOKING_DETAIL.Cust_Code,TSPL_BOOKING_MATSER.TCSAmount,TSPL_BOOKING_MATSER.Against_DemandBooking_No
) XYZ
group by XYZ.Cust_Code
) as prevtcs on xx.Cust_Code=prevtcs.Cust_Code  )XXXFinal
left join TSPL_CUSTOMER_MASTER on XXXFinal.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code"
                    ElseIf rbtn_Ambient.IsChecked Then
                        qry = "select XXXFinal.*,TSPL_CUSTOMER_MASTER.Display_Seq,TSPL_CUSTOMER_MASTER.Credit_Customer  from(  select XXFinal.Cust_Code as Cust_Code, max(XXFinal.ShiftType) as ShiftType, XXFinal.Sku_Seq as Sku_Seq ,max(XXFinal.Document_Date) as Document_Date, max(XXFinal.Short_Description)+' '+ max(XXFinal.Unit_code) as Short_Description, sum(XXFinal.Qty) as Qty, max(XXFinal.Unit_code) as Unit_code, sum(XXFinal.ItemNetAmount) as ItemNetAmount,max(XXFinal.Route_No) as Route_No, max(XXFinal.Route_Desc) as Route_Desc, max(XXFinal.CompanyName) as CompanyName, max(XXFinal.TranspoterName) as TranspoterName, max(XXFinal.DriverName) as DriverName,max(XXFinal.Vehicle_No) as Vehicle_No, max(XXFinal.Item_Rate) as Item_Rate, max(XXFinal.CFForKG) as CFForKG, max(XXFinal.Conversion_Factor) as Conversion_Factor,'" + Posted + "' as DocStatus from (
selectTSPL_DEMAND_BOOKING_DETAIL.Cust_Code,  TSPL_DEMAND_BOOKING_DETAIL.ShiftType,  TSPL_ITEM_MASTER.Sku_Seq,TSPL_DEMAND_BOOKING_MASTER.Document_Date,TSPL_ITEM_MASTER.Short_Description,TSPL_DEMAND_BOOKING_DETAIL.Qty as Qty, 
TSPL_DEMAND_BOOKING_DETAIL.Unit_code,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_MASTER.Route_No,TSPL_ROUTE_MASTER.Route_Desc,Isnull(TSPL_COMPANY_MASTER.Comp_Name, 'Jaipur Zila Dugdh Utpadak Sahakari Sangh Ltd.') as CompanyName,TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No, 
TSPL_DEMAND_BOOKING_DETAIL.Item_Rate,ITEMDETAIL.CFForKG,TSPL_ITEM_UOM_DETAIL.Conversion_Factor,Convert(decimal(18, 2),(TSPL_DEMAND_BOOKING_DETAIL.Qty * TSPL_ITEM_UOM_DETAIL.Conversion_Factor)/ ITEMDETAIL.CFForKG) As QTYKg
from TSPL_DEMAND_BOOKING_DETAIL 
Left join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No = TSPL_DEMAND_BOOKING_DETAIL.Document_No 
Left join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code = TSPL_DEMAND_BOOKING_DETAIL.Item_Code 
Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
Left Join ( select  Conversion_factor AS CFForKG,TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'KG') as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
Left Join TSPL_COMPANY_MASTER on TSPL_DEMAND_BOOKING_MASTER.Comp_Code = TSPL_COMPANY_MASTER.Comp_Code
where 2=2 "
                        If rbtnMorning.IsChecked Then
                            qry += " and TSPL_DEMAND_BOOKING_DETAIL.ShiftType = 'Morning' "
                        ElseIf rbtnEvening.IsChecked Then
                            qry += " and TSPL_DEMAND_BOOKING_DETAIL.ShiftType = 'Evening' "
                        End If
                        qry += " and (CONVERT(date, TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)= '" & clsCommon.GetPrintDate(txtDate.Value) & "') 
and TSPL_DEMAND_BOOKING_MASTER.Route_No = '" & txtRouteNo.Value & "' and TSPL_ITEM_MASTER.Is_Ambient=1  and TSPL_DEMAND_BOOKING_MASTER.IsIndividualCustomer=0 )XXFinal
Group by XXFinal.Cust_Code,XXFinal.Sku_Seq ,XXFinal.Short_Description )XXXFinal
left join TSPL_CUSTOMER_MASTER on XXXFinal.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code"
                    End If
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry & " order by TSPL_CUSTOMER_MASTER.Credit_Customer,TSPL_CUSTOMER_MASTER.Display_Seq")
                        If rbtn_Fresh.IsChecked Then
                            If ItemCount > 0 AndAlso ItemCount <= 9 Then
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingFUDP", "Demand Booking")
                            Else
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingFUDP9", "Demand Booking")
                            End If
                        ElseIf rbtn_Ambient.IsChecked Then
                            If ItemCount > 0 AndAlso ItemCount <= 13 Then
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingAUDP", "Demand Booking")
                            Else
                                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingAUDP14", "Demand Booking")
                            End If
                        End If
                    Else
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry + " order by TSPL_CUSTOMER_MASTER.Display_Seq")
                        If rbtn_Fresh.IsChecked Then
                            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBooking", "Demand Booking")
                        ElseIf rbtn_Ambient.IsChecked Then
                            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptDemandBookingAUDP", "Demand Booking")
                        End If
                    End If
                    frmCRV = Nothing
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    '    Private Sub CreateDotMatrixReport(ByVal Baseqry As String)
    '        Dim qry As String = " select Short_Description,max(Route_No) as Route_No,max(Route_Desc) as Route_Desc,max(TranspoterName) as TranspoterName,max(DriverName) as DriverName,MAX(Vehicle_No) as Vehicle_No,max(Document_Date) as Document_Date,max(ShiftType) as ShiftType,max(DocStatus) as DocStatus from (" + Baseqry + " )xx group by Short_Description order by max(Sku_Seq)"
    '        Dim dtItem As DataTable = clsDBFuncationality.GetDataTable(qry)
    '        If dtItem IsNot Nothing AndAlso dtItem.Rows.Count <= 0 Then
    '            Throw New Exception("No Data found to print")
    '        End If
    '        qry = "select Cust_Code,case when Credit_Customer='Y' then 'Department Booth' else 'Normal Booth' end as Credit_Customer "
    '        For Each drItem As DataRow In dtItem.Rows
    '            qry += ",sum((case when Credit_Customer='Y' then QTYLtr else Crate end) * (case when Short_Description='" + clsCommon.myCstr(drItem("Short_Description")) + "' then 1 else 0 end)) as [" + clsCommon.myCstr(drItem("Short_Description")) + "] "
    '        Next
    '        qry += ",sum(case when Credit_Customer='Y' then QTYLtr else Crate end) as [TotalCrate]
    ',sum(ItemNetAmount) as ItemNetAmount
    ',sum(AmountBE) as AmountBE
    ',sum(TotalTCSAmt) as TotalTCSAmt
    ',sum(TotalCollectCrate) as TotalCollectCrate
    'from (
    '" + Baseqry + "
    ')xx Group by Cust_Code,Credit_Customer
    'order by xx.Credit_Customer,max(Display_Seq)"
    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
    '#Region "Add Grand Total"
    '        qry = "select  sum( QTYLtr ) as [TOTLTR] "
    '        For Each drItem As DataRow In dtItem.Rows
    '            qry += ",sum( QTYLtr * (case when Short_Description='" + clsCommon.myCstr(drItem("Short_Description")) + "' then 1 else 0 end)) as [LTR#$#" + clsCommon.myCstr(drItem("Short_Description")) + "]  
    ',sum( Crate * (case when Short_Description='" + clsCommon.myCstr(drItem("Short_Description")) + "' then 1 else 0 end)) as [CRATE#$#" + clsCommon.myCstr(drItem("Short_Description")) + "]  
    ',sum( Pouch * (case when Short_Description='" + clsCommon.myCstr(drItem("Short_Description")) + "' then 1 else 0 end)) as [POUCH#$#" + clsCommon.myCstr(drItem("Short_Description")) + "]  "
    '        Next
    '        qry += " from ( " + Baseqry + " )  xx "
    '        Dim dtTotal As DataTable = clsDBFuncationality.GetDataTable(qry)
    '        Dim drLtr As DataRow = dt.NewRow
    '        drLtr("Cust_Code") = "Litre"
    '        drLtr("Credit_Customer") = "Grand Total"
    '        drLtr("TotalCrate") = dtTotal.Rows(0)("TOTLTR")
    '        Dim drCrate As DataRow = dt.NewRow
    '        drCrate("Cust_Code") = "Crate"
    '        drCrate("Credit_Customer") = "Grand Total"
    '        drCrate("TotalCrate") = 0
    '        Dim drPourch As DataRow = dt.NewRow
    '        drPourch("Cust_Code") = "Pouch"
    '        drPourch("Credit_Customer") = "Grand Total"
    '        drPourch("TotalCrate") = 0
    '        For Each drItem As DataRow In dtItem.Rows
    '            drLtr(clsCommon.myCstr(drItem("Short_Description"))) = dtTotal.Rows(0)("LTR#$#" + clsCommon.myCstr(drItem("Short_Description")))
    '            drCrate(clsCommon.myCstr(drItem("Short_Description"))) = dtTotal.Rows(0)("CRATE#$#" + clsCommon.myCstr(drItem("Short_Description")))
    '            drPourch(clsCommon.myCstr(drItem("Short_Description"))) = dtTotal.Rows(0)("POUCH#$#" + clsCommon.myCstr(drItem("Short_Description")))
    '            Dim Quotient As Integer = clsCommon.myCDecimal(drPourch(clsCommon.myCstr(drItem("Short_Description"))) / 12)
    '            Dim Reminder As Integer = drPourch(clsCommon.myCstr(drItem("Short_Description"))) Mod 12
    '            drCrate(clsCommon.myCstr(drItem("Short_Description"))) += Quotient
    '            drPourch(clsCommon.myCstr(drItem("Short_Description"))) = Reminder
    '            drCrate("TotalCrate") += drCrate(clsCommon.myCstr(drItem("Short_Description")))
    '            drPourch("TotalCrate") += drPourch(clsCommon.myCstr(drItem("Short_Description")))
    '        Next
    '        dt.Rows.Add(drLtr)
    '        dt.Rows.Add(drCrate)
    '        dt.Rows.Add(drPourch)
    '        dt.AcceptChanges()
    '#End Region
    '        Dim obj As clsDosPrint = New clsDosPrint()
    '        obj.ReportName = objCommonVar.CurrentCompanyName
    '        obj.ReportName1 = "DAILY TENTATIVE DEMAND SHEET FOR AREA NO: " + clsCommon.myCstr(dtItem.Rows(0)("Route_No")) + " Date: " + clsCommon.GetPrintDate(dtItem.Rows(0)("Document_Date"), "dd/MM/yyyy") + " Shift: " + clsCommon.myCstr(dtItem.Rows(0)("ShiftType")) + " Status: " + clsCommon.myCstr(dtItem.Rows(0)("DocStatus"))
    '        obj.ReportName2 = clsCommon.myCstr(dtItem.Rows(0)("Route_Desc")) + " ( ROUTE: " + clsCommon.myCstr(dtItem.Rows(0)("Route_No")) + ")  " + clsCommon.myCstr(dtItem.Rows(0)("TranspoterName")) + "  DRIVER: " + clsCommon.myCstr(dtItem.Rows(0)("DriverName")) + "  VEHICLE NO:" + clsCommon.myCstr(dtItem.Rows(0)("Vehicle_No")) + "  PRINT AT: " + clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(), "dd/MM/yyyy hh:mm tt")
    '        obj.HideGroupHeader = True
    '        obj.HideLastGroupTotal = True
    '        'obj.ShowPageNo = True
    '        obj.PageSetupCustomizeCharColumn = 140
    '        obj.PageSetupCustomizeCharRows = 70
    '        obj.arrGroup = New List(Of clsDosPrintGroup)()
    '        obj.arrGroup.Add(clsDosPrintGroup.GetObject("Credit_Customer", "Details of", ""))
    '        obj.arrFilter = New List(Of clsDosPrintHeaderFilter)()
    '        obj.arrColumn = New List(Of clsDosPrintColumn)()
    '        obj.arrColumn.Add(clsDosPrintColumn.SetColumn("Cust_Code", "Booth", True, DosPrintAlignment.Left, 10, False, DecimalPlaces.NA))
    '        For Each drItem As DataRow In dtItem.Rows
    '            obj.arrColumn.Add(clsDosPrintColumn.SetColumn(clsCommon.myCstr(drItem("Short_Description")), clsCommon.myCstr(drItem("Short_Description")), False, DosPrintAlignment.Right, 10, True, DecimalPlaces.One))
    '        Next
    '        obj.arrColumn.Add(clsDosPrintColumn.SetColumn("TotalCrate", "Total", False, DosPrintAlignment.Right, 10, True, DecimalPlaces.Zero))
    '        obj.arrColumn.Add(clsDosPrintColumn.SetColumn("ItemNetAmount", "Shift Amt", False, DosPrintAlignment.Right, 12, True, DecimalPlaces.Two))
    '        obj.arrColumn.Add(clsDosPrintColumn.SetColumn("AmountBE", "  Total Amt", False, DosPrintAlignment.Right, 12, True, DecimalPlaces.Two))
    '        obj.arrColumn.Add(clsDosPrintColumn.SetColumn("TotalTCSAmt", "TCS", False, DosPrintAlignment.Right, 10, True, DecimalPlaces.Two))
    '        obj.arrColumn.Add(clsDosPrintColumn.SetColumn("TotalCollectCrate", "Crate Coll.", False, DosPrintAlignment.Right, 10, True, DecimalPlaces.Zero))
    '        obj.Print(obj, dt, PageSetup.Landscap)
    '    End Sub
    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Export(EnumExportTo.Excel)
    End Sub
    Private Sub Export(ByVal exporter As EnumExportTo)
        Try
            If gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If
            'Dim sfd As SaveFileDialog = New SaveFileDialog()
            'Dim filePath As String = ""
            'sfd.FileName = Me.Text
            'sfd.Filter = "Excel 97-2003 (*.xls) |*.xls;|Excel 2007 (*.xlsx)|*.xlsx;|CSV Files (*.csv) |*.csv"
            'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            '    filePath = sfd.FileName
            'Else
            'End If
            'If Not filePath.Equals(String.Empty) Then
            If exporter = EnumExportTo.Excel Then
                transportSql.applyExportTemplate(gv1, PageSetupReport_ID)
                transportSql.exportdata(gv1, "", Me.Text, , Nothing, False, False, True)
            End If
            ' End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub
    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim gvImport As New UserControls.MyRadGridView
        Dim arrVisbleColumns As New List(Of Integer)
        Try
            If clsCommon.myLen(gv1.Rows(0).Cells(colCustCode).Value) = 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Import", Me.Text)
                Exit Sub
            End If
            Me.Controls.Add(gvImport)
            If clsDemandBookingImport.importExcel(gvImport) Then
                LoadBlankGrid()
                RouteData(False, False)
                For ii As Integer = 0 To gv1.Columns.Count - 1
                    If gv1.Columns(ii).IsVisible Then
                        arrVisbleColumns.Add(ii)
                    End If
                Next
                If gvImport.Rows.Count > 0 Then
                    If clsCommon.CompairString(gv1.Rows.Count - 1, gvImport.Rows.Count - 2) = CompairStringResult.Equal Then
                        Try
                            Dim arrCustCodeExist As New List(Of String)
                            For i As Integer = 0 To gv1.Rows.Count - 1
                                arrCustCodeExist.Add(gv1.Rows(i).Cells(1).Value)
                            Next
                            isInsideLoadData = True
                            clsCommon.ProgressBarPercentShow()
                            For ii As Integer = 1 To gvImport.Rows.Count - 2
                                clsCommon.ProgressBarPercentUpdate(((ii + 1) * 100) / gvImport.RowCount, clsCommon.myCstr((ii + 1)) & "/" & clsCommon.myCstr(gvImport.RowCount))
                                For jj As Integer = 0 To gv1.Rows.Count - 1
                                    Dim code As String = clsCommon.myCstr(gvImport.Rows(ii).Cells(1).Value)
                                    If arrCustCodeExist.Contains(code) Then
                                        If clsCommon.CompairString(clsCommon.myCstr(gvImport.Rows(ii).Cells(2).Value), clsCommon.myCstr(gv1.Rows(jj).Cells(colCustCode).Value)) = CompairStringResult.Equal Then
                                            For kk As Integer = 5 To arrVisbleColumns.Count - 11
                                                If clsCommon.myCDecimal(gv1.Rows(jj).Cells(arrVisbleColumns(kk)).Value) <> clsCommon.myCDecimal(gvImport.Rows(ii).Cells(kk).Value) Then
                                                    gv1.Rows(jj).Cells(arrVisbleColumns(kk)).Value = gvImport.Rows(ii).Cells(kk).Value
                                                End If
                                                'clsCommon.ProgressBarPercentUpdate(((ii + 1) * 100) / gvImport.RowCount, clsCommon.myCstr((ii + 1)) + "/" + clsCommon.myCstr(gvImport.RowCount) + "  " + clsCommon.myCstr(gvImport.Rows(ii).Cells(kk).Value))
                                            Next
                                        End If
                                    Else
                                        clsCommon.MyMessageBoxShow("Default Customer '" & clsCommon.myCstr(gvImport.Rows(ii).Cells(1).Value) & "' Does Not Exist", Me.Text)
                                        Exit Sub
                                    End If
                                Next
                            Next
                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        Finally
                            clsCommon.ProgressBarPercentHide()
                        End Try
                        isInsideLoadData = False
                        UpdateAllTotals(False)
                        clsCommon.MyMessageBoxShow(Me, "Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
                    Else
                        clsCommon.MyMessageBoxShow("You cannot import quantity because both Import and Export Data is different", Me.Text)
                    End If
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        Finally
            Me.Controls.Remove(gvImport)
        End Try
    End Sub
    Public Function GetOutStandingBal(ByVal VendorNo As String, ByVal dtDoc As DateTime) As Decimal
        Dim OSBal As Decimal = 0
        Try
            Dim strcustomerfilter As String = String.Empty
            strcustomerfilter = "'" & VendorNo & "'"
            'Dim qry = ""
            Dim qry As String = "Select  case when (( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt)) )>=0 then -abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) else abs(( SUM(convert(decimal(18,2),OpngBal)) + SUM(convert(decimal(18,2),DrAmt)) ) -SUM(convert(decimal(18,2),CrAmt))) end  as BalAmt From ( " &
                    "Select MAX(TSPL_CUSTOMER_MASTER.Cust_Group_Code) as Cust_Group_Code, ACode, MAX(TSPL_CUSTOMER_MASTER.Customer_Name) as AName, '' as CurrencyCode,  " &
                    "null as ConvRate, SUM(DrAmt* Final.ConvRate)-SUM(CrAmt) as OpngBal, 0 as DrAmt, 0 as CrAmt, 0 as [Sales], 0 as CollectionRefund, 0 as DrNote,  " &
                    "0 as CrNote, MAX(tspl_customer_master.Cust_Category_Code) as Cust_Category_Code,MAX(CUST_CATEGORY_DESC) as Cust_Category_Desc,  " &
                    "MAX(tspl_customer_master.Cust_Type_Code) As Cust_Type_Code,MAX(Cust_Type_Desc) As Cust_Type_Desc from   " &
                    "(" & clsCustomerMaster.GetCustomerBaseQry(False, False, "", False, "ConvRate", strcustomerfilter, True, clsCommon.GetPrintDate(dtDoc.AddDays(1), "dd/MMM/yyyy"), "", False, False, True, Nothing, False) & "   " &
                    " ) Final left outer join TSPL_CUSTOMER_MASTER on final.ACode=TSPL_CUSTOMER_MASTER.Cust_Code LEFT OUTER JOIN TSPL_CUSTOMER_GROUP_MASTER ON TSPL_CUSTOMER_GROUP_MASTER.Cust_Group_Code=TSPL_CUSTOMER_MASTER.Cust_Group_Code " &
                    "Left outer join TSPL_RECEIPT_HEADER on TSPL_RECEIPT_HEADER.Receipt_No =Final.DocNo  LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE=Final.Bank_Code " &
                    "where  CONVERT(DATE,final.DocDate,103) <= '" & clsCommon.GetPrintDate(dtDoc, "dd/MMM/yyyy") & "' AND LEN(ACode)>0 and ACode in ('" & VendorNo & "')   AND TSPL_CUSTOMER_MASTER.Status='N' GROUP BY ACode " &
                    ") XXX GROUP BY ACode ORDER BY ACode"
            Dim dblBal As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, Nothing))
            OSBal = dblBal
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return OSBal
    End Function
    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles btnFullMode.Click
        Try
            If gvFullMode Then
                rgbDemandHead.Visible = True
                RadGroupBox2.Location = New System.Drawing.Point(6, 190)
                RadGroupBox2.Size = New System.Drawing.Size(1114, 400)
                gvFullMode = False
            Else
                gvFullMode = True
                rgbDemandHead.Visible = False
                RadGroupBox2.Location = New System.Drawing.Point(6, 6)
                RadGroupBox2.Size = New System.Drawing.Size(1114, 600)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnQuickDemand_Click(sender As Object, e As EventArgs) Handles btnQuickDemand.Click
        Try
            Dim qry As String = "select Route_No,max(demand_date) as DemandDate,max(shiftType) as ShiftType from TSPL_DEMAND_SHEET where convert(date,demand_date,103)='" + clsCommon.GetPrintDate(txtDate.Value, "dd/MMM/yyyy") + "' and ShiftType='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "'  group by Route_No"  ''and Created_By='" + objCommonVar.CurrentUserCode + "'
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                clsCommon.ProgressBarShow()
                For Each dr As DataRow In dt.Rows
                    AddNew()
                    txtDocNo.Value = ""
                    txtDate.Value = clsCommon.GetPrintDate(dr.Item("DemandDate"))
                    If clsCommon.CompairString(clsCommon.myCstr(dr.Item("ShiftType")), "Morning") = CompairStringResult.Equal Then
                        rbtnMorning.IsChecked = True
                    Else
                        rbtnEvening.IsChecked = True
                    End If
                    QuickDemamd(clsCommon.GetPrintDate(dr.Item("DemandDate")), objCommonVar.CurrentUserCode, clsCommon.myCstr(dr.Item("ShiftType")), clsCommon.myCstr(dr.Item("Route_No")))
                Next
                AddNew()
                clsCommon.ProgressBarHide()
                clsCommon.MyMessageBoxShow(Me, "Data Saved  ", Me.Text)
            Else
                clsCommon.MyMessageBoxShow(Me, "Demand Sheet Found ", Me.Text)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub QuickDemamd(ByVal DocDate As Date, ByVal CurrUser As String, ByVal ShiftType As String, ByVal RouteNo As String)
        Try
            Dim qry As String = " select Document_No from TSPL_DEMAND_BOOKING_MASTER where convert(date,Document_Date,103)='" + clsCommon.GetPrintDate(DocDate, "dd/MMM/yyyy") + "' and ShiftType='" + ShiftType + "' and Route_No='" + RouteNo + "'"
            Dim DocumentNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
            If clsCommon.myLen(DocumentNo) > 0 Then
                qry = "select Posted from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + DocumentNo + "' "
                Dim isPosted As Boolean = IIf(clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) = 1, True, False)
                If Not isPosted Then
                    LoadData(DocumentNo, NavigatorType.Current)
                    FillQuickDemandData(DocDate, ShiftType)
                    SaveData(True)
                Else
                    clsCommon.MyMessageBoxShow(Me, "Document Posted for Route No [ " & RouteNo & " ]", Me.Text)
                End If
            Else
                Dim isClicked As Boolean = False
                txtRouteNo.Value = RouteNo
                RouteData(isClicked, True)
                FillQuickDemandData(DocDate, ShiftType)
                SaveData(True)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Public Sub FillQuickDemandData(ByVal docDate As Date, ByVal ShiftType As String)
        Try
            Dim Qty As Integer = 0
            Dim qry As String = "select Cust_Code,Item_Code,Qty from TSPL_DEMAND_SHEET where convert(date,demand_date,103)='" + clsCommon.GetPrintDate(docDate, "dd/MMM/yyyy") + "' and ShiftType='" + ShiftType + "' and Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' "  ' and Created_By='" + objCommonVar.CurrentUserCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    For dblrow As Integer = 0 To gv1.Rows.Count - 2
                        If clsCommon.CompairString(gv1.Rows(dblrow).Cells(colCustCode).Value, clsCommon.myCstr(dr.Item("Cust_Code"))) = CompairStringResult.Equal Then
                            Dim k As Integer = 1
                            For dblcolumns As Integer = 9 To gv1.Columns.Count - 11
                                Dim obj1 As ItemValueClass = Nothing
                                Try
                                    obj1 = TryCast(gv1.Columns(colItemCode + clsCommon.myCstr(k)).Tag, ItemValueClass)
                                Catch ex As Exception
                                End Try
                                If obj1 IsNot Nothing AndAlso clsCommon.CompairString(clsCommon.myCstr(obj1.itemCode), clsCommon.myCstr(dr.Item("Item_Code"))) = CompairStringResult.Equal AndAlso clsCommon.CompairString("Crate", clsCommon.myCstr(obj1.Unit_code)) = CompairStringResult.Equal Then
                                    'If clsCommon.CompairString(clsCommon.myCstr(obj1.itemCode), clsCommon.myCstr(dr.Item("Item_Code"))) = CompairStringResult.Equal AndAlso clsCommon.CompairString("Crate", clsCommon.myCstr(obj1.Unit_code)) = CompairStringResult.Equal Then
                                    Qty = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select top 1 Qty from TSPL_DEMAND_SHEET where convert(date,demand_date,103)='" & clsCommon.GetPrintDate(docDate, "dd/MMM/yyyy") & "' and ShiftType='" & ShiftType & "'  and Item_Code='" & clsCommon.myCstr(dr.Item("Item_Code")) & "' and Cust_Code='" & clsCommon.myCstr(dr.Item("Cust_Code")) & "'  order by Modify_Date desc"))
                                    If Qty > 0 Then
                                        gv1.Rows(dblrow).Cells(dblcolumns).Value = Qty
                                    Else
                                        gv1.Rows(dblrow).Cells(dblcolumns).Value = ""
                                    End If
                                    'End If
                                End If
                                k = k + 1
                            Next
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub gv1_CellClick(sender As Object, e As GridViewCellEventArgs) Handles gv1.CellClick
        If e.Column Is gv1.Columns(colbtncol) Then
            Try
                If clsCommon.myLen(txtDocNo.Value) > 0 Then
                    If common.clsCommon.MyMessageBoxShow(Me, "Do You Want to Reset Demand for Booth  " & clsCommon.myCstr(gv1.CurrentRow.Cells(colCustCode).Value) & Environment.NewLine & "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        DeleteBoothDemand(txtDocNo.Value, gv1.CurrentRow.Cells(colCustCode).Value, IIf(rbtnMorning.IsChecked, "Morning", "Evening"))
                    End If
                Else
                    Throw New Exception("Document not Found!")
                End If
            Catch ex As Exception
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End Try
        End If
    End Sub
    Public Sub DeleteBoothDemand(ByVal DocNo As String, ByVal cust_code As String, ByVal ShiftType As String)
        Try
            If clsDemandBookingSale.DeleteBoothDemand(DocNo, cust_code, ShiftType, False) Then
                clsCommon.MyMessageBoxShow(Me, "Demand Reset for Booth No :" & clsCommon.myCstr(cust_code), Me.Text)
                LoadData(DocNo, NavigatorType.Current)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub ResetDemandOnSave(ByVal DocNo As String)
        Try
            For dblrows As Integer = 0 To gv1.Rows.Count - 1
                If clsCommon.CompairString(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colItemExist).Value), "No") = CompairStringResult.Equal AndAlso clsCommon.myLen(clsCommon.myCstr(gv1.Rows(dblrows).Cells(colCustCode).Value)) > 0 Then
                    clsDemandBookingSale.DeleteBoothDemand(DocNo, gv1.Rows(dblrows).Cells(colCustCode).Value, IIf(rbtnMorning.IsChecked = True, "Morning", "Evening"), False)
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub txtCustomerNo_Load(sender As Object, e As EventArgs) Handles txtCustomerNo.Load
    End Sub
    Private Sub btnPrintChallan_Click(sender As Object, e As EventArgs) Handles btnPrintChallan.Click
        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                PrintChallan()
            Else
                Throw New Exception("Please select Document")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Public Sub PrintChallan()
        Try

            Dim qry As String = " select  max(TSPL_COMPANY_MASTER.Comp_Name) as Comp_Name,( max(TSPL_COMPANY_MASTER.Add1) + max(TSPL_COMPANY_MASTER.Add2) + Max(TSPL_COMPANY_MASTER.Add3)) as Company_Address,
max(TSPL_DEMAND_BOOKING_MASTER.Document_Date) as Document_Date,max(TSPL_DEMAND_BOOKING_MASTER.ShiftType) as ShiftType,TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,
max(TSPL_CUSTOMER_MASTER.Customer_Name) as Customer_Name,max(TSPL_DEMAND_BOOKING_MASTER.Route_No) as Route_No,max(TSPL_Route_Master.Route_Desc)  as Route_Desc,
max(TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code) as Vehicle_Code,max(TSPL_VEHICLE_MASTER.Number) as Vehicle_No,max(TSPL_ITEM_MASTER.HSN_Code) as HSN_Code,
TSPL_DEMAND_BOOKING_DETAIL.Item_Code,max(TSPL_ITEM_MASTER.Short_Description) as Short_Description,TSPL_DEMAND_BOOKING_DETAIL.Unit_code,sum(TSPL_DEMAND_BOOKING_DETAIL.Qty) as Qty ,
sum(TSPL_DEMAND_BOOKING_DETAIL.Item_Rate) as Item_Rate,sum(TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount) as ItemNetAmount
"
            If isDepartmentRouteSetting Then
                qry += " ,convert(int,sum(TSPL_DEMAND_BOOKING_DETAIL.Qty)*max(ITEMDETAIL.CFForLTR)/ max(ITEMDETAILInCrate.CFForLTR)) as Crate,
  (sum(TSPL_DEMAND_BOOKING_DETAIL.Qty)* max(ITEMDETAIL.CFForLTR)/ max(ITEMDETAILInpouch.CFForLTR) -((convert(int,sum(TSPL_DEMAND_BOOKING_DETAIL.Qty)* max(ITEMDETAIL.CFForLTR)/ max(ITEMDETAILInCrate.CFForLTR)))* max(ITEMDETAILInCrate.CFForLTR))) AS Pouch "
            End If

            qry += " from TSPL_DEMAND_BOOKING_MASTER
left join TSPL_DEMAND_BOOKING_DETAIL on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left join TSPL_ITEM_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
left join TSPL_CUSTOMER_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
left outer join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Location_Code=TSPL_DEMAND_BOOKING_MASTER.location_code 
Left Outer Join TSPL_Route_Master On TSPL_Route_Master.Route_No=TSPL_DEMAND_BOOKING_MASTER.Route_No 
Left Outer Join TSPL_VEHICLE_MASTER On TSPL_VEHICLE_MASTER.Vehicle_Id=TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code 
left outer join TSPL_COMPANY_MASTER on  TSPL_COMPANY_MASTER.Comp_Code1 = '" & objCommonVar.CurrComp_Code1 & "' "
            If isDepartmentRouteSetting Then
                qry += " Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code 
  Left Join (select Conversion_factor AS CFForLTR, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'LTR'  ) as ITEMDETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAIL.Item_code 
  Left Join (select Conversion_factor AS CFForLTR, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'Crate'  ) as ITEMDETAILInCrate on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAILInCrate.Item_code 
                            
Left Join (select Conversion_factor AS CFForLTR, TSPL_ITEM_UOM_DETAIL.Item_code from TSPL_ITEM_UOM_DETAIL where UOM_code = 'Pouch'  ) as ITEMDETAILInpouch on TSPL_ITEM_UOM_DETAIL.Item_Code = ITEMDETAILInpouch.Item_code "
            End If
            qry += " where TSPL_DEMAND_BOOKING_MASTER.Document_No='" & txtDocNo.Value & "' and TSPL_DEMAND_BOOKING_MASTER.Posted=1 and TSPL_CUSTOMER_MASTER.Credit_Customer='Y'
group by TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_DEMAND_BOOKING_DETAIL.Unit_code "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                If isDepartmentRouteSetting Then
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptDemandBookingChallanForDepartment", "Demand Booking Challan For Department")
                Else
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "crptDemandBookingChallan", "Demand Booking Challan")
                End If
                'frmCRV.funsubreportWithdt(CrystalReportFolder.KwalitySalesReport, dt, dt2, "rptDemandBooking", "Demand Booking", "rptSubDemandBooking")
                frmCRV = Nothing
            Else
                Throw New Exception("Data Not Found!")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub ShowRemarks()
        Try
            Dim Reason As String = ""
            Dim frm As New FrmFreeTxtBox1
            frm.Text = "Remarks for Unpost"
            frm.ShowDialog()
            If clsCommon.myLen(frm.strRmks) <= 0 Then
                Exit Sub
            Else
                Reason = frm.strRmks
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub brnReverseAndUnpost_Click(sender As Object, e As EventArgs) Handles btnReverseAndUnpost.Click

        Dim frm1 As New FrmPWD(Nothing)
        frm1.strType = clsFixedParameterType.Transactionupdate
        frm1.strCode = clsFixedParameterCode.DemandUnpost
        frm1.ShowDialog()
        If frm1.isPasswordCorrect Then
            Reverse()
            OneTimeCheck = True
        End If

    End Sub
    Sub Reverse()
        'Dim NextDayDocNo As String = ""
        'Try
        '    If clsCommon.myLen(txtDocNo.Value) > 0 Then
        '        Dim isDispatch As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select COUNT(*) from TSPL_SD_SHIPMENT_BOOKING_DETAIL where Booking_TR_Code in(select TR_Code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + txtDocNo.Value + "')"))
        '        If isDispatch >= 1 Then
        '            Throw New Exception("Dispatch already Created!")
        '        End If
        '    Else
        '        Throw New Exception("Please Select Document")

        '    End If
        '    If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, Nothing)) = 1 Or clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, Nothing)) = 1 Then
        '        NextDayDocNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 103 )='" + clsCommon.GetPrintDate(txtDate.Value.AddDays(1)) + "') and location_code='" + clsCommon.myCstr(txtLocation.Value) + "' and ShiftType='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' and IsIndividualCustomer=0 ")
        '    End If
        '    'Dim dt As DataTable = clsDBFuncationality.GetDataTable("select location_code,Document_Date from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + NextDayDocNo + "'", "")
        '    'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '    'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Dairy Sale", "Demand Booking", txtLocation.Value, txtDate.Value, trans)
        '    'End If
        '    If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document " + IIf(clsCommon.myLen(NextDayDocNo) > 0, "and Delete Next Day Document [" + NextDayDocNo + "]", "") + " " + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
        '        ' REASON FOR DELETE 
        '        Dim Reason As String = ""
        '        Dim qry As String = ""
        '        Dim frm As New FrmFreeTxtBox1
        '        frm.Text = "Remarks for Reverse"
        '        frm.ShowDialog()
        '        If clsCommon.myLen(frm.strRmks) <= 0 Then
        '            Exit Sub
        '        Else
        '            Reason = frm.strRmks
        '        End If
        '        If clsCommon.myLen(clsCommon.myCstr(NextDayDocNo)) > 0 Then
        '            qry = "select Posted from TSPL_Demand_BOOKING_MAstER where Document_No='" + NextDayDocNo + "'"
        '            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) = 1 Then
        '                Throw New Exception("Please Reverse/Unpost Document No: [ " + NextDayDocNo + " ]")
        '            End If
        '            Dim dt As DataTable = Nothing
        '            '' to check gatepass or truck sheet generated
        '            Dim strDocNoForGatePass As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & NextDayDocNo & "' and IsGatePassGenerated='Y' "))
        '            Dim strDocNoForTrucksheet As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & NextDayDocNo & "' and  IsTruckSheetGenerated ='Y'  "))
        '            If clsCommon.myLen(clsCommon.myCstr(strDocNoForGatePass)) > 0 Then
        '                Throw New Exception("Demand cannot be reverse because Next Day Demand Gate Pass has generated.")
        '            End If
        '            If clsCommon.myLen(clsCommon.myCstr(strDocNoForTrucksheet)) > 0 Then
        '                Throw New Exception("Demand cannot be reverse because Next Day Demand Gate Pass has generated.")
        '            End If
        '        End If
        '        If clsCommon.myLen(NextDayDocNo) > 0 Then
        '            If clsDemandBookingSale.DeleteData(NextDayDocNo) Then
        '                If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
        '                    saveCancelLog(Reason, "Reverse And Recreate", Nothing)
        '                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
        '                    LoadData(txtDocNo.Value, NavigatorType.Current)
        '                End If
        '            End If
        '        Else
        '            If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
        '                saveCancelLog(Reason, "Reverse And Recreate", Nothing)
        '                common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
        '                LoadData(txtDocNo.Value, NavigatorType.Current)
        '            End If
        '        End If
        '    End If
        'Catch ex As Exception
        '    clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        'End Try
        Dim NextDayDocNo As String = ""

        Try
            If clsCommon.myLen(txtDocNo.Value) > 0 Then
                Dim isDispatch As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select COUNT(*) from TSPL_SD_SHIPMENT_BOOKING_DETAIL where Booking_TR_Code in(select TR_Code from TSPL_DEMAND_BOOKING_DETAIL where Document_No='" + txtDocNo.Value + "')"))
                If isDispatch >= 1 Then
                    Throw New Exception("Dispatch already Created!")
                End If
            Else
                Throw New Exception("Please Select Document")

            End If
            If objCommonVar.ApplyBoothRouteMapping Then
                If Not chkIndividualCustomer.Checked Then
                    NextDayDocNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 103 )='" + clsCommon.GetPrintDate(txtDate.Value.AddDays(1)) + "') and location_code='" + clsCommon.myCstr(txtLocation.Value) + "' and ShiftType='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' and IsIndividualCustomer=0 ")

                End If
                If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document " + IIf(clsCommon.myLen(NextDayDocNo) > 0, "and Delete Next Day Document [" + NextDayDocNo + "]", "") + " " + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    Dim Reason As String = ""
                    Dim qry As String = ""
                    Dim frm As New FrmFreeTxtBox1
                    frm.Text = "Remarks for Reverse"
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 Then
                        Exit Sub
                    Else
                        Reason = frm.strRmks
                    End If
                    If chkIndividualCustomer.Checked Then
                        If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
                            saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                            common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                            LoadData(txtDocNo.Value, NavigatorType.Current)
                        End If
                    Else
                        If clsDemandBookingSale.ReverseMultipleDOC(txtDocNo.Value, txtRouteNo.Value, clsCommon.GetPrintDate(txtDate.Value.AddDays(1)), txtLocation.Value, IIf(rbtnMorning.IsChecked, "Morning", "Evening")) Then
                            saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                            common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                            LoadData(txtDocNo.Value, NavigatorType.Current)
                        End If
                    End If

                End If

            Else
                If clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandAll, clsFixedParameterCode.ApplyDemandAll, Nothing)) = 1 Or clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyDemandCustomerWise, clsFixedParameterCode.ApplyDemandCustomerWise, Nothing)) = 1 Then
                    If Not chkIndividualCustomer.Checked Then
                        NextDayDocNo = clsDBFuncationality.getSingleValue("select Document_No from TSPL_DEMAND_BOOKING_MASTER where Route_No='" + clsCommon.myCstr(txtRouteNo.Value) + "' and ( CONVERT( date, TSPL_DEMAND_BOOKING_MASTER.Document_Date, 103 )='" + clsCommon.GetPrintDate(txtDate.Value.AddDays(1)) + "') and location_code='" + clsCommon.myCstr(txtLocation.Value) + "' and ShiftType='" + IIf(rbtnMorning.IsChecked, "Morning", "Evening") + "' and IsIndividualCustomer=0 ")

                    End If
                End If
                'Dim dt As DataTable = clsDBFuncationality.GetDataTable("select location_code,Document_Date from TSPL_DEMAND_BOOKING_MASTER where Document_No='" + NextDayDocNo + "'", "")
                'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                'clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, "Dairy Sale", "Demand Booking", txtLocation.Value, txtDate.Value, trans)
                'End If
                If common.clsCommon.MyMessageBoxShow(Me, "Reverse and Unpost the Current Document " + IIf(clsCommon.myLen(NextDayDocNo) > 0, "and Delete Next Day Document [" + NextDayDocNo + "]", "") + " " + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                    '' REASON FOR DELETE 
                    Dim Reason As String = ""
                    Dim qry As String = ""
                    Dim frm As New FrmFreeTxtBox1
                    frm.Text = "Remarks for Reverse"
                    frm.ShowDialog()
                    If clsCommon.myLen(frm.strRmks) <= 0 Then
                        Exit Sub
                    Else
                        Reason = frm.strRmks
                    End If
                    If clsCommon.myLen(clsCommon.myCstr(NextDayDocNo)) > 0 Then
                        qry = "select Posted from TSPL_Demand_BOOKING_MAstER where Document_No='" + NextDayDocNo + "'"
                        If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry)) = 1 Then
                            Throw New Exception("Please Reverse/Unpost Document No: [ " + NextDayDocNo + " ]")
                        End If
                        ' Dim dt As DataTable = Nothing
                        '' to check gatepass or truck sheet generated
                        Dim strDocNoForGatePass As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & NextDayDocNo & "' and IsGatePassGenerated='Y' "))
                        Dim strDocNoForTrucksheet As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select top 1 Document_No  from TSPL_DEMAND_BOOKING_DETAIL where document_No='" & NextDayDocNo & "' and  IsTruckSheetGenerated ='Y'  "))
                        If clsCommon.myLen(clsCommon.myCstr(strDocNoForGatePass)) > 0 Then
                            Throw New Exception("Demand cannot be reverse because Next Day Demand Gate Pass has generated.")
                        End If
                        If clsCommon.myLen(clsCommon.myCstr(strDocNoForTrucksheet)) > 0 Then
                            Throw New Exception("Demand cannot be reverse because Next Day Demand Gate Pass has generated.")
                        End If
                    End If
                    If clsCommon.myLen(NextDayDocNo) > 0 Then
                        If clsDemandBookingSale.DeleteData(NextDayDocNo) AndAlso clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
                            'If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
                            saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                            common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                            LoadData(txtDocNo.Value, NavigatorType.Current)
                            'End If
                        End If
                    Else
                        If clsDemandBookingSale.ReverseAndUnpost(txtDocNo.Value) Then
                            saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                            common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                            LoadData(txtDocNo.Value, NavigatorType.Current)
                        End If
                    End If
                End If


            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnSplitPrint_Click(sender As Object, e As EventArgs) Handles btnSplitPrint.Click
        Try
            Dim qry As String = Nothing
            'Dim SubRptQry As String = Nothing
            Dim ShiftType As String = ""
            Dim shiftAMPMType As String = ""
            Dim PreshiftAMPMType As String = ""
            Dim Previous_Shift As String = ""
            Dim Previous_Date As String
            'Dim ItemCount As Double = 0
            If rbtnEvening.IsChecked Then
                ShiftType = "Evening"
                shiftAMPMType = "PM"
                PreshiftAMPMType = "AM"
                Previous_Shift = "Morning"
                Previous_Date = clsCommon.myCDate(txtDate.Value).AddDays(1)
            Else
                ShiftType = "Morning"
                Previous_Shift = "Evening"
                shiftAMPMType = "AM"
                PreshiftAMPMType = "PM"
                ' Previous_Date = clsDBFuncationality.getSingleValue("select CONVERT(varchar, DATEADD(DAY, -1, convert(Nvarchar, '" & txtDate.Value & "' ,112)),21) as Previous_Date")
                Previous_Date = clsCommon.myCDate(txtDate.Value).AddDays(-1)
            End If
            If clsCommon.myLen(txtDocNo.Value) <= 0 Then
                Throw New Exception("Booking not found to Print")
            End If
            qry = " Select Count(1) from TSPL_ROUTE_MASTER Where Split_Print=1 And Route_No='" & txtRouteNo.Value & "'"
            If clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry)) <= 0 Then
                Throw New Exception("Split data not found to print on Route No: " & clsCommon.myCstr(txtRouteNo.Value) & ".")
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                Dim arrRoute As New ArrayList
                arrRoute.Add(txtRouteNo.Value)
                clsDemandBookingSale.PrintDOSData(arrRoute, ShiftType, txtDate.Value, rbtn_Fresh.IsChecked, rbtn_Ambient.IsChecked, chkIndividualCustomer.Checked, 107, 48, DosPaperSize.A4, PageSetup.Landscap, True, isDepartmentRouteSetting) ''
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub CrateHisTable()

        Dim coll As Dictionary(Of String, String)
        coll = New Dictionary(Of String, String)()
        coll.Add("PK_Id", "integer NOT NULL identity NOT FOR REPLICATION primary key")
        coll.Add("TR_Code", "varchar(30) NOT NULL ")
        coll.Add("Document_No", "varchar(30) NOT NULL")
        coll.Add("Document_Date", "datetime not NULL")
        coll.Add("Route_No", "varchar(12) NULL")
        coll.Add("Line_No", "integer not null default 0")
        coll.Add("Cust_Code", "Varchar(12) null ")
        coll.Add("Item_Code", "Varchar(50) null ")
        coll.Add("Qty", "decimal(18,2) null")
        coll.Add("Unit_code", "Varchar(12) null")
        coll.Add("Vehicle_Code", "Varchar(12) null")
        coll.Add("Item_Rate", "decimal(18,2) not null default 0")
        coll.Add("Price_code", "varchar(12) NULL")
        coll.Add("ShiftType", "varchar(20) NULL")
        coll.Add("IsItemUpdate", "int not null default 0")
        coll.Add("TotalCrates_ItemWise", "decimal(18,2) null")
        coll.Add("TotalLtr_ItemWise", "decimal(18,2) null")
        coll.Add("ItemNetAmount", "decimal(18,2) null")
        coll.Add("IsGatePassGenerated", "char(1) not null default 'N'")
        coll.Add("IsTruckSheetGenerated", "char(1) not null default 'N'")
        coll.Add("Production_Remarks", "varchar(200) NULL")
        coll.Add("GPCode", "varchar(30) NULL")
        coll.Add("Is_Posted", "char(1) not null default 'N'")
        coll.Add("Trip_No", "integer null")
        coll.Add("TAX_Group", "varchar(12) NULL")
        coll.Add("TAX1", "varchar(12) NULL")
        coll.Add("TAX1_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX1_Rate", "decimal(18, 4) NULL")
        coll.Add("TAX1_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX2", "varchar(12) NULL")
        coll.Add("TAX2_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX2_Rate", "decimal(18, 4) NULL")
        coll.Add("TAX2_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX3", "varchar(12) NULL")
        coll.Add("TAX3_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX3_Rate", "decimal(18, 4) NULL")
        coll.Add("TAX3_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX4", "varchar(12) NULL")
        coll.Add("TAX4_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX4_Rate", "decimal(18, 4) NULL")
        coll.Add("TAX4_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX5", "varchar(12) NULL")
        coll.Add("TAX5_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX5_Rate", "decimal(18, 4) NULL")
        coll.Add("TAX5_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX6", "varchar(12) NULL")
        coll.Add("TAX6_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX6_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX6_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX7", "varchar(12) NULL")
        coll.Add("TAX7_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX7_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX7_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX8", "varchar(12) NULL")
        coll.Add("TAX8_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX8_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX8_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX9", "varchar(12) NULL")
        coll.Add("TAX9_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX9_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX9_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX10", "varchar(12) NULL")
        coll.Add("TAX10_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX10_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX10_Amt", "decimal(18, 2) NULL")
        coll.Add("Created_By", "varchar(30) NULL")
        coll.Add("change_timestamp", "DateTime Default CURRENT_TIMESTAMP")
        coll.Add("operation_type", "VARCHAR(50)")
        coll.Add("operation_Source", "VARCHAR(50)")
        coll.Add("Hist_Version", "integer NOT NULL")
        coll.Add("Hist_By", "VARCHAR(50) NOT NULL")

        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_DEMAND_BOOKING_DETAIL_HISTORY", coll, "", False, False, "TSPL_DEMAND_BOOKING_MASTER", "Document_No", "")
    End Sub

    Private Sub btnShuffle_Click(sender As Object, e As EventArgs) Handles btnShuffle.Click
        Try
            clsCommon.ProgressBarShow()

            clsDemandBookingSale.ShuffleBoothRouteData(txtShuffleDate.Value, cmbShift.Text, txtShuffleRoute.arrValueMember)
            clsCommon.ProgressBarHide()
            clsCommon.MyMessageBoxShow(Me, "Shuffled Successffuly")

        Catch ex As Exception
            clsCommon.ProgressBarHide()

            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmi_Indent_PDF_Click(sender As Object, e As EventArgs) Handles rmi_Indent_PDF.Click
        Try
            ExportPDF()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmi_Indent_Excel_Click(sender As Object, e As EventArgs) Handles rmi_Indent_Excel.Click
        Try
            exportExcel()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtShuffleRoute__My_Click(sender As Object, e As EventArgs) Handles txtShuffleRoute._My_Click
        Try
            Dim qry As String = ""
            qry = " select Route_No as Code,Route_Desc from TSPL_ROUTE_MASTER where Status='A' "

            txtShuffleRoute.arrValueMember = clsCommon.ShowMultipleSelectForm("Routeno@ShuffleDemand", qry, "Code", "Code", txtShuffleRoute.arrValueMember, txtShuffleRoute.arrDispalyMember)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Sub BoothSlipExport()
        Try
            Dim ShiftType As String = ""
            Dim qry As String = ""
            Dim itemqry As String = ""
            Dim Freshitem As String = ""
            Dim ProductItem As String = ""

            Freshitem = "Select max(TSPL_ITEM_MASTER.Short_Description)Fresh_Item,max(TSPL_ITEM_MASTER.Item_Desc)Item_Desc,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Print_Sequence)Print_Sequence from TSPL_DEMAND_BOOKING_DETAIL
LEFT OUTER JOIN TSPL_DEMAND_BOOKING_MASTER ON TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "' AND ((TSPL_ITEM_MASTER.Is_FreshItem = 1 ) or (TSPL_ITEM_MASTER.Is_FreshItem = 1 and Is_CrateType = 1))
 group by TSPL_ITEM_MASTER.Item_Code ORDER BY Sku_Seq "
            ProductItem = "Select max(TSPL_ITEM_MASTER.Short_Description)Product_Item,max(TSPL_ITEM_MASTER.Item_Desc)Item_Desc,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Print_Sequence)Print_Sequence from TSPL_DEMAND_BOOKING_DETAIL
LEFT OUTER JOIN TSPL_DEMAND_BOOKING_MASTER ON TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "' AND TSPL_ITEM_MASTER.Is_Ambient = 1 
 group by TSPL_ITEM_MASTER.Item_Code ORDER BY Sku_Seq "



            Dim BaseItemQry As String = "Select TSPL_ITEM_MASTER.Item_Code,max(TSPL_ITEM_MASTER.Short_Description)Short_Description,max(TSPL_ITEM_MASTER.Item_Desc)Item_Description,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Alies_Name2)Alies_Name2,MAX(TSPL_ITEM_MASTER.Alies_Name3)Alies_Name3,Convert(varchar,MAX(TSPL_ITEM_MASTER.Print_Sequence))Print_Sequence,Sum(TSPL_DEMAND_BOOKING_DETAIL.Qty)Qty from TSPL_DEMAND_BOOKING_DETAIL
LEFT OUTER JOIN TSPL_DEMAND_BOOKING_MASTER ON TSPL_DEMAND_BOOKING_MASTER.DOCUMENT_NO=TSPL_DEMAND_BOOKING_DETAIL.DOCUMENT_NO
LEFT OUTER JOIN TSPL_ITEM_MASTER ON TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
WHERE TSPL_DEMAND_BOOKING_MASTER.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "'  
 group by TSPL_ITEM_MASTER.Item_Code " ' ORDER BY Sku_Seq"
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                BaseItemQry += " union
 Select TSPL_ITEM_MASTER.Item_Code,max(TSPL_ITEM_MASTER.Short_Description)Short_Description,max(TSPL_ITEM_MASTER.Item_Desc)Item_Description,max(TSPL_ITEM_MASTER.Sku_Seq)Sku_Seq,MAX(TSPL_ITEM_MASTER.Alies_Name2)Alies_Name2,MAX(TSPL_ITEM_MASTER.Alies_Name3)Alies_Name3,
Convert(varchar,MAX(TSPL_ITEM_MASTER.Print_Sequence))Print_Sequence,0 As Qty from TSPL_ITEM_MASTER
LEFT OUTER JOIN TSPL_DEMAND_BOOKING_DETAIL On TSPL_DEMAND_BOOKING_DETAIL.Item_Code=TSPL_ITEM_MASTER.Item_Code
LEFT OUTER JOIN TSPL_DEMAND_BOOKING_MASTER ON TSPL_DEMAND_BOOKING_MASTER.DOCUMENT_NO=TSPL_DEMAND_BOOKING_DETAIL.DOCUMENT_NO
WHERE TSPL_ITEM_MASTER.Print_Sequence is not null and TSPL_ITEM_MASTER.Active=1
 group by TSPL_ITEM_MASTER.Item_Code "
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                itemqry = "Select Max(Short_Description)Short_Description,Max(Item_Description)Item_Description,Max(Sku_Seq)Sku_Seq,Max(Alies_Name2)Alies_Name2,Max(Alies_Name3)Alies_Name3,Max(Print_Sequence)Print_Sequence,Sum(Qty)Qty from (" & BaseItemQry & ") xyz Group By Item_Code Order By Sku_Seq"
            Else
                itemqry = "Select * from (" & BaseItemQry & ") xyz Order By Sku_Seq"
            End If

            Dim sbitemName2 As New StringBuilder()
            Dim sbitemName1 As New StringBuilder()
            Dim sbitemNames1 As New StringBuilder()
            Dim sbitemNames2 As New StringBuilder()
            Dim sbitemNamesQty As New StringBuilder()
            Dim sbitemNamesAmt As New StringBuilder()
            Dim sbFinalItemNamesQty As New StringBuilder()
            Dim sbFinalItemNamesAmt As New StringBuilder()
            Dim sbProductIemName As New StringBuilder()
            Dim sbFreshItemName As New StringBuilder()
            Dim sbFreshItemsName As New StringBuilder()
            Dim sbProductIemsName As New StringBuilder()
            Dim sbitemNamesFresh As New StringBuilder()
            Dim sbitemNamesProduct As New StringBuilder()
            Dim sbFreshItemNameMax As New StringBuilder()
            Dim sbProductItemNameMax As New StringBuilder()


            Dim itemName2 As String = Nothing
            Dim itemName1 As String = Nothing
            Dim itemNames1 As String = Nothing
            Dim itemNames2 As String = Nothing
            Dim itemNamesQty As String = Nothing
            Dim itemNamesAmt As String = Nothing
            Dim FinalItemNamesQty As String = Nothing
            Dim FinalItemNamesAmt As String = Nothing
            Dim ProductIemName As String = Nothing
            Dim FreshItemName As String = Nothing
            Dim FreshItemsName As String = Nothing
            Dim ProductIemsName As String = Nothing
            'Dim ProductItemsAmt As String = Nothing
            'Dim ItemSubGroup As String = Nothing
            'Dim ItemSubGroupAvg As String = Nothing
            'Dim ItemsSubGroup As String = Nothing
            Dim itemNamesFresh As String = Nothing
            Dim itemNamesProduct As String = Nothing
            Dim FreshItemNameMax As String = Nothing
            Dim ProductItemNameMax As String = Nothing
            Dim dtitemName As DataTable = clsDBFuncationality.GetDataTable(itemqry)
            Dim dtFresh As DataTable = clsDBFuncationality.GetDataTable(Freshitem)
            Dim dtProduct As DataTable = clsDBFuncationality.GetDataTable(ProductItem)

            If dtitemName.Rows.Count > 0 Then
                For i As Integer = 0 To dtitemName.Rows.Count - 1
                    'itemName1 += "Sum(IsNull([" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "],0)) As [" + clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) + "]" + ","
                    sbitemName2.Append("Sum(IsNull([" & clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) & "],0)) As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "1]" & ",")
                    sbFinalItemNamesQty.Append("SUM(XXFINAL.[" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "]) As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "]" & ",")
                    sbFinalItemNamesAmt.Append("SUM(XXFINAL.[" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "1]) As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "1]" & ",")

                    If i = 0 Then
                        sbitemNamesQty.Append("ISNULL([" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "],0)")
                        sbitemNamesAmt.Append("ISNULL([" & clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) & "],0)")
                        sbitemNames1.Append("[" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "] ")
                        sbitemNames2.Append("[" & clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) & "] ")
                        If clsCommon.myCDecimal(dtitemName.Rows(i)("Qty")) > 0 Then
                            sbitemName1.Append("Sum(IsNull([" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "],0)) As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "]")
                        Else
                            sbitemName1.Append("0 As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "]")
                        End If
                    Else
                        sbitemNamesQty.Append("+" & "ISNULL([" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "],0)")
                        sbitemNamesAmt.Append("+" & "ISNULL([" & clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) & "],0)")
                        sbitemNames1.Append(", [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "] ")
                        sbitemNames2.Append(", [" & clsCommon.myCstr(dtitemName.Rows(i)("Item_Description")) & "] ")
                        If clsCommon.myCDecimal(dtitemName.Rows(i)("Qty")) > 0 Then
                            sbitemName1.Append(", Sum(IsNull([" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "],0)) As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "]")
                        Else
                            sbitemName1.Append(", 0 As [" & clsCommon.myCstr(dtitemName.Rows(i)("Short_Description")) & "]")
                        End If
                    End If
                Next
            End If
            If dtFresh.Rows.Count > 0 Then
                For i As Integer = 0 To dtFresh.Rows.Count - 1
                    sbFreshItemName.Append(" Sum(IsNull([" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "],0)) As [" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "]" & ",")
                    sbFreshItemNameMax.Append("max(IsNull([" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "],0)) As [" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "]" & ",")
                    If i = 0 Then
                        sbitemNamesFresh.Append("ISNULL([" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "],0)")
                        sbFreshItemsName.Append("[" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "] ")
                    Else
                        sbitemNamesFresh.Append("+" & "ISNULL([" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "],0)")
                        sbFreshItemsName.Append(", [" & clsCommon.myCstr(dtFresh.Rows(i)("Fresh_Item")) & "] ")
                    End If
                Next
            End If
            If dtProduct.Rows.Count > 0 Then
                For i As Integer = 0 To dtProduct.Rows.Count - 1
                    sbProductIemName.Append("Sum(IsNull([" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "],0)) As [" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "]" & ",")
                    sbProductItemNameMax.Append(" max(IsNull([" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "],0)) As [" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "]" & ",")
                    If i = 0 Then
                        sbitemNamesProduct.Append("ISNULL([" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "],0)")
                        sbProductIemsName.Append("[" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "] ")
                    Else
                        sbitemNamesProduct.Append("+" & "ISNULL([" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "],0)")

                        sbProductIemsName.Append(", [" & clsCommon.myCstr(dtProduct.Rows(i)("Product_Item")) & "] ")
                    End If
                Next
            End If
            itemName2 = sbitemName2.ToString()
            itemName1 = sbitemName1.ToString()
            itemNames1 = sbitemNames1.ToString()
            itemNames2 = sbitemNames2.ToString()
            itemNamesQty = sbitemNamesQty.ToString()
            itemNamesAmt = sbitemNamesAmt.ToString()
            FinalItemNamesQty = sbFinalItemNamesQty.ToString()
            FinalItemNamesAmt = sbFinalItemNamesAmt.ToString()
            ProductIemName = sbProductIemName.ToString
            FreshItemName = sbFreshItemName.ToString()
            FreshItemsName = sbFreshItemsName.ToString()
            ProductIemsName = sbProductIemsName.ToString()
            itemNamesFresh = sbitemNamesFresh.ToString()
            itemNamesProduct = sbitemNamesProduct.ToString()
            FreshItemNameMax = sbFreshItemNameMax.ToString()
            ProductItemNameMax = sbProductItemNameMax.ToString()

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                qry = "Select Convert(Varchar,ROW_NUMBER() Over (Order By (Select 1))) As [SR.],max(Customer_Name)OUTLET,max(Display_Seq)as Display_Seq, " & itemName1 & ",sum(ItemNetAmount) as Amount from (select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name,max(XXFinal.Display_Seq) as Display_Seq, max(XXFinal.Short_Description) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount,sum(LTR_QTY)LTR_QTY,sum(KG_QTY)KG_QTY,max(Fresh_Item)Fresh_Item,max(Product_Item)Product_Item

from (select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,IsNull(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name,isnull(TSPL_CUSTOMER_MASTER.Display_Seq,0) as Display_Seq,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_DEMAND_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No ,
  Convert(Decimal(18,2),(isnull(TSPL_DEMAND_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY ,
  Convert(Decimal(18,2),(isnull(TSPL_DEMAND_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY1,
  Convert(Decimal(18,2),(isnull(TSPL_DEMAND_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[LTR]) as LTR_QTY,
case when TSPL_ITEM_MASTER.Is_Ambient = 1 then TSPL_ITEM_MASTER.Short_Description  end AS Product_Item,
				case when (TSPL_ITEM_MASTER.Is_FreshItem = 1  ) or (TSPL_ITEM_MASTER.Is_FreshItem = 1 and Is_CrateType = 1) then  TSPL_ITEM_MASTER.Short_Description end as Fresh_Item
    from  TSPL_DEMAND_BOOKING_DETAIL
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" & objCommonVar.CurrComp_Code1 & "'
  left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) I ON TSPL_ITEM_MASTER.Item_Code=I.item_code 

where IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N' And TSPL_DEMAND_BOOKING_MASTER.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "'
Union
select TSPL_CUSTOMER_MASTER.Cust_Code,
IsNull(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name,isnull(TSPL_CUSTOMER_MASTER.Display_Seq,0) as Display_Seq,'' As Item_Code,'' As Short_Description
,0 As Sku_Seq,0 As Qty,0 As ItemNetAmount,'' As Unit_code ,'' As ShiftType,
   TSPL_ROUTE_MASTER.Route_No, TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No ,
  0 as KG_QTY ,
  0 as KG_QTY1,
  0 as LTR_QTY,
'' AS Product_Item,
'' as Fresh_Item

from TSPL_CUSTOMER_MASTER
Left Outer Join TSPL_ROUTE_MASTER on TSPL_CUSTOMER_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
Left Join TSPL_VEHICLE_MASTER on TSPL_ROUTE_MASTER.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id
LEFT Outer Join (Select Document_No,Document_Date,Route_No,Max(ShiftType)ShiftType from TSPL_DEMAND_BOOKING_MASTER 
--Where CONVERT(date,TSPL_DEMAND_BOOKING_MASTER.Document_Date,103)='01-Apr-2024'
Group By Document_No,Document_Date,Route_No )TSPL_DEMAND_BOOKING_MASTER On TSPL_DEMAND_BOOKING_MASTER.Route_No=TSPL_ROUTE_MASTER.Route_No
Left Outer Join TSPL_DEMAND_BOOKING_DETAIL On TSPL_DEMAND_BOOKING_DETAIL.Document_No=TSPL_DEMAND_BOOKING_MASTER.Document_No And TSPL_DEMAND_BOOKING_DETAIL.Cust_Code=TSPL_CUSTOMER_MASTER.Cust_Code
--Left Join TSPL_SD_SHIPMENT_BOOKING_DETAIL ON TSPL_SD_SHIPMENT_BOOKING_DETAIL.Booking_TR_Code=TSPL_DEMAND_BOOKING_DETAIL.TR_Code  
Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" & objCommonVar.CurrComp_Code1 & "'
----left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) I ON TSPL_ITEM_MASTER.Item_Code=I.item_code 

where IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N' and TSPL_CUSTOMER_MASTER.Status='N'  And 
TSPL_ROUTE_MASTER.Route_No='" & clsCommon.myCstr(txtRouteNo.Value) & "'
)XXFinal
group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code )XXXX "

                If dtFresh.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(LTR_QTY)  For Fresh_Item In (" & FreshItemsName & ") ) As pivot_fresh "
                End If
                If dtProduct.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(KG_QTY)   For Product_Item In (" & ProductIemsName & ") ) As  pivot_Product "
                End If
                qry += " group by Cust_Code "

                qry += " Union all 
                       Select '' As [SR.],'TOTAL QNTY' as OUTLET,100000 as Display_Seq ," & itemName1 & " ,sum(Amount) as Amount
from (Select 1 AS Sno,Cust_Code,max(Customer_Name)Customer_Name,max(Display_Seq)as Display_Seq, " & itemName1 & " ,sum(ItemNetAmount) as Amount from (select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name,max(XXFinal.Display_Seq) as Display_Seq, max(XXFinal.Short_Description) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount,sum(LTR_QTY)LTR_QTY,sum(KG_QTY)KG_QTY,max(Fresh_Item)Fresh_Item,max(Product_Item)Product_Item

from (select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,IsNull(TSPL_CUSTOMER_MASTER.Customer_Name_Hindi,TSPL_CUSTOMER_MASTER.Customer_Name)Customer_Name, isnull(TSPL_CUSTOMER_MASTER.Display_Seq,0)as Display_Seq, TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_DEMAND_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No ,
Convert(Decimal(18,2),(isnull(TSPL_DEMAND_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY ,
  Convert(Decimal(18,2),(isnull(TSPL_DEMAND_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[KG]) as KG_QTY1,
  Convert(Decimal(18,2),(isnull(TSPL_DEMAND_BOOKING_DETAIL.qty,0) *isnull(TSPL_ITEM_UOM_DETAIL.Conversion_Factor,1))/I.[LTR]) as LTR_QTY,
case when TSPL_ITEM_MASTER.Is_Ambient = 1  then TSPL_ITEM_MASTER.Short_Description  end AS Product_Item,
				case when (TSPL_ITEM_MASTER.Is_FreshItem = 1 ) or (TSPL_ITEM_MASTER.Is_FreshItem = 1 and Is_CrateType = 1) then  TSPL_ITEM_MASTER.Short_Description end as Fresh_Item
    from TSPL_DEMAND_BOOKING_DETAIL 
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" & objCommonVar.CurrComp_Code1 & "'
  left join (  SELECT * FROM ( select item_code,uom_code,conversion_factor from TSPL_ITEM_UOM_DETAIL) I  PIVOT (Max(conversion_factor) FOR uom_code IN ( [KG],[LTR] )) P ) I ON TSPL_ITEM_MASTER.Item_Code = I.item_code 

where IsNull(TSPL_CUSTOMER_MASTER.Credit_Customer,'N')='N' And TSPL_DEMAND_BOOKING_MASTER.Document_No='" & clsCommon.myCstr(txtDocNo.Value) & "'
)XXFinal
group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code )XXXX "
                If dtFresh.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(LTR_QTY)  For Fresh_Item In (" & FreshItemsName & ") ) As pivot_fresh "
                End If
                If dtProduct.Rows.Count > 0 Then
                    qry += " PIVOT (SUM(KG_QTY)   For Product_Item In (" & ProductIemsName & ") ) As  pivot_Product "
                End If

                qry += " group by Cust_Code )XX group by SNo "

                qry = "select * from (" & qry & ") XXXFinal order by Display_Seq "
            Else
                qry = "select XXFinal.Cust_Code as Cust_Code,max(XXFinal.Customer_Name) as Customer_Name, max(XXFinal.Short_Description) +' '+max(XXFinal.Unit_code) as Short_Description,
sum(XXFinal.Qty) as Qty,sum(XXFinal.ItemNetAmount) as ItemNetAmount
from (select TSPL_DEMAND_BOOKING_DETAIL.Cust_Code,TSPL_CUSTOMER_MASTER.Customer_Name,TSPL_DEMAND_BOOKING_DETAIL.Item_Code,TSPL_ITEM_MASTER.Short_Description,TSPL_ITEM_MASTER.Sku_Seq,
TSPL_DEMAND_BOOKING_DETAIL.Qty,TSPL_DEMAND_BOOKING_DETAIL.ItemNetAmount,TSPL_DEMAND_BOOKING_DETAIL.Unit_code ,TSPL_DEMAND_BOOKING_DETAIL.ShiftType,
   TSPL_DEMAND_BOOKING_MASTER.Route_No,   TSPL_ROUTE_MASTER.Route_Desc,    TSPL_COMPANY_MASTER.Comp_Name  as CompanyName,  TSPL_TRANSPORT_MASTER.Transporter_Name as TranspoterName, 
  TSPL_VEHICLE_MASTER.DriverName,TSPL_VEHICLE_MASTER.Number as Vehicle_No from  TSPL_DEMAND_BOOKING_DETAIL 
left outer join TSPL_DEMAND_BOOKING_MASTER on TSPL_DEMAND_BOOKING_MASTER.Document_No=TSPL_DEMAND_BOOKING_DETAIL.Document_No
left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code=TSPL_DEMAND_BOOKING_DETAIL.Item_Code
left outer join TSPL_CUSTOMER_MASTER  on TSPL_CUSTOMER_MASTER.Cust_Code=TSPL_DEMAND_BOOKING_DETAIL.Cust_Code
  Left Join TSPL_ITEM_UOM_DETAIL on TSPL_ITEM_UOM_DETAIL.Item_Code = TSPL_ITEM_MASTER.Item_Code 
  And TSPL_ITEM_UOM_DETAIL.UOM_Code = TSPL_DEMAND_BOOKING_DETAIL.Unit_code  
  Left Join TSPL_VEHICLE_MASTER on TSPL_DEMAND_BOOKING_DETAIL.Vehicle_Code = TSPL_VEHICLE_MASTER.Vehicle_Id 
  Left Join TSPL_ROUTE_MASTER on TSPL_DEMAND_BOOKING_MASTER.Route_No = TSPL_ROUTE_MASTER.Route_No 
  Left Join TSPL_TRANSPORT_MASTER on TSPL_VEHICLE_MASTER.Transport_Id = TSPL_TRANSPORT_MASTER.Transport_Id 
  Left Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" & objCommonVar.CurrComp_Code1 & "'

where TSPL_DEMAND_BOOKING_MASTER.DOCUMENT_NO='" & clsCommon.myCstr(txtDocNo.Value) & "'
)XXFinal
group by XXFinal.Cust_Code,XXFinal.Item_Code,XXFinal.Sku_Seq,XXFinal.Unit_code "

            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

            If rbtnMorning.IsChecked Then
                ShiftType = "Morning"
            Else
                ShiftType = "Evening"
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                Dim dr As DataRow = dt.NewRow
                dr("OUTLET") = clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy") & ", " & ShiftType
                For ii As Integer = 0 To dtitemName.Rows.Count - 1
                    'dr(clsCommon.myCstr(dtitemName.Rows(ii)("Short_Description"))) = clsCommon.myCstr(dtitemName.Rows(ii)("Print_Sequence"))
                    'Dim dr As DataRow = dt.NewRow
                    Dim colName As String = clsCommon.myCstr(dtitemName.Rows(ii)("Short_Description"))
                    Dim value As Decimal = clsCommon.myCDecimal(dtitemName.Rows(ii)("Print_Sequence"))

                    ' Check if value is numeric before assigning
                    If IsNumeric(value) AndAlso value > 0 Then
                        dr(colName) = clsCommon.myCDecimal(value)
                    Else
                        dr(colName) = DBNull.Value ' Or handle accordingly
                    End If
                Next
                dt.Rows.InsertAt(dr, 0)
                dt.AcceptChanges()
            End If


            MyRadGridView1.DataSource = Nothing
            MyRadGridView1.Rows.Clear()
            MyRadGridView1.Columns.Clear()
            MyRadGridView1.GroupDescriptors.Clear()
            MyRadGridView1.MasterView.Refresh()
            MyRadGridView1.GroupDescriptors.Clear()
            MyRadGridView1.EnableFiltering = True
            MyRadGridView1.MasterTemplate.SummaryRowsBottom.Clear()
            If dt.Rows.Count > 0 Then

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    ' Create a new DataTable to store converted data
                    Dim dtConverted As New DataTable()
                    ' Convert all columns to String type
                    For Each col As DataColumn In dt.Columns
                        If clsCommon.CompairString(clsCommon.myCstr(col.ColumnName), "Display_Seq") = CompairStringResult.Equal Then
                            Continue For
                        End If
                        dtConverted.Columns.Add(col.ColumnName, GetType(String))
                    Next
                    ' Copy data with replacements
                    For Each row As DataRow In dt.Rows
                        Dim newRow As DataRow = dtConverted.NewRow()

                        For Each col As DataColumn In dt.Columns
                            If clsCommon.CompairString(clsCommon.myCstr(col.ColumnName), "Display_Seq") = CompairStringResult.Equal Then
                                Continue For
                            End If
                            Dim cellValue As Object = row(col)
                            ' If numeric, check for 0 and replace
                            If IsNumeric(cellValue) Then
                                If Convert.ToDouble(cellValue) = 0 Then
                                    'newRow(col.ColumnName) = "-" ' Replace 0 with "-"
                                Else
                                    newRow(col.ColumnName) = cellValue.ToString() ' Convert to string
                                End If
                            Else
                                newRow(col.ColumnName) = cellValue.ToString() ' Convert non-numeric to string
                            End If
                        Next
                        dtConverted.Rows.Add(newRow)
                    Next
                    ' Bind the converted DataTable to RadGridView
                    MyRadGridView1.DataSource = dtConverted
                Else
                    MyRadGridView1.DataSource = dt
                End If
                MyRadGridView1.AutoSizeColumnsMode = GridViewAutoSizeColumnsMode.None
                MyRadGridView1.MasterTemplate.Refresh()


                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    ApplyFormattingManually()
                    For i As Integer = 0 To dtitemName.Rows.Count - 1
                        MyRadGridView1.Columns("" & clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) & "").FormatString = "{0:n2}"
                        If clsCommon.myLen(clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name2"))) > 0 AndAlso clsCommon.myLen(clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name3"))) > 0 Then
                            MyRadGridView1.Columns("" & clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) & "").HeaderText = clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name2")) & Environment.NewLine & clsCommon.myCstr(dtitemName.Rows(i).Item("Alies_Name3"))
                        Else
                            MyRadGridView1.Columns("" & clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) & "").HeaderText = clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description"))
                        End If
                        MyRadGridView1.Columns("" & clsCommon.myCstr(dtitemName.Rows(i).Item("Short_Description")) & "").HeaderTextAlignment = ContentAlignment.MiddleCenter
                    Next
                    MyRadGridView1.Columns("Amount").FormatString = "{0:n2}"
                End If
                MyRadGridView1.MasterTemplate.AutoExpandGroups = True
                MyRadGridView1.BestFitColumns()
            Else
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
#Disable Warning S3385 ' "Exit" statements should not be used
                Exit Sub
#Enable Warning S3385 ' "Exit" statements should not be used
            End If
            Dim arrHeader As List(Of String) = New List(Of String)()
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                arrHeader.Add("Supply Chart")
                arrHeader.Add("Transpoter : " & clsCommon.myCstr(lblTransporterName.Text) & "" & "     Date: " & clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy") & "   " & ShiftType & "   " & "Route :" & clsCommon.myCstr(lblRouteDesc.Text) & "    ")

                'arrHeader.Add(("Vehicle No: " + clsCommon.myCstr(lblVehicleDesc.Text) + "  "))
                'arrHeader.Add(("Shift Type: " + ShiftType + "  "))
                'arrHeader.Add(("Driver No: " + clsCommon.myCstr(txtDriverMobNo.Text) + "  "))
            Else
                arrHeader.Add("Name : " & clsDBFuncationality.getSingleValue("select program_name from tspl_program_Master where program_cODE='" & clsUserMgtCode.rptBookingQtyAmtReport & "'"))
                arrHeader.Add("Company : " & objCommonVar.CurrentCompanyName)
                arrHeader.Add(("Date: " & clsCommon.GetPrintDate(txtDate.Value, "dd/MM/yyyy") & "  "))
                arrHeader.Add(("Route No: " & clsCommon.myCstr(txtRouteNo.Value) & "  "))
                arrHeader.Add(("Route Name: " & clsCommon.myCstr(lblRouteDesc.Text) & "  "))
                arrHeader.Add(("Vehicle No: " & clsCommon.myCstr(lblVehicleNo.Text) & "  "))
                arrHeader.Add(("Shift Type: " & ShiftType & "  "))

                arrHeader.Add(("Transpoter Name: " & clsCommon.myCstr(lblTransporterName.Text) & "  "))
                arrHeader.Add(("Vehicle_No: " & clsCommon.myCstr(lblVehicleNo.Text) & "  "))
                'arrHeader.Add(("Driver: " + clsCommon.myCstr(txtDriverName.Text) + "  "))
                'arrHeader.Add(("Driver No: " + clsCommon.myCstr(txtDriverMobNo.Text) + "  "))

            End If
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    transportSql.exportdataBoothSlipGNG(Nothing, MyRadGridView1, "", "Supply Chart", 0, MyRadGridView1.Rows.Count, False, arrHeader, False, False, False, False, False, Nothing, True, True)
                Else
                    clsCommon.MyExportToExcelGrid("Supply Chart", MyRadGridView1, arrHeader, "Supply Chart")
                End If
            End If
        Catch ex As InvalidOperationException
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, RadMessageIcon.Error)
        End Try
    End Sub

    Public Sub ApplyFormattingManually()
        For Each row As GridViewRowInfo In MyRadGridView1.Rows
            For Each cell As GridViewCellInfo In row.Cells
                If cell IsNot Nothing AndAlso IsNumeric(cell.Value) Then
                    Dim value As Double = clsCommon.myCdbl(cell.Value)
                    If value > 0 Then
                        ' Apply formatting directly
                        If value = Math.Floor(value) Then
                            cell.Value = value.ToString("0") ' No decimals
                        ElseIf value * 10 = Math.Floor(value * 10) Then
                            cell.Value = value.ToString("0.0") ' One decimal place
                        Else
                            cell.Value = value.ToString("0.00") ' Two decimal places
                        End If
                    End If
                End If
            Next
        Next
    End Sub
    Private Sub rmi_BoothSlipExcel_Click(sender As Object, e As EventArgs) Handles rmi_BoothSlipExcel.Click
        Try
            BoothSlipExport()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnGVTruckSheet_Click(sender As Object, e As EventArgs) Handles btnGVTruckSheet.Click
        Try
            If btnGVTruckSheet.Text = "Show Truck Sheet" Then
                GVTruckSheet.Visible = True
                GVTruckSheet.Dock = DockStyle.Fill
                btnGVTruckSheet.Text = "Hide Truck Sheet"
            ElseIf btnGVTruckSheet.Text = "Hide Truck Sheet" Then
                GVTruckSheet.Visible = False
                GVTruckSheet.Dock = DockStyle.None
                btnGVTruckSheet.Text = "Show Truck Sheet"
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class
Public Class ItemValueClass
    Public itemCode As String = String.Empty
    Public itemDesc As String = String.Empty
    Public Unit_code As String = String.Empty
    Public IsFreshAmbient As String = String.Empty
    Public ShortDesc As String = String.Empty
    Public ItemRate As Double = 0
    Public ItemTotAmt As Double = 0
    Public FreshItem_QtyInCrates As Double = 0
    Public FreshItem_QtyInLitres As Double = 0
    Public Conversion_Factor As Double = 0
    Public Stocking_Unit As String = String.Empty
    Public TAX_Group As String = ""
    Public TAX1 As String = Nothing
    Public TAX1_Base_Amt As Double = 0
    Public TAX1_Rate As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Base_Amt As Double = 0
    Public TAX2_Rate As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Base_Amt As Double = 0
    Public TAX3_Rate As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Base_Amt As Double = 0
    Public TAX4_Rate As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Base_Amt As Double = 0
    Public TAX5_Rate As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Base_Amt As Double = 0
    Public TAX6_Rate As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Base_Amt As Double = 0
    Public TAX7_Rate As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Base_Amt As Double = 0
    Public TAX8_Rate As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Base_Amt As Double = 0
    Public TAX9_Rate As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Base_Amt As Double = 0
    Public TAX10_Rate As Double = 0
    Public TAX10_Amt As Double = 0
End Class
