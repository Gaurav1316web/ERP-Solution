Imports System.Data.SqlClient
Imports common
Imports System.IO

Public Class frmMPDCSIncentiveReco
    Inherits FrmMainTranScreen
#Region "Variables"
    Public Const colPKID As String = "colPKID"
    Public Const colSlNo As String = "SLNO"
    Public Const colRecoStatus As String = "colRecoStatus"
    Public Const colBMCCode As String = "colBMCCode"
    Public Const colBMCName As String = "colBMCName"
    Public Const colRouteCode As String = "colRouteCode"
    Public Const colRouteName As String = "colRouteName"
    Public Const colZoneCode As String = "colZoneCode"
    Public Const colZoneName As String = "colZoneName"
    Public Const colVLCCode As String = "colVLCCode"
    Public Const colVLCUploaderCode As String = "colVLCUploaderCode"
    Public Const colVLCName As String = "colVLCName"
    Public Const colYear As String = "Year"
    Public Const colMonth As String = "Month"
    Public Const colCycleNo As String = "Cycle No"
    Public Const colQty As String = "Qty"
    Public Const colEditQty As String = "colEditQty"
    Public Const colUOM As String = "colUOM"
    Public Const colFAT As String = "colFAT"
    Public Const colSNF As String = "colSNF"
    Public Const colAmount As String = "colAmount"
    Public Const colMPCount As String = "colMPCount"
    Public Const colMPQty As String = "colMPQty"
    Public Const colMPFAT As String = "colMPFAT"
    Public Const colMPSNF As String = "colMPSNF"
    Public Const colMPAmt As String = "colMPAmt"

    Public Const colDiffQty As String = "colDiffQty"
    Public Const colDiffFAT As String = "colDiffFAT"
    Public Const colDiffSNF As String = "colDiffSNF"
    Public Const colDiffAmt As String = "colDiffAmt"



    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Dim isCellValueChangedOpen As Boolean = False
    Dim IsinsideLoadData As Boolean = False
    Dim Qry As String
    Dim isFlag As Boolean = False
    Dim arrLoc As String = Nothing

    Dim strICode As String = ""
    Dim SettRefreshDBTReco As Boolean = False
    Dim SettDCSMPIncetiveReco As Boolean = False
    Dim SettMatchingUOM As String = ""
    Dim SettAllowMPQtyGreaterThanDCSQty As Boolean = False
    Dim SettAllowMPQtyLessThanDCSQty As Boolean = False
    Dim SettAllowMPQtyEqualToDCSQty As Boolean = False
    Dim SettMPIncentiveEntryApplyMonthly As Boolean = False
    Dim SettApplyZoneOnDBT As Boolean = False
    Dim SettPickMilkPurchaseInvoiceQtyOrRecoQty As Boolean = False
    Dim ApplyMonthlySetting As Boolean = False
    Dim SettDBTByDailyFarmerMilkCollection As Boolean = False
    Dim SettDCSQtyDecimalPlaces As Integer = 2
#End Region
    Public Sub New()
        InitializeComponent()
    End Sub
    Private Sub FrmVLCDataUploaderManual_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SettDBTByDailyFarmerMilkCollection = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DBTByDailyFarmerMilkCollection, clsFixedParameterCode.DBTByDailyFarmerMilkCollection, Nothing)) > 0)
        SettMPIncentiveEntryApplyMonthly = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, Nothing)) > 0)
        SettDCSQtyDecimalPlaces = clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DCSQtyDecimalPlaces, clsFixedParameterCode.DCSQtyDecimalPlaces, Nothing))
        UcAttachment1.Form_ID = MyBase.Form_ID
        UcAttachment1.isDeleteTheAttachment = False
        UcAttachment1.settAutoAttachment = True

        strICode = clsFixedParameter.GetData(clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, Nothing)
        SettRefreshDBTReco = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.RefreshDBTReco, clsFixedParameterCode.RefreshDBTReco, Nothing)) > 0)
        SettApplyZoneOnDBT = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.ApplyZoneInDBT, clsFixedParameterCode.ApplyZoneInDBT, Nothing)) > 0)
        txtZone.MendatroryField = SettApplyZoneOnDBT
        SettDCSMPIncetiveReco = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MandatoryDCSMPIncetiveReco, Nothing)) = 1)
        SettMatchingUOM = clsFixedParameter.GetData(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.MatchingUOM, Nothing)
        SettPickMilkPurchaseInvoiceQtyOrRecoQty = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.PickMilkPurchaseInvoiceQtyOrRecoQty, clsFixedParameterCode.PickMilkPurchaseInvoiceQtyOrRecoQty, Nothing)) = 1)
        SettAllowMPQtyGreaterThanDCSQty = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyGreaterThanDCSQty, Nothing)) = 1)
        SettAllowMPQtyLessThanDCSQty = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyLessThanDCSQty, Nothing)) = 1)
        SettAllowMPQtyEqualToDCSQty = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.DCSRecoCondition, clsFixedParameterCode.AllowMPQtyEqualToDCSQty, Nothing)) = 1)
        ApplyMonthlySetting = (clsCommon.myCDecimal(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, Nothing)) = 1)

        SetUserMgmtNew()
        Reset()

        RadButton4.Visible = SettPickMilkPurchaseInvoiceQtyOrRecoQty
        RadGroupBox1.Visible = SettPickMilkPurchaseInvoiceQtyOrRecoQty
        RadButton1.Visible = Not SettPickMilkPurchaseInvoiceQtyOrRecoQty
        RadButton2.Visible = Not SettPickMilkPurchaseInvoiceQtyOrRecoQty
        ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Transaction")
        ButtonToolTip.SetToolTip(btndelete, "Press Alt+D Delete Transaction")
        ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
        ButtonToolTip.SetToolTip(btnReset, "Press Alt+N New Transaction")
        ButtonToolTip.SetToolTip(btnPost, "Press Alt+P Post Trasnaction")
        If SettApplyZoneOnDBT Then
            SplitContainer3.Panel1Collapsed = False
            SplitContainer3.Panel2Collapsed = True
        ElseIf clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
            SplitContainer3.Panel1Collapsed = True
            SplitContainer3.Panel2Collapsed = False

            RadButton4.Visible = False
            RadGroupBox1.Visible = False
            RadButton1.Visible = False
            RadButton2.Visible = False
        Else
            SplitContainer3.Panel1Collapsed = False
            SplitContainer3.Panel2Collapsed = True
        End If

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
            SaveData(False)
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btndelete.Enabled Then
            DeleteData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnclose.Enabled Then
            CloseForm()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso btnclose.Enabled Then
            'PostData()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then
                ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                      "========Table Name=========" + Environment.NewLine +
                      "TSPL_DCS_MP_INCENTIVE_RECO_HEAD" + Environment.NewLine +
                      "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL" + Environment.NewLine)
                Dim frm As New FrmPWD(Nothing)
                frm.strType = "MulProcDedReversAndCreate"
                frm.strCode = "MulProcDedReversAndCreate"
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverse.Visible = True
                    btnEditDCSQty.Visible = True
                    Try
                        gvItem.Columns(colEditQty).ReadOnly = False
                        gvItem.Columns(colEditQty).IsVisible = True
                    Catch ex As Exception
                    End Try
                End If
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
            End If
        End If
    End Sub
    Private Sub RadButton2_Click(sender As Object, e As EventArgs) Handles RadButton2.Click
        Try
            Qry = "select TSPL_VLC_MASTER_HEAD.MCC as [BMC Code],TSPL_MCC_MASTER.MCC_NAME as  [BMC Name],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as  [DCS Code],TSPL_VLC_MASTER_HEAD.VLC_Name as  [DCS Name],DATEPART(YEAR, GETDATE()) as [Year],DATEPART(MONTH, GETDATE()) [Month],0 as [Cycle No],0 as Qty,'' as UOM,0 as FAT,0 as  SNF,0 as  Amount
from TSPL_VLC_MASTER_HEAD
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC "
            transportSql.ExporttoExcel(Qry, "", "[BMC Name],[DCS Code]", Me)
        Catch ex As Exception
            clsCommon.ProgressBarHide()
            Throw New Exception(ex.Message)
        End Try
    End Sub
    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        rmimport_Click()
    End Sub
    Private Sub rmimport_Click()

        UcAttachment1.BlankAllControls()
        Dim dgv As New RadGridView
        Dim gvExcel As New RadGridView
        Me.Controls.Add(dgv)
        Me.Controls.Add(gvExcel)
        Dim ii As Integer = 0

        'ii = ii + 1

        Dim dtError As New DataTable
        dtError.Columns.Add("DCS Code", GetType(String))
        dtError.Columns.Add("Year", GetType(String))
        dtError.Columns.Add("Month", GetType(String))
        dtError.Columns.Add("Cycle No", GetType(String))
        dtError.Columns.Add("Qty", GetType(String))
        dtError.Columns.Add("Error", GetType(String))

        Dim result As DialogResult = clsCommon.MyMessageBoxShow(Me, "Do you want to use the customize Format for importing?", "Import Format", MessageBoxButtons.YesNo)
        'Dim gvExcel As New RadGridView
        If result = DialogResult.Yes Then

            Dim Filename As String = ""
            Dim SafeFileName As String = ""

            If transportSql.importExcel(Filename, SafeFileName, gvExcel) Then
                Dim dt As New DataTable()
                dt.Columns.Add("Code", GetType(String))
                For Each column As GridViewDataColumn In gvExcel.Columns
                    Dim dr As DataRow = dt.NewRow()
                    dr("Code") = column.Name
                    dt.Rows.Add(dr)
                Next

                Dim frm As New Excelform
                frm.comboDataSource = dt
                frm.ShowDialog()

                If frm.collReturn IsNot Nothing AndAlso frm.collReturn.Count > 0 Then
                    ''Do you import function.

                    Try

                        If Not SettDCSMPIncetiveReco Then
                            Throw New Exception("This facility is not for you")
                        End If
                        UcAttachment1.AddAttachment(Filename, SafeFileName)
                        For ii = 0 To gvExcel.Rows.Count - 1
                            Try

                                If clsCommon.myLen(gvExcel.Rows(ii).Cells(frm.collReturn.Item("DCS Code")).Value) > 0 AndAlso clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Year")).Value) > 0 AndAlso clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Month")).Value) > 0 AndAlso clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Cycle No")).Value) > 0 AndAlso clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Qty")).Value) > 0 Then
                                    Qry = "select TSPL_VLC_MASTER_HEAD.MCC,TSPL_MCC_MASTER.MCC_NAME,TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VENDOR_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description as Zone_Name , TSPL_Mcc_UOM_DETAIL.UOM_Code as UOM from TSPL_VLC_MASTER_HEAD 
                                           left outer join TSPL_Mcc_UOM_DETAIL on TSPL_Mcc_UOM_DETAIL.MCC_CODE = TSPL_VLC_MASTER_HEAD.MCC and TSPL_Mcc_UOM_DETAIL.Stocking_Unit = 'Y'
                                           left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VLC_MASTER_HEAD.VSP_Code
                                           left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
                                           left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
                                           where TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader ='" + clsCommon.myCstr(gvExcel.Rows(ii).Cells(frm.collReturn.Item("DCS Code")).Value) + "'"

                                    Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(Qry)
                                    If dt1 Is Nothing OrElse dt1.Rows.Count <= 0 Then
                                        Throw New Exception("Invalid DCS [" + clsCommon.myCstr(gvExcel.Rows(ii).Cells(frm.collReturn.Item("DCS Code")).Value) + "]")
                                    End If
                                    If clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Year")).Value) <> txtFromDate.Value.Year Then
                                        Throw New Exception("Year is Mismatched")
                                    End If
                                    If clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Month")).Value) <> txtFromDate.Value.Month Then
                                        Throw New Exception("Month is Mismatched")
                                    End If
                                    If clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Cycle No")).Value) <> CycleNo Then
                                        Throw New Exception("Cycle No is Mismatched")
                                    End If
                                    Qry = "select PK_Id from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL  where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code not in ('" + txtDocumentNo.Value + "') and Cycle_Year='" + clsCommon.myCstr(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Year")).Value) + "' and Cycle_Month='" + clsCommon.myCstr(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Month")).Value) + "' and Cycle_No='" + clsCommon.myCstr(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Cycle No")).Value) + "' and  VLC_Code='" + clsCommon.myCstr(dt1.Rows(0)("VLC_Code")) + "'"
                                    Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable(Qry)
                                    If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                                        Continue For
                                    End If


                                    gvItem.Rows.AddNew()
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSlNo).Value = gvItem.Rows.Count

                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCCode).Value = clsCommon.myCstr(dt1.Rows(0)("MCC"))
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCName).Value = clsCommon.myCstr(dt1.Rows(0)("MCC_NAME"))

                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneCode).Value = clsCommon.myCstr(dt1.Rows(0)("Zone_Code"))
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneName).Value = clsCommon.myCstr(dt1.Rows(0)("Zone_Name"))

                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCCode).Value = clsCommon.myCstr(dt1.Rows(0)("VLC_Code"))
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCUploaderCode).Value = clsCommon.myCstr(dt1.Rows(0)("VLC_Code_VLC_Uploader"))
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCName).Value = clsCommon.myCstr(dt1.Rows(0)("VLC_Name"))

                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colYear).Value = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Year")).Value)
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMonth).Value = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Month")).Value)
                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colCycleNo).Value = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Cycle No")).Value)

                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Qty")).Value)



                                    If clsCommon.myLen(dt1.Rows(0)("UOM")) <= 0 Then
                                        Throw New Exception("Please define UOM")
                                    End If

                                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colUOM).Value = clsCommon.myCstr(dt1.Rows(0)("UOM"))
                                    'gvItem.Rows(gvItem.Rows.Count - 1).Cells(colFAT).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("FAT").Value)
                                    'gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSNF).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("SNF").Value)
                                    'gvItem.Rows(gvItem.Rows.Count - 1).Cells(colAmount).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("Amount").Value)

                                    FillFarmerInfo(gvItem.Rows.Count - 1)
                                End If
                            Catch ex As Exception

                                Dim dr As DataRow = dtError.NewRow()
                                dr("DCS Code") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("DCS Code")).Value)
                                dr("Year") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Year")).Value)
                                dr("Month") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Month")).Value)
                                dr("Cycle No") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Cycle No")).Value)
                                dr("Qty") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Qty")).Value)
                                dr("Error") = "Error At Row No [" + clsCommon.myCstr(ii) + "] " + "Invalid DCS"
                                dtError.Rows.Add(dr)
                            End Try
                        Next
                        ' End If
                        'clsCommon.MyMessageBoxShow(Me, "Data Imported", Me.Text, MessageBoxButtons.OK)
                    Catch ex As Exception

                        'Dim dr As DataRow = dtError.NewRow()
                        'dr("DCS Code") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("DCS Code")).Value)
                        'dr("Year") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Year")).Value)
                        'dr("Month") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Month")).Value)
                        'dr("Cycle No") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Cycle No")).Value)
                        'dr("Qty") = clsCommon.myCDecimal(gvExcel.Rows(ii).Cells(frm.collReturn.Item("Qty")).Value)
                        'dr("Error") = "Error At Row No [" + clsCommon.myCstr(ii) + "] " + ex.Message
                        'dtError.Rows.Add(dr)

                        'loadBlankGrid()
                        'clsCommon.MyMessageBoxShow(Me, "Error at Row No [" + clsCommon.myCstr(ii) + "]" + Environment.NewLine + ex.Message, Me.Text)
                    Finally
                        Me.Controls.Remove(dgv)
                    End Try

                    Try

                        If dtError.Rows.Count > 0 Then
                            Dim ff As New FrmFreeGrid
                            ff.ReportID = "MULPROD"
                            ff.Text = "Multiple Deduction Fill Errors"
                            ff.dt = dtError
                            ff.ShowDialog()
                        Else
                            clsCommon.MyMessageBoxShow(Me, "Data Imported", Me.Text, MessageBoxButtons.OK)
                        End If

                    Catch ex As Exception
                        clsCommon.MyMessageBoxShow(Me, "Error at Row No [" + clsCommon.myCstr(ii) + "]" + Environment.NewLine + ex.Message, Me.Text)
                    End Try
                End If
            End If

        Else

            Try

                If Not SettDCSMPIncetiveReco Then
                    Throw New Exception("This facility is not for you")
                End If
                loadBlankGrid()
                Dim FileName As String = ""
                Dim SafeFileName As String = ""
                If transportSql.importExcel(FileName, SafeFileName, dgv, "BMC Code", "BMC Name", "DCS Code", "DCS Name", "Year", "Month", "Cycle No", "Qty", "UOM", "FAT", "SNF", "Amount") Then
                    UcAttachment1.AddAttachment(FileName, SafeFileName)
                    For ii = 0 To dgv.Rows.Count - 1
                        If clsCommon.myLen(dgv.Rows(ii).Cells("BMC Code").Value) > 0 AndAlso clsCommon.myLen(dgv.Rows(ii).Cells("DCS Code").Value) > 0 AndAlso clsCommon.myCDecimal(dgv.Rows(ii).Cells("Qty").Value) > 0 Then
                            Qry = "select TSPL_VLC_MASTER_HEAD.MCC,TSPL_MCC_MASTER.MCC_NAME,TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VENDOR_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description as Zone_Name from TSPL_VLC_MASTER_HEAD 
                        left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code=TSPL_VLC_MASTER_HEAD.VSP_Code
                        left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code
                        left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
                        where TSPL_VLC_MASTER_HEAD.MCC='" + clsCommon.myCstr(dgv.Rows(ii).Cells("BMC Code").Value) + "' and TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader='" + clsCommon.myCstr(dgv.Rows(ii).Cells("DCS Code").Value) + "'"
                            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
                            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                                Throw New Exception("Invalid combinition of BMC [" + clsCommon.myCstr(dgv.Rows(ii).Cells("BMC Code").Value) + "] and DCS [" + clsCommon.myCstr(dgv.Rows(ii).Cells("DCS Code").Value) + "]")
                            End If
                            If clsCommon.myCDecimal(dgv.Rows(ii).Cells("Year").Value) <> txtFromDate.Value.Year Then
                                Throw New Exception("Year is Mismatched")
                            End If
                            If clsCommon.myCDecimal(dgv.Rows(ii).Cells("Month").Value) <> txtFromDate.Value.Month Then
                                Throw New Exception("Month is Mismatched")
                            End If
                            If clsCommon.myCDecimal(dgv.Rows(ii).Cells("Cycle No").Value) <> CycleNo Then
                                Throw New Exception("Cycle No is Mismatched")
                            End If
                            Qry = "select PK_Id from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL  where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code not in ('" + txtDocumentNo.Value + "') and Cycle_Year='" + clsCommon.myCstr(dgv.Rows(ii).Cells("Year").Value) + "' and Cycle_Month='" + clsCommon.myCstr(dgv.Rows(ii).Cells("Month").Value) + "' and Cycle_No='" + clsCommon.myCstr(dgv.Rows(ii).Cells("Cycle No").Value) + "' and  VLC_Code='" + clsCommon.myCstr(dt.Rows(0)("VLC_Code")) + "'"
                            Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable(Qry)
                            If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                                Continue For
                            End If


                            gvItem.Rows.AddNew()
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSlNo).Value = gvItem.Rows.Count

                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCCode).Value = clsCommon.myCstr(dt.Rows(0)("MCC"))
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCName).Value = clsCommon.myCstr(dt.Rows(0)("MCC_NAME"))

                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneCode).Value = clsCommon.myCstr(dt.Rows(0)("Zone_Code"))
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneName).Value = clsCommon.myCstr(dt.Rows(0)("Zone_Name"))

                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCCode).Value = clsCommon.myCstr(dt.Rows(0)("VLC_Code"))
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCUploaderCode).Value = clsCommon.myCstr(dt.Rows(0)("VLC_Code_VLC_Uploader"))
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCName).Value = clsCommon.myCstr(dt.Rows(0)("VLC_Name"))

                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colYear).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("Year").Value)
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMonth).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("Month").Value)
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colCycleNo).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("Cycle No").Value)

                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("Qty").Value)
                            If clsCommon.myLen(dgv.Rows(ii).Cells("UOM").Value) <= 0 Then
                                Throw New Exception("Please define UOM")
                            End If
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colUOM).Value = clsCommon.myCstr(dgv.Rows(ii).Cells("UOM").Value)
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colFAT).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("FAT").Value)
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSNF).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("SNF").Value)
                            gvItem.Rows(gvItem.Rows.Count - 1).Cells(colAmount).Value = clsCommon.myCDecimal(dgv.Rows(ii).Cells("Amount").Value)

                            FillFarmerInfo(gvItem.Rows.Count - 1)
                        End If
                    Next
                End If
                clsCommon.MyMessageBoxShow(Me, "Data Imported", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                loadBlankGrid()
                clsCommon.MyMessageBoxShow(Me, "Error at Row No [" + clsCommon.myCstr(ii) + "]" + Environment.NewLine + ex.Message, Me.Text)
            Finally
                Me.Controls.Remove(dgv)
            End Try

        End If

    End Sub

    Private Sub FillFarmerInfo(ByVal Indx As Integer)
        Dim Qry As String = "select UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where Item_Code='" + strICode + "' and UOM_Code='" + clsCommon.myCstr(gvItem.Rows(Indx).Cells(colUOM).Value) + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            Throw New Exception("Invalid UOM [" + clsCommon.myCstr(gvItem.Rows(Indx).Cells(colUOM).Value) + "] for Item [" + strICode + "]")
        End If
        Dim MulConvFat As Decimal = clsCommon.myCDecimal(dt.Rows(0)("Conversion_Factor"))
        gvItem.Rows(Indx).Cells(colUOM).Value = clsCommon.myCstr(dt.Rows(0)("UOM_Code"))

        Qry = "select sum(isnull(Qty,0)) as Qty,sum(FAT_Kg) as FAT_Kg,sum(SNF_Kg) as SNF_Kg,sum(Amount) as Amount,count(MP_Code) as No_MP_Code from (select Qty,Amount,FAT_Kg,SNF_Kg,MP_Code from (
select  '" + strICode + "' as Item,TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,Qty,case when len(isnull(UOM,''))<=0 then 'KG' else UOM end as UOM ,Amount_Actual as Amount,isnull(TSPL_MP_INCENTIVE_ENTRY_DETAIL.FAT_Kg,0) as FAT_Kg,isnull(TSPL_MP_INCENTIVE_ENTRY_DETAIL.SNF_Kg,0) as SNF_Kg from TSPL_MP_INCENTIVE_ENTRY_DETAIL  where VLC_Code='" + clsCommon.myCstr(gvItem.Rows(Indx).Cells(colVLCCode).Value) + "' and Cycle_Month=" + clsCommon.myCstr(clsCommon.myCDecimal(gvItem.Rows(Indx).Cells(colMonth).Value)) + " and Cycle_Year=" + clsCommon.myCstr(clsCommon.myCDecimal(gvItem.Rows(Indx).Cells(colYear).Value)) + " and Cycle_No=" + clsCommon.myCstr(clsCommon.myCDecimal(gvItem.Rows(Indx).Cells(colCycleNo).Value)) + " 
)xx  )xxx"
        dt = clsDBFuncationality.GetDataTable(Qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            gvItem.Rows(Indx).Cells(colMPCount).Value = Math.Round(clsCommon.myCDecimal(dt.Rows(0)("No_MP_Code")), 0, MidpointRounding.ToEven)
            gvItem.Rows(Indx).Cells(colMPQty).Value = Math.Round(clsCommon.myCDecimal(dt.Rows(0)("Qty")), 2, MidpointRounding.ToEven)
            gvItem.Rows(Indx).Cells(colMPFAT).Value = Math.Round(clsCommon.myCDecimal(dt.Rows(0)("FAT_Kg")), 2, MidpointRounding.ToEven)
            gvItem.Rows(Indx).Cells(colMPSNF).Value = Math.Round(clsCommon.myCDecimal(dt.Rows(0)("SNF_Kg")), 2, MidpointRounding.ToEven)
            gvItem.Rows(Indx).Cells(colMPAmt).Value = Math.Round(clsCommon.myCDecimal(dt.Rows(0)("Amount")), 2, MidpointRounding.ToEven)
        End If
        If clsCommon.myLen(SettMatchingUOM) > 0 Then
            If Not clsCommon.CompairString(clsCommon.myCstr(gvItem.Rows(Indx).Cells(colUOM).Value), SettMatchingUOM) = CompairStringResult.Equal Then
                Qry = "select UOM_Code,Conversion_Factor from TSPL_ITEM_UOM_DETAIL where Item_Code='" + strICode + "' and UOM_Code='" + SettMatchingUOM + "'"
                dt = clsDBFuncationality.GetDataTable(Qry)
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    Throw New Exception("Invalid UOM [" + clsCommon.myCstr(gvItem.Rows(Indx).Cells(colUOM).Value) + "] for Item [" + strICode + "]")
                End If
                Dim DivConvFat As Decimal = clsCommon.myCDecimal(dt.Rows(0)("Conversion_Factor"))
                If DivConvFat <= 0 Then
                    Throw New Exception("Convertsion Factor Can't be [" + clsCommon.myCstr(DivConvFat) + "]  For Item [" + strICode + "] and UOM [" + clsCommon.myCstr(gvItem.Rows(Indx).Cells(colUOM).Value) + "] ")
                End If
                ' gvItem.Rows(Indx).Cells(colQty).Value = Math.Round(clsCommon.myCDecimal(gvItem.Rows(Indx).Cells(colQty).Value) * MulConvFat / DivConvFat, 2, MidpointRounding.AwayFromZero)
                gvItem.Rows(Indx).Cells(colUOM).Value = clsCommon.myCstr(dt.Rows(0)("UOM_Code"))
            End If
            'Qry += " ((Qty*MulConv.Conversion_Factor)/DivConv.Conversion_Factor) as Qty "
        End If

        UpdateCurrentRow(Indx, False)


    End Sub

    Private Sub UpdateCurrentRow(indx As Integer, isEditQty As Boolean)
        gvItem.Rows(indx).Cells(colDiffQty).Value = clsCommon.myCDecimal(gvItem.Rows(indx).Cells(IIf(isEditQty, colEditQty, colQty)).Value) - clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPQty).Value)
        gvItem.Rows(indx).Cells(colDiffFAT).Value = clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colFAT).Value) - clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPFAT).Value)
        gvItem.Rows(indx).Cells(colDiffSNF).Value = clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colSNF).Value) - clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPSNF).Value)
        gvItem.Rows(indx).Cells(colDiffAmt).Value = clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colAmount).Value) - clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPAmt).Value)


        Dim RecoStatus As Boolean = False
        If clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPQty).Value) > 0 Then
            If SettAllowMPQtyGreaterThanDCSQty And (clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPQty).Value) > clsCommon.myCDecimal(gvItem.Rows(indx).Cells(IIf(isEditQty, colEditQty, colQty)).Value)) Then
                RecoStatus = True
            End If
            If SettAllowMPQtyLessThanDCSQty And (clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPQty).Value) < clsCommon.myCDecimal(gvItem.Rows(indx).Cells(IIf(isEditQty, colEditQty, colQty)).Value)) Then
                RecoStatus = True
            End If
            If SettAllowMPQtyEqualToDCSQty And (clsCommon.myCDecimal(gvItem.Rows(indx).Cells(colMPQty).Value) = clsCommon.myCDecimal(gvItem.Rows(indx).Cells(IIf(isEditQty, colEditQty, colQty)).Value)) Then
                RecoStatus = True
            End If
        End If
        gvItem.Rows(indx).Cells(colRecoStatus).Value = RecoStatus
    End Sub

    Sub loadBlankGrid()
        gvItem.Rows.Clear()
        gvItem.Columns.Clear()
        gvItem.DataSource = Nothing
        Dim farmercode As New GridViewTextBoxColumn()

        Dim lineNo As New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "PKID."
        lineNo.Name = colPKID
        lineNo.IsVisible = False
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        lineNo = New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "SNo."
        lineNo.Name = colSlNo
        lineNo.Width = 60
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        Dim repoIsSurTax1 As New GridViewCheckBoxColumn()
        repoIsSurTax1.HeaderText = "Status"
        repoIsSurTax1.Name = colRecoStatus
        repoIsSurTax1.ReadOnly = True
        repoIsSurTax1.IsVisible = True
        repoIsSurTax1.Width = 50
        repoIsSurTax1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax1) '30



        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "BMC Code"
        farmercode.Name = colBMCCode
        farmercode.ReadOnly = True
        farmercode.Width = 100
        farmercode.IsVisible = True
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)



        Dim farmername As New GridViewTextBoxColumn()
        farmername.FormatString = ""
        farmername.HeaderText = "BMC Name"
        farmername.Name = colBMCName
        farmername.Width = 150
        farmername.ReadOnly = True
        farmername.IsVisible = True
        farmername.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmername)

        Dim routecode As New GridViewTextBoxColumn()
        routecode.FormatString = ""
        routecode.HeaderText = "Route Code"
        routecode.Name = colRouteCode
        routecode.Width = 150
        routecode.ReadOnly = True
        routecode.IsVisible = True
        routecode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(routecode)

        Dim routename As New GridViewTextBoxColumn()
        routename.FormatString = ""
        routename.HeaderText = "Route Name"
        routename.Name = colRouteName
        routename.Width = 150
        routename.ReadOnly = True
        routename.IsVisible = True
        routename.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(routename)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Zone Code"
        farmercode.Name = colZoneCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "Zone"
        farmercode.Name = colZoneName
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "DCS"
        farmercode.Name = colVLCUploaderCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 80
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "DCS Code"
        farmercode.Name = colVLCCode
        farmercode.ReadOnly = True
        farmercode.IsVisible = False
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        farmercode = New GridViewTextBoxColumn()
        farmercode.FormatString = ""
        farmercode.HeaderText = "DCS Name"
        farmercode.Name = colVLCName
        farmercode.ReadOnly = True
        farmercode.IsVisible = True
        farmercode.Width = 150
        farmercode.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(farmercode)

        lineNo = New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "Year"
        lineNo.Name = colYear
        lineNo.Width = 60
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        lineNo = New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "Month"
        lineNo.Name = colMonth
        lineNo.Width = 60
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        lineNo = New GridViewTextBoxColumn()
        lineNo.FormatString = ""
        lineNo.HeaderText = "Cycle No"
        lineNo.Name = colCycleNo
        lineNo.Width = 60
        lineNo.ReadOnly = True
        lineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(lineNo)

        Dim Qty As New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "Qty"
        Qty.Name = colQty
        Qty.Width = 100
        Qty.ReadOnly = True
        Qty.FormatString = "{0:n3}"
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "Edit Qty"
        Qty.Name = colEditQty
        Qty.Width = 100
        Qty.FormatString = "{0:n3}"
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        Qty.IsVisible = btnEditDCSQty.Visible
        Qty.ReadOnly = False
        gvItem.Columns.Add(Qty)

        Dim UOM As New GridViewTextBoxColumn()
        UOM.FormatString = ""
        UOM.HeaderText = "UOM"
        UOM.Name = colUOM
        UOM.Width = 80
        UOM.ReadOnly = True
        UOM.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.Columns.Add(UOM)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "FAT"
        Qty.Name = colFAT
        Qty.Width = 80
        Qty.ReadOnly = True
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "SNF"
        Qty.Name = colSNF
        Qty.Width = 80
        Qty.ReadOnly = True
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        Dim Amount As New GridViewDecimalColumn
        Amount.FormatString = ""
        Amount.HeaderText = "Amount"
        Amount.Name = colAmount
        Amount.Width = 100
        Amount.FormatString = "{0:n2}"
        Amount.ReadOnly = True
        Amount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Amount)

        Dim AmountActual As New GridViewDecimalColumn
        AmountActual.FormatString = ""
        AmountActual.HeaderText = "No Of Farmers"
        AmountActual.Name = colMPCount
        AmountActual.Width = 100
        AmountActual.FormatString = "{0:n0}"
        AmountActual.ReadOnly = True
        AmountActual.IsVisible = True
        AmountActual.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(AmountActual)


        AmountActual = New GridViewDecimalColumn
        AmountActual.FormatString = ""
        AmountActual.HeaderText = "MP Qty"
        AmountActual.Name = colMPQty
        AmountActual.Width = 100
        AmountActual.FormatString = "{0:n2}"
        AmountActual.ReadOnly = True
        AmountActual.IsVisible = True
        AmountActual.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(AmountActual)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "MP FAT"
        Qty.Name = colMPFAT
        Qty.Width = 80
        Qty.ReadOnly = True
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "MP SNF"
        Qty.Name = colMPSNF
        Qty.Width = 80
        Qty.ReadOnly = True
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        AmountActual = New GridViewDecimalColumn
        AmountActual.FormatString = ""
        AmountActual.HeaderText = "MP Amount"
        AmountActual.Name = colMPAmt
        AmountActual.Width = 100
        AmountActual.FormatString = "{0:n2}"
        AmountActual.ReadOnly = True
        AmountActual.IsVisible = True
        AmountActual.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(AmountActual)


        AmountActual = New GridViewDecimalColumn
        AmountActual.FormatString = ""
        AmountActual.HeaderText = "Diff Qty"
        AmountActual.Name = colDiffQty
        AmountActual.Width = 100
        AmountActual.FormatString = "{0:n2}"
        AmountActual.ReadOnly = True
        AmountActual.IsVisible = True
        AmountActual.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(AmountActual)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "Diff FAT"
        Qty.Name = colDiffFAT
        Qty.Width = 80
        Qty.ReadOnly = True
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        Qty = New GridViewDecimalColumn
        Qty.FormatString = ""
        Qty.HeaderText = "Diff SNF"
        Qty.Name = colDiffSNF
        Qty.Width = 80
        Qty.ReadOnly = True
        Qty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(Qty)

        AmountActual = New GridViewDecimalColumn
        AmountActual.FormatString = ""
        AmountActual.HeaderText = "Diff Amount"
        AmountActual.Name = colDiffAmt
        AmountActual.Width = 100
        AmountActual.FormatString = "{0:n2}"
        AmountActual.ReadOnly = True
        AmountActual.IsVisible = True
        AmountActual.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.Columns.Add(AmountActual)


        gvItem.AllowAddNewRow = False
        gvItem.AllowDeleteRow = False
        gvItem.AllowRowReorder = False
        gvItem.ShowGroupPanel = False
        gvItem.EnableFiltering = True
        gvItem.ShowFilteringRow = True
        gvItem.EnableSorting = False
        gvItem.EnableGrouping = False
        gvItem.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvItem.GridBehavior = New MyBehavior()

        ReStoreGridLayout()

        gvItem.MasterTemplate.SummaryRowsBottom.Clear()
        Dim summaryRowItem As New GridViewSummaryRowItem()
        Dim Smitem As New GridViewSummaryItem(colQty, "{0:n3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colEditQty, "{0:n3}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colAmount, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colMPCount, "{0:n0}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colMPQty, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colMPAmt, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colDiffQty, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        Smitem = New GridViewSummaryItem(colDiffAmt, "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItem.Add(Smitem)

        gvItem.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
        gvItem.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            '  If rbtnNEFT.IsChecked Then
            If clsCommon.myLen(Me.Form_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(Me.Form_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= gvItem.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To gvItem.Columns.Count - 1 Step ii + 1
                        gvItem.Columns(ii).IsVisible = False
                        gvItem.Columns(ii).VisibleInColumnChooser = True
                    Next
                    gvItem.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub
    Private Sub SetUserMgmtNew()
        ''MyBase.SetUserMgmt(clsUserMgtCode.FrmVLCDataUploaderManual)
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btndelete.Visible = MyBase.isDeleteFlag
        btnPost.Visible = MyBase.isPostFlag
        RadButton3.Visible = MyBase.isExport
        btnReverse.Visible = False
        'If MyBase.isReverse Then
        '    btnReverse.Enabled = True
        'Else
        '    btnReverse.Enabled = False
        'End If
    End Sub
    Sub Reset()
        loadBlankGrid()
        UcAttachment1.BlankAllControls()
        Dim dt As Date = clsCommon.GETSERVERDATE()

        txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dt) & "/" & DatePart(DateInterval.Year, dt)
        If SettMPIncentiveEntryApplyMonthly Then
            txtToDate.Value = txtFromDate.Value.AddMonths(1).AddDays(-1)
        Else
            txtToDate.Value = "01/" & DatePart(DateInterval.Month, dt) & "/" & DatePart(DateInterval.Year, dt)
        End If


        txtDocumentNo.Value = ""
        txtdate.Value = dt
        txtZone.Value = ""
        lblZone.Text = ""
        txtDocumentNo.MyReadOnly = False
        btnsave.Text = "Save"
        btndelete.Enabled = False
        btnsave.Enabled = True
        btnPost.Enabled = False
        txtdate.Focus()
        EnableInputDataField()
        isNewEntry = True
        IsinsideLoadData = False
        lblPending.Status = ERPTransactionStatus.Pending
        chkFarmerCollection.Checked = SettDBTByDailyFarmerMilkCollection
        txtFatPer.Tag = 0
        txtFatPer.Value = 0
        txtSnfPer.Tag = 0
        txtSnfPer.Value = 0
    End Sub
    Private Function AllowToSave(ByVal isByPost As Boolean) As Boolean
        Xtra.TransactionValidity(txtdate.Value)
        If AllowFutureDateTransaction(txtdate.Value, Nothing) = False Then
            txtdate.Focus()
            Return False
        End If
        If Not isByPost Then
            If SettApplyZoneOnDBT Then
                If clsCommon.myLen(txtZone.Value) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Please select " + txtZone.MyLinkLable2.Text)
                    txtZone.Focus()
                    Return False
                End If
            End If
        End If
        UcAttachment1.AllowToSave()
        SetToDate()
        If Not isNewEntry Then
            If SettRefreshDBTReco Then
                RefershFarmerInfo()
            End If
        End If
        Return True
    End Function
    Sub SaveData(ByVal isByPost As Boolean)
        Try
            If AllowToSave(isByPost) Then
                Dim obj As New clsMPDCSInsentiveReco()
                obj.Document_Code = txtDocumentNo.Value
                obj.Document_Date = txtdate.Value
                obj.Zone_Code = txtZone.Value
                obj.Reco_Date = txtFromDate.Value
                obj.Reco_Date_To = txtToDate.Value
                obj.Apply_FAT_Above = clsCommon.myCDecimal(txtFatPer.Tag)
                obj.Apply_SNF_Above = clsCommon.myCDecimal(txtSnfPer.Tag)
                obj.Farmer_Collection = chkFarmerCollection.Checked
                Dim objTr As New clsMPDCSInsentiveRecoDetail
                obj.arr = New List(Of clsMPDCSInsentiveRecoDetail)
                For Each grow As GridViewRowInfo In gvItem.Rows
                    If clsCommon.myLen(grow.Cells(colQty).Value) > 0 Then
                        objTr = New clsMPDCSInsentiveRecoDetail()
                        objTr.PK_Id = clsCommon.myCDecimal(grow.Cells(colPKID).Value)
                        objTr.SNo = obj.arr.Count + 1
                        objTr.Reco_Staus = clsCommon.myCBool(grow.Cells(colRecoStatus).Value)
                        objTr.MCC_Code = clsCommon.myCstr(grow.Cells(colBMCCode).Value)
                        objTr.VLC_Code = clsCommon.myCstr(grow.Cells(colVLCCode).Value)
                        objTr.Cycle_Year = clsCommon.myCstr(grow.Cells(colYear).Value)
                        objTr.Cycle_Month = clsCommon.myCstr(grow.Cells(colMonth).Value)
                        objTr.Cycle_No = clsCommon.myCstr(grow.Cells(colCycleNo).Value)
                        objTr.Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                        objTr.UOM = clsCommon.myCstr(grow.Cells(colUOM).Value)
                        objTr.FAT = clsCommon.myCDecimal(grow.Cells(colFAT).Value)
                        objTr.SNF = clsCommon.myCDecimal(grow.Cells(colSNF).Value)
                        objTr.Amount = clsCommon.myCDecimal(grow.Cells(colAmount).Value)
                        objTr.MP_Count = clsCommon.myCDecimal(grow.Cells(colMPCount).Value)
                        objTr.MP_Qty = clsCommon.myCDecimal(grow.Cells(colMPQty).Value)
                        objTr.MP_FAT = clsCommon.myCDecimal(grow.Cells(colMPFAT).Value)
                        objTr.MP_SNF = clsCommon.myCDecimal(grow.Cells(colMPSNF).Value)
                        objTr.MP_Amount = clsCommon.myCDecimal(grow.Cells(colMPAmt).Value)
                        objTr.Diff_Qty = clsCommon.myCDecimal(grow.Cells(colDiffQty).Value)
                        objTr.Diff_FAT = clsCommon.myCDecimal(grow.Cells(colDiffFAT).Value)
                        objTr.Diff_SNF = clsCommon.myCDecimal(grow.Cells(colDiffSNF).Value)
                        objTr.Diff_Amount = clsCommon.myCDecimal(grow.Cells(colDiffAmt).Value)
                        If Not isByPost Then
                            If SettApplyZoneOnDBT Then
                                If Not clsCommon.CompairString(clsCommon.myCstr(clsCommon.myCstr(grow.Cells(colZoneCode).Value)), txtZone.Value) = CompairStringResult.Equal Then
                                    Throw New Exception("Zone Should be [" + txtZone.Value + "] at [" + clsCommon.myCstr(obj.arr.Count + 1) + "]")
                                End If
                            End If
                        End If
                        obj.arr.Add(objTr)
                    End If
                Next
                If obj.arr.Count <= 0 Then
                    Throw New Exception("No data found to save")
                End If
                Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
                Try
                    clsMPDCSInsentiveReco.SaveData(obj, isNewEntry, trans)
                    UcAttachment1.SaveData(obj.Document_Code, True, trans)
                    trans.Commit()
                    If Not isByPost Then
                        clsCommon.MyMessageBoxShow(Me, "Data saved successfully", Me.Text)
                        LoadData(obj.Document_Code, NavigatorType.Current)
                    End If
                Catch ex As Exception
                    trans.Rollback()
                    Throw New Exception(ex.Message)
                End Try
            End If
        Catch ex As Exception
            If isByPost Then
                Throw New Exception("Error while updating data " + ex.Message)
            Else
                clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            End If
        End Try
    End Sub
    Private Sub DeleteData()
        Try
            If (deleteConfirm()) Then
                If (clsMPDCSInsentiveReco.DeleteData(txtDocumentNo.Value)) Then
                    clsCommon.MyMessageBoxShow(Me, "Data Deleted Successfully ", Me.Text)
                    Reset()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub LoadData(ByVal strCode As String, ByVal NavTyep As NavigatorType)
        IsinsideLoadData = True
        Reset()
        Dim obj As clsMPDCSInsentiveReco = clsMPDCSInsentiveReco.GetData(strCode, NavTyep, objCommonVar.strCurrUserZones, Nothing)
        If obj IsNot Nothing Then
            DisableInputDataField()
            isNewEntry = False
            txtDocumentNo.Value = obj.Document_Code
            txtdate.Value = obj.Document_Date
            txtFromDate.Value = obj.Reco_Date
            txtToDate.Value = obj.Reco_Date_To
            lblPending.Status = obj.Status
            txtZone.Value = obj.Zone_Code
            lblZone.Text = ClsZoneMaster.GetName(obj.Zone_Code)
            txtFatPer.Tag = obj.Apply_FAT_Above
            txtFatPer.Value = obj.Apply_FAT_Above

            txtSnfPer.Tag = obj.Apply_SNF_Above
            txtSnfPer.Value = obj.Apply_SNF_Above
            chkFarmerCollection.Checked = obj.Farmer_Collection
            If obj.arr IsNot Nothing AndAlso obj.arr.Count > 0 Then
                CycleNo = obj.arr(0).Cycle_No
                For Each objTr As clsMPDCSInsentiveRecoDetail In obj.arr
                    gvItem.Rows.AddNew()
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colPKID).Value = objTr.PK_Id

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSlNo).Value = gvItem.Rows.Count
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colRecoStatus).Value = objTr.Reco_Staus
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colRecoStatus).Tag = objTr.Reco_Staus
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCCode).Value = objTr.MCC_Code
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCName).Value = objTr.MCC_Name

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colRouteCode).Value = objTr.Route_Code
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colRouteName).Value = objTr.Route_Name

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneCode).Value = objTr.Zone_Code
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneName).Value = objTr.Zone_Name

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCCode).Value = objTr.VLC_Code
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCUploaderCode).Value = objTr.VLC_Uploader_Code
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCName).Value = objTr.VLC_Name

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colYear).Value = objTr.Cycle_Year
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMonth).Value = objTr.Cycle_Month
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colCycleNo).Value = objTr.Cycle_No

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colQty).Value = objTr.Qty
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colEditQty).Value = objTr.Qty
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colUOM).Value = objTr.UOM
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colFAT).Value = objTr.FAT
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSNF).Value = objTr.SNF
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colAmount).Value = objTr.Amount

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMPCount).Value = objTr.MP_Count
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMPQty).Value = objTr.MP_Qty
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMPFAT).Value = objTr.MP_FAT
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMPSNF).Value = objTr.MP_SNF
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMPAmt).Value = objTr.MP_Amount

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colDiffQty).Value = objTr.Diff_Qty
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colDiffFAT).Value = objTr.Diff_FAT
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colDiffSNF).Value = objTr.Diff_SNF
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colDiffAmt).Value = objTr.Diff_Amount

                    'FillFarmerInfo(gvItem.Rows.Count - 1)
                Next
            End If
            UcAttachment1.LoadData(obj.Document_Code)
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
            Dim qry As String = "select count(*) from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Document_Code='" + txtDocumentNo.Value + "' "
            Dim check As Integer = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(qry))
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
            txtDocumentNo.Value = clsMPDCSInsentiveReco.getFinder("", txtDocumentNo.Value, isButtonClicked)
            LoadData(txtDocumentNo.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub DisableInputDataField()
        txtdate.Enabled = False
        txtFromDate.Enabled = False
        txtZone.Enabled = False
    End Sub
    Sub EnableInputDataField()
        txtdate.Enabled = True
        txtFromDate.Enabled = True
        txtZone.Enabled = True
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
        SaveData(False)
    End Sub
    Private Sub btnPost_Click(sender As Object, e As EventArgs) Handles btnPost.Click
        Try
            Dim qry As String = ""
            Dim msg As String = ""
            Dim dt As DataTable = Nothing

            If (myMessages.postConfirm()) Then
                If SettRefreshDBTReco Then
                    SaveData(True)
                End If
                clsMPDCSInsentiveReco.PostData(txtDocumentNo.Value)
                clsCommon.MyMessageBoxShow(Me, "Successfully Posted", Me.Text)
                LoadData(txtDocumentNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub

    Private Sub RadButton3_Click(sender As Object, e As EventArgs) Handles RadButton3.Click
        Try
            transportSql.exportdata(gvItem, Me.Text, "")
            'clsCommon.MyExportToExcel("", gvItem, Nothing, Me.Text)
        Catch ex As Exception
            clsCommon.ProgressBarHide()
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub txtFromDate_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles txtFromDate.Validating
        If txtFromDate.Enabled Then
            SetToDate()
            GetDocNoAndLoad()
        End If
    End Sub

    Sub GetDocNoAndLoad()
        Dim qry As String = "select max(Document_Code) as Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_HEAD where Reco_Date='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "'"
        If SettApplyZoneOnDBT Then
            qry += " and TSPL_DCS_MP_INCENTIVE_RECO_HEAD.Zone_Code='" + txtZone.Value + "'"
        End If
        Dim strDocNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
        If clsCommon.myLen(strDocNo) > 0 Then
            LoadData(strDocNo, NavigatorType.Current)
        End If
    End Sub

    Dim CycleNo As Integer = 0
    Sub SetToDate()
        If Not IsinsideLoadData Then
            Dim PaymentCycleType As String = ""
            Dim PaymentCycleValue As Integer = 0
            If SettMPIncentiveEntryApplyMonthly Then
                txtFromDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1)
                txtToDate.Value = txtFromDate.Value.AddMonths(1).AddDays(-1)
                CycleNo = 1
            Else
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(" select top 1 TSPL_MCC_MASTER.Payment_Cycle,TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE  from TSPL_MCC_MASTER left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle   where TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE is not null") ''TSPL_MCC_MASTER.MCC_Code  in ('" + txtMCC.Value + "') 
                If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "No Payment Cycle found on current MCC/Location", Me.Text)
                    Exit Sub
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
                    CycleNo = (txtFromDate.Value.Day / PaymentCycleValue) + 1
                ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
                    If clsCommon.myCdbl(clsCommon.GetPrintDate(txtFromDate.Value, "dd")) <> 1 Then
                        clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month, Because MCC has payment Cycle of Month Type", Me.Text)
                        txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                        txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                        Exit Sub
                    End If
                    txtToDate.Value = DateAdd(DateInterval.Month, PaymentCycleValue, txtFromDate.Value)
                    CycleNo = 1
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
    End Sub

    Private Sub RadButton5_Click(sender As Object, e As EventArgs) Handles RadButton5.Click
        RefershFarmerInfo()
    End Sub

    Private Sub RefershFarmerInfo()
        For ii As Integer = 0 To gvItem.Rows.Count - 1
            FillFarmerInfo(ii)
        Next
    End Sub

    Private Sub RadButton4_Click(sender As Object, e As EventArgs) Handles RadButton4.Click
        Dim ii As Integer = 0
        Try
            If Not SettDCSMPIncetiveReco Then
                Throw New Exception("This facility is not for you")
            End If
            Dim whrQry As String = Nothing
            Dim whrMonthCycleQry As String = Nothing
            If clsCommon.myLen(txtZone.Value) > 0 Then
                whrQry = " where TSPL_ZONE_MASTER.Zone_Code='" + clsCommon.myCstr(txtZone.Value) + "'"
            End If
            If ApplyMonthlySetting Then
                whrMonthCycleQry = " where convert(Date, TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) >='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' And convert(Date, TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) <='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' "
            Else
                whrMonthCycleQry = " where convert(Date, TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE,103) ='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' "
            End If
            loadBlankGrid()

            Dim qry As String = "select TSPL_MCC_MASTER.MCC_Code,TSPL_MCC_MASTER.MCC_NAME,TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,xx.CycleYear,xx.CycleMonth,Convert(decimal(18," + clsCommon.myCstr(SettDCSQtyDecimalPlaces) + "),xx.Qty) As Qty,xx.UOM_Code,xx.FAT_KG,xx.SNF_KG,xx.AMOUNT,TSPL_VENDOR_MASTER.Zone_Code,TSPL_ZONE_MASTER.Description as Zone_Name from (
select VSP_CODE,CycleYear,CycleMonth,sum(Qty+RQty) as Qty,max(UOM_Code) as UOM_Code,sum(FAT_KG) as FAT_KG,sum(SNF_KG) as SNF_KG,sum(AMOUNT) as AMOUNT  from (
select TSPL_MILK_PURCHASE_INVOICE_HEAD.MCC_CODE,TSPL_MILK_PURCHASE_INVOICE_HEAD.VSP_CODE,DATEPART(YEAR, TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE) as CycleYear,
DATEPART(MONTH, TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_DATE) as CycleMonth
,IsNull((case when TSPL_MILK_REJECT_TYPE.Code is null then TSPL_MILK_SRN_DETAIL.ACC_Qty_LTR else 0 end),0) AS Qty
,IsNull((case when TSPL_MILK_REJECT_TYPE.Code is not null then TSPL_MILK_SRN_DETAIL.ACC_Qty_LTR else 0 end),0) AS RQty
,TSPL_MILK_SRN_DETAIL.UOM_Code,TSPL_MILK_SRN_DETAIL.FAT_KG,TSPL_MILK_SRN_DETAIL.SNF_KG,TSPL_MILK_SRN_DETAIL.AMOUNT from 
TSPL_MILK_PURCHASE_INVOICE_DETAIL
left outer join TSPL_MILK_PURCHASE_INVOICE_HEAD on TSPL_MILK_PURCHASE_INVOICE_HEAD.DOC_CODE=TSPL_MILK_PURCHASE_INVOICE_DETAIL.DOC_CODE 
left outer join TSPL_MILK_SRN_DETAIL on TSPL_MILK_SRN_DETAIL.DOC_CODE=TSPL_MILK_PURCHASE_INVOICE_DETAIL.SRN_CODE
left outer join TSPL_MILK_SRN_HEAD  on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE
left outer join TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL on TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Uploader_TR_No
left outer join TSPL_MILK_SHIFT_UPLOADER_DETAIL on TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No
left outer join TSPL_MILK_REJECT_TYPE on TSPL_MILK_REJECT_TYPE.Code= (case when TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No is not null then TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type else TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type end )
" + whrMonthCycleQry + "
and 2=(case when len(isnull(TSPL_MILK_REJECT_TYPE.Code,''))>0 then (case when isnull(TSPL_MILK_REJECT_TYPE.Include_In_DBT,0)=1 then 2 else 3 end ) else 2 end)
)X Group by VSP_CODE,CycleYear,CycleMonth
)xx 
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code=xx.VSP_CODE
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_VLC_MASTER_HEAD.MCC
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_VLC_MASTER_HEAD.VSP_Code
left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_VENDOR_MASTER.Zone_Code" + whrQry
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            txtFatPer.Tag = txtFatPer.Value
            txtSnfPer.Tag = txtSnfPer.Value
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                For ii = 0 To dt.Rows.Count - 1
                    If clsCommon.myCDecimal(dt.Rows(ii)("CycleYear")) <> txtFromDate.Value.Year Then
                        Throw New Exception("Year is Mismatched")
                    End If
                    If clsCommon.myCDecimal(dt.Rows(ii)("CycleMonth")) <> txtFromDate.Value.Month Then
                        Throw New Exception("Month is Mismatched")
                    End If
                    'If clsCommon.myCDecimal(dt.Rows(ii)("Cycle No")) <> CycleNo Then
                    '    Throw New Exception("Cycle No is Mismatched")
                    'End If
                    qry = "select PK_Id from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL  where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code not in ('" + txtDocumentNo.Value + "') and Cycle_Year='" + clsCommon.myCstr(dt.Rows(ii)("CycleYear")) + "' and Cycle_Month='" + clsCommon.myCstr(dt.Rows(ii)("CycleMonth")) + "' and Cycle_No='" + clsCommon.myCstr(CycleNo) + "' and  VLC_Code='" + clsCommon.myCstr(dt.Rows(ii)("VLC_Code")) + "'"
                    Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable(qry)
                    If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                        Continue For
                    End If
                    gvItem.Rows.AddNew()
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSlNo).Value = gvItem.Rows.Count

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCCode).Value = clsCommon.myCstr(dt.Rows(ii)("MCC_CODE"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colBMCName).Value = clsCommon.myCstr(dt.Rows(ii)("MCC_NAME"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneCode).Value = clsCommon.myCstr(dt.Rows(ii)("Zone_Code"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colZoneName).Value = clsCommon.myCstr(dt.Rows(ii)("Zone_Name"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCCode).Value = clsCommon.myCstr(dt.Rows(ii)("VLC_Code"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCUploaderCode).Value = clsCommon.myCstr(dt.Rows(ii)("VLC_Code_VLC_Uploader"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colVLCName).Value = clsCommon.myCstr(dt.Rows(ii)("VLC_Name"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colYear).Value = clsCommon.myCDecimal(dt.Rows(ii)("CycleYear"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colMonth).Value = clsCommon.myCDecimal(dt.Rows(ii)("CycleMonth"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colCycleNo).Value = CycleNo

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colQty).Value = clsCommon.myCDecimal(dt.Rows(ii)("Qty"))
                    If clsCommon.myLen(dt.Rows(ii)("UOM_Code")) <= 0 Then
                        Throw New Exception("Please define UOM")
                    End If

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colUOM).Value = clsCommon.myCstr(dt.Rows(ii)("UOM_Code"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colFAT).Value = clsCommon.myCDecimal(dt.Rows(ii)("FAT_KG"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colSNF).Value = clsCommon.myCDecimal(dt.Rows(ii)("SNF_KG"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(colAmount).Value = clsCommon.myCDecimal(dt.Rows(ii)("AMOUNT"))


                    FillFarmerInfo(gvItem.Rows.Count - 1)

                Next
            Else
                clsCommon.MyMessageBoxShow(Me, "Data Not Found", Me.Text)
            End If
        Catch ex As Exception
            loadBlankGrid()
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub RadButton6_Click(sender As Object, e As EventArgs) Handles RadButton6.Click
        Try
            If chkFarmerCollection.Checked Then
                Throw New Exception("Update is not allowed for the Day-wise Farmer Collection document.")
            End If

            If Not lblPending.Status = ERPTransactionStatus.Pending Then
                Throw New Exception("Trasanction should be pending for update")
            End If
            If SettApplyZoneOnDBT Then
                If clsCommon.myLen(txtZone.Value) <= 0 Then
                    Throw New Exception("Please select " + txtZone.MyLinkLable2.Text)
                End If
            End If
            Dim Arr As New List(Of clsMPDCSInsentiveRecoDetail)
            For Each grow As GridViewRowInfo In gvItem.Rows
                If clsCommon.myLen(grow.Cells(colQty).Value) > 0 Then
                    Dim objTr = New clsMPDCSInsentiveRecoDetail()
                    objTr.PK_Id = clsCommon.myCDecimal(grow.Cells(colPKID).Value)
                    'objTr.SNo = obj.arr.Count + 1
                    objTr.Reco_Staus = clsCommon.myCBool(grow.Cells(colRecoStatus).Value)
                    objTr.MCC_Code = clsCommon.myCstr(grow.Cells(colBMCCode).Value)
                    objTr.VLC_Code = clsCommon.myCstr(grow.Cells(colVLCCode).Value)
                    objTr.Cycle_Year = clsCommon.myCstr(grow.Cells(colYear).Value)
                    objTr.Cycle_Month = clsCommon.myCstr(grow.Cells(colMonth).Value)
                    objTr.Cycle_No = clsCommon.myCstr(grow.Cells(colCycleNo).Value)
                    objTr.Qty = clsCommon.myCDecimal(grow.Cells(colQty).Value)
                    objTr.UOM = clsCommon.myCstr(grow.Cells(colUOM).Value)
                    objTr.FAT = clsCommon.myCDecimal(grow.Cells(colFAT).Value)
                    objTr.SNF = clsCommon.myCDecimal(grow.Cells(colSNF).Value)
                    objTr.Amount = clsCommon.myCDecimal(grow.Cells(colAmount).Value)
                    objTr.MP_Count = clsCommon.myCDecimal(grow.Cells(colMPCount).Value)
                    objTr.MP_Qty = clsCommon.myCDecimal(grow.Cells(colMPQty).Value)
                    objTr.MP_FAT = clsCommon.myCDecimal(grow.Cells(colMPFAT).Value)
                    objTr.MP_SNF = clsCommon.myCDecimal(grow.Cells(colMPSNF).Value)
                    objTr.MP_Amount = clsCommon.myCDecimal(grow.Cells(colMPAmt).Value)
                    objTr.Diff_Qty = clsCommon.myCDecimal(grow.Cells(colDiffQty).Value)
                    objTr.Diff_FAT = clsCommon.myCDecimal(grow.Cells(colDiffFAT).Value)
                    objTr.Diff_SNF = clsCommon.myCDecimal(grow.Cells(colDiffSNF).Value)
                    objTr.Diff_Amount = clsCommon.myCDecimal(grow.Cells(colDiffAmt).Value)
                    If SettApplyZoneOnDBT Then
                        If Not clsCommon.CompairString(clsCommon.myCstr(clsCommon.myCstr(grow.Cells(colZoneCode).Value)), txtZone.Value) = CompairStringResult.Equal Then
                            Throw New Exception("Zone Should be [" + txtZone.Value + "] at [" + clsCommon.myCstr(clsCommon.myCDecimal(grow.Cells(colSlNo).Value)) + "]")
                        End If
                    End If
                    Arr.Add(objTr)
                End If
            Next
            If Arr.Count <= 0 Then
                Throw New Exception("No data found to save")
            End If
            clsMPDCSInsentiveRecoDetail.saveDataZone(txtDocumentNo.Value, objCommonVar.strCurrUserZones, Arr, chkFarmerCollection.Checked)
            clsCommon.MyMessageBoxShow(Me, "Data updated successfully", Me.Text)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        Try
            If clsCommon.myLen(Me.Form_ID) > 0 Then
                gvItem.MasterTemplate.FilterDescriptors.Clear()
                Dim obj As New clsGridLayout()
                obj.ReportID = Me.Form_ID
                obj.UserID = objCommonVar.CurrentUserCode
                obj.GridLayout = New MemoryStream()
                gvItem.SaveLayout(obj.GridLayout)
                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                obj.GridColumns = gvItem.ColumnCount
                If obj.SaveData() Then
                    clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
                End If
                obj.GridLayout.Close()
                obj.GridLayout.Dispose()
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        Try
            clsGridLayout.DeleteData(Me.Form_ID, objCommonVar.CurrentUserCode)
            clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtZone__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtZone._MYValidating
        Dim whr As String = "2=3"
        If SettApplyZoneOnDBT Then
            whr = "2=2"
        End If
        txtZone.Value = ClsZoneMaster.getFinder(whr, txtZone.Value, isButtonClicked)
        lblZone.Text = ClsZoneMaster.GetName(txtZone.Value)
        GetDocNoAndLoad()
    End Sub

    Private Sub btnReverse_Click(sender As Object, e As EventArgs) Handles btnReverse.Click
        Try
            If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                Throw New Exception("No document found to unpost")
            End If
            If clsCommon.MyMessageBoxShow("Do you want to Reverse and unpost the current Document [" + txtDocumentNo.Value + "] " + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo, RadMessageIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
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
                clsMPDCSInsentiveReco.ReverseAndUnpost(txtDocumentNo.Value)
                clsCommon.MyMessageBoxShow("Task done Successfully", Me.Text)
                LoadData(txtDocumentNo.Value, NavigatorType.Current)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnEditDCSQty_Click(sender As Object, e As EventArgs) Handles btnEditDCSQty.Click
        Try
            If Not MyBase.isPostFlag Then
                Throw New Exception("Unauthorized user")
            End If


            Dim Arr As New List(Of clsMPDCSInsentiveRecoDetail)
            Dim dclTotalEditQtyOLD As Decimal = 0
            Dim dclTotalEditQtyNew As Decimal = 0
            For ii As Integer = 0 To gvItem.Rows.Count - 1
                If clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colQty).Value) <> clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colEditQty).Value) Then
                    If clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colEditQty).Value) < clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colMPQty).Value) Then
                        Throw New Exception("DCS [" + clsCommon.myCstr(gvItem.Rows(ii).Cells(colVLCUploaderCode).Value) + "] Edit DCS Qty [" + clsCommon.myCstr(gvItem.Rows(ii).Cells(colEditQty).Value) + "] Can't be less than MP Qty [" + clsCommon.myCstr(gvItem.Rows(ii).Cells(colMPQty).Value) + "]  ")
                    End If

                    dclTotalEditQtyOLD += clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colQty).Value)
                    dclTotalEditQtyNew += clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colEditQty).Value)


                    Dim objTr As New clsMPDCSInsentiveRecoDetail()
                    objTr.Reco_Staus_OLD = clsCommon.myCBool(gvItem.Rows(ii).Cells(colRecoStatus).Tag)
                    UpdateCurrentRow(ii, True)
                    objTr.SNo = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colSlNo).Value)
                    objTr.PK_Id = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colPKID).Value)
                    objTr.Reco_Staus = clsCommon.myCBool(gvItem.Rows(ii).Cells(colRecoStatus).Value)
                    objTr.MCC_Code = clsCommon.myCstr(gvItem.Rows(ii).Cells(colBMCCode).Value)
                    objTr.VLC_Code = clsCommon.myCstr(gvItem.Rows(ii).Cells(colVLCCode).Value)
                    objTr.Cycle_Year = clsCommon.myCstr(gvItem.Rows(ii).Cells(colYear).Value)
                    objTr.Cycle_Month = clsCommon.myCstr(gvItem.Rows(ii).Cells(colMonth).Value)
                    objTr.Cycle_No = clsCommon.myCstr(gvItem.Rows(ii).Cells(colCycleNo).Value)
                    objTr.Qty = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colEditQty).Value)
                    objTr.UOM = clsCommon.myCstr(gvItem.Rows(ii).Cells(colUOM).Value)
                    objTr.FAT = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colFAT).Value)
                    objTr.SNF = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colSNF).Value)
                    objTr.Amount = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colAmount).Value)
                    objTr.MP_Count = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colMPCount).Value)
                    objTr.MP_Qty = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colMPQty).Value)
                    objTr.MP_FAT = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colMPFAT).Value)
                    objTr.MP_SNF = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colMPSNF).Value)
                    objTr.MP_Amount = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colMPAmt).Value)
                    objTr.Diff_Qty = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colDiffQty).Value)
                    objTr.Diff_FAT = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colDiffFAT).Value)
                    objTr.Diff_SNF = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colDiffSNF).Value)
                    objTr.Diff_Amount = clsCommon.myCDecimal(gvItem.Rows(ii).Cells(colDiffAmt).Value)
                    Arr.Add(objTr)
                End If
            Next
            If Arr IsNot Nothing AndAlso Arr.Count > 0 Then
                If dclTotalEditQtyOLD <> dclTotalEditQtyNew Then
                    Throw New Exception("Sum of Qty [" + clsCommon.myCstr(dclTotalEditQtyOLD) + "] and Edit Qty [" + clsCommon.myCstr(dclTotalEditQtyNew) + "] should be same")
                End If
                Dim frm As New FrmFreeTxtBox1
                frm.Text = "Remarks for Update"
                frm.ShowDialog()
                If clsCommon.myLen(frm.strRmks) <= 0 Then
                    Throw New Exception("Please provice remarks to Edit Qty")
                Else
                    clsMPDCSInsentiveRecoDetail.EditDCSQty(txtDocumentNo.Value, frm.strRmks, Arr)
                    clsCommon.MyMessageBoxShow("Qty Edited Successfully", Me.Text)
                    LoadData(txtDocumentNo.Value, NavigatorType.Current)
                End If
            Else
                Throw New Exception("No Data edit")
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
End Class

