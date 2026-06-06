Imports common
Imports System.Data.SqlClient
Public Class frmCorrectionFarmer
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isNewEntry As Boolean = False
    Public errorControl As clsErrorControl = New clsErrorControl()
    Dim Remark As String
#End Region

    Private Sub frmMilkGateEntryIn_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim coll As New Dictionary(Of String, String)()
            coll.Add("Code", "Varchar(10) not null PRIMARY KEY")
            coll.Add("Description", "nvarchar(1000) null")
            clsCommonFunctionality.CreateOrAlterTable(False, False, "TSPL_CUSTOM_MSG", coll, "", False, False, "", "", "", False)

            Dim sql As String = "SELECT count(1) FROM INFORMATION_SCHEMA.COLUMNS WHERE  TABLE_NAME = 'TSPL_DCS_MP_INCENTIVE_RECO_DETAIL' AND COLUMN_NAME = 'Status' "
            Dim CodeExists As Decimal = clsCommon.myCDecimal(clsDBFuncationality.getSingleValue(sql))

            coll = New Dictionary(Of String, String)()
            coll.Add("Status", "int Null")
            coll.Add("Posted_By", "varchar(12) NULL")
            coll.Add("Posting_Date", "Datetime NULL")
            coll.Add("Loc_Latitude", "varchar(20) NULL")
            coll.Add("Loc_Longitude", "varchar(20) NULL")
            coll.Add("Loc_Address", "nvarchar(1000) null")
            clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL", coll, "unique(Cycle_Year,Cycle_Month,Cycle_No,VLC_Code)", True, True, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "", True)
            'clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID", coll, "", True, True, "TSPL_DCS_MP_INCENTIVE_RECO_HEAD", "Document_Code", "", True)

            Try
                If CodeExists = 0 Then
                    sql = "update TSPL_DCS_MP_INCENTIVE_RECO_DETAIL set TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Status=1 from (
select Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_HEAD  where Status=1
)xx inner join TSPL_DCS_MP_INCENTIVE_RECO_DETAIL on TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Document_Code=xx.Document_Code"
                    clsDBFuncationality.ExecuteNonQuery(sql)
                End If
            Catch ex As Exception
            End Try

            ButtonToolTip.SetToolTip(btnSave, "Press Alt+S for Save/Update ")
            ButtonToolTip.SetToolTip(btnclose, "Press Alt+C Close the Window")
            txtShiftDate.Value = clsCommon.GETSERVERDATE()
            RadGroupBox1.Enabled = False
            RadGroupBox2.Enabled = True
            txtShiftDate.Focus()
            RadPageView1.SelectedPage = RadPageViewPage1
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnSave.Visible = MyBase.isModifyFlag
        'btnExport.Visible = MyBase.isExport
        'btnImport.Visible = MyBase.isExport
    End Sub

    Private Sub txtMCC__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtMCC._MYValidating
        'Dim qry As String = ""
        'Dim arrMCCRights As ArrayList = clsMCCCodes.GetUserHavingMCCRights()

        'qry = "select * from ( select tspl_mcc_master.MCC_Code as [Code] ,tspl_mcc_master.MCC_Type as [Mcc Type] ,tspl_mcc_master.MCC_NAME as [Mcc Name] ,tspl_mcc_master.Chilling_Vendor as [Chilling Vendor] ,tspl_mcc_master.Add1 as [Address1] ,tspl_mcc_master.Add2 as [Address2] ,tspl_mcc_master.Tehsil as [Tehsil] ,tspl_mcc_master.City_code as [City Code] ,tspl_mcc_master.State_Code as [State Code] ,tspl_mcc_master.Country_code as [Country Code] ,tspl_mcc_master.Pin_code as [Pin Code],tspl_mcc_master.Pan_No as [Pan No] ,tspl_mcc_master.Telphone as [Telphone] ,tspl_mcc_master.Email as [Email] ,tspl_mcc_master.Fax as [Fax] ,tspl_mcc_master.MCC_Area as [Mcc Area] ,tspl_mcc_master.Area_Of_Store as [Area Of Store] ,tspl_mcc_master.Area_Of_Office as [Area Of Office] ,tspl_mcc_master.Open_Area_For_tanker as [Open Area For Tanker] ,tspl_mcc_master.Area_Of_LAB as [Area Of Lab] ,tspl_mcc_master.No_Of_SILO as [No Of Silo] ,tspl_mcc_master.Total_Storage_capacity as [Total Storage Capacity] ,tspl_mcc_master.Area_Of_Receiving_DOCK as [Area Of Receiving Dock] ,tspl_mcc_master.No_Of_Chiller as [No Of Chiller] ,tspl_mcc_master.Chiller_Brand_Name as [Chiller Brand Name] ,tspl_mcc_master.Chiller_Capacity as [Chiller Capacity] ,tspl_mcc_master.No_Of_MilkPump as [No Of Milkpump] ,tspl_mcc_master.MilkPump_Capacity as [Milkpump Capacity] ,tspl_mcc_master.DripSaver as [Drip Saver] ,tspl_mcc_master.CanWasher as [Can Washer] ,tspl_mcc_master.CanScrubber as [Can Scrubber] ,tspl_mcc_master.FSSAI_NO as [FSSAI No] ,tspl_mcc_master.ETP as [ETP] ,tspl_mcc_master.Earthing as [Earthing] ,tspl_mcc_master.Coil_Length as [Coil Length] ,tspl_mcc_master.Electricity_Connection as [Electricity Connection] ,tspl_mcc_master.Boiler as [Boiler] ,tspl_mcc_master.NoOfDG as [No. of DG] ,tspl_mcc_master.NoOfCompressor as [No. of Compressor] ,tspl_mcc_master.PayeeName as [Payee Name] ,tspl_mcc_master.BankName as [Bank Name] ,tspl_mcc_master.BankBranch as [Bank Branch] ,tspl_mcc_master.BankCityCode as [Bank City Code] ,tspl_mcc_master.BankStateCode as [Bank State Code] ,tspl_mcc_master.IFCICode as [IFCI Code] ,tspl_mcc_master.AccountNO as [Account No] ,tspl_mcc_master.Created_By as [Created By] ,tspl_mcc_master.Created_Date as [Created Date] ,tspl_mcc_master.Modified_By as [Modified By] ,tspl_mcc_master.Modified_Date as [Modified Date] ,tspl_mcc_master.Comp_Code as [Company Code],tspl_mcc_master.mcc_code_vlc_uploader as [MCC Code For VLC Uploder],tspl_mcc_master.Plant_Code AS [Plant Code],TSPL_LOCATION_MASTER_PLANT.Location_Desc AS [Plant Name] from tspl_mcc_master LEFT JOIN TSPL_LOCATION_MASTER as TSPL_LOCATION_MASTER_PLANT ON TSPL_LOCATION_MASTER_PLANT.Location_Code=tspl_mcc_master.Plant_Code  inner join tspl_location_master on tspl_location_master.location_Code= tspl_mcc_master.mcc_Code where tspl_mcc_master.mcc_Code in (" & StrPermission & ") " _
        '& " and (  tspl_mcc_master.mcc_Code in (" & clsCommon.GetMulcallString(arrMCCRights) & ")))xx "

        'txtMCC.Value = clsCommon.ShowSelectForm("frmCorrection@MCC", qry, "Code", "", txtMCC.Value, "", isButtonClicked)
        'If txtMCC.Value IsNot Nothing AndAlso clsCommon.myLen(txtMCC.Value) > 0 Then
        '    lblMcc.Text = clsDBFuncationality.getSingleValue(" select MCC_NAME from TSPL_Mcc_MASTER where MCC_Code = '" + txtMCC.Value + "'", Nothing)
        'End If
    End Sub

    Private Sub txtVLC__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtVLC._MYValidating
        vlcUploaderFinder(txtVLC, lblVLC, isButtonClicked)
    End Sub

    Sub vlcUploaderFinder(ByVal finder As common.UserControls.txtFinder, ByVal label As common.Controls.MyLabel, ByVal isButtonClicked As Boolean)
        Try
            Dim qry As String = "select TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [UploaderCode], TSPL_VLC_MASTER_HEAD.VLC_Code AS [DCS Code],TSPL_VLC_MASTER_HEAD.VLC_Name as [DCS NAME] from TSPL_VLC_MASTER_HEAD "
            Dim whrCls As String = "  isnull(TSPL_VLC_MASTER_HEAD.IsSuspense,0)=0  "
            finder.Value = clsCommon.ShowSelectForm("SsaNUdC", qry, "UploaderCode", whrCls, finder.Value, "UploaderCode", isButtonClicked)

            qry = "select TSPL_VLC_MASTER_HEAD.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_VLC_MASTER_HEAD.MCC from TSPL_VLC_MASTER_HEAD where TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader='" + finder.Value + "' "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                finder.Tag = clsCommon.myCstr(dt.Rows(0)("VLC_Code"))
                label.Text = clsCommon.myCstr(dt.Rows(0)("VLC_Name"))
            Else
                finder.Tag = Nothing
                label.Text = ""
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub TxtFinder1__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles TxtFinder1._MYValidating
        vlcUploaderFinder(TxtFinder1, MyLabel5, isButtonClicked)
    End Sub
    Private Sub txtRoute__MYValidating_1(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtRoute._MYValidating
        Try
            Dim qry As String = " select ROUTE_NO as Code,ROUTE_NAME as Name from  TSPL_BULK_ROUTE_MASTER "
            Dim whrCls As String = ""
            txtRoute.Value = clsCommon.ShowSelectForm("corRoutFnd", qry, "Code", whrCls, txtRoute.Value, "Code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles RadButton1.Click
        Try
            If clsCommon.myLen(txtMCC.Value) <= 0 Then
                txtMCC.Focus()
                Throw New Exception("Please enter MCC")
            End If
            If clsCommon.myLen(cboShift.SelectedValue) <= 0 Then
                cboShift.Focus()
                Throw New Exception("Please Select shift")
            End If
            If chkAddMissingSample.Checked Then
                RadGroupBox2.Enabled = False
                RadGroupBox1.Enabled = True
                Exit Sub
            End If
            If clsCommon.myLen(txtVLC.Tag) <= 0 Then
                txtVLC.Focus()
                Throw New Exception("Please enter VLC")
            End If
            Dim qry As String = Nothing
            qry = "select TSPL_MILK_SRN_HEAD.DOC_CODE as SRNNo,TSPL_MILK_SRN_HEAD.Dock_Collection_Milk_Type as MilkType
,TSPL_MILK_SRN_DETAIL.Qty,TSPL_MILK_SRN_DETAIL.UOM_Code,TSPL_MILK_SRN_DETAIL.FAT_PER,TSPL_MILK_SRN_DETAIL.SNF_PER
,TSPL_MILK_SRN_DETAIL.Retesting_FAT,TSPL_MILK_SRN_DETAIL.Retesting_SNF,TSPL_MILK_SRN_DETAIL.Retesting_OR_Correction_Status
,(Case When Retesting_OR_Correction_Status=1 Then TSPL_MILK_SRN_DETAIL.Retesting_FAT Else (Case When Retesting_OR_Correction_Status=2 Then TSPL_MILK_SRN_DETAIL.FAT_PER Else TSPL_MILK_SRN_DETAIL.Retesting_FAT End)End) As FAT
,(Case When Retesting_OR_Correction_Status=1 Then TSPL_MILK_SRN_DETAIL.Retesting_SNF Else (Case When Retesting_OR_Correction_Status=2 Then TSPL_MILK_SRN_DETAIL.SNF_PER Else TSPL_MILK_SRN_DETAIL.Retesting_SNF End) End) As SNF 
,case when TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No is not null then TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Reject_Type else TSPL_MILK_SHIFT_UPLOADER_DETAIL.Reject_Type end Reject_Type
,TSPL_MILK_SRN_HEAD.ROUTE_CODE
from TSPL_MILK_SRN_DETAIL
left outer join TSPL_MILK_SRN_HEAD on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE
left outer join TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL on TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Uploader_TR_No
left outer join TSPL_MILK_SHIFT_UPLOADER_DETAIL on TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No"
            Dim whr As String = " convert(date, TSPL_MILK_SRN_HEAD.DOC_DATE,106)='" + clsCommon.GetPrintDate(txtShiftDate.Value, "dd/MMM/yyyy") + "' and TSPL_MILK_SRN_HEAD.SHIFT='" + clsCommon.myCstr(cboShift.SelectedValue) + "' and TSPL_MILK_SRN_HEAD.VLC_CODE='" + clsCommon.myCstr(txtVLC.Tag) + "'"


            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry + " Where " + whr)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("No Milk SRN found")
            End If
            Dim srnNo As String = ""
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If dt.Rows.Count = 1 Then
                    srnNo = clsCommon.myCstr(dt.Rows(0)("SRNNo"))
                Else
                    srnNo = clsCommon.ShowSelectForm("SRNCorrf", qry, "SRNNo", whr, srnNo, "SRNNo", True)
                End If
            End If
            If clsCommon.myLen(srnNo) > 0 Then
                dt = clsDBFuncationality.GetDataTable(qry + " where " + whr + " and TSPL_MILK_SRN_HEAD.DOC_CODE='" + srnNo + "'")
                lblSRNNo.Text = clsCommon.myCstr(dt.Rows(0)("SRNNo"))
                txtQty.Value = clsCommon.myCdbl(dt.Rows(0)("Qty"))
                txtQty.Tag = clsCommon.myCdbl(dt.Rows(0)("Qty"))
                lblUOM.Text = clsCommon.myCstr(dt.Rows(0)("UOM_Code"))

                If clsCommon.myCdbl(dt.Rows(0)("Retesting_OR_Correction_Status")) > 0 Then
                    txtFAT.Value = clsCommon.myCdbl(dt.Rows(0)("FAT"))
                    txtFAT.Tag = clsCommon.myCdbl(dt.Rows(0)("FAT"))
                    txtSNF.Value = clsCommon.myCdbl(dt.Rows(0)("SNF"))
                    txtSNF.Tag = clsCommon.myCdbl(dt.Rows(0)("SNF"))
                Else
                    txtFAT.Value = clsCommon.myCdbl(dt.Rows(0)("FAT_PER"))
                    txtFAT.Tag = clsCommon.myCdbl(dt.Rows(0)("FAT_PER"))
                    txtSNF.Value = clsCommon.myCdbl(dt.Rows(0)("SNF_PER"))
                    txtSNF.Tag = clsCommon.myCdbl(dt.Rows(0)("SNF_PER"))
                End If
                cboMilkType.SelectedValue = clsCommon.myCstr(dt.Rows(0)("MilkType"))
                cboMilkType.Tag = clsCommon.myCstr(dt.Rows(0)("MilkType"))
                cboRejectType.SelectedValue = clsCommon.myCstr(dt.Rows(0)("Reject_Type"))
                cboRejectType.Tag = clsCommon.myCstr(dt.Rows(0)("Reject_Type"))
                txtRoute.Value = clsCommon.myCstr(dt.Rows(0)("ROUTE_CODE"))

                TxtFinder1.Value = txtVLC.Value
                TxtFinder1.Tag = txtVLC.Tag
                MyLabel5.Text = lblVLC.Text
                RadGroupBox2.Enabled = False
                RadGroupBox1.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If clsCommon.CompairString(Form_ID, clsUserMgtCode.MilkRetesting) = CompairStringResult.Equal Then
            ShowRemarks()
        Else
            SaveData()
        End If
    End Sub
    Private Sub ShowRemarks()
        Try
            Dim obj As New clsMilkSRNMCC
            Dim qry As String = ""
            Dim Reason As String = ""
            Dim frm As New FrmFreeTxtBox1
            frm.Text = "Remarks for Retesting"
            frm.ShowDialog()
            If clsCommon.myLen(frm.strRmks) <= 0 Then
                Exit Sub
            Else
                If frm.strRmks IsNot Nothing Then
                    obj.Reason = "1"

                End If
            End If
            SaveData()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub SaveData()
        Try
            If isNewEntry = False AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                'Dim obj As New clsMilkCollectionMCC
                Dim qry As String = ""
                Dim Reason As String = ""
                Dim frm As New FrmFreeTxtBox1
                frm.Text = "Remarks for Update"
                frm.ShowDialog()
                If clsCommon.myLen(frm.strRmks) <= 0 Then
                    Exit Sub
                Else
                    Remark = frm.strRmks
                End If
            End If
            If chkAddMissingSample.Checked Then
                If txtQty.Value <= 0 Then
                    txtQty.Focus()
                    Throw New Exception("Please Enter Qty")
                End If
                If txtFAT.Value <= 0 Then
                    txtFAT.Focus()
                    Throw New Exception("Please Enter FAT %")
                End If
                If txtSNF.Value <= 0 Then
                    txtSNF.Focus()
                    Throw New Exception("Please Enter SNF %")
                End If
                If clsCommon.myLen(cboMilkType.SelectedValue) <= 0 Then
                    cboMilkType.Focus()
                    Throw New Exception("Please Enter Milk Type")
                End If
                If clsCommon.myLen(TxtFinder1.Value) <= 0 OrElse clsCommon.myLen(TxtFinder1.Tag) <= 0 Then
                    TxtFinder1.Focus()
                    Throw New Exception("Please Enter VLC")
                End If
                If clsCommon.myLen(txtRoute.Value) <= 0 Then
                    txtRoute.Focus()
                    Throw New Exception("Please Enter route")
                End If
                Dim obj As New clsMilkProcurementUploaderHead()
                obj.Document_No = "" ''To be Generated
                obj.Document_Date = clsCommon.GETSERVERDATE()
                obj.Description = "Missing Sample Added By " + objCommonVar.CurrentUserCode + "[" + objCommonVar.CurrentUser + "]"
                obj.MCC_Code = txtMCC.Value
                obj.Dock_Code = ""
                obj.Reject = False
                obj.Arr = New List(Of clsMilkProcurementUploaderDetail)

                Dim objTr As New clsMilkProcurementUploaderDetail()
                objTr.SNo = 1
                objTr.Shift_Date = txtShiftDate.Value
                objTr.Shift = clsCommon.myCstr(cboShift.SelectedValue)
                objTr.Dock_Collection_Milk_Type = clsCommon.myCstr(cboMilkType.SelectedValue)
                objTr.VLC_Code = clsCommon.myCstr(TxtFinder1.Tag)
                objTr.No_Of_Cans = 1
                objTr.Milk_Weight = txtQty.Value
                objTr.FAT = Math.Round(txtFAT.Value, 1, MidpointRounding.ToEven)
                objTr.SNF = Math.Round(txtSNF.Value, IIf(objCommonVar.MilkProcurementSNF2DecimalPlaces, 2, 1), MidpointRounding.ToEven)
                'objTr.Reject_Defaulter = clsCommon.myCstr(gv1.Rows(ii).Cells(colRejectDefaulter).Value)
                objTr.Reject_Type = clsCommon.myCstr(cboRejectType.SelectedValue)
                objTr.Bulk_Route_Code = txtRoute.Value
                'objTr.arrQCParameter = GetParamCollection(ii)
                obj.Arr.Add(objTr)
                If (obj.Arr Is Nothing OrElse obj.Arr.Count <= 0) Then
                    Throw New Exception("Please Fill at list one Item")
                End If
                Dim tran As SqlTransaction = clsDBFuncationality.GetTransactin
                Try
                    obj.SaveData(obj, True, tran)
                    clsMilkProcurementUploaderHead.PostData(obj.Document_No, tran)
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw New Exception(ex.Message)
                End Try
                Dim qry As String = "Select TSPL_MILK_SRN_HEAD.DOC_CODE from  TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL" + Environment.NewLine +
                "left outer join TSPL_MILK_RECEIPT_DETAIL On TSPL_MILK_RECEIPT_DETAIL.Against_Uploader_TR_No=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No" + Environment.NewLine +
                "left outer join TSPL_MILK_SAMPLE_HEAD On TSPL_MILK_SAMPLE_HEAD.MILK_RECEIPT_CODE=TSPL_MILK_RECEIPT_DETAIL.DOC_CODE" + Environment.NewLine +
                "left outer join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.MILK_SAMPLE_CODE=TSPL_MILK_SAMPLE_HEAD.DOC_CODE And TSPL_MILK_SRN_HEAD.SAMPLE_NO=TSPL_MILK_RECEIPT_DETAIL.SAMPLE_NO" + Environment.NewLine +
                "where TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Document_No='" + obj.Document_No + "'"
                lblSRNNo.Text = clsDBFuncationality.getSingleValue(qry)
                If clsCommon.myLen(lblSRNNo.Text) > 0 Then
                    chkAddMissingSample.Checked = False
                End If
            Else
                If clsCommon.myLen(txtSuspenceRemarks.Text) <= 0 Then
                    If chkMarkAsAdulteration.Checked Then
                        txtSuspenceRemarks.Focus()
                        Throw New Exception("Please fill remarks")
                    Else
                        txtSuspenceRemarks.Text = ""
                    End If
                End If

                Dim CorrTypeSRNQty As Boolean = True
                Dim CorrTypeSRNFATSNF As Boolean = True
                Dim CorrTypeSRNVLC As Boolean = True

                If clsCommon.myCdbl(txtQty.Tag) = txtQty.Value Then
                    CorrTypeSRNQty = False
                End If
                If clsCommon.myCdbl(txtFAT.Tag) = txtFAT.Value AndAlso clsCommon.myCdbl(txtSNF.Tag) = txtSNF.Value AndAlso clsCommon.CompairString(clsCommon.myCstr(cboMilkType.Tag), clsCommon.myCstr(cboMilkType.SelectedValue)) = CompairStringResult.Equal Then
                    CorrTypeSRNFATSNF = False
                End If
                If clsCommon.CompairString(clsCommon.myCstr(TxtFinder1.Tag), clsCommon.myCstr(txtVLC.Tag)) = CompairStringResult.Equal Then
                    CorrTypeSRNVLC = False
                End If
                Dim qry As String = "select TSPL_MILK_COLLECTION_DCS_DETAIL.Suspence_VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name
from TSPL_MILK_SRN_DETAIL
left outer join TSPL_MILK_SRN_HEAD on TSPL_MILK_SRN_HEAD.DOC_CODE=TSPL_MILK_SRN_DETAIL.DOC_CODE
left outer join TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL on TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Uploader_TR_No
left outer join TSPL_MILK_SHIFT_UPLOADER_DETAIL on TSPL_MILK_SHIFT_UPLOADER_DETAIL.TR_No=TSPL_MILK_SRN_HEAD.Against_Shift_Uploader_TR_No 
left outer join TSPL_MILK_COLLECTION_DCS_DETAIL on TSPL_MILK_COLLECTION_DCS_DETAIL.PK_Id=TSPL_MILK_PROCUREMENT_UPLOADER_DETAIL.Against_Milk_Collection_DCS_Detail or TSPL_MILK_COLLECTION_DCS_DETAIL.PK_Id=TSPL_MILK_SHIFT_UPLOADER_DETAIL.Against_Milk_Collection_DCS_Detail
left outer join TSPL_VLC_MASTER_HEAD  on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MILK_COLLECTION_DCS_DETAIL.Suspence_VLC_Code
where TSPL_MILK_SRN_HEAD.DOC_CODE='" + lblSRNNo.Text + "' and TSPL_MILK_COLLECTION_DCS_DETAIL.Suspence=1 "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    If Not clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Suspence_VLC_Code")), clsCommon.myCstr(TxtFinder1.Tag)) = CompairStringResult.Equal Then
                        qry = "SRN [" + lblSRNNo.Text + "] is belongs to suspence.Original DCS is " + clsCommon.myCstr(dt.Rows(0)("VLC_Code_VLC_Uploader")) + " [" + clsCommon.myCstr(dt.Rows(0)("VLC_Name")) + "] is diffent from selected DCS " + Environment.NewLine + " Do you want to continue ? "
                        If clsCommon.MyMessageBoxShow(Me, qry, Me.Text, MessageBoxButtons.OK, RadMessageIcon.Error) Then
                            Exit Sub
                        End If
                    End If
                End If
                Dim tran As SqlTransaction = clsDBFuncationality.GetTransactin
                Try
                    clsMilkSRNMCC.Correction(lblSRNNo.Text, CorrTypeSRNQty, CorrTypeSRNFATSNF, CorrTypeSRNVLC, txtQty.Value, clsCommon.myCstr(cboMilkType.SelectedValue), txtFAT.Value, txtSNF.Value, TxtFinder1.Value, False, tran, False, Form_ID, clsCommon.myCstr(cboRejectType.SelectedValue), Remark, chkMarkAsSuspence.Checked, chkMarkAsAdulteration.Checked, txtSuspenceRemarks.Text, txtRoute.Value)
                    tran.Commit()
                Catch ex As Exception
                    tran.Rollback()
                    Throw New Exception(ex.Message)
                End Try
            End If
            clsCommon.MyMessageBoxShow(Me, "Data corrected sucessfully", Me.Text)
            btnSave.Enabled = False
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub btnnew_Click_1(sender As Object, e As EventArgs) Handles btnnew.Click
        lblSRNNo.Text = Nothing
        txtQty.Value = Nothing
        txtQty.Tag = Nothing
        txtFAT.Value = Nothing
        txtFAT.Tag = Nothing
        txtSNF.Value = Nothing
        txtSNF.Tag = Nothing
        cboMilkType.SelectedValue = Nothing
        cboMilkType.Tag = Nothing
        cboRejectType.SelectedValue = Nothing
        cboRejectType.Tag = Nothing
        txtRoute.Value = Nothing

        TxtFinder1.Value = Nothing
        TxtFinder1.Tag = Nothing
        MyLabel5.Text = Nothing
        RadGroupBox2.Enabled = True
        RadGroupBox1.Enabled = False
        btnSave.Enabled = True
        chkRetesting.Checked = False
        chkRetesting.Visible = False
        chkCorrection.Visible = True
        chkCorrection.Checked = True
    End Sub
    Private Sub btnclose_Click(sender As Object, e As EventArgs) Handles btnclose.Click
        CloseForm()
    End Sub

    Sub CloseForm()
        Me.Close()
        GC.Collect()
    End Sub



















End Class