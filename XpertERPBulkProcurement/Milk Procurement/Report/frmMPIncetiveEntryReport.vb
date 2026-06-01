Imports common
Imports System.ComponentModel
Imports System.IO

Public Class frmMPIncetiveEntryReport
    Inherits FrmMainTranScreen

    Dim dt As DataTable = Nothing
    Dim arr As New Dictionary(Of Integer, DataRow)
    Dim strColumnForTotal As String = Nothing
    Dim SettMPIncentiveEntryApplyMonthly As Boolean = False
    'Dim ApplyZoneWiseVSP As Boolean = False
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnExp.Visible = MyBase.isExport
    End Sub
    Private Sub rptMilkBillProcurementSummary_Load(sender As Object, e As EventArgs) Handles Me.Load
        SettMPIncentiveEntryApplyMonthly = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MPIncentiveEntryApplyMonthly, clsFixedParameterCode.MPIncentiveEntryApplyMonthly, Nothing))
        'ApplyZoneWiseVSP = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyZoneWiseVSP, clsFixedParameterCode.ApplyZoneWiseVSP, Nothing)) > 0)
        ReportType()
        If Not clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDP") = CompairStringResult.Equal Then
            lblType.Visible = False
            ddlType.Visible = False
        End If
        Reset()
        '       If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
        '           Dim qry As String = " select TSPL_USER_CUSTOMER_ZONE.Zone_Code as Code , TSPL_ZONE_MASTER.Description as Name from TSPL_ZONE_MASTER  left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code=TSPL_USER_CUSTOMER_ZONE.Zone_Code
        'where 2=2 "
        '           If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
        '               qry += "  and TSPL_ZONE_MASTER.Zone_Code in (" + objCommonVar.strCurrUserZones + ") "
        '           End If
        '           Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        '           If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '               Dim arr As New ArrayList
        '               For Each dr As DataRow In dt.Rows
        '                   arr.Add(dr("Code"))
        '               Next
        '               txtZone.arrValueMember = arr
        '               txtZone.Enabled = False
        '           End If
        '       End If
        If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
            Dim qry As String = " select TSPL_ZONE_MASTER.Zone_Code as Code , TSPL_ZONE_MASTER.Description as Name from TSPL_ZONE_MASTER where 2=2 "
            If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
                qry += "  and TSPL_ZONE_MASTER.Zone_Code in (" + objCommonVar.strCurrUserZones + ") "
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim arr As New ArrayList
                For Each dr As DataRow In dt.Rows
                    arr.Add(dr("Code"))
                Next
                txtZone.arrValueMember = arr
                txtZone.Enabled = False
            End If
        End If

    End Sub

    Private Sub ReportType()
        Dim dt As DataTable
        Dim dr As DataRow
        dt = New DataTable()
        dt.Columns.Add(New DataColumn("Code", System.Type.GetType("System.String")))

        dr = dt.NewRow()
        dr("Code") = "All"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Mineral Mixture"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Pashu Aahar"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Rahat Kampekat Feed"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Code") = "Sailej"
        dt.Rows.Add(dr)

        ddlType.DataSource = dt
        ddlType.ValueMember = "Code"
        ddlType.DisplayMember = "Code"
        ddlType.SelectedIndex = 0
    End Sub

    Sub Reset()
        txtFromDate.Value = clsCommon.GETSERVERDATE()
        txtToDate.Value = clsCommon.GETSERVERDATE()
        btnGo.Enabled = True
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()
        dt = Nothing
        arr = New Dictionary(Of Integer, DataRow)
        strColumnForTotal = Nothing
        txtMCC.arrDispalyMember = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
    End Sub
    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        GetReportID()
        Print(False)
    End Sub

    Sub GetReportID()
        Dim VarID As String = ""
        If rbtnBankWise.IsChecked Then
            VarID += "_BW"
        ElseIf rbtnMCCWise.IsChecked Then
            VarID += "_MW"
        ElseIf rbtnDetails.IsChecked Then
            VarID += "_D"
        ElseIf rbtnNEFT.IsChecked Then
            VarID += "_N"
        ElseIf rbtnPerNEFT.IsChecked Then
            VarID += "_PN"
        ElseIf rbtnDuplicateACNo.IsChecked Then
            VarID += "_DAN"
        ElseIf rbtnDuplicateJanAdharNo.IsChecked Then
            VarID += "_DJA"
        ElseIf rbtnFarmerBankWiseDetail.IsChecked Then
            VarID += "_FBWD"
        ElseIf rbtnFarmerBankWiseSummary.IsChecked Then
            VarID += "_FBWS"
        End If
        Gv1.VarID = VarID
    End Sub

    Sub Print(ByVal isPrint As Boolean)
        Try
            ' If SetToDate() Then
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.DataSource = Nothing
            Gv1.Rows.Clear()
            Gv1.Columns.Clear()
            Gv1.GroupDescriptors.Clear()
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterView.Refresh()


            PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID)
            If rbtnBankWise.IsChecked = True Then
                PageSetupReport_ID = PageSetupReport_ID + "_B"
            ElseIf rbtnMCCWise.IsChecked = True Then
                PageSetupReport_ID = PageSetupReport_ID + "_M" + clsCommon.myCstr(ddlType.SelectedValue)
            ElseIf rbtnNEFT.IsChecked = True Then
                PageSetupReport_ID = PageSetupReport_ID + "_N"
            ElseIf rbtnNEFT.IsChecked = True Then
                PageSetupReport_ID = PageSetupReport_ID + "_PN"
            End If


            'If txtMCC.arrValueMember Is Nothing OrElse txtMCC.arrValueMember.Count <= 0 Then
            '    txtMCC.Focus()
            '    Throw New Exception("Please select MCC")
            'End If
            Dim whre As String = ""
            If txtZone.arrValueMember IsNot Nothing AndAlso txtZone.arrValueMember.Count > 0 Then
                whre += " and  TSPL_VENDOR_MASTER.Zone_Code in (" + clsCommon.GetMulcallString(txtZone.arrValueMember) + ") "
            End If
            If txtBlock.arrValueMember IsNot Nothing AndAlso txtBlock.arrValueMember.Count > 0 Then
                whre += " and  TSPL_MP_MASTER.BLOCK_CODE in (" + clsCommon.GetMulcallString(txtBlock.arrValueMember) + ") "
            End If
            If txtRevenueVillage.arrValueMember IsNot Nothing AndAlso txtRevenueVillage.arrValueMember.Count > 0 Then
                whre += " and  TSPL_MP_MASTER.REVENUE_VILLAGE_CODE in (" + clsCommon.GetMulcallString(txtRevenueVillage.arrValueMember) + ") "
            End If
            If txtGrampanchayat.arrValueMember IsNot Nothing AndAlso txtGrampanchayat.arrValueMember.Count > 0 Then
                whre += " and  TSPL_MP_MASTER.GRAMPANCHAYAT_CODE in (" + clsCommon.GetMulcallString(txtGrampanchayat.arrValueMember) + ") "
            End If

            If txtPanchayatSamiti.arrValueMember IsNot Nothing AndAlso txtPanchayatSamiti.arrValueMember.Count > 0 Then
                whre += " and  TSPL_MP_MASTER.PANCHAYAT_SAMITI_CODE in (" + clsCommon.GetMulcallString(txtPanchayatSamiti.arrValueMember) + ") "
            End If

            If txtVidhanSabha.arrValueMember IsNot Nothing AndAlso txtVidhanSabha.arrValueMember.Count > 0 Then
                whre += " and  TSPL_MP_MASTER.VIDHAN_SABHA_CODE in (" + clsCommon.GetMulcallString(txtVidhanSabha.arrValueMember) + ") "
            End If
            If txtDistrict.arrValueMember IsNot Nothing AndAlso txtDistrict.arrValueMember.Count > 0 Then
                whre += " and  TSPL_MP_MASTER.DISTRICT_Code in (" + clsCommon.GetMulcallString(txtDistrict.arrValueMember) + ") "
            End If
            'If ApplyZoneWiseVSP = True Then
            '    If Not clsCommon.CompairString(objCommonVar.CurrentUserCode, "Admin") = CompairStringResult.Equal Then
            '        whre += " and TSPL_VENDOR_MASTER.Zone_Code in  (" + objCommonVar.strCurrUserZones + ")"
            '    End If
            'End If

            Dim BaseQry As String = "select TSPL_COMPANY_MASTER.Comp_Name,TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,
TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end  as Comp_address,
TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code as Doc_No,convert(varchar, TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,103) +' - '+ convert(varchar,TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date,103) as Date_Range,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code,tspl_MCC_Master.MCC_Name,TSPL_VLC_MASTER_HEAD.Route_Code,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME ,TSPL_VENDOR_MASTER.SupervisorOrRP,RPMaster.Emp_Name as SupervisorOrRPName,TSPL_MP_MASTER.Jan_Aadhar_No,TSPL_MP_MASTER.JA_acc,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MP_MASTER.MP_Code,TSPL_MP_MASTER.BankStateCode,TSPL_MP_MASTER.MP_Code_VLC_Uploader as VLC_CODE_Uploader,TSPL_MP_MASTER.PayeeName as Payee_Joint_Name,case when isnull (TSPL_MP_MASTER.Jan_Aadhar_No_Verified,'')>0 then 'Y' else 'N' end as [Jan Aadhar Verified],TSPL_MP_MASTER.JA_aadhar, TSPL_MP_MASTER.JA_age, TSPL_MP_MASTER.JA_bankBranch, TSPL_MP_MASTER.JA_bankName, TSPL_MP_MASTER.JA_ifsc, TSPL_MP_MASTER.JA_caste, TSPL_MP_MASTER.JA_categoryDescEng, TSPL_MP_MASTER.JA_disability, TSPL_MP_MASTER.JA_disabilityPercentage, TSPL_MP_MASTER.JA_disabilityType, TSPL_MP_MASTER.JA_dlNo, TSPL_MP_MASTER.JA_dob, TSPL_MP_MASTER.JA_eid, TSPL_MP_MASTER.JA_email, TSPL_MP_MASTER.JA_enrId, TSPL_MP_MASTER.JA_fnameEng, TSPL_MP_MASTER.JA_fnameHnd, TSPL_MP_MASTER.JA_gender, TSPL_MP_MASTER.JA_income, TSPL_MP_MASTER.JA_isorphan, TSPL_MP_MASTER.JA_isStateGovtEmp, TSPL_MP_MASTER.JA_jan_mid, TSPL_MP_MASTER.JA_janaadhaarId, TSPL_MP_MASTER.JA_maritalStatus, TSPL_MP_MASTER.JA_micr, TSPL_MP_MASTER.JA_mnameEng, TSPL_MP_MASTER.JA_mnameHnd, TSPL_MP_MASTER.JA_mobile, TSPL_MP_MASTER.JA_nameEng, TSPL_MP_MASTER.JA_nameHnd, TSPL_MP_MASTER.JA_occupation, TSPL_MP_MASTER.JA_relationTyp, TSPL_MP_MASTER.JA_panNo, TSPL_MP_MASTER.JA_passport, TSPL_MP_MASTER.JA_qualification, TSPL_MP_MASTER.JA_rghs_no, TSPL_MP_MASTER.JA_snameEng, TSPL_MP_MASTER.JA_snameHnd, TSPL_MP_MASTER.JA_voterId,TSPL_MP_MASTER.BankName as Bank_Code,TSPL_MP_MASTER.BankName as Bank_Code_Desc,TSPL_MP_MASTER.AccountNO as Payee_Joint_Account_No,TSPL_MP_MASTER.IFCICode as Payee_Joint_IFSC_Code,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Amount as Payable_Amount,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_Entry_Source,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_By,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_Date,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Modified_Entry_Source,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Modified_By,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Modified_Date,case when TABReco.Document_Code is null then 'Pending' else TABReco.Reco_Status end as Reco_Status,TABReco.Document_Code as Reco_Document_Code 
,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mineral_Mixture_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mineral_Mixture_Amount,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code as DocumentCode
,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Pashu_Aahar_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Pashu_Aahar_Amount
,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Rahat_Kampekat_Feed_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Rahat_Kampekat_Feed_Amount
,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Sailej_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Sailej_Amount, 
TSPL_MP_MASTER.DISTRICT_Code as [District Code],TSPL_DISTRICT_MASTER.Name as [District Name],TSPL_MP_MASTER.Zone_Code as [Zone Code], TSPL_ZONE_MASTER.Description as [Zone Name],TSPL_MP_MASTER.BLOCK_CODE as [Block Code],TSPL_BLOCK_MASTER.BLOCK_NAME as [Block Name] ,TSPL_MP_MASTER.REVENUE_VILLAGE_CODE as [Revenue Village Code],TSPL_REVENUE_VILLAGE_MASTER.REVENUE_VILLAGE_NAME as [Revenue Village Name],TSPL_MP_MASTER.GRAMPANCHAYAT_CODE as [Grampanchayat Code],TSPL_GRAMPANCHAYAT_MASTER.GRAMPANCHAYAT_NAME as [Grampanchayat Name],TSPL_MP_MASTER.PANCHAYAT_SAMITI_CODE as [Panchayat Samiti Code],TSPL_PANCHAYAT_SAMITI_MASTER.PANCHAYAT_SAMITI_NAME as [Panchayat Samiti Name],TSPL_MP_MASTER.VIDHAN_SABHA_CODE as [Vidhan Sabha Code], TSPL_VIDHAN_SABHA_MASTER.VIDHAN_SABHA_NAME as [Vidhan Sabha Name],FORMAT((TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date), 'MM-yyyy') as [Created Month-Year],TSPL_MP_MASTER.THIRD_PARTY_CODE
from TSPL_MP_INCENTIVE_ENTRY_DETAIL 
left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
left outer join tspl_MCC_Master on tspl_MCC_Master.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
left outer join(select MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No,Document_Code,'Done' as Reco_Status from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL
union all 
select MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No,Document_Code,'Invalid' as Reco_Status from (
select  MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No,max(Document_Code) as Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID 
group by MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No
)xxx 
where not exists(select 1 from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.MCC_Code=xxx.MCC_Code and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code=xxx.VLC_Code and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Year=xxx.Cycle_Year and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Month=xxx.Cycle_Month and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_No=xxx.Cycle_No)) as TABReco on TABReco.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code and TABReco.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code and TABReco.Cycle_Year=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Year and TABReco.Cycle_Month=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Month and TABReco.Cycle_No=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_No
left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code = TSPL_MP_MASTER.Zone_Code
left outer join TSPL_BLOCK_MASTER on TSPL_BLOCK_MASTER.BLOCK_CODE = TSPL_MP_MASTER.BLOCK_CODE
left outer join TSPL_REVENUE_VILLAGE_MASTER on TSPL_REVENUE_VILLAGE_MASTER.REVENUE_VILLAGE_CODE = TSPL_MP_MASTER.REVENUE_VILLAGE_CODE
left outer join TSPL_GRAMPANCHAYAT_MASTER on TSPL_GRAMPANCHAYAT_MASTER.GRAMPANCHAYAT_CODE = TSPL_MP_MASTER.GRAMPANCHAYAT_CODE
left outer join TSPL_PANCHAYAT_SAMITI_MASTER on TSPL_PANCHAYAT_SAMITI_MASTER.PANCHAYAT_SAMITI_CODE = TSPL_MP_MASTER.PANCHAYAT_SAMITI_CODE
left outer join TSPL_VIDHAN_SABHA_MASTER on TSPL_VIDHAN_SABHA_MASTER.VIDHAN_SABHA_CODE = TSPL_MP_MASTER.VIDHAN_SABHA_CODE
left outer join TSPL_DISTRICT_MASTER on TSPL_DISTRICT_MASTER.Code = TSPL_MP_MASTER.DISTRICT_Code
left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =  TSPL_VLC_MASTER_HEAD.VSP_Code
left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code
left outer join TSPL_EMPLOYEE_MASTER as RPMaster on RPMaster.EMP_CODE=TSPL_VENDOR_MASTER.SupervisorOrRP
where 2=2 and CONVERT(date,TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' and CONVERT(date,TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' " + whre + " "
            If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                BaseQry += " and TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code in (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ") "
            End If
            If rbtnMCCWise.IsChecked Or rbtnDetails.IsChecked Then
                If clsCommon.CompairString(ddlType.SelectedValue, "Mineral Mixture") = CompairStringResult.Equal Then
                    BaseQry += " and isnull(TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mineral_Mixture_Amount,0) > 0 "
                ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Pashu Aahar") = CompairStringResult.Equal Then
                    BaseQry += " and isnull(TSPL_MP_INCENTIVE_ENTRY_DETAIL.Pashu_Aahar_Amount,0) > 0 "
                ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Rahat Kampekat Feed") = CompairStringResult.Equal Then
                    BaseQry += " and isnull(TSPL_MP_INCENTIVE_ENTRY_DETAIL.Rahat_Kampekat_Feed_Amount,0) > 0 "
                ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Sailej") = CompairStringResult.Equal Then
                    BaseQry += " and isnull(TSPL_MP_INCENTIVE_ENTRY_DETAIL.Sailej_Amount,0) > 0 "
                End If
            End If

            Dim Qry As String = ""
            strColumnForTotal = ""
            If rbtnBankWise.IsChecked Then
                strColumnForTotal = "Bank_Code"
                Qry = "select max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address,max(Date_Range) as Date_Range, MCC_Code,max(MCC_Name) as MCC_Name,max(Route_Code) as Route_Code,max(ROUTE_NAME) as ROUTE_NAME,max(SupervisorOrRP) as SupervisorOrRP,max(SupervisorOrRPName) as  SupervisorOrRPName,VLC_Code,max(VLC_Code_VLC_Uploader) as VLC_Code_VLC_Uploader,max(VLC_Name) as VLC_Name,MP_Code,max(VLC_CODE_Uploader) as VLC_CODE_Uploader,max(Payee_Joint_Name) as Payee_Joint_Name,max([District Code]) as [District Code],max([District Name]) as [District Name] ,max([Zone Code]) as [Zone Code], max([Zone Name]) as [Zone Name],max([Block Code]) as [Block Code],max([Block Name]) as [Block Name] ,max([Revenue Village Code]) as [Revenue Village Code],max([Revenue Village Name]) as [Revenue Village Name],max([Grampanchayat Code]) as [Grampanchayat Code],max([Grampanchayat Name]) as [Grampanchayat Name],max([Panchayat Samiti Code]) as [Panchayat Samiti Code],max([Panchayat Samiti Name]) as [Panchayat Samiti Name],max([Vidhan Sabha Code]) as [Vidhan Sabha Code], max([Vidhan Sabha Name]) as [Vidhan Sabha Name],max([Created Month-Year]) as [Created Month-Year],max(THIRD_PARTY_CODE) as THIRD_PARTY_CODE,Bank_Code,max(Bank_Code_Desc) as Bank_Code_Desc,max(Payee_Joint_Account_No) as Payee_Joint_Account_No,max(Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,sum(Qty) as Quantity,sum(Payable_Amount) as Payable_Amount,max(THIRD_PARTY_CODE) as [Third Party Code] from (" + BaseQry + ")xx group by MCC_Code,VLC_Code,MP_Code,Bank_Code order by [Created Month-Year],Bank_Code,MCC_Code, VLC_Code_VLC_Uploader,VLC_CODE_Uploader"

            ElseIf rbtnMCCWise.IsChecked Or rbtnDetails.IsChecked Then
                strColumnForTotal = "MCC_Code"
                Qry = "select ROW_NUMBER() OVER(ORDER BY (SELECT 1)) As S_No,xx.DocumentCode, max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address,xx.From_Date,max(Date_Range) as Date_Range, MCC_Code,max(MCC_Name) as MCC_Name,max(Route_Code) as Route_Code,max(ROUTE_NAME) as ROUTE_NAME,max(SupervisorOrRP) as SupervisorOrRP,max(SupervisorOrRPName) as  SupervisorOrRPName,max(Jan_Aadhar_No) as Jan_Aadhar_No,max(JA_acc)Jan_Aadhar_Account_No,VLC_Code,max(VLC_Code_VLC_Uploader) as VLC_Code_VLC_Uploader,max(VLC_Name) as VLC_Name,MP_Code,max(VLC_CODE_Uploader) as VLC_CODE_Uploader,max(Payee_Joint_Name) as Payee_Joint_Name,max([Jan Aadhar Verified]) as Jan_Aadhar_No_Verified,max([JA_bankBranch]) as [JA_bankBranch], max ([JA_bankName]) as [JA_bankName], max ([JA_ifsc]) as [JA_ifsc], max ([JA_caste]) as [JA_caste], max ([JA_categoryDescEng]) as [JA_categoryDescEng], max ([JA_disability]) as [JA_disability], max ([JA_disabilityPercentage]) as [JA_disabilityPercentage], max ([JA_disabilityType]) as [JA_disabilityType],
max ([JA_dlNo]) as [JA_dlNo], max ([JA_dob]) as [JA_dob], max ([JA_eid]) as [JA_eid], max ([JA_email]) as [JA_email], max ([JA_enrId]) as [JA_enrId], max ([JA_fnameEng]) as [JA_fnameEng], max ([JA_fnameHnd]) as [JA_fnameHnd], max ([JA_gender]) as [JA_gender], max ([JA_income]) as [JA_income], max ([JA_isorphan]) as [JA_isorphan], max ([JA_isStateGovtEmp]) as [JA_isStateGovtEmp], max ([JA_jan_mid]) as [JA_jan_mid], max ([JA_janaadhaarId]) as [JA_janaadhaarId], max ([JA_maritalStatus]) as [JA_maritalStatus], max ([JA_micr]) as [JA_micr], max ([JA_mobile]) as [JA_mobile], max ([JA_nameEng]) as [JA_nameEng], max ([JA_nameHnd]) as [JA_nameHnd], max ([JA_occupation]) as [JA_occupation], max ([JA_relationTyp]) as [JA_relationTyp], max ([JA_panNo]) as [JA_panNo], max ([JA_passport]) as [JA_passport], max ([JA_qualification]) as [JA_qualification], max ([JA_snameEng]) as [JA_snameEng], max ([JA_snameHnd]) as [JA_snameHnd], max ([JA_snameHnd]) as [JA_snameHnd], max ([JA_voterId]) as [JA_voterId], max([District Code]) as [District Code],max([District Name]) as [District Name] ,max([Zone Code]) as [Zone Code], max([Zone Name]) as [Zone Name],max([Block Code]) as [Block Code],max([Block Name]) as [Block Name] ,max([Revenue Village Code]) as [Revenue Village Code],max([Revenue Village Name]) as [Revenue Village Name],max([Grampanchayat Code]) as [Grampanchayat Code],max([Grampanchayat Name]) as [Grampanchayat Name],max([Panchayat Samiti Code]) as [Panchayat Samiti Code],max([Panchayat Samiti Name]) as [Panchayat Samiti Name],max([Vidhan Sabha Code]) as [Vidhan Sabha Code], max([Vidhan Sabha Name]) as [Vidhan Sabha Name],max([Created Month-Year]) as [Created Month-Year], max(THIRD_PARTY_CODE) as [THIRD PARTY CODE], max(Bank_Code) as Bank_Code,max(Bank_Code_Desc) as Bank_Code_Desc,max(Payee_Joint_Account_No) as Payee_Joint_Account_No,max(Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(BankStateCode) as [Bank State],sum(Qty) as Quantity,  CASE WHEN (SUM(Payable_Amount)) % 1 >= 0.5 THEN CAST(CEILING(SUM(Payable_Amount)) AS DECIMAL(10, 2)) ELSE CAST(FLOOR(SUM(Payable_Amount)) AS DECIMAL(10, 2)) END AS Payable_Amount,SUM(Payable_Amount) as Original_Payable_Amount  ,  max(xx.Created_Entry_Source) as Created_Entry_Source,max(xx.Created_By) as Created_By,max(convert(varchar, xx.Created_Date,103)+' '+SUBSTRING(convert(varchar, xx.Created_Date,100),13,20) ) as Created_Date,max(xx.Modified_Entry_Source) as Modified_Entry_Source,max(xx.Modified_By) as Modified_By,max(convert(varchar, xx.Modified_Date,103)+' '+SUBSTRING(convert(varchar, xx.Modified_Date,100),13,20)) as Modified_Date,max(xx.Reco_Status) as Reco_Status,max(xx.Reco_Document_Code) as Reco_Document_Code
                           ,sum(Mineral_Mixture_Qty) as Mineral_Mixture_Qty,sum(Mineral_Mixture_Amount) as Mineral_Mixture_Amount,sum(Pashu_Aahar_Qty) as Pashu_Aahar_Qty,sum(Pashu_Aahar_Amount) as Pashu_Aahar_Amount,sum(Rahat_Kampekat_Feed_Qty) as Rahat_Kampekat_Feed_Qty,sum(Rahat_Kampekat_Feed_Amount) as Rahat_Kampekat_Feed_Amount,sum(Sailej_Qty) as Sailej_Qty,sum(Sailej_Amount) as Sailej_Amount, "

                If clsCommon.CompairString(objCommonVar.CurrDatabase, "BKN") = CompairStringResult.Equal Then
                    Qry += " CASE  WHEN (sum(Payable_Amount) + sum(Mineral_Mixture_Amount) + sum(Pashu_Aahar_Amount) + sum(Rahat_Kampekat_Feed_Amount) + sum(Sailej_Amount)) % 1 >= 0.5    THEN CEILING(sum(Payable_Amount) + sum(Mineral_Mixture_Amount) + sum(Pashu_Aahar_Amount) + sum(Rahat_Kampekat_Feed_Amount) + sum(Sailej_Amount))  Else FLOOR(sum(Payable_Amount) + sum(Mineral_Mixture_Amount) + sum(Pashu_Aahar_Amount) + sum(Rahat_Kampekat_Feed_Amount) + sum(Sailej_Amount)) End As Total_Amount "
                Else
                    Qry += " (sum(Payable_Amount)+sum(Mineral_Mixture_Amount)+sum(Pashu_Aahar_Amount)+sum(Rahat_Kampekat_Feed_Amount)+sum(Sailej_Amount)) as Total_Amount "
                End If


                Qry += " ,max(THIRD_PARTY_CODE)THIRD_PARTY_CODE from (" + BaseQry + ")xx group by MCC_Code,VLC_Code,MP_Code,from_date,DocumentCode  order by [Created Month-Year],MCC_Code, VLC_Code_VLC_Uploader , VLC_CODE_Uploader"
            ElseIf rbtnNEFT.IsChecked Then
                Qry = "select   TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code,TSPL_MCC_MASTER.MCC_Name ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_Code_VLC_Uploader,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as VLC_CODE_Uploader,TSPL_MP_MASTER.PayeeName as Payee_Joint_Name,TSPL_MP_MASTER.BankName as Bank_Code,TSPL_MP_MASTER.BankName as Bank_Code_Desc,TSPL_MP_MASTER.AccountNO as Payee_Joint_Account_No, TSPL_MP_MASTER.IFCICode as Payee_Joint_IFSC_Code,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty as Quantity,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Amount as Payable_Amount,(case when  isnull(TabLatestDBT.DBT_NEFT_REJECT_DETAIL_PK_Id,0)>0 and isnull(TabLatestDBT.DBT_NEFT_REJECT_Status,0)=0 then 'DBT Rejected' else (case when isnull(TabLatestDBT.DBT_NEFT_REJECT_Status,0)=1 then 'Pending For NEFT' else  (case when isnull(TabLatestDBT.DBT_NEFT_Status,0)=1 then 'DBT Succeed' else 'Pending For NEFT'  end) end)  end) as Status  from 
TSPL_MP_INCENTIVE_ENTRY_DETAIL
Left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
left outer join (    
select TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR,TSPL_DBT_NEFT_DETAIL.PK_Id as DBT_NEFT_PK_Id,TSPL_DBT_NEFT.Status as DBT_NEFT_Status,TSPL_DBT_NEFT_REJECT_DETAIL.PK_Id as DBT_NEFT_REJECT_DETAIL_PK_Id,TSPL_DBT_NEFT_REJECT.Status as DBT_NEFT_REJECT_Status from 
(select TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR,max(TSPL_DBT_NEFT_DETAIL.PK_Id) as PK_Id from TSPL_DBT_NEFT_DETAIL group by Against_MP_Incentive_TR  ) as x
inner join TSPL_DBT_NEFT_DETAIL on TSPL_DBT_NEFT_DETAIL.PK_Id=x.PK_Id
left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code=TSPL_DBT_NEFT_DETAIL.Document_Code
left outer join TSPL_DBT_NEFT_REJECT_DETAIL on TSPL_DBT_NEFT_REJECT_DETAIL.Against_DBT_NEFT_TR=TSPL_DBT_NEFT_DETAIL.PK_Id
left outer join TSPL_DBT_NEFT_REJECT on TSPL_DBT_NEFT_REJECT.Document_Code=TSPL_DBT_NEFT_REJECT_DETAIL.Document_Code 
) as TabLatestDBT on TabLatestDBT.Against_MP_Incentive_TR=TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id
where TSPL_MP_INCENTIVE_ENTRY_HEAD.Status=1  and CONVERT(date,TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,103)='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' and CONVERT(date,TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date,103)='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' "
                If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                    Qry += " and TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code in (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ")"
                End If

            ElseIf rbtnDuplicateACNo.IsChecked Then
                Qry = "WITH CTE AS (SELECT *, COUNT(*) OVER (PARTITION BY Payee_Joint_Account_No) AS DuplicateCount FROM 
                (select ROW_NUMBER() OVER(PARTITION BY TSPL_MP_MASTER.AccountNo ORDER BY TSPL_MP_MASTER.AccountNO) AS S_No, TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code as VLC_Code,TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_Code_VLC_Uploader, TSPL_MP_MASTER.AccountNO as Payee_Joint_Account_No, TSPL_COMPANY_MASTER.Comp_Name as Comp_Name,TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date as From_Date,
                TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+ isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end  as Comp_address,
                TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code as Doc_No,convert(varchar, TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,103) +' - '+ convert(varchar,TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date,103) as Date_Range,TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code as MCC_Code ,tspl_MCC_Master.MCC_Name as MCC_Name,TSPL_VLC_MASTER_HEAD.Route_Code,TSPL_BULK_ROUTE_MASTER.ROUTE_NAME ,TSPL_VENDOR_MASTER.SupervisorOrRP,RPMaster.Emp_Name as SupervisorOrRPName, TSPL_MP_MASTER.Jan_Aadhar_No,  TSPL_MP_MASTER.JA_acc,TSPL_VLC_MASTER_HEAD.VLC_Name,TSPL_MP_MASTER.MP_Code,TSPL_MP_MASTER.MP_Code_VLC_Uploader as VLC_CODE_Uploader,TSPL_MP_MASTER.PayeeName as Payee_Joint_Name, case when isnull (TSPL_MP_MASTER.Jan_Aadhar_No_Verified,'')>0 then 'Y' else 'N' end as [Jan Aadhar Verified],TSPL_MP_MASTER.JA_aadhar, TSPL_MP_MASTER.JA_age, TSPL_MP_MASTER.JA_bankBranch, TSPL_MP_MASTER.JA_bankName, TSPL_MP_MASTER.JA_ifsc, TSPL_MP_MASTER.JA_caste, TSPL_MP_MASTER.JA_categoryDescEng, TSPL_MP_MASTER.JA_disability, TSPL_MP_MASTER.JA_disabilityPercentage, TSPL_MP_MASTER.JA_disabilityType, TSPL_MP_MASTER.JA_dlNo, TSPL_MP_MASTER.JA_dob, TSPL_MP_MASTER.JA_eid, TSPL_MP_MASTER.JA_email, TSPL_MP_MASTER.JA_enrId, TSPL_MP_MASTER.JA_fnameEng, TSPL_MP_MASTER.JA_fnameHnd, TSPL_MP_MASTER.JA_gender, TSPL_MP_MASTER.JA_income, TSPL_MP_MASTER.JA_isorphan, TSPL_MP_MASTER.JA_isStateGovtEmp, TSPL_MP_MASTER.JA_jan_mid, TSPL_MP_MASTER.JA_janaadhaarId, TSPL_MP_MASTER.JA_maritalStatus, TSPL_MP_MASTER.JA_micr, TSPL_MP_MASTER.JA_mnameEng, TSPL_MP_MASTER.JA_mnameHnd, TSPL_MP_MASTER.JA_mobile, TSPL_MP_MASTER.JA_nameEng, TSPL_MP_MASTER.JA_nameHnd, TSPL_MP_MASTER.JA_occupation, TSPL_MP_MASTER.JA_relationTyp, TSPL_MP_MASTER.JA_panNo, TSPL_MP_MASTER.JA_passport, TSPL_MP_MASTER.JA_qualification, TSPL_MP_MASTER.JA_rghs_no, TSPL_MP_MASTER.JA_snameEng, TSPL_MP_MASTER.JA_snameHnd, TSPL_MP_MASTER.JA_voterId,TSPL_MP_MASTER.BankName as Bank_Code,TSPL_MP_MASTER.BankName as Bank_Code_Desc,TSPL_MP_MASTER.IFCICode as Payee_Joint_IFSC_Code,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Amount as Payable_Amount,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_Entry_Source,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_By,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_Date,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Modified_Entry_Source,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Modified_By,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Modified_Date,case when TABReco.Document_Code is null then 'Pending' else TABReco.Reco_Status end as Reco_Status,TABReco.Document_Code as Reco_Document_Code 
                ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mineral_Mixture_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Mineral_Mixture_Amount,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code as DocumentCode
                ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Pashu_Aahar_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Pashu_Aahar_Amount
                ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Rahat_Kampekat_Feed_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Rahat_Kampekat_Feed_Amount
                ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Sailej_Qty ,TSPL_MP_INCENTIVE_ENTRY_DETAIL.Sailej_Amount, 
                TSPL_MP_MASTER.DISTRICT_Code as [District Code],TSPL_DISTRICT_MASTER.Name as [District Name],TSPL_MP_MASTER.Zone_Code as [Zone Code], TSPL_ZONE_MASTER.Description as [Zone Name],TSPL_MP_MASTER.BLOCK_CODE as [Block Code],TSPL_BLOCK_MASTER.BLOCK_NAME as [Block Name] ,TSPL_MP_MASTER.REVENUE_VILLAGE_CODE as [Revenue Village Code],TSPL_REVENUE_VILLAGE_MASTER.REVENUE_VILLAGE_NAME as [Revenue Village Name],TSPL_MP_MASTER.GRAMPANCHAYAT_CODE as [Grampanchayat Code],TSPL_GRAMPANCHAYAT_MASTER.GRAMPANCHAYAT_NAME as [Grampanchayat Name],TSPL_MP_MASTER.PANCHAYAT_SAMITI_CODE as [Panchayat Samiti Code],TSPL_PANCHAYAT_SAMITI_MASTER.PANCHAYAT_SAMITI_NAME as [Panchayat Samiti Name],TSPL_MP_MASTER.VIDHAN_SABHA_CODE as [Vidhan Sabha Code], TSPL_VIDHAN_SABHA_MASTER.VIDHAN_SABHA_NAME as [Vidhan Sabha Name],FORMAT((TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date), 'MM-yyyy') as [Created Month-Year],TSPL_MP_MASTER.THIRD_PARTY_CODE
                from TSPL_MP_INCENTIVE_ENTRY_DETAIL 
                left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
                left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
                left outer join TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code
                left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
                left outer join tspl_MCC_Master on tspl_MCC_Master.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code
                left outer join(select MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No,Document_Code,'Done' as Reco_Status from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL
                union all 
                select MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No,Document_Code,'Invalid' as Reco_Status from (
                select  MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No,max(Document_Code) as Document_Code from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL_INVALID 
                group by MCC_Code,VLC_Code,Cycle_Year,Cycle_Month,Cycle_No
                )xxx 
                where not exists(select 1 from TSPL_DCS_MP_INCENTIVE_RECO_DETAIL where TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.MCC_Code=xxx.MCC_Code and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.VLC_Code=xxx.VLC_Code and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Year=xxx.Cycle_Year and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_Month=xxx.Cycle_Month and TSPL_DCS_MP_INCENTIVE_RECO_DETAIL.Cycle_No=xxx.Cycle_No)) as TABReco on TABReco.MCC_Code=TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code and TABReco.VLC_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.VLC_Code and TABReco.Cycle_Year=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Year and TABReco.Cycle_Month=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_Month and TABReco.Cycle_No=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Cycle_No
                left outer join TSPL_ZONE_MASTER on TSPL_ZONE_MASTER.Zone_Code = TSPL_MP_MASTER.Zone_Code
                left outer join TSPL_BLOCK_MASTER on TSPL_BLOCK_MASTER.BLOCK_CODE = TSPL_MP_MASTER.BLOCK_CODE
                left outer join TSPL_REVENUE_VILLAGE_MASTER on TSPL_REVENUE_VILLAGE_MASTER.REVENUE_VILLAGE_CODE = TSPL_MP_MASTER.REVENUE_VILLAGE_CODE
                left outer join TSPL_GRAMPANCHAYAT_MASTER on TSPL_GRAMPANCHAYAT_MASTER.GRAMPANCHAYAT_CODE = TSPL_MP_MASTER.GRAMPANCHAYAT_CODE
                left outer join TSPL_PANCHAYAT_SAMITI_MASTER on TSPL_PANCHAYAT_SAMITI_MASTER.PANCHAYAT_SAMITI_CODE = TSPL_MP_MASTER.PANCHAYAT_SAMITI_CODE
                left outer join TSPL_VIDHAN_SABHA_MASTER on TSPL_VIDHAN_SABHA_MASTER.VIDHAN_SABHA_CODE = TSPL_MP_MASTER.VIDHAN_SABHA_CODE
                left outer join TSPL_DISTRICT_MASTER on TSPL_DISTRICT_MASTER.Code = TSPL_MP_MASTER.DISTRICT_Code
                left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =  TSPL_VLC_MASTER_HEAD.VSP_Code
                left outer join TSPL_BULK_ROUTE_MASTER on TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_VLC_MASTER_HEAD.Route_Code
                left outer join TSPL_EMPLOYEE_MASTER as RPMaster on RPMaster.EMP_CODE=TSPL_VENDOR_MASTER.SupervisorOrRP
                where 2=2 and CONVERT(date,TSPL_MP_INCENTIVE_ENTRY_HEAD.From_Date,103)>='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' and CONVERT(date,TSPL_MP_INCENTIVE_ENTRY_HEAD.To_Date,103)<='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "' " + whre + " )xx   
                )SELECT * FROM CTE WHERE DuplicateCount > 1 ORDER BY Payee_Joint_Account_No"

            ElseIf rbtnDuplicateJanAdharNo.IsChecked Then
                Qry = "select Jan_Aadhar_No,sum(1) as Repeated from (" + BaseQry + ")xx group by Jan_Aadhar_No  having len(isnull(Jan_Aadhar_No,''))>0 and sum(1)>1"
            ElseIf rbtnFarmerBankWiseDetail.IsChecked Then
                Qry = "select  ROW_NUMBER() over (Partition BY max(Bank_Code) order by max(Bank_Code)) as SNo,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address,max(Date_Range) as Date_Range, max(MCC_Code) as MCC_Code,max(MCC_Name) as MCC_Name ,VLC_Code,max(VLC_Code_VLC_Uploader) as VLC_Code_VLC_Uploader,max(VLC_Name) as VLC_Name,MP_Code,max(VLC_CODE_Uploader) as VLC_CODE_Uploader,max(Payee_Joint_Name) as Payee_Joint_Name,max([Jan Aadhar Verified]) as Jan_Aadhar_No_Verified,max(Bank_Code) as Bank_Code,max(Bank_Code_Desc) as Bank_Code_Desc,max(Payee_Joint_Account_No) as Payee_Joint_Account_No,max(Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,sum(Qty) as Quantity,sum(Payable_Amount) as Payable_Amount,max(THIRD_PARTY_CODE)THIRD_PARTY_CODE from (" + BaseQry + ")xx group by VLC_Code,MP_Code order by Bank_Code, VLC_Code_VLC_Uploader,VLC_CODE_Uploader"
            ElseIf rbtnFarmerBankWiseSummary.IsChecked Then
                Qry = "select max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address,max(MCC_Code) as MCC_Code,max(MCC_Name) as MCC_Name,max(Date_Range) as Date_Range,ROW_NUMBER() over ( order by Bank_Code) as SNO,Bank_Code,max(Bank_Code_Desc) as Bank_Code_Desc,sum(Qty) as Quantity,sum(Payable_Amount) as Payable_Amount,max(THIRD_PARTY_CODE)THIRD_PARTY_CODE from (" + BaseQry + ")xx group by Bank_Code order by Bank_Code"
                'Qry = "select max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address,max(MCC_Code) as MCC_Code,max(MCC_Name) as MCC_Name,max(Date_Range) as Date_Range,ROW_NUMBER() over ( order by Bank_Code) as SNO,Bank_Code,max(Bank_Code_Desc) as Bank_Code_Desc,max(Qty) + min(Qty) as Quantity ,max(Payable_Amount)+min(Payable_Amount) as Payable_Amount from (" + BaseQry + ")xx group by Bank_Code order by Bank_Code"
            ElseIf rbtnPerNEFT.IsChecked Then
                    Qry = "SELECT 
                       FORMAT(TSPL_MP_INCENTIVE_ENTRY_DETAIL.Created_Date, 'MMM-dd')Created_Date, TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code ,TSPL_VLC_MASTER_HEAD.VLC_Name,
                      TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code,TSPL_DBT_NEFT_DETAIL.MP_Uploader_Code,TSPL_DBT_NEFT_DETAIL.MP_NAME, 
                      TSPL_MP_INCENTIVE_ENTRY_DETAIL.Qty,TSPL_DBT_NEFT_DETAIL.Amount,TSPL_MP_MASTER.THIRD_PARTY_CODE
                       FROM   TSPL_DBT_NEFT_DETAIL
                      LEFT OUTER JOIN TSPL_DBT_NEFT_REJECT_DETAIL ON TSPL_DBT_NEFT_REJECT_DETAIL.Against_DBT_NEFT_TR = TSPL_DBT_NEFT_DETAIL.PK_Id
	                  LEFT OUTER JOIN TSPL_VLC_MASTER_HEAD ON TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader=TSPL_DBT_NEFT_DETAIL.VLC_Uploader_Code
	                  left outer join TSPL_MP_INCENTIVE_ENTRY_DETAIL on TSPL_MP_INCENTIVE_ENTRY_DETAIL.PK_Id=TSPL_DBT_NEFT_DETAIL.Against_MP_Incentive_TR
                	    left outer join TSPL_MP_INCENTIVE_ENTRY_HEAD on TSPL_MP_INCENTIVE_ENTRY_HEAD.Document_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.Document_Code
                     left outer join TSPL_MP_MASTER on TSPL_MP_MASTER.MP_Code=TSPL_MP_INCENTIVE_ENTRY_DETAIL.MP_Code
				     left outer join TSPL_DBT_NEFT on TSPL_DBT_NEFT.Document_Code=TSPL_DBT_NEFT_DETAIL.Document_Code
                   WHERE 
                     TSPL_DBT_NEFT_DETAIL.PK_Id IS NOT NULL
                     AND TSPL_DBT_NEFT_DETAIL.PK_Id NOT IN (
                     SELECT Against_DBT_NEFT_TR 
                     FROM TSPL_DBT_NEFT_REJECT_DETAIL
                     WHERE Against_DBT_NEFT_TR IS NOT NULL 
    and  TSPL_DBT_NEFT.Status=1  and CONVERT(date,TSPL_DBT_NEFT.From_Date,103)='" + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MMM/yyyy") + "' and CONVERT(date,TSPL_DBT_NEFT.To_Date,103)='" + clsCommon.GetPrintDate(txtToDate.Value, "dd/MMM/yyyy") + "') " + whre + "  

    "
                If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                    Qry += " and TSPL_MP_INCENTIVE_ENTRY_HEAD.MCC_Code in (" + clsCommon.GetMulcallString(txtMCC.arrValueMember) + ")"
                End If
            Else
                Throw New Exception("Wrong Method")
            End If
            dt = Nothing
            dt = clsDBFuncationality.GetDataTable(Qry)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                Throw New Exception("No Data Found to Display")
            End If

            If isPrint Then
                If rbtnFarmerBankWiseDetail.IsChecked Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptDBTFarmerWiseBankAdvice", "Farmer Bank Wise Details")
                    frmCRV = Nothing
                ElseIf rbtnFarmerBankWiseSummary.IsChecked Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptDBTFarmerWiseBankSummary", "Farmer Bank Wise Summary")
                    frmCRV = Nothing
                End If
            End If

            'If Not (rbtnNEFT.IsChecked OrElse rbtnDuplicateACNo.IsChecked OrElse rbtnDuplicateJanAdharNo.IsChecked OrElse rbtnDetails.IsChecked OrElse rbtnFarmerBankWiseDetail.IsChecked OrElse rbtnFarmerBankWiseSummary.IsChecked) Then
            '    AddTotalRows() 'strColumnForTotal, dt
            'End If


            Gv1.DataSource = dt

            Gv1.MasterTemplate.AutoExpandGroups = False

            RadPageView1.SelectedPage = RadPageViewPage2
            SetGridFormat(Gv1)
            EnableDisaableControls(False)
            'If rbtnNEFT.IsChecked Then
            ReStoreGridLayout()
            'End If

            ' End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ReStoreGridLayout()
        Try
            '  If rbtnNEFT.IsChecked Then
            If clsCommon.myLen(PageSetupReport_ID) > 0 Then
                Dim obj As clsGridLayout = New clsGridLayout()
                obj = CType(obj.GetData(PageSetupReport_ID, "", objCommonVar.CurrentUserCode), clsGridLayout)
                If Not obj Is Nothing AndAlso obj.GridColumns >= Gv1.ColumnCount Then
                    Dim ii As Integer
                    For ii = 0 To Gv1.Columns.Count - 1 Step ii + 1
                        Gv1.Columns(ii).IsVisible = False
                        Gv1.Columns(ii).VisibleInColumnChooser = True
                    Next
                    Gv1.LoadLayout(obj.GridLayout)
                    obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                End If
            End If
        Catch err As Exception
            MessageBox.Show(err.Message)
        End Try
    End Sub
    Sub AddTotalRows() 'ByVal strColumnForTotal As String, ByRef dt As DataTable
        'Dim arr As New Dictionary(Of Integer, DataRow)
        arr = New Dictionary(Of Integer, DataRow)
        Dim Total As Decimal = 0
        Dim TotalQty As Decimal = 0
        Dim strPreviousBank As String = clsCommon.myCstr(dt.Rows(0)(strColumnForTotal))
        For ii As Integer = 0 To dt.Rows.Count - 1
            Dim flag As Boolean = False
            If clsCommon.CompairString(strPreviousBank, clsCommon.myCstr(dt.Rows(ii)(strColumnForTotal))) = CompairStringResult.Equal Then
                Total += clsCommon.myCdbl(dt.Rows(ii)("Payable_Amount"))
                TotalQty += clsCommon.myCdbl(dt.Rows(ii)("Quantity"))
            Else
                Dim drTS As DataRow = dt.NewRow()
                drTS(strColumnForTotal) = strPreviousBank
                drTS("Payee_Joint_IFSC_Code") = "Total"
                drTS("Payable_Amount") = Total
                drTS("Quantity") = TotalQty
                arr.Add(ii, drTS)

                Total = clsCommon.myCdbl(dt.Rows(ii)("Payable_Amount"))
                TotalQty = clsCommon.myCdbl(dt.Rows(ii)("Quantity"))
            End If
            strPreviousBank = clsCommon.myCstr(dt.Rows(ii)(strColumnForTotal))
            If dt.Rows.Count - 1 = ii Then
                Dim drTS As DataRow = dt.NewRow()
                drTS(strColumnForTotal) = strPreviousBank
                drTS("Payee_Joint_IFSC_Code") = "Total"
                drTS("Payable_Amount") = Total
                drTS("Quantity") = TotalQty

                If rbtnMCCWise.IsChecked = True Or rbtnDetails.IsChecked = True Then
                    drTS("Mineral_Mixture_Qty") = clsCommon.myCdbl(dt.Compute("SUM(Mineral_Mixture_Qty)", " Mineral_Mixture_Qty is not null"))
                    drTS("Mineral_Mixture_Amount") = clsCommon.myCdbl(dt.Compute("SUM(Mineral_Mixture_Amount)", " Mineral_Mixture_Amount is not null"))
                    drTS("Pashu_Aahar_Qty") = clsCommon.myCdbl(dt.Compute("SUM(Pashu_Aahar_Qty)", " Pashu_Aahar_Qty is not null"))
                    drTS("Pashu_Aahar_Amount") = clsCommon.myCdbl(dt.Compute("SUM(Pashu_Aahar_Amount)", " Pashu_Aahar_Amount is not null"))
                    drTS("Rahat_Kampekat_Feed_Qty") = clsCommon.myCdbl(dt.Compute("SUM(Rahat_Kampekat_Feed_Qty)", " Rahat_Kampekat_Feed_Qty is not null"))
                    drTS("Rahat_Kampekat_Feed_Amount") = clsCommon.myCdbl(dt.Compute("SUM(Rahat_Kampekat_Feed_Amount)", " Rahat_Kampekat_Feed_Amount is not null"))
                    drTS("Sailej_Qty") = clsCommon.myCdbl(dt.Compute("SUM(Sailej_Qty)", " Sailej_Qty is not null"))
                    drTS("Sailej_Amount") = clsCommon.myCdbl(dt.Compute("SUM(Sailej_Amount)", " Sailej_Amount is not null"))
                    drTS("Total_Amount") = clsCommon.myCdbl(dt.Compute("SUM(Total_Amount)", " Total_Amount is not null"))
                End If

                arr.Add(ii + 1, drTS)
            End If

        Next

        For ii As Integer = arr.Count - 1 To 0 Step -1
            Dim Key As Integer = arr.Keys(ii)
            dt.Rows.InsertAt(arr(Key), Key)
        Next

    End Sub
    Sub SetGridFormat(ByRef Gv1 As RadGridView)
        Gv1.ShowGroupPanel = True
        Gv1.ShowRowHeaderColumn = True
        Gv1.AllowAddNewRow = False
        Gv1.AllowDeleteRow = False
        Gv1.EnableFiltering = True
        Gv1.ShowFilteringRow = True
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()

        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).IsVisible = rbtnNEFT.IsChecked
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).BestFit()
        Next
        If rbtnBankWise.IsChecked Then


            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Company Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Cycle Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_Code").IsVisible = False

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = True

            Gv1.Columns("Route_Code").HeaderText = "Route Code"
            Gv1.Columns("Route_Code").IsVisible = False

            Gv1.Columns("ROUTE_NAME").HeaderText = "Route"
            Gv1.Columns("ROUTE_NAME").IsVisible = True

            Gv1.Columns("SupervisorOrRP").HeaderText = "RP Code"
            Gv1.Columns("SupervisorOrRP").IsVisible = False

            Gv1.Columns("SupervisorOrRPName").HeaderText = "RP"
            Gv1.Columns("SupervisorOrRPName").IsVisible = True

            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Code").IsVisible = False

            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Society"
            Gv1.Columns("VLC_Code_VLC_Uploader").IsVisible = True

            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Name").IsVisible = False


            Gv1.Columns("MP_Code").HeaderText = "MP Code"
            Gv1.Columns("MP_Code").IsVisible = False

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "MP Uploader Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Payee_Joint_Name").HeaderText = "MP Name"
            Gv1.Columns("Payee_Joint_Name").IsVisible = True

            'Gv1.Columns("Jan_Aadhar_No_Verified").HeaderText = "Jan Aadhar Verified"
            'Gv1.Columns("Jan_Aadhar_No_Verified").IsVisible = True
            '======================================
            Gv1.Columns("District Code").HeaderText = "District Code"
            Gv1.Columns("District Code").IsVisible = False

            Gv1.Columns("District Name").HeaderText = "District Name"
            Gv1.Columns("District Name").IsVisible = True

            Gv1.Columns("Zone Code").HeaderText = "Zone Code"
            Gv1.Columns("Zone Code").IsVisible = True

            Gv1.Columns("Zone Name").HeaderText = "Zone Name"
            Gv1.Columns("Zone Name").IsVisible = True

            Gv1.Columns("Block Code").HeaderText = "Block Code"
            Gv1.Columns("Block Code").IsVisible = False

            Gv1.Columns("Block Name").HeaderText = "Block Name"
            Gv1.Columns("Block Name").IsVisible = True

            Gv1.Columns("Revenue Village Code").HeaderText = "Revenue Village Code"
            Gv1.Columns("Revenue Village Code").IsVisible = False

            Gv1.Columns("Revenue Village Name").HeaderText = "Revenue Village Name"
            Gv1.Columns("Revenue Village Name").IsVisible = True

            Gv1.Columns("Grampanchayat Code").HeaderText = "Grampanchayat Code"
            Gv1.Columns("Grampanchayat Code").IsVisible = False

            Gv1.Columns("Grampanchayat Name").HeaderText = "Grampanchayat Name"
            Gv1.Columns("Grampanchayat Name").IsVisible = True

            Gv1.Columns("Panchayat Samiti Code").HeaderText = "Panchayat Samiti Code"
            Gv1.Columns("Panchayat Samiti Code").IsVisible = False

            Gv1.Columns("Panchayat Samiti Name").HeaderText = "Panchayat Samiti Name"
            Gv1.Columns("Panchayat Samiti Name").IsVisible = True

            Gv1.Columns("Vidhan Sabha Code").HeaderText = "Vidhan Sabha Code"
            Gv1.Columns("Vidhan Sabha Code").IsVisible = False

            Gv1.Columns("Vidhan Sabha Name").HeaderText = "Vidhan Sabha Name"
            Gv1.Columns("Vidhan Sabha Name").IsVisible = True

            Gv1.Columns("Created Month-Year").HeaderText = "Created Month-Year"
            Gv1.Columns("Created Month-Year").IsVisible = True
            '====================================

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = False

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = True

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "A/C No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            'Gv1.Columns("Bank State").HeaderText = "Bank State"
            'Gv1.Columns("Bank State").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True
            Gv1.Columns("Quantity").IsVisible = True
            Gv1.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"

            Gv1.Columns("THIRD_PARTY_CODE").IsVisible = False
            Gv1.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True

            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item1 As New GridViewSummaryItem("Quantity", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item1)
            Dim item2 As New GridViewSummaryItem("Payable_Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item2)

            Gv1.ShowGroupPanel = False
            Gv1.MasterTemplate.AutoExpandGroups = True
            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        ElseIf rbtnPerNEFT.IsChecked Then
            Gv1.Columns("Created_Date").HeaderText = "Month & Year"
            Gv1.Columns("Created_Date").IsVisible = True
            Gv1.Columns("VLC_Uploader_Code").HeaderText = "DCS Code"
            Gv1.Columns("VLC_Uploader_Code").IsVisible = True
            Gv1.Columns("VLC_Name").HeaderText = "DCS Name"
            Gv1.Columns("VLC_Name").IsVisible = True
            Gv1.Columns("MP_Code").HeaderText = "ERP MP Code"
            Gv1.Columns("MP_Code").IsVisible = True
            Gv1.Columns("MP_Uploader_Code").HeaderText = "MP Code"
            Gv1.Columns("MP_Uploader_Code").IsVisible = True
            Gv1.Columns("MP_NAME").HeaderText = "MP Name"
            Gv1.Columns("MP_NAME").IsVisible = True
            Gv1.Columns("Qty").HeaderText = "Qty in Ltr"
            Gv1.Columns("Qty").IsVisible = True
            Gv1.Columns("Amount").HeaderText = "Amount"
            Gv1.Columns("Amount").IsVisible = True
            Gv1.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            Gv1.Columns("THIRD_PARTY_CODE").IsVisible = False
            Gv1.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True

            Dim summaryRowItem As New GridViewSummaryRowItem()
            'Dim item1 As New GridViewSummaryItem("Qty", "{0:F2}", GridAggregateFunction.Sum)
            'summaryRowItem.Add(item1)
            Dim item2 As New GridViewSummaryItem("Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item2)

            Gv1.ShowGroupPanel = True
            Gv1.MasterTemplate.AutoExpandGroups = True
            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        ElseIf rbtnDetails.IsChecked Then
            Gv1.Columns("S_No").HeaderText = "S.No."
            Gv1.Columns("S_No").IsVisible = True

            Gv1.Columns("DocumentCode").HeaderText = "Document code"
            Gv1.Columns("DocumentCode").IsVisible = True

            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Company Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Cycle Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_Code").IsVisible = False

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = True

            Gv1.Columns("Route_Code").HeaderText = "Route Code"
            Gv1.Columns("Route_Code").IsVisible = True

            Gv1.Columns("ROUTE_NAME").HeaderText = "Route Name"
            Gv1.Columns("ROUTE_NAME").IsVisible = True

            Gv1.Columns("SupervisorOrRP").HeaderText = "RP Code"
            Gv1.Columns("SupervisorOrRP").IsVisible = False

            Gv1.Columns("SupervisorOrRPName").HeaderText = "RP"
            Gv1.Columns("SupervisorOrRPName").IsVisible = True

            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Code").IsVisible = False

            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Society"
            Gv1.Columns("VLC_Code_VLC_Uploader").IsVisible = True

            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Name").IsVisible = False

            Gv1.Columns("MP_Code").HeaderText = "MP Code"
            Gv1.Columns("MP_Code").IsVisible = False

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "MP Uploader Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Payee_Joint_Name").HeaderText = "MP Name"
            Gv1.Columns("Payee_Joint_Name").IsVisible = True

            '======================================
            Gv1.Columns("District Code").HeaderText = "District Code"
            Gv1.Columns("District Code").IsVisible = False

            Gv1.Columns("District Name").HeaderText = "District Name"
            Gv1.Columns("District Name").IsVisible = True

            Gv1.Columns("Zone Code").HeaderText = "Zone Code"
            Gv1.Columns("Zone Code").IsVisible = True

            Gv1.Columns("Zone Name").HeaderText = "Zone Name"
            Gv1.Columns("Zone Name").IsVisible = True

            Gv1.Columns("Block Code").HeaderText = "Block Code"
            Gv1.Columns("Block Code").IsVisible = False

            Gv1.Columns("Block Name").HeaderText = "Block Name"
            Gv1.Columns("Block Name").IsVisible = True

            Gv1.Columns("Revenue Village Code").HeaderText = "Revenue Village Code"
            Gv1.Columns("Revenue Village Code").IsVisible = False

            Gv1.Columns("Revenue Village Name").HeaderText = "Revenue Village Name"
            Gv1.Columns("Revenue Village Name").IsVisible = True

            Gv1.Columns("Grampanchayat Code").HeaderText = "Grampanchayat Code"
            Gv1.Columns("Grampanchayat Code").IsVisible = False

            Gv1.Columns("Grampanchayat Name").HeaderText = "Grampanchayat Name"
            Gv1.Columns("Grampanchayat Name").IsVisible = True

            Gv1.Columns("Panchayat Samiti Code").HeaderText = "Panchayat Samiti Code"
            Gv1.Columns("Panchayat Samiti Code").IsVisible = False

            Gv1.Columns("Panchayat Samiti Name").HeaderText = "Panchayat Samiti Name"
            Gv1.Columns("Panchayat Samiti Name").IsVisible = True

            Gv1.Columns("Vidhan Sabha Code").HeaderText = "Vidhan Sabha Code"
            Gv1.Columns("Vidhan Sabha Code").IsVisible = False

            Gv1.Columns("Vidhan Sabha Name").HeaderText = "Vidhan Sabha Name"
            Gv1.Columns("Vidhan Sabha Name").IsVisible = True

            Gv1.Columns("Created Month-Year").HeaderText = "Created Month-Year"
            Gv1.Columns("Created Month-Year").IsVisible = True
            '====================================

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = False

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = True

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "A/C No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = False

            Gv1.Columns("Original_Payable_Amount").HeaderText = "Original Amount"
            Gv1.Columns("Original_Payable_Amount").IsVisible = False

            Gv1.Columns("Quantity").IsVisible = False

            Gv1.Columns("Created_Entry_Source").HeaderText = "Created Source"
            Gv1.Columns("Created_Entry_Source").IsVisible = True

            Gv1.Columns("Created_By").HeaderText = "Created By"
            Gv1.Columns("Created_By").IsVisible = True

            Gv1.Columns("Created_Date").HeaderText = "Created On"
            Gv1.Columns("Created_Date").IsVisible = True

            Gv1.Columns("Modified_Entry_Source").HeaderText = "Modify Source"
            Gv1.Columns("Modified_Entry_Source").IsVisible = True

            Gv1.Columns("Modified_By").HeaderText = "Modify By"
            Gv1.Columns("Modified_By").IsVisible = True

            Gv1.Columns("Modified_Date").HeaderText = "Modify On"
            Gv1.Columns("Modified_Date").IsVisible = True

            Gv1.Columns("Reco_Status").HeaderText = "Reco Status"
            Gv1.Columns("Reco_Status").IsVisible = True

            Gv1.Columns("Reco_Document_Code").HeaderText = "Reco No"
            Gv1.Columns("Reco_Document_Code").IsVisible = True

            '=======================================

            Gv1.Columns("Jan_Aadhar_No").HeaderText = "Jan Aadhar No"
            Gv1.Columns("Jan_Aadhar_No").IsVisible = True

            Gv1.Columns("Jan_Aadhar_No_Verified").HeaderText = "Jan Aadhar Verified"
            Gv1.Columns("Jan_Aadhar_No_Verified").IsVisible = True

            Gv1.Columns("Jan_Aadhar_Account_No").HeaderText = "Jan Aadhar Account"
            Gv1.Columns("Jan_Aadhar_Account_No").IsVisible = True

            Gv1.Columns("JA_bankBranch").HeaderText = "JA Bank Branch"
            Gv1.Columns("JA_bankBranch").IsVisible = False

            Gv1.Columns("JA_bankName").HeaderText = "JA Bank Name"
            Gv1.Columns("JA_bankName").IsVisible = False

            Gv1.Columns("JA_ifsc").HeaderText = "JA IFSC"
            Gv1.Columns("JA_ifsc").IsVisible = False

            Gv1.Columns("JA_caste").HeaderText = "JA Caste"
            Gv1.Columns("JA_caste").IsVisible = False

            Gv1.Columns("JA_categoryDescEng").HeaderText = "JA Category Desc Eng"
            Gv1.Columns("JA_categoryDescEng").IsVisible = False

            Gv1.Columns("JA_disability").HeaderText = "JA Disability"
            Gv1.Columns("JA_disability").IsVisible = False

            Gv1.Columns("JA_disabilityPercentage").HeaderText = "JA Disability Percentage"
            Gv1.Columns("JA_disabilityPercentage").IsVisible = False

            Gv1.Columns("JA_disabilityType").HeaderText = "JA Disability Type"
            Gv1.Columns("JA_disabilityType").IsVisible = False

            Gv1.Columns("JA_dlNo").HeaderText = "JA DL No"
            Gv1.Columns("JA_dlNo").IsVisible = False

            Gv1.Columns("JA_dob").HeaderText = "JA DOB"
            Gv1.Columns("JA_dob").IsVisible = False

            Gv1.Columns("JA_eid").HeaderText = "JA EID"
            Gv1.Columns("JA_eid").IsVisible = False

            Gv1.Columns("JA_email").HeaderText = "JA Email"
            Gv1.Columns("JA_email").IsVisible = False

            Gv1.Columns("JA_enrId").HeaderText = "JA Enr ID"
            Gv1.Columns("JA_enrId").IsVisible = False

            Gv1.Columns("JA_fnameEng").HeaderText = "JA FName Eng"
            Gv1.Columns("JA_fnameEng").IsVisible = False

            Gv1.Columns("JA_fnameHnd").HeaderText = "JA Fname Hnd"
            Gv1.Columns("JA_fnameHnd").IsVisible = False

            Gv1.Columns("JA_gender").HeaderText = "JA Gender"
            Gv1.Columns("JA_gender").IsVisible = False

            Gv1.Columns("JA_income").HeaderText = "JA Income"
            Gv1.Columns("JA_income").IsVisible = False

            Gv1.Columns("JA_isorphan").HeaderText = "JA Isorphan"
            Gv1.Columns("JA_isorphan").IsVisible = False

            Gv1.Columns("JA_isStateGovtEmp").HeaderText = "JA Is State Govt Emp"
            Gv1.Columns("JA_isStateGovtEmp").IsVisible = False

            Gv1.Columns("JA_jan_mid").HeaderText = "JA Jan Mid"
            Gv1.Columns("JA_jan_mid").IsVisible = False

            Gv1.Columns("JA_janaadhaarId").HeaderText = "JA Janaadhaar Id"
            Gv1.Columns("JA_janaadhaarId").IsVisible = False

            Gv1.Columns("JA_maritalStatus").HeaderText = "JA Marital Status"
            Gv1.Columns("JA_maritalStatus").IsVisible = False

            Gv1.Columns("JA_micr").HeaderText = "JA MICR"
            Gv1.Columns("JA_micr").IsVisible = False

            Gv1.Columns("JA_mobile").HeaderText = "JA Mobile"
            Gv1.Columns("JA_mobile").IsVisible = False

            Gv1.Columns("JA_nameEng").HeaderText = "JA Name Eng"
            Gv1.Columns("JA_nameEng").IsVisible = False

            Gv1.Columns("JA_nameHnd").HeaderText = "JA Name Hnd"
            Gv1.Columns("JA_nameHnd").IsVisible = False

            Gv1.Columns("JA_occupation").HeaderText = "JA Occupation"
            Gv1.Columns("JA_occupation").IsVisible = False

            Gv1.Columns("JA_relationTyp").HeaderText = "JA Relation Type"
            Gv1.Columns("JA_relationTyp").IsVisible = False

            Gv1.Columns("JA_panNo").HeaderText = "JA Pan No"
            Gv1.Columns("JA_panNo").IsVisible = False

            Gv1.Columns("JA_passport").HeaderText = "JA Passport"
            Gv1.Columns("JA_passport").IsVisible = False

            Gv1.Columns("JA_qualification").HeaderText = "JA Qualification"
            Gv1.Columns("JA_qualification").IsVisible = False

            Gv1.Columns("JA_snameEng").HeaderText = "JA SName Eng"
            Gv1.Columns("JA_snameEng").IsVisible = False

            Gv1.Columns("JA_snameHnd").HeaderText = "JA SName Hnd"
            Gv1.Columns("JA_snameHnd").IsVisible = False

            Gv1.Columns("JA_voterId").HeaderText = "JA Voter ID"
            Gv1.Columns("JA_voterId").IsVisible = False

            Gv1.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            Gv1.Columns("THIRD_PARTY_CODE").IsVisible = False
            Gv1.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True


            Gv1.Columns("Mineral_Mixture_Qty").HeaderText = "Mineral Mixture Qty"
            Gv1.Columns("Mineral_Mixture_Amount").HeaderText = "Mineral Mixture Amt"
            Gv1.Columns("Pashu_Aahar_Qty").HeaderText = "Pashu Aahar Qty"
            Gv1.Columns("Pashu_Aahar_Amount").HeaderText = "Pashu Aahar Amt"
            Gv1.Columns("Rahat_Kampekat_Feed_Qty").HeaderText = "Rahat Kampekat Feed Qty"
            Gv1.Columns("Rahat_Kampekat_Feed_Amount").HeaderText = "Rahat Kampekat Feed Amt"
            Gv1.Columns("Sailej_Qty").HeaderText = "Sailej Qty"
            Gv1.Columns("Sailej_Amount").HeaderText = "Sailej Amount"
            Gv1.Columns("Total_Amount").HeaderText = "Total Amount"
            Gv1.Columns("Total_Amount").IsVisible = False

            If clsCommon.CompairString(ddlType.SelectedValue, "Mineral Mixture") = CompairStringResult.Equal Then
                Gv1.Columns("Mineral_Mixture_Qty").IsVisible = True
                Gv1.Columns("Mineral_Mixture_Amount").IsVisible = True
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Pashu Aahar") = CompairStringResult.Equal Then
                Gv1.Columns("Pashu_Aahar_Qty").IsVisible = True
                Gv1.Columns("Pashu_Aahar_Amount").IsVisible = True
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Rahat Kampekat Feed") = CompairStringResult.Equal Then
                Gv1.Columns("Rahat_Kampekat_Feed_Qty").IsVisible = True
                Gv1.Columns("Rahat_Kampekat_Feed_Amount").IsVisible = True
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Sailej") = CompairStringResult.Equal Then
                Gv1.Columns("Sailej_Qty").IsVisible = True
                Gv1.Columns("Sailej_Amount").IsVisible = True
            Else
                Gv1.Columns("Mineral_Mixture_Qty").IsVisible = True
                Gv1.Columns("Mineral_Mixture_Amount").IsVisible = True
                Gv1.Columns("Pashu_Aahar_Qty").IsVisible = True
                Gv1.Columns("Pashu_Aahar_Amount").IsVisible = True
                Gv1.Columns("Rahat_Kampekat_Feed_Qty").IsVisible = True
                Gv1.Columns("Rahat_Kampekat_Feed_Amount").IsVisible = True
                Gv1.Columns("Sailej_Qty").IsVisible = True
                Gv1.Columns("Sailej_Amount").IsVisible = True
                Gv1.Columns("Total_Amount").IsVisible = True
                Gv1.Columns("Payable_Amount").IsVisible = True
                Gv1.Columns("Original_Payable_Amount").IsVisible = False
                Gv1.Columns("Quantity").IsVisible = True
            End If
            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item1 As New GridViewSummaryItem("Quantity", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item1)
            Dim item2 As New GridViewSummaryItem("Payable_Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item2)
            Dim item3 As New GridViewSummaryItem("Original_Payable_Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item3)

            Gv1.ShowGroupPanel = False
            Gv1.MasterTemplate.AutoExpandGroups = True
            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        ElseIf rbtnMCCWise.IsChecked Then

            If rbtnDetails.IsChecked Then
                Gv1.Columns("S_No").HeaderText = "S.No."
                Gv1.Columns("S_No").IsVisible = True
                Gv1.Columns("Jan_Aadhar_No").IsVisible = False
            End If

            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Company Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Cycle Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_Code").IsVisible = False

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = True


            Gv1.Columns("Route_Code").HeaderText = "Route Code"
            Gv1.Columns("Route_Code").IsVisible = False

            Gv1.Columns("ROUTE_NAME").HeaderText = "Route"
            Gv1.Columns("ROUTE_NAME").IsVisible = True

            Gv1.Columns("SupervisorOrRP").HeaderText = "RP Code"
            Gv1.Columns("SupervisorOrRP").IsVisible = False

            Gv1.Columns("SupervisorOrRPName").HeaderText = "RP"
            Gv1.Columns("SupervisorOrRPName").IsVisible = True

            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Code").IsVisible = False

            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Society"
            Gv1.Columns("VLC_Code_VLC_Uploader").IsVisible = True

            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Name").IsVisible = False

            Gv1.Columns("MP_Code").HeaderText = "MP Code"
            Gv1.Columns("MP_Code").IsVisible = False

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "MP Uploader Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Payee_Joint_Name").HeaderText = "MP Name"
            Gv1.Columns("Payee_Joint_Name").IsVisible = True

            '======================================
            Gv1.Columns("District Code").HeaderText = "District Code"
            Gv1.Columns("District Code").IsVisible = False

            Gv1.Columns("District Name").HeaderText = "District Name"
            Gv1.Columns("District Name").IsVisible = True

            Gv1.Columns("Zone Code").HeaderText = "Zone Code"
            Gv1.Columns("Zone Code").IsVisible = True

            Gv1.Columns("Zone Name").HeaderText = "Zone Name"
            Gv1.Columns("Zone Name").IsVisible = True

            Gv1.Columns("Block Code").HeaderText = "Block Code"
            Gv1.Columns("Block Code").IsVisible = False

            Gv1.Columns("Block Name").HeaderText = "Block Name"
            Gv1.Columns("Block Name").IsVisible = True

            Gv1.Columns("Revenue Village Code").HeaderText = "Revenue Village Code"
            Gv1.Columns("Revenue Village Code").IsVisible = False

            Gv1.Columns("Revenue Village Name").HeaderText = "Revenue Village Name"
            Gv1.Columns("Revenue Village Name").IsVisible = True

            Gv1.Columns("Grampanchayat Code").HeaderText = "Grampanchayat Code"
            Gv1.Columns("Grampanchayat Code").IsVisible = False

            Gv1.Columns("Grampanchayat Name").HeaderText = "Grampanchayat Name"
            Gv1.Columns("Grampanchayat Name").IsVisible = True

            Gv1.Columns("Panchayat Samiti Code").HeaderText = "Panchayat Samiti Code"
            Gv1.Columns("Panchayat Samiti Code").IsVisible = False

            Gv1.Columns("Panchayat Samiti Name").HeaderText = "Panchayat Samiti Name"
            Gv1.Columns("Panchayat Samiti Name").IsVisible = True

            Gv1.Columns("Vidhan Sabha Code").HeaderText = "Vidhan Sabha Code"
            Gv1.Columns("Vidhan Sabha Code").IsVisible = False

            Gv1.Columns("Vidhan Sabha Name").HeaderText = "Vidhan Sabha Name"
            Gv1.Columns("Vidhan Sabha Name").IsVisible = True

            Gv1.Columns("Created Month-Year").HeaderText = "Created Month-Year"
            Gv1.Columns("Created Month-Year").IsVisible = True
            '====================================

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = False

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = True

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "A/C No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = False

            Gv1.Columns("Original_Payable_Amount").HeaderText = "Original Amount"
            Gv1.Columns("Original_Payable_Amount").IsVisible = False

            Gv1.Columns("Quantity").IsVisible = False


            Gv1.Columns("Created_Entry_Source").HeaderText = "Created Source"
            Gv1.Columns("Created_Entry_Source").IsVisible = True

            Gv1.Columns("Created_By").HeaderText = "Created By"
            Gv1.Columns("Created_By").IsVisible = True

            Gv1.Columns("Created_Date").HeaderText = "Created On"
            Gv1.Columns("Created_Date").IsVisible = True

            Gv1.Columns("Modified_Entry_Source").HeaderText = "Modify Source"
            Gv1.Columns("Modified_Entry_Source").IsVisible = True

            Gv1.Columns("Modified_By").HeaderText = "Modify By"
            Gv1.Columns("Modified_By").IsVisible = True

            Gv1.Columns("Modified_Date").HeaderText = "Modify On"
            Gv1.Columns("Modified_Date").IsVisible = True

            Gv1.Columns("Reco_Status").HeaderText = "Reco Status"
            Gv1.Columns("Reco_Status").IsVisible = True

            Gv1.Columns("Reco_Document_Code").HeaderText = "Reco No"
            Gv1.Columns("Reco_Document_Code").IsVisible = True

            Gv1.Columns("Jan_Aadhar_No").HeaderText = "Jan Aadhar No"
            Gv1.Columns("Jan_Aadhar_No").IsVisible = True

            Gv1.Columns("Mineral_Mixture_Qty").HeaderText = "Mineral Mixture Qty"
            Gv1.Columns("Mineral_Mixture_Amount").HeaderText = "Mineral Mixture Amt"
            Gv1.Columns("Pashu_Aahar_Qty").HeaderText = "Pashu Aahar Qty"
            Gv1.Columns("Pashu_Aahar_Amount").HeaderText = "Pashu Aahar Amt"
            Gv1.Columns("Rahat_Kampekat_Feed_Qty").HeaderText = "Rahat Kampekat Feed Qty"
            Gv1.Columns("Rahat_Kampekat_Feed_Amount").HeaderText = "Rahat Kampekat Feed Amt"
            Gv1.Columns("Sailej_Qty").HeaderText = "Sailej Qty"
            Gv1.Columns("Sailej_Amount").HeaderText = "Sailej Amount"
            Gv1.Columns("Total_Amount").HeaderText = "Total Amount"

            Gv1.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            Gv1.Columns("THIRD_PARTY_CODE").IsVisible = False
            Gv1.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True


            If clsCommon.CompairString(ddlType.SelectedValue, "Mineral Mixture") = CompairStringResult.Equal Then
                Gv1.Columns("Mineral_Mixture_Qty").IsVisible = True
                Gv1.Columns("Mineral_Mixture_Amount").IsVisible = True
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Pashu Aahar") = CompairStringResult.Equal Then
                Gv1.Columns("Pashu_Aahar_Qty").IsVisible = True
                Gv1.Columns("Pashu_Aahar_Amount").IsVisible = True
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Rahat Kampekat Feed") = CompairStringResult.Equal Then
                Gv1.Columns("Rahat_Kampekat_Feed_Qty").IsVisible = True
                Gv1.Columns("Rahat_Kampekat_Feed_Amount").IsVisible = True
            ElseIf clsCommon.CompairString(ddlType.SelectedValue, "Sailej") = CompairStringResult.Equal Then
                Gv1.Columns("Sailej_Qty").IsVisible = True
                Gv1.Columns("Sailej_Amount").IsVisible = True
            Else
                Gv1.Columns("Mineral_Mixture_Qty").IsVisible = True
                Gv1.Columns("Mineral_Mixture_Amount").IsVisible = True
                Gv1.Columns("Pashu_Aahar_Qty").IsVisible = True
                Gv1.Columns("Pashu_Aahar_Amount").IsVisible = True
                Gv1.Columns("Rahat_Kampekat_Feed_Qty").IsVisible = True
                Gv1.Columns("Rahat_Kampekat_Feed_Amount").IsVisible = True
                Gv1.Columns("Sailej_Qty").IsVisible = True
                Gv1.Columns("Sailej_Amount").IsVisible = True
                Gv1.Columns("Total_Amount").IsVisible = True
                Gv1.Columns("Payable_Amount").IsVisible = True
                Gv1.Columns("Quantity").IsVisible = True
            End If
            Dim summaryRowItem As New GridViewSummaryRowItem()
            Dim item1 As New GridViewSummaryItem("Quantity", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item1)
            Dim item2 As New GridViewSummaryItem("Payable_Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item2)
            Dim item3 As New GridViewSummaryItem("Original_Payable_Amount", "{0:F2}", GridAggregateFunction.Sum)
            summaryRowItem.Add(item3)

            Gv1.ShowGroupPanel = False
            Gv1.MasterTemplate.AutoExpandGroups = True
            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItem)
            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        ElseIf rbtnNEFT.IsChecked Then
            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_Code").IsVisible = False

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = True

            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Code").IsVisible = False

            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Society"
            Gv1.Columns("VLC_Code_VLC_Uploader").IsVisible = True

            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Name").IsVisible = False

            Gv1.Columns("MP_Code").HeaderText = "MP Code"
            Gv1.Columns("MP_Code").IsVisible = False

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "MP Uploader Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Payee_Joint_Name").HeaderText = "MP Name"
            Gv1.Columns("Payee_Joint_Name").IsVisible = True

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = False

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = True

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "A/C No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True

            Gv1.Columns("Quantity").IsVisible = True

            Gv1.Columns("Status").HeaderText = "Status"
            Gv1.Columns("Status").IsVisible = True

        ElseIf rbtnDuplicateACNo.IsChecked Then
            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "Account No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("VLC_Code").HeaderText = "VLC Code"
            Gv1.Columns("VLC_Code").IsVisible = True

            Gv1.Columns("VLC_Name").HeaderText = "VLC Name"
            Gv1.Columns("VLC_Name").IsVisible = True

            Gv1.Columns("MP_Code").HeaderText = "MP Code"
            Gv1.Columns("MP_Code").IsVisible = True

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "MP Uploader Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Payee_Joint_Name").HeaderText = "MP Name"
            Gv1.Columns("Payee_Joint_Name").IsVisible = True

            'Gv1.Columns("Repeated").HeaderText = "Repeated"
            'Gv1.Columns("Repeated").IsVisible = True

            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Company Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Cycle Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("MCC_Code").HeaderText = "MCC_Code"
            Gv1.Columns("MCC_Code").IsVisible = True

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = True

            Gv1.Columns("Route_Code").HeaderText = "Route Code"
            Gv1.Columns("Route_Code").IsVisible = True

            Gv1.Columns("ROUTE_NAME").HeaderText = "Route"
            Gv1.Columns("ROUTE_NAME").IsVisible = True

            Gv1.Columns("SupervisorOrRP").HeaderText = "RP Code"
            Gv1.Columns("SupervisorOrRP").IsVisible = True

            Gv1.Columns("SupervisorOrRPName").HeaderText = "RP"
            Gv1.Columns("SupervisorOrRPName").IsVisible = True

            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "Society"
            Gv1.Columns("VLC_Code_VLC_Uploader").IsVisible = True

            '======================================
            Gv1.Columns("District Code").HeaderText = "District Code"
            Gv1.Columns("District Code").IsVisible = True

            Gv1.Columns("District Name").HeaderText = "District Name"
            Gv1.Columns("District Name").IsVisible = True

            Gv1.Columns("Zone Code").HeaderText = "Zone Code"
            Gv1.Columns("Zone Code").IsVisible = True

            Gv1.Columns("Zone Name").HeaderText = "Zone Name"
            Gv1.Columns("Zone Name").IsVisible = True

            Gv1.Columns("Block Code").HeaderText = "Block Code"
            Gv1.Columns("Block Code").IsVisible = True

            Gv1.Columns("Block Name").HeaderText = "Block Name"
            Gv1.Columns("Block Name").IsVisible = True

            Gv1.Columns("Revenue Village Code").HeaderText = "Revenue Village Code"
            Gv1.Columns("Revenue Village Code").IsVisible = True

            Gv1.Columns("Revenue Village Name").HeaderText = "Revenue Village Name"
            Gv1.Columns("Revenue Village Name").IsVisible = True

            Gv1.Columns("Grampanchayat Code").HeaderText = "Grampanchayat Code"
            Gv1.Columns("Grampanchayat Code").IsVisible = True

            Gv1.Columns("Grampanchayat Name").HeaderText = "Grampanchayat Name"
            Gv1.Columns("Grampanchayat Name").IsVisible = True

            Gv1.Columns("Panchayat Samiti Code").HeaderText = "Panchayat Samiti Code"
            Gv1.Columns("Panchayat Samiti Code").IsVisible = True

            Gv1.Columns("Panchayat Samiti Name").HeaderText = "Panchayat Samiti Name"
            Gv1.Columns("Panchayat Samiti Name").IsVisible = True

            Gv1.Columns("Vidhan Sabha Code").HeaderText = "Vidhan Sabha Code"
            Gv1.Columns("Vidhan Sabha Code").IsVisible = True

            Gv1.Columns("Vidhan Sabha Name").HeaderText = "Vidhan Sabha Name"
            Gv1.Columns("Vidhan Sabha Name").IsVisible = True

            Gv1.Columns("Created Month-Year").HeaderText = "Created Month-Year"
            Gv1.Columns("Created Month-Year").IsVisible = True
            '====================================

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = True

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = True

            'Gv1.Columns("Payee_Joint_Account_No").HeaderText = "A/C No"
            'Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True

            Gv1.Columns("Qty").HeaderText = "Qty"
            Gv1.Columns("Qty").IsVisible = True

            Gv1.Columns("Created_Entry_Source").HeaderText = "Created Source"
            Gv1.Columns("Created_Entry_Source").IsVisible = True

            Gv1.Columns("Created_By").HeaderText = "Created By"
            Gv1.Columns("Created_By").IsVisible = True

            Gv1.Columns("Created_Date").HeaderText = "Created On"
            Gv1.Columns("Created_Date").IsVisible = True

            Gv1.Columns("Modified_Entry_Source").HeaderText = "Modify Source"
            Gv1.Columns("Modified_Entry_Source").IsVisible = True

            Gv1.Columns("Modified_By").HeaderText = "Modify By"
            Gv1.Columns("Modified_By").IsVisible = True

            Gv1.Columns("Modified_Date").HeaderText = "Modify On"
            Gv1.Columns("Modified_Date").IsVisible = True

            Gv1.Columns("Reco_Status").HeaderText = "Reco Status"
            Gv1.Columns("Reco_Status").IsVisible = True

            Gv1.Columns("Reco_Document_Code").HeaderText = "Reco No"
            Gv1.Columns("Reco_Document_Code").IsVisible = True

            Gv1.Columns("Jan_Aadhar_No").HeaderText = "Jan Aadhar No"
            Gv1.Columns("Jan_Aadhar_No").IsVisible = True

            Gv1.Columns("Mineral_Mixture_Qty").HeaderText = "Mineral Mixture Qty"
            Gv1.Columns("Mineral_Mixture_Amount").HeaderText = "Mineral Mixture Amt"
            Gv1.Columns("Pashu_Aahar_Qty").HeaderText = "Pashu Aahar Qty"
            Gv1.Columns("Pashu_Aahar_Amount").HeaderText = "Pashu Aahar Amt"
            Gv1.Columns("Rahat_Kampekat_Feed_Qty").HeaderText = "Rahat Kampekat Feed Qty"
            Gv1.Columns("Rahat_Kampekat_Feed_Amount").HeaderText = "Rahat Kampekat Feed Amt"
            Gv1.Columns("Sailej_Qty").HeaderText = "Sailej Qty"
            Gv1.Columns("Sailej_Amount").HeaderText = "Sailej Amount"
            ' Gv1.Columns("Total_Amount").HeaderText = "Total Amount"

        ElseIf rbtnDuplicateJanAdharNo.IsChecked Then
            Gv1.Columns("Jan_Aadhar_No").HeaderText = "Janadhar No"
            Gv1.Columns("Jan_Aadhar_No").IsVisible = True

            Gv1.Columns("Repeated").HeaderText = "Repeated"
            Gv1.Columns("Repeated").IsVisible = True

        ElseIf rbtnFarmerBankWiseDetail.IsChecked Then
            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Comp Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Date Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_Code").IsVisible = False

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = False

            Gv1.Columns("VLC_Code").HeaderText = "DCS"
            Gv1.Columns("VLC_Code").IsVisible = False

            Gv1.Columns("VLC_Code_VLC_Uploader").HeaderText = "DCS"
            Gv1.Columns("VLC_Code_VLC_Uploader").IsVisible = True

            Gv1.Columns("VLC_Name").HeaderText = "DCS Name"
            Gv1.Columns("VLC_Name").IsVisible = False

            Gv1.Columns("MP_Code").HeaderText = "Farmer Code"
            Gv1.Columns("MP_Code").IsVisible = False

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "Farmer Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Bank_Code").HeaderText = "Farmer Code"
            Gv1.Columns("Bank_Code").IsVisible = False

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "MM Name"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("SNO").HeaderText = "SNo"
            Gv1.Columns("SNO").IsVisible = True

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Quantity").HeaderText = "Quantity"
            Gv1.Columns("Quantity").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True

            Gv1.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            Gv1.Columns("THIRD_PARTY_CODE").IsVisible = False
            Gv1.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True

        ElseIf rbtnFarmerBankWiseSummary.IsChecked Then
            Gv1.Columns("Comp_Name").HeaderText = "VLC Code"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Comp Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("MCC_Code").HeaderText = "MCC Code"
            Gv1.Columns("MCC_Code").IsVisible = False

            Gv1.Columns("MCC_Name").HeaderText = "BMC"
            Gv1.Columns("MCC_Name").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Date Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("SNO").HeaderText = "SNo"
            Gv1.Columns("SNO").IsVisible = True

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = False


            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True

            Gv1.Columns("Quantity").HeaderText = "Quantity"
            Gv1.Columns("Quantity").IsVisible = True
            Gv1.Columns("THIRD_PARTY_CODE").HeaderText = "Third Party Code"
            Gv1.Columns("THIRD_PARTY_CODE").IsVisible = False
            Gv1.Columns("THIRD_PARTY_CODE").VisibleInColumnChooser = True

        End If
        Gv1.AutoSizeRows = False
        Gv1.BestFitColumns()
        Gv1.MasterTemplate.AutoExpandGroups = True
    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        EnableDisaableControls(True)
    End Sub
    Sub EnableDisaableControls(ByVal flag As Boolean)
        txtMCC.Enabled = flag
        'txtFromDate.Enabled = flag
        'txtToDate.Enabled = flag
        ddlType.Enabled = flag
        GroupBox1.Enabled = flag
        txtBlock.Enabled = flag
        If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
            txtZone.Enabled = False
        End If
        txtRevenueVillage.Enabled = flag
        txtGrampanchayat.Enabled = flag
        txtPanchayatSamiti.Enabled = flag
        txtVidhanSabha.Enabled = flag
        txtDistrict.Enabled = flag
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    Private Sub ExportGrid(ByVal exporter As EnumExportTo)
        Try
            If Gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If
            Dim strHeading As String = "MP Incetive Entry Register"

            Dim arrHeader As List(Of String) = New List(Of String)()
            If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                arrHeader.Add("Mcc : " + clsCommon.GetMulcallStringWithComma(txtMCC.arrValueMember))
            End If
            arrHeader.Add("Date Range from : " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy"))

            If rbtnMCCWise.IsChecked = True Then
                If Not clsCommon.CompairString(ddlType.SelectedValue, "All") = CompairStringResult.Equal Then
                    arrHeader.Add("Type : " + clsCommon.myCstr(ddlType.SelectedValue))
                End If
            End If

            If exporter = EnumExportTo.Excel Then
                'Dim TempPath As String = ""
                'Dim sfd As SaveFileDialog = New SaveFileDialog()
                'sfd.FileName = strHeading
                'If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                '    TempPath = sfd.FileName
                'Else
                '    Return
                'End If

                'Dim dttemp As New DataTable
                'Dim gv As New UserControls.MyRadGridView
                'Dim totqty As Double = 0
                'Me.Controls.Add(gv)

                'For ii As Integer = 0 To arr.Count - 1
                '    gv.DataSource = Nothing
                '    dttemp = Nothing
                '    Dim Temprows As DataRow = arr.Item(arr.Keys(ii))
                '    Dim KeyCol As String = Temprows.Item(strColumnForTotal)
                '    Dim rows As DataRow() = dt.Select("" + strColumnForTotal + "='" + KeyCol + "'")
                '    If rows Is Nothing OrElse rows.Length > 0 Then
                '        dttemp = rows.CopyToDataTable()
                '        gv.DataSource = dttemp
                '        SetGridFormat(gv)
                '        ReStoreGridLayout()
                '        clsCommon.MyExportToExcelGrid(strHeading, gv, arrHeader, Me.Text + KeyCol, False, "", TempPath)
                '    End If
                'Next

                'Me.Controls.Remove(gv)
                transportSql.QuickExportToExcel(Gv1, "", Me.Text, , arrHeader)
            Else
                clsCommon.MyExportToPDF(strHeading, Gv1, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub rmiExcel_Click(sender As Object, e As EventArgs) Handles rmiExcel.Click
        ExportGrid(EnumExportTo.Excel)
    End Sub
    Private Sub rmiPDF_Click(sender As Object, e As EventArgs) Handles rmiPDF.Click
        ExportGrid(EnumExportTo.PDF)
    End Sub
    Private Sub RmiExcelGrid_Click(sender As Object, e As EventArgs) Handles rmiExcelGrid.Click
        Try
            If Gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If

            Dim strHeading As String = clsCommon.myCstr("MP Incetive Entry Register")
            If rbtnNEFT.IsChecked Then
                strHeading = objCommonVar.CurrentCompanyName
            End If
            Dim arrHeader As List(Of String) = New List(Of String)()
            If txtMCC.arrValueMember IsNot Nothing AndAlso txtMCC.arrValueMember.Count > 0 Then
                arrHeader.Add("Mcc : " + clsCommon.GetMulcallStringWithComma(txtMCC.arrValueMember))
            End If

            arrHeader.Add("Date Range from : " + clsCommon.GetPrintDate(txtFromDate.Value, "dd/MM/yyyy") + " To " + clsCommon.GetPrintDate(txtToDate.Value, "dd/MM/yyyy"))
            If rbtnMCCWise.IsChecked = True Then
                If Not clsCommon.CompairString(ddlType.SelectedValue, "All") = CompairStringResult.Equal Then
                    arrHeader.Add("Type : " + clsCommon.myCstr(ddlType.SelectedValue))
                End If
            End If

            clsCommon.MyExportToExcelGrid(strHeading, Gv1, arrHeader, Me.Text, True)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
        End Try
    End Sub
    Private Sub TxtMultiSelectFinder1__My_Click(sender As Object, e As EventArgs) Handles txtMCC._My_Click
        Dim qry As String = "select MCC_Code as [MCC Code],MCC_NAME as [MCC Name],TSPL_MCC_MASTER.plant_code as [Plant Code],tspl_location_master.location_desc as [Plant Name] from TSPL_MCC_MASTER left join tspl_location_master on tspl_location_master.location_code=TSPL_MCC_MASTER.plant_code"
        txtMCC.arrValueMember = clsCommon.ShowMultipleSelectForm("MBPSMCC4", qry, "MCC Code", "MCC Name", txtMCC.arrValueMember, txtMCC.arrDispalyMember)
    End Sub

    'Private Sub txtFromDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtFromDate.Validating
    '    'SetToDate()
    'End Sub
    'Private Sub txtFromDate_Leave(sender As Object, e As EventArgs) Handles txtFromDate.Leave
    '    'SetToDate()
    'End Sub

    'Function SetToDate() As Boolean
    '    Try
    '        Dim PaymentCycleType As String = ""
    '        Dim PaymentCycleValue As Integer = 0
    '        'If txtMCC.arrValueMember Is Nothing OrElse txtMCC.arrValueMember.Count <= 0 Then
    '        '    Throw New Exception("Please select the MCC first")
    '        'End If
    '        Dim dt As DataTable = clsDBFuncationality.GetDataTable(" select Payment_Cycle,PC_TYPE,PC_VALUE from ( select TSPL_MCC_MASTER.Payment_Cycle,TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE  from TSPL_MCC_MASTER left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle   where TSPL_MCC_MASTER.MCC_Code  in (select top 1 mcc_code from tspl_mcc_master order by MCC_Code) ) xx group by Payment_Cycle,PC_TYPE,PC_VALUE")
    '        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
    '            Throw New Exception("No Payment Cycle found on current MCC/Location")
    '        End If
    '        If dt.Rows.Count > 1 Then
    '            Throw New Exception("Selected MCC's Payment Cycle Should be Same")
    '        End If
    '        If SettMPIncentiveEntryApplyMonthly Then
    '            txtFromDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1)
    '            txtToDate.Value = txtFromDate.Value.AddMonths(1).AddDays(-1)
    '        Else
    '            PaymentCycleType = clsCommon.myCstr(dt.Rows(0)("PC_TYPE"))
    '            PaymentCycleValue = clsCommon.myCdbl(dt.Rows(0)("PC_VALUE"))
    '            Dim dtCurr As DateTime = clsCommon.GETSERVERDATE()
    '            If clsCommon.CompairString(PaymentCycleType, "Day") = CompairStringResult.Equal Then
    '                If txtFromDate.Value.Day Mod PaymentCycleValue <> 1 And (Not PaymentCycleValue = 1) Then
    '                    txtFromDate.Value = New Date(dtCurr.Year, dtCurr.Month, 1)
    '                    txtToDate.Value = txtFromDate.Value
    '                    Throw New Exception("Date can only be first day of month or at interval of " & PaymentCycleValue & " Day, Because MCC has payment Cycle of " & PaymentCycleValue & " Day ")
    '                End If
    '                txtToDate.Value = txtFromDate.Value.AddDays(PaymentCycleValue - 1)
    '                'txtToDate.Value = clsCommon.GETSERVERDATE()

    '                If txtFromDate.Value.Month <> txtToDate.Value.Month Then
    '                    txtToDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
    '                End If
    '                Dim dtNxtPay As DateTime = txtToDate.Value.AddDays(Math.Ceiling(PaymentCycleValue / 2.0))
    '                If txtFromDate.Value.Month <> dtNxtPay.Month Then
    '                    txtToDate.Value = New Date(txtFromDate.Value.Year, txtFromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
    '                End If
    '            ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
    '                If clsCommon.myCdbl(clsCommon.GetPrintDate(txtFromDate.Value, "dd")) <> 1 Then

    '                    txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
    '                    txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
    '                    Throw New Exception("Date can only be first day of month, Because MCC has payment Cycle of Month Type")
    '                End If
    '                txtToDate.Value = DateAdd(DateInterval.Month, PaymentCycleValue, txtFromDate.Value)
    '            ElseIf clsCommon.CompairString(PaymentCycleType, "Year") = CompairStringResult.Equal Then
    '                If clsCommon.myCdbl(clsCommon.GetPrintDate(txtFromDate.Value, "dd")) <> 1 Then
    '                    txtFromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
    '                    txtToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
    '                    Throw New Exception("Date can only be first day of month, Because MCC has payment Cycle of Year Type")
    '                End If
    '                txtToDate.Value = DateAdd(DateInterval.Year, PaymentCycleValue, txtFromDate.Value)
    '            ElseIf clsCommon.CompairString(PaymentCycleType, "Week") = CompairStringResult.Equal Then
    '                Dim today As Date = txtFromDate.Value
    '                Dim dayDiff As Integer = today.DayOfWeek - IIf(PaymentCycleValue = 1, DayOfWeek.Sunday, IIf(PaymentCycleValue = 2, DayOfWeek.Monday, IIf(PaymentCycleValue = 3, DayOfWeek.Tuesday, IIf(PaymentCycleValue = 4, DayOfWeek.Wednesday, IIf(PaymentCycleValue = 5, DayOfWeek.Thursday, IIf(PaymentCycleValue = 6, DayOfWeek.Friday, DayOfWeek.Saturday))))))
    '                txtFromDate.Value = today.AddDays(-dayDiff)
    '                txtToDate.Value = txtFromDate.Value.AddDays(6)
    '            End If
    '        End If

    '    Catch ex As Exception
    '        clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
    '        Return False
    '    End Try
    '    Return True
    'End Function

    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
        Try
            If clsCommon.myLen(PageSetupReport_ID) > 0 Then
                Gv1.MasterTemplate.FilterDescriptors.Clear()
                Dim obj As New clsGridLayout()
                obj.ReportID = PageSetupReport_ID
                obj.UserID = objCommonVar.CurrentUserCode
                obj.GridLayout = New MemoryStream()
                Gv1.SaveLayout(obj.GridLayout)
                obj.GridLayout.Seek(0, System.IO.SeekOrigin.Begin)
                obj.GridColumns = Gv1.ColumnCount
                If obj.SaveData() Then
                    common.clsCommon.MyMessageBoxShow(Me, "Layout saved successfully", Me.Text)
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
            clsGridLayout.DeleteData(PageSetupReport_ID, objCommonVar.CurrentUserCode)
            common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rbtnMCCWise_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnMCCWise.ToggleStateChanged
        Try
            If rbtnMCCWise.IsChecked = True Then
                lblType.Visible = True
                ddlType.Visible = True
            ElseIf rbtnMCCWise.IsChecked = False Then
                lblType.Visible = False
                ddlType.Visible = False
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    '    Dim value As Double = 15.35
    'Dim roundedValue As Double = Math.Ceiling(value)

    'Console.WriteLine("Original Value: " & value)
    'Console.WriteLine("Rounded Up Value: " & roundedValue)

    Private Sub txtBlock__My_Click(sender As Object, e As EventArgs) Handles txtBlock._My_Click
        Dim qry As String = " select TSPL_BLOCK_MASTER.BLOCK_CODE as Code , TSPL_BLOCK_MASTER.BLOCK_NAME as Name from TSPL_BLOCK_MASTER "
        txtBlock.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelBlock@MPIncentiveEntryRPT", qry, "Code", "Code", txtBlock.arrValueMember, txtBlock.arrDispalyMember)
    End Sub

    Private Sub txtZone__My_Click(sender As Object, e As EventArgs) Handles txtZone._My_Click
        Dim qry As String = " select TSPL_ZONE_MASTER.Zone_Code as Code , TSPL_ZONE_MASTER.Description as Name from TSPL_ZONE_MASTER where 2=2 "
        If clsCommon.myLen(objCommonVar.strCurrUserZones) > 0 Then
            qry += "  and TSPL_ZONE_MASTER.Zone_Code in (" + objCommonVar.strCurrUserZones + ") "
        End If
        txtZone.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelZone@MPIncentiveEntryRPT", qry, "Code", "Code", txtZone.arrValueMember, txtZone.arrDispalyMember)
    End Sub

    Private Sub txtRevenueVillage__My_Click(sender As Object, e As EventArgs) Handles txtRevenueVillage._My_Click
        Dim qry As String = " select TSPL_REVENUE_VILLAGE_MASTER.REVENUE_VILLAGE_CODE as Code , TSPL_REVENUE_VILLAGE_MASTER.REVENUE_VILLAGE_NAME as Name from TSPL_REVENUE_VILLAGE_MASTER "
        txtRevenueVillage.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelRevenueVillage@MPIncentiveEntryRPT", qry, "Code", "Code", txtRevenueVillage.arrValueMember, txtRevenueVillage.arrDispalyMember)
    End Sub

    Private Sub txtGrampanchayat__My_Click(sender As Object, e As EventArgs) Handles txtGrampanchayat._My_Click
        Dim qry As String = " select TSPL_GRAMPANCHAYAT_MASTER.GRAMPANCHAYAT_CODE as Code , TSPL_GRAMPANCHAYAT_MASTER.GRAMPANCHAYAT_NAME as Name from TSPL_GRAMPANCHAYAT_MASTER "
        txtGrampanchayat.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelGrampanchayat@MPIncentiveEntryRPT", qry, "Code", "Code", txtGrampanchayat.arrValueMember, txtGrampanchayat.arrDispalyMember)
    End Sub

    Private Sub txtPanchayatSamiti__My_Click(sender As Object, e As EventArgs) Handles txtPanchayatSamiti._My_Click
        Dim qry As String = " select TSPL_PANCHAYAT_SAMITI_MASTER.PANCHAYAT_SAMITI_CODE as Code , TSPL_PANCHAYAT_SAMITI_MASTER.PANCHAYAT_SAMITI_NAME as Name from TSPL_PANCHAYAT_SAMITI_MASTER "
        txtPanchayatSamiti.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelPanchayatSamiti@MPIncentiveEntryRPT", qry, "Code", "Code", txtPanchayatSamiti.arrValueMember, txtPanchayatSamiti.arrDispalyMember)
    End Sub

    Private Sub txtVidhanSabha__My_Click(sender As Object, e As EventArgs) Handles txtVidhanSabha._My_Click
        Dim qry As String = " select TSPL_VIDHAN_SABHA_MASTER.VIDHAN_SABHA_CODE as Code , TSPL_VIDHAN_SABHA_MASTER.VIDHAN_SABHA_NAME as Name from TSPL_VIDHAN_SABHA_MASTER "
        txtVidhanSabha.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelVidhanSabha@MPIncentiveEntryRPT", qry, "Code", "Code", txtVidhanSabha.arrValueMember, txtVidhanSabha.arrDispalyMember)
    End Sub

    Private Sub txtDistrict__My_Click(sender As Object, e As EventArgs) Handles txtDistrict._My_Click
        Dim qry As String = " select TSPL_DISTRICT_MASTER.Code as Code , TSPL_DISTRICT_MASTER.Name as Name from TSPL_DISTRICT_MASTER "
        txtDistrict.arrValueMember = clsCommon.ShowMultipleSelectForm("MulSelDistrict@MPIncentiveEntryRPT", qry, "Code", "Code", txtDistrict.arrValueMember, txtDistrict.arrDispalyMember)
    End Sub

    Private Sub RadButton1_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        Print(True)
    End Sub

    Private Sub rbtnDuplicateACNo_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnBankWise.ToggleStateChanged, rbtnMCCWise.ToggleStateChanged, rbtnDetails.ToggleStateChanged, rbtnNEFT.ToggleStateChanged, rbtnDuplicateACNo.ToggleStateChanged, rbtnDuplicateJanAdharNo.ToggleStateChanged, rbtnFarmerBankWiseDetail.ToggleStateChanged, rbtnFarmerBankWiseSummary.ToggleStateChanged
        btnPrint.Visible = False
        If rbtnFarmerBankWiseDetail.IsChecked OrElse rbtnFarmerBankWiseSummary.IsChecked Then
            btnPrint.Visible = True
        End If
    End Sub
End Class
