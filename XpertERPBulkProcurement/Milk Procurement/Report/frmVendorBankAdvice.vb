Imports common
Imports System.ComponentModel
Imports System.IO

'by Sanjay - Create new report 
Public Class frmVendorBankAdvice
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim MultipleFinderFillAuto As Boolean = False
    Dim DCSWiseFilterEnable As Boolean = False
    Dim ConvertVlcCodeUploaderToInt As Boolean = False
    Dim AreaWiseBilling As Boolean = False
    Dim VendorBankAdviceForSWM As Boolean = False
    Dim StrPermission As String
    Dim dtREJECT As DataTable
    Dim IsBankAdviseStartDate As String
    Dim ExportBankWiseQry As String = Nothing
    Public DocNo As String = Nothing
    Public MCC As String = Nothing
    Dim SettVSPHoldPaymentNotCompanyBank As Boolean = False
    Dim IsSeprateBankForDCSBankAdvice As String
    Dim dtBankAdvise As DataTable = New DataTable
    Public isSendMail As Boolean = False
    Public returnBankAdviseQry As String = Nothing

#End Region
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnExp.Visible = MyBase.isExport
    End Sub

    Private Sub rptMilkBillProcurementSummary_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            FormLoad()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub FormLoad()
        SettVSPHoldPaymentNotCompanyBank = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.VSPHoldPaymentNotCompanyBank, clsFixedParameterCode.VSPHoldPaymentNotCompanyBank, Nothing)) = 1)
        IsSeprateBankForDCSBankAdvice = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.SeprateBankForDCSBankAdvice, clsFixedParameterCode.SeprateBankForDCSBankAdvice, Nothing))

        DCSWiseFilterEnable = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DCSWiseFilterEnableOnSavingCheck, clsFixedParameterCode.DCSWiseFilterEnableOnSavingCheck, Nothing)) = 0)
        IsBankAdviseStartDate = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.BankAdviseRequired, clsFixedParameterCode.BankAdviseRequired, Nothing))
        MultipleFinderFillAuto = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.MultipleFinderFillAuto, clsFixedParameterCode.MultipleFinderFillAuto, Nothing)) = 1)
        ConvertVlcCodeUploaderToInt = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ConvertVlcCodeUploaderToInt, clsFixedParameterCode.ConvertVlcCodeUploaderToInt, Nothing)) = 1)
        StrPermission = clsERPFuncationality.UserWiseAvailableLocationCode()
        Reset()
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") = CompairStringResult.Equal Then
            btnExportBankWise.Visible = True
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
            btnBankSummary.Visible = True
        End If
        'RadGroupBox1.Visible = Not MultipleFinderFillAuto
        txtPaymentCycleFrom.Enabled = Not MultipleFinderFillAuto
        txtPaymentCycleTo.Enabled = Not MultipleFinderFillAuto
        VendorBankAdviceForSWM = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.VendorBankAdviceForSWM, clsFixedParameterCode.VendorBankAdviceForSWM, Nothing)) = 1)
        btnPrintSWM.Visible = VendorBankAdviceForSWM
        'RadGroupBox3.Visible =MultipleFinderFillAuto
        'If MultipleFinderFillAuto = False Then
        '    RadGroupBox3.Visible = True
        'Else
        '    RadGroupBox3.Visible = MultipleFinderFillAuto
        'End If
        'RadGroupBox3.Visible = True
        txtMCC.Enabled = Not MultipleFinderFillAuto
        'rbtnSaving.Visible = MultipleFinderFillAuto
        AreaWiseBilling = (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AreaWiseBilling, clsFixedParameterCode.AreaWiseBilling, Nothing)) = 1)
        fndArea.Visible = AreaWiseBilling
        lblArea.Visible = AreaWiseBilling

        txtPaymentCycleFrom.Enabled = Not AreaWiseBilling
        txtPaymentCycleTo.Enabled = Not AreaWiseBilling
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BRN") = CompairStringResult.Equal Then
            RadGroupBox4.Visible = True
        Else
            RadGroupBox4.Visible = False
        End If
        If clsCommon.myLen(DocNo) > 0 Then
            Print(True, DocNo, MCC)
        End If
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
            'RadPageViewPage3
            'RadPageView1
            RadPageViewPage3.Visible = True
        Else
            RadPageViewPage3.Visible = False
        End If
    End Sub


    Sub Reset()
        fromDate.Value = clsCommon.GETSERVERDATE()
        ToDate.Value = clsCommon.GETSERVERDATE()
        dtREJECT = Nothing
        txtFiscalYear.Value = objCommonVar.CurrFiscalYear
        txtPaymentCycleFrom.Value = ""
        txtPaymentCycleTo.Value = ""
        btnGo.Enabled = True
        Gv1.DataSource = Nothing
        Gv1.Rows.Clear()
        Gv1.Columns.Clear()
        rbtnSaving.IsChecked = False
        rbtnCompulsoryWiseSummary.IsChecked = False
        rbtnCompulsory.IsChecked = False
        rbtnBankAdvice.IsChecked = True
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()
        txtMCC.Value = ""
        RadPageView1.SelectedPage = RadPageViewPage1
        lblBankAdviseNo.Visible = False
        txtBankAdviseNo.Visible = False
    End Sub

    Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
        Try
            If rbtnBankAdvice.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_CBA"
            ElseIf rbtnBankWiseSummary.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_CBS"
            ElseIf rbtnCurrentBankWiseSummary.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_CCBS"
            End If

            If rbtnSaving.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_SA"
            ElseIf rbtnCompulsory.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_SC"
            ElseIf rbtnCompulsoryWiseSummary.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_SBS"
            ElseIf rbtnSavingSummary.IsChecked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "_SS"
            Else
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID)
            End If

            If ChkIFSCCode.Checked = True Then
                PageSetupReport_ID = clsCommon.myCstr(MyBase.Form_ID) + "IFSC"
            End If
            GetReportID()
            If dtBankAdvise IsNot Nothing AndAlso dtBankAdvise.Rows.Count > 1 Then
                If rbtnBothSavCur.IsChecked Then
                    If clsCommon.myLen(txtBankAdviseNo.Value) <= 0 Then
                        Throw New Exception("Please select atleast one Bank Advise No")
                    End If
                    Print(False, "", Nothing, txtBankAdviseNo.Value)
                End If
            Else
                Print(False, Nothing, Nothing)
            End If


            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                OneDayStop(txtBankAdviseNo.Value)
            End If
            ReStoreGridLayout()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub OneDayStop(ByVal BankAdviseDocNo As String)
        Try
            Dim FinalQuery As String = ""
            Dim dt As New DataTable
            FinalQuery = "Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,xxxfinal.* from (" + GetSavingCurrentALW(BankAdviseDocNo) + ")xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='" + objCommonVar.CurrComp_Code1 + "' where xxxfinal.DCSCount=1 "
            dt = clsDBFuncationality.GetDataTable(FinalQuery)
            Gv2.MasterTemplate.SummaryRowsBottom.Clear()
            Gv2.MasterTemplate.FilterDescriptors.Clear()
            Gv2.DataSource = Nothing
            Gv2.Rows.Clear()
            Gv2.Columns.Clear()
            Gv2.GroupDescriptors.Clear()
            Gv2.MasterTemplate.SummaryRowsBottom.Clear()
            Gv2.MasterView.Refresh()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found/Posted to Display", Me.Text)
                Exit Sub
            Else
                Gv2.DataSource = dt
                RadPageView1.SelectedPage = RadPageViewPage2
                SetGridFormatOneDayStop()
                EnableDisaableControls(False)
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub SetGridFormatOneDayStop()
        Gv2.ShowGroupPanel = False
        Gv2.ShowRowHeaderColumn = False
        Gv2.AllowAddNewRow = False
        Gv2.AllowDeleteRow = False
        Gv2.EnableFiltering = True
        Gv2.ShowFilteringRow = True
        Gv2.MasterTemplate.SummaryRowsBottom.Clear()

        For ii As Integer = 0 To Gv2.Columns.Count - 1
            Gv2.Columns(ii).ReadOnly = True
            Gv2.Columns(ii).BestFit()
        Next

        If Gv2.Columns.Contains("DCSCount") Then
            Gv2.Columns("DCSCount").HeaderText = "DCS Count"
            Gv2.Columns("DCSCount").IsVisible = False
        End If

        Gv2.Columns("VLC_CODE_Uploader").HeaderText = "DCS Code"
        Gv2.Columns("VLC_CODE_Uploader").IsVisible = True

        'FOR ALW Temp Work
        Gv2.Columns("RI").IsVisible = False
        Gv2.Columns("Branch_Name").IsVisible = False
        Gv2.Columns("DateRange").IsVisible = False

        Gv2.Columns("BankCode").HeaderText = "Bank Code"
        Gv2.Columns("BankCode").IsVisible = True

        Gv2.Columns("BankName").HeaderText = "Bank Name"
        Gv2.Columns("BankName").IsVisible = True

        Gv2.Columns("DCS_Name").HeaderText = "DCS Name"
        Gv2.Columns("DCS_Name").IsVisible = True

        Gv2.Columns("Bill_Amt").HeaderText = "Bill Amount"
        Gv2.Columns("Bill_Amt").IsVisible = True

        Gv2.Columns("SavingAccountNo").HeaderText = "Saving Account No."
        Gv2.Columns("SavingAccountNo").IsVisible = True

        Gv2.Columns("SavingAmt").HeaderText = "Saving Amount"
        Gv2.Columns("SavingAmt").IsVisible = True

        Gv2.Columns("CurrentAccountNo").HeaderText = "Current Account No."
        Gv2.Columns("CurrentAccountNo").IsVisible = True

        Gv2.Columns("CurrentAmt").HeaderText = "Current Amount"
        Gv2.Columns("CurrentAmt").IsVisible = True

        Dim summaryRowItemB As New GridViewSummaryRowItem()
        'Dim MilkTypeB As New GridViewSummaryItem("Payable_Amount", "{0:n0}", GridAggregateFunction.Sum)
        Dim Bill_Amt As New GridViewSummaryItem("Bill_Amt", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(Bill_Amt)
        Dim SavingAmt As New GridViewSummaryItem("SavingAmt", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(SavingAmt)
        Dim CurrentAmt As New GridViewSummaryItem("CurrentAmt", "{0:n2}", GridAggregateFunction.Sum)
        summaryRowItemB.Add(CurrentAmt)
        Gv2.MasterTemplate.SummaryRowsBottom.Add(summaryRowItemB)
        Gv2.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom

    End Sub
    Sub GetReportID()
        Dim VarID As String = ""
        If rbtnBankAdvice.IsChecked Then
            VarID += "_BA"
        ElseIf rbtnBankWiseSummary.IsChecked Then
            VarID += "_BWS"
        ElseIf rbtnCurrentBankWiseSummary.IsChecked Then
            VarID += "_CBWS"
        End If

        If rbtnSaving.IsChecked Then
            VarID += "_S"
        ElseIf rbtnSavingSummary.IsChecked Then
            VarID += "_SS"
        End If

        If rbtnCompulsory.IsChecked Then
            VarID += "_C"
        ElseIf rbtnCompulsoryWiseSummary.IsChecked Then
            VarID += "_CWS"
        End If

        If rbtnBothSavCur.IsChecked Then
            VarID += "_BS"
        End If

        If ChkIFSCCode.Checked Then
            VarID += "_IFSC"
        End If
        Gv1.VarID = VarID
    End Sub

    Public Sub Print(ByVal isPrint As Boolean, ByVal BankAdviseDocNo As String, ByVal MCC As String)
        Print(isPrint, BankAdviseDocNo, MCC, "")
    End Sub
    Public Sub Print(ByVal isPrint As Boolean, ByVal BankAdviseDocNo As String, ByVal MCC As String, ByVal LocalBankAdviseDocNo As String)
        Try
            If clsCommon.myLen(BankAdviseDocNo) > 0 Then
                rbtnBankAdvice.IsChecked = True
            End If

            Dim vSeprateBankForDCSBankAdvice As String = Nothing
            If clsCommon.myLen(IsSeprateBankForDCSBankAdvice) > 0 Then
                'If clsCommon.myLen(IsSeprateBankForDCSBankAdvice) > 0 Then
                vSeprateBankForDCSBankAdvice = IsSeprateBankForDCSBankAdvice
                ' If
            End If


            If clsCommon.myLen(BankAdviseDocNo) <= 0 Then
                If rbtnBankAdvice.IsChecked AndAlso clsCommon.myLen(IsBankAdviseStartDate) < 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Plz Create Bank Advice." + Environment.NewLine + " Bank Advice Creation Date Start On '" + IsBankAdviseStartDate + "'.", Me.Text)
                    Exit Sub
                End If
                If clsCommon.myLen(txtMCC.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
                    clsCommon.MyMessageBoxShow(Me, "Plz Select MCC.", Me.Text)
                    txtMCC.Focus()
                    Exit Sub
                End If

                If clsCommon.myLen(txtFiscalYear.Value) <= 0 Then
                    clsCommon.MyMessageBoxShow(Me, "Plz Select Fiscal Year First.", Me.Text)
                    txtFiscalYear.Focus()
                    Exit Sub
                End If
                If clsCommon.myLen(txtPaymentCycleFrom.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
                    'clsCommon.MyMessageBoxShow("Plz Select Payment Cycle From First.", Me.Text)
                    'txtPaymentCycleFrom.Focus()
                    'Exit Sub
                End If
                If clsCommon.myLen(txtPaymentCycleTo.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
                    'clsCommon.MyMessageBoxShow("Plz Select Payment Cycle To First.", Me.Text)
                    'txtPaymentCycleTo.Focus()
                    'Exit Sub
                End If

                If clsCommon.myCdbl(txtPaymentCycleFrom.Value) > clsCommon.myCdbl(txtPaymentCycleTo.Value) AndAlso MultipleFinderFillAuto = False Then
                    common.clsCommon.MyMessageBoxShow(Me, "From Payment Cycle can not be greater then to Payment Cycle", Me.Text)
                    txtPaymentCycleFrom.Focus()
                    Exit Sub
                End If
                If MultipleFinderFillAuto = False Then
                    Patment_Cycle_changed()
                End If
                If AreaWiseBilling = False Then
                    Patment_Cycle_changed()
                End If
                If clsCommon.myCdbl(txtPaymentCycleFrom.Value) > clsCommon.myCdbl(txtPaymentCycleTo.Value) AndAlso AreaWiseBilling = False Then
                    common.clsCommon.MyMessageBoxShow(Me, "From Payment Cycle can not be greater then to Payment Cycle", Me.Text)
                    txtPaymentCycleFrom.Focus()
                    Exit Sub
                End If
            End If
            If clsCommon.myLen(LocalBankAdviseDocNo) > 0 Then
                BankAdviseDocNo = LocalBankAdviseDocNo
            End If
            Dim strCycleRange As String = txtPaymentCycleTo.Value
            If clsCommon.myCdbl(txtPaymentCycleTo.Value) > clsCommon.myCdbl(txtPaymentCycleFrom.Value) Then
                strCycleRange = txtPaymentCycleTo.Value + " To  " + txtPaymentCycleFrom.Value
            End If
            Dim BaseQry As String = ""
            Dim Doc_No_Value As String = "select Doc_No from TSPL_PAYMENT_PROCESS_Invoice"


            If rbtnCompulsoryWiseSummary.IsChecked AndAlso rbtnCurrentBankWiseSummary.IsChecked Then
                clsCommon.MyMessageBoxShow(Me, "You have to select only option at a time ", Me.Text)
                Exit Sub
            End If
            If rbtnSaving.IsChecked OrElse rbtnSavingSummary.IsChecked Then
                BaseQry = " select  TSPL_BANK_ADVISE.Document_No As Bank_Advise_No,Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As Bank_Advise_Date,'" + fromDate.Value + "' as FD,'" + ToDate.Value + "' AS TD,'' AS CycleRange,TSPL_Vendor_MASTER.BankCode2 as GRPColumn,TSPL_COMPANY_MASTER.Comp_Name
,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No,TSPL_Fiscal_Year_Master.Fiscal_Name
,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader
,coalesce(TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name_Saving,TSPL_VENDOR_MASTER.VSP_Payee_Name,Mp_V.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code
,TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name, TSPL_BANK_MASTER2.DESCRIPTION as Bank_Code_Desc
, case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment, Mp_V.vsp_payment),'')= 'Self' then coalesce(TSPL_VENDOR_MASTER.IFSCCode2, mp_V.IFSC_Code ) else coalesce( TSPL_VENDOR_MASTER.Joint_IFSC_Code,mp_v.Joint_IFSC_Code) end as Payee_Joint_IFSC_Code
, case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment, Mp_V.vsp_payment), '')= 'Self' then coalesce(TSPL_VENDOR_MASTER.AccNo2, mp_V.Account_No) else coalesce( TSPL_VENDOR_MASTER.Joint_Account_No, mp_V.Joint_Account_No) end as Payee_Joint_Account_No,"
                If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                    BaseQry += " Round(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as Payable_Amount "
                Else
                    BaseQry += " TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount "
                End If
                BaseQry += " from TSPL_PAYMENT_PROCESS_SAVING 
left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
left outer join TSPL_PAYMENT_PROCESS_DETAIL on TSPL_PAYMENT_PROCESS_DETAIL.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No and TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
left outer join TSPL_MP_MASTER mp on mp.MP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
Left outer join tspl_vendor_master Mp_V on mp_V.Vendor_Code=mp.MP_Code  
left outer join ( select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_SAVING.VSP_CODE
left outer join TSPL_BANK_MASTER as TSPL_BANK_MASTER2 on TSPL_BANK_MASTER2.BANK_CODE=TSPL_Vendor_MASTER.BankCode2 
left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Bank_Code
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No "
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    BaseQry += " left outer join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
                                where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_VENDOR_INVOICE_DETAIL.DeductionCode <>'PDP' And TSPL_PAYMENT_PROCESS_HEAD.From_Date >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0  "
                Else
                    BaseQry += " where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 And TSPL_PAYMENT_PROCESS_HEAD.From_Date >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0  "
                End If


                If clsCommon.myLen(txtMCC.Value) > 0 Then
                    BaseQry += " And TSPL_MCC_MASTER.MCC_Code = '" + txtMCC.Value + "' "
                End If
                If clsCommon.myLen(fndArea.Value) > 0 Then
                    BaseQry += " And TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code = '" + fndArea.Value + "' "
                End If
                If txtDCS.arrValueMember IsNot Nothing AndAlso txtDCS.arrValueMember.Count > 0 Then
                    BaseQry += " and TSPL_VLC_MASTER_HEAD.VSP_Code in (" + clsCommon.GetMulcallString(txtDCS.arrValueMember) + ")  "
                End If
                'If rbtnSavingSummary.IsChecked Then
                '    BaseQry = " select  ROW_NUMBER() over ( order by Bank_Code) as SNO ,max(Bank_Code) as GRPColumn,sum(Payable_Amount) as Payable_Amount ,max(Branch_Name) as Bank_Desc from (" + BaseQry + ") xx group by xx.Bank_Code "
                'End If
            ElseIf rbtnCompulsory.IsChecked Then
                BaseQry = " select max(x.CycleRange)CycleRange,x.GRPColumn,max(x.[Company Bank])[Company Bank],max(x.[Company Bank Account No])[Company Bank Account No],
                            max(x.Comp_Name)Comp_Name,max(x.Comp_address)Comp_address,max(x.CompPhone)CompPhone,max(x.Regn_No)Regn_No,max(x.MCC_NAME)MCC_NAME,
                            max(x.From_Date)From_Date,max(x.GSTReg_No)GSTReg_No,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.CycleNo)CycleNo,
                            max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,
                            max(x.Bank_Code_Desc)Bank_Code_Desc,max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount from
                            (select  '' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code+(case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment,Mp_V.vsp_payment),'')='Self' then coalesce(TSPL_VENDOR_MASTER .IFSC_Code,mp_V.IFSC_Code)   else coalesce(TSPL_VENDOR_MASTER .Joint_IFSC_Code,mp_v.Joint_IFSC_Code)   end)  as GRPColumn,TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
                            TSPL_COMPANY_MASTER.Comp_Name
                            ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name
                            ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, coalesce(TSPL_VENDOR_MASTER.VSP_Payee_Name,Mp_V.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
  TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name, 
  TSPL_Vendor_MASTER.BankCode2 as Bank_Code_Desc, 
  case when isnull(
    coalesce(
      TSPL_VENDOR_MASTER.vsp_payment, 
      Mp_V.vsp_payment
    ), 
    ''
  )= 'Self' then coalesce(
    TSPL_VENDOR_MASTER.IFSCCode2, mp_V.IFSC_Code
  ) else coalesce(
    TSPL_VENDOR_MASTER.Joint_IFSC_Code, 
    mp_v.Joint_IFSC_Code
  ) end as Payee_Joint_IFSC_Code, 
  case when isnull(
    coalesce(
      TSPL_VENDOR_MASTER.vsp_payment, 
      Mp_V.vsp_payment
    ), 
    ''
  )= 'Self' then coalesce( 
    TSPL_VENDOR_MASTER.AccNo2, mp_V.Account_No
  ) else coalesce(
    TSPL_VENDOR_MASTER.Joint_Account_No, 
    mp_V.Joint_Account_No
  ) end as Payee_Joint_Account_No,"

                If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                    BaseQry += " Round(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as Payable_Amount "
                Else
                    BaseQry += " TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount  "

                End If

                BaseQry += " from TSPL_PAYMENT_PROCESS_COMPULSORY 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No
                            left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No
                            left outer join TSPL_MP_MASTER mp on mp.MP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
                            Left outer join tspl_vendor_master Mp_V on mp_V.Vendor_Code=mp.MP_Code  
                            left outer join ( select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                            left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_COMPULSORY.VSP_CODE
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'"
                If clsCommon.myLen(txtMCC.Value) > 0 Then
                    BaseQry += " And TSPL_MCC_MASTER.MCC_Code = '" + txtMCC.Value + "' "
                End If
                If clsCommon.myLen(fndArea.Value) > 0 Then
                    BaseQry += " And TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code = '" + fndArea.Value + "' "
                End If
                BaseQry += " and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0 )x
							group by x.GRPColumn,x.VLC_CODE_Uploader,x.Payee_Joint_Account_No "

            ElseIf rbtnCompulsoryWiseSummary.IsChecked Then
                BaseQry = "select  '" + strCycleRange + "' AS CycleRange,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code+TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code as GRPColumn,"
                Else
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code as GRPColumn,"
                End If

                BaseQry += "TSPL_COMPANY_MASTER.Comp_Name
,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME
,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No," + IIf(MultipleFinderFillAuto = True, "", " TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc, ") + " TSPL_Fiscal_Year_Master.Fiscal_Name
,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_BANK_MASTER.DESCRIPTION as Bank_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,TSPL_PAYMENT_PROCESS_DETAIL.SNo,TSPL_COMPANY_MASTER.logo_img,"

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " Round(isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                Else
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        BaseQry += " Round(isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                    Else
                        BaseQry += " isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0) as Payable_Amount  "

                    End If

                End If
                BaseQry += " from TSPL_PAYMENT_PROCESS_DETAIL 
left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank 
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected

" + IIf(MultipleFinderFillAuto = True, "    ", " left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' ") + "
left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) " + IIf(MultipleFinderFillAuto = True, "  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected  ", " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code ") + " 
where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, "and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + txtMCC.Value + "' And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + "   and TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount>0 "
                'If clsCommon.myLen(txtMCC.Value) > 0 Then
                '    BaseQry += " And TSPL_MCC_MASTER.MCC_Code = '" + txtMCC.Value + "' "
                'End If

                If clsCommon.myLen(fndArea.Value) > 0 Then
                    BaseQry += " And TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code = '" + fndArea.Value + "' "
                End If
            End If

            If rbtnBankAdvice.IsChecked OrElse rbtnBankWiseSummary.IsChecked Then
                ''Note IF You do any changes than change in function clsBankAdvise.CreateEmailContent(ByVal strDateRange As String, trans As SqlTransaction)
                If rbtnSaving.IsChecked = False AndAlso rbtnCompulsory.IsChecked = False AndAlso rbtnCompulsoryWiseSummary.IsChecked = False Then
                    '                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal AndAlso clsCommon.myLen(IsSeprateBankForDCSBankAdvice) > 0 Then
                    '                        BaseQry = " select max(xxx.CycleRange)CycleRange,	max(xxx.GRPColumn)GRPColumn	,xxx.GRPColumns	,max(xxx.FD)FD,max(xxx.TD)TD,max(xxx.Bank_Name)Bank_Name,	max(xxx.BankAccountNo)BankAccountNo,max(xxx.BankIFSCCode)BankIFSCCode,max(xxx.BankBranchAddress)BankBranchAddress,	max([Company Bank])	as [Company Bank], max([Company Bank Account No]) as [Company Bank Account No],	max(Comp_Name)Comp_Name,	max(Comp_address)Comp_address,	max(CompPhone)CompPhone,	max(Regn_No)Regn_No,max(MCC_NAME)MCC_NAME,max(From_Date)From_Date,	max(GSTReg_No)GSTReg_No,	max(Doc_No)Doc_No,	max(Location_Code)Location_Code,max(Location_Desc)Location_Desc,	max(Fiscal_Name)Fiscal_Name,max(CycleNo)CycleNo,	max(Date_Range)Date_Range,	max(VLC_CODE_Uploader)VLC_CODE_Uploader,	max(Payee_Joint_Name)Payee_Joint_Name,max(Bank_Code)Bank_Code,	max(Branch_Name)Branch_Name,max(Bank_Code_Desc)Bank_Code_Desc,	max(Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,	max(Payee_Joint_Account_No)Payee_Joint_Account_No,	sum(Payable_Amount)Payable_Amount,	max([Bank Advise No]) as [Bank Advise No],	max([Bank Advise Date]) as [Bank Advise Date]	
                    ',max([Bank Advice Status])[Bank Advice Status] FROM( "
                    '                    End If


                    BaseQry = "select  '" + strCycleRange + "' AS CycleRange,"
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        BaseQry += " TSPL_Vendor_MASTER.Bank_Code+TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code as GRPColumn,"

                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal AndAlso SettVSPHoldPaymentNotCompanyBank = True Then
                        BaseQry += "  TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Branch_Name as GRPColumn,"
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                        BaseQry += "  CASE WHEN TSPL_Vendor_MASTER.Bank_Code='CASH' THEN 'CASH' WHEN TSPL_Vendor_MASTER.Bank_Code='CENTRAL CO-OPERATIVE BANK' THEN 'CENTRAL CO-OPERATIVE BANK' ELSE 'Other Banks' END AS GRPColumn,"
                    Else
                        BaseQry += " TSPL_Vendor_MASTER.Bank_Code as GRPColumn,"
                    End If
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal AndAlso clsCommon.myLen(IsSeprateBankForDCSBankAdvice) > 0 Then
                        BaseQry += " CASE WHEN TSPL_Vendor_MASTER.Bank_Code LIKE '" + vSeprateBankForDCSBankAdvice + "' THEN 'PNB Bank' ELSE 'Other Banks' END AS GRPColumns,"

                    Else
                        BaseQry += " CASE WHEN TSPL_Vendor_MASTER.Bank_Code LIKE 'PNB%' THEN 'PNB Bank' ELSE 'Other Banks' END AS GRPColumns,"

                    End If
                    BaseQry += "'" + fromDate.Value + "' as FD,'" + ToDate.Value + "' AS TD,TSPL_COMPANY_MASTER.Bank_Name,TSPL_COMPANY_MASTER.BankAccountNo,TSPL_COMPANY_MASTER.BankIFSCCode,TSPL_COMPANY_MASTER.BankBranchAddress,
                               TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
                               TSPL_COMPANY_MASTER.Comp_Name
                               ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
                               ,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,"
                    If AreaWiseBilling = True Then
                        BaseQry += " TSPL_LOCATION_MASTER.Location_Desc AS MCC_Name "
                    Else
                        BaseQry += " TSPL_MCC_MASTER.MCC_NAME "
                    End If
                    BaseQry += ",TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No," + IIf(MultipleFinderFillAuto = True, "", " TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc, ") + " TSPL_Fiscal_Year_Master.Fiscal_Name
,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,"

                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal AndAlso SettVSPHoldPaymentNotCompanyBank = True Then
                        'If SettVSPHoldPaymentNotCompanyBank = True Then
                        BaseQry += "  TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Branch_Name ,"
                        'End If
                    Else
                        BaseQry += "  TSPL_Vendor_MASTER.Bank_Code,"
                    End If
                    BaseQry += " TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,"

                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        BaseQry += " Round((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)),0) as Payable_Amount "
                    Else
                        If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                            BaseQry += " Round((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)-isnull(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0)),0) as Payable_Amount  "
                        Else
                            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal Then
                                ' BaseQry += " Cast((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)) as decimal(18)) as Payable_Amount "
                                BaseQry += " (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0))  as Payable_Amount  "

                            Else
                                BaseQry += " (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)-isnull(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0))  as Payable_Amount  "
                            End If
                        End If

                    End If
                End If
                BaseQry += ",TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_BANK_ADVISE.Status IS NULL OR TSPL_BANK_ADVISE.Status =0 Then 'Pending' Else 'Approved' End As [Bank Advice Status] "
                BaseQry += " from TSPL_PAYMENT_PROCESS_DETAIL 
                                left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                                left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
                                left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
                                left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                                left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current "
                If AreaWiseBilling = True Then
                    BaseQry += " LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code"
                Else
                    BaseQry += " left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected"

                End If

                BaseQry += " left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                                left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No "

                BaseQry += "" + IIf(MultipleFinderFillAuto = True, "    ", " left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' ") + "
left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) " + IIf(MultipleFinderFillAuto = True, "  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected  ", " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code ") + " "

                If clsCommon.myLen(BankAdviseDocNo) <= 0 Then
                    BaseQry += " where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, "and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + txtMCC.Value + "' And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + " "
                    'BaseQry += " where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, " And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + " "
                Else
                    BaseQry += " where TSPL_BANK_ADVISE.Document_No='" + BankAdviseDocNo + "' "
                    If clsCommon.myLen(MCC) > 0 AndAlso MultipleFinderFillAuto = False Then
                        BaseQry += " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + MCC + "' And TSPL_MCC_MASTER.MCC_Code='" + MCC + "' "
                    End If
                    'If clsCommon.myLen(MCC) > 0 AndAlso MultipleFinderFillAuto = False Then
                    '    BaseQry += " And TSPL_MCC_MASTER.MCC_Code='" + MCC + "' "
                    'End If
                End If

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHT") <> CompairStringResult.Equal Then
                    BaseQry += " And (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0"
                End If

                'and (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0 "
                'If clsCommon.myLen(txtMCC.Value) > 0 Then
                '    BaseQry += " And TSPL_MCC_MASTER.MCC_Code = '" + txtMCC.Value + "' "
                'End If

                If clsCommon.myLen(fndArea.Value) > 0 Then
                    BaseQry += " And TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code = '" + fndArea.Value + "' "
                End If

                If clsCommon.myLen(IsBankAdviseStartDate) > 0 Then
                    BaseQry += "and 2=(case when TSPL_PAYMENT_PROCESS_HEAD.To_Date<'" + clsCommon.GetPrintDate(IsBankAdviseStartDate, "dd/MMM/yyyy") + "' then 2 else ( case when len(isnull(TSPL_BANK_ADVISE.Document_No,''))>0 then 2 else 3 end ) end) "
                End If

            ElseIf rbtnCurrentBankWiseSummary.IsChecked Then
                BaseQry = "select  '" + strCycleRange + "' AS CycleRange,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code+TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code as GRPColumn,"
                Else
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code as GRPColumn,"
                End If
                BaseQry += "TSPL_COMPANY_MASTER.Comp_Name
,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME
,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No," + IIf(MultipleFinderFillAuto = True, "", " TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc, ") + " TSPL_Fiscal_Year_Master.Fiscal_Name
,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc , TSPL_BANK_MASTER.DESCRIPTION as Bank_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " Round(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                Else
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        BaseQry += " Round(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                    Else
                        BaseQry += " (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)) as Payable_Amount "
                    End If
                End If
                BaseQry += " from TSPL_PAYMENT_PROCESS_DETAIL 
left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
" + IIf(MultipleFinderFillAuto = True, "    ", " left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' ") + "
left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) " + IIf(MultipleFinderFillAuto = True, "  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected  ", " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code ") + " 
where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, "and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + txtMCC.Value + "' And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + "   and TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount>0 "
                'If clsCommon.myLen(txtMCC.Value) > 0 Then
                '    BaseQry += " And TSPL_MCC_MASTER.MCC_Code = '" + txtMCC.Value + "' "
                'End If

                If clsCommon.myLen(fndArea.Value) > 0 Then
                    BaseQry += " And TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code = '" + fndArea.Value + "' "
                End If
            End If

            If clsCommon.myLen(txtBankGroup.Value) > 0 Then
                BaseQry = BaseQry + " AND TSPL_BANK_MASTER.BANK_GROUP_CODE='" + txtBankGroup.Value + "' "
            End If

            Dim FinalQuery As String = ""
            If rbtnBankAdvice.IsChecked OrElse rbtnSaving.IsChecked Then
                ''Note IF You do any changes than change in function clsBankAdvise.CreateEmailContent(ByVal strDateRange As String, trans As SqlTransaction)
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                    FinalQuery = " select "
                    If rbtnBankAdvice.IsChecked Then
                        FinalQuery += " max([Company Bank])[Company Bank],max([Company Bank Account No])[Company Bank Account No],"
                    End If
                    FinalQuery += " max(CycleRange)CycleRange,max(GRPColumn)GRPColumn,max(Comp_Name)Comp_Name,max(Comp_address)Comp_address,
                                  max(CompPhone)CompPhone,max(Regn_No)Regn_No,max(MCC_NAME)MCC_NAME,max(From_Date)From_Date,max(GSTReg_No)GSTReg_No,
                                  max(Doc_No)Doc_No,max(Fiscal_Name)Fiscal_Name,max(CycleNo)CycleNo,max(Date_Range)Date_Range,VLC_CODE_Uploader,
                                  max(Payee_Joint_Name)Payee_Joint_Name,max(Bank_Code)Bank_Code,max(Branch_Name)Branch_Name,max(Bank_Code_Desc)Bank_Code_Desc,
                                  max(Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No)Payee_Joint_Account_No,sum(Payable_Amount)Payable_Amount,
                                  max(FD)FD,max(TD)TD"

                    If rbtnSaving.IsChecked Then
                        FinalQuery += " ,max(Bank_Advise_No)Bank_Advise_No,max(Bank_Advise_Date)Bank_Advise_Date  "
                    End If
                    If rbtnBankAdvice.IsChecked Then
                        FinalQuery += " ,max([Bank Advise No])[Bank Advise No],max([Bank Advise Date])[Bank Advise Date],max([Bank Advice Status])[Bank Advice Status] "
                    End If

                    FinalQuery += "from ( " + BaseQry + ")xxx group by xxx.VLC_CODE_Uploader order by Payee_Joint_Account_No asc "
                Else
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal AndAlso clsCommon.myLen(IsSeprateBankForDCSBankAdvice) > 0 Then
                        FinalQuery = "  select max(xxx.CycleRange)CycleRange,	max(xxx.GRPColumn)GRPColumn	,xxx.GRPColumns	,max(xxx.FD)FD,max(xxx.TD)TD,max(xxx.Bank_Name)Bank_Name,	max(xxx.BankAccountNo)BankAccountNo,max(xxx.BankIFSCCode)BankIFSCCode,max(xxx.BankBranchAddress)BankBranchAddress,	max([Company Bank])	as [Company Bank], max([Company Bank Account No]) as [Company Bank Account No],	max(Comp_Name)Comp_Name,	max(Comp_address)Comp_address,	max(CompPhone)CompPhone,	max(Regn_No)Regn_No,max(MCC_NAME)MCC_NAME,max(From_Date)From_Date,	max(GSTReg_No)GSTReg_No,	max(Doc_No)Doc_No,	max(Location_Code)Location_Code,max(Location_Desc)Location_Desc,	max(Fiscal_Name)Fiscal_Name,max(CycleNo)CycleNo,	max(Date_Range)Date_Range,	max(VLC_CODE_Uploader)VLC_CODE_Uploader,	max(Payee_Joint_Name)Payee_Joint_Name,max(Bank_Code)Bank_Code,	max(Branch_Name)Branch_Name,max(Bank_Code_Desc)Bank_Code_Desc,	max(Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,	max(Payee_Joint_Account_No)Payee_Joint_Account_No,	sum(Payable_Amount)Payable_Amount,	max([Bank Advise No]) as [Bank Advise No],	max([Bank Advise Date]) as [Bank Advise Date]	
,max([Bank Advice Status])[Bank Advice Status] FROM( " + BaseQry + ")xxx group by XXX.GRPColumns,XXX.VLC_CODE_Uploader ORDER BY MAX(Bank_Code),  "
                    Else
                        FinalQuery = "select * from ( " + BaseQry + ")xxx order by Bank_Code, "
                    End If
                    ' FinalQuery = "select * from ( " + BaseQry + ")xxx order by Bank_Code, "
                    If ConvertVlcCodeUploaderToInt = True Then
                        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "SWM") = CompairStringResult.Equal AndAlso clsCommon.myLen(IsSeprateBankForDCSBankAdvice) > 0 Then
                            FinalQuery += " MAX(CAST(VLC_CODE_Uploader AS INT)) "
                        Else
                            FinalQuery += " cast(VLC_Code_Uploader as Int) "
                        End If
                        'FinalQuery += " cast(VLC_Code_Uploader as Int) "
                    Else
                        FinalQuery += " VLC_Code_Uploader "
                    End If
                End If
            ElseIf rbtnBankWiseSummary.IsChecked Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                    FinalQuery += "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select max(CycleRange) as CycleRange, max(GRPColumn) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,max(Branch_Name)Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,sum(Payable_Amount) as Payable_Amount
,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
from (" + Environment.NewLine + BaseQry + Environment.NewLine + " )xxx group by Bank_Code )xxxx order by GRPColumn "
                Else
                    FinalQuery += "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select max(CycleRange) as CycleRange, max(GRPColumn) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,sum(Payable_Amount) as Payable_Amount
,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
from (" + Environment.NewLine + BaseQry + Environment.NewLine + " )xxx group by Bank_Code,Branch_Name )xxxx order by GRPColumn "
                End If

            ElseIf rbtnCurrentBankWiseSummary.IsChecked OrElse rbtnCompulsoryWiseSummary.IsChecked Then
                FinalQuery += "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select  max(GRPColumn) as GRPColumn,sum(Payable_Amount) as Payable_Amount , Bank_Desc
from (" + Environment.NewLine + BaseQry + Environment.NewLine + " )xxx group by Bank_Desc )xxxx order by GRPColumn "
            ElseIf rbtnSavingSummary.IsChecked Then
                FinalQuery += "select ROW_NUMBER() over ( order by Bank_Code) as SNO , max(CycleRange) as CycleRange, max(Bank_Code) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,max (Bank_Code_Desc)Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,sum(Payable_Amount) as Payable_Amount
,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME,max(FD)FD,max(TD)TD
from (" + Environment.NewLine + BaseQry + Environment.NewLine + "   )xxx group by Bank_Code "
            End If
            If rbtnCompulsory.IsChecked Then
                FinalQuery = BaseQry
            End If

            If rbtnBothSavCur.IsChecked AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BRN") = CompairStringResult.Equal Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    FinalQuery = GetSavingCurrentALW(BankAdviseDocNo)
                Else
                    FinalQuery = GetSavingCurrent()
                End If

                If ConvertVlcCodeUploaderToInt Then
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                        'FinalQuery += " Select * from (" + FinalQuery + ")xx where  xx.dcscount <> 1 order by cast(VLC_CODE_Uploader as int) "
                        FinalQuery = " Select * from (" + GetSavingCurrentALW(BankAdviseDocNo) + ")xx where  xx.dcscount <> 1 order by cast(VLC_CODE_Uploader as int) "
                    Else
                        FinalQuery += " order by cast(VLC_CODE_Uploader as int) "
                    End If
                    'FinalQuery += " order by cast(VLC_CODE_Uploader as int) "
                Else
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                        FinalQuery += " Select * from (" + GetSavingCurrentALW(BankAdviseDocNo) + ")xx where xx.dcscount <> 1 order by  VLC_CODE_Uploader  "
                    Else
                        FinalQuery += " order by  VLC_CODE_Uploader  "
                    End If
                    ' FinalQuery += " order by  VLC_CODE_Uploader  "
                End If
            End If
            If isSendMail Then
                returnBankAdviseQry = FinalQuery
                Exit Sub
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(FinalQuery)
            ExportBankWiseQry = FinalQuery

            Dim dt1 As DataTable = Nothing
            Dim dt2 As DataTable = Nothing
            Dim dt3 As DataTable = Nothing
            Dim dt7 As DataTable = Nothing
            Dim dt8 As DataTable = Nothing
            If isPrint AndAlso rbtnBothSavCur.IsChecked AndAlso clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BRN") = CompairStringResult.Equal Then
                FinalQuery = ""
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    FinalQuery = "Select ROW_NUMBER() OVER(Partition by xxxfinal.bankcode ORDER BY xxxfinal.bankcode) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,xxxfinal.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,
                              TSPL_COMPANY_MASTER.Regn_No,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (" + GetSavingCurrentALW(BankAdviseDocNo) + ")xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='" + objCommonVar.CurrComp_Code1 + "' where xxxfinal.DCSCount>1 "

                Else
                    FinalQuery = "Select ROW_NUMBER() OVER(Partition by xxxfinal.bankcode ORDER BY xxxfinal.bankcode) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,xxxfinal.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,
                              TSPL_COMPANY_MASTER.Regn_No,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (" + GetSavingCurrent() + ")xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='" + objCommonVar.CurrComp_Code1 + "' where xxxfinal.DCSCount>1 "

                End If
                If ConvertVlcCodeUploaderToInt Then
                    FinalQuery += "order by cast(VLC_CODE_Uploader as int) "
                Else
                    FinalQuery += "order by VLC_CODE_Uploader "
                End If
                dt1 = clsDBFuncationality.GetDataTable(FinalQuery)

                FinalQuery = ""
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    FinalQuery = "Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,xxxfinal.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,
                              TSPL_COMPANY_MASTER.Regn_No,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (" + GetSavingCurrentALW(BankAdviseDocNo) + ")xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='" + objCommonVar.CurrComp_Code1 + "' where xxxfinal.DCSCount=1 "
                Else
                    FinalQuery = "Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,xxxfinal.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,
                              TSPL_COMPANY_MASTER.Regn_No,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (" + GetSavingCurrent() + ")xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='" + objCommonVar.CurrComp_Code1 + "' where xxxfinal.DCSCount=1 "
                End If

                dt2 = clsDBFuncationality.GetDataTable(FinalQuery)

                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                    FinalQuery = ""
                    FinalQuery = "Select ('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,final.BankCode,Max(final.BankName)BankName,Sum(final.Bill_Amt)TotalBillAmt,Sum(final.SavingAmt)TotalSavingAmt,Sum(final.CurrentAmt)CurrentAmt,max(final.DCSCount)DCSCount,max(Comp_Name)Comp_Name,max(CompPhone)CompPhone,max(Regn_No)Regn_No,max(GSTReg_No)GSTReg_No,max(Pincode)Pincode from (Select ROW_NUMBER() OVER(Partition by xxxfinal.bankcode ORDER BY xxxfinal.bankcode) As SNo,('16/08/2025'+' to '+'31/08/2025') As DateRange,xxxfinal.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,
                              TSPL_COMPANY_MASTER.Regn_No,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (   --select * from bal

SELECT ca.RI,t.[Bank Advise No],t.[Bank Advise Date],t.VLC_CODE_Uploader,ca.Bank as BankCode,ca.Bank as [Branch_Name],ca.Bank as [BankName],t.DCS_Name

,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Bill_Amt] else (case when AccountType='Saving' then t.[Saving_Amount] else  t.[CurrentAmt] end) end )As [Bill_Amt]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Bank_Account_No_Saving] else (case when AccountType='Saving' then t.[Bank_Account_No_Saving] else '' end) end ) as [SavingAccountNo]
,ROUND(CASE WHEN t.Bank_Code = t.Bank_Code_Saving THEN t.Saving_Amount ELSE CASE WHEN AccountType = 'Saving' THEN t.Saving_Amount ELSE 0 END END , 0) AS [SavingAmt]


,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Payee_Joint_Account_No]  else (case when AccountType='Current' then t.[Payee_Joint_Account_No]  else '' end) end ) as [CurrentAccountNo]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[CurrentAmt]  else (case when AccountType='Current' then t.[CurrentAmt]  else 0 end) end ) as [CurrentAmt],t.DCSCount

 FROM(
    SELECT Head_Load_Rate,[Bank Advise No],[Bank Advise Date],[Account_Type],[From_Date],[Doc_No],[Fiscal_Name],[Date_Range],[accno2],[VLC_CODE_Uploader],[Payee_Joint_Name],[Bank_Code],[Branch_Name],[Bank_Code_Desc],[Payee_Joint_IFSC_Code],[Payee_Joint_Account_No],[Payable_Amount],[Saving_Amount],[Bank_Account_No_Saving],[Bank_Code_Saving],[Bank_Desc_Saving],Round(IsNull(Payable_Amount,0)-ISNULL(Head_Load_Rate,0)-IsNull(Saving_Amount,0),0) as CurrentAmt,
	Round(Case When IsNull(Payable_Amount,0)>0 Then (IsNull(Payable_Amount,0)-ISNULL(Head_Load_Rate,0)) Else IsNull(Saving_Amount,0) End,0) As Bill_Amt,(VLC_CODE_Uploader +' '+ Payee_Joint_Name) As DCS_Name,isnull(Bank_Code_Desc,Bank_Desc_Saving) as BankName,DCSCount from (
    select  SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
                            convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VENDOR_MASTER.accno2,
                            TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,"
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PickcompanyBankCodeFromPaymentProcess, clsFixedParameterCode.PickcompanyBankCodeFromPaymentProcess, Nothing)) = "1" Then
                        FinalQuery += " TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code_Saving,
                     TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc_Saving, TSPL_VENDOR_MASTER.Branch_Name, TSPL_VENDOR_MASTER.IFSC_Code, TSPL_VENDOR_MASTER.IFSCCode2, TSPL_PAYMENT_PROCESS_DETAIL.Bank_Account_No_Saving "
                    Else
                        FinalQuery += " TSPL_Vendor_MASTER.Bank_Code,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,
                     TSPL_VENDOR_MASTER.BankCode2 As Bank_Code_Saving, TSPL_VENDOR_MASTER.BankName2 As Bank_Desc_Saving, TSPL_VENDOR_MASTER.Branch_Name, TSPL_VENDOR_MASTER.IFSC_Code, TSPL_VENDOR_MASTER.IFSCCode2, TSPL_VENDOR_MASTER.AccNo2 as Bank_Account_No_Saving "
                    End If

                    FinalQuery += " ,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,
                            (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount,
							isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0) as Saving_Amount
							  from TSPL_PAYMENT_PROCESS_DETAIL 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                            --left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                           left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
                           where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) 
                           and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
                           group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code   left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'   )xx

Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
                            (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
                            TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
                            From TSPL_MILK_SRN_DETAIL 
                            Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
                            Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
                            left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
                            left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
                            Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
                             where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                             and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                             ) As final
							 where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xx.VLC_CODE_Uploader

) t
CROSS APPLY
(
    SELECT 'Current' AS AccountType, t.Bank_Code AS Bank, t.CurrentAmt AS Amount,1 as RI 
    UNION ALL
    SELECT 'Saving', t.Bank_Code_Saving, t.Saving_Amount,1 as RI
) ca
WHERE t.Bank_Code <> t.Bank_Code_Saving
   OR (t.Bank_Code = t.Bank_Code_Saving AND ca.AccountType = 'Current') )xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='ALW' where xxxfinal.DCSCount>1 

   union all

   Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,xxxfinal.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,
                              TSPL_COMPANY_MASTER.Regn_No,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (   --select * from bal

SELECT ca.RI,t.[Bank Advise No],t.[Bank Advise Date],t.VLC_CODE_Uploader,'ONE DAY STOP' AS BankCode,ca.Bank as [Branch_Name],ca.Bank as [BankName],t.DCS_Name

,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Bill_Amt] else (case when AccountType='Saving' then t.[Saving_Amount] else  t.[CurrentAmt] end) end )As [Bill_Amt]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Bank_Account_No_Saving] else (case when AccountType='Saving' then t.[Bank_Account_No_Saving] else '' end) end ) as [SavingAccountNo]
,ROUND(CASE WHEN t.Bank_Code = t.Bank_Code_Saving THEN t.Saving_Amount ELSE CASE WHEN AccountType = 'Saving' THEN t.Saving_Amount ELSE 0 END END , 0) AS [SavingAmt]


,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Payee_Joint_Account_No]  else (case when AccountType='Current' then t.[Payee_Joint_Account_No]  else '' end) end ) as [CurrentAccountNo]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[CurrentAmt]  else (case when AccountType='Current' then t.[CurrentAmt]  else 0 end) end ) as [CurrentAmt],t.DCSCount

 FROM(
    SELECT Head_Load_Rate,[Bank Advise No],[Bank Advise Date],[Account_Type],[From_Date],[Doc_No],[Fiscal_Name],[Date_Range],[accno2],[VLC_CODE_Uploader],[Payee_Joint_Name],[Bank_Code],[Branch_Name],[Bank_Code_Desc],[Payee_Joint_IFSC_Code],[Payee_Joint_Account_No],[Payable_Amount],[Saving_Amount],[Bank_Account_No_Saving],[Bank_Code_Saving],[Bank_Desc_Saving],Round(IsNull(Payable_Amount,0)-ISNULL(Head_Load_Rate,0)-IsNull(Saving_Amount,0),0) as CurrentAmt,
	Round(Case When IsNull(Payable_Amount,0)>0 Then (IsNull(Payable_Amount,0)-ISNULL(Head_Load_Rate,0)) Else IsNull(Saving_Amount,0) End,0) As Bill_Amt,(VLC_CODE_Uploader +' '+ Payee_Joint_Name) As DCS_Name,isnull(Bank_Code_Desc,Bank_Desc_Saving) as BankName,DCSCount,IFSC_Code,IFSCCode2 from (
    select  SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
                            convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VENDOR_MASTER.accno2,
                            TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,"
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PickcompanyBankCodeFromPaymentProcess, clsFixedParameterCode.PickcompanyBankCodeFromPaymentProcess, Nothing)) = "1" Then
                        FinalQuery += " TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code_Saving,
                     TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc_Saving, TSPL_VENDOR_MASTER.Branch_Name, TSPL_VENDOR_MASTER.IFSC_Code, TSPL_VENDOR_MASTER.IFSCCode2, TSPL_PAYMENT_PROCESS_DETAIL.Bank_Account_No_Saving "
                    Else
                        FinalQuery += " TSPL_Vendor_MASTER.Bank_Code,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,
                     TSPL_VENDOR_MASTER.BankCode2 As Bank_Code_Saving, TSPL_VENDOR_MASTER.BankName2 As Bank_Desc_Saving, TSPL_VENDOR_MASTER.Branch_Name, TSPL_VENDOR_MASTER.IFSC_Code, TSPL_VENDOR_MASTER.IFSCCode2, TSPL_VENDOR_MASTER.AccNo2 as Bank_Account_No_Saving "
                    End If

                    FinalQuery += "  ,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,
                            (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount,
							isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0) as Saving_Amount
						    from TSPL_PAYMENT_PROCESS_DETAIL 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                            --left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                           left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
                           where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) 
                           and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
                           group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code   left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'   )xx

Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
                            (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
                            TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
                            From TSPL_MILK_SRN_DETAIL 
                            Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
                            Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
                            left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
                            left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
                            Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
                             where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                             and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                             ) As final
							 where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xx.VLC_CODE_Uploader

) t
CROSS APPLY
(
    SELECT 'Current' AS AccountType, t.Bank_Code AS Bank, t.CurrentAmt AS Amount,1 as RI 
    UNION ALL
    SELECT 'Saving', t.Bank_Code_Saving, t.Saving_Amount,1 as RI
) ca
WHERE t.Bank_Code <> t.Bank_Code_Saving
   OR (t.Bank_Code = t.Bank_Code_Saving AND ca.AccountType = 'Current') )xxxfinal Left Outer Join TSPL_Company_Master On TSPL_Company_Master.comp_code1='ALW' where xxxfinal.DCSCount=1 )final group by final.bankcode"


                Else
                    FinalQuery = ""
                    FinalQuery = " select ROW_NUMBER() Over (Order By (Select 1)) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,BankCode,BankName,TotalBillAmt,TotalSavingAmt,CurrentAmt,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No
                                            ,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2 from (SELECT max(SNo)SNo,BankCode,max(BankName)BankName,sum(TotalBillAmt)TotalBillAmt,sum(TotalSavingAmt)TotalSavingAmt,sum(CurrentAmt)CurrentAmt FROM (Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,Final.*   from ( Select xxx.BankCode,Max(xxx.BankName)BankName,Sum(xxx.Bill_Amt)TotalBillAmt,Sum(xxx.SavingAmt)TotalSavingAmt,Sum(xxx.CurrentAmt)CurrentAmt,max(xxx.DCSCount)DCSCount from (Select * from (Select xxxFinal.[Bank Advise No],xxxFinal.[Bank Advise Date],xxxfinal.VLC_CODE_Uploader,	xxxfinal.BankCode,	xxxfinal.BankName,	xxxfinal.DCS_Name,	Round(xxxfinal.Bill_Amt,0)Bill_Amt,	xxxfinal.SavingAccountNo,	Round(xxxfinal.SavingAmt,0)SavingAmt,	xxxfinal.CurrentAccountNo,	Case When xxxfinal.CurrentAmt>0 Then Round(xxxfinal.CurrentAmt,0) Else 0 End As CurrentAmt,xxxDCSCount.DCSCount 
                                            from 
                                            (Select [Bank Advise No],[Bank Advise Date],xxxSaving.VLC_CODE_Uploader,IsNull(xxxSaving.Bank_Code,xxxCurrent.Bank_Code) As BankCode,IsNull(xxxsaving.Bank_Code_Desc,xxxCurrent.Bank_Code_Desc) As BankName,(xxxSaving.VLC_CODE_Uploader +' '+ xxxSaving.Payee_Joint_Name) As DCS_Name,
                							Case When IsNull(xxxCurrent.Payable_Amount,0)>0 Then (IsNull(xxxCurrent.Payable_Amount,0)-ISNULL(xxxCurrent.Head_Load_Rate,0)) Else IsNull(xxxSaving.Payable_Amount,0) End As Bill_Amt, 
                                            Case When xxxSaving.Account_Type='Sav' Then xxxSaving.Payee_Joint_Account_No Else '' End As SavingAccountNo,
                							IsNull(xxxSaving.Payable_Amount,0) As SavingAmt,
                                            Case When xxxCurrent.Account_Type='Cur' Then xxxCurrent.Payee_Joint_Account_No Else '' End As CurrentAccountNo,
                							IsNull(xxxCurrent.Payable_Amount,0)-IsNull(xxxSaving.Payable_Amount,0)-ISNULL(xxxCurrent.Head_Load_Rate,0) As CurrentAmt
                                            from (select Max(AccountType)As Account_Type, x.GRPColumn,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,
                                            max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,
                                            max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount 
                                            from
                                            (select  Case When TSPL_VENDOR_MASTER.Account_Type='sav' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='sav' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As AccountType ,'' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code+(case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')='Self' then (TSPL_VENDOR_MASTER.IFSC_Code)   else (TSPL_VENDOR_MASTER.Joint_IFSC_Code)   end)  as GRPColumn,
                                            TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, (TSPL_VENDOR_MASTER.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
                                              TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code_Desc,
                                              case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')= 'Self' then (TSPL_VENDOR_MASTER.IFSCCode2) else (TSPL_VENDOR_MASTER.Joint_IFSC_Code) end as Payee_Joint_IFSC_Code, 
                                              case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')= 'Self' then (TSPL_VENDOR_MASTER.AccNo2) else (TSPL_VENDOR_MASTER.Joint_Account_No) end as Payee_Joint_Account_No, 
                                              TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount   
                                              from TSPL_PAYMENT_PROCESS_SAVING 
                                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
                                            left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No                                                       
                                            left outer join (select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_COMPULSORY.VSP_CODE
                                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)
                                            and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank
                                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                                            and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0 
                                            ) x group by x.GRPColumn,x.VLC_CODE_Uploader,x.Payee_Joint_Account_No )xxxSaving
                                            Left Outer Join
                                            (select MAX([Bank Advise No])[Bank Advise No],MAX([Bank Advise Date])[Bank Advise Date],Max(Account_Type)As Account_Type,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount,sum(Head_Load_Rate)Head_Load_Rate from
                                            (select   SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
                                            convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, 
                                            TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,
                                            case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,
                                            (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount 
                                            from TSPL_PAYMENT_PROCESS_DETAIL 
                                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
                                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                           left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
                                           where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) 
                                            and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
                                           group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code
                                            left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  ) x group by x.VLC_CODE_Uploader,x.Payee_Joint_Account_No ) xxxCurrent On xxxCurrent.VLC_CODE_Uploader=xxxSaving.VLC_CODE_Uploader) xxxFinal
                                            Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
                                            (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
                                            TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
                                            From TSPL_MILK_SRN_DETAIL 
                                            Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
                                            Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
                                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
                                            left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
                                            left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
                                            Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
                                             where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                            and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                                             ) As final
                                             where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY  GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xxxFinal.VLC_CODE_Uploader  )xxxfinal )xxx
                                            where xxx.DCSCount>1 Group By xxx.BankCode)Final 

                							union all

                							Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,xxxfinal.*
                							from (Select 'ONE DAY STOP' AS BankCode,

                							xxxfinal.BankName,	Round(xxxfinal.Bill_Amt,0)Bill_Amt,	Round(xxxfinal.SavingAmt,0)SavingAmt,		Case When xxxfinal.CurrentAmt>0 Then Round(xxxfinal.CurrentAmt,0) Else 0 End As CurrentAmt,xxxDCSCount.DCSCount 
                                            from 
                                            (Select [Bank Advise No],[Bank Advise Date],xxxSaving.VLC_CODE_Uploader,IsNull(xxxSaving.Bank_Code,xxxCurrent.Bank_Code) As BankCode,IsNull(xxxsaving.Bank_Code_Desc,xxxCurrent.Bank_Code_Desc) As BankName,(xxxSaving.VLC_CODE_Uploader +' '+ xxxSaving.Payee_Joint_Name) As DCS_Name,
                							Case When IsNull(xxxCurrent.Payable_Amount,0)>0 Then (IsNull(xxxCurrent.Payable_Amount,0)-ISNULL(xxxCurrent.Head_Load_Rate,0)) Else IsNull(xxxSaving.Payable_Amount,0) End As Bill_Amt, 
                                            Case When xxxSaving.Account_Type='Sav' Then xxxSaving.Payee_Joint_Account_No Else '' End As SavingAccountNo,
                							IsNull(xxxSaving.Payable_Amount,0) As SavingAmt,
                                            Case When xxxCurrent.Account_Type='Cur' Then xxxCurrent.Payee_Joint_Account_No Else '' End As CurrentAccountNo,
                							IsNull(xxxCurrent.Payable_Amount,0)-IsNull(xxxSaving.Payable_Amount,0)-ISNULL(xxxCurrent.Head_Load_Rate,0) As CurrentAmt
                                            from (select Max(AccountType)As Account_Type, x.GRPColumn,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,
                                            max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,
                                            max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount 
                                            from
                                            (select  Case When TSPL_VENDOR_MASTER.Account_Type='sav' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='sav' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As AccountType ,'' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code+(case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')='Self' then (TSPL_VENDOR_MASTER.IFSC_Code)   else (TSPL_VENDOR_MASTER.Joint_IFSC_Code)   end)  as GRPColumn,
                                            TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, (TSPL_VENDOR_MASTER.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
                                              TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code_Desc,
                                              case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')= 'Self' then (TSPL_VENDOR_MASTER.IFSCCode2) else (TSPL_VENDOR_MASTER.Joint_IFSC_Code) end as Payee_Joint_IFSC_Code, 
                                              case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')= 'Self' then (TSPL_VENDOR_MASTER.AccNo2) else (TSPL_VENDOR_MASTER.Joint_Account_No) end as Payee_Joint_Account_No, 
                                              TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount   
                                              from TSPL_PAYMENT_PROCESS_SAVING 
                                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
                                            left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No                                                       
                                            left outer join (select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_COMPULSORY.VSP_CODE
                                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank
                                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                                            and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0 
                                            ) x group by x.GRPColumn,x.VLC_CODE_Uploader,x.Payee_Joint_Account_No )xxxSaving
                                            Left Outer Join
                                            (select MAX([Bank Advise No])[Bank Advise No],MAX([Bank Advise Date])[Bank Advise Date],Max(Account_Type)As Account_Type,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount,sum(Head_Load_Rate)Head_Load_Rate from
                                            (select  SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
                                            convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, 
                                            TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,
                                            case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,
                                            (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount 
                                            from TSPL_PAYMENT_PROCESS_DETAIL 
                                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
                                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                           left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
                                           where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) 
                                           and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
                                           group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code
                                            left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                           and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  ) x group by x.VLC_CODE_Uploader,x.Payee_Joint_Account_No ) xxxCurrent On xxxCurrent.VLC_CODE_Uploader=xxxSaving.VLC_CODE_Uploader) xxxFinal
                                            Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
                                            (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
                                            TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
                                            From TSPL_MILK_SRN_DETAIL 
                                            Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
                                            Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
                                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
                                            left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
                                            left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
                                            Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
                                             where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                             ) As final
                                             where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY  GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xxxFinal.VLC_CODE_Uploader  )xxxfinal 
                							 where xxxfinal.DCSCount=1 )XXY Group by BankCode)zz
                							 Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1= '" + objCommonVar.CurrComp_Code1 + "' "
                    'FinalQuery = "Select ROW_NUMBER() Over (Order By (Select 1)) As SNo,('" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MM/yyyy") + "'+' to '+'" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MM/yyyy") + "') As DateRange,Final.*,TSPL_COMPANY_MASTER.Comp_Name,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No
                    '            ,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_COMPANY_MASTER.Pincode,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2   from ( Select xxx.BankCode,Max(xxx.BankName)BankName,Sum(xxx.Bill_Amt)TotalBillAmt,Sum(xxx.SavingAmt)TotalSavingAmt,Sum(xxx.CurrentAmt)CurrentAmt from (Select * from (" + GetSavingCurrent() + ")xxxfinal )xxx
                    '            where xxx.DCSCount>1 Group By xxx.BankCode)Final Left Outer Join TSPL_COMPANY_MASTER On TSPL_COMPANY_MASTER.Comp_Code1='" + objCommonVar.CurrComp_Code1 + "'"
                End If
                'FinalQuery = ""


                dt3 = clsDBFuncationality.GetDataTable(FinalQuery)
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                If rbtnBankAdvice.IsChecked Then
                    Dim query As String = Nothing
                    'Dim dt7 As DataTable = Nothing
                    query = "  select Bank_Code,max(Branch_Name)Branch_Name, sum(Payable_Amount) as Payable_Amount
                                from (select TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name, "
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        query += " Round((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)-isnull(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0)),0) as Payable_Amount  "
                    Else
                        query += " (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)-isnull(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0))  as Payable_Amount  "
                    End If
                    query += " from TSPL_PAYMENT_PROCESS_DETAIL 
                                left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                                left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP'
                                left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
                                left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                                left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current  left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                                left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                                left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected 
                                where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	
                                TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'   And (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0
                                )xxx group by Bank_Code   "

                    dt7 = clsDBFuncationality.GetDataTable(query)
                End If
                If rbtnSaving.IsChecked Then
                    Dim qry As String = Nothing
                    qry = "  select Bank_Code,sum(Payable_Amount) as Payable_Amount
                                 from (select TSPL_Vendor_MASTER.BankCode2 as Bank_Code, "
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        qry += " Round(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as Payable_Amount "
                    Else
                        qry += " TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount "
                    End If
                    qry += " from TSPL_PAYMENT_PROCESS_SAVING 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
                            left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
                            left outer join TSPL_MP_MASTER mp on mp.MP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
                            Left outer join tspl_vendor_master Mp_V on mp_V.Vendor_Code=mp.MP_Code  
                            left outer join ( select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                            left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP'
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
                            left outer join TSPL_BANK_MASTER as TSPL_BANK_MASTER2 on TSPL_BANK_MASTER2.BANK_CODE=TSPL_Vendor_MASTER.BankCode2 
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Bank_Code
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>=  '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <=  '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0  
                               )xxx group by Bank_Code "

                    dt8 = clsDBFuncationality.GetDataTable(qry)
                End If
            End If
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterTemplate.FilterDescriptors.Clear()
            Gv1.DataSource = Nothing
            Gv1.Rows.Clear()
            Gv1.Columns.Clear()
            Gv1.GroupDescriptors.Clear()
            Gv1.MasterTemplate.SummaryRowsBottom.Clear()
            Gv1.MasterView.Refresh()

            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found/Posted to Display", Me.Text)
                Exit Sub
            Else
                If clsCommon.myLen(BankAdviseDocNo) <= 0 OrElse clsCommon.myLen(LocalBankAdviseDocNo) > 0 Then
                    Gv1.DataSource = dt
                    RadPageView1.SelectedPage = RadPageViewPage2
                    SetGridFormat()
                    EnableDisaableControls(False)
                End If
            End If



            If isPrint Then
                If rbtnBankAdvice.IsChecked Then
                    ''Note IF You do any changes than change in function clsBankAdvise.CreateEmailContent(ByVal strDateRange As String, trans As SqlTransaction)
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdvice", "Bank Advice")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewJPR", "Bank Advice")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "UDP") = CompairStringResult.Equal AndAlso VendorBankAdviceForSWM = True Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewSWM", "Bank Advice")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, dt7, "crptBankAdviceNewTNK", "Bank Advice", "crptBankAdviceNewTNKSubReport.rpt")
                    ElseIf clsCommon.myLen(vSeprateBankForDCSBankAdvice) > 0 Then
                        'SWM SAME PRINT

                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewSWMNEW1", "Bank Advice")
                        'BANKDVIEPRINT FORMATE
                        'frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceSWM11", "Bank Advice")

                    Else
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNew", "Bank Advice")
                        frmCRV = Nothing
                    End If
                ElseIf rbtnBankWiseSummary.IsChecked OrElse rbtnSavingSummary.IsChecked Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankWiseSummary", "Bank Wise Summary")
                    Else
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankWiseSummaryNew", "Bank Wise Summary")
                    End If

                    frmCRV = Nothing
                ElseIf rbtnBothSavCur.IsChecked Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    'frmCRV.funsubreportWithdt(CrystalReportFolder.MilkProcurement, dt, "", "crptBankAdviseReportALW", "Bank Advice Report", "", "crptBankAdviseSubReportALW.rpt", "crptBankAdviseBankWiseSubReportALW.rpt", "", Nothing, Nothing, Nothing)
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "ALW") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt1, dt2, "crptBankAdviseReportALW", Nothing, "crptBankAdviseSubReportALW.rpt", "crptBankAdviseBankWiseSubReportALW.rpt", dt3, Nothing, Nothing, Nothing, Nothing)
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JHL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BRN") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt1, dt2, "crptBankAdviseReportNew", Nothing, "crptBankAdviseSubReportALW.rpt", "crptBankAdviseBankWiseSubReportALW.rpt", dt3, Nothing, Nothing, Nothing, Nothing)

                    Else
                        clsCommon.MyMessageBoxShow(Me, "Data Not Found", Me.Text)
                    End If
                    frmCRV = Nothing
                ElseIf rbtnSaving.IsChecked Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "TNK") = CompairStringResult.Equal Then
                        frmCRV.funsubreportWithdt(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, dt8, "crptBankAdviceNewSavingTNK", "Bank Advice Saving", "crptBankAdviceNewSavingTNKSubReport.rpt")
                    ElseIf clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewSavingJPR", "Bank Advice Saving")
                    Else
                        frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewSaving", "Bank Advice Saving")
                        frmCRV = Nothing
                    End If
                    'frmCRV.funreport(CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewSaving", "Bank Advice Saving")
                    'frmCRV = Nothing
                End If
            End If
            'ReStoreGridLayout()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Function GetSavingCurrentALW(ByVal BankAdviseDocNo As String) As String
        Dim qry As String = "   --select * from bal

SELECT ca.RI,t.[Bank Advise No],t.[Bank Advise Date],t.VLC_CODE_Uploader,ca.Bank as BankCode,ca.Bank as [Branch_Name],ca.Bank as [BankName],t.DCS_Name

,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Bill_Amt] else (case when AccountType='Saving' then t.[Saving_Amount] else  t.[CurrentAmt] end) end )As [Bill_Amt]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Bank_Account_No_Saving] else (case when AccountType='Saving' then t.[Bank_Account_No_Saving] else '' end) end ) as [SavingAccountNo]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.IFSCCode2 else (case when AccountType='Saving' then t.IFSCCode2 else '' end) end ) as [SavingIFSCCode]
,ROUND(CASE WHEN t.Bank_Code = t.Bank_Code_Saving THEN t.Saving_Amount ELSE CASE WHEN AccountType = 'Saving' THEN t.Saving_Amount ELSE 0 END END , 0) AS [SavingAmt]


,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[Payee_Joint_Account_No]  else (case when AccountType='Current' then t.[Payee_Joint_Account_No]  else '' end) end ) as [CurrentAccountNo]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.IFSC_Code  else (case when AccountType='Current' then t.IFSC_Code  else '' end) end ) as [CurrentIFSCCode]
,(case when t.Bank_Code=t.Bank_Code_Saving then  t.[CurrentAmt]  else (case when AccountType='Current' then t.[CurrentAmt]  else 0 end) end ) as [CurrentAmt],t.DCSCount

 FROM(
    SELECT Head_Load_Rate,[Bank Advise No],[Bank Advise Date],[Account_Type],[From_Date],[Doc_No],[Fiscal_Name],[Date_Range],[accno2],[VLC_CODE_Uploader],[Payee_Joint_Name],[Bank_Code],[Branch_Name],[Bank_Code_Desc],[Payee_Joint_IFSC_Code],[Payee_Joint_Account_No],[Payable_Amount],[Saving_Amount],[Bank_Account_No_Saving],[Bank_Code_Saving],[Bank_Desc_Saving],Round(IsNull(Payable_Amount,0)-ISNULL(Head_Load_Rate,0)-IsNull(Saving_Amount,0),0) as CurrentAmt,
	Round(Case When IsNull(Payable_Amount,0)>0 Then (IsNull(Payable_Amount,0)-ISNULL(Head_Load_Rate,0)) Else IsNull(Saving_Amount,0) End,0) As Bill_Amt,(VLC_CODE_Uploader +' '+ Payee_Joint_Name) As DCS_Name,isnull(Bank_Code_Desc,Bank_Desc_Saving) as BankName,DCSCount,IFSC_Code,IFSCCode2 from (
    select  SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],
    Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,
    TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
    convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VENDOR_MASTER.accno2,
    TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name, "
        If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.PickcompanyBankCodeFromPaymentProcess, clsFixedParameterCode.PickcompanyBankCodeFromPaymentProcess, Nothing)) = "1" Then
            qry += " TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code_Saving,
                     TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc_Saving,TSPL_VENDOR_MASTER.Branch_Name,TSPL_VENDOR_MASTER.IFSC_Code,TSPL_VENDOR_MASTER.IFSCCode2,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Account_No_Saving "
        Else
            qry += " TSPL_Vendor_MASTER.Bank_Code,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,
                     TSPL_VENDOR_MASTER.BankCode2 as Bank_Code_Saving,TSPL_VENDOR_MASTER.BankName2 as Bank_Desc_Saving,TSPL_VENDOR_MASTER.Branch_Name,TSPL_VENDOR_MASTER.IFSC_Code,TSPL_VENDOR_MASTER.IFSCCode2,TSPL_VENDOR_MASTER.AccNo2 as Bank_Account_No_Saving "
        End If
        qry += " ,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,"
        If dtBankAdvise IsNot Nothing AndAlso dtBankAdvise.Rows.Count > 0 Then
            qry += " (isnull(TSPL_BANK_ADVISE_DETAIL.Partial_Amt,0)+isnull(TSPL_BANK_ADVISE_DETAIL.Saving_Partial_Amt,0))  as Payable_Amount,isnull(TSPL_BANK_ADVISE_DETAIL.Saving_Partial_Amt,0) as Saving_Amount "
        Else
            qry += "   (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount,	isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0) as Saving_Amount "
        End If
        qry += "	 from TSPL_PAYMENT_PROCESS_DETAIL 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                            --left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                           left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
                           where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) 
                           and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
                           group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code   left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No  
                           left outer join TSPL_BANK_ADVISE_DETAIL On TSPL_BANK_ADVISE_DETAIL.Payment_Process_PP_Detail_No=TSPL_PAYMENT_PROCESS_DETAIL.PP_Detail_No
						   and TSPL_BANK_ADVISE_DETAIL.Document_No=TSPL_BANK_ADVISE.Document_No
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' "
        If clsCommon.myLen(BankAdviseDocNo) > 0 Then
            qry += " and TSPL_BANK_ADVISE.Document_No='" + BankAdviseDocNo + "' "
        End If

        qry += "   )xx

Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
                            (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
                            TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
                            From TSPL_MILK_SRN_DETAIL 
                            Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
                            Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
                            left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
                            left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
                            Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
                             where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                             and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
                             ) As final
							 where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xx.VLC_CODE_Uploader

) t
CROSS APPLY
(
    SELECT 'Current' AS AccountType, t.Bank_Code AS Bank, t.CurrentAmt AS Amount,1 as RI 
    UNION ALL
    SELECT 'Saving', t.Bank_Code_Saving, t.Saving_Amount,1 as RI
) ca
WHERE t.Bank_Code <> t.Bank_Code_Saving
   OR (t.Bank_Code = t.Bank_Code_Saving AND ca.AccountType = 'Current') "
        Return qry
    End Function
    Private Function GetSavingCurrent() As String
        Dim Qry As String = " Select xxxfinal.[Bank Advise No],xxxfinal.[Bank Advise Date],xxxfinal.VLC_CODE_Uploader,isnull(xxxfinal.Bank_Code,xxxfinal.Bank_Code_Saving) as BankCode,xxxfinal.Branch_Name,isnull(xxxfinal.Bank_Code_Desc,xxxfinal.Bank_Desc_Saving) as BankName,(xxxfinal.VLC_CODE_Uploader +' '+ xxxfinal.Payee_Joint_Name) As DCS_Name,
Round(Case When IsNull(xxxfinal.Payable_Amount,0)>0 Then (IsNull(xxxfinal.Payable_Amount,0)-ISNULL(xxxfinal.Head_Load_Rate,0)) Else IsNull(xxxfinal.Saving_Amount,0) End,0) As Bill_Amt,
xxxfinal.Bank_Account_No_Saving as SavingAccountNo,IsNull(xxxfinal.Saving_Amount,0) as SavingAmt,xxxfinal.Payee_Joint_Account_No as CurrentAccountNo, "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " case when AccNo2 is null or AccNo2='' then Round(IsNull(xxxfinal.Payable_Amount,0)-ISNULL(xxxfinal.Head_Load_Rate,0)-IsNull(xxxfinal.Saving_Amount,0)+isnull(xxxfinal.Amt1,0),0) else  
                     Round(IsNull(xxxfinal.Payable_Amount,0)-ISNULL(xxxfinal.Head_Load_Rate,0)-IsNull(xxxfinal.Saving_Amount,0),0) end as CurrentAmt "
        Else
            Qry += " Round(IsNull(xxxfinal.Payable_Amount,0)-ISNULL(xxxfinal.Head_Load_Rate,0)-IsNull(xxxfinal.Saving_Amount,0),0) as CurrentAmt "
        End If

        Qry += " ,DCSCount
from (select MAX([Bank Advise No])[Bank Advise No],MAX([Bank Advise Date])[Bank Advise Date],Max(Account_Type)As Account_Type,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(Head_Load_Rate)Head_Load_Rate,
MAX(Bank_Account_No_Saving)Bank_Account_No_Saving,max(Bank_Code_Saving)Bank_Code_Saving,max(Bank_Desc_Saving)Bank_Desc_Saving, "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " Max(x.Payable_Amount)Payable_Amount, max(Saving_Amount)Saving_Amount,sum(Amount)Amt1,max(accno2)AccNo2 "
        Else
            Qry += " sum(x.Payable_Amount)Payable_Amount, sum(Saving_Amount)Saving_Amount "
        End If

        Qry += " from
                            (select  SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
                            convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VENDOR_MASTER.accno2,
                            TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,
                            case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,
                            (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount,
							isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0) as Saving_Amount,
							TSPL_PAYMENT_PROCESS_DETAIL.Bank_Account_No_Saving,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Code_Saving,TSPL_PAYMENT_PROCESS_DETAIL.Bank_Desc_Saving "
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " ,TabDCSdrcr.Amount "
        End If
        Qry += " from TSPL_PAYMENT_PROCESS_DETAIL 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                            --left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                           left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
                           where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) 
                           and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
                           group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code "

        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "KTA") = CompairStringResult.Equal Then
            Qry += " left outer join (Select InvoiceNo,DCSDescription,case when DCSDescription LIKE 'Purchase%' THEN Amount end as Amount from(
				                    select TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.InvoiceNo,TSPL_DCS_ADDITION_DEDUCTION.Description As DCSDescription,TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.Amount from TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE
				                    left outer join TSPL_DCS_ADDITION_DEDUCTION on TSPL_DCS_ADDITION_DEDUCTION.Code = TSPL_MILK_PURCHASE_INVOICE_DCS_ADD_DED_DONT_GENERATE_DR_CR_NOTE.DCS_Addition_Deduction
				                    )x) As TabDCSdrcr on TabDCSdrcr.InvoiceNo = TSPL_PAYMENT_PROCESS_DETAIL.Milk_Purchase_Invoice_No "
        End If

        Qry += "  left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                            and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
							) x group by x.VLC_CODE_Uploader,x.Payee_Joint_Account_No ) xxxfinal

							Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
                            (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
                            TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
                            From TSPL_MILK_SRN_DETAIL 
                            Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
                            Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
                            Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
                            left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
                            left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
                            Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
                             where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                             and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                             ) As final
							 where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xxxFinal.VLC_CODE_Uploader "
        '     "Select xxxFinal.[Bank Advise No],xxxFinal.[Bank Advise Date],xxxfinal.VLC_CODE_Uploader,	xxxfinal.BankCode,	xxxfinal.BankName,	xxxfinal.DCS_Name,	Round(xxxfinal.Bill_Amt,0)Bill_Amt,	xxxfinal.SavingAccountNo,	Round(xxxfinal.SavingAmt,0)SavingAmt,	xxxfinal.CurrentAccountNo,	Case When xxxfinal.CurrentAmt>0 Then Round(xxxfinal.CurrentAmt,0) Else 0 End As CurrentAmt,xxxDCSCount.DCSCount 
        '                     from 
        '                     (Select [Bank Advise No],[Bank Advise Date],xxxSaving.VLC_CODE_Uploader,IsNull(xxxSaving.Bank_Code,xxxCurrent.Bank_Code) As BankCode,IsNull(xxxsaving.Bank_Code_Desc,xxxCurrent.Bank_Code_Desc) As BankName,(xxxSaving.VLC_CODE_Uploader +' '+ xxxSaving.Payee_Joint_Name) As DCS_Name,
        'Case When IsNull(xxxCurrent.Payable_Amount,0)>0 Then (IsNull(xxxCurrent.Payable_Amount,0)-ISNULL(xxxCurrent.Head_Load_Rate,0)) Else IsNull(xxxSaving.Payable_Amount,0) End As Bill_Amt, 
        '                     Case When xxxSaving.Account_Type='Sav' Then xxxSaving.Payee_Joint_Account_No Else '' End As SavingAccountNo,
        'IsNull(xxxSaving.Payable_Amount,0) As SavingAmt,
        '                     Case When xxxCurrent.Account_Type='Cur' Then xxxCurrent.Payee_Joint_Account_No Else '' End As CurrentAccountNo,
        'IsNull(xxxCurrent.Payable_Amount,0)-IsNull(xxxSaving.Payable_Amount,0)-ISNULL(xxxCurrent.Head_Load_Rate,0) As CurrentAmt
        '                     from (select Max(AccountType)As Account_Type, x.GRPColumn,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,
        '                     max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,
        '                     max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount 
        '                     from
        '                     (select  Case When TSPL_VENDOR_MASTER.Account_Type='sav' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='sav' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As AccountType ,'' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code+(case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')='Self' then (TSPL_VENDOR_MASTER.IFSC_Code)   else (TSPL_VENDOR_MASTER.Joint_IFSC_Code)   end)  as GRPColumn,
        '                     TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, (TSPL_VENDOR_MASTER.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
        '                       TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code_Desc,
        '                       case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')= 'Self' then (TSPL_VENDOR_MASTER.IFSCCode2) else (TSPL_VENDOR_MASTER.Joint_IFSC_Code) end as Payee_Joint_IFSC_Code, 
        '                       case when isnull(TSPL_VENDOR_MASTER.vsp_payment,'')= 'Self' then (TSPL_VENDOR_MASTER.AccNo2) else (TSPL_VENDOR_MASTER.Joint_Account_No) end as Payee_Joint_Account_No, 
        '                       TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount   
        '                       from TSPL_PAYMENT_PROCESS_SAVING 
        '                     left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
        '                     left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No                                                       
        '                     left outer join (select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
        '                     left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_COMPULSORY.VSP_CODE
        '                     left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
        '                     left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
        '                     left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank
        '                     left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
        '                     where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'
        '                     and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0 --and TSPL_Vendor_MASTER.Bank_Code='504' 
        '                     ) x group by x.GRPColumn,x.VLC_CODE_Uploader,x.Payee_Joint_Account_No )xxxSaving
        '                     Left Outer Join
        '                     (select MAX([Bank Advise No])[Bank Advise No],MAX([Bank Advise Date])[Bank Advise Date],Max(Account_Type)As Account_Type,max(x.From_Date)From_Date,max(x.Doc_No)Doc_No,max(x.Fiscal_Name)Fiscal_Name,max(x.Date_Range)Date_Range,x.VLC_CODE_Uploader,max(x.Payee_Joint_Name)Payee_Joint_Name,max(x.Bank_Code)Bank_Code,max(x.Branch_Name)Branch_Name,max(x.Bank_Code_Desc)Bank_Code_Desc,max(x.Payee_Joint_IFSC_Code)Payee_Joint_IFSC_Code,x.Payee_Joint_Account_No,sum(x.Payable_Amount)Payable_Amount,sum(Head_Load_Rate)Head_Load_Rate from
        '                     (select  SavingDetail.Head_Load_Rate,TSPL_BANK_ADVISE.Document_No As [Bank Advise No],Convert(Varchar(10),TSPL_BANK_ADVISE.Document_Date,103) As [Bank Advise Date],Case When TSPL_VENDOR_MASTER.Account_Type='cur' Then TSPL_VENDOR_MASTER.Account_Type Else Case When TSPL_VENDOR_MASTER.AccountType2='cur' Then TSPL_VENDOR_MASTER.AccountType2 else '' End  End As Account_Type,TSPL_PAYMENT_PROCESS_HEAD.From_Date,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name,
        '                     convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, 
        '                     TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,
        '                     case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,
        '                     (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)+isnull(TSPL_PAYMENT_PROCESS_DETAIL.Saving_Amount,0))  as Payable_Amount 
        '                     from TSPL_PAYMENT_PROCESS_DETAIL 
        '                     left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
        '                     left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE								
        '                     left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
        '                     left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
        '                     left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
        '                     --left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
        '                    left outer join( select Vendor_Code,Sum(Amount)Head_Load_Rate from TSPL_TRANSFER_TO_SAVING_DETAIL left outer join TSPL_TRANSFER_TO_SAVING on TSPL_TRANSFER_TO_SAVING.Document_No = TSPL_TRANSFER_TO_SAVING_DETAIL.Document_No
        '                    where convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103)>=convert(date,('" + fromDate.Value + "'),103) and convert(date,TSPL_TRANSFER_TO_SAVING.Document_Date,103) <=convert(date,('" + ToDate.Value + "'),103)
        '                    group by Vendor_Code )SavingDetail on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = SavingDetail.Vendor_Code
        '                     left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
        '                     left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
        '                     where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' "
        ' ' and (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0 
        ' Qry += " ) x group by x.VLC_CODE_Uploader,x.Payee_Joint_Account_No ) xxxCurrent On xxxCurrent.VLC_CODE_Uploader=xxxSaving.VLC_CODE_Uploader) xxxFinal
        '                     Inner Join(SELECT YYY.[Vlc Uploader Code],YYY.[VLC Name],COUNT(YYY.[Vlc Uploader Code]) As DCSCount FROM (Select final.[Doc Date] ,(final.[Vlc Uploader Code])[Vlc Uploader Code] ,final.[VLC Name] From 
        '                     (Select Convert(varchar,TSPL_MILK_SRN_HEAD.DOC_DATE,103) As [Doc Date],TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader As [Vlc Uploader Code], 
        '                     TSPL_VLC_MASTER_HEAD.VLC_Name As [VLC Name]
        '                     From TSPL_MILK_SRN_DETAIL 
        '                     Left Outer Join TSPL_MILK_SRN_HEAD On TSPL_MILK_SRN_HEAD.DOC_CODE = TSPL_MILK_SRN_DETAIL.DOC_CODE 
        '                     Left Outer Join TSPL_VLC_MASTER_HEAD On TSPL_VLC_MASTER_HEAD.VLC_Code = TSPL_MILK_SRN_HEAD.VLC_CODE
        '                     Left Outer Join TSPL_VENDOR_MASTER On TSPL_VENDOR_MASTER.Vendor_Code = TSPL_MILK_SRN_HEAD.VSP_CODE
        '                     left outer join TSPL_VENDOR_GROUP on TSPL_VENDOR_MASTER.Vendor_Group_Code = TSPL_VENDOR_GROUP.Ven_Group_Code 
        '                     left outer join TSPL_BULK_ROUTE_MASTER On TSPL_BULK_ROUTE_MASTER.ROUTE_NO=TSPL_MILK_SRN_HEAD.ROUTE_CODE 
        '                     Left Outer Join TSPL_MCC_ROUTE_MASTER On TSPL_MCC_ROUTE_MASTER.Route_Code = TSPL_MILK_SRN_HEAD.ROUTE_CODE
        '                      where 2 = 2  and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and Cast(TSPL_MILK_SRN_HEAD.DOC_DATE as Date) <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
        '                      ) As final
        '                      where 2=2  GROUP BY FINAL.[Doc Date],FINAL.[VLC Name],FINAL.[Vlc Uploader Code] )YYY  GROUP BY [VLC Name],[Vlc Uploader Code])xxxDCSCount On xxxDCSCount.[Vlc Uploader Code]=xxxFinal.VLC_CODE_Uploader  "
        Return Qry
    End Function

    Sub SetGridFormat()
        Gv1.ShowGroupPanel = False
        Gv1.ShowRowHeaderColumn = False
        Gv1.AllowAddNewRow = False
        Gv1.AllowDeleteRow = False
        Gv1.EnableFiltering = True
        Gv1.ShowFilteringRow = True
        Gv1.MasterTemplate.SummaryRowsBottom.Clear()

        For ii As Integer = 0 To Gv1.Columns.Count - 1
            Gv1.Columns(ii).ReadOnly = True
            Gv1.Columns(ii).BestFit()
        Next
        If rbtnSaving.IsChecked Then
            Gv1.Columns("CycleRange").HeaderText = "Cycle Range"
            Gv1.Columns("CycleRange").IsVisible = False

            Gv1.Columns("GRPColumn").HeaderText = "Group Range"
            Gv1.Columns("GRPColumn").IsVisible = False

            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Company Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("CompPhone").HeaderText = "Phone No"
            Gv1.Columns("CompPhone").IsVisible = False

            Gv1.Columns("Regn_No").HeaderText = "Regn No"
            Gv1.Columns("Regn_No").IsVisible = False

            Gv1.Columns("MCC_NAME").HeaderText = "Area"
            Gv1.Columns("MCC_NAME").IsVisible = False

            Gv1.Columns("From_Date").HeaderText = "Date"
            Gv1.Columns("From_Date").IsVisible = False

            Gv1.Columns("GSTReg_No").HeaderText = "GSTIN"
            Gv1.Columns("GSTReg_No").IsVisible = False

            Gv1.Columns("Doc_No").HeaderText = "Documant No"
            Gv1.Columns("Doc_No").IsVisible = False

            Gv1.Columns("Fiscal_Name").HeaderText = "Fiscal Year"
            Gv1.Columns("Fiscal_Name").IsVisible = False

            Gv1.Columns("CycleNo").HeaderText = "Cycle No"
            Gv1.Columns("CycleNo").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Date Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "DCS Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            Gv1.Columns("Payee_Joint_Name").HeaderText = "Society Name"
            Gv1.Columns("Payee_Joint_Name").IsVisible = True

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = False

            Gv1.Columns("Branch_Name").HeaderText = "Branch"
            Gv1.Columns("Branch_Name").IsVisible = False

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = MultipleFinderFillAuto

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = MultipleFinderFillAuto

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "Account No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = True

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True

            Gv1.Columns("FD").HeaderText = "FD"
            Gv1.Columns("FD").IsVisible = False

            Gv1.Columns("TD").HeaderText = "TD"
            Gv1.Columns("TD").IsVisible = False

            Gv1.Columns("Bank_Advise_No").HeaderText = "Bank Advise No"
            Gv1.Columns("Bank_Advise_No").IsVisible = False

            Gv1.Columns("Bank_Advise_Date").HeaderText = "Bank Advise Date"
            Gv1.Columns("Bank_Advise_Date").IsVisible = False

        ElseIf rbtnBankAdvice.IsChecked OrElse rbtnCompulsory.IsChecked Then
            If Gv1.Columns.Contains("CycleRange") Then
                Gv1.Columns("CycleRange").HeaderText = "Cycle Range"
                Gv1.Columns("CycleRange").IsVisible = False
            End If
            If Gv1.Columns.Contains("GRPColumn") Then
                Gv1.Columns("GRPColumn").HeaderText = "Group Range"
                Gv1.Columns("GRPColumn").IsVisible = False
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                Gv1.Columns("Payee_Joint_Branch_Name").HeaderText = "Payee_Joint_Branch_Name"
                Gv1.Columns("Payee_Joint_Branch_Name").IsVisible = False
            End If
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
                Gv1.Columns("Bank_Code").HeaderText = "Bank"
                Gv1.Columns("Bank_Code").IsVisible = False

                Gv1.Columns("Branch_Name").HeaderText = "Branch"
                Gv1.Columns("Branch_Name").IsVisible = False
            End If

            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") <> CompairStringResult.Equal Then
                If Gv1.Columns.Contains("GRPColumns") Then
                    Gv1.Columns("GRPColumns").HeaderText = "Group Ranges"
                    Gv1.Columns("GRPColumns").IsVisible = False
                End If

                If Gv1.Columns.Contains("Bank_Name") Then
                    Gv1.Columns("Bank_Name").HeaderText = "Bank_Name"
                    Gv1.Columns("Bank_Name").IsVisible = False
                End If

                If Gv1.Columns.Contains("BankAccountNo") Then
                    Gv1.Columns("BankAccountNo").HeaderText = "BankAccountNo"
                    Gv1.Columns("BankAccountNo").IsVisible = False
                End If

                If Gv1.Columns.Contains("BankIFSCCode") Then
                    Gv1.Columns("BankIFSCCode").HeaderText = "BankIFSCCode"
                    Gv1.Columns("BankIFSCCode").IsVisible = False
                End If

                If Gv1.Columns.Contains("BankBranchAddress") Then
                    Gv1.Columns("BankBranchAddress").HeaderText = "BankBranchAddress"
                    Gv1.Columns("BankBranchAddress").IsVisible = False
                End If

                If Gv1.Columns.Contains("Company Bank") Then
                    Gv1.Columns("Company Bank").HeaderText = "Company Bank"
                    Gv1.Columns("Company Bank").IsVisible = False
                End If

                If Gv1.Columns.Contains("Company Bank Account No") Then
                    Gv1.Columns("Company Bank Account No").HeaderText = "Company Bank Account No"
                    Gv1.Columns("Company Bank Account No").IsVisible = False
                End If

                If Gv1.Columns.Contains("Bank_Code") Then
                    Gv1.Columns("Bank_Code").HeaderText = "Bank"
                    Gv1.Columns("Bank_Code").IsVisible = True
                End If

                If Gv1.Columns.Contains("Branch_Name") Then
                    Gv1.Columns("Branch_Name").HeaderText = "Branch"
                    Gv1.Columns("Branch_Name").IsVisible = True
                End If
            End If

            If Gv1.Columns.Contains("Comp_Name") Then
                Gv1.Columns("Comp_Name").HeaderText = "Company Name"
                Gv1.Columns("Comp_Name").IsVisible = False
            End If

            If Gv1.Columns.Contains("Comp_address") Then
                Gv1.Columns("Comp_address").HeaderText = "Company Address"
                Gv1.Columns("Comp_address").IsVisible = False
            End If

            If Gv1.Columns.Contains("From_Date") Then
                Gv1.Columns("From_Date").HeaderText = "Date"
                Gv1.Columns("From_Date").IsVisible = False
            End If

            If Gv1.Columns.Contains("GSTReg_No") Then
                Gv1.Columns("GSTReg_No").HeaderText = "GSTIN"
                Gv1.Columns("GSTReg_No").IsVisible = False
            End If

            If Gv1.Columns.Contains("Doc_No") Then
                Gv1.Columns("Doc_No").HeaderText = "Documant No"
                Gv1.Columns("Doc_No").IsVisible = False
            End If

            If Gv1.Columns.Contains("CompPhone") Then
                Gv1.Columns("CompPhone").HeaderText = "Phone No"
                Gv1.Columns("CompPhone").IsVisible = False
            End If

            If Gv1.Columns.Contains("Regn_No") Then
                Gv1.Columns("Regn_No").HeaderText = "Regn No"
                Gv1.Columns("Regn_No").IsVisible = False
            End If

            If Gv1.Columns.Contains("MCC_NAME") Then
                Gv1.Columns("MCC_NAME").HeaderText = "Area"
                Gv1.Columns("MCC_NAME").IsVisible = False
            End If

            If Gv1.Columns.Contains("FD") Then
                Gv1.Columns("FD").HeaderText = "FD"
                Gv1.Columns("FD").IsVisible = False
            End If
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then

            If Gv1.Columns.Contains("TD") Then
                Gv1.Columns("TD").HeaderText = "TD"
                Gv1.Columns("TD").IsVisible = False
            End If



            'Gv1.Columns("[Bank Advise No]").HeaderText = "Bank Advise No"
            'Gv1.Columns("[Bank Advise No]").IsVisible = True

            'Gv1.Columns("[Bank Advise Date]").HeaderText = "Bank Advise Date"
            'Gv1.Columns("[Bank Advise Date]").IsVisible = False

            'Gv1.Columns("Bank Advice Status").HeaderText = "Bank Advice Status"
            'Gv1.Columns("Bank Advice Status").IsVisible = False
            'End If

            If MultipleFinderFillAuto Then
            Else
                If clsCommon.myLen(Gv1.Columns("Location_Code")) Then
                    Gv1.Columns("Location_Code").HeaderText = "Location"
                    Gv1.Columns("Location_Code").IsVisible = False
                End If
                If clsCommon.myLen(Gv1.Columns("Location_Desc")) Then
                    Gv1.Columns("Location_Desc").HeaderText = "Location Name"
                    Gv1.Columns("Location_Desc").IsVisible = False
                End If
            End If

            If Gv1.Columns.Contains("Fiscal_Name") Then
                Gv1.Columns("Fiscal_Name").HeaderText = "Fiscal Year"
                Gv1.Columns("Fiscal_Name").IsVisible = False
            End If

            If Gv1.Columns.Contains("CycleNo") Then
                Gv1.Columns("CycleNo").HeaderText = "Cycle No"
                Gv1.Columns("CycleNo").IsVisible = False
            End If

            If Gv1.Columns.Contains("Date_Range") Then
                Gv1.Columns("Date_Range").HeaderText = "Date Range"
                Gv1.Columns("Date_Range").IsVisible = False
            End If

            If Gv1.Columns.Contains("VLC_CODE_Uploader") Then
                Gv1.Columns("VLC_CODE_Uploader").HeaderText = "DCS Code"
                Gv1.Columns("VLC_CODE_Uploader").IsVisible = True
            End If

            If Gv1.Columns.Contains("Payee_Joint_Name") Then
                Gv1.Columns("Payee_Joint_Name").HeaderText = "Society Name"
                Gv1.Columns("Payee_Joint_Name").IsVisible = True
            End If

            If Gv1.Columns.Contains("Bank_Code_Desc") Then
                Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
                Gv1.Columns("Bank_Code_Desc").IsVisible = MultipleFinderFillAuto
            End If

            If Gv1.Columns.Contains("Payee_Joint_IFSC_Code") Then
                Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
                Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = MultipleFinderFillAuto
            End If

            If Gv1.Columns.Contains("Payee_Joint_Account_No") Then
                Gv1.Columns("Payee_Joint_Account_No").HeaderText = "Account No"
                Gv1.Columns("Payee_Joint_Account_No").IsVisible = True
            End If

            If Gv1.Columns.Contains("Payable_Amount") Then
                Gv1.Columns("Payable_Amount").HeaderText = "Amount"
                Gv1.Columns("Payable_Amount").IsVisible = True
            End If

        ElseIf rbtnBankWiseSummary.IsChecked Then

            Gv1.Columns("SNO").HeaderText = "SNo"
            Gv1.Columns("SNO").IsVisible = True

            Gv1.Columns("CycleRange").HeaderText = "Cycle Range"
            Gv1.Columns("CycleRange").IsVisible = False

            Gv1.Columns("GRPColumn").HeaderText = "Group Range"
            Gv1.Columns("GRPColumn").IsVisible = False

            Gv1.Columns("Comp_Name").HeaderText = "Company Name"
            Gv1.Columns("Comp_Name").IsVisible = False

            Gv1.Columns("Comp_address").HeaderText = "Company Address"
            Gv1.Columns("Comp_address").IsVisible = False

            Gv1.Columns("From_Date").HeaderText = "Date"
            Gv1.Columns("From_Date").IsVisible = False

            Gv1.Columns("GSTReg_No").HeaderText = "GSTIN"
            Gv1.Columns("GSTReg_No").IsVisible = False

            Gv1.Columns("Fiscal_Name").HeaderText = "Fiscal Year"
            Gv1.Columns("Fiscal_Name").IsVisible = False

            Gv1.Columns("CycleNo").HeaderText = "Cycle No"
            Gv1.Columns("CycleNo").IsVisible = False

            Gv1.Columns("Date_Range").HeaderText = "Date Range"
            Gv1.Columns("Date_Range").IsVisible = False

            Gv1.Columns("Bank_Code").HeaderText = "Bank"
            Gv1.Columns("Bank_Code").IsVisible = True

            Gv1.Columns("Branch_Name").HeaderText = "Branch"
            Gv1.Columns("Branch_Name").IsVisible = True

            Gv1.Columns("Bank_Code_Desc").HeaderText = "Bank Name"
            Gv1.Columns("Bank_Code_Desc").IsVisible = MultipleFinderFillAuto

            Gv1.Columns("Payee_Joint_IFSC_Code").HeaderText = "IFSC Code"
            Gv1.Columns("Payee_Joint_IFSC_Code").IsVisible = MultipleFinderFillAuto

            Gv1.Columns("Payee_Joint_Account_No").HeaderText = "Account No"
            Gv1.Columns("Payee_Joint_Account_No").IsVisible = MultipleFinderFillAuto

            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True
            Gv1.Columns("Payable_Amount").FormatString = "{0:n2}"
            'Gv1.Columns("Payable_Amount").FormatString = "{0:n2}"
        ElseIf rbtnCurrentBankWiseSummary.IsChecked OrElse rbtnCompulsoryWiseSummary.IsChecked Then
            Gv1.Columns("Payable_Amount").HeaderText = "Amount"
            Gv1.Columns("Payable_Amount").IsVisible = True
            Gv1.Columns("Payable_Amount").FormatString = "{0:n2}"
            'Gv1.Columns("Payable_Amount").FormatString = "{0:n2}"
            Gv1.Columns("Bank_Desc").IsVisible = True
            Gv1.Columns("Bank_Desc").HeaderText = "Bank Description"
            Gv1.Columns("GRPColumn").IsVisible = False
        End If

        If rbtnBothSavCur.IsChecked Then
            If Gv1.Columns.Contains("DCSCount") Then
                Gv1.Columns("DCSCount").HeaderText = "DCS Count"
                Gv1.Columns("DCSCount").IsVisible = False
            End If

            Gv1.Columns("VLC_CODE_Uploader").HeaderText = "DCS Code"
            Gv1.Columns("VLC_CODE_Uploader").IsVisible = True

            'FOR ALW Temp Work
            Gv1.Columns("RI").IsVisible = False
            Gv1.Columns("Branch_Name").IsVisible = False

            Gv1.Columns("BankCode").HeaderText = "Bank Code"
            Gv1.Columns("BankCode").IsVisible = True

            Gv1.Columns("BankName").HeaderText = "Bank Name"
            Gv1.Columns("BankName").IsVisible = True

            Gv1.Columns("DCS_Name").HeaderText = "DCS Name"
            Gv1.Columns("DCS_Name").IsVisible = True

            Gv1.Columns("Bill_Amt").HeaderText = "Bill Amount"
            Gv1.Columns("Bill_Amt").IsVisible = True

            Gv1.Columns("SavingAccountNo").HeaderText = "Saving Account No."
            Gv1.Columns("SavingAccountNo").IsVisible = True

            Gv1.Columns("SavingIFSCCode").HeaderText = "Saving IFSC Code"
            Gv1.Columns("SavingIFSCCode").IsVisible = True

            Gv1.Columns("SavingAmt").HeaderText = "Saving Amount"
            Gv1.Columns("SavingAmt").IsVisible = True

            Gv1.Columns("CurrentAccountNo").HeaderText = "Current Account No."
            Gv1.Columns("CurrentAccountNo").IsVisible = True

            Gv1.Columns("CurrentIFSCCode").HeaderText = "Current IFSC Code"
            Gv1.Columns("CurrentIFSCCode").IsVisible = True

            Gv1.Columns("CurrentAmt").HeaderText = "Current Amount"
            Gv1.Columns("CurrentAmt").IsVisible = True
        End If
        If rbtnBothSavCur.IsChecked Then
            Dim summaryRowItemB As New GridViewSummaryRowItem()
            'Dim MilkTypeB As New GridViewSummaryItem("Payable_Amount", "{0:n0}", GridAggregateFunction.Sum)
            Dim Bill_Amt As New GridViewSummaryItem("Bill_Amt", "{0:n2}", GridAggregateFunction.Sum)
            summaryRowItemB.Add(Bill_Amt)
            Dim SavingAmt As New GridViewSummaryItem("SavingAmt", "{0:n2}", GridAggregateFunction.Sum)
            summaryRowItemB.Add(SavingAmt)
            Dim CurrentAmt As New GridViewSummaryItem("CurrentAmt", "{0:n2}", GridAggregateFunction.Sum)
            summaryRowItemB.Add(CurrentAmt)
            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItemB)
            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom
        Else
            Dim summaryRowItemB As New GridViewSummaryRowItem()
            'Dim MilkTypeB As New GridViewSummaryItem("Payable_Amount", "{0:n0}", GridAggregateFunction.Sum)
            Dim MilkTypeB As New GridViewSummaryItem("Payable_Amount", "{0:n2}", GridAggregateFunction.Sum)
            summaryRowItemB.Add(MilkTypeB)
            Gv1.MasterTemplate.SummaryRowsBottom.Add(summaryRowItemB)
            Gv1.MasterView.SummaryRows(0).PinPosition = PinnedRowPosition.Bottom

        End If

        Gv1.AutoSizeRows = True
        Gv1.BestFitColumns()
        Gv1.MasterTemplate.AutoExpandGroups = True
    End Sub
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        EnableDisaableControls(True)
        Gv1.DataSource = Nothing
        Gv2.DataSource = Nothing
        RadPageView1.SelectedPage = RadPageViewPage1
        lblBankAdviseNo.Visible = False
        txtBankAdviseNo.Visible = False
        txtBankAdviseNo.Value = ""
    End Sub
    Sub EnableDisaableControls(ByVal flag As Boolean)
        txtMCC.Enabled = flag
        txtFiscalYear.Enabled = flag
        txtPaymentCycleFrom.Enabled = flag
        txtPaymentCycleTo.Enabled = flag
        txtbankgroupname.Enabled = flag
        txtBankGroup.Enabled = flag
        ChkIFSCCode.Enabled = flag
        fndArea.Enabled = flag
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
            Dim strUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select City_Code from TSPL_COMPANY_MASTER where Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "))
            Dim strHeading As String = ""
            Dim strHeadingGV As String = ""
            If MultipleFinderFillAuto Then
                strHeading = clsCommon.myCstr("BankAdvice Report(OHC) from " + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + " to " + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + " for '" + strUnit + "' unit  ")
            Else
                strHeading = clsCommon.myCstr("Milk Bill Procurement Summary For Cycle : " + txtPaymentCycleFrom.Value + " To " + txtPaymentCycleTo.Value + ", " + txtFiscalYear.Value + "")
            End If
            strHeadingGV = clsCommon.myCstr("One Day Stop Summary For Cycle : " + txtPaymentCycleFrom.Value + " To " + txtPaymentCycleTo.Value + ", " + txtFiscalYear.Value + "")

            Dim arrHeader As List(Of String) = New List(Of String)()
            Dim arrHeaderGV As List(Of String) = New List(Of String)()
            If MultipleFinderFillAuto Then
                If rbtnSaving.IsChecked = True Then
                    arrHeader.Add("Data: Saving ")
                End If
                If rbtnBankAdvice.IsChecked = True Then
                    arrHeader.Add("Report Type: Bank Advice")
                End If
                If rbtnBankWiseSummary.IsChecked = True Then
                    arrHeader.Add("Report Type: Bank Wise Summary")
                End If
            Else
                arrHeader.Add("Mcc : " + txtMCC.Value)
            End If
            arrHeaderGV.Add("Report Type: One Day Stop Report")
            If exporter = EnumExportTo.Excel Then
                clsCommon.MyExportToExcelGrid(strHeading, Gv1, arrHeader, Me.Text)
                clsCommon.MyExportToExcelGrid(strHeadingGV, Gv2, arrHeaderGV, Me.Text)
            Else
                clsCommon.MyExportToPDF(strHeading, Gv1, arrHeader, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
                clsCommon.MyExportToPDF(strHeadingGV, Gv2, arrHeaderGV, Me.Text, PageSetupReport_ID, objCommonVar.CurrentUserCode)
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




    'Private Sub Txtmccode__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean)
    '    Dim whrcls As String = ""

    '    txtmccode.Value = clsMccMaster.getFinder(whrcls, txtmccode.Value, isButtonClicked)
    '    If clsCommon.myLen(txtmccode.Value) > 0 Then
    '        lblmccname.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select mcc_name from tspl_mcc_master where mcc_code='" + txtmccode.Value + "'"))
    '    Else
    '        txtmccode.Value = ""
    '        lblmccname.Text = ""
    '    End If
    'End Sub

    Private Sub TxtFiscalYear__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtFiscalYear._MYValidating
        Try
            Dim qry As String = "select Fiscal_Code,Fiscal_Name,Start_Date,End_Date from TSPL_FISCAL_YEAR_MASTER"
            txtFiscalYear.Value = clsCommon.ShowSelectForm("LRFY", qry, "Fiscal_Code", "", txtFiscalYear.Value, "", isButtonClicked)
            'fromDate.Value = clsCommon.GETSERVERDATE()
            'ToDate.Value = clsCommon.GETSERVERDATE()
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub TxtPaymentCycleFrom__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtPaymentCycleFrom._MYValidating
        Try
            If clsCommon.myLen(txtFiscalYear.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Plz Select Fiscal Year First.", Me.Text)
                txtFiscalYear.Focus()
                Exit Sub
            End If
            Dim whrcls As String = " Fiscal_Code='" + txtFiscalYear.Value + "' "
            Dim qry As String = "SELECT distinct convert(int,name) as Code,convert(varchar,From_Date,103) as [From Date],convert(varchar,To_Date,103) as [To Date] FROM TSPL_PAYMENT_CYCLE_GENERATED"
            txtPaymentCycleFrom.Value = clsCommon.ShowSelectForm("LRPCF", qry, "Code", whrcls, txtPaymentCycleFrom.Value, "Code", isButtonClicked)

            If clsCommon.myLen(txtPaymentCycleFrom.Value) > 0 Then
                txtPaymentCycleTo.Value = txtPaymentCycleFrom.Value
            End If

        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK)
        End Try

    End Sub

    Private Sub Patment_Cycle_changed()
        Try
            Dim dt As DataTable
            Dim qry As String = "SELECT Name ,From_Date,To_Date FROM TSPL_PAYMENT_CYCLE_GENERATED where Fiscal_Code='" + txtFiscalYear.Value + "' and Name='" + txtPaymentCycleFrom.Value + "' "
            dt = clsDBFuncationality.GetDataTable(qry)
            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                'clsCommon.MyMessageBoxShow("No Payment Cycle found for Selected Fiscal Year")
                Exit Sub
            End If

            fromDate.Value = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select From_Date from TSPL_PAYMENT_CYCLE_GENERATED where Fiscal_Code='" + txtFiscalYear.Value + "' and Name='" + txtPaymentCycleFrom.Value + "' "))
            ToDate.Value = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select To_Date from TSPL_PAYMENT_CYCLE_GENERATED where Fiscal_Code='" + txtFiscalYear.Value + "' and Name='" + txtPaymentCycleTo.Value + "' "))
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub TxtPaymentCycleTo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtPaymentCycleTo._MYValidating
        Try
            If clsCommon.myLen(txtFiscalYear.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Plz Select Fiscal Year First.", Me.Text)
                txtFiscalYear.Focus()
                Exit Sub
            End If
            Dim whrcls As String = " Fiscal_Code='" + txtFiscalYear.Value + "' "
            Dim qry As String = "SELECT distinct convert(int,name) as Code,convert(varchar,From_Date,103) as [From Date],convert(varchar,To_Date,103) as [To Date] FROM TSPL_PAYMENT_CYCLE_GENERATED"
            txtPaymentCycleTo.Value = clsCommon.ShowSelectForm("LRPCT", qry, "Code", whrcls, txtPaymentCycleTo.Value, "Code", isButtonClicked)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub RmiExcelGrid_Click(sender As Object, e As EventArgs) Handles rmiExcelGrid.Click
        Try
            If Gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If
            Dim strUnit As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select City_Code from TSPL_COMPANY_MASTER where Comp_Code = '" + objCommonVar.CurrentCompanyCode + "' "))
            Dim strHeading As String = ""
            If MultipleFinderFillAuto Then
                strHeading = clsCommon.myCstr("BankAdvice Report(OHC) from " + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + " to " + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") + " for '" + strUnit + "' unit  ")
            Else
                strHeading = clsCommon.myCstr("Milk Bill Procurement Summary For Cycle : " + txtPaymentCycleFrom.Value + " To " + txtPaymentCycleTo.Value + ", " + txtFiscalYear.Value + "")
            End If
            'Dim strHeading As String = clsCommon.myCstr("Milk Bill Procurement Summary For Cycle : " + txtPaymentCycleFrom.Value + " To " + txtPaymentCycleTo.Value + ", " + txtFiscalYear.Value + "")
            Dim arrHeader As List(Of String) = New List(Of String)()
            If MultipleFinderFillAuto Then
            Else
                arrHeader.Add("Mcc : " + txtMCC.Value)
            End If

            clsCommon.MyExportToExcelGrid(strHeading, Gv1, arrHeader, Me.Text, True)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, "Error", MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
        End Try
    End Sub

    Private Sub txtMCC__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtMCC._MYValidating
        Try
            'Dim qry As String = "select MCC_Code as Code,MCC_NAME as Name,TSPL_MCC_MASTER.plant_code as [Plant Code],tspl_location_master.location_desc as [Plant Name] from TSPL_MCC_MASTER left join tspl_location_master on tspl_location_master.location_code=TSPL_MCC_MASTER.plant_code"

            Dim qry As String = "select TSPL_MCC_MASTER.MCC_Code as Code,MCC_NAME as Name,TSPL_MCC_MASTER.plant_code as [Plant Code],tspl_location_master.location_desc as [Plant Name] from TSPL_MCC_MASTER 
            Left join tspl_location_master on tspl_location_master.location_code=TSPL_MCC_MASTER.plant_code "
            If fndArea.Value IsNot Nothing AndAlso fndArea.Value.Count > 0 Then
                qry += " and TSPL_MCC_MASTER.Area_Location_Code ='" + clsCommon.myCstr(fndArea.Value) + "' "
            End If
            Dim strWhrcls As String = "location_code ='" + clsCommon.myCstr(txtMCC.Value) + "'"
            txtMCC.Value = clsCommon.ShowSelectForm("vendorBadvice", qry, "Code", "", txtMCC.Value, "Code", isButtonClicked)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub RadMenuItem2_Click(sender As Object, e As EventArgs) Handles RadMenuItem2.Click
        If ChkIFSCCode.Checked Then
            Printt()
        Else
            If rbtnBothSavCur.IsChecked Then
                If dtBankAdvise IsNot Nothing AndAlso dtBankAdvise.Rows.Count > 1 Then
                    If clsCommon.myLen(txtBankAdviseNo.Value) <= 0 Then
                        Throw New Exception("Please select atleast one Bank Advise No")
                    Else
                        Print(True, "", Nothing, txtBankAdviseNo.Value)
                    End If
                Else
                    Print(True, Nothing, Nothing)
                End If
            Else
                Print(True, Nothing, Nothing)
            End If
            ' Print(True, Nothing, Nothing)
        End If
    End Sub

    Sub SetToDate()
        'If MultipleFinderFillAuto Then
        Dim PaymentCycleType As String = ""
        Dim PaymentCycleValue As Integer = 0

        'If clsCommon.myLen(fndLoc.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
        '    clsCommon.MyMessageBoxShow("Please select the Location first")
        '    Exit Sub
        'End If
        'If MultipleFinderFillAuto = True Then
        '    If mfndMcc.arrValueMember Is Nothing OrElse mfndMcc.arrValueMember.Count <= 0 Then
        '        clsCommon.MyMessageBoxShow("Please select the Location first")
        '        Exit Sub
        '    End If
        'End If
        Dim strMCCcode = ""
        If clsCommon.myLen(txtMCC.Value) Then
            strMCCcode = " location_Code in ( '" + clsCommon.myCstr(txtMCC.Value) + "')  and "
        End If


        Dim dt As DataTable = clsDBFuncationality.GetDataTable(" select TSPL_MCC_MASTER.Payment_Cycle,TSPL_PAYMENT_CYCLE_MASTER.PC_TYPE,TSPL_PAYMENT_CYCLE_MASTER.PC_VALUE  from TSPL_MCC_MASTER left outer join TSPL_PAYMENT_CYCLE_MASTER on TSPL_PAYMENT_CYCLE_MASTER.PC_CODE=TSPL_MCC_MASTER.Payment_Cycle   where TSPL_MCC_MASTER.MCC_Code  in (select Location_Code  from TSPL_LOCATION_MASTER where " + strMCCcode + "  Location_Category='MCC' and Rejected_Type='N') ")
        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
            clsCommon.MyMessageBoxShow(Me, "No Payment Cycle found on current MCC/Location", Me.Text)
            Exit Sub
        End If
        PaymentCycleType = clsCommon.myCstr(dt.Rows(0)("PC_TYPE"))
        PaymentCycleValue = clsCommon.myCdbl(dt.Rows(0)("PC_VALUE"))
        Dim dtCurr As DateTime = clsCommon.GETSERVERDATE()
        If clsCommon.CompairString(PaymentCycleType, "Day") = CompairStringResult.Equal Then
            If fromDate.Value.Day Mod PaymentCycleValue <> 1 And (Not PaymentCycleValue = 1) Then
                clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month or at interval of " & PaymentCycleValue & " Day, Because MCC has payment Cycle of " & PaymentCycleValue & " Day ")
                fromDate.Value = New Date(dtCurr.Year, dtCurr.Month, 1)
                ToDate.Value = fromDate.Value
                Exit Sub
            End If
            ToDate.Value = fromDate.Value.AddDays(PaymentCycleValue - 1)

            If fromDate.Value.Month <> ToDate.Value.Month Then
                ToDate.Value = New Date(fromDate.Value.Year, fromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
            End If
            Dim dtNxtPay As DateTime = ToDate.Value.AddDays(Math.Ceiling(PaymentCycleValue / 2.0))
            If fromDate.Value.Month <> dtNxtPay.Month Then
                ToDate.Value = New Date(fromDate.Value.Year, fromDate.Value.Month, 1).AddMonths(1).AddDays(-1)
            End If
        ElseIf clsCommon.CompairString(PaymentCycleType, "Month") = CompairStringResult.Equal Then
            If clsCommon.myCdbl(clsCommon.GetPrintDate(fromDate.Value, "dd")) <> 1 Then
                clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month, Because MCC has payment Cycle of Month Type")
                fromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                ToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                Exit Sub
            End If
            ToDate.Value = DateAdd(DateInterval.Month, PaymentCycleValue, fromDate.Value)
        ElseIf clsCommon.CompairString(PaymentCycleType, "Year") = CompairStringResult.Equal Then
            If clsCommon.myCdbl(clsCommon.GetPrintDate(fromDate.Value, "dd")) <> 1 Then
                clsCommon.MyMessageBoxShow(Me, "Date can only be first day of month, Because MCC has payment Cycle of Year Type")
                fromDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                ToDate.Value = "01/" & DatePart(DateInterval.Month, dtCurr) & "/" & DatePart(DateInterval.Year, dtCurr)
                Exit Sub
            End If
            ToDate.Value = DateAdd(DateInterval.Year, PaymentCycleValue, fromDate.Value)
        ElseIf clsCommon.CompairString(PaymentCycleType, "Week") = CompairStringResult.Equal Then
            Dim today As Date = fromDate.Value
            Dim dayDiff As Integer = today.DayOfWeek - IIf(PaymentCycleValue = 1, DayOfWeek.Sunday, IIf(PaymentCycleValue = 2, DayOfWeek.Monday, IIf(PaymentCycleValue = 3, DayOfWeek.Tuesday, IIf(PaymentCycleValue = 4, DayOfWeek.Wednesday, IIf(PaymentCycleValue = 5, DayOfWeek.Thursday, IIf(PaymentCycleValue = 6, DayOfWeek.Friday, DayOfWeek.Saturday))))))
            fromDate.Value = today.AddDays(-dayDiff)
            ToDate.Value = fromDate.Value.AddDays(6)
        End If
        ' End If
        'If clsCommon.myLen(txtMCC.Text) > 0 Then
        '    txtPaymentCycleNo.Text = clsGenratePaymentCycles.GetPaymentCycleNo(txtMCC.Text, dtpToDate.Value)
        '    txtFiscalYear.Text = clsGenratePaymentCycles.GetPaymentFiscalCode(txtMCC.Text, dtpToDate.Value)
        'Else
        '    txtPaymentCycleNo.Text = clsGenratePaymentCycles.GetPaymentCycleNo(fndLoc.Value, dtpToDate.Value)
        '    txtFiscalYear.Text = clsGenratePaymentCycles.GetPaymentFiscalCode(fndLoc.Value, dtpToDate.Value)
        'End If
        'End If
    End Sub

    Private Sub fromDate_Leave(sender As Object, e As EventArgs) Handles fromDate.Leave
        If MultipleFinderFillAuto Then
            SetToDate()
        Else
            SetToDate()
        End If
        LoadBankAdviseNo()
    End Sub
    Sub LoadBankAdviseNo()
        dtBankAdvise = clsDBFuncationality.GetDataTable("select TSPL_BANK_ADVISE_DETAIL.Document_No from TSPL_BANK_ADVISE_DETAIL left join TSPL_BANK_ADVISE on TSPL_BANK_ADVISe.Document_No = TSPL_BANK_ADVISE_DETAIL.Document_No left join TSPL_PAYMENT_PROCESS_HEAD ON TSPL_PAYMENT_PROCESS_HEAD.Doc_No =TSPL_BANK_ADVISE.Payment_Process_Document_No where TSPL_PAYMENT_PROCESS_HEAD.From_Date >= '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") & "' and TSPL_PAYMENT_PROCESS_HEAD.To_Date <= '" & clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") & "' and TSPL_BANK_ADVISE.Status = 1 group by TSPL_BANK_ADVISE_DETAIL.Document_No ")
        If dtBankAdvise IsNot Nothing AndAlso dtBankAdvise.Rows.Count > 1 Then
            lblBankAdviseNo.Visible = True
            txtBankAdviseNo.Visible = True
        Else
            txtBankAdviseNo.Value = ""
            lblBankAdviseNo.Visible = False
            txtBankAdviseNo.Visible = False
        End If
    End Sub

    Private Sub fromDate_Validating(sender As Object, e As CancelEventArgs) Handles fromDate.Validating
        If MultipleFinderFillAuto Then
            SetToDate()
        Else
            SetToDate()
        End If
        LoadBankAdviseNo()
    End Sub

    Private Sub txtBankGroup__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtBankGroup._MYValidating
        Try
            Dim qry As String = "select BANK_GROUP_CODE as Code,DESCRIPTION as Name from TSPL_BANK_GROUP_MASTER"
            txtBankGroup.Value = clsCommon.ShowSelectForm("fmVBA", qry, "Code", "", txtBankGroup.Value, "BANK_GROUP_CODE", isButtonClicked)
            qry = "select TSPL_BANK_GROUP_MASTER.DESCRIPTION from TSPL_BANK_GROUP_MASTER where BANK_GROUP_CODE ='" + txtBankGroup.Value + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                txtbankgroupname.Text = clsCommon.myCstr(dt.Rows(0)("description"))
            Else
                txtbankgroupname.Text = ""
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Dim flagSaving As Boolean = False
    Dim flagCurrent As Boolean = False
    Dim flagBothSavingCurrent As Boolean = False
    Dim flagSavingCurrent As Boolean = False
    Private Sub rbtnBankAdvice_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnBankAdvice.ToggleStateChanged, rbtnBankWiseSummary.ToggleStateChanged
        If rbtnBankAdvice.IsChecked OrElse rbtnBankWiseSummary.IsChecked OrElse rbtnCurrentBankWiseSummary.IsChecked Then
            rbtnBothSavCur.IsChecked = False
            If Not flagSaving Then
                flagCurrent = True
                rbtnSaving.IsChecked = False
                rbtnCompulsory.IsChecked = False
                rbtnCompulsoryWiseSummary.IsChecked = False
                rbtnBothSavCur.IsChecked = False
                flagCurrent = False
            End If
        End If




    End Sub

    Private Sub rbtnSaving_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnSaving.ToggleStateChanged, rbtnCompulsory.ToggleStateChanged, rbtnCompulsoryWiseSummary.ToggleStateChanged
        If rbtnSaving.IsChecked OrElse rbtnCompulsory.IsChecked OrElse rbtnCompulsoryWiseSummary.IsChecked Then
            rbtnBothSavCur.IsChecked = False

            If Not flagCurrent Then
                flagSaving = True
                rbtnBankAdvice.IsChecked = False
                rbtnBankWiseSummary.IsChecked = False
                rbtnCurrentBankWiseSummary.IsChecked = False
                rbtnBothSavCur.IsChecked = False
                flagSaving = False
            End If
        End If
        If rbtnSaving.IsChecked AndAlso DCSWiseFilterEnable = True Then
            txtDCS.Visible = True
            MyLabel15.Visible = True
        End If




    End Sub
    Private Sub ReStoreGridLayout()
        Try
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
            common.clsCommon.MyMessageBoxShow(Me, err.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmsaveLayout_Click(sender As Object, e As EventArgs) Handles rmsaveLayout.Click
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
    End Sub

    Private Sub rmDeleteLayout_Click(sender As Object, e As EventArgs) Handles rmDeleteLayout.Click
        clsGridLayout.DeleteData(MyBase.Form_ID, objCommonVar.CurrentUserCode)
        common.clsCommon.MyMessageBoxShow(Me, "Layout Delete successfully", Me.Text)
    End Sub

    Sub Printt()
        Try


            If clsCommon.myLen(txtMCC.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
                clsCommon.MyMessageBoxShow(Me, "Plz Select MCC.", Me.Text)
                txtMCC.Focus()
                Exit Sub
            End If

            If clsCommon.myLen(txtFiscalYear.Value) <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "Plz Select Fiscal Year First.", Me.Text)
                txtFiscalYear.Focus()
                Exit Sub
            End If
            If clsCommon.myLen(txtPaymentCycleFrom.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
                'clsCommon.MyMessageBoxShow("Plz Select Payment Cycle From First.", Me.Text)
                'txtPaymentCycleFrom.Focus()
                'Exit Sub
            End If
            If clsCommon.myLen(txtPaymentCycleTo.Value) <= 0 AndAlso MultipleFinderFillAuto = False Then
                'clsCommon.MyMessageBoxShow("Plz Select Payment Cycle To First.", Me.Text)
                'txtPaymentCycleTo.Focus()
                'Exit Sub
            End If

            If clsCommon.myCdbl(txtPaymentCycleFrom.Value) > clsCommon.myCdbl(txtPaymentCycleTo.Value) AndAlso MultipleFinderFillAuto = False Then
                common.clsCommon.MyMessageBoxShow(Me, "From Payment Cycle can not be greater then to Payment Cycle", Me.Text)
                txtPaymentCycleFrom.Focus()
                Exit Sub
            End If
            If MultipleFinderFillAuto = False Then
                Patment_Cycle_changed()
            End If


            Dim strCycleRange As String = txtPaymentCycleTo.Value
            If clsCommon.myCdbl(txtPaymentCycleTo.Value) > clsCommon.myCdbl(txtPaymentCycleFrom.Value) Then
                strCycleRange = txtPaymentCycleTo.Value + " To  " + txtPaymentCycleFrom.Value
            End If
            Dim BaseQry As String = ""

            If rbtnCompulsoryWiseSummary.IsChecked AndAlso rbtnCurrentBankWiseSummary.IsChecked Then
                clsCommon.MyMessageBoxShow(Me, "You have to select only option at a time ", Me.Text)
                Exit Sub
            End If
            If rbtnSaving.IsChecked Then
                BaseQry = " select  '' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code+(case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment,Mp_V.vsp_payment),'')='Self' then coalesce(TSPL_VENDOR_MASTER .IFSC_Code,mp_V.IFSC_Code)   else coalesce(TSPL_VENDOR_MASTER .Joint_IFSC_Code,mp_v.Joint_IFSC_Code)   end)  as GRPColumn,
                                        TSPL_COMPANY_MASTER.Comp_Name
                                        ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name
                                        ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, coalesce(TSPL_VENDOR_MASTER.VSP_Payee_Name,Mp_V.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
              TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name, 
              TSPL_Vendor_MASTER.BankCode2 as Bank_Code_Desc, 
              case when isnull(
                coalesce(
                  TSPL_VENDOR_MASTER.vsp_payment, 
                  Mp_V.vsp_payment
                ), 
                ''
              )= 'Self' then coalesce(
                TSPL_VENDOR_MASTER.IFSCCode2, mp_V.IFSC_Code
              ) else coalesce(
                TSPL_VENDOR_MASTER.Joint_IFSC_Code, 
                mp_v.Joint_IFSC_Code
              ) end as Payee_Joint_IFSC_Code, 
              case when isnull(
                coalesce(
                  TSPL_VENDOR_MASTER.vsp_payment, 
                  Mp_V.vsp_payment
                ), 
                ''
              )= 'Self' then coalesce( 
                TSPL_VENDOR_MASTER.AccNo2, mp_V.Account_No
              ) else coalesce(
                TSPL_VENDOR_MASTER.Joint_Account_No, 
                mp_V.Joint_Account_No
              ) end as Payee_Joint_Account_No,"
                If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                    BaseQry += " Round(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as Payable_Amount "
                Else
                    BaseQry += " TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount "

                End If

                BaseQry += "from TSPL_PAYMENT_PROCESS_SAVING 
                                        left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
                                        left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
                                        left outer join TSPL_MP_MASTER mp on mp.MP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
                                        Left outer join tspl_vendor_master Mp_V on mp_V.Vendor_Code=mp.MP_Code  
                                        left outer join ( select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                                        left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
                                        left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_SAVING.VSP_CODE
                                        left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                                        left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                                        left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Bank_Code
                                        left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                        where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0  "

            ElseIf rbtnCompulsory.IsChecked Then

                BaseQry = " select  '' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code+(case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment,Mp_V.vsp_payment),'')='Self' then coalesce(TSPL_VENDOR_MASTER .IFSC_Code,mp_V.IFSC_Code)   else coalesce(TSPL_VENDOR_MASTER .Joint_IFSC_Code,mp_v.Joint_IFSC_Code)   end)  as GRPColumn,TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
                                        TSPL_COMPANY_MASTER.Comp_Name
                                        ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name
                                        ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, coalesce(TSPL_VENDOR_MASTER.VSP_Payee_Name,Mp_V.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
              TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name, 
              TSPL_Vendor_MASTER.BankCode2 as Bank_Code_Desc, 
              case when isnull(
                coalesce(
                  TSPL_VENDOR_MASTER.vsp_payment, 
                  Mp_V.vsp_payment
                ), 
                ''
              )= 'Self' then coalesce(
                TSPL_VENDOR_MASTER.IFSCCode2, mp_V.IFSC_Code
              ) else coalesce(
                TSPL_VENDOR_MASTER.Joint_IFSC_Code, 
                mp_v.Joint_IFSC_Code
              ) end as Payee_Joint_IFSC_Code, 
              case when isnull(
                coalesce(
                  TSPL_VENDOR_MASTER.vsp_payment, 
                  Mp_V.vsp_payment
                ), 
                ''
              )= 'Self' then coalesce( 
                TSPL_VENDOR_MASTER.AccNo2, mp_V.Account_No
              ) else coalesce(
                TSPL_VENDOR_MASTER.Joint_Account_No, 
                mp_V.Joint_Account_No
              ) end as Payee_Joint_Account_No,"
                If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                    BaseQry += " Round(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) as Payable_Amount "
                Else
                    BaseQry += " TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount "

                End If

                BaseQry += "from TSPL_PAYMENT_PROCESS_COMPULSORY 
                                        left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_COMPULSORY.Doc_No
                                        left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_COMPULSORY.AP_Invoice_No
                                        left outer join TSPL_MP_MASTER mp on mp.MP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
                                        Left outer join tspl_vendor_master Mp_V on mp_V.Vendor_Code=mp.MP_Code  
                                        left outer join ( select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                                        left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
                                        left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_COMPULSORY.VSP_CODE
                                        left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                                        left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                                        left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank
                                        left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                        where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'  and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0  "

            ElseIf rbtnCompulsoryWiseSummary.IsChecked Then

                BaseQry = "select  '" + strCycleRange + "' AS CycleRange,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code+TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code as GRPColumn,"
                Else
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code as GRPColumn,"
                End If

                BaseQry += "SPL_COMPANY_MASTER.Comp_Name
            ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
            ,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME
            ,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No," + IIf(MultipleFinderFillAuto = True, "", " TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc, ") + " TSPL_Fiscal_Year_Master.Fiscal_Name
            ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_BANK_MASTER.DESCRIPTION as Bank_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " Round(isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                Else
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        BaseQry += " Round(isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                    Else
                        BaseQry += " isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0) as Payable_Amount "

                    End If

                End If
                BaseQry += " from TSPL_PAYMENT_PROCESS_DETAIL 
            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
            left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank 
            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
            " + IIf(MultipleFinderFillAuto = True, "    ", " left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' ") + "
            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) " + IIf(MultipleFinderFillAuto = True, "  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected  ", " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code ") + " 
            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, "and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + txtMCC.Value + "' And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + "   and TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount>0 "

            End If

            If rbtnBankAdvice.IsChecked OrElse rbtnBankWiseSummary.IsChecked Then
                If rbtnSaving.IsChecked = False AndAlso rbtnCompulsory.IsChecked = False AndAlso rbtnCompulsoryWiseSummary.IsChecked = False Then
                    BaseQry = "select '" + strCycleRange + "' AS CycleRange,"
                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        BaseQry += " TSPL_Vendor_MASTER.Bank_Code+TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code as GRPColumn,"
                    Else
                        BaseQry += " TSPL_Vendor_MASTER.Bank_Code as GRPColumn,"
                    End If
                    BaseQry += "TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
            TSPL_COMPANY_MASTER.Comp_Name
            ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
            ,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No
            ,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No," + IIf(MultipleFinderFillAuto = True, "", " TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc, ") + " TSPL_Fiscal_Year_Master.Fiscal_Name
            ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,"
                    'BaseQry += " And tspl_mcc_master.Area_Location_Code ='" + clsCommon.myCstr(fndArea.Value) + "' "
                    If AreaWiseBilling = True Then
                        BaseQry += " TSPL_LOCATION_MASTER.Location_Desc AS MCC_Name, "
                    Else
                        BaseQry += " TSPL_MCC_MASTER.MCC_NAME, "
                    End If

                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                        BaseQry += " Round((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)),0) as Payable_Amount "
                    Else
                        If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                            BaseQry += " Round((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)),0) as Payable_Amount "
                        Else
                            'BaseQry += " Cast((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)) as decimal(18)) as Payable_Amount "
                            BaseQry += " (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)) as Payable_Amount "
                        End If

                    End If
                    BaseQry += " from TSPL_PAYMENT_PROCESS_DETAIL 
            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
            left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date"
                    If AreaWiseBilling = True Then
                        BaseQry += " LEFT OUTER JOIN TSPL_LOCATION_MASTER ON TSPL_LOCATION_MASTER.Location_Code=TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code"
                    Else
                        BaseQry += " left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected"
                    End If
                    BaseQry += " left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
            " + IIf(MultipleFinderFillAuto = True, "    ", " left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' ") + "
            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) " + IIf(MultipleFinderFillAuto = True, "  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected  ", " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code ") + " 
            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, "and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + txtMCC.Value + "' And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + "   "


                    If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal AndAlso clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        BaseQry += " And (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0.99 "
                    Else
                        BaseQry += " And (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0 "
                    End If


                    If clsCommon.myLen(fndArea.Value) > 0 Then
                        BaseQry += " And TSPL_PAYMENT_PROCESS_HEAD.Area_Location_Code = '" + fndArea.Value + "' "
                    End If
                End If
            ElseIf rbtnCurrentBankWiseSummary.IsChecked Then
                BaseQry = "select  '" + strCycleRange + "' AS CycleRange,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code+TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code as GRPColumn,"
                Else
                    BaseQry += " TSPL_Vendor_MASTER.Bank_Code as GRPColumn,"
                End If

                BaseQry += "TSPL_COMPANY_MASTER.Comp_Name
            ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
            ,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME
            ,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No," + IIf(MultipleFinderFillAuto = True, "", " TSPL_Location_MASTER.Location_Code,TSPL_Location_MASTER.Location_Desc, ") + " TSPL_Fiscal_Year_Master.Fiscal_Name
            ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc , TSPL_BANK_MASTER.DESCRIPTION as Bank_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No,"
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JSL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrComp_Code1, "NAG") = CompairStringResult.Equal Then
                    BaseQry += " Round(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                Else
                    If clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.RoundOffBankAdvice, clsFixedParameterCode.RoundOffBankAdvice, Nothing)) = "1" Then
                        BaseQry += " Round(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0),0) as Payable_Amount "
                    Else
                        BaseQry += " isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0) as Payable_Amount "
                    End If

                End If
                BaseQry += " from TSPL_PAYMENT_PROCESS_DETAIL 
            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
            left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='" + objCommonVar.CurrentCompanyCode + "'
            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current 
            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
            " + IIf(MultipleFinderFillAuto = True, "    ", " left outer join TSPL_Location_MASTER on TSPL_Location_MASTER.Loc_Segment_Code=TSPL_PAYMENT_PROCESS_HEAD.Loc_Seg_Code and  TSPL_Location_MASTER.Rejected_Type='N' and TSPL_Location_MASTER.Location_Category='MCC' ") + "
            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) " + IIf(MultipleFinderFillAuto = True, "  and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected  ", " and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code=TSPL_Location_MASTER.Location_Code ") + " 
            where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' " + IIf(clsCommon.myLen(txtMCC.Value) > 0, "and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code='" + txtMCC.Value + "' And TSPL_MCC_MASTER.MCC_Code='" + txtMCC.Value + "' ", "") + "   and TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount>0 "

            End If

            If clsCommon.myLen(txtBankGroup.Value) > 0 Then
                BaseQry = BaseQry + " AND TSPL_BANK_MASTER.BANK_GROUP_CODE='" + txtBankGroup.Value + "' "
            End If
            Dim FinalQuery As String = ""

            If rbtnBankAdvice.IsChecked OrElse rbtnCompulsory.IsChecked OrElse rbtnSaving.IsChecked Then
                FinalQuery = BaseQry + " order by TSPL_Vendor_MASTER.Bank_Code"
            ElseIf rbtnBankWiseSummary.IsChecked Then
                If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "CHU") = CompairStringResult.Equal Then
                    FinalQuery += "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select max(CycleRange) as CycleRange, max(GRPColumn) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,max(Branch_Name)Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,sum(Payable_Amount) as Payable_Amount
            ,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
            from (" + Environment.NewLine + BaseQry + Environment.NewLine + " )xxx group by Bank_Code )xxxx order by GRPColumn "
                Else
                    FinalQuery += "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select max(CycleRange) as CycleRange, max(GRPColumn) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,sum(Payable_Amount) as Payable_Amount
            ,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
            from (" + Environment.NewLine + BaseQry + Environment.NewLine + " )xxx group by Bank_Code,Branch_Name )xxxx order by GRPColumn "
                End If

            ElseIf rbtnCurrentBankWiseSummary.IsChecked OrElse rbtnCompulsoryWiseSummary.IsChecked Then
                FinalQuery += "select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select  max(GRPColumn) as GRPColumn,sum(Payable_Amount) as Payable_Amount , Bank_Desc
            from (" + Environment.NewLine + BaseQry + Environment.NewLine + " )xxx group by Bank_Desc )xxxx order by GRPColumn "
            End If


            Dim dt As DataTable = clsDBFuncationality.GetDataTable(FinalQuery)


            If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Display", Me.Text)
                Exit Sub

            End If


            Dim frmCRV As New frmCrystalReportViewer()
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "GNG") = CompairStringResult.Equal Then
            'If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "JPR") = CompairStringResult.Equal Then
            'frmCRV.funreport(CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewJPR", "Bank Advice")
            ' Else
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "BKN") = CompairStringResult.Equal Then
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNewwBKN", "Bank Advice New")
            Else
                frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankAdviceNeww", "Bank Advice New")
            End If

            'End If


        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    Private Sub fndArea__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndArea._MYValidating
        Try
            Dim sQuery As String = " Select TSPL_LOCATION_MASTER.Location_Code as Code ,  TSPL_LOCATION_MASTER.Location_Desc, Type from TSPL_LOCATION_MASTER "
            fndArea.Value = clsCommon.ShowSelectForm("Location@Plant@Master", sQuery, "Code", "TSPL_LOCATION_MASTER.Type <> 'PLANT' OR TSPL_LOCATION_MASTER.Location_Category <> 'Mcc'", fndArea.Value, "Code", isButtonClicked)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.ToString, Me.Text)
        End Try
    End Sub

    Private Sub rbtnBothSavCur_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles rbtnBothSavCur.ToggleStateChanged
        'If Not flagBothSavingCurrent Then
        '    flagSavingCurrent = True
        '    rbtnSaving.IsChecked = False
        '    rbtnSavingCompulsory.IsChecked = False
        '    rbtnSavingBankWiseSummary.IsChecked = False
        '    rbtnBankAdvice.IsChecked = False
        '    rbtnBankWiseSummary.IsChecked = False
        '    rbtnCurrentBankWiseSummary.IsChecked = False
        '    flagSavingCurrent = False
        'End If
        If rbtnBothSavCur.IsChecked Then
            rbtnSaving.IsChecked = False
            rbtnCompulsory.IsChecked = False
            rbtnCompulsoryWiseSummary.IsChecked = False
            rbtnBankAdvice.IsChecked = False
            rbtnBankWiseSummary.IsChecked = False
            rbtnCurrentBankWiseSummary.IsChecked = False
        End If
    End Sub

    Private Sub btnPrintSWM_Click(sender As Object, e As EventArgs) Handles btnPrintSWM.Click
        If ChkIFSCCode.Checked Then
            Printt()
        Else
            Print(True, Nothing, Nothing)
        End If
    End Sub

    Private Sub btnExportBankWise_Click(sender As Object, e As EventArgs) Handles btnExportBankWise.Click
        Try
            Dim dt As DataTable = Nothing
            Dim dtBank As DataTable = Nothing

            If Gv1.Rows.Count <= 0 Then
                clsCommon.MyMessageBoxShow(Me, "No Data Found to Export", Me.Text)
                Exit Sub
            End If
            If rbtnSaving.IsChecked OrElse rbtnBankAdvice.IsChecked Then
                dt = New DataTable()
                dtBank = New DataTable()
                dt = Gv1.DataSource
                dtBank = dt.DefaultView.ToTable(True, "GRPColumn")

                If dtBank.Rows.Count > 0 Then
                    For Each dr In dtBank.Rows
                        Gv1.MasterTemplate.FilterDescriptors.Clear()
                        Dim filter As New FilterDescriptor("GRPColumn", FilterOperator.Contains, dr("GRPColumn"))
                        Gv1.MasterTemplate.FilterDescriptors.Add(filter)

                        Dim strHeading As String = ""
                        If MultipleFinderFillAuto Then
                            strHeading = clsCommon.myCstr("BankAdvice Report Bank Wise from " + clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") + " to " + clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy"))
                        End If

                        Dim arrHeader As List(Of String) = New List(Of String)()
                        If MultipleFinderFillAuto Then
                            If rbtnSaving.IsChecked = True Then
                                arrHeader.Add("Data: Saving ")
                            End If
                            If rbtnBankAdvice.IsChecked = True Then
                                arrHeader.Add("Report Type: Bank Advice")
                            End If
                        Else
                            arrHeader.Add("Mcc : " + txtMCC.Value)
                        End If

                        clsCommon.MyExportToExcelGrid(strHeading, Gv1, arrHeader, "Bank Advise  BankWise")
                    Next

                End If
            End If
            clsCommon.MyMessageBoxShow(Me, "Export Successfully", Me.Text)
            Gv1.MasterTemplate.FilterDescriptors.Clear()

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnBankSummary_Click(sender As Object, e As EventArgs) Handles btnBankSummary.Click
        Try
            Dim qry As String = Nothing
            If rbtnSavingSummary.IsChecked Then
                qry = " select ROW_NUMBER() over ( order by Bank_Code) as SNO , max(CycleRange) as CycleRange, max(Bank_Code) as GRPColumn,max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,max (Bank_Code_Desc)Branch_Name,max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,sum(Payable_Amount) as Payable_Amount
                        ,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
                        from (
                         select  '' AS CycleRange,TSPL_Vendor_MASTER.BankCode2 as GRPColumn,TSPL_COMPANY_MASTER.Comp_Name
                        ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address, case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No,TSPL_MCC_MASTER.MCC_NAME, TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No,TSPL_Fiscal_Year_Master.Fiscal_Name
                        ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as VLC_CODE_Uploader, coalesce(TSPL_VENDOR_MASTER.VSP_Payee_Name,Mp_V.VSP_Payee_Name)  as Payee_Joint_Name,TSPL_Vendor_MASTER.BankCode2 as Bank_Code, 
                          TSPL_VENDOR_MASTER.BankBranch2 as Branch_Name, 
                          TSPL_BANK_MASTER2.DESCRIPTION as Bank_Code_Desc, 
                          case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment, Mp_V.vsp_payment), '')= 'Self' then coalesce(TSPL_VENDOR_MASTER.IFSCCode2, mp_V.IFSC_Code) else coalesce(
                            TSPL_VENDOR_MASTER.Joint_IFSC_Code,mp_v.Joint_IFSC_Code) end as Payee_Joint_IFSC_Code, 
                          case when isnull(coalesce(TSPL_VENDOR_MASTER.vsp_payment,Mp_V.vsp_payment),'')= 'Self' then coalesce(TSPL_VENDOR_MASTER.AccNo2, mp_V.Account_No) else coalesce(
                            TSPL_VENDOR_MASTER.Joint_Account_No, mp_V.Joint_Account_No) end as Payee_Joint_Account_No, TSPL_VENDOR_INVOICE_HEAD.Document_Total as Payable_Amount 
	                        from TSPL_PAYMENT_PROCESS_SAVING 
                            left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_SAVING.Doc_No
                            left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_PAYMENT_PROCESS_SAVING.AP_Invoice_No
                            left outer join TSPL_MP_MASTER mp on mp.MP_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code 
                            Left outer join tspl_vendor_master Mp_V on mp_V.Vendor_Code=mp.MP_Code  
                            left outer join ( select distinct VSP_Code, VLC_Code_VLC_Uploader  from TSPL_VLC_MASTER_HEAD) as TSPL_VLC_MASTER_HEAD on TSPL_VLC_MASTER_HEAD.VSP_Code =  TSPL_VENDOR_INVOICE_HEAD.Vendor_Code
                            left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP'
                            left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code= TSPL_VENDOR_INVOICE_HEAD.Vendor_Code --TSPL_PAYMENT_PROCESS_SAVING.VSP_CODE
                            left outer join TSPL_BANK_MASTER as TSPL_BANK_MASTER2 on TSPL_BANK_MASTER2.BANK_CODE=TSPL_Vendor_MASTER.BankCode2 
                            left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date    
                            left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   
                            left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Bank_Code
                            left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected
                                where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and TSPL_PAYMENT_PROCESS_HEAD.From_Date>=  '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' 
                                and	TSPL_PAYMENT_PROCESS_HEAD.To_Date <=  '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and TSPL_VENDOR_INVOICE_HEAD.Document_Total>0  
                               )xxx group by Bank_Code  "

                Dim dt1 As DataTable = clsDBFuncationality.GetDataTable(qry)

                If dt1 IsNot Nothing And dt1.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt1, "crptBankSavingSummary", "Bank Saving Summary")
                    frmCRV = Nothing
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
                End If

            Else
                qry = " select ROW_NUMBER() over ( order by GRPColumn) as SNO , * from ( select max(CycleRange) as CycleRange, max(GRPColumn) as GRPColumn,
                                 max(Comp_Name) as Comp_Name,max(Comp_address) as Comp_address, max(From_Date) as From_Date,max(GSTReg_No) as GSTReg_No,
                                 max(Fiscal_Name) as Fiscal_Name,max(CycleNo) as CycleNo,max(Date_Range) as Date_Range,Bank_Code,max(Branch_Name)Branch_Name,
                                 max(Bank_Code_Desc) as Bank_Code_Desc, max (Payee_Joint_IFSC_Code) as Payee_Joint_IFSC_Code,max(Payee_Joint_Account_No) as Payee_Joint_Account_No ,
                                 sum(Payable_Amount) as Payable_Amount,max(CompPhone) as CompPhone,max(Regn_No) as Regn_No,max(MCC_NAME) as MCC_NAME
                                from (select  '' AS CycleRange, TSPL_Vendor_MASTER.Bank_Code as GRPColumn, CASE WHEN TSPL_Vendor_MASTER.Bank_Code LIKE 'PNB%' THEN 'PNB Bank' ELSE 'Other Banks' END AS GRPColumns,TSPL_COMPANY_MASTER.Bank_Name,TSPL_COMPANY_MASTER.BankAccountNo,TSPL_COMPANY_MASTER.BankIFSCCode,TSPL_COMPANY_MASTER.BankBranchAddress,
                               TSPL_BANK_MASTER.DESCRIPTION as [Company Bank], TSPL_BANK_MASTER.BANKACCNUMBER as [Company Bank Account No],
                               TSPL_COMPANY_MASTER.Comp_Name
                               ,TSPL_COMPANY_MASTER.add1 +case when len(TSPL_COMPANY_MASTER.add2)>0 then ', '+TSPL_COMPANY_MASTER.add2 else '' end +case when LEN(isnull(TSPL_COMPANY_MASTER.Add3,''))>0 then ', '+isnull(TSPL_COMPANY_MASTER.Add3,'') else ' ' end  + case when len(TSPL_COMPANY_MASTER.State )>0 then TSPL_COMPANY_MASTER.State else '' end as Comp_address
                               ,case when ISNULL(TSPL_COMPANY_MASTER.Phone1,'')='(+__)__________' then '' else TSPL_COMPANY_MASTER.Phone1 end +  Case When ISNULL (TSPL_COMPANY_MASTER.Phone2,'')<>'(+__)__________' Then ', '+ TSPL_COMPANY_MASTER.Phone2 Else'' End as CompPhone ,TSPL_COMPANY_MASTER.Regn_No, TSPL_MCC_MASTER.MCC_NAME ,TSPL_PAYMENT_PROCESS_HEAD.From_Date,'GSTIN : '+ TSPL_COMPANY_MASTER.GSTReg_No as GSTReg_No,TSPL_PAYMENT_PROCESS_HEAD.Doc_No, TSPL_Fiscal_Year_Master.Fiscal_Name
                                ,TSPL_PAYMENT_CYCLE_GENERATED.Name as CycleNo ,convert(varchar, TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) +' To '+ convert(varchar,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103) as Date_Range, TSPL_PAYMENT_PROCESS_DETAIL.VLC_CODE_Uploader,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Name,TSPL_Vendor_MASTER.Bank_Code,TSPL_VENDOR_MASTER.Branch_Name,case when isnull(TSPL_Vendor_MASTER.Bank_Name,'')  = '' then  TSPL_Vendor_MASTER.Bank_Code else TSPL_Vendor_MASTER.Bank_Name end as Bank_Code_Desc,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_IFSC_Code,TSPL_PAYMENT_PROCESS_DETAIL.Payee_Joint_Account_No, Round((isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0)-isnull(TSPL_TRANSFER_TO_SAVING_DETAIL.Amount,0)),0) as Payable_Amount  ,Case When TSPL_BANK_ADVISE.Status IS NULL OR TSPL_BANK_ADVISE.Status =0 Then 'Pending' Else 'Approved' End As [Bank Advice Status]  from TSPL_PAYMENT_PROCESS_DETAIL 
                                left outer join TSPL_PAYMENT_PROCESS_HEAD on TSPL_PAYMENT_PROCESS_HEAD.Doc_No=TSPL_PAYMENT_PROCESS_DETAIL.Doc_No
                                left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code='UDP'
                                left outer join TSPL_Vendor_MASTER on TSPL_Vendor_MASTER.Vendor_Code=TSPL_PAYMENT_PROCESS_DETAIL.VSP_CODE
                                left outer join TSPL_Fiscal_Year_Master on TSPL_Fiscal_Year_Master.Start_Date<=TSPL_PAYMENT_PROCESS_HEAD.From_Date and TSPL_Fiscal_Year_Master.End_Date>=TSPL_PAYMENT_PROCESS_HEAD.From_Date
                                left outer join TSPL_BANK_MASTER ON TSPL_BANK_MASTER.BANK_CODE = TSPL_Vendor_MASTER.Company_Bank_Current  left outer join TSPL_MCC_MASTER on TSPL_MCC_MASTER.MCC_Code=TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected left outer join TSPL_TRANSFER_TO_SAVING_DETAIL  on TSPL_PAYMENT_PROCESS_DETAIL.VSP_Code = TSPL_TRANSFER_TO_SAVING_DETAIL.Vendor_Code 
                                left outer join TSPL_BANK_ADVISE On TSPL_BANK_ADVISE.Payment_Process_Document_No=TSPL_PAYMENT_PROCESS_HEAD.Doc_No     
                                left outer join TSPL_PAYMENT_CYCLE_GENERATED on convert(date, TSPL_PAYMENT_CYCLE_GENERATED.From_Date,103)<=convert(date,TSPL_PAYMENT_PROCESS_HEAD.From_Date,103) and convert(date,TSPL_PAYMENT_CYCLE_GENERATED.To_Date,103)>=convert(date,TSPL_PAYMENT_PROCESS_HEAD.To_Date,103)   and TSPL_PAYMENT_CYCLE_GENERATED.MCC_Code = TSPL_PAYMENT_PROCESS_HEAD.MCC_Code_Selected   where TSPL_PAYMENT_PROCESS_HEAD.isPrePosted = 1 and  TSPL_PAYMENT_PROCESS_HEAD.From_Date>= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(fromDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "' and	TSPL_PAYMENT_PROCESS_HEAD.To_Date<= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate.Value), "dd/MMM/yyyy hh:mm:ss tt") + "'   And (isnull(TSPL_PAYMENT_PROCESS_DETAIL.Payable_Amount,0)-isnull(TSPL_PAYMENT_PROCESS_DETAIL.Compulsory_Amount,0))>0
                                )xxx group by Bank_Code )xxxx order by GRPColumn  "

                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)

                If dt IsNot Nothing And dt.Rows.Count > 0 Then
                    Dim frmCRV As New frmCrystalReportViewer()
                    frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.MilkProcurement, dt, "crptBankDraftSummary", "Bank Draft Summary")
                    frmCRV = Nothing
                Else
                    clsCommon.MyMessageBoxShow(Me, "No Data Found", Me.Text)
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtDCS__My_Click(sender As Object, e As EventArgs) Handles txtDCS._My_Click
        Try
            'If txtRoute.arrValueMember Is Nothing OrElse txtRoute.arrValueMember.Count <= 0 Then
            '    txtRoute.Focus()
            '    Throw New Exception("Please select at least route")
            'End If
            Dim qry As String = " select  TSPL_VLC_MASTER_HEAD.VLC_Code_VLC_Uploader as [DCS Code] ,TSPL_VLC_MASTER_HEAD.VLC_Name as [DCS Name],TSPL_VLC_MASTER_HEAD.VSP_Code,isnull(TSPL_VENDOR_MASTER.Zone_Code,'') AS Zone
		                          from TSPL_VLC_MASTER_HEAD 
		                          left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = TSPL_VLC_MASTER_HEAD.VSP_Code"
            'If txtRoute.arrValueMember IsNot Nothing AndAlso txtRoute.arrValueMember.Count > 0 Then
            '    qry += " and TSPL_VLC_MASTER_HEAD.Route_Code in (" + clsCommon.GetMulcallString(txtRoute.arrValueMember) + ") "
            'End If

            txtDCS.arrValueMember = clsCommon.ShowMultipleSelectForm("PCUVLC1", qry, "VSP_Code", "DCS Name", txtDCS.arrValueMember, Nothing)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub txtBankAdvise__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtBankAdviseNo._MYValidating
        Try
            Dim qry As String = "select Document_No as DocNo,convert(varchar,Document_date,103) as DocDate,Payment_Process_Document_No as [Payment Process Document_No] from TSPL_BANK_ADVISE left join TSPL_PAYMENT_PROCESS_HEAD ON TSPL_PAYMENT_PROCESS_HEAD.Doc_No =TSPL_BANK_ADVISE.Payment_Process_Document_No "
            Dim strWhrcls As String = " TSPL_PAYMENT_PROCESS_HEAD.From_Date >= '" & clsCommon.GetPrintDate(fromDate.Value, "dd/MMM/yyyy") & "' and TSPL_PAYMENT_PROCESS_HEAD.To_Date <= '" & clsCommon.GetPrintDate(ToDate.Value, "dd/MMM/yyyy") & "' and TSPL_BANK_ADVISE.Status = 1 "
            txtBankAdviseNo.Value = clsCommon.ShowSelectForm("BnkAdviceNo", qry, "DocNo", strWhrcls, txtBankAdviseNo.Value, "DocNo", isButtonClicked)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text, MessageBoxButtons.OK)
        End Try
    End Sub
End Class
