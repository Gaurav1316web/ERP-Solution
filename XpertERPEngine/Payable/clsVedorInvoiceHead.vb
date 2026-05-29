'--Updations By [Pankaj Kumar Chaudhary]-Against Ticket No -[BM00000001707],BM00000007903,BM00000007899
Imports common
Imports System.Data.SqlClient
Public Class clsVedorInvoiceHead
#Region "Variables"
    Public Vendor_Bank_Name As String = Nothing
    Public Vendor_Bank_Code As String = Nothing
    Public Branch_IFSC_Code As String = Nothing
    Public Branch_Name As String = Nothing
    Public Against_MCC_Material_Sale As String = Nothing
    Public Against_MCC_Material_Sale_Return As String = Nothing
    Public Vendor_Bank_ACNo As String = Nothing
    Public Addition_Doc_Type As String = Nothing
    Public PROJECT_ID As String = Nothing
    Public Document_No As String = Nothing
    Public Invoice_Entry_Date As String = Nothing
    Public Invoice_Entry_Date_New As DateTime = Nothing
    Public Vendor_Code As String = Nothing
    Public Vendor_Name As String = Nothing
    Public DateAndTime As DateTime?
    Public TapalNo As String = String.Empty
    Public Vendor_Invoice_No As String = Nothing
    Public Vendor_Invoice_Date As String = Nothing
    Public Posting_Date As String = Nothing
    Public Account_Set As String = Nothing
    Public Document_Type As String = Nothing
    Public GSTRegistered As Integer = 0
    Public PO_Number As String = Nothing
    Public Order_No As String = Nothing
    Public Document_Total As Double = 0
    Public On_Hold As Boolean = Nothing
    Public Remarks As String = Nothing
    Public Description As String = Nothing
    Public Tax_Group As String = Nothing
    Public RefDocType As String = Nothing
    Public RefDocNo As String = Nothing
    Public Ref_SNo As Integer = Nothing
    Public Against_MillkPurchaseInvoice_No As String = Nothing
    Public Against_VSPItemIssue_No As String = Nothing
    Public Against_VSP_Asset_Issue As String = Nothing
    Public Against_BulkMillkPurchaseInvoice_No As String = Nothing
    Public Total_Taxable_Amount As Decimal
    Public TAX1 As String = Nothing
    Public TAX1_Rate As Double = 0
    Public Tax1_BAmount As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Rate As Double = 0
    Public Tax2_BAmount As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Rate As Double = 0
    Public Tax3_BAmount As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Rate As Double = 0
    Public Tax4_BAmount As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Rate As Double = 0
    Public Tax5_BAmount As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Rate As Double = 0
    Public Tax6_BAmount As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Rate As Double = 0
    Public Tax7_BAmount As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Rate As Double = 0
    Public Tax8_BAmount As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Rate As Double = 0
    Public Tax9_BAmount As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Rate As Double = 0
    Public Tax10_BAmount As Double = 0
    Public TAX10_Amt As Double = 0
    Public Total_Tax As Double = 0
    Public Terms_Code As String = Nothing
    Public Terms_Description As String = Nothing
    Public Due_Date As String = Nothing
    Public VI_Due_Date As String = Nothing
    Public Discount_Percentage As Double = 0
    Public Discount_Base As Double = 0
    Public Discount_Amount As Double = 0
    Public Amount_Less_Discount As Double = 0
    Public TDS_Base_Actual_Amount As Double = 0
    Public TDS_Base_Calculated_Amount As Double = 0
    Public TDS_Percentage As Double = 0
    Public TDS_Actual_Amount As Double = 0
    Public TDS_Calculated_Amount As Double = 0
    Public Nature_of_deduction As String = Nothing
    Public Branch_Code As String = Nothing
    Public Section_Code As String = Nothing
    'Public Created_By As String = Nothing
    'Public Created_Date As String = Nothing
    'Public Modify_By As String = Nothing
    'Public Modify_Date As String = Nothing
    Public Comp_Code As String = Nothing
    Public Balance_Amt As Double = 0
    Public Arr As List(Of clsVedorInvoiceDetail) = Nothing
    Public RemittanceObject As clsRemittance = Nothing

    Public Vendor_Control_AC As String = Nothing
    Public Discount_GL_AC As String = Nothing
    Public TAX1_GLAC As String = Nothing
    Public TAX2_GLAC As String = Nothing
    Public TAX3_GLAC As String = Nothing
    Public TAX4_GLAC As String = Nothing
    Public TAX5_GLAC As String = Nothing
    Public TAX6_GLAC As String = Nothing
    Public TAX7_GLAC As String = Nothing
    Public TAX8_GLAC As String = Nothing
    Public TAX9_GLAC As String = Nothing
    Public TAX10_GLAC As String = Nothing
    Public Against_POInvoice_No As String = Nothing
    Public Against_PurchaseReturn_No As String = Nothing
    Public MCC_Code As String = String.Empty
    Public MCC_Name As String = String.Empty
    Public Against_Acquisition As String = Nothing

    Public Add_Charge_Code1 As String = Nothing
    Public Add_Charge_Name1 As String = Nothing
    Public Add_Charge_Amt1 As Double = 0
    Public Add_Charge_Code2 As String = Nothing
    Public Add_Charge_Name2 As String = Nothing
    Public Add_Charge_Amt2 As Double = 0
    Public Add_Charge_Code3 As String = Nothing
    Public Add_Charge_Name3 As String = Nothing
    Public Add_Charge_Amt3 As Double = 0
    Public Add_Charge_Code4 As String = Nothing
    Public Add_Charge_Name4 As String = Nothing
    Public Add_Charge_Amt4 As Double = 0
    Public Add_Charge_Code5 As String = Nothing
    Public Add_Charge_Name5 As String = Nothing
    Public Add_Charge_Amt5 As Double = 0
    Public Add_Charge_Code6 As String = Nothing
    Public Add_Charge_Name6 As String = Nothing
    Public Add_Charge_Amt6 As Double = 0
    Public Add_Charge_Code7 As String = Nothing
    Public Add_Charge_Name7 As String = Nothing
    Public Add_Charge_Amt7 As Double = 0
    Public Add_Charge_Code8 As String = Nothing
    Public Add_Charge_Name8 As String = Nothing
    Public Add_Charge_Amt8 As Double = 0
    Public Add_Charge_Code9 As String = Nothing
    Public Add_Charge_Name9 As String = Nothing
    Public Add_Charge_Amt9 As Double = 0
    Public Add_Charge_Code10 As String = Nothing
    Public Add_Charge_Name10 As String = Nothing
    Public Add_Charge_Amt10 As Double = 0
    Public Total_Add_Charge As Double = 0
    Public loc_code As String = Nothing
    Public Location_Code_Prefix As String = Nothing
    Public Irregular_loc_code As String = Nothing
    Public Tax_Calculation_Type As EnumTaxCalucationType

    Public TAX1_GLAC_Amt As Double = 0
    Public TAX1_GLAC2 As String = Nothing
    Public TAX1_GLAC2_Amt As Double = 0
    Public TAX1_GLAC3 As String = Nothing
    Public TAX1_GLAC3_Amt As Double = 0
    Public TAX1_GLAC4 As String = Nothing
    Public TAX1_GLAC4_Amt As Double = 0
    Public TAX1_GLAC5 As String = Nothing
    Public TAX1_GLAC5_Amt As Double = 0

    Public irn_no As String = Nothing
    Public Ack_No As String = Nothing
    Public QR_Code As String = Nothing
    Public Ack_Date As DateTime?

    Public TAX2_GLAC_Amt As Double = 0
    Public TAX2_GLAC2 As String = Nothing
    Public TAX2_GLAC2_Amt As Double = 0
    Public TAX2_GLAC3 As String = Nothing
    Public TAX2_GLAC3_Amt As Double = 0
    Public TAX2_GLAC4 As String = Nothing
    Public TAX2_GLAC4_Amt As Double = 0
    Public TAX2_GLAC5 As String = Nothing
    Public TAX2_GLAC5_Amt As Double = 0

    Public TAX3_GLAC_Amt As Double = 0
    Public TAX3_GLAC2 As String = Nothing
    Public TAX3_GLAC2_Amt As Double = 0
    Public TAX3_GLAC3 As String = Nothing
    Public TAX3_GLAC3_Amt As Double = 0
    Public TAX3_GLAC4 As String = Nothing
    Public TAX3_GLAC4_Amt As Double = 0
    Public TAX3_GLAC5 As String = Nothing
    Public TAX3_GLAC5_Amt As Double = 0

    Public TAX4_GLAC_Amt As Double = 0
    Public TAX4_GLAC2 As String = Nothing
    Public TAX4_GLAC2_Amt As Double = 0
    Public TAX4_GLAC3 As String = Nothing
    Public TAX4_GLAC3_Amt As Double = 0
    Public TAX4_GLAC4 As String = Nothing
    Public TAX4_GLAC4_Amt As Double = 0
    Public TAX4_GLAC5 As String = Nothing
    Public TAX4_GLAC5_Amt As Double = 0

    Public TAX5_GLAC_Amt As Double = 0
    Public TAX5_GLAC2 As String = Nothing
    Public TAX5_GLAC2_Amt As Double = 0
    Public TAX5_GLAC3 As String = Nothing
    Public TAX5_GLAC3_Amt As Double = 0
    Public TAX5_GLAC4 As String = Nothing
    Public TAX5_GLAC4_Amt As Double = 0
    Public TAX5_GLAC5 As String = Nothing
    Public TAX5_GLAC5_Amt As Double = 0

    Public TAX6_GLAC_Amt As Double = 0
    Public TAX6_GLAC2 As String = Nothing
    Public TAX6_GLAC2_Amt As Double = 0
    Public TAX6_GLAC3 As String = Nothing
    Public TAX6_GLAC3_Amt As Double = 0
    Public TAX6_GLAC4 As String = Nothing
    Public TAX6_GLAC4_Amt As Double = 0
    Public TAX6_GLAC5 As String = Nothing
    Public TAX6_GLAC5_Amt As Double = 0

    Public TAX7_GLAC_Amt As Double = 0
    Public TAX7_GLAC2 As String = Nothing
    Public TAX7_GLAC2_Amt As Double = 0
    Public TAX7_GLAC3 As String = Nothing
    Public TAX7_GLAC3_Amt As Double = 0
    Public TAX7_GLAC4 As String = Nothing
    Public TAX7_GLAC4_Amt As Double = 0
    Public TAX7_GLAC5 As String = Nothing
    Public TAX7_GLAC5_Amt As Double = 0

    Public TAX8_GLAC_Amt As Double = 0
    Public TAX8_GLAC2 As String = Nothing
    Public TAX8_GLAC2_Amt As Double = 0
    Public TAX8_GLAC3 As String = Nothing
    Public TAX8_GLAC3_Amt As Double = 0
    Public TAX8_GLAC4 As String = Nothing
    Public TAX8_GLAC4_Amt As Double = 0
    Public TAX8_GLAC5 As String = Nothing
    Public TAX8_GLAC5_Amt As Double = 0

    Public TAX9_GLAC_Amt As Double = 0
    Public TAX9_GLAC2 As String = Nothing
    Public TAX9_GLAC2_Amt As Double = 0
    Public TAX9_GLAC3 As String = Nothing
    Public TAX9_GLAC3_Amt As Double = 0
    Public TAX9_GLAC4 As String = Nothing
    Public TAX9_GLAC4_Amt As Double = 0
    Public TAX9_GLAC5 As String = Nothing
    Public TAX9_GLAC5_Amt As Double = 0

    Public TAX10_GLAC_Amt As Double = 0
    Public TAX10_GLAC2 As String = Nothing
    Public TAX10_GLAC2_Amt As Double = 0
    Public TAX10_GLAC3 As String = Nothing
    Public TAX10_GLAC3_Amt As Double = 0
    Public TAX10_GLAC4 As String = Nothing
    Public TAX10_GLAC4_Amt As Double = 0
    Public TAX10_GLAC5 As String = Nothing
    Public TAX10_GLAC5_Amt As Double = 0
    Public Empty_Amount As Double = 0
    Public Empty_Account As String = Nothing
    Public Total_Landed_Amt As Double = 0
    Public is_For_TDS As Integer = 0

    Public CURRENCY_CODE As String = ""
    Public ConvRate As Decimal
    Public ApplicableFrom As Date? = Nothing
    Public Is_ProRated As Char = "N"
    Public Form_ID As String = ""
    Public arrCustomFields As List(Of clsCustomFieldValues) = Nothing
    Public Invoice_Type As String = Nothing

    Public RoundOffAmount As Double = 0
    Public ISProcurementDeduction As Integer = 0
    Public is_For_Provision As Integer = 0
    Public isFreightProvisionAccount As Boolean = False
    Public Prov_From_Date As Date
    Public Prov_To_Date As Date
    Public Prov_Amt As Double = 0
    Public arrProvDocNo As List(Of String) = Nothing
    Public Security As Integer
    Public isDeduction As Integer
    Public is_Secondary_Transporter_Deduction As Boolean = False
    Public Against_Asset_Work As String = ""
    Public Against_TransferToSavingPKID As Integer
    Public Against_VCGL As String = Nothing
    Public Hirerachy_Level_Code As String = String.Empty
    Public Cost_Centre_Fin_Level_Code As String = String.Empty
    Public Employee_Type As String = String.Empty
    '=========added by shivani tyagi
    '======================
    Public Cash_Discount_Amt As Double = 0
    Public Amt_After_Cash_Discount As Double = 0
    Public ServiceBill_No As String = Nothing
    Public ServiceBill_Date As String = Nothing
    '===============================================
    Public ArrAssetEMI As List(Of clsAPInvoiceAssetEMIDetails)
    Public ArrAdvanceInterest As List(Of clsAPInvoiceAdvanceInterest)
    Public RCM As Boolean = False
    Public TDS_Provision As Boolean = False
    Public Transfer_To_Saving As Boolean = False
    Public IsEInvoice As Boolean = False


    Public No_GST_Credit As Boolean = False
    Public Purchase_Tax_Invoice As String
    Public Purchase_Tax_Invoice_Type As String
    Public LckType As String = Nothing
    Public Against_Salary_Generation_Code As String = String.Empty
    Public ITC_Elibible As Integer = 0
    Public ITC_Type As String = Nothing
    Public ITC_Type_Category As String = Nothing
    '=========Added by preeti Gupta Against ticket no[22/02/2019]
    Public arrInvoiceNo As List(Of String) = Nothing

    Public Main_VSP_Milk_AP_Invoice_No As String = Nothing
    '===============
    Public Saving As Integer = 0
    Public arrSecurityAdjustment As ArrayList = Nothing
#End Region

    Public Shared Function funAPInvoiceEntryPrint(ByVal Form_ID As String, ByVal strCancelDelete As String, ByVal strDate As DateTime, ByVal StrDocNo As String, ByVal SettingCostCenterlevel As Boolean) As Boolean
        Try
            Dim TSPL_VENDOR_INVOICE_HEAD As String = Nothing
            Dim TSPL_VENDOR_INVOICE_DETAIL As String = Nothing
            Dim TSPL_JOURNAL_DETAILS As String = Nothing
            Dim TSPL_JOURNAL_MASTER As String = Nothing
            If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal Then
                TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD_Cancel_data"
                TSPL_VENDOR_INVOICE_DETAIL = "TSPL_VENDOR_INVOICE_detail_cancel_data"
                TSPL_JOURNAL_DETAILS = "TSPL_JOURNAL_DETAILS_Cancel_Data"
                TSPL_JOURNAL_MASTER = "TSPL_JOURNAL_MASTER_Cancel_Data"
            Else
                TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD"
                TSPL_VENDOR_INVOICE_DETAIL = "TSPL_VENDOR_INVOICE_detail"
                TSPL_JOURNAL_DETAILS = "TSPL_JOURNAL_DETAILS"
                TSPL_JOURNAL_MASTER = "TSPL_JOURNAL_MASTER"
            End If




            Dim qry As String = "select"
            If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal OrElse clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
                qry += " 'Cancelled' As Report_Status, "
            Else
                qry += " '' As Report_Status, "
            End If

            qry += "    isnull (TBL_Remarks.Rmarks,'') as AccCodeRemarks, '' as SSC_Code,tspl_state_master_for_location_state.GST_STATE_Code as LOC_GST_State_Code, TSPL_LOCATION_MASTER.GSTNO as Loc_GstInNo ,TSPL_VENDOR_MASTER.GSTFinalNo AS Vendor_GSTIN_NO,TSPL_STATE_MASTER.GST_STATE_Code AS Vendor_GST_StateCode,xxxx.RefDocNo ,xxxx.RefDocType ,xxxx.Document_No ,xxxx.ACCode ,xxxx.ACName ,xxxx.DrAmt ,xxxx.CrAmt ,xxxx.Comp_Name ,xxxx.Vendor_Code, TSPL_VENDOR_MASTER.Vendor_Name,xxxx.Invoice_Entry_Date ,xxxx.Due_Date ,xxxx.Description ,xxxx.Vendor_Invoice_Date ,xxxx.CreateBy ,xxxx.ApproveBy ,xxxx.InvStatus ,xxxx.InvoiceType ,xxxx.Vendor_Invoice_No , '' as Remarks,TSPL_LOCATION_MASTER.Location_Desc ,TSPL_LOCATION_MASTER.Add1 ,TSPL_LOCATION_MASTER.Add2 ,TSPL_LOCATION_MASTER.Add3 ,TSPL_LOCATION_MASTER.Add4 ,XXXX.Document_Type,TSPL_COMPANY_MASTER .Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,XXXX.TapalNo,XXXX.DateAndTime,XXXX.Vendor_Bank_Code,XXXX.Vendor_Bank_Name,XXXX.Branch_IFSC_Code,XXXX.Branch_Name,XXXX.Vendor_Bank_ACNo from (select MAX( FromDate) as FromDate,max(ToDate) as ToDate,max(Location) as Location,max(Document_Type) as Document_Type,max(Loc_Code) as Loc_Code,max(Vendor) as Vendor,max(Document) as Document,max(Document_No) as Document_No," & Environment.NewLine &
  " max(ACCode) as ACCode,max(ACName) as ACName, case when SUM(DrAmt-CrAmt)>0 then sum(DrAmt-CrAmt) else 0 end   as DrAmt ,case when SUM( CrAmt-DrAmt)>0 then SUM( CrAmt-DrAmt) else 0 end   as CrAmt,max(Comp_Code) as Comp_Code,max(Comp_Name) as Comp_Name,max(Vendor_Code) as Vendor_Code,max(Invoice_Entry_Date) as Invoice_Entry_Date,max(RefDocNo) as RefDocNo,max(RefDocType) as RefDocType,max(Vendor_Name) as Vendor_NameOLD,max(Created_By) as Created_By,max(Created_Date) as Created_Date,max(Due_Date) as Due_Date,max(Description) as Description,max(Vendor_Invoice_No) as Vendor_Invoice_No,max(Vendor_Invoice_Date) as Vendor_Invoice_Date,max(CreateBy) as CreateBy,max(ApproveBy) as ApproveBy,max(InvStatus) as InvStatus,max(InvoiceType) as InvoiceType,max(RefDocDescription) as RefDocDescription,max(Hirerachy_Code) as Hirerachy_Code ,max(Cost_Centre_Code) as Cost_Centre_Code,MAX(Hirerachy_Code3) AS Hirerachy_Code3,MAX(Hirerachy_Code4) AS Hirerachy_Code4,MAX(TapalNo) AS TapalNo,MAX(DateAndTime) AS DateAndTime,MAX(Vendor_Bank_Code) as Vendor_Bank_Code,max(Vendor_Bank_Name) as Vendor_Bank_Name,max(Branch_IFSC_Code) as Branch_IFSC_Code,max(Branch_Name) as Branch_Name,max(Vendor_Bank_ACNo) as Vendor_Bank_ACNo  from(select  '' as FromDate,'' as ToDate, '' as Location," + TSPL_VENDOR_INVOICE_HEAD + ".Loc_Code," + TSPL_VENDOR_INVOICE_HEAD + ".Document_Type,'' as Vendor, '" & StrDocNo & "' as Document, Final.Document_No,final.ACCode,Final.ACName,DrAmt,CrAmt,Hirerachy_Code,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name as Cost_Centre_Code,Hirerachy_Code3,Hirerachy_Code4  ,TSPL_COMPANY_MASTER.Comp_Code,TSPL_COMPANY_MASTER.Logo_Img,TSPL_COMPANY_MASTER.Logo_Img2,TSPL_COMPANY_MASTER.Comp_Name," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Code," + TSPL_VENDOR_INVOICE_HEAD + ".Invoice_Entry_Date," + TSPL_VENDOR_INVOICE_HEAD + ".RefDocNo ," + TSPL_VENDOR_INVOICE_HEAD + ".RefDocType, " + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Name ," + TSPL_VENDOR_INVOICE_HEAD + ".Created_By  ," + TSPL_VENDOR_INVOICE_HEAD + ".Created_Date," + TSPL_VENDOR_INVOICE_HEAD + ".Due_Date," + TSPL_VENDOR_INVOICE_HEAD + ".Description," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Invoice_No," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Invoice_Date,CreatedBy.User_Name as CreateBy,AuthorisedBy.User_Name as ApproveBy,case when ((" + TSPL_VENDOR_INVOICE_HEAD + ".Posting_Date IS null ) Or (" + TSPL_VENDOR_INVOICE_HEAD + ".Posting_Date='') ) then 'Pending' else 'Posted' end as InvStatus,case when " + TSPL_VENDOR_INVOICE_HEAD + ".Document_Type='I' then 'Bill Inward Voucher' else case when " + TSPL_VENDOR_INVOICE_HEAD + ".Document_Type='D' then 'Debit Note' else case when " + TSPL_VENDOR_INVOICE_HEAD + ".Document_Type='C' then 'Credit Note' else '' end end end as InvoiceType  , ( select case when LEN(ISNULL(RefDocType,''))>0 then case when RefDocType='S' then 'SRN' else case when RefDocType='AP' then 'AP Invoice' end end +' : '+RefDocNo  +' - ' +(case when RefDocType='S' then (Select top 1 convert(varchar(100),SRN_Date,110)   from TSPL_SRN_HEAD where SRN_No =RefDocNo) else (select top 1 convert(varchar(100),Invoice_Entry_Date,110)  from " + TSPL_VENDOR_INVOICE_HEAD + " where RefDocNo  = RefDocNo) end)  else '' end from " + TSPL_VENDOR_INVOICE_HEAD + " where Document_No=Final.Document_No) as RefDocDescription,final.Detail_Line_No," + TSPL_VENDOR_INVOICE_HEAD + ".TapalNo," + TSPL_VENDOR_INVOICE_HEAD + ".DateAndTime ," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Bank_Code," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Bank_Name," + TSPL_VENDOR_INVOICE_HEAD + ".Branch_IFSC_Code," + TSPL_VENDOR_INVOICE_HEAD + ".Branch_Name," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Bank_ACNo   from ( " & Environment.NewLine &
  " SELECT " + TSPL_JOURNAL_DETAILS + ".Detail_Line_No ,VI.RefDocNo,VI.RefDocType, " + TSPL_JOURNAL_MASTER + ".Source_doc_No as Document_No," + TSPL_JOURNAL_DETAILS + ".Account_code as ACCode, " + TSPL_JOURNAL_DETAILS + ".Account_Desc as ACName,case when " + TSPL_JOURNAL_DETAILS + ".Amount>=0 then  " + TSPL_JOURNAL_DETAILS + ".Amount else 0 end as DrAmt,case when " + TSPL_JOURNAL_DETAILS + ".Amount<0 then -1 * " + TSPL_JOURNAL_DETAILS + ".Amount else 0 end as CrAmt," + TSPL_JOURNAL_DETAILS + ".Hirerachy_Code as Hirerachy_Code," + TSPL_JOURNAL_DETAILS + ".Cost_Centre_Code as Cost_Centre_Code, VI.Hirerachy_Code3 ,VI.Hirerachy_Code4  " & Environment.NewLine &
  " FROM " + TSPL_JOURNAL_MASTER + " INNER JOIN " + TSPL_JOURNAL_DETAILS + " ON " + TSPL_JOURNAL_MASTER + ".Journal_No = " + TSPL_JOURNAL_DETAILS + ".Journal_No  AND  " + TSPL_JOURNAL_MASTER + ".Voucher_No = " + TSPL_JOURNAL_DETAILS + ".Voucher_No  left outer join  TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code  =" + TSPL_JOURNAL_MASTER + ".Comp_Code  "


            qry = qry + " LEFT OUTER JOIN (select Detail_Line_No," + TSPL_VENDOR_INVOICE_DETAIL + ".Document_No," + TSPL_VENDOR_INVOICE_DETAIL + ".GL_Account_Code," & Environment.NewLine &
            " " + TSPL_VENDOR_INVOICE_DETAIL + ".Cost_Centre_Code, " + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code," + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code3," + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code4, " + TSPL_VENDOR_INVOICE_HEAD + ".RefDocNo, " + TSPL_VENDOR_INVOICE_HEAD + ".RefDocType " & Environment.NewLine &
            " from " + TSPL_VENDOR_INVOICE_HEAD + " left outer join " + TSPL_VENDOR_INVOICE_DETAIL + " on " + TSPL_VENDOR_INVOICE_DETAIL + ".Document_No = " + TSPL_VENDOR_INVOICE_HEAD + ".Document_No  " & Environment.NewLine &
            " GROUP BY " + TSPL_VENDOR_INVOICE_DETAIL + ".Document_No," + TSPL_VENDOR_INVOICE_DETAIL + ".GL_Account_Code, " & Environment.NewLine &
            " " + TSPL_VENDOR_INVOICE_DETAIL + ".Cost_Centre_Code," + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code, " + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code3," + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code4," + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code3," + TSPL_VENDOR_INVOICE_DETAIL + ".Hirerachy_Code4," + TSPL_VENDOR_INVOICE_HEAD + ".RefDocNo," + TSPL_VENDOR_INVOICE_HEAD + ".RefDocType,Detail_Line_No ) VI" & Environment.NewLine &
            "  on VI.Document_No  =" + TSPL_JOURNAL_MASTER + ".Source_Doc_No and VI.GL_Account_Code =" + TSPL_JOURNAL_DETAILS + ".Account_code AND " + TSPL_JOURNAL_DETAILS + ".Detail_Line_No=VI.Detail_Line_No " & Environment.NewLine &
            " left join TSPL_COST_CENTRE_FINANCIAL on TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code =VI.Cost_Centre_Code  left join TSPL_HIRERACHY_LEVEL_MASTER on TSPL_HIRERACHY_LEVEL_MASTER.Hirerachy_Code =VI.Hirerachy_Code  where  " + TSPL_JOURNAL_MASTER + ".Source_Doc_No = '" & StrDocNo & "' " & Environment.NewLine &
            " group by " + TSPL_JOURNAL_DETAILS + ".Detail_Line_No ,VI.RefDocNo,VI.RefDocType, " + TSPL_JOURNAL_MASTER + ".Source_doc_No ," + TSPL_JOURNAL_DETAILS + ".Account_code , " + TSPL_JOURNAL_DETAILS + ".Account_Desc, " + TSPL_JOURNAL_DETAILS + ".Amount," + TSPL_JOURNAL_DETAILS + ".Hirerachy_Code " & Environment.NewLine &
            " ," + TSPL_JOURNAL_DETAILS + ".Cost_Centre_Code,VI.Hirerachy_Code3,VI.Hirerachy_Code4 " & Environment.NewLine &
             " )Final left outer join " + TSPL_VENDOR_INVOICE_HEAD + " on " + TSPL_VENDOR_INVOICE_HEAD + ".Document_No=final.Document_No left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Comp_Code left outer join TSPL_USER_MASTER as CreatedBy on CreatedBy.User_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Created_By left outer join TSPL_USER_MASTER as AuthorisedBy on AuthorisedBy .User_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Modify_By left join TSPL_COST_CENTRE_FINANCIAL on TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code =final.Cost_Centre_Code where 2=2  and " + TSPL_VENDOR_INVOICE_HEAD + ".Document_No in ('" & StrDocNo & "') " & Environment.NewLine &
            " )xxx     group by Document_No  ,ACCode ,Detail_Line_No " & Environment.NewLine &
            " )xxxx 
         left outer join (SELECT  TSPL_VENDOR_INVOICE_DETAIL2.GL_Account_Code , STUFF((SELECT distinct ',' + (" + TSPL_VENDOR_INVOICE_DETAIL + ".Remarks) as Remarks   from " + TSPL_VENDOR_INVOICE_DETAIL + "   where TSPL_VENDOR_INVOICE_DETAIL2.GL_Account_Code = " + TSPL_VENDOR_INVOICE_DETAIL + ".GL_Account_Code and len(isnull (Remarks,'')) > 0 and " + TSPL_VENDOR_INVOICE_DETAIL + ".Document_No  = '" + StrDocNo + "' FOR XML PATH(''), TYPE ).value('.', 'NVARCHAR(MAX)') ,1,1,'')  as Rmarks
         from " + TSPL_VENDOR_INVOICE_DETAIL + "  as TSPL_VENDOR_INVOICE_DETAIL2  where  TSPL_VENDOR_INVOICE_DETAIL2.Document_No =  '" + StrDocNo + "'  group by GL_Account_Code)  TBL_Remarks on TBL_Remarks.GL_Account_Code = xxxx.ACCode 

        left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER .Comp_Code =xxxx.Comp_Code  left join TSPL_LOCATION_MASTER on left(TSPL_LOCATION_MASTER.Location_Code,3)=xxxx.Loc_Code left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code = xxxx.Vendor_Code   left outer join TSPL_STATE_MASTER on TSPL_VENDOR_MASTER.State_Code = TSPL_STATE_MASTER.STATE_CODE left outer join tspl_state_master as tspl_state_master_for_location_state on  tspl_state_master_for_location_state.state_code=tspl_location_master.state   order by xxxx.Document_No,xxxx.DrAmt desc,xxxx.CrAmt desc"



            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            Dim qry1 As String = "select  TSPL_ITEM_MASTER.HSN_Code, TSPL_SRN_DETAIL.Item_Code ,TSPL_SRN_DETAIL.Item_Desc," + TSPL_VENDOR_INVOICE_HEAD + ".Description ," + TSPL_VENDOR_INVOICE_HEAD + ".RefDocNo ," + TSPL_VENDOR_INVOICE_HEAD + ".RefDocType   from TSPL_SRN_DETAIL left outer join " + TSPL_VENDOR_INVOICE_HEAD + " on TSPL_SRN_DETAIL .SRN_No =" + TSPL_VENDOR_INVOICE_HEAD + ".RefDocNo  left outer join TSPL_ITEM_MASTER on TSPL_ITEM_MASTER.Item_Code =TSPL_SRN_DETAIL.Item_Code where RefDocType ='S'and " + TSPL_VENDOR_INVOICE_HEAD + ".Document_No in('" & StrDocNo & "') and ISNULL(Against_POInvoice_No,'')= '' and ISNULL(Against_PurchaseReturn_No,'')= ''  "
            If dt.Rows.Count > 0 Then
                Dim frmCRV As New frmCrystalReportViewer()
                If SettingCostCenterlevel AndAlso clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "UDL") = CompairStringResult.Equal Then
                    frmCRV.funsubreport(Form_ID, CrystalReportFolder.Purchase, qry, qry1, "rptAPInvoice_Hierarchy", "AP Invoice", "AP_InvoiceDetails.rpt", clsCommon.myCDate(strDate))
                Else
                    frmCRV.funsubreport(Form_ID, CrystalReportFolder.Purchase, qry, qry1, "rptAPInvoice", "AP Invoice", "AP_InvoiceDetails.rpt", clsCommon.myCDate(strDate))
                End If
                frmCRV = Nothing
            Else
                'clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
                clsCommon.MyMessageBoxShow("No Data Found")
            End If
        Catch ex As Exception

        End Try
        Return True
    End Function

    Public Shared Function funVendorServicePrint(ByVal Form_ID As String, ByVal strCancelDelete As String, ByVal strDate As DateTime, ByVal StrDocNo As String) As Boolean
        Dim TSPL_VENDOR_INVOICE_HEAD As String = Nothing
        Dim TSPL_VENDOR_INVOICE_detail As String = Nothing
        If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal Then
            TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD_Cancel_data"
            TSPL_VENDOR_INVOICE_detail = "TSPL_VENDOR_INVOICE_detail_cancel_data"
        ElseIf clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
            TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD_Delete_data"
            TSPL_VENDOR_INVOICE_detail = "TSPL_VENDOR_INVOICE_detail_Delete_data"
        Else
            TSPL_VENDOR_INVOICE_HEAD = "TSPL_VENDOR_INVOICE_HEAD"
            TSPL_VENDOR_INVOICE_detail = "TSPL_VENDOR_INVOICE_detail"
        End If


        Dim qry As String = " select " + TSPL_VENDOR_INVOICE_HEAD + ".IsEInvoice,"
        If clsCommon.CompairString(strCancelDelete, "Cancel") = CompairStringResult.Equal OrElse clsCommon.CompairString(strCancelDelete, "Delete") = CompairStringResult.Equal Then
            qry += " 'Cancelled' As Report_Status, "
        Else
            qry += " '' As Report_Status, "
        End If

        qry += "" + TSPL_VENDOR_INVOICE_HEAD + ".Invoice_Entry_Date, TSPL_Additional_Charges.SAC_Code,(TSPL_VENDOR_MASTER.Add1+' '+TSPL_VENDOR_MASTER.ADD2+' '+TSPL_VENDOR_MASTER.Add3) AS ADD1,TSPL_VENDOR_MASTER.PAN ,TSPL_VENDOR_MASTER.Pin_Code,TSPL_VENDOR_MASTER.State_Code,TSPL_VENDOR_MASTER.GSTFinalNo,TSPL_VENDOR_MASTER.City_Code_Desc,right(" + TSPL_VENDOR_INVOICE_HEAD + ".document_no,4) as Gatepass ," + TSPL_VENDOR_INVOICE_HEAD + ".Loc_Code as from_location ,TSPL_VENDOR_MASTER.GSTFinalNo as Cust_GstInNo," + TSPL_VENDOR_INVOICE_HEAD + ".document_no"
        If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RCDFCF") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(strCancelDelete, "Cancel") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(strCancelDelete, "Delete") <> CompairStringResult.Equal Then
            qry += " ,TSPL_VENDOR_INVOICE_HEAD.Document_Type,TSPL_VENDOR_INVOICE_HEAD.RefDocType,TSPL_VENDOR_INVOICE_HEAD.RefDocNo,TSPL_VENDOR_INVOICE_HEAD.Remarks "
        End If
        qry += " , cast(
        " + TSPL_VENDOR_INVOICE_HEAD + ".BarCode_Img as image
      ) As BarCode_Img ," + TSPL_VENDOR_INVOICE_detail + ".Discount,	" + TSPL_VENDOR_INVOICE_detail + ".Amount_less_Discount	," + TSPL_VENDOR_INVOICE_detail + ".TAX1	," + TSPL_VENDOR_INVOICE_detail + ".TAX1_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX1_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".TAX2,	" + TSPL_VENDOR_INVOICE_detail + ".TAX2_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX2_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".TAX3	," + TSPL_VENDOR_INVOICE_detail + ".TAX3_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX3_Amt	," + TSPL_VENDOR_INVOICE_detail + ".TAX4,	" + TSPL_VENDOR_INVOICE_detail + ".TAX4_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX4_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".TAX5,	" + TSPL_VENDOR_INVOICE_detail + ".TAX5_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX5_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".TAX6,	" + TSPL_VENDOR_INVOICE_detail + ".TAX6_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX6_Amt	," + TSPL_VENDOR_INVOICE_detail + ".TAX7,	" + TSPL_VENDOR_INVOICE_detail + ".TAX7_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX7_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".TAX8,	" + TSPL_VENDOR_INVOICE_detail + ".TAX8_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX8_Amt	," + TSPL_VENDOR_INVOICE_detail + ".TAX9,	" + TSPL_VENDOR_INVOICE_detail + ".TAX9_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX9_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".TAX10,	" + TSPL_VENDOR_INVOICE_detail + ".TAX10_Rate,	" + TSPL_VENDOR_INVOICE_detail + ".TAX10_Amt,	" + TSPL_VENDOR_INVOICE_detail + ".Total_Tax,	" + TSPL_VENDOR_INVOICE_detail + ".Total_Amount	," + TSPL_VENDOR_INVOICE_detail + ".Remarks,	" + TSPL_VENDOR_INVOICE_detail + ".Comments	,	" + TSPL_VENDOR_INVOICE_detail + ".Invoice_Type,	" + TSPL_VENDOR_INVOICE_detail + ".Landed_Amount
	  ," + TSPL_VENDOR_INVOICE_HEAD + ".IRN_No," + TSPL_VENDOR_INVOICE_HEAD + ".Ack_No," + TSPL_VENDOR_INVOICE_HEAD + ".Ack_Date," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Invoice_No," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Invoice_Date," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Name," + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Code," + TSPL_VENDOR_INVOICE_detail + ".AddChargeDesc," + TSPL_VENDOR_INVOICE_detail + ".Amount
		,tspl_company_master.comp_Name, " + TSPL_VENDOR_INVOICE_detail + ".GL_Account_Code," + TSPL_VENDOR_INVOICE_detail + ".GL_Account_Desc,
      TSPL_COMPANY_MASTER.Add1 as Comp_Add1, 
      TSPL_COMPANY_MASTER.Add2 as Comp_Add2, 
      TSPL_COMPANY_MASTER.Add3 as Comp_Add3, 
      TSPL_COMPANY_MASTER.Email as Comp_Email, 
      TSPL_COMPANY_MASTER.Phone1 as Comp_Phone1, 
      TSPL_COMPANY_MASTER.Phone2 as Comp_Phone2, 
      TSPL_COMPANY_MASTER.Pan_No as Comp_Pan_No, 
      cast(
        TSPL_COMPANY_MASTER.Logo_Img as image
      ) as Logo_Img, 
      cast(
        TSPL_COMPANY_MASTER.Logo_Img2 as image
      ) as Logo_Img2, 
      TSPL_COMPANY_MASTER.GSTREg_No as Comp_GSTREg_No, 
      TSPL_COMPANY_MASTER.CINNO as Comp_CINNO, 
      TSPL_COMPANY_MASTER.Access_Officer as Comp_Access_Officer , TSPL_LOCATION_MASTER.accountholdername, 
      TSPL_LOCATION_MASTER.Bank, 
      TSPL_LOCATION_MASTER.Branch, 
      TSPL_LOCATION_MASTER.ACType, 
      TSPL_LOCATION_MASTER.bankaccno, 
      TSPL_LOCATION_MASTER.bankifsccode ,
	  TSPL_LOCATION_MASTER.Pin_Code as PinNo,
	  TSPL_LOCATION_MASTER.Phone1 as LPhone,
	  TSPL_LOCATION_MASTER.Registration_Number as Registration_No,TSPL_LOCATION_MASTER.GSTNO as From_Loc_GstinNo, TSPL_LOCATION_MASTER.HOAdd1 as frmHO1, 
      TSPL_LOCATION_MASTER.HOAdd2 as frmHO2, TSPL_LOCATION_MASTER.Location_Desc as [From Location Desc], 
      (
        TSPL_LOCATION_MASTER.Add1 + TSPL_LOCATION_MASTER.Add2 + TSPL_LOCATION_MASTER.Add3 + TSPL_LOCATION_MASTER.Add4
      ) as [From Address], 
     
      TSPL_LOCATION_MASTER.TIN_No, 
      TSPL_LOCATION_MASTER.CST_No, 
	  TSPL_VENDOR_MASTER.Pin_Code as [To Pin Code], 
	  TSPL_VENDOR_MASTER.PAN as Cust_Pan,
      TSPL_VENDOR_MASTER.TIN_No as [To TIN No], 
      TSPL_VENDOR_MASTER.CST as [To CST No], 
      TSPL_VENDOR_MASTER.Phone1 as [To phone], 
      TSPL_LOCATION_MASTER.State as From_State  from " + TSPL_VENDOR_INVOICE_HEAD + "
	  left outer join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Comp_Code
	  left  join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Loc_Segment_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Loc_Code and TSPL_LOCATION_MASTER.IsEinvoice=1
	  --left Outer Join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Loc_Segment_Code = " + TSPL_VENDOR_INVOICE_HEAD + ".Loc_code and TSPL_LOCATION_MASTER.isEinvoice=1
	  inner join " + TSPL_VENDOR_INVOICE_detail + " on " + TSPL_VENDOR_INVOICE_detail + ".document_no=" + TSPL_VENDOR_INVOICE_HEAD + ".document_no
	 -- inner join tspl_customer_master on tspl_customer_master.Cust_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Code
	  inner join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=" + TSPL_VENDOR_INVOICE_HEAD + ".Vendor_Code
	  left join TSPL_Additional_Charges on TSPL_Additional_Charges.Code=" + TSPL_VENDOR_INVOICE_detail + ".AddChargeCode  where " + TSPL_VENDOR_INVOICE_HEAD + ".Document_No ='" + StrDocNo + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim frmCRV As New frmCrystalReportViewer()
            If clsCommon.CompairString(objCommonVar.CurrComp_Code1, "RCDFCF") = CompairStringResult.Equal Then
                frmCRV.funreport(Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptVendorServiceInvoice_RCDFCF", "VendorService")
            Else
                frmCRV.funreport(Form_ID, CrystalReportFolder.KwalitySalesReport, dt, "rptVendorServiceInvoice_Union", "VendorService")

            End If
            frmCRV = Nothing
        Else
            clsCommon.MyMessageBoxShow("No Data Found")
        End If
        Return True
    End Function
    Public Function SaveData(ByVal obj As clsVedorInvoiceHead, ByVal isNewEntry As Boolean, Optional ByVal IsDirectEntry As Boolean = False) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = obj.SaveData(obj, isNewEntry, trans, "", False, IsDirectEntry)
            If isSaved AndAlso obj.is_For_Provision = 1 Then
                If obj.isFreightProvisionAccount Then
                    'clsProvisionEntryAdvanceKnockOff.SaveData(obj.Document_No,obj.Discount_Base, arrProvDocNo, trans)
                Else
                    isSaved = isSaved AndAlso clsProvisionEntry.SaveData(obj.Document_No, arrProvDocNo, trans)
                End If
            End If
            If (isSaved) Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch err As Exception
            trans.Rollback()
            Throw New Exception(err.Message)
        End Try
        Return isSaved
    End Function

    Public Function SaveData(ByVal obj As clsVedorInvoiceHead, ByVal isNewEntry As Boolean, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoRecreateOnly As String = Nothing, Optional ByVal strasset As Boolean = False, Optional ByVal IsDirectEntry As Boolean = False) As Boolean
        ''update TSPL_DOCPREFIX_MASTER set Doc_Trans_Type='Direct AP' where Doc_Type in ('AP Debit Note','AP Credit Note','AP Invoice') 

        If clsCommon.myLen(obj.loc_code) <= 0 Then
            Throw New Exception("Please first select Location")
        End If
        If clsCommon.CompairString(obj.Form_ID, clsUserMgtCode.frmProcurementDeduction) = CompairStringResult.Equal Then
            clsMCCPaymentCycleLockForScheduler.CheckForSchedulerLock(obj.MCC_Code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
        End If
        If clsCommon.myCstr(obj.Invoice_Type) = "VS" Then
            clsERPFuncationality.ValidateLocationCode(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePayable, clsUserMgtCode.FrmVendorService, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
        ElseIf obj.is_For_TDS = 1 Then
            clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleTDS, clsUserMgtCode.mbtnAPInvoiceEntryTDS, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
        Else
            clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePayable, clsUserMgtCode.mbtnAPInvoiceEntry, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
        End If
        If Not isNewEntry Then
            clsCommonFunctionality.SaveHistoryData(EnumSaveType.History, objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", "TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL", "AP_Invoice_No", "TSPL_AP_Invoice_Asset_EMI_Details", "AP_Invoice_No", "", "", "", "", "", "", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_REMITTANCE", "Document_No", "TSPL_AP_INVOICE_ADVANCE_INTEREST", "AP_Invoice_No", "TSPL_PROVISION_ENTRY_KNOCKOFF", "AP_Invoice_No", trans)
        End If

        Dim qry As String = "delete from TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL where AP_Invoice_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_VENDOR_INVOICE_DETAIL where Document_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_REMITTANCE where Document_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_AP_Invoice_Asset_EMI_Details where AP_Invoice_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "delete from TSPL_AP_INVOICE_ADVANCE_INTEREST where AP_Invoice_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "Delete from TSPL_PROVISION_ENTRY_KNOCKOFF where AP_Invoice_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        qry = "Delete from TSPL_VENDOR_INVOICE_SECURITY_REFUND where AP_Invoice_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)

        If obj.Arr.Count <= 0 Then
            Throw New Exception("Please fill at list one Account")
        End If
        Dim strLocation As String
        Dim isLocationCodeSegment As Boolean

        If objCommonVar.ApplyLocationWisePrefix Then
            strLocation = obj.Location_Code_Prefix
            isLocationCodeSegment = False
        Else
            isLocationCodeSegment = True
            strLocation = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))
            If clsCommon.myLen(strLocation) <= 0 Then
                Throw New Exception("Please enter account with location segment")
            End If
        End If



        Dim strDocNo As String = ""
        Dim strPrefixTransType As String = clsDocTransactionType.DirectAP
        If clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 Then
            strPrefixTransType = clsDocTransactionType.MccProc
        ElseIf clsCommon.myLen(obj.Against_BulkMillkPurchaseInvoice_No) > 0 Then
            strPrefixTransType = clsDocTransactionType.BulkProc
        ElseIf clsCommon.myLen(obj.Against_POInvoice_No) > 0 OrElse clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
            strPrefixTransType = clsDocTransactionType.GeneralPurchase
        End If


        If (isNewEntry) Then
            Dim CreateSeperateSeriesforRefDocAPinvforCreritdebit As Integer = 0
            CreateSeperateSeriesforRefDocAPinvforCreritdebit = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateSeperateSeriesforRefDocAPinvforCreditdebit, clsFixedParameterCode.CreateSeperateSeriesforRefDocAPinvforCreditdebit, trans))
            If clsCommon.myLen(obj.Document_No) > 0 Then
                ''No need to generate Code.
            ElseIf (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) Then
                obj.Document_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.APInvoice, strPrefixTransType, strLocation, isLocationCodeSegment)
            ElseIf (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                If clsCommon.myLen(RefDocNo) > 0 And IsDirectEntry = True And CreateSeperateSeriesforRefDocAPinvforCreritdebit = 1 Then
                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.DebitNote, clsDocTransactionType.DebitRefDoc, strLocation, isLocationCodeSegment)
                Else
                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.DebitNote, strPrefixTransType, strLocation, isLocationCodeSegment)
                End If
            ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                If clsCommon.myLen(RefDocNo) > 0 And IsDirectEntry = True And CreateSeperateSeriesforRefDocAPinvforCreritdebit = 1 Then
                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.CreditNote, clsDocTransactionType.CreditRefDoc, strLocation, isLocationCodeSegment)
                Else
                    obj.Document_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.CreditNote, strPrefixTransType, strLocation, isLocationCodeSegment)
                End If

            ElseIf (clsCommon.CompairString(obj.Document_Type, "P") = CompairStringResult.Equal) Then
                obj.Document_No = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaserOrder, strPrefixTransType, strLocation, isLocationCodeSegment)
            End If
        End If
        If (clsCommon.myLen(obj.Document_No) <= 0) Then
            Throw New Exception("Error in Document Code Generation")
        End If
        strLocation = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))

        ''richa agarwal 
        Dim strcurrentfisyearstartdate As Date? = Nothing
        Dim strcurrentfisyearenddate As Date? = Nothing
        Dim strcurrentfisyear As String = String.Empty
        Dim strmonth As Integer = Convert.ToDateTime(obj.Invoice_Entry_Date).Month()
        Dim stryear As Integer = Convert.ToDateTime(obj.Invoice_Entry_Date).Year()
        If strmonth <= 3 Then
            strcurrentfisyearstartdate = "01-Apr-" + clsCommon.myCstr(stryear - 1)
            strcurrentfisyearenddate = "31-Mar-" + clsCommon.myCstr(stryear)
            strcurrentfisyear = clsCommon.myCstr(stryear - 1) + "-" + clsCommon.myCstr(stryear)
        Else
            strcurrentfisyearstartdate = "01-Apr-" + clsCommon.myCstr(stryear)
            strcurrentfisyearenddate = "31-Mar-" + clsCommon.myCstr(stryear + 1)
            strcurrentfisyear = clsCommon.myCstr(stryear) + "-" + clsCommon.myCstr(stryear + 1)
        End If

        If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Vendor_Invoice_No) > 0 Then
            ''===============03/10/2016---------------------------
            Dim addwhrCls As String = ""
            If clsCommon.myLen(obj.Addition_Doc_Type) > 0 Then
                addwhrCls = " and Addition_Doc_Type='" + obj.Addition_Doc_Type + "' "
            End If
            '===========================================

            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*)  from TSPL_VENDOR_INVOICE_HEAD where convert(date,Invoice_Entry_Date,103)>= convert(date,'" & strcurrentfisyearstartdate & "',103)  and convert(datetime,Invoice_Entry_Date,103)<= convert(datetime ,'" & strcurrentfisyearenddate & "',103)  and isnull(Vendor_Invoice_No,'') ='" & obj.Vendor_Invoice_No & "' and Vendor_Code ='" & obj.Vendor_Code & "' and Document_Type ='I' and Document_No<>'" & obj.Document_No & "' " + addwhrCls + " ", trans)) >= 1 Then
                Throw New Exception("Vendor Invoice No. " & obj.Vendor_Invoice_No & " is already exist for Vendor " & obj.Vendor_Code & " in current financial year " & strcurrentfisyear & ". ")
            End If
        End If

        ''------------------



        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Invoice_Entry_Date", clsCommon.GetPrintDate(obj.Invoice_Entry_Date, "dd/MM/yyyy"))

        Dim ServerTime As DateTime = Nothing
        If isNewEntry Then
            ServerTime = clsCommon.GETSERVERDATE(trans)
            obj.Invoice_Entry_Date_New = New DateTime(clsCommon.myCDate(obj.Invoice_Entry_Date).Year, clsCommon.myCDate(obj.Invoice_Entry_Date).Month, clsCommon.myCDate(obj.Invoice_Entry_Date).Day, ServerTime.Hour, ServerTime.Minute, ServerTime.Second)
        End If

        clsCommon.AddColumnsForChange(coll, "Invoice_Entry_Date_New", clsCommon.GetPrintDate(obj.Invoice_Entry_Date_New, "dd/MMM/yyyy hh:mm tt"))
        clsCommon.AddColumnsForChange(coll, "Vendor_Code", obj.Vendor_Code)
        clsCommon.AddColumnsForChange(coll, "Vendor_Name", obj.Vendor_Name)
        clsCommon.AddColumnsForChange(coll, "Vendor_Invoice_No", obj.Vendor_Invoice_No)
        ''richa agarwal 14 Sep, 2016
        If clsCommon.myLen(obj.Vendor_Invoice_Date) <= 0 Then
            obj.Vendor_Invoice_Date = Invoice_Entry_Date
        End If
        ''----------

        ''added by Monika=================================
        clsCommon.AddColumnsForChange(coll, "Addition_Doc_Type", obj.Addition_Doc_Type, True)
        clsCommon.AddColumnsForChange(coll, "Employee_Type", obj.Employee_Type, True)
        ''===================================================

        clsCommon.AddColumnsForChange(coll, "Vendor_Invoice_Date", clsCommon.GetPrintDate(obj.Vendor_Invoice_Date, "dd/MM/yyyy"))
        clsCommon.AddColumnsForChange(coll, "Account_Set", obj.Account_Set)
        clsCommon.AddColumnsForChange(coll, "Document_Type", obj.Document_Type)
        clsCommon.AddColumnsForChange(coll, "PO_Number", obj.PO_Number)
        '------addesd by usha
        clsCommon.AddColumnsForChange(coll, "Loc_Code", obj.loc_code)
        clsCommon.AddColumnsForChange(coll, "IRregular_loc_code", obj.Irregular_loc_code)
        '------------
        clsCommon.AddColumnsForChange(coll, "Location_Code_Prefix", obj.Location_Code_Prefix, True)
        clsCommon.AddColumnsForChange(coll, "RCM", IIf(obj.RCM, 1, 0))
        clsCommon.AddColumnsForChange(coll, "TDS_Provision", IIf(obj.TDS_Provision, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Transfer_To_Saving", IIf(obj.Transfer_To_Saving, 1, 0))
        clsCommon.AddColumnsForChange(coll, "IsEInvoice", IIf(obj.IsEInvoice, 1, 0))

        'clsCommon.AddColumnsForChange(coll, "TDS_Provision", clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Is_Provisional from TSPL_VENDOR_MASTER where Vendor_Code='" + obj.Vendor_Code + "'", trans)))
        clsCommon.AddColumnsForChange(coll, "MCC_Code", obj.MCC_Code, True)
        clsCommon.AddColumnsForChange(coll, "MCC_Name", obj.MCC_Name, True)

        clsCommon.AddColumnsForChange(coll, "Vendor_Bank_Code", obj.Vendor_Bank_Code, True)
        clsCommon.AddColumnsForChange(coll, "Vendor_Bank_Name", obj.Vendor_Bank_Name, True)
        clsCommon.AddColumnsForChange(coll, "Branch_IFSC_Code", obj.Branch_IFSC_Code, True)
        clsCommon.AddColumnsForChange(coll, "Branch_Name", obj.Branch_Name, True)
        clsCommon.AddColumnsForChange(coll, "Vendor_Bank_ACNo", obj.Vendor_Bank_ACNo, True)
        clsCommon.AddColumnsForChange(coll, "Against_MCC_Material_Sale", obj.Against_MCC_Material_Sale, True)
        clsCommon.AddColumnsForChange(coll, "Against_MCC_Material_Sale_Return", obj.Against_MCC_Material_Sale_Return, True)

        Dim ApplyNoGSTCreditIndependentlyOnVendorServiceCharge As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, clsFixedParameterCode.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, trans)) = 1, True, False)
        If ApplyNoGSTCreditIndependentlyOnVendorServiceCharge = True Then
            clsCommon.AddColumnsForChange(coll, "No_GST_Credit", IIf(obj.No_GST_Credit, 1, 0))
        Else
            clsCommon.AddColumnsForChange(coll, "No_GST_Credit", IIf(obj.RCM AndAlso obj.No_GST_Credit, 1, 0))
        End If

        '''''added by priit 
        clsCommon.AddColumnsForChange(coll, "RefDocType", obj.RefDocType)
        clsCommon.AddColumnsForChange(coll, "RefDocNo", obj.RefDocNo)
        clsCommon.AddColumnsForChange(coll, "Against_MillkPurchaseInvoice_No", obj.Against_MillkPurchaseInvoice_No, True)
        clsCommon.AddColumnsForChange(coll, "Against_VSPItemIssue_No", obj.Against_VSPItemIssue_No, True)
        clsCommon.AddColumnsForChange(coll, "Against_VSP_Asset_Issue", obj.Against_VSP_Asset_Issue, True)
        clsCommon.AddColumnsForChange(coll, "Against_BulkMillkPurchaseInvoice_No", obj.Against_BulkMillkPurchaseInvoice_No, True)
        '''' ends here
        clsCommon.AddColumnsForChange(coll, "Order_No", obj.Order_No)
        clsCommon.AddColumnsForChange(coll, "Document_Total", obj.Document_Total)
        clsCommon.AddColumnsForChange(coll, "RoundOffAmount", clsCommon.myCdbl(obj.RoundOffAmount))
        clsCommon.AddColumnsForChange(coll, "GSTRegistered", obj.GSTRegistered)
        clsCommon.AddColumnsForChange(coll, "ISProcurementDeduction", obj.ISProcurementDeduction)
        Dim strOnHold As String = "N"
        If (obj.On_Hold) Then
            strOnHold = "Y"
        End If
        clsCommon.AddColumnsForChange(coll, "is_For_TDS", obj.is_For_TDS)
        clsCommon.AddColumnsForChange(coll, "On_Hold", strOnHold)
        clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
        clsCommon.AddColumnsForChange(coll, "Description", obj.Description)
        clsCommon.AddColumnsForChange(coll, "Tax_Group", obj.Tax_Group)
        clsCommon.AddColumnsForChange(coll, "is_For_Provision", obj.is_For_Provision)
        clsCommon.AddColumnsForChange(coll, "isDeduction", obj.isDeduction)
        clsCommon.AddColumnsForChange(coll, "is_Secondary_Transporter_Deduction", IIf(obj.is_Secondary_Transporter_Deduction, 1, 0))
        'If obj.isDeduction = 1 Then
        '    clsCommon.AddColumnsForChange(coll, "Deduction_Code", clsCommon.myCstr(obj.Deduction_Code))
        '    clsCommon.AddColumnsForChange(coll, "Deduction_Desc", clsCommon.myCstr(obj.Deduction_Desc))
        'End If
        If obj.is_For_Provision = 1 Then
            clsCommon.AddColumnsForChange(coll, "Prov_From_Date", clsCommon.GetPrintDate(obj.Prov_From_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Prov_To_Date", clsCommon.GetPrintDate(obj.Prov_To_Date, "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Prov_Amt", clsCommon.myCdbl(obj.Prov_Amt))
        End If

        clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
        clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax1_BAmount", obj.Tax1_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
        clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax2_BAmount", obj.Tax2_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
        clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax3_BAmount", obj.Tax3_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
        clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax4_BAmount", obj.Tax4_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
        clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax5_BAmount", obj.Tax5_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
        clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax6_BAmount", obj.Tax6_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
        clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax7_BAmount", obj.Tax7_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
        clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax8_BAmount", obj.Tax8_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
        clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax9_BAmount", obj.Tax9_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
        clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
        clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
        clsCommon.AddColumnsForChange(coll, "Tax10_BAmount", obj.Tax10_BAmount)
        clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
        clsCommon.AddColumnsForChange(coll, "Total_Tax", obj.Total_Tax)
        clsCommon.AddColumnsForChange(coll, "Ref_SNo", obj.Ref_SNo)
        If clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) <= 0 Then
            obj.Vendor_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Vendor_Control_AC, strLocation, True, trans)
        End If
        clsCommon.AddColumnsForChange(coll, "Vendor_Control_AC", obj.Vendor_Control_AC, True)
        obj.Discount_GL_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Discount_GL_AC, strLocation, True, trans)
        clsCommon.AddColumnsForChange(coll, "Discount_GL_AC", obj.Discount_GL_AC, True)
        clsCommon.AddColumnsForChange(coll, "Empty_Amount", obj.Empty_Amount)
        clsCommon.AddColumnsForChange(coll, "Empty_Account", obj.Empty_Account, True)
        clsCommon.AddColumnsForChange(coll, "PROJECT_ID", obj.PROJECT_ID, True)
        clsCommon.AddColumnsForChange(coll, "Invoice_Type", obj.Invoice_Type)
        clsCommon.AddColumnsForChange(coll, "Is_Security", obj.Security)
        clsCommon.AddColumnsForChange(coll, "Against_Salary_Generation_Code", obj.Against_Salary_Generation_Code, True)
        clsCommon.AddColumnsForChange(coll, "Against_Asset_Work", obj.Against_Asset_Work, True)
        clsCommon.AddColumnsForChange(coll, "Against_TransferToSavingPKID", obj.Against_TransferToSavingPKID, True)
        UpdateAllTax(obj.Document_No, trans)    '---This Function Updates All TaxGLAccount='' And TaxGLAmt=0

        Dim objTM As clsTaxMaster = clsTaxMaster.GetData(obj.TAX1, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC_Amt", Math.Round(obj.TAX1_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC2_Amt", Math.Round(obj.TAX1_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC3_Amt", Math.Round(obj.TAX1_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC4_Amt", Math.Round(obj.TAX1_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX1_GLAC5_Amt", Math.Round(obj.TAX1_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If

        End If

        objTM = clsTaxMaster.GetData(obj.TAX2, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC_Amt", Math.Round(obj.TAX2_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC2_Amt", Math.Round(obj.TAX2_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC3_Amt", Math.Round(obj.TAX2_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC4_Amt", Math.Round(obj.TAX2_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX2_GLAC5_Amt", Math.Round(obj.TAX2_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX3, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC_Amt", Math.Round(obj.TAX3_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC2_Amt", Math.Round(obj.TAX3_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC3_Amt", Math.Round(obj.TAX3_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC4_Amt", Math.Round(obj.TAX3_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX3_GLAC5_Amt", Math.Round(obj.TAX3_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX4, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC_Amt", Math.Round(obj.TAX4_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC2_Amt", Math.Round(obj.TAX4_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC3_Amt", Math.Round(obj.TAX4_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC4_Amt", Math.Round(obj.TAX4_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX4_GLAC5_Amt", Math.Round(obj.TAX4_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX5, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC_Amt", Math.Round(obj.TAX5_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC2_Amt", Math.Round(obj.TAX5_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC3_Amt", Math.Round(obj.TAX5_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC4_Amt", Math.Round(obj.TAX5_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX5_GLAC5_Amt", Math.Round(obj.TAX5_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX6, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC_Amt", Math.Round(obj.TAX6_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC2_Amt", Math.Round(obj.TAX6_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC3_Amt", Math.Round(obj.TAX6_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC4_Amt", Math.Round(obj.TAX6_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX6_GLAC5_Amt", Math.Round(obj.TAX6_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX7, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC_Amt", Math.Round(obj.TAX7_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC2_Amt", Math.Round(obj.TAX7_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC3_Amt", Math.Round(obj.TAX7_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC4_Amt", Math.Round(obj.TAX7_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX7_GLAC5_Amt", Math.Round(obj.TAX7_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX8, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC_Amt", Math.Round(obj.TAX8_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC2_Amt", Math.Round(obj.TAX8_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC3_Amt", Math.Round(obj.TAX8_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC4_Amt", Math.Round(obj.TAX8_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX8_GLAC5_Amt", Math.Round(obj.TAX8_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX9, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC_Amt", Math.Round(obj.TAX9_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC2_Amt", Math.Round(obj.TAX9_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)

                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC3_Amt", Math.Round(obj.TAX9_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC4_Amt", Math.Round(obj.TAX9_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX9_GLAC5_Amt", Math.Round(obj.TAX9_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If
        objTM = clsTaxMaster.GetData(obj.TAX10, trans)
        If objTM IsNot Nothing AndAlso objTM.Tax_Recoverable Then
            If clsCommon.myLen(objTM.Tax_Recoverable_Account) > 0 Then
                objTM.Tax_Recoverable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC", objTM.Tax_Recoverable_Account, True)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC_Amt", Math.Round(obj.TAX10_Amt * objTM.Tax_Recover_Rate / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account2) > 0 Then
                objTM.Tax_Recoverable_Account2 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account2, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC2", objTM.Tax_Recoverable_Account2, True)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC2_Amt", Math.Round(obj.TAX10_Amt * objTM.Tax_Recover_Rate2 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account3) > 0 Then
                objTM.Tax_Recoverable_Account3 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account3, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC3", objTM.Tax_Recoverable_Account3, True)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC3_Amt", Math.Round(obj.TAX10_Amt * objTM.Tax_Recover_Rate3 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account4) > 0 Then
                objTM.Tax_Recoverable_Account4 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account4, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC4", objTM.Tax_Recoverable_Account4, True)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC4_Amt", Math.Round(obj.TAX10_Amt * objTM.Tax_Recover_Rate4 / 100, 2, MidpointRounding.ToEven))
            End If
            If clsCommon.myLen(objTM.Tax_Recoverable_Account5) > 0 Then
                objTM.Tax_Recoverable_Account5 = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.Tax_Recoverable_Account5, strLocation, True, trans)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC5", objTM.Tax_Recoverable_Account5, True)
                clsCommon.AddColumnsForChange(coll, "TAX10_GLAC5_Amt", Math.Round(obj.TAX10_Amt * objTM.Tax_Recover_Rate5 / 100, 2, MidpointRounding.ToEven))
            End If
        End If

        ''richa agarwal changes on 12 Aug,2016 against AP Aging report for invoice type
        If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) AndAlso clsCommon.myLen(obj.Terms_Code) <= 0 Then
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("Select TSPL_VENDOR_MASTER.Terms_Code ,TSPL_TERMS_MASTER.Terms_Desc,TSPL_TERMS_MASTER.No_Days   from TSPL_VENDOR_MASTER left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code =TSPL_VENDOR_MASTER.Terms_Code where TSPL_VENDOR_MASTER.Vendor_Code ='" & clsCommon.myCstr(obj.Vendor_Code) & "'", trans)
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                obj.Terms_Code = clsCommon.myCstr(dt1.Rows(0)("Terms_Code"))
                obj.Terms_Description = clsCommon.myCstr(dt1.Rows(0)("Terms_Desc"))
                obj.Due_Date = clsCommon.myCDate(obj.Invoice_Entry_Date).AddDays(clsCommon.myCdbl(dt1.Rows(0)("No_Days")))
            Else
                Throw New Exception("Please enter Terms Code for Vendor " & obj.Vendor_Code & " in Vendor Master")
            End If
        End If
        ''----------------------

        clsCommon.AddColumnsForChange(coll, "Terms_Code", obj.Terms_Code)
        clsCommon.AddColumnsForChange(coll, "Terms_Description", obj.Terms_Description)

        If obj.Due_Date IsNot Nothing Then
            clsCommon.AddColumnsForChange(coll, "Due_Date", clsCommon.GetPrintDate(obj.Due_Date, "dd/MMM/yyyy"))
        End If

        ''richa agarwal changes on 14 Aug,2016 against AP Aging report for invoice type
        If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) Then
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("Select TSPL_VENDOR_MASTER.Terms_Code ,TSPL_TERMS_MASTER.Terms_Desc,TSPL_TERMS_MASTER.No_Days   from TSPL_VENDOR_MASTER left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code =TSPL_VENDOR_MASTER.Terms_Code where TSPL_VENDOR_MASTER.Vendor_Code ='" & clsCommon.myCstr(obj.Vendor_Code) & "'", trans)
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                obj.VI_Due_Date = clsCommon.myCDate(obj.Vendor_Invoice_Date).AddDays(clsCommon.myCdbl(dt1.Rows(0)("No_Days")))
            Else
                Throw New Exception("Please enter Terms Code for Vendor " & obj.Vendor_Code & " in Vendor Master")
            End If
        Else
            obj.VI_Due_Date = obj.Invoice_Entry_Date
        End If
        If (clsCommon.CompairString(obj.Document_Type, "P") = CompairStringResult.Equal) Then
            Dim dt1 As DataTable = clsDBFuncationality.GetDataTable("Select TSPL_VENDOR_MASTER.Terms_Code ,TSPL_TERMS_MASTER.Terms_Desc,TSPL_TERMS_MASTER.No_Days   from TSPL_VENDOR_MASTER left outer join TSPL_TERMS_MASTER on TSPL_TERMS_MASTER.Terms_Code =TSPL_VENDOR_MASTER.Terms_Code where TSPL_VENDOR_MASTER.Vendor_Code ='" & clsCommon.myCstr(obj.Vendor_Code) & "'", trans)
            If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 Then
                obj.VI_Due_Date = clsCommon.myCDate(obj.Vendor_Invoice_Date).AddDays(clsCommon.myCdbl(dt1.Rows(0)("No_Days")))
            Else
                Throw New Exception("Please enter Terms Code for Vendor " & obj.Vendor_Code & " in Vendor Master")
            End If

        End If
        If obj.VI_Due_Date IsNot Nothing Then
            clsCommon.AddColumnsForChange(coll, "VI_Due_Date", clsCommon.GetPrintDate(obj.VI_Due_Date, "dd/MMM/yyyy"))
        End If
        ''----------------------
        clsCommon.AddColumnsForChange(coll, "TapalNo", obj.TapalNo, True)
        If clsCommon.myLen(obj.DateAndTime) > 0 Then
            clsCommon.AddColumnsForChange(coll, "DateAndTime", clsCommon.GetPrintDate(obj.DateAndTime, "dd/MMM/yyyy hh:mm tt"))
        Else
            clsCommon.AddColumnsForChange(coll, "DateAndTime", Nothing, True)
        End If

        clsCommon.AddColumnsForChange(coll, "Discount_Percentage", obj.Discount_Percentage)
        clsCommon.AddColumnsForChange(coll, "Discount_Base", obj.Discount_Base)
        clsCommon.AddColumnsForChange(coll, "Discount_Amount", obj.Discount_Amount)
        clsCommon.AddColumnsForChange(coll, "Amount_Less_Discount", obj.Amount_Less_Discount)
        clsCommon.AddColumnsForChange(coll, "Total_Taxable_Amount", obj.Total_Taxable_Amount)
        clsCommon.AddColumnsForChange(coll, "TDS_Base_Actual_Amount", obj.TDS_Base_Actual_Amount)
        clsCommon.AddColumnsForChange(coll, "TDS_Base_Calculated_Amount", obj.TDS_Base_Calculated_Amount)
        clsCommon.AddColumnsForChange(coll, "TDS_Percentage", obj.TDS_Percentage)
        'clsCommon.AddColumnsForChange(coll, "TDS_Actual_Amount", Math.Round(obj.TDS_Actual_Amount, 0, MidpointRounding.AwayFromZero))
        clsCommon.AddColumnsForChange(coll, "TDS_Calculated_Amount", obj.TDS_Calculated_Amount)
        clsCommon.AddColumnsForChange(coll, "Nature_of_deduction", obj.Nature_of_deduction)
        clsCommon.AddColumnsForChange(coll, "Branch_Code", obj.Branch_Code)
        clsCommon.AddColumnsForChange(coll, "Section_Code", obj.Section_Code)
        'clsCommon.AddColumnsForChange(coll, "Balance_Amt", IIf(obj.is_For_TDS = 0, (obj.Document_Total - Math.Round(obj.TDS_Actual_Amount, 0, MidpointRounding.AwayFromZero)), obj.Document_Total))
        clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
        clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
        clsCommon.AddColumnsForChange(coll, "Against_POInvoice_No", obj.Against_POInvoice_No, True)
        clsCommon.AddColumnsForChange(coll, "Against_PurchaseReturn_No", obj.Against_PurchaseReturn_No, True)
        '' Anubhooti 19-Mar-2015 (Against_VCGL)
        clsCommon.AddColumnsForChange(coll, "Against_VCGL", obj.Against_VCGL, True)
        clsCommon.AddColumnsForChange(coll, "Hirerachy_Level_Code", obj.Hirerachy_Level_Code, True)
        clsCommon.AddColumnsForChange(coll, "Cost_Centre_Fin_Level_Code", obj.Cost_Centre_Fin_Level_Code, True)
        ''
        clsCommon.AddColumnsForChange(coll, "Against_Acquisition", obj.Against_Acquisition)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code1", obj.Add_Charge_Code1)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name1", obj.Add_Charge_Name1)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt1", obj.Add_Charge_Amt1)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code2", obj.Add_Charge_Code2)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name2", obj.Add_Charge_Name2)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt2", obj.Add_Charge_Amt2)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code3", obj.Add_Charge_Code3)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name3", obj.Add_Charge_Name3)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt3", obj.Add_Charge_Amt3)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code4", obj.Add_Charge_Code4)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name4", obj.Add_Charge_Name4)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt4", obj.Add_Charge_Amt4)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code5", obj.Add_Charge_Code5)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name5", obj.Add_Charge_Name5)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt5", obj.Add_Charge_Amt5)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code6", obj.Add_Charge_Code6)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name6", obj.Add_Charge_Name6)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt6", obj.Add_Charge_Amt6)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code7", obj.Add_Charge_Code7)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name7", obj.Add_Charge_Name7)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt7", obj.Add_Charge_Amt7)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code8", obj.Add_Charge_Code8)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name8", obj.Add_Charge_Name8)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt8", obj.Add_Charge_Amt8)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code9", obj.Add_Charge_Code9)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name9", obj.Add_Charge_Name9)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt9", obj.Add_Charge_Amt9)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Code10", obj.Add_Charge_Code10)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Name10", obj.Add_Charge_Name10)
        clsCommon.AddColumnsForChange(coll, "Add_Charge_Amt10", obj.Add_Charge_Amt10)
        clsCommon.AddColumnsForChange(coll, "Total_Add_Charge", obj.Total_Add_Charge)
        clsCommon.AddColumnsForChange(coll, "Total_Landed_Amt", obj.Total_Landed_Amt)
        clsCommon.AddColumnsForChange(coll, "Tax_Calculation_Type", IIf(obj.Tax_Calculation_Type = EnumTaxCalucationType.Automatic, 0, 1))
        clsCommon.AddColumnsForChange(coll, "Is_ProRated", obj.Is_ProRated)
        '======================
        clsCommon.AddColumnsForChange(coll, "Cash_Discount_Amt", obj.Cash_Discount_Amt)
        clsCommon.AddColumnsForChange(coll, "Amt_After_Cash_Discount", obj.Amt_After_Cash_Discount)
        '============
        '' currencyconversion
        If clsCommon.myLen(obj.CURRENCY_CODE) = 0 Then
            'obj.CURRENCY_CODE = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select BaseCurrencyCode from TSPL_COMPANY_MASTER where Comp_Code='" & objCommonVar.CurrentCompanyCode & "'", trans))
            'clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            'clsCommon.AddColumnsForChange(coll, "ConvRate", 1)

            obj.CURRENCY_CODE = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(CURRENCY_CODE,'') as CURRENCY_CODE from TSPL_VENDOR_MASTER where Vendor_Code ='" & obj.Vendor_Code & "'", trans))
            If clsCommon.myLen(obj.CURRENCY_CODE) > 0 Then
                Dim dt = clsModuleCurrencyMapping.GetLatestCurConvRateDT(obj.Invoice_Entry_Date, obj.CURRENCY_CODE, trans)
                If dt.Rows.Count = 0 Then
                    If obj.CURRENCY_CODE = objCommonVar.BaseCurrencyCode Then
                        obj.ConvRate = 1
                    Else
                        Throw New Exception("Conversion rate not entered for currency '" & obj.CURRENCY_CODE & "'")
                    End If
                Else
                    obj.ConvRate = dt.Rows(0).Item("Rate")
                    obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0).Item("FROM_DATE"), "dd-MMM-yyyy")
                End If
            Else
                Throw New Exception("Please enter currency for Vendor'" & obj.Vendor_Code & "' in vendor master.")
            End If
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
        Else
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE, True)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
        End If
        ' clsCommon.AddColumnsForChange(coll, "ApplicableFrom", obj.ApplicableFrom, True)

        If clsCommon.CompairString(obj.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) = CompairStringResult.Equal Then
            clsCommon.AddColumnsForChange(coll, "TDS_Actual_Amount", Math.Round(obj.TDS_Actual_Amount, 0, MidpointRounding.AwayFromZero))
            clsCommon.AddColumnsForChange(coll, "Balance_Amt", IIf(obj.is_For_TDS = 0, (obj.Document_Total - Math.Round(obj.TDS_Actual_Amount, 0, MidpointRounding.AwayFromZero)), obj.Document_Total))
        Else
            clsCommon.AddColumnsForChange(coll, "TDS_Actual_Amount", Math.Round(obj.TDS_Actual_Amount, 2, MidpointRounding.AwayFromZero))
            clsCommon.AddColumnsForChange(coll, "Balance_Amt", IIf(obj.is_For_TDS = 0, (obj.Document_Total - Math.Round(obj.TDS_Actual_Amount, 2, MidpointRounding.AwayFromZero)), obj.Document_Total))
        End If
        If Not obj.ApplicableFrom Is Nothing Then
            clsCommon.AddColumnsForChange(coll, "ApplicableFrom", clsCommon.GetPrintDate(obj.ApplicableFrom, "dd-MMM-yyyy"), True)
        End If

        '' End currencyconversion

        clsCommon.AddColumnsForChange(coll, "ITC_Elibible", IIf(obj.ITC_Elibible, 1, 0))
        clsCommon.AddColumnsForChange(coll, "ITC_Type", obj.ITC_Type)
        clsCommon.AddColumnsForChange(coll, "ITC_Type_Category", obj.ITC_Type_Category)
        clsCommon.AddColumnsForChange(coll, "Main_VSP_Milk_AP_Invoice_No", obj.Main_VSP_Milk_AP_Invoice_No, True)
        clsCommon.AddColumnsForChange(coll, "Saving", obj.Saving)
        If isNewEntry Then
            clsCommon.AddColumnsForChange(coll, "Document_No", obj.Document_No)
            clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MM/yyyy"))
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_INVOICE_HEAD", OMInsertOrUpdate.Insert, "", trans)
        Else
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_INVOICE_HEAD", OMInsertOrUpdate.Update, "Document_No='" + obj.Document_No + "'", trans)
        End If



        clsVedorInvoiceDetail.SaveData(obj.Document_No, Arr, trans)

        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", trans)

        clsRemittance.SaveData(obj.RemittanceObject, obj.Document_No, strLocation, trans)
        clsCustomFieldValues.SaveData(obj.Form_ID, obj.Document_No, obj.arrCustomFields, trans)
        clsAPInvoiceAssetEMIDetails.SaveData(obj.Document_No, obj.ArrAssetEMI, trans)
        clsAPInvoiceAdvanceInterest.SaveData(obj.Document_No, obj.ArrAdvanceInterest, trans)
        clsApprovalScreen.SaveApprovalAtTransLevel(obj.Form_ID, "Document_No", obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", trans)
        If obj.is_For_Provision AndAlso obj.isFreightProvisionAccount Then
            clsProvisionEntryAdvanceKnockOff.SaveData(obj, trans)
        End If
        If obj.arrSecurityAdjustment IsNot Nothing AndAlso obj.arrSecurityAdjustment.Count > 0 Then
            qry = clsVedorInvoiceHead.GetRefundBaseQuery(obj.Vendor_Code, obj.Document_No, " TSPL_VENDOR_INVOICE_HEAD.Document_No in (" + clsCommon.GetMulcallString(obj.arrSecurityAdjustment) + ")")
            Dim dtSec As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            Dim balAmt As Decimal = obj.Discount_Base
            For Each drSec As DataRow In dtSec.Rows
                balAmt -= clsCommon.myCDecimal(drSec("BalanceAmount"))
                coll = New Hashtable()
                clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", obj.Document_No)
                clsCommon.AddColumnsForChange(coll, "Against_Security", clsCommon.myCstr(drSec("DocumentCode")))
                If balAmt <= 0 Then
                    clsCommon.AddColumnsForChange(coll, "Amount", balAmt + clsCommon.myCDecimal(drSec("BalanceAmount")))
                Else
                    clsCommon.AddColumnsForChange(coll, "Amount", clsCommon.myCDecimal(drSec("BalanceAmount")))
                End If
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_INVOICE_SECURITY_REFUND", OMInsertOrUpdate.Insert, "", trans)
            Next
            If balAmt > 0 Then
                Throw New Exception("Document Amount is Greater than Balance of Security documents")
            End If
        End If

        obj = clsVedorInvoiceHead.GetData(obj.Document_No, "", trans)
        ''richa 24 Dec,2018  TEC/02/11/18-000361 create journal entry for opening in case of Credit or debit note as Journal Master table instead of journal master op table
        'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, trans)
        Dim JEWithOPening As Boolean = False
        If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
            Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
            If clsCommon.myCDate(clsCommon.GetPrintDate(obj.Invoice_Entry_Date, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                JEWithOPening = True
            End If
        End If

        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, trans)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "C") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "D") = CompairStringResult.Equal) And JEWithOPening = True Then
            clsVedorInvoiceHead.CreateJournalEntryForOpening(obj, Form_ID, obj.Invoice_Entry_Date, trans, strVoucherNoRecreateOnly, True, strasset)
        Else
            'update by preeti gupta Against ticket No [BM00000008498]
            clsVedorInvoiceHead.CreateJournalEntry(obj, Form_ID, obj.Invoice_Entry_Date, trans, strVoucherNoRecreateOnly, True, strasset)
        End If
        ''-----------------

        Return True
    End Function
    Public Shared Function UpdateEInvoiceAfterPosting(ByVal obj As clsVedorInvoiceHead, ByVal strDocumentNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            If obj IsNot Nothing And clsCommon.myLen(strDocumentNo) > 0 Then
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "IRN_No", obj.irn_no)
                clsCommon.AddColumnsForChange(coll, "Ack_No", obj.Ack_No)
                clsCommon.AddColumnsForChange(coll, "Ack_Date", clsCommon.GetPrintDate(obj.Ack_Date, "dd/MMM/yyyy hh:mm tt"))
                clsCommon.AddColumnsForChange(coll, "QR_Code", obj.QR_Code)


                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_INVOICE_HEAD", OMInsertOrUpdate.Update, "TSPL_VENDOR_INVOICE_HEAD.Document_No='" + strDocumentNo + "'", trans)
                '    clsCommon.AddColumnsForChange(coll, "Electronic_Ref_No", obj.Electronic_Ref_No)
                '    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_SCRAPSALE_HEAD", OMInsertOrUpdate.Update, "TSPL_SCRAPSALE_HEAD.shipment_No='" + ShipmentNo + "'", trans)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    '-29/06/2013-------------Created by--Pankaj kumar
    Public Shared Function UpdateAllTax(ByVal DocNo As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim Qry As String = "Update TSPL_VENDOR_INVOICE_HEAD Set "

            Qry += " TAX1_GLAC= NULL, TAX1_GLAC_Amt=0, TAX1_GLAC2= NULL, TAX1_GLAC2_Amt=0, TAX1_GLAC3= NULL, TAX1_GLAC3_Amt=0, TAX1_GLAC4= NULL, TAX1_GLAC4_Amt=0, TAX1_GLAC5= NULL, TAX1_GLAC5_Amt=0,"
            Qry += " TAX2_GLAC= NULL, TAX2_GLAC_Amt=0, TAX2_GLAC2= NULL, TAX2_GLAC2_Amt=0, TAX2_GLAC3= NULL, TAX2_GLAC3_Amt=0, TAX2_GLAC4= NULL, TAX2_GLAC4_Amt=0, TAX2_GLAC5= NULL, TAX2_GLAC5_Amt=0,"
            Qry += " TAX3_GLAC= NULL, TAX3_GLAC_Amt=0, TAX3_GLAC2= NULL, TAX3_GLAC2_Amt=0, TAX3_GLAC3= NULL, TAX3_GLAC3_Amt=0, TAX3_GLAC4= NULL, TAX3_GLAC4_Amt=0, TAX3_GLAC5= NULL, TAX3_GLAC5_Amt=0,"
            Qry += " TAX4_GLAC= NULL, TAX4_GLAC_Amt=0, TAX4_GLAC2= NULL, TAX4_GLAC2_Amt=0, TAX4_GLAC3= NULL, TAX4_GLAC3_Amt=0, TAX4_GLAC4= NULL, TAX4_GLAC4_Amt=0, TAX4_GLAC5= NULL, TAX4_GLAC5_Amt=0,"
            Qry += " TAX5_GLAC= NULL, TAX5_GLAC_Amt=0, TAX5_GLAC2= NULL, TAX5_GLAC2_Amt=0, TAX5_GLAC3= NULL, TAX5_GLAC3_Amt=0, TAX5_GLAC4= NULL, TAX5_GLAC4_Amt=0, TAX5_GLAC5= NULL, TAX5_GLAC5_Amt=0,"
            Qry += " TAX6_GLAC= NULL, TAX6_GLAC_Amt=0, TAX6_GLAC2= NULL, TAX6_GLAC2_Amt=0, TAX6_GLAC3= NULL, TAX6_GLAC3_Amt=0, TAX6_GLAC4= NULL, TAX6_GLAC4_Amt=0, TAX6_GLAC5= NULL, TAX6_GLAC5_Amt=0,"
            Qry += " TAX7_GLAC= NULL, TAX7_GLAC_Amt=0, TAX7_GLAC2= NULL, TAX7_GLAC2_Amt=0, TAX7_GLAC3= NULL, TAX7_GLAC3_Amt=0, TAX7_GLAC4= NULL, TAX7_GLAC4_Amt=0, TAX7_GLAC5= NULL, TAX7_GLAC5_Amt=0,"
            Qry += " TAX8_GLAC= NULL, TAX8_GLAC_Amt=0, TAX8_GLAC2= NULL, TAX8_GLAC2_Amt=0, TAX8_GLAC3= NULL, TAX8_GLAC3_Amt=0, TAX8_GLAC4= NULL, TAX8_GLAC4_Amt=0, TAX8_GLAC5= NULL, TAX8_GLAC5_Amt=0,"
            Qry += " TAX9_GLAC= NULL, TAX9_GLAC_Amt=0, TAX9_GLAC2= NULL, TAX9_GLAC2_Amt=0, TAX9_GLAC3= NULL, TAX9_GLAC3_Amt=0, TAX9_GLAC4= NULL, TAX9_GLAC4_Amt=0, TAX9_GLAC5= NULL, TAX9_GLAC5_Amt=0,"
            Qry += " TAX10_GLAC= NULL, TAX10_GLAC_Amt=0, TAX10_GLAC2= NULL, TAX10_GLAC2_Amt=0, TAX10_GLAC3= NULL, TAX10_GLAC3_Amt=0, TAX10_GLAC4= NULL, TAX10_GLAC4_Amt=0, TAX10_GLAC5= NULL, TAX10_GLAC5_Amt=0 "
            Qry += " WHERE TSPL_VENDOR_INVOICE_HEAD.Document_No='" + DocNo + "'"

            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function CancelData(ByVal Form_Id As String, ByVal Doc_No As String, ByVal InvoiceNo As String, ByVal NavType As NavigatorType) As Boolean
        '' created by Richa Agarwal against ticket No-ERO/09/09/19-001022  on date 09-09-2019
        Dim qry As String = ""
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            '' table list 
            ''1. TSPL_SD_SALE_INVOICE_HEAD
            ''2. TSPL_SD_SALE_INVOICE_DETAIL
            ''3. TSPL_EX_SALE_INVOICE_NOTIFY_PARTY_DETAIL
            ''4.TSPL_CUSTOM_FIELD_VALUES
            ''4. TSPL_INVENTORY_MOVEMENT     ( no need of history)
            ''5. TSPL_Customer_Invoice_Head
            ''6. TSPL_Customer_Invoice_Detail
            ''7. TSPL_JOURNAL_DETAILS
            ''8. TSPL_JOURNAL_MASTER
            ''9. TSPL_BATCH_ITEM ( no need of history)
            '' steps for checking the items stock and batch wise stock

            Dim obj As clsVedorInvoiceHead = clsVedorInvoiceHead.GetData(Doc_No, "VS", trans)

            If obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0 Then
                Throw New Exception("Document- " & Doc_No & " not found")
            End If
            Dim strInvTransType As String = String.Empty
            'If clsCommon.CompairString(obj.Document_Type, "MT") = CompairStringResult.Equal Then
            '    strInvTransType = "MT_SALE_IN"
            'Else
            '    strInvTransType = "EX_SALE_IN"
            'End If

            ''richa agarwal 24 Dec,2020
            Dim dtirn As DataTable = clsDBFuncationality.GetDataTable("select Einvoice_type,IRN_No,IsEInvoice,Loc_Code from TSPL_VENDOR_INVOICE_HEAD  where Document_No='" & InvoiceNo & "'", trans)
            If dtirn IsNot Nothing AndAlso dtirn.Rows.Count > 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(dtirn.Rows(0)("Einvoice_type")), "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(dtirn.Rows(0)("IsEInvoice")), "1") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Invoice_Entry_Date, trans) = True Then
                    If ClsEInvoiceOFAPIs.EInvoice_Cancellation(InvoiceNo, clsCommon.myCstr(dtirn.Rows(0)("IRN_No")), clsCommon.myCstr(dtirn.Rows(0)("Loc_Code")), trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If
            ''----------

            'clsItemLocationDetails.CheckCancelInventoryBalance(Form_Id, Doc_No, trans)
            '' transfer data into cancel table
            ' clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Doc_No, "tspl_sd_shipment_head", "Document_Code", "tspl_sd_shipment_DETAIL", "Document_Code", trans)
            clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, InvoiceNo, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", trans)


            '' cancel customer invoice data
            'qry = "select Document_No from TSPL_Customer_Invoice_Head  where against_Sale_no='" & InvoiceNo & "'"
            'Dim Document_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            'If clsCommon.myLen(Document_No) > 0 Then
            '    clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Document_No, "TSPL_Customer_Invoice_Head", "Document_No", "TSPL_Customer_Invoice_Detail", "Document_No", trans)
            'End If

            '' cancel journal master data invoice
            qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No ='" & InvoiceNo & "'"
            Dim Voucher_No As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            If clsCommon.myLen(Voucher_No) > 0 Then
                clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            End If

            '' cancel journal master data shipment
            'qry = "select Voucher_No from TSPL_JOURNAL_MASTER  where Source_Doc_No='" & Doc_No & "'"
            'Voucher_No = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            'If clsCommon.myLen(Voucher_No) > 0 Then
            '    clsCommonFunctionality.SaveCancelData(objCommonVar.CurrentUserCode, Voucher_No, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
            'End If

            '' cancel custom field data
            clsCommonFunctionality.SaveCancelDataMultKey(objCommonVar.CurrentUserCode, InvoiceNo, "TSPL_CUSTOM_FIELD_VALUES", "Transaction_Code", "Program_Code", Form_Id, trans)


            ''delete data from multiple tables

            qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No ='" & InvoiceNo & "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No ='" & InvoiceNo & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_VENDOR_INVOICE_DETAIL where Document_No ='" & InvoiceNo & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & InvoiceNo & "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "delete from TSPL_CUSTOM_FIELD_VALUES where Transaction_Code in ('" & InvoiceNo & "','" & Doc_No & "') "
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            '----- shipment data


            'qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & Doc_No & "' and Trans_Type='FS-SH'"
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_BATCH_ITEM where  Document_Code='" & Doc_No & "' and Document_Type='FS-SH'"
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_BATCH_ITEM where  Document_Code='" & Doc_No & "' and Document_Type in ('PS-SH', 'FS-SH')"
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_INVENTORY_MOVEMENT where Source_Doc_No='" & Doc_No & "' and Trans_Type in ('PS-SH', 'FS-SH')"
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No ='" & Doc_No & "')"
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_JOURNAL_MASTER where Source_Doc_No  ='" & Doc_No & "'"
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_SD_SALE_INVOICE_DETAIL where Document_Code='" & InvoiceNo & "' "
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from TSPL_SD_SALE_INVOICE_HEAD where Document_Code='" & InvoiceNo & "' "
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from tspl_sd_shipment_detail where Document_Code='" & Doc_No & "' "
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            'qry = "delete from tspl_sd_shipment_head where Document_Code='" & Doc_No & "' "
            'clsDBFuncationality.ExecuteNonQuery(qry, trans)

            trans.Commit()
            '' release objects 
            obj = Nothing
            qry = Nothing

        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal strInvoiceType As String) As clsVedorInvoiceHead
        Return GetData(strDocumentNo, strInvoiceType, Nothing)
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal strInvoiceType As String, ByVal trans As SqlTransaction) As clsVedorInvoiceHead

        Dim obj As clsVedorInvoiceHead = Nothing
        Dim WhrClause As String = ""
        If clsCommon.myLen(strInvoiceType) > 0 Then
            WhrClause = " and TSPL_VENDOR_INVOICE_HEAD.Invoice_Type in ('" & strInvoiceType & "')"
        End If
        Dim qry As String = "Select *  from TSPL_VENDOR_INVOICE_HEAD"
        qry += " where Document_No='" + strDocumentNo + "' " & WhrClause & "  "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            obj = New clsVedorInvoiceHead()
            obj.RemittanceObject = clsRemittance.GetData(strDocumentNo, trans)
            obj.Document_No = clsCommon.myCstr(dt.Rows(0)("Document_No"))
            obj.Invoice_Entry_Date = clsCommon.myCstr(dt.Rows(0)("Invoice_Entry_Date"))
            If dt.Rows(0)("Invoice_Entry_Date_New") IsNot DBNull.Value Then
                obj.Invoice_Entry_Date_New = clsCommon.myCDate(dt.Rows(0)("Invoice_Entry_Date_New"))
            End If

            obj.Vendor_Code = clsCommon.myCstr(dt.Rows(0)("Vendor_Code"))
            obj.Vendor_Name = clsCommon.myCstr(dt.Rows(0)("Vendor_Name"))
            obj.Vendor_Invoice_No = clsCommon.myCstr(dt.Rows(0)("Vendor_Invoice_No"))
            obj.Vendor_Invoice_Date = clsCommon.myCstr(dt.Rows(0)("Vendor_Invoice_Date"))
            obj.Posting_Date = clsCommon.myCstr(dt.Rows(0)("Posting_Date"))
            obj.is_For_TDS = clsCommon.myCdbl(dt.Rows(0)("is_For_TDS"))
            obj.Account_Set = clsCommon.myCstr(dt.Rows(0)("Account_Set"))
            obj.Document_Type = clsCommon.myCstr(dt.Rows(0)("Document_Type"))
            obj.PO_Number = clsCommon.myCstr(dt.Rows(0)("PO_Number"))
            obj.Invoice_Type = clsCommon.myCstr(dt.Rows(0)("Invoice_Type"))
            obj.Employee_Type = clsCommon.myCstr(dt.Rows(0)("Employee_Type"))
            '---------added by usha
            obj.loc_code = clsCommon.myCstr(dt.Rows(0)("Loc_Code"))
            obj.MCC_Code = clsCommon.myCstr(dt.Rows(0)("MCC_Code"))
            obj.MCC_Name = clsCommon.myCstr(dt.Rows(0)("MCC_Name"))
            obj.Irregular_loc_code = clsCommon.myCstr(dt.Rows(0)("Irregular_Loc_Code"))
            obj.Location_Code_Prefix = clsCommon.myCstr(dt.Rows(0)("Location_Code_Prefix"))
            obj.Vendor_Bank_Code = clsCommon.myCstr(dt.Rows(0)("Vendor_Bank_Code"))
            obj.Vendor_Bank_Name = clsCommon.myCstr(dt.Rows(0)("Vendor_Bank_Name"))
            obj.Branch_IFSC_Code = clsCommon.myCstr(dt.Rows(0)("Branch_IFSC_Code"))
            obj.Branch_Name = clsCommon.myCstr(dt.Rows(0)("Branch_Name"))
            obj.Vendor_Bank_ACNo = clsCommon.myCstr(dt.Rows(0)("Vendor_Bank_ACNo"))
            obj.Against_MCC_Material_Sale = clsCommon.myCstr(dt.Rows(0)("Against_MCC_Material_Sale"))
            obj.Against_MCC_Material_Sale_Return = clsCommon.myCstr(dt.Rows(0)("Against_MCC_Material_Sale_Return"))
            '--------------
            obj.RCM = IIf(clsCommon.myCdbl(dt.Rows(0)("RCM")) = 1, True, False)
            obj.IsEInvoice = IIf(clsCommon.myCdbl(dt.Rows(0)("IsEInvoice")) = 1, True, False)
            obj.TDS_Provision = IIf(clsCommon.myCdbl(dt.Rows(0)("TDS_Provision")) = 1, True, False)
            obj.Transfer_To_Saving = IIf(clsCommon.myCdbl(dt.Rows(0)("Transfer_To_Saving")) = 1, True, False)
            obj.No_GST_Credit = IIf(clsCommon.myCdbl(dt.Rows(0)("No_GST_Credit")) = 1, True, False)
            ''======added by monika========================
            obj.Addition_Doc_Type = clsCommon.myCstr(dt.Rows(0)("Addition_Doc_Type"))
            ''=================================
            obj.GSTRegistered = clsCommon.myCdbl(dt.Rows(0)("GSTRegistered"))
            '' priti starts here
            obj.Purchase_Tax_Invoice = clsCommon.myCstr(dt.Rows(0)("Purchase_Tax_Invoice"))
            obj.Purchase_Tax_Invoice_Type = clsCommon.myCstr(dt.Rows(0)("Purchase_Tax_Invoice_Type"))
            obj.RefDocType = clsCommon.myCstr(dt.Rows(0)("RefDocType"))
            obj.RefDocNo = clsCommon.myCstr(dt.Rows(0)("RefDocNo"))
            obj.Ref_SNo = clsCommon.myCdbl(dt.Rows(0)("Ref_SNo"))
            obj.Against_MillkPurchaseInvoice_No = clsCommon.myCstr(dt.Rows(0)("Against_MillkPurchaseInvoice_No"))
            obj.Against_VSPItemIssue_No = clsCommon.myCstr(dt.Rows(0)("Against_VSPItemIssue_No"))
            obj.Against_VSP_Asset_Issue = clsCommon.myCstr(dt.Rows(0)("Against_VSP_Asset_Issue"))
            If IsDBNull(dt.Rows(0)("DateAndTime")) = True Then
                obj.DateAndTime = Nothing
            Else
                obj.DateAndTime = clsCommon.myCstr(dt.Rows(0)("DateAndTime"))
            End If
            obj.TapalNo = clsCommon.myCstr(dt.Rows(0)("TapalNo"))
            obj.Against_Salary_Generation_Code = clsCommon.myCstr(dt.Rows(0)("Against_Salary_Generation_Code"))
            '' priti ends here
            ' ==================
            If IsDBNull(dt.Rows(0)("Ack_Date")) = True Then
                obj.Ack_Date = Nothing
            Else
                obj.Ack_Date = clsCommon.myCstr(dt.Rows(0)("Ack_Date"))
            End If
            obj.irn_no = clsCommon.myCstr(dt.Rows(0)("irn_no"))
            obj.Ack_No = clsCommon.myCstr(dt.Rows(0)("Ack_No"))
            obj.QR_Code = clsCommon.myCstr(dt.Rows(0)("QR_Code"))
            obj.Order_No = clsCommon.myCstr(dt.Rows(0)("Order_No"))
            obj.Document_Total = clsCommon.myCdbl(dt.Rows(0)("Document_Total"))
            obj.RoundOffAmount = clsCommon.myCdbl(dt.Rows(0)("RoundOffAmount"))
            obj.On_Hold = (clsCommon.CompairString("Y", clsCommon.myCstr(dt.Rows(0)("On_Hold"))) = CompairStringResult.Equal)
            obj.is_For_Provision = clsCommon.myCdbl(dt.Rows(0)("is_For_Provision"))
            obj.isDeduction = clsCommon.myCdbl(dt.Rows(0)("isDeduction"))
            obj.is_Secondary_Transporter_Deduction = IIf(clsCommon.myCdbl(dt.Rows(0)("is_Secondary_Transporter_Deduction")) = 1, True, False)
            '=========================
            obj.Cash_Discount_Amt = clsCommon.myCdbl(dt.Rows(0)("Cash_Discount_Amt"))
            obj.Amt_After_Cash_Discount = clsCommon.myCdbl(dt.Rows(0)("Amt_After_Cash_Discount"))
            '==============
            If obj.is_For_Provision = 1 Then
                obj.Prov_Amt = clsCommon.myCdbl(dt.Rows(0)("Prov_Amt"))
                obj.Prov_From_Date = clsCommon.myCDate(dt.Rows(0)("Prov_From_Date"))
                obj.Prov_To_Date = clsCommon.myCDate(dt.Rows(0)("Prov_To_Date"))

                qry = "select Provision_No from TSPL_PROVISION_ENTRY_KNOCKOFF where AP_Invoice_No='" + obj.Document_No + "'"
                Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    obj.arrProvDocNo = New List(Of String)
                    For Each dr As DataRow In dtTemp.Rows
                        obj.arrProvDocNo.Add(clsCommon.myCstr(dr("Provision_No")))
                    Next
                Else
                    obj.arrProvDocNo = clsProvisionEntry.getProvisionDocNo(obj.Document_No, trans)
                End If
            End If
            obj.Total_Taxable_Amount = clsCommon.myCdbl(dt.Rows(0)("Total_Taxable_Amount"))
            obj.Remarks = clsCommon.myCstr(dt.Rows(0)("Remarks"))
            obj.Description = clsCommon.myCstr(dt.Rows(0)("Description"))
            obj.Tax_Group = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
            obj.TAX1 = clsCommon.myCstr(dt.Rows(0)("TAX1"))
            obj.TAX1_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX1_Rate"))
            obj.Tax1_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax1_BAmount"))
            obj.TAX1_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
            obj.TAX2 = clsCommon.myCstr(dt.Rows(0)("TAX2"))
            obj.TAX2_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX2_Rate"))
            obj.Tax2_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax2_BAmount"))
            obj.TAX2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
            obj.TAX3 = clsCommon.myCstr(dt.Rows(0)("TAX3"))
            obj.Tax3_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax3_BAmount"))
            obj.TAX3_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX3_Rate"))
            obj.TAX3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
            obj.TAX4 = clsCommon.myCstr(dt.Rows(0)("TAX4"))
            obj.TAX4_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX4_Rate"))
            obj.Tax4_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax4_BAmount"))
            obj.TAX4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
            obj.TAX5 = clsCommon.myCstr(dt.Rows(0)("TAX5"))
            obj.TAX5_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX5_Rate"))
            obj.Tax5_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax5_BAmount"))
            obj.TAX5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
            obj.TAX6 = clsCommon.myCstr(dt.Rows(0)("TAX6"))
            obj.TAX6_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX6_Rate"))
            obj.Tax6_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax6_BAmount"))
            obj.TAX6_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt"))
            obj.TAX7 = clsCommon.myCstr(dt.Rows(0)("TAX7"))
            obj.TAX7_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX7_Rate"))
            obj.Tax7_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax7_BAmount"))
            obj.TAX7_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt"))
            obj.TAX8 = clsCommon.myCstr(dt.Rows(0)("TAX8"))
            obj.TAX8_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX8_Rate"))
            obj.Tax8_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax8_BAmount"))
            obj.TAX8_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt"))
            obj.TAX9 = clsCommon.myCstr(dt.Rows(0)("TAX9"))
            obj.TAX9_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX9_Rate"))
            obj.Tax9_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax9_BAmount"))
            obj.TAX9_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt"))
            obj.TAX10 = clsCommon.myCstr(dt.Rows(0)("TAX10"))
            obj.TAX10_Rate = clsCommon.myCdbl(dt.Rows(0)("TAX10_Rate"))
            obj.Tax10_BAmount = clsCommon.myCdbl(dt.Rows(0)("Tax10_BAmount"))
            obj.TAX10_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt"))
            obj.Total_Tax = clsCommon.myCdbl(dt.Rows(0)("Total_Tax"))
            obj.Terms_Code = clsCommon.myCstr(dt.Rows(0)("Terms_Code"))
            obj.Terms_Description = clsCommon.myCstr(dt.Rows(0)("Terms_Description"))
            obj.Due_Date = clsCommon.myCstr(dt.Rows(0)("Due_Date"))
            obj.Discount_Percentage = clsCommon.myCdbl(dt.Rows(0)("Discount_Percentage"))
            obj.Discount_Base = clsCommon.myCdbl(dt.Rows(0)("Discount_Base"))
            obj.Discount_Amount = clsCommon.myCdbl(dt.Rows(0)("Discount_Amount"))
            obj.Amount_Less_Discount = clsCommon.myCdbl(dt.Rows(0)("Amount_Less_Discount"))
            obj.TDS_Base_Actual_Amount = clsCommon.myCdbl(dt.Rows(0)("TDS_Base_Actual_Amount"))
            obj.TDS_Base_Calculated_Amount = clsCommon.myCdbl(dt.Rows(0)("TDS_Base_Calculated_Amount"))
            obj.TDS_Percentage = clsCommon.myCdbl(dt.Rows(0)("TDS_Percentage"))
            obj.TDS_Actual_Amount = clsCommon.myCdbl(dt.Rows(0)("TDS_Actual_Amount"))
            obj.TDS_Calculated_Amount = clsCommon.myCdbl(dt.Rows(0)("TDS_Calculated_Amount"))
            obj.Nature_of_deduction = clsCommon.myCstr(dt.Rows(0)("Nature_of_deduction"))
            obj.Branch_Code = clsCommon.myCstr(dt.Rows(0)("Branch_Code"))
            obj.Section_Code = clsCommon.myCstr(dt.Rows(0)("Section_Code"))
            obj.Comp_Code = clsCommon.myCstr(dt.Rows(0)("Comp_Code"))
            obj.Balance_Amt = clsCommon.myCdbl(dt.Rows(0)("Balance_Amt"))
            obj.Vendor_Control_AC = clsCommon.myCstr(dt.Rows(0)("Vendor_Control_AC"))
            obj.Discount_GL_AC = clsCommon.myCstr(dt.Rows(0)("Discount_GL_AC"))

            obj.TAX1_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX1_GLAC"))
            obj.TAX1_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_GLAC_Amt"))
            obj.TAX1_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX1_GLAC2"))
            obj.TAX1_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_GLAC2_Amt"))
            obj.TAX1_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX1_GLAC3"))
            obj.TAX1_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_GLAC3_Amt"))
            obj.TAX1_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX1_GLAC4"))
            obj.TAX1_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_GLAC4_Amt"))
            obj.TAX1_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX1_GLAC5"))
            obj.TAX1_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX1_GLAC5_Amt"))


            obj.TAX2_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX2_GLAC"))
            obj.TAX2_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_GLAC_Amt"))
            obj.TAX2_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX2_GLAC2"))
            obj.TAX2_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_GLAC2_Amt"))
            obj.TAX2_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX2_GLAC3"))
            obj.TAX2_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_GLAC3_Amt"))
            obj.TAX2_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX2_GLAC4"))
            obj.TAX2_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_GLAC4_Amt"))
            obj.TAX2_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX2_GLAC5"))
            obj.TAX2_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX2_GLAC5_Amt"))

            obj.TAX3_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX3_GLAC"))
            obj.TAX3_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_GLAC_Amt"))
            obj.TAX3_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX3_GLAC2"))
            obj.TAX3_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_GLAC2_Amt"))
            obj.TAX3_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX3_GLAC3"))
            obj.TAX3_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_GLAC3_Amt"))
            obj.TAX3_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX3_GLAC4"))
            obj.TAX3_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_GLAC4_Amt"))
            obj.TAX3_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX3_GLAC5"))
            obj.TAX3_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX3_GLAC5_Amt"))

            obj.TAX4_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX4_GLAC"))
            obj.TAX4_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_GLAC_Amt"))
            obj.TAX4_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX4_GLAC2"))
            obj.TAX4_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_GLAC2_Amt"))
            obj.TAX4_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX4_GLAC3"))
            obj.TAX4_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_GLAC3_Amt"))
            obj.TAX4_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX4_GLAC4"))
            obj.TAX4_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_GLAC4_Amt"))
            obj.TAX4_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX4_GLAC5"))
            obj.TAX4_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX4_GLAC5_Amt"))

            obj.TAX5_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX5_GLAC"))
            obj.TAX5_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_GLAC_Amt"))
            obj.TAX5_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX5_GLAC2"))
            obj.TAX5_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_GLAC2_Amt"))
            obj.TAX5_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX5_GLAC3"))
            obj.TAX5_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_GLAC3_Amt"))
            obj.TAX5_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX5_GLAC4"))
            obj.TAX5_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_GLAC4_Amt"))
            obj.TAX5_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX5_GLAC5"))
            obj.TAX5_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX5_GLAC5_Amt"))

            obj.TAX6_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX6_GLAC"))
            obj.TAX6_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_GLAC_Amt"))
            obj.TAX6_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX6_GLAC2"))
            obj.TAX6_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_GLAC2_Amt"))
            obj.TAX6_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX6_GLAC3"))
            obj.TAX6_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_GLAC3_Amt"))
            obj.TAX6_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX6_GLAC4"))
            obj.TAX6_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_GLAC4_Amt"))
            obj.TAX6_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX6_GLAC5"))
            obj.TAX6_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX6_GLAC5_Amt"))

            obj.TAX7_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX7_GLAC"))
            obj.TAX7_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_GLAC_Amt"))
            obj.TAX7_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX7_GLAC2"))
            obj.TAX7_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_GLAC2_Amt"))
            obj.TAX7_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX7_GLAC3"))
            obj.TAX7_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_GLAC3_Amt"))
            obj.TAX7_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX7_GLAC4"))
            obj.TAX7_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_GLAC4_Amt"))
            obj.TAX7_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX7_GLAC5"))
            obj.TAX7_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX7_GLAC5_Amt"))

            obj.TAX8_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX8_GLAC"))
            obj.TAX8_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_GLAC_Amt"))
            obj.TAX8_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX8_GLAC2"))
            obj.TAX8_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_GLAC2_Amt"))
            obj.TAX8_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX8_GLAC3"))
            obj.TAX8_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_GLAC3_Amt"))
            obj.TAX8_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX8_GLAC4"))
            obj.TAX8_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_GLAC4_Amt"))
            obj.TAX8_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX8_GLAC5"))
            obj.TAX8_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX8_GLAC5_Amt"))

            obj.TAX9_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX9_GLAC"))
            obj.TAX9_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_GLAC_Amt"))
            obj.TAX9_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX9_GLAC2"))
            obj.TAX9_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_GLAC2_Amt"))
            obj.TAX9_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX9_GLAC3"))
            obj.TAX9_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_GLAC3_Amt"))
            obj.TAX9_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX9_GLAC4"))
            obj.TAX9_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_GLAC4_Amt"))
            obj.TAX9_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX9_GLAC5"))
            obj.TAX9_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX9_GLAC5_Amt"))

            obj.TAX10_GLAC = clsCommon.myCstr(dt.Rows(0)("TAX10_GLAC"))
            obj.TAX10_GLAC_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_GLAC_Amt"))
            obj.TAX10_GLAC2 = clsCommon.myCstr(dt.Rows(0)("TAX10_GLAC2"))
            obj.TAX10_GLAC2_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_GLAC2_Amt"))
            obj.TAX10_GLAC3 = clsCommon.myCstr(dt.Rows(0)("TAX10_GLAC3"))
            obj.TAX10_GLAC3_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_GLAC3_Amt"))
            obj.TAX10_GLAC4 = clsCommon.myCstr(dt.Rows(0)("TAX10_GLAC4"))
            obj.TAX10_GLAC4_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_GLAC4_Amt"))
            obj.TAX10_GLAC5 = clsCommon.myCstr(dt.Rows(0)("TAX10_GLAC5"))
            obj.TAX10_GLAC5_Amt = clsCommon.myCdbl(dt.Rows(0)("TAX10_GLAC5_Amt"))

            obj.PROJECT_ID = clsCommon.myCstr(dt.Rows(0)("PROJECT_ID"))
            obj.Against_POInvoice_No = clsCommon.myCstr(dt.Rows(0)("Against_POInvoice_No"))
            obj.Against_PurchaseReturn_No = clsCommon.myCstr(dt.Rows(0)("Against_PurchaseReturn_No"))
            '' Anubhooti 19-Mar-2015 (Against_VCGL)
            obj.Against_VCGL = clsCommon.myCstr(dt.Rows(0)("Against_VCGL"))
            obj.Hirerachy_Level_Code = clsCommon.myCstr(dt.Rows(0)("Hirerachy_Level_Code"))
            obj.Cost_Centre_Fin_Level_Code = clsCommon.myCstr(dt.Rows(0)("Cost_Centre_Fin_Level_Code"))
            ''
            obj.Against_BulkMillkPurchaseInvoice_No = clsCommon.myCstr(dt.Rows(0)("Against_BulkMillkPurchaseInvoice_No"))
            obj.Against_Acquisition = clsCommon.myCstr(dt.Rows(0)("Against_Acquisition"))
            obj.Add_Charge_Code1 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code1"))
            obj.Add_Charge_Name1 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name1"))
            obj.Add_Charge_Amt1 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt1"))

            obj.Add_Charge_Code2 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code2"))
            obj.Add_Charge_Name2 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name2"))
            obj.Add_Charge_Amt2 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt2"))

            obj.Add_Charge_Code3 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code3"))
            obj.Add_Charge_Name3 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name3"))
            obj.Add_Charge_Amt3 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt3"))

            obj.Add_Charge_Code4 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code4"))
            obj.Add_Charge_Name4 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name4"))
            obj.Add_Charge_Amt4 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt4"))

            obj.Add_Charge_Code5 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code5"))
            obj.Add_Charge_Name5 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name5"))
            obj.Add_Charge_Amt5 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt5"))

            obj.Add_Charge_Code6 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code6"))
            obj.Add_Charge_Name6 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name6"))
            obj.Add_Charge_Amt6 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt6"))

            obj.Add_Charge_Code7 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code7"))
            obj.Add_Charge_Name7 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name7"))
            obj.Add_Charge_Amt7 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt7"))

            obj.Add_Charge_Code8 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code8"))
            obj.Add_Charge_Name8 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name8"))
            obj.Add_Charge_Amt8 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt8"))

            obj.Add_Charge_Code9 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code9"))
            obj.Add_Charge_Name9 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name9"))
            obj.Add_Charge_Amt9 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt9"))

            obj.Add_Charge_Code10 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Code10"))
            obj.Add_Charge_Name10 = clsCommon.myCstr(dt.Rows(0)("Add_Charge_Name10"))
            obj.Add_Charge_Amt10 = clsCommon.myCdbl(dt.Rows(0)("Add_Charge_Amt10"))

            obj.Empty_Amount = clsCommon.myCdbl(dt.Rows(0)("Empty_Amount"))
            obj.Empty_Account = clsCommon.myCstr(dt.Rows(0)("Empty_Account"))
            obj.Total_Landed_Amt = clsCommon.myCdbl(dt.Rows(0)("Total_Landed_Amt"))

            obj.Total_Add_Charge = clsCommon.myCdbl(dt.Rows(0)("Total_Add_Charge"))
            obj.Tax_Calculation_Type = IIf(clsCommon.myCdbl(dt.Rows(0)("Tax_Calculation_Type")) = 0, EnumTaxCalucationType.Automatic, EnumTaxCalucationType.Mannual)
            obj.Is_ProRated = clsCommon.myCstr(dt.Rows(0)("Is_ProRated"))

            '' CURRENCYCONVERSION 
            obj.ITC_Type_Category = clsCommon.myCstr(dt.Rows(0)("ITC_Type_Category"))
            obj.ITC_Elibible = IIf(clsCommon.myCdbl(dt.Rows(0)("ITC_Elibible")) = 1, 1, 0)
            obj.ITC_Type = clsCommon.myCdbl(dt.Rows(0)("ITC_Type"))
            obj.Saving = clsCommon.myCDecimal(dt.Rows(0)("Saving"))
            obj.CURRENCY_CODE = clsCommon.myCstr(dt.Rows(0)("CURRENCY_CODE"))
            obj.ConvRate = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
            obj.Security = clsCommon.myCdbl(dt.Rows(0)("Is_Security"))

            If IsDBNull(dt.Rows(0)("ApplicableFrom")) = True Then
                obj.ApplicableFrom = Nothing
            Else
                obj.ApplicableFrom = clsCommon.GetPrintDate(dt.Rows(0)("ApplicableFrom"), "dd/MMM/yyyy")
            End If

            obj.Main_VSP_Milk_AP_Invoice_No = clsCommon.myCstr(dt.Rows(0)("Main_VSP_Milk_AP_Invoice_No"))
            '' END CURRENCYCONVERSION
            obj.Against_Asset_Work = clsCommon.myCstr(dt.Rows(0)("Against_Asset_Work"))
            obj.Against_TransferToSavingPKID = clsCommon.myCDecimal(dt.Rows(0)("Against_TransferToSavingPKID"))
            obj.ArrAssetEMI = clsAPInvoiceAssetEMIDetails.GetData(obj.Document_No, trans)
            obj.ArrAdvanceInterest = clsAPInvoiceAdvanceInterest.GetData(obj.Document_No, trans)

            qry = "Select TSPL_VENDOR_INVOICE_SECURITY_REFUND.Against_Security from TSPL_VENDOR_INVOICE_SECURITY_REFUND  where AP_Invoice_No='" + strDocumentNo + "' ORDER BY PK_ID"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.arrSecurityAdjustment = New ArrayList
                For Each dr As DataRow In dt.Rows
                    obj.arrSecurityAdjustment.Add(dr("Against_Security"))
                Next
            End If

            qry = "Select TSPL_VENDOR_INVOICE_DETAIL.*,TSPL_SAC_MASTER.Code as SAC_Code,TSPL_SAC_MASTER.Description as SAC_Name,TSPL_TDS_DEDUCTION_HEAD.Description as Deduction_Name,TSPL_TDS_DEDUCTION_HEAD.TDS_Section as Deduction_Section,TSPL_ACQUISITION_DETAIL.Asset_Name,case when len(isnull(TSPL_VENDOR_INVOICE_DETAIL.Deduction_Desc,''))>0 then TSPL_VENDOR_INVOICE_DETAIL.Deduction_Desc else TSPL_DEDUCTION_MASTER.Description end as DEDUCTION_DESC_New from TSPL_VENDOR_INVOICE_DETAIL left outer join TSPL_TDS_DEDUCTION_HEAD on TSPL_TDS_DEDUCTION_HEAD.Deduction_Code=TSPL_VENDOR_INVOICE_DETAIL.Deduction_Code left outer join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER.Code=TSPL_VENDOR_INVOICE_DETAIL.DeductionCode left join TSPL_ACQUISITION_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Asset_Code=TSPL_ACQUISITION_DETAIL.Asset_Code left outer join TSPL_SAC_MASTER on TSPL_SAC_MASTER.Code=TSPL_VENDOR_INVOICE_DETAIL.SAC_Code where Document_No='" + strDocumentNo + "' ORDER BY Detail_Line_No"
            dt = New DataTable()
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                obj.Arr = New List(Of clsVedorInvoiceDetail)
                Dim objTr As clsVedorInvoiceDetail
                For Each dr As DataRow In dt.Rows
                    objTr = New clsVedorInvoiceDetail

                    '**********************************************************************
                    objTr.chrgcatcode = clsCommon.myCstr(dr("charge_cat_code"))
                    objTr.chrgcatdesc = clsCommon.myCstr(dr("charge_cat_desc"))
                    objTr.chritemcode = clsCommon.myCstr(dr("item_code"))
                    objTr.chritemdesc = clsCommon.myCstr(dr("item_desc"))
                    objTr.chrgcatvalue = clsCommon.myCstr(dr("charge_cat_charges"))
                    '*****************************************************************
                    objTr.DeductionCode = clsCommon.myCstr(dr("DeductionCode"))
                    objTr.DeductionDesc = clsCommon.myCstr(dr("DEDUCTION_DESC_New"))

                    ''richa agarwal
                    objTr.AgainstPayment_No = clsCommon.myCstr(dr("AgainstPayment_No"))
                    objTr.Payment_Amount = clsCommon.myCdbl(dr("Payment_Amount"))
                    objTr.TDS_Per = clsCommon.myCdbl(dr("TDS_Per"))
                    objTr.SAC_Code = clsCommon.myCstr(dr("SAC_Code"))
                    objTr.SAC_Name = clsCommon.myCstr(dr("SAC_Name"))
                    '--------------------------

                    objTr.Document_No = clsCommon.myCstr(dr("Document_No"))
                    objTr.Detail_Line_No = clsCommon.myCstr(dr("Detail_Line_No"))

                    objTr.Deduction_Code = clsCommon.myCstr(dr("Deduction_Code"))
                    objTr.Deduction_Name = clsCommon.myCstr(dr("Deduction_Name"))
                    objTr.Deduction_Section = clsCommon.myCstr(dr("Deduction_Section"))
                    objTr.DCS_Addition_Deduction = clsCommon.myCstr(dr("DCS_Addition_Deduction"))
                    objTr.GL_Account_Code = clsCommon.myCstr(dr("GL_Account_Code"))
                    objTr.GL_Account_Desc = clsCommon.myCstr(dr("GL_Account_Desc"))
                    objTr.Amount = clsCommon.myCdbl(dr("Amount"))
                    objTr.Discount_Per = clsCommon.myCdbl(dr("Discount_Per"))
                    objTr.Discount = clsCommon.myCdbl(dr("Discount"))
                    objTr.Amount_less_Discount = clsCommon.myCdbl(dr("Amount_less_Discount"))

                    objTr.Taxable_Amount = clsCommon.myCdbl(dr("Taxable_Amount"))
                    objTr.Taxable_Amount_Per = clsCommon.myCdbl(dr("Taxable_Amount_Per"))

                    objTr.TAX1 = clsCommon.myCstr(dr("TAX1"))
                    objTr.TAX1_Rate = clsCommon.myCdbl(dr("TAX1_Rate"))
                    objTr.TAX1_Amt = clsCommon.myCdbl(dr("TAX1_Amt"))
                    objTr.TAX2 = clsCommon.myCstr(dr("TAX2"))
                    objTr.TAX2_Rate = clsCommon.myCdbl(dr("TAX2_Rate"))
                    objTr.TAX2_Amt = clsCommon.myCdbl(dr("TAX2_Amt"))
                    objTr.TAX3 = clsCommon.myCstr(dr("TAX3"))
                    objTr.TAX3_Rate = clsCommon.myCdbl(dr("TAX3_Rate"))
                    objTr.TAX3_Amt = clsCommon.myCdbl(dr("TAX3_Amt"))
                    objTr.TAX4 = clsCommon.myCstr(dr("TAX4"))
                    objTr.TAX4_Rate = clsCommon.myCdbl(dr("TAX4_Rate"))
                    objTr.TAX4_Amt = clsCommon.myCdbl(dr("TAX4_Amt"))
                    objTr.TAX5 = clsCommon.myCstr(dr("TAX5"))
                    objTr.TAX5_Rate = clsCommon.myCdbl(dr("TAX5_Rate"))
                    objTr.TAX5_Amt = clsCommon.myCdbl(dr("TAX5_Amt"))
                    objTr.TAX6 = clsCommon.myCstr(dr("TAX6"))
                    objTr.TAX6_Rate = clsCommon.myCdbl(dr("TAX6_Rate"))
                    objTr.TAX6_Amt = clsCommon.myCdbl(dr("TAX6_Amt"))
                    objTr.TAX7 = clsCommon.myCstr(dr("TAX7"))
                    objTr.TAX7_Rate = clsCommon.myCdbl(dr("TAX7_Rate"))
                    objTr.TAX7_Amt = clsCommon.myCdbl(dr("TAX7_Amt"))
                    objTr.TAX8 = clsCommon.myCstr(dr("TAX8"))
                    objTr.TAX8_Rate = clsCommon.myCdbl(dr("TAX8_Rate"))
                    objTr.TAX8_Amt = clsCommon.myCdbl(dr("TAX8_Amt"))
                    objTr.TAX9 = clsCommon.myCstr(dr("TAX9"))
                    objTr.TAX9_Rate = clsCommon.myCdbl(dr("TAX9_Rate"))
                    objTr.TAX9_Amt = clsCommon.myCdbl(dr("TAX9_Amt"))
                    objTr.TAX10 = clsCommon.myCstr(dr("TAX10"))
                    objTr.TAX10_Rate = clsCommon.myCdbl(dr("TAX10_Rate"))
                    objTr.TAX10_Amt = clsCommon.myCdbl(dr("TAX10_Amt"))
                    objTr.Total_Tax = clsCommon.myCdbl(dr("Total_Tax"))
                    objTr.Total_Amount = clsCommon.myCdbl(dr("Total_Amount"))
                    objTr.Remarks = clsCommon.myCstr(dr("Remarks"))
                    objTr.Comments = clsCommon.myCstr(dr("Comments"))
                    objTr.AddChargeCode = clsCommon.myCstr(dr("AddChargeCode"))
                    objTr.AddChargeDesc = clsCommon.myCstr(dr("AddChargeDesc"))
                    objTr.is_Unclaimed_Tax = IIf(clsCommon.myCdbl(dr("is_Unclaimed_Tax") = 1), True, False)
                    objTr.Invoice_No = clsCommon.myCstr(dr("Invoice_No"))
                    objTr.Invoice_Type = clsCommon.myCstr(dr("Invoice_Type"))
                    objTr.Landed_Amount = clsCommon.myCdbl(dr("Landed_Amount"))
                    objTr.Item_Rate = clsCommon.myCdbl(dr("Item_Rate"))
                    objTr.Abatement_Per = clsCommon.myCdbl(dr("Abatement_Per"))
                    objTr.Abatement_Amt = clsCommon.myCdbl(dr("Abatement_Amt"))

                    objTr.Reverse_Charge_Per = clsCommon.myCdbl(dr("Reverse_Charge_Per"))
                    objTr.Reverse_Charge_Amount = clsCommon.myCdbl(dr("Reverse_Charge_Amount"))

                    objTr.TAX1_Base_Amt = clsCommon.myCdbl(dr("TAX1_Base_Amt"))
                    objTr.TAX2_Base_Amt = clsCommon.myCdbl(dr("TAX2_Base_Amt"))
                    objTr.TAX3_Base_Amt = clsCommon.myCdbl(dr("TAX3_Base_Amt"))
                    objTr.TAX4_Base_Amt = clsCommon.myCdbl(dr("TAX4_Base_Amt"))
                    objTr.TAX5_Base_Amt = clsCommon.myCdbl(dr("TAX5_Base_Amt"))
                    objTr.TAX6_Base_Amt = clsCommon.myCdbl(dr("TAX6_Base_Amt"))
                    objTr.TAX7_Base_Amt = clsCommon.myCdbl(dr("TAX7_Base_Amt"))
                    objTr.TAX8_Base_Amt = clsCommon.myCdbl(dr("TAX8_Base_Amt"))
                    objTr.TAX9_Base_Amt = clsCommon.myCdbl(dr("TAX9_Base_Amt"))
                    objTr.TAX10_Base_Amt = clsCommon.myCdbl(dr("TAX10_Base_Amt"))
                    ''shivani
                    objTr.Hirerachy_Code = clsCommon.myCstr(dr("Hirerachy_Code"))
                    objTr.Cost_Centre_Code = clsCommon.myCstr(dr("Cost_Centre_Code"))
                    objTr.Hirerachy_Code1 = clsCommon.myCstr(dr("Hirerachy_Code1"))
                    objTr.Hirerachy_Code2 = clsCommon.myCstr(dr("Hirerachy_Code2"))
                    objTr.Hirerachy_Code3 = clsCommon.myCstr(dr("Hirerachy_Code3"))
                    objTr.Hirerachy_Code4 = clsCommon.myCstr(dr("Hirerachy_Code4"))
                    '' PanchRaj
                    objTr.Item_Type = clsCommon.myCstr(dr("Item_Type"))
                    objTr.Asset_Code = clsCommon.myCstr(dr("Asset_Code"))
                    objTr.Asset_Desc = clsCommon.myCstr(dr("Asset_Name"))
                    objTr.PK_Id = clsCommon.myCdbl(dr("PK_Id"))
                    If obj.is_Secondary_Transporter_Deduction Then
                        objTr.arrSTDed = clsAPSecondaryTranporterDeductionDetail.GetData(objTr.Document_No, objTr.PK_Id, trans)
                    End If
                    obj.Arr.Add(objTr)
                Next
            End If
        End If
        Return obj
    End Function

    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal strRefDocNo As String) As Boolean
        Dim isSaved As Boolean = False
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            isSaved = clsVedorInvoiceHead.PostData(FormId, strDocNo, strRefDocNo, trans)
            If isSaved Then
                trans.Commit()
            Else
                trans.Rollback()
            End If
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return isSaved
    End Function

    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal strRefDocNo As String, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoRecreatedOnly As String = Nothing, Optional ByVal strasset As Boolean = False) As Boolean
        Dim strPostDate As Date = clsCommon.myCDate(clsDBFuncationality.getSingleValue("select Invoice_Entry_Date from TSPL_VENDOR_INVOICE_HEAD where Document_No='" + strDocNo + "'", trans))
        Return PostData(FormId, strDocNo, strRefDocNo, trans, strPostDate, strVoucherNoRecreatedOnly)
    End Function

    Public Shared Function PostData(ByVal FormId As String, ByVal strDocNo As String, ByVal strRefDocNo As String, ByVal trans As SqlTransaction, ByVal strPostDate As Date, Optional ByVal strVoucherNoRecreatedOnly As String = Nothing, Optional ByVal strasset As Boolean = False) As Boolean
        Dim qry As String = ""
        Dim isTaxRecoverable As Boolean = False
        If (clsCommon.myLen(strDocNo) <= 0) Then
            Throw New Exception("Document No not found to Post")
        End If

        Dim obj As clsVedorInvoiceHead = clsVedorInvoiceHead.GetData(strDocNo, "", trans)
        If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0) Then
            Throw New Exception("No Data found to Post")
        End If
        ' ''richa agarwal 
        'If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*)  from TSPL_VENDOR_INVOICE_HEAD where convert(date,Invoice_Entry_Date,103)>= convert(date,'" & objCommonVar.CurrFiscalStartDate & "',103)  and convert(datetime,Invoice_Entry_Date,103)<= convert(datetime ,'" & objCommonVar.CurrFiscalEndDate & "',103)  and isnull(Vendor_Invoice_No,'') ='" & obj.Vendor_Invoice_No & "' and Vendor_Code ='" & obj.Vendor_Code & "' and Document_Type ='I' and Document_No<>'" & obj.Document_No & "'  ", trans)) >= 1 Then
        '    Throw New Exception("Vendor Invoice No. " & obj.Vendor_Invoice_No & " is already exist for Vendor " & obj.Vendor_Code & " in current financial year " & objCommonVar.CurrFiscalYear & ". ")
        'End If
        ' ''-------------------

        ''richa agarwal 
        Dim strcurrentfisyearstartdate As Date? = Nothing
        Dim strcurrentfisyearenddate As Date? = Nothing
        Dim strcurrentfisyear As String = String.Empty

        Dim strmonth As Integer = Convert.ToDateTime(obj.Invoice_Entry_Date).Month()
        Dim stryear As Integer = Convert.ToDateTime(obj.Invoice_Entry_Date).Year()
        If strmonth <= 3 Then
            strcurrentfisyearstartdate = "01-Apr-" + clsCommon.myCstr(stryear - 1)
            strcurrentfisyearenddate = "31-Mar-" + clsCommon.myCstr(stryear)
            strcurrentfisyear = clsCommon.myCstr(stryear - 1) + "-" + clsCommon.myCstr(stryear)
        Else
            strcurrentfisyearstartdate = "01-Apr-" + clsCommon.myCstr(stryear)
            strcurrentfisyearenddate = "31-Mar-" + clsCommon.myCstr(stryear + 1)
            strcurrentfisyear = clsCommon.myCstr(stryear) + "-" + clsCommon.myCstr(stryear + 1)
        End If
        If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Vendor_Invoice_No) > 0 Then
            ''=================03/10/2016===================
            Dim addwhrCls As String = ""
            If clsCommon.myLen(obj.Addition_Doc_Type) > 0 Then
                addwhrCls = " and Addition_Doc_Type='" + obj.Addition_Doc_Type + "' "
            End If
            ''=============================================

            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select count(*)  from TSPL_VENDOR_INVOICE_HEAD where convert(date,Invoice_Entry_Date,103)>= convert(date,'" & strcurrentfisyearstartdate & "',103)  and convert(datetime,Invoice_Entry_Date,103)<= convert(datetime ,'" & strcurrentfisyearenddate & "',103)  and isnull(Vendor_Invoice_No,'') ='" & obj.Vendor_Invoice_No & "' and Vendor_Code ='" & obj.Vendor_Code & "' and Document_Type ='I' and Document_No<>'" & obj.Document_No & "' " + addwhrCls + " ", trans)) >= 1 Then
                Throw New Exception("Vendor Invoice No. " & obj.Vendor_Invoice_No & " is already exist for Vendor " & obj.Vendor_Code & " in current financial year " & strcurrentfisyear & ". ")
            End If
        End If

        ''------------------

        clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePayable, clsUserMgtCode.mbtnAPInvoiceEntry, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
        clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleTDS, clsUserMgtCode.mbtnAPInvoiceEntryTDS, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
        If (clsCommon.myLen(obj.Posting_Date) > 0) Then
            Throw New Exception("Already Post on :" + obj.Posting_Date)
        End If
        If (obj.On_Hold) Then
            Throw New Exception("Document No " + obj.Document_No + " Is currently On Hold.Can't Post it")
        End If
        If clsCommon.myLen(obj.Vendor_Control_AC) <= 0 Then
            Throw New Exception("Vendor's Control A/C Not found")
        End If

        Dim isResult As Boolean = clsApprovalScreen.CheckApprovalLevel(FormId, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", strDocNo, trans)
        If isResult = False Then
            trans.Commit()
            Return False
        End If


        ''richa 24 Dec,2018  TEC/02/11/18-000361 create journal entry for opening in case of Credit or debit note as Journal Master table instead of journal master op table
        'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, trans)
        Dim JEWithOPening As Boolean = False
        If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
            Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
            If clsCommon.myCDate(clsCommon.GetPrintDate(obj.Invoice_Entry_Date, "dd/MM/yyyy")) <= clsCommon.myCDate(clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy")) Then
                JEWithOPening = True
            End If
        End If

        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, trans)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "C") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "D") = CompairStringResult.Equal) And JEWithOPening = True Then
            CreateJournalEntryForOpening(obj, FormId, strPostDate, trans, strVoucherNoRecreatedOnly, False, strasset)
        Else
            CreateJournalEntry(obj, FormId, strPostDate, trans, strVoucherNoRecreatedOnly, False, strasset)
        End If
        ''-----------------
        If obj.TDS_Provision Then
            qry = "update TSPL_REMITTANCE set Remit_TDS='N',Is_TDS_Provision='Y' where Document_No='" + strDocNo + "'"
        Else
            qry = "update TSPL_REMITTANCE set Remit_TDS='N',Is_TDS_Provision='N' where Document_No='" + strDocNo + "'"

        End If
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Dim pono As String = obj.Against_POInvoice_No
        Dim amt As Decimal = obj.Document_Total
        Dim UpdatePRAPInvoiceBalanceAmt As Boolean = False
        If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
            UpdatePRAPInvoiceBalanceAmt = True
            If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurReturnEvenIfPaymentDone, clsFixedParameterCode.AllowPurReturnEvenIfPaymentDone, trans)) = 1 Then
                Dim strAgainst_PI As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Against_PI from TSPL_PR_HEAD where PR_No='" + obj.Against_PurchaseReturn_No + "'", trans))
                Dim dblPendingBalance As Double = clsPurchasReturnHead.GetPIBalance(obj.Vendor_Code, obj.Against_PurchaseReturn_No, strAgainst_PI, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
                If dblPendingBalance = 0 Then
                    UpdatePRAPInvoiceBalanceAmt = False
                End If
                If obj.Document_Total > dblPendingBalance Then
                    UpdatePRAPInvoiceBalanceAmt = False
                End If
            End If

            qry = "update TSPL_VENDOR_INVOICE_HEAD set Update_PR_APInvoice_Balance_Amt=" + IIf(UpdatePRAPInvoiceBalanceAmt, "1", "0") + "  "

            If UpdatePRAPInvoiceBalanceAmt Then
                qry += ",Balance_Amt=Balance_Amt - " + clsCommon.myCstr(obj.Document_Total - obj.TDS_Actual_Amount)
            End If
            qry += " where Against_POInvoice_No  in ( select Against_PI from TSPL_PR_HEAD where PR_No='" + obj.Against_PurchaseReturn_No + "')"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        End If

        If clsCommon.CompairString(obj.RefDocType, "AP") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.RefDocNo) > 0 Then
            Dim strOpearateor = "+"
            If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                strOpearateor = "-"
            End If
            Dim dblRemAmt As Double = 0
            If (obj.RemittanceObject IsNot Nothing) Then
                dblRemAmt = obj.RemittanceObject.Actual_Total_TDS
            End If

            qry = "update TSPL_VENDOR_INVOICE_HEAD set Balance_Amt=Balance_Amt" + strOpearateor + " " + clsCommon.myCstr(obj.Document_Total - dblRemAmt) + " where Document_No='" + obj.RefDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)

            qry = "select Balance_Amt  from TSPL_VENDOR_INVOICE_HEAD where  Document_No='" + obj.RefDocNo + "'"
            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans)) < 0 Then
                Throw New Exception("Balance is Going to Negative For AP Invoice No : " + obj.RefDocNo)
            End If
        End If
        Dim objCostAdj As ClsAdjustments = Nothing
        If clsCommon.CompairString(obj.Is_ProRated, "Y") = CompairStringResult.Equal Then
            Dim Purchase_Control_Account As String = ""
            If clsCommon.CompairString(obj.RefDocType, "BS") = CompairStringResult.Equal Then
                Dim ObjSRN As clsBulkMilkSRN = clsBulkMilkSRN.getData(obj.RefDocNo, NavigatorType.Current, False, trans)
                Dim PurchaseControlAmount As Decimal = 0.0
                If ObjSRN IsNot Nothing And clsCommon.myLen(ObjSRN.SRN_NO) > 0 Then
                    objCostAdj = New ClsAdjustments()
                    objCostAdj.Arr = New List(Of ClsAdjustmentsDetails)
                    For Each objTR As clsVedorInvoiceDetail In obj.Arr
                        Dim intCount As Integer = obj.Arr.Count
                        Dim dblLedgeerNonRecoverableAmt As Double = clsVedorInvoiceHead.GetTaxAmt(objTR, trans)
                        Dim dblAddionalCost As Double = Math.Round((obj.Total_Add_Charge / intCount), 6)
                        Dim tempAmt As Double = objTR.Amount_less_Discount + dblAddionalCost + dblLedgeerNonRecoverableAmt
                        PurchaseControlAmount += tempAmt
                    Next

                    '-------------Cost Adjustment Entry----------------
                    objCostAdj.Reference_Document = ObjSRN.SRN_NO
                    objCostAdj.Adjustment_Date = obj.Invoice_Entry_Date
                    objCostAdj.Against_AP_Invoice_No = obj.Document_No
                    objCostAdj.Reference = "Cost Adjustment Against AP Invoice : " + obj.Document_No + " & Bulk SRN : " + ObjSRN.SRN_NO + "."
                    objCostAdj.Description = "Cost Adjustment Against AP Invoice : " + obj.Document_No + " & Bulk SRN : " + ObjSRN.SRN_NO + "."

                    objCostAdj.Unit_Code = "ALL"
                    objCostAdj.IsMilkType = 1
                    objCostAdj.ItemType = ""
                    objCostAdj.Loc_Code = obj.loc_code
                    objCostAdj.Loc_Desc = clsLocation.GetName(obj.loc_code, trans)
                    If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                        objCostAdj.Trans_Type = "Out"
                    ElseIf clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal Then
                        objCostAdj.Trans_Type = "In"
                    End If



                    Dim objCostAdjTr As New ClsAdjustmentsDetails()
                    objCostAdjTr.Adjustment_Line_No = 1
                    objCostAdjTr.Item_Code = ObjSRN.Item_Code
                    objCostAdjTr.Item_Description = ObjSRN.Item_Desc
                    If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                        objCostAdjTr.Adjustment_Type = "CD"
                    ElseIf clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal Then
                        objCostAdjTr.Adjustment_Type = "CI"
                    End If

                    objCostAdjTr.Item_Quantity = 0
                    objCostAdjTr.Item_Cost = PurchaseControlAmount
                    objCostAdjTr.Unit_Code = ObjSRN.UOM
                    objCostAdjTr.Remarks = ""
                    objCostAdjTr.Comments = ""
                    'objCostAdjTr.mrp = clsCommon.myCdbl(ObjSRN.MRP)

                    objCostAdjTr.BreakageType = ""
                    objCostAdjTr.Breakage = 0.0
                    objCostAdjTr.Breakage_Cost = 0.0
                    objCostAdjTr.LeakageQty = 0.0
                    objCostAdjTr.fat_pers = ObjSRN.fat_per
                    objCostAdjTr.fat_kg = ObjSRN.fat_KG
                    objCostAdjTr.snf_pers = ObjSRN.snf_Per
                    objCostAdjTr.snf_kg = ObjSRN.SNF_KG
                    'objCostAdjTr.ItemType = ObjSRN.Item_Type
                    'objCostAdjTr.Basic_Price = clsItemBasicPrice.GetBasicPrice(objCostAdjTr.Item_Code, objCostAdjTr.mrp, trans)
                    'objCostAdjTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))
                    objCostAdj.Arr.Add(objCostAdjTr)
                End If
            Else
                Dim ObjSRN As clsSRNHead = clsSRNHead.GetData(obj.RefDocNo, NavigatorType.Current, trans)
                Dim Index As Integer = 1
                Dim PurchaseControlAmount As Decimal = 0.0
                If ObjSRN IsNot Nothing And clsCommon.myLen(ObjSRN.SRN_No) > 0 Then
                    objCostAdj = New ClsAdjustments()
                    objCostAdj.Arr = New List(Of ClsAdjustmentsDetails)
                    For Each objSRNDetail As clsSRNDetail In ObjSRN.Arr
                        '-------------Cost Adjustment Entry----------------
                        objCostAdj.Reference_Document = ObjSRN.SRN_No
                        objCostAdj.Adjustment_Date = obj.Invoice_Entry_Date
                        objCostAdj.Reference = "Cost Adjustment Against AP Invoice : " + obj.Document_No + " & SRN : " + ObjSRN.SRN_No + "."
                        objCostAdj.Description = "Cost Adjustment Against AP Invoice : " + obj.Document_No + " & SRN : " + ObjSRN.SRN_No + "."

                        objCostAdj.Unit_Code = "ALL"
                        objCostAdj.ItemType = IIf(clsCommon.CompairString(ObjSRN.Item_Type, "F") = CompairStringResult.Equal, "FT", "OT")
                        objCostAdj.Loc_Code = obj.loc_code
                        objCostAdj.Loc_Desc = clsLocation.GetName(obj.loc_code, trans)
                        objCostAdj.Trans_Type = "In"
                        If clsCommon.CompairString(ObjSRN.Item_Type, "F") = CompairStringResult.Equal Then
                            objCostAdj.Stock_Type = ""
                        End If

                        Dim objTr As New ClsAdjustmentsDetails()
                        objTr.Adjustment_Line_No = Index
                        objTr.Item_Code = objSRNDetail.Item_Code
                        objTr.Item_Description = objSRNDetail.Item_Desc
                        objTr.Adjustment_Type = "CI"
                        objTr.Item_Quantity = 0
                        objTr.Item_Cost = PurchaseControlAmount
                        objTr.Unit_Code = objSRNDetail.Unit_code
                        objTr.Remarks = ""
                        objTr.Comments = ""
                        objTr.mrp = clsCommon.myCdbl(objSRNDetail.MRP)

                        objTr.BreakageType = ""
                        objTr.Breakage = 0.0
                        objTr.Breakage_Cost = 0.0
                        objTr.LeakageQty = 0.0
                        objTr.ItemType = ObjSRN.Item_Type
                        objTr.Basic_Price = clsItemBasicPrice.GetBasicPrice(objTr.Item_Code, objTr.mrp, trans)
                        'objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))
                        objCostAdj.Arr.Add(objTr)
                    Next
                    '---------------------------------------------------------
                End If
            End If
        Else

        End If

        'qry = "Update TSPL_VENDOR_INVOICE_HEAD set Posting_Date='" + clsCommon.GetPrintDate(clsCommon.myCDate(strPostDate), "dd/MMM/yyyy") + "',Modify_By='" + objCommonVar.CurrentUserCode + "' where Document_No='" + strDocNo + "'"
        'clsDBFuncationality.ExecuteNonQuery(qry, trans)
        If clsCommon.CompairString(obj.Is_ProRated, "Y") = CompairStringResult.Equal Then
            objCostAdj.SaveData(objCostAdj, True, "", trans)
            If clsCommon.CompairString(obj.RefDocType, "BS") = CompairStringResult.Equal Then
                ClsAdjustments.PostData(objCostAdj.Adjustment_No, "Store Adjustment", trans)
            End If
        End If
        ''Cost adjustment of provision entry
        If obj.is_For_Provision = 1 Then
            Dim TempAmt As Double = obj.Prov_Amt - obj.Discount_Base
            If TempAmt <> 0 Then
                qry = "select Doc_No,Vendor_Type from TSPL_PROVISION_ENTRY  where Status_Update_Doc_No='" + obj.Document_No + "' "
                Dim dtPE As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dtPE IsNot Nothing AndAlso dtPE.Rows.Count > 0 Then
                    Dim isApplyPurchaseAccounting As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowPurchaseAccounting, clsFixedParameterCode.AllowPurchaseAccounting, trans)) = 0, False, True)
                    If Not isApplyPurchaseAccounting Then
                        ''added by richa agarwal 13 March,2019 allow journal entry for provision entry of Dairy Sale Type against BHA/22/02/19-000819
                        If clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryMCCLeaseVendor) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryOthers) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryPrimaryTransporter) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntrySecondaryTransporter) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForBulkSale) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForFreshSale) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForProductSale) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForTransfer) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForCSATransfer) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForGateentry) = CompairStringResult.Equal Then
                        ElseIf clsCommon.CompairString(clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), clsFixedParameterCode.CreateJEForProvisionEntryTransporterForDairySale) = CompairStringResult.Equal Then
                        Else
                            Throw New Exception("No Setting found for vendor type " + clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")))
                        End If
                        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CreateJEForProvisionEntry, clsCommon.myCstr(dtPE.Rows(0)("Vendor_Type")), trans)) = 1 Then
                            objCostAdj = New ClsAdjustments()
                            objCostAdj.Arr = New List(Of ClsAdjustmentsDetails)
                            Dim strProvNo As String = ""
                            For Each dr As DataRow In dtPE.Rows
                                If clsCommon.myLen(strProvNo) > 0 Then
                                    strProvNo += ","
                                End If
                                strProvNo += clsCommon.myCstr(dr("Doc_No"))
                            Next
                            '-------------Cost Adjustment Entry----------------
                            objCostAdj.Reference_Document = obj.Document_No
                            objCostAdj.Adjustment_Date = obj.Invoice_Entry_Date
                            objCostAdj.Reference = strProvNo
                            objCostAdj.Description = "Cost Adjustment Against AP Invoice : " + obj.Document_No + " & Provision Entry No : " + strProvNo + "."

                            objCostAdj.Unit_Code = "ALL"
                            Dim strICode As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.MCCDefaultMilkItem, clsFixedParameterCode.MilkSetting, trans))
                            If clsCommon.myLen(strICode) <= 0 Then
                                Throw New Exception("MCC default item not found ")
                            End If
                            Dim strUOM As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Unit_Code from TSPL_ITEM_MASTER where Item_Code='" + strICode + "'", trans))

                            Dim strItemType As String = clsItemMaster.GetItemType(strICode, trans)
                            objCostAdj.IsMilkType = 1
                            objCostAdj.ItemType = IIf(clsCommon.CompairString(strItemType, "F") = CompairStringResult.Equal, "FT", "OT")
                            objCostAdj.Loc_Code = obj.loc_code
                            objCostAdj.Loc_Desc = clsLocation.GetName(obj.loc_code, trans)
                            objCostAdj.MainLocationCode = obj.loc_code
                            objCostAdj.MainLocationDesc = objCostAdj.Loc_Desc

                            If clsCommon.CompairString(strICode, "F") = CompairStringResult.Equal Then
                                objCostAdj.Stock_Type = ""
                            End If

                            Dim objTr As New ClsAdjustmentsDetails()
                            objTr.Adjustment_Line_No = 1
                            objTr.Item_Code = strICode
                            objTr.Item_Description = clsItemMaster.GetItemName(strICode, trans)
                            objTr.Item_Quantity = 0


                            If TempAmt > 0 Then
                                objTr.Item_Cost = TempAmt
                                objTr.Adjustment_Type = "CD"
                                objCostAdj.Trans_Type = "Out"
                            Else
                                objTr.Item_Cost = -1 * TempAmt
                                objTr.Adjustment_Type = "CI"
                                objCostAdj.Trans_Type = "In"
                            End If


                            objTr.Unit_Code = strUOM
                            objTr.Remarks = ""
                            objTr.Comments = ""
                            objTr.mrp = 0

                            objTr.BreakageType = ""
                            objTr.Breakage = 0.0
                            objTr.Breakage_Cost = 0.0
                            objTr.LeakageQty = 0.0
                            objTr.ItemType = strItemType
                            objTr.Basic_Price = clsItemBasicPrice.GetBasicPrice(objTr.Item_Code, objTr.mrp, trans)
                            'objTr.arrBatchItem = TryCast(grow.Cells(colICode).Tag, List(Of clsBatchInventory))
                            objCostAdj.Arr.Add(objTr)

                            objCostAdj.SaveData(objCostAdj, True, "", trans)
                            ClsAdjustments.PostData(objCostAdj.Adjustment_No, "Store Adjustment", trans)
                        End If
                    End If
                End If
            End If
        End If
        ''End of Cost adjustment of provision entry
        clsVedorInvoiceHead.CreateAdditionalCostEntry(strRefDocNo, obj, trans)
        ' done by Panch Raj on 06/07/2016 only for Vendor Service charge against asset for assembled asset
        If clsCommon.CompairString(obj.Invoice_Type, "VS") = CompairStringResult.Equal Then
            For Each objtr As clsVedorInvoiceDetail In obj.Arr
                If clsCommon.myLen(objtr.Asset_Code) > 0 Then
                    qry = " UPDATE TSPL_ACQUISITION_DETAIL SET Book_Source_value=Book_Source_value+" & IIf(obj.Document_Type = "I" Or obj.Document_Type = "C", objtr.Amount_less_Discount, -objtr.Amount_less_Discount) & ",Book_Source_Original_value=Book_Source_Original_value+" & IIf(obj.Document_Type = "I" Or obj.Document_Type = "C", objtr.Amount_less_Discount, -objtr.Amount_less_Discount) & ",Book_Salvage_Value = (Book_Source_value+" & IIf(obj.Document_Type = "I" Or obj.Document_Type = "C", objtr.Amount_less_Discount, -objtr.Amount_less_Discount) & ") * Book_Salvage_Rate / 100  where Asset_Code='" & objtr.Asset_Code & "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If
            Next
        End If

        CreatePrimaryTransporterDebitNoteOfSecurityDeduction(obj, strPostDate, trans)
        clsCommonFunctionality.SaveHistoryData(EnumSaveType.History, objCommonVar.CurrentUserCode, strDocNo, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", "TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL", "AP_Invoice_No", "TSPL_AP_Invoice_Asset_EMI_Details", "AP_Invoice_No", "", "", "", "", "", "", trans)
        clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_REMITTANCE", "Document_No", "TSPL_AP_INVOICE_ADVANCE_INTEREST", "AP_Invoice_No", "TSPL_PROVISION_ENTRY_KNOCKOFF", "AP_Invoice_No", trans)

        '' For EInvoice
        Dim ECustomerType As String = ""
        If obj.IsEInvoice Then
            ECustomerType = clsERPFuncationality.GetVendorEInvoiceType(obj.Vendor_Code, trans)
            If clsCommon.CompairString(ECustomerType, "BC") = CompairStringResult.Equal Then
                Throw New Exception("e-Invoice will be not generated for unregistered vendor")
            End If
            ''richa agarwal 31 Dec,2020 check eInvoice Implementation
            If clsCommon.myLen(clsCommon.myCstr(obj.Tax_Group)) > 0 Then
                Dim isTaxTaxable As String = "N"
                Dim LocationCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Location_Code from TSPL_Location_Master where Loc_Segment_Code='" + obj.loc_code + "' and IsEinvoice=1", trans))
                isTaxTaxable = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select 'Y' from TSPL_TAX_GROUP_MASTER where Tax_Group_Code ='" & obj.Tax_Group & "' and Is_Tax_Exempted =0 and Tax_Group_Type ='P'", trans))
                ''If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(isTaxTaxable), "Y") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.AgainstServiceInvoice), "Y") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "I") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Document_Date, trans) = True Then
                If clsCommon.CompairString(ECustomerType, "BB") = CompairStringResult.Equal AndAlso clsCommon.CompairString(clsCommon.myCstr(isTaxTaxable), "Y") = CompairStringResult.Equal AndAlso clsERPFuncationality.GetEInvoiceStatus(obj.Invoice_Entry_Date, trans) = True Then
                    If clsVedorInvoiceHead.EInvoice_Implementation(obj.Document_No, LocationCode, trans) = True Then
                    Else
                        Throw New Exception("Invalid JSON Value")
                    End If
                End If
            End If
            ''------------------------------
        End If

        qry = "Update TSPL_VENDOR_INVOICE_HEAD set Posting_Date='" + clsCommon.GetPrintDate(clsCommon.myCDate(strPostDate), "dd/MMM/yyyy") + "',Modify_By='" + objCommonVar.CurrentUserCode + "' where Document_No='" + strDocNo + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
        Return True

    End Function
    Public Shared Function EInvoice_Implementation(ByVal strDocNo As String, ByVal strLocation As String, ByVal trans As SqlTransaction) As Boolean
        Try
            Dim isSaved As Boolean = True
            If (clsCommon.myLen(strDocNo) <= 0) Then
                Throw New Exception("Code not found to Post")
            End If

            Dim strtoken As String = ClsEInvoiceOFAPIs.IsGenerateAuthTokenNo_Required(objCommonVar.CurrentCompanyCode, strLocation, trans)
            If clsCommon.myLen(strtoken) > 0 Then
                Dim strQry As String = "select TSPL_VENDOR_master.Vendor_Code, TSPL_VENDOR_Invoice_Head.document_No as DocNo, convert( date, TSPL_VENDOR_Invoice_Head.Invoice_Entry_Date, 103 ) as DocDate, case when TSPL_VENDOR_Invoice_Head.Document_Type = 'D' then 'DBN' when TSPL_VENDOR_Invoice_Head.Document_Type = 'C' then 'CRN' else 'INV' end as DocType, 'B2B' as SupTyp, 'N' as IgstOnIntra, TSPL_LOCATION_MASTER.GSTNO as SellerGSTINNo, TSPL_LOCATION_MASTER.location_desc as SellerLglNm, TSPL_COMPANY_MASTER.Comp_Name as SellerTrdNm, TSPL_LOCATION_MASTER.Add1 as SellerAdd1, TSPL_LOCATION_MASTER.Add2 as SellerAdd2, TSPL_LOCATION_MASTER.City_Code as SellerLoc, TSPL_LOCATION_MASTER.Pin_Code as SellerPincode, Location_State_Master.GST_STATE_Code as SellerStcd, TSPL_LOCATION_MASTER.Phone1 as SellerPhone, TSPL_LOCATION_MASTER.Email as SellerEmail, TSPL_VENDOR_master.GSTFinalNo as BuyerGSTINNo, TSPL_VENDOR_master.Vendor_Name as BuyerLglNm, TSPL_VENDOR_master.Alies_name as BuyerTrdNm, Customer_State_Master.GST_STATE_Code as BuyerPOS, TSPL_VENDOR_master.Add1 as BuyerAdd1, TSPL_VENDOR_master.Add2 as BuyerAdd2, tspl_city_master.City_Name as BuyerLoc, cast( TSPL_VENDOR_master.Pin_Code as int ) as BuyerPincode, Customer_State_Master.GST_STATE_Code as BuyerStcd, TSPL_VENDOR_master.Phone1 as BuyerPhone, TSPL_VENDOR_master.Email as BuyerEmail, TSPL_VENDOR_Invoice_Detail.Detail_Line_No as ItemSlNo, 'Y' as ItemIsServc, TSPL_ADDITIONAL_charges.Description AS ItemPrdDesc, TSPL_ADDITIONAL_charges.Sac_code AS ItemHsnCd, 0 as ItemQty, 'OTH' as ItemUnit, 0 as ItemUnitPrice, TSPL_VENDOR_Invoice_Detail.Amount as ItemTotAmt, TSPL_VENDOR_Invoice_Detail.Discount as ItemDiscount, TSPL_VENDOR_Invoice_Detail.Amount_Less_Discount as ItemAssAmt, case when ISNULL( TSPL_VENDOR_Invoice_Detail.tax1, '' ) = 'IGST' THEN TSPL_VENDOR_Invoice_Detail.TAX1_Rate when ISNULL( TSPL_VENDOR_Invoice_Detail.tax1, '' ) = 'CGST' AND ISNULL( TSPL_VENDOR_Invoice_Detail.tax2, '' ) = 'SGST' THEN TSPL_VENDOR_Invoice_Detail.TAX1_Rate + TSPL_VENDOR_Invoice_Detail.TAX2_Rate ELSE 0 end as ItemGstRt, case when TSPL_VENDOR_Invoice_Detail.TAX1 = 'SGST' AND TSPL_VENDOR_Invoice_Detail.TAX2 = 'CGST' then TSPL_VENDOR_Invoice_Detail.TAX1_Amt when TSPL_VENDOR_Invoice_Detail.TAX1 = 'CGST' AND TSPL_VENDOR_Invoice_Detail.TAX2 = 'SGST' then TSPL_VENDOR_Invoice_Detail.TAX2_Amt else 0 end ItemSgstAmt, case when TSPL_VENDOR_Invoice_Detail.TAX1 = 'IGST' then TSPL_VENDOR_Invoice_Detail.TAX1_Amt else 0 end ItemIgstAmt, case when TSPL_VENDOR_Invoice_Detail.TAX1 = 'CGST' AND TSPL_VENDOR_Invoice_Detail.TAX2 = 'SGST' then TSPL_VENDOR_Invoice_Detail.TAX1_Amt when TSPL_VENDOR_Invoice_Detail.TAX1 = 'SGST' AND TSPL_VENDOR_Invoice_Detail.TAX2 = 'CGST' then TSPL_VENDOR_Invoice_Detail.TAX2_Amt else 0 end ItemCgstAmt, 0 as ItemOthChrg, TSPL_VENDOR_Invoice_Detail.total_amount - case when isnull(TCS1.is_tcs, '')= 'Y' THEN TSPL_VENDOR_Invoice_Detail.TAX2_AMT when isnull(TCS2.is_tcs, '')= 'Y' THEN TSPL_VENDOR_Invoice_Detail.TAX3_AMT ELSE 0 END as ItemTotItemVal, TSPL_VENDOR_Invoice_Head.Discount_Base as ValDtlsAssVal, case when TSPL_VENDOR_Invoice_Head.TAX1 = 'CGST' AND TSPL_VENDOR_Invoice_Head.TAX2 = 'SGST' then TSPL_VENDOR_Invoice_Head.TAX1_Amt when TSPL_VENDOR_Invoice_Head.TAX1 = 'SGST' AND TSPL_VENDOR_Invoice_Head.TAX2 = 'CGST' then TSPL_VENDOR_Invoice_Head.TAX2_Amt else 0 end ValDtlsCgstVal, case when TSPL_VENDOR_Invoice_Head.TAX1 = 'SGST' AND TSPL_VENDOR_Invoice_Head.TAX2 = 'CGST' then TSPL_VENDOR_Invoice_Head.TAX1_Amt when TSPL_VENDOR_Invoice_Head.TAX1 = 'CGST' AND TSPL_VENDOR_Invoice_Head.TAX2 = 'SGST' then TSPL_VENDOR_Invoice_Head.TAX2_Amt else 0 end ValDtlsSgstVal, case when TSPL_VENDOR_Invoice_Head.TAX1 = 'IGST' then TSPL_VENDOR_Invoice_Head.TAX1_Amt else 0 end ValDtlsIgstVal, TSPL_VENDOR_Invoice_Head.Discount_Amount as ValDtlsDiscount, case when isnull(TCS1.is_tcs, '')= 'Y' THEN TSPL_VENDOR_Invoice_Head.TAX2_AMT when isnull(TCS2.is_tcs, '')= 'Y' THEN TSPL_VENDOR_Invoice_Head.TAX3_AMT ELSE 0 END as ValDtlsOthChrg, TSPL_VENDOR_Invoice_Head.document_total as ValDtlsTotInvVal, TSPL_VENDOR_Invoice_Head.RoundOffAmount as ValDtlsRndOffAmt 
                from TSPL_VENDOR_Invoice_Head 
                Left Outer Join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code = '" & objCommonVar.CurrentCompanyCode & "' 
                Left Outer Join TSPL_VENDOR_master on TSPL_VENDOR_master.Vendor_Code = TSPL_VENDOR_Invoice_Head.Vendor_Code
                left Outer Join TSPL_LOCATION_MASTER on TSPL_LOCATION_MASTER.Loc_Segment_Code = TSPL_VENDOR_Invoice_Head.Loc_code and TSPL_LOCATION_MASTER.IsEinvoice=1
                left outer join TSPL_VENDOR_Invoice_Detail on TSPL_VENDOR_Invoice_Detail.document_No = TSPL_VENDOR_Invoice_Head.document_No 
                left outer join TSPL_ADDITIONAL_charges on TSPL_ADDITIONAL_charges.Code = TSPL_VENDOR_Invoice_Detail.AddChargeCode 
                left outer join TSPL_STATE_MASTER as LOCATION_State_Master on LOCATION_State_Master.STATE_CODE = TSPL_LOCATION_MASTER.State 
                left outer join TSPL_STATE_MASTER as Customer_State_Master on Customer_State_Master.STATE_CODE = TSPL_VENDOR_master.State_Code 
                left outer join tspl_city_master on tspl_city_master.city_code = TSPL_VENDOR_master.City_Code 
                left outer join tspl_tax_master as TCS1 on TCS1.Tax_Code = TSPL_VENDOR_Invoice_Head.Tax2 left outer join tspl_tax_master as TCS2 on TCS2.Tax_Code = TSPL_VENDOR_Invoice_Head.Tax3 where TSPL_VENDOR_Invoice_Head.document_No = '" & strDocNo & "'"

                '               

                Dim objResult As Object = ClsEInvoiceOFAPIs.PostAuthTokenNo_withInvoiceData(objCommonVar.CurrentCompanyCode, strtoken, strQry, strLocation, trans, True)
                If objResult IsNot Nothing Then
                    'assign to variable
                    Dim AckNo As String = objResult.SelectToken("AckNo").ToString
                    Dim AckDt As String = objResult.SelectToken("AckDt").ToString
                    Dim Irn As String = objResult.SelectToken("Irn").ToString
                    Dim SignedQRCode As String = objResult.SelectToken("SignedQRCode").ToString
                    clsDBFuncationality.ExecuteNonQuery("update TSPL_VENDOR_Invoice_Head set  IRN_No ='" & Irn & "',qr_code='" & SignedQRCode & "',ack_no='" & AckNo & "',ack_date='" & clsCommon.GetPrintDate(AckDt, "dd/MMM/yyyy hh:mm tt") & "' where document_No ='" & strDocNo & "'", trans)

                    Dim TempByte As Byte() = clsERPFuncationalityOLD.GenerateMyQCCode(SignedQRCode)
                    clsDBFuncationality.UpdateImage("BarCode_Img", TempByte, "TSPL_VENDOR_Invoice_Head", "TSPL_VENDOR_Invoice_Head.document_No='" & strDocNo & "'", trans)
                Else
                    Return False
                End If
            Else
                Return False
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Public Shared Function GetVendorRetention(ByVal rdbDetail As Boolean, ByVal fromDate As DateTime, ByVal ToDate As DateTime, ByVal cbgVendor As ArrayList, ByVal txtVendorGroupMult As ArrayList, ByVal txtLocationMult As ArrayList) As String
        Dim qry As String = ""
        Dim strWhrClause As String = ""
        If cbgVendor IsNot Nothing Then
            If cbgVendor.Count > 0 Then
                strWhrClause += "and xxx.Vendor_Code  IN (" + clsCommon.GetMulcallString(cbgVendor) + ") "
            End If
        End If
        If txtLocationMult IsNot Nothing AndAlso txtLocationMult.Count > 0 Then
            strWhrClause += " and xxx.Loc_Code in (" + clsCommon.GetMulcallString(txtLocationMult) + ") " + Environment.NewLine
        End If
        If txtVendorGroupMult IsNot Nothing AndAlso txtVendorGroupMult.Count > 0 Then
            strWhrClause += " and Vendor_Group_Code in (" + clsCommon.GetMulcallString(txtVendorGroupMult) + ") " + Environment.NewLine
        End If

        Dim squery As String
        Dim query As String
        Dim companyname As String = objCommonVar.CurrentCompanyName
        query = " select TSPL_VENDOR_INVOICE_HEAD.Document_No,TSPL_VENDOR_INVOICE_HEAD.Description as Comments, 
                  TSPL_VENDOR_INVOICE_HEAD.Posting_Date as DocDate, 'AP Invoice' as Document_Type ,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code,
                  Loc_Code ,TSPL_GL_SEGMENT_CODE.Description,TSPL_DEDUCTION_MASTER.Description as Is_Retention,
                  case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then Amount Else 0 End as DrAmt,
                  case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then Amount Else 0 End as CrAmt  from TSPL_VENDOR_INVOICE_HEAD   left join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No =TSPL_VENDOR_INVOICE_HEAD.Document_No "
        query += " left join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER .Code =TSPL_VENDOR_INVOICE_DETAIL.DeductionCode 
                   left join TSPL_GL_SEGMENT_CODE on TSPL_VENDOR_INVOICE_HEAD.Loc_Code=TSPL_GL_SEGMENT_CODE.Segment_code 
                   where TSPL_VENDOR_INVOICE_HEAD.Document_Type in ('D','C') AND Is_Retention =1"
        query += " union all"
        query += " select TSPL_PAYMENT_HEADER.Payment_No,TSPL_PAYMENT_HEADER.Entry_Desc as Comments ,
                   TSPL_PAYMENT_HEADER.Payment_Date  ,'Payment Entry' as Payment_Type ,TSPL_PAYMENT_HEADER.Vendor_Code,
                   TSPL_PAYMENT_HEADER.Location_GL_Code as Loc_Code, TSPL_GL_SEGMENT_CODE.Description,'Retention' as Is_Retention,
                   (case when TSPL_PAYMENT_HEADER.Payment_Type='SR' then Payment_Amount else 0 end) as DrAmt,
                   (case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then Payment_Amount else 0 end)  as CrAmt   from TSPL_PAYMENT_HEADER 
                   left join TSPL_GL_SEGMENT_CODE on TSPL_PAYMENT_HEADER.Location_GL_Code =TSPL_GL_SEGMENT_CODE.Segment_code 
                   where TSPL_PAYMENT_HEADER.Payment_Type IN ('RC','SR') and Is_Retention =1"
        query += " union all"
        query += " select TSPL_BANK_REVERSE.Reverse_Code,TSPL_BANK_REVERSE.Reason as Comments ,TSPL_BANK_REVERSE.Reversal_Date,
                   'Bank Reverse' as Payment_Type,TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_PAYMENT_HEADER.Location_GL_Code as Loc_Code,
                   TSPL_GL_SEGMENT_CODE.Description, 'Retention' as Is_Retention, TSPL_PAYMENT_HEADER.Payment_Amount as DrAmt,
                   0 as CrAmt  from TSPL_BANK_REVERSE"
        query += " left join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No =TSPL_BANK_REVERSE.Document_No
                   left join TSPL_GL_SEGMENT_CODE on TSPL_PAYMENT_HEADER.Location_GL_Code =TSPL_GL_SEGMENT_CODE.Segment_code"
        query += " where Reverse_Document ='Payments' and TSPL_PAYMENT_HEADER.Is_Retention =1"


        squery = "select '" & companyname & "' as companyname,'" & clsCommon.GetPrintDate(fromDate, "dd/MM/yyyy") & "'  as fromDate,
                 '" & clsCommon.GetPrintDate(ToDate, "dd/MM/yyyy") & "'  as Todate ,ROW_NUMBER() OVER ( ORDER BY final.Vendor_Code) as RowNo,
                 convert(varchar,DocDate,103) as DocDate, Document_No,Comments, Document_Type, Vendor_Code, Vendor_Name,Loc_Code,
                 Description,Vendor_Group_Code,Group_Desc,Is_Retention, Opening ,DrAmt,CrAmt,SUM(TempClosing) OVER (Partition BY final.Vendor_Code ORDER BY RowNo) as Closing   from (Select *,Opening-DrAmt+CrAmt as TempClosing, ROW_NUMBER() OVER (Partition BY Vendor_Code ORDER BY Vendor_Code, DocDate) as RowNo from ("
        squery += " Select XXX.Vendor_Code, MAX(TSPL_VENDOR_MASTER.Vendor_Name) as Vendor_Name,Loc_Code ,max(xxx.Description ) as Description,
                    (TSPL_VENDOR_MASTER.Vendor_Group_Code) as Vendor_Group_Code ,max(TSPL_VENDOR_GROUP.Group_Desc) as Group_Desc,
                    '' as Document_No,'' as Comments, NULL as DocDate, 'Opening' as Document_Type, SUM(CrAmt)-SUM(DrAmt) as Opening,
                    0 as DrAmt, 0 as CrAmt,(Is_Retention ) as Is_Retention  from ( "
        squery += " " + query + " "
        squery += " )XXX LEFT OUTER JOIN TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=XXX.Vendor_Code
                    left join TSPL_VENDOR_GROUP on TSPL_VENDOR_GROUP.Ven_Group_Code =TSPL_VENDOR_MASTER.Vendor_Group_Code
                    WHERE 2=2 and CONVERT(Date,DocDate,103)< CONVERT(Date,'" + fromDate + "',103) "
        squery += " " + strWhrClause + " "
        squery += "  Group By XXX.Vendor_Code,xxx.Loc_Code,TSPL_VENDOR_MASTER.Vendor_Group_Code,Is_Retention"
        squery += " union all "
        squery += " Select XXX.Vendor_Code, TSPL_VENDOR_MASTER.Vendor_Name,xxx.Loc_Code,xxx.Description,Vendor_Group_Code,
                    Group_Desc, Document_No,Comments, DocDate, Document_Type, 0 as Opening, DrAmt, CrAmt,Is_Retention from ("
        squery += " " + query + " "
        squery += " ) XXX LEFT OUTER JOIN TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=XXX.Vendor_Code 
                    left join TSPL_VENDOR_GROUP on TSPL_VENDOR_GROUP.Ven_Group_Code =TSPL_VENDOR_MASTER.Vendor_Group_Code
                    WHERE 2=2 and CONVERT(Date,DocDate,103)>=CONVERT(Date,'" + fromDate + "',103)
                    AND CONVERT(Date,DocDate,103)<=CONVERT(Date,'" + ToDate + "',103)"
        squery += " " + strWhrClause + " "
        squery += ") YYY )final"
        If rdbDetail Then
            squery += " ORDER BY Vendor_Code "
        End If

        If rdbDetail Then
            qry = "" + squery + ""
        Else
            qry = "select Vendor_Code ,max(Vendor_Name ) as Vendor_Name,max(companyname) as companyname, max(fromDate) as fromDate,
            max(Todate) as Todate,max(Is_Retention) as Is_Retention  ,Loc_Code,max(Description) as Description,Vendor_Group_Code,
            max(Group_Desc) as Group_Desc,sum(Opening ) as Opening, sum(DrAmt) as DrAmt,sum(CrAmt) as CrAmt,
            SUM(Opening)-SUM(DrAmt)+SUM(CrAmt) as Closing from ("
            qry += "" & squery & ""

            qry += " ) summary group by summary.Vendor_Code,summary .Loc_Code,summary .Vendor_Group_Code "
        End If
        Return qry
    End Function

    Public Shared Function GetVendorSecurity(ByVal rdbDetail As Boolean, ByVal fromDate As DateTime, ByVal ToDate As DateTime, ByVal cbgVendor As ArrayList, ByVal txtVendorGroupMult As ArrayList, ByVal txtLocationMult As ArrayList) As String
        Dim qry As String = ""
        Dim strWhrClause As String = ""
        If cbgVendor IsNot Nothing Then
            If cbgVendor.Count > 0 Then
                strWhrClause += "and xxx.Vendor_Code  IN (" + clsCommon.GetMulcallString(cbgVendor) + ") "
            End If
        End If
        If txtLocationMult IsNot Nothing AndAlso txtLocationMult.Count > 0 Then
            strWhrClause += " and xxx.Loc_Code in (" + clsCommon.GetMulcallString(txtLocationMult) + ") " + Environment.NewLine
        End If
        If txtVendorGroupMult IsNot Nothing AndAlso txtVendorGroupMult.Count > 0 Then
            strWhrClause += " and Vendor_Group_Code in (" + clsCommon.GetMulcallString(txtVendorGroupMult) + ") " + Environment.NewLine
        End If

        Dim squery As String
        Dim query As String
        Dim companyname As String = objCommonVar.CurrentCompanyName
        query = " select TSPL_VENDOR_INVOICE_HEAD.Document_No,TSPL_VENDOR_INVOICE_HEAD.Description as Comments, TSPL_VENDOR_INVOICE_HEAD.Posting_Date as DocDate, 'AP Invoice' as Document_Type ,TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ,Loc_Code ,TSPL_GL_SEGMENT_CODE.Description,TSPL_DEDUCTION_MASTER.Description as Is_Security, case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then TSPL_VENDOR_INVOICE_DETAIL.Amount Else 0 End as DrAmt, case when TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then TSPL_VENDOR_INVOICE_DETAIL.Amount Else 0 End as CrAmt, ISNULL(TSPL_VENDOR_INVOICE_SECURITY_REFUND.Amount,0) as Adjust_Bal   from TSPL_VENDOR_INVOICE_HEAD   left join TSPL_VENDOR_INVOICE_DETAIL on TSPL_VENDOR_INVOICE_DETAIL.Document_No =TSPL_VENDOR_INVOICE_HEAD.Document_No  left outer join TSPL_VENDOR_INVOICE_SECURITY_REFUND on TSPL_VENDOR_INVOICE_SECURITY_REFUND.Against_Security=TSPL_VENDOR_INVOICE_HEAD.Document_No "
        query += " left join TSPL_DEDUCTION_MASTER on TSPL_DEDUCTION_MASTER .Code =TSPL_VENDOR_INVOICE_DETAIL.DeductionCode left join TSPL_GL_SEGMENT_CODE on TSPL_VENDOR_INVOICE_HEAD.Loc_Code=TSPL_GL_SEGMENT_CODE.Segment_code  where TSPL_VENDOR_INVOICE_HEAD.Document_Type in ('D','C') AND Is_Security =1"
        query += " union all"
        query += " select TSPL_PAYMENT_HEADER.Payment_No,TSPL_PAYMENT_HEADER.Entry_Desc as Comments ,TSPL_PAYMENT_HEADER.Payment_Date  ,'Payment Entry' as Payment_Type ,TSPL_PAYMENT_HEADER.Vendor_Code ,TSPL_PAYMENT_HEADER.Location_GL_Code as Loc_Code, TSPL_GL_SEGMENT_CODE.Description,'Security' as Is_Security, (case when TSPL_PAYMENT_HEADER.Payment_Type='SR' then Payment_Amount else 0 end) as DrAmt,  (case when TSPL_PAYMENT_HEADER.Payment_Type='RC' then Payment_Amount else 0 end)  as CrAmt ,0 AS ADJUST_BAL   from TSPL_PAYMENT_HEADER left join TSPL_GL_SEGMENT_CODE on TSPL_PAYMENT_HEADER.Location_GL_Code =TSPL_GL_SEGMENT_CODE.Segment_code where TSPL_PAYMENT_HEADER.Payment_Type IN ('RC','SR') and Is_Security =1"
        query += " union all"
        query += " select TSPL_BANK_REVERSE.Reverse_Code,TSPL_BANK_REVERSE.Reason as Comments ,TSPL_BANK_REVERSE.Reversal_Date, 'Bank Reverse' as Payment_Type,TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_PAYMENT_HEADER.Location_GL_Code as Loc_Code, TSPL_GL_SEGMENT_CODE.Description, 'Security' as Is_Security, TSPL_PAYMENT_HEADER.Payment_Amount as DrAmt, 0 as CrAmt,0 as adjust_bal  from TSPL_BANK_REVERSE"
        query += " left join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No =TSPL_BANK_REVERSE.Document_No left join TSPL_GL_SEGMENT_CODE on TSPL_PAYMENT_HEADER.Location_GL_Code =TSPL_GL_SEGMENT_CODE.Segment_code"
        query += " where Reverse_Document ='Payments' and TSPL_PAYMENT_HEADER.Is_Security =1"


        squery = "select '" & companyname & "' as companyname,'" & clsCommon.GetPrintDate(fromDate, "dd/MM/yyyy") & "'  as fromDate,'" & clsCommon.GetPrintDate(ToDate, "dd/MM/yyyy") & "'  as Todate ,ROW_NUMBER() OVER ( ORDER BY final.Vendor_Code) as RowNo,convert(varchar,DocDate,103) as DocDate, Document_No,Comments, Document_Type, Vendor_Code, Vendor_Name,Loc_Code,Description,Vendor_Group_Code,Group_Desc,Is_Security, Opening ,DrAmt,CrAmt,SUM(TempClosing) OVER (Partition BY final.Vendor_Code ORDER BY RowNo) as Closing,Adjust_Bal    from (Select *,Opening-DrAmt+CrAmt as TempClosing, ROW_NUMBER() OVER (Partition BY Vendor_Code ORDER BY Vendor_Code, DocDate) as RowNo from ("
        squery += " Select XXX.Vendor_Code, MAX(TSPL_VENDOR_MASTER.Vendor_Name) as Vendor_Name,Loc_Code ,max(xxx.Description ) as Description,(TSPL_VENDOR_MASTER.Vendor_Group_Code) as Vendor_Group_Code ,max(TSPL_VENDOR_GROUP.Group_Desc) as Group_Desc, '' as Document_No,'' as Comments, NULL as DocDate, 'Opening' as Document_Type, SUM(CrAmt)-SUM(DrAmt) as Opening, 0 as DrAmt, 0 as CrAmt,SUM(XXX.Adjust_Bal)Adjust_Bal,(Is_Security ) as Is_Security  from (select xxx.Document_No,xxx.Comments,xxx.DocDate,xxx.Document_Type,xxx.Vendor_Code,xxx.Loc_Code,xxx.Description,xxx.Is_Security,xxx.DrAmt,xxx.CrAmt,
                case when xxx.Adjust_Bal>0 then xxx.CrAmt-xxx.Adjust_Bal else CrAmt end as Adjust_Bal
                from ( "
        squery += " " + query + " "
        squery += " )XXX)xxx LEFT OUTER JOIN TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=XXX.Vendor_Code left join TSPL_VENDOR_GROUP on TSPL_VENDOR_GROUP.Ven_Group_Code =TSPL_VENDOR_MASTER.Vendor_Group_Code WHERE 2=2 and CONVERT(Date,DocDate,103)< CONVERT(Date,'" + fromDate + "',103) "
        squery += " " + strWhrClause + " "
        squery += "  Group By XXX.Vendor_Code,xxx.Loc_Code,TSPL_VENDOR_MASTER.Vendor_Group_Code,Is_Security"
        squery += " union all "
        squery += " Select XXX.Vendor_Code, TSPL_VENDOR_MASTER.Vendor_Name,xxx.Loc_Code,xxx.Description,Vendor_Group_Code ,Group_Desc, Document_No,Comments, DocDate, Document_Type, 0 as Opening, DrAmt, CrAmt,Adjust_Bal,Is_Security from ("
        squery += " select xxx.Document_No,xxx.Comments,xxx.DocDate,xxx.Document_Type,xxx.Vendor_Code,xxx.Loc_Code,xxx.Description,xxx.Is_Security,xxx.DrAmt,xxx.CrAmt,
                case  when xxx.Adjust_Bal>0 then xxx.CrAmt-xxx.Adjust_Bal else CrAmt end as Adjust_Bal
                from (    " + query + " "
        squery += " ) XXX)XXX LEFT OUTER JOIN TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code=XXX.Vendor_Code left join TSPL_VENDOR_GROUP on TSPL_VENDOR_GROUP.Ven_Group_Code =TSPL_VENDOR_MASTER.Vendor_Group_Code WHERE 2=2 and CONVERT(Date,DocDate,103)>=CONVERT(Date,'" + fromDate + "',103) AND CONVERT(Date,DocDate,103)<=CONVERT(Date,'" + ToDate + "',103)"
        squery += " " + strWhrClause + " "
        squery += ") YYY )final"
        If rdbDetail Then
            squery += " ORDER BY Vendor_Code "
        End If

        If rdbDetail Then
            qry = "" + squery + ""
        Else
            qry = "select Vendor_Code ,max(Vendor_Name ) as Vendor_Name,max(companyname) as companyname, max(fromDate) as fromDate , max(Todate) as Todate,max(Is_Security) as Is_Security  ,Loc_Code,max(Description) as Description,Vendor_Group_Code,max(Group_Desc) as Group_Desc,sum(Opening ) as Opening, sum(DrAmt) as DrAmt,sum(CrAmt) as CrAmt,SUM(Opening)-SUM(DrAmt)+SUM(CrAmt) as Closing from ("
            qry += "" & squery & ""

            qry += " ) summary group by summary.Vendor_Code,summary .Loc_Code,summary .Vendor_Group_Code "
        End If
        Return qry
    End Function


    Public Shared Sub CreatePrimaryTransporterDebitNoteOfSecurityDeduction(ByVal obj As clsVedorInvoiceHead, ByVal strPostDate As Date, ByVal trans As SqlTransaction)
        If clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal Then
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select Security_Amount,Security_Deduction_Amount from TSPL_VENDOR_MASTER where Form_Type='PTM' and Vendor_Code='" + obj.Vendor_Code + "'", trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim dclSecurityAmount As Decimal = clsCommon.myCdbl(dt.Rows(0)("Security_Amount"))
                Dim dblAmount As Decimal = clsCommon.myCdbl(dt.Rows(0)("Security_Deduction_Amount"))
                If dclSecurityAmount > 0 AndAlso dblAmount > 0 Then

                    dt = clsDBFuncationality.GetDataTable("select sum(Document_Total) as Document_Total,max(Posting_Date) as Posting_Date  from tspl_vendor_invoice_head where Document_Type='D' and Vendor_Code='" + obj.Vendor_Code + "' and RefDocType='PTM-SEC' and RefDocNo='" + obj.Vendor_Code + "'", trans)


                    Dim arrVendor As New ArrayList
                    arrVendor.Add(obj.Vendor_Code)
                    Dim qry As String = "select Vendor_Code,sum(Closing)  as Closing from (" + clsVedorInvoiceHead.GetVendorSecurity(False, strPostDate.AddDays(-1), strPostDate, arrVendor, Nothing, Nothing) + ")xxxx group by Vendor_Code"
                    Dim dtBalance As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                    Dim dclAmtDeducted As Decimal = 0
                    If dtBalance IsNot Nothing AndAlso dtBalance.Rows.Count > 0 Then
                        dclAmtDeducted = clsCommon.myCdbl(dtBalance.Rows(0)("Closing"))
                    Else
                        dclAmtDeducted = clsCommon.myCdbl(dt.Rows(0)("Document_Total"))
                    End If
                    Dim flag As Boolean = True
                    If dt.Rows(0)("Posting_Date") IsNot DBNull.Value Then
                        Dim LastDateDeducted As Date = clsCommon.myCDate(dt.Rows(0)("Posting_Date"))
                        If strPostDate < LastDateDeducted Then
                            flag = False
                        ElseIf LastDateDeducted.Year = strPostDate.Year AndAlso LastDateDeducted.Month = strPostDate.Month Then
                            flag = False
                        End If
                    End If
                    If flag Then
                        If ((dclSecurityAmount - dclAmtDeducted) < dblAmount) Then
                            dblAmount = dclSecurityAmount - dclAmtDeducted
                        End If
                        If obj.Document_Total < dblAmount Then
                            dblAmount = obj.Document_Total
                        End If
                        If dblAmount > 0 Then
                            Dim strVendor As String = obj.Vendor_Code

                            Dim objVendorInvHead As New clsVedorInvoiceHead()
                            Dim objVendorInvDetail As New clsVedorInvoiceDetail()
                            'objVendorInvHead.isDeduction = 1
                            objVendorInvHead.Security = 1
                            'Dim dtDed As DataTable = clsDBFuncationality.GetDataTable("select code from TSPL_DEDUCTION_MASTER  where Is_Default_Security_Deduction=1", trans)
                            'If dtDed Is Nothing OrElse dtDed.Rows.Count <= 0 Then
                            '    Throw New Exception("Please set default Security deduction code")
                            'End If
                            'objVendorInvDetail.DeductionCode = clsCommon.myCstr(dtDed.Rows(0)("code"))

                            'objVendorInvHead.Document_No = txtDocNo.Value'ToBeGenerated
                            objVendorInvHead.Invoice_Entry_Date = clsCommon.GetPrintDate(strPostDate, "dd/MMM/yyyy")
                            objVendorInvHead.Vendor_Code = strVendor
                            objVendorInvHead.Vendor_Name = clsVendorMaster.GetName(strVendor, trans)
                            objVendorInvHead.Vendor_Invoice_No = "" ''No Need to send vendor invoice no because it is of debit note type
                            objVendorInvHead.Invoice_Type = "AP"
                            objVendorInvHead.Vendor_Invoice_Date = strPostDate
                            objVendorInvHead.loc_code = obj.loc_code
                            'objVendorInvHead.Irregular_loc_code = obj.Irregular_MCC_CODE
                            objVendorInvHead.Description = "AP Debit Note Against Primary Transporter Deduction"
                            'objVendorInvHead.PROJECT_ID = 1 'obj.PROJECT_ID
                            objVendorInvHead.Account_Set = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select  Vendor_Account from TSPL_VENDOR_MASTER where Vendor_Code ='" + objVendorInvHead.Vendor_Code + "'", trans))
                            If (clsCommon.myLen(objVendorInvHead.Account_Set) < 0) Then
                                Throw New Exception("Please set the vendor Account Set For Vendor : " + objVendorInvHead.Vendor_Name)
                            End If

                            objVendorInvHead.Document_Type = "D" ''For Purchase Invoice Type
                            ''objVendorInvHead.PO_Number = obj.p

                            '' ''added by priti
                            objVendorInvHead.RefDocType = "PTM-SEC"
                            objVendorInvHead.RefDocNo = obj.Vendor_Code
                            'objVendorInvHead.Ref_SNo = ""
                            '' '' priti ends here
                            'objVendorInvHead.Order_No = txtOrderNo.Text
                            ' objVendorInvHead.Total_Tax = 0

                            objVendorInvHead.On_Hold = False
                            'Dim srndate As String = ""
                            'Dim srncode As String = ""
                            'Dim Vlc_Code As String = ""
                            'Dim Vlc_Name As String = ""
                            'For Each objTr As clsMilkPurchaseInvoiceMCCDetail In obj.ObjList
                            '    If clsCommon.myLen(objTr.SRN_CODE) > 0 Then
                            '        Dim query As String = "select doc_date,vd.VLC_Code,VLC_Name from TSPL_Milk_SRN_HEAD sh inner join TSPL_VLC_MASTER_HEAD vd on sh.VLC_CODE=vd.VLC_Code where DOc_Code ='" + objTr.SRN_CODE + "' "
                            '        Dim Dt_SRN As DataTable = clsDBFuncationality.GetDataTable(query, trans)
                            '        srndate = IIf(srndate = "", clsCommon.myCDate(CStr(Dt_SRN.Rows(0).Item("Doc_Date")), "dd/MMM/yyyy"), srndate & "," & clsCommon.myCDate(CStr(Dt_SRN.Rows(0).Item("Doc_Date")), "dd/MMM/yyyy"))
                            '        srncode = IIf(srncode = "", objTr.SRN_CODE, srncode & "," & objTr.SRN_CODE)
                            '        Vlc_Code = IIf(Vlc_Code = "", Dt_SRN.Rows(0).Item("VLC_Code").ToString, Vlc_Code & "," & Dt_SRN.Rows(0).Item("VLC_Code").ToString)
                            '        Vlc_Name = IIf(Vlc_Name = "", Dt_SRN.Rows(0).Item("VLC_Name").ToString, Vlc_Name & "," & Dt_SRN.Rows(0).Item("VLC_name").ToString)
                            '    End If
                            'Next



                            'objVendorInvHead.Description = "VSP : " + obj.VSP_CODE + "/" + vendor_name + "VLC : " + Vlc_Code + "/" + Vlc_Name + " .Against PI Invoice No " + obj.DOC_CODE + "-" + srncode + "-" + srndate
                            'objVendorInvHead.Tax_Calculation_Type = Nothing
                            'objVendorInvHead.Tax_Group = Nothing
                            'If (clsCommon.myLen(obj.TAX1) > 0) Then
                            '    objVendorInvHead.TAX1 = obj.TAX1
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX1, trans) Then
                            '        objVendorInvHead.TAX1_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX1, trans)
                            '        objVendorInvHead.TAX1_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX1_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX1_Rate = obj.TAX1_Rate
                            '    objVendorInvHead.Tax1_BAmount = obj.TAX1_Base_Amt
                            '    objVendorInvHead.TAX1_Amt = obj.TAX1_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX2) > 0) Then
                            '    objVendorInvHead.TAX2 = obj.TAX2
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX2, trans) Then
                            '        objVendorInvHead.TAX2_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX2, trans)
                            '        objVendorInvHead.TAX2_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX2_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX2_Rate = obj.TAX2_Rate
                            '    objVendorInvHead.Tax2_BAmount = obj.TAX2_Base_Amt
                            '    objVendorInvHead.TAX2_Amt = obj.TAX2_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX3) > 0) Then
                            '    objVendorInvHead.TAX3 = obj.TAX3
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX3, trans) Then
                            '        objVendorInvHead.TAX3_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX3, trans)
                            '        objVendorInvHead.TAX3_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX3_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX3_Rate = obj.TAX3_Rate
                            '    objVendorInvHead.Tax3_BAmount = obj.TAX3_Base_Amt
                            '    objVendorInvHead.TAX3_Amt = obj.TAX3_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX4) > 0) Then
                            '    objVendorInvHead.TAX4 = obj.TAX4
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX4, trans) Then
                            '        objVendorInvHead.TAX4_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX4, trans)
                            '        objVendorInvHead.TAX4_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX4_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX4_Rate = obj.TAX4_Rate
                            '    objVendorInvHead.Tax4_BAmount = obj.TAX4_Base_Amt
                            '    objVendorInvHead.TAX4_Amt = obj.TAX4_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX5) > 0) Then
                            '    objVendorInvHead.TAX5 = obj.TAX5
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX5, trans) Then
                            '        objVendorInvHead.TAX5_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX5, trans)
                            '        objVendorInvHead.TAX5_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX5_GLAC, obj.MCC_CODE, trans)

                            '    End If
                            '    objVendorInvHead.TAX5_Rate = obj.TAX5_Rate
                            '    objVendorInvHead.Tax5_BAmount = obj.TAX5_Base_Amt
                            '    objVendorInvHead.TAX5_Amt = obj.TAX5_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX6) > 0) Then
                            '    objVendorInvHead.TAX6 = obj.TAX6
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX6, trans) Then
                            '        objVendorInvHead.TAX6_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX6, trans)
                            '        objVendorInvHead.TAX6_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX6_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX6_Rate = obj.TAX6_Rate
                            '    objVendorInvHead.Tax6_BAmount = obj.TAX6_Base_Amt
                            '    objVendorInvHead.TAX6_Amt = obj.TAX6_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX7) > 0) Then
                            '    objVendorInvHead.TAX7 = obj.TAX7
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX7, trans) Then
                            '        objVendorInvHead.TAX7_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX7, trans)
                            '        objVendorInvHead.TAX7_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX7_GLAC, obj.MCC_CODE, trans)

                            '    End If
                            '    objVendorInvHead.TAX7_Rate = obj.TAX7_Rate
                            '    objVendorInvHead.Tax7_BAmount = obj.TAX7_Base_Amt
                            '    objVendorInvHead.TAX7_Amt = obj.TAX7_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX8) > 0) Then
                            '    objVendorInvHead.TAX8 = obj.TAX8
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX8, trans) Then
                            '        objVendorInvHead.TAX8_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX8, trans)
                            '        objVendorInvHead.TAX8_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX8_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX8_Rate = obj.TAX8_Rate
                            '    objVendorInvHead.Tax8_BAmount = obj.TAX8_Base_Amt
                            '    objVendorInvHead.TAX8_Amt = obj.TAX8_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX9) > 0) Then
                            '    objVendorInvHead.TAX9 = obj.TAX9
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX9, trans) Then
                            '        objVendorInvHead.TAX9_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX9, trans)
                            '        objVendorInvHead.TAX9_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX9_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX9_Rate = obj.TAX9_Rate
                            '    objVendorInvHead.Tax9_BAmount = obj.TAX9_Base_Amt
                            '    objVendorInvHead.TAX9_Amt = obj.TAX9_Amt
                            'End If
                            'If (clsCommon.myLen(obj.TAX10) > 0) Then
                            '    objVendorInvHead.TAX10 = obj.TAX10
                            '    If clsTaxMaster.IsTaxRecoverableAC(obj.TAX10, trans) Then
                            '        objVendorInvHead.TAX10_GLAC = clsTaxMaster.GetTaxRecoverableAC(obj.TAX10, trans)
                            '        objVendorInvHead.TAX10_GLAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.TAX10_GLAC, obj.MCC_CODE, trans)
                            '    End If
                            '    objVendorInvHead.TAX10_Rate = obj.TAX10_Rate
                            '    objVendorInvHead.Tax10_BAmount = obj.TAX10_Base_Amt
                            '    objVendorInvHead.TAX10_Amt = obj.TAX10_Amt
                            'End If

                            'objVendorInvHead.Terms_Code = obj.Terms_Code
                            'objVendorInvHead.Terms_Description = obj.TermsName
                            objVendorInvHead.Due_Date = strPostDate

                            'objVendorInvHead.Against_POInvoice_No = obj.DOC_CODE
                            'objVendorInvHead.Against_MillkPurchaseInvoice_No = obj.DOC_CODE

                            dt = clsDBFuncationality.GetDataTable("select Acct_Set_Code,Payable_Account,Discount_Account,Deduction_ACCOUNT,SECURITY_ACCOUNT from TSPL_VENDOR_ACCOUNT_SET  where Acct_Set_Code='" + objVendorInvHead.Account_Set + "'", trans)
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                objVendorInvHead.Vendor_Control_AC = clsCommon.myCstr(dt.Rows(0)("Payable_Account"))
                                objVendorInvHead.Vendor_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(objVendorInvHead.Vendor_Control_AC, obj.loc_code, True, trans)

                            End If
                            If clsCommon.myLen(objVendorInvHead.Vendor_Control_AC) <= 0 Then
                                Throw New Exception("Please set the vendor payable Account")
                            End If

                            'objVendorInvHead.Total_Add_Charge = obj.Total_Add_Charge

                            'objVendorInvHead.Add_Charge_Code1 = obj.Add_Charge_Code1
                            'objVendorInvHead.Add_Charge_Name1 = obj.Add_Charge_Name1
                            'objVendorInvHead.Add_Charge_Amt1 = obj.Add_Charge_Amt1

                            'objVendorInvHead.Add_Charge_Code2 = obj.Add_Charge_Code2
                            'objVendorInvHead.Add_Charge_Name2 = obj.Add_Charge_Name2
                            'objVendorInvHead.Add_Charge_Amt2 = obj.Add_Charge_Amt2

                            'objVendorInvHead.Add_Charge_Code3 = obj.Add_Charge_Code3
                            'objVendorInvHead.Add_Charge_Name3 = obj.Add_Charge_Name3
                            'objVendorInvHead.Add_Charge_Amt3 = obj.Add_Charge_Amt3

                            'objVendorInvHead.Add_Charge_Code4 = obj.Add_Charge_Code4
                            'objVendorInvHead.Add_Charge_Name4 = obj.Add_Charge_Name4
                            'objVendorInvHead.Add_Charge_Amt4 = obj.Add_Charge_Amt4

                            'objVendorInvHead.Add_Charge_Code5 = obj.Add_Charge_Code5
                            'objVendorInvHead.Add_Charge_Name5 = obj.Add_Charge_Name5
                            'objVendorInvHead.Add_Charge_Amt5 = obj.Add_Charge_Amt5

                            'objVendorInvHead.Add_Charge_Code6 = obj.Add_Charge_Code6
                            'objVendorInvHead.Add_Charge_Name6 = obj.Add_Charge_Name6
                            'objVendorInvHead.Add_Charge_Amt6 = obj.Add_Charge_Amt6

                            'objVendorInvHead.Add_Charge_Code7 = obj.Add_Charge_Code7
                            'objVendorInvHead.Add_Charge_Name7 = obj.Add_Charge_Name7
                            'objVendorInvHead.Add_Charge_Amt7 = obj.Add_Charge_Amt7

                            'objVendorInvHead.Add_Charge_Code8 = obj.Add_Charge_Code8
                            'objVendorInvHead.Add_Charge_Name8 = obj.Add_Charge_Name8
                            'objVendorInvHead.Add_Charge_Amt8 = obj.Add_Charge_Amt8

                            'objVendorInvHead.Add_Charge_Code9 = obj.Add_Charge_Code9
                            'objVendorInvHead.Add_Charge_Name9 = obj.Add_Charge_Name9
                            'objVendorInvHead.Add_Charge_Amt9 = obj.Add_Charge_Amt9

                            'objVendorInvHead.Add_Charge_Code10 = obj.Add_Charge_Code10
                            'objVendorInvHead.Add_Charge_Name10 = obj.Add_Charge_Name10
                            'objVendorInvHead.Add_Charge_Amt10 = obj.Add_Charge_Amt10


                            objVendorInvHead.Arr = New List(Of clsVedorInvoiceDetail)
                            Dim ii As Integer = 0
                            Dim isFirstTime As Boolean = True
                            ' Dim strFirstItemCode As String = GetFirstItemCode(obj.ObjList)
                            'objVendorInvHead.Empty_Amount = obj.Tot_Empty_Amount
                            objVendorInvHead.Total_Landed_Amt = 0

                            objVendorInvHead.ArrAssetEMI = New List(Of clsAPInvoiceAssetEMIDetails)()



                            ''Set AP Invvoice Detail Table

                            Dim strInvCtrlAC As String = clsCommon.myCstr(dt.Rows(0)("SECURITY_ACCOUNT"))
                            If clsCommon.myLen(strInvCtrlAC) <= 0 Then
                                Throw New Exception("Please set Security Account for Vendor Account :" + clsCommon.myCstr(dt.Rows(0)("Acct_Set_Code")))
                            End If
                            strInvCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvCtrlAC, obj.loc_code, True, trans)




                            ii = ii + 1
                            objVendorInvDetail.Detail_Line_No = ii
                            objVendorInvDetail.GL_Account_Code = strInvCtrlAC
                            objVendorInvDetail.GL_Account_Desc = clsGLAccount.GetName(strInvCtrlAC, trans)
                            objVendorInvDetail.Amount = dblAmount

                            objVendorInvDetail.Discount_Per = 0
                            objVendorInvDetail.Discount = 0
                            objVendorInvDetail.Amount_less_Discount = dblAmount
                            objVendorInvDetail.Total_Tax = 0
                            objVendorInvDetail.Total_Amount = dblAmount
                            objVendorInvDetail.Landed_Amount = dblAmount
                            ''End of Set AP Invvoice Detail Table

                            If (clsCommon.myLen(objVendorInvDetail.GL_Account_Code) > 0) Then
                                objVendorInvHead.Arr.Add(objVendorInvDetail)
                            End If

                            ''Set AP Invvoice Header Table
                            objVendorInvHead.Total_Landed_Amt += dblAmount
                            objVendorInvHead.Discount_Base += dblAmount
                            objVendorInvHead.Discount_Amount += 0
                            objVendorInvHead.Amount_Less_Discount += dblAmount
                            objVendorInvHead.Document_Total += dblAmount
                            objVendorInvHead.Balance_Amt += dblAmount
                            ''End of Set AP Invvoice Header Table

                            objVendorInvHead.Empty_Amount = 0 'obj.Tot_Empty_Amount
                            If objVendorInvHead.Empty_Amount > 0 Then
                                If clsCommon.myLen(objVendorInvHead.Empty_Account) <= 0 Then
                                    Throw New Exception("Please set Inventory Control Empties")
                                End If
                                objVendorInvHead.Document_Total += objVendorInvHead.Empty_Amount
                            End If
                            If (objVendorInvHead.Arr Is Nothing OrElse objVendorInvHead.Arr.Count <= 0) Then
                                Throw New Exception("No GL Account Found For AP Invoice")
                            End If
                            ''multicurrency
                            'objVendorInvHead.CURRENCY_CODE = obj.CURRENCY_CODE
                            'objVendorInvHead.ConvRate = 1
                            objVendorInvHead.ApplicableFrom = clsCommon.GetPrintDate(strPostDate, "dd/MMM/yyyy")
                            ''end multicurrency

                            objVendorInvHead.SaveData(objVendorInvHead, True, trans)
                            clsVedorInvoiceHead.PostData("", objVendorInvHead.Document_No, "", trans, strPostDate)
                        End If
                    End If
                End If




            End If
        End If


    End Sub


    Private Shared Sub BlankRegisteredVendorColumn(ByVal obj As clsVedorInvoiceHead, ByVal tempRegVendor As String, ByVal trans As SqlTransaction)
        Dim qry As String
        If clsCommon.myLen(obj.Against_POInvoice_No) > 0 Then
            qry = "Update TSPL_PI_HEAD set GSTRegistered=" + tempRegVendor + ", Purchase_Tax_Invoice=null,Purchase_Tax_Invoice_Type=null where PI_No='" + obj.Against_POInvoice_No + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        ElseIf clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 Then
            qry = "Update TSPL_MILK_PURCHASE_INVOICE_HEAD set GSTRegistered=" + tempRegVendor + ", Purchase_Tax_Invoice=null,Purchase_Tax_Invoice_Type=null where DOC_Code='" + obj.Against_MillkPurchaseInvoice_No + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        ElseIf clsCommon.myLen(obj.Against_BulkMillkPurchaseInvoice_No) > 0 Then
            qry = "Update TSPL_Bulk_MILK_PURCHASE_INVOICE_HEAD set GSTRegistered=" + tempRegVendor + ", Purchase_Tax_Invoice=null,Purchase_Tax_Invoice_Type=null where DOC_no='" + obj.Against_BulkMillkPurchaseInvoice_No + "'"
            clsDBFuncationality.ExecuteNonQuery(qry, trans)
        End If
        qry = "Update TSPL_VENDOR_INVOICE_HEAD set GSTRegistered=" + tempRegVendor + ",Purchase_Tax_Invoice=null,Purchase_Tax_Invoice_Type=null where Document_No='" + obj.Document_No + "'"
        clsDBFuncationality.ExecuteNonQuery(qry, trans)
    End Sub
    ''richa TEC/02/11/18-000361 on 24 Dec,2018
    Public Shared Function CreateJournalEntryForOpening(ByVal obj As clsVedorInvoiceHead, ByVal FormId As String, ByVal strPostDate As Date, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoRecreatedOnly As String = Nothing, Optional ByVal isForUnpostedTransaction As Boolean = False, Optional ByVal strasset As Boolean = False) As Boolean
        Try
            Dim qry As String = "select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No='" + obj.Document_No + "' "
            Dim strVoucherNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))

            If strVoucherNoRecreatedOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoRecreatedOnly) > 0 Then
                strVoucherNo = strVoucherNoRecreatedOnly
            End If

            ''===========these column is for GL entry=========do for multicurrency==if currency is differ form basecurrency=======================================
            Dim coll As New Hashtable()
            If clsCommon.myLen(obj.CURRENCY_CODE) > 0 AndAlso clsCommon.CompairString(obj.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) <> CompairStringResult.Equal Then
                coll = New Hashtable()
                clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE)
                clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
                clsCommon.AddColumnsForChange(coll, "ConvRateOld", obj.ConvRate)
            End If
            ''===============================================================================================


            Dim isTaxRecoverable As Boolean = False
            Dim NonRecovTaxAmount As Decimal = 0
            Dim Tax_Liability_Account As String = ""
            Dim strEntryDesc As String = ""
            Dim strSrcType As String = ""
            Dim strSrcDesc As String = ""
            Dim strRemarks As String = ""

            If (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                strEntryDesc = "Vendor Opening"
                strSrcType = "GL-JE"
                strSrcDesc = "AP Debit"
                strRemarks = " AP Invoice for Vendor: " + obj.Vendor_Code + " - " + obj.Vendor_Name + "  "
            ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                strEntryDesc = "Vendor Opening"
                strSrcType = "GL-JE"
                strSrcDesc = "AP Credit Note"
                strRemarks = " AP Invoice for Vendor: " + obj.Vendor_Code + " - " + obj.Vendor_Name + " "
            End If

            Dim strQ As String = Nothing
            Dim ArryLst As ArrayList = New ArrayList()
            If (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
                Dim strPayable_Account As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_VENDOR_ACCOUNT_SET.Payable_Account ,'') as Payable_Account from TSPL_VENDOR_ACCOUNT_SET left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER.Vendor_Account  =TSPL_VENDOR_ACCOUNT_SET.Acct_Set_Code  where TSPL_VENDOR_MASTER.Vendor_Code ='" & obj.Vendor_Code & "'", trans))
                If clsCommon.myLen(strPayable_Account) = 0 Then
                    Throw New Exception("Please set Payable Control Account for Vendor - " + obj.Vendor_Code)
                End If
                strPayable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(strPayable_Account, obj.loc_code, True, trans)
                Dim Acc1() As String = {strPayable_Account, (obj.Document_Total)}
                ArryLst.Add(Acc1)


                Dim strVendor_Opening_Clearing_AC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_VENDOR_ACCOUNT_SET.Opening_Clearing ,'') as Vendor_Opening_Clearing_AC from TSPL_VENDOR_ACCOUNT_SET left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER.Vendor_Account  =TSPL_VENDOR_ACCOUNT_SET.Acct_Set_Code  where TSPL_VENDOR_MASTER.Vendor_Code ='" & obj.Vendor_Code & "'", trans))
                If clsCommon.myLen(strVendor_Opening_Clearing_AC) = 0 Then
                    Throw New Exception("Please set Vendor Opening Clearing Account for Vendor - " + obj.Vendor_Code)
                End If
                strVendor_Opening_Clearing_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(strVendor_Opening_Clearing_AC, obj.loc_code, True, trans)
                Dim Acc2() As String = {strVendor_Opening_Clearing_AC, (obj.Document_Total) * -1}
                ArryLst.Add(Acc2)

            ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
                Dim strPayable_Account As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_VENDOR_ACCOUNT_SET.Payable_Account ,'') as Payable_Account from TSPL_VENDOR_ACCOUNT_SET left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER.Vendor_Account  =TSPL_VENDOR_ACCOUNT_SET.Acct_Set_Code  where TSPL_VENDOR_MASTER.Vendor_Code ='" & obj.Vendor_Code & "'", trans))
                If clsCommon.myLen(strPayable_Account) = 0 Then
                    Throw New Exception("Please set Payable Control Account for Vendor - " + obj.Vendor_Code)
                End If
                strPayable_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(strPayable_Account, obj.loc_code, True, trans)
                Dim Acc1() As String = {strPayable_Account, (obj.Document_Total) * -1}
                ArryLst.Add(Acc1)


                Dim strVendor_Opening_Clearing_AC As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull(TSPL_VENDOR_ACCOUNT_SET.Opening_Clearing ,'') as Vendor_Opening_Clearing_AC from TSPL_VENDOR_ACCOUNT_SET left outer join TSPL_VENDOR_MASTER  on TSPL_VENDOR_MASTER.Vendor_Account  =TSPL_VENDOR_ACCOUNT_SET.Acct_Set_Code  where TSPL_VENDOR_MASTER.Vendor_Code ='" & obj.Vendor_Code & "'", trans))
                If clsCommon.myLen(strVendor_Opening_Clearing_AC) = 0 Then
                    Throw New Exception("Please set Vendor Opening Clearing Account for Vendor - " + obj.Vendor_Code)
                End If
                strVendor_Opening_Clearing_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(strVendor_Opening_Clearing_AC, obj.loc_code, True, trans)
                Dim Acc2() As String = {strVendor_Opening_Clearing_AC, (obj.Document_Total)}
                ArryLst.Add(Acc2)

            End If
            Dim strPrefixTransType As String = clsDocTransactionType.DirectAP

            ''clsJournalMaster.FunGrnlEntryWithTrans(True, 0, "", "N", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, obj.Document_Date, strEntryDesc, strSrcType, strSrcDesc, obj.Document_No, obj.Description, "C", obj.Customer_Code, obj.Customer_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, "", strRemarks, Nothing, coll)
            clsJournalMaster.FunGrnlEntryWithTrans(True, 0, strPrefixTransType, "", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, strPostDate, strEntryDesc, strSrcType, strSrcDesc, obj.Document_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, "", strRemarks, Nothing, coll)
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
        Return True
    End Function
    ''-------------------------------------------------
    Public Shared Function CreateJournalEntry(ByVal obj As clsVedorInvoiceHead, ByVal FormId As String, ByVal strPostDate As Date, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoRecreatedOnly As String = Nothing, Optional ByVal isForUnpostedTransaction As Boolean = False, Optional ByVal strasset As Boolean = False) As Boolean
        Dim qry As String = "select Voucher_No from TSPL_JOURNAL_MASTER where Source_Doc_No='" + obj.Document_No + "' and Source_Code in ('AP-IN','AP-CN','AP-DN')"
        Dim strVoucherNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        Dim ApplyNoGSTCreditIndependentlyOnVendorServiceCharge As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, clsFixedParameterCode.ApplyNoGSTCreditIndependentlyOnVendorServiceCharge, trans)) = 1, True, False)
        If strVoucherNoRecreatedOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoRecreatedOnly) > 0 Then
            strVoucherNo = strVoucherNoRecreatedOnly
        End If
        ''GST Should applicable and Not a registed vendor
        Dim isAddGSTAccounts As Boolean = clsERPFuncationality.GetGSTStatus(clsCommon.myCDate(obj.Invoice_Entry_Date))
        Dim isGSTRegisteredVendor As Boolean
        If clsCommon.myLen(obj.Against_POInvoice_No) > 0 Then
            isGSTRegisteredVendor = IIf(obj.GSTRegistered = 1, True, False)
        ElseIf obj.RCM Then ''LLLLLLLLLLast Work Balwinder
            qry = "select 1 from TSPL_VENDOR_MASTER where Vendor_Code='" + obj.Vendor_Code + "' and GSTRegistered=1 "
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                isGSTRegisteredVendor = True
            Else
                isGSTRegisteredVendor = False
            End If
        Else
            If ApplyNoGSTCreditIndependentlyOnVendorServiceCharge = True Then
                qry = "select 1 from TSPL_VENDOR_MASTER where Vendor_Code='" + obj.Vendor_Code + "' and GSTRegistered=1 "
                Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    isGSTRegisteredVendor = True
                Else
                    isGSTRegisteredVendor = False
                End If
            Else
                isGSTRegisteredVendor = clsVendorMaster.IsGSTRegisteredVendor(obj.Vendor_Code, trans)
            End If
        End If
        Dim tempRegVendor As String = "1"
        If isAddGSTAccounts Then ''If GST Start Date is OFF tempRegVendor should be 1
            If Not isGSTRegisteredVendor Then
                tempRegVendor = "0" ''If GST Start Date is ON and Vendor unregister then tempRegVendor should be 0
            End If
        End If
        If obj.RCM Then
            isAddGSTAccounts = isAddGSTAccounts AndAlso True
        ElseIf Not (isGSTRegisteredVendor) Then
            isAddGSTAccounts = isAddGSTAccounts AndAlso True
        Else
            isAddGSTAccounts = False
        End If
        isAddGSTAccounts = isAddGSTAccounts AndAlso clsTaxGroupMaster.IsHavingRecoverableTaxAuthority(obj.Tax_Group, "P", trans)

        If (obj.TAX1_GLAC_Amt + obj.TAX1_GLAC2_Amt + obj.TAX1_GLAC3_Amt + obj.TAX1_GLAC4_Amt + obj.TAX1_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX2_GLAC_Amt + obj.TAX2_GLAC2_Amt + obj.TAX2_GLAC3_Amt + obj.TAX2_GLAC4_Amt + obj.TAX2_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX3_GLAC_Amt + obj.TAX3_GLAC2_Amt + obj.TAX3_GLAC3_Amt + obj.TAX3_GLAC4_Amt + obj.TAX3_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX4_GLAC_Amt + obj.TAX4_GLAC2_Amt + obj.TAX4_GLAC3_Amt + obj.TAX4_GLAC4_Amt + obj.TAX4_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX5_GLAC_Amt + obj.TAX5_GLAC2_Amt + obj.TAX5_GLAC3_Amt + obj.TAX5_GLAC4_Amt + obj.TAX5_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX6_GLAC_Amt + obj.TAX6_GLAC2_Amt + obj.TAX6_GLAC3_Amt + obj.TAX6_GLAC4_Amt + obj.TAX6_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX7_GLAC_Amt + obj.TAX7_GLAC2_Amt + obj.TAX7_GLAC3_Amt + obj.TAX7_GLAC4_Amt + obj.TAX7_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX8_GLAC_Amt + obj.TAX8_GLAC2_Amt + obj.TAX8_GLAC3_Amt + obj.TAX8_GLAC4_Amt + obj.TAX8_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX9_GLAC_Amt + obj.TAX9_GLAC2_Amt + obj.TAX9_GLAC3_Amt + obj.TAX9_GLAC4_Amt + obj.TAX9_GLAC5_Amt) <> 0 OrElse + _
            (obj.TAX10_GLAC_Amt + obj.TAX10_GLAC2_Amt + obj.TAX10_GLAC3_Amt + obj.TAX10_GLAC4_Amt + obj.TAX10_GLAC5_Amt) <> 0 Then
            isAddGSTAccounts = isAddGSTAccounts AndAlso True
        Else
            isAddGSTAccounts = isAddGSTAccounts AndAlso False
        End If

        Dim objBalAdvTaxAmt As clsPOAdvanceAdjustmentKnockOff = Nothing
        Dim strPurchaseTaxInvoiceType As String = ""
        Dim isSkipGST As Boolean = False
        If clsCommon.myLen(obj.Against_POInvoice_No) > 0 OrElse clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
            qry = "select sum(case when isnull( Skip_GST,0)=1 then 1 else 0 end) as NoOfSkipGSTItem,sum(case when isnull( Skip_GST,0)=0 then 1 else 0 end) as NoOfNonSkipGSTItem from tspl_item_master where item_Code in "
            If clsCommon.myLen(obj.Against_POInvoice_No) > 0 Then
                qry += "(select Item_Code from TSPL_PI_DETAIL where PI_No='" + obj.Against_POInvoice_No + "')"
            ElseIf clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
                qry += "(select Item_Code from TSPL_PR_DETAIL where PR_No='" + obj.Against_PurchaseReturn_No + "')"
            End If
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                If clsCommon.myCdbl(dt.Rows(0)("NoOfSkipGSTItem")) > 0 Then
                    isSkipGST = True
                    isAddGSTAccounts = isAddGSTAccounts AndAlso False
                End If
            End If
            dt = Nothing
        End If
        Dim strPIType As String = ""
        If clsCommon.myLen(obj.Against_POInvoice_No) > 0 Then
            strPIType = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_Type from TSPL_pi_head where PI_No='" & obj.Against_POInvoice_No & "'", trans))
        End If
        If (Not isSkipGST AndAlso (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) OrElse (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal)) Then
            If isAddGSTAccounts Then
                strPurchaseTaxInvoiceType = "P" ''Purchase Tax Invoice Document
                Dim strDocNo As String = obj.Against_POInvoice_No
                Dim strTransType As String = "PI"
                If clsCommon.CompairString(obj.RefDocType, "WO") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.RefDocNo) > 0 Then
                    strDocNo = obj.Document_No
                    strTransType = obj.RefDocType
                End If

                objBalAdvTaxAmt = clsPOAdvanceAdjustmentKnockOff.GetBalanceAdvanceAmt(strDocNo, strTransType, trans)
                If Not (clsCommon.myLen(obj.Purchase_Tax_Invoice) > 0 AndAlso clsCommon.CompairString(obj.Purchase_Tax_Invoice_Type, strPurchaseTaxInvoiceType) = CompairStringResult.Equal) Then
                    If obj.RCM AndAlso isGSTRegisteredVendor Then
                        BlankRegisteredVendorColumn(obj, tempRegVendor, trans)
                    Else
                        If clsCommon.myLen(obj.Against_POInvoice_No) > 0 Then
                            If Not clsCommon.CompairString(strPIType, "MT") = CompairStringResult.Equal Then
                                obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseTaxInvoice, clsDocTransactionType.GeneralPurchase, obj.loc_code, True)
                                qry = "Update TSPL_PI_HEAD set GSTRegistered=0, Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where PI_No='" + obj.Against_POInvoice_No + "'"
                                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                            End If
                        ElseIf clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 Then
                            obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseTaxInvoice, clsDocTransactionType.MilkPurchase, obj.loc_code, True)
                            qry = "Update TSPL_MILK_PURCHASE_INVOICE_HEAD set GSTRegistered=0, Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where DOC_Code='" + obj.Against_MillkPurchaseInvoice_No + "'"
                            clsDBFuncationality.ExecuteNonQuery(qry, trans)
                        ElseIf clsCommon.myLen(obj.Against_BulkMillkPurchaseInvoice_No) > 0 Then
                            obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseTaxInvoice, clsDocTransactionType.BulkMilkTaxInvoice, obj.loc_code, True)
                            qry = "Update TSPL_Bulk_MILK_PURCHASE_INVOICE_HEAD set GSTRegistered=0, Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where DOC_no='" + obj.Against_BulkMillkPurchaseInvoice_No + "'"
                            clsDBFuncationality.ExecuteNonQuery(qry, trans)
                        Else
                            obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseTaxInvoice, clsDocTransactionType.DirectPurchaseFromAP, obj.loc_code, True)
                        End If
                        If Not clsCommon.CompairString(strPIType, "MT") = CompairStringResult.Equal Then
                            If clsCommon.myLen(obj.Purchase_Tax_Invoice) <= 0 Then
                                Throw New Exception("Error in document Generation 'Purchase Tax Invoice' ")
                            End If
                            qry = "Update TSPL_VENDOR_INVOICE_HEAD set GSTRegistered=0,Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where Document_No='" + obj.Document_No + "'"
                            clsDBFuncationality.ExecuteNonQuery(qry, trans)
                        End If

                    End If
                End If
            ElseIf (clsERPFuncationality.GetGSTStatus(clsCommon.myCDate(obj.Invoice_Entry_Date)) AndAlso (Not isGSTRegisteredVendor)) Then
                strPurchaseTaxInvoiceType = "B" ''Bill of Supply Document
                If Not (clsCommon.myLen(obj.Purchase_Tax_Invoice) > 0 AndAlso clsCommon.CompairString(obj.Purchase_Tax_Invoice_Type, strPurchaseTaxInvoiceType) = CompairStringResult.Equal) Then
                    If clsCommon.myLen(obj.Against_POInvoice_No) > 0 Then
                        If Not clsCommon.CompairString(strPIType, "MT") = CompairStringResult.Equal Then
                            obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseBillOfSupply, clsDocTransactionType.GeneralPurchase, obj.loc_code, True)
                            qry = "Update TSPL_PI_HEAD set GSTRegistered=0, Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where PI_No='" + obj.Against_POInvoice_No + "'"
                            clsDBFuncationality.ExecuteNonQuery(qry, trans)
                        End If
                    ElseIf clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 Then
                        obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseBillOfSupply, clsDocTransactionType.MilkPurchase, obj.loc_code, True)
                        qry = "Update TSPL_MILK_PURCHASE_INVOICE_HEAD set GSTRegistered=0, Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where DOC_Code='" + obj.Against_MillkPurchaseInvoice_No + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    ElseIf clsCommon.myLen(obj.Against_BulkMillkPurchaseInvoice_No) > 0 Then
                        qry = "select isSRNTradeInvoice  from TSPL_Bulk_MILK_PURCHASE_INVOICE_HEAD where DOC_NO='" + obj.Against_BulkMillkPurchaseInvoice_No + "'"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
                        If dt Is Nothing OrElse dt.Rows.Count <= 0 Then
                            Throw New Exception("Bulk Milk purchase invoice no -" + obj.Against_BulkMillkPurchaseInvoice_No + " not found")
                        End If
                        If clsCommon.myCdbl(dt.Rows(0)("isSRNTradeInvoice")) = 1 Then
                            obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseBillOfSupply, clsDocTransactionType.BulkMilkTaxInvoiceTrade, obj.loc_code, True)
                        Else
                            Dim IsAgainstJobWork As Integer = 0
                            Dim Joblocation_Code As String = Nothing

                            qry = "select Joblocation_Code,IsAgainstJobWork  from TSPL_Bulk_MILK_PURCHASE_INVOICE_HEAD where DOC_NO='" + obj.Against_BulkMillkPurchaseInvoice_No + "'"
                            dt = clsDBFuncationality.GetDataTable(qry, trans)
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                IsAgainstJobWork = clsCommon.myCdbl(dt.Rows(0)("IsAgainstJobWork"))
                                Joblocation_Code = clsLocation.GetSegmentCode(clsCommon.myCstr(dt.Rows(0)("Joblocation_Code")), trans)
                            End If
                            If IsAgainstJobWork = 0 Then
                                obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseBillOfSupply, clsDocTransactionType.BulkMilkTaxInvoice, obj.loc_code, True)
                            Else
                                obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseBillOfSupply, clsDocTransactionType.BulkMilkTaxInvoice, Joblocation_Code, True)
                            End If

                        End If
                        qry = "Update TSPL_Bulk_MILK_PURCHASE_INVOICE_HEAD set GSTRegistered=0, Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where DOC_no='" + obj.Against_BulkMillkPurchaseInvoice_No + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    Else
                        obj.Purchase_Tax_Invoice = clsERPFuncationality.GetNextCode(trans, clsCommon.myCDate(obj.Invoice_Entry_Date), clsDocType.PurchaseBillOfSupply, clsDocTransactionType.DirectPurchaseFromAP, obj.loc_code, True)
                    End If
                    If Not clsCommon.CompairString(strPIType, "MT") = CompairStringResult.Equal Then
                        If clsCommon.myLen(obj.Purchase_Tax_Invoice) <= 0 Then
                            Throw New Exception("Error in document Generation 'Bill of Supply' ")
                        End If

                        qry = "Update TSPL_VENDOR_INVOICE_HEAD set GSTRegistered=0,Purchase_Tax_Invoice='" + obj.Purchase_Tax_Invoice + "',Purchase_Tax_Invoice_Type='" + strPurchaseTaxInvoiceType + "' where Document_No='" + obj.Document_No + "'"
                        clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    End If

                End If
            Else
                BlankRegisteredVendorColumn(obj, tempRegVendor, trans)
            End If
        Else
            ''done by richa discussed with ranjana mam because at the time of debitnote journal entry were created wrongly it was not considered payable control account of taxes
            'isAddGSTAccounts = False
        End If



        ''End of GST Should applicable and Not a registed vendor
        Dim objTM As clsTaxMaster = Nothing
        'Dim qry As String = ""
        Dim ArryLst As ArrayList = New ArrayList()
        Dim strLocation As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code='" + obj.Arr(0).GL_Account_Code + "'", trans))

        Dim dblConvRate As Double = IIf(obj.ConvRate = 0, 1, obj.ConvRate)

        Dim isTaxRecoerable As Boolean = False
        If obj.TAX1_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX1_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX1_GLAC, obj.TAX1_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX1_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX1_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX1_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX1_GLAC2, obj.TAX1_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX1_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX1_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX1_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX1_GLAC3, obj.TAX1_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX1_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX1_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX1_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX1_GLAC4, obj.TAX1_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX1_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX1_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX1_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX1_GLAC5, obj.TAX1_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX1_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX1_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX1, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX1)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX1)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX1_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX1_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX1_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX1_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX1)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX1_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If

        If obj.TAX2_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX2_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX2_GLAC, obj.TAX2_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX2_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX2_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX2_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX2_GLAC2, obj.TAX2_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX2_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX2_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX2_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX2_GLAC3, obj.TAX2_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX2_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX2_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX2_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX2_GLAC4, obj.TAX2_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX2_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX2_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX2_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX2_GLAC5, obj.TAX2_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX2_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX2_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX2, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX2)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX2)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX2_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX2_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX2_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX2_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX2)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX2_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If

        If obj.TAX3_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX3_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX3_GLAC, obj.TAX3_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX3_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX3_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX3_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX3_GLAC2, obj.TAX3_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX3_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX3_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX3_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX3_GLAC3, obj.TAX3_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX3_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX3_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX3_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX3_GLAC4, obj.TAX3_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX3_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX3_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX3_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX3_GLAC5, obj.TAX3_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX3_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX3_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX3, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX3)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX3)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX3_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX3_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX3_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX3_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX3)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX3_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If

        If obj.TAX4_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX4_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX4_GLAC, obj.TAX4_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX4_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX4_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX4_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX4_GLAC2, obj.TAX4_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX4_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX4_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX4_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX4_GLAC3, obj.TAX4_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX4_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX4_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX4_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX4_GLAC4, obj.TAX4_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX4_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX4_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX4_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX4_GLAC5, obj.TAX4_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX4_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX4_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX4, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX4)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX4)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX4_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX4_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX4_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX4_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX4)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX4_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If

        If obj.TAX5_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX5_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX5_GLAC, obj.TAX5_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX5_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX5_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX5_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX5_GLAC2, obj.TAX5_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX5_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX5_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX5_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX5_GLAC3, obj.TAX5_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX5_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX5_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX5_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX5_GLAC4, obj.TAX5_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX5_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX5_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX5_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX5_GLAC5, obj.TAX5_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX5_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX5_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX5, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX5)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX5)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX5_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX5_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX5_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX5_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX5)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX5_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If
        If obj.TAX6_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX6_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX6_GLAC, obj.TAX6_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX6_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX6_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX6_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX6_GLAC2, obj.TAX6_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX6_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX6_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX6_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX6_GLAC3, obj.TAX6_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX6_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX6_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX6_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX6_GLAC4, obj.TAX6_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX6_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX6_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX6_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX6_GLAC5, obj.TAX6_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX6_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX6_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX6, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX6)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX6)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX6_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX6_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX6_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX6_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX6)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX6_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If
        If obj.TAX7_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX7_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX7_GLAC, obj.TAX7_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX7_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX7_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX7_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX7_GLAC2, obj.TAX7_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX7_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX7_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX7_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX7_GLAC3, obj.TAX7_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX7_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX7_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX7_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX7_GLAC4, obj.TAX7_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX7_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX7_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX7_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX7_GLAC5, obj.TAX7_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX7_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX7_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX7, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX7)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX7)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX7_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX7_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX7_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX7_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX7)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX7_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If
        If obj.TAX8_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX8_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX8_GLAC, obj.TAX8_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX8_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX8_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX8_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX8_GLAC2, obj.TAX8_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX8_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX8_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX8_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX8_GLAC3, obj.TAX8_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX8_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX8_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX8_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX8_GLAC4, obj.TAX8_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX8_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX8_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX8_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX8_GLAC5, obj.TAX8_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX8_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX8_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX8, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX8)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX8)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX8_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX8_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX8_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX8_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX8)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX8_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If
        If obj.TAX9_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX9_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX9_GLAC, obj.TAX9_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX9_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX9_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX9_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX9_GLAC2, obj.TAX9_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX9_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX9_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX9_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX9_GLAC3, obj.TAX9_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX9_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX9_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX9_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX9_GLAC4, obj.TAX9_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX9_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX9_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX9_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX9_GLAC5, obj.TAX9_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX9_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX9_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX9, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX9)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX9)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX9_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX9_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX9_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX9_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX9)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX9_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If


        If obj.TAX10_Amt <> 0 Then
            If obj.No_GST_Credit Then
                isTaxRecoerable = True
            Else
                isTaxRecoerable = False
                If clsCommon.myLen(obj.TAX10_GLAC) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX10_GLAC, obj.TAX10_GLAC_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX10_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX10_GLAC_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX10_GLAC2) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX10_GLAC2, obj.TAX10_GLAC2_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX10_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX10_GLAC2_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX10_GLAC3) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX10_GLAC3, obj.TAX10_GLAC3_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX10_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX10_GLAC3_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX10_GLAC4) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX10_GLAC4, obj.TAX10_GLAC4_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX10_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX10_GLAC4_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
                If clsCommon.myLen(obj.TAX10_GLAC5) > 0 Then
                    Dim AccInvDR() As String = {obj.TAX10_GLAC5, obj.TAX10_GLAC5_Amt * dblConvRate}
                    ArryLst.Add(AccInvDR)
                    If obj.TAX10_Amt < 0 Then
                        Dim AccInvDR1() As String = {obj.Vendor_Control_AC, -1 * obj.TAX10_GLAC5_Amt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    End If
                    isTaxRecoerable = True
                End If
            End If
            If isAddGSTAccounts AndAlso isTaxRecoerable Then
                objTM = clsTaxMaster.GetData(obj.TAX10, trans)
                If objTM Is Nothing OrElse clsCommon.myLen(objTM.Tax_Code) <= 0 Then
                    Throw New Exception("Tax authority not found:" + obj.TAX10)
                End If

                If clsCommon.myLen(objTM.PayableControl) <= 0 Then
                    Throw New Exception("Please map payable control A/C in Tax authority :" + obj.TAX10)
                End If
                objTM.PayableControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.PayableControl, obj.loc_code, True, trans)
                Dim str1() As String = {objTM.PayableControl, -1 * obj.TAX10_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing

                str1 = {obj.Vendor_Control_AC, obj.TAX10_Amt * dblConvRate}
                ArryLst.Add(str1)
                str1 = Nothing
                If objBalAdvTaxAmt IsNot Nothing Then
                    If objBalAdvTaxAmt.TAX10_Amt <> 0 Then
                        str1 = {objTM.PayableControl, objBalAdvTaxAmt.TAX10_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing

                        If clsCommon.myLen(objTM.DepositControl) <= 0 Then
                            Throw New Exception("Please map deposit control A/C in Tax authority :" + obj.TAX10)
                        End If
                        objTM.DepositControl = clsERPFuncationality.ChangeGLAccountLocationSegment(objTM.DepositControl, obj.loc_code, True, trans)
                        str1 = {objTM.DepositControl, -1 * objBalAdvTaxAmt.TAX10_Amt * dblConvRate}
                        ArryLst.Add(str1)
                        str1 = Nothing
                    End If
                End If
            End If
        End If

        Dim strShipTolocation As String = ""
        Dim strBrachAC As String = ""
        Dim strBrachACofShiptoLoc As String = ""
        If clsCommon.myLen(obj.Against_POInvoice_No) > 0 OrElse clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
            If clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
                strShipTolocation = clsDBFuncationality.getSingleValue("select Ship_To_Location from TSPL_PR_HEAD where PR_No='" & obj.Against_PurchaseReturn_No & "'", trans)
            Else
                strShipTolocation = clsDBFuncationality.getSingleValue("select Ship_To_Location from TSPL_PI_HEAD where PI_No='" & obj.Against_POInvoice_No & "'", trans)
            End If

            '' find Location Segment of ship to loc
            Dim strShipTolocationSeg As String = clsLocation.GetSegmentCode(strShipTolocation, trans)
            '' obj.Loc_Code is already a location segment because in vendorinvoice head table location segment is saved 
            '' so changed the query against the bug raised for a PI post in MPD.(Panch Raj)
            If clsCommon.myLen(strShipTolocation) > 0 Then
                If clsCommon.CompairString(strShipTolocationSeg, obj.loc_code) = CompairStringResult.Equal Then
                    strShipTolocation = ""
                Else
                    qry = "select Branch_Account from TSPL_BRANCH_ACCOUNT_MAPPING where From_Location='" & strShipTolocationSeg & "' and To_Location='" & obj.loc_code & "'"
                    strBrachAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                    If clsCommon.myLen(strBrachAC) <= 0 Then
                        Throw New Exception("Please set Brach account with From Location=" & strShipTolocationSeg & " and To Location=" & obj.loc_code & "")
                    End If
                End If
            End If
        End If
        If clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 And obj.RefDocType = "MI-PI" Then
            ''IN RCDF No concept of Irregular_Mcc_Code
            'strShipTolocation = clsDBFuncationality.getSingleValue("select coalesce(Irregular_Mcc_Code,'') from TSPL_MILK_PURCHASE_INVOICE_HEAD where Doc_Code='" + obj.Against_MillkPurchaseInvoice_No + "'", trans)
            'If clsCommon.myLen(strShipTolocation) > 0 Then
            '    If clsCommon.CompairString(obj.loc_code, strShipTolocation) <> CompairStringResult.Equal Then
            '        Dim StrShipment As String = clsERPFuncationality.GetLocationSegment(strShipTolocation, trans)
            '        Dim strMainLoc As String = clsERPFuncationality.GetLocationSegment(obj.loc_code, trans)
            '        qry = "select Branch_Account from TSPL_BRANCH_ACCOUNT_MAPPING where From_Location='" + StrShipment + "' and To_Location='" + strMainLoc + "'"
            '        strBrachACofShiptoLoc = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            '        If clsCommon.myLen(strBrachACofShiptoLoc) <= 0 Then
            '            Throw New Exception("Please set Brach account with From Location=" + StrShipment + " and To Location=" + strMainLoc + "")
            '        End If
            '        qry = "select Branch_Account from TSPL_BRANCH_ACCOUNT_MAPPING where To_Location='" + StrShipment + "' and From_Location='" + strMainLoc + "'"
            '        strBrachAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
            '        If clsCommon.myLen(strBrachAC) <= 0 Then
            '            Throw New Exception("Please set Brach account with  From Location=" + strMainLoc + " and To Location=" + StrShipment + " ")
            '        End If
            '    End If
            'Else
            '    strShipTolocation = ""
            'End If
        End If

        '' Dim isFirstTime As Boolean = True
        If clsCommon.CompairString(obj.Is_ProRated, "Y") = CompairStringResult.Equal Then
            Dim Purchase_Control_Account As String = ""
            If clsCommon.CompairString(obj.RefDocType, "BS") = CompairStringResult.Equal Then
                Dim ObjSRN As clsBulkMilkSRN = clsBulkMilkSRN.getData(obj.RefDocNo, NavigatorType.Current, False, trans)
                If ObjSRN IsNot Nothing And clsCommon.myLen(ObjSRN.SRN_NO) > 0 Then
                    For Each objTR As clsVedorInvoiceDetail In obj.Arr
                        Dim intCount As Integer = obj.Arr.Count
                        Dim dblLedgeerNonRecoverableAmt As Double = clsVedorInvoiceHead.GetTaxAmt(objTR, trans)
                        Dim dblAddionalCost As Double = Math.Round((obj.Total_Add_Charge / intCount), 6)
                        Dim tempAmt As Double = objTR.Amount_less_Discount + dblAddionalCost + dblLedgeerNonRecoverableAmt
                        ''richa agarwal 21/122016
                        objTR.GL_Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(objTR.GL_Account_Code, obj.loc_code, trans)
                        ''-------------------
                        Dim AccInvDR1() As String = {objTR.GL_Account_Code, tempAmt * dblConvRate}
                        ArryLst.Add(AccInvDR1)
                    Next
                End If
            Else
                Dim ObjSRN As clsSRNHead = clsSRNHead.GetData(obj.RefDocNo, NavigatorType.Current, trans)
                Dim Index As Integer = 1
                Dim PurchaseControlAmount As Decimal = 0.0
                If ObjSRN IsNot Nothing And clsCommon.myLen(ObjSRN.SRN_No) > 0 Then
                    For Each objSRNDetail As clsSRNDetail In ObjSRN.Arr
                        Purchase_Control_Account = "Select Purchase_Control_Account from TSPL_PURCHASE_ACCOUNTS WHERE Purchase_Class_Code= (Select Purchase_Class_Code from TSPL_ITEM_MASTER WHERE Item_Code='" & objSRNDetail.Item_Code & "')"
                        Purchase_Control_Account = clsCommon.myCstr(clsDBFuncationality.getSingleValue(Purchase_Control_Account, trans))
                        PurchaseControlAmount = Math.Round((objSRNDetail.Item_Net_Amt / (ObjSRN.SRN_Total_Amt - ObjSRN.Total_Add_Charge)) * obj.Amount_Less_Discount, 2)
                        If clsCommon.myLen(Purchase_Control_Account) <= 0 Then
                            Throw New Exception("Purchase Control Account Not Found For Item : " & objSRNDetail.Item_Code & "")
                        End If
                        Purchase_Control_Account = clsERPFuncationality.ChangeGLAccountLocationSegment(Purchase_Control_Account, obj.loc_code, trans)
                        Dim AccInvDR() As String = {Purchase_Control_Account, PurchaseControlAmount * dblConvRate}
                        ArryLst.Add(AccInvDR)
                    Next
                End If
            End If
        Else
            For Each objTR As clsVedorInvoiceDetail In obj.Arr
                Dim intCount As Integer = obj.Arr.Count
                Dim dblLedgeerNonRecoverableAmt As Double = clsVedorInvoiceHead.GetTaxAmt(objTR, trans)
                Dim dblAddionalCost As Double = Math.Round((obj.Total_Add_Charge / intCount), 6)
                Dim dtbTempAmt As Double = 0
                dtbTempAmt = objTR.Amount_less_Discount + IIf(clsCommon.myLen(obj.Against_POInvoice_No) > 0 OrElse clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 OrElse clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0, 0, dblAddionalCost) + IIf(clsCommon.myLen(obj.Against_POInvoice_No) > 0 OrElse clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0, objTR.Landed_Amount, dblLedgeerNonRecoverableAmt)
                If ApplyNoGSTCreditIndependentlyOnVendorServiceCharge Then
                    If obj.RCM AndAlso obj.No_GST_Credit Then
                        dtbTempAmt += objTR.Total_Tax
                    ElseIf obj.No_GST_Credit Then
                        dtbTempAmt += objTR.Total_Tax
                    End If
                Else
                    If obj.RCM AndAlso obj.No_GST_Credit Then
                        dtbTempAmt += objTR.Total_Tax
                    End If
                End If

                Dim StrRecoControlAccount As String = ""
                If clsCommon.CompairString(objTR.Comments, "Y") = CompairStringResult.Equal Then
                    StrRecoControlAccount = objTR.Comments
                End If

                ' Dim AccInvDR() As String = {objTR.GL_Account_Code, dtbTempAmt, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code}
                If clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 AndAlso clsCommon.myLen(strShipTolocation) > 0 Then
                    strBrachAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strBrachAC, obj.loc_code, trans)
                    Dim AccInvDR() As String = {strBrachAC, dtbTempAmt * dblConvRate, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code, objTR.Hirerachy_Code3, objTR.Hirerachy_Code4, StrRecoControlAccount}
                    ArryLst.Add(AccInvDR)
                Else
                    Dim AccInvDR() As String = {objTR.GL_Account_Code, dtbTempAmt * dblConvRate, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code, objTR.Hirerachy_Code3, objTR.Hirerachy_Code4, StrRecoControlAccount}
                    ArryLst.Add(AccInvDR)
                End If


                If clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) <= 0 Or obj.RefDocType <> "MI-PI" Then
                    If clsCommon.myLen(strShipTolocation) > 0 Then
                        Dim strPaybleCleanigCtrlAC As String = clsERPFuncationality.ChangeGLAccountLocationSegment(objTR.GL_Account_Code, strShipTolocation, trans)
                        ' Dim AccCr2() As String = {strPaybleCleanigCtrlAC, dtbTempAmt, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code}
                        Dim AccCr2() As String = {strPaybleCleanigCtrlAC, dtbTempAmt * dblConvRate, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code, objTR.Hirerachy_Code3, objTR.Hirerachy_Code4}
                        ArryLst.Add(AccCr2)

                        If clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
                            strBrachAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strBrachAC, strShipTolocation, trans)
                        End If
                        Dim AccCr3() As String = {strBrachAC, -1 * dtbTempAmt * dblConvRate}
                        ArryLst.Add(AccCr3)
                    End If
                Else
                    If clsCommon.myLen(strShipTolocation) > 0 Then
                        Dim strPaybleCleanigCtrlAC As String = clsERPFuncationality.ChangeGLAccountLocationSegment(objTR.GL_Account_Code, strShipTolocation, trans)
                        Dim AccCr3() As String = {strBrachACofShiptoLoc, -1 * dtbTempAmt * dblConvRate}
                        ArryLst.Add(AccCr3)
                        strPaybleCleanigCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(objTR.GL_Account_Code, obj.loc_code, trans)
                        ' Dim AccCr6() As String = {strPaybleCleanigCtrlAC, dtbTempAmt, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code}
                        Dim AccCr6() As String = {strPaybleCleanigCtrlAC, dtbTempAmt * dblConvRate, "", "", objTR.Hirerachy_Code, objTR.Cost_Centre_Code, objTR.Hirerachy_Code3, objTR.Hirerachy_Code4}
                        ArryLst.Add(AccCr6)
                    End If
                End If
            Next
        End If
        If strasset = True Then
            qry = " Select TSPL_PURCHASE_ACCOUNTS.Inv_Control_Account,TSPL_ACQUISITION_DETAIL.Item_Net_Amt  from TSPL_ACQUISITION_DETAIL left outer join tspl_item_master on tspl_item_master.Item_Code =TSPL_ACQUISITION_DETAIL.Item_Code left outer join TSPL_PURCHASE_ACCOUNTS on TSPL_PURCHASE_ACCOUNTS.Purchase_Class_Code =tspl_item_master.Purchase_Class_Code " &
             " where Acquisition_Code ='" + obj.Against_Acquisition + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt Is Nothing OrElse dt.Rows.Count > 0 Then
                Dim strInvControlAcc As String = String.Empty
                For i = 0 To dt.Rows.Count - 1
                    strInvControlAcc = clsCommon.myCstr(dt.Rows(i)("Inv_Control_Account"))
                    strInvControlAcc = clsERPFuncationality.ChangeGLAccountLocationSegment(strInvControlAcc, obj.loc_code, trans)
                    If strInvControlAcc = "" Then
                        Throw New Exception("Please set Inventory Control Account in Purchase Account set")
                    End If
                    Dim AccInvCR() As String = {strInvControlAcc, -1 * (clsCommon.myCdbl(dt.Rows(i)("Item_Net_Amt"))) * dblConvRate}
                    ArryLst.Add(AccInvCR)
                Next
            Else
                Throw New Exception("Please set Inventory Control Account in Purchase Account set")
            End If
        Else
            ''richa agarwal 21/122016
            If clsCommon.myLen(obj.Vendor_Control_AC) <= 0 Then
                Throw New Exception("Please set Vendor Control Account in Vendor Account set")
            End If
            obj.Vendor_Control_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Vendor_Control_AC, obj.loc_code, True, trans)
            ''-------------------
            Dim AccInvCR() As String = {obj.Vendor_Control_AC, -1 * (obj.Document_Total + IIf(obj.TAX1_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX1, trans), -1 * obj.TAX1_Amt, 0) + IIf(obj.TAX2_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX2, trans), -1 * obj.TAX2_Amt, 0) + IIf(obj.TAX3_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX3, trans), -1 * obj.TAX3_Amt, 0) + IIf(obj.TAX4_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX4, trans), -1 * obj.TAX4_Amt, 0) + IIf(obj.TAX5_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX5, trans), -1 * obj.TAX5_Amt, 0) + IIf(obj.TAX6_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX6, trans), -1 * obj.TAX6_Amt, 0) + IIf(obj.TAX7_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX7, trans), -1 * obj.TAX7_Amt, 0) + IIf(obj.TAX8_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX8, trans), -1 * obj.TAX8_Amt, 0) + IIf(obj.TAX9_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX9, trans), -1 * obj.TAX9_Amt, 0) + IIf(obj.TAX10_Amt < 0 And clsTaxMaster.ISTaxRecoverableAC(obj.TAX10, trans), -1 * obj.TAX10_Amt, 0)) * dblConvRate}
            ArryLst.Add(AccInvCR)
        End If


        ''=====================added by shivani 
        If obj.Cash_Discount_Amt <> 0 Then
            Dim strDiscountCtrlAC As String = "select Discount_Account from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code = '" + obj.Account_Set + "'"
            strDiscountCtrlAC = clsCommon.myCstr(clsDBFuncationality.getSingleValue(strDiscountCtrlAC, trans))
            If clsCommon.myLen(strDiscountCtrlAC) <= 0 Then
                Throw New Exception("Please set the Discount account ")
            End If
            strDiscountCtrlAC = clsERPFuncationality.ChangeGLAccountLocationSegment(strDiscountCtrlAC, obj.loc_code, True, trans)
            Dim AccInvCR1() As String = {strDiscountCtrlAC, -1 * (obj.Cash_Discount_Amt) * dblConvRate}
            ArryLst.Add(AccInvCR1)
            Dim AccInvDR1() As String = {obj.Vendor_Control_AC, 1 * (obj.Cash_Discount_Amt) * dblConvRate}
            ArryLst.Add(AccInvDR1)
        End If
        '=============================================================

        If obj.RoundOffAmount <> 0 Then
            Dim dtTemp As DataTable = clsDBFuncationality.GetDataTable("select Round_Off,Short_Excess from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code = '" + obj.Account_Set + "'", trans)
            If dtTemp Is Nothing OrElse dtTemp.Rows.Count <= 0 Then
                Throw New Exception("Please set account set for vendor " + obj.Vendor_Code)
            End If

            Dim strACRoundInvCr As String = clsCommon.myCstr(dtTemp.Rows(0)("Round_Off"))
            If clsCommon.myLen(strACRoundInvCr) <= 0 Then
                Throw New Exception("Please set the Roundoff account in vendor account set for-" + obj.Account_Set)
            End If
            strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, strLocation, True, trans)
            Dim AccRoundInvCR() As String = {strACRoundInvCr, obj.RoundOffAmount * dblConvRate}
            ArryLst.Add(AccRoundInvCR)

            ''Comment by balwinder on 25/04/2019 because it take double diffenrence in Payable clearing report.
            'If clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 Then
            '    ''Short Excess acount will pick only in AP Invoice Against Milk purchase Invoice
            '    ''GKD/04/06/18-000142 by balwinder on 11/06/2018
            '    strACRoundInvCr = clsCommon.myCstr(dtTemp.Rows(0)("Short_Excess"))
            '    If clsCommon.myLen(strACRoundInvCr) <= 0 Then
            '        Throw New Exception("Please set the Short Excess account in vendor account set for-" + obj.Account_Set)
            '    End If
            '    strACRoundInvCr = clsERPFuncationality.ChangeGLAccountLocationSegment(strACRoundInvCr, strLocation, True, trans)
            '    Dim AccRoundInvDR12() As String = {strACRoundInvCr, -1 * obj.RoundOffAmount * dblConvRate}
            '    ArryLst.Add(AccRoundInvDR12)
            'End If

        End If



        If obj.Empty_Amount > 0 Then
            If clsCommon.myLen(obj.Empty_Account) < 0 Then
                Throw New Exception("Plese set items Container Debosit account")
            End If

            Dim AccInvDR() As String = {obj.Empty_Account, obj.Empty_Amount * dblConvRate}
            ArryLst.Add(AccInvDR)
        End If


        If (obj.RemittanceObject IsNot Nothing AndAlso obj.is_For_TDS = 0) AndAlso Not obj.TDS_Provision Then ''is_For_TDS Entry made by ap invoice is come in this section

            If obj.TDS_Actual_Amount <> 0 Then
                'Dim AccToInsertDR() As String = {obj.Vendor_Control_AC, obj.TDS_Actual_Amount }
                'ArryLst.Add(AccToInsertDR)
                'obj.RemittanceObject.Branch_GL_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.RemittanceObject.Branch_GL_AC, strLocation, True, trans)
                'Dim AccToInsertCR() As String = {obj.RemittanceObject.Branch_GL_AC, -1 * obj.TDS_Actual_Amount }
                'ArryLst.Add(AccToInsertCR)

                Dim AccToInsertDR() As String = {obj.Vendor_Control_AC, Math.Round(obj.TDS_Actual_Amount * dblConvRate, 0, MidpointRounding.AwayFromZero)}
                ArryLst.Add(AccToInsertDR)
                obj.RemittanceObject.Branch_GL_AC = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.RemittanceObject.Branch_GL_AC, strLocation, True, trans)
                Dim AccToInsertCR() As String = {obj.RemittanceObject.Branch_GL_AC, -1 * Math.Round(obj.TDS_Actual_Amount * dblConvRate, 0, MidpointRounding.AwayFromZero)}
                ArryLst.Add(AccToInsertCR)
            End If
        End If

        Dim pjvNOVochdesc As String
        Dim strVoucher_Desc As String
        If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
            strVoucher_Desc = "AP Debit Note Against " + obj.Document_No
        ElseIf clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal Then
            strVoucher_Desc = "AP Credit Note Against " + obj.Document_No
        ElseIf clsCommon.CompairString(FormId, "CSA-SALE") = CompairStringResult.Equal Then
            strVoucher_Desc = "AP Invoice Against CSA Sale Patti No. " + obj.Vendor_Invoice_No
        ElseIf clsCommon.CompairString(FormId, "VSP-MP") = CompairStringResult.Equal Then
            strVoucher_Desc = "AP Debit Note Against VSP and MP Collection. " + obj.Vendor_Invoice_No
        ElseIf clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 And clsCommon.CompairString(obj.RefDocType, "MI-PI") = CompairStringResult.Equal Then
            strVoucher_Desc = "Milk Purchase Invoice Against AP Invoice - " + obj.Document_No
            obj.Description = ""
        ElseIf clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 And clsCommon.CompairString(obj.RefDocType, "MI-CO") = CompairStringResult.Equal Then
            strVoucher_Desc = "EMP Against AP Invoice - " + obj.Document_No
            obj.Description = ""
        ElseIf clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 And clsCommon.CompairString(obj.RefDocType, "MI-IN") = CompairStringResult.Equal Then
            strVoucher_Desc = "Incentive Against AP Invoice - " + obj.Document_No
            obj.Description = ""
        ElseIf clsCommon.myLen(obj.Against_BulkMillkPurchaseInvoice_No) > 0 And clsCommon.CompairString(obj.RefDocType, "BM-PI") = CompairStringResult.Equal Then
            strVoucher_Desc = "Bulk Milk Purchase Invoice No. " & obj.Against_BulkMillkPurchaseInvoice_No & ", Against AP Invoice - " + obj.Document_No
            obj.Description = ""
        ElseIf clsCommon.myLen(obj.Against_Acquisition) > 0 AndAlso clsCommon.myLen(obj.Description) > 0 Then
            strVoucher_Desc = obj.Description
        Else
            strVoucher_Desc = "AP Invoice Against " + obj.Document_No
        End If
        If Len(obj.Against_POInvoice_No) > 0 Then
            Dim qry1 As String = " select pjv_no from TSPL_PJV_HEAD where Invoice_No ='" + obj.Against_POInvoice_No + "'"

            Dim pjvno As String = clsDBFuncationality.getSingleValue(qry1, trans)
            pjvNOVochdesc = "PJVNO-" + pjvno + "-" + strVoucher_Desc
        Else
            pjvNOVochdesc = strVoucher_Desc
        End If

        '========================do code for multicurrency conversion at gl entry-----16/04/2015-Monika----------------------
        Dim coll As New Hashtable()
        If clsCommon.myLen(obj.CURRENCY_CODE) > 0 AndAlso clsCommon.CompairString(obj.CURRENCY_CODE, objCommonVar.BaseCurrencyCode) <> CompairStringResult.Equal Then
            coll = New Hashtable()
            clsCommon.AddColumnsForChange(coll, "CURRENCY_CODE", obj.CURRENCY_CODE)
            clsCommon.AddColumnsForChange(coll, "ConvRate", obj.ConvRate)
            clsCommon.AddColumnsForChange(coll, "ConvRateOld", obj.ConvRate)
        End If
        '===========================end here=========================================================================================

        Dim strPrefixTransType As String = clsDocTransactionType.DirectAP
        If clsCommon.myLen(obj.Against_MillkPurchaseInvoice_No) > 0 Then
            strPrefixTransType = clsDocTransactionType.MccProc
        ElseIf clsCommon.myLen(obj.Against_BulkMillkPurchaseInvoice_No) > 0 Then
            strPrefixTransType = clsDocTransactionType.BulkProc
        ElseIf clsCommon.myLen(obj.Against_POInvoice_No) > 0 OrElse clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
            strPrefixTransType = clsDocTransactionType.GeneralPurchase
        End If
        If obj.is_For_Provision Then
            qry = "select sum(Knockoff_Amount) as Knockoff_Amount from TSPL_PROVISION_ENTRY_KNOCKOFF where AP_Invoice_No='" + obj.Document_No + "'"
            Dim KOAmt As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
            If KOAmt > 0 Then
                Dim DetailAccount As String = clsERPFuncationality.ChangeGLAccountLocationSegment(obj.Arr(0).GL_Account_Code, obj.loc_code, True, trans)
                Dim AccCr2() As String = {DetailAccount, -1 * KOAmt * dblConvRate, "", "", obj.Arr(0).Hirerachy_Code, obj.Arr(0).Cost_Centre_Code}
                ArryLst.Add(AccCr2)

                qry = "select Freight_Provision from TSPL_VENDOR_ACCOUNT_SET where Acct_Set_Code = '" + obj.Account_Set + "'"
                Dim FreightProvision As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
                If clsCommon.myLen(FreightProvision) <= 0 Then
                    Throw New Exception("Please set Freight Provision account for vendor account set :" + obj.Account_Set + " , vendor:" + obj.Vendor_Code)
                End If

                FreightProvision = clsERPFuncationality.ChangeGLAccountLocationSegment(FreightProvision, obj.loc_code, True, trans)
                Dim AccCr3() As String = {FreightProvision, KOAmt * dblConvRate}
                ArryLst.Add(AccCr3)

            End If
        End If
        ''richa 6 Feb,2020
        Dim objJE As New clsJEExtraColumns
        If clsCommon.myLen(obj.TapalNo) > 0 Or clsCommon.myLen(obj.DateAndTime) > 0 Then
            objJE.TapalNo = clsCommon.myCstr(obj.TapalNo)
            If clsCommon.myLen(obj.DateAndTime) > 0 Then
                objJE.DateAndTime = obj.DateAndTime
            End If
        End If
        If (clsCommon.CompairString(obj.Document_Type, "I") = CompairStringResult.Equal) Then
            clsJournalMaster.FunGrnlEntryWithTrans(strPrefixTransType, "", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, strPostDate, pjvNOVochdesc, "AP-IN", "AP Invoice", obj.Document_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, Nothing, Nothing, Nothing, coll, objJE)
        ElseIf (clsCommon.CompairString(obj.Document_Type, "C") = CompairStringResult.Equal) Then
            clsJournalMaster.FunGrnlEntryWithTrans(strPrefixTransType, "", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, strPostDate, pjvNOVochdesc, "AP-CN", "AP Invoice", obj.Document_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, Nothing, Nothing, Nothing, coll, objJE)
        ElseIf (clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal) Then
            Dim ArryLstNew As ArrayList = New ArrayList()
            For Each Str() As String In ArryLst
                Dim strNew() As String = {Str(0), -1 * Str(1), If(Str.Length >= 3, Str(2), ""), If(Str.Length >= 4, Str(3), ""), If(Str.Length >= 5, Str(4), ""), If(Str.Length >= 6, Str(5), "")}
                ArryLstNew.Add(strNew)
            Next
            clsJournalMaster.FunGrnlEntryWithTrans(strPrefixTransType, "", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, strPostDate, pjvNOVochdesc, "AP-DN", "AP Invoice", obj.Document_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLstNew, Nothing, Nothing, Nothing, coll, objJE)
        ElseIf (clsCommon.CompairString(obj.Document_Type, "P") = CompairStringResult.Equal) Then
            clsJournalMaster.FunGrnlEntryWithTrans(strPrefixTransType, "", obj.loc_code, True, isForUnpostedTransaction, strVoucherNo, trans, strPostDate, pjvNOVochdesc, "AP-IN", "AP Invoice", obj.Document_No, obj.Description, "V", obj.Vendor_Code, obj.Vendor_Name, objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst, Nothing, Nothing, Nothing, coll, objJE)
        Else
            Throw New Exception("Invoice Type not found to Post")
        End If
        Return True
    End Function

    Private Shared Function CreateAdditionalCostEntry(ByVal strRefDocNo As String, ByVal obj As clsVedorInvoiceHead, ByVal trans As SqlTransaction, Optional ByVal strVoucherNoForRecreatedOnly As String = Nothing) As Boolean
        Try
            ''''''''' priti starts here when REFerence Document No is not blank
            If strRefDocNo <> "" Then
                Dim strQry As String
                Dim strRecoverable, strInvAcc, strAdjAcc As String
                Dim dt As New DataTable
                Dim decTaxAmt As Double = 0
                Dim Totalamt As Double
                Dim totalunit_cost_tax As Double = 0
                Dim AddCost As Double = 0
                Dim TotalAddCost As Double = 0
                Dim AddCostOld As Double = 0
                Dim TotalAddCostOld As Double = 0
                Dim Landed_Cost As Double = 0
                Dim TotLanded_Cost As Double = 0
                Dim Landed_CostOld As Double = 0
                Dim TotLanded_CostOld As Double = 0
                Dim intLineNo As Integer = 0
                Dim ArryLst1 As ArrayList = New ArrayList()

                For Each objTR As clsVedorInvoiceDetail In obj.Arr
                    decTaxAmt = 0
                    If objTR.TAX1 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX1 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'obj = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + obj.TAX1_Amt
                                decTaxAmt = objTR.Amount_less_Discount + decTaxAmt
                            Else
                                decTaxAmt = objTR.Amount_less_Discount + decTaxAmt
                            End If
                        End If
                    Else
                        decTaxAmt = objTR.Amount_less_Discount + decTaxAmt
                    End If
                    If objTR.TAX2 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX2 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'obj = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX2_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX3 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX3 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX3_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX4 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX4 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX4_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX5 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX5 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX5_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX6 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX6 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX6_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX7 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX7 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX7_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX8 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX8 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX8_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX9 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX9 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX9_Amt
                            End If
                        End If
                    End If
                    If objTR.TAX10 <> "" Then
                        strQry = "select Tax_Recoverable from TSPL_TAX_MASTER  where Tax_Code='" & objTR.TAX10 & "'"
                        dt = clsDBFuncationality.GetDataTable(strQry, trans)
                        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                            'objTR = New clsVedorInvoiceHead()
                            strRecoverable = clsCommon.myCstr(dt.Rows(0)("Tax_Recoverable"))
                            If strRecoverable = "N" Then
                                decTaxAmt = decTaxAmt + objTR.TAX10_Amt
                            End If
                        End If
                    End If

                    Totalamt = decTaxAmt
                    'Dim strGlAcct As String = obj.Arr(0).GL_Account_Code

                    Dim strSql As String = "select Item_Code,Item_Desc,Location from TSPL_SRN_DETAIL where SRN_No='" & strRefDocNo & "' and Line_No=1"
                    dt = clsDBFuncationality.GetDataTable(strSql, trans)
                    If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                        obj.Arr = New List(Of clsVedorInvoiceDetail)
                        Dim objTr1 As clsVedorInvoiceDetail
                        For Each dr As DataRow In dt.Rows
                            intLineNo = intLineNo + 1
                            objTr1 = New clsVedorInvoiceDetail
                            objTr1.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                            objTr1.Item_Desc = clsCommon.myCstr(dr("Item_Desc"))
                            objTr1.Location = clsCommon.myCstr(dr("Location"))


                            '==========================Added by preeti Gupta=====================
                            Dim strsegment As String = obj.loc_code 'clsDBFuncationality.getSingleValue("select Loc_Segment_Code  from TSPL_LOCATION_MASTER  where Location_Code='" + objTr1.Location + "'", trans)
                            Dim dt2 As DataTable = clsDBFuncationality.GetDataTable("select  replace (Inv_Control_Account,RIGHT(Inv_Control_Account,3),'" + strsegment + "')AS Inv_Control_Account ,'" & objTR.GL_Account_Code & "' as Adjustment_Account from TSPL_PURCHASE_ACCOUNTS where Purchase_Class_Code in " &
                               "(select Purchase_Class_Code from TSPL_ITEM_MASTER where Item_Code='" + objTr1.Item_Code + "')", trans)
                            If dt2 Is Nothing AndAlso dt2.Rows.Count <= 0 Then
                                Throw New Exception("Error")
                            End If

                            For Each DAtarow As DataRow In dt2.Rows
                                strInvAcc = DAtarow(0).ToString()
                                strAdjAcc = DAtarow(1).ToString()
                                Dim Acc1() As String = {strInvAcc, Totalamt}
                                Dim Acc2() As String = {strAdjAcc, Totalamt * -1}
                                ArryLst1.Add(Acc1)
                                ArryLst1.Add(Acc2)
                            Next
                        Next
                    End If
                Next

                '==== Added by preeti
                If ArryLst1 IsNot Nothing AndAlso ArryLst1.Count > 0 Then
                    If strVoucherNoForRecreatedOnly IsNot Nothing AndAlso clsCommon.myLen(strVoucherNoForRecreatedOnly) > 0 Then
                        clsJournalMaster.FunGrnlEntryWithTrans(obj.loc_code, True, strVoucherNoForRecreatedOnly, trans, clsCommon.GETSERVERDATE(trans), "Additional Cost against SRN No '" & strRefDocNo & "' for Vendor '" & obj.Vendor_Name & "'", "AP-IN", "I/C Adjustments", obj.Document_No, obj.Description, "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst1, strRefDocNo)
                    Else
                        clsJournalMaster.FunGrnlEntryWithTrans(obj.loc_code, True, trans, clsCommon.GETSERVERDATE(trans), "Additional Cost against SRN No '" & strRefDocNo & "' for Vendor '" & obj.Vendor_Name & "'", "AP-IN", "I/C Adjustments", obj.Document_No, obj.Description, "O", "", "", objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode, ArryLst1, strRefDocNo)
                    End If

                End If
            End If
            ''''''''' priti ends here
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function

    Private Shared Function GetTaxAmt(ByVal objPIDetail As clsVedorInvoiceDetail, ByVal tans As SqlTransaction) As Double
        Dim dblTotalTax As Double = 0
        Dim isTaxRecoverable As Boolean = False
        If objPIDetail.TAX1_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX1, tans) Then
            dblTotalTax += objPIDetail.TAX1_Amt
        End If
        If objPIDetail.TAX2_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX2, tans) Then
            dblTotalTax += objPIDetail.TAX2_Amt
        End If
        If objPIDetail.TAX3_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX3, tans) Then
            dblTotalTax += objPIDetail.TAX3_Amt
        End If
        If objPIDetail.TAX4_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX4, tans) Then
            dblTotalTax += objPIDetail.TAX4_Amt
        End If
        If objPIDetail.TAX5_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX5, tans) Then
            dblTotalTax += objPIDetail.TAX5_Amt
        End If
        If objPIDetail.TAX6_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX6, tans) Then
            dblTotalTax += objPIDetail.TAX6_Amt
        End If
        If objPIDetail.TAX7_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX7, tans) Then
            dblTotalTax += objPIDetail.TAX7_Amt
        End If
        If objPIDetail.TAX8_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX8, tans) Then
            dblTotalTax += objPIDetail.TAX8_Amt
        End If
        If objPIDetail.TAX9_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX9, tans) Then
            dblTotalTax += objPIDetail.TAX9_Amt
        End If
        If objPIDetail.TAX10_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX10, tans) Then
            dblTotalTax += objPIDetail.TAX10_Amt
        End If
        Return dblTotalTax
    End Function

    Public Function GetTaxAmtNonShared(ByVal objPIDetail As clsVedorInvoiceDetail, ByVal tans As SqlTransaction) As Double
        Dim dblTotalTax As Double = 0
        Dim isTaxRecoverable As Boolean = False
        If objPIDetail.TAX1_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX1, tans) Then
            dblTotalTax += objPIDetail.TAX1_Amt
        End If
        If objPIDetail.TAX2_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX2, tans) Then
            dblTotalTax += objPIDetail.TAX2_Amt
        End If
        If objPIDetail.TAX3_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX3, tans) Then
            dblTotalTax += objPIDetail.TAX3_Amt
        End If
        If objPIDetail.TAX4_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX4, tans) Then
            dblTotalTax += objPIDetail.TAX4_Amt
        End If
        If objPIDetail.TAX5_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX5, tans) Then
            dblTotalTax += objPIDetail.TAX5_Amt
        End If
        If objPIDetail.TAX6_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX6, tans) Then
            dblTotalTax += objPIDetail.TAX6_Amt
        End If
        If objPIDetail.TAX7_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX7, tans) Then
            dblTotalTax += objPIDetail.TAX7_Amt
        End If
        If objPIDetail.TAX8_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX8, tans) Then
            dblTotalTax += objPIDetail.TAX8_Amt
        End If
        If objPIDetail.TAX9_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX9, tans) Then
            dblTotalTax += objPIDetail.TAX9_Amt
        End If
        If objPIDetail.TAX10_Amt <> 0 AndAlso Not clsTaxMaster.ISTaxRecoverableAC(objPIDetail.TAX10, tans) Then
            dblTotalTax += objPIDetail.TAX10_Amt
        End If
        Return dblTotalTax
    End Function
    ''added by richa agarwal against ticket no TEC/01/04/19-000464
    Public Shared Function DeleteData(ByVal strCode As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            DeleteData(strCode, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function DeleteData(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean

        Dim isSaved As Boolean = False
        If (clsCommon.myLen(strDocNo) <= 0) Then
            Throw New Exception("Document No not found to Delete")
        End If
        Dim obj As clsVedorInvoiceHead = clsVedorInvoiceHead.GetData(strDocNo, "", trans)
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Document_No) > 0) Then
            Try
                clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModulePayable, clsUserMgtCode.mbtnAPInvoiceEntry, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)
                clsERPFuncationality.ValidateLocationSegment(objCommonVar.CurrentCompanyCode, clsUserMgtCode.ModuleTDS, clsUserMgtCode.mbtnAPInvoiceEntryTDS, obj.loc_code, clsCommon.myCDate(obj.Invoice_Entry_Date), trans)

                If (clsCommon.myLen(obj.Posting_Date) > 0) Then
                    Throw New Exception("Already Post on :" + obj.Posting_Date)
                End If
                clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", trans)
                'clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", trans)

                clsCommonFunctionality.SaveHistoryData(EnumSaveType.History, objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", "TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL", "AP_Invoice_No", "TSPL_AP_Invoice_Asset_EMI_Details", "AP_Invoice_No", "", "", "", "", "", "", trans)
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_REMITTANCE", "Document_No", "TSPL_AP_INVOICE_ADVANCE_INTEREST", "AP_Invoice_No", "TSPL_PROVISION_ENTRY_KNOCKOFF", "AP_Invoice_No", trans)

                Dim qry As String = "delete from TSPL_VENDOR_INVOICE_DETAIL where Document_No='" + strDocNo + "'"
                isSaved = clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "Delete from TSPL_PROVISION_ENTRY_KNOCKOFF where AP_Invoice_No='" + obj.Document_No + "'"
                clsDBFuncationality.ExecuteNonQuery(qry, trans)
                ''changes done by richa agarwal on 11 Aug,2016 to delete journal entry on the behalf of ap invoice type
                '' changed by Panch Raj against Ticket No: BM00000008161

                '' delete journal details
                ' qry = "delete from TSPL_JOURNAL_DETAILS where Journal_No in (select Journal_No from TSPL_JOURNAL_MASTER where Source_Code in ('AP-IN','AP-CN','AP-DN') and Source_Doc_No='" + strDocNo + "')"
                'qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code in ('AP-IN','AP-CN','AP-DN') and Source_Doc_No='" + strDocNo + "')"
                'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                'qry = "delete from TSPL_JOURNAL_MASTER   where Source_Code in ('AP-IN','AP-CN','AP-DN') and Source_Doc_No='" + strDocNo + "'"
                'isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                'If clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "D") = CompairStringResult.Equal Then
                '    qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code IN ('AP-DN','GL-JE') and Source_Doc_No='" + strDocNo + "')"
                '    isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                '    qry = "delete from TSPL_JOURNAL_MASTER   where Source_Code  IN ('AP-DN','GL-JE') and Source_Doc_No='" + strDocNo + "'"
                '    isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                'ElseIf clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "C") = CompairStringResult.Equal Then
                '    qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='AP-CN' and Source_Doc_No='" + strDocNo + "')"
                '    isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                '    qry = "delete from TSPL_JOURNAL_MASTER   where Source_Code='AP-CN' and Source_Doc_No='" + strDocNo + "'"
                '    isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                'ElseIf clsCommon.CompairString(clsCommon.myCstr(obj.Document_Type), "I") = CompairStringResult.Equal Then
                '    qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No in (select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code='AP-IN' and Source_Doc_No='" + strDocNo + "')"
                '    isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)

                '    qry = "delete from TSPL_JOURNAL_MASTER   where Source_Code='AP-IN' and Source_Doc_No='" + strDocNo + "'"
                '    isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                'End If

                '-----Delete Main Journal ENtry----
                Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where (TSPL_JOURNAL_MASTER.Source_Code like 'AP%' or TSPL_JOURNAL_MASTER.Source_Code='GL-JE' ) and Source_Doc_No='" + strDocNo + "'", trans)
                If clsCommon.myLen(VoucherNo) > 0 Then
                    clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                    qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If

                '' TO DELETE DATA FROM JOURNAL MASTER OPENING TABLE ''richa TEC/02/11/18-000360 
                Dim VoucherNoOP As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER_OP where (TSPL_JOURNAL_MASTER_OP.Source_Code like 'AP%' or TSPL_JOURNAL_MASTER_OP.Source_Code='GL-JE' ) AND  TSPL_JOURNAL_MASTER_OP.Source_Doc_No='" + strDocNo + "'", trans)
                If clsCommon.myLen(VoucherNoOP) > 0 Then
                    qry = "delete from TSPL_JOURNAL_DETAILS_OP where Voucher_No ='" + VoucherNoOP + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                    qry = "delete from TSPL_JOURNAL_MASTER_OP where Voucher_No ='" + VoucherNoOP + "'"
                    clsDBFuncationality.ExecuteNonQuery(qry, trans)
                End If

                'Ticket No-TEC/06/09/19-001003,Save Deleted data ,sanjay
                ' clsCommonFunctionality.SaveDeletedData(objCommonVar.CurrentUserCode, strDocNo, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", trans)

                qry = "delete from TSPL_REMITTANCE where Document_No='" + strDocNo + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                qry = "delete from TSPL_VENDOR_INVOICE_HEAD where Document_No='" + strDocNo + "'"
                isSaved = isSaved AndAlso clsDBFuncationality.ExecuteNonQuery(qry, trans)
                isSaved = isSaved AndAlso clsCustomFieldValues.DeleteData(obj.Form_ID, strDocNo, trans)
                If obj.is_For_Provision = 1 Then
                    isSaved = isSaved AndAlso clsProvisionEntry.SaveData(obj.Document_No, Nothing, trans)
                End If
                'If (isSaved) Then
                '    trans.Commit()
                'Else
                '    trans.Rollback()
                'End If
            Catch ex As Exception
                'trans.Rollback()
                Throw New Exception(ex.Message)
            End Try
        End If
        Return isSaved
    End Function
    ''richa agarwal 08 Apr,2019 TEC/01/04/19-000464
    Public Shared Function ReverseAndUnpost(ByVal strCode As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            ReverseAndUnpost(strCode, trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function
    Public Shared Function ReverseAndUnpost(ByVal strDocNo As String, ByVal trans As SqlTransaction) As Boolean
        'Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim Qry As String = "Select Posting_Date from TSPL_VENDOR_INVOICE_HEAD WHERE Document_No='" + strDocNo + "'"
            If clsCommon.myLen(clsDBFuncationality.getSingleValue(Qry, trans)) <= 0 Then
                Throw New Exception("Transaction status should be posted for reverse and unpost")
            End If

            Dim obj As clsVedorInvoiceHead = clsVedorInvoiceHead.GetData(strDocNo, "", trans)
            If (obj Is Nothing OrElse clsCommon.myLen(obj.Document_No) <= 0) Then
                Throw New Exception("No Data found to Reverse And UnPost")
            End If
            'clsCommonFunctionality.SaveHistoryData(EnumSaveType.History, objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", "TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL", "AP_Invoice_No", "TSPL_AP_Invoice_Asset_EMI_Details", "AP_Invoice_No", "", "", "", "", "", "", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_REMITTANCE", "Document_No", "TSPL_AP_INVOICE_ADVANCE_INTEREST", "AP_Invoice_No", "TSPL_PROVISION_ENTRY_KNOCKOFF", "AP_Invoice_No", trans)

            '' Get Payment Entry Against AP Invoice

            Qry = "select Payment_No from TSPL_PAYMENT_DETAIL  where Document_No in ('" + obj.Document_No + "')"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Qry = "AP-Invoice " + obj.Document_No + " is used in following Payment -"
                For Each drAP As DataRow In dt.Rows
                    Qry += Environment.NewLine + clsCommon.myCstr(drAP("Payment_No"))
                Next
                Throw New Exception(Qry)
            End If

            ''richa agarwal 26 Nov,2019 ERO/19/11/19-001120
            If clsCommon.myLen(obj.Against_Salary_Generation_Code) > 0 Then
                Throw New Exception("Transaction  cannot be Reversed because it has created against salary generation")
            End If

            ''to check  document used into payment process 
            Qry = "select Doc_No from TSPL_PAYMENT_PROCESS_DEDUCTION where AP_Invoice_No ='" & strDocNo & "'"
            dt = clsDBFuncationality.GetDataTable(Qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Qry = "Current document is used in following Payment Process -"
                For Each dr As DataRow In dt.Rows
                    Qry += Environment.NewLine + clsCommon.myCstr(dr("Doc_No"))
                Next
                Throw New Exception(Qry)
            End If


            '-----Delete Main Journal ENtry----
            Dim VoucherNo As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where (TSPL_JOURNAL_MASTER.Source_Code like 'AP%' or TSPL_JOURNAL_MASTER.Source_Code='GL-JE' ) and Source_Doc_No='" + obj.Document_No + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, VoucherNo, "TSPL_JOURNAL_MASTER", "Voucher_No", "TSPL_JOURNAL_DETAILS", "Voucher_No", trans)
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            '' TO DELETE DATA FROM JOURNAL MASTER OPENING TABLE ''richa TEC/02/11/18-000360 
            Dim VoucherNoOP As String = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER_OP where (TSPL_JOURNAL_MASTER_OP.Source_Code like 'AP%' or TSPL_JOURNAL_MASTER_OP.Source_Code='GL-JE' ) AND  TSPL_JOURNAL_MASTER_OP.Source_Doc_No='" + obj.Document_No + "'", trans)
            If clsCommon.myLen(VoucherNoOP) > 0 Then
                Qry = "delete from TSPL_JOURNAL_DETAILS_OP where Voucher_No ='" + VoucherNoOP + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER_OP where Voucher_No ='" + VoucherNoOP + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            '-----Delete Additional Cost ENtry----
            VoucherNo = clsDBFuncationality.getSingleValue("select Voucher_No from TSPL_JOURNAL_MASTER where Source_Code like 'AP%' and Source_Doc_No='" + obj.RefDocNo + "'", trans)
            If clsCommon.myLen(VoucherNo) > 0 Then
                Qry = "delete from TSPL_JOURNAL_DETAILS where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
                Qry = "delete from TSPL_JOURNAL_MASTER where Voucher_No ='" + VoucherNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            Qry = "update TSPL_REMITTANCE set Remit_TDS=NULL where Document_No='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)

            If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.Against_PurchaseReturn_No) > 0 Then
                Qry = "update TSPL_VENDOR_INVOICE_HEAD set Balance_Amt=Balance_Amt+ " + clsCommon.myCstr(obj.Document_Total - obj.TDS_Actual_Amount) + " where Against_POInvoice_No  in ( select Against_PI from TSPL_PR_HEAD where PR_No='" + obj.Against_PurchaseReturn_No + "')"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            End If

            If clsCommon.CompairString(obj.RefDocType, "AP") = CompairStringResult.Equal AndAlso clsCommon.myLen(obj.RefDocNo) > 0 Then
                Dim strOpearateor = "-"
                If clsCommon.CompairString(obj.Document_Type, "D") = CompairStringResult.Equal Then
                    strOpearateor = "+"
                End If
                Dim dblRemAmt As Double = 0
                If (obj.RemittanceObject IsNot Nothing) Then
                    dblRemAmt = obj.RemittanceObject.Actual_Total_TDS
                End If

                Qry = "update TSPL_VENDOR_INVOICE_HEAD set Balance_Amt=Balance_Amt" + strOpearateor + " " + clsCommon.myCstr(obj.Document_Total - dblRemAmt) + " where Document_No='" + obj.RefDocNo + "'"
                clsDBFuncationality.ExecuteNonQuery(Qry, trans)

                Qry = "select Balance_Amt  from TSPL_VENDOR_INVOICE_HEAD where  Document_No='" + obj.RefDocNo + "'"
                If clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry, trans)) < 0 Then
                    Throw New Exception("Balance is Going to Negative For AP Invoice No : " + obj.RefDocNo)
                End If

            End If

            Qry = "Update TSPL_VENDOR_INVOICE_HEAD set Posting_Date=NULL, Modify_By='" + objCommonVar.CurrentUserCode + "' where Document_No='" + strDocNo + "'"
            clsDBFuncationality.ExecuteNonQuery(Qry, trans)
            'clsCommonFunctionality.SaveHistoryData(EnumSaveType.History, objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", "TSPL_VENDOR_INVOICE_DETAIL", "Document_No", "TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL", "AP_Invoice_No", "TSPL_AP_Invoice_Asset_EMI_Details", "AP_Invoice_No", "", "", "", "", "", "", trans)
            'clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_REMITTANCE", "Document_No", "TSPL_AP_INVOICE_ADVANCE_INTEREST", "AP_Invoice_No", "TSPL_PROVISION_ENTRY_KNOCKOFF", "AP_Invoice_No", trans)
            clsCommonFunctionality.SaveHistoryData(objCommonVar.CurrentUserCode, obj.Document_No, "TSPL_VENDOR_INVOICE_HEAD", "Document_No", trans)

            'trans.Commit()
            Return True
        Catch ex As Exception
            'trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
    End Function
    Public Shared Function UpdateAfterPost(obj As clsVedorInvoiceHead) As Boolean
        '' PanchRaj
        Dim isSaved As Boolean = True
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            For Each objTr As clsVedorInvoiceDetail In obj.Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "Item_Type", objTr.Item_Type, True)
                clsCommon.AddColumnsForChange(coll, "Asset_Code", objTr.Asset_Code, True)
                isSaved = isSaved AndAlso clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_INVOICE_DETAIL", OMInsertOrUpdate.Update, "TSPL_VENDOR_INVOICE_DETAIL.Document_No='" & objTr.Document_No & "' and Detail_Line_No='" & objTr.Detail_Line_No & "'", trans)
                trans.Commit()
            Next
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try
        Return True

    End Function
    Public Shared Function GetWorkOrderBalanceAmount(ByVal Doc_Code As String, ByVal strAPInvoice As String, ByVal trans As SqlTransaction) As Decimal
        Dim qry As String = GetWorkOrderBalanceAmountBaseQry(Doc_Code, strAPInvoice)
        'qry = " SELECT SUM(PO_Total_Amt) AS Balance_WO_Amt from (" & _
        '      " select Vendor_Code as [Vendor code],PO_Total_Amt " & _
        '      " from TSPL_PURCHASE_ORDER_HEAD WHERE PurchaseOrder_No='" & Doc_Code & "' " & _
        '      " UNION ALL " & _
        '      " SELECT Vendor_Code,(CASE WHEN Document_Type IN ('I','C') THEN -Document_Total WHEN Document_Type='D' THEN Document_Total ELSE 0 END) AS Document_Total " & _
        '      " FROM TSPL_VENDOR_INVOICE_HEAD WHERE RefDocType='WO' AND RefDocNo='" & Doc_Code & "'" & _
        '      " ) as Final"
        qry = "select sum(Balance_WO_Amt) as Balance_WO_Amt from (" & qry & ") as Final"
        Dim Bal As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
        Return Bal
    End Function
    Public Shared Function GetWorkOrderBalanceAmountBaseQry(ByVal Doc_Code As String, ByVal strAPInvoice As String) As String
        Dim qry As String = ""
        qry = " SELECT PurchaseOrder_No,SUM(PO_Total_Amt) AS Balance_WO_Amt from (" &
              " select PurchaseOrder_No,Vendor_Code as [Vendor code],PO_Total_Amt " &
              " from TSPL_PURCHASE_ORDER_HEAD where 2=2 and TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Type ='J' " & If(clsCommon.myLen(Doc_Code) > 0, "and TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No='" & Doc_Code & "'", "") & "" &
              " UNION ALL " &
              " SELECT RefDocNo as PurchaseOrder_No,Vendor_Code,(CASE WHEN Document_Type IN ('I','C') THEN -Document_Total WHEN Document_Type='D' THEN Document_Total ELSE 0 END) AS Document_Total " &
              " FROM TSPL_VENDOR_INVOICE_HEAD WHERE RefDocType='WO' and TSPL_VENDOR_INVOICE_HEAD.Document_No<>'" & clsCommon.myCstr(strAPInvoice) & "' " &
              " ) as Final where 2=2 "
        If clsCommon.myLen(Doc_Code) > 0 Then
            qry = qry & " and purchaseOrder_No='" & Doc_Code & "'"
        End If
        qry = qry & " group by PurchaseOrder_No"

        Return qry
    End Function


    Public Function GetBalanceAmtOfAPInvoice(ByVal strAPInvoiceNo As String, ByVal trans As SqlTransaction) As Decimal
        Dim qry As String = ""
        Return clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qry, trans))
    End Function

    Public Shared Function GetMainVSPMilkAPInvoiceNo(ByVal DocDate As Date, ByVal VSPCode As String, ByVal tran As SqlTransaction) As String
        Return clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Document_No from TSPL_VENDOR_INVOICE_HEAD where Vendor_Code='" + VSPCode + "' and Against_MillkPurchaseInvoice_No is not null and Posting_Date='" + clsCommon.GetPrintDate(DocDate, "dd/MMM/yyyy") + "'", tran))
    End Function

    Public Shared Function GetRefundBaseQuery(Vendorcode As String, currDocNo As String, whrclas As String) As String
        Dim qry As String = "select Document_No as DocumentCode,max(convert(varchar,Posting_Date,103)) as DocumentDate,sum(Document_Total * case when RI=1 then 1 else 0 end) as DocumentAmount,sum(Document_Total * case when RI=-1 then 1 else 0 end) as AdjustAmount,sum(Document_Total*RI) as BalanceAmount from (
select Document_No,Posting_Date,TSPL_VENDOR_INVOICE_HEAD.Document_Total ,1 as RI,1 as Chk  from TSPL_VENDOR_INVOICE_HEAD 
where  Document_Type='D' and Is_Security=1 and LEN(Posting_Date)>0 and Vendor_Code='" + Vendorcode + "' "
        If clsCommon.myLen(whrclas) > 0 Then
            qry += " and " + whrclas
        End If
        qry += " union all
select TSPL_VENDOR_INVOICE_SECURITY_REFUND.Against_Security as Document_No,null as Posting_Date,Amount as Document_Total,-1 as RI,0 as Chk 
from TSPL_VENDOR_INVOICE_SECURITY_REFUND
left outer join TSPL_VENDOR_INVOICE_HEAD on TSPL_VENDOR_INVOICE_HEAD.Document_No=TSPL_VENDOR_INVOICE_SECURITY_REFUND.Against_Security
where TSPL_VENDOR_INVOICE_HEAD.Vendor_Code='" + Vendorcode + "' and TSPL_VENDOR_INVOICE_SECURITY_REFUND.AP_Invoice_No not in ('" + currDocNo + "')
) xx group by Document_No having sum(Document_Total*RI)>0 and sum(Chk)>0  order by max(Posting_Date)"
        Return qry
    End Function
End Class

Public Class clsVedorInvoiceDetail
#Region "Variables"
    Public chrgcatcode As String = Nothing
    Public chrgcatdesc As String = Nothing
    Public chrgcatvalue As String = Nothing
    Public chritemcode As String = Nothing
    Public chritemdesc As String = Nothing

    Public SAC_Code As String = String.Empty
    Public SAC_Name As String = String.Empty
    '' priti starts here
    Public Document_No As String = Nothing
    Public Unit_code As String = Nothing
    Public Location As String = Nothing
    Public Detail_Line_No As Integer = 0
    Public unit_cost_tax As Double = 0
    Public Item_Code As String = Nothing
    Public Item_Desc As String = Nothing
    Public MRP As Double = 0
    Public Batch_No As String = Nothing
    Public MFG_Date As String = Nothing
    Public Expiry_Date As String = Nothing
    Public Adjustment_Account As String = Nothing
    Public Description As String = Nothing


    '' priti ends here
    Public Deduction_Code As String = Nothing
    Public Deduction_Name As String = Nothing ''Not a table column
    Public Deduction_Section As String = Nothing ''Not a table column


    Public GL_Account_Code As String = Nothing
    Public GL_Account_Desc As String = Nothing
    Public Amount As Double = 0
    Public Discount_Per As Double = 0
    Public Discount As Double = 0
    Public Amount_less_Discount As Double
    Public Taxable_Amount_Per As Decimal
    Public Taxable_Amount As Decimal
    Public TAX1 As String = Nothing
    Public TAX1_Rate As Double = 0
    Public TAX1_Amt As Double = 0
    Public TAX2 As String = Nothing
    Public TAX2_Rate As Double = 0
    Public TAX2_Amt As Double = 0
    Public TAX3 As String = Nothing
    Public TAX3_Rate As Double = 0
    Public TAX3_Amt As Double = 0
    Public TAX4 As String = Nothing
    Public TAX4_Rate As Double = 0
    Public TAX4_Amt As Double = 0
    Public TAX5 As String = Nothing
    Public TAX5_Rate As Double = 0
    Public TAX5_Amt As Double = 0
    Public TAX6 As String = Nothing
    Public TAX6_Rate As Double = 0
    Public TAX6_Amt As Double = 0
    Public TAX7 As String = Nothing
    Public TAX7_Rate As Double = 0
    Public TAX7_Amt As Double = 0
    Public TAX8 As String = Nothing
    Public TAX8_Rate As Double = 0
    Public TAX8_Amt As Double = 0
    Public TAX9 As String = Nothing
    Public TAX9_Rate As Double = 0
    Public TAX9_Amt As Double = 0
    Public TAX10 As String = Nothing
    Public TAX10_Rate As Double = 0
    Public TAX10_Amt As Double = 0
    Public Total_Tax As Double = 0
    Public Total_Amount As Double = 0
    Public Remarks As String = Nothing
    Public Comments As String = Nothing
    Public AddChargeCode As String = Nothing
    Public AddChargeDesc As String = Nothing
    Public is_Unclaimed_Tax As Boolean = False


    Public Invoice_No As String = Nothing
    Public Invoice_Type As String = Nothing
    Public Landed_Amount As Decimal = 0

    Public Item_Rate As Decimal = 0
    Public Abatement_Per As Double = 0
    Public Abatement_Amt As Double = 0
    Public Reverse_Charge_Per As Double = 0
    Public Reverse_Charge_Amount As Double = 0

    Public TAX1_Base_Amt As Double = 0
    Public TAX2_Base_Amt As Double = 0
    Public TAX3_Base_Amt As Double = 0
    Public TAX4_Base_Amt As Double = 0
    Public TAX5_Base_Amt As Double = 0
    Public TAX6_Base_Amt As Double = 0
    Public TAX7_Base_Amt As Double = 0
    Public TAX8_Base_Amt As Double = 0
    Public TAX9_Base_Amt As Double = 0
    Public TAX10_Base_Amt As Double = 0
    Public DeductionCode As String = ""
    Public DeductionDesc As String = ""
    ''richa agarwal 
    Public AgainstPayment_No As String = Nothing
    Public Payment_Amount As Double = 0
    Public TDS_Per As Double = 0
    '' shivani tyagi
    Public Hirerachy_Code As String = String.Empty
    Public Cost_Centre_Code As String = String.Empty

    Public Hirerachy_Code1 As String = String.Empty
    Public Hirerachy_Code2 As String = String.Empty
    Public Hirerachy_Code3 As String = String.Empty
    Public Hirerachy_Code4 As String = String.Empty
    Public Against_Milk_SRN_Commission_No As String = String.Empty
    '' PanchRaj
    Public Item_Type As String = ""
    Public Asset_Code As String = ""
    Public Asset_Desc As String = ""
    Public PK_Id As Integer
    Public arrSTDed As List(Of clsAPSecondaryTranporterDeductionDetail) = Nothing

    Public DCS_Addition_Deduction As String = ""
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal Arr As List(Of clsVedorInvoiceDetail), ByVal trans As SqlTransaction) As Boolean
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            For Each obj As clsVedorInvoiceDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "item_code", obj.chritemcode)
                clsCommon.AddColumnsForChange(coll, "item_desc", obj.chritemdesc)
                clsCommon.AddColumnsForChange(coll, "charge_cat_code", obj.chrgcatcode)
                clsCommon.AddColumnsForChange(coll, "charge_cat_desc", obj.chrgcatdesc)
                clsCommon.AddColumnsForChange(coll, "charge_cat_charges", obj.chrgcatvalue)
                clsCommon.AddColumnsForChange(coll, "DeductionCode", obj.DeductionCode)
                clsCommon.AddColumnsForChange(coll, "Deduction_Desc", obj.DeductionDesc)
                clsCommon.AddColumnsForChange(coll, "Document_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Detail_Line_No", obj.Detail_Line_No)
                clsCommon.AddColumnsForChange(coll, "Deduction_Code", obj.Deduction_Code, True)
                clsCommon.AddColumnsForChange(coll, "DCS_Addition_Deduction", obj.DCS_Addition_Deduction, True)
                clsCommon.AddColumnsForChange(coll, "AgainstPayment_No", obj.AgainstPayment_No, True)
                clsCommon.AddColumnsForChange(coll, "Payment_Amount", obj.Payment_Amount)
                clsCommon.AddColumnsForChange(coll, "TDS_Per", obj.TDS_Per)
                clsCommon.AddColumnsForChange(coll, "SAC_Code", obj.SAC_Code, True)
                clsCommon.AddColumnsForChange(coll, "GL_Account_Code", obj.GL_Account_Code)
                clsCommon.AddColumnsForChange(coll, "GL_Account_Desc", obj.GL_Account_Desc)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommon.AddColumnsForChange(coll, "Discount_Per", obj.Discount_Per)
                clsCommon.AddColumnsForChange(coll, "Discount", obj.Discount)
                clsCommon.AddColumnsForChange(coll, "Amount_less_Discount", obj.Amount_less_Discount)
                clsCommon.AddColumnsForChange(coll, "Taxable_Amount_Per", obj.Taxable_Amount_Per)
                clsCommon.AddColumnsForChange(coll, "Taxable_Amount", obj.Taxable_Amount)
                clsCommon.AddColumnsForChange(coll, "TAX1", obj.TAX1)
                clsCommon.AddColumnsForChange(coll, "TAX1_Rate", obj.TAX1_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX1_Amt", obj.TAX1_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2", obj.TAX2)
                clsCommon.AddColumnsForChange(coll, "TAX2_Rate", obj.TAX2_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX2_Amt", obj.TAX2_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3", obj.TAX3)
                clsCommon.AddColumnsForChange(coll, "TAX3_Rate", obj.TAX3_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX3_Amt", obj.TAX3_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4", obj.TAX4)
                clsCommon.AddColumnsForChange(coll, "TAX4_Rate", obj.TAX4_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX4_Amt", obj.TAX4_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5", obj.TAX5)
                clsCommon.AddColumnsForChange(coll, "TAX5_Rate", obj.TAX5_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX5_Amt", obj.TAX5_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6", obj.TAX6)
                clsCommon.AddColumnsForChange(coll, "TAX6_Rate", obj.TAX6_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX6_Amt", obj.TAX6_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7", obj.TAX7)
                clsCommon.AddColumnsForChange(coll, "TAX7_Rate", obj.TAX7_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX7_Amt", obj.TAX7_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8", obj.TAX8)
                clsCommon.AddColumnsForChange(coll, "TAX8_Rate", obj.TAX8_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX8_Amt", obj.TAX8_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9", obj.TAX9)
                clsCommon.AddColumnsForChange(coll, "TAX9_Rate", obj.TAX9_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX9_Amt", obj.TAX9_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10", obj.TAX10)
                clsCommon.AddColumnsForChange(coll, "TAX10_Rate", obj.TAX10_Rate)
                clsCommon.AddColumnsForChange(coll, "TAX10_Amt", obj.TAX10_Amt)
                clsCommon.AddColumnsForChange(coll, "Total_Tax", obj.Total_Tax)
                clsCommon.AddColumnsForChange(coll, "Total_Amount", obj.Total_Amount)
                clsCommon.AddColumnsForChange(coll, "Remarks", obj.Remarks)
                clsCommon.AddColumnsForChange(coll, "Comments", obj.Comments)
                clsCommon.AddColumnsForChange(coll, "AddChargeCode", obj.AddChargeCode)
                clsCommon.AddColumnsForChange(coll, "AddChargeDesc", obj.AddChargeDesc)
                clsCommon.AddColumnsForChange(coll, "is_Unclaimed_Tax", IIf(obj.is_Unclaimed_Tax, 1, 0))
                clsCommon.AddColumnsForChange(coll, "Invoice_No", obj.Invoice_No)
                clsCommon.AddColumnsForChange(coll, "Invoice_Type", obj.Invoice_Type)
                clsCommon.AddColumnsForChange(coll, "Landed_Amount", obj.Landed_Amount)
                clsCommon.AddColumnsForChange(coll, "Item_Rate", obj.Item_Rate)
                clsCommon.AddColumnsForChange(coll, "Abatement_Per", obj.Abatement_Per)
                clsCommon.AddColumnsForChange(coll, "Abatement_Amt", obj.Abatement_Amt)
                clsCommon.AddColumnsForChange(coll, "Reverse_Charge_Per", obj.Reverse_Charge_Per)
                clsCommon.AddColumnsForChange(coll, "Reverse_Charge_Amount", obj.Reverse_Charge_Amount)
                clsCommon.AddColumnsForChange(coll, "TAX1_Base_Amt", obj.TAX1_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX2_Base_Amt", obj.TAX2_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX3_Base_Amt", obj.TAX3_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX4_Base_Amt", obj.TAX4_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX5_Base_Amt", obj.TAX5_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX6_Base_Amt", obj.TAX6_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX7_Base_Amt", obj.TAX7_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX8_Base_Amt", obj.TAX8_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX9_Base_Amt", obj.TAX9_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "TAX10_Base_Amt", obj.TAX10_Base_Amt)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code", obj.Hirerachy_Code, True)
                clsCommon.AddColumnsForChange(coll, "Cost_Centre_Code", obj.Cost_Centre_Code, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code1", obj.Hirerachy_Code1, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code2", obj.Hirerachy_Code2, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code3", obj.Hirerachy_Code3, True)
                clsCommon.AddColumnsForChange(coll, "Hirerachy_Code4", obj.Hirerachy_Code4, True)
                clsCommon.AddColumnsForChange(coll, "Against_Milk_SRN_Commission_No", obj.Against_Milk_SRN_Commission_No, True)
                clsCommon.AddColumnsForChange(coll, "Item_Type", obj.Item_Type, True)
                clsCommon.AddColumnsForChange(coll, "Asset_Code", obj.Asset_Code, True)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_VENDOR_INVOICE_DETAIL", OMInsertOrUpdate.Insert, "", trans)

                obj.PK_Id = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select max(PK_Id) as PK_Id from TSPL_VENDOR_INVOICE_DETAIL where Document_No='" + strDocNo + "'", trans))
                clsAPSecondaryTranporterDeductionDetail.SaveData(strDocNo, obj.DeductionCode, obj.PK_Id, obj.arrSTDed, trans)
            Next
        End If
        Return True
    End Function


End Class

Public Class clsAPInvoiceAssetEMIDetails
#Region "Variables"
    Public AP_Invoice_No As String = Nothing
    Public Asset_Issue_No As String = Nothing
    Public Item_Code As String = Nothing
    Public Installment_Amount As Double = 0
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal ArrAssetEMI As List(Of clsAPInvoiceAssetEMIDetails), ByVal trans As SqlTransaction) As Boolean
        If (ArrAssetEMI IsNot Nothing AndAlso ArrAssetEMI.Count > 0) Then
            For Each obj As clsAPInvoiceAssetEMIDetails In ArrAssetEMI
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Asset_Issue_No", obj.Asset_Issue_No)
                clsCommon.AddColumnsForChange(coll, "Item_Code", obj.Item_Code)
                clsCommon.AddColumnsForChange(coll, "Installment_Amount", obj.Installment_Amount)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_AP_Invoice_Asset_EMI_Details", OMInsertOrUpdate.Insert, "", trans)
            Next
        End If
        Return True
    End Function
    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal trans As SqlTransaction) As List(Of clsAPInvoiceAssetEMIDetails)
        Dim qry As String = "Select TSPL_AP_Invoice_Asset_EMI_Details.* from TSPL_AP_Invoice_Asset_EMI_Details where AP_Invoice_No='" + strDocumentNo + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Dim arr As List(Of clsAPInvoiceAssetEMIDetails) = Nothing
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            arr = New List(Of clsAPInvoiceAssetEMIDetails)
            Dim objTr As clsAPInvoiceAssetEMIDetails
            For Each dr As DataRow In dt.Rows
                objTr = New clsAPInvoiceAssetEMIDetails
                objTr.AP_Invoice_No = clsCommon.myCstr(dr("AP_Invoice_No"))
                objTr.Asset_Issue_No = clsCommon.myCstr(dr("Asset_Issue_No"))
                objTr.Item_Code = clsCommon.myCstr(dr("Item_Code"))
                objTr.Installment_Amount = clsCommon.myCdbl(dr("Installment_Amount"))
                arr.Add(objTr)
            Next
        End If
        Return arr
    End Function

    Public Shared Function GetVSPAssetEMIQuery(ByVal strVSPCode As String) As String
        Return GetVSPAssetEMIQuery(strVSPCode, "", "")
    End Function

    Public Shared Function GetVSPAssetEMIQuery(ByVal strAssetIssueCode As String, ByVal strItemCode As String) As String
        Return GetVSPAssetEMIQuery("", strAssetIssueCode, strItemCode)
    End Function

    Private Shared Function GetVSPAssetEMIQuery(ByVal strVSPCode As String, ByVal strAssetIssueCode As String, ByVal strItemCode As String) As String
        Dim qry As String = "select * from (  select Doc_No,Item_Code,case when EMI_Asset_Value<Installment_Amt or ABS(EMI_Asset_Value-Installment_Amt)<=0.09 then EMI_Asset_Value else Installment_Amt end as Installment_Amt from (" + Environment.NewLine +
        " select Doc_No,Item_Code,sum(EMI_Asset_Value*RI) as EMI_Asset_Value,max(Installment_Amt) as Installment_Amt  from ( " + Environment.NewLine +
        " select Doc_No,Item_Code,sum(EMI_Asset_Value*RI) as EMI_Asset_Value,max(EMI_No_Of_Payment_Cycle) as EMI_No_Of_Payment_Cycle,max(Installment_Amt) as Installment_Amt, 1 as RI,1 as Chk  from ( " + Environment.NewLine +
        " select TSPL_VSPAsset_DETAIL.Doc_No ,TSPL_VSPAsset_DETAIL.Item_Code,TSPL_VSPAsset_DETAIL.EMI_Asset_Value,TSPL_VSPAsset_DETAIL.EMI_No_Of_Payment_Cycle,1 as RI,convert(decimal(18,2), (EMI_Asset_Value/EMI_No_Of_Payment_Cycle)) as Installment_Amt " + Environment.NewLine +
        " from TSPL_VSPAsset_DETAIL " + Environment.NewLine +
        " left outer join TSPL_VSPAsset_HEAD on TSPL_VSPAsset_HEAD.Doc_No=TSPL_VSPAsset_DETAIL.Doc_No " + Environment.NewLine +
        " where Doc_Type='Issue' and Status='1' and TSPL_VSPAsset_DETAIL.EMI_Asset_Value>0 and EMI_No_Of_Payment_Cycle>0 " + Environment.NewLine
        If clsCommon.myLen(strVSPCode) > 0 Then
            qry += " and Issue_To='" + strVSPCode + "'"
        End If
        If clsCommon.myLen(strAssetIssueCode) > 0 Then
            qry += " and TSPL_VSPAsset_DETAIL.Doc_No='" + strAssetIssueCode + "' and TSPL_VSPAsset_DETAIL.Item_Code='" + strItemCode + "'"
        End If

        qry += " union all " + Environment.NewLine +
        " select TSPL_VSPAsset_HEAD.Issue_No ,TSPL_VSPAsset_DETAIL.Item_Code,TSPL_VSPAsset_DETAIL.EMI_Asset_Value,0 as  EMI_No_Of_Payment_Cycle,-1 as RI,0 as Installment_Amt " + Environment.NewLine +
        " from TSPL_VSPAsset_DETAIL " + Environment.NewLine +
        " left outer join TSPL_VSPAsset_HEAD on TSPL_VSPAsset_HEAD.Doc_No=TSPL_VSPAsset_DETAIL.Doc_No " + Environment.NewLine +
        " where Doc_Type='Return'  and Status='1' and TSPL_VSPAsset_DETAIL.EMI_Asset_Value>0" + Environment.NewLine
        If clsCommon.myLen(strVSPCode) > 0 Then
            qry += " and Issue_To='" + strVSPCode + "' "
        End If
        If clsCommon.myLen(strAssetIssueCode) > 0 Then
            qry += " and TSPL_VSPAsset_HEAD.Issue_No='" + strAssetIssueCode + "'"
        End If

        qry += " ) xx group by Doc_No,Item_Code " + Environment.NewLine +
        " union all " + Environment.NewLine +
        " select Asset_Issue_No,Item_Code,Installment_Amount as EMI_Asset_Value,0 as EMI_No_Of_Payment_Cycle,0 as Installment_Amt, -1 as RI,0 as Chk from TSPL_AP_Invoice_Asset_EMI_Details  " + Environment.NewLine +
        " )xxx " + Environment.NewLine +
        " group by Doc_No,Item_Code having sum(chk)>0 " + Environment.NewLine +
         " )xx)xxxx where Installment_Amt>0"
        Return qry
    End Function

End Class

Public Class clsAPSecondaryTranporterDeductionDetail
#Region "Variables"
    Public Code As String = Nothing
    Public AP_Invoice_No As String = Nothing
    Public AP_PK_Id As String = Nothing
    Public DC_Challan_No As String = Nothing
    Public Deduction_Code As String = Nothing
    Public SNO As String = Nothing
    Public Amount As Double = 0
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal strDeductionCode As String, ByVal intPK_ID As Integer, ByVal Arr As List(Of clsAPSecondaryTranporterDeductionDetail), ByVal trans As SqlTransaction) As Boolean
        If (Arr IsNot Nothing AndAlso Arr.Count > 0) Then
            Dim Count As Integer = 1
            For Each obj As clsAPSecondaryTranporterDeductionDetail In Arr
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "AP_PK_Id", intPK_ID)
                clsCommon.AddColumnsForChange(coll, "DC_Challan_No", obj.DC_Challan_No)
                clsCommon.AddColumnsForChange(coll, "Deduction_Code", strDeductionCode)
                clsCommon.AddColumnsForChange(coll, "SNO", Count)
                clsCommon.AddColumnsForChange(coll, "Amount", obj.Amount)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL", OMInsertOrUpdate.Insert, "", trans)
                Count += 1
            Next
        End If
        Return True
    End Function

    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal intPK_ID As Integer, ByVal trans As SqlTransaction) As List(Of clsAPSecondaryTranporterDeductionDetail)
        Dim qry As String = "Select * from TSPL_AP_INVOICE_SECONDARY_TRANSPORTER_DEDUTION_DETAIL where AP_Invoice_No='" + strDocumentNo + "' and AP_PK_Id='" + clsCommon.myCstr(intPK_ID) + "' order by SNO"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Dim arr As List(Of clsAPSecondaryTranporterDeductionDetail) = Nothing
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            arr = New List(Of clsAPSecondaryTranporterDeductionDetail)
            Dim objTr As clsAPSecondaryTranporterDeductionDetail
            For Each dr As DataRow In dt.Rows
                objTr = New clsAPSecondaryTranporterDeductionDetail
                objTr.AP_Invoice_No = clsCommon.myCstr(dr("AP_Invoice_No"))
                objTr.AP_PK_Id = clsCommon.myCstr(dr("AP_PK_Id"))
                objTr.DC_Challan_No = clsCommon.myCstr(dr("DC_Challan_No"))
                objTr.Deduction_Code = clsCommon.myCstr(dr("Deduction_Code"))
                objTr.SNO = clsCommon.myCdbl(dr("SNO"))
                objTr.Amount = clsCommon.myCdbl(dr("Amount"))
                arr.Add(objTr)
            Next
        End If
        Return arr
    End Function
End Class


Public Class clsAPInvoiceAdvanceInterest
#Region "Variables"
    Public AP_Invoice_No As String = Nothing
    Public Payment_No As String = Nothing
    Public Interest_Amount As Double = 0
#End Region

    Public Shared Function SaveData(ByVal strDocNo As String, ByVal ArrAdvanceInterest As List(Of clsAPInvoiceAdvanceInterest), ByVal trans As SqlTransaction) As Boolean
        If (ArrAdvanceInterest IsNot Nothing AndAlso ArrAdvanceInterest.Count > 0) Then
            For Each obj As clsAPInvoiceAdvanceInterest In ArrAdvanceInterest
                Dim coll As New Hashtable()
                clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", strDocNo)
                clsCommon.AddColumnsForChange(coll, "Payment_No", obj.Payment_No)
                clsCommon.AddColumnsForChange(coll, "Interest_Amount", obj.Interest_Amount)
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_AP_Invoice_Advance_Interest", OMInsertOrUpdate.Insert, "", trans)
            Next
        End If
        Return True
    End Function
    Public Shared Function GetData(ByVal strDocumentNo As String, ByVal trans As SqlTransaction) As List(Of clsAPInvoiceAdvanceInterest)
        Dim qry As String = "Select * from TSPL_AP_Invoice_Advance_Interest where AP_Invoice_No='" + strDocumentNo + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
        Dim arr As List(Of clsAPInvoiceAdvanceInterest) = Nothing
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            arr = New List(Of clsAPInvoiceAdvanceInterest)
            Dim objTr As clsAPInvoiceAdvanceInterest
            For Each dr As DataRow In dt.Rows
                objTr = New clsAPInvoiceAdvanceInterest
                objTr.AP_Invoice_No = clsCommon.myCstr(dr("AP_Invoice_No"))
                objTr.Payment_No = clsCommon.myCstr(dr("Payment_No"))
                objTr.Interest_Amount = clsCommon.myCdbl(dr("Interest_Amount"))
                arr.Add(objTr)
            Next
        End If
        Return arr
    End Function

    Public Shared Function GetAdvancePaymentQry(ByVal strVendorCode As String, ByVal dtFrom As DateTime?, ByVal dtTo As DateTime, ByVal skipPreviousDocumeent As Boolean, ByVal tran As SqlTransaction) As String
        Return GetAdvancePaymentQry(strVendorCode, dtFrom, dtTo, skipPreviousDocumeent, tran, True)
    End Function
    Public Shared Function GetAdvancePaymentQry(ByVal strVendorCode As String, ByVal dtFrom As DateTime?, ByVal dtTo As DateTime, ByVal skipPreviousDocumeent As Boolean, ByVal tran As SqlTransaction, ByVal isOrderBy As Boolean) As String
        Dim dtToDateForQry As DateTime = dtTo ''BHA/14/09/18-000549 by balwinder on 18/09/2018
        If (clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.DoNotConsiderTheFutureDateOfAdvancePayment, clsFixedParameterCode.DoNotConsiderTheFutureDateOfAdvancePayment, tran)) = 0) Then
            dtToDateForQry = clsCommon.GETSERVERDATE(tran)
        End If

        Dim qry As String = "Select * from (" &
               " Select TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_PAYMENT_HEADER.Vendor_Name,TSPL_PAYMENT_HEADER.Payment_No,TSPL_PAYMENT_HEADER.Payment_Date,TSPL_PAYMENT_HEADER.Payment_Amount+isnull(TDS_Amount ,0) as Payment_Amount," &
               " Balance_Amt+isnull(TDS_Amount ,0)-ISNULL((Select SUM(Payment_Amount)+sum(isnull(TDS_Amount ,0)) from TSPL_PAYMENT_HEADER PH WHERE PH.Posted<>'1' AND PH.Payment_Type='AD' AND PH.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No),0) as Balance_Amt,TSPL_PAYMENT_HEADER.Interest_Rate,TSPL_PAYMENT_HEADER.No_Of_EMI  from TSPL_PAYMENT_HEADER WHERE Posted='1' "
        If clsCommon.myLen(strVendorCode) <= 0 Then
        Else
            qry += " AND Vendor_Code in  (" + strVendorCode + ") "
        End If
        qry += " AND   IsChkReverse='N' and Payment_Type IN ('AV' "
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PickOnAccountPaymentForAdvanceKnockOff, clsFixedParameterCode.PickOnAccountPaymentForAdvanceKnockOff, tran)) > 0 Then
            qry += " ,'OA' "
        End If
        qry += ")"
        If skipPreviousDocumeent Then
            qry += " and TSPL_PAYMENT_HEADER.Payment_Date >= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(dtFrom), "dd/MMM/yyyy hh:mm tt") + "' "
        End If
        qry += " and TSPL_PAYMENT_HEADER.Payment_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtToDateForQry), "dd/MMM/yyyy hh:mm tt") + "' " &
       " ) Final where Balance_Amt>0 "
        If isOrderBy Then
            qry += " order by Payment_Date"
        End If
        Return qry
    End Function

End Class

Public Class clsProvisionEntryAdvanceKnockOff
#Region "Variables"
    Public Document_No As String = Nothing
    Public AP_Invoice_No As String = Nothing
    Public Provision_No As String = Nothing
    Public Balance_Amount As Decimal = 0
    Public Knockoff_Amount As Decimal = 0
#End Region

    Public Shared Function SaveData(ByVal obj As clsVedorInvoiceHead, ByVal trans As SqlTransaction) As Boolean
        If (obj.arrProvDocNo IsNot Nothing AndAlso obj.arrProvDocNo.Count > 0) Then
            Dim CalculateProvisionOnGateePass As Integer = 0 ' added by preeti gupta Against ticket no[UDL/10/01/19-000252]
            CalculateProvisionOnGateePass = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CalculateProvisionOnGateePass, clsFixedParameterCode.CalculateProvisionOnGateePass, trans))
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(GetBalanceQry(obj.Document_No, obj.Vendor_Code, Nothing, Nothing, obj.arrProvDocNo, obj.arrInvoiceNo, obj.loc_code, CalculateProvisionOnGateePass), trans)
            Dim balAmt As Decimal = obj.Discount_Base
            For Each dr As DataRow In dt.Rows
                balAmt -= clsCommon.myCdbl(dr("Amount"))
                Dim coll As New Hashtable()
                Dim strCode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select max(Document_No) from TSPL_PROVISION_ENTRY_KNOCKOFF", trans))
                If clsCommon.myLen(strCode) <= 0 Then
                    strCode = "PEKO0000000001"
                Else
                    strCode = clsCommon.incval(strCode)
                End If
                If clsCommon.myLen(strCode) <= 0 Then
                    Throw New Exception("Error in code genereation of Provision knock off")
                End If
                clsCommon.AddColumnsForChange(coll, "Document_No", strCode)
                clsCommon.AddColumnsForChange(coll, "AP_Invoice_No", obj.Document_No)
                clsCommon.AddColumnsForChange(coll, "Provision_No", clsCommon.myCstr(dr("Doc_No")))
                clsCommon.AddColumnsForChange(coll, "Invoice_No", clsCommon.myCstr(dr("Ref_Doc_No")))
                clsCommon.AddColumnsForChange(coll, "Balance_Amount", clsCommon.myCstr(dr("Amount")))
                If balAmt <= 0 Then
                    clsCommon.AddColumnsForChange(coll, "Knockoff_Amount", balAmt + clsCommon.myCdbl(dr("Amount")))
                Else
                    clsCommon.AddColumnsForChange(coll, "Knockoff_Amount", clsCommon.myCdbl(dr("Amount")))
                End If
                clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PROVISION_ENTRY_KNOCKOFF", OMInsertOrUpdate.Insert, "", trans)
                If balAmt <= 0 Then
                    Exit For
                End If
            Next
        End If
        clsDBFuncationality.ExecuteNonQuery("update TSPL_VENDOR_INVOICE_HEAD set Prov_Amt=( select sum(TSPL_PROVISION_ENTRY_KNOCKOFF.Knockoff_Amount) as Knockoff_Amount from TSPL_PROVISION_ENTRY_KNOCKOFF where TSPL_PROVISION_ENTRY_KNOCKOFF.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No) where TSPL_VENDOR_INVOICE_HEAD.Document_No='" + obj.Document_No + "'", trans)
        Return True
    End Function
    'Public Shared Function GetData(ByVal strDocumentNo As String, ByVal trans As SqlTransaction) As List(Of clsAPInvoiceAdvanceInterest)
    '    Dim qry As String = "Select * from TSPL_AP_Invoice_Advance_Interest where AP_Invoice_No='" + strDocumentNo + "'"
    '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
    '    Dim arr As List(Of clsAPInvoiceAdvanceInterest) = Nothing
    '    If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
    '        arr = New List(Of clsAPInvoiceAdvanceInterest)
    '        Dim objTr As clsAPInvoiceAdvanceInterest
    '        For Each dr As DataRow In dt.Rows
    '            objTr = New clsAPInvoiceAdvanceInterest
    '            objTr.AP_Invoice_No = clsCommon.myCstr(dr("AP_Invoice_No"))
    '            objTr.Payment_No = clsCommon.myCstr(dr("Payment_No"))
    '            objTr.Interest_Amount = clsCommon.myCdbl(dr("Interest_Amount"))
    '            arr.Add(objTr)
    '        Next
    '    End If
    '    Return arr
    'End Function

    Public Shared Function GetBalanceQry(ByVal strAPInvNo As String, ByVal VendorCode As String, ByVal FromDate? As DateTime, ByVal ToDate? As DateTime, ByVal arrProvision As List(Of String), ByVal arrInvoice As List(Of String), ByVal location As String, ByVal CalculateProvisionOnGateePass As Integer) As String

        Dim qry As String = "select max( Prog_Code) as Prog_Code,max(Doc_No) as Doc_No,max(Doc_Date) as Doc_Date,max(Loc_Code) as Loc_Code,max(Loc_Desc) as Loc_Desc,max(Vendor_Code) as Vendor_Code,max(Vendor_Desc) as Vendor_Desc,max(Vendor_Type) as Vendor_Type,max(Route_Code) as Route_Code,max(Ref_Doc_No) as Ref_Doc_No,max(Prov_type) as Prov_type" + Environment.NewLine +
           ",max(Status) as Status,max(Prov_Month) as Prov_Month,max(Prov_Year) as Prov_Year,max(Comp_Code) as Comp_Code,max(Created_by) as Created_by,max(Created_Date) as  Created_Date,max(Modified_By) as Modified_By,max(Modified_Date) as Modified_Date,max(isPosted) as isPosted,max(Posting_Date) Posting_Date" + Environment.NewLine +
           ",sum( Amount * RI) as Amount from ("
        If CalculateProvisionOnGateePass = 0 Then
            qry += "select Prog_Code,Doc_No,Doc_Date,Loc_Code,Loc_Desc,Vendor_Code,Vendor_Desc,Vendor_Type,Route_Code,Ref_Doc_No,Prov_type,Amount ,1 as RI,1 as chk" + Environment.NewLine +
           ",Status,Prov_Month,Prov_Year,Comp_Code,Created_by, Created_Date,Modified_By,Modified_Date,isPosted,Posting_Date" + Environment.NewLine +
           "from tspl_provision_entry " + Environment.NewLine +
           "where Vendor_Code='" + VendorCode + "'  and (loc_Code in (select location_code from TSPL_LOCATION_MASTER   where Loc_Segment_Code='" + location + "') or tspl_provision_entry.GL_Location_Seg='" + location + "'  ) and isPosted='1'  "
        Else
            qry += "select Prog_Code,Doc_No,Doc_Date,Loc_Code,Loc_Desc,Vendor_Code,Vendor_Desc,Vendor_Type,Route_Code,TSPL_GATEPASS_DETAIL_PRODUCTSALE.Invoice_No as Ref_Doc_No,Prov_type,TSPL_GATEPASS_DETAIL_PRODUCTSALE.Provision_Amount as Amount ,1 as RI,1 as chk" + Environment.NewLine +
                      ",Status,Prov_Month,Prov_Year,tspl_provision_entry.Comp_Code,tspl_provision_entry.Created_by, tspl_provision_entry.Created_Date,tspl_provision_entry.Modified_By,tspl_provision_entry.Modified_Date,isPosted,Posting_Date" + Environment.NewLine +
                      "from tspl_provision_entry " + Environment.NewLine +
                     " left join TSPL_GATEPASS_master_PRODUCTSALE on TSPL_GATEPASS_master_PRODUCTSALE.Provision_No =tspl_provision_entry.Doc_No " + Environment.NewLine +
                        " left join TSPL_GATEPASS_DETAIL_PRODUCTSALE on TSPL_GATEPASS_DETAIL_PRODUCTSALE.GPCode =TSPL_GATEPASS_master_PRODUCTSALE.GPCode " + Environment.NewLine +
                      "where Vendor_Code='" + VendorCode + "'  and (loc_Code in (select location_code from TSPL_LOCATION_MASTER   where Loc_Segment_Code='" + location + "') or tspl_provision_entry.GL_Location_Seg='" + location + "' ) and isPosted='1'  "
            If arrInvoice IsNot Nothing AndAlso arrInvoice.Count > 0 Then
                qry += " and TSPL_GATEPASS_DETAIL_PRODUCTSALE.Invoice_No in (" + clsCommon.GetMulcallString(arrInvoice) + ")"
            End If
        End If

        If FromDate IsNot Nothing Then
            qry += " and Doc_Date >='" + clsCommon.GetPrintDate(clsCommon.GetDateWithStartTime(FromDate), "dd/MMM/yyyy hh:mm:ss tt") + "' "
        End If
        If ToDate IsNot Nothing Then
            qry += " and Doc_Date <='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(ToDate), "dd/MMM/yyyy hh:mm:ss tt") + "'"
        End If
        If arrProvision IsNot Nothing AndAlso arrProvision.Count > 0 Then
            qry += " and tspl_provision_entry.Doc_No in (" + clsCommon.GetMulcallString(arrProvision) + ")"
        End If

        qry += " union all" + Environment.NewLine +
         " select null as Prog_Code,Provision_No as Doc_No,null as Doc_Date,null as Loc_Code,null as Loc_Desc,null as Vendor_Code, null as Vendor_Desc,null as Vendor_Type,null as Route_Code,TSPL_PROVISION_ENTRY_KNOCKOFF.Invoice_No as Ref_Doc_No,null as Prov_type,Knockoff_Amount as Amount ,-1 as RI ,0 as chk" + Environment.NewLine +
       ",null as Status,null as Prov_Month,null as Prov_Year,null as Comp_Code,null as Created_by,null as  Created_Date,null as Modified_By,null as Modified_Date,null as isPosted,null Posting_Date" + Environment.NewLine +
       " from TSPL_PROVISION_ENTRY_KNOCKOFF where AP_Invoice_No not in ('" + strAPInvNo + "')" + Environment.NewLine +
        " )xx "
        If CalculateProvisionOnGateePass = 0 Then
            qry += " group by Doc_No having sum(chk)>0 and sum(Amount*RI)>0 order by Doc_Date"
        Else
            qry += " group by Doc_No,Ref_Doc_No having sum(chk)>0 and sum(Amount*RI)>0 order by Doc_Date"
        End If


        Return qry
    End Function
End Class

