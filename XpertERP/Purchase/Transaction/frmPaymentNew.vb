Imports common
Imports System.Data.SqlClient
Imports System
Imports XpertERPEngine

Public Class FrmPaymentNew
    Inherits FrmMainTranScreen
#Region "Variables"
    Dim ButtonToolTip As ToolTip = New ToolTip()
    Private isSettlementBankOnly As Boolean = False
    Public strPaymentNo As String = Nothing
    Private isCellValueChangedTaxOpen As Boolean = False
    '-----------Used in Bank Reco---------------
    Dim isFlag As Boolean = False
    Public ChequeNo As String = Nothing
    Public ChequeDate As Date? = Nothing
    Public Amount As Decimal = 0
    Public EntryDesc As String = Nothing
    '-------------------------------------------
    Dim iStxtTaxGroup_TxtChangedComplete As Boolean = True
    Private IsInsideLoadData As Boolean = False
    Private isNewEntry As Boolean = False
    Private objRemittance As clsRemittance
    Dim Qry As String = ""
    Dim dt As DataTable
    Dim btnToolTip As ToolTip = New ToolTip()
    Dim IsPaymentTypeChanged As Boolean = False
    Dim GSTStatus As Boolean = False
    '==============Grid Columns===================
    Const colApply As String = "Apply"
    Const colPI_No As String = "PurchaseInvoice"
    Const colDocType As String = "DocType"
    Const colPINo As String = "PurchaseInvoice"
    Const colDocNo As String = "DocNo"
    Const colAPTapalNo As String = "APTapalNo"
    Const colDocDate As String = "VendorDocDate"
    Const colDocumentDate As String = "DocumentDate"
    Const colVendorInvNo As String = "VendorInvNo"
    Const colNetAmt As String = "NetAmount"
    Const colPendingAmt As String = "PendingAmt"
    Const colAppliedAmt As String = "AppliedAmt"
    Const colSecurityAmt As String = "SecurityAmt"
    Const colOriginalAmt As String = "OriginalAmt"
    Const colTDSAmt As String = "TDSAmt"
    Const colComment As String = "Comment"
    Const colTemp As String = "Temp"

    Const colLineNo As String = "LineNo"
    Const colGLAccount As String = "GLAccount"
    Const colGLType As String = "colGLType"
    Const colAccDesc As String = "AccDesc"
    Const colHirerachyCenter As String = "colHirerachyCenter"
    Const colHirerachyName As String = "colHirerachyName"
    Const colCostCenter As String = "colCostCenter"
    Const colCostCenterName As String = "colCostCenterName"
    Const colAmount As String = "Amount"
    Const colRemark As String = "Remark"
    Const colExpenseCode As String = "ExpenseCode"
    Public StrdocNo As String = ""
    Const colAdjustedAmt As String = "Adjusted Amount"
    '---------------------------------------------
    Dim isApplyBranchAccounting As Boolean
    Dim TagExemptedtaxgroupincaseofBankChargesinPaymentEntry As Boolean = False
    Dim isApplyCostCenter As Boolean
    Dim deadLockCounter As Integer
    Dim Arr As ArrayList
    Public PDCSetting As Boolean = False
    Dim ApplyBankChargesasperSlabonBankMaster As Boolean = False
    Dim AllowSameChequeNoForMultiplePaymentEntry As Boolean = False
    Dim EnableBankFromMaster As Boolean = False
    Dim autoadjustdebitcreditnote As Boolean = False
    ' -------------- purchase order item grid details 
    Const AdcolDocument_Code As String = "AdcolDocument_Code"
    Const AdcolLine_No As String = "AdcolLine_No"
    Const AdcolRow_Type As String = "AdcolRow_Type"
    Const AdcolItem_Code As String = "AdcolItem_Code"
    Const adcolIName As String = "adcolIName"
    Const adcolIHSN As String = "adcolIHSN"
    Const AdcolQty As String = "AdcolQty"
    Const AdcolBalance_Qty As String = "AdcolBalance_Qty"
    Const AdcolUnit_code As String = "AdcolUnit_code"
    Const AdcolItem_Cost As String = "AdcolItem_Cost"
    Const AdcolTAX1 As String = "COLTAX1"
    Const AdcolTAX1_Amt As String = "COLTAXAMT1"
    Const AdcolTAX1_Base_Amt As String = "COLTAXBASEAMT1"
    Const AdcolTAX1_Rate As String = "COLTAXRATE1"
    Const Adcoltax2 As String = "COLTAX2"
    Const AdcolTAX2_Base_Amt As String = "COLTAXBASEAMT2"
    Const AdcolTAX2_Rate As String = "COLTAXRATE2"
    Const AdcolTAX2_Amt As String = "COLTAXAMT2"
    Const AdcolTAX3 As String = "COLTAX3"
    Const AdcolTAX3_Base_Amt As String = "COLTAXBASEAMT3"
    Const AdcolTAX3_Rate As String = "COLTAXRATE3"
    Const AdcolTAX3_Amt As String = "COLTAXAMT3"
    Const AdcolTAX4 As String = "COLTAX4"
    Const AdcolTAX4_Base_Amt As String = "COLTAXBASEAMT4"
    Const AdcolTAX4_Rate As String = "COLTAXRATE4"
    Const AdcolTAX4_Amt As String = "COLTAXAMT4"
    Const AdcolTAX5 As String = "COLTAX5"
    Const AdcolTAX5_Base_Amt As String = "COLTAXBASEAMT5"
    Const AdcolTAX5_Rate As String = "COLTAXRATE5"
    Const AdcolTAX5_Amt As String = "COLTAXAMT5"
    Const AdcolTAX6 As String = "COLTAX6"
    Const AdcolTAX6_Base_Amt As String = "COLTAXBASEAMT6"
    Const AdcolTAX6_Rate As String = "COLTAXRATE6"
    Const AdcolTAX6_Amt As String = "COLTAXAMT6"
    Const AdcolTAX7 As String = "COLTAX7"
    Const AdcolTAX7_Base_Amt As String = "COLTAXBASEAMT7"
    Const AdcolTAX7_Rate As String = "COLTAXRATE7"
    Const AdcolTAX7_Amt As String = "COLTAXAMT7"
    Const AdcolTAX8 As String = "COLTAX8"
    Const AdcolTAX8_Base_Amt As String = "COLTAXBASEAMT8"
    Const AdcolTAX8_Rate As String = "COLTAXRATE8"
    Const AdcolTAX8_Amt As String = "COLTAXAMT8"
    Const AdcolTAX9 As String = "COLTAX9"
    Const AdcolTAX9_Base_Amt As String = "COLTAXBASEAMT9"
    Const AdcolTAX9_Rate As String = "COLTAXRATE9"
    Const AdcolTAX9_Amt As String = "COLTAXAMT9"
    Const AdcolTAX10 As String = "COLTAX10"
    Const AdcolTAX10_Base_Amt As String = "COLTAXBASEAMT10"
    Const AdcolTAX10_Rate As String = "COLTAXRATE10"
    Const AdcolTAX10_Amt As String = "COLTAXAMT10"
    Const AdcolAmount As String = "AdcolAmount"
    Const AdcolDisc_Per As String = "AdcolDisc_Per"
    Const AdcolDisc_Amt As String = "AdcolDisc_Amt"
    Const AdcolAmt_Less_Discount As String = "AdcolAmt_Less_Discount"
    Const AdcolTotal_Tax_Amt As String = "AdcolTotal_Tax_Amt"
    Const AdcolItem_Net_Amt As String = "AdcolItem_Net_Amt"

    Const AdcolTAX1_Amt_Payment As String = "COLTAXAMTPAYMENT1"
    Const AdcolTAX2_Amt_Payment As String = "COLTAXAMTPAYMENT2"
    Const AdcolTAX3_Amt_Payment As String = "COLTAXAMTPAYMENT3"
    Const AdcolTAX4_Amt_Payment As String = "COLTAXAMTPAYMENT4"
    Const AdcolTAX5_Amt_Payment As String = "COLTAXAMTPAYMENT5"
    Const AdcolTAX6_Amt_Payment As String = "COLTAXAMTPAYMENT6"
    Const AdcolTAX7_Amt_Payment As String = "COLTAXAMTPAYMENT7"
    Const AdcolTAX8_Amt_Payment As String = "COLTAXAMTPAYMENT8"
    Const AdcolTAX9_Amt_Payment As String = "COLTAXAMTPAYMENT9"
    Const AdcolTAX10_Amt_Payment As String = "COLTAXAMTPAYMENT10"

    Const AdcolPaymentAdvance As String = "AdcolPaymentAdvance"
    Const AdcolPaymentTotalTax As String = "AdcolPaymentTotalTax"
    Const AdcolPaymentTotalAdvanceAmt As String = "AdcolPaymentTotalAdvanceAmt"

    '' grid detial of tax 

    Const colTTaxAutCode As String = "TAXAUTCODE"
    Const colTTaxAutName As String = "TAXAUTNAME"
    Const colTTaxRate As String = "TAXRATE"
    Const colTBaseAmt As String = "TAXBASEAMT"
    Const colTTaxAmt As String = "TAXAMT"


    Const colIsTaxable1 As String = "ISTAXABLE1"
    Const colIsSurTax1 As String = "ISSURTAX1"
    Const colSurTaxCode1 As String = "SURTAXCODE1"
    Const colIsTaxable2 As String = "ISTAXABLE2"
    Const colIsSurTax2 As String = "ISSURTAX2"
    Const colSurTaxCode2 As String = "SURTAXCODE2"
    Const colIsTaxable3 As String = "ISTAXABLE3"
    Const colIsSurTax3 As String = "ISSURTAX3"
    Const colSurTaxCode3 As String = "SURTAXCODE3"
    Const colIsTaxable4 As String = "ISTAXABLE4"
    Const colIsSurTax4 As String = "ISSURTAX4"
    Const colSurTaxCode4 As String = "SURTAXCODE4"
    Const colIsTaxable5 As String = "ISTAXABLE5"
    Const colIsSurTax5 As String = "ISSURTAX5"
    Const colSurTaxCode5 As String = "SURTAXCODE5"
    Const colIsTaxable6 As String = "ISTAXABLE6"
    Const colIsSurTax6 As String = "ISSURTAX6"
    Const colSurTaxCode6 As String = "SURTAXCODE6"
    Const colIsTaxable7 As String = "ISTAXABLE7"
    Const colIsSurTax7 As String = "ISSURTAX7"
    Const colSurTaxCode7 As String = "SURTAXCODE7"
    Const colIsTaxable8 As String = "ISTAXABLE8"
    Const colIsSurTax8 As String = "ISSURTAX8"
    Const colSurTaxCode8 As String = "SURTAXCODE8"
    Const colIsTaxable9 As String = "ISTAXABLE9"
    Const colIsSurTax9 As String = "ISSURTAX9"
    Const colSurTaxCode9 As String = "SURTAXCODE9"
    Const colIsTaxable10 As String = "ISTAXABLE10"
    Const colIsSurTax10 As String = "ISSURTAX10"
    Const colSurTaxCode10 As String = "SURTAXCODE10"


    ''----------
    Const colBCaxAutCode As String = "TAXBCAUTCODE"
    Const colBCaxAutName As String = "TAXBCAUTNAME"
    Const colBCLine_No As String = "colBCLine_No"
    Const colBCaxRate As String = "TAXBCRATE"
    Const colBCBaseAmt As String = "TAXBCBASEAMT"
    Const colBCaxAmt As String = "TAXBCAMT"
    Dim ERPStartDate As Date
    Dim CostCenterAndHirerachyCodeUpdateAfterPost As Boolean = False
#End Region


    Private Sub FrmPaymentNew_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim coll As Dictionary(Of String, String)

        coll = New Dictionary(Of String, String)()
        coll.Add("Transport_Id", "varchar(12) null")
        coll.Add("Debit_Account", "Varchar(50) null")
        coll.Add("Credit_Account", "Varchar(50) null")
        coll.Add("Balance_Amt", "Decimal(18,2) null")
        coll.Add("FIFO_Balance", "Decimal(18,2) not null default 0.00")
        coll.Add("QuickEntryNo", "varchar(30) Null ")
        coll.Add("LoadOutNo", "varchar(30) Null ")
        coll.Add("Salesman_Code", "varchar(30) Null ")
        coll.Add("Salesman_Name", "varchar(200) Null ")
        coll.Add("Route_NO", "varchar(30) Null ")
        coll.Add("Route_Description", "varchar(200) Null ")
        coll.Add("Location_Code", "varchar(30) Null ")
        coll.Add("Location_Description", "varchar(200) Null ")
        coll.Add("IsRecoCleared", "char(1) not Null default 'N' ")
        coll.Add("IsChkReverse", "char(1) not Null default 'N' ")
        coll.Add("Loadout_No", "varchar(30) Null ")
        coll.Add("Bank_Charges_Ac", "Varchar(50) null")
        coll.Add("Bank_Charges", "Decimal(18,2) null")
        coll.Add("CURRENCY_CODE", "VARCHAR(30)  NULL REFERENCES TSPL_CURRENCY_MASTER(CURRENCY_CODE) ")
        coll.Add("ConvRate", "Decimal(18,6) NOT NULL DEFAULT 1")
        coll.Add("ApplicableFrom", "DATE NULL")
        coll.Add("BASE_CURRENCY_CODE", "VARCHAR(30)  NULL REFERENCES TSPL_CURRENCY_MASTER(CURRENCY_CODE) ")
        coll.Add("PAYMENT_AMOUNT_BASE_CURRENCY", "Decimal(18,2) NOT NULL DEFAULT 0")
        coll.Add("EXCHANGE_LOSS_AMT", "Decimal(18,2) NOT NULL DEFAULT 0")
        coll.Add("EXCHANGE_GAIN_AMT", "Decimal(18,2) NOT NULL DEFAULT 0")
        coll.Add("EXCHANGE_LOSS_ACCOUNT", "Varchar(50) null References TSPL_GL_ACCOUNTS(Account_Code)")
        coll.Add("EXCHANGE_GAIN_ACCOUNT", "Varchar(50) null References TSPL_GL_ACCOUNTS(Account_Code)")
        coll.Add("ConvRateOld", "Decimal(18,6) NOT NULL DEFAULT 1")
        coll.Add("CFormRecd", "char(1)  Null default 'N'")
        coll.Add("CForm_InvoiceNo", "varchar(30) Null ")
        coll.Add("EMP_CODE", "VARCHAR(12)  NULL REFERENCES TSPL_EMPLOYEE_MASTER(EMP_CODE)")
        coll.Add("PROJECT_CODE", "Varchar(30) NULL References TSPL_PJC_PROJECT(PROJECT_CODE)")
        coll.Add("PDC_Cheque", "char(1) Not Null default 'N'")
        coll.Add("Document_No", "Varchar(30) null References TSPL_PJC_EXPENSE_HEADER(Document_No)")
        coll.Add("CHECK_PRINT", "INT Not Null default 0")
        coll.Add("CHECK_CODE", "VARCHAR(30) NULL REFERENCES TSPL_BANK_CHECK_PRINTING(CHECK_CODE)")
        coll.Add("memorandum_amt", "float NOT NULL DEFAULT 0")
        coll.Add("Applied_Payment", "char(30) NOT NULL Default ''")
        coll.Add("Account_Payee", "BIT NOT NULL DEFAULT 0")
        coll.Add("PurchaseOrder_No", "Varchar(30) null References TSPL_PURCHASE_ORDER_HEAD(PurchaseOrder_No)")
        coll.Add("Is_Security", "INTEGER NOT NULL DEFAULT 0")
        coll.Add("Is_Retention", "INTEGER NOT NULL DEFAULT 0")
        coll.Add("Account_Payee_Name", "varchar(200) Null ")
        coll.Add("Total_Security_Amount", "Decimal(18,2) NOT NULL DEFAULT 0")
        coll.Add("Location_GL_Code", "varchar(12) NULL")
        coll.Add("Advance_Against_Salary", "BIT NOT NULL DEFAULT 0")
        coll.Add("is_Opening", "integer not null default 0")
        coll.Add("Payment_No", "varchar(30)  NOT NULL PRIMARY KEY ")
        coll.Add("Payment_Date", "datetime  NOT NULL")
        coll.Add("Payment_Post_Date", "datetime NULL")
        coll.Add("Bank_Code", "varchar(12)  NOT NULL REFERENCES TSPL_BANK_MASTER (BANK_CODE)")
        coll.Add("Payment_Type", "char(2)  NOT NULL")
        coll.Add("Vendor_Code", "varchar(12) NULL")
        coll.Add("Vendor_Name", "varchar(200) NULL")
        coll.Add("Remit_To", "varchar(60) NULL")
        coll.Add("Entry_Desc", "varchar(250) NULL")
        coll.Add("Reference", "varchar(100) NULL")
        coll.Add("Narration", "varchar(200) NULL")
        coll.Add("Payment_Code", "varchar(12)  NOT NULL")
        coll.Add("Cheque_No", "varchar(20) NULL")
        coll.Add("Cheque_Date", "date NULL")
        coll.Add("Payment_Amount", "decimal (18,2)  NOT NULL")
        coll.Add("Vendor_Account_Set", "varchar(12)  NOT NULL")
        coll.Add("TDS_Amount", "decimal (18,2) NULL")
        coll.Add("Total_Prepayment", "decimal (18,2) NULL")
        coll.Add("Apply_By", "varchar(30) NULL")
        coll.Add("Apply_To", "varchar(30) NULL")
        coll.Add("Posted", "char(1) NULL")
        coll.Add("Created_By", "varchar(12)  NOT NULL")
        coll.Add("Created_Date", "datetime  NOT NULL")
        coll.Add("Modify_By", "varchar(12)  NOT NULL")
        coll.Add("Modify_Date", "datetime  NOT NULL")
        coll.Add("Level1_User_code", "varchar(12) NULL")
        coll.Add("Level2_User_code", "varchar(12) NULL")
        coll.Add("Level3_User_code", "varchar(12) NULL")
        coll.Add("Level4_User_code", "varchar(12) NULL")
        coll.Add("Level5_User_code", "varchar(12) NULL")
        coll.Add("Comp_Code", "varchar(8)  NOT NULL")
        coll.Add("Total_Applied_Amount", "decimal (18,2) NULL")
        coll.Add("Loan_Code", "varchar(30) NULL REFERENCES TSPL_LOAN_APPLICATION (Loan_Code)")
        coll.Add("Against_TDS_PAYMENT_No", "varchar(30) NULL references TSPL_TDS_Payment_Header(Document_No) ")
        coll.Add("IsApplyDocAuto", "int not null default 0")
        coll.Add("Interest_Rate", "decimal(18,2) NULL")
        coll.Add("No_Of_EMI", "decimal(18,2) NULL")
        coll.Add("PurchaseOrder_No_GST", "Varchar(30) null References TSPL_PURCHASE_ORDER_HEAD(PurchaseOrder_No)")
        coll.Add("Tax_Group", "varchar(12) NULL")
        coll.Add("PurchaseOrder_Amount", "float not null default 0")
        coll.Add("Tax_Amount_Advance", "float not null default 0")
        coll.Add("PurchaseOrder_Add_Amount", "float not null default 0")
        coll.Add("PO_Location_Code", "varchar(12) NULL")
        coll.Add("TAX1", "varchar(12) NULL")
        coll.Add("TAX1_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX1_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX1_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX2", "varchar(12) NULL")
        coll.Add("TAX2_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX2_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX2_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX3", "varchar(12) NULL")
        coll.Add("TAX3_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX3_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX3_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX4", "varchar(12) NULL")
        coll.Add("TAX4_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX4_Rate", "decimal(18, 2) NULL")
        coll.Add("TAX4_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX5", "varchar(12) NULL")
        coll.Add("TAX5_Base_Amt", "decimal(18, 2) NULL")
        coll.Add("TAX5_Rate", "decimal(18, 2) NULL")
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
        coll.Add("GSTRegistered", "Integer not null default 1")
        coll.Add("Tax_Group_BankCharges", "Varchar(12)  NULL ")
        coll.Add("Bank_Charges_Tax", "decimal(18, 2) NULL")
        coll.Add("Employee_Type", "char(5) null")
        coll.Add("Employee_Advance_Type", "char(5) null")
        coll.Add("WaveOFFBankCharges", "char(1) null ")
        coll.Add("Against_Salary_Generation_Code", "Varchar(30) null references TSPL_GENERATE_SALARY(SALARY_GENERATION_CODE)")
        coll.Add("DateAndTime", "DATETIME NULL")
        coll.Add("TapalNo", "VARCHAR(100) NULL")
        coll.Add("MP_Code_For_Advance", "Varchar(30) null References TSPL_MP_MASTER(MP_Code)")
        coll.Add("isReceipt", "Integer not null default 0")
        coll.Add("isFarmerLoanPayment", "Integer not null default 0")
        coll.Add("Vendor_Bank_Code", "varchar(50) NULL")
        coll.Add("Vendor_Bank_Name", "varchar(50) NULL ")
        coll.Add("Vendor_IFSC_Code", "varchar(50) NULL")
        coll.Add("Vendor_Branch_Name", "varchar(150) NULL")
        coll.Add("Vendor_Bank_ACNo", "varchar(50) NULL")
        coll.Add("Against_VSP_Asset_Lost", "Varchar(30) null References TSPL_VSPAsset_HEAD(Doc_No)")
        coll.Add("Saving", "INTEGER NULL")
        coll.Add("TDS_Provision", "INTEGER not null default 0")
        coll.Add("Location_Code_Prefix", "VARCHAR(12) NULL REFERENCES TSPL_LOCATION_MASTER(LOCATION_CODE)")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_PAYMENT_HEADER", coll, Nothing, True, True, "", "Payment_No", "Payment_Date", True)


        ''UDL/24/07/18-000208 richa add foreign key with payment no.
        coll = New Dictionary(Of String, String)()
        coll.Add("ESI_WCT_Percentage", "decimal(18,2) not null default 0.0")
        coll.Add("Post", "char(1)  not null default 'N'")
        coll.Add("Settlement_code", "varchar(50) null")
        coll.Add("Settlement_Description", "varchar(100) null")
        coll.Add("EXPENSE_CODE", "VARCHAR(30)  NULL References TSPL_EXPENSE_MASTER(EXPENSE_CODE)")
        coll.Add("ConvRateOld", "Decimal(18,6) NOT NULL DEFAULT 1")
        coll.Add("Security_Amount", "Decimal(18,2) NOT NULL DEFAULT 0")
        coll.Add("AP_Invoice_No", "varchar(100) null")
        coll.Add("Payment_No", "varchar(30) not null References TSPL_PAYMENT_HEADER(Payment_No)")
        coll.Add("Payment_Line_No", "int  NOT NULL")
        coll.Add("Apply", "char(1) NULL")
        coll.Add("Payment_Type", "char(2) NULL")
        coll.Add("Document_No", "varchar(30) NULL")
        coll.Add("Vendor_Invoice_No", "varchar(30) NULL")
        coll.Add("Pending_Balance", "decimal (18,2)  NOT NULL")
        coll.Add("Applied_Amount", "decimal (18,2)  NOT NULL")
        coll.Add("Original_Invoice_Amt", "decimal (18,2) NULL")
        coll.Add("TDS_Amount", "decimal (18,2) NULL")
        coll.Add("Net_Balance", "decimal (18,2) NULL")
        coll.Add("Account_Code", "varchar(50) NULL")
        coll.Add("Description", "varchar(100) NULL")
        coll.Add("Remarks", "varchar(100) NULL")
        coll.Add("Comment", "varchar(100) NULL")
        coll.Add("Hirerachy_Level_Code", "Varchar(30) null References TSPL_HIRERACHY_LEVEL_MASTER(Hirerachy_Code)")
        coll.Add("Cost_Center_Fin_Code", "Varchar(30) null References TSPL_COST_CENTRE_FINANCIAL(Cost_Center_Fin_Code)")
        clsCommonFunctionality.CreateOrAlterTable(True, False, "TSPL_PAYMENT_DETAIL", coll, Nothing, True, True, "TSPL_PAYMENT_HEADER", "Payment_No", "", True)

        '' PJC tables ends here
        Try
            ERPStartDate = clsCommon.myCDate(objCommonVar.ERPStartDate)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow("Invalid ERP Start Date")
            Me.Close()
        End Try
        PDCSetting = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PDCSetting, clsFixedParameterCode.PDCSetting, Nothing)) = 1, True, False)
        isApplyBranchAccounting = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, Nothing)) = 1, True, False)
        isApplyCostCenter = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowHierarchyAndCostCenterInAPInvoiceEntry, clsFixedParameterCode.ShowHierarchyAndCostCenterInAPInvoiceEntry, Nothing)) = 1
        ApplyBankChargesasperSlabonBankMaster = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBankChargesasperSlabonBankMaster, clsFixedParameterCode.ApplyBankChargesasperSlabonBankMaster, Nothing)) = 1, True, False)
        AllowSameChequeNoForMultiplePaymentEntry = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowSameChequeNoForMultiplePaymentEntry, clsFixedParameterCode.AllowSameChequeNoForMultiplePaymentEntry, Nothing)) = 1, True, False)
        autoadjustdebitcreditnote = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.Autoadjustdebitcreditnote, clsFixedParameterCode.Autoadjustdebitcreditnote, Nothing)) = 1, True, False)
        EnableBankFromMaster = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.EnableBankFromMaster, clsFixedParameterCode.EnableBankFromMaster, Nothing)) = 1, True, False)
        CostCenterAndHirerachyCodeUpdateAfterPost = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.CostCenterAndHirerachyCodeUpdateAfterPost, clsFixedParameterCode.CostCenterAndHirerachyCodeUpdateAfterPost, Nothing)) = 1, True, False)
        SetUserMgmtNew()
        If objCommonVar.ApplyLocationWisePrefix Then
            pnlLocation.Visible = True
        Else
            pnlLocation.Visible = False
        End If
        loadPaymentType()
        loadEmployeeType()
        SetTollTip()
        Reset()
        SetLength()
        txtPaymentAmt.ReadOnly = True

        payment()
        RadPageView1.Pages("pvpCustomFields").Item.Visibility = MyBase.customFieldTabProperty
        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.Report_ID = MyBase.Form_ID
            UcCustomFields1.LoadCustomControls()
        End If

        If clsCommon.myLen(strPaymentNo) > 0 Then
            LoadData(strPaymentNo, NavigatorType.Current)
        End If
        txtConversionRate.Text = 1
        SetMultiCurrencyVisibility()
        Me.txtBaseCurrency.Value = objCommonVar.BaseCurrencyCode
        If objCommonVar.IsDemoERP Then
            UcAttachment1.Form_ID = MyBase.Form_ID
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Visible
        Else
            RadPageView1.Pages("Attachments").Item.Visibility = ElementVisibility.Collapsed
        End If

        pnlPJC.Visible = False
        chkCForm.Visible = False
        loadEmployeeAdvanceType()
        txtTotalPaymentBaseCurr.Enabled = False
        If clsCommon.myLen(Me.Tag) > 0 Then
            LoadData(clsCommon.myCstr(Me.Tag), NavigatorType.Current)
        End If

        GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
        RadPageView1.Pages("TabForGST").Item.Visibility = ElementVisibility.Collapsed
        If ApplyBankChargesasperSlabonBankMaster Then
            txtBankCharges.Enabled = False
            chkBankChargesWaveOff.Enabled = True
        Else
            txtBankCharges.Enabled = True
            chkBankChargesWaveOff.Enabled = False
        End If
        TagExemptedtaxgroupincaseofBankChargesinPaymentEntry = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.TagExemptedtaxgroupincaseofBankChargesinPaymentEntry, clsFixedParameterCode.TagExemptedtaxgroupincaseofBankChargesinPaymentEntry, Nothing)) = 1, True, False)

    End Sub
    Sub SetMultiCurrencyVisibility()
        Dim strq As String = ""
        Dim Currency As Integer = clsModuleCurrencyMapping.GetmulticurrencyDecimalPlaces()
        txtConversionRate.DecimalPlaces = clsCommon.myCdbl(Currency)
        If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
            pnlCurrConv.Visible = True

            If clsCommon.myLen(Me.txtVendorCode.Value) > 0 Then
                strq = "select currency_code from TSPL_VENDOR_MASTER where VENDOR_CODE='" & clsCommon.myCstr(Me.txtVendorCode.Value) & "'"
                Me.txtCurrencyCode.Value = clsDBFuncationality.getSingleValue(strq).ToString
                ShowCurrencyDetail()
            End If
            ShowCurrencyDetail()
        Else
            'SplitContainer2.SplitterDistance = SplitContainer2.SplitterDistance - pnlCurrConv.Height
            'pnlCurrConv.Height = 0
            pnlCurrConv.Visible = False

        End If

    End Sub
    Sub ShowCurrencyDetail()
        Dim dt As DataTable
        If clsCommon.myLen(clsCommon.myCstr(Me.txtVendorCode.Value)) = 0 Then
            Me.txtCurrencyCode.Value = objCommonVar.BaseCurrencyCode
            Me.txtConversionRate.Text = 1
            Me.txtApplicableFrom.Text = ""
            Me.txtConversionRate.ReadOnly = True
            Exit Sub
        End If

        If clsCommon.myLen(txtCurrencyCode.Value) > 0 Then
            dt = clsModuleCurrencyMapping.GetLatestCurConvRateDT(Me.dtpPayment.Value, txtCurrencyCode.Value)
            If dt.Rows.Count = 0 Then
                If Me.txtCurrencyCode.Value = objCommonVar.BaseCurrencyCode Then
                    Me.txtConversionRate.Text = 1
                    Me.txtApplicableFrom.Text = ""
                    Me.txtConversionRate.ReadOnly = True
                Else
                    clsCommon.MyMessageBoxShow("Conversion rate not entered for currency '" & Me.txtCurrencyCode.Value & "'")
                    Exit Sub
                End If
            Else
                Me.txtConversionRate.Text = dt.Rows(0).Item("Rate")
                Me.txtConversionRate.ReadOnly = False
                Me.txtApplicableFrom.Text = clsCommon.GetPrintDate(dt.Rows(0).Item("FROM_DATE"), "dd/MMM/yyyy")
            End If
        Else
            Me.txtConversionRate.Text = 1
            Me.txtApplicableFrom.Text = ""
            Me.txtConversionRate.ReadOnly = True
        End If

    End Sub
    Private Sub SetUserMgmtNew()
        If Not (MyBase.isReadFlag) Then
            Throw New Exception("Permission Denied")
        End If
        btnsave.Visible = MyBase.isModifyFlag
        btnpost.Visible = MyBase.isPostFlag
        btndelete.Visible = MyBase.isDeleteFlag
        btnprint.Visible = MyBase.isPrintFlag
        'RadMenu1.Visible = MyBase.isExport
        If MyBase.isExport = True Then
            rmiExport.Enabled = True
            rmiImport.Enabled = True
        Else
            rmiExport.Enabled = False
            rmiImport.Enabled = False
        End If
        'If MyBase.isReverse Then
        '    btnReverse.Enabled = True
        'Else
        '    btnReverse.Enabled = False
        'End If
        btnReverse.Visible = False

    End Sub

    Private Sub SetTollTip()
        btnToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Transaction")
        btnToolTip.SetToolTip(btnpost, "Press Alt+P Post Transaction")
        btnToolTip.SetToolTip(btndelete, "Press Alt+D Delete Transaction")
        btnToolTip.SetToolTip(btnclose, "Press Esc Close the Window")
        btnToolTip.SetToolTip(btnNew, "Press Alt+N Adding New Transaction")
    End Sub

    Private Sub Reset()
        butCostCenterAndHirerachy_Update_AfterPost.Visible = False
        chkFarmerLoanPayment.Checked = False
        txtDataAndTimeSelection.Value = clsCommon.GETSERVERDATE()
        txtTapalNo.Text = ""
        txtDataAndTimeSelection.Checked = False
        ddlPaymentType.SelectedValue = "PY"
        fndloanNo.Value = ""
        txtLocationPrefix.Value = ""
        txtLocationPrefixName.Text = ""
        chkmemorndm.Checked = False
        txtmemoamt.Text = ""
        isNewEntry = True
        dtpPayment.Enabled = True
        txtPaymentNo.Value = ""
        txtDescription.Text = ""
        chkSaving.Checked = False
        dtpPayment.Value = clsCommon.GETSERVERDATE()
        txtVendorCode.Value = ""
        lblVendorName.Text = ""
        txtChequeNo.Text = ""
        dtpChequeDate.Value = clsCommon.GETSERVERDATE()
        txtPaymentAmt.Text = ""
        txtTDSAmt.Text = "0"
        txtNetPayableAmt.Text = "0"
        txtTotalAppliedAmt.Text = "0"
        txtBankCharges.Text = ""
        txtRemitTo.Text = ""
        txtPONo.Value = ""
        txtmulPI.arrValueMember = Nothing
        LoadBlankGrid(ddlPaymentType.SelectedValue)
        dtpPayment.Focus()
        txtBankCode.Enabled = True
        txtVendorCode.Enabled = True
        ddlPaymentType.Enabled = True
        chkSaving.Enabled = True
        txtLoadOutno.Value = ""
        txtMPAdv.Value = ""
        chkIsReceipt.Checked = False
        chkIsReceipt.Visible = False
        btnsave.Enabled = True
        btnPrintCheck.Enabled = False
        btnVoidCheck.Enabled = False
        txtDocumentNo.Value = ""
        lblBalAmt.Text = ""
        UsLock1.Status = ERPTransactionStatus.Pending
        btnsave.Text = "Save"
        ChkSecurity.Checked = False
        ChkRetention.Checked = False
        txtlocation.Value = ""
        LblLocDesp.Text = ""
        txtlocation.Enabled = True
        lblApplyLoanNo.Visible = False
        fndloanNo.Visible = False
        chkTDSProvision.Checked = False
        If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "MI") = CompairStringResult.Equal Then
            gvDetails.Rows.AddNew()
        End If
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then 'If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
            ChkSecurity.Visible = True
            ChkRetention.Visible = True
        Else
            ChkSecurity.Visible = False
            ChkRetention.Visible = False
        End If
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
            If isApplyBranchAccounting = True Then
                RadLabel18.Visible = True
                txtlocation.Visible = True
                LblLocDesp.Visible = True
                txtlocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code in (select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' ) "))
                If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
                    LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
                Else
                    LblLocDesp.Text = ""
                End If
            Else
                RadLabel18.Visible = False
                txtlocation.Visible = False
                LblLocDesp.Visible = False
                txtlocation.Value = ""
                LblLocDesp.Text = ""
            End If
            ChkAdvSalary.Visible = True
            ChkAdvSalary.Checked = False
        Else
            ChkAdvSalary.Visible = False
            ChkAdvSalary.Checked = False
        End If

        If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
            UcCustomFields1.BlankAllControls()
        End If
        lblEmpCode.Text = ""
        lblEmpDesc.Text = ""
        lblProjCode.Text = ""
        lblProjDesc.Text = ""
        chkPDC.Checked = False
        LblAccPayee.Text = ""
        lblCustomerOutStanding.Text = "0"
        txtPaymentNo.MyReadOnly = False
        chkCheckPrint.Checked = False
        UcAttachment1.BlankAllControls()
        chkOpening.Checked = False
        btnChqUpdate.Visible = False
        If ChkAdvSalary.Checked = True Then
            lblApplyLoanNo.Visible = True
            fndloanNo.Visible = True
        Else
            lblApplyLoanNo.Visible = False
            fndloanNo.Visible = False
        End If
        txtConversionRate.Text = 1
        SetMultiCurrencyVisibility()
        Me.txtBaseCurrency.Value = objCommonVar.BaseCurrencyCode
        txtNoOfEMI.Value = 0
        txtInterestRate.Value = 0
        btnpost.Enabled = False

        txtPONo_GST.Value = ""
        txtTaxGroup.Value = ""
        lblTaxGrpName.Text = ""
        LblPOTotalAmount.Text = 0
        lblPOTotalAdditionalCharge.Text = 0
        lblPOTotalTaxAmt.Text = 0
        LBLPO_Location_GST.Text = ""
        TxtPO_Location_GST.Value = ""
        LoadBlankGridPOItemDetail()
        LoadBlankGridTax()
        txtTaxGroupBankCharges.Value = Nothing
        gv2.Rows.Clear()
        LoadBlankGridBankChargeTax()
        Dim EmployeeSalaryGeneration As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowtoEmployeeSalaryIntegration, clsFixedParameterCode.AllowtoEmployeeSalaryIntegration, Nothing)) = 1, True, False))
        If EmployeeSalaryGeneration = True Then
            'loadEmployeeType()
            ddlEmployeeType.Visible = True
            MyLabel13.Visible = True
            'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
            ddlEmployeeAdvanceType.Visible = True
            MyLabel14.Visible = True
        Else
            'loadEmployeeType()
            ddlEmployeeType.Visible = False
            MyLabel13.Visible = False
            'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
            ddlEmployeeAdvanceType.Visible = False
            MyLabel14.Visible = False
        End If

        ddlEmployeeType.Enabled = False
        'loadEmployeeType()
        ddlEmployeeAdvanceType.Enabled = False
        'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
        If ApplyBankChargesasperSlabonBankMaster Then
            txtBankCharges.Enabled = False
            chkBankChargesWaveOff.Enabled = True
        Else
            txtBankCharges.Enabled = True
            chkBankChargesWaveOff.Enabled = False
        End If
        chkBankChargesWaveOff.Checked = False

        txtVendor_bankcode.Text = ""
        TxtVendor_BankName.Text = ""
        TxtVendorBank_IFSCCode.Text = ""
        txtVendorBank_branchname.Text = ""
        txtVendor_Bank_ACNo.Text = ""
        If EnableBankFromMaster = True Then
            grpVendorBankDetails.Visible = False
        Else
            grpVendorBankDetails.Visible = True
        End If
        ddlEmployeeAdvanceType.SelectedValue = ""
        ddlEmployeeType.SelectedValue = ""

    End Sub

    Sub SetLength()
        txtDescription.MaxLength = 250
        txtChequeNo.MaxLength = 6
        txtRemitTo.MaxLength = 60
    End Sub
    Private Sub loadPaymentType()
        IsPaymentTypeChanged = False
        dt = New DataTable
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Value", GetType(String))

        dt.Rows.Add("Payment", "PY")
        dt.Rows.Add("Advance", "AV")
        dt.Rows.Add("On Account", "OA")
        dt.Rows.Add("Miscellaneous", "MI")
        dt.Rows.Add("Receipt", "RC")
        dt.Rows.Add("Apply Document", "AD")
        dt.Rows.Add("Receipt Against Manual Deduction", "MD")
        ddlPaymentType.DataSource = dt
        ddlPaymentType.DisplayMember = "Code"
        ddlPaymentType.ValueMember = "Value"
        LoadBlankGrid(ddlPaymentType.SelectedValue)
        IsPaymentTypeChanged = True
    End Sub
    Sub payment()

        pnlAdvance.Visible = False
        pnlMiscPayment.Visible = False
        pnlVendor.Visible = True
        txtVendorCode.Value = ""
        lblVendorName.Text = ""
        lblLoadOutNo.Visible = False
        txtLoadOutno.Visible = False
        lblMPAdv.Visible = False
        txtMPAdv.Visible = False

        txtBankCharges.Text = ""
        txtBankCharges.Enabled = True
        pnlCform.Visible = False

        pnlPJC.Visible = False
        pnlmemorndm.Visible = False
        lblDocumentNo.Visible = False
        txtDocumentNo.Visible = False
        lblBalAmt.Visible = False
        LblPONo.Visible = False
        txtPONo.Visible = False
        btnOk.Visible = False
        RadLabel18.Visible = False
        txtlocation.Value = ""
        LblLocDesp.Text = ""
        txtlocation.Visible = False
        LblLocDesp.Visible = False


        LoadBlankGrid(ddlPaymentType.SelectedValue)
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
            lblTotPayment.Text = "Total Payment"

            lblpaymentcode.Visible = True
            txtPaymentMode.Visible = True
            pnlCheque.Visible = True
            txtBankCharges.Visible = True
            MyLabel3.Visible = True
            ''
            ChkSecurity.Visible = False
            ChkRetention.Visible = False
            gvDetails.Visible = True
            btnViewTDSDetails.Enabled = False
            chkCheckPrint.Visible = True
            btnOk.Visible = False
            If autoadjustdebitcreditnote = True Then
                txtmulPI.Visible = True
                MyLabel15.Visible = True
            Else
                txtmulPI.Visible = False
                MyLabel15.Visible = False
            End If
        End If
    End Sub
    Private Sub ddlPaymentType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.Data.PositionChangedEventArgs) Handles ddlPaymentType.SelectedIndexChanged
        If IsPaymentTypeChanged = True Then
            pnlAdvance.Visible = False
            pnlMiscPayment.Visible = False
            pnlVendor.Visible = True
            txtVendorCode.Value = ""
            lblVendorName.Text = ""
            lblLoadOutNo.Visible = False
            txtLoadOutno.Visible = False
            lblMPAdv.Visible = False
            txtMPAdv.Visible = False
            chkIsReceipt.Visible = False
            txtBankCharges.Text = ""
            txtBankCharges.Enabled = True
            pnlCform.Visible = False

            pnlPJC.Visible = False
            pnlmemorndm.Visible = False
            lblDocumentNo.Visible = False
            txtDocumentNo.Visible = False
            lblBalAmt.Visible = False
            '' Anubhooti 21-Aug-2014
            LblPONo.Visible = False
            txtPONo.Visible = False
            btnOk.Visible = False
            RadLabel18.Visible = False
            txtlocation.Value = ""
            LblLocDesp.Text = ""
            txtlocation.Visible = False
            LblLocDesp.Visible = False
            pnlEMI.Visible = False

            LoadBlankGrid(ddlPaymentType.SelectedValue)
            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
                lblTotPayment.Text = "Total Payment"
                lblpaymentcode.Visible = True
                txtPaymentMode.Visible = True
                pnlCheque.Visible = True
                txtBankCharges.Visible = True
                MyLabel3.Visible = True
                ChkSecurity.Visible = False
                ChkRetention.Visible = False
                gvDetails.Visible = True
                btnViewTDSDetails.Enabled = False
                chkCheckPrint.Visible = True
                chkSaving.Visible = True
                btnOk.Visible = False
                If autoadjustdebitcreditnote = True Then
                    txtmulPI.Visible = True
                    MyLabel15.Visible = True
                End If

            ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                lblTotPayment.Text = "Total Payment"
                txtPaymentAmt.ReadOnly = True
                gvDetails.Visible = True
                btnViewTDSDetails.Enabled = False
                chkCheckPrint.Visible = True
                lblDocumentNo.Visible = True
                txtDocumentNo.Visible = True
                lblBalAmt.Visible = True
                ChkSecurity.Visible = False
                ChkRetention.Visible = False
                lblpaymentcode.Visible = False
                txtPaymentMode.Visible = False
                pnlCheque.Visible = False
                txtBankCharges.Visible = False
                MyLabel3.Visible = False
                txtPaymentMode.Value = ""
                txtBankCharges.Text = "0"
                ChkAdvSalary.Visible = False
                ChkAdvSalary.Checked = False
                ' If autoadjustdebitcreditnote = True Then
                txtmulPI.Visible = False
                'End If

                MyLabel15.Visible = False
            ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                pnlAdvance.Visible = True
                lblTotPayment.Text = "Payment Amount"
                lblpaymentcode.Visible = True
                txtPaymentMode.Visible = True
                pnlCheque.Visible = True
                txtBankCharges.Visible = True
                MyLabel3.Visible = True
                ChkSecurity.Visible = False
                ChkRetention.Visible = False
                txtPaymentAmt.ReadOnly = False
                gvDetails.Visible = False
                btnViewTDSDetails.Enabled = False
                ChkAdvSalary.Visible = True
                pnlEMI.Visible = True
                chkSaving.Visible = False

                txtmulPI.Visible = False

                MyLabel15.Visible = False
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
                    pnlmemorndm.Visible = True
                    lblpaymentcode.Visible = True
                    txtPaymentMode.Visible = True
                    pnlCheque.Visible = True
                    txtBankCharges.Visible = True
                    MyLabel3.Visible = True
                    ChkSecurity.Visible = False
                    ChkRetention.Visible = False
                    LblPONo.Visible = True
                    txtPONo.Visible = True
                    txtmulPI.Visible = False

                    MyLabel15.Visible = False
                Else
                    chkSaving.Visible = True
                End If
                If objCommonVar.IsDemoERP = True Then
                    pnlCform.Visible = True
                Else
                    pnlCform.Visible = False
                End If
                chkCheckPrint.Visible = True
                If isApplyBranchAccounting = True Then
                    RadLabel18.Visible = True
                    txtlocation.Visible = True
                    LblLocDesp.Visible = True
                    txtlocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code in (select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' ) "))
                    If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
                        LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
                    Else
                        LblLocDesp.Text = ""
                    End If
                Else
                    RadLabel18.Visible = False
                    txtlocation.Visible = False
                    LblLocDesp.Visible = False
                    txtlocation.Value = ""
                    LblLocDesp.Text = ""
                End If

            ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
                lblTotPayment.Text = "Payment Amount"
                lblpaymentcode.Visible = True
                txtPaymentMode.Visible = True
                pnlCheque.Visible = True
                txtBankCharges.Visible = True
                MyLabel3.Visible = True
                txtPaymentAmt.ReadOnly = False
                gvDetails.Visible = False
                btnViewTDSDetails.Enabled = False
                txtBankCharges.Enabled = False
                chkCheckPrint.Checked = False
                chkCheckPrint.Visible = False
                ChkSecurity.Visible = True
                ChkRetention.Visible = True
                ChkAdvSalary.Visible = True
                ChkAdvSalary.Checked = False
                txtmulPI.Visible = False

                MyLabel15.Visible = False
                If isApplyBranchAccounting = True Then
                    RadLabel18.Visible = True
                    txtlocation.Visible = True
                    LblLocDesp.Visible = True
                    txtlocation.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Loc_Segment_Code from TSPL_LOCATION_MASTER where Location_Code in (select Default_Location from TSPL_USER_MASTER where User_Code='" + objCommonVar.CurrentUserCode + "' ) "))
                    If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
                        LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
                    Else
                        LblLocDesp.Text = ""
                    End If
                Else
                    RadLabel18.Visible = False
                    txtlocation.Visible = False
                    LblLocDesp.Visible = False
                    txtlocation.Value = ""
                    LblLocDesp.Text = ""
                End If

            ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
                pnlMiscPayment.Visible = True
                lblTotPayment.Text = " Total Amount"
                lblpaymentcode.Visible = True
                txtPaymentMode.Visible = True
                pnlCheque.Visible = True
                txtBankCharges.Visible = True
                MyLabel3.Visible = True

                txtmulPI.Visible = False

                MyLabel15.Visible = False
                txtPaymentAmt.ReadOnly = False
                gvDetails.Visible = True
                btnViewTDSDetails.Enabled = False
                pnlVendor.Visible = False
                lblLoadOutNo.Visible = True
                txtLoadOutno.Visible = True
                lblMPAdv.Visible = True
                txtMPAdv.Visible = True
                chkIsReceipt.Visible = True
                If objCommonVar.IsDemoERP = True Then
                    pnlPJC.Visible = True
                Else
                    pnlPJC.Visible = False
                End If
                chkCheckPrint.Visible = True
                ChkSecurity.Visible = False
                ChkRetention.Visible = False
                ChkAdvSalary.Visible = False
                ChkAdvSalary.Checked = False
                txtBankCharges.Enabled = False
            ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MD") = CompairStringResult.Equal Then
                txtPaymentAmt.ReadOnly = False
                lblTotPayment.Text = "Payment Amount"
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "MI") = CompairStringResult.Equal Then
            gvDetails.Rows.AddNew()
            ChkSecurity.Visible = False
            ChkRetention.Visible = False
        End If
        If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "SR") = CompairStringResult.Equal Then
            gvDetails.Visible = True
            gvDetails.Rows.AddNew()
            ChkSecurity.Visible = True
            ChkRetention.Visible = True
            txtPaymentAmt.ReadOnly = True
            lblTotPayment.Text = "Total Payment"
            LoadPaymentEntries(txtVendorCode.Value, dtpPayment.Value, False)
        End If


        GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
        If GSTStatus Then
            If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "AV") = CompairStringResult.Equal Then
                RadPageView1.Pages("TabForGST").Item.Visibility = ElementVisibility.Visible
            Else
                RadPageView1.Pages("TabForGST").Item.Visibility = ElementVisibility.Collapsed
            End If
        End If
        If ApplyBankChargesasperSlabonBankMaster Then
            txtBankCharges.Enabled = False
            chkBankChargesWaveOff.Enabled = True
        Else
            txtBankCharges.Enabled = True
            chkBankChargesWaveOff.Enabled = False
        End If
    End Sub

#Region "Finders"
    Private Sub txtBankCode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtBankCode._MYValidating
        Dim strWhrclas As String = ""
        Dim strWhereClause As String = ""
        Qry = clsERPFuncationality.glbankqueryNew(strWhrclas)
        If isNewEntry = False Then
            strWhrclas += "  and Bank_type=(Select Bank_type from TSPL_BANK_MASTER WHERE BANK_CODE=(Select Bank_Code from TSPL_PAYMENT_HEADER WHERE Payment_No='" & txtPaymentNo.Value & "')) AND RIGHT(BANKACC,3)=(Select RIGHT(BANKACC,3) from TSPL_BANK_MASTER WHERE BANK_CODE=(Select Bank_Code from TSPL_PAYMENT_HEADER WHERE Payment_No='" & txtPaymentNo.Value & "'))"
        End If
        If isSettlementBankOnly Then
            strWhrclas += " and TSPL_BANK_MASTER.bank_type='S'"
            strWhereClause += "  TSPL_BANK_MASTER.bank_type='S'"
        Else
            strWhrclas += "  and TSPL_BANK_MASTER.bank_type<>'S'"
            strWhereClause += "   TSPL_BANK_MASTER.bank_type<>'S'"
        End If
        If objCommonVar.RCDFCFP = True Then
        Else
            'strWhrclas += " and User_Code='" + objCommonVar.CurrentUserCode + "' "
            strWhereClause += " and User_Code='" + objCommonVar.CurrentUserCode + "' "
        End If
        strWhrclas += " and TSPL_bank_master.INACTIVE ='Active' "
        strWhereClause += " and TSPL_bank_master.INACTIVE ='Active' "

        Dim query As String = Qry & "where " & strWhereClause
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(query)
        If dt.Rows.Count > 0 Then
            txtBankCode.Value = clsCommon.ShowSelectForm("BankSlctr@Payment", Qry, "Code", strWhereClause, txtBankCode.Value, "Code", isButtonClicked)
            lblBankDesc.Text = connectSql.RunScalar("select description from TSPL_BANK_MASTER where bank_code = '" + txtBankCode.Value + "'")

        Else
            Dim qrys As String = "select BANK_CODE as [Code], DESCRIPTION,BANKACCNUMBER as [Bank Account No]  from TSPL_BANK_MASTER "
            txtBankCode.Value = clsCommon.ShowSelectForm("BankSlctr@Payment", qrys, "Code", strWhrclas, txtBankCode.Value, "Code", isButtonClicked)
            lblBankDesc.Text = connectSql.RunScalar("select description from TSPL_BANK_MASTER where bank_code = '" + txtBankCode.Value + "'")

        End If
        txtPaymentMode.Value = connectSql.RunScalar("select TSPL_PAYMENT_CODE.Payment_Code   from TSPL_PAYMENT_CODE Where TSPL_PAYMENT_CODE.Payment_Code=  (select DISTINCT (case when Bank_type = 'C' THEN 'CASH' WHEN BANK_TYPE = 'B' THEN 'CHEQUE' WHEN BANK_TYPE = 'O' THEN 'OTHER' WHEN Bank_type = 'P' THEN 'PETTYCASH' END ) AS [Paymet Type] from TSPL_BANK_MASTER Where BANK_CODE='" + txtBankCode.Value + "' )")
        HandleCheque()

        ''richa 05 feb,2019  TEC/05/02/19-000412 check for opening in case of Miscellenous
        'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, Nothing)
        Dim JEWithOPening As Boolean = False
        If clsCommon.myLen(clsCommon.GetDateWithStartTime(ERPStartDate)) > 0 Then
            If clsCommon.GetDateWithStartTime(dtpPayment.Value) < clsCommon.GetDateWithStartTime(ERPStartDate) Then
                JEWithOPening = True
            End If
        End If
        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, Nothing)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "MI") = CompairStringResult.Equal) And JEWithOPening = True Then
            LoadBlankGrid(ddlPaymentType.SelectedValue)
            gvDetails.Rows.AddNew()
        End If

    End Sub

    Private Sub txtPaymentMode__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtPaymentMode._MYValidating
        Dim strbankcode As String
        If Not String.IsNullOrEmpty(connectSql.RunScalar("select bank_type from tspl_bank_master where bank_code = '" + txtBankCode.Value + "'")) Then
            strbankcode = connectSql.RunScalar("select bank_type from tspl_bank_master where bank_code = '" + txtBankCode.Value + "'")
            If strbankcode.Trim() = "C" Then
                Dim Qry1 As String = "select Payment_Code as [PaymentMode], Payment_Desc as [Description], Payment_Type  as [PaymentType]  from TSPL_PAYMENT_CODE "
                txtPaymentMode.Value = clsCommon.ShowSelectForm("PaymentCode Selector1", Qry1, "PaymentMode", "PAYMENT_TYPE = 'CASH'", txtPaymentMode.Value, "PaymentMode", isButtonClicked)
            ElseIf strbankcode.Trim() = "P" Then
                Dim Qry1 As String = "select Payment_Code as [PaymentMode], Payment_Desc as [Description], Payment_Type  as [PaymentType]  from TSPL_PAYMENT_CODE "
                txtPaymentMode.Value = clsCommon.ShowSelectForm("PaymentCode Selector2", Qry1, "PaymentMode", "PAYMENT_TYPE = 'Petty Cash'", txtPaymentMode.Value, "PaymentMode", isButtonClicked)
                '' Anubhooti 08-Sep-2014 (Added 'NEFT','RTGS' with 'Cheque' 'Other' as discussed with mam)
            ElseIf strbankcode = "B" Then
                Dim Qry1 As String = "select Payment_Code as [PaymentMode], Payment_Desc as [Description], Payment_Type  as [PaymentType]  from TSPL_PAYMENT_CODE "
                txtPaymentMode.Value = clsCommon.ShowSelectForm("PaymentCode Selector3", Qry1, "PaymentMode", "PAYMENT_TYPE IN ('Cheque', 'Other','NEFT','RTGS')", txtPaymentMode.Value, "PaymentMode", isButtonClicked)
            Else
                Dim Qry1 As String = "select Payment_Code as [PaymentMode], Payment_Desc as [Description], Payment_Type  as [PaymentType]  from TSPL_PAYMENT_CODE "
                txtPaymentMode.Value = clsCommon.ShowSelectForm("PaymentCode Selector4", Qry1, "PaymentMode", "PAYMENT_TYPE = 'Other'", txtPaymentMode.Value, "PaymentMode", isButtonClicked)
            End If
            HandleCheque()
        End If
    End Sub

    Public Sub HandleCheque()
        If clsCommon.CompairString(txtPaymentMode.Value, "Cheque") = CompairStringResult.Equal Or clsCommon.CompairString(txtPaymentMode.Value, "DD") = CompairStringResult.Equal Then
            pnlCheque.Visible = True
            txtChequeNo.Text = ""
            dtpChequeDate.Value = clsCommon.GETSERVERDATE
            txtChequeNo.Enabled = True
            dtpChequeDate.Enabled = True
        Else
            pnlCheque.Visible = False
            txtChequeNo.Text = ""
            dtpChequeDate.Value = Nothing
            txtChequeNo.Enabled = False
            dtpChequeDate.Enabled = False
        End If
    End Sub

    Private Sub txtVendorCode__MYValidating_1(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtVendorCode._MYValidating
        Dim VSPAccPayee As String = ""
        Dim Qry As String = clsERPFuncationality.glvendorqueryNew
        Dim Whrcls As String = "  m.Status='N' "
        If chkSaving.Checked Then
            Whrcls += " and m.Form_Type='VSP' "
        End If
        '' and m.Is_Inactive_In_Milk_Procurement=0 removed by Panch Raj on behalf of Balwinder Sir
        txtVendorCode.Value = clsCommon.ShowSelectForm("VNDRSlctr@Payment", Qry, "Code", Whrcls, txtVendorCode.Value, "Code", isButtonClicked)
        lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Name from TSPL_VENDOR_MASTER WHERE Vendor_Code='" + txtVendorCode.Value + "'"))
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
            If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Is_Provisional from TSPL_VENDOR_MASTER where Vendor_Code='" + txtVendorCode.Value + "'")) = 1 Then
                If common.clsCommon.MyMessageBoxShow("Do You Want to Apply TDS Provision", "", MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes Then
                    chkTDSProvision.Checked = True
                Else
                    chkTDSProvision.Checked = False
                End If
            End If
        End If

        '' Tax on Bank Charges
        If clsCommon.myCdbl(txtBankCharges.Text) > 0 Then
            txtChangeVendorNo()
        End If
        ''richa agarwal 25 June 2020 , when Vendor code finder is used and document no is filled then if we change vendor then document no will also be change
        txtDocumentNo.Value = ""
        If (clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal) Then
            LoadVendorInvoices(txtVendorCode.Value)
        ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MD") = CompairStringResult.Equal Then
            LoadVendorInvoiceOfManualDeduction(txtVendorCode.Value)
        ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "SR") = CompairStringResult.Equal Then
            LoadPaymentEntries(txtVendorCode.Value, dtpPayment.Value, False)
        ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then  ' Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
            Dim objVendor As clsTDSVendorDetails = clsTDSVendorDetails.GetData(txtVendorCode.Value)
            If objVendor IsNot Nothing Then
                btnViewTDSDetails.Enabled = True
            Else
                btnViewTDSDetails.Enabled = False
            End If
            SetVendorTDSDetails()
        End If
        SetMultiCurrencyVisibility()
        '' Anubhooti 25-Sep-2014 BM00000004050
        VSPAccPayee = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT CASE WHEN vsp_payment ='Self' THEN Vendor_Name WHEN vsp_payment ='Different' THEN VSP_Payee_Name END As Payee_Name FROM TSPL_VENDOR_MASTER WHERE Vendor_Code ='" & clsCommon.myCstr(txtVendorCode.Value) & "' And ISNULL(CASE WHEN vsp_payment ='Self' THEN Vendor_Name WHEN vsp_payment ='Different' THEN VSP_Payee_Name END,'') <> ISNULL(Vendor_Name ,'')"))
        If clsCommon.myLen(VSPAccPayee) > 0 Then
            LblAccPayee.Text = VSPAccPayee
        Else
            LblAccPayee.Text = ""
        End If

        ''richa 13 oct,2017 work related to employee salary integration
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                ddlEmployeeType.Enabled = True
                'loadEmployeeType()
                ddlEmployeeAdvanceType.Enabled = False
                'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
            Else
                ddlEmployeeType.Enabled = False
                'loadEmployeeType()
                ddlEmployeeAdvanceType.Enabled = False
                'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
            End If
        ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                ddlEmployeeAdvanceType.Enabled = True
                'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
                ddlEmployeeType.Enabled = False
                'loadEmployeeType()
            Else
                ddlEmployeeAdvanceType.Enabled = False
                'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
                ddlEmployeeType.Enabled = False
                'loadEmployeeType()
            End If
        ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                ddlEmployeeAdvanceType.Enabled = True
                'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
                ddlEmployeeType.Enabled = True
                'loadEmployeeType()
            Else
                ddlEmployeeAdvanceType.Enabled = False
                'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
                ddlEmployeeType.Enabled = False
                'loadEmployeeType()
            End If
        Else
            ddlEmployeeType.Enabled = False
            'loadEmployeeType()
            ddlEmployeeAdvanceType.Enabled = False
            'loadEmployeeAdvanceType()''Balwinder on 28/06/2022
        End If
        ''------------------

        Qry = "select  Bank_Name,Bank_Code ,ifsc_code,Branch_Name,Account_No from TSPL_VENDOR_MASTER where Vendor_Code ='" + clsCommon.myCstr(txtVendorCode.Value) + "'"
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            txtVendor_bankcode.Text = clsCommon.myCstr(dt.Rows(0)("Bank_Code"))
            TxtVendor_BankName.Text = clsCommon.myCstr(dt.Rows(0)("Bank_Name"))
            TxtVendorBank_IFSCCode.Text = clsCommon.myCstr(dt.Rows(0)("ifsc_code"))
            txtVendorBank_branchname.Text = clsCommon.myCstr(dt.Rows(0)("Branch_Name"))
            txtVendor_Bank_ACNo.Text = clsCommon.myCstr(dt.Rows(0)("Account_No"))
        Else
            txtVendor_bankcode.Text = ""
            TxtVendor_BankName.Text = ""
            TxtVendorBank_IFSCCode.Text = ""
            txtVendorBank_branchname.Text = ""
            txtVendor_Bank_ACNo.Text = ""
        End If

    End Sub
#End Region

    Private Sub FillCustomerOutStanding(ByVal strVendorCode As String)
        Try
            Arr = New ArrayList
            Arr.Add("I")
            Arr.Add("C")
            Arr.Add("D")
            Arr.Add("AV")
            Arr.Add("OA")
            Arr.Add("P")
            Arr.Add("RC")
            Qry = clsCustomerMasterNew.GetOutStandingQry(dtpPayment.Value, dtpPayment.Value, " AND TSPL_CUSTOMER_MASTER.Status='N'", Arr, "DocumentDate", "ConvRate")
            Qry = "Select SUM([Due Amount]) from (" & Qry & ") ZZZ WHERE [Customer Id]=(Select Top 1 Cust_Code from TSPL_CUSTOMER_VENDOR_MAPPING WHERE Vendor_Code='" + strVendorCode + "')"
            lblCustomerOutStanding.Text = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Sub SetVendorTDSDetails()
        btnViewTDSDetails.Enabled = False
        objRemittance = Nothing
        Dim objVendor As clsTDSVendorDetails = clsTDSVendorDetails.GetData(txtVendorCode.Value)

        If objVendor IsNot Nothing Then
            ''richa agarwal 29/07/2015
            Dim ISOverride As Integer = 0
            If clsCommon.myLen(txtPaymentNo.Value) > 0 Then
                objVendor.Nature_Of_Deduction = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Deduction_Code  from TSPL_REMITTANCE where Document_No ='" & txtPaymentNo.Value & "' "))
            End If
            If clsCommon.myLen(clsCommon.myCstr(objVendor.Nature_Of_Deduction)) <= 0 Then
                objVendor.Nature_Of_Deduction = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select ISNULL(Deduction_Code,'')  from tspl_vendor_master where Vendor_Code ='" & txtVendorCode.Value & "'"))
            End If
            btnViewTDSDetails.Enabled = True
            Dim objDedDetails As clsTDSDeductionDetails = clsTDSDeductionDetails.GetApplicableTDRate(objVendor.Nature_Of_Deduction, clsCommon.myCdbl(txtPaymentAmt.Text), Nothing, False, txtVendorCode.Value)
            If (objDedDetails IsNot Nothing) Then
                objRemittance = New clsRemittance()
                objRemittance.Branch_Code = objVendor.Branch_Code
                objRemittance.Deduction_Code = objVendor.Nature_Of_Deduction
                objRemittance.TDS_Per = objDedDetails.TDS
                objRemittance.Surcharge_Per = objDedDetails.Surcharge
                objRemittance.Edu_Cess_Per = objDedDetails.Educess
                objRemittance.Sec_Educess_Per = objDedDetails.Seceducess
                objRemittance.Section_Code = objVendor.TDSSection
                objRemittance.Section_Description = objVendor.TDSSectionDescription
                objRemittance.Select_By = objVendor.VendorTypeCode
                Dim dt As DataTable = clsDBFuncationality.GetDataTable("Select case when Actual_TDS_Base=Document_Amount then 0 else 1  end as isoverride,Actual_TDS,Actual_TDS_Base ,Actual_Total_TDS  from TSPL_REMITTANCE where Document_No ='" & txtPaymentNo.Value & "' ")
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    objRemittance.Actual_TDS_Base = clsCommon.myCdbl(dt.Rows(0)("Actual_TDS_Base"))
                    objRemittance.Actual_TDS = clsCommon.myCdbl(dt.Rows(0)("Actual_TDS"))
                    objRemittance.Actual_Total_TDS = clsCommon.myCdbl(dt.Rows(0)("Actual_Total_TDS"))
                    ISOverride = clsCommon.myCdbl(dt.Rows(0)("isoverride"))
                End If
                If ISOverride = 1 Then
                    objRemittance.IsTDSOverride = True
                Else
                    objRemittance.IsTDSOverride = False
                End If
                If isNewEntry Then
                    objRemittance.IsApplyTDS = True
                Else
                    objRemittance.IsApplyTDS = clsPIRemittance.IsTDSApplied(txtPaymentNo.Value)
                End If
                Dim qry As String = "select Year_Name from TSPL_TDS_FINANCIAL_YEAR where convert(date,'" + dtpPayment.Value + "',103)>=  convert(date,From_Date,103)  and convert(date,'" + dtpPayment.Value + "',103)<=convert(date,To_Date,103) "
                objRemittance.Fiscal_Year = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry))
                objRemittance.Quarter = "First"

            End If
        End If
    End Sub

#Region "Grid Functionality"

    Sub LoadBlankGrid(ByVal paymentType As String)
        gvDetails.Rows.Clear()
        gvDetails.Columns.Clear()

        If (clsCommon.CompairString(paymentType, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(paymentType, "AD") = CompairStringResult.Equal Or clsCommon.CompairString(paymentType, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(paymentType, "MD") = CompairStringResult.Equal) Then
            gvDetails.AllowDeleteRow = True
            gvDetails.AllowAddNewRow = False

            Dim apply As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            apply.FormatString = ""
            apply.HeaderText = colApply
            apply.Name = colApply
            apply.Width = 50
            apply.ReadOnly = True
            apply.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
            gvDetails.MasterTemplate.Columns.Add(apply)

            Dim docType As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            docType.FormatString = ""
            docType.HeaderText = "Document Type"
            docType.Name = colDocType
            docType.Width = 100
            docType.ReadOnly = True
            gvDetails.MasterTemplate.Columns.Add(docType)

            Dim PINo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            PINo.FormatString = ""
            PINo.HeaderText = "Document No"
            PINo.Name = colPINo
            PINo.Width = 150
            PINo.ReadOnly = True
            gvDetails.MasterTemplate.Columns.Add(PINo)

            Dim docNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            docNo.FormatString = ""
            docNo.HeaderText = "AP Document No"
            docNo.Name = colDocNo
            docNo.Width = 150
            docNo.ReadOnly = True
            docNo.IsVisible = False
            gvDetails.MasterTemplate.Columns.Add(docNo)



            Dim documentDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            documentDate.FormatString = ""
            documentDate.HeaderText = "Document Date"
            documentDate.Name = colDocumentDate
            documentDate.Width = 150
            documentDate.ReadOnly = True
            gvDetails.MasterTemplate.Columns.Add(documentDate)

            Dim APTapalNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            APTapalNo.FormatString = ""
            APTapalNo.HeaderText = "Tapal No"
            APTapalNo.Name = colAPTapalNo
            APTapalNo.Width = 150
            APTapalNo.ReadOnly = True
            APTapalNo.IsVisible = True
            gvDetails.MasterTemplate.Columns.Add(APTapalNo)

            Dim docDate As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            docDate.FormatString = ""
            docDate.HeaderText = "Vendor Invoice Date"
            docDate.Name = colDocDate
            docDate.Width = 150
            docDate.ReadOnly = True
            gvDetails.MasterTemplate.Columns.Add(docDate)

            Dim vendorInvNo As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            vendorInvNo.FormatString = ""
            vendorInvNo.HeaderText = "Vendor Invoice No"
            vendorInvNo.Name = colVendorInvNo
            vendorInvNo.Width = 100
            vendorInvNo.ReadOnly = True
            gvDetails.MasterTemplate.Columns.Add(vendorInvNo)

            Dim originalInvAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            originalInvAmt = New GridViewDecimalColumn()
            originalInvAmt.FormatString = ""
            originalInvAmt.HeaderText = "Original Inv Amt"
            originalInvAmt.Name = colOriginalAmt
            originalInvAmt.Width = 100
            originalInvAmt.ReadOnly = True
            originalInvAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(originalInvAmt)

            Dim tdsAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            tdsAmt = New GridViewDecimalColumn()
            tdsAmt.FormatString = ""
            tdsAmt.HeaderText = "TDS Amount"
            tdsAmt.Name = colTDSAmt
            tdsAmt.Width = 100
            tdsAmt.ReadOnly = True
            tdsAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(tdsAmt)

            Dim payableAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            payableAmt = New GridViewDecimalColumn()
            payableAmt.FormatString = ""
            payableAmt.HeaderText = "Net Amount"
            payableAmt.Name = colNetAmt
            payableAmt.Width = 100
            payableAmt.ReadOnly = True
            payableAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(payableAmt)

            Dim SecurityAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            SecurityAmt = New GridViewDecimalColumn()
            SecurityAmt.FormatString = ""
            SecurityAmt.HeaderText = "Security Amount"
            SecurityAmt.Name = colSecurityAmt
            SecurityAmt.Width = 0
            SecurityAmt.IsVisible = False
            SecurityAmt.ReadOnly = False
            SecurityAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(SecurityAmt)

            Dim pendingAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            pendingAmt = New GridViewDecimalColumn()
            pendingAmt.FormatString = ""
            pendingAmt.HeaderText = "Current Pending"
            pendingAmt.Name = colPendingAmt
            pendingAmt.Width = 100
            pendingAmt.ReadOnly = True
            pendingAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(pendingAmt)

            Dim appliedAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            appliedAmt = New GridViewDecimalColumn()
            appliedAmt.FormatString = ""
            appliedAmt.HeaderText = "Applied Amount"
            appliedAmt.Name = colAppliedAmt
            appliedAmt.Width = 100
            appliedAmt.ReadOnly = False
            appliedAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(appliedAmt)

            '' Anubhooti 14-Nov-2014
            Dim AdjustedAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
            AdjustedAmt = New GridViewDecimalColumn()
            AdjustedAmt.FormatString = ""
            AdjustedAmt.HeaderText = "Adjusted/Paid Amount"
            AdjustedAmt.Name = colAdjustedAmt
            AdjustedAmt.Width = 100
            AdjustedAmt.ReadOnly = True
            AdjustedAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(AdjustedAmt)
            ''
            Dim temp As GridViewDecimalColumn = New GridViewDecimalColumn()
            temp = New GridViewDecimalColumn()
            temp.FormatString = ""
            temp.HeaderText = "Temporary"
            temp.Name = colTemp
            temp.Width = 100
            temp.ReadOnly = True
            temp.IsVisible = False
            temp.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(temp)

            Dim comment As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            comment.FormatString = ""
            comment.HeaderText = "Comment"
            comment.Name = colComment
            comment.Width = 150
            comment.ReadOnly = False
            gvDetails.MasterTemplate.Columns.Add(comment)
        ElseIf clsCommon.CompairString(paymentType, "MI") = CompairStringResult.Equal Then
            gvDetails.AllowDeleteRow = True
            gvDetails.AllowAddNewRow = True
            gvDetails.AddNewRowPosition = SystemRowPosition.Bottom

            Dim LineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
            LineNo.FormatString = ""
            LineNo.HeaderText = "Line No"
            LineNo.Name = colLineNo
            LineNo.Width = 60
            LineNo.ReadOnly = True
            LineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
            gvDetails.MasterTemplate.Columns.Add(LineNo)

            Dim GLAccount As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            GLAccount.FormatString = ""
            GLAccount.HeaderText = "GL Account"
            GLAccount.Name = colGLAccount
            GLAccount.Width = 200
            GLAccount.ReadOnly = False
            gvDetails.MasterTemplate.Columns.Add(GLAccount)

            Dim AccDesc As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            AccDesc.FormatString = ""
            AccDesc.HeaderText = "Account Description"
            AccDesc.Name = colAccDesc
            AccDesc.Width = 350
            AccDesc.ReadOnly = True
            gvDetails.MasterTemplate.Columns.Add(AccDesc)

            Dim GLType As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            GLType = New GridViewTextBoxColumn()
            GLType.FormatString = ""
            GLType.HeaderText = "GL Type"
            GLType.Name = colGLType
            GLType.Width = 100
            GLType.ReadOnly = True
            GLType.IsVisible = False
            gvDetails.MasterTemplate.Columns.Add(GLType)

            GLAccount = New GridViewTextBoxColumn()
            GLAccount.FormatString = ""
            GLAccount.HeaderText = "Hierarchy Level"
            GLAccount.Name = colHirerachyCenter
            GLAccount.Width = 100
            GLAccount.ReadOnly = False
            GLAccount.IsVisible = isApplyCostCenter
            gvDetails.MasterTemplate.Columns.Add(GLAccount)

            AccDesc = New GridViewTextBoxColumn()
            AccDesc.FormatString = ""
            AccDesc.HeaderText = "Hierarchy Level Description"
            AccDesc.Name = colHirerachyName
            AccDesc.Width = 200
            AccDesc.ReadOnly = True
            AccDesc.IsVisible = isApplyCostCenter
            gvDetails.MasterTemplate.Columns.Add(AccDesc)

            GLAccount = New GridViewTextBoxColumn()
            GLAccount.FormatString = ""
            GLAccount.HeaderText = "Cost Center Code"
            GLAccount.Name = colCostCenter
            GLAccount.Width = 100
            GLAccount.ReadOnly = False
            GLAccount.IsVisible = isApplyCostCenter
            gvDetails.MasterTemplate.Columns.Add(GLAccount)

            AccDesc = New GridViewTextBoxColumn()
            AccDesc.FormatString = ""
            AccDesc.HeaderText = "Cost Center Description"
            AccDesc.Name = colCostCenterName
            AccDesc.Width = 200
            AccDesc.ReadOnly = True
            AccDesc.IsVisible = isApplyCostCenter
            gvDetails.MasterTemplate.Columns.Add(AccDesc)

            Dim Amount As GridViewDecimalColumn = New GridViewDecimalColumn()
            Amount = New GridViewDecimalColumn()
            Amount.FormatString = ""
            Amount.HeaderText = "Amount"
            Amount.Name = colAmount
            Amount.Width = 200
            Amount.ReadOnly = False
            Amount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
            gvDetails.MasterTemplate.Columns.Add(Amount)

            Dim remark As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            remark.FormatString = ""
            remark.HeaderText = "Paid To"
            remark.Name = colRemark
            remark.Width = 200
            remark.ReadOnly = False
            gvDetails.MasterTemplate.Columns.Add(remark)

            Dim ExpenseCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
            ExpenseCode.FormatString = ""
            ExpenseCode.HeaderText = "ExpenseCode"
            ExpenseCode.Name = colExpenseCode
            ExpenseCode.Width = 200
            ExpenseCode.ReadOnly = False
            gvDetails.MasterTemplate.Columns.Add(ExpenseCode)

        End If
        clsCustomFieldGrid.LoadBlankGrid(gvDetails, MyBase.ArrDetailFields)
        gvDetails.ShowGroupPanel = False
        gvDetails.AllowColumnReorder = False
        gvDetails.AllowRowReorder = False
        gvDetails.EnableSorting = False
        gvDetails.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvDetails.MasterTemplate.ShowRowHeaderColumn = False
        gvDetails.AllowAddNewRow = False
    End Sub
    Private Sub LoadVendorInvoices(ByVal strVendorCode As String)
        LoadVendorInvoices(strVendorCode, dtpPayment.Value.Date, False)
    End Sub

    Private Sub LoadVendorInvoiceOfManualDeduction(ByVal strVendorCode As String)
        LoadVendorInvoiceOfManualDeduction(strVendorCode, dtpPayment.Value.Date, False)
    End Sub

    Private Sub LoadVendorInvoiceOfManualDeduction(ByVal strVendorCode As String, ByVal InvoiceDate As Date, ByVal isForUpdate As Boolean)
        Try
            If Not isForUpdate Then
                LoadBlankGrid(ddlPaymentType.SelectedValue)
            End If

            Dim strQryForRejectedAmt As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
           "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No ) and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "

            '' this code is wrritten for that debit note which is created auto through PI and Pr is not created against that PI
            Dim strQryForRejectedAmtforNonPR As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
          "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "


            Dim strTaxRecovarableQuery As String = " - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then " &
            " ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'  " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end) "

            If clsCommon.myLen(strVendorCode) > 0 Then
                Qry = "Select * from ( select TSPL_VENDOR_INVOICE_HEAD.ISProcurementDeduction,Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType],  " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_No, Case When ISNULL(Against_POInvoice_No,'')<>'' Then Against_POInvoice_No When ISNULL(Against_PurchaseReturn_No,'')<> '' Then Against_PurchaseReturn_No Else TSPL_VENDOR_INVOICE_HEAD.Document_No End as PurchaseInvoice," &
           " Case When ISNULL(Against_POInvoice_No,'')<>'' Then (Select convert(varchar,PI_Date,103)  from TSPL_PI_HEAD where PI_No =Against_POInvoice_No)  When ISNULL(Against_PurchaseReturn_No,'')<> '' Then (Select convert(varchar,PR_Date,103) from TSPL_PR_HEAD where PR_No  =Against_PurchaseReturn_No ) Else TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date End as DocumentDate, " &
           " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date as [DocDate] ,  " &
           " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No as [VendorInvoiceNo], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " as [OriginalAmt]  ," &
           " case when isnull(TSPL_REMITTANCE.Is_TDS_Provision,'N')='N' then TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount else 0 end as [TDSAmt], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total-isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) as [NetAmount], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total-isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) as [PendingAmt],TSPL_VENDOR_INVOICE_HEAD.ConvRate " &
           " ,isnull(( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and TSPL_REVALUATION_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation " &
           " from TSPL_VENDOR_INVOICE_HEAD " &
           " Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  
left outer join TSPL_REMITTANCE on TSPL_REMITTANCE.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
            WHERE TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,Invoice_Entry_Date,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' And ((ISNULL(RefDocNo,'')= '' or ISNULL((select VI.Document_Type from TSPL_VENDOR_INVOICE_HEAD VI where VI.Document_No =TSPL_VENDOR_INVOICE_HEAD .RefDocNo ),'') in ('I','')) AND ISNULL(Against_PurchaseReturn_No,'')= '')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D') " &
          " and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_Salary_Generation_Code,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='C') "
                If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                    Qry += "  and TSPL_VENDOR_INVOICE_HEAD.Loc_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
                End If

                If isForUpdate Then
                    Dim arrDocNo As New List(Of String)
                    For ii As Integer = 0 To gvDetails.Rows.Count - 1
                        arrDocNo.Add(clsCommon.myCstr(gvDetails.Rows(ii).Cells(colDocNo).Value))
                    Next
                    If arrDocNo.Count > 0 Then
                        Qry += " and TSPL_VENDOR_INVOICE_HEAD.Document_No not in (" + clsCommon.GetMulcallString(arrDocNo) + ")"
                    End If
                End If
                Qry += " and TSPL_VENDOR_INVOICE_HEAD.RefDocType<>'REVALUATION ENTRY'"

                ''richa 13/10/2017
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "TD") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='TD' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "T") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='T' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "S") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='S' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "I") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='I' "
                    Else
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='' "
                    End If
                End If
                ''--------
                Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.posting_date,'')<>''" & Environment.NewLine
                If chkSaving.Checked Then
                    Qry += " and ISNULL(TSPL_VENDOR_INVOICE_HEAD.Saving,0)=1"
                Else
                    Qry += " and ISNULL(TSPL_VENDOR_INVOICE_HEAD.Saving,0)=0"
                End If
                Qry += "  union all " & Environment.NewLine &
                " select  0 as ISProcurementDeduction,'Receipt Note' as [DocType], Payment_No  as Document_No, Payment_No as [PurchaseInvoice],convert(varchar ,Payment_Date ,103) as [DocumentDate] ,convert(varchar,Payment_Date ,103) as [DocDate] ,'' as VendorInvoiceNo,Payment_Amount  as [OriginalAmt],TDS_Amount as TDSAmt ,Payment_Amount - TDS_Amount as NetAmount," & Environment.NewLine &
                " (Payment_Amount - TDS_Amount -isnull((select sum(isnull(Applied_Amount,0)) from TSPL_PAYMENT_DETAIL where TSPL_PAYMENT_DETAIL.Document_No=TSPL_PAYMENT_HEADER.Payment_No  and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No )),0) ) as [PendingAmt] " & Environment.NewLine &
                " , ConvRate ,1 as ConvRateRevaluation from TSPL_PAYMENT_HEADER WHERE Payment_Type  ='RC' and IsChkReverse ='N' AND Is_Security=0 AND ISNULL(TSPL_PAYMENT_HEADER.Applied_Payment  ,'')='' and " & Environment.NewLine &
                " TSPL_PAYMENT_HEADER.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,TSPL_PAYMENT_HEADER.Payment_Date ,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' "
                If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                    Qry += "  and TSPL_PAYMENT_HEADER.Location_GL_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
                End If
                Qry += ") FINALQRY WHERE FINALQRY.PendingAmt>0  and Document_No <>'" & txtDocumentNo.Value & "' and DocType='Debit Note' and ISProcurementDeduction=1 ORDER BY Convert(Date,FINALQRY.DocDate,103)  "
                dt = clsDBFuncationality.GetDataTable(Qry)
                IsInsideLoadData = True
                For Each dr As DataRow In dt.Rows
                    gvDetails.Rows.AddNew()
                    gvDetails.CurrentRow.Cells(colApply).Value = "No"
                    gvDetails.CurrentRow.Cells(colDocType).Value = clsCommon.myCstr(dr("DocType"))
                    gvDetails.CurrentRow.Cells(colPINo).Value = clsCommon.myCstr(dr("PurchaseInvoice"))
                    gvDetails.CurrentRow.Cells(colDocNo).Value = clsCommon.myCstr(dr("Document_No"))
                    If clsCommon.myCdbl(dr("ConvRateRevaluation")) > 0 Then
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRateRevaluation"))
                    Else
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRate"))
                    End If

                    gvDetails.CurrentRow.Cells(colAPTapalNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull (TapalNo,'') as TapalNo  from TSPL_VENDOR_INVOICE_HEAD where Document_No = '" + clsCommon.myCstr(dr("Document_No")) + "'"))
                    gvDetails.CurrentRow.Cells(colDocDate).Value = clsCommon.myCstr(dr("DocDate"))
                    gvDetails.CurrentRow.Cells(colDocumentDate).Value = clsCommon.myCstr(dr("DocumentDate"))
                    gvDetails.CurrentRow.Cells(colVendorInvNo).Value = clsCommon.myCstr(dr("VendorInvoiceNo"))
                    gvDetails.CurrentRow.Cells(colOriginalAmt).Value = clsCommon.myCdbl(dr("OriginalAmt"))
                    gvDetails.CurrentRow.Cells(colTDSAmt).Value = clsCommon.myCstr(dr("TDSAmt"))
                    gvDetails.CurrentRow.Cells(colNetAmt).Value = clsCommon.myCdbl(dr("NetAmount"))
                    gvDetails.CurrentRow.Cells(colPendingAmt).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colTemp).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = True
                    '' Anubhooti 14-Nov-2014 BM00000004636
                    gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = clsCommon.myCdbl(dr("NetAmount")) - clsCommon.myCdbl(dr("PendingAmt"))
                Next
                If dt.Rows.Count > 0 Then
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
                        ddlPaymentType.Enabled = False
                        chkSaving.Enabled = False
                    ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                        ddlPaymentType.Enabled = False
                        chkSaving.Enabled = False
                    End If
                Else
                    txtVendorCode.Enabled = True
                    txtBankCode.Enabled = True
                    ddlPaymentType.Enabled = True
                    chkSaving.Enabled = True
                End If
                IsInsideLoadData = False
            End If

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Private Sub LoadVendorInvoicesOLD(ByVal strVendorCode As String, ByVal InvoiceDate As Date, ByVal isForUpdate As Boolean)
        Try
            If Not isForUpdate Then
                LoadBlankGrid(ddlPaymentType.SelectedValue)
            End If

            Dim strQryForRejectedAmt As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
           "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No ) and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "

            '' this code is wrritten for that debit note which is created auto through PI and Pr is not created against that PI
            Dim strQryForRejectedAmtforNonPR As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
          "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "


            Dim strTaxRecovarableQuery As String = " - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then " &
            " ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'  " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end) "

            If clsCommon.myLen(strVendorCode) > 0 Then
                Qry = "Select * from ( select Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType],  " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_No, Case When ISNULL(Against_POInvoice_No,'')<>'' Then Against_POInvoice_No When ISNULL(Against_PurchaseReturn_No,'')<> '' Then Against_PurchaseReturn_No Else TSPL_VENDOR_INVOICE_HEAD.Document_No End as PurchaseInvoice," &
           " Case When ISNULL(Against_POInvoice_No,'')<>'' Then (Select convert(varchar,PI_Date,103)  from TSPL_PI_HEAD where PI_No =Against_POInvoice_No)  When ISNULL(Against_PurchaseReturn_No,'')<> '' Then (Select convert(varchar,PR_Date,103) from TSPL_PR_HEAD where PR_No  =Against_PurchaseReturn_No ) Else TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date End as DocumentDate, " &
           " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date as [DocDate] ,  " &
           " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No as [VendorInvoiceNo], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " as [OriginalAmt]  ," &
           " case when isnull(TSPL_REMITTANCE.Is_TDS_Provision,'N')='N' then TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount else 0 end as [TDSAmt], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total-isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) as [NetAmount], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total-isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) as [PendingAmt],TSPL_VENDOR_INVOICE_HEAD.ConvRate " &
           " ,isnull(( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and TSPL_REVALUATION_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation " &
           " from TSPL_VENDOR_INVOICE_HEAD " &
           " Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  
left outer join TSPL_REMITTANCE on TSPL_REMITTANCE.Document_No=TSPL_VENDOR_INVOICE_HEAD.Document_No
            WHERE TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,Invoice_Entry_Date,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' And ((ISNULL(RefDocNo,'')= '' or ISNULL((select VI.Document_Type from TSPL_VENDOR_INVOICE_HEAD VI where VI.Document_No =TSPL_VENDOR_INVOICE_HEAD .RefDocNo ),'') in ('I','')) AND ISNULL(Against_PurchaseReturn_No,'')= '')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D') " &
          " and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_Salary_Generation_Code,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='C') "
                If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                    Qry += "  and TSPL_VENDOR_INVOICE_HEAD.Loc_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
                End If

                If isForUpdate Then
                    Dim arrDocNo As New List(Of String)
                    For ii As Integer = 0 To gvDetails.Rows.Count - 1
                        arrDocNo.Add(clsCommon.myCstr(gvDetails.Rows(ii).Cells(colDocNo).Value))
                    Next
                    If arrDocNo.Count > 0 Then
                        Qry += " and TSPL_VENDOR_INVOICE_HEAD.Document_No not in (" + clsCommon.GetMulcallString(arrDocNo) + ")"
                    End If
                End If
                Qry += " and TSPL_VENDOR_INVOICE_HEAD.RefDocType<>'REVALUATION ENTRY'"

                ''richa 13/10/2017
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "TD") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='TD' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "T") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='T' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "S") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='S' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "I") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='I' "
                    Else
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='' "
                    End If
                End If
                ''--------
                Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.posting_date,'')<>''" & Environment.NewLine
                If chkSaving.Checked Then
                    Qry += " and ISNULL(TSPL_VENDOR_INVOICE_HEAD.Saving,0)=1"
                Else
                    Qry += " and ISNULL(TSPL_VENDOR_INVOICE_HEAD.Saving,0)=0"
                End If
                Qry += "  union all " & Environment.NewLine &
                " select  'Receipt Note' as [DocType], Payment_No  as Document_No, Payment_No as [PurchaseInvoice],convert(varchar ,Payment_Date ,103) as [DocumentDate] ,convert(varchar,Payment_Date ,103) as [DocDate] ,'' as VendorInvoiceNo,Payment_Amount  as [OriginalAmt],TDS_Amount as TDSAmt ,Payment_Amount - TDS_Amount as NetAmount," & Environment.NewLine &
                " (Payment_Amount - TDS_Amount -isnull((select sum(isnull(Applied_Amount,0)) from TSPL_PAYMENT_DETAIL where TSPL_PAYMENT_DETAIL.Document_No=TSPL_PAYMENT_HEADER.Payment_No  and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No )),0) ) as [PendingAmt] " & Environment.NewLine &
                " , ConvRate ,1 as ConvRateRevaluation from TSPL_PAYMENT_HEADER WHERE Payment_Type  ='RC' and IsChkReverse ='N' AND Is_Security=0 AND ISNULL(TSPL_PAYMENT_HEADER.Applied_Payment  ,'')='' and " & Environment.NewLine &
                " TSPL_PAYMENT_HEADER.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,TSPL_PAYMENT_HEADER.Payment_Date ,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' "
                If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                    Qry += "  and TSPL_PAYMENT_HEADER.Location_GL_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
                End If
                Qry += ") FINALQRY WHERE FINALQRY.PendingAmt>0  and Document_No <>'" & txtDocumentNo.Value & "' ORDER BY Convert(Date,FINALQRY.DocDate,103)  "
                dt = clsDBFuncationality.GetDataTable(Qry)
                IsInsideLoadData = True
                For Each dr As DataRow In dt.Rows
                    gvDetails.Rows.AddNew()
                    gvDetails.CurrentRow.Cells(colApply).Value = "No"
                    gvDetails.CurrentRow.Cells(colDocType).Value = clsCommon.myCstr(dr("DocType"))
                    gvDetails.CurrentRow.Cells(colPINo).Value = clsCommon.myCstr(dr("PurchaseInvoice"))
                    gvDetails.CurrentRow.Cells(colDocNo).Value = clsCommon.myCstr(dr("Document_No"))
                    If clsCommon.myCdbl(dr("ConvRateRevaluation")) > 0 Then
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRateRevaluation"))
                    Else
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRate"))
                    End If

                    gvDetails.CurrentRow.Cells(colAPTapalNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull (TapalNo,'') as TapalNo  from TSPL_VENDOR_INVOICE_HEAD where Document_No = '" + clsCommon.myCstr(dr("Document_No")) + "'"))
                    gvDetails.CurrentRow.Cells(colDocDate).Value = clsCommon.myCstr(dr("DocDate"))
                    gvDetails.CurrentRow.Cells(colDocumentDate).Value = clsCommon.myCstr(dr("DocumentDate"))
                    gvDetails.CurrentRow.Cells(colVendorInvNo).Value = clsCommon.myCstr(dr("VendorInvoiceNo"))
                    gvDetails.CurrentRow.Cells(colOriginalAmt).Value = clsCommon.myCdbl(dr("OriginalAmt"))
                    gvDetails.CurrentRow.Cells(colTDSAmt).Value = clsCommon.myCstr(dr("TDSAmt"))
                    gvDetails.CurrentRow.Cells(colNetAmt).Value = clsCommon.myCdbl(dr("NetAmount"))
                    gvDetails.CurrentRow.Cells(colPendingAmt).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colTemp).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = True
                    '' Anubhooti 14-Nov-2014 BM00000004636
                    gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = clsCommon.myCdbl(dr("NetAmount")) - clsCommon.myCdbl(dr("PendingAmt"))
                Next
                If dt.Rows.Count > 0 Then
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
                        ddlPaymentType.Enabled = False
                        chkSaving.Enabled = False
                    ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                        ddlPaymentType.Enabled = False
                        chkSaving.Enabled = False
                    End If
                Else
                    txtVendorCode.Enabled = True
                    txtBankCode.Enabled = True
                    ddlPaymentType.Enabled = True
                    chkSaving.Enabled = True
                End If
                IsInsideLoadData = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub LoadPaymentEntries(ByVal strVendorCode As String, ByVal InvoiceDate As Date, ByVal isForUpdate As Boolean)
        Try
            If Not isForUpdate Then
                LoadBlankGrid(ddlPaymentType.SelectedValue)
            End If
            If clsCommon.myLen(strVendorCode) > 0 Then
                Qry = "select Case When TSPL_PAYMENT_HEADER.Payment_type='RC' Then 'Receipt' End As [DocType],   TSPL_PAYMENT_HEADER.Payment_No Document_No," _
                & " TSPL_PAYMENT_HEADER.Payment_No as PurchaseInvoice, TSPL_PAYMENT_HEADER.Payment_date as [DocDate] ,   TSPL_PAYMENT_HEADER.Payment_No as " _
                & " [VendorInvoiceNo],  TSPL_PAYMENT_HEADER.Payment_amount as [OriginalAmt] , TSPL_PAYMENT_HEADER.TDS_Amount as [TDSAmt], " _
                & " (TSPL_PAYMENT_HEADER.Payment_amount-TSPL_PAYMENT_HEADER.TDS_Amount) as [NetAmount],  TSPL_PAYMENT_HEADER.Payment_amount - ISNULL((Select" _
                & " SUM(Applied_Amount) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No = TSPL_PAYMENT_HEADER.payment_No " _
                & " ),0) as [PendingAmt],TSPL_PAYMENT_HEADER.ConvRate  from TSPL_PAYMENT_HEADER "
                Qry += " WHERE Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,TSPL_PAYMENT_HEADER.Payment_date,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' and balance_amt <> '0.00' AND Payment_type='RC' "

                If isForUpdate Then
                    Dim arrDocNo As New List(Of String)
                    For ii As Integer = 0 To gvDetails.Rows.Count - 1
                        arrDocNo.Add(clsCommon.myCstr(gvDetails.Rows(ii).Cells(colDocNo).Value))
                    Next
                    If arrDocNo.Count > 0 Then
                        Qry += " and TSPL_PAYMENT_HEADER.Payment_No not in (" + clsCommon.GetMulcallString(arrDocNo) + ")"
                    End If
                End If

                Qry += " ORDER BY Convert(Date,TSPL_PAYMENT_HEADER.Payment_date,103)  DESC "
                dt = clsDBFuncationality.GetDataTable(Qry)
                IsInsideLoadData = True
                For Each dr As DataRow In dt.Rows
                    gvDetails.Rows.AddNew()
                    gvDetails.CurrentRow.Cells(colApply).Value = "No"
                    gvDetails.CurrentRow.Cells(colDocType).Value = clsCommon.myCstr(dr("DocType"))
                    gvDetails.CurrentRow.Cells(colPINo).Value = clsCommon.myCstr(dr("PurchaseInvoice"))
                    gvDetails.CurrentRow.Cells(colDocNo).Value = clsCommon.myCstr(dr("Document_No"))
                    gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRate"))
                    gvDetails.CurrentRow.Cells(colDocDate).Value = clsCommon.myCstr(dr("DocDate"))
                    gvDetails.CurrentRow.Cells(colVendorInvNo).Value = clsCommon.myCstr(dr("VendorInvoiceNo"))
                    gvDetails.CurrentRow.Cells(colOriginalAmt).Value = clsCommon.myCdbl(dr("OriginalAmt"))
                    gvDetails.CurrentRow.Cells(colTDSAmt).Value = clsCommon.myCstr(dr("TDSAmt"))
                    gvDetails.CurrentRow.Cells(colNetAmt).Value = clsCommon.myCdbl(dr("NetAmount"))
                    gvDetails.CurrentRow.Cells(colPendingAmt).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colTemp).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = True
                    '' Anubhooti 14-Nov-2014 BM00000004636
                    gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = clsCommon.myCdbl(dr("NetAmount")) - clsCommon.myCdbl(dr("PendingAmt"))
                Next
                If dt.Rows.Count > 0 Then
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
                        'txtVendorCode.Enabled = False
                        ddlPaymentType.Enabled = False
                        chkSaving.Enabled = False
                    ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                        'txtBankCode.Enabled = False
                        ddlPaymentType.Enabled = False
                        chkSaving.Enabled = False
                    End If
                Else
                    txtVendorCode.Enabled = True
                    txtBankCode.Enabled = True
                    ddlPaymentType.Enabled = True
                    chkSaving.Enabled = True
                End If
                IsInsideLoadData = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub gvDetails_CellDoubleClick(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gvDetails.CellDoubleClick
        If autoadjustdebitcreditnote = False Then
            If (e.Column Is gvDetails.Columns(colApply)) Then
                If gvDetails.CurrentRow.Cells(colApply).Value = "No" Then
                    gvDetails.CurrentRow.Cells(colApply).Value = "Yes"
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = False
                Else
                    gvDetails.CurrentRow.Cells(colApply).Value = "No"
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = True
                End If
            End If
            If gvDetails.CurrentColumn Is gvDetails.Columns(colDocNo) AndAlso clsCommon.myLen(gvDetails.CurrentRow.Cells(colDocNo).Value) > 0 AndAlso ddlPaymentType.SelectedValue = "PY" Then

                clsOpenTransactionForm.OpenTransacionForm(clsUserMgtCode.mbtnAPInvoiceEntryTDS, gvDetails.CurrentRow.Cells(colDocNo).Value)
            End If
        End If
    End Sub

    Dim IsCellValueChanged As Boolean = True
    Private Sub gvDetails_CellValueChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gvDetails.CellValueChanged
        Try
            If Not IsInsideLoadData Then
                If IsCellValueChanged Then
                    IsCellValueChanged = True
                    If e.Column.FieldName.StartsWith("_CFLD_") Then
                        clsCustomFieldGrid.getFinderForCustomFieldGrid(gvDetails, e.Column.Name.ToString, MyBase.Form_ID)
                    End If
                    If (clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "SR") = CompairStringResult.Equal) Then
                        If (e.Column Is gvDetails.Columns(colApply)) Then
                            If gvDetails.CurrentRow.Cells(colApply).Value = "Yes" Then
                                gvDetails.CurrentRow.Cells(colAppliedAmt).Value = gvDetails.CurrentRow.Cells(colPendingAmt).Value
                            ElseIf e.Row.Cells(colApply).Value = "No" Then
                                gvDetails.CurrentRow.Cells(colAppliedAmt).Value = 0
                            End If
                        End If

                        If (e.Column Is gvDetails.Columns(colAppliedAmt)) Or (e.Column Is gvDetails.Columns(colSecurityAmt)) Then
                            If gvDetails.CurrentRow.Cells(colAppliedAmt).Value > gvDetails.CurrentRow.Cells(colTemp).Value Then
                                common.clsCommon.MyMessageBoxShow(Me, "Applied Amount Can Not Be More Than Pending Amount", Me.Text)
                                gvDetails.CurrentRow.Cells(colAppliedAmt).Value = gvDetails.CurrentRow.Cells(colTemp).Value
                            End If
                            Dim PaymentAmount As Double = 0
                            For Each grow As GridViewRowInfo In gvDetails.Rows
                                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Debit Note") = CompairStringResult.Equal Then
                                    PaymentAmount -= clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                                Else
                                    PaymentAmount += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value) + clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
                                End If
                            Next
                            txtPaymentAmt.Text = PaymentAmount
                            gvDetails.CurrentRow.Cells(colPendingAmt).Value = gvDetails.CurrentRow.Cells(colTemp).Value - (gvDetails.CurrentRow.Cells(colAppliedAmt).Value + gvDetails.CurrentRow.Cells(colSecurityAmt).Value)
                        End If
                        '============Added by Rohit to Save Security Amount against Payment===========
                        If (e.Column Is gvDetails.Columns(colSecurityAmt)) Then
                            If (clsCommon.myCdbl(gvDetails.CurrentRow.Cells(colAppliedAmt).Value) + clsCommon.myCdbl(gvDetails.CurrentRow.Cells(colSecurityAmt).Value)) > gvDetails.CurrentRow.Cells(colTemp).Value Then
                                common.clsCommon.MyMessageBoxShow(Me, "Applied Amount + Security Amount Can Not Be More Than Pending Amount", Me.Text)
                                gvDetails.CurrentRow.Cells(colAppliedAmt).Value = gvDetails.CurrentRow.Cells(colTemp).Value
                                gvDetails.CurrentRow.Cells(colSecurityAmt).Value = 0
                            End If
                        End If
                        '============================================

                        IsCellValueChanged = True
                    ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
                        If clsCommon.myCstr(txtBankCode.Value) = "" Then
                            clsCommon.MyMessageBoxShow("Please Select Bank Code")
                            txtBankCode.Focus()
                            Exit Sub
                        End If
                        If gvDetails.CurrentColumn Is gvDetails.Columns(colGLAccount) Then
                            If clsCommon.myLen(txtMPAdv.Value) <= 0 Then
                                OpenGLAccount(False)
                                gvDetails.CurrentRow.Cells(colAccDesc).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_GL_ACCOUNTS where Account_Code='" + clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLAccount).Value) + "'"))
                                Dim grouptype As String = ""
                                grouptype = clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLType).Value)
                                If clsCommon.CompairString(grouptype, "Balance Sheet") = CompairStringResult.Equal Then
                                    gvDetails.CurrentRow.Cells(colHirerachyCenter).ReadOnly = True
                                    gvDetails.CurrentRow.Cells(colCostCenter).ReadOnly = True
                                Else
                                    gvDetails.CurrentRow.Cells(colHirerachyCenter).ReadOnly = False
                                    gvDetails.CurrentRow.Cells(colCostCenter).ReadOnly = False
                                End If
                            End If
                        ElseIf gvDetails.CurrentColumn Is gvDetails.Columns(colHirerachyCenter) Then
                            Dim grouptype As String = ""
                            grouptype = clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLType).Value)
                            If clsCommon.CompairString(grouptype, "Balance Sheet") <> CompairStringResult.Equal Then
                                OpenHierarchyCode(False)
                            End If

                        ElseIf gvDetails.CurrentColumn Is gvDetails.Columns(colCostCenter) Then
                            Dim grouptype As String = ""
                            grouptype = clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLType).Value)
                            If clsCommon.CompairString(grouptype, "Balance Sheet") <> CompairStringResult.Equal Then
                                OpenCostCenterCode(False)
                            End If

                        End If
                        If gvDetails.CurrentColumn Is gvDetails.Columns(colAmount) Then
                            CalculateTotalAppliedAmt()
                        End If
                        CalculateTotalAppliedAmt()
                    End If
                End If
                IsCellValueChanged = True
            End If
        Catch ex As Exception
            IsCellValueChanged = True
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub OpenHierarchyCode(ByVal isButtonClick As Boolean)
        Dim qry As String = " select TSPL_HIRERACHY_LEVEL_MASTER.HIRERACHY_CODE as [Code] ,TSPL_HIRERACHY_LEVEL_MASTER.Description as [Description],TSPL_HIRERACHY_LEVEL_MASTER.Level as [Level] ,TSPL_HIRERACHY_LEVEL_MASTER.Created_By as [Created By] ,CONVERT(VARCHAR,TSPL_HIRERACHY_LEVEL_MASTER.Created_Date,103) as [Created Date] ,TSPL_HIRERACHY_LEVEL_MASTER.Modified_By as [Modified By] ,CONVERT(VARCHAR,TSPL_HIRERACHY_LEVEL_MASTER.Modified_Date,103) as [Modified Date]  From TSPL_HIRERACHY_LEVEL_MASTER  "
        gvDetails.CurrentRow.Cells(colHirerachyCenter).Value = clsCommon.ShowSelectForm("HierarchyPN", qry, "Code", "", clsCommon.myCstr(gvDetails.CurrentRow.Cells(colHirerachyCenter).Value), "Code", isButtonClick)
        gvDetails.CurrentRow.Cells(colHirerachyName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Description from TSPL_HIRERACHY_LEVEL_MASTER where HIRERACHY_CODE='" + clsCommon.myCstr(gvDetails.CurrentRow.Cells(colHirerachyCenter).Value) + "'"))
    End Sub
    Private Sub OpenCostCenterCode(ByVal isButtonClick As Boolean)
        If clsCommon.myLen(clsCommon.myCstr(gvDetails.CurrentRow.Cells(colHirerachyCenter).Value)) > 0 Then
            Dim DBLevel As String = String.Empty
            DBLevel = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select ISNULL(Level,'') AS Level from TSPL_HIRERACHY_LEVEL_MASTER Where HIRERACHY_CODE='" + clsCommon.myCstr(gvDetails.CurrentRow.Cells(colHirerachyCenter).Value) + "' "))
            Dim qry As String = "select TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Code as [Code] ,TSPL_COST_CENTRE_FINANCIAL.Cost_Center_Fin_Name,ISNULL(TSPL_COST_CENTRE_FINANCIAL.Hirerachy_Level_Code,'') AS [Hirerachy Level Code],ISNULL(TSPL_COST_CENTRE_FINANCIAL.Cost_Centre_Fin_Level_Code,'') AS [Cost Centre Fin Level Code],ISNULL(TSPL_COST_CENTRE_FINANCIAL.Hirerachy_Level,'') AS [Hirerachy Level] ,TSPL_COST_CENTRE_FINANCIAL.Created_By as [Created By] ,Convert(varchar,TSPL_COST_CENTRE_FINANCIAL.Created_Date,103) as [Created Date] ,TSPL_COST_CENTRE_FINANCIAL.Modified_By as [Modified By] ,Convert(varchar,TSPL_COST_CENTRE_FINANCIAL.Modified_Date,103) as [Modified Date]  From TSPL_COST_CENTRE_FINANCIAL "
            gvDetails.CurrentRow.Cells(colCostCenter).Value = clsCommon.ShowSelectForm("HierarchyPNCc", qry, "Code", " Hirerachy_Level = '" + DBLevel + "'", clsCommon.myCstr(gvDetails.CurrentRow.Cells(colCostCenter).Value), "Code", isButtonClick)
            gvDetails.CurrentRow.Cells(colCostCenterName).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Cost_Center_Fin_Name from TSPL_COST_CENTRE_FINANCIAL where Cost_Center_Fin_Code='" + clsCommon.myCstr(gvDetails.CurrentRow.Cells(colCostCenter).Value) + "'"))
        Else
            clsCommon.MyMessageBoxShow(Me, "Please select hirerachy level first.", Me.Text)
        End If
    End Sub

    Private Sub gvDetails_CurrentColumnChanged(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.CurrentColumnChangedEventArgs) Handles gvDetails.CurrentColumnChanged
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
            If gvDetails.RowCount > 0 Then
                Dim intCurrRow As Integer = gvDetails.CurrentRow.Index
                gvDetails.CurrentRow.Cells(colLineNo).Value = clsCommon.myCdbl(intCurrRow + 1)
                If intCurrRow = gvDetails.Rows.Count - 1 Then
                    gvDetails.Rows.AddNew()
                    gvDetails.CurrentRow = gvDetails.Rows(intCurrRow)
                End If
            End If
        End If
    End Sub

    Private Sub gvDetails_CellEditorInitialized(ByVal sender As Object, ByVal e As Telerik.WinControls.UI.GridViewCellEventArgs) Handles gvDetails.CellEditorInitialized
        If TypeOf Me.gvDetails.CurrentColumn Is GridViewMultiComboBoxColumn Then
            Dim editor As RadMultiColumnComboBoxElement = DirectCast(Me.gvDetails.ActiveEditor, RadMultiColumnComboBoxElement)
            editor.AutoSizeDropDownToBestFit = True
            editor.EditorControl.MasterTemplate.BestFitColumns()
            editor.DropDownStyle = RadDropDownStyle.DropDown
            editor.AutoFilter = True
            If editor.EditorControl.MasterTemplate.FilterDescriptors.Count = 0 Then
                Dim autoFilter As FilterDescriptor = New FilterDescriptor("Description", FilterOperator.StartsWith, "")
                autoFilter.IsFilterEditor = True
                editor.EditorControl.FilterDescriptors.Add(autoFilter)
            End If
        End If
    End Sub

    Private Sub OpenGLAccount(ByVal isButtonClick As Boolean)
        Dim bankseg As String = " select right(BANKACC,3) as segment,BANKACC,* from TSPL_BANK_MASTER where BANK_CODE='" + txtBankCode.Value + "'"
        Dim val As String = clsDBFuncationality.getSingleValue(bankseg)
        Dim qry As String
        Dim whrcls As String
        Dim arr As New ArrayList()
        arr = clsERPFuncationality.glaccountquery(objCommonVar.CurrentUserCode)
        qry = arr.Item(0) + " inner join TSPL_GL_STRUCTURE on TSPL_GL_ACCOUNTS .Str_Code=TSPL_GL_STRUCTURE.Str_Code "
        whrcls = arr.Item(1)

        If whrcls = "" Then
        Else
            whrcls = "(" + whrcls + ")"
        End If

        If whrcls Is Nothing OrElse whrcls = "" Then
            whrcls = " 1<>(isnull(Seg_No1,0) +isnull(Seg_No2,0) +isnull(Seg_No3,0) +isnull(Seg_No4,0) +isnull(Seg_No5,0) +isnull(Seg_No6,0) +isnull(Seg_No7,0) +isnull(Seg_No8,0) +isnull(Seg_No9,0) +isnull(Seg_No10,0) )"
        Else
            whrcls = whrcls + " and 1<>(isnull(Seg_No1,0) +isnull(Seg_No2,0) +isnull(Seg_No3,0) +isnull(Seg_No4,0) +isnull(Seg_No5,0) +isnull(Seg_No6,0) +isnull(Seg_No7,0) +isnull(Seg_No8,0) +isnull(Seg_No9,0) +isnull(Seg_No10,0) )"
        End If
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, Nothing)) = 0 Then
            whrcls += "   and TSPL_GL_ACCOUNTS.Account_Seg_Code7='" + val + "'"
        End If
        whrcls += "  and TSPL_GL_ACCOUNTS.ControlAccount='N' "

        If clsCommon.GetDateWithStartTime(dtpPayment.Value) < clsCommon.GetDateWithStartTime(ERPStartDate) Then
            whrcls = clsCommon.ReplaceString(whrcls, "TSPL_GL_ACCOUNTS.ControlAccount='N'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "ControlAccount ='N'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "( ControlAccount ='N')", " ")

            whrcls = clsCommon.ReplaceString(whrcls, "TSPL_GL_ACCOUNTS.ControlAccount<>'Y'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "ControlAccount<>'Y'", " 2=2 ")
            whrcls = clsCommon.ReplaceString(whrcls, "( ControlAccount<>'Y')", " ")
        End If

        ''richa 05 feb,2019  TEC/05/02/19-000412 check for opening in case of miscelleonous
        Dim JEWithOPening As Boolean = False
        If clsCommon.myLen(clsCommon.GetDateWithStartTime(ERPStartDate)) > 0 Then
            If clsCommon.GetDateWithStartTime(dtpPayment.Value) < clsCommon.GetDateWithStartTime(ERPStartDate) Then
                JEWithOPening = True
            End If
        End If
        Dim strCustomerOpeningAccount As String = String.Empty
        If clsCommon.CompairString(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.CreateOpeningEntryAutomatically, clsFixedParameterCode.CreateOpeningEntryAutomatically, Nothing)), "1") = CompairStringResult.Equal AndAlso (clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "MI") = CompairStringResult.Equal) And JEWithOPening = True Then
            If clsCommon.myLen(txtBankCode.Value) <= 0 Then
                common.clsCommon.MyMessageBoxShow(Me, "Please first select Bank", Me.Text)
                txtBankCode.Focus()
                Return
            End If
            strCustomerOpeningAccount = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select ISNULL(Bank_Opening_Clearing_Account ,'') from tspl_bank_master where BANK_CODE ='" & txtBankCode.Value & "'"))
            If clsCommon.myLen(strCustomerOpeningAccount) = 0 Then
                common.clsCommon.MyMessageBoxShow("Please set Bank Opening Clearing Account for Bank - " + txtBankCode.Value)
                Return
            End If
            whrcls += " and Account_Code='" & strCustomerOpeningAccount & "'"
        End If

        'richa 17 SEp,2019 TEC/03/07/19-000927
        Dim strqry As String = " Select Account_Code,Description from (" & qry & " where " & whrcls & Environment.NewLine &
            " UNION All " & Environment.NewLine &
            " select Account_Code , Description  from TSPL_GL_ACCOUNTS " & Environment.NewLine &
" left outer join (select TSPL_GL_SEGMENT_CODE.Account_Code as AccCode from TSPL_GL_SEGMENT_CODE where TSPL_GL_SEGMENT_CODE.Seg_No='7' " & Environment.NewLine &
" and len(isnull(TSPL_GL_SEGMENT_CODE.Account_Code,''))>0 ) as segTable  on segTable.AccCode=TSPL_GL_ACCOUNTS.Account_Code " & Environment.NewLine &
    " inner join TSPL_GL_STRUCTURE on TSPL_GL_ACCOUNTS .Str_Code=TSPL_GL_STRUCTURE.Str_Code where ( 2=2  and TSPL_GL_ACCOUNTS.Status='Y' and ( segTable.AccCode is null  ))" & Environment.NewLine &
    " and 1<>(isnull(Seg_No1,0) +isnull(Seg_No2,0) +isnull(Seg_No3,0) +isnull(Seg_No4,0) +isnull(Seg_No5,0) +isnull(Seg_No6,0) +isnull(Seg_No7,0) +isnull(Seg_No8,0) +isnull(Seg_No9,0) +isnull(Seg_No10,0) ) " & Environment.NewLine &
    " and TSPL_GL_ACCOUNTS.Account_Code in (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForPayment =1) "
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, Nothing)) = 0 Then
            strqry += "   and TSPL_GL_ACCOUNTS.Account_Seg_Code7='" + val + "'"
        End If
        If clsCommon.myLen(strCustomerOpeningAccount) > 0 Then
            strqry += " and Account_Code='" & strCustomerOpeningAccount & "'"
        End If
        strqry += " ) Final "
        gvDetails.CurrentRow.Cells(colGLAccount).Value = clsCommon.ShowSelectForm("PaymentGLACFND1", strqry, "Account_Code", "", clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLAccount).Value), "Account_Code", isButtonClick)
        'gvDetails.CurrentRow.Cells(colGLAccount).Value = clsCommon.ShowSelectForm("PaymentGLACFND1", qry, "Account_Code", whrcls, clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLAccount).Value), "Account_Code", isButtonClick)
        If clsCommon.myLen(gvDetails.CurrentRow.Cells(colGLAccount).Value) > 0 Then
            gvDetails.CurrentRow.Cells(colGLType).Value = clsPaymentHeader.CheckGLAccountType(clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLAccount).Value), Nothing)
        Else
            gvDetails.CurrentRow.Cells(colGLType).Value = ""
        End If
    End Sub

    Private Sub gvDetails_UserDeletingRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowCancelEventArgs) Handles gvDetails.UserDeletingRow
        If gvDetails.Rows.Count > 1 Then
            If common.clsCommon.MyMessageBoxShow("Delete The Current Row." + Environment.NewLine + "Are you sure?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                e.Cancel = True
            End If
        Else
            clsCommon.MyMessageBoxShow(Me, "Sorry! Yoy can not delete last row", Me.Text)
            e.Cancel = True
        End If
    End Sub

    Private Sub gvDetails_UserDeletedRow(ByVal sender As System.Object, ByVal e As Telerik.WinControls.UI.GridViewRowEventArgs) Handles gvDetails.UserDeletedRow
        CalculateTotalAppliedAmt()
    End Sub

    Private Sub CalculateTotalAppliedAmt()
        Dim TotalAppliedAmt As Double = 0.0
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
            For Each Grow As GridViewRowInfo In gvDetails.Rows
                TotalAppliedAmt = TotalAppliedAmt + clsCommon.myCdbl(Grow.Cells(colAppliedAmt).Value)
            Next
            txtPaymentAmt.Text = clsCommon.myCstr(TotalAppliedAmt)
        ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
            For Each Grow As GridViewRowInfo In gvDetails.Rows
                TotalAppliedAmt = TotalAppliedAmt + clsCommon.myCdbl(Grow.Cells(colAmount).Value)
            Next
            txtTotalAppliedAmt.Text = clsCommon.myCstr(TotalAppliedAmt)
        End If
    End Sub

#End Region

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub
    Function AllowToSave1() As Boolean
        Try
            Xtra.TransactionValidity(dtpPayment.Value)
            Dim isAllowFutureTransForPDCCheque As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowtoFutureDateTransForPDCCheque, clsFixedParameterCode.AllowtoFutureDateTransForPDCCheque, Nothing)) = 1, True, False)
            If isAllowFutureTransForPDCCheque = True AndAlso chkPDC.Checked = True Then
            Else
                If AllowFutureDateTransaction(dtpPayment.Value, Nothing) = False Then
                    dtpPayment.Focus()
                    Return False
                End If
            End If


            Return True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
    End Function
    Private Sub btnsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsave.Click
        If AllowToSave1() Then
            SaveData1()
        End If
    End Sub

    Private Function SaveData1()
        Try
            deadLockCounter = 1
            If btnsave.Text = "Update" And clsCommon.myLen(txtPaymentNo.Value) > 0 Then
                Dim strchk As String = "select Posted from TSPL_PAYMENT_HEADER where Payment_No='" + txtPaymentNo.Value + "'"
                Dim chkpost As String = clsDBFuncationality.getSingleValue(strchk)
                If clsCommon.CompairString(chkpost, "1") = CompairStringResult.Equal Then
                    clsCommon.MyMessageBoxShow(Me, "Transaction already posted", Me.Text)
                    Return False
                End If
            End If
            SaveData2()
            If deadLockCounter < 15 Then
                clsCommon.MyMessageBoxShow(Me, "Data Saved Successfully", Me.Text)
            End If
            Return True
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
            Return False
        End Try
    End Function

    Private Function SaveData2() As Boolean
        Try
            SaveData()
        Catch ex As Exception
            If ex.Message.Contains("deadlocked") Then
                deadLockCounter += 1
                If deadLockCounter >= 15 Then
                    Me.Close()
                    Exit Function
                End If
                System.Threading.Thread.Sleep(3000)
                SaveData2()
            Else
                Throw New Exception(ex.Message)
            End If
        End Try
        Return True
    End Function

    Private Function SaveData(Optional ByVal isPosted As Boolean = False) As Boolean
        Try

            If (AllowToSave()) Then
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
                    clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtPaymentNo.Value))
                End If

                Dim obj As New clsPaymentHeader()
                obj.Payment_No = clsCommon.myCstr(txtPaymentNo.Value)
                obj.Entry_Desc = clsCommon.myCstr(txtDescription.Text)
                obj.Payment_Date = clsCommon.myCDate(dtpPayment.Value)
                obj.Payment_Post_Date = clsCommon.myCDate(dtpPayment.Value)
                obj.Bank_Code = clsCommon.myCstr(txtBankCode.Value)
                obj.Payment_Type = clsCommon.myCstr(ddlPaymentType.SelectedValue)
                obj.Vendor_Code = clsCommon.myCstr(txtVendorCode.Value)
                obj.Vendor_Name = clsCommon.myCstr(lblVendorName.Text)
                obj.Payment_Code = clsCommon.myCstr(txtPaymentMode.Value)
                obj.TDS_Provision = chkTDSProvision.Checked
                '' Anubhooti 22-July-2014 BM00000003161
                If ChkAccPayee.Checked Then
                    obj.Account_Payee = 1
                Else
                    obj.Account_Payee = 0
                End If
                If chkIsReceipt.Checked Then
                    obj.isReceipt = 1
                Else
                    obj.isReceipt = 0
                End If
                obj.TapalNo = clsCommon.myCstr(txtTapalNo.Text)
                If txtDataAndTimeSelection.Checked Then
                    obj.DateAndTime = txtDataAndTimeSelection.Value
                End If
                obj.WaveOFFBankCharges = clsCommon.myCstr(IIf(chkBankChargesWaveOff.Checked = True, "Y", "N"))
                '' Anubhooti 07-Jan-2014 BM00000005309
                If clsCommon.myLen(txtlocation.Value) > 0 Then
                    obj.Location_GL_Code = txtlocation.Value
                Else
                    obj.Location_GL_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code in (select BANKACC from TSPL_BANK_MASTER where BANK_CODE='" + txtBankCode.Value + "')"))
                End If
                obj.Location_Code_Prefix = txtLocationPrefix.Value
                obj.Employee_Type = clsCommon.myCstr(ddlEmployeeType.SelectedValue)
                obj.Employee_Advance_Type = clsCommon.myCstr(ddlEmployeeAdvanceType.SelectedValue)
                obj.memorndmamt = "0"
                If chkmemorndm.Checked = True Then
                    obj.memorndmamt = clsCommon.myCdbl(txtmemoamt.Text)
                End If

                obj.Vendor_Bank_Code = txtVendor_bankcode.Text
                obj.Vendor_Bank_Name = TxtVendor_BankName.Text
                obj.Vendor_IFSC_Code = TxtVendorBank_IFSCCode.Text
                obj.Vendor_Branch_Name = txtVendorBank_branchname.Text
                obj.Vendor_Bank_ACNo = txtVendor_Bank_ACNo.Text

                If clsCommon.CompairString(obj.Payment_Code, "Cheque") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Code, "DD") = CompairStringResult.Equal Then
                    obj.Cheque_No = clsCommon.myCstr(txtChequeNo.Text)
                    obj.Cheque_Date = clsCommon.myCDate(dtpChequeDate.Value)
                    If chkPDC.Checked Then
                        obj.PDC_Cheque = "Y"
                    End If
                    If chkCheckPrint.Checked Then
                        obj.CHECK_PRINT = 1
                    Else
                        obj.CHECK_PRINT = 0
                    End If
                Else
                    obj.Cheque_No = ""
                    obj.Cheque_Date = Nothing
                End If
                Dim OutstandingAmt As Decimal = 0
                If Not (clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal) Then
                    '' Anubhooti 27-Nov-2014 BM00000004644 (Remove Outstanding popup)
                    'If clsCommon.myCdbl(lblCustomerOutStanding.Text) <> 0 Then
                    '    If common.clsCommon.MyMessageBoxShow("OutStanding of vendor is (" & lblCustomerOutStanding.Text & "). Do You Want To consider it ?", "Payment", MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    '        OutstandingAmt = clsCommon.myCdbl(lblCustomerOutStanding.Text)
                    '    End If
                    'End If

                End If
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "MD") = CompairStringResult.Equal Then
                    obj.Payment_Amount = clsCommon.myCdbl(txtPaymentAmt.Text) - OutstandingAmt
                    obj.Balance_Amt = clsCommon.myCdbl(txtPaymentAmt.Text) - OutstandingAmt
                    ' '' Anubhooti 24-Sep-2014 
                    'If ChkSecurity.Checked Then
                    '    obj.Is_Security = 1
                    'Elserc
                    '    obj.Is_Security = 0
                    'End If
                ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    obj.Total_Prepayment = clsCommon.myCdbl(txtPaymentAmt.Text) - OutstandingAmt
                    obj.Payment_Amount = clsCommon.myCdbl(txtNetPayableAmt.Text) - OutstandingAmt
                    obj.Balance_Amt = clsCommon.myCdbl(txtNetPayableAmt.Text) - OutstandingAmt
                    obj.TDS_Amount = clsCommon.myCdbl(txtTDSAmt.Text)
                ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
                    obj.Payment_Amount = clsCommon.myCdbl(txtPaymentAmt.Text) - OutstandingAmt
                    obj.Total_Applied_Amount = clsCommon.myCdbl(txtPaymentAmt.Text) - OutstandingAmt
                    obj.Remit_To = clsCommon.myCstr(txtRemitTo.Text)
                    obj.Loadout_No = txtLoadOutno.Value
                    obj.MP_Code_For_Advance = txtMPAdv.Value
                End If
                '' Anubhooti 09-Oct-2014 
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    If ChkSecurity.Checked Then
                        obj.Is_Security = 1
                    Else
                        obj.Is_Security = 0
                    End If
                    If ChkRetention.Checked Then
                        obj.Is_Retention = 1
                    Else
                        obj.Is_Retention = 0
                    End If
                    '' Anubhooti 27-Mar-2015 (Advance/OnAccount Against Salary checkbox)
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") Then
                        If ChkAdvSalary.Checked Then
                            obj.Advance_Against_Salary = 1
                        Else
                            obj.Advance_Against_Salary = 0
                        End If
                    End If
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                        obj.Saving = chkSaving.Checked
                    End If
                End If

                obj.isFarmerLoanPayment = IIf(chkFarmerLoanPayment.Checked = True, 1, 0)
                ''
                obj.IsChkReverse = "N"
                obj.Bank_Charges = clsCommon.myCdbl(txtBankCharges.Text)
                obj.objRemittance = objRemittance
                '' Anubhooti 23-Nov-2014 BM00000004668 (Remove C-Form)
                'If chkCForm.Checked = True Then
                '    obj.CFormRecd = "Y"
                'Else
                '    obj.CFormRecd = "N"
                'End If
                '' Anubhooti 08-Dec-2014 (Remove Invoice No.)
                'obj.CForm_InvoiceNo = txtCFormInvNo.Value
                '' Anubhooti 21-Aug-2014
                obj.PurchaseOrder_No = clsCommon.myCstr(txtPONo.Value)
                '' Anubhooti 25-Sep-2014 BM00000004050
                obj.Account_Payee_Name = clsCommon.myCstr(LblAccPayee.Text)

                obj.Applied_Payment = clsCommon.myCstr(txtDocumentNo.Value)


                obj.Interest_Rate = txtInterestRate.Value
                obj.No_Of_EMI = txtNoOfEMI.Value

                obj.ArrTr = New List(Of clsPaymentDetail)
                ''===shivani
                obj.Loan_Code = clsCommon.myCstr(fndloanNo.Value)

                ''For Custom Fields
                obj.Form_ID = MyBase.Form_ID
                obj.arrCustomFields = New List(Of clsCustomFieldValues)
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.GetData(obj.arrCustomFields)
                End If
                ''End of For Custom Fields
                'obj.is_Opening = chkOpening.Checked

                ''richa 17 July,2019 
                'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, Nothing)
                obj.is_Opening = 0
                If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                    Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                    If clsCommon.myCDate(obj.Payment_Date, "dd/MM/yyyy") <= clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy") Then
                        obj.is_Opening = 1
                    End If
                End If

                '============================Detail Section==============================
                If (clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "SR") = CompairStringResult.Equal) Then
                    Dim TotalSecurityAmount As Double = 0
                    For Each grow As GridViewRowInfo In gvDetails.Rows
                        If clsCommon.myCstr(grow.Cells(colApply).Value) = "Yes" Then
                            Dim objTr As New clsPaymentDetail()
                            objTr.Apply = "1"
                            objTr.Payment_Type = clsCommon.myCstr(ddlPaymentType.SelectedValue)
                            objTr.Document_No = clsCommon.myCstr(grow.Cells(colDocNo).Value)
                            If grow.Cells(colDocNo).Tag Is Nothing Then
                                objTr.ConvRateOld = 1
                            Else
                                objTr.ConvRateOld = IIf(clsCommon.myCdbl(grow.Cells(colDocNo).Tag) = 0, 1, clsCommon.myCdbl(grow.Cells(colDocNo).Tag))
                            End If

                            objTr.Vendor_Invoice_No = clsCommon.myCstr(grow.Cells(colVendorInvNo).Value)
                            objTr.Original_Invoice_Amt = clsCommon.myCdbl(grow.Cells(colOriginalAmt).Value)
                            objTr.TDS_Amount = clsCommon.myCdbl(grow.Cells(colTDSAmt).Value)
                            objTr.Applied_Amount = clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                            ''changed by richa agarwal on 07/01/2015
                            'objTr.Pending_Balance = clsCommon.myCdbl(grow.Cells(colTemp).Value)
                            'objTr.Net_Balance = clsCommon.myCdbl(grow.Cells(colPendingAmt).Value)
                            objTr.Pending_Balance = clsCommon.myCdbl(grow.Cells(colPendingAmt).Value)
                            objTr.Net_Balance = clsCommon.myCdbl(grow.Cells(colOriginalAmt).Value)
                            ''----------------------
                            objTr.Security_Amount = clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
                            TotalSecurityAmount += objTr.Security_Amount
                            objTr.Comment = clsCommon.myCstr(grow.Cells(colComment).Value)
                            obj.ArrTr.Add(objTr)
                        End If
                    Next
                    obj.Total_Security_Amount = TotalSecurityAmount
                    If (obj.ArrTr Is Nothing OrElse obj.ArrTr.Count <= 0) Then
                        Throw New Exception("Please Fill at list one Item")
                        Return False
                    End If

                    If MyBase.ArrDetailFields IsNot Nothing AndAlso MyBase.ArrDetailFields.Count > 0 Then
                        clsCustomFieldGrid.GetData(obj.arrCustomFields, gvDetails, MyBase.ArrDetailFields, colDocNo)
                    End If
                ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
                    Dim ESiAmt As Decimal = 0.0
                    Dim MiscAmt As Decimal = 0.0
                    Dim ESI_Percent As Decimal = 0.0
                    For i As Integer = 0 To gvDetails.Rows.Count - 1
                        If Not IsDBNull(gvDetails.Rows(i).Cells(colGLAccount).Value) Or Not IsDBNull(gvDetails.Rows(i).Cells(colAmount).Value) Then
                            If gvDetails.Rows(i).Cells(colGLAccount).Value <> "" Then
                                If clsCommon.myCdbl(gvDetails.Rows(i).Cells(3).Value) > 0 Then
                                    MiscAmt = MiscAmt + clsCommon.myCdbl(gvDetails.Rows(i).Cells(3).Value)
                                    ESiAmt = ESiAmt + (clsCommon.myCdbl(gvDetails.Rows(i).Cells(3).Value) * -1)
                                End If
                            End If
                        End If
                    Next
                    If MiscAmt = 0 Then
                        ESI_Percent = 0
                    Else
                        ESI_Percent = (ESiAmt / MiscAmt) * 100
                    End If
                    Dim iiRowNo As Integer = 1
                    For Each grow As GridViewRowInfo In gvDetails.Rows
                        iiRowNo += 1
                        If grow.Cells(colGLAccount).Value <> "" Then
                            Dim objTr As New clsPaymentDetail()
                            objTr.Payment_Type = clsCommon.myCstr(ddlPaymentType.SelectedValue)
                            objTr.Account_Code = clsCommon.myCstr(grow.Cells(colGLAccount).Value)
                            objTr.Description = clsCommon.myCstr(grow.Cells(colAccDesc).Value)
                            'objTr.Applied_Amount = clsCommon.myCdbl(grow.Cells(colAmount).Value)
                            objTr.Net_Balance = clsCommon.myCdbl(grow.Cells(colAmount).Value)
                            objTr.Remarks = clsCommon.myCstr(grow.Cells(colRemark).Value)
                            objTr.ESI_WCT_Percentage = ESI_Percent

                            objTr.Hirerachy_Level_Code = clsCommon.myCstr(grow.Cells(colHirerachyCenter).Value)
                            objTr.Cost_Center_Fin_Code = clsCommon.myCstr(grow.Cells(colCostCenter).Value)
                            If isApplyCostCenter Then
                                Dim grouptype As String = ""
                                grouptype = clsPaymentHeader.CheckGLAccountType(clsCommon.myCstr(grow.Cells(colGLAccount).Value), Nothing)
                                If clsCommon.CompairString(grouptype, "Balance Sheet") = CompairStringResult.Equal Then
                                Else
                                    If clsCommon.myLen(clsCommon.myCstr(txtMPAdv.Value)) <= 0 Then
                                        If clsCommon.myLen(objTr.Hirerachy_Level_Code) <= 0 Then
                                            Throw New Exception("Please provide the Hierarchy Level " + clsCommon.myCstr(iiRowNo))
                                        ElseIf clsCommon.myLen(objTr.Cost_Center_Fin_Code) <= 0 Then
                                            Throw New Exception("Please provide the Cost Center " + clsCommon.myCstr(iiRowNo))
                                        End If
                                    End If

                                End If

                            End If
                            obj.ArrTr.Add(objTr)
                        End If
                    Next
                    If MyBase.ArrDetailFields IsNot Nothing AndAlso MyBase.ArrDetailFields.Count > 0 Then
                        clsCustomFieldGrid.GetData(obj.arrCustomFields, gvDetails, MyBase.ArrDetailFields, colGLAccount)
                    End If
                End If
                '==================Detail Section Ends Here=======================

                '' CurrencyConversion

                If clsModuleCurrencyMapping.CheckMultiCurrency(Me.Module_Code) = True Then
                    obj.CURRENCY_CODE = Me.txtCurrencyCode.Value
                    obj.ConvRate = clsCommon.myCdbl(Me.txtConversionRate.Text)
                    obj.ConvRateOld = clsCommon.myCdbl(Me.txtConversionRate.Text)
                    If clsCommon.myLen(txtApplicableFrom.Text) > 0 Then
                        obj.ApplicableFrom = clsCommon.myCDate(Me.txtApplicableFrom.Text)
                    Else
                        obj.ApplicableFrom = Nothing
                    End If
                    obj.BASE_CURRENCY_CODE = clsCommon.myCstr(Me.txtBaseCurrency.Value)
                    obj.PAYMENT_AMOUNT_BASE_CURRENCY = clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text)

                    ''richa agarwal 18/05/2015

                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                        If isPosted AndAlso clsCommon.CompairString(obj.CURRENCY_CODE, obj.BASE_CURRENCY_CODE) <> CompairStringResult.Equal Then
                            '' only during posting of transaction
                            '' gather information for exchange gain and loss account
                            Dim dt As DataTable
                            dt = clsPaymentHeader.GetExchangeDetailDt(txtVendorCode.Value)
                            If dt.Rows.Count > 0 Then
                                obj.EXCHANGE_GAIN_ACCOUNT = clsCommon.myCstr(dt.Rows(0).Item("EXCHANGE_GAIN_ACCOUNT"))
                                obj.EXCHANGE_LOSS_ACCOUNT = clsCommon.myCstr(dt.Rows(0).Item("EXCHANGE_LOSS_ACCOUNT"))
                            Else
                                obj.EXCHANGE_GAIN_ACCOUNT = Nothing
                                obj.EXCHANGE_LOSS_ACCOUNT = Nothing
                            End If
                            Dim dtLastRate As DataTable
                            '' gather conv rate and amount of transaction to calculate exchange loss and gain
                            Dim strInvoiceNo As String = String.Empty
                            Dim lossorgainamount As Double = 0
                            Dim Totallossorgainamount As Double = 0
                            Dim InvoiceType As String = ""
                            For i As Integer = 0 To gvDetails.Rows.Count - 1
                                strInvoiceNo = clsCommon.myCstr(gvDetails.Rows(i).Cells(colDocNo).Value)
                                InvoiceType = clsCommon.myCstr(gvDetails.Rows(i).Cells(colDocType).Value)
                                dtLastRate = clsPaymentHeader.GetExchangeRateAmount(strInvoiceNo, dtpPayment.Value)
                                If clsCommon.CompairString(InvoiceType, "Debit Note") = CompairStringResult.Equal Then
                                    lossorgainamount = clsCommon.myCdbl(gvDetails.Rows(i).Cells(colAppliedAmt).Value) * dtLastRate.Rows(0).Item("ConvRate") * -1
                                Else
                                    lossorgainamount = clsCommon.myCdbl(gvDetails.Rows(i).Cells(colAppliedAmt).Value) * dtLastRate.Rows(0).Item("ConvRate")
                                End If
                                'lossorgainamount = clsCommon.myCdbl(gvDetails.Rows(i).Cells(colAppliedAmt).Value) * dtLastRate.Rows(0).Item("ConvRate")
                                Totallossorgainamount = Totallossorgainamount + lossorgainamount
                            Next
                            Dim diff As Double = 0.0
                            If Totallossorgainamount <> 0 Then
                                'diff = obj.PAYMENT_AMOUNT_BASE_CURRENCY - Totallossorgainamount
                                diff = Math.Round(obj.PAYMENT_AMOUNT_BASE_CURRENCY, 2, MidpointRounding.AwayFromZero) - Math.Round(Totallossorgainamount, 2, MidpointRounding.AwayFromZero)
                                diff = Math.Round(diff, 2, MidpointRounding.AwayFromZero)
                            End If

                            If diff = 0 Then
                                obj.EXCHANGE_LOSS_AMT = 0
                                obj.EXCHANGE_GAIN_AMT = 0
                            ElseIf diff > 0 Then
                                If clsCommon.myLen(obj.EXCHANGE_LOSS_ACCOUNT) = 0 Then
                                    clsCommon.MyMessageBoxShow(Me, "Exchange Loss Account not defined.", Me.Text)
                                    Return False
                                End If
                                obj.EXCHANGE_LOSS_AMT = diff
                                obj.EXCHANGE_GAIN_AMT = 0
                            Else
                                If clsCommon.myLen(obj.EXCHANGE_GAIN_ACCOUNT) = 0 Then
                                    clsCommon.MyMessageBoxShow(Me, "Exchange Gain Account not defined.", Me.Text)
                                    Return False
                                End If
                                obj.EXCHANGE_LOSS_AMT = 0
                                obj.EXCHANGE_GAIN_AMT = -diff
                            End If
                        End If
                    Else
                        obj.EXCHANGE_LOSS_AMT = 0
                        obj.EXCHANGE_GAIN_AMT = 0
                        obj.EXCHANGE_GAIN_ACCOUNT = Nothing
                        obj.EXCHANGE_LOSS_ACCOUNT = Nothing
                    End If

                    '' End If


                    'If isPosted Then
                    '    '' only during posting of transaction
                    '    '' gather information for exchange gain and loss account
                    '    Dim dt As DataTable
                    '    dt = clsPaymentHeader.GetExchangeDetailDt(txtVendorCode.Value)
                    '    If dt.Rows.Count > 0 Then
                    '        obj.EXCHANGE_GAIN_ACCOUNT = clsCommon.myCstr(dt.Rows(0).Item("EXCHANGE_GAIN_ACCOUNT"))
                    '        obj.EXCHANGE_LOSS_ACCOUNT = clsCommon.myCstr(dt.Rows(0).Item("EXCHANGE_LOSS_ACCOUNT"))
                    '    Else
                    '        obj.EXCHANGE_GAIN_ACCOUNT = Nothing
                    '        obj.EXCHANGE_LOSS_ACCOUNT = Nothing
                    '    End If
                    '    Dim dtLastRate As DataTable
                    '    '' gather conv rate and amount of transaction to calculate exchange loss and gain
                    '    dtLastRate = clsPaymentHeader.GetExchangeRateAmount(clsCommon.myCstr(txtPaymentNo.Value))
                    '    If dtLastRate.Rows.Count > 0 Then
                    '        obj.ConvRateOld = dtLastRate.Rows(0).Item("ConvRate")
                    '        Dim diff As Double = 0.0
                    '        diff = obj.PAYMENT_AMOUNT_BASE_CURRENCY - dtLastRate.Rows(0).Item("payment_amount_base_currency")
                    '        If diff = 0 Then
                    '            obj.EXCHANGE_LOSS_AMT = 0
                    '            obj.EXCHANGE_GAIN_AMT = 0
                    '        ElseIf diff > 0 Then
                    '            If clsCommon.myLen(obj.EXCHANGE_LOSS_ACCOUNT) = 0 Then
                    '                clsCommon.MyMessageBoxShow("Exchange Loss Account not defined.")
                    '                Return False
                    '            End If
                    '            obj.EXCHANGE_LOSS_AMT = diff
                    '            obj.EXCHANGE_GAIN_AMT = 0
                    '        Else
                    '            If clsCommon.myLen(obj.EXCHANGE_GAIN_ACCOUNT) = 0 Then
                    '                clsCommon.MyMessageBoxShow("Exchange Gain Account not defined.")
                    '                Return False
                    '            End If
                    '            obj.EXCHANGE_LOSS_AMT = 0
                    '            obj.EXCHANGE_GAIN_AMT = -diff
                    '        End If
                    '    Else
                    '        obj.EXCHANGE_LOSS_AMT = 0
                    '        obj.EXCHANGE_GAIN_AMT = 0
                    '    End If
                    'Else
                    '    obj.EXCHANGE_LOSS_AMT = 0
                    '    obj.EXCHANGE_GAIN_AMT = 0
                    '    obj.EXCHANGE_GAIN_ACCOUNT = Nothing
                    '    obj.EXCHANGE_LOSS_ACCOUNT = Nothing
                    'End If


                Else
                    obj.CURRENCY_CODE = Nothing
                    obj.ConvRate = 1
                    obj.ConvRateOld = 1
                    obj.ApplicableFrom = Nothing
                    obj.BASE_CURRENCY_CODE = Nothing
                    obj.PAYMENT_AMOUNT_BASE_CURRENCY = 0
                    obj.EXCHANGE_LOSS_AMT = 0
                    obj.EXCHANGE_GAIN_AMT = 0
                    obj.EXCHANGE_GAIN_ACCOUNT = Nothing
                    obj.EXCHANGE_LOSS_ACCOUNT = Nothing
                End If
                '' end CurrencyConversion
                ''RICHA AGARWAL 22/07/2015
                ' obj.PAYMENT_AMOUNT_BASE_CURRENCY = (clsCommon.myCdbl(txtNetPayableAmt.Text) * clsCommon.myCdbl(txtConversionRate.Value))
                obj.PAYMENT_AMOUNT_BASE_CURRENCY = (clsCommon.myCdbl(txtPaymentAmt.Text) - clsCommon.myCdbl(txtTDSAmt.Text)) * clsCommon.myCdbl(txtConversionRate.Value)
                ''--------------------------

                ''richa agarwal
                GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
                If GSTStatus Then
                    If clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Then
                        obj.PurchaseOrder_Amount = clsCommon.myCdbl(LblPOTotalAmount.Text)
                        obj.PurchaseOrder_Add_Amount = clsCommon.myCdbl(lblPOTotalAdditionalCharge.Text)
                        obj.Tax_Amount_Advance = clsCommon.myCdbl(lblPOTotalTaxAmt.Text)
                        obj.Tax_Group = clsCommon.myCstr(txtTaxGroup.Value)
                        obj.PurchaseOrder_No_GST = clsCommon.myCstr(txtPONo_GST.Value)
                        obj.GSTRegistered = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("Select GSTRegistered  from TSPL_PURCHASE_ORDER_HEAD where PurchaseOrder_No ='" & obj.PurchaseOrder_No_GST & "' "))
                        obj.PO_Location_Code = clsCommon.myCstr(TxtPO_Location_GST.Value)
                        obj.ArrTrGST = New List(Of clsPaymentDetailGST)
                        Dim i As Integer = 0
                        For i = 0 To gvItem.Rows.Count - 1
                            Dim objTr As New clsPaymentDetailGST()

                            objTr.Document_Code = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolDocument_Code).Value)
                            objTr.Payment_No = clsCommon.myCstr(txtPaymentNo.Value)
                            objTr.Line_No = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolLine_No).Value)
                            objTr.Row_Type = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolRow_Type).Value)
                            objTr.Item_Code = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolItem_Code).Value)
                            objTr.Qty = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolQty).Value)
                            objTr.Balance_Qty = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolBalance_Qty).Value)
                            objTr.Item_Cost = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolItem_Cost).Value)
                            objTr.Unit_code = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolUnit_code).Value)
                            objTr.TAX1 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX1).Value)
                            objTr.TAX1_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Amt).Value)
                            objTr.TAX1_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Base_Amt).Value)
                            objTr.TAX1_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Rate).Value)
                            objTr.tax2 = clsCommon.myCstr(gvItem.Rows(i).Cells(Adcoltax2).Value)
                            objTr.TAX2_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Base_Amt).Value)
                            objTr.TAX2_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Rate).Value)
                            objTr.TAX2_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Amt).Value)
                            objTr.TAX3 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX3).Value)
                            objTr.TAX3_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Base_Amt).Value)
                            objTr.TAX3_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Rate).Value)
                            objTr.TAX3_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Amt).Value)
                            objTr.TAX4 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX4).Value)
                            objTr.TAX4_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Base_Amt).Value)
                            objTr.TAX4_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Rate).Value)
                            objTr.TAX4_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Amt).Value)
                            objTr.tax5 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX5).Value)
                            objTr.TAX5_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Base_Amt).Value)
                            objTr.TAX5_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Rate).Value)
                            objTr.TAX5_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Amt).Value)
                            objTr.TAX6_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Base_Amt).Value)
                            objTr.tax6 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX6).Value)
                            objTr.TAX6_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Rate).Value)
                            objTr.TAX6_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Amt).Value)
                            objTr.tax7 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX7).Value)
                            objTr.TAX7_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Base_Amt).Value)
                            objTr.TAX7_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Rate).Value)
                            objTr.TAX7_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Amt).Value)
                            objTr.tax8 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX8).Value)
                            objTr.TAX8_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Base_Amt).Value)
                            objTr.TAX8_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Rate).Value)
                            objTr.TAX8_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Amt).Value)
                            objTr.tax9 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX9).Value)
                            objTr.TAX9_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Base_Amt).Value)
                            objTr.TAX9_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Amt).Value)
                            objTr.TAX9_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Rate).Value)
                            objTr.tax10 = clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolTAX10).Value)
                            objTr.TAX10_Base_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Base_Amt).Value)
                            objTr.TAX10_Rate = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Rate).Value)
                            objTr.TAX10_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Amt).Value)
                            objTr.Amount = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolAmount).Value)
                            objTr.Disc_Per = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolDisc_Per).Value)
                            objTr.Disc_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolDisc_Amt).Value)
                            objTr.Amt_Less_Discount = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolAmt_Less_Discount).Value)
                            objTr.Total_Tax_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTotal_Tax_Amt).Value)
                            objTr.Item_Net_Amt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolItem_Net_Amt).Value)

                            objTr.TAX1_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Amt_Payment).Value)
                            objTr.TAX2_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Amt_Payment).Value)
                            objTr.TAX3_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Amt_Payment).Value)
                            objTr.TAX4_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Amt_Payment).Value)
                            objTr.TAX5_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Amt_Payment).Value)
                            objTr.TAX6_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Amt_Payment).Value)
                            objTr.TAX7_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Amt_Payment).Value)
                            objTr.TAX8_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Amt_Payment).Value)
                            objTr.TAX9_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Amt_Payment).Value)
                            objTr.TAX10_Amt_Payment = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Amt_Payment).Value)
                            objTr.PaymentAdvance = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value)
                            objTr.PaymentTotalTax = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalTax).Value)
                            objTr.PaymentTotalAdvanceAmt = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value)
                            obj.ArrTrGST.Add(objTr)
                        Next

                        For i = 0 To gvTaxDetail.Rows.Count - 1
                            If i = 0 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.TAX1 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX1_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX1_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX1_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If

                            If i = 1 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax2 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX2_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX2_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX2_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If

                            If i = 2 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.TAX3 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX3_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX3_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX3_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If

                            If i = 3 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.TAX4 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX4_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX4_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX4_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If

                            If i = 4 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax5 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX5_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX5_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX5_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If
                            If i = 5 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax6 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX6_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX6_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX6_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If

                            If i = 6 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax7 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX7_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX7_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX7_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If
                            If i = 7 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax8 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX8_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX8_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX8_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If
                            If i = 8 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax9 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX9_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX9_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX9_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If
                            If i = 9 Then
                                If clsCommon.myLen(clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)) > 0 Then
                                    obj.tax10 = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAutCode).Value)
                                    obj.TAX10_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxAmt).Value)
                                    obj.TAX10_Base_Amt = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTBaseAmt).Value)
                                    obj.TAX10_Rate = clsCommon.myCstr(gvTaxDetail.Rows(i).Cells(colTTaxRate).Value)
                                End If
                            End If


                        Next
                    End If

                End If

                ''----------
                If clsCommon.CompairString(obj.Payment_Type, "AD") = CompairStringResult.Equal Then
                    If clsCommon.myCstr(txtDescription.Text).Contains("Applied against Document No" + obj.Applied_Payment) Then
                        obj.Entry_Desc = txtDescription.Text
                    Else
                        If clsCommon.myCstr(txtDescription.Text).Contains("Applied") Then
                            Dim strindex As Integer = clsCommon.myCstr(txtDescription.Text).IndexOf("Applied")
                            txtDescription.Text = txtDescription.Text.Remove(6, txtDescription.Text.Length - 6)
                        End If
                        obj.Entry_Desc = txtDescription.Text + " Applied against Document No  " + obj.Applied_Payment
                    End If

                End If
                '' done by Panch Raj on 22-08-2017:tax on Bank Charges
                obj.Tax_Group_BankCharges = txtTaxGroupBankCharges.Value
                Dim objBCTtr As clsPaymentBankChargesTax
                Dim Bank_Charges_Tax As Double = 0
                For Each grow As GridViewRowInfo In gv2.Rows
                    objBCTtr = New clsPaymentBankChargesTax
                    objBCTtr.Payment_No = obj.Payment_No
                    objBCTtr.Line_No = gv2.Rows.IndexOf(grow) + 1
                    objBCTtr.Tax_Code = grow.Cells(colBCaxAutCode).Value
                    objBCTtr.Tax_Rate = grow.Cells(colBCaxRate).Value
                    objBCTtr.Tax_Base_Amount = grow.Cells(colBCBaseAmt).Value
                    objBCTtr.Tax_Amount = grow.Cells(colBCaxAmt).Value
                    Bank_Charges_Tax = Bank_Charges_Tax + objBCTtr.Tax_Amount
                    obj.objBCT.Add(objBCTtr)
                Next
                Dim PI_NO As New List(Of String)
                If txtmulPI.arrValueMember IsNot Nothing Then
                    For i As Integer = 0 To txtmulPI.arrValueMember.Count - 1
                        PI_NO.Add(txtmulPI.arrValueMember(i))
                    Next
                End If
                obj.ArrMul = New List(Of clsMultipleInvoicedetail)
                For i As Integer = 0 To PI_NO.Count - 1
                    Dim objtr As New clsMultipleInvoicedetail
                    objtr.PI_No = PI_NO(i)
                    obj.ArrMul.Add(objtr)
                Next
                obj.Bank_Charges_Tax = Bank_Charges_Tax

                obj.SaveData(obj, isNewEntry)
                UcAttachment1.SaveData(obj.Payment_No)
                ''approval work 11/02/2020
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
                    Dim xNewDesc As String = ""
                    xNewDesc = "Party Name : " + obj.Vendor_Name
                    ''=====================capex cond==============

                    xNewDesc = xNewDesc + Environment.NewLine + "Description : " + obj.Entry_Desc
                    clsApply_Approval.CheckApprovalRequired(MyBase.Form_ID, txtPaymentNo.Value, dtpPayment.Value, clsCommon.myCstr(xNewDesc), clsCommon.myCstr(txtDescription.Text), clsCommon.myCdbl(obj.Payment_Amount), 0, "")
                End If
                '================================================================

                LoadData(obj.Payment_No, NavigatorType.Current)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return True
    End Function


    Private Function AllowToSave() As Boolean
        Try
            'Dim isApplyBrachAccounting As Boolean = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ApplyBrachAccounting, clsFixedParameterCode.ApplyBrachAccounting, Nothing)) = 1, True, False)
            Dim Venagreement As String = ""
            Dim PaymentDate As Date = dtpPayment.Value.Date
            '' Anubhooti 13-Nov-2014 (Case:Apply Doc bank code should not be manadatory)
            btnsave.Focus()
            If clsCommon.myLen(txtDescription.Text) <= 0 And clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "TSDDCF") = CompairStringResult.Equal Then
                txtDescription.Focus()
                Throw New Exception("Description can't be blank")
            End If
            If txtBankCode.Value = "" Then
                txtBankCode.Focus()
                Throw New Exception("Bank Code can't be blank")
            ElseIf Not clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal And txtVendorCode.Value = "" Then
                txtVendorCode.Focus()
                Throw New Exception("Vendor Code can't be blank")
            ElseIf txtPaymentMode.Value = "" AndAlso clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
                txtPaymentMode.Focus()
                Throw New Exception("Payment Mode can't be blank")
            End If
            If objCommonVar.ApplyLocationWisePrefix Then
                If clsCommon.myLen(txtLocationPrefix.Value) <= 0 Then
                    txtLocationPrefix.Focus()
                    Throw New Exception("Please select Location")
                End If
            End If

            'Comment below check, asked by Ranjana mam
            'If clsCommon.myLen(txtVendorCode.Value) > 0 Then
            '    Venagreement = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Agreement  from TSPL_VENDOR_MASTER where Vendor_Code ='" & clsCommon.myCstr(txtVendorCode.Value) & "' And Form_Type ='PTM'"))
            '    If clsCommon.CompairString(Venagreement, "NO") = CompairStringResult.Equal Then
            '        txtPaymentMode.Focus()
            '        Throw New Exception("You can not make this entry because vendor doesn't has agreement")
            '    End If
            'End If

            ''richa agarwal 14/10/2014 aginst ticket no. BM00000004298
            'If clsCommon.myLen(txtVendorCode.Value) > 0 Then
            '    Venagreement = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Agreement  from TSPL_VENDOR_MASTER where Vendor_Code ='" & clsCommon.myCstr(txtVendorCode.Value) & "' And Form_Type ='TTM'"))
            '    If clsCommon.CompairString(Venagreement, "NO") = CompairStringResult.Equal Then
            '        txtPaymentMode.Focus()
            '        Throw New Exception("You can not make this entry because vendor doesn't has agreement")
            '    End If
            'End If
            ''============================
            '' Anubhooti 07-Apr-2015 (Conversion Rate Can not be <=0)
            If clsCommon.myCdbl(txtConversionRate.Text) <= 0 Then
                txtConversionRate.Focus()
                Throw New Exception("Conversion rate can not be zero or less than zero")
            End If
            If clsCommon.myCdbl(txtBankCharges.Text) > 0 Then
                If clsCommon.myLen(txtTaxGroupBankCharges.Value) <= 0 Then
                    txtTaxGroupBankCharges.Focus()
                    Throw New Exception("Select Tax Group for Bank Charges.")
                End If
            End If

            ''richa 17 July,2019 TEC/24/07/19-000958
            'Dim strERPStartDate As String = clsFixedParameter.GetData(clsFixedParameterType.ERPStartDate, clsFixedParameterCode.ERPStartDate, Nothing)
            If clsCommon.myLen(objCommonVar.ERPStartDate) > 0 Then
                Dim dtERPStartDate As DateTime = clsCommon.GetDateWithEndTime(objCommonVar.ERPStartDate).AddDays(-1)
                If clsCommon.myCDate(dtpPayment.Value, "dd/MM/yyyy") <= clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy") Then
                    If (clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal AndAlso ChkSecurity.Checked = True) Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
                    Else
                        ddlPaymentType.Focus()
                        Throw New Exception("Document Date should be Greater Than ERP Start Date " & clsCommon.GetPrintDate(dtERPStartDate, "dd/MM/yyyy") & "")
                    End If
                End If
            End If

            If (clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal) Then  '-----Payment
                If clsCommon.myLen(gvDetails.Rows.Count) <= 0 Then
                    Throw New Exception("Please fill atleast one row on grid.")
                End If

                If clsCommon.myCdbl(txtPaymentAmt.Text) < 0 Then
                    Throw New Exception("Payment amount should not be negative.")
                End If



                Dim Counter As Integer = 0
                Dim EnteredAmt As Double
                '----------Checks Balance Amount of Vendor Invoice-----------------
                Dim InvoiceDate As Date
                Dim isDebitNoteExist As Boolean = False
                Dim isCreditNoteExist As Boolean = False
                Dim isInvoiceExist As Boolean = False
                Dim isReceiptExist As Boolean = False
                Dim isDataInGrid As Boolean = False
                For Each grow As GridViewRowInfo In gvDetails.Rows
                    If clsCommon.CompairString("Yes", clsCommon.myCstr(grow.Cells(colApply).Value)) = CompairStringResult.Equal Then
                        isDataInGrid = True
                        If clsCommon.CompairString("Invoice", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            isInvoiceExist = True
                        ElseIf clsCommon.CompairString("Debit Note", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            isDebitNoteExist = True
                        ElseIf clsCommon.CompairString("Credit Note", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            isCreditNoteExist = True
                        ElseIf clsCommon.CompairString("Receipt Note", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            isReceiptExist = True
                        End If

                        Counter = Counter + 1
                        InvoiceDate = grow.Cells(colDocDate).Value
                        If CDate(Format(PaymentDate, "dd-MMM-yyyy")) < CDate(Format(InvoiceDate, "dd-MMM-yyyy")) Then
                            Throw New Exception("Invoice Date is '" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "', You can not pay for it on '" + clsCommon.GetPrintDate(PaymentDate, "dd/MMM/yyyy") + "'")
                        End If
                        'EnteredAmt += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)

                        If clsCommon.CompairString("Invoice", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            EnteredAmt += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        ElseIf clsCommon.CompairString("Debit Note", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            EnteredAmt -= clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        ElseIf clsCommon.CompairString("Credit Note", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            EnteredAmt += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        ElseIf clsCommon.CompairString("Receipt Note", clsCommon.myCstr(grow.Cells(colDocType).Value)) = CompairStringResult.Equal Then
                            EnteredAmt += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        End If



                        'Dim qry As String = "Select Balance_Amt-(Select ISNULL(SUM(Applied_Amount),0) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No= TSPL_VENDOR_INVOICE_HEAD.Document_No AND Payment_No<>'" + txtPaymentNo.Value + "' AND Post Not in ('1','Y')) as BalAmt  from "
                        'qry += " TSPL_VENDOR_INVOICE_HEAD WHere TSPL_VENDOR_INVOICE_HEAD.Document_No ='" + grow.Cells(colDocNo).Value + "'"
                        'Dim qry As String = "Select TSPL_VENDOR_INVOICE_HEAD.Document_Total -isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount ,0) -(Select ISNULL(SUM(Applied_Amount),0) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No= TSPL_VENDOR_INVOICE_HEAD.Document_No AND Payment_No<>'" + txtPaymentNo.Value + "' AND Post Not in ('1','Y')) as BalAmt  from " & _
                        ' " TSPL_VENDOR_INVOICE_HEAD WHere TSPL_VENDOR_INVOICE_HEAD.Document_No ='" + grow.Cells(colDocNo).Value + "'"
                        ''richa ERO/07/02/20-001184
                        Dim qry As String = "sELECT zzz.BalAmt from (Select TSPL_VENDOR_INVOICE_HEAD.Document_Total - CASE WHEN TSPL_VENDOR_INVOICE_HEAD.is_fOR_tds=0 THEN isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount ,0) ELSE 0 END  -(Select ISNULL(SUM(Applied_Amount),0) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No= TSPL_VENDOR_INVOICE_HEAD.Document_No AND Payment_No<>'" + txtPaymentNo.Value + "' AND Post Not in ('1','Y')) as BalAmt  from " &
                        " TSPL_VENDOR_INVOICE_HEAD WHere TSPL_VENDOR_INVOICE_HEAD.Document_No ='" + grow.Cells(colDocNo).Value + "'" & Environment.NewLine &
                        " Union All " & Environment.NewLine &
                        "  select  Payment_Amount -isnull((select sum(isnull(Applied_Amount,0)) from TSPL_PAYMENT_DETAIL where TSPL_PAYMENT_DETAIL.Document_No=TSPL_PAYMENT_HEADER.Payment_No  and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No ) AND TSPL_PAYMENT_DETAIL.Payment_No <>'" + txtPaymentNo.Value + "'),0)  as BalAmt from TSPL_PAYMENT_HEADER  WHERE  Payment_Type  ='RC' and IsChkReverse ='N' AND TSPL_PAYMENT_HEADER.Payment_No ='" + grow.Cells(colDocNo).Value + "')zzz"


                        If txtmulPI.arrValueMember IsNot Nothing AndAlso txtmulPI.arrValueMember.Count <= 0 Then
                            Dim availBal As Double = clsDBFuncationality.getSingleValue(qry)
                            If availBal < clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value) Then
                                Throw New Exception("Available Balance:  " + clsCommon.myCstr(availBal) + " Applied Amount: " + clsCommon.myCstr(grow.Cells(colAppliedAmt).Value) + " against Invoice: " + clsCommon.myCstr(grow.Cells(colDocNo).Value) + "")
                                grow.Cells(colTemp).Value = availBal
                                grow.Cells(colAppliedAmt).Value = 0
                                grow.Cells(colPendingAmt).Value = availBal
                                Return False
                            End If
                        End If
                    End If
                Next
                If isDataInGrid Then
                    If Not isInvoiceExist And Not isCreditNoteExist And Not isReceiptExist Then
                        If isDebitNoteExist Then
                            Throw New Exception("Can not Apply Only Debit note type document")
                        End If
                    End If
                End If

                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then '' Apply Document
                    If clsCommon.myLen(txtDocumentNo.Value) <= 0 Then
                        txtDocumentNo.Focus()
                        Throw New Exception("Document No can't be left blank")
                    ElseIf Math.Round(EnteredAmt, 2) > Math.Round(clsCommon.myCdbl(lblBalAmt.Text), 2) Then
                        '' Anubhooti 13-Nov-2014
                        Throw New Exception("Applied amount can't be greater than advance document amount")
                    End If
                End If
                '--------------------------------------------------------------------
                If Counter <= 0 Then
                    Throw New Exception("Please apply amount for atleast single document")
                End If
            End If
            'If clsCommon.myCdbl(txtPaymentAmt.Text) <= 0 Then
            '    Throw New Exception("Total Payment Amount Can Not Be 0 Or Negative")
            'End If
            '----------------This Cheks Wherther Applied Amount Is Greater Than Payment Amt Or Not
            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal Then
                If clsCommon.myLen(gvDetails.Rows(0).Cells(colGLAccount).Value) <= 0 Then
                    Throw New Exception("Please fill atleast one row on grid.")
                End If
                Dim totalAppliedAmt As Double = 0
                Dim arrAccountCode As New List(Of String)
                If gvDetails.Rows.Count > 0 Then
                    For Each grow As GridViewRowInfo In gvDetails.Rows
                        If clsCommon.myCstr(grow.Cells(colGLAccount).Value) <> "" Then
                            If clsCommon.myCdbl(grow.Cells(colAmount).Value) = 0 Then
                                Throw New Exception("Please enter amount against Account '" + grow.Cells(colGLAccount).Value + "' at line no '" + clsCommon.myCstr(grow.Cells(colLineNo).Value) + "'.")
                            End If
                        End If
                        If clsCommon.myCstr(grow.Cells(colGLAccount).Value) <> "" And clsCommon.myCstr(grow.Cells(colGLAccount)) <> "" Then
                            totalAppliedAmt = totalAppliedAmt + clsCommon.myCdbl(grow.Cells(colAmount).Value)
                        End If
                        Dim strACode As String = clsCommon.myCstr(grow.Cells(colGLAccount).Value)
                        If clsCommon.myLen(strACode) > 0 Then
                            If Not arrAccountCode.Contains(strACode) Then
                                arrAccountCode.Add(strACode)
                            End If
                        End If
                    Next
                Else
                    Throw New Exception("Please Enter Detail")
                End If

                If Math.Round(totalAppliedAmt, 2) <> Math.Round(clsCommon.myCdbl(txtPaymentAmt.Text), 2) Then
                    Throw New Exception("Total Applied Amount should be equal to Total Payment Amount.")
                End If
                If arrAccountCode IsNot Nothing AndAlso arrAccountCode.Count > 0 Then
                    If clsCommon.GetDateWithStartTime(dtpPayment.Value) >= clsCommon.GetDateWithStartTime(ERPStartDate) Then
                        ''RICHA AGARWAL TEC/03/07/19-000927
                        Dim qry As String = "select Account_Code from TSPL_GL_ACCOUNTS where Account_Code in (" + clsCommon.GetMulcallString(arrAccountCode) + ") and ControlAccount<>'N' AND TSPL_GL_ACCOUNTS.Account_Code NOT IN  (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForPayment =1)"
                        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
                        If clsCommon.myLen(clsCommon.myCstr(txtMPAdv.Value)) <= 0 Then
                            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                Throw New Exception("Can not select control Account -" + clsCommon.myCstr(dt.Rows(0)("Account_Code")))
                            End If
                        End If
                    End If
                End If
            End If
            '==================Addded By Rohit==================
            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then

                Dim TotalDebitAmt As Double = 0
                Dim TotalCreInvAmt As Double = 0
                Dim totalAppliedAmt As Double = 0
                Dim totalSecurityAmt As Double = 0
                'BM00000007735--By ANUBHUTI UDL/20/09/18-000221
                If gvDetails.Rows.Count > 0 Then
                    For Each grow As GridViewRowInfo In gvDetails.Rows
                        'If clsCommon.myCstr(grow.Cells(colGLAccount).Value) <> "" And clsCommon.myCstr(grow.Cells(colGLAccount)) <> "" Then

                        If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Debit Note") = CompairStringResult.Equal Then
                            TotalDebitAmt = TotalDebitAmt + clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        End If
                        If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Credit Note") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Invoice") = CompairStringResult.Equal OrElse clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Receipt Note") = CompairStringResult.Equal Then
                            TotalCreInvAmt = TotalCreInvAmt + clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        End If
                        ' totalAppliedAmt = totalAppliedAmt + clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        totalAppliedAmt = TotalCreInvAmt - TotalDebitAmt

                        totalSecurityAmt = totalSecurityAmt + clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
                        'End If
                    Next
                Else
                    Throw New Exception("Please Enter Detail")
                End If
                If txtmulPI.arrValueMember IsNot Nothing AndAlso txtmulPI.arrValueMember.Count <= 0 Then
                    If Math.Round(totalAppliedAmt + totalSecurityAmt, 2) <> Math.Round(clsCommon.myCdbl(txtPaymentAmt.Text), 2) Then
                        Throw New Exception("(Total Applied Amounty + Total Security Amount) should be equal to Total Payment Amount.")
                    End If
                End If
            End If
            '==================================================

            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                If clsCommon.myCdbl(txtPaymentAmt.Text) <= 0 Then
                    Throw New Exception("Please enter Total Payament Amount.")
                End If
                '' Anubhooti 08-Jan-2014 (BM00000005309)
                If clsCommon.myLen(txtlocation.Value) <= 0 AndAlso isApplyBranchAccounting = True Then
                    txtlocation.Focus()
                    Throw New Exception("Please fill location.")
                End If
            End If

            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
                If chkmemorndm.Checked = False AndAlso clsCommon.myCdbl(txtPaymentAmt.Text) <= 0 Then
                    Throw New Exception("Please enter Total Payment Amount.")
                End If
                '' Anubhooti 08-Jan-2014 (BM00000005309)
                If clsCommon.myLen(txtlocation.Value) <= 0 AndAlso isApplyBranchAccounting = True Then
                    txtlocation.Focus()
                    Throw New Exception("Please fill location.")
                End If
            End If

            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
                '' Rohiit 29-Apr-2014 (BM00000005309)
                If clsCommon.myLen(txtlocation.Value) <= 0 AndAlso isApplyBranchAccounting = True Then
                    txtlocation.Focus()
                    Throw New Exception("Please fill location.")
                End If
                If txtPaymentAmt.Value <= 0 Then
                    txtPaymentAmt.Focus()
                    Throw New Exception("Please enter Total Payment Amount.")
                End If
            End If

            '------Checks Whether ChequeNo is Blank Or Not, If CHeck ==Is Not Blank Then It Is Already Used Or Not---------
            Dim strcheckcode As String = connectSql.RunScalar("select Payment_Type  from TSPL_PAYMENT_CODE  where Payment_Code ='" + txtPaymentMode.Value + "'")
            If clsCommon.CompairString(strcheckcode, "Cheque") = CompairStringResult.Equal Or clsCommon.CompairString(txtPaymentMode.Value, "DD") = CompairStringResult.Equal Then
                If txtChequeNo.Text = "" Then
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
                        txtChequeNo.Focus()
                        Throw New Exception("Cheque No can't be blank")
                    Else
                        If (chkCheckPrint.Visible And chkCheckPrint.Checked = False) Then
                            txtChequeNo.Focus()
                            Throw New Exception("Cheque No can't be blank")
                        End If
                    End If


                Else

                    If AllowSameChequeNoForMultiplePaymentEntry = False Then
                        ''richa 9 Jan,2019 apply bank code and cheque no combination to check existing cheque number(as per Ranjana mam and Shruti Mam)  UDL/09/01/19-000250
                        Dim check As String = "Select Payment_No from TSPL_PAYMENT_HEADER Where Cheque_No='" + txtChequeNo.Text + "' and Bank_code='" & txtBankCode.Value & "' AND Payment_No <> '" + txtPaymentNo.Value + "'"
                        Dim chk As String = clsDBFuncationality.getSingleValue(check)
                        If clsCommon.myLen(chk) > 0 Then
                            txtChequeNo.Text = ""
                            txtChequeNo.Focus()
                            Throw New Exception("This Cheque No '" + txtChequeNo.Text + "' is Already Exists Against Payment No '" + chk + "'")
                        End If
                    End If


                    '' Anubhooti 08-Dec-2014 (Cheque No. Length is 6)

                    If clsCommon.myLen(txtChequeNo.Text) > 0 Then
                        If clsCommon.myLen(txtChequeNo.Text) < 6 Then
                            txtChequeNo.Focus()
                            Throw New Exception("Length of Cheque No should be of 6 digits.")
                        End If
                    End If

                End If
            ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
                'If chkCForm.Checked = True Then
                '' Anubhooti 08-Dec-2014 (Remove Invoice No.)
                'If clsCommon.myLen(txtCFormInvNo.Value) <= 0 Then
                '    common.clsCommon.MyMessageBoxShow("Please Select Invoice No for applying C Form.")
                '    Return False
                'End If
                'End If

            End If


            '----------------------------in case of memorandum--------------
            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal AndAlso chkmemorndm.Checked = True AndAlso clsCommon.myLen(txtmemoamt.Text.Trim().TrimEnd()) <= 0 Then
                Throw New Exception("For Memorandum Entry Please Fill Memorandum Amount")
                txtmemoamt.Focus()
                txtmemoamt.Select()

            End If

            '---------if payment amt enter then it should be equal to memorandum amt----------
            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal AndAlso chkmemorndm.Checked = True AndAlso clsCommon.myCdbl(txtPaymentAmt.Text) > 0 AndAlso clsCommon.myCdbl(txtPaymentAmt.Text) <> clsCommon.myCdbl(txtmemoamt.Text) Then
                Throw New Exception("For Memorandum Entry If Want To Enter Payment Amount Then It Should Be Equal To Memorandum Amount")
                txtPaymentAmt.Focus()
                txtPaymentAmt.Select()

            End If
            '---------------------------------------------------------
            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select bank_type from tspl_bank_master where BANK_CODE='" + txtBankCode.Value + "'")), "S") = CompairStringResult.Equal Then
                If Not clsCommon.CompairString("OA", clsCommon.myCstr(ddlPaymentType.SelectedValue)) = CompairStringResult.Equal Then
                    txtBankCode.Focus()
                    Throw New Exception("Transaction type should be on Account for settlement bank.")
                End If
            End If

            GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
            If GSTStatus Then
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal And clsCommon.myLen(txtPONo_GST.Value) > 0 Then
                    If clsCommon.myCdbl(txtPaymentAmt.Value) > clsCommon.myCdbl(LblPOTotalAmount.Text) Then
                        txtPaymentAmt.Focus()
                        Throw New Exception("Payment Amount cannot be greater than PO Amount")
                    End If

                    Dim strCustmer As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Code from TSPL_PURCHASE_ORDER_HEAD where PurchaseOrder_No='" & clsCommon.myCstr(txtPONo_GST.Value) & "' "))
                    If clsCommon.CompairString(strCustmer, txtVendorCode.Value) <> CompairStringResult.Equal Then
                        txtVendorCode.Focus()
                        Throw New Exception("Vendor of Payment should be same as PO")
                    End If
                End If
            End If

            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(ddlEmployeeAdvanceType.SelectedValue, "") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(ddlEmployeeType.SelectedValue, "") <> CompairStringResult.Equal Then
                        Throw New Exception("Please select any one option at a time Employee Expense type/Employee Advance type.")
                    End If
                End If
            End If
            ''richa agarwal 22 May,2018 ERO/21/05/18-000318,ERO/20/06/18-000357
            Dim AllowtoSetPaymentAmountForCashTransaction As Double = clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.AllowtoSetPaymentAmountForCashTransaction, clsFixedParameterCode.AllowtoSetPaymentAmountForCashTransaction, Nothing))
            If AllowtoSetPaymentAmountForCashTransaction > 0 And clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
                If clsCommon.CompairString(clsCommon.myCstr(txtPaymentMode.Value), clsDBFuncationality.getSingleValue("select Payment_Code from tspl_payment_code where Payment_Type ='Cash'")) = CompairStringResult.Equal Then
                    Dim dblTotalreceiptCashAmount As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("sELECT ISNULL(sum(Payment_Amount+TDS_Amount),0) FROM TSPL_PAYMENT_HEADER WHERE Vendor_Code='" & clsCommon.myCstr(txtVendorCode.Value) & "' AND convert(date,Payment_Date,103) ='" & clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") & "' and Payment_Code='CASH' and TSPL_PAYMENT_HEADER.Payment_Type  <>'AD' and Payment_No <>'" & clsCommon.myCstr(txtPaymentNo.Value) & "'"))
                    If (dblTotalreceiptCashAmount + clsCommon.myCdbl(txtPaymentAmt.Value)) > AllowtoSetPaymentAmountForCashTransaction Then
                        Throw New Exception("You cannot create Payment entry of amount greater than " & AllowtoSetPaymentAmountForCashTransaction & " for Payment mode Cash for each Vendor.")
                    End If
                End If
            End If
            ''---------------
            'If clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Is_Provisional from TSPL_VENDOR_MASTER where Vendor_Code='" + txtVendorCode.Value + "'")) = 1 Then
            '    If common.clsCommon.MyMessageBoxShow("Do You Want to Apply TDS Provision", "", MessageBoxButtons.YesNo, RadMessageIcon.Question) = DialogResult.Yes Then
            '        chkTDSProvision.Checked = True
            '    Else
            '        chkTDSProvision.Checked = False
            '    End If

            'End If

            ''richa agarwal  to check bank balance
            If Not isFlag Then
                CheckNegativeBankBalance()
            End If

            If ChkSecurity.Checked And ChkRetention.Checked Then
                Throw New Exception("Select One Checkbox At a time.")
            End If

            ''--------
            UcCustomFields1.AllowToSave()
            UcAttachment1.AllowToSave()
            Return True
            '---------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        End Try
        Return True
    End Function
    ''richa agarwal create function to check bank balance on Delete 
    Public Function CheckNegativeBankBalanceondelete() As Boolean
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
            Dim Bank_Type_Check As String = "0"
            Bank_Type_Check = clsFixedParameter.GetData(clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, Nothing)
            Dim Bank_Type As String = clsBankMaster.GetBankType(txtBankCode.Value, Nothing)
            Dim Bank_Balance As Decimal = 0
            Dim Bank_Location As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code in (select BANKACC from TSPL_BANK_MASTER where BANK_CODE='" & txtBankCode.Value & "')", Nothing))
            If clsCommon.CompairString(Bank_Type_Check, "0") = CompairStringResult.Equal Then
                '' allow for all
            ElseIf clsCommon.CompairString(Bank_Type_Check, "1") = CompairStringResult.Equal Then
                If clsCommon.CompairString(Bank_Type, "C") = CompairStringResult.Equal OrElse clsCommon.CompairString(Bank_Type, "P") = CompairStringResult.Equal Then
                    Bank_Balance = clsPaymentHeader.GetBankBalance(txtPaymentNo.Value, dtpPayment.Value, txtBankCode.Value, Bank_Location, Nothing, True)
                    If Bank_Balance < clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) Then
                        Throw New Exception("Payment Amount : " & clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) & " Available Bank Balance(" & Bank_Location & ") : " & Bank_Balance & "")
                    End If
                End If
            ElseIf clsCommon.CompairString(Bank_Type_Check, "2") = CompairStringResult.Equal Then
                If clsCommon.CompairString(Bank_Type, "B") = CompairStringResult.Equal Then
                    Bank_Balance = clsPaymentHeader.GetBankBalance(txtPaymentNo.Value, dtpPayment.Value, txtBankCode.Value, Bank_Location, Nothing, True)
                    If Bank_Balance < clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) Then
                        Throw New Exception("Payment Amount : " & clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) & " Available Bank Balance(" & Bank_Location & ") : " & Bank_Balance & "")
                    End If
                End If
            ElseIf clsCommon.CompairString(Bank_Type_Check, "3") = CompairStringResult.Equal Then
                Bank_Balance = clsPaymentHeader.GetBankBalance(txtPaymentNo.Value, dtpPayment.Value, txtBankCode.Value, Bank_Location, Nothing, True)
                If Bank_Balance < clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) Then
                    Throw New Exception("Payment Amount : " & clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) & " Available Bank Balance(" & Bank_Location & ") : " & Bank_Balance & "")
                End If
            End If
        End If
        Return True
    End Function



    ''richa agarwal create function to check bank balance on save 
    Public Function CheckNegativeBankBalance() As Boolean
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
            Dim Bank_Type_Check As String = "0"
            Bank_Type_Check = clsFixedParameter.GetData(clsFixedParameterType.StopNegativeBankBalance, clsFixedParameterCode.StopNegativeBankBalance, Nothing)
            Dim Bank_Type As String = clsBankMaster.GetBankType(txtBankCode.Value, Nothing)
            Dim Bank_Balance As Decimal = 0
            Dim Bank_Location As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Account_Seg_Code7 from TSPL_GL_ACCOUNTS where Account_Code in (select BANKACC from TSPL_BANK_MASTER where BANK_CODE='" & txtBankCode.Value & "')", Nothing))
            If clsCommon.CompairString(Bank_Type_Check, "0") = CompairStringResult.Equal Then
                '' allow for all
            ElseIf clsCommon.CompairString(Bank_Type_Check, "1") = CompairStringResult.Equal Then
                If clsCommon.CompairString(Bank_Type, "C") = CompairStringResult.Equal OrElse clsCommon.CompairString(Bank_Type, "P") = CompairStringResult.Equal Then
                    Bank_Balance = clsPaymentHeader.GetBankBalance(txtPaymentNo.Value, dtpPayment.Value, txtBankCode.Value, Bank_Location, Nothing)
                    If Bank_Balance < clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) Then
                        Throw New Exception("Payment Amount : " & clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) & " Available Bank Balance(" & Bank_Location & ") : " & Bank_Balance & "")
                    End If
                End If
            ElseIf clsCommon.CompairString(Bank_Type_Check, "2") = CompairStringResult.Equal Then
                If clsCommon.CompairString(Bank_Type, "B") = CompairStringResult.Equal Then
                    Bank_Balance = clsPaymentHeader.GetBankBalance(txtPaymentNo.Value, dtpPayment.Value, txtBankCode.Value, Bank_Location, Nothing)
                    If Bank_Balance < clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) Then
                        Throw New Exception("Payment Amount : " & clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) & " Available Bank Balance(" & Bank_Location & ") : " & Bank_Balance & "")
                    End If
                End If
            ElseIf clsCommon.CompairString(Bank_Type_Check, "3") = CompairStringResult.Equal Then
                Bank_Balance = clsPaymentHeader.GetBankBalance(txtPaymentNo.Value, dtpPayment.Value, txtBankCode.Value, Bank_Location, Nothing)
                If Bank_Balance < clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) Then
                    Throw New Exception("Payment Amount : " & clsCommon.myCdbl(Me.txtTotalPaymentBaseCurr.Text) & " Available Bank Balance(" & Bank_Location & ") : " & Bank_Balance & "")
                End If
            End If
        End If
        Return True
    End Function


    ''--------------
    Private Sub txtPaymentNo__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtPaymentNo._MYValidating
        'done by priti KDI/04/07/18-000388 for updating vendor name from master
        Dim qry As String = "select TSPL_PAYMENT_HEADER.Payment_No as [Code] ,convert(varchar,TSPL_PAYMENT_HEADER.Payment_Date,103) as [Date],FORMAT(CAST(Payment_Date AS DATETIME),'dd/MM/yyyy hh:mm tt') as [Date & Time] ,TSPL_PAYMENT_HEADER.Entry_Desc as [Description],TSPL_PAYMENT_HEADER.Vendor_Code as [Vendor Code] ,TSPL_VENDOR_MASTER.Vendor_Name as [Vendor Name],TSPL_PAYMENT_HEADER.Location_Code_Prefix as [Location Code Prefix],ISNULL(TSPL_VENDOR_MASTER.alies_name,'') As [Alies Name] ,TSPL_PAYMENT_HEADER.Payment_Post_Date as [Payment Post Date] ,TSPL_PAYMENT_HEADER.Bank_Code as [Bank Code] ,TSPL_PAYMENT_HEADER.Payment_Type as [Payment Type] ,TSPL_PAYMENT_HEADER.Remit_To as [Remit To]  ,TSPL_PAYMENT_HEADER.Reference as [Reference] ,TSPL_PAYMENT_HEADER.Narration as [Narration] ,TSPL_PAYMENT_HEADER.Payment_Code as [Payment Code] ,TSPL_PAYMENT_HEADER.Cheque_No as [Cheque No] ,TSPL_PAYMENT_HEADER.Cheque_Date as [Cheque Date] ,TSPL_PAYMENT_HEADER.Payment_Amount as [Payment Amount] ,TSPL_PAYMENT_HEADER.Vendor_Account_Set as [Vendor Account Set] ,TSPL_PAYMENT_HEADER.TDS_Amount as [Tds Amount] ,TSPL_PAYMENT_HEADER.Total_Prepayment as [Total Prepayment] ,TSPL_PAYMENT_HEADER.Apply_By as [Apply By] ,TSPL_PAYMENT_HEADER.Apply_To as [Apply To] ,TSPL_PAYMENT_HEADER.Posted as [Posted] ,TSPL_PAYMENT_HEADER.Created_By as [Created By] ,TSPL_PAYMENT_HEADER.Created_Date as [Created Date] ,TSPL_PAYMENT_HEADER.Modify_By as [Modify By] ,TSPL_PAYMENT_HEADER.Modify_Date as [Modify Date] ,TSPL_PAYMENT_HEADER.Level1_User_code as [Level1 User Code] ,TSPL_PAYMENT_HEADER.Level2_User_code as [Level2 User Code] ,TSPL_PAYMENT_HEADER.Level3_User_code as [Level3 User Code] ,TSPL_PAYMENT_HEADER.Level4_User_code as [Level4 User Code] ,TSPL_PAYMENT_HEADER.Level5_User_code as [Level5 User Code] ,TSPL_PAYMENT_HEADER.Comp_Code as [Comp Code] ,TSPL_PAYMENT_HEADER.Debit_Account as [Debit Account] ,TSPL_PAYMENT_HEADER.Credit_Account as [Credit Account] ,TSPL_PAYMENT_HEADER.Balance_Amt as [Balance Amt] ,TSPL_PAYMENT_HEADER.Total_Applied_Amount as [Total Applied Amount] ,TSPL_PAYMENT_HEADER.Transport_Id as [Transport Id] ,TSPL_PAYMENT_HEADER.FIFO_Balance as [Fifo Balance] ,TSPL_PAYMENT_HEADER.QuickEntryNo as [Quickentryno] ,TSPL_PAYMENT_HEADER.LoadOutNo as [Loadoutno] ,TSPL_PAYMENT_HEADER.Salesman_Code as [Salesman Code] ,TSPL_PAYMENT_HEADER.Salesman_Name as [Salesman Name] ,TSPL_PAYMENT_HEADER.Route_NO as [Route No] ,TSPL_PAYMENT_HEADER.Route_Description as [Route Description] ,TSPL_PAYMENT_HEADER.Location_Code as [Location Code] ,TSPL_PAYMENT_HEADER.Location_Description as [Location Description] ,TSPL_PAYMENT_HEADER.IsRecoCleared as [Isrecocleared] ,TSPL_PAYMENT_HEADER.IsChkReverse as [Ischkreverse] ,TSPL_PAYMENT_HEADER.Loadout_No as [Loadout No] ,TSPL_PAYMENT_HEADER.Bank_Charges_Ac as [Bank Charges Ac] ,TSPL_PAYMENT_HEADER.Bank_Charges as [Bank Charges] ,TSPL_PAYMENT_HEADER.CURRENCY_CODE as [Currency Code] ,TSPL_PAYMENT_HEADER.ConvRate as [Convrate] ,TSPL_PAYMENT_HEADER.ApplicableFrom as [Applicablefrom] ,TSPL_PAYMENT_HEADER.BASE_CURRENCY_CODE as [Base Currency Code] ,TSPL_PAYMENT_HEADER.PAYMENT_AMOUNT_BASE_CURRENCY as [Payment Amount Base Currency] ,TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_AMT as [Exchange Loss Amt] ,TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_AMT as [Exchange Gain Amt] ,TSPL_PAYMENT_HEADER.EXCHANGE_LOSS_ACCOUNT as [Exchange Loss Account] ,TSPL_PAYMENT_HEADER.EXCHANGE_GAIN_ACCOUNT as [Exchange Gain Account] ,TSPL_PAYMENT_HEADER.ConvRateOld as [Convrateold] ,TSPL_PAYMENT_HEADER.CFormRecd as [Cformrecd] ,TSPL_PAYMENT_HEADER.CForm_InvoiceNo as [Cform Invoiceno] ,TSPL_PAYMENT_HEADER.EMP_CODE as [Emp Code] ,TSPL_PAYMENT_HEADER.PROJECT_CODE as [Project Code] ,TSPL_PAYMENT_HEADER.PDC_Cheque as [Pdc Cheque] ,TSPL_PAYMENT_HEADER.Document_No as [Document No] ,TSPL_PAYMENT_HEADER.CHECK_PRINT as [Check Print] ,TSPL_PAYMENT_HEADER.CHECK_CODE as [Check Code] ,TSPL_PAYMENT_HEADER.memorandum_amt as [Memorandum Amt] ,TSPL_PAYMENT_HEADER.Applied_Payment as [Applied Payment],TSPL_PAYMENT_HEADER.PurchaseOrder_No_GST AS [Purchase Order No]  From TSPL_PAYMENT_HEADER " &
        " LEFT OUTER JOIN TSPL_BANK_MASTER ON TSPL_BANK_MASTER.Bank_Code=TSPL_PAYMENT_HEADER.Bank_Code"
        '' Anubhooti 12-Mar-2015 (Fetch Alies Name On Vendor Finder)
        qry += " LEFT OUTER JOIN TSPL_VENDOR_MASTER ON TSPL_VENDOR_MASTER.Vendor_Code = TSPL_PAYMENT_HEADER.Vendor_Code "

        Dim Bank_Code As String = FrmMainTranScreen.bankPermission(Nothing)
        Dim strWhrclas As String = "1=1"
        If clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PermissionSettingForTransactionWithBank, clsFixedParameterType.PermissionSettingForTransactionWithBank, Nothing)) = 1 Then
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                strWhrclas += " AND RIGHT(TSPL_BANK_MASTER.BANKACC,3) in (" + objCommonVar.strCurrUserLocationsSegment + ")"
            End If
        ElseIf clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.PermissionSettingForTransactionWithBank, clsFixedParameterType.PermissionSettingForTransactionWithBank, Nothing)) = 1 Then
            If clsCommon.myLen(Bank_Code) > 0 Then
                strWhrclas += " AND TSPL_PAYMENT_HEADER.Bank_Code in ( " + Bank_Code + " )"
            End If
        End If
        LoadData(clsCommon.ShowSelectForm("PMNTFND@Payment", qry, "Code", strWhrclas, txtPaymentNo.Value, "Code", isButtonClicked, "TSPL_PAYMENT_HEADER.Payment_Date"), NavigatorType.Current)
    End Sub

    Private Sub txtPaymentNo__MYNavigator(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal NavType As common.NavigatorType) Handles txtPaymentNo._MYNavigator
        Try
            Dim qst As String = "select count(*) from TSPL_PAYMENT_HEADER where Payment_No='" + txtPaymentNo.Value + "'"
            Dim count As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(qst))
            If count = 0 Then
                txtPaymentNo.MyReadOnly = False
            Else
                txtPaymentNo.MyReadOnly = True
            End If
            LoadData(txtPaymentNo.Value, NavType)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub LoadData(ByVal strDocumentNo As String, ByVal navType As common.NavigatorType)
        Try
            btnViewTDSDetails.Enabled = False
            loadPaymentType()
            Dim DblPaymentTaxAmt1 As Double = 0
            Dim DblPaymentTaxAmt2 As Double = 0
            Dim DblPaymentTaxAm3 As Double = 0
            Dim DblPaymentTaxAmt4 As Double = 0
            Dim DblPaymentTaxAmt5 As Double = 0
            Dim DblPaymentTaxAmt6 As Double = 0
            Dim DblPaymentTaxAmt7 As Double = 0
            Dim DblPaymentTaxAmt8 As Double = 0
            Dim DblPaymentTaxAmt9 As Double = 0
            Dim DblPaymentTaxAmt10 As Double = 0
            Dim obj As New clsPaymentHeader()
            obj = clsPaymentHeader.GetData(strDocumentNo, navType)
            If (obj IsNot Nothing AndAlso clsCommon.myLen(obj.Payment_No) > 0) Then

                txtLocationPrefix.Value = obj.Location_Code_Prefix
                If clsCommon.myLen(clsCommon.myCstr(obj.Location_Code_Prefix)) > 0 Then
                    txtLocationPrefixName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" select Location_Desc  from TSPL_LOCATION_MASTER WHERE  Location_Code='" & txtLocationPrefix.Value & "' "))
                Else
                    txtLocationPrefixName.Text = ""
                End If
                txtPaymentNo.Value = obj.Payment_No
                txtDescription.Text = obj.Entry_Desc
                dtpPayment.Value = obj.Payment_Date
                chkIsReceipt.Visible = False
                If clsFixedParameter.GetData(clsFixedParameterType.AlowwdateChangeinPaymentEntry, clsFixedParameterCode.AlowwdateChangeinPaymentEntry, Nothing) Then
                    dtpPayment.Enabled = False
                End If
                txtBankCode.Value = obj.Bank_Code
                lblBankDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select DESCRIPTION from TSPL_BANK_MASTER  WHere BANK_CODE='" + obj.Bank_Code + "'"))
                ChkSecurity.Checked = False
                ChkRetention.Checked = False
                If clsCommon.CompairString(obj.Payment_Type, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(obj.Payment_Type), "RC") = CompairStringResult.Equal Then 'If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    ChkSecurity.Visible = True
                    ChkSecurity.Checked = IIf(obj.Is_Security = 1, True, False)
                Else
                    ChkSecurity.Visible = False
                End If
                If clsCommon.CompairString(obj.Payment_Type, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(clsCommon.myCstr(obj.Payment_Type), "RC") = CompairStringResult.Equal Then 'If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    ChkRetention.Visible = True
                    ChkRetention.Checked = IIf(obj.Is_Retention = 1, True, False)
                Else
                    ChkRetention.Visible = False
                End If
                chkSaving.Checked = obj.Saving
                chkFarmerLoanPayment.Checked = IIf(obj.isFarmerLoanPayment = 1, True, False)

                ddlPaymentType.SelectedValue = clsCommon.myCstr(obj.Payment_Type)
                txtPaymentMode.Value = obj.Payment_Code
                Dim strcheckcode As String = connectSql.RunScalar("select Payment_Type  from TSPL_PAYMENT_CODE  where Payment_Code ='" + txtPaymentMode.Value + "'")
                If clsCommon.CompairString(strcheckcode, "Cheque") = CompairStringResult.Equal Then
                    If txtChequeNo.Text = "" Then
                        btnChqUpdate.Visible = True
                    Else
                        btnChqUpdate.Visible = False
                    End If
                End If
                chkOpening.Checked = obj.is_Opening
                ddlEmployeeType.SelectedValue = obj.Employee_Type
                ddlEmployeeAdvanceType.SelectedValue = obj.Employee_Advance_Type
                If clsCommon.myLen(obj.DateAndTime) > 0 Then
                    txtDataAndTimeSelection.Value = obj.DateAndTime
                    txtDataAndTimeSelection.Checked = True
                End If
                txtTapalNo.Text = clsCommon.myCstr(obj.TapalNo)
                '' Anubhooti 13-Nov-2014
                If clsCommon.CompairString(clsCommon.myCstr(obj.Payment_Type), "AD") = CompairStringResult.Equal Then
                    'txtBankCode.MendatroryField = False
                    '' Anubhooti 14-Nov-2014
                    txtPaymentMode.Value = ""
                    'lblbankcode.Visible = False
                    'txtBankCode.Visible = False
                    'lblBankDesc.Visible = False
                    lblpaymentcode.Visible = False
                    txtPaymentMode.Visible = False
                    pnlCheque.Visible = False
                    txtBankCharges.Visible = False
                    MyLabel3.Visible = False
                    btnprint.Enabled = False
                Else
                    'txtBankCode.MendatroryField = True
                    '' Anubhooti 14-Nov-2014
                    'lblbankcode.Visible = True
                    'txtBankCode.Visible = True
                    'lblBankDesc.Visible = True
                    lblpaymentcode.Visible = True
                    txtPaymentMode.Visible = True
                    pnlCheque.Visible = True
                    txtBankCharges.Visible = True
                    MyLabel3.Visible = True
                    btnprint.Enabled = True
                End If
                ''
                HandleCheque()  '-----------------Enables/Disables Cheque Utilities
                txtChequeNo.Text = obj.Cheque_No
                '' check printing
                chkCheckPrint.Checked = IIf(obj.CHECK_PRINT = 1, True, False)
                If (obj.CHECK_PRINT = 1 Or clsCommon.myLen(obj.Cheque_No) > 0) Then
                    Me.btnPrintCheck.Enabled = True
                    btnVoidCheck.Enabled = True
                Else
                    Me.btnPrintCheck.Enabled = False
                    btnVoidCheck.Enabled = False
                End If


                txtVendor_bankcode.Text = obj.Vendor_Bank_Code
                TxtVendor_BankName.Text = obj.Vendor_Bank_Name
                TxtVendorBank_IFSCCode.Text = obj.Vendor_IFSC_Code
                txtVendorBank_branchname.Text = obj.Vendor_Branch_Name
                txtVendor_Bank_ACNo.Text = obj.Vendor_Bank_ACNo

                ChkAccPayee.Checked = IIf(obj.Account_Payee = 1, True, False)
                chkTDSProvision.Checked = obj.TDS_Provision
                If clsCommon.myLen(obj.Cheque_Date) > 0 Then
                    dtpChequeDate.Value = obj.Cheque_Date
                    If clsCommon.CompairString(obj.PDC_Cheque, "Y") = CompairStringResult.Equal Then
                        chkPDC.Checked = True
                    Else
                        chkPDC.Checked = False
                    End If
                Else
                    chkPDC.Checked = False
                    dtpChequeDate.Value = Nothing
                End If
                txtVendorCode.Value = obj.Vendor_Code
                FillCustomerOutStanding(txtVendorCode.Value) '-----------------This Finction--Loads Vendor's Outstanding as Customer.
                'txtVendorCode.Enabled = False
                lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Name from TSPL_VENDOR_MASTER Where Vendor_Code='" + obj.Vendor_Code + "'"))
                If (clsCommon.CompairString(obj.Payment_Type, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "AD") = CompairStringResult.Equal) Then ' Or clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal) Then
                    txtPaymentAmt.Text = obj.Payment_Amount
                    If clsCommon.CompairString(obj.Payment_Type, "AD") = CompairStringResult.Equal Then
                        txtDocumentNo.Value = obj.Applied_Payment
                        lblBalAmt.Text = obj.Balance_Amt
                    End If
                    '' Anubhooti 27-Mar-2015 (Advance against salary checkbox will only be visible in case of Advance & On-Account)
                    If clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                        ChkAdvSalary.Visible = True
                        ChkAdvSalary.Checked = IIf(obj.Advance_Against_Salary = 1, True, False)
                    Else
                        ChkAdvSalary.Visible = False
                        ChkAdvSalary.Checked = False
                    End If
                ElseIf (clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal) Then
                    txtPaymentAmt.Text = clsCommon.myCdbl(obj.Total_Prepayment)
                    txtTDSAmt.Text = clsCommon.myCdbl(obj.TDS_Amount)
                    txtNetPayableAmt.Text = clsCommon.myCstr(obj.Payment_Amount)
                    Dim objVendor As clsTDSVendorDetails = clsTDSVendorDetails.GetData(txtVendorCode.Value)
                    If objVendor IsNot Nothing Then
                        btnViewTDSDetails.Enabled = True
                    End If
                    '' Anubhooti 27-Mar-2015 (Advance against salary checkbox will only be visible in case of Advance & On-Account)
                    ChkAdvSalary.Visible = True
                    ChkAdvSalary.Checked = IIf(obj.Advance_Against_Salary = 1, True, False)
                    ''
                ElseIf clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "SR") = CompairStringResult.Equal Then
                    txtPaymentAmt.Text = clsCommon.myCstr(obj.Payment_Amount)
                    If clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                        ChkAdvSalary.Visible = True
                        ChkAdvSalary.Checked = IIf(obj.Advance_Against_Salary = 1, True, False)
                    End If
                ElseIf clsCommon.CompairString(obj.Payment_Type, "MI") = CompairStringResult.Equal Then
                    txtPaymentAmt.Text = obj.Payment_Amount
                    txtTotalAppliedAmt.Text = obj.Total_Applied_Amount
                    If clsCommon.myLen(obj.Loadout_No) > 0 Then
                        txtLoadOutno.Value = obj.Loadout_No
                    Else
                        txtLoadOutno.Value = ""
                    End If
                    txtMPAdv.Value = obj.MP_Code_For_Advance
                    chkIsReceipt.Checked = IIf(obj.isReceipt = 1, True, False)
                    chkIsReceipt.Visible = True
                    lblEmpCode.Text = obj.EMP_CODE
                    lblProjCode.Text = obj.PROJECT_CODE
                    lblEmpDesc.Text = clsDBFuncationality.getSingleValue("select Emp_Name from TSPL_EMPLOYEE_MASTER where EMP_CODE='" + lblEmpCode.Text + "'")
                    lblProjDesc.Text = clsDBFuncationality.getSingleValue("select SPECIFICATION from TSPL_PJC_PROJECT where PROJECT_CODE='" + lblProjCode.Text + "'")
                    If clsCommon.myLen(lblProjCode.Text) > 0 Then
                        pnlPJC.Visible = True
                        lblEmpCode.Visible = True
                    End If

                End If

                txtRemitTo.Text = clsCommon.myCstr(obj.Remit_To)
                txtBankCharges.Text = obj.Bank_Charges
                If clsCommon.CompairString(obj.Posted, "P") = CompairStringResult.Equal OrElse clsCommon.CompairString(obj.Posted, "1") = CompairStringResult.Equal Then
                    UsLock1.Status = ERPTransactionStatus.Approved
                    btnsave.Enabled = False
                    btnpost.Enabled = False
                    btndelete.Enabled = False
                    If CostCenterAndHirerachyCodeUpdateAfterPost = True Then
                        butCostCenterAndHirerachy_Update_AfterPost.Visible = True
                    End If
                Else
                    UsLock1.Status = ERPTransactionStatus.Pending
                    btnsave.Enabled = True
                    btndelete.Enabled = True
                    btnpost.Enabled = True
                    butCostCenterAndHirerachy_Update_AfterPost.Visible = False
                End If

                chkmemorndm.Checked = False
                txtmemoamt.Text = obj.memorndmamt
                If clsCommon.myCdbl(txtmemoamt.Text) > 0 Then
                    chkmemorndm.Checked = True
                End If
                '' Anubhooti 23-Nov-2014 BM00000004668 (Remove C-From)
                'chkCForm.IsChecked = IIf(obj.CFormRecd = "Y", True, False)
                '' Anubhooti 08-Dec-2014 (Remove Invoice No.)
                'txtCFormInvNo.Value = obj.CForm_InvoiceNo
                '' Anubhooti 21-Aug-2014
                txtPONo.Value = obj.PurchaseOrder_No

                '' Anubhooti 25-Sep-2014 BM00000004050
                LblAccPayee.Text = clsCommon.myCstr(obj.Account_Payee_Name)
                ''shivani
                fndloanNo.Value = obj.Loan_Code

                txtInterestRate.Value = obj.Interest_Rate
                txtNoOfEMI.Value = obj.No_Of_EMI

                '' Anubhooti 17-Nov-2014 BM00000005309
                If ((clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Or (clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal) Or clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal)) AndAlso isApplyBranchAccounting = True Then
                    RadLabel18.Visible = True
                    txtlocation.Visible = True
                    LblLocDesp.Visible = True
                    'txtlocation.Enabled = False
                    txtlocation.Value = clsCommon.myCstr(obj.Location_GL_Code)
                    If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
                        LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
                    Else
                        LblLocDesp.Text = ""
                    End If

                Else
                    RadLabel18.Visible = False
                    txtlocation.Visible = False
                    LblLocDesp.Visible = False
                    txtlocation.Enabled = True

                End If
                chkBankChargesWaveOff.Checked = IIf(obj.WaveOFFBankCharges = "Y", True, False)
                ''
                IsInsideLoadData = True
                LoadBlankGrid(ddlPaymentType.SelectedValue)
                If (clsCommon.CompairString(obj.Payment_Type, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "SR") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "AD") = CompairStringResult.Equal) Then
                    For Each objTr As clsPaymentDetail In obj.ArrTr
                        gvDetails.Rows.AddNew()
                        gvDetails.CurrentRow.Cells(colApply).Value = "Yes"
                        gvDetails.CurrentRow.Cells(colPINo).Value = objTr.PurchaseInvoice
                        gvDetails.CurrentRow.Cells(colDocNo).Value = objTr.Document_No
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = objTr.ConvRateOld
                        gvDetails.CurrentRow.Cells(colVendorInvNo).Value = objTr.Vendor_Invoice_No
                        gvDetails.CurrentRow.Cells(colAPTapalNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull (TapalNo,'') as TapalNo  from TSPL_VENDOR_INVOICE_HEAD where Document_No = '" + clsCommon.myCstr(objTr.Document_No) + "'"))
                        If clsCommon.CompairString(obj.Payment_Type, "SR") = CompairStringResult.Equal Then
                            gvDetails.CurrentRow.Cells(colDocType).Value = "Receipt"
                            'gvDetails.CurrentRow.Cells(colDocDate).Value = clsDBFuncationality.getSingleValue("SELECT Payment_Date as DocDate from TSPL_Payment_Header WHere Payment_No='" + objTr.Vendor_Invoice_No + "'")
                            gvDetails.CurrentRow.Cells(colDocDate).Value = objTr.PaymentDate
                            gvDetails.CurrentRow.Cells(colDocumentDate).Value = objTr.DocumentDate
                        Else
                            gvDetails.CurrentRow.Cells(colDocType).Value = clsDBFuncationality.getSingleValue("SELECT Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType]From TSPL_VENDOR_INVOICE_HEAD Where Vendor_Invoice_No='" + objTr.Vendor_Invoice_No + "' AND Document_No='" + objTr.Document_No + "'")
                            'gvDetails.CurrentRow.Cells(colDocDate).Value = clsDBFuncationality.getSingleValue("SELECT Invoice_Entry_Date as DocDate from TSPL_VENDOR_INVOICE_HEAD WHere Document_No='" + objTr.Document_No + "'")
                            If clsCommon.myLen(gvDetails.CurrentRow.Cells(colDocType).Value) <= 0 Then
                                gvDetails.CurrentRow.Cells(colDocType).Value = clsDBFuncationality.getSingleValue(" Select 'Receipt Note' from TSPL_PAYMENT_HEADER where Payment_No ='" & objTr.Document_No & "' and Payment_Type  ='RC'")
                                gvDetails.CurrentRow.Cells(colDocDate).Value = obj.Payment_Date
                                gvDetails.CurrentRow.Cells(colDocumentDate).Value = obj.Payment_Date
                            Else
                                gvDetails.CurrentRow.Cells(colDocDate).Value = objTr.PaymentDate
                                gvDetails.CurrentRow.Cells(colDocumentDate).Value = objTr.DocumentDate
                            End If

                        End If
                        gvDetails.CurrentRow.Cells(colOriginalAmt).Value = objTr.Original_Invoice_Amt
                        gvDetails.CurrentRow.Cells(colTDSAmt).Value = objTr.TDS_Amount
                        'changed by richa agarwal
                        ' gvDetails.CurrentRow.Cells(colNetAmt).Value = objTr.Pending_Balance
                        ' gvDetails.CurrentRow.Cells(colNetAmt).Value = objTr.Original_Invoice_Amt
                        gvDetails.CurrentRow.Cells(colNetAmt).Value = objTr.Original_Invoice_Amt - objTr.TDS_Amount
                        '----------
                        gvDetails.CurrentRow.Cells(colAppliedAmt).Value = objTr.Applied_Amount
                        'changed by richa agarwal
                        'gvDetails.CurrentRow.Cells(colPendingAmt).Value = objTr.Net_Balance
                        gvDetails.CurrentRow.Cells(colPendingAmt).Value = objTr.Pending_Balance
                        ''-------------------
                        gvDetails.CurrentRow.Cells(colSecurityAmt).Value = objTr.Security_Amount
                        'changed by richa agarwal
                        'gvDetails.CurrentRow.Cells(colTemp).Value = objTr.Applied_Amount + objTr.Net_Balance
                        gvDetails.CurrentRow.Cells(colTemp).Value = objTr.Applied_Amount + objTr.Pending_Balance
                        ''--------------------
                        gvDetails.CurrentRow.Cells(colComment).Value = objTr.Comment
                        '' Anubhooti 14-Nov-2014 BM00000004636
                        If clsCommon.myCdbl(objTr.Pending_Balance) > 0 Then
                            gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = clsCommon.myCdbl(objTr.Net_Balance) - clsCommon.myCdbl(objTr.Pending_Balance)
                        Else
                            gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = clsCommon.myCdbl(objTr.Net_Balance)
                        End If

                    Next

                ElseIf clsCommon.CompairString(obj.Payment_Type, "MI") = CompairStringResult.Equal Then
                    For Each objTr As clsPaymentDetail In obj.ArrTr
                        gvDetails.Rows.AddNew()
                        gvDetails.CurrentRow.Cells(colLineNo).Value = objTr.Payment_Line_No
                        gvDetails.CurrentRow.Cells(colGLAccount).Value = objTr.Account_Code
                        gvDetails.CurrentRow.Cells(colAccDesc).Value = objTr.Description
                        If clsCommon.myLen(gvDetails.Rows(gvDetails.Rows.Count - 1).Cells(colGLAccount).Value) > 0 Then
                            gvDetails.Rows(gvDetails.Rows.Count - 1).Cells(colGLType).Value = clsPaymentHeader.CheckGLAccountType(clsCommon.myCstr(gvDetails.Rows(gvDetails.Rows.Count - 1).Cells(colGLAccount).Value), Nothing)
                        Else
                            gvDetails.Rows(gvDetails.Rows.Count - 1).Cells(colGLType).Value = ""
                        End If
                        gvDetails.CurrentRow.Cells(colAmount).Value = objTr.Net_Balance
                        gvDetails.CurrentRow.Cells(colRemark).Value = objTr.Remarks
                        If clsCommon.myLen(lblProjCode.Text) > 0 Then
                            gvDetails.CurrentRow.Cells(colExpenseCode).Value = objTr.EXPENSE_CODE
                            gvDetails.Columns(colExpenseCode).IsVisible = True
                        Else
                            gvDetails.Columns(colExpenseCode).IsVisible = False
                        End If

                        gvDetails.CurrentRow.Cells(colHirerachyCenter).Value = objTr.Hirerachy_Level_Code
                        gvDetails.CurrentRow.Cells(colHirerachyName).Value = objTr.Hirerachy_Level_Name
                        gvDetails.CurrentRow.Cells(colCostCenter).Value = objTr.Cost_Center_Fin_Code
                        gvDetails.CurrentRow.Cells(colCostCenterName).Value = objTr.Cost_Center_Fin_Name
                    Next
                End If

                IsInsideLoadData = False
                isNewEntry = False
                btnsave.Text = "Update"
                'txtBankCode.Enabled = False
                ddlPaymentType.Enabled = False
                chkSaving.Enabled = False

                'If (clsCommon.CompairString(obj.Payment_Type, "PY") = CompairStringResult.Equal AndAlso UsLock1.Status = ERPTransactionStatus.Pending) Then
                'If (clsCommon.CompairString(obj.Payment_Type, "PY") = CompairStringResult.Equal AndAlso UsLock1.Status = ERPTransactionStatus.Approved) Then
                ' // comment by suhail 08/05/2024 
                'If (clsCommon.CompairString(obj.Payment_Type, "PY") = CompairStringResult.Equal AndAlso UsLock1.Status = ERPTransactionStatus.Pending) Then
                '    LoadVendorInvoices(obj.Vendor_Code, obj.Payment_Date, True)
                'End If

                If (clsCommon.CompairString(obj.Payment_Type, "SR") = CompairStringResult.Equal AndAlso UsLock1.Status = ERPTransactionStatus.Pending) Then
                    LoadPaymentEntries(obj.Vendor_Code, obj.Payment_Date, True)
                End If
                If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "MI") = CompairStringResult.Equal Then
                    gvDetails.Rows.AddNew()
                End If

                '' MULTICURRENCY
                Me.txtCurrencyCode.Value = obj.CURRENCY_CODE
                Me.txtConversionRate.Text = obj.ConvRate

                If obj.ApplicableFrom IsNot Nothing Then
                    Me.txtApplicableFrom.Text = clsCommon.GetPrintDate(obj.ApplicableFrom, "dd/MMM/yyyy")
                Else
                    Me.txtApplicableFrom.Text = ""
                End If

                ''richa agarwal 30/09/2015 against ticket no BM00000008007
                '  Me.txtTotalPaymentBaseCurr.Text = obj.PAYMENT_AMOUNT_BASE_CURRENCY
                Me.txtTotalPaymentBaseCurr.Text = (clsCommon.myCdbl(txtPaymentAmt.Text) - obj.TDS_Amount) * obj.ConvRate
                '' end  MULTICURRENCY
                ''---------------------- richa employee salary integration
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                        ddlEmployeeType.Enabled = True
                        ddlEmployeeAdvanceType.Enabled = False
                    Else
                        ddlEmployeeType.Enabled = False
                        ddlEmployeeAdvanceType.Enabled = False
                    End If
                ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                        ddlEmployeeAdvanceType.Enabled = True
                        ddlEmployeeType.Enabled = False
                    Else
                        ddlEmployeeAdvanceType.Enabled = False
                        ddlEmployeeType.Enabled = False
                    End If
                ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & txtVendorCode.Value & "' ")), "1") = CompairStringResult.Equal Then
                        ddlEmployeeAdvanceType.Enabled = True
                        ddlEmployeeType.Enabled = True
                    Else
                        ddlEmployeeAdvanceType.Enabled = False
                        ddlEmployeeType.Enabled = False
                    End If
                Else
                    ddlEmployeeType.Enabled = False
                    ddlEmployeeAdvanceType.Enabled = False
                End If
                ''---------------

                GSTStatus = clsERPFuncationality.GetGSTStatus(obj.Payment_Date)
                If GSTStatus Then
                    LoadBlankGridTax()
                    LoadBlankGridPOItemDetail()
                    LblPOTotalAmount.Text = obj.PurchaseOrder_Amount
                    lblPOTotalAdditionalCharge.Text = obj.PurchaseOrder_Add_Amount
                    lblPOTotalTaxAmt.Text = obj.Tax_Amount_Advance
                    txtTaxGroup.Value = obj.Tax_Group
                    txtPONo_GST.Value = obj.PurchaseOrder_No_GST
                    TxtPO_Location_GST.Value = obj.PO_Location_Code
                    LBLPO_Location_GST.Text = clsLocation.GetName(obj.PO_Location_Code, Nothing)

                    lblTaxGrpName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" Select Tax_Group_Desc from TSPL_TAX_GROUP_MASTER where Tax_Group_Code ='" & clsCommon.myCstr(txtTaxGroup.Value) & "'"))


                    For Each objTr As clsPaymentDetailGST In obj.ArrTrGST
                        gvItem.Rows.AddNew()
                        gvItem.CurrentRow.Cells(AdcolDocument_Code).Value = objTr.Document_Code
                        gvItem.CurrentRow.Cells(AdcolLine_No).Value = objTr.Line_No
                        gvItem.CurrentRow.Cells(AdcolRow_Type).Value = objTr.Row_Type
                        gvItem.CurrentRow.Cells(AdcolItem_Code).Value = objTr.Item_Code
                        gvItem.CurrentRow.Cells(adcolIName).Value = clsItemMaster.GetItemName(objTr.Item_Code, Nothing)
                        gvItem.CurrentRow.Cells(adcolIHSN).Value = clsItemMaster.GetItemHSNCode(objTr.Item_Code, Nothing)
                        gvItem.CurrentRow.Cells(AdcolQty).Value = objTr.Qty
                        gvItem.CurrentRow.Cells(AdcolBalance_Qty).Value = objTr.Balance_Qty
                        gvItem.CurrentRow.Cells(AdcolItem_Cost).Value = objTr.Item_Cost
                        gvItem.CurrentRow.Cells(AdcolUnit_code).Value = objTr.Unit_code
                        gvItem.CurrentRow.Cells(AdcolTAX1).Value = objTr.TAX1
                        gvItem.CurrentRow.Cells(AdcolTAX1_Amt).Value = objTr.TAX1_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX1_Base_Amt).Value = objTr.TAX1_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX1_Rate).Value = objTr.TAX1_Rate
                        gvItem.CurrentRow.Cells(Adcoltax2).Value = objTr.tax2
                        gvItem.CurrentRow.Cells(AdcolTAX2_Base_Amt).Value = objTr.TAX2_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX2_Rate).Value = objTr.TAX2_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX2_Amt).Value = objTr.TAX2_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX3).Value = objTr.TAX3
                        gvItem.CurrentRow.Cells(AdcolTAX3_Base_Amt).Value = objTr.TAX3_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX3_Rate).Value = objTr.TAX3_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX3_Amt).Value = objTr.TAX3_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX4).Value = objTr.TAX4
                        gvItem.CurrentRow.Cells(AdcolTAX4_Base_Amt).Value = objTr.TAX4_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX4_Rate).Value = objTr.TAX4_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX4_Amt).Value = objTr.TAX4_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX5).Value = objTr.tax5
                        gvItem.CurrentRow.Cells(AdcolTAX5_Base_Amt).Value = objTr.TAX5_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX5_Rate).Value = objTr.TAX5_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX5_Amt).Value = objTr.TAX5_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX6_Base_Amt).Value = objTr.TAX6_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX6).Value = objTr.tax6
                        gvItem.CurrentRow.Cells(AdcolTAX6_Rate).Value = objTr.TAX6_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX6_Amt).Value = objTr.TAX6_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX7).Value = objTr.tax7
                        gvItem.CurrentRow.Cells(AdcolTAX7_Base_Amt).Value = objTr.TAX7_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX7_Rate).Value = objTr.TAX7_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX7_Amt).Value = objTr.TAX7_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX8).Value = objTr.tax8
                        gvItem.CurrentRow.Cells(AdcolTAX8_Base_Amt).Value = objTr.TAX8_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX8_Rate).Value = objTr.TAX8_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX8_Amt).Value = objTr.TAX8_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX9).Value = objTr.tax9
                        gvItem.CurrentRow.Cells(AdcolTAX9_Base_Amt).Value = objTr.TAX9_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX9_Amt).Value = objTr.TAX9_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX9_Rate).Value = objTr.TAX9_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX10).Value = objTr.tax10
                        gvItem.CurrentRow.Cells(AdcolTAX10_Base_Amt).Value = objTr.TAX10_Base_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX10_Rate).Value = objTr.TAX10_Rate
                        gvItem.CurrentRow.Cells(AdcolTAX10_Amt).Value = objTr.TAX10_Amt
                        gvItem.CurrentRow.Cells(AdcolAmount).Value = objTr.Amount
                        gvItem.CurrentRow.Cells(AdcolDisc_Per).Value = objTr.Disc_Per
                        gvItem.CurrentRow.Cells(AdcolDisc_Amt).Value = objTr.Disc_Amt
                        gvItem.CurrentRow.Cells(AdcolAmt_Less_Discount).Value = objTr.Amt_Less_Discount
                        gvItem.CurrentRow.Cells(AdcolTotal_Tax_Amt).Value = objTr.Total_Tax_Amt
                        gvItem.CurrentRow.Cells(AdcolItem_Net_Amt).Value = objTr.Item_Net_Amt
                        gvItem.CurrentRow.Cells(AdcolTAX1_Amt_Payment).Value = objTr.TAX1_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX2_Amt_Payment).Value = objTr.TAX2_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX3_Amt_Payment).Value = objTr.TAX3_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX4_Amt_Payment).Value = objTr.TAX4_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX5_Amt_Payment).Value = objTr.TAX5_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX6_Amt_Payment).Value = objTr.TAX6_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX7_Amt_Payment).Value = objTr.TAX7_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX8_Amt_Payment).Value = objTr.TAX8_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX9_Amt_Payment).Value = objTr.TAX9_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolTAX10_Amt_Payment).Value = objTr.TAX10_Amt_Payment
                        gvItem.CurrentRow.Cells(AdcolPaymentAdvance).Value = objTr.PaymentAdvance
                        gvItem.CurrentRow.Cells(AdcolPaymentTotalTax).Value = objTr.PaymentTotalTax
                        gvItem.CurrentRow.Cells(AdcolPaymentTotalAdvanceAmt).Value = objTr.PaymentTotalAdvanceAmt


                        DblPaymentTaxAmt1 = DblPaymentTaxAmt1 + clsCommon.myCdbl(objTr.TAX1_Amt_Payment)
                        DblPaymentTaxAmt2 = DblPaymentTaxAmt2 + clsCommon.myCdbl(objTr.TAX2_Amt_Payment)
                        DblPaymentTaxAm3 = DblPaymentTaxAm3 + clsCommon.myCdbl(objTr.TAX3_Amt_Payment)
                        DblPaymentTaxAmt4 = DblPaymentTaxAmt4 + clsCommon.myCdbl(objTr.TAX4_Amt_Payment)
                        DblPaymentTaxAmt5 = DblPaymentTaxAmt5 + clsCommon.myCdbl(objTr.TAX5_Amt_Payment)
                        DblPaymentTaxAmt6 = DblPaymentTaxAmt6 + clsCommon.myCdbl(objTr.TAX6_Amt_Payment)
                        DblPaymentTaxAmt7 = DblPaymentTaxAmt7 + clsCommon.myCdbl(objTr.TAX7_Amt_Payment)
                        DblPaymentTaxAmt8 = DblPaymentTaxAmt8 + clsCommon.myCdbl(objTr.TAX8_Amt_Payment)
                        DblPaymentTaxAmt9 = DblPaymentTaxAmt9 + clsCommon.myCdbl(objTr.TAX9_Amt_Payment)
                        DblPaymentTaxAmt10 = DblPaymentTaxAmt10 + clsCommon.myCdbl(objTr.TAX10_Amt_Payment)

                    Next

                    For Each objTr As clsPaymentDetailGST In obj.ArrTrGST
                        If (clsCommon.myLen(clsCommon.myCstr(objTr.TAX1)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.TAX1)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.TAX1))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX1_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt1)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt1 * 100) / objTr.TAX1_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax2)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax2)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax2))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX2_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt2)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt2 * 100) / objTr.TAX2_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.TAX3)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.TAX3)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.TAX3))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX3_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAm3)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAm3 * 100) / objTr.TAX3_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.TAX4)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.TAX4)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.TAX4))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX4_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt4)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt4 * 100) / objTr.TAX4_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax5)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax5)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax5))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX5_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt5)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt5 * 100) / objTr.TAX5_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax6)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax6)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax6))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX6_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt6)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt6 * 100) / objTr.TAX6_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax7)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax7)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax7))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX7_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt7)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt7 * 100) / objTr.TAX7_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax8)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax8)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax8))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX8_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt8)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt8 * 100) / objTr.TAX8_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax9)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax9)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax9))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX9_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt9)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt9 * 100) / objTr.TAX9_Rate)
                        End If

                        If (clsCommon.myLen(clsCommon.myCstr(objTr.tax10)) > 0) Then
                            gvTaxDetail.Rows.AddNew()
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutCode).Value = clsCommon.myCstr(objTr.tax10)
                            gvTaxDetail.CurrentRow.Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(objTr.tax10))
                            gvTaxDetail.CurrentRow.Cells(colTTaxRate).Value = clsCommon.myCdbl(objTr.TAX10_Rate)
                            'gvTaxDetail.CurrentRow.Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblPaymentTaxAmt10)
                            'gvTaxDetail.CurrentRow.Cells(colTBaseAmt).Value = clsCommon.myCdbl((DblPaymentTaxAmt10 * 100) / objTr.TAX10_Rate)
                        End If
                        Exit For
                    Next
                    SetitemWiseTaxSetting(False, False)
                    CalculateTaxDetailForTaxgrid()
                End If

                ''--------------------- end of gst work
                '' Panch Raj :Tax on Bank Charges
                gv2.Rows.Clear()
                txtTaxGroupBankCharges.Value = obj.Tax_Group_BankCharges
                For Each objtr As clsPaymentBankChargesTax In obj.objBCT
                    gv2.Rows.AddNew()
                    'gv2.Rows(gv2.Rows.Count - 1).Cells(colBCLine_No).Value = objtr.Line_No
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colBCaxAutCode).Value = objtr.Tax_Code
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colBCaxAutName).Value = clsTaxMaster.GetName(objtr.Tax_Code)

                    gv2.Rows(gv2.Rows.Count - 1).Cells(colBCaxRate).Value = objtr.Tax_Rate
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colBCBaseAmt).Value = objtr.Tax_Base_Amount
                    gv2.Rows(gv2.Rows.Count - 1).Cells(colBCaxAmt).Value = objtr.Tax_Amount
                Next
                Dim dtProductionCode As DataTable = clsDBFuncationality.GetDataTable(" select PI_No from TSPL_PAYMENT_MULTIPLE_INVOICE_DETAIL  where Payment_No='" + obj.Payment_No + "'")

                Dim Arr_Prod As New ArrayList
                For Each dr As DataRow In dtProductionCode.Rows
                    Arr_Prod.Add(dr("PI_No"))
                Next
                'If obj.Arr_Prod IsNot Nothing AndAlso obj.Arr_Prod.Count > 0 Then
                '    For Each objPRD As clsProductionEntry In obj.Arr_Prod
                '        For i As Integer = 0 To txtprodCode.arrValueMember.Count - 1
                '            ' Assuming Arr_Prod is a List(Of String)
                '            Arr_Prod.Add(txtprodCode.arrValueMember(i))
                '        Next
                '    Next
                txtmulPI.arrValueMember = Arr_Prod

                '=====================if document go for approval then no post button visible or if document contain related setting
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
                    btnpost.Visible = MyBase.isPostFlag
                    If Not clsApply_Approval.Visibility_PostButtonForApproval(MyBase.Form_ID, clsCommon.myCstr(txtPaymentNo.Value), clsCommon.myCdbl(obj.Payment_Amount), 0, "") Then
                        btnpost.Visible = False
                        If UsLock1.Status = ERPTransactionStatus.Pending Then
                            UsLock1.Status = clsApply_Approval.ApprovalCondCheck_Doc(MyBase.Form_ID, clsCommon.myCstr(txtPaymentNo.Value), Nothing)
                        End If
                    End If
                End If
                '============================================

                MyBase.ReStoreGridLayoutMain(gvItem)

                ''For Custom Fields
                If MyBase.customFieldTabProperty = ElementVisibility.Visible Then
                    UcCustomFields1.LoadData(obj.Payment_No)
                End If
                clsCustomFieldGrid.FillDataInGrid(obj.Payment_No, MyBase.Form_ID, gvDetails)
                ''End of For Custom Fields
                UcAttachment1.LoadData(obj.Payment_No)

                End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

#Region "TDS Detail"
    Private Sub btnViewTDSDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewTDSDetails.Click
        viewtds()
        'If Not chkTDSProvision.Checked Then
        If objRemittance IsNot Nothing Then
            txtTDSAmt.Text = clsCommon.myCstr(Math.Round(objRemittance.Actual_Total_TDS, 0, MidpointRounding.AwayFromZero))
        Else
            txtTDSAmt.Text = Nothing
        End If
        'End If

    End Sub

    Public Sub viewtds()
        Try
            Dim frm As New FrmViewTDS()
            UpdateTDSAmount()
            frm.ObjIn = objRemittance
            frm.ShowDialog()
            'If (frm.ObjReturn IsNot Nothing) Then
            objRemittance = frm.ObjReturn
            'End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub UpdateTDSAmount()
        Dim tdsamt As Decimal = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select TDS_Amount  from TSPL_PAYMENT_HEADER where Payment_No ='" + txtPaymentNo.Value + "'"))
        Dim dblTotAmt As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select Payment_Amount  from TSPL_PAYMENT_HEADER where Payment_No ='" + txtPaymentNo.Value + "'"))
        'Dim dbTotAmtNew As Double = clsCommon.myCdbl(txtPaymentAmt.Text)
        'If dblTotAmt = dbTotAmtNew Then
        '    dblTotAmt = clsCommon.myCdbl(txtPaymentAmt.Text) + tdsamt
        'Else
        '    dblTotAmt = clsCommon.myCdbl(txtPaymentAmt.Text)
        'End If
        ''richa agarwal 30/09/2015 against ticket no BM00000008007
        '  Dim dbTotAmtNew As Double = clsCommon.myCdbl(txtNetPayableAmt.Text)
        Dim dbTotAmtNew As Double = clsCommon.myCdbl(txtPaymentAmt.Text)
        'If dblTotAmt = dbTotAmtNew Then
        '    dblTotAmt = clsCommon.myCdbl(txtNetPayableAmt.Text) + tdsamt
        'Else
        '    dblTotAmt = clsCommon.myCdbl(txtNetPayableAmt.Text)
        'End If

        dblTotAmt = clsCommon.myCdbl(txtPaymentAmt.Text)
        If (objRemittance Is Nothing) Then
            SetVendorTDSDetails()
        Else
            Dim objDedDetails As clsTDSDeductionDetails = clsTDSDeductionDetails.GetApplicableTDRate(objRemittance.Deduction_Code, dblTotAmt, Nothing, False, txtVendorCode.Value)
            If (objDedDetails IsNot Nothing) Then
                objRemittance.TDS_Per = objDedDetails.TDS
                objRemittance.Surcharge_Per = objDedDetails.Surcharge
                objRemittance.Edu_Cess_Per = objDedDetails.Educess
                objRemittance.Sec_Educess_Per = objDedDetails.Seceducess
            End If
        End If

        If (objRemittance IsNot Nothing) Then
            objRemittance.Vendor_Code = txtVendorCode.Value
            objRemittance.Vendor_Name = lblVendorName.Text
            objRemittance.Document_Date = dtpPayment.Value
            objRemittance.Document_Type = clsCommon.myCstr(ddlPaymentType.SelectedValue)
            objRemittance.Document_Amount = dblTotAmt
            objRemittance.Calculated_TDS_Base = dblTotAmt
            If Not objRemittance.IsTDSOverride Then
                objRemittance.Actual_TDS_Base = dblTotAmt
            End If

            objRemittance.Calculated_TDS = (objRemittance.Calculated_TDS_Base * objRemittance.TDS_Per) / 100
            objRemittance.Actual_TDS = (objRemittance.Actual_TDS_Base * objRemittance.TDS_Per) / 100

            objRemittance.Calculated_Surcharge = (objRemittance.Calculated_TDS_Base * objRemittance.Surcharge_Per) / 100
            objRemittance.Actual_Surcharge = (objRemittance.Actual_TDS_Base * objRemittance.Surcharge_Per) / 100

            objRemittance.Calculated_Edu_Cess = (objRemittance.Calculated_TDS_Base * objRemittance.Edu_Cess_Per) / 100
            objRemittance.Actual_Edu_Cess = (objRemittance.Actual_TDS_Base * objRemittance.Edu_Cess_Per) / 100

            objRemittance.Calculated_Sec_Educess = (objRemittance.Calculated_TDS_Base * objRemittance.Sec_Educess_Per) / 100
            objRemittance.Actual_Sec_Educess = (objRemittance.Actual_TDS_Base * objRemittance.Sec_Educess_Per) / 100

            objRemittance.Calculated_Total_TDS = objRemittance.Calculated_TDS + objRemittance.Calculated_Surcharge + objRemittance.Calculated_Edu_Cess + objRemittance.Calculated_Sec_Educess
            objRemittance.Actual_Total_TDS = objRemittance.Actual_TDS + objRemittance.Actual_Surcharge + objRemittance.Actual_Edu_Cess + objRemittance.Actual_Sec_Educess
        End If
    End Sub

    Private Sub txtPaymentAmt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPaymentAmt.TextChanged
        If txtmulPI.arrValueMember IsNot Nothing AndAlso txtmulPI.arrValueMember.Count <= 0 Then
            If clsCommon.myCdbl(txtPaymentAmt.Text) > 0 Then
                If Not clsCommon.CompairString(ddlPaymentType.SelectedValue, "MI") = CompairStringResult.Equal And clsCommon.myLen(txtVendorCode.Value) <= 0 Then
                    txtPaymentAmt.Text = "0"
                    RadMessageBox.Show("Please Select Vendor")
                    txtVendorCode.Focus()
                    Exit Sub
                End If
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    SetVendorTDSDetails()
                    If objRemittance IsNot Nothing Then
                        UpdateTDSAmount()
                    End If
                    txtTDSAmt.Text = Math.Round(objRemittance.Actual_Total_TDS, 0, MidpointRounding.AwayFromZero)

                    If Not chkTDSProvision.Checked Then
                        txtNetPayableAmt.Text = clsCommon.myCstr(clsCommon.myCdbl(txtPaymentAmt.Text) - clsCommon.myCdbl(txtTDSAmt.Text))
                    Else
                        txtNetPayableAmt.Text = clsCommon.myCstr(clsCommon.myCdbl(txtPaymentAmt.Text))

                    End If

                End If
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
                    GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
                    If GSTStatus AndAlso clsCommon.myLen(txtPONo_GST.Value) > 0 Then
                        CalculateTaxAmountInAdvnce()
                    End If
                End If
                ''richa BHA/28/08/18-000491
                BankChargerasperSlabWise()
            Else
                txtTDSAmt.Text = "0"
                txtNetPayableAmt.Text = "0"
            End If
            ' Me.txtTotalPaymentBaseCurr.Value = (clsCommon.myCdbl(txtNetPayableAmt.Text) * clsCommon.myCdbl(txtConversionRate.Value))
            If Not chkTDSProvision.Checked Then
                Me.txtTotalPaymentBaseCurr.Value = (clsCommon.myCdbl(txtPaymentAmt.Text) - clsCommon.myCdbl(txtTDSAmt.Text)) * clsCommon.myCdbl(txtConversionRate.Value)
            Else
                Me.txtTotalPaymentBaseCurr.Value = (clsCommon.myCdbl(txtPaymentAmt.Text) - clsCommon.myCdbl(txtTDSAmt.Text)) * clsCommon.myCdbl(txtConversionRate.Value)

            End If
        End If
    End Sub
    ''richa BHA/28/08/18-000491
    Sub BankChargerasperSlabWise()

        If ApplyBankChargesasperSlabonBankMaster = True Then
            If clsCommon.myLen(txtBankCode.Value) > 0 Then
                txtBankCharges.Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select isnull(charges,0) as charges from TSPL_BANK_MASTER_SLAB where fromSlab<=" & clsCommon.myCdbl(txtPaymentAmt.Value) & " and ToSlab >=" & clsCommon.myCdbl(txtPaymentAmt.Value) & " and BANK_CODE ='" & txtBankCode.Value & "'"))
            End If
            If clsCommon.myCdbl(txtBankCharges.Text) <= 0 Then
                txtTaxGroupBankCharges.Value = ""
                txtTaxGroup_TxtChanged()
            Else
                Dim Qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER WHERE " &
           "(  select count(TSPL_TAX_GROUP_DETAILS.Tax_Code)  from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_MASTER on " &
           " TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code=TSPL_TAX_GROUP_MASTER.Tax_Group_Code " &
           "  )=(  select count(TSPL_TAX_GROUP_DETAILS.Tax_Code)  from TSPL_TAX_GROUP_DETAILS left outer join " &
         " TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where " &
           " TSPL_TAX_GROUP_DETAILS.Tax_Group_Code=TSPL_TAX_GROUP_MASTER.Tax_Group_Code   ) and Tax_Group_Type='P' " &
          " and 2=(case when (select Description from TSPL_FIXED_PARAMETER where Type='Show only Active Taxes/Rates/Groups for GST' and Code='Show only Active Taxes/Rates/Groups for GST')=1 then case when TSPL_TAX_GROUP_MASTER.Active=1 then 2 else 1 end else 2 end)  AND TSPL_TAX_GROUP_MASTER.Is_Tax_Exempted =1"
                txtTaxGroupBankCharges.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue(Qry))
                txtTaxGroup_TxtChanged()
            End If
        End If
    End Sub

    Private Sub txtTDSAmt_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTDSAmt.TextChanged
        'txtNetPayableAmt.Text = clsCommon.myCstr(clsCommon.myCdbl(txtPaymentAmt.Text) - clsCommon.myCdbl(txtTDSAmt.Text))
    End Sub
#End Region

    Private Sub btnpost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpost.Click
        PostData()
    End Sub

    Private Function PostData()
        Try
            iiDeadlockErrors = 1
            If myMessages.postConfirm AndAlso PostData1() Then
                If common.clsCommon.MyMessageBoxShow("Record Posted Successfully. Do You Want To Print Payment Voucher ?", "Payment", MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    PrintData()
                End If
                btndelete.Enabled = False
                btnpost.Enabled = False
                btnsave.Enabled = False
                btnViewTDSDetails.Enabled = True
                UsLock1.Status = ERPTransactionStatus.Approved
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
        Return True
    End Function

    Dim iiDeadlockErrors As Integer
    Private Function PostData1() As Boolean
        Try
            '' Anubhooti 27-Nov-2014 BM00000004642
            If objCommonVar.RCDFCFP = True Then
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    Dim objVendor As clsTDSVendorDetails = clsTDSVendorDetails.GetData(txtVendorCode.Value)
                    If objVendor IsNot Nothing AndAlso clsCommon.myCdbl(txtTDSAmt.Text) > 0 Then
                        If (common.clsCommon.MyMessageBoxShow("TDS " & clsCommon.myCstr(txtTDSAmt.Text) & " rupees has been deducted.Do you want to deduct it?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                        Else
                            txtTDSAmt.Text = Nothing
                        End If
                    End If
                End If
            Else
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Then
                    Dim objVendor As clsTDSVendorDetails = clsTDSVendorDetails.GetData(txtVendorCode.Value)
                    If objVendor IsNot Nothing AndAlso clsCommon.myCdbl(txtTDSAmt.Text) > 0 Then
                        If (common.clsCommon.MyMessageBoxShow("TDS " & clsCommon.myCstr(txtTDSAmt.Text) & " rupees has been deducted.Do you want to change it?", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes) Then
                            viewtds()
                            If objRemittance IsNot Nothing Then
                                txtTDSAmt.Text = clsCommon.myCstr(Math.Round(objRemittance.Actual_Total_TDS, 0, MidpointRounding.AwayFromZero))
                            Else
                                txtTDSAmt.Text = Nothing
                            End If
                        End If
                    Else
                    End If
                    ' SetVendorTDSDetails()
                End If
            End If

            ''
            ''richa agarwal 09/06/2016
            isFlag = True
            ''----------------
            If SaveData(True) = False Then
                Return False
            End If

            If clsPaymentHeader.PostData(txtPaymentNo.Value, Me.Module_Code) Then
                Return True
            End If
        Catch ex As Exception
            If ex.Message.Contains("deadlocked") Then
                iiDeadlockErrors += 1
                If iiDeadlockErrors >= 15 Then
                    Me.Close()
                    Exit Function
                End If
                System.Threading.Thread.Sleep(3000)
                PostData1()
            Else
                Throw New Exception(ex.Message)
            End If
        Finally
            isFlag = False
        End Try
    End Function

    Private Sub btndelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btndelete.Click
        deletedata()
    End Sub
    Public Sub deletedata()
        Try
            Dim Reason As String = ""
            If myMessages.deleteConfirm Then
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
                If CheckNegativeBankBalanceondelete() Then
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") <> CompairStringResult.Equal Then
                        clsApply_Approval.CheckUpdate_Doc_Valid(MyBase.Form_ID, clsCommon.myCstr(txtPaymentNo.Value))
                    End If
                    fundelete()
                    saveCancelLog(Reason, "Delete", Nothing)
                    Reset()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
    Function saveCancelLog(ByVal Reason As String, ByVal Activity_Type As String, Optional ByVal trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim obj As New clsCancelLog
        obj.Program_Code = Form_ID
        obj.DOCUMENT_NO = clsCommon.myCstr(Me.txtPaymentNo.Value)
        obj.REASON = Reason
        obj.ACTIVITY_TYPE = Activity_Type
        Return clsCancelLog.SaveData(obj, True, trans)
    End Function

    Private Sub fundelete()
        Try
            If clsPaymentHeader.fundelete(ddlPaymentType.SelectedValue, txtPaymentNo.Value, txtVendorCode.Value) Then
                myMessages.delete()
            End If
        Catch ex As Exception
            myMessages.myExceptions(ex)
        End Try
    End Sub

    Private Sub FrmPaymentNew_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Alt AndAlso e.KeyCode = Keys.N Then
            Reset()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.S AndAlso MyBase.isModifyFlag AndAlso btnsave.Enabled Then
            If AllowToSave1() Then
                SaveData1()
            End If
        ElseIf e.Alt AndAlso e.KeyCode = Keys.P AndAlso MyBase.isPostFlag AndAlso btnpost.Enabled Then
            PostData()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.D AndAlso MyBase.isDeleteFlag AndAlso btndelete.Enabled Then
            deletedata()
        ElseIf e.Alt AndAlso e.KeyCode = Keys.C AndAlso btnclose.Enabled Then
            Me.Close()
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F12 Then
            If MyBase.isReverse Then

                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SIR
                frm.strCode = clsFixedParameterCode.SIReversAndCreate
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    btnReverse.Visible = True
                End If

                ButtonToolTip.SetToolTip(btnsave, "Press Alt+S for Save/Update Trasnaction" + Environment.NewLine +
                               "TSPL_PAYMENT_HEADER " + Environment.NewLine +
                               "TSPL_PAYMENT_DETAIL " + Environment.NewLine +
                               "TSPL_PAYMENT_DETAIL_GST  " + Environment.NewLine +
                               "TSPL_PAYMENT_BANK_CHARGES_TAX " + Environment.NewLine +
                               "TSPL_REMITTANCE (For Remmitance Entry)" + Environment.NewLine +
                               "TSPL_CUSTOM_FIELD_VALUES " + Environment.NewLine +
                               "TSPL_JOURNAL_MASTER ( For Journal Entry)" + Environment.NewLine +
                               "TSPL_VENDOR_INVOICE_HEAD ( update during Journal Entry) " + Environment.NewLine +
                               "TSPL_Bulk_MILK_PURCHASE_INVOICE_HEAD  ( update during Journal Entry)" + Environment.NewLine +
                               "TSPL_MILK_PURCHASE_INVOICE_HEAD ( update during Journal Entry) " + Environment.NewLine +
                               "TSPL_PI_HEAD ( update during Journal Entry) " + Environment.NewLine +
                               "TSPL_VENDOR_INVOICE_HEAD ( update during Journal Entry) " + Environment.NewLine +
                               "TSPL_PJC_EXPENSE_HEADER ")
            Else
                clsCommon.MyMessageBoxShow(Me, "You are not authorized to perform this action.", Me.Text, MessageBoxButtons.OK, Telerik.WinControls.RadMessageIcon.Error)
                'MessageBox.Show("You are not authorized to perform this action.", "Unauthorized Access", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.F10 Then
            If Not isSettlementBankOnly Then
                Dim frm As New FrmPWD(Nothing)
                frm.strType = clsFixedParameterType.SettlementBankOnlyPWD
                frm.strCode = clsFixedParameterCode.SettlementBankOnlyPWD
                frm.ShowDialog()
                If frm.isPasswordCorrect Then
                    isSettlementBankOnly = True
                End If
            Else
                isSettlementBankOnly = False
            End If
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.I Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.StoreADJExportImportAfterPost
            frm.strCode = clsFixedParameterCode.StoreADJExportImportAfterPost
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                rbtnImportPosted.Visible = True
            End If
        ElseIf e.Alt AndAlso e.Shift AndAlso e.Control And e.KeyCode = Keys.E Then
            Dim frm As New FrmPWD(Nothing)
            frm.strType = clsFixedParameterType.StoreADJExportImportAfterPost
            frm.strCode = clsFixedParameterCode.StoreADJExportImportAfterPost
            frm.ShowDialog()
            If frm.isPasswordCorrect Then
                rbtnExportPosted.Visible = True
            End If

        End If
    End Sub

    Private Sub BtnBank_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBank.Click
        Try
            Dim frm As New FrmBankTransfer(objCommonVar.CurrentUserCode, objCommonVar.CurrentCompanyCode)
            frm.SetUserMgmt(clsUserMgtCode.bankTransfer)
            frm.ShowDialog()
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnprint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnprint.Click
        PrintData()
    End Sub

    Private Sub PrintData()
        If clsCommon.myLen(txtPaymentNo.Value) <= 0 Then
            common.clsCommon.MyMessageBoxShow("Please select Document no to Print", Me.Text)
            Return
        End If
        Dim DocDate As Date?
        DocDate = Nothing
        DocDate = dtpPayment.Value
        If clsERPFuncationality.GetGSTStatus(DocDate) = True AndAlso clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "AV") = CompairStringResult.Equal AndAlso clsCommon.myLen(clsCommon.myCstr(txtPONo_GST.Value)) > 0 Then

            clsPaymentHeader.funPaymentEntyPrint(MyBase.Form_ID, False, DocDate, txtPaymentNo.Value)

            '    Dim StrQuery As String = Nothing
            '    StrQuery = "select TSPL_VENDOR_MASTER.State_Code as Vendor_StateCode,TSPL_LOCATION_MASTER.State AS Location_State, TSPL_PAYMENT_HEADER.PurchaseOrder_No_GST, '' as EwayBillNo,'' as Electronic_Ref_No,'' AS Vechil_No,null as SupllyDate,TSPL_PAYMENT_HEADER.Payment_No,convert(varchar,TSPL_PAYMENT_HEADER.Payment_Date,103) as Payment_Date,TSPL_PAYMENT_HEADER.Payment_Amount + isnull(TSPL_PAYMENT_HEADER.TDS_Amount ,0) as Payment_Amount, TSPL_PAYMENT_HEADER.Vendor_Name,TSPL_COMPANY_MASTER.Comp_Name,Comp_State_Master.State_Name as Comp_State," &
            '    " COMP_Address=TSPL_COMPANY_MASTER.Add1 + CASE WHEN ISNULL(TSPL_COMPANY_MASTER.Add2,'')<>'' THEN ','+TSPL_COMPANY_MASTER.Add2 WHEN ISNULL(TSPL_COMPANY_MASTER.Add3,'')<>'' THEN ','+TSPL_COMPANY_MASTER.Add3 END ," &
            '    " TSPL_COMPANY_MASTER.Pan_No as Comp_PanNo,TSPL_COMPANY_MASTER.GSTReg_No  as Comp_GSTIN_NO ,Comp_State_Master.GST_STATE_Code as Comp_GST_StateCode," &
            '    " Loc_Address=TSPL_LOCATION_MASTER.Add1 + CASE WHEN ISNULL(TSPL_LOCATION_MASTER.Add2,'')<>'' THEN ','+TSPL_LOCATION_MASTER.Add2 WHEN ISNULL(TSPL_LOCATION_MASTER.Add3,'')<>'' THEN ','+TSPL_LOCATION_MASTER.Add3  END ," &
            '    " Receiver_Add=Ship_to_Location_master.Add1 + CASE WHEN ISNULL(Ship_to_Location_master.Add2,'')<>'' THEN ','+Ship_to_Location_master.Add2 WHEN ISNULL(Ship_to_Location_master.Add3,'')<>'' THEN ','+Ship_to_Location_master.Add3  END, " &
            '    " ISNULL(TSPL_PURCHASE_ORDER_HEAD.Ship_To_Location,'') AS Ship_To_Location,	" &
            '    " case when isnull(Ship_to_Location_master.Location_Desc,'')<>'' then Ship_to_Location_master.GSTNO else TSPL_LOCATION_MASTER.GSTNO  end as Loc_GSTIN_NO ," &
            '    " case when isnull(Ship_to_Location_master.Location_Desc,'')<>'' then Ship_to_State_Master.STATE_NAME else Location_State_Master.STATE_NAME end as Loc_State," &
            '    " case when isnull(Ship_to_Location_master.Location_Desc,'')<>'' then Ship_to_State_Master.GST_STATE_Code else Location_State_Master.GST_STATE_Code end as Loc_GST_StateCode," &
            '    " TSPL_VENDOR_MASTER.Vendor_Name as  Supplier_Name,Vendor_State_Master.STATE_NAME as Supl_State,Vendor_State_Master.GST_STATE_Code as Sup_GST_State , " &
            '    " Supplier_Add=TSPL_VENDOR_MASTER.Add1 + CASE WHEN ISNULL(TSPL_VENDOR_MASTER.Add2,'')<>'' THEN ','+TSPL_VENDOR_MASTER.Add2 WHEN ISNULL(TSPL_VENDOR_MASTER.Add3,'')<>'' THEN ','+TSPL_VENDOR_MASTER.Add3  END," &
            '    " (case when isnull(Ship_to_Location_master.Location_Desc,'')<>'' then Ship_to_Location_master.Location_Desc else TSPL_LOCATION_MASTER.Location_Desc end )as Place_of_Supply," &
            '    " TSPL_VENDOR_MASTER.GSTFinalNo as Supl_GSTIN_NO,TSPL_PURCHASE_ORDER_HEAD.Mode_Of_Transport,TSPL_LOCATION_MASTER.Location_Desc," &
            '    " TSPL_ITEM_MASTER.Item_Desc,TSPL_ITEM_MASTER.HSN_Code,TSPL_PAYMENT_DETAIL_GST.Qty,TSPL_PAYMENT_DETAIL_GST.Unit_code,TSPL_PAYMENT_DETAIL_GST.Item_Cost," &
            '    " TSPL_PAYMENT_DETAIL_GST.Amount,TSPL_PAYMENT_DETAIL_GST.PaymentAdvance,TSPL_PAYMENT_DETAIL_GST.PaymentTotalTax,TSPL_PAYMENT_DETAIL_GST.PaymentTotalAdvanceAmt, " &
            '    " DTAX1.Type AS TAX1,DTAX2.TYPE AS TAX2,DTAX3.TYPE AS TAX3,DTAX4.Type AS TAX4,DTAX5.TYPE AS TAX5,DTAX6.TYPE AS TAX6,DTAX7.TYPE AS TAX7,DTAX8.TYPE AS TAX8,DTAX9.TYPE AS TAX9,DTAX10.Type AS TAX10," &
            '    " isnull(TSPL_PAYMENT_DETAIL_GST.TAX1_Amt_Payment,0) as TAX1_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX1_Rate,0) as TAX1_Rate,  isnull(TSPL_PAYMENT_DETAIL_GST.TAX2_Amt_Payment,0) as TAX2_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX2_Rate,0) as TAX2_Rate,  isnull(TSPL_PAYMENT_DETAIL_GST.TAX3_Amt_Payment,0) as TAX3_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX3_Rate,0) as TAX3_Rate,  isnull(TSPL_PAYMENT_DETAIL_GST.TAX4_Amt_Payment,0) as TAX4_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX4_Rate,0) as TAX4_Rate,isnull(TSPL_PAYMENT_DETAIL_GST.TAX5_Amt_Payment,0) as TAX5_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX5_Rate,0) as TAX5_Rate,  isnull(TSPL_PAYMENT_DETAIL_GST.TAX6_Amt_Payment,0) as TAX6_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX6_Rate,0) as TAX6_Rate,  isnull(TSPL_PAYMENT_DETAIL_GST.TAX7_Amt_Payment,0) as TAX7_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX7_Rate,0) as TAX7_Rate, isnull(TSPL_PAYMENT_DETAIL_GST.TAX8_Amt_Payment,0) as TAX8_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX8_Rate,0) as TAX8_Rate , isnull(TSPL_PAYMENT_DETAIL_GST.TAX9_Amt_Payment,0) as TAX9_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX9_Rate,0) as TAX9_Rate,isnull(TSPL_PAYMENT_DETAIL_GST.TAX10_Amt_Payment,0) as TAX10_Amt, isnull(TSPL_PAYMENT_DETAIL_GST.TAX10_Rate,0) as TAX10_Rate ," &
            '    " DTAX1.TYPE AS Tax1Type,DTAX2.TYPE AS Tax2Type,DTAX3.TYPE AS Tax3Type,DTAX4.TYPE AS Tax4Type,DTAX5.TYPE AS Tax5Type,DTAX6.TYPE AS Tax6Type,DTAX7.TYPE AS Tax7Type,DTAX8.TYPE AS Tax8Type,DTAX9.TYPE AS Tax9Type,DTAX10.TYPE AS Tax10Type," &
            '    " isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt1,0) as Add_Charge_Amt1 ,isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt2,0) as Add_Charge_Amt2 ," &
            '    " isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt3,0) as Add_Charge_Amt3 ,isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt4,0) as Add_Charge_Amt4 ," &
            '    " isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt5,0) as Add_Charge_Amt5 ,isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt6,0) as Add_Charge_Amt6 ," &
            '    " isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt7,0) as Add_Charge_Amt7 ,isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt8,0) as Add_Charge_Amt8 ,isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt9,0) as Add_Charge_Amt9,isnull(TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Amt10,0) as Add_Charge_Amt10," &
            '    " TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name1,TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name2," &
            '    " TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name3,TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name4," &
            '    " TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name5,TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name6," &
            '    " TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name7,TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name8," &
            '    " TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name9,TSPL_PURCHASE_ORDER_HEAD.Add_Charge_Name10 " &
            '    " from TSPL_PAYMENT_HEADER left join TSPL_PAYMENT_DETAIL_GST on TSPL_PAYMENT_HEADER.Payment_No =TSPL_PAYMENT_DETAIL_GST.Payment_No " &
            '    " left join TSPL_COMPANY_MASTER on TSPL_PAYMENT_HEADER.Comp_Code=TSPL_COMPANY_MASTER.Comp_Code " &
            '    " left join TSPL_ITEM_MASTER on TSPL_PAYMENT_DETAIL_GST.item_code=TSPL_ITEM_MASTER.Item_Code " &
            '    " left join TSPL_STATE_MASTER as Comp_State_Master on TSPL_COMPANY_MASTER.State=Comp_State_Master.STATE_CODE" &
            '    " left join TSPL_VENDOR_MASTER on TSPL_PAYMENT_HEADER.Vendor_Code =TSPL_VENDOR_MASTER.Vendor_Code " &
            '    " left join TSPL_STATE_MASTER AS Vendor_State_Master on TSPL_VENDOR_MASTER.State_Code=Vendor_State_Master.STATE_CODE " &
            '    " left join TSPL_PURCHASE_ORDER_HEAD on TSPL_PAYMENT_HEADER.PurchaseOrder_No_GST=TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No" &
            '    " left join TSPL_LOCATION_MASTER on TSPL_PURCHASE_ORDER_HEAD.Bill_To_Location=TSPL_LOCATION_MASTER.Location_Code" &
            '    " left join TSPL_STATE_MASTER Location_State_Master on TSPL_LOCATION_MASTER.state= Location_State_Master.STATE_CODE " &
            '      " LEFT JOIN TSPL_LOCATION_MASTER Ship_to_Location_master on TSPL_PURCHASE_ORDER_HEAD.Ship_To_Location =Ship_to_Location_master.Location_Code" &
            '    "  left outer join TSPL_STATE_MASTER as Ship_to_State_Master on Ship_to_Location_master.State=Ship_to_State_Master.STATE_CODE " &
            '    " LEFT JOIN TSPL_TAX_MASTER DTAX1 ON DTAX1.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX1 " &
            '" LEFT JOIN TSPL_TAX_MASTER DTAX2 ON DTAX2.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX2 " &
            '" LEFT JOIN TSPL_TAX_MASTER DTAX3 ON DTAX3.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX3 " &
            '" LEFT JOIN TSPL_TAX_MASTER DTAX4 ON DTAX4.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX4 " &
            '" LEFT JOIN TSPL_TAX_MASTER DTAX5 ON DTAX5.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX5 " &
            '" LEFT JOIN TSPL_TAX_MASTER DTAX6 ON DTAX6.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX6 " &
            '" LEFT JOIN TSPL_TAX_MASTER DTAX7 ON DTAX7.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX7 " &
            ' " LEFT JOIN TSPL_TAX_MASTER DTAX8 ON DTAX8.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX8 " &
            '  " LEFT JOIN TSPL_TAX_MASTER DTAX9 ON DTAX9.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX9 " &
            '  " LEFT JOIN TSPL_TAX_MASTER DTAX10 ON DTAX10.Tax_Code=TSPL_PAYMENT_DETAIL_GST.TAX10 " &
            '    " WHERE TSPL_PAYMENT_HEADER.Payment_No='" + txtPaymentNo.Value + "'"

            '    Dim dt As DataTable = clsDBFuncationality.GetDataTable(StrQuery)
            '    If dt.Rows.Count > 0 Then
            '        Dim frmCRV As New frmCrystalReportViewer()
            '        If clsCommon.CompairString(clsCommon.myCstr(dt.Rows(0)("Vendor_StateCode")), clsCommon.myCstr(dt.Rows(0)("Location_State"))) = CompairStringResult.Equal Then
            '            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, "RptPaymentVoucher_SGST_CGST", "Payment Voucher", DocDate)
            '        Else
            '            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.SalesReport, dt, "RptPaymentVoucher_IGST", "Payment Voucher", DocDate)
            '        End If
            '        frmCRV = Nothing
            '    End If
        Else
            Dim frm As New FrmPaymentEntry
            FrmPaymentEntry.strCode = txtPaymentNo.Value
            Dim arr As New ArrayList()
            arr.Add(txtPaymentNo.Value)
            FrmPaymentEntry.funReport("", "", arr, Nothing, Nothing)
        End If

    End Sub



    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        ChequeNo = txtChequeNo.Text
        ChequeDate = dtpChequeDate.Value
        strPaymentNo = txtPaymentNo.Value
        Amount = clsCommon.myCdbl(txtPaymentAmt.Text)
        EntryDesc = txtDescription.Text
        Me.Close()
    End Sub

    Private Sub txtLoadOutno__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtLoadOutno._MYValidating
        Dim Qry As String = "select Transfer_No as [TransferNo],Transfer_Date as [TransferDate],Salesmancode ,Reference_Doc_No as [Ref. Document No] from tspl_transfer_head"
        Dim strWhrcls As String = "Location_Type='Logical' and Post='Y' and Transfer_Type  ='Lo' "
        txtLoadOutno.Value = clsCommon.ShowSelectForm("LoadOut", Qry, "TransferNo", strWhrcls, txtLoadOutno.Value, "TransferNo", isButtonClicked)

    End Sub

    Private Sub btnReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReverse.Click
        Try
            If common.clsCommon.MyMessageBoxShow("Reverse and Unpost the Current Document" + Environment.NewLine + "Are you sure", Me.Text, MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
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
                If clsPaymentHeader.ReverseAndUnpost(txtPaymentNo.Value) Then
                    saveCancelLog(Reason, "Reverse And Recreate", Nothing)
                    common.clsCommon.MyMessageBoxShow(Me, "Successfully Reversed and Recreated", Me.Text)
                    LoadData(txtPaymentNo.Value, NavigatorType.Current)
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

    Private Sub txtConversionRate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtConversionRate.TextChanged
        Me.txtTotalPaymentBaseCurr.Text = clsCommon.myCdbl(txtPaymentAmt.Text) * clsCommon.myCdbl(txtConversionRate.Text)
    End Sub
    Private Sub txtCFormInvNo__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean)
        '' Anubhooti 23-Nov-2014 BM00000004668 (Remove C-From) And Anubhooti 08-Dec-2014 (Remove Invoice No.)
        'If Not String.IsNullOrEmpty(ddlPaymentType.SelectedValue) Then 'And chkCForm.Checked = True Then
        '    If Not String.IsNullOrEmpty(txtVendorCode.Value) Then
        '        Dim qry As String = "select PI_No,PI_Date from TSPL_PI_HEAD "
        '        txtCFormInvNo.Value = clsCommon.ShowSelectForm("InvoiceNo", qry, "PI_No", "Posting_Date is not null and Against_C_Form=1 and CFormRecd=0 and CFormApplied=0 and Vendor_Code='" & txtVendorCode.Value & " ' and PI_No not in (select isnull(CForm_InvoiceNo,'')  from  TSPL_Payment_HEADER) ", txtCFormInvNo.Value, "PI_No", isButtonClicked)
        '    Else
        '        clsCommon.MyMessageBoxShow("Please select Vendor before selecting Invoice No")
        '    End If
        'End If
    End Sub



    Private Sub chkCheckPrint_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkCheckPrint.ToggleStateChanged
        If Me.chkCheckPrint.Checked Then
            Me.txtChequeNo.Enabled = False
            'Me.txtChequeNo.Text = ""
        Else
            Me.txtChequeNo.Enabled = True
        End If
    End Sub

    Private Sub btnPrintCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintCheck.Click
        If chkCheckPrint.Checked Then
            '' get data of payment entry
            Dim obj As New clsPaymentHeader()
            obj = clsPaymentHeader.GetData(Me.txtPaymentNo.Value, NavigatorType.Current)
            Dim frm As New frmPrintCheck
            frmPrintCheck.Manual_Print = 0
            'frm.Manual_Print = 0
            frmPrintCheck.BankCode = obj.Bank_Code
            frmPrintCheck.CheckCode = obj.CHECK_CODE
            frmPrintCheck.lblCheckDesc.Text = connectSql.RunScalar("select description from TSPL_BANK_CHECK_PRINTING where CHECK_CODE = '" + obj.CHECK_CODE + "'")
            frmPrintCheck.DocumentType = "Payment Entry"
            frmPrintCheck.DocumentCode = obj.Payment_No
            '' Anubhooti 22-July-2014 BM00000003161
            frm.chkAccPayee.Checked = IIf(obj.Account_Payee = 1, True, False)
            If clsCommon.myLen(obj.CHECK_CODE) > 0 Then
                frm.btnPrint.Text = "RePrint"
            End If
            ''-------richa 
            If obj.memorndmamt = 0 Then
                frm.chkNotForHighValue.Enabled = True
            Else
                frm.chkNotForHighValue.Enabled = False
            End If
            frm.chkNotForHighValue.Checked = False
            If obj.Account_Payee = 1 Then
                frm.chkAccPayee.Checked = True
            Else
                frm.chkAccPayee.Checked = False
            End If
            ''-----------------
            frm.Show()
        Else
            '' get data of payment entry

            Dim obj As New clsPaymentHeader()
            obj = clsPaymentHeader.GetData(Me.txtPaymentNo.Value, NavigatorType.Current)
            If clsPrintCheck.CheckforVoidCheck(obj.Bank_Code, obj.Cheque_No) Then
                clsCommon.MyMessageBoxShow("Please enter valid Cheque No, Entered Cheque No -" & obj.Cheque_No & " is Void.")
                Exit Sub
            End If
            Dim frm As New frmPrintCheck
            frmPrintCheck.BankCode = obj.Bank_Code
            frmPrintCheck.CheckCode = "" ''obj.CHECK_CODE
            frmPrintCheck.fndCheckCode.Enabled = False
            'frm.lblCheckDesc.Text = connectSql.RunScalar("select description from TSPL_BANK_CHECK_PRINTING where CHECK_CODE = '" + obj.CHECK_CODE + "'")
            frmPrintCheck.DocumentType = "Payment Entry"
            frmPrintCheck.DocumentCode = obj.Payment_No
            '' Anubhooti 22-July-2014 BM00000003161
            frmPrintCheck.chkAccPayee.Checked = IIf(obj.Account_Payee = 1, True, False)
            frmPrintCheck.Manual_Print = 1
            frmPrintCheck.Manual_Check_No = txtChequeNo.Text
            'If clsCommon.myLen(obj.CHECK_CODE) > 0 Then
            '    frm.btnPrint.Text = "RePrint"
            'End If
            ''--------richa 
            If obj.memorndmamt <> 0 Then
                frm.chkNotForHighValue.Enabled = True
            Else
                frm.chkNotForHighValue.Enabled = False
            End If
            frm.chkNotForHighValue.Checked = False
            If obj.Account_Payee = 1 Then
                frm.chkAccPayee.Checked = True
            Else
                frm.chkAccPayee.Checked = False
            End If
            ''---------------
            frm.Show()
        End If


    End Sub

    Private Sub txtmemoamt_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtmemoamt.Validating
        Try
            Convert.ToDecimal(txtmemoamt.Text)
        Catch ex As Exception
            txtmemoamt.Text = "0"
        End Try
    End Sub

    Private Sub chkmemorndm_ToggleStateChanged(ByVal sender As System.Object, ByVal args As Telerik.WinControls.UI.StateChangedEventArgs) Handles chkmemorndm.ToggleStateChanged
        txtmemoamt.Enabled = False
        If chkmemorndm.Checked = True Then
            txtmemoamt.Enabled = True
        End If
    End Sub

    Private Sub txtDocumentNo__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtDocumentNo._MYValidating
        ''richa agarwal against ticket no BM00000008630 on 07-Jan-2016
        'Qry = "Select * from (" & _
        '    " Select Payment_No as [Code], Entry_Desc as [Description], Payment_Date as [Payment Date], Case When Payment_Type='AV' Then 'Advance' Else 'On Account' End As [Payment Type], Payment_Amount as [Payment Amt]," & _
        '    " Balance_Amt-ISNULL((Select SUM(Payment_Amount) from TSPL_PAYMENT_HEADER PH WHERE PH.Posted<>'1' AND PH.Payment_Type='AD' AND PH.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'" + txtPaymentNo.Value + "'),0) as [Bal Amt] from TSPL_PAYMENT_HEADER WHERE Posted='1' AND Vendor_Code='" + txtVendorCode.Value + "' AND Payment_Type IN ('AV','OA') AND Payment_No <> '" + txtPaymentNo.Value + "'" & _
        '    " ) Final"
        '===============================update by preeti gupta Aginst ticket no[BM00000009001]
        'Qry = "Select * from (" & _
        '    " Select Payment_No as [Code], Entry_Desc as [Description], Payment_Date as [Payment Date], Case When Payment_Type='AV' Then 'Advance' Else 'On Account' End As [Payment Type], Payment_Amount+isnull(TDS_Amount ,0) as [Payment Amt]," & _
        '    " Balance_Amt+isnull(TDS_Amount ,0)-ISNULL((Select SUM(Payment_Amount)+sum(isnull(TDS_Amount ,0)) from TSPL_PAYMENT_HEADER PH WHERE PH.Posted<>'1' AND PH.Payment_Type='AD' AND PH.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'" + txtPaymentNo.Value + "'),0) as [Bal Amt]  from TSPL_PAYMENT_HEADER WHERE Posted='1' AND Vendor_Code='" + txtVendorCode.Value + "' AND   IsChkReverse='N' and Payment_Type IN ('AV','OA') AND Payment_No <> '" + txtPaymentNo.Value + "'" & _
        '    " ) Final"

        'Qry = "Select * from (" & _
        '  " Select Payment_No as [Code], Entry_Desc as [Description], Payment_Date as [Payment Date], Case When Payment_Type='AV' Then 'Advance' Else 'On Account' End As [Payment Type], Payment_Amount+isnull(TDS_Amount ,0) as [Payment Amt]," & _
        '  " Payment_Amount+isnull(TDS_Amount ,0)-ISNULL((Select SUM(Payment_Amount)+sum(isnull(TDS_Amount ,0)) from TSPL_PAYMENT_HEADER PH WHERE PH.Payment_Type='AD' AND PH.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'" + txtPaymentNo.Value + "' and PH.IsChkReverse='N' ),0) as [Bal Amt]  from TSPL_PAYMENT_HEADER WHERE Posted='1' AND Vendor_Code='" + txtVendorCode.Value + "' AND   IsChkReverse='N' and Payment_Type IN ('AV','OA') AND Payment_No <> '" + txtPaymentNo.Value + "'" & _
        '  " ) Final"



        '' richa 20 Aug,2017

        Dim strQryForRejectedAmt As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
          "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No ) and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "

        '' this code is wrritten for that debit note which is created auto through PI and Pr is not created against that PI
        Dim strQryForRejectedAmtforNonPR As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
      "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "


        Dim strTaxRecovarableQuery As String = " - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then " &
        " ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'  " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end + " &
        " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y' " &
        " then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end) "



        '' UDL/10/07/18-000202 RICHA 
        '  Qry = "Select * from (" & _
        '" Select Payment_No as [Code], Entry_Desc as [Description], CONVERT(VARCHAR,Payment_Date,103) as [Payment Date], Case When Payment_Type='AV' Then 'Advance' Else 'On Account' End As [Payment Type], Payment_Amount+isnull(TDS_Amount ,0) as [Payment Amt]," & _
        '" Payment_Amount+isnull(TDS_Amount ,0)-ISNULL((Select SUM(Payment_Amount)+sum(isnull(TDS_Amount ,0)) from TSPL_PAYMENT_HEADER PH WHERE PH.Payment_Type='AD' AND PH.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'" + txtPaymentNo.Value + "' and PH.IsChkReverse='N' ),0) as [Bal Amt]  from TSPL_PAYMENT_HEADER WHERE Posted='1' AND Vendor_Code='" + txtVendorCode.Value + "' AND   IsChkReverse='N' and Payment_Type IN ('AV','OA') AND Payment_No <> '" + txtPaymentNo.Value + "'  and Convert(Date,Payment_Date,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' " & Environment.NewLine & _
        '" UNION ALL " & Environment.NewLine & _
        '" Select Document_No AS Code,'' as Description,DocumentDate as [Payment Date],DocType as [Payment Type],OriginalAmt as [Payment Amt],PendingAmt as [Bal Amt] from ( select Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType],  " & _
        '" TSPL_VENDOR_INVOICE_HEAD.Document_No, Case When ISNULL(Against_POInvoice_No,'')<>'' Then Against_POInvoice_No When ISNULL(Against_PurchaseReturn_No,'')<> '' Then Against_PurchaseReturn_No Else Document_No End as PurchaseInvoice," & _
        '" Case When ISNULL(Against_POInvoice_No,'')<>'' Then (Select convert(varchar,PI_Date,103)  from TSPL_PI_HEAD where PI_No =Against_POInvoice_No)  When ISNULL(Against_PurchaseReturn_No,'')<> '' Then (Select convert(varchar,PR_Date,103) from TSPL_PR_HEAD where PR_No  =Against_PurchaseReturn_No ) Else TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date End as DocumentDate, " & _
        '" TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date as [DocDate] ,  " & _
        '" TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No as [VendorInvoiceNo], " & _
        '" TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " as [OriginalAmt]  ," & _
        '" TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount as [TDSAmt], " & _
        '" (TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end ) " + strTaxRecovarableQuery + "  as [NetAmount], " & _
        '" TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end " + strTaxRecovarableQuery + " - ISNULL((Select SUM(Applied_Amount) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No = TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No) ),0) " & _
        ' " -isnull((select sum(isnull(Payment_Amount,0)) from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Applied_Payment =TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_HEADER.Payment_No ) and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' and isnull(TSPL_PAYMENT_HEADER.Applied_Payment ,'')<>'' AND Payment_No <> '" + txtPaymentNo.Value + "'),0)  " & Environment.NewLine & _
        '" " + strQryForRejectedAmt + " " + strQryForRejectedAmtforNonPR + " " & _
        '" -ISNULL((Select SUM(isnull(TSPL_PAYMENT_ADJUSTMENT_DETAIL.Amount,0)) from TSPL_PAYMENT_ADJUSTMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_ADJUSTMENT_DETAIL on TSPL_PAYMENT_ADJUSTMENT_Header.Adjustment_No=TSPL_PAYMENT_ADJUSTMENT_DETAIL.Adjustment_No Where isnull(TSPL_PAYMENT_ADJUSTMENT_Header.Doc_No,'') = isnull(TSPL_VENDOR_INVOICE_HEAD.Document_No,'') ),0) " & _
        '" as [PendingAmt],TSPL_VENDOR_INVOICE_HEAD.ConvRate " & _
        '" ,isnull(( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and TSPL_REVALUATION_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation " & _
        '" from TSPL_VENDOR_INVOICE_HEAD " & _
        '" Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  " & _
        '" WHERE TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,Invoice_Entry_Date,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' and ((ISNULL(RefDocNo,'')= '' or ISNULL((select VI.Document_Type from TSPL_VENDOR_INVOICE_HEAD VI where VI.Document_No =TSPL_VENDOR_INVOICE_HEAD .RefDocNo ),'') in ('I','')) AND ISNULL(Against_PurchaseReturn_No,'')= '')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D') " & _
        ' " and TSPL_VENDOR_INVOICE_HEAD.RefDocType<>'REVALUATION ENTRY'" & _
        ' " and isnull(TSPL_VENDOR_INVOICE_HEAD.posting_date,'')<>'') FINALQRY WHERE FINALQRY.PendingAmt>0  AND DocType ='Debit Note' and Convert(Date,DocumentDate,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' " & _
        '" ) Final"

        Qry = "Select * from (" &
    " Select Payment_No as [Code], Entry_Desc as [Description], CONVERT(VARCHAR,Payment_Date,103) as [Payment Date], Case When Payment_Type='AV' Then 'Advance' Else 'On Account' End As [Payment Type], Payment_Amount+isnull(TDS_Amount ,0) as [Payment Amt]," &
    " Payment_Amount+isnull(TDS_Amount ,0)-ISNULL((Select SUM(Payment_Amount)+sum(isnull(TDS_Amount ,0)) from TSPL_PAYMENT_HEADER PH WHERE PH.Payment_Type='AD' AND PH.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code AND PH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No AND PH.Payment_No<>'" + txtPaymentNo.Value + "' and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No in (select PPH.Payment_No from TSPL_PAYMENT_HEADER PPH where PPH.Applied_Payment=TSPL_PAYMENT_HEADER.Payment_No and isnull(TSPL_BANK_REVERSE.POST,'')='P' and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "')) ),0) as [Bal Amt]  from TSPL_PAYMENT_HEADER WHERE Posted='1' AND Vendor_Code='" + txtVendorCode.Value + "' AND   IsChkReverse='N' and Payment_Type IN ('AV','OA') AND Payment_No <> '" + txtPaymentNo.Value + "'  and Convert(Date,Payment_Date,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' " & Environment.NewLine &
    " UNION ALL " & Environment.NewLine &
    " Select Document_No AS Code,'' as Description,DocumentDate as [Payment Date],DocType as [Payment Type],OriginalAmt as [Payment Amt],PendingAmt as [Bal Amt] from ( select Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType],  " &
    " TSPL_VENDOR_INVOICE_HEAD.Document_No, Case When ISNULL(Against_POInvoice_No,'')<>'' Then Against_POInvoice_No When ISNULL(Against_PurchaseReturn_No,'')<> '' Then Against_PurchaseReturn_No Else Document_No End as PurchaseInvoice," &
    " Case When ISNULL(Against_POInvoice_No,'')<>'' Then (Select convert(varchar,PI_Date,103)  from TSPL_PI_HEAD where PI_No =Against_POInvoice_No)  When ISNULL(Against_PurchaseReturn_No,'')<> '' Then (Select convert(varchar,PR_Date,103) from TSPL_PR_HEAD where PR_No  =Against_PurchaseReturn_No ) Else TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date End as DocumentDate, " &
    " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date as [DocDate] ,  " &
    " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No as [VendorInvoiceNo], " &
    " TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " as [OriginalAmt]  ," &
    " TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount as [TDSAmt], " &
    " (TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end ) " + strTaxRecovarableQuery + "  as [NetAmount], " &
    " TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end " + strTaxRecovarableQuery + " - ISNULL((Select SUM(Applied_Amount) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No = TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No) ),0) " &
     " -isnull((select sum(isnull(Payment_Amount,0)) from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Applied_Payment =TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_HEADER.Payment_No ) and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' and isnull(TSPL_PAYMENT_HEADER.Applied_Payment ,'')<>'' AND Payment_No <> '" + txtPaymentNo.Value + "'),0)  " & Environment.NewLine &
    " " + strQryForRejectedAmt + " " + strQryForRejectedAmtforNonPR + " " &
    " -ISNULL((Select SUM(isnull(TSPL_PAYMENT_ADJUSTMENT_DETAIL.Amount,0)) from TSPL_PAYMENT_ADJUSTMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_ADJUSTMENT_DETAIL on TSPL_PAYMENT_ADJUSTMENT_Header.Adjustment_No=TSPL_PAYMENT_ADJUSTMENT_DETAIL.Adjustment_No Where isnull(TSPL_PAYMENT_ADJUSTMENT_Header.Doc_No,'') = isnull(TSPL_VENDOR_INVOICE_HEAD.Document_No,'') ),0) " &
    " as [PendingAmt],TSPL_VENDOR_INVOICE_HEAD.ConvRate " &
    " ,isnull(( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and TSPL_REVALUATION_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation " &
    " from TSPL_VENDOR_INVOICE_HEAD " &
    " Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  " &
    " WHERE TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,Invoice_Entry_Date,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' and ((ISNULL(RefDocNo,'')= '' or ISNULL((select VI.Document_Type from TSPL_VENDOR_INVOICE_HEAD VI where VI.Document_No =TSPL_VENDOR_INVOICE_HEAD .RefDocNo ),'') in ('I','')) AND ISNULL(Against_PurchaseReturn_No,'')= '')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D') " &
     " and TSPL_VENDOR_INVOICE_HEAD.RefDocType<>'REVALUATION ENTRY'" &
     " and isnull(TSPL_VENDOR_INVOICE_HEAD.posting_date,'')<>'') FINALQRY WHERE FINALQRY.PendingAmt>0  AND DocType ='Debit Note' and Convert(Date,DocumentDate,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' " &
    " ) Final"


        txtDocumentNo.Value = clsCommon.ShowSelectForm("Payment@Payment", Qry, "Code", "[Bal Amt]>0", txtPaymentNo.Value, "Code", isButtonClicked)
        lblBalAmt.Text = clsPaymentHeader.GetBalance(txtDocumentNo.Value, txtPaymentNo.Value, Nothing)

        Dim strdocumentType As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Document_Type from TSPL_VENDOR_INVOICE_HEAD where Document_No='" & txtDocumentNo.Value & "'"))

        '' done by Panch Raj ticket no: BM00000008678 
        If clsCommon.myLen(txtDocumentNo.Value) > 0 Then
            Dim dt As DataTable = clsDBFuncationality.GetDataTable("select xx.Bank_Code,xx.Payment_Code,xx.DESCRIPTION,case when isnull(ConvRateRevaluation,0)<>0 then ConvRateRevaluation else ConvRate end as ConvRate  from ( SELECT TSPL_PAYMENT_HEADER.Bank_Code,TSPL_PAYMENT_HEADER.Payment_Code,TSPL_BANK_MASTER.DESCRIPTION,TSPL_PAYMENT_HEADER.ConvRate,isnull( ( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate  from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No and TSPL_REVALUATION_HEAD.Status=1 and  TSPL_REVALUATION_HEAD.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "'  order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation FROM TSPL_PAYMENT_HEADER left join TSPL_BANK_MASTER on TSPL_PAYMENT_HEADER.Bank_Code=TSPL_BANK_MASTER.BANK_CODE  WHERE TSPL_PAYMENT_HEADER.Payment_No ='" & txtDocumentNo.Value & "')xx")
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                txtBankCode.Value = clsCommon.myCstr(dt.Rows(0)("Bank_Code"))
                lblBankDesc.Text = clsCommon.myCstr(dt.Rows(0)("DESCRIPTION"))
                txtPaymentMode.Value = clsCommon.myCstr(dt.Rows(0)("Payment_Code"))
                txtConversionRate.Value = clsCommon.myCdbl(dt.Rows(0)("ConvRate"))
                txtBankCode.Enabled = False
                txtPaymentMode.Enabled = False
                txtConversionRate.ReadOnly = True

            Else
                txtBankCode.Value = ""
                lblBankDesc.Text = ""
                txtPaymentMode.Value = ""
                txtBankCode.Enabled = True
                txtPaymentMode.Enabled = True
                txtConversionRate.ReadOnly = False
            End If
        Else
            txtBankCode.Value = ""
            txtPaymentMode.Value = ""
            txtBankCode.Enabled = True
            txtPaymentMode.Enabled = True
        End If

        If clsCommon.CompairString(strdocumentType, "D") = CompairStringResult.Equal Then
            'txtBankCode.Value = clsCommon.myCstr(dt.Rows(0)("Bank_Code"))
            'lblBankDesc.Text = clsCommon.myCstr(dt.Rows(0)("DESCRIPTION"))
            'txtPaymentMode.Value = clsCommon.myCstr(dt.Rows(0)("Payment_Code"))
            Dim strbancode As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, Nothing))
            If clsCommon.myLen(strbancode) > 0 Then
                Dim bankcode As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select BANK_CODE as [Code] from TSPL_BANK_MASTEr where TSPL_bank_master.INACTIVE ='Active' and TSPL_BANK_MASTER.bank_type<>'S' and Bank_Code='" & clsCommon.myCstr(strbancode) & "' "))
                If clsCommon.myLen(bankcode) <= 0 Then
                    common.clsCommon.MyMessageBoxShow("Please enter Bank Code into fixed parameter.", Me.Text)
                End If
            Else
                common.clsCommon.MyMessageBoxShow("Please enter Bank Code into fixed parameter.", Me.Text)
            End If

            If clsCommon.myLen(strdocumentType) > 0 AndAlso clsCommon.CompairString(strdocumentType, "D") = CompairStringResult.Equal AndAlso clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                txtBankCode.Value = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowDefaultBankCodeforCreditNote, clsFixedParameterCode.AllowDefaultBankCodeforCreditNote, Nothing))
                lblBankDesc.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("sELECT DESCRIPTION FROM TSPL_BANK_MASTER WHERE BANK_CODE ='" & clsCommon.myCstr(txtBankCode.Value) & "'"))
                txtPaymentMode.Value = "NEFT"
            End If
            txtConversionRate.Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select case when isnull(ConvRateRevaluation,0)<>0 then ConvRateRevaluation else ConvRate end as ConvRate from ( Select isnull(TSPL_VENDOR_INVOICE_HEAD.ConvRate,1) as ConvRate, isnull( ( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate  from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and  TSPL_REVALUATION_HEAD.Document_Date <= '" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "'  order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation from TSPL_VENDOR_INVOICE_HEAD where Document_No ='" & txtDocumentNo.Value & "' )xx"))
            txtBankCode.Enabled = False
            txtPaymentMode.Enabled = False
            txtConversionRate.ReadOnly = True

            Qry = "select  Vendor_Bank_Code,Vendor_Bank_Name ,Branch_IFSC_Code,Branch_Name,Vendor_Bank_ACNo from TSPL_VENDOR_INVOICE_HEAD where Document_No ='" + clsCommon.myCstr(txtDocumentNo.Value) + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                txtVendor_bankcode.Text = clsCommon.myCstr(dt.Rows(0)("Vendor_Bank_Code"))
                TxtVendor_BankName.Text = clsCommon.myCstr(dt.Rows(0)("Vendor_Bank_Name"))
                TxtVendorBank_IFSCCode.Text = clsCommon.myCstr(dt.Rows(0)("Branch_IFSC_Code"))
                txtVendorBank_branchname.Text = clsCommon.myCstr(dt.Rows(0)("Branch_Name"))
                txtVendor_Bank_ACNo.Text = clsCommon.myCstr(dt.Rows(0)("Vendor_Bank_ACNo"))
            Else
                txtVendor_bankcode.Text = ""
                TxtVendor_BankName.Text = ""
                TxtVendorBank_IFSCCode.Text = ""
                txtVendorBank_branchname.Text = ""
                txtVendor_Bank_ACNo.Text = ""
            End If

        End If

        LoadVendorInvoices(txtVendorCode.Value)
        'If clsCommon.myCdbl(lblCustomerOutStanding.Text) <> 0 AndAlso common.clsCommon.MyMessageBoxShow("OutStanding of vendor is (" & lblCustomerOutStanding.Text & "). Do You Want To consider it ?", "Payment", MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes Then
        '    ApplyAmount(clsCommon.myCdbl(lblBalAmt.Text) - clsCommon.myCdbl(lblCustomerOutStanding.Text))
        'Else
        ApplyAmount(clsCommon.myCdbl(lblBalAmt.Text))
        'End If
    End Sub

    Private Sub ApplyAmount(ByVal BalAmt As Double)
        Try
            Dim tempAmt As Decimal = BalAmt
            Dim PaymentAmt As Decimal = 0
            For Each grow As GridViewRowInfo In gvDetails.Rows
                If tempAmt > 0 Then
                    grow.Cells(colApply).Value = "Yes"
                    If clsCommon.myCdbl(grow.Cells(colTemp).Value) <= tempAmt Then
                        grow.Cells(colAppliedAmt).Value = grow.Cells(colTemp).Value
                        grow.Cells(colPendingAmt).Value = 0.0
                    ElseIf clsCommon.myCdbl(grow.Cells(colTemp).Value) > tempAmt Then
                        grow.Cells(colAppliedAmt).Value = tempAmt
                        grow.Cells(colPendingAmt).Value = clsCommon.myCdbl(grow.Cells(colTemp).Value) - tempAmt
                    End If
                    grow.Cells(colAppliedAmt).ReadOnly = False
                    If Not clsCommon.CompairString(grow.Cells(colDocType).Value, "Debit Note") = CompairStringResult.Equal Then
                        tempAmt = tempAmt - clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        PaymentAmt = PaymentAmt + clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                    Else
                        tempAmt = tempAmt + clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                        PaymentAmt = PaymentAmt - clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                    End If
                Else
                    Exit For
                End If
            Next
            txtPaymentAmt.Text = clsCommon.myCstr(PaymentAmt)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    '' Anubhooti 21-Aug-2014
    Private Sub txtPONo__MYValidating(ByVal sender As Object, ByVal e As System.EventArgs, ByVal isButtonClicked As Boolean) Handles txtPONo._MYValidating
        If Not String.IsNullOrEmpty(txtVendorCode.Value) Then
            Dim qry As String = "select PurchaseOrder_No As PONo,PurchaseOrder_Date,PurchaseOrder_Type,Vendor_Code,Description,Tax_Group ,Bill_To_Location,Ship_To_Location from TSPL_PURCHASE_ORDER_HEAD "
            txtPONo.Value = clsCommon.ShowSelectForm("PONo", qry, "PONo", "Vendor_Code='" & txtVendorCode.Value & " ' ", txtPONo.Value, "PONo", isButtonClicked)
        Else
            clsCommon.MyMessageBoxShow("Please select Vendor before selecting Purchase Order No", Me.Text)
        End If
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Try
            'If clsCommon.myCdbl(lblCustomerOutStanding.Text) <> 0 AndAlso common.clsCommon.MyMessageBoxShow("OutStanding of vendor is (" & lblCustomerOutStanding.Text & "). Do You Want To consider it ?", "Payment", MessageBoxButtons.YesNo, RadMessageIcon.Question) = Windows.Forms.DialogResult.Yes Then
            '    ApplyAmount(clsCommon.myCdbl(txtPaymentAmt.Text) - clsCommon.myCdbl(lblCustomerOutStanding.Text))
            'Else
            ApplyAmount(clsCommon.myCdbl(txtPaymentAmt.Text))
            ' End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub ChkSecurity_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkSecurity.CheckStateChanged
        Try
            If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "PY") = CompairStringResult.Equal Then
                If ChkSecurity.Checked Then
                    gvDetails.Columns(colSecurityAmt).IsVisible = True
                    gvDetails.Columns(colSecurityAmt).Width = 100
                Else
                    gvDetails.Columns(colSecurityAmt).IsVisible = False
                    gvDetails.Columns(colSecurityAmt).Width = 0
                End If
            End If
            securityRefund()
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    Sub securityRefund()
        If ChkSecurity.Checked Then
            IsPaymentTypeChanged = False
            dt = New DataTable
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Value", GetType(String))

            dt.Rows.Add("Payment", "PY")
            dt.Rows.Add("Advance", "AV")
            dt.Rows.Add("On Account", "OA")
            dt.Rows.Add("Miscellaneous", "MI")
            dt.Rows.Add("Receipt", "RC")
            dt.Rows.Add("Apply Document", "AD")
            dt.Rows.Add("Security Refund", "SR")

            ddlPaymentType.DataSource = dt
            ddlPaymentType.DisplayMember = "Code"
            ddlPaymentType.ValueMember = "Value"
            ddlPaymentType.SelectedValue = "RC"
            LoadBlankGrid(ddlPaymentType.SelectedValue)
            IsPaymentTypeChanged = True
        Else
            IsPaymentTypeChanged = False
            dt = New DataTable
            dt.Columns.Add("Code", GetType(String))
            dt.Columns.Add("Value", GetType(String))

            dt.Rows.Add("Payment", "PY")
            dt.Rows.Add("Advance", "AV")
            dt.Rows.Add("On Account", "OA")
            dt.Rows.Add("Miscellaneous", "MI")
            dt.Rows.Add("Receipt", "RC")
            dt.Rows.Add("Apply Document", "AD")

            ddlPaymentType.DataSource = dt
            ddlPaymentType.DisplayMember = "Code"
            ddlPaymentType.ValueMember = "Value"
            LoadBlankGrid(ddlPaymentType.SelectedValue)
            IsPaymentTypeChanged = True
        End If

    End Sub
    '' Anubhooti 07-Jan-2014 (BM00000005309)
    Private Sub txtlocation__MYValidating(ByVal sender As System.Object, ByVal e As System.EventArgs, ByVal isButtonClicked As System.Boolean) Handles txtlocation._MYValidating
        Try
            Dim qry As String = "select distinct(Segment_code) as Code ,Description  from TSPL_GL_SEGMENT_CODE left outer join TSPL_LOCATION_MASTER on TSPL_GL_SEGMENT_CODE .Segment_code =TSPL_LOCATION_MASTER .Loc_Segment_Code "
            Dim WhrCls As String = "Seg_No = '7' AND GIT='N'"
            If clsCommon.myLen(objCommonVar.strCurrUserLocations) > 0 Then
                WhrCls += "  and  TSPL_LOCATION_MASTER.Location_Code in (" + objCommonVar.strCurrUserLocations + ")"
            End If
            txtlocation.Value = clsCommon.ShowSelectForm("PELoc", qry, "Code", WhrCls, txtlocation.Value, "Code", isButtonClicked)
            If clsCommon.myLen(clsCommon.myCstr(txtlocation.Value)) > 0 Then
                LblLocDesp.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("SELECT ISNULL(Description,'') As Description FROM TSPL_GL_SEGMENT_CODE WHERE Segment_code ='" & clsCommon.myCstr(txtlocation.Value) & "'"))
            Else
                LblLocDesp.Text = ""
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkPDC_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPDC.CheckStateChanged
        If PDCSetting Then
            If chkPDC.Checked = True AndAlso clsCommon.myLen(dtpChequeDate.Value) > 0 Then
                dtpPayment.Value = dtpChequeDate.Value
                'Else
                '    dtpPayment.Value = clsCommon.GETSERVERDATE()
            End If
        End If

    End Sub

    Private Sub dtpChequeDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpChequeDate.Validating
        If chkPDC.Checked = True AndAlso clsCommon.myLen(dtpChequeDate.Value) > 0 Then
            dtpPayment.Value = dtpChequeDate.Value
            'Else
            'dtpPayment.Value = clsCommon.GETSERVERDATE()
        End If
    End Sub

    Private Sub dtpChequeDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpChequeDate.ValueChanged
        'If chkPDC.Checked = True AndAlso clsCommon.myLen(dtpChequeDate.Value) > 0 Then
        '    dtpPayment.Value = dtpChequeDate.Value
        'End If
    End Sub

    Private Sub dtpPayment_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpPayment.Validating
        If PDCSetting Then
            If chkPDC.Checked = True AndAlso clsCommon.myLen(dtpChequeDate.Value) > 0 Then
                dtpPayment.Value = dtpChequeDate.Value
                'Else
                'dtpPayment.Value = clsCommon.GETSERVERDATE()
            End If
        End If
    End Sub

    Private Sub dtpPayment_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpPayment.ValueChanged
        'If chkPDC.Checked = True AndAlso clsCommon.myLen(dtpChequeDate.Value) > 0 Then
        '    dtpPayment.Value = dtpChequeDate.Value
        'End If
    End Sub


    Private Sub rmiExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rmiExport.Click
        Try
            Qry = "Select '' as [Payment Date], '' as [Description], '' as [Vendor Code], '' as [Bank Code], '' as [Payment Type(A/O/R/M)], '' as [Payment Mode], '' as [Cheque No], '' as [Cheque Date], 0 as Amount,'' as [Location Code],0 as [Advance Against Salary],0 as [Is Opening],0 as [Bank Charges],0 as [Is Security],1 as [Conv Rate],'' as GLAccount,'' AS [Employee Advance Type(S/T/I)],'' AS [Employee Expense Type(TD/S/T/I)],'' as [MP Code],'' as [VLC Uploader Code]"
            transportSql.ExporttoExcel(Qry, "", Me)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub rmiImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rmiImport.Click
        Try
            funImport(False)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Public Sub funImport(ByVal IsForPost As Boolean)
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim Count As String = ""
        If transportSql.importExcel(gv, "Payment Date", "Description", "Vendor Code", "Bank Code", "Payment Type(A/O/R/M)", "Payment Mode", "Cheque No", "Cheque Date", "Amount", "Location Code", "Advance Against Salary", "Is Opening", "Bank Charges", "Is Security", "Conv Rate", "GLAccount", "Employee Advance Type(S/T/I)", "Employee Expense Type(TD/S/T/I)", "MP Code", "VLC Uploader Code") Then
            Dim trans As SqlTransaction = Nothing
            Dim linno As Integer = 0
            Try
                connectSql.OpenConnection()
                trans = clsDBFuncationality.GetTransactin()
                clsCommon.ProgressBarPercentShow()
                Dim obj As clsPaymentHeader
                For Each grow As GridViewRowInfo In gv.Rows
                    clsCommon.ProgressBarPercentUpdate((grow.Index + 1) * 100 / (gv.Rows.Count + 1), "Importing  : " & (grow.Index + 1) & "/" & gv.Rows.Count & "")
                    Count = clsCommon.myCstr(grow.Index + 2)
                    If clsCommon.myLen(grow.Cells("Payment Date").Value) > 0 Then
                        obj = New clsPaymentHeader()
                        obj.Payment_No = ""

                        obj.Entry_Desc = clsCommon.myCstr(grow.Cells("Description").Value)
                        If clsCommon.myLen(obj.Entry_Desc) > 250 Then
                            Throw New Exception("Description Length can not be more than 250.")
                        End If

                        obj.Payment_Date = clsCommon.myCDate(grow.Cells("Payment Date").Value)
                        obj.Payment_Post_Date = obj.Payment_Date

                        obj.Payment_Type = clsCommon.myCstr(grow.Cells("Payment Type(A/O/R/M)").Value)
                        If clsCommon.CompairString(obj.Payment_Type, "O") = CompairStringResult.Equal Then
                            obj.Payment_Type = "OA"
                        ElseIf clsCommon.CompairString(obj.Payment_Type, "A") = CompairStringResult.Equal Then
                            obj.Payment_Type = "AV"
                        ElseIf clsCommon.CompairString(obj.Payment_Type, "R") = CompairStringResult.Equal Then
                            obj.Payment_Type = "RC"
                        ElseIf clsCommon.CompairString(obj.Payment_Type, "M") = CompairStringResult.Equal Then
                            obj.Payment_Type = "MI"
                        Else
                            Throw New Exception("Payment type can be 'O' or 'A' or 'R' or 'M'.")
                        End If
                        Dim strVendorCode As String = String.Empty
                        If clsCommon.myLen(clsCommon.myCstr(grow.Cells("VLC Uploader Code").Value)) > 0 Then
                            If clsCommon.myLen(clsCommon.myCstr(grow.Cells("Location Code").Value)) > 0 Then
                                strVendorCode = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select VSP_Code  from TSPL_VLC_MASTER_HEAD where vlc_code_vlc_uploader='" & clsCommon.myCstr(grow.Cells("VLC Uploader Code").Value) & "' and MCC= (select Location_Code  from TSPL_LOCATION_MASTER where loc_segment_code ='" & clsCommon.myCstr(grow.Cells("Location Code").Value) & "') ", trans))
                                If clsCommon.myLen(strVendorCode) <= 0 Then
                                    Throw New Exception("VSP not exists for VLC Uploader Code  " & clsCommon.myCstr(grow.Cells("VLC Uploader Code").Value) & " and Location " & clsCommon.myCstr(grow.Cells("Location Code").Value) & ".")
                                End If
                            Else
                                Throw New Exception("Please enter Location Code.")
                            End If

                        End If

                        If clsCommon.myLen(strVendorCode) > 0 Then
                            obj.Vendor_Code = strVendorCode
                        Else
                            obj.Vendor_Code = clsCommon.myCstr(grow.Cells("Vendor Code").Value)
                        End If

                        If clsCommon.CompairString(obj.Payment_Type, "MI") = CompairStringResult.Equal Then
                            obj.Vendor_Code = ""
                            obj.Vendor_Name = ""
                            obj.MP_Code_For_Advance = clsCommon.myCstr(grow.Cells("MP Code").Value)
                            If clsCommon.myLen(obj.MP_Code_For_Advance) > 0 Then
                                obj.MP_Code_For_Advance = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select MP_Code from TSPL_MP_MASTER where MP_Code='" + obj.MP_Code_For_Advance + "'", trans))
                                If clsCommon.myLen(obj.MP_Code_For_Advance) <= 0 Then
                                    Throw New Exception("MP Code [" + clsCommon.myCstr(grow.Cells("MP Code").Value) + "] not exists in MP Master.")
                                End If
                            End If
                        Else
                            If clsCommon.myLen(obj.Vendor_Code) > 0 Then
                                obj.Vendor_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Code from TSPL_VENDOR_MASTER WHERE Vendor_Code='" + obj.Vendor_Code + "'", trans))
                                If clsCommon.myLen(obj.Vendor_Code) <= 0 Then
                                    Throw New Exception("Vendor Code does not exist.")
                                End If
                                obj.Vendor_Name = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Name from TSPL_VENDOR_MASTER WHERE Vendor_Code='" + obj.Vendor_Code + "'", trans))
                            Else
                                Throw New Exception("Please select Vendor Code.")
                            End If
                        End If


                        obj.Bank_Code = clsCommon.myCstr(grow.Cells("Bank Code").Value)
                        If clsCommon.myLen(obj.Bank_Code) > 0 Then
                            obj.Bank_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select BANK_CODE from TSPL_BANK_MASTER WHERE Bank_Code='" + obj.Bank_Code + "'", trans))
                            If clsCommon.myLen(obj.Bank_Code) <= 0 Then
                                Throw New Exception("Bank Code does not exist.")
                            End If

                            ''richa agarwal 13 May, 2017
                            If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select TSPL_bank_master.INACTIVE from TSPL_BANK_MASTER Where Bank_Code='" & obj.Bank_Code & "'", trans)), "Active") <> CompairStringResult.Equal Then
                                Throw New Exception("Bank Code should be Active .")
                            End If
                            ''--------------
                        Else
                            Throw New Exception("Please select Bank Code.")
                        End If



                        obj.Payment_Code = clsCommon.myCstr(grow.Cells("Payment Mode").Value)
                        If clsCommon.myLen(obj.Payment_Code) > 0 Then
                            obj.Payment_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_PAYMENT_CODE.Payment_Code from TSPL_PAYMENT_CODE Where TSPL_PAYMENT_CODE.Payment_Code='" + obj.Payment_Code + "'", trans))
                            If clsCommon.myLen(obj.Payment_Code) <= 0 Then
                                Throw New Exception("Payment Mode does not exist.")
                            End If
                        Else
                            Throw New Exception("Enter Payment Mode.")
                        End If


                        If clsCommon.CompairString(obj.Payment_Code, "Cheque") = CompairStringResult.Equal Then
                            obj.Cheque_No = clsCommon.myCstr(grow.Cells("Cheque No").Value)
                            If clsCommon.myLen(obj.Cheque_No) > 0 Then
                                If clsCommon.myLen(obj.Cheque_No) < 6 Or clsCommon.myLen(obj.Cheque_No) > 20 Then
                                    Throw New Exception("Length of Cheque No should between 6-20.")
                                End If
                                Qry = "Select Payment_No from TSPL_PAYMENT_HEADER Where Cheque_No='" & obj.Cheque_No & "'"
                                Qry = clsCommon.myCstr(clsDBFuncationality.getSingleValue(Qry, trans))
                                If clsCommon.myLen(Qry) > 0 Then
                                    Throw New Exception("Cheque '" & obj.Cheque_No & "' is Already Exists Against Payment '" & Qry & "'")
                                End If
                                '----------------------Cheque Date---------------------
                                If clsCommon.myLen(grow.Cells("Cheque Date").Value) > 0 Then
                                    obj.Cheque_Date = CDate(grow.Cells("Cheque Date").Value)
                                Else
                                    Throw New Exception("Please enter Cheque Date.")
                                End If
                                '------------------------------------------------------
                            Else
                                Throw New Exception("Cheque No can't be blank")
                            End If
                        End If

                        If clsCommon.myCdbl(grow.Cells("Amount").Value) <= 0 Then
                            Throw New Exception("Enter Payment Amount.")
                        End If

                        obj.Account_Payee = 0
                        obj.Location_GL_Code = ""

                        obj.memorndmamt = 0.0

                        obj.CHECK_PRINT = 0

                        ''richa agarwal 09/07/2015
                        Dim dblAdvanceAgainstSalary As Double = clsCommon.myCdbl(grow.Cells("Advance Against Salary").Value)
                        If dblAdvanceAgainstSalary = 0 Or dblAdvanceAgainstSalary = 1 Then
                            obj.Advance_Against_Salary = dblAdvanceAgainstSalary
                        Else
                            Throw New Exception("Advance Against Salary should be 0 or 1.")
                        End If
                        Dim dblIsOpening As Double = clsCommon.myCdbl(grow.Cells("Is Opening").Value)
                        If dblIsOpening = 0 Or dblIsOpening = 1 Then
                            obj.is_Opening = dblIsOpening
                        Else
                            Throw New Exception("Is Opening Advance Against Salary should be 0 or 1.")
                        End If
                        ''--------------------
                        ''richa agarwal against advance type of employee salary integration
                        Dim EmployeeSalaryGeneration As Boolean = clsCommon.myCBool(IIf(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.AllowtoEmployeeSalaryIntegration, clsFixedParameterCode.AllowtoEmployeeSalaryIntegration, trans)) = 1, True, False))
                        If EmployeeSalaryGeneration = True Then
                            If clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                                obj.Employee_Advance_Type = clsCommon.myCstr(grow.Cells("Employee Advance Type(S/T/I)").Value)
                                obj.Employee_Type = clsCommon.myCstr(grow.Cells("Employee Expense Type(TD/S/T/I)").Value)
                                If clsCommon.CompairString(obj.Employee_Advance_Type, "") <> CompairStringResult.Equal AndAlso clsCommon.CompairString(obj.Employee_Type, "") <> CompairStringResult.Equal Then
                                    Throw New Exception("Please enter any one option at a time Employee Expense type/Employee Advance type.")
                                End If
                            End If
                            If clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & obj.Vendor_Code & "' ", trans)), "1") = CompairStringResult.Equal Then
                                    obj.Employee_Advance_Type = clsCommon.myCstr(grow.Cells("Employee Advance Type(S/T/I)").Value)
                                    If clsCommon.CompairString(obj.Employee_Advance_Type, "S") = CompairStringResult.Equal Then
                                        obj.Employee_Advance_Type = "S"
                                        obj.Advance_Against_Salary = 1
                                    ElseIf clsCommon.CompairString(obj.Employee_Advance_Type, "T") = CompairStringResult.Equal Then
                                        obj.Employee_Advance_Type = "T"
                                    ElseIf clsCommon.CompairString(obj.Employee_Advance_Type, "I") = CompairStringResult.Equal Then
                                        obj.Employee_Advance_Type = "I"
                                    ElseIf clsCommon.CompairString(obj.Employee_Advance_Type, "") = CompairStringResult.Equal Then
                                        obj.Employee_Advance_Type = ""
                                    Else
                                        Throw New Exception("Payment type can be 'S' or 'T' or 'I' or ''.")
                                    End If

                                    If clsCommon.CompairString(obj.Advance_Against_Salary, "1") = CompairStringResult.Equal And clsCommon.CompairString(obj.Employee_Advance_Type, "S") <> CompairStringResult.Equal Then
                                        Throw New Exception("Please enter Employee Advance Type 'S' or Advance Against Salary '0'.")
                                    End If
                                End If
                            Else
                                obj.Employee_Advance_Type = ""
                            End If
                            If clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                                If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("select IsEmployee from TSPL_VENDOR_MASTER where Vendor_Code ='" & obj.Vendor_Code & "' ", trans)), "1") = CompairStringResult.Equal Then
                                    obj.Employee_Type = clsCommon.myCstr(grow.Cells("Employee Expense Type(TD/S/T/I)").Value)
                                    If clsCommon.CompairString(obj.Employee_Type, "TD") = CompairStringResult.Equal Then
                                        obj.Employee_Type = "TD"
                                    ElseIf clsCommon.CompairString(obj.Employee_Type, "S") = CompairStringResult.Equal Then
                                        obj.Employee_Type = "S"
                                    ElseIf clsCommon.CompairString(obj.Employee_Type, "T") = CompairStringResult.Equal Then
                                        obj.Employee_Type = "T"
                                    ElseIf clsCommon.CompairString(obj.Employee_Type, "I") = CompairStringResult.Equal Then
                                        obj.Employee_Type = "I"
                                    ElseIf clsCommon.CompairString(obj.Employee_Type, "") = CompairStringResult.Equal Then
                                        obj.Employee_Type = ""
                                    Else
                                        Throw New Exception("Payment type can be 'TD' or 'S' or 'T' or 'I' or ''.")
                                    End If
                                End If
                            Else
                                obj.Employee_Type = ""
                            End If
                        End If
                        ''----------------end of employee salary integration

                        '-----------Location Code--------------------------------
                        Dim Loc_Code As String = clsCommon.myCstr(grow.Cells("Location Code").Value)
                        If clsCommon.myLen(Loc_Code) > 0 Then
                            Loc_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct(Segment_code) as Code ,Description  from TSPL_GL_SEGMENT_CODE left " _
                                       & " outer join TSPL_LOCATION_MASTER on TSPL_GL_SEGMENT_CODE .Segment_code =TSPL_LOCATION_MASTER .Loc_Segment_Code  Where Seg_No = '7' AND GIT='N' and Segment_code = '" & Loc_Code & "'", trans))
                            If Loc_Code = "" Then
                                Throw New Exception("Please Check Location Code dose not Exits") ' + LineNo + " does not exist. ")
                            End If
                        Else
                            Throw New Exception("Insert Location Code in All Rows ") ' + LineNo + ".")
                        End If
                        '--------------------------------------------------------

                        If clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "MI") = CompairStringResult.Equal Then
                            obj.Total_Prepayment = clsCommon.myCdbl(grow.Cells("Amount").Value)
                            obj.Payment_Amount = clsCommon.myCdbl(grow.Cells("Amount").Value)
                            obj.Balance_Amt = clsCommon.myCdbl(grow.Cells("Amount").Value)
                            obj.TDS_Amount = 0
                        End If

                        If clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                            obj.Is_Security = clsCommon.myCdbl(grow.Cells("Is Security").Value)
                            obj.Is_Retention = clsCommon.myCdbl(grow.Cells("Is Retention").Value)
                        Else
                            obj.Is_Security = 0
                            obj.Is_Retention = 0
                        End If
                        obj.IsChkReverse = "N"
                        obj.Bank_Charges = clsCommon.myCdbl(grow.Cells("Bank Charges").Value)
                        obj.objRemittance = Nothing
                        obj.PurchaseOrder_No = ""
                        obj.Loan_Code = ""
                        obj.Account_Payee_Name = ""
                        obj.CURRENCY_CODE = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select currency_code from TSPL_VENDOR_MASTER where VENDOR_CODE='" & obj.Vendor_Code & "'", trans))
                        obj.BASE_CURRENCY_CODE = objCommonVar.BaseCurrencyCode
                        obj.PAYMENT_AMOUNT_BASE_CURRENCY = obj.Payment_Amount
                        obj.ConvRateOld = 1
                        obj.ConvRate = IIf(clsCommon.myCdbl(grow.Cells("Conv Rate").Value) <= 0, 1, clsCommon.myCdbl(grow.Cells("Conv Rate").Value))
                        If clsCommon.CompairString(obj.CURRENCY_CODE, obj.BASE_CURRENCY_CODE) = CompairStringResult.Equal Then
                            obj.ConvRate = 1
                        End If
                        obj.PAYMENT_AMOUNT_BASE_CURRENCY = obj.Payment_Amount * obj.ConvRate

                        obj.Location_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select RIGHT(BANKACC,3) from TSPL_Bank_Master Where Bank_Code='" + obj.Bank_Code + "'", trans))
                        obj.Location_GL_Code = obj.Location_Code
                        obj.PDC_Cheque = "N'"
                        obj.Applied_Payment = ""
                        obj.ArrTr = Nothing
                        obj.Location_GL_Code = Loc_Code
                        If clsCommon.CompairString(obj.Payment_Type, "MI") = CompairStringResult.Equal Then
                            obj.ArrTr = New List(Of clsPaymentDetail)

                            Dim objTr As New clsPaymentDetail()
                            objTr.Payment_Type = obj.Payment_Type
                            If clsCommon.myLen(obj.MP_Code_For_Advance) > 0 Then
                                Dim BankSeg As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" SELECT Loc_Segment_Code  FROM TSPL_LOCATION_MASTER WHERE LOCATION_CODE IN ( SELECT MCC  FROM TSPL_VLC_MASTER_HEAD WHERE VLC_Code IN (select VLC_Code  from TSPL_MP_MASTER WHERE MP_Code ='" + obj.MP_Code_For_Advance + "'))", trans))
                                Dim strDiscountCode As String = clsFixedParameter.GetData(clsFixedParameterType.DiscountCodeForMPAdj, clsFixedParameterCode.DiscountCodeForMPAdj, trans)
                                If clsCommon.myLen(strDiscountCode) <= 0 Then
                                    Throw New Exception("Please first set Value of DiscountCodeForMPAdj in fixed parameter")
                                End If

                                Dim qry As String = "SELECT  TSPL_DISCOUNT_MASTER.Account_Code FROM TSPL_DISCOUNT_MASTER" + Environment.NewLine +
                                "where Code='" + strDiscountCode + "' "
                                qry = clsDBFuncationality.getSingleValue(qry, trans)
                                If clsCommon.myLen(qry) > 0 Then

                                    objTr.Account_Code = clsERPFuncationality.ChangeGLAccountLocationSegment(qry, BankSeg, True, trans)
                                    objTr.Description = clsGLAccount.GetName(clsCommon.myCstr(objTr.Account_Code), trans)
                                Else
                                    Throw New Exception("Please set GL Account in discount Master [" + strDiscountCode + "]")
                                End If
                            Else
                                objTr.Account_Code = clsCommon.myCstr(grow.Cells("GLAccount").Value)
                                If clsCommon.myLen(objTr.Account_Code) <= 0 Then
                                    Throw New Exception("Please enter account code")
                                End If
                                Dim dttmp As DataTable = clsDBFuncationality.GetDataTable("select  Account_Code,Description from TSPL_GL_ACCOUNTS where Account_Code='" + objTr.Account_Code + "'", trans)
                                If clsCommon.myLen(objTr.Account_Code) <= 0 Then
                                    Throw New Exception("Please enter valid account code")
                                End If
                                If clsCommon.GetDateWithStartTime(obj.Payment_Date) >= clsCommon.GetDateWithStartTime(ERPStartDate) Then
                                    ''RICHA AGARWAL TEC/03/07/19-000927
                                    Qry = "select TSPL_GL_ACCOUNTS.Account_Code,Description from TSPL_GL_ACCOUNTS WHERE TSPL_GL_ACCOUNTS.Account_Code='" & objTr.Account_Code & "' and ( ControlAccount='N' or TSPL_GL_ACCOUNTS.Account_Code IN (select TSPL_CONTROL_ACC_MAPPING.Account_Code  from TSPL_CONTROL_ACC_MAPPING where IsForPayment =1))"
                                    dttmp = clsDBFuncationality.GetDataTable(Qry, trans)
                                    If dttmp Is Nothing OrElse dttmp.Rows.Count <= 0 Then
                                        Throw New Exception("Account should be type of non Control Account.")
                                    End If
                                End If
                                objTr.Account_Code = clsCommon.myCstr(dttmp.Rows(0)("Account_Code"))
                                objTr.Description = clsCommon.myCstr(dttmp.Rows(0)("Description"))
                                'objTr.Applied_Amount = clsCommon.myCdbl(grow.Cells(colAmount).Value)

                            End If
                            objTr.Net_Balance = obj.Payment_Amount

                            obj.ArrTr.Add(objTr)
                        End If
                        obj.SaveData1(obj, True, trans)
                    End If
                Next
                trans.Commit()
                clsCommon.ProgressBarPercentHide()
                common.clsCommon.MyMessageBoxShow("Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                trans.Rollback()
                clsCommon.ProgressBarPercentHide()
                clsCommon.MyMessageBoxShow("Erro at Line: " + Count + Environment.NewLine + ex.Message)
            Finally
                Me.Controls.Remove(gv)
            End Try
        End If
    End Sub

    Private Sub btnChqUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnChqUpdate.Click
        Try
            If clsCommon.myLen(clsCommon.myCstr(txtChequeNo.Text)) <= 0 Then
                Throw New Exception("Please enter Cheque No")
            End If
            Dim qry As String = "update TSPL_PAYMENT_HEADER set cheque_No='" & clsCommon.myCstr(txtChequeNo.Text) & "' , cheque_date =' " & clsCommon.GetPrintDate(dtpChequeDate.Value, "dd/MMM/yyyy") & "' where payment_no='" & clsCommon.myCstr(txtPaymentNo.Value) & "'"
            clsDBFuncationality.ExecuteNonQuery(qry)
            myMessages.update()
            LoadData(txtPaymentNo.Value, NavigatorType.Current)
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub btnVoidCheck_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnVoidCheck.Click
        If clsCommon.myLen(txtPaymentNo.Value) <= 0 Then
            clsCommon.MyMessageBoxShow("Select document no to void check.", Me.Text)
            Exit Sub
        End If
        Dim obj As New clsPaymentHeader()
        obj = clsPaymentHeader.GetData(Me.txtPaymentNo.Value, NavigatorType.Current)
        If clsCommon.myLen(obj.Bank_Code) <= 0 Then
            clsCommon.MyMessageBoxShow("Bank Code not found for selected document.", Me.Text)
            Exit Sub
            'ElseIf clsCommon.myLen(obj.CHECK_CODE) <= 0 Then
            '    clsCommon.MyMessageBoxShow("Check Code not found for selected document.")
            '    Exit Sub
        End If
        If clsPrintCheck.VoidCheck(obj.Bank_Code, obj.CHECK_CODE, "Payment Entry", obj.Payment_No) Then
            clsCommon.MyMessageBoxShow("done successfully", Me.Text)
        End If
        ''or  clscommon.myLen(obj.Bank_Code )<=0
    End Sub

    Private Sub btnRtgs_Click(sender As Object, e As EventArgs) Handles btnRtgs.Click
        'Dim Qry1 As String = "select TSPL_PAYMENT_HEADER.Payment_No as Doc_No,Payment_Date,Payment_Amount,Debit_Account,TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_VENDOR_MASTER.Vendor_Name,TSPL_VENDOR_MASTER.bank_code as Vendor_Bank_Code,TSPL_VENDOR_MASTER.IFSC_Code,VendorBank.DESCRIPTION as Vendor_Bank_Name,TSPL_BANK_BRANCH_MASTER.Branch_code as Vendor_Branch_code,TSPL_BANK_BRANCH_MASTER.Branch_Name as Vendor_Branch_Name ,TSPL_PAYMENT_HEADER.Payment_Code,Credit_Account,TSPL_PAYMENT_HEADER.Bank_Code,Bank.DESCRIPTION as Bank_Name,(Bank.Add1+Bank.Add2+Bank.Add3+Bank.Add4) as Address,TSPL_COMPANY_MASTER.Comp_Name from"
        'Qry1 += " TSPL_PAYMENT_HEADER left join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No "
        'Qry1 += " left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code "
        'Qry1 += " left outer join tspl_bank_master as VendorBank on VendorBank.bank_code=TSPL_VENDOR_MASTER.bank_code"
        'Qry1 += " left join TSPL_BANK_BRANCH_MASTER on TSPL_BANK_BRANCH_MASTER.Bank_code=VendorBank.bank_code"
        'Qry1 += " left join TSPL_bank_master as Bank on Bank.bank_code=TSPL_PAYMENT_HEADER.Bank_Code"
        'Qry1 += " left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.comp_code=TSPL_PAYMENT_HEADER.comp_code"
        'Qry1 += " WHERE TSPL_PAYMENT_HEADER.Payment_Code='RTGS' and TSPL_PAYMENT_HEADER.Payment_No='" & txtPaymentNo.Value & "'"

        'Dim Qry As String = "select  Value ,TSPL_PAYMENT_HEADER.Payment_No as Doc_No,convert(varchar,Payment_Date,103)as Payment_Date,Payment_Amount,bank.BANKACCNUMBER as Debit_Account,TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_VENDOR_MASTER.Vendor_Name,TSPL_VENDOR_MASTER.bank_code as Vendor_Bank_Code,TSPL_VENDOR_MASTER.IFSC_Code,VendorBank.DESCRIPTION as Vendor_Bank_Name,TSPL_BANK_BRANCH_MASTER.Branch_code as Vendor_Branch_code,TSPL_BANK_BRANCH_MASTER.Branch_Name as Vendor_Branch_Name ,TSPL_PAYMENT_HEADER.Payment_Code,Account_No as Credit_Account,TSPL_PAYMENT_HEADER.Bank_Code,Bank.DESCRIPTION as Bank_Name,Bank.Add1,Bank.Add2,Bank.Add3,Bank.Add4 ,TSPL_COMPANY_MASTER.Comp_Name from TSPL_PAYMENT_HEADER left join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No  left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code  left outer join tspl_bank_master as VendorBank on VendorBank.bank_code=TSPL_VENDOR_MASTER.bank_code left join TSPL_BANK_BRANCH_MASTER on TSPL_BANK_BRANCH_MASTER.Bank_code=VendorBank.bank_code left join TSPL_bank_master as Bank on Bank.bank_code=TSPL_PAYMENT_HEADER.Bank_Code left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.comp_code=TSPL_PAYMENT_HEADER.comp_code left join (select Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code left join TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code=TSPL_CUSTOM_FIELD_VALUES.Program_Code where TSPL_CUSTOM_FIELD_VALUES.Program_Code='PYMT-NEW'  and "
        'Qry += " Name ='Purpose')tt on tt.transaction_Code=TSPL_PAYMENT_HEADER.Payment_No WHERE TSPL_PAYMENT_HEADER.Payment_Code='RTGS' and TSPL_PAYMENT_HEADER.Payment_No='" & txtPaymentNo.Value & "'"

        'Dim Qry As String = "select Value, TSPL_PAYMENT_HEADER.Payment_No as Doc_No,convert(varchar,Payment_Date,103)as Payment_Date,Payment_Amount,bank.BANKACCNUMBER as Debit_Account,TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_VENDOR_MASTER.Vendor_Name,TSPL_VENDOR_MASTER.bank_code as Vendor_Bank_Code,TSPL_Vendor_Bank_Branch_Details.Bank_IFSC_Code as IFSC_Code,tspl_vendor_bank_Master.Bank_Name as Vendor_Bank_Name,TSPL_Vendor_Bank_Branch_Details.Branch_name as Vendor_Branch_Name,TSPL_PAYMENT_HEADER.Payment_Code,Account_No as Credit_Account,TSPL_PAYMENT_HEADER.Bank_Code,Bank.DESCRIPTION as Bank_Name,Bank.Add1,Bank.Add2,Bank.Add3,Bank.Add4 ,TSPL_COMPANY_MASTER.Comp_Name from TSPL_PAYMENT_HEADER left join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No  left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code"
        'Qry += " left outer join tspl_vendor_bank_Master on  tspl_vendor_bank_Master.bank_code=TSPL_VENDOR_MASTER.bank_code "
        'Qry += " left join  TSPL_Vendor_Bank_Branch_Details on TSPL_Vendor_Bank_Branch_Details.Bank_code=tspl_vendor_bank_Master.bank_code"
        'Qry += " left join TSPL_bank_master as Bank on Bank.bank_code=TSPL_PAYMENT_HEADER.Bank_Code "
        'Qry += " left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.comp_code=TSPL_PAYMENT_HEADER.comp_code "
        'Qry += " left join (select Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code left join TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code=TSPL_CUSTOM_FIELD_VALUES.Program_Code where TSPL_CUSTOM_FIELD_VALUES.Program_Code='PYMT-NEW'  and  Name ='Purpose')tt on tt.transaction_Code=TSPL_PAYMENT_HEADER.Payment_No WHERE TSPL_PAYMENT_HEADER.Payment_Code='RTGS' and TSPL_PAYMENT_HEADER.Payment_No ='" & txtPaymentNo.Value & "'"
        'Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(Qry)
        Dim Qry As String = "select TSPL_PAYMENT_HEADER.Entry_Desc, Value, TSPL_PAYMENT_HEADER.Payment_No as Doc_No,convert(varchar,Payment_Date,103)as Payment_Date,Payment_Amount,bank.BANKACCNUMBER as Debit_Account,TSPL_PAYMENT_HEADER.Vendor_Code,TSPL_VENDOR_MASTER.Cheque_In_Favour_Of as Vendor_Name,TSPL_VENDOR_MASTER.bank_code as Vendor_Bank_Code,TSPL_Vendor_Bank_Branch_Details.Bank_IFSC_Code as IFSC_Code,tspl_vendor_bank_Master.Bank_Name as Vendor_Bank_Name,TSPL_Vendor_Bank_Branch_Details.Branch_name as Vendor_Branch_Name,TSPL_PAYMENT_HEADER.Payment_Code,Account_No as Credit_Account,TSPL_PAYMENT_HEADER.Bank_Code,Bank.DESCRIPTION as Bank_Name,Bank.Add1,Bank.Add2,Bank.Add3,Bank.Add4 ,TSPL_COMPANY_MASTER.Comp_Name from TSPL_PAYMENT_HEADER left join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No  left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code"
        Qry += " left outer join tspl_vendor_bank_Master on  tspl_vendor_bank_Master.bank_code=TSPL_VENDOR_MASTER.bank_code "
        Qry += " left join  TSPL_Vendor_Bank_Branch_Details on TSPL_Vendor_Bank_Branch_Details.Bank_code=tspl_vendor_bank_Master.bank_code"
        Qry += " left join TSPL_bank_master as Bank on Bank.bank_code=TSPL_PAYMENT_HEADER.Bank_Code "
        Qry += " left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.comp_code=TSPL_PAYMENT_HEADER.comp_code "
        Qry += " left join (select Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code left join TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code=TSPL_CUSTOM_FIELD_VALUES.Program_Code where TSPL_CUSTOM_FIELD_VALUES.Program_Code='PYMT-NEW'  and  Name ='Purpose')tt on tt.transaction_Code=TSPL_PAYMENT_HEADER.Payment_No WHERE TSPL_PAYMENT_HEADER.Payment_Code='RTGS' and TSPL_PAYMENT_HEADER.Payment_No  = '" & txtPaymentNo.Value & "'"
        Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(Qry)
        If dt2.Rows.Count > 0 AndAlso dt2 IsNot Nothing Then
            Dim frmCRV As New frmCrystalReportViewer()
            frmCRV.funreport(MyBase.Form_ID, CrystalReportFolder.Purchase, dt2, EnumTecxpertPaperSize.NA, "PaymentRTGS", "Payment Details", False, dtpPayment.Value)
            frmCRV = Nothing
        End If
    End Sub

    Private Sub btnPaymentAdvice_Click(sender As Object, e As EventArgs) Handles btnPaymentAdvice.Click
        'Dim qry As String = " select TSPL_PAYMENT_HEADER.Payment_No,convert(varchar,TSPL_PAYMENT_HEADER.Payment_Date,103)as Payment_Date ,TSPL_PAYMENT_HEADER.Vendor_Code ,TSPL_VENDOR_MASTER.vendor_name,(TSPL_VENDOR_MASTER.Add1+TSPL_VENDOR_MASTER.Add2+TSPL_VENDOR_MASTER.Add3)as Vendor_Address,Comp_Name,Cheque_No ,convert(varchar,Cheque_Date,103)as Cheque_Date,TSPL_VENDOR_MASTER.Bank_Code, tspl_vendor_bank_Master.BAnk_Name,Payment_Amount, TSPL_PAYMENT_DETAIL.Document_No,Vendor_invoice_No,Original_Invoice_Amt,Net_Balance,TSPL_PAYMENT_DETAIL.TDS_Amount,Value,Net_Balance-Pending_Balance as adjAmt   from TSPL_PAYMENT_HEADER left join TSPL_PAYMENT_DETAIL on TSPL_PAYMENT_DETAIL.Payment_No =TSPL_PAYMENT_HEADER.Payment_No left join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code left join TSPL_STATE_MASTER on TSPL_STATE_MASTER.STATE_CODE =TSPL_VENDOR_MASTER.State_Code "
        'qry += " left join TSPL_COMPANY_MASTER on TSPL_COMPANY_MASTER.Comp_Code=TSPL_PAYMENT_HEADER.Comp_Code left join tspl_vendor_bank_Master on tspl_vendor_bank_Master.Bank_Code=TSPL_VENDOR_MASTER.Bank_Code"
        'qry += " where TSPL_PAYMENT_HEADER.Payment_No ='" & txtPaymentNo.Value & "' and TSPL_PAYMENT_HEADER.Payment_Type='PY'"

        'Dim qry As String = ""

        'qry += " Select TSPL_PAYMENT_HEADER.Payment_No, convert(varchar,TSPL_PAYMENT_HEADER.Payment_Date,103) as Payment_Date, TSPL_PAYMENT_HEADER.Vendor_Code, TSPL_PAYMENT_HEADER.Vendor_Name, TSPL_PAYMENT_DETAIL.Document_No as InvoiceNo, YYY.Vendor_Invoice_No, TSPL_PAYMENT_HEADER.Payment_Code as PaymentMode, YYY.InvoiceAmt, YYY.DrNoteAmt, CrNoteAmt, TSPL_PAYMENT_DETAIL.Applied_Amount as PaymentAmt, TSPL_PAYMENT_DETAIL.TDS_Amount, (YYY.InvoiceAmt-YYY.DrNoteAmt+CrNoteAmt-TSPL_PAYMENT_DETAIL.TDS_Amount-TSPL_PAYMENT_DETAIL.Applied_Amount) as [RemainingAmt], TSPL_PAYMENT_DETAIL.Comment,(TSPL_VENDOR_MASTER.Add1+TSPL_VENDOR_MASTER.Add2+TSPL_VENDOR_MASTER.Add3)as Vendor_Address,Comp_Name, Cheque_No, convert(varchar,Cheque_Date,103)as Cheque_Date, TSPL_VENDOR_MASTER.Bank_Code, tspl_vendor_bank_Master.BAnk_Name,case  when TSPL_PAYMENT_HEADER.Payment_Code='Cheque' then TSPL_PAYMENT_HEADER.Cheque_No +'  ' +'Dated' +' ' + convert(varchar,Cheque_Date,103) +' '+'Drawn on' when  TSPL_PAYMENT_HEADER.Payment_Code='DD' then value +' '+'Drawn on' else 'Drawn on' end as Payment  from TSPL_PAYMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No"
        'qry += " LEFT OUTER JOIN (Select InvoiceNo, MAX(Vendor_Invoice_No) as Vendor_Invoice_No, MAX(InvoiceAmt) as InvoiceAmt, SUM(DrNoteAmt) as DrNoteAmt, SUM(CrNoteAmt) as CrNoteAmt from ("
        'qry += " Select TSPL_VENDOR_INVOICE_HEAD.Document_No as InvoiceNo, TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No, TSPL_VENDOR_INVOICE_HEAD.Document_Total as InvoiceAmt, ISNULL(DRHEAD.Document_No,'') as DrNote, ISNULL(DRHEAD.Document_Total,0) as DrNoteAmt, ISNULL(CRHEAD.Document_No,'') as CrNote, ISNULL(CRHEAD.Document_Total,0) as CrNoteAmt from TSPL_VENDOR_INVOICE_HEAD"
        'qry += " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD DRHEAD ON DRHEAD.RefDocNo=TSPL_VENDOR_INVOICE_HEAD.Document_No AND DRHEAD.RefDocType='AP' AND DRHEAD.Document_Type='D'"
        'qry += " LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD CRHEAD ON CRHEAD.RefDocNo=TSPL_VENDOR_INVOICE_HEAD.Document_No AND CRHEAD.RefDocType='AP' AND CRHEAD.Document_Type='C'"
        'qry += " ) XXX GROUP BY InvoiceNo"

        'qry += " ) YYY ON YYY.InvoiceNo=TSPL_PAYMENT_DETAIL.Document_No"
        'qry += " LEFT OUTER JOIn TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=TSPL_PAYMENT_HEADER.Comp_Code"
        'qry += " left join TSPL_VENDOR_MASTEr on  TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code "
        'qry += " left join tspl_vendor_bank_Master on tspl_vendor_bank_Master.Bank_Code=TSPL_VENDOR_MASTER.Bank_Code "
        'qry += " left join (select Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code left join TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code=TSPL_CUSTOM_FIELD_VALUES.Program_Code where TSPL_CUSTOM_FIELD_VALUES.Program_Code='PYMT-NEW'  and  Name ='DD Number')tt"
        'qry += "   on tt.transaction_Code=TSPL_PAYMENT_HEADER.Payment_No"
        'qry += "  WHERE TSPL_PAYMENT_HEADER.Payment_No= '" & txtPaymentNo.Value & "' and TSPL_PAYMENT_HEADER.Payment_Type='PY'"
        '====changes by shivani against[BM00000008656]
        'qry += " Select TSPL_PAYMENT_HEADER.Bank_Code,TSpl_BANK_MASTER.DESCRIPTION as Bank_Name,case when TSPL_PAYMENT_HEADER.Payment_Type ='PY' then 'Your Inv. No' else 'Po No.' end as PO_No ,case when TSPL_PAYMENT_HEADER.Payment_Type ='PY' then 'Your Inv. Amt' else 'Adv Amt' end as Adv_Amt ,  TSPL_PAYMENT_HEADER.Payment_No, convert(varchar,TSPL_PAYMENT_HEADER.Payment_Date,103) as Payment_Date, TSPL_PAYMENT_HEADER.Vendor_Code, TSPL_PAYMENT_HEADER.Vendor_Name, TSPL_PAYMENT_DETAIL.Document_No as InvoiceNo, YYY.Vendor_Invoice_No, TSPL_PAYMENT_HEADER.Payment_Code as PaymentMode, YYY.InvoiceAmt, YYY.DrNoteAmt, CrNoteAmt,case when TSPL_PAYMENT_HEADER.Payment_Type ='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount else Balance_Amt end as PaymentAmt, TSPL_PAYMENT_DETAIL.TDS_Amount, (YYY.InvoiceAmt-YYY.DrNoteAmt+CrNoteAmt-TSPL_PAYMENT_DETAIL.TDS_Amount-TSPL_PAYMENT_DETAIL.Applied_Amount) as [RemainingAmt], TSPL_PAYMENT_DETAIL.Comment,(TSPL_VENDOR_MASTER.Add1+TSPL_VENDOR_MASTER.Add2+TSPL_VENDOR_MASTER.Add3)as Vendor_Address,Comp_Name, Cheque_No, convert(varchar,Cheque_Date,103)as Cheque_Date,case  when TSPL_PAYMENT_HEADER.Payment_Code='Cheque' then TSPL_PAYMENT_HEADER.Cheque_No +'  ' +'Dated' +' ' + convert(varchar,Cheque_Date,103) +' '+'Drawn on' when  TSPL_PAYMENT_HEADER.Payment_Code='DD' then value +' '+'Drawn on' else 'Drawn on' end as Payment  from TSPL_PAYMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No LEFT OUTER JOIN (Select InvoiceNo, MAX(Vendor_Invoice_No) as Vendor_Invoice_No, MAX(InvoiceAmt) as InvoiceAmt, SUM(DrNoteAmt) as DrNoteAmt, SUM(CrNoteAmt) as CrNoteAmt from ( Select TSPL_VENDOR_INVOICE_HEAD.Document_No as InvoiceNo, TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No, TSPL_VENDOR_INVOICE_HEAD.Document_Total as InvoiceAmt, ISNULL(DRHEAD.Document_No,'') as DrNote, ISNULL(DRHEAD.Document_Total,0) as DrNoteAmt, ISNULL(CRHEAD.Document_No,'') as CrNote, ISNULL(CRHEAD.Document_Total,0) as CrNoteAmt from TSPL_VENDOR_INVOICE_HEAD LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD DRHEAD ON DRHEAD.RefDocNo=TSPL_VENDOR_INVOICE_HEAD.Document_No AND DRHEAD.RefDocType='AP' AND DRHEAD.Document_Type='D' LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD CRHEAD ON CRHEAD.RefDocNo=TSPL_VENDOR_INVOICE_HEAD.Document_No AND CRHEAD.RefDocType='AP' AND CRHEAD.Document_Type='C' ) XXX GROUP BY InvoiceNo ) YYY ON YYY.InvoiceNo=TSPL_PAYMENT_DETAIL.Document_No LEFT OUTER JOIn TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=TSPL_PAYMENT_HEADER.Comp_Code left join TSpl_BANK_MASTER on TSpl_BANK_MASTER.BANK_CODE =TSPL_PAYMENT_HEADER.Bank_Code "
        'qry += " left join TSPL_VENDOR_MASTEr on  TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code  left join tspl_vendor_bank_Master on tspl_vendor_bank_Master.Bank_Code=TSPL_VENDOR_MASTER.Bank_Code  left join (select Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code left join TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code=TSPL_CUSTOM_FIELD_VALUES.Program_Code where TSPL_CUSTOM_FIELD_VALUES.Program_Code='PYMT-NEW'  and  Name ='DD Number')tt   on tt.transaction_Code=TSPL_PAYMENT_HEADER.Payment_No  WHERE TSPL_PAYMENT_HEADER.Payment_No = '" & txtPaymentNo.Value & "' and TSPL_PAYMENT_HEADER.Payment_Type in ('PY','OA','AV')"
        'richa UDL/31/03/19-000283
        'qry += " Select TSPL_PAYMENT_HEADER.Bank_Code,TSpl_BANK_MASTER.DESCRIPTION as Bank_Name,case when TSPL_PAYMENT_HEADER.Payment_Type ='PY' then 'Your Inv. No' else 'Po No.' end as PO_No ,case when TSPL_PAYMENT_HEADER.Payment_Type ='PY' then 'Your Inv. Amt' else 'Adv Amt' end as Adv_Amt ,  TSPL_PAYMENT_HEADER.Payment_No, convert(varchar,TSPL_PAYMENT_HEADER.Payment_Date,103) as Payment_Date, TSPL_PAYMENT_HEADER.Vendor_Code, TSPL_PAYMENT_HEADER.Vendor_Name, TSPL_PAYMENT_DETAIL.Document_No as InvoiceNo, YYY.Vendor_Invoice_No, TSPL_PAYMENT_HEADER.Payment_Code as PaymentMode, YYY.InvoiceAmt, YYY.DrNoteAmt, CrNoteAmt,case when TSPL_PAYMENT_HEADER.Payment_Type ='PY' then TSPL_PAYMENT_DETAIL.Applied_Amount else Balance_Amt end * (CASE WHEN YYY.Document_Type ='D' THEN -1 ELSE 1 END)  as PaymentAmt, TSPL_PAYMENT_DETAIL.TDS_Amount, (YYY.InvoiceAmt-YYY.DrNoteAmt+CrNoteAmt-TSPL_PAYMENT_DETAIL.TDS_Amount-TSPL_PAYMENT_DETAIL.Applied_Amount) as [RemainingAmt], TSPL_PAYMENT_DETAIL.Comment,(TSPL_VENDOR_MASTER.Add1+TSPL_VENDOR_MASTER.Add2+TSPL_VENDOR_MASTER.Add3)as Vendor_Address,Comp_Name, Cheque_No, convert(varchar,Cheque_Date,103)as Cheque_Date,case  when TSPL_PAYMENT_HEADER.Payment_Code='Cheque' then TSPL_PAYMENT_HEADER.Cheque_No +'  ' +'Dated' +' ' + convert(varchar,Cheque_Date,103) +' '+'Drawn on' when  TSPL_PAYMENT_HEADER.Payment_Code='DD' then value +' '+'Drawn on' else 'Drawn on' end as Payment  from TSPL_PAYMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_DETAIL ON TSPL_PAYMENT_DETAIL.Payment_No=TSPL_PAYMENT_HEADER.Payment_No LEFT OUTER JOIN (Select MAX(Document_Type) as Document_Type ,InvoiceNo, MAX(Vendor_Invoice_No) as Vendor_Invoice_No, MAX(InvoiceAmt) as InvoiceAmt, SUM(DrNoteAmt) as DrNoteAmt, SUM(CrNoteAmt) as CrNoteAmt from ( Select TSPL_VENDOR_INVOICE_HEAD.Document_Type ,TSPL_VENDOR_INVOICE_HEAD.Document_No as InvoiceNo, TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No, TSPL_VENDOR_INVOICE_HEAD.Document_Total as InvoiceAmt, ISNULL(DRHEAD.Document_No,'') as DrNote, case when  TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' then ISNULL(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0) ELSE 0 end as DrNoteAmt, ISNULL(CRHEAD.Document_No,'') as CrNote, case when  TSPL_VENDOR_INVOICE_HEAD.Document_Type<>'D' then ISNULL(TSPL_VENDOR_INVOICE_HEAD.Document_Total,0)  ELSE 0  end as CrNoteAmt from TSPL_VENDOR_INVOICE_HEAD LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD DRHEAD ON DRHEAD.RefDocNo=TSPL_VENDOR_INVOICE_HEAD.Document_No AND DRHEAD.RefDocType='AP' AND DRHEAD.Document_Type='D' LEFT OUTER JOIN TSPL_VENDOR_INVOICE_HEAD CRHEAD ON CRHEAD.RefDocNo=TSPL_VENDOR_INVOICE_HEAD.Document_No AND CRHEAD.RefDocType='AP' AND CRHEAD.Document_Type='C' ) XXX GROUP BY InvoiceNo ) YYY ON YYY.InvoiceNo=TSPL_PAYMENT_DETAIL.Document_No LEFT OUTER JOIn TSPL_COMPANY_MASTER ON TSPL_COMPANY_MASTER.Comp_Code=TSPL_PAYMENT_HEADER.Comp_Code left join TSpl_BANK_MASTER on TSpl_BANK_MASTER.BANK_CODE =TSPL_PAYMENT_HEADER.Bank_Code "
        'qry += " left join TSPL_VENDOR_MASTEr on  TSPL_VENDOR_MASTER.Vendor_Code=TSPL_PAYMENT_HEADER.Vendor_Code  left join tspl_vendor_bank_Master on tspl_vendor_bank_Master.Bank_Code=TSPL_VENDOR_MASTER.Bank_Code  left join (select Transaction_code,Value from TSPL_CUSTOM_FIELD_VALUES left join TSPL_CUSTOM_FIELD_HEAD on TSPL_CUSTOM_FIELD_HEAD.Code=TSPL_CUSTOM_FIELD_VALUES.Custom_Field_Code left join TSPL_PROGRAM_MASTER on TSPL_PROGRAM_MASTER.Program_Code=TSPL_CUSTOM_FIELD_VALUES.Program_Code where TSPL_CUSTOM_FIELD_VALUES.Program_Code='PYMT-NEW'  and  Name ='DD Number')tt   on tt.transaction_Code=TSPL_PAYMENT_HEADER.Payment_No  WHERE TSPL_PAYMENT_HEADER.Payment_No = '" & txtPaymentNo.Value & "' and TSPL_PAYMENT_HEADER.Payment_Type in ('PY','OA','AV')"

        'Dim dt2 As DataTable = clsDBFuncationality.GetDataTable(qry)
        'If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
        '    Dim frmCRV As New frmCrystalReportViewer()
        '    frmCRV.funreport(CrystalReportFolder.Purchase, dt2, EnumTecxpertPaperSize.NA, "crptPaymentAdvice", "Payment Advice", clsCommon.myCDate(dt2.Rows(0)("Payment_Date")))
        '    frmCRV = Nothing
        'Else
        '    clsCommon.MyMessageBoxShow("Payment Advice Not Available this mode.", Me.Text)
        'End If
        clsPaymentHeader.PrintPaymentAdvice(txtPaymentNo.Value, Nothing, False)

    End Sub

    Private Sub fndloanNo__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles fndloanNo._MYValidating

        Dim whrcls As String = "Loan_Status='Open'"
        fndloanNo.Value = clsApplyLoan.GetFinder(whrcls, fndloanNo.Value, isButtonClicked)
        txtDescription.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select LOAN_DESCRIPTION  from TSPL_LOAN_APPLICATION where Loan_Code='" & fndloanNo.Value & "'"))
        txtVendorCode.Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select emp_Code from TSPL_LOAN_APPLICATION where Loan_Code='" & fndloanNo.Value & "'"))
        lblVendorName.Text = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Emp_Name  from TSPL_EMPLOYEE_MASTER where EMP_CODE ='" & txtVendorCode.Value & "'"))
        txtPaymentAmt.Text = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select loan_amount from TSPL_LOAN_APPLICATION where Loan_Code='" & fndloanNo.Value & "'"))
        txtVendorCode.MyReadOnly = True
        txtPaymentAmt.ReadOnly = True
    End Sub

    Private Sub ChkAdvSalary_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles ChkAdvSalary.ToggleStateChanged
        If ChkAdvSalary.Checked Then
            fndloanNo.Visible = True
            lblApplyLoanNo.Visible = True
        Else
            fndloanNo.Visible = False
            lblApplyLoanNo.Visible = False

        End If
    End Sub

    Private Sub rbtnExportPosted_Click(sender As Object, e As EventArgs) Handles rbtnExportPosted.Click
        Try
            Dim qryExport As String
            qryExport = " select Payment_No as [Payment No],convert(varchar,Payment_Date,103) as [Payment Date],Entry_Desc as [Description],Vendor_Code as [Vendor Code],Bank_Code as [Bank Code]," &
                        " Payment_Type as [Payment Type(A/O/R)],Payment_Code as [Payment Mode],Cheque_No as [Cheque No],Cheque_Date as [Cheque Date],Payment_Amount as Amount, " &
                        " coalesce(Location_Code,Location_GL_Code) as [Location Code],Advance_Against_Salary as [Advance Against Salary],is_Opening as [Is Opening], " &
                        " Bank_Charges as [Bank Charges],Is_Security as [Is Security],ConvRate as [Conv Rate] from TSPL_PAYMENT_HEADER "
            Dim Cond As String = " and  Payment_Type in ('AD','OA','RC') and Posted=1 "
            transportSql.ExporttoExcel(qryExport, Cond, "[Payment Date]", Me)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Payment Entry")
        End Try
    End Sub

    Private Sub rbtnImportPosted_Click(sender As Object, e As EventArgs) Handles rbtnImportPosted.Click
        '' done by panch raj against ticket No: 
        Dim gv As New UserControls.MyRadGridView
        Me.Controls.Add(gv)
        Dim Count As String = ""
        If transportSql.importExcel(gv, "Payment No", "Payment Date", "Description", "Vendor Code", "Bank Code", "Payment Type(A/O/R)", "Payment Mode", "Cheque No", "Cheque Date", "Amount", "Location Code", "Advance Against Salary", "Is Opening", "Bank Charges", "Is Security", "Conv Rate") Then
            Dim trans As SqlTransaction = Nothing
            Dim linno As Integer = 0
            Try
                connectSql.OpenConnection()
                trans = clsDBFuncationality.GetTransactin()
                clsCommon.ProgressBarPercentShow()
                Dim obj As clsPaymentHeader
                For Each grow As GridViewRowInfo In gv.Rows
                    Count = clsCommon.myCstr(grow.Index + 2)
                    If clsCommon.myLen(grow.Cells("Payment No").Value) > 0 Then
                        obj = New clsPaymentHeader()
                        obj = clsPaymentHeader.GetData(grow.Cells("Payment No").Value, NavigatorType.Current, trans)
                        If Not (clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "AD") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal) Then
                            Throw New Exception("Payment Type must be (OA,AD,RC) for Import after post")
                        End If
                        obj.Entry_Desc = clsCommon.myCstr(grow.Cells("Description").Value)
                        If clsCommon.myLen(obj.Entry_Desc) > 250 Then
                            Throw New Exception("Description Length can not be more than 250.")
                        End If

                        obj.Payment_Date = clsCommon.myCDate(grow.Cells("Payment Date").Value)
                        obj.Payment_Post_Date = obj.Payment_Date
                        obj.Vendor_Code = clsCommon.myCstr(grow.Cells("Vendor Code").Value)
                        If clsCommon.myLen(obj.Vendor_Code) > 0 Then
                            obj.Vendor_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Code from TSPL_VENDOR_MASTER WHERE Vendor_Code='" + obj.Vendor_Code + "'", trans))
                            If clsCommon.myLen(obj.Vendor_Code) <= 0 Then
                                Throw New Exception("Vendor Code does not exist.")
                            End If
                            obj.Vendor_Name = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Vendor_Name from TSPL_VENDOR_MASTER WHERE Vendor_Code='" + obj.Vendor_Code + "'", trans))
                        Else
                            Throw New Exception("Please select Vendor Code.")
                        End If
                        obj.Bank_Code = clsCommon.myCstr(grow.Cells("Bank Code").Value)
                        If clsCommon.myLen(obj.Bank_Code) > 0 Then
                            obj.Bank_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select BANK_CODE from TSPL_BANK_MASTER WHERE Bank_Code='" + obj.Bank_Code + "'", trans))
                            If clsCommon.myLen(obj.Bank_Code) <= 0 Then
                                Throw New Exception("Bank Code does not exist.")
                            End If
                        Else
                            Throw New Exception("Please select Bank Code.")
                        End If

                        obj.Payment_Type = clsCommon.myCstr(grow.Cells("Payment Type(A/O/R)").Value)
                        If clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal Then
                            obj.Payment_Type = "OA"
                        ElseIf clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Then
                            obj.Payment_Type = "AV"
                        ElseIf clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                            obj.Payment_Type = "RC"
                        Else
                            Throw New Exception("Payment type can be 'OA' or 'AV' or 'RC'.")
                        End If

                        obj.Payment_Code = clsCommon.myCstr(grow.Cells("Payment Mode").Value)
                        If clsCommon.myLen(obj.Payment_Code) > 0 Then
                            obj.Payment_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select TSPL_PAYMENT_CODE.Payment_Code from TSPL_PAYMENT_CODE Where TSPL_PAYMENT_CODE.Payment_Code='" + obj.Payment_Code + "'", trans))
                            If clsCommon.myLen(obj.Payment_Code) <= 0 Then
                                Throw New Exception("Payment Mode does not exist.")
                            End If
                        Else
                            Throw New Exception("Enter Payment Mode.")
                        End If


                        If clsCommon.CompairString(obj.Payment_Code, "Cheque") = CompairStringResult.Equal Then
                            obj.Cheque_No = clsCommon.myCstr(grow.Cells("Cheque No").Value)
                            If clsCommon.myLen(obj.Cheque_No) > 0 Then
                                If clsCommon.myLen(obj.Cheque_No) < 6 Or clsCommon.myLen(obj.Cheque_No) > 20 Then
                                    Throw New Exception("Length of Cheque No should between 6-20.")
                                End If
                                Qry = "Select Payment_No from TSPL_PAYMENT_HEADER Where Cheque_No='" & obj.Cheque_No & "' and Payment_No not in ('" & obj.Payment_No & "')"
                                Qry = clsCommon.myCstr(clsDBFuncationality.getSingleValue(Qry, trans))
                                If clsCommon.myLen(Qry) > 0 Then
                                    Throw New Exception("Cheque '" & obj.Cheque_No & "' is Already Exists Against Payment '" & Qry & "'")
                                End If
                                '----------------------Cheque Date---------------------
                                If clsCommon.myLen(grow.Cells("Cheque Date").Value) > 0 Then
                                    obj.Cheque_Date = clsCommon.GetPrintDate(obj.Cheque_Date, "dd/MMM/yyyy") ' CDate(grow.Cells("Cheque Date").Value)
                                Else
                                    Throw New Exception("Please enter Cheque Date.")
                                End If
                                '------------------------------------------------------
                            Else
                                Throw New Exception("Cheque No can't be blank")
                            End If
                        End If

                        If clsCommon.myCdbl(grow.Cells("Amount").Value) <= 0 Then
                            Throw New Exception("Enter Payment Amount.")
                        End If

                        obj.Account_Payee = 0
                        obj.Location_GL_Code = ""

                        obj.memorndmamt = 0.0

                        obj.CHECK_PRINT = 0

                        ''richa agarwal 09/07/2015
                        Dim dblAdvanceAgainstSalary As Double = clsCommon.myCdbl(grow.Cells("Advance Against Salary").Value)
                        If dblAdvanceAgainstSalary = 0 Or dblAdvanceAgainstSalary = 1 Then
                            obj.Advance_Against_Salary = dblAdvanceAgainstSalary
                        Else
                            Throw New Exception("Advance Against Salary should be 0 or 1.")
                        End If
                        Dim dblIsOpening As Double = clsCommon.myCdbl(grow.Cells("Is Opening").Value)
                        If dblIsOpening = 0 Or dblIsOpening = 1 Then
                            obj.is_Opening = dblIsOpening
                        Else
                            Throw New Exception("Is Opening Advance Against Salary should be 0 or 1.")
                        End If
                        ''--------------------

                        '-----------Location Code--------------------------------
                        Dim Loc_Code As String = clsCommon.myCstr(grow.Cells("Location Code").Value)
                        If clsCommon.myLen(Loc_Code) > 0 Then
                            Loc_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select distinct(Segment_code) as Code ,Description  from TSPL_GL_SEGMENT_CODE left " _
                                       & " outer join TSPL_LOCATION_MASTER on TSPL_GL_SEGMENT_CODE .Segment_code =TSPL_LOCATION_MASTER .Loc_Segment_Code  Where Seg_No = '7' AND GIT='N' and Segment_code = '" & Loc_Code & "'", trans))
                            If Loc_Code = "" Then
                                Throw New Exception("Please Check Location Code dose not Exits") ' + LineNo + " does not exist. ")
                            End If
                        Else
                            'Throw New Exception("Insert Location Code in All Rows ") ' + LineNo + ".")
                        End If
                        '--------------------------------------------------------

                        If clsCommon.CompairString(obj.Payment_Type, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "OA") = CompairStringResult.Equal Or clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                            obj.Total_Prepayment = clsCommon.myCdbl(grow.Cells("Amount").Value)
                            obj.Payment_Amount = clsCommon.myCdbl(grow.Cells("Amount").Value)
                            obj.Balance_Amt = clsCommon.myCdbl(grow.Cells("Amount").Value)
                            obj.TDS_Amount = 0
                        End If

                        If clsCommon.CompairString(obj.Payment_Type, "RC") = CompairStringResult.Equal Then
                            obj.Is_Security = clsCommon.myCdbl(grow.Cells("Is Security").Value)
                            obj.Is_Retention = clsCommon.myCdbl(grow.Cells("Is Retention").Value)
                        Else
                            obj.Is_Security = 0
                            obj.Is_Retention = 0
                        End If
                        obj.IsChkReverse = "N"
                        obj.Bank_Charges = clsCommon.myCdbl(grow.Cells("Bank Charges").Value)
                        obj.objRemittance = Nothing
                        obj.PurchaseOrder_No = ""
                        obj.Loan_Code = ""
                        obj.Account_Payee_Name = ""
                        obj.CURRENCY_CODE = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select currency_code from TSPL_VENDOR_MASTER where VENDOR_CODE='" & obj.Vendor_Code & "'", trans))
                        obj.BASE_CURRENCY_CODE = objCommonVar.BaseCurrencyCode
                        obj.PAYMENT_AMOUNT_BASE_CURRENCY = obj.Payment_Amount
                        obj.ConvRateOld = 1
                        obj.ConvRate = IIf(clsCommon.myCdbl(grow.Cells("Conv Rate").Value) <= 0, 1, clsCommon.myCdbl(grow.Cells("Conv Rate").Value))
                        If clsCommon.CompairString(obj.CURRENCY_CODE, obj.BASE_CURRENCY_CODE) = CompairStringResult.Equal Then
                            obj.ConvRate = 1
                        End If
                        obj.PAYMENT_AMOUNT_BASE_CURRENCY = obj.Payment_Amount * obj.ConvRate
                        If clsCommon.myLen(grow.Cells("Location Code").Value) <= 0 Then
                            obj.Location_Code = clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select RIGHT(BANKACC,3) from TSPL_Bank_Master Where Bank_Code='" + obj.Bank_Code + "'", trans))
                        Else
                            obj.Location_Code = clsCommon.myCstr(grow.Cells("Location Code").Value)
                        End If

                        obj.Location_GL_Code = obj.Location_Code
                        obj.PDC_Cheque = "N'"
                        obj.Applied_Payment = ""
                        obj.ArrTr = Nothing
                        'obj.Location_GL_Code = Loc_Code
                        obj.UpdatePostedData(obj, trans)
                    End If
                Next
                trans.Commit()
                clsCommon.ProgressBarPercentHide()
                common.clsCommon.MyMessageBoxShow("Data Transfer Completed!", Me.Text, MessageBoxButtons.OK)
            Catch ex As Exception
                trans.Rollback()
                clsCommon.ProgressBarPercentHide()
                clsCommon.MyMessageBoxShow("Line: " + Count + " - " + ex.Message)
            Finally
                Me.Controls.Remove(gv)
            End Try
        End If
    End Sub

    Private Sub gvDetails_CellFormatting(sender As Object, e As CellFormattingEventArgs) Handles gvDetails.CellFormatting
        Try
            If e.Column.Index >= 0 Then
                If clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "MI") = CompairStringResult.Equal Then
                    If clsCommon.myLen(gvDetails.CurrentRow.Cells(colGLAccount).Value) > 0 Then
                        Dim grouptype As String = ""
                        grouptype = clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLType).Value)
                        If clsCommon.CompairString(grouptype, "Balance Sheet") = CompairStringResult.Equal Then
                            gvDetails.CurrentRow.Cells(colHirerachyCenter).ReadOnly = True
                            gvDetails.CurrentRow.Cells(colCostCenter).ReadOnly = True
                        Else
                            gvDetails.CurrentRow.Cells(colHirerachyCenter).ReadOnly = False
                            gvDetails.CurrentRow.Cells(colCostCenter).ReadOnly = False
                        End If
                    End If
                    If e.Column Is gvDetails.Columns(colGLAccount) Then
                        gvDetails.CurrentRow.Cells(colGLAccount).ReadOnly = (clsCommon.myLen(txtMPAdv.Value) > 0)
                    End If
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtPONo_GST__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtPONo_GST._MYValidating
        Try
            GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
            If GSTStatus AndAlso clsCommon.CompairString(clsCommon.myCstr(ddlPaymentType.SelectedValue), "AV") = CompairStringResult.Equal Then
                Dim whrcls As String = String.Empty
                Dim gstdate As String = clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterType.GSTApplicableDate, clsFixedParameterCode.GSTApplicableDate, Nothing))




                Qry = " Select Final.* from (Select zz.Document_Code,max(zz.Vendor_Code )as Vendor_Code ,sum(isnull(zz.PO_Total_Amt ,0)) as PO_Total_Amt,sum(isnull(zz.Paymentamt ,0)) as Paymentamt from ( " & Environment.NewLine &
              " sELECT TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No AS Document_Code  ,TSPL_PURCHASE_ORDER_HEAD.Vendor_Code  ,TSPL_PURCHASE_ORDER_HEAD.PO_Total_Amt-ISNULL(TSPL_PURCHASE_ORDER_HEAD.Total_Tax_Amt,0)  as PO_Total_Amt,0 as Paymentamt  " & Environment.NewLine &
              " FROM TSPL_PURCHASE_ORDER_HEAD left outer join TSPL_SRN_DETAIL on TSPL_SRN_DETAIL.PO_Id =TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No" & Environment.NewLine &
              " left outer join tspl_vendor_master on tspl_vendor_master.vendor_code= TSPL_PURCHASE_ORDER_HEAD.Vendor_Code " + Environment.NewLine +
              " where TSPL_PURCHASE_ORDER_HEAD.Status=1 and  isnull(TSPL_PURCHASE_ORDER_HEAD.close_yn ,'N')='N' " & Environment.NewLine &
              " and isnull(TSPL_SRN_DETAIL.PO_Id,'') <> TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No AND " + IIf(objCommonVar.TreatUnregisteredVendorAsRegisteredVendor, " TSPL_VENDOR_MASTER.GSTRegistered = 0 ", " TSPL_PURCHASE_ORDER_HEAD.GSTRegistered=0 ") + "  " & Environment.NewLine &
              " AND TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_Date  >=convert(date,'" & gstdate & "',103) " & Environment.NewLine &
              " UNION ALL " & Environment.NewLine &
              " sELECT PurchaseOrder_No_GST as Document_Code,Vendor_Code  as Customer_Code ,0 as PO_Total_Amt,Payment_Amount+TDS_Amount  as Paymentamt  FROM TSPL_PAYMENT_HEADER where isnull(TSPL_PAYMENT_HEADER. PurchaseOrder_No_GST,'')<>'' and Payment_Type  ='AV' and TSPL_PAYMENT_HEADER.Payment_No <>'" & clsCommon.myCstr(txtPaymentNo.Value) & "' and isnull(TSPL_PAYMENT_HEADER. IsChkReverse,'') ='N' ) zz " & Environment.NewLine &
              " group by Document_Code) Final"



                whrcls = " (Final.PO_Total_Amt - Final.Paymentamt) > 0 "
                If clsCommon.myLen(clsCommon.myCstr(txtVendorCode.Value)) > 0 Then
                    whrcls += " and Final.Vendor_Code ='" & clsCommon.myCstr(txtVendorCode.Value) & "' "
                End If


                txtPONo_GST.Value = clsCommon.ShowSelectForm("PurchaseOrderFinder_GSt", Qry, "Document_Code", whrcls, clsCommon.myCstr(txtPONo_GST.Value), "Document_Code", isButtonClicked)
                If clsCommon.myLen(txtPONo_GST.Value) > 0 Then
                    FillPOofProductItems()
                Else
                    txtPONo_GST.Value = ""
                    txtTaxGroup.Value = ""
                    lblTaxGrpName.Text = ""
                    LblPOTotalAmount.Text = 0
                    lblPOTotalAdditionalCharge.Text = 0
                    lblPOTotalTaxAmt.Text = 0
                    LBLPO_Location_GST.Text = ""
                    TxtPO_Location_GST.Value = ""
                    LoadBlankGridPOItemDetail()
                    LoadBlankGridTax()
                End If
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Sub FillPOofProductItems()
        LoadBlankGridPOItemDetail()
        LoadBlankGridTax()

        Dim DblReceiptTaxAmt1 As Double = 0
        Dim DblReceiptTaxAmt2 As Double = 0
        Dim DblReceiptTaxAm3 As Double = 0
        Dim DblReceiptTaxAmt4 As Double = 0
        Dim DblReceiptTaxAmt5 As Double = 0
        Dim DblReceiptTaxAmt6 As Double = 0
        Dim DblReceiptTaxAmt7 As Double = 0
        Dim DblReceiptTaxAmt8 As Double = 0
        Dim DblReceiptTaxAmt9 As Double = 0
        Dim DblReceiptTaxAmt10 As Double = 0



        If clsCommon.myLen(txtPONo_GST.Value) > 0 Then
            Qry = " Select Final.* from (Select zz.Document_Code,max(zz.Vendor_Code )as Vendor_Code ,sum(isnull(zz.PO_Total_Amt ,0)) as PO_Total_Amt,sum(isnull(zz.Paymentamt ,0)) as Paymentamt, max(zz.Tax_Group )as Tax_Group,max(zz.Tax_Group_Desc )as Tax_Group_Desc,sum(AddCharge) as AddCharge,MAX(PO_Location_Code) AS PO_Location_Code from ( " & Environment.NewLine &
          " sELECT TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No AS Document_Code  ,TSPL_PURCHASE_ORDER_HEAD.Vendor_Code  ,TSPL_PURCHASE_ORDER_HEAD.PO_Total_Amt-ISNULL(TSPL_PURCHASE_ORDER_HEAD .Total_Tax_Amt,0)  as PO_Total_Amt,0 as Paymentamt,TSPL_PURCHASE_ORDER_HEAD.Tax_Group,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc, TSPL_PURCHASE_ORDER_HEAD.Total_Add_Charge as AddCharge,TSPL_PURCHASE_ORDER_HEAD.Bill_To_Location AS PO_Location_Code " & Environment.NewLine &
          " FROM TSPL_PURCHASE_ORDER_HEAD left outer join TSPL_SRN_DETAIL on TSPL_SRN_DETAIL.PO_Id =TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.vendor_code= TSPL_PURCHASE_ORDER_HEAD.Vendor_Code  " & Environment.NewLine &
          " Left Outer Join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code =TSPL_PURCHASE_ORDER_HEAD.Tax_Group and isnull(TSPL_TAX_GROUP_MASTER.Tax_Group_Type,'')='P' " & Environment.NewLine &
          " where TSPL_PURCHASE_ORDER_HEAD.Status=1 and  isnull(TSPL_PURCHASE_ORDER_HEAD.close_yn ,'N')='N' " & Environment.NewLine &
          " and   isnull(TSPL_SRN_DETAIL.PO_Id,'') <>TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No " & Environment.NewLine &
          " AND TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No  ='" & clsCommon.myCstr(txtPONo_GST.Value) & "'  AND  " + IIf(objCommonVar.TreatUnregisteredVendorAsRegisteredVendor, " TSPL_VENDOR_MASTER.GSTRegistered = 0 ", "TSPL_PURCHASE_ORDER_HEAD.GSTRegistered=0") + "   " & Environment.NewLine &
          " UNION ALL " & Environment.NewLine &
          " sELECT PurchaseOrder_No_GST as Document_Code,Vendor_Code  as Customer_Code ,0 as PO_Total_Amt,Payment_Amount+TDS_Amount  as Paymentamt,TSPL_PAYMENT_HEADER.Tax_Group,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc, 0 as AddCharge, ISNULL(TSPL_PAYMENT_HEADER.PO_Location_Code,'') AS PO_Location_Code  FROM TSPL_PAYMENT_HEADER " &
          " Left Outer Join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code =TSPL_PAYMENT_HEADER.Tax_Group and isnull(TSPL_TAX_GROUP_MASTER.Tax_Group_Type,'')='P' " & Environment.NewLine &
          " where isnull(TSPL_PAYMENT_HEADER. PurchaseOrder_No_GST,'')='" & clsCommon.myCstr(txtPONo_GST.Value) & "'  and Payment_Type  ='AV' and TSPL_PAYMENT_HEADER.Payment_No <>'" & clsCommon.myCstr(txtPaymentNo.Value) & "' and isnull(TSPL_PAYMENT_HEADER. IsChkReverse,'') ='N' ) zz " & Environment.NewLine &
          " group by Document_Code) Final where  (Final.PO_Total_Amt - Final.Paymentamt)>0 "


            Dim dt As DataTable = clsDBFuncationality.GetDataTable(Qry)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                txtTaxGroup.Value = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
                lblTaxGrpName.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
                LblPOTotalAmount.Text = clsCommon.myCdbl(dt.Rows(0)("PO_Total_Amt")) - clsCommon.myCdbl(dt.Rows(0)("Paymentamt"))
                lblPOTotalAdditionalCharge.Text = clsCommon.myCdbl(dt.Rows(0)("AddCharge"))
                TxtPO_Location_GST.Value = clsCommon.myCstr(dt.Rows(0)("PO_Location_Code"))
                LBLPO_Location_GST.Text = clsLocation.GetName(clsCommon.myCstr(dt.Rows(0)("PO_Location_Code")), Nothing)
            End If

            If clsCommon.myLen(txtTaxGroup.Value) > 0 Then
                Qry = "select PurchaseOrder_No ,Line_No ,Item_Code ,Item_Desc ,PurchaseOrder_Qty ,Unit_code ,Balance_Qty,Item_Cost ,TAX1 ,TAX1_Amt ,TAX1_Base_Amt ,TAX1_Rate , " &
          " tax2,TAX2_Base_Amt,TAX2_Rate,TAX2_Amt,TAX3,TAX3_Base_Amt,TAX3_Rate,TAX3_Amt,TAX4,TAX4_Base_Amt,TAX4_Rate,TAX4_Amt,TAX5,TAX5_Base_Amt,TAX5_Rate," &
          " TAX5_Amt,TAX6,TAX6_Base_Amt,TAX6_Rate,TAX6_Amt,TAX7,TAX7_Base_Amt,TAX7_Rate,TAX7_Amt,TAX8," &
          " TAX8_Base_Amt,TAX8_Rate,TAX8_Amt,TAX9,TAX9_Base_Amt,TAX9_Rate,TAX9_Amt,TAX10,TAX10_Base_Amt,TAX10_Rate,TAX10_Amt," &
          " Amount ,Disc_Per ,Disc_Amt ,Amt_Less_Discount ,Total_Tax_Amt ,Item_Net_Amt ,Row_Type ,AbatementRate " &
          " from TSPL_PURCHASE_ORDER_detail where PurchaseOrder_No ='" & clsCommon.myCstr(txtPONo_GST.Value) & "'"

                dt = clsDBFuncationality.GetDataTable(Qry)

                For i As Integer = 0 To dt.Rows.Count - 1
                    gvItem.Rows.AddNew()
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolDocument_Code).Value = clsCommon.myCstr(dt.Rows(i)("PurchaseOrder_No"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolLine_No).Value = clsCommon.myCdbl(dt.Rows(i)("Line_No"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolRow_Type).Value = clsCommon.myCstr(dt.Rows(i)("Row_Type"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolItem_Code).Value = clsCommon.myCstr(dt.Rows(i)("Item_Code"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(adcolIName).Value = clsCommon.myCstr(dt.Rows(i)("Item_Desc"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(adcolIHSN).Value = clsItemMaster.GetItemHSNCode(clsCommon.myCstr(dt.Rows(i)("Item_Code")), Nothing)
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolQty).Value = clsCommon.myCdbl(dt.Rows(i)("PurchaseOrder_Qty"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolBalance_Qty).Value = clsCommon.myCdbl(dt.Rows(i)("Balance_Qty"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolUnit_code).Value = clsCommon.myCstr(dt.Rows(i)("Unit_code"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolItem_Cost).Value = clsCommon.myCdbl(dt.Rows(i)("Item_Cost"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX1).Value = clsCommon.myCstr(dt.Rows(i)("TAX1"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX1_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX1_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX1_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX1_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX1_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX1_Rate"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(Adcoltax2).Value = clsCommon.myCstr(dt.Rows(i)("tax2"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX2_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX2_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX2_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX2_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX2_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX2_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX3).Value = clsCommon.myCstr(dt.Rows(i)("TAX3"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX3_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX3_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX3_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX3_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX3_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX3_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX4).Value = clsCommon.myCstr(dt.Rows(i)("TAX4"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX4_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX4_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX4_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX4_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX4_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX4_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX5).Value = clsCommon.myCstr(dt.Rows(i)("TAX5"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX5_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX5_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX5_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX5_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX5_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX5_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX6).Value = clsCommon.myCstr(dt.Rows(i)("TAX6"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX6_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX6_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX6_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX6_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX6_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX6_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX7).Value = clsCommon.myCstr(dt.Rows(i)("TAX7"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX7_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX7_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX7_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX7_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX7_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX7_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX8).Value = clsCommon.myCstr(dt.Rows(i)("TAX8"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX8_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX8_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX8_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX8_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX8_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX8_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX9).Value = clsCommon.myCstr(dt.Rows(i)("TAX9"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX9_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX9_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX9_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX9_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX9_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX9_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX10).Value = clsCommon.myCstr(dt.Rows(i)("TAX10"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX10_Base_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX10_Base_Amt"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX10_Rate).Value = clsCommon.myCdbl(dt.Rows(i)("TAX10_Rate"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTAX10_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("TAX10_Amt"))

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolAmount).Value = clsCommon.myCdbl(dt.Rows(i)("Amount"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolDisc_Per).Value = clsCommon.myCdbl(dt.Rows(i)("Disc_Per"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolDisc_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("Disc_Amt"))

                    Qry = " Select sum(Amt_Less_Discount) as Amt_Less_Discount  from " & Environment.NewLine &
                  "(select PurchaseOrder_No AS Document_Code ,Line_No ,Row_Type ,Item_Code ,Amt_Less_Discount  from TSPL_PURCHASE_ORDER_detail where PurchaseOrder_No ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "' " & Environment.NewLine &
                  " union all" & Environment.NewLine &
                  " select PurchaseOrder_No_GST AS Document_Code ,Line_No ,Row_Type ,Item_Code , PaymentAdvance * -1 as Amt_Less_Discount  from TSPL_PAYMENT_DETAIL_GST Left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No =TSPL_PAYMENT_DETAIL_GST.Payment_No  where isnull(TSPL_PAYMENT_HEADER.PurchaseOrder_No_GST,'')  ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "'  and  TSPL_PAYMENT_HEADER.Payment_No<>'" & clsCommon.myCstr(txtPaymentNo.Value) & "' and isnull(TSPL_PAYMENT_HEADER. IsChkReverse,'') ='N' )zz " & Environment.NewLine &
                  "group by zz.Document_Code ,zz.Item_Code,Line_No  order by Line_No "

                    ' gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolAmt_Less_Discount).Value = clsCommon.myCdbl(dt.Rows(i)("Amt_Less_Discount"))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolAmt_Less_Discount).Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))
                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolTotal_Tax_Amt).Value = clsCommon.myCdbl(dt.Rows(i)("Total_Tax_Amt"))

                    Qry = " Select sum(Item_Net_Amt) as Item_Net_Amt  from " & Environment.NewLine &
                    "(select PurchaseOrder_No AS Document_Code ,Line_No ,Row_Type ,Item_Code ,Amt_Less_Discount+isnull(Total_ItemAdd_Charge,0) AS Item_Net_Amt  from TSPL_PURCHASE_ORDER_detail where PurchaseOrder_No ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "' " & Environment.NewLine &
                    " union all" & Environment.NewLine &
                    " select PurchaseOrder_No_GST AS Document_Code ,Line_No ,Row_Type ,Item_Code , PaymentTotalAdvanceAmt * -1  from TSPL_PAYMENT_DETAIL_GST Left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No =TSPL_PAYMENT_DETAIL_GST.Payment_No  where isnull(TSPL_PAYMENT_HEADER.PurchaseOrder_No_GST,'')  ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "'  and  TSPL_PAYMENT_HEADER.Payment_No<>'" & clsCommon.myCstr(txtPaymentNo.Value) & "' and isnull(TSPL_PAYMENT_HEADER. IsChkReverse,'') ='N' )zz " & Environment.NewLine &
                    "group by zz.Document_Code ,zz.Item_Code,Line_No  order by Line_No "

                    '   Qry = " Select sum(Item_Net_Amt) as Item_Net_Amt  from " & Environment.NewLine & _
                    '"(select PurchaseOrder_No AS Document_Code ,Line_No ,Row_Type ,Item_Code ,Item_Net_Amt AS Item_Net_Amt  from TSPL_PURCHASE_ORDER_detail where PurchaseOrder_No ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "' " & Environment.NewLine & _
                    '" union all" & Environment.NewLine & _
                    '" select PurchaseOrder_No_GST AS Document_Code ,Line_No ,Row_Type ,Item_Code , PaymentTotalAdvanceAmt * -1  from TSPL_PAYMENT_DETAIL_GST Left outer join TSPL_PAYMENT_HEADER on TSPL_PAYMENT_HEADER.Payment_No =TSPL_PAYMENT_DETAIL_GST.Payment_No  where isnull(TSPL_PAYMENT_HEADER.PurchaseOrder_No_GST,'')  ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(dt.Rows(i)("Item_Code")) & "'  and  TSPL_PAYMENT_HEADER.Payment_No<>'" & clsCommon.myCstr(txtPaymentNo.Value) & "' and isnull(TSPL_PAYMENT_HEADER. IsChkReverse,'') ='N' )zz " & Environment.NewLine & _
                    '"group by zz.Document_Code ,zz.Item_Code,Line_No  order by Line_No "

                    gvItem.Rows(gvItem.Rows.Count - 1).Cells(AdcolItem_Net_Amt).Value = clsCommon.myCdbl(clsDBFuncationality.getSingleValue(Qry))


                Next

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX1"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX1"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX1")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX1_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX1_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX1_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX2"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("tax2"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX2")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX2_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX2_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX2_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX3"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX3"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX3")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX3_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX3_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX3_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX4"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX4"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX4")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX4_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX4_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX4_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX5"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX5"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX5")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX5_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX5_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX5_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX6"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX6"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX6")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX6_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX6_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX6_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX7"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX7"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX7")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX7_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX7_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX7_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX8"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX8"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX8")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX8_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX8_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX8_Amt"))
                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX9"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX9"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX9")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX9_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX9_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX9_Amt"))

                End If

                If (clsCommon.myLen(clsCommon.myCstr(dt.Rows(0)("TAX10"))) > 0) Then
                    gvTaxDetail.Rows.AddNew()
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutCode).Value = clsCommon.myCstr(dt.Rows(0)("TAX10"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAutName).Value = clsTaxMaster.GetName(clsCommon.myCstr(dt.Rows(0)("TAX10")))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxRate).Value = clsCommon.myCdbl(dt.Rows(0)("TAX10_Rate"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTBaseAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX10_Base_Amt"))
                    gvTaxDetail.Rows(gvTaxDetail.Rows.Count - 1).Cells(colTTaxAmt).Value = clsCommon.myCdbl(dt.Rows(0)("TAX10_Amt"))
                End If

                SetitemWiseTaxSetting(False, False)
                For ii As Integer = 0 To gvItem.RowCount - 1
                    UpdateCurrentRow(ii)
                Next

                CalculateTaxAmountInAdvnce()
            End If


        End If

    End Sub

    Sub CalculateTaxAmountInAdvnce()
        Dim DblPaymentAmt As Double = 0
        Dim dblPOTotalAmt As Double = 0
        Dim dblTotalRateOfAllTaxes As Double = 0
        Dim dblPOTotalAddCharge As Double = 0
        DblPaymentAmt = clsCommon.myCdbl(txtPaymentAmt.Value)
        dblPOTotalAmt = clsCommon.myCdbl(LblPOTotalAmount.Text)
        dblPOTotalAddCharge = clsCommon.myCdbl(lblPOTotalAdditionalCharge.Text)
        'dblPOTotalAmt = dblPOTotalAmt - dblPOTotalAddCharge
        Dim dbltotalTaxamount As Double = 0


        'DblPaymentAmt = DblPaymentAmt - dblPOTotalAddCharge
        If DblPaymentAmt <= 0 Then
            DblPaymentAmt = dblPOTotalAmt
        End If

        If DblPaymentAmt > 0 Then

            For i As Integer = 0 To gvItem.Rows.Count - 1
                'dblTotalRateOfAllTaxes = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Rate).Value) + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Rate).Value)
                'gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value = clsCommon.myCdbl((DblPaymentAmt * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolItem_Net_Amt).Value)) / dblPOTotalAmt)
                'gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value = clsCommon.myCdbl((clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value) * 100.0) / (100.0 + clsCommon.myCdbl(dblTotalRateOfAllTaxes)))
                'gvItem.Rows(i).Cells(AdcolTAX1_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX2_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX3_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX4_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX5_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX6_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX7_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX8_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX9_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolTAX10_Amt_Payment).Value = (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value) * clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Rate).Value)) / 100.0
                'gvItem.Rows(i).Cells(AdcolPaymentTotalTax).Value = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX2_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX3_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX4_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX5_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX6_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX7_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX8_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX9_Amt_Payment).Value + gvItem.Rows(i).Cells(AdcolTAX10_Amt_Payment).Value)

                Dim dbladdchargeitemwise As Double = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select isnull(Total_ItemAdd_Charge,0)  from TSPL_PURCHASE_ORDER_detail where PurchaseOrder_No ='" & clsCommon.myCstr(txtPONo_GST.Value) & "' and Item_Code ='" & clsCommon.myCstr(gvItem.Rows(i).Cells(AdcolItem_Code).Value) & "' "))

                'gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value = clsCommon.myCdbl((DblPaymentAmt * (clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolItem_Net_Amt).Value) - dbladdchargeitemwise)) / dblPOTotalAmt)
                '  dblTotalRateOfAllTaxes = getcurrentTaxReceiptRateNonTaxable(i)
                'gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value = clsCommon.myCdbl((clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value) * 100.0) / (100.0 + clsCommon.myCdbl(dblTotalRateOfAllTaxes)))

                gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value = clsCommon.myCdbl((clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolAmt_Less_Discount).Value) * DblPaymentAmt) / (clsCommon.myCdbl(dblPOTotalAmt)))
                gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value = clsCommon.myCdbl((clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolItem_Net_Amt).Value) * DblPaymentAmt) / (clsCommon.myCdbl(dblPOTotalAmt)))

                Dim dblTotTaxAmtTaxInclusive As Double = 0
                Dim dblTotTaxAmtForTaxable As Double = 0
                For ii As Integer = 1 To 10
                    Dim Strii As String = clsCommon.myCstr(ii)
                    Dim dbltaxamt As Double = 0
                    ' Dim dblreceiptamtforcalculation As Double = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value) - clsCommon.myCdbl(lblPOTotalAdditionalCharge.Text)
                    Dim dblreceiptamtforcalculation As Double = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value)
                    Dim dblTotTaxRate As Double = dblTotalRateOfAllTaxes
                    Dim strTaxCode As String = clsCommon.myCstr(gvItem.Rows(i).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                    If clsCommon.myLen(strTaxCode) > 0 Then
                        Dim dblTaxRate As Double = clsCommon.myCdbl(gvItem.Rows(i).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                        If objCommonVar.TreatUnregisteredVendorAsRegisteredVendor Then
                            dblTaxRate = 0
                            gvItem.Rows(i).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value = dblTaxRate
                        End If

                        Dim IsTaxable As Boolean = clsCommon.myCBool(gvItem.Rows(i).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                        If clsCommon.myCBool(gvItem.Rows(i).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value) Then
                            'dbltaxamt = clsCommon.myCdbl(clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value) * dblTaxRate) / (100)
                            dbltaxamt = clsCommon.myCdbl(dblreceiptamtforcalculation * dblTaxRate) / (100)
                            dblTotTaxAmtForTaxable += dbltaxamt
                        Else
                            dbltaxamt = clsCommon.myCdbl((dblreceiptamtforcalculation + dblTotTaxAmtForTaxable) * dblTaxRate) / 100
                            'dblTotTaxRate = dblTotTaxRate - dblTaxRate
                        End If
                        dblTotTaxAmtTaxInclusive += dbltaxamt
                        gvItem.Rows(i).Cells(clsCommon.myCstr("COLTAXAMTPAYMENT" + Strii)).Value = Math.Round(dbltaxamt, 4)
                        ' dbltaxamt = dblreceiptamtforcalculation - dbltaxamt
                    End If
                Next
                'gvItem.Rows(i).Cells(AdcolReceiptAdvance).Value = DblReceiptAmt - dblTotTaxAmtTaxInclusive

                'gvItem.Rows(i).Cells(AdcolPaymentAdvance).Value = clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalAdvanceAmt).Value)
                gvItem.Rows(i).Cells(AdcolPaymentTotalTax).Value = dblTotTaxAmtTaxInclusive

            Next

            For i As Integer = 0 To gvItem.Rows.Count - 1
                dbltotalTaxamount = dbltotalTaxamount + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolPaymentTotalTax).Value)
            Next

            lblPOTotalTaxAmt.Text = Math.Round(clsCommon.myCdbl(dbltotalTaxamount), 4)

            CalculateTaxDetailForTaxgrid()
        End If
    End Sub



    Sub CalculateTaxDetailForTaxgrid()
        Dim DblReceiptTaxAmt1 As Double = 0
        Dim DblReceiptTaxAmt2 As Double = 0
        Dim DblReceiptTaxAm3 As Double = 0
        Dim DblReceiptTaxAmt4 As Double = 0
        Dim DblReceiptTaxAmt5 As Double = 0
        Dim DblReceiptTaxAmt6 As Double = 0
        Dim DblReceiptTaxAmt7 As Double = 0
        Dim DblReceiptTaxAmt8 As Double = 0
        Dim DblReceiptTaxAmt9 As Double = 0
        Dim DblReceiptTaxAmt10 As Double = 0
        If Not objCommonVar.TreatUnregisteredVendorAsRegisteredVendor Then
            For i As Integer = 0 To gvItem.Rows.Count - 1
                DblReceiptTaxAmt1 = DblReceiptTaxAmt1 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX1_Amt_Payment).Value)
                DblReceiptTaxAmt2 = DblReceiptTaxAmt2 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX2_Amt_Payment).Value)
                DblReceiptTaxAm3 = DblReceiptTaxAm3 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX3_Amt_Payment).Value)
                DblReceiptTaxAmt4 = DblReceiptTaxAmt4 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX4_Amt_Payment).Value)
                DblReceiptTaxAmt5 = DblReceiptTaxAmt5 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX5_Amt_Payment).Value)
                DblReceiptTaxAmt6 = DblReceiptTaxAmt6 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX6_Amt_Payment).Value)
                DblReceiptTaxAmt7 = DblReceiptTaxAmt7 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX7_Amt_Payment).Value)
                DblReceiptTaxAmt8 = DblReceiptTaxAmt8 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX8_Amt_Payment).Value)
                DblReceiptTaxAmt9 = DblReceiptTaxAmt9 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX9_Amt_Payment).Value)
                DblReceiptTaxAmt10 = DblReceiptTaxAmt10 + clsCommon.myCdbl(gvItem.Rows(i).Cells(AdcolTAX10_Amt_Payment).Value)
            Next
        End If
        If clsCommon.myCdbl(gvItem.Rows.Count) > 0 Then
            If clsCommon.myCdbl(gvItem.Rows.Count) > 0 Then
                Dim DblReceiptTaxAmt As Double = 0
                Dim dblTotTaxAmtForTaxable As Double = 0
                For ii As Integer = 0 To 9
                    Dim Strii As String = clsCommon.myCstr(ii + 1)
                    Dim strTaxCode As String = clsCommon.myCstr(gvItem.Rows(0).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                    Select Case (ii + 1)
                        Case 1
                            DblReceiptTaxAmt = DblReceiptTaxAmt1
                        Case 2
                            DblReceiptTaxAmt = DblReceiptTaxAmt2
                        Case 3
                            DblReceiptTaxAmt = DblReceiptTaxAm3
                        Case 4
                            DblReceiptTaxAmt = DblReceiptTaxAmt4
                        Case 5
                            DblReceiptTaxAmt = DblReceiptTaxAmt5
                        Case 6
                            DblReceiptTaxAmt = DblReceiptTaxAmt6
                        Case 7
                            DblReceiptTaxAmt = DblReceiptTaxAmt7
                        Case 8
                            DblReceiptTaxAmt = DblReceiptTaxAmt8
                        Case 9
                            DblReceiptTaxAmt = DblReceiptTaxAmt9
                        Case 10
                            DblReceiptTaxAmt = DblReceiptTaxAmt10
                    End Select

                    If clsCommon.myLen(strTaxCode) > 0 Then
                        Dim dblTaxRate As Double = clsCommon.myCdbl(gvItem.Rows(0).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                        If objCommonVar.TreatUnregisteredVendorAsRegisteredVendor Then
                            dblTaxRate = 0
                            gvTaxDetail.Rows(ii).Cells(colTTaxRate).Value = dblTaxRate
                        End If
                        Dim IsTaxable As Boolean = clsCommon.myCBool(gvItem.Rows(0).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                        If clsCommon.myCBool(gvItem.Rows(0).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value) Then
                            gvTaxDetail.Rows(ii).Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblReceiptTaxAmt)
                            gvTaxDetail.Rows(ii).Cells(colTBaseAmt).Value = Math.Round(clsCommon.myCdbl((DblReceiptTaxAmt * 100) / dblTaxRate), 2)
                            dblTotTaxAmtForTaxable += clsCommon.myCdbl(DblReceiptTaxAmt)
                        Else
                            gvTaxDetail.Rows(ii).Cells(colTTaxAmt).Value = clsCommon.myCdbl(DblReceiptTaxAmt)
                            gvTaxDetail.Rows(ii).Cells(colTBaseAmt).Value = Math.Round((clsCommon.myCdbl((DblReceiptTaxAmt * 100) / dblTaxRate)), 2)
                        End If

                    End If
                Next

            End If

        End If
    End Sub
    Private Function getcurrentTaxReceiptRateNonTaxable(ByVal IntRowNo As Integer) As Double
        Dim dblrate As Double = 0
        Try
            Dim arrTaxableAuth As New List(Of String)
            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)
                Dim strTaxCode As String = clsCommon.myCstr(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                If clsCommon.myLen(strTaxCode) > 0 Then
                    Dim IsTaxable As Boolean = clsCommon.myCBool(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                    If Not IsTaxable Then
                        dblrate = dblrate + clsCommon.myCdbl(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + Strii)).Value)
                    End If
                End If
            Next
            Return dblrate
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Function

    Sub SetitemWiseTaxSetting(ByVal isChangeRate As Boolean, ByVal isForCurrentRow As Boolean)
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetailsscrap(txtTaxGroup.Value)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If isForCurrentRow Then
                If clsCommon.myLen(gvItem.CurrentRow.Cells(AdcolItem_Code)) > 0 Then
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        Dim strII As String = clsCommon.myCstr(ii)
                        'gv1.CurrentRow.Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                        'If isChangeRate Then
                        '    gv1.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))
                        'End If
                        gvItem.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        gvItem.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        gvItem.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        'gvItem.CurrentRow.Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        'gvItem.CurrentRow.Cells(clsCommon.myCstr("RECOVERTABLETAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_Recoverable")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                End If
            Else
                For intRowNo As Integer = 0 To gvItem.Rows.Count - 1
                    ' BlankTaxDetails(intRowNo, isChangeRate)
                    If clsCommon.myLen(gvItem.Rows(intRowNo).Cells(AdcolItem_Code)) > 0 Then
                        Dim ii As Integer = 1
                        For Each dr As DataRow In dt.Rows
                            Dim strII As String = clsCommon.myCstr(ii)
                            gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("colTax" + strII)).Value = clsCommon.myCstr(dr("Tax_Code"))
                            'If isChangeRate Then
                            '    gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxRate" + strII)).Value = clsCommon.myCdbl(dr("TaxRate"))
                            'End If
                            gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                            gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                            gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                            'gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                            'gvItem.Rows(intRowNo).Cells(clsCommon.myCstr("RECOVERTABLETAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Tax_Recoverable")), "Y") = CompairStringResult.Equal)
                            ii = ii + 1
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub UpdateCurrentRow(ByVal IntRowNo As Integer)
        Try
            Dim arrTaxableAuth As New List(Of String)


            Dim dblAmt As Double = clsCommon.myCdbl(gvItem.Rows(IntRowNo).Cells(AdcolAmt_Less_Discount).Value)
            Dim dblAmtAfterDis As Double = dblAmt
            For ii As Integer = 1 To 10
                Dim Strii As String = clsCommon.myCstr(ii)
                Dim strTaxCode As String = clsCommon.myCstr(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAX" + Strii)).Value)
                If clsCommon.myLen(strTaxCode) > 0 Then
                    Dim dblTaxRate As Double = clsCommon.myCdbl(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value)
                    Dim IsSurTax As Boolean = clsCommon.myCBool(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value)
                    Dim strSurTaxCode As String = clsCommon.myCstr(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value)
                    Dim IsTaxable As Boolean = clsCommon.myCBool(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                    'Dim IsExcisable As Boolean = clsCommon.myCBool(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value)
                    Dim dblBaseAmt As Double = 0
                    Dim dblTaxAmt As Double = 0
                    If IsSurTax Then
                        Dim dblSurTaxAmt As Double = GetCurrentRowSurTaxAmt(IntRowNo, ii, strSurTaxCode)
                        dblBaseAmt = dblSurTaxAmt
                    Else
                        Dim dblOtherTaxAmt As Double = 0

                        dblOtherTaxAmt = GetCurrentRowOtherTaxAmt(IntRowNo, Strii, arrTaxableAuth)

                        dblBaseAmt = (dblAmtAfterDis + dblOtherTaxAmt)
                        'If gvItem.Rows(IntRowNo).Cells(AdcolFOC_Item).Value = 0 Then
                        '    dblBaseAmt = (dblAmtAfterDis + dblOtherTaxAmt)
                        'End If

                    End If

                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Math.Round(dblBaseAmt, 2)
                    dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Math.Round(dblTaxAmt, 6)
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMTPAYMENT" + Strii)).Value = Math.Round(dblTaxAmt, 6)
                    If IsTaxable AndAlso Not arrTaxableAuth.Contains(strTaxCode.ToUpper()) Then
                        arrTaxableAuth.Add(strTaxCode.ToUpper())
                    End If

                Else
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXBASEAMT" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxRate" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISSURTAX" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value = Nothing
                    gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("COLTAXAMTPAYMENT" + Strii)).Value = Nothing
                    'gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value = Nothing
                End If

            Next



            'Dim dblTotTaxAmt As Double = GetCurrentRowTotalTaxAmt(IntRowNo)
            'Dim dblAmtAfterTax As Double = dblAmtAfterDis + dblTotTaxAmt

        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub


    Private Function GetCurrentRowSurTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal strSurTaxCode As String) As Double
        For ii As Integer = 1 To intEndCol
            Dim strii As String = clsCommon.myCstr(ii)
            If IntRowNo < 0 Then
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            Else
                If clsCommon.CompairString(strSurTaxCode, clsCommon.myCstr(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                    Return clsCommon.myCdbl(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                End If
            End If
        Next
        Return 0
    End Function

    Private Function GetCurrentRowOtherTaxAmt(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Double
        Dim dblRet As Double = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If IntRowNo < 0 Then
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                Else
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxAmt" + strii)).Value)
                    End If
                End If
            Next
        Next
        Return dblRet
    End Function
    Private Function GetCurrentRowTaxRate(ByVal IntRowNo As Integer, ByVal intEndCol As Integer, ByVal arrTaxableAuth As List(Of String)) As Double
        Dim dblRet As Double = 0
        For Each strTaxAuth As String In arrTaxableAuth
            For ii As Integer = 1 To intEndCol
                Dim strii As String = clsCommon.myCstr(ii)
                If IntRowNo < 0 Then
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gvItem.CurrentRow.Cells(clsCommon.myCstr("colTaxRate" + strii)).Value)
                    End If
                Else
                    If clsCommon.CompairString(strTaxAuth, clsCommon.myCstr(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTax" + strii)).Value)) = CompairStringResult.Equal Then
                        dblRet = dblRet + clsCommon.myCdbl(gvItem.Rows(IntRowNo).Cells(clsCommon.myCstr("colTaxRate" + strii)).Value)
                    End If
                End If
            Next
        Next
        Return dblRet
    End Function



    Sub LoadBlankGridTax()
        gvTaxDetail.Rows.Clear()
        gvTaxDetail.Columns.Clear()

        Dim repoTaxAuthCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthCode.FormatString = ""
        repoTaxAuthCode.HeaderText = "Tax Authority Code"
        repoTaxAuthCode.Name = colTTaxAutCode
        repoTaxAuthCode.Width = 150
        repoTaxAuthCode.ReadOnly = True
        gvTaxDetail.MasterTemplate.Columns.Add(repoTaxAuthCode)

        Dim repoTaxAuthName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthName.FormatString = ""
        repoTaxAuthName.HeaderText = "Tax Authority"
        repoTaxAuthName.Name = colTTaxAutName
        repoTaxAuthName.Width = 200
        repoTaxAuthName.ReadOnly = True
        gvTaxDetail.MasterTemplate.Columns.Add(repoTaxAuthName)

        Dim repoTaxBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt.FormatString = "{0:n4}"
        repoTaxBaseAmt.HeaderText = "Base Amount"
        repoTaxBaseAmt.Name = colTBaseAmt
        repoTaxBaseAmt.Width = 100
        repoTaxBaseAmt.ReadOnly = True
        repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvTaxDetail.MasterTemplate.Columns.Add(repoTaxBaseAmt)

        Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate.FormatString = "{0:n2}"
        repoTaxRate.HeaderText = "Tax Rate"
        repoTaxRate.Name = colTTaxRate
        repoTaxRate.Width = 100
        repoTaxRate.ReadOnly = True
        repoTaxRate.IsVisible = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, Nothing)) = 1, True, False)
        repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvTaxDetail.MasterTemplate.Columns.Add(repoTaxRate)

        Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt.FormatString = "{0:n4}"
        repoTaxAmt.HeaderText = "Tax Amount"
        repoTaxAmt.Name = colTTaxAmt
        repoTaxAmt.Width = 100
        repoTaxAmt.ReadOnly = False
        repoTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvTaxDetail.MasterTemplate.Columns.Add(repoTaxAmt)

        gvTaxDetail.AllowDeleteRow = True
        gvTaxDetail.AllowAddNewRow = False
        gvTaxDetail.ShowGroupPanel = False
        gvTaxDetail.AllowColumnReorder = False
        gvTaxDetail.AllowRowReorder = False
        gvTaxDetail.EnableSorting = False
        gvTaxDetail.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvTaxDetail.MasterTemplate.ShowRowHeaderColumn = False

    End Sub

    Sub LoadBlankGridPOItemDetail()
        gvItem.Rows.Clear()
        gvItem.Columns.Clear()

        Dim repoLineNo As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoLineNo = New GridViewDecimalColumn()
        repoLineNo.FormatString = ""
        repoLineNo.HeaderText = "Line No"
        repoLineNo.Name = AdcolLine_No
        repoLineNo.Width = 50
        repoLineNo.ReadOnly = True
        repoLineNo.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoLineNo)

        Dim repoDocCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoDocCode.FormatString = ""
        repoDocCode.HeaderText = "Document No"
        repoDocCode.Width = 70
        repoDocCode.Name = AdcolDocument_Code
        repoDocCode.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoDocCode)

        Dim repoRowType As GridViewComboBoxColumn = New GridViewComboBoxColumn()
        repoRowType.FormatString = ""
        repoRowType.HeaderText = "Row Type"
        repoRowType.Name = AdcolRow_Type
        repoRowType.Width = 50
        repoRowType.ReadOnly = True
        repoRowType.TextAlignment = System.Drawing.ContentAlignment.MiddleLeft
        gvItem.MasterTemplate.Columns.Add(repoRowType)

        Dim repoICode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoICode.FormatString = ""
        repoICode.HeaderText = "Item Code"
        repoICode.Name = AdcolItem_Code
        repoICode.HeaderImage = Global.ERP.My.Resources.Resources.search4
        repoICode.TextImageRelation = TextImageRelation.TextBeforeImage
        repoICode.Width = 100
        repoICode.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoICode)

        Dim repoIName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIName.FormatString = ""
        repoIName.HeaderText = "Item Description"
        repoIName.Name = adcolIName
        repoIName.Width = 150
        repoIName.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoIName)


        Dim repoIHSN As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoIHSN.FormatString = ""
        repoIHSN.HeaderText = "HSN Code"
        repoIHSN.Name = adcolIHSN
        repoIHSN.Width = 150
        repoIHSN.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoIHSN)


        Dim repoBalQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoBalQty.FormatString = ""
        repoBalQty.WrapText = True
        repoBalQty.HeaderText = "Balance Quantity"
        repoBalQty.Name = AdcolBalance_Qty
        repoBalQty.Width = 80
        repoBalQty.Minimum = 0
        repoBalQty.IsVisible = False
        repoBalQty.ReadOnly = True
        repoBalQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoBalQty)

        Dim repoQty As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoQty.FormatString = ""
        repoQty.WrapText = True
        repoQty.HeaderText = "Qty"
        repoQty.Name = AdcolQty
        repoQty.Width = 80
        repoQty.Minimum = 0
        repoQty.ReadOnly = True
        repoQty.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoQty)

        Dim repoUnitcode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoUnitcode.FormatString = ""
        repoUnitcode.HeaderText = "Unit Code"
        repoUnitcode.Name = AdcolUnit_code
        repoUnitcode.IsVisible = False
        repoUnitcode.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoUnitcode)

        Dim repoItemCost As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoItemCost = New GridViewDecimalColumn()
        repoItemCost.FormatString = ""
        repoItemCost.HeaderText = "Item Cost"
        repoItemCost.Name = AdcolItem_Cost
        repoItemCost.IsVisible = False
        repoItemCost.Minimum = 0
        repoItemCost.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoItemCost.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoItemCost)



        Dim repoTax1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax1.FormatString = ""
        repoTax1.HeaderText = "Tax 1"
        repoTax1.Name = AdcolTAX1
        repoTax1.ReadOnly = True
        repoTax1.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax1)

        Dim repoTaxBaseAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt1.FormatString = ""
        repoTaxBaseAmt1.HeaderText = "Tax Base Amount 1"
        repoTaxBaseAmt1.Name = AdcolTAX1_Base_Amt
        repoTaxBaseAmt1.ReadOnly = True
        repoTaxBaseAmt1.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt1)

        Dim repoTaxRate1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate1 = New GridViewDecimalColumn()
        repoTaxRate1.FormatString = ""
        repoTaxRate1.HeaderText = "Tax Rate 1"
        repoTaxRate1.Name = AdcolTAX1_Rate
        repoTaxRate1.IsVisible = False
        repoTaxRate1.ReadOnly = True
        repoTaxRate1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate1)

        Dim repoTaxAmt1 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt1 = New GridViewDecimalColumn()
        repoTaxAmt1.FormatString = ""
        repoTaxAmt1.HeaderText = "Tax Amt 1"
        repoTaxAmt1.Name = AdcolTAX1_Amt
        repoTaxAmt1.IsVisible = False
        repoTaxAmt1.ReadOnly = True
        repoTaxAmt1.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt1)

        Dim repoIsSurTax1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax1.HeaderText = "Is Surtax 1"
        repoIsSurTax1.Name = colIsSurTax1
        repoIsSurTax1.ReadOnly = True
        repoIsSurTax1.IsVisible = False
        repoIsSurTax1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax1)

        Dim repoSurTaxCode1 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode1.FormatString = ""
        repoSurTaxCode1.HeaderText = "Surtax 1"
        repoSurTaxCode1.Name = colSurTaxCode1
        repoSurTaxCode1.ReadOnly = True
        repoSurTaxCode1.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode1)

        Dim repoIsTaxable1 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable1.HeaderText = "Is Taxable 1"
        repoIsTaxable1.Name = colIsTaxable1
        repoIsTaxable1.ReadOnly = True
        repoIsTaxable1.IsVisible = False
        repoIsTaxable1.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable1)

        Dim repoTaxAmt1Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt1Receipt = New GridViewDecimalColumn()
        repoTaxAmt1Receipt.FormatString = ""
        repoTaxAmt1Receipt.HeaderText = "Tax Amt 1 Payment"
        repoTaxAmt1Receipt.Name = AdcolTAX1_Amt_Payment
        repoTaxAmt1Receipt.IsVisible = False
        repoTaxAmt1Receipt.ReadOnly = True
        repoTaxAmt1Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt1Receipt)


        Dim repoTax2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax2.FormatString = ""
        repoTax2.HeaderText = "Tax 2"
        repoTax2.Name = Adcoltax2
        repoTax2.ReadOnly = True
        repoTax2.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax2)

        Dim repoTaxBaseAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt2.FormatString = ""
        repoTaxBaseAmt2.HeaderText = "Tax Base Amount 2"
        repoTaxBaseAmt2.Name = AdcolTAX2_Base_Amt
        repoTaxBaseAmt2.ReadOnly = True
        repoTaxBaseAmt2.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt2)

        Dim repoTaxRate2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate2 = New GridViewDecimalColumn()
        repoTaxRate2.FormatString = ""
        repoTaxRate2.HeaderText = "Tax Rate 2"
        repoTaxRate2.Name = AdcolTAX2_Rate
        repoTaxRate2.IsVisible = False
        repoTaxRate2.ReadOnly = True
        repoTaxRate2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate2)

        Dim repoTaxAmt2 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt2 = New GridViewDecimalColumn()
        repoTaxAmt2.FormatString = ""
        repoTaxAmt2.HeaderText = "Tax Amt 2"
        repoTaxAmt2.Name = AdcolTAX2_Amt
        repoTaxAmt2.IsVisible = False
        repoTaxAmt2.ReadOnly = True
        repoTaxAmt2.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt2)

        Dim repoIsSurTax2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax2.HeaderText = "Is Surtax 2"
        repoIsSurTax2.Name = colIsSurTax2
        repoIsSurTax2.ReadOnly = True
        repoIsSurTax2.IsVisible = False
        repoIsSurTax2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax2)

        Dim repoSurTaxCode2 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode2.FormatString = ""
        repoSurTaxCode2.HeaderText = "Surtax 2"
        repoSurTaxCode2.Name = colSurTaxCode2
        repoSurTaxCode2.ReadOnly = True
        repoSurTaxCode2.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode2)

        Dim repoIsTaxable2 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable2.HeaderText = "Is Taxable 2"
        repoIsTaxable2.Name = colIsTaxable2
        repoIsTaxable2.ReadOnly = True
        repoIsTaxable2.IsVisible = False
        repoIsTaxable2.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable2)

        Dim repoTaxAmt2Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt2Receipt = New GridViewDecimalColumn()
        repoTaxAmt2Receipt.FormatString = ""
        repoTaxAmt2Receipt.HeaderText = "Tax Amt 2 Payment"
        repoTaxAmt2Receipt.Name = AdcolTAX2_Amt_Payment
        repoTaxAmt2Receipt.IsVisible = False
        repoTaxAmt2Receipt.ReadOnly = True
        repoTaxAmt2Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt2Receipt)


        Dim repoTax3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax3.FormatString = ""
        repoTax3.HeaderText = "Tax 3"
        repoTax3.Name = AdcolTAX3
        repoTax3.ReadOnly = True
        repoTax3.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax3)

        Dim repoTaxBaseAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt3.FormatString = ""
        repoTaxBaseAmt3.HeaderText = "Tax Base Amount 3"
        repoTaxBaseAmt3.Name = AdcolTAX3_Base_Amt
        repoTaxBaseAmt3.ReadOnly = True
        repoTaxBaseAmt3.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt3)

        Dim repoTaxRate3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate3 = New GridViewDecimalColumn()
        repoTaxRate3.FormatString = ""
        repoTaxRate3.HeaderText = "Tax Rate 3"
        repoTaxRate3.Name = AdcolTAX3_Rate
        repoTaxRate3.IsVisible = False
        repoTaxRate3.ReadOnly = True
        repoTaxRate3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate3)

        Dim repoTaxAmt3 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt3 = New GridViewDecimalColumn()
        repoTaxAmt3.FormatString = ""
        repoTaxAmt3.HeaderText = "Tax Amt 3"
        repoTaxAmt3.Name = AdcolTAX3_Amt
        repoTaxAmt3.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt3.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt3)

        Dim repoIsSurTax3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax3.HeaderText = "Is Surtax 3"
        repoIsSurTax3.Name = colIsSurTax3
        repoIsSurTax3.ReadOnly = True
        repoIsSurTax3.IsVisible = False
        repoIsSurTax3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax3)

        Dim repoSurTaxCode3 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode3.FormatString = ""
        repoSurTaxCode3.HeaderText = "Surtax 3"
        repoSurTaxCode3.Name = colSurTaxCode3
        repoSurTaxCode3.ReadOnly = True
        repoSurTaxCode3.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode3)

        Dim repoIsTaxable3 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable3.HeaderText = "Is Taxable 3"
        repoIsTaxable3.Name = colIsTaxable3
        repoIsTaxable3.ReadOnly = True
        repoIsTaxable3.IsVisible = False
        repoIsTaxable3.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable3)

        Dim repoTaxAmt3Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt3Receipt = New GridViewDecimalColumn()
        repoTaxAmt3Receipt.FormatString = ""
        repoTaxAmt3Receipt.HeaderText = "Tax Amt 3 Payment"
        repoTaxAmt3Receipt.Name = AdcolTAX3_Amt_Payment
        repoTaxAmt3Receipt.IsVisible = False
        repoTaxAmt3Receipt.ReadOnly = True
        repoTaxAmt3Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt3Receipt)


        Dim repoTax4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax4.FormatString = ""
        repoTax4.HeaderText = "Tax 4"
        repoTax4.Name = AdcolTAX4
        repoTax4.ReadOnly = True
        repoTax4.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax4)

        Dim repoTaxBaseAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt4.FormatString = ""
        repoTaxBaseAmt4.HeaderText = "Tax Base Amount 4"
        repoTaxBaseAmt4.Name = AdcolTAX4_Base_Amt
        repoTaxBaseAmt4.ReadOnly = True
        repoTaxBaseAmt4.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt4)

        Dim repoTaxRate4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate4 = New GridViewDecimalColumn()
        repoTaxRate4.FormatString = ""
        repoTaxRate4.HeaderText = "Tax Rate 4"
        repoTaxRate4.Name = AdcolTAX4_Rate
        repoTaxRate4.IsVisible = False
        repoTaxRate4.ReadOnly = True
        repoTaxRate4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate4)

        Dim repoTaxAmt4 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt4 = New GridViewDecimalColumn()
        repoTaxAmt4.FormatString = ""
        repoTaxAmt4.HeaderText = "Tax Amt 4"
        repoTaxAmt4.Name = AdcolTAX4_Amt
        repoTaxAmt4.IsVisible = False
        repoTaxAmt4.ReadOnly = True
        repoTaxAmt4.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt4)

        Dim repoIsSurTax4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax4.HeaderText = "Is Surtax 4"
        repoIsSurTax4.Name = colIsSurTax4
        repoIsSurTax4.ReadOnly = True
        repoIsSurTax4.IsVisible = False
        repoIsSurTax4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax4)

        Dim repoSurTaxCode4 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode4.FormatString = ""
        repoSurTaxCode4.HeaderText = "Surtax 4"
        repoSurTaxCode4.Name = colSurTaxCode4
        repoSurTaxCode4.ReadOnly = True
        repoSurTaxCode4.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode4)

        Dim repoIsTaxable4 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable4.HeaderText = "Is Taxable 4"
        repoIsTaxable4.Name = colIsTaxable4
        repoIsTaxable4.ReadOnly = True
        repoIsTaxable4.IsVisible = False
        repoIsTaxable4.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable4)

        Dim repoTaxAmt4Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt4Receipt = New GridViewDecimalColumn()
        repoTaxAmt4Receipt.FormatString = ""
        repoTaxAmt4Receipt.HeaderText = "Tax Amt 4 Payment"
        repoTaxAmt4Receipt.Name = AdcolTAX4_Amt_Payment
        repoTaxAmt4Receipt.IsVisible = False
        repoTaxAmt4Receipt.ReadOnly = True
        repoTaxAmt4Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt4Receipt)


        Dim repoTax5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax5.FormatString = ""
        repoTax5.HeaderText = "Tax 5"
        repoTax5.Name = AdcolTAX5
        repoTax5.ReadOnly = True
        repoTax5.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax5)

        Dim repoTaxBaseAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt5.FormatString = ""
        repoTaxBaseAmt5.HeaderText = "Tax Base Amount 5"
        repoTaxBaseAmt5.Name = AdcolTAX5_Base_Amt
        repoTaxBaseAmt5.ReadOnly = True
        repoTaxBaseAmt5.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt5)

        Dim repoTaxRate5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate5 = New GridViewDecimalColumn()
        repoTaxRate5.FormatString = ""
        repoTaxRate5.HeaderText = "Tax Rate 5"
        repoTaxRate5.Name = AdcolTAX5_Rate
        repoTaxRate5.IsVisible = False
        repoTaxRate5.ReadOnly = True
        repoTaxRate5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate5)

        Dim repoTaxAmt5 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt5 = New GridViewDecimalColumn()
        repoTaxAmt5.FormatString = ""
        repoTaxAmt5.HeaderText = "Tax Amt 5"
        repoTaxAmt5.Name = AdcolTAX5_Amt
        repoTaxAmt5.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt5.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt5)

        Dim repoIsSurTax5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax5.HeaderText = "Is Surtax 5"
        repoIsSurTax5.Name = colIsSurTax5
        repoIsSurTax5.ReadOnly = True
        repoIsSurTax5.IsVisible = False
        repoIsSurTax5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax5)

        Dim repoSurTaxCode5 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode5.FormatString = ""
        repoSurTaxCode5.HeaderText = "Surtax 5"
        repoSurTaxCode5.Name = colSurTaxCode5
        repoSurTaxCode5.ReadOnly = True
        repoSurTaxCode5.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode5)

        Dim repoIsTaxable5 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable5.HeaderText = "Is Taxable 5"
        repoIsTaxable5.Name = colIsTaxable5
        repoIsTaxable5.ReadOnly = True
        repoIsTaxable5.IsVisible = False
        repoIsTaxable5.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable5)

        Dim repoTaxAmt5Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt5Receipt = New GridViewDecimalColumn()
        repoTaxAmt5Receipt.FormatString = ""
        repoTaxAmt5Receipt.HeaderText = "Tax Amt 5 Payment"
        repoTaxAmt5Receipt.Name = AdcolTAX5_Amt_Payment
        repoTaxAmt5Receipt.IsVisible = False
        repoTaxAmt3.ReadOnly = True
        repoTaxAmt5Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt5Receipt)


        Dim repoTax6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax6.FormatString = ""
        repoTax6.HeaderText = "Tax 6"
        repoTax6.Name = AdcolTAX6
        repoTax6.ReadOnly = True
        repoTax6.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax6)

        Dim repoTaxBaseAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt6.FormatString = ""
        repoTaxBaseAmt6.HeaderText = "Tax Base Amount 6"
        repoTaxBaseAmt6.Name = AdcolTAX6_Base_Amt
        repoTaxBaseAmt6.ReadOnly = True
        repoTaxBaseAmt6.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt6)

        Dim repoTaxRate6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate6 = New GridViewDecimalColumn()
        repoTaxRate6.FormatString = ""
        repoTaxRate6.HeaderText = "Tax Rate 6"
        repoTaxRate6.Name = AdcolTAX6_Rate
        repoTaxRate6.IsVisible = False
        repoTaxRate6.ReadOnly = True
        repoTaxRate6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate6)

        Dim repoTaxAmt6 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt6 = New GridViewDecimalColumn()
        repoTaxAmt6.FormatString = ""
        repoTaxAmt6.HeaderText = "Tax Amt 6"
        repoTaxAmt6.Name = AdcolTAX6_Amt
        repoTaxAmt6.IsVisible = False
        repoTaxAmt6.ReadOnly = True
        repoTaxAmt6.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt6)

        Dim repoIsSurTax6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax6.HeaderText = "Is Surtax 6"
        repoIsSurTax6.Name = colIsSurTax6
        repoIsSurTax6.ReadOnly = True
        repoIsSurTax6.IsVisible = False
        repoIsSurTax6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax6)

        Dim repoSurTaxCode6 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode6.FormatString = ""
        repoSurTaxCode6.HeaderText = "Surtax 6"
        repoSurTaxCode6.Name = colSurTaxCode6
        repoSurTaxCode6.ReadOnly = True
        repoSurTaxCode6.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode6)

        Dim repoIsTaxable6 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable6.HeaderText = "Is Taxable 6"
        repoIsTaxable6.Name = colIsTaxable6
        repoIsTaxable6.ReadOnly = True
        repoIsTaxable6.IsVisible = False
        repoIsTaxable6.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable6)

        Dim repoTaxAmt6Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt6Receipt = New GridViewDecimalColumn()
        repoTaxAmt6Receipt.FormatString = ""
        repoTaxAmt6Receipt.HeaderText = "Tax Amt 6 Payment"
        repoTaxAmt6Receipt.Name = AdcolTAX6_Amt_Payment
        repoTaxAmt6Receipt.IsVisible = False
        repoTaxAmt6Receipt.ReadOnly = True
        repoTaxAmt6Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt6Receipt)


        Dim repoTax7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax7.FormatString = ""
        repoTax7.HeaderText = "Tax 7"
        repoTax7.Name = AdcolTAX7
        repoTax7.ReadOnly = True
        repoTax7.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax7)

        Dim repoTaxBaseAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt7.FormatString = ""
        repoTaxBaseAmt7.HeaderText = "Tax Base Amount 7"
        repoTaxBaseAmt7.Name = AdcolTAX7_Base_Amt
        repoTaxBaseAmt7.ReadOnly = True
        repoTaxBaseAmt7.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt7)

        Dim repoTaxRate7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate7 = New GridViewDecimalColumn()
        repoTaxRate7.FormatString = ""
        repoTaxRate7.HeaderText = "Tax Rate 7"
        repoTaxRate7.Name = AdcolTAX7_Rate
        repoTaxRate7.IsVisible = False
        repoTaxRate7.ReadOnly = True
        repoTaxRate7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate7)

        Dim repoTaxAmt7 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt7 = New GridViewDecimalColumn()
        repoTaxAmt7.FormatString = ""
        repoTaxAmt7.HeaderText = "Tax Amt 7"
        repoTaxAmt7.Name = AdcolTAX7_Amt
        repoTaxAmt7.IsVisible = False
        repoTaxAmt7.ReadOnly = True
        repoTaxAmt7.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt7)

        Dim repoIsSurTax7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax7.HeaderText = "Is Surtax 7"
        repoIsSurTax7.Name = colIsSurTax7
        repoIsSurTax7.ReadOnly = True
        repoIsSurTax7.IsVisible = False
        repoIsSurTax7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax7)

        Dim repoSurTaxCode7 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode7.FormatString = ""
        repoSurTaxCode7.HeaderText = "Surtax 7"
        repoSurTaxCode7.Name = colSurTaxCode7
        repoSurTaxCode7.ReadOnly = True
        repoSurTaxCode7.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode7)

        Dim repoIsTaxable7 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable7.HeaderText = "Is Taxable 7"
        repoIsTaxable7.Name = colIsTaxable7
        repoIsTaxable7.ReadOnly = True
        repoIsTaxable7.IsVisible = False
        repoIsTaxable7.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable7)

        Dim repoTaxAmt7Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt7Receipt = New GridViewDecimalColumn()
        repoTaxAmt7Receipt.FormatString = ""
        repoTaxAmt7Receipt.HeaderText = "Tax Amt 7 Payment"
        repoTaxAmt7Receipt.Name = AdcolTAX7_Amt_Payment
        repoTaxAmt7Receipt.IsVisible = False
        repoTaxAmt7Receipt.ReadOnly = True
        repoTaxAmt7Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt7Receipt)



        Dim repoTax8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax8.FormatString = ""
        repoTax8.HeaderText = "Tax 8"
        repoTax8.Name = AdcolTAX8
        repoTax8.ReadOnly = True
        repoTax8.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax8)

        Dim repoTaxBaseAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt8.FormatString = ""
        repoTaxBaseAmt8.HeaderText = "Tax Base Amount 8"
        repoTaxBaseAmt8.Name = AdcolTAX8_Base_Amt
        repoTaxBaseAmt8.ReadOnly = True
        repoTaxBaseAmt8.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt8)

        Dim repoTaxRate8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate8 = New GridViewDecimalColumn()
        repoTaxRate8.FormatString = ""
        repoTaxRate8.HeaderText = "Tax Rate 8"
        repoTaxRate8.Name = AdcolTAX8_Rate
        repoTaxRate8.IsVisible = False
        repoTaxRate8.ReadOnly = True
        repoTaxRate8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate8)

        Dim repoTaxAmt8 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt8 = New GridViewDecimalColumn()
        repoTaxAmt8.FormatString = ""
        repoTaxAmt8.HeaderText = "Tax Amt 8"
        repoTaxAmt8.Name = AdcolTAX8_Amt
        repoTaxAmt8.IsVisible = False
        repoTaxAmt8.ReadOnly = True
        repoTaxAmt8.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt8)

        Dim repoIsSurTax8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax8.HeaderText = "Is Surtax 8"
        repoIsSurTax8.Name = colIsSurTax8
        repoIsSurTax8.ReadOnly = True
        repoIsSurTax8.IsVisible = False
        repoIsSurTax8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax8)

        Dim repoSurTaxCode8 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode8.FormatString = ""
        repoSurTaxCode8.HeaderText = "Surtax 8"
        repoSurTaxCode8.Name = colSurTaxCode8
        repoSurTaxCode8.ReadOnly = True
        repoSurTaxCode8.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode8)

        Dim repoIsTaxable8 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable8.HeaderText = "Is Taxable 8"
        repoIsTaxable8.Name = colIsTaxable8
        repoIsTaxable8.ReadOnly = True
        repoIsTaxable8.IsVisible = False
        repoIsTaxable8.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable8)

        Dim repoTaxAmt8Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt8Receipt = New GridViewDecimalColumn()
        repoTaxAmt8Receipt.FormatString = ""
        repoTaxAmt8Receipt.HeaderText = "Tax Amt 8 Payment"
        repoTaxAmt8Receipt.Name = AdcolTAX8_Amt_Payment
        repoTaxAmt8Receipt.IsVisible = False
        repoTaxAmt8Receipt.ReadOnly = True
        repoTaxAmt8Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt8Receipt)

        Dim repoTax9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax9.FormatString = ""
        repoTax9.HeaderText = "Tax 9"
        repoTax9.Name = AdcolTAX9
        repoTax9.ReadOnly = True
        repoTax9.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax9)

        Dim repoTaxBaseAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt9.FormatString = ""
        repoTaxBaseAmt9.HeaderText = "Tax Base Amount 9"
        repoTaxBaseAmt9.Name = AdcolTAX9_Base_Amt
        repoTaxBaseAmt9.ReadOnly = True
        repoTaxBaseAmt9.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt9)

        Dim repoTaxRate9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate9 = New GridViewDecimalColumn()
        repoTaxRate9.FormatString = ""
        repoTaxRate9.HeaderText = "Tax Rate 9"
        repoTaxRate9.Name = AdcolTAX9_Rate
        repoTaxRate9.IsVisible = False
        repoTaxRate9.ReadOnly = True
        repoTaxRate9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate9)

        Dim repoTaxAmt9 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt9 = New GridViewDecimalColumn()
        repoTaxAmt9.FormatString = ""
        repoTaxAmt9.HeaderText = "Tax Amt 9"
        repoTaxAmt9.Name = AdcolTAX9_Amt
        repoTaxAmt9.IsVisible = False
        repoTaxAmt9.ReadOnly = True
        repoTaxAmt9.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt9)

        Dim repoIsSurTax9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax9.HeaderText = "Is Surtax 9"
        repoIsSurTax9.Name = colIsSurTax9
        repoIsSurTax9.ReadOnly = True
        repoIsSurTax9.IsVisible = False
        repoIsSurTax9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax9)

        Dim repoSurTaxCode9 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode9.FormatString = ""
        repoSurTaxCode9.HeaderText = "Surtax 9"
        repoSurTaxCode9.Name = colSurTaxCode9
        repoSurTaxCode9.ReadOnly = True
        repoSurTaxCode9.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode9)

        Dim repoIsTaxable9 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable9.HeaderText = "Is Taxable 9"
        repoIsTaxable9.Name = colIsTaxable9
        repoIsTaxable9.ReadOnly = True
        repoIsTaxable9.IsVisible = False
        repoIsTaxable9.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable9)

        Dim repoTaxAmt9Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt9Receipt = New GridViewDecimalColumn()
        repoTaxAmt9Receipt.FormatString = ""
        repoTaxAmt9Receipt.HeaderText = "Tax Amt 9 Payment"
        repoTaxAmt9Receipt.Name = AdcolTAX9_Amt_Payment
        repoTaxAmt9Receipt.IsVisible = False
        repoTaxAmt9Receipt.ReadOnly = True
        repoTaxAmt9Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt9Receipt)


        Dim repoTax10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTax10.FormatString = ""
        repoTax10.HeaderText = "Tax 10"
        repoTax10.Name = AdcolTAX10
        repoTax10.ReadOnly = True
        repoTax10.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTax10)

        Dim repoTaxBaseAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt10.FormatString = ""
        repoTaxBaseAmt10.HeaderText = "Tax Base Amount 10"
        repoTaxBaseAmt10.Name = AdcolTAX10_Base_Amt
        repoTaxBaseAmt10.ReadOnly = True
        repoTaxBaseAmt10.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoTaxBaseAmt10)

        Dim repoTaxRate10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate10 = New GridViewDecimalColumn()
        repoTaxRate10.FormatString = ""
        repoTaxRate10.HeaderText = "Tax Rate 10"
        repoTaxRate10.Name = AdcolTAX10_Rate
        repoTaxRate10.IsVisible = False
        repoTaxRate10.ReadOnly = True
        repoTaxRate10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxRate10)

        Dim repoTaxAmt10 As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt10 = New GridViewDecimalColumn()
        repoTaxAmt10.FormatString = ""
        repoTaxAmt10.HeaderText = "Tax Amt 10"
        repoTaxAmt10.Name = AdcolTAX10_Amt
        repoTaxAmt10.IsVisible = False
        repoTaxAmt10.ReadOnly = True
        repoTaxAmt10.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt10)

        Dim repoIsSurTax10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsSurTax10.HeaderText = "Is Surtax 10"
        repoIsSurTax10.Name = colIsSurTax10
        repoIsSurTax10.ReadOnly = True
        repoIsSurTax10.IsVisible = False
        repoIsSurTax10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsSurTax10)

        Dim repoSurTaxCode10 As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoSurTaxCode10.FormatString = ""
        repoSurTaxCode10.HeaderText = "Surtax 10"
        repoSurTaxCode10.Name = colSurTaxCode10
        repoSurTaxCode10.ReadOnly = True
        repoSurTaxCode10.IsVisible = False
        gvItem.MasterTemplate.Columns.Add(repoSurTaxCode10)

        Dim repoIsTaxable10 As GridViewCheckBoxColumn = New GridViewCheckBoxColumn()
        repoIsTaxable10.HeaderText = "Is Taxable 10"
        repoIsTaxable10.Name = colIsTaxable10
        repoIsTaxable10.ReadOnly = True
        repoIsTaxable10.IsVisible = False
        repoIsTaxable10.TextAlignment = System.Drawing.ContentAlignment.MiddleCenter
        gvItem.MasterTemplate.Columns.Add(repoIsTaxable10)

        Dim repoTaxAmt10Receipt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt10Receipt = New GridViewDecimalColumn()
        repoTaxAmt10Receipt.FormatString = ""
        repoTaxAmt10Receipt.HeaderText = "Tax Amt 10 Payment"
        repoTaxAmt10Receipt.Name = AdcolTAX10_Amt_Payment
        repoTaxAmt10Receipt.IsVisible = False
        repoTaxAmt10Receipt.ReadOnly = True
        repoTaxAmt10Receipt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTaxAmt10Receipt)


        Dim repoAmount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmount = New GridViewDecimalColumn()
        repoAmount.FormatString = ""
        repoAmount.HeaderText = "Amount"
        repoAmount.Name = AdcolAmount
        repoAmount.WrapText = True
        repoAmount.Width = 80
        repoAmount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        repoAmount.VisibleInColumnChooser = False
        repoAmount.ReadOnly = True
        gvItem.MasterTemplate.Columns.Add(repoAmount)


        Dim repoDisc_Per As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisc_Per.FormatString = ""
        repoDisc_Per.HeaderText = "Disc Per"
        repoDisc_Per.Name = AdcolDisc_Per
        repoDisc_Per.Width = 80
        repoDisc_Per.ReadOnly = True
        repoDisc_Per.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoDisc_Per)

        Dim repoDisc_Amt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoDisc_Amt.FormatString = ""
        repoDisc_Amt.HeaderText = "Disc Amount"
        repoDisc_Amt.Name = AdcolDisc_Amt
        repoDisc_Amt.Width = 80
        repoDisc_Amt.ReadOnly = True
        repoDisc_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoDisc_Amt)

        Dim repoAmt_Less_Discount As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoAmt_Less_Discount.FormatString = ""
        repoAmt_Less_Discount.HeaderText = "Amount Less Discount"
        repoAmt_Less_Discount.Name = AdcolAmt_Less_Discount
        repoAmt_Less_Discount.Width = 80
        repoAmt_Less_Discount.ReadOnly = True
        repoAmt_Less_Discount.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoAmt_Less_Discount)

        Dim repoTotal_Tax_Amt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTotal_Tax_Amt.FormatString = ""
        repoTotal_Tax_Amt.HeaderText = "Total Tax Amount"
        repoTotal_Tax_Amt.Name = AdcolTotal_Tax_Amt
        repoTotal_Tax_Amt.Width = 80
        repoTotal_Tax_Amt.ReadOnly = True
        repoTotal_Tax_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoTotal_Tax_Amt)


        Dim repoItem_Net_Amt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoItem_Net_Amt.FormatString = ""
        repoItem_Net_Amt.HeaderText = "Item Net Amount"
        repoItem_Net_Amt.Name = AdcolItem_Net_Amt
        repoItem_Net_Amt.Width = 80
        repoItem_Net_Amt.ReadOnly = True
        repoItem_Net_Amt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoItem_Net_Amt)

        Dim repoReceiptAdvance As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoReceiptAdvance.FormatString = "{0:n4}"
        repoReceiptAdvance.HeaderText = "Payment Advance Taxable"
        repoReceiptAdvance.Name = AdcolPaymentAdvance
        repoReceiptAdvance.Width = 80
        repoReceiptAdvance.ReadOnly = True
        repoReceiptAdvance.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoReceiptAdvance)

        Dim repoReceiptTotalTax As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoReceiptTotalTax.FormatString = "{0:n4}"
        repoReceiptTotalTax.HeaderText = "Payment Total Tax"
        repoReceiptTotalTax.Name = AdcolPaymentTotalTax
        repoReceiptTotalTax.Width = 80
        repoReceiptTotalTax.ReadOnly = True
        repoReceiptTotalTax.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoReceiptTotalTax)

        Dim repoReceiptTotalAdvanceAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoReceiptTotalAdvanceAmt.FormatString = "{0:n4}"
        repoReceiptTotalAdvanceAmt.HeaderText = "Payment Total Advance"
        repoReceiptTotalAdvanceAmt.Name = AdcolPaymentTotalAdvanceAmt
        repoReceiptTotalAdvanceAmt.Width = 80
        repoReceiptTotalAdvanceAmt.ReadOnly = True
        repoReceiptTotalAdvanceAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gvItem.MasterTemplate.Columns.Add(repoReceiptTotalAdvanceAmt)


        gvItem.AllowAddNewRow = False
        gvItem.ShowGroupPanel = False
        gvItem.AllowColumnReorder = True
        gvItem.AllowRowReorder = False
        gvItem.EnableSorting = False
        gvItem.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gvItem.MasterTemplate.ShowRowHeaderColumn = False
        gvItem.TableElement.TableHeaderHeight = 40

        MyBase.ReStoreGridLayoutMain(gvItem)
    End Sub



    Private Sub txtTaxGroupBankCharges__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtTaxGroupBankCharges._MYValidating
        If clsCommon.myCdbl(txtBankCharges.Text) <= 0 Then
            txtTaxGroupBankCharges.Value = ""
            txtTaxGroup_TxtChanged()
        End If

        GSTStatus = clsERPFuncationality.GetGSTStatus(dtpPayment.Value)
        If GSTStatus = True Then
            txtTaxGroupBankCharges.Value = clsLocationWiseTax.FinderForTaxGroupFinance("P", txtTaxGroupBankCharges.Value, isButtonClicked)
            txtTaxGroup_TxtChanged()
        Else
            Dim Qry As String = "select Tax_Group_Code as Code,Tax_Group_Desc as Description from TSPL_TAX_GROUP_MASTER "
            Dim WhrClause As String = "(  select count(TSPL_TAX_GROUP_DETAILS.Tax_Code)  from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_MASTER on "
            WhrClause += " TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code=TSPL_TAX_GROUP_MASTER.Tax_Group_Code "
            ''richa agarwal 26/06/2015 BM00000007205
            ' WhrClause += " and  TSPL_TAX_MASTER.Type = 'S'   )=(  select count(TSPL_TAX_GROUP_DETAILS.Tax_Code)  from TSPL_TAX_GROUP_DETAILS left outer join "
            WhrClause += "  )=(  select count(TSPL_TAX_GROUP_DETAILS.Tax_Code)  from TSPL_TAX_GROUP_DETAILS left outer join "
            ''-----------------------------------
            WhrClause += " TSPL_TAX_MASTER on TSPL_TAX_MASTER.Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code where "
            WhrClause += " TSPL_TAX_GROUP_DETAILS.Tax_Group_Code=TSPL_TAX_GROUP_MASTER.Tax_Group_Code   ) and Tax_Group_Type='P'"
            txtTaxGroupBankCharges.Value = clsCommon.ShowSelectForm("TaxGrpSFND", Qry, "Code", WhrClause, txtTaxGroupBankCharges.Value, "Code", isButtonClicked)
            txtTaxGroup_TxtChanged()
        End If

    End Sub
    Private Sub txtTaxGroup_TxtChanged()
        If Not IsInsideLoadData AndAlso iStxtTaxGroup_TxtChangedComplete Then
            iStxtTaxGroup_TxtChangedComplete = False
            LoadBlankGridBankChargeTax()
            Dim qry As String = "select TSPL_TAX_GROUP_DETAILS.Tax_Group_Code ,TSPL_TAX_GROUP_MASTER.Tax_Group_Desc,Tax_Code,Tax_Code_Desc,Surtax,Surtax_Tax_Code,(select Tax_Rate from TSPL_TAX_RATES WHERE Tax_Rate_Code=1 AND Tax_Code=TSPL_TAX_GROUP_DETAILS.Tax_Code and TSPL_TAX_RATES.Tax_Type='P') AS TaxRate,Taxable from TSPL_TAX_GROUP_DETAILS left outer join TSPL_TAX_GROUP_MASTER on TSPL_TAX_GROUP_MASTER.Tax_Group_Code=TSPL_TAX_GROUP_DETAILS.Tax_Group_Code where TSPL_TAX_GROUP_DETAILS.Tax_Group_Code='" + txtTaxGroupBankCharges.Value + "' and TSPL_TAX_GROUP_MASTER.Tax_Group_Type='P' and TSPL_TAX_GROUP_DETAILS.Tax_Group_Type='P' order by Trans_Code"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                If (dt.Rows.Count > 10) Then
                    MessageBox.Show("Can't Handle More than 10 Tax Types in a Group")
                    Return
                End If
                Dim ii As Integer = 0
                txtTaxGroupBankCharges.Value = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Code"))
                lblBankChargTaxGroup.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
                For Each dr As DataRow In dt.Rows
                    gv2.Rows.AddNew()
                    gv2.Rows(ii).Cells(colBCaxAutCode).Value = clsCommon.myCstr(dr("Tax_Code"))
                    gv2.Rows(ii).Cells(colBCaxAutName).Value = clsCommon.myCstr(dr("Tax_Code_Desc"))
                    If rbtnTaxCalAutomatic.IsChecked Then
                        gv2.Rows(ii).Cells(colBCaxRate).Value = clsCommon.myCdbl(dr("TaxRate"))
                    ElseIf rbtnTaxCalManual.IsChecked Then
                        gv2.Rows(ii).Cells(colBCaxRate).Value = Nothing
                    End If
                    ii = ii + 1
                Next
                SetitemWiseTaxSettingBankCharges(True, False)
            Else
                lblTaxGrpName.Text = ""
            End If

            'For ii As Integer = 0 To gv1.Rows.Count - 1
            '    UpdateCurrentRowWihtRowNo(ii)
            'Next
            'UpdateAllTotals()
            iStxtTaxGroup_TxtChangedComplete = True
        End If
    End Sub
    Sub LoadBlankGridBankChargeTax()
        gv2.Rows.Clear()
        gv2.Columns.Clear()

        Dim repoTaxAuthCode As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthCode.FormatString = ""
        repoTaxAuthCode.HeaderText = "Tax Code"
        repoTaxAuthCode.Name = colBCaxAutCode
        repoTaxAuthCode.Width = 150
        repoTaxAuthCode.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthCode)

        Dim repoTaxAuthName As GridViewTextBoxColumn = New GridViewTextBoxColumn()
        repoTaxAuthName.FormatString = ""
        repoTaxAuthName.HeaderText = "Tax Name"
        repoTaxAuthName.Name = colBCaxAutName
        repoTaxAuthName.Width = 200
        repoTaxAuthName.ReadOnly = True
        gv2.MasterTemplate.Columns.Add(repoTaxAuthName)

        Dim repoTaxBaseAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxBaseAmt.FormatString = "{0:n4}"
        repoTaxBaseAmt.HeaderText = "Base Amount"
        repoTaxBaseAmt.Name = colBCBaseAmt
        repoTaxBaseAmt.Width = 100
        repoTaxBaseAmt.ReadOnly = True
        repoTaxBaseAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxBaseAmt)

        Dim repoTaxRate As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxRate.FormatString = "{0:n2}"
        repoTaxRate.HeaderText = "Tax Rate"
        repoTaxRate.Name = colBCaxRate
        repoTaxRate.Width = 100
        repoTaxRate.ReadOnly = True
        repoTaxRate.IsVisible = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.ShowTaxRateColumnOnTransaction, clsFixedParameterCode.ShowTaxRateColumnOnTransaction, Nothing)) = 1, True, False)
        repoTaxRate.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxRate)

        Dim repoTaxAmt As GridViewDecimalColumn = New GridViewDecimalColumn()
        repoTaxAmt.FormatString = "{0:n4}"
        repoTaxAmt.HeaderText = "Tax Amount"
        repoTaxAmt.Name = colBCaxAmt
        repoTaxAmt.Width = 100
        repoTaxAmt.ReadOnly = False
        repoTaxAmt.TextAlignment = System.Drawing.ContentAlignment.MiddleRight
        gv2.MasterTemplate.Columns.Add(repoTaxAmt)

        gv2.AllowDeleteRow = True
        gv2.AllowAddNewRow = False
        gv2.ShowGroupPanel = False
        gv2.AllowColumnReorder = False
        gv2.AllowRowReorder = False
        gv2.EnableSorting = False
        gv2.AddNewRowPosition = Telerik.WinControls.UI.SystemRowPosition.Bottom
        gv2.MasterTemplate.ShowRowHeaderColumn = False

    End Sub
    Private Sub txtChangeVendorNo()
        gv2.Rows.Clear()
        If clsCommon.myLen(txtVendorCode.Value) <= 0 Then
            SetitemWiseTaxSettingBankCharges(True, False)
            Exit Sub
        End If
        If clsCommon.myCdbl(txtBankCharges.Text) <= 0 Then
            txtTaxGroupBankCharges.Value = ""
            lblBankChargTaxGroup.Text = ""
            SetitemWiseTaxSettingBankCharges(True, False)
            Exit Sub
        End If
        If Not IsInsideLoadData Then
            Dim qry As String = "select  Vendor_Code,Vendor_Name,Terms_Code,Terms_Code_Desc ,Vendor_Account ,Tax_Group,Tax_Group_Desc from TSPL_VENDOR_MASTER where Vendor_Code ='" & clsCommon.myCstr(txtVendorCode.Value) & "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry)
            If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
                txtTaxGroupBankCharges.Value = clsCommon.myCstr(dt.Rows(0)("Tax_Group"))
                lblBankChargTaxGroup.Text = clsCommon.myCstr(dt.Rows(0)("Tax_Group_Desc"))
            Else
                txtTaxGroupBankCharges.Value = ""
                lblBankChargTaxGroup.Text = ""
            End If
            SetitemWiseTaxSettingBankCharges(True, False)
        End If

    End Sub
    Sub SetitemWiseTaxSettingBankCharges(ByVal isChangeRate As Boolean, ByVal isForCurrentRow As Boolean)
        Dim dt As DataTable = clsTaxGroupMaster.GetTaxDetails(txtTaxGroupBankCharges.Value)
        If (dt IsNot Nothing AndAlso dt.Rows.Count > 0) Then
            If isForCurrentRow Then
                BlankTaxDetails(gv2.CurrentRow.Index)
                Dim ii As Integer = 1
                For Each dr As DataRow In dt.Rows
                    Dim strII As String = clsCommon.myCstr(ii)
                    gv2.CurrentRow.Cells(colBCaxAutCode).Value = clsCommon.myCstr(dr("Tax_Code"))
                    gv2.CurrentRow.Cells(colBCaxAutName).Value = clsCommon.myCstr(dr("Tax_Code_Desc"))
                    If isChangeRate Then
                        gv2.CurrentRow.Cells(colBCaxRate).Value = clsCommon.myCdbl(dr("TaxRate"))
                    End If
                    'gv1.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                    'gv1.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                    'gv1.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                    'gv1.CurrentRow.Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                    ii = ii + 1
                Next
            Else
                For intRowNo As Integer = 0 To gv2.Rows.Count - 1
                    Dim ii As Integer = 1
                    For Each dr As DataRow In dt.Rows
                        If clsCommon.CompairString(gv2.Rows(intRowNo).Cells(colBCaxAutCode).Value, clsCommon.myCstr(dr("Tax_Code"))) = CompairStringResult.Equal Then
                            Dim strII As String = clsCommon.myCstr(ii)
                            If isChangeRate Then
                                gv2.Rows(intRowNo).Cells(colBCaxRate).Value = clsCommon.myCdbl(dr("TaxRate"))
                            End If
                        End If

                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Taxable")), "Y") = CompairStringResult.Equal)
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Surtax")), "Y") = CompairStringResult.Equal)
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = clsCommon.myCstr(dr("Surtax_Tax_Code"))
                        'gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + strII)).Value = clsCommon.myCBool(clsCommon.CompairString(clsCommon.myCstr(dr("Excisable")), "Y") = CompairStringResult.Equal)
                        ii = ii + 1
                    Next
                Next
            End If
        End If
        UpdateBankChargeTax()
    End Sub
    Private Sub BlankTaxDetails(ByVal intRowNo As Integer)
        For ii As Integer = 1 To 10
            Dim strII As String = clsCommon.myCstr(ii)
            If intRowNo < 0 Then
                gv2.CurrentRow.Cells(colBCaxAutCode).Value = Nothing
                gv2.CurrentRow.Cells(colBCaxAutName).Value = Nothing
                gv2.CurrentRow.Cells(colBCBaseAmt).Value = Nothing
                gv2.CurrentRow.Cells(colBCaxRate).Value = Nothing
                gv2.CurrentRow.Cells(colBCaxAmt).Value = Nothing
                'gv2.CurrentRow.Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                'gv2.CurrentRow.Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                'gv2.CurrentRow.Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            Else
                gv2.Rows(intRowNo).Cells(colBCaxAutCode).Value = Nothing
                gv2.Rows(intRowNo).Cells(colBCaxAutName).Value = Nothing
                gv2.Rows(intRowNo).Cells(colBCBaseAmt).Value = Nothing
                gv2.Rows(intRowNo).Cells(colBCaxRate).Value = Nothing
                gv2.Rows(intRowNo).Cells(colBCaxAmt).Value = Nothing
                'gv2.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + strII)).Value = Nothing
                'gv2.Rows(intRowNo).Cells(clsCommon.myCstr("ISSURTAX" + strII)).Value = Nothing
                'gv2.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + strII)).Value = Nothing
            End If
        Next
    End Sub
    Private Sub UpdateBankChargeTax()
        'Dim arrTaxableAuth As New List(Of String)
        Dim isUnClaimedTax As Boolean = False 'clsCommon.myCBool(gv1.Rows(intRowNo).Cells(colIsUnclaimedTax).Value)
        'Dim dblRate As Double = clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(colRate).Value)
        'Dim dblAmt As Double = dblRate ''+ dblFAmt
        'gv1.Rows(intRowNo).Cells(colAmt).Value = Math.Round(dblAmt, 2)
        'Dim dblDisPer As Double = clsCommon.myCdbl(gv1.Rows(intRowNo).Cells(colDisPer).Value)
        'Dim dblDisAmt As Double = (dblAmt * dblDisPer) / 100
        'Dim dblAmtAfterDis As Double = dblAmt - dblDisAmt
        'Dim dblAbatementRate As Double = gv1.Rows(intRowNo).Cells(colAbatementPer).Value
        'Dim dblAbatementAmt As Double = ((dblRate * dblAbatementRate) / 100)
        'gv1.Rows(intRowNo).Cells(colAbatementAmount).Value = Math.Round(dblAbatementAmt, 2)

        'Dim dblReverseRate As Double = gv1.Rows(intRowNo).Cells(colReverserChargePer).Value
        Dim dblReverseAmount As Double = 0 '((dblAbatementAmt * dblReverseRate) / 100)
        'gv1.Rows(intRowNo).Cells(colReverserChargeAmount).Value = Math.Round(dblReverseAmount, 2)

        Dim dblBasicAmt As Double = clsCommon.myCdbl(txtBankCharges.Text)
        ''--------------------------
        For Each grow As GridViewRowInfo In gv2.Rows
            If rbtnTaxCalAutomatic.IsChecked Then
                Dim strTaxCode As String = clsCommon.myCstr(grow.Cells(colBCaxAutCode).Value)
                If clsCommon.myLen(strTaxCode) > 0 Then
                    Dim dblTaxRate As Double = clsCommon.myCdbl(grow.Cells(colBCaxRate).Value)
                    Dim IsSurTax As Boolean = False
                    Dim strSurTaxCode As String = "" ''clsCommon.myCstr(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("SURTAXCODE" + Strii)).Value)
                    Dim IsTaxable As Boolean = False 'clsCommon.myCBool(gv1.Rows(intRowNo).Cells(clsCommon.myCstr("ISTAXABLE" + Strii)).Value)
                    'Dim IsExcisable As Boolean = clsCommon.myCBool(gv1.Rows(IntRowNo).Cells(clsCommon.myCstr("ISEXCISABLE" + Strii)).Value)
                    Dim dblBaseAmt As Double = 0
                    Dim dblTaxAmt As Double = 0
                    If IsSurTax Then
                        Dim dblSurTaxAmt As Double = 0 'GetCurrentRowSurTaxAmt(intRowNo, ii, strSurTaxCode)
                        dblBaseAmt = dblSurTaxAmt
                    Else
                        Dim dblOtherTaxAmt As Double = 0
                        dblOtherTaxAmt = 0 'GetCurrentRowOtherTaxAmt(intRowNo, Strii, arrTaxableAuth)                        
                        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select isnull(Type,'') from TSPL_TAX_MASTER where Tax_Code ='" & strTaxCode & "' ")), "S") = CompairStringResult.Equal Then
                            dblBaseAmt = (dblReverseAmount + dblOtherTaxAmt)
                        Else
                            dblBaseAmt = dblBasicAmt
                        End If
                        'dblBaseAmt = (dblReverseAmount + dblOtherTaxAmt)
                        ''---------------------

                        ''End If
                    End If

                    If isUnClaimedTax Then
                        dblBaseAmt = 0
                    End If
                    grow.Cells(colBCBaseAmt).Value = dblBaseAmt
                    dblTaxAmt = (dblBaseAmt * dblTaxRate) / 100
                    grow.Cells(colBCaxAmt).Value = dblTaxAmt

                Else
                    grow.Cells(colBCaxAutCode).Value = ""
                    grow.Cells(colBCaxAutName).Value = ""
                    grow.Cells(colBCaxRate).Value = 0
                    grow.Cells(colBCBaseAmt).Value = 0
                    grow.Cells(colBCaxAmt).Value = 0

                End If
            ElseIf rbtnTaxCalManual.IsChecked Then
                'If gv2.Rows.Count >= ii Then
                '    Dim dblTaxAmt As Double = clsCommon.myCdbl(gv2.Rows(ii - 1).Cells(colTTaxAmt).Value)
                '    Dim dblTaxRate As Double = clsCommon.myCdbl(gv2.Rows(ii - 1).Cells(colTTaxRate).Value)
                '    Dim dblCurrRowAmt As Double = clsCommon.myCdbl(gv1.Rows(clsCommon.myCdbl(intRowNo)).Cells(colAmt).Value)
                '    Dim dblTotAmt As Double = 0
                '    For jj As Integer = 0 To gv1.Rows.Count - 1
                '        dblTotAmt += clsCommon.myCdbl(gv1.Rows(jj).Cells(colAmt).Value)
                '    Next
                '    Dim dblCurrCalTax As Double = 0
                '    If dblTotAmt <> 0 Then
                '        dblCurrCalTax = Math.Round(clsCommon.myCdbl(dblTaxAmt * dblCurrRowAmt / dblTotAmt), 2, MidpointRounding.ToEven)
                '    End If
                '    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("colTaxAmt" + Strii)).Value = dblCurrCalTax
                '    gv1.Rows(intRowNo).Cells(clsCommon.myCstr("COLTAXRATE" + Strii)).Value = dblTaxRate
                'End If
            End If
        Next

    End Sub

    Private Sub txtBankCharges_LostFocus(sender As Object, e As EventArgs) Handles txtBankCharges.LostFocus
        UpdateBankChargeTax()
        If TagExemptedtaxgroupincaseofBankChargesinPaymentEntry = True Then
            txtTaxGroupBankCharges.Value = clsDBFuncationality.getSingleValue("select Tax_Group_Code from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='EXEMPTED' and Tax_Group_Type='P'")
            lblBankChargTaxGroup.Text = clsDBFuncationality.getSingleValue("select Tax_Group_Code from TSPL_TAX_GROUP_MASTER where Tax_Group_Code='EXEMPTED' and Tax_Group_Type='P'")
            txtTaxGroup_TxtChanged()
        End If
    End Sub

    Private Sub gv2_CellFormatting(sender As Object, e As CellFormattingEventArgs) Handles gv2.CellFormatting
        Try
            If e.Column.Index >= 0 Then
                If (e.Column Is gv2.Columns(colBCaxAmt)) Then
                    gv2.CurrentRow.Cells(colBCaxAmt).ReadOnly = rbtnTaxCalAutomatic.IsChecked
                ElseIf (e.Column Is gv2.Columns(colBCaxRate)) Then
                    gv2.CurrentRow.Cells(colBCaxRate).ReadOnly = rbtnTaxCalAutomatic.IsChecked
                End If

            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub gv2_CellValueChanged(sender As Object, e As GridViewCellEventArgs) Handles gv2.CellValueChanged
        Try
            If (Not IsInsideLoadData) Then
                If Not isCellValueChangedTaxOpen Then
                    isCellValueChangedTaxOpen = True
                    If (e.Column Is (gv2.Columns(colTTaxAmt)) AndAlso rbtnTaxCalManual.IsChecked) Then
                        SetitemWiseTaxSettingBankCharges(True, False)
                    End If
                    isCellValueChangedTaxOpen = False
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub gv2_DoubleClick(sender As Object, e As EventArgs) Handles gv2.DoubleClick
        Try
            If rbtnTaxCalAutomatic.IsChecked Then
                Dim qry As String = "select Tax_Rate_Code as [Rate Code],Tax_Rate_Desc as [Rate Description],Tax_Rate as [Rate] from TSPL_TAX_RATES "
                Dim dblNewRate As Double = clsCommon.myCdbl(clsCommon.ShowSelectForm("FndVendorTaxRateIO", qry, "Rate", "Tax_Code='" + clsCommon.myCstr(gv2.CurrentRow.Cells(colBCaxAutCode).Value) + "' and Tax_Type='P'", "", "", True))
                Dim intRowNo As Integer = gv2.CurrentRow.Index
                If gv2.RowCount > 0 AndAlso intRowNo >= 0 Then
                    gv2.CurrentRow.Cells(colBCaxRate).Value = dblNewRate
                    UpdateBankChargeTax()
                End If
            End If
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub


    Private Sub loadEmployeeType()
        Dim dt As DataTable = Nothing
        dt = New DataTable
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Value", GetType(String))

        dt.Rows.Add("Select", "")
        dt.Rows.Add("TA-DA", "TD")
        dt.Rows.Add("Salary", "S")
        dt.Rows.Add("Travelling", "T")
        dt.Rows.Add("Imprest", "I")
        ddlEmployeeType.DataSource = dt
        ddlEmployeeType.DisplayMember = "Code"
        ddlEmployeeType.ValueMember = "Value"

    End Sub

    Public Sub loadEmployeeAdvanceType()
        Dim dt As DataTable = Nothing
        dt = New DataTable
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Value", GetType(String))
        dt.Rows.Add("Select", "")
        dt.Rows.Add("Advance Against Salary", "S")
        dt.Rows.Add("Advance Against Travel", "T")
        dt.Rows.Add("Advance Against Imprest", "I")
        ddlEmployeeAdvanceType.DataSource = dt
        ddlEmployeeAdvanceType.DisplayMember = "Code"
        ddlEmployeeAdvanceType.ValueMember = "Value"
    End Sub

    Private Sub ddlEmployeeType_SelectedValueChanged(sender As Object, e As EventArgs) Handles ddlEmployeeType.SelectedValueChanged
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
            If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                LoadVendorInvoices(txtVendorCode.Value)
            End If
        Else
            ddlEmployeeType.Enabled = False
            'loadEmployeeType()
        End If
    End Sub

    Private Sub ddlEmployeeAdvanceType_SelectedValueChanged(sender As Object, e As EventArgs) Handles ddlEmployeeAdvanceType.SelectedValueChanged
        If clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "RC") = CompairStringResult.Equal Then
            If clsCommon.CompairString(ddlEmployeeAdvanceType.SelectedValue, "S") = CompairStringResult.Equal Then
                ChkAdvSalary.Checked = True
                ChkAdvSalary.Enabled = False
            Else
                ChkAdvSalary.Checked = False
                ChkAdvSalary.Enabled = True
            End If
        Else
            ddlEmployeeAdvanceType.Enabled = False
            'loadEmployeeAdvanceType() ''Balwinder on 28/06/2022
        End If
    End Sub
    ''richa BHA/28/08/18-000491
    Private Sub chkBankChargesWaveOff_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkBankChargesWaveOff.ToggleStateChanged
        Try
            If ApplyBankChargesasperSlabonBankMaster Then
                If chkBankChargesWaveOff.Checked = True Then
                    txtBankCharges.Value = 0
                Else
                    BankChargerasperSlabWise()
                End If
                txtBankCharges.Enabled = False
            Else
                txtBankCharges.Enabled = True
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub
    ' Ticket No : TEC/07/05/19-000475
    Private Sub btnOpenBankCashBook_Click(sender As Object, e As EventArgs) Handles btnOpenBankCashBook.Click
        clsOpenBankCashBook.ShowBankCashBookDatails(txtPaymentNo.Value)
    End Sub

    Private Sub txtMPAdv__MYValidating(sender As Object, e As EventArgs, isButtonClicked As Boolean) Handles txtMPAdv._MYValidating
        Try
            If clsCommon.myLen(txtBankCode.Value) <= 0 Then
                Throw New Exception("Please First select Bank")
            End If
            txtMPAdv.Value = clsMpMaster.getFinder("", txtMPAdv.Value, isButtonClicked)

            Dim BankSeg As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(" SELECT Loc_Segment_Code  FROM TSPL_LOCATION_MASTER WHERE LOCATION_CODE IN ( SELECT MCC  FROM TSPL_VLC_MASTER_HEAD WHERE VLC_Code IN (select VLC_Code  from TSPL_MP_MASTER WHERE MP_Code ='" + txtMPAdv.Value + "'))"))

            'changes done by richa as per Ranjana Mam
            Dim dblFarmerCount As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(isFarmer) from tspl_vendor_account_set where isFarmer=1"))
            If dblFarmerCount <= 0 Then
                Throw New Exception("Please set Farmer Account Set as Farmer Type")
            End If
            Dim qry As String = clsPaymentHeader.GetFarmerAccountQry(chkFarmerLoanPayment.Checked)

            qry = clsDBFuncationality.getSingleValue(qry)
            If clsCommon.myLen(qry) > 0 Then
                IsCellValueChanged = True
                Try
                    gvDetails.CurrentRow.Cells(colGLAccount).Value = clsERPFuncationality.ChangeGLAccountLocationSegment(qry, BankSeg, True, Nothing)
                    gvDetails.CurrentRow.Cells(colAccDesc).Value = clsGLAccount.GetName(clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLAccount).Value))
                    If clsCommon.myLen(gvDetails.CurrentRow.Cells(colGLAccount).Value) > 0 Then
                        gvDetails.CurrentRow.Cells(colGLType).Value = clsPaymentHeader.CheckGLAccountType(clsCommon.myCstr(gvDetails.CurrentRow.Cells(colGLAccount).Value), Nothing)
                    Else
                        gvDetails.CurrentRow.Cells(colGLType).Value = ""
                    End If
                Catch ex As Exception
                Finally
                    IsCellValueChanged = False
                End Try
            Else
                If chkFarmerLoanPayment.Checked Then
                    Throw New Exception("Please set Loan Advance Account into Farmer Account set for Farmer Type")
                Else
                    Throw New Exception("Please set Advance Account into Farmer Account set for Farmer Type")
                End If

            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try
    End Sub

    Private Sub chkFarmerLoanPayment_ToggleStateChanged(sender As Object, args As StateChangedEventArgs) Handles chkFarmerLoanPayment.ToggleStateChanged
        txtMPAdv.Value = ""
        LoadBlankGrid("MI")
        gvDetails.Rows.AddNew()
    End Sub

    Private Sub butCostCenterAndHirerachy_Update_AfterPost_Click(sender As Object, e As EventArgs) Handles butCostCenterAndHirerachy_Update_AfterPost.Click
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim strJEStatus As String = clsDBFuncationality.getSingleValue("select Authorized from TSPL_JOURNAL_MASTER where Source_Doc_No ='" + txtPaymentNo.Value + "' ", trans)
            If clsCommon.CompairString(strJEStatus, "A") = CompairStringResult.Equal Then
                clsDBFuncationality.ExecuteNonQuery("ALTER TABLE TSPL_JOURNAL_DETAILS DISABLE TRIGGER TRG_JD_FiscaYearEndNoUpdateNoDelete", trans)
            End If

            For Each grow As GridViewRowInfo In gvDetails.Rows
                Dim coll As New Hashtable()
                If clsCommon.myLen(clsCommon.myCstr(grow.Cells(colGLAccount).Value)) > 0 Then

                    Dim strGLAccountCode As String = clsCommon.myCstr(grow.Cells(colGLAccount).Value)
                    clsCommon.AddColumnsForChange(coll, "Hirerachy_Level_Code", clsCommon.myCstr(grow.Cells(colHirerachyCenter).Value), True)
                    clsCommon.AddColumnsForChange(coll, "Cost_Center_Fin_Code", clsCommon.myCstr(grow.Cells(colCostCenter).Value), True)
                    clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PAYMENT_DETAIL", OMInsertOrUpdate.Update, "Payment_No='" + txtPaymentNo.Value + "' and Account_Code = '" + strGLAccountCode + "'", trans)
                    Dim strVoucherNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select Voucher_No  from TSPL_JOURNAL_MASTER where Source_Doc_No = '" + txtPaymentNo.Value + "' ", trans))
                    Dim qry As String = "update TSPL_JOURNAL_DETAILS SET Hirerachy_Code='" + clsCommon.myCstr(grow.Cells(colHirerachyCenter).Value) + "',Cost_Centre_Code='" + clsCommon.myCstr(grow.Cells(colCostCenter).Value) + "' WHERE Voucher_No='" + strVoucherNo + "' and Account_code='" + strGLAccountCode + "' "
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

    Private Sub txtmulPI__My_Click(sender As Object, e As EventArgs) Handles txtmulPI._My_Click
        Try
            If clsCommon.myLen(txtVendorCode.Value) <= 0 Then
                RadMessageBox.Show("Please Select Vendor")
                txtVendorCode.Focus()
                Exit Sub
            End If
            Dim qry As String = " select  DISTINCT TSPL_PI_HEAD.PI_No as PurchaseInvoice ,PI_Date as Date,TSPL_PI_HEAD.Vendor_Code as [Vendor Code],TSPL_PI_HEAD.Vendor_Name as [Vendor Name],TSPL_PI_DETAIL.ITEM_CODE,TSPL_PURCHASE_ORDER_HEAD.RefTendorNo AS RAL,TSPL_PI_HEAD.Against_PO as PO from TSPL_PI_DETAIL
                                LEFT OUTER JOIN TSPL_PI_HEAD ON TSPL_PI_HEAD.PI_No=TSPL_PI_DETAIL.PI_No
                                left outer join TSPL_PURCHASE_ORDER_HEAD on TSPL_PURCHASE_ORDER_HEAD.PurchaseOrder_No=TSPL_PI_HEAD.Against_PO
                                where TSPL_PI_HEAD.Status=1 and TSPL_PI_HEAD.Vendor_Code='" + txtVendorCode.Value + "'  "
            txtmulPI.arrValueMember = clsCommon.ShowMultipleSelectForm("ItemMulSel", qry, "PurchaseInvoice", "", txtmulPI.arrValueMember, txtmulPI.arrDispalyMember)
            If txtmulPI.arrValueMember IsNot Nothing AndAlso txtmulPI.arrValueMember.Count > 0 Then
                LoadBlankGrid(ddlPaymentType.SelectedValue)
                For jj As Integer = 0 To txtmulPI.arrValueMember.Count - 1
                    LOADPI(clsCommon.myCstr(txtmulPI.arrValueMember(jj)))
                Next
            End If
            'LOADPI()
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub
    Sub loadgrid(ByVal strcode As String)
        Try
            Dim PaymentAmount As Double = 0
            'For Each grow As GridViewRowInfo In gvDetails.Rows
            '    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Debit Note") = CompairStringResult.Equal Then
            '        PaymentAmount -= clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
            '    Else
            '        PaymentAmount += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value) + clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
            '    End If
            'Next



            For Each grow As GridViewRowInfo In gvDetails.Rows
                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colPI_No).Value), strcode) = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Debit Note") = CompairStringResult.Equal Then
                        PaymentAmount -= clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Credit Note") = CompairStringResult.Equal Then
                        PaymentAmount += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value) + clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
                    End If
                End If
            Next
            'txtPaymentAmt.Text = PaymentAmount

            'For Each grow As GridViewRowInfo In gvDetails.Rows
            '    ' Skip rows where the current column is one of the excluded columns
            '    If Not (grow.Cells(colDocNo).ColumnInfo Is colDocNo OrElse grow.Cells(colPINo).ColumnInfo Is colPINo) Then
            '        If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Debit Note") = CompairStringResult.Equal Then
            '            PaymentAmount -= clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
            '        ElseIf clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Credit Note") = CompairStringResult.Equal Then
            '            PaymentAmount += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value) + clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
            '        End If
            '    End If
            'Next


            'For Each grow As GridViewRowInfo In gvDetails.Rows
            '    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colPI_No).Value), strcode) = CompairStringResult.Equal Then
            '        If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Invoice") = CompairStringResult.Equal Then
            '            grow.Cells(colAdjustedAmt).Value = PaymentAmount
            '            If PaymentAmount > 0 Then
            '                grow.Cells(colAppliedAmt).Value = grow.Cells(colNetAmt).Value - grow.Cells(colAdjustedAmt).Value
            '            Else
            '                grow.Cells(colAppliedAmt).Value = grow.Cells(colNetAmt).Value + grow.Cells(colAdjustedAmt).Value
            '            End If
            '            txtPaymentAmt.Text = grow.Cells(colAppliedAmt).Value
            '        End If
            '    End If
            'Next
            For Each grow As GridViewRowInfo In gvDetails.Rows
                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colPI_No).Value), strcode) = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Invoice") = CompairStringResult.Equal Then
                        If PaymentAmount < 0 Then
                            grow.Cells(colAdjustedAmt).Value = PaymentAmount * -1
                        End If
                    End If
                End If
            Next
            Dim payment As Double = 0
            For Each grow As GridViewRowInfo In gvDetails.Rows
                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Invoice") = CompairStringResult.Equal Then
                    payment += grow.Cells(colAppliedAmt).Value
                    txtPaymentAmt.Text = grow.Cells(colAppliedAmt).Value
                End If
            Next
            Dim payamt As Double = 0
            For Each grow As GridViewRowInfo In gvDetails.Rows
                If clsCommon.CompairString(clsCommon.myCstr(grow.Cells(colDocType).Value), "Debit Note") = CompairStringResult.Equal Then
                    payamt -= clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value)
                Else
                    payamt += clsCommon.myCdbl(grow.Cells(colAppliedAmt).Value) + clsCommon.myCdbl(grow.Cells(colSecurityAmt).Value)
                End If
            Next
            txtPaymentAmt.Text = payamt
            'gvDetails.CurrentRow.Cells(colPendingAmt).Value = gvDetails.CurrentRow.Cells(colTemp).Value - (gvDetails.CurrentRow.Cells(colAppliedAmt).Value + gvDetails.CurrentRow.Cells(colSecurityAmt).Value)
        Catch ex As Exception
            common.clsCommon.MyMessageBoxShow(Me, ex.Message, Me.Text)
        End Try

    End Sub

    '    gvDetails.CurrentRow.Cells(colPendingAmt).Value = gvDetails.CurrentRow.Cells(colTemp).Value - (gvDetails.CurrentRow.Cells(colAppliedAmt).Value + gvDetails.CurrentRow.Cells(colSecurityAmt).Value)

    Sub LOADPI(ByVal strcode As String)
        Try

            Dim whr As String = " "
            Dim str As String = ""

            If txtmulPI.arrValueMember IsNot Nothing AndAlso txtmulPI.arrValueMember.Count <= 0 Then
                RadMessageBox.Show("Please select at least one invoice ")
                txtmulPI.Focus()
                Exit Sub
            End If

            Dim query As String = "Select NEW.FINAL_PI,FINALQRY.DocType,(FINALQRY.Document_No)Document_No,FINALQRY.PurchaseInvoice,(FINALQRY.DocumentDate)DocumentDate,(FINALQRY.DocDate)DocDate,(FINALQRY.VendorInvoiceNo)VendorInvoiceNo,(FINALQRY.TDSAmt)TDSAmt,(FINALQRY.[OriginalAmt])[OriginalAmt],(FINALQRY.NetAmount)NetAmount,(FINALQRY.PendingAmt)PendingAmt,(FINALQRY.ConvRate)ConvRate,(FINALQRY.ConvRateRevaluation)ConvRateRevaluation    from ( select Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType],   TSPL_VENDOR_INVOICE_HEAD.Document_No, Case When ISNULL(Against_POInvoice_No,'')<>'' Then Against_POInvoice_No When ISNULL(Against_PurchaseReturn_No,'')<> '' Then Against_PurchaseReturn_No Else Document_No End as PurchaseInvoice, Case When ISNULL(Against_POInvoice_No,'')<>'' Then (Select convert(varchar,PI_Date,103)  from TSPL_PI_HEAD where PI_No =Against_POInvoice_No)  When ISNULL(Against_PurchaseReturn_No,'')<> '' Then (Select convert(varchar,PR_Date,103) from TSPL_PR_HEAD where PR_No  =Against_PurchaseReturn_No ) Else TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date End as DocumentDate,  TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date as [DocDate] ,   TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No as [VendorInvoiceNo],  TSPL_VENDOR_INVOICE_HEAD.Document_Total  - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then  ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'   then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end)  as [OriginalAmt]  , TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount as [TDSAmt],  (TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end )  - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then  ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'   then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end)   as [NetAmount],  TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end  - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then  ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'   then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end +  case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y'  then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end)  - ISNULL((Select SUM(Applied_Amount) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No = TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No and isnull(TSPL_BANK_REVERSE.POST,'')='P' and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' )),0)  -isnull((select sum(isnull(Payment_Amount,0)) from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Applied_Payment =TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_HEADER.Payment_No and isnull(TSPL_BANK_REVERSE.POST,'')='P'  and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "') and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' and isnull(TSPL_PAYMENT_HEADER.Applied_Payment ,'')<>'' ),0)  
                         - ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No ) and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0)  - ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0)   -ISNULL((Select SUM(isnull(TSPL_PAYMENT_ADJUSTMENT_DETAIL.Amount,0)) from TSPL_PAYMENT_ADJUSTMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_ADJUSTMENT_DETAIL on TSPL_PAYMENT_ADJUSTMENT_Header.Adjustment_No=TSPL_PAYMENT_ADJUSTMENT_DETAIL.Adjustment_No Where isnull(TSPL_PAYMENT_ADJUSTMENT_Header.Doc_No,'') = isnull(TSPL_VENDOR_INVOICE_HEAD.Document_No,'') ),0)  as [PendingAmt],TSPL_VENDOR_INVOICE_HEAD.ConvRate  ,isnull(( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and TSPL_REVALUATION_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation  from TSPL_VENDOR_INVOICE_HEAD  Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code   WHERE TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,Invoice_Entry_Date,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "' And ((ISNULL(RefDocNo,'')= '' or ISNULL((select VI.Document_Type from TSPL_VENDOR_INVOICE_HEAD VI where VI.Document_No =TSPL_VENDOR_INVOICE_HEAD .RefDocNo ),'') in ('I','')) AND ISNULL(Against_PurchaseReturn_No,'')= '')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_Salary_Generation_Code,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='C')   "
            If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                query += "  and TSPL_VENDOR_INVOICE_HEAD.Loc_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
            End If
            query += " And TSPL_VENDOR_INVOICE_HEAD.RefDocType <>'REVALUATION ENTRY' and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')=''  and isnull(TSPL_VENDOR_INVOICE_HEAD.posting_date,'')<>''
                         And TSPL_VENDOR_INVOICE_HEAD.Document_No In (
                                                    select Document_No
                                                    from TSPL_VENDOR_INVOICE_HEAD where RefDocType in ('QC-DED','SCH-PNT','RE-TEN','REV-SPT','SEC-DED')  and        RefDocNo in ('" + strcode + "')  )
                                                    OR
                                                     TSPL_VENDOR_INVOICE_HEAD.Document_No IN 
                                                    (select Document_No
                                                    from TSPL_VENDOR_INVOICE_HEAD where  Document_Type='I'   and Against_POInvoice_No in ('" + strcode + "')    )
                                                    OR
                                                    TSPL_VENDOR_INVOICE_HEAD.Document_No in (
                                                    select Document_No
                                                    from TSPL_VENDOR_INVOICE_HEAD where RefDocType in('REV-SPT') and RefDocNo in (select Document_No from TSPL_VENDOR_INVOICE_HEAD where       RefDocNo in ('" + strcode + "')    ) 
                                                     )
                          union all 
                                                     select  'Receipt Note' as [DocType], Payment_No  as Document_No, Payment_No as [PurchaseInvoice],convert(varchar ,Payment_Date ,103) as [DocumentDate] ,convert(varchar,Payment_Date ,103) as [DocDate] ,'' as VendorInvoiceNo,Payment_Amount  as [OriginalAmt],TDS_Amount as TDSAmt ,Payment_Amount - TDS_Amount as NetAmount,
                                                     (Payment_Amount - TDS_Amount -isnull((select sum(isnull(Applied_Amount,0)) from TSPL_PAYMENT_DETAIL where TSPL_PAYMENT_DETAIL.Document_No=TSPL_PAYMENT_HEADER.Payment_No  and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No )),0) ) as [PendingAmt] 
                                                     , ConvRate ,1 as ConvRateRevaluation from TSPL_PAYMENT_HEADER WHERE Payment_Type  ='RC' and IsChkReverse ='N' AND Is_Security=0 AND ISNULL(TSPL_PAYMENT_HEADER.Applied_Payment  ,'')='' and 
                                                     TSPL_PAYMENT_HEADER.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,TSPL_PAYMENT_HEADER.Payment_Date ,103)<='" + clsCommon.GetPrintDate(dtpPayment.Value, "dd/MMM/yyyy") + "'  "
            If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                query += "  and TSPL_PAYMENT_HEADER.Location_GL_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
            End If
            query += ") FINALQRY 
						  LEFT OUTER JOIN 
						 (
                        select case when ab.Document_Type<>'C' then ab.NEWFIELD else abc.RefDocNo end FINAL_PI,AB.Document_No from (
                        select Against_POInvoice_No AS NEWFIELD,Against_POInvoice_No,Document_No,RefDocNo,Document_Type,RefDocType from TSPL_VENDOR_INVOICE_HEAD where Against_POInvoice_No IN ('" + strcode + "') and Document_Type='I'
                        union all
                        select RefDocNo AS NEWFIELD,Against_POInvoice_No,Document_No,RefDocNo,Document_Type,RefDocType from TSPL_VENDOR_INVOICE_HEAD where RefDocNo IN ('" + strcode + "') AND RefDocType in ('QC-DED','SCH-PNT','RE-TEN','SEC-DED') AND Document_Type='D'
                        UNION ALL

                        SELECT RefDocNo AS NEWFIELD,Against_POInvoice_No,Document_No,RefDocNo,Document_Type,RefDocType FROM TSPL_VENDOR_INVOICE_HEAD WHERE RefDocNo IN (SELECT Document_No FROM TSPL_VENDOR_INVOICE_HEAD WHERE RefDocNo IN ('" + strcode + "') )) ab
                        left outer join (

                        SELECT Document_No,RefDocNo,Document_Type,RefDocType  FROM TSPL_VENDOR_INVOICE_HEAD WHERE Document_No IN (select Against_POInvoice_No AS NEWFILED from TSPL_VENDOR_INVOICE_HEAD where Against_POInvoice_No IN ('" + strcode + "') and Document_Type='I'
                        union all
                        select RefDocNo AS NEWFIELD from TSPL_VENDOR_INVOICE_HEAD where RefDocNo IN ('" + strcode + "') AND  RefDocType in ('QC-DED','SCH-PNT','RE-TEN','SEC-DED') AND Document_Type='D'
                        UNION ALL

                        SELECT RefDocNo AS NEWFIELD FROM TSPL_VENDOR_INVOICE_HEAD WHERE RefDocNo IN (SELECT Document_No FROM TSPL_VENDOR_INVOICE_HEAD WHERE RefDocNo IN ('" + strcode + "') ))) abc
                        on abc.Document_No=ab.NEWFIELD
                        ) NEW ON NEW.Document_No=FINALQRY.Document_No
						 WHERE FINALQRY.PendingAmt>0  and FINALQRY.Document_No <>'" + txtDocumentNo.Value + "'    "
            dt = clsDBFuncationality.GetDataTable(query)
            IsInsideLoadData = True
            For Each dr As DataRow In dt.Rows
                gvDetails.Rows.AddNew()
                gvDetails.CurrentRow.Cells(colApply).Value = "Yes"
                gvDetails.CurrentRow.Cells(colDocType).Value = clsCommon.myCstr(dr("DocType"))
                gvDetails.CurrentRow.Cells(colPINo).Value = clsCommon.myCstr(dr("PurchaseInvoice"))
                gvDetails.CurrentRow.Cells(colDocNo).Value = clsCommon.myCstr(dr("Document_No"))
                If clsCommon.myCdbl(dr("ConvRateRevaluation")) > 0 Then
                    gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRateRevaluation"))
                Else
                    gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRate"))
                End If
                gvDetails.CurrentRow.Cells(colPI_No).Value = clsCommon.myCstr(dr("FINAL_PI"))
                gvDetails.CurrentRow.Cells(colAPTapalNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull (TapalNo,'') as TapalNo  from TSPL_VENDOR_INVOICE_HEAD where Document_No = '" + clsCommon.myCstr(dr("Document_No")) + "'"))
                gvDetails.CurrentRow.Cells(colDocDate).Value = clsCommon.myCstr(dr("DocDate"))
                gvDetails.CurrentRow.Cells(colDocumentDate).Value = clsCommon.myCstr(dr("DocumentDate"))
                gvDetails.CurrentRow.Cells(colVendorInvNo).Value = clsCommon.myCstr(dr("VendorInvoiceNo"))
                gvDetails.CurrentRow.Cells(colOriginalAmt).Value = clsCommon.myCdbl(dr("OriginalAmt"))
                gvDetails.CurrentRow.Cells(colTDSAmt).Value = clsCommon.myCstr(dr("TDSAmt"))
                gvDetails.CurrentRow.Cells(colNetAmt).Value = clsCommon.myCdbl(dr("NetAmount")) 
                ' gvDetails.CurrentRow.Cells(colPendingAmt).Value = clsCommon.myCdbl(dr("PendingAmt"))
                gvDetails.CurrentRow.Cells(colTemp).Value = clsCommon.myCdbl(dr("PendingAmt"))
                If gvDetails.CurrentRow.Cells(colApply).Value = "Yes" Then
                    gvDetails.CurrentRow.Cells(colAppliedAmt).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colPendingAmt).Value = 0
                End If
                ' gvDetails.CurrentRow.Cells(colAppliedAmt).Value = gvDetails.CurrentRow.Cells(colPendingAmt).Value
                'gvDetails.CurrentRow.Cells(colPendingAmt).Value = gvDetails.CurrentRow.Cells(colTemp).Value - (gvDetails.CurrentRow.Cells(colAppliedAmt).Value + gvDetails.CurrentRow.Cells(colSecurityAmt).Value)

                ' For Each grow As GridViewRowInfo In gvDetails.Rows    
                '  Next
                If clsCommon.CompairString(clsCommon.myCstr(gvDetails.CurrentRow.Cells(colDocType).Value), "Invoice") = CompairStringResult.Equal Then
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = False
                Else
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = True
                End If

                '' Anubhooti 14-Nov-2014 BM00000004636
                ' gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = gvDetails.CurrentRow.Cells(colNetAmt).Value - gvDetails.CurrentRow.Cells(colPendingAmt).Value
            Next
            If dt.Rows.Count > 0 Then
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
                    ddlPaymentType.Enabled = False
                    chkSaving.Enabled = False
                ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                    ddlPaymentType.Enabled = False
                    chkSaving.Enabled = False
                End If

                loadgrid(strcode)

            Else
                txtVendorCode.Enabled = True
                txtBankCode.Enabled = True
                ddlPaymentType.Enabled = True
                chkSaving.Enabled = True
            End If
            IsInsideLoadData = False

        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoadVendorInvoices(ByVal strVendorCode As String, ByVal InvoiceDate As Date, ByVal isForUpdate As Boolean)
        Try
            If Not isForUpdate Then
                LoadBlankGrid(ddlPaymentType.SelectedValue)
            End If

            Dim strQryForRejectedAmt As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
           "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_PurchaseReturn_No  in (SELECT PR_No  FROM TSPL_PR_HEAD WHERE Against_PI =TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No ) and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "

            '' this code is wrritten for that debit note which is created auto through PI and Pr is not created against that PI
            Dim strQryForRejectedAmtforNonPR As String = "- ISNULL(case when ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='I' then (Select sum (z.Document_Total)-sum (isnull(z.TaxAmount,0)) from ( select  isnull(Document_Total,0) as Document_Total,(case when inn.GSTRegistered =0 or inn.RCM=1 then  ( case when len(isnull(inn.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax1)='Y'  then inn.TAX1_Amt else 0 end +  case when len(isnull(inn.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX2)='Y'  then inn.TAX2_Amt else 0 end +  case when len(isnull(inn.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX3)='Y'  then inn.TAX3_Amt else 0 end +  case when len(isnull(inn.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX4)='Y'  then inn.TAX4_Amt else 0 end +  case when len(isnull(inn.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX5)='Y'  then inn.TAX5_Amt else 0 end +  case when len(isnull(inn.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX6)='Y'   then inn.TAX6_Amt else 0 end +  case when len(isnull(inn.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX7)='Y'  then inn.TAX7_Amt else 0 end +  case when len(isnull(inn.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX8)='Y'  then inn.TAX8_Amt else 0 end +  case when len(isnull(inn.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.TAX9)='Y'  then inn.TAX9_Amt else 0 end +  case when len(isnull(inn.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =inn.Tax10)='Y'  then inn.TAX10_Amt else 0 end ) else 0 end) as TaxAmount " &
          "from TSPL_VENDOR_INVOICE_HEAD as inn where inn.Against_POInvoice_No=TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No and inn.Document_Type='D' and inn.Vendor_Code=TSPL_VENDOR_INVOICE_HEAD.Vendor_Code )z ) else 0 end,0) "


            Dim strTaxRecovarableQuery As String = " - (case when TSPL_VENDOR_INVOICE_HEAD.GSTRegistered =0 or TSPL_VENDOR_INVOICE_HEAD.RCM=1 then " &
            " ( case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax1,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax1)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX1_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX2,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX2)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX2_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX3,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX3)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX3_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX4,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX4)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX4_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX5,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX5)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX5_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX6,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX6)='Y'  " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX6_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX7,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX7)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX7_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX8 ,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX8)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX8_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.TAX9,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.TAX9)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX9_Amt else 0 end + " &
            " case when len(isnull(TSPL_VENDOR_INVOICE_HEAD.Tax10,''))> 0 and (Select Tax_Recoverable from TSPL_TAX_MASTER where Tax_Code =TSPL_VENDOR_INVOICE_HEAD.Tax10)='Y' " &
            " then TSPL_VENDOR_INVOICE_HEAD.TAX10_Amt else 0 end ) else 0 end) "

            If clsCommon.myLen(strVendorCode) > 0 Then

                Qry = "Select * from ( select Case When TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' Then 'Debit Note' Else case  When TSPL_VENDOR_INVOICE_HEAD.Document_Type='C' Then 'Credit Note' Else 'Invoice' End End As [DocType],  " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_No, Case When ISNULL(Against_POInvoice_No,'')<>'' Then Against_POInvoice_No When ISNULL(Against_PurchaseReturn_No,'')<> '' Then Against_PurchaseReturn_No Else Document_No End as PurchaseInvoice," &
           " Case When ISNULL(Against_POInvoice_No,'')<>'' Then (Select convert(varchar,PI_Date,103)  from TSPL_PI_HEAD where PI_No =Against_POInvoice_No)  When ISNULL(Against_PurchaseReturn_No,'')<> '' Then (Select convert(varchar,PR_Date,103) from TSPL_PR_HEAD where PR_No  =Against_PurchaseReturn_No ) Else TSPL_VENDOR_INVOICE_HEAD.Invoice_Entry_Date End as DocumentDate, " &
           " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_Date as [DocDate] ,  " &
           " TSPL_VENDOR_INVOICE_HEAD.Vendor_Invoice_No as [VendorInvoiceNo], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total " + strTaxRecovarableQuery + " as [OriginalAmt]  ," &
           " TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount as [TDSAmt], " &
           " (TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end ) " + strTaxRecovarableQuery + "  as [NetAmount], " &
           " TSPL_VENDOR_INVOICE_HEAD.Document_Total - case when isnull(TSPL_VENDOR_INVOICE_HEAD .is_For_TDS,0) =1 and TSPL_VENDOR_INVOICE_HEAD.Document_Type ='D' then 0 else isnull(TSPL_VENDOR_INVOICE_HEAD.TDS_Actual_Amount,0) end " + strTaxRecovarableQuery + " - ISNULL((Select SUM(Applied_Amount) from TSPL_PAYMENT_DETAIL Where TSPL_PAYMENT_DETAIL.Document_No = TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No and isnull(TSPL_BANK_REVERSE.POST,'')='P' and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' ) ),0) " &
          " -isnull((select sum(isnull(Payment_Amount,0)) from TSPL_PAYMENT_HEADER where TSPL_PAYMENT_HEADER.Applied_Payment =TSPL_VENDOR_INVOICE_HEAD.Document_No and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_HEADER.Payment_No and isnull(TSPL_BANK_REVERSE.POST,'')='P'  and Convert(Date,TSPL_BANK_REVERSE.Reversal_Date ,103)<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "') and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D' and isnull(TSPL_PAYMENT_HEADER.Applied_Payment ,'')<>'' ),0)  " & Environment.NewLine &
          " " + strQryForRejectedAmt + " " + strQryForRejectedAmtforNonPR + " " &
           " -ISNULL((Select SUM(isnull(TSPL_PAYMENT_ADJUSTMENT_DETAIL.Amount,0)) from TSPL_PAYMENT_ADJUSTMENT_HEADER LEFT OUTER JOIN TSPL_PAYMENT_ADJUSTMENT_DETAIL on TSPL_PAYMENT_ADJUSTMENT_Header.Adjustment_No=TSPL_PAYMENT_ADJUSTMENT_DETAIL.Adjustment_No Where isnull(TSPL_PAYMENT_ADJUSTMENT_Header.Doc_No,'') = isnull(TSPL_VENDOR_INVOICE_HEAD.Document_No,'') ),0) " &
           " as [PendingAmt],TSPL_VENDOR_INVOICE_HEAD.ConvRate " &
           " ,isnull(( select top 1 TSPL_REVALUATION_HEAD.Currency_Rate from TSPL_REVALUATION_DETAIL left outer join TSPL_REVALUATION_HEAD on TSPL_REVALUATION_HEAD.Document_No=TSPL_REVALUATION_DETAIL.Document_No  where TSPL_REVALUATION_DETAIL.AP_Invoice_No=TSPL_VENDOR_INVOICE_HEAD.Document_No and TSPL_REVALUATION_HEAD.Status=1 and TSPL_REVALUATION_HEAD.Document_Date<='" + clsCommon.GetPrintDate(clsCommon.GetDateWithEndTime(dtpPayment.Value), "dd/MMM/yyyy hh:mm tt") + "' order by TSPL_REVALUATION_HEAD.Document_Date desc),0) as ConvRateRevaluation " &
           " from TSPL_VENDOR_INVOICE_HEAD " &
           " Left outer join TSPL_VENDOR_MASTER on TSPL_VENDOR_MASTER.Vendor_Code =TSPL_VENDOR_INVOICE_HEAD.Vendor_Code  " &
          " WHERE TSPL_VENDOR_INVOICE_HEAD.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,Invoice_Entry_Date,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' And ((ISNULL(RefDocNo,'')= '' or ISNULL((select VI.Document_Type from TSPL_VENDOR_INVOICE_HEAD VI where VI.Document_No =TSPL_VENDOR_INVOICE_HEAD .RefDocNo ),'') in ('I','')) AND ISNULL(Against_PurchaseReturn_No,'')= '')  and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_POInvoice_No,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='D') " &
          " and not ( ISNULL( TSPL_VENDOR_INVOICE_HEAD.Against_Salary_Generation_Code,'')<>'' and TSPL_VENDOR_INVOICE_HEAD.Document_Type='C') "
                If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                    Qry += "  and TSPL_VENDOR_INVOICE_HEAD.Loc_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
                End If



                If isForUpdate Then
                    Dim arrDocNo As New List(Of String)
                    For ii As Integer = 0 To gvDetails.Rows.Count - 1
                        arrDocNo.Add(clsCommon.myCstr(gvDetails.Rows(ii).Cells(colDocNo).Value))
                    Next
                    If arrDocNo.Count > 0 Then
                        Qry += " and TSPL_VENDOR_INVOICE_HEAD.Document_No not in (" + clsCommon.GetMulcallString(arrDocNo) + ")"
                    End If
                End If
                Qry += " and TSPL_VENDOR_INVOICE_HEAD.RefDocType<>'REVALUATION ENTRY'"

                ''richa 13/10/2017
                If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal OrElse clsCommon.CompairString(ddlPaymentType.SelectedValue, "AD") = CompairStringResult.Equal Then
                    If clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "TD") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='TD' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "T") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='T' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "S") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='S' "
                    ElseIf clsCommon.CompairString(clsCommon.myCstr(ddlEmployeeType.SelectedValue), "I") = CompairStringResult.Equal Then
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='I' "
                    Else
                        Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.Employee_Type,'')='' "
                    End If
                End If
                ''--------


                Qry += " and isnull(TSPL_VENDOR_INVOICE_HEAD.posting_date,'')<>''" & Environment.NewLine &
                "  union all " & Environment.NewLine &
                " select  'Receipt Note' as [DocType], Payment_No  as Document_No, Payment_No as [PurchaseInvoice],convert(varchar ,Payment_Date ,103) as [DocumentDate] ,convert(varchar,Payment_Date ,103) as [DocDate] ,'' as VendorInvoiceNo,Payment_Amount  as [OriginalAmt],TDS_Amount as TDSAmt ,Payment_Amount - TDS_Amount as NetAmount," & Environment.NewLine &
                " (Payment_Amount - TDS_Amount -isnull((select sum(isnull(Applied_Amount,0)) from TSPL_PAYMENT_DETAIL where TSPL_PAYMENT_DETAIL.Document_No=TSPL_PAYMENT_HEADER.Payment_No  and not exists (select 1 from TSPL_BANK_REVERSE where TSPL_BANK_REVERSE.Reverse_Document='Payments' and TSPL_BANK_REVERSE.Document_No=TSPL_PAYMENT_DETAIL.Payment_No )),0) ) as [PendingAmt] " & Environment.NewLine &
                " , ConvRate ,1 as ConvRateRevaluation from TSPL_PAYMENT_HEADER WHERE Payment_Type  ='RC' and IsChkReverse ='N' AND Is_Security=0 AND ISNULL(TSPL_PAYMENT_HEADER.Applied_Payment  ,'')='' and " & Environment.NewLine &
                " TSPL_PAYMENT_HEADER.Vendor_Code ='" + txtVendorCode.Value + "' AND Convert(Date,TSPL_PAYMENT_HEADER.Payment_Date ,103)<='" + clsCommon.GetPrintDate(InvoiceDate, "dd/MMM/yyyy") + "' "
                If objCommonVar.ApplyLocationFilterBasedOnPermission = True AndAlso clsCommon.myLen(objCommonVar.strCurrUserLocationsSegment) > 0 Then
                    Qry += "  and TSPL_PAYMENT_HEADER.Location_GL_Code in (" + objCommonVar.strCurrUserLocationsSegment + ")"
                End If
                Qry += ") FINALQRY WHERE FINALQRY.PendingAmt>0  and Document_No <>'" & txtDocumentNo.Value & "' ORDER BY Convert(Date,FINALQRY.DocDate,103)  "
                dt = clsDBFuncationality.GetDataTable(Qry)
                IsInsideLoadData = True
                For Each dr As DataRow In dt.Rows
                    gvDetails.Rows.AddNew()
                    gvDetails.CurrentRow.Cells(colApply).Value = "No"
                    gvDetails.CurrentRow.Cells(colDocType).Value = clsCommon.myCstr(dr("DocType"))
                    gvDetails.CurrentRow.Cells(colPINo).Value = clsCommon.myCstr(dr("PurchaseInvoice"))
                    gvDetails.CurrentRow.Cells(colDocNo).Value = clsCommon.myCstr(dr("Document_No"))
                    If clsCommon.myCdbl(dr("ConvRateRevaluation")) > 0 Then
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRateRevaluation"))
                    Else
                        gvDetails.CurrentRow.Cells(colDocNo).Tag = clsCommon.myCdbl(dr("ConvRate"))
                    End If

                    gvDetails.CurrentRow.Cells(colAPTapalNo).Value = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select isnull (TapalNo,'') as TapalNo  from TSPL_VENDOR_INVOICE_HEAD where Document_No = '" + clsCommon.myCstr(dr("Document_No")) + "'"))
                    gvDetails.CurrentRow.Cells(colDocDate).Value = clsCommon.myCstr(dr("DocDate"))
                    gvDetails.CurrentRow.Cells(colDocumentDate).Value = clsCommon.myCstr(dr("DocumentDate"))
                    gvDetails.CurrentRow.Cells(colVendorInvNo).Value = clsCommon.myCstr(dr("VendorInvoiceNo"))
                    gvDetails.CurrentRow.Cells(colOriginalAmt).Value = clsCommon.myCdbl(dr("OriginalAmt"))
                    gvDetails.CurrentRow.Cells(colTDSAmt).Value = clsCommon.myCstr(dr("TDSAmt"))
                    gvDetails.CurrentRow.Cells(colNetAmt).Value = clsCommon.myCdbl(dr("NetAmount"))
                    gvDetails.CurrentRow.Cells(colPendingAmt).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colTemp).Value = clsCommon.myCdbl(dr("PendingAmt"))
                    gvDetails.CurrentRow.Cells(colAppliedAmt).ReadOnly = True
                    '' Anubhooti 14-Nov-2014 BM00000004636
                    gvDetails.CurrentRow.Cells(colAdjustedAmt).Value = clsCommon.myCdbl(dr("NetAmount")) - clsCommon.myCdbl(dr("PendingAmt"))
                Next
                If dt.Rows.Count > 0 Then
                    If clsCommon.CompairString(ddlPaymentType.SelectedValue, "PY") = CompairStringResult.Equal Then
                        'txtVendorCode.Enabled = False
                        ddlPaymentType.Enabled = False
                    ElseIf clsCommon.CompairString(ddlPaymentType.SelectedValue, "AV") = CompairStringResult.Equal Or clsCommon.CompairString(ddlPaymentType.SelectedValue, "OA") = CompairStringResult.Equal Then
                        'txtBankCode.Enabled = False
                        ddlPaymentType.Enabled = False
                    End If
                Else
                    txtVendorCode.Enabled = True
                    txtBankCode.Enabled = True
                    ddlPaymentType.Enabled = True
                End If
                IsInsideLoadData = False
            End If
        Catch ex As Exception
            clsCommon.MyMessageBoxShow(ex.Message)
        End Try
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        Try
            If clsCommon.myLen(txtPaymentNo.Value) <= 0 Then
                clsCommon.MyMessageBoxShow("Select Document No")
                Exit Sub
            End If
            clsERPFuncationalityOLD.ShowTransHistoryData(txtPaymentNo.Value, "Payment_No", "TSPL_PAYMENT_HEADER", "TSPL_PAYMENT_DETAIL")
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
End Class
