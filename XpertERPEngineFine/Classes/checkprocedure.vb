Imports System.Data.SqlClient

Public Class checkprocedure
    Public Shared Sub adjustment()
        Dim dic As New Hashtable()
        dic.Add("", "")

    End Sub
End Class

Public Class ProgramCodeNew
    Public Shared Sub SetProgramCode()

        InsertDefaultValue(clsUserMgtCode.ExpertERP, "PERT ERP", "0", Nothing, 32)
        '--------------Favourite----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleFavourite, "Favourite", "0.1", clsUserMgtCode.ExpertERP, "M", 31)
#Region "System Administrator"
        '--------------System Administrator----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleSystemAdmin, "System Administrator", "1.00", clsUserMgtCode.ExpertERP, "M", 23)
        InsertDefaultValue(clsUserMgtCode.SubModuleSystemAdminSetup, "Setup", "1.00.01", clsUserMgtCode.ModuleSystemAdmin, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.EmployeeMaster, "Employee Master", "1.00.01.01", clsUserMgtCode.SubModuleSystemAdminSetup, 28)
        InsertDefaultValue(clsUserMgtCode.DesignationMaster, "Designation Master", "1.00.01.02", clsUserMgtCode.SubModuleSystemAdminSetup, 28)
        InsertDefaultValue(clsUserMgtCode.userMaster, "User Master", "1.00.01.03", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.userGroupMaster, "User Group Master", "1.00.01.04", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.userGroupMapping, "User Group Mapping", "1.00.01.05", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.groupProgramMapping, "Group Program Mapping", "1.00.01.06", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmScheduling, "Screen Notification Scheduling", "1.00.01.07", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmApprovalLevelScreen, "Approval Level Screen", "1.00.01.08", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmLocationSetting, "Location Setting", "1.00.01.09", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        'InsertDefaultValue(clsUserMgtCode.frmSynchronization, "Sync Master Data", "1.00.01.10", clsUserMgtCode.SubModuleSystemAdminSetup, "", 29, 0, "", "", "", True, True, False, Nothing)
        'InsertDefaultValue(clsUserMgtCode.frmTranReverse, "Transaction Reverse", "1.00.01.11", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.FisaclYearEndProcess, "Financial Year End", "1.00.01.12", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmAppIntegrator, "Application Integrator", "1.00.01.13", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmApprovalAlertSumm, "Approval Alert", "1.00.01.14", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmNotification, "Notification", "1.00.01.15", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.FrmOfficeOrderTemplate, "Office Order Template", "1.00.01.16", clsUserMgtCode.SubModuleSystemAdminSetup, 29)

        'InsertDefaultValue(clsUserMgtCode.TimeTable, "Time Table", "1.00.01.15", clsUserMgtCode.SubModuleSystemAdminSetup, 30)
        'InsertDefaultValue(clsUserMgtCode.frmConfigureSynchronization, "Configure Synchronization", "1.00.01.16", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        'InsertDefaultValue(clsUserMgtCode.FrmSendSMSMultipleUser, "Send SMS", "1.00.01.17", clsUserMgtCode.SubModuleSystemAdminSetup, "", 29, 0, "", "", "", True, True, False, Nothing)
        'InsertDefaultValue(clsUserMgtCode.UtilityImportExport, "Utility Import/Export", "1.00.01.18", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        InsertDefaultValue(clsUserMgtCode.frmChangePasswordScreen, "Change Password", "1.00.01.20", clsUserMgtCode.SubModuleSystemAdminSetup, 29)
        ' Transation

        InsertDefaultValue(clsUserMgtCode.SubModuleSystemAdminTransation, "Transaction", "1.00.02", clsUserMgtCode.ModuleSystemAdmin, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmSendSMSEmailSetting, "SMS/Email Setting", "1.00.02.01", clsUserMgtCode.SubModuleSystemAdminTransation, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmRCDFRateControl, "RCDF Rate Control", "1.00.02.02", clsUserMgtCode.SubModuleSystemAdminTransation, 8)
        InsertDefaultValue(clsUserMgtCode.frmOfficeOrder, "Office Order", "1.00.02.03", clsUserMgtCode.SubModuleSystemAdminTransation, 8)


        '=======Preeti Gupta==========
        ' Reports 
        InsertDefaultValue(clsUserMgtCode.SubModuleSystemAdminSetupReport, "Report", "1.00.03", clsUserMgtCode.ModuleSystemAdmin, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.Security_Matr, "Security Matrix Report", "1.00.03.01", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDocumentVersionReport, "Document Version Report", "1.00.03.02", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPendingDocumentList, "Pending Document List", "1.00.03.03", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptApprovalReport, "Document Approval Status", "1.00.03.04", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptUserScreenRightsReport, "User Screen Rights Report", "1.00.03.05", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSMSDetailsReport, "SMS Details Report", "1.00.03.06", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptScreenSettingReport, "Screen Setting Report", "1.00.03.07", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptActiveUsers, "Active Users", "1.00.03.08", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailSystemAdmin, "System Administrator Audit Trail Report", "1.00.03.09", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDeleteHistoryReport, "Delete History Report", "1.00.03.10", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.RCDFDashboard, "Dashboard", "1.00.03.11", clsUserMgtCode.SubModuleSystemAdminSetupReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.RCDFUnionDashboard, "Union Dashboard", "1.00.03.12", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.Complainfeedback, "Complain feedback", "1.00.03.13", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.DashboardMilkUnion, "Union Dashboard", "1.00.03.14", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.DairySaleDashboard, "Dairy Dashboard", "1.00.03.15", clsUserMgtCode.SubModuleSystemAdminSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.SalesMarketingDashboard, "Sales And Marketing Dashboard", "1.00.03.16", clsUserMgtCode.SubModuleSystemAdminSetupReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDBTDashboard, "DBT Dashboard", "1.00.03.17", clsUserMgtCode.SubModuleSystemAdminSetupReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.MostUserScreen, "Frequently Used Screens", "1.00.03.18", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSystemAdministratorTool, "System Admin Tool", "1.00.03.19", clsUserMgtCode.SubModuleSystemAdminSetupReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)


        '-------------------------------------------------------------------End SYSTEM ADMINISTRATOR---------------------------------------

#End Region

#Region "Common Services"
        '--------------Common Services----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleCommonServices, "Common Services", "1.01", clsUserMgtCode.ExpertERP, "M", 23)

        '----------------------------------Common Services------------------------------------
        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleCommonServicesSetup, "Setup", "1.01.01", clsUserMgtCode.ModuleCommonServices, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.taxAuthority, "Tax Authority", "1.01.01.01", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.taxRate, "Tax Rate", "1.01.01.02", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.taxGroup, "Tax Group", "1.01.01.03", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.paymentCodes, "Payment Code", "1.01.01.04", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.cityMaster, "City Master", "1.01.01.05", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmAbateMentMaster, "Abatment Master", "1.01.01.06", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmCompanyMaster, "Company Master", "1.01.01.07", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmAdditionalCharges, "Additional Charges", "1.01.01.08", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.PrefixGeneration, "Prefix Generation", "1.01.01.09", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FormMaster, "Form Master", "1.01.01.10", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.ChangePwd, "Change Password", "1.01.01.11", clsUserMgtCode.SubModuleCommonServicesSetup, "", 27, False, "", "", "", True, False, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmBulkPostingNew, "Bulk Posting(All Modules)", "1.01.01.12", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnPendingApproval1, "Bulk Posting", "1.01.01.13", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.BankGroupMaster, "Bank Group Master", "1.01.01.14", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.bankMaster, "Bank Master", "1.01.01.15", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmCurrencyMaster, "Currency Master", "1.01.01.16", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmCurrencyConversion, "Currency Conversion Rate", "1.01.01.17", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.CustomFieldMaster, "Custom Field Master", "1.01.01.18", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.CustomFieldMapping, "Custom Field Mapping", "1.01.01.19", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmModuleCurrencyMapping, "Module Currency Mapping", "1.01.01.20", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.lockTransaction, "Lock Transaction", "1.01.01.21", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnTakeBackup, "Backup", "1.01.01.22", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CommonServicesSetting, "Common Service Setting", "1.01.01.23", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRegionMaster, "Region Master", "1.01.01.24", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCountryMaster1, "Country Master", "1.01.01.25", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmStateMaster1, "State Master", "1.01.01.26", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DistrictMaster, "District Master", "1.01.01.27", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.bankBranchMaster, "Bank  Branch Master", "1.01.01.28", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmCalScreen, "CalCulator Screen", "1.01.01.29", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmNotificationScreen, "Notification Setting Screen", "1.01.01.30", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmFormSerialNoMaster, "Form Nos. Entry", "1.01.01.31", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmBranchAccountMapping, "Branch Account Mapping", "1.01.01.32", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmTransactionApproval, "Transaction Approval", "1.01.01.33", clsUserMgtCode.SubModuleCommonServicesSetup, "", 27, 0, "", "", "", True, True, Nothing)
        'InsertDefaultValue(clsUserMgtCode.FrmTransactionApproval, "Transaction Approval", "1.01.01.32", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMCCDiscountMaster, "Discount Master", "1.01.01.34", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.AreaMaster, "Area Master", "1.01.01.35", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPromptMsgRelatedtopending, "Prompt Message", "1.01.01.36", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSAC, "SAC Master", "1.01.01.37", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSACWiseTax, "SAC Wise Tax Master", "1.01.01.38", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmItemTaxRate, "Item Tax Rate", "1.01.01.37", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmItemWiseTax, "Item Wise Tax Master", "1.01.01.39", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmRackBinMaster, "Rack/Bin Master", "1.01.01.40", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmDeptHeadCustomerMapping, "Dept Head Customer Mapping", "1.01.01.40", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.paymentTerms, "Payment Terms", "1.01.01.41", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmLocationDistanceMapping, "EWay Bill Distance Master", "1.01.01.42", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDisplaySquenece, "Display Squenece", "1.01.01.43", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBlockMaster, "Block Master", "1.01.01.44", clsUserMgtCode.SubModuleCommonServicesSetup, 27)

        InsertDefaultValue(clsUserMgtCode.frmRequestMaster, "Request Master", "1.01.01.49", clsUserMgtCode.SubModuleCommonServicesSetup, 27)
        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleCommonServicesTransaction, "Transaction", "1.01.02", clsUserMgtCode.ModuleCommonServices, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.reverseTransaction, "Bank Reverse Entry", "1.01.02.01", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.bankTransfer, "Bank Transfer", "1.01.02.02", clsUserMgtCode.SubModuleCommonServicesTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.FrmBankReco, "Bank Reconciliation", "1.01.02.03", clsUserMgtCode.SubModuleCommonServicesTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.FrmCFormEntry, "CForm Entry", "1.01.02.04", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmBankGuaranteeMaster, "Bank Guarantee Master", "1.01.02.05", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.BankOpeningReco, "Bank Opening Reco", "1.01.02.06", clsUserMgtCode.SubModuleCommonServicesTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.RevaluationEntry, "Revaluation Entry", "1.01.02.07", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.ItemRackBinMapping, "Item Rack Bin Mapping", "1.01.02.08", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.LoanEntry, "Loan Entry", "1.01.02.09", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.LoanInstallment, "Loan Installment Entry", "1.01.02.10", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmRequestAproval, "Request Approval", "1.01.02.11", clsUserMgtCode.SubModuleCommonServicesTransaction, 8)
        ' Reports  
        InsertDefaultValue(clsUserMgtCode.SubModuleCommonServicesReport, "Report", "1.01.03", clsUserMgtCode.ModuleCommonServices, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.FrmBankReverse, "Bank Reverse", "1.01.03.01", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankBook, "Bank Cash Book", "1.01.03.02", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankBookDayWise, "Bank/Cash Book DayWise", "1.01.03.03", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCashVoucher, "Contra Voucher", "1.01.03.04", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.fundprovision, "Fund Provision", "1.01.03.05", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.DVAT30, "DVAT-30", "1.01.03.06", clsUserMgtCode.SubModuleCommonServicesReport, 37) ''expert
        InsertDefaultValue(clsUserMgtCode.DVAT31, "DVAT-31", "1.01.03.07", clsUserMgtCode.SubModuleCommonServicesReport, 37) ''expert
        InsertDefaultValue(clsUserMgtCode.TaxTracking, "Tax Tracking", "1.01.03.08", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        'InsertDefaultValue(clsUserMgtCode.ExciseSummary1, "Excise Summary", "1.01.03.08", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmExciseChapterWise, " ER1 Report", "1.01.03.09", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.CrptRG1Detail1, "RG-1", "1.01.03.10", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmOpeningBalance, "Opening Balance", "1.01.03.10", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCFormReport, "C-Form Report", "1.01.03.11", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptServiceTaxDetail, "Service Tax Report", "1.01.03.13", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBankWiseChequeIssue, "Bank Wise Cheque Issue Report", "1.01.03.12", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankBookLocationDetail, "Bank Cash Book Location Detail", "1.01.03.13", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPromptMsgRelatedtopending, "Promtmsg", "1.01.03.14", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBranchAccountMapping, "Branch Account Mapping Report", "1.01.03.15", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmLockTransactionReport, "Lock Transaction Report", "1.01.03.16", clsUserMgtCode.SubModuleCommonServicesReport, 37)

        InsertDefaultValue(clsUserMgtCode.RptInventoryMovement, "Inventory Movement Report", "1.01.03.17", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptItemWiseTaxMasterReport, "Item Wise Tax Master Report", "1.01.03.18", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmReconcilationRpt, "Reconcilation Report", "1.01.03.24", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmBankBookRecoReport, "Bank Book Reco Report", "1.01.03.25", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailCommonServices, "Common Services Audit Trail Report", "1.01.03.19", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDasboard, "Dash Board Report", "1.01.03.20", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDasboardCombine, "Combine Dash Board", "1.01.03.21", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptDataEntryTracingReport, "Data Entry Tracing Report", "1.01.03.27", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDocumentCancelledReport, "Document Cancelled Report", "1.01.03.22", clsUserMgtCode.SubModuleCommonServicesReport, 37)
        InsertDefaultValue(clsUserMgtCode.DashboardMilkProcurement, "Dashboard Milk Procurement", "1.01.03.23", clsUserMgtCode.SubModuleCommonServicesReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.JanaadharStatusReport, "Janaadhar Status Report", "1.01.03.24", clsUserMgtCode.SubModuleCommonServicesReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmGSTDocumentList, "GST Document List", "1.01.03.25", clsUserMgtCode.SubModuleCommonServicesReport, 37)

#End Region

#Region "Milk Procurement and DBT"
        '----------------------------------Milk Procurement---------------------------------------'
        InsertDefaultValue(clsUserMgtCode.ModuleMCCMilkProcurement, "Milk Procurement and DBT", "1.03", clsUserMgtCode.ExpertERP, "M", 23)

        '======================MCC Pro Setup
        InsertDefaultValue(clsUserMgtCode.SubModuleMCCMilkProcurementSetup, "Setup", "1.03.01", clsUserMgtCode.ModuleMCCMilkProcurement, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmMCCMaster, "MCC/BMC Master", "1.03.01.01", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVSP_VLCMaster, "DCS Master", "1.03.01.02", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMPMaster, "Farmer Master", "1.03.01.03", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmHeadLoadMaster, "Head Load Master", "1.03.01.04", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartMaster, "Milk Price Chart Master", "1.03.01.05", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.MilkPricePlanning, "Milk Price Planning", "1.03.01.06", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmPriceChartUploader, "Price Chart Uploader", "1.03.01.07", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBulkRoutMaster, "BMC Route Master", "1.03.01.08", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmZoneMaster, "Zone Master", "1.03.01.09", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmTankerTransporterMaster, "Transporter Master", "1.03.01.10", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmTankerMaster, "Tanker Master", "1.03.01.11", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VerifyMPIFSC, "Farmer IFSC Verification", "1.03.01.12", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDeductionMaster, "Deduction Master", "1.03.01.13", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DCSAdditionDeduction, "Auto Addition/Deduction", "1.03.01.14", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.OwnBMCGainLossRate, "Own BMC Gain Loss Master", "1.03.01.15", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmOwnBmcExpanse, "Own Bmc Expanse", "1.03.01.16", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRevenueVillageMaster, "Revenue Village Master", "1.03.01.17", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmGrampanchayatMaster, "Grampanchayat Master", "1.03.01.18", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPanchayatSamitiMaster, "Panchayat Samiti Master", "1.03.01.19", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVidhanSabhaMaster, "Vidhan Sabha Master", "1.03.01.20", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVillageMaster, "Village Master", "1.03.01.21", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVSPMaster, "VSP Master", "1.03.01.22", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVLCMaster, "VLC Master", "1.03.01.23", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.MilkRejectType, "Milk Reject Type", "1.03.01.24", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPaymentCycleMaster, "Payment Cycle Master", "1.03.01.25", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.GenratePaymentCycle, "Generate Payment Cycle", "1.03.01.26", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.GazeReading, "Gaize Master", "1.03.01.27", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.ChillingChargesSlab, "Chilling Charges Slab", "1.03.01.28", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CappingMaster, "DBT Capping Master", "1.03.01.29", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSalePriceChart, "BMC Item Price Chart", "1.03.01.30", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkRouteMaster, "Milk Route Master", "1.03.01.31", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVLCUploader, "VLC Uploader", "1.03.01.32", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPrimaryTransporterMaster, "Primary Transporter Master", "1.03.01.33", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPrimaryTransporterVehicalMaster, "Primary Transporter Vehicle Master", "1.03.01.34", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVLCRouteShiftMaster, "VLC Route Shift Master", "1.03.01.35", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartMaster_Bulk, "Bulk Milk Price Master", "1.03.01.36", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkCollectionArea, "Milk Collection Area", "1.03.01.37", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkCollectionLevels, "Milk Collection Levels", "1.03.01.38", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        ' InsertDefaultValue(clsUserMgtCode.frmMilkCollectionCenters, "Milk Collection Centers", "1.18.01.18", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkVehicleTypeMaster, "Vehicle Type Master", "1.03.01.39", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkVehicleMaster, "Vehicle Master", "1.03.01.40", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransportRateMaster, "Transport Rate Master", "1.03.01.41", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkComponentMaster, "Milk Component Master", "1.03.01.42", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkComponentRateList, "Milk Component Rate List", "1.03.01.43", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkAdvanceMaster, "Advance Master", "1.03.01.44", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkRateTypeMaster, "Milk Rate Type Master", "1.03.01.45", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftMaster, "Shift Master", "1.03.01.46", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSeasonMaster, "Season Master", "1.03.01.47", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmUOMMaster, "UOM Master", "1.03.01.48", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmdeductionGroup, "Deduction Group master", "1.03.01.49", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmGroupOfDeduction, " Group Of Deduction", "1.03.01.50", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DeductionMapping, "Deduction Mapping", "1.03.01.51", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CanMaster, "Can Master", "1.03.01.52", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DockMaster, "Dock Master", "1.03.01.53", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmTankerDispatchPriceMaster, "Tanker Dispatch Price Master", "1.03.01.54", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDeductionTypeMaster, "Deduction Type Master", "1.03.01.55", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkReasonMaster, "Mcc Reason Master", "1.03.01.56", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmIncentiveMaster, "Incentive Master", "1.03.01.57", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.LocationDistanceMapping, "Location Distance Mapping", "1.03.01.58", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMCCSMSSettiing, " MCC Mail/SMS Setting", "1.03.01.59", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShiftManual, "MCC Shift Setting", "1.03.01.60", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmPortSettings, "Port Settings", "1.03.01.61", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.MCCSetting, "Milk Procurement Setting", "1.03.01.62", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmVSPIncentiveTagging, "VSP Incentive Tagging", "1.03.01.63", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMCCTransactionApproval, "MCC Transaction Approval", "1.03.01.64", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVLCMasterTarget, "VLC/MP Target Master ", "1.03.01.65", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmItemChargeCategoryMaster, "Charge Category Master", "1.03.01.66", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMCCFarmerMapping, "MCC Farmer Mapping", "1.03.01.67", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VLCMappingForMPAmount, "VLC Mapping For MP Amount", "1.03.01.68", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FreightChargesMaster, "Freight Charges Master", "1.03.01.69", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VLCMappingForMP_PP, "VLC Mapping For MP Payment Process", "1.03.01.70", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VSP_Commsission, "VSP Commission", "1.03.01.71", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VSP_Deduction, "VSP Deduction", "1.03.01.72", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VSP_DayWiseIncetive, "VSP Day Wise Incetive", "1.03.01.73", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VSP_Mapping, "VSP Mapping", "1.03.01.74", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FarmerProMaster, "Farmer Pro Master", "1.03.01.75", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.MPIncetiveSlab, "MP Incentive Slab", "1.03.01.76", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DCSFinancialHead, "DCS Financial Head", "1.03.01.77", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DCSSupervisorTagging, "DCS Supervisor Tagging", "1.03.01.78", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.MPMilkPrice, "MP Milk Price", "1.03.01.79", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmUnpostBmcDcs, "Unpost BMC DCS", "1.03.01.80", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDCSTransportationCharges, "DCS Transportation Charges", "1.03.01.81", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmManualIoTFarmerCollection, "Manual/IoT Farmer Collection", "1.03.01.82", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBMCCategory, "BMC Category", "1.03.01.83", clsUserMgtCode.SubModuleMCCMilkProcurementSetup, 27)

        '=============================Mcc Transaction===========================================================
        InsertDefaultValue(clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "MCC Transaction", "1.03.02", clsUserMgtCode.ModuleMCCMilkProcurement, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCGateEntry, "BMC Milk Gate Entry", "1.03.02.01", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, False, False, Nothing)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCMultipleDays, "BMC Milk Collection Multiple Days", "1.03.02.02", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCC, "BMC Milk Collection", "1.03.02.03", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionMCCSample, "BMC Milk QC", "1.03.02.04", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, False, False, "", Nothing, "", "", "", "Exception", "", False, "", 0, False, False)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCSMultipleDays, "DCS Milk Collection Multiple Days", "1.03.02.05", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCS, "DCS Truck sheet", "1.03.02.06", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkShiftUploader, "Dock Entry", "1.03.02.07", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkProcurementUploader, "Milk Uploader", "1.03.02.08", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkProcurementCorrection, "Milk Correction", "1.03.02.09", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.MilkRetesting, "Milk Retesting", "1.03.02.10", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMultipleProcDeduction, "Manual Deduction entry", "1.03.02.11", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkVSPPayment, "DCS Bill Processing", "1.03.02.12", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmPaymentProcess, "DCS Payment Process", "1.03.02.13", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmBankAdvise, "DCS Bank Advice", "1.03.02.14", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSendBillToDCS, "Send Milk bill to Mobile", "1.03.02.15", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.DCSMPIncentiveReco, "DCS Farmer Reconciliation", "1.03.02.16", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.DBTMonthlyFarmerMilk, "DBT Monthly Farmer Milk", "1.03.02.17", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntry, "Farmer DBT Entry", "1.03.02.18", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.DBTCappingCheck, "DBT Capping Check", "1.03.02.19", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.DBTCappingIncrease, "DBT Capping Increase", "1.03.02.20", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTUploader, "DBT NEFT Uploader", "1.03.02.20", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.DBTNEFTReject, "DBT NEFT Reject", "1.03.02.22", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDCSSaleEntry, "DCS Sale Entry", "1.03.02.23", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterial, "DCS Sale", "1.03.02.24", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturn, "DCS SaleReturn", "1.03.02.25", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSendSMSToDCS, "Send SMS To Mobile", "1.03.02.26", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "Exception", "Tanker P/L SMS", "Tanker QC SMS", "", "", False, "")
        InsertDefaultValue(clsUserMgtCode.frmVlcdataUploadar, "DCS Data Uploader", "1.03.02.27", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmReviseMilkBill, "Revise Milk Bill", "1.03.02.28", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "Exception", "Tanker P/L SMS", "Tanker QC SMS", "", "", False, "")
        InsertDefaultValue(clsUserMgtCode.BMC_Transporter_Bill, "BMC Transporter Bill", "1.03.02.29", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMobileDCSDemand, "Mobile DCS Demand", "1.03.02.30", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmTransferToSaving, "Transfer To Saving", "1.03.02.31", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMccMatSaleUploader, "DCS Sale Uploader", "1.03.02.32", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkSRN, "Milk SRN", "1.03.02.33", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkPurchaseInvoice, "Milk Purchase Invoice", "1.03.02.34", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmOpenMCCShift, "Open MCC Shift", "1.03.02.35", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryIn, "Milk Gate Entry In", "1.03.02.36", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryWeightment, "Milk Gate Entry Weighment", "1.03.02.37", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkReceipt, "Milk Receipt", "1.03.02.38", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkReject, "Milk Reject", "1.03.02.39", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkSample, "Milk Sample", "1.03.02.40", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkGateEntryOut, "Milk Gate Entry Out", "1.03.02.41", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkShiftEndMCC, "Milk Shift End", "1.03.02.42", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmVLCDataUploaderManual, "Village Data", "1.03.02.43", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCDispatch, "Tanker Dispatch", "1.03.02.44", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmVSPAssetIssue, "VSP Asset Issue", "1.03.02.45", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkVSPIssuePayment, "VSP Issue Payment Processing", "1.03.02.46", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, "Provision Entry", "1.03.02.47", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.IncentiveEntry, "Incentive Entry", "1.03.02.48", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMccScrapGatePass, "Mcc/Scrap GatePass", "1.03.02.49", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmProcurementDeduction, "Procurement Deduction", "1.03.02.50", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.PrimaryTransportProvisionCorrection, "Primary Transport Provision Correction", "1.03.02.51", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmOutputEntry, "Output Entry", "1.03.02.52", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAcknowledgementEntry, "Acknowledgement Entry", "1.03.02.53", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCreateBMCDCSbyMobile, "BMC DCS Mobile Correction", "1.03.02.54", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.DBTPayment, "DBT Payment", "1.03.02.55", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.DCSFinancialEntry, "DCS Financial Entry", "1.03.02.56", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMakeSavingPayment, "Make Saving Payment", "1.03.02.57", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.SendDBTToJanaadhar, "Send DBT To Janaadhar", "1.03.02.58", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmCattleFeedSalePurchaseUploader, "Cattle Feed Sale/Purchase Uploader", "1.03.02.59", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.AMCUMaster, "AMCU Master", "1.03.02.60", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCancelDCSSale, "Cancel DCS Sale", "1.03.02.61", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.WEIGHINGMaster, "WEIGHING Master", "1.03.02.62", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmTransporterPaymentProcess, "Transporter Payment Process", "1.03.02.64", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.MilkProcurementFarmerCorrection, "Farmer Milk Correction", "1.03.02.65", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 8)


        '=========================================================================================
        '--Reportsmdi
        InsertDefaultValue(clsUserMgtCode.SubModuleMCCMilkProcurementReport, "Report", "1.03.03", clsUserMgtCode.ModuleMCCMilkProcurement, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.MCCMilkRegister, "Milk Purchase Register", "1.03.03.001", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDailyQtyReport, "Daily Qty Report", "1.03.03.002", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptBMCDCSAbsent, "Absent BMC/DCS Report", "1.03.03.003", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTempTruckSheetCollectionReport, "BMC Truck sheet report", "1.03.03.004", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerUtilization, "Tanker Approval Utilization Report", "1.03.03.005", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTemporaryPaymentDeductionSummary, "DCS Payment Deduction Summary", "1.03.03.006", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptMarginAmt, "DCS Margin Report", "1.03.03.007", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTruckSheetDailySummaryReport, "BMC Daily Summary Report", "1.03.03.008", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmVendorBankAdvice, "DCS Bank Advice", "1.03.03.009", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessRouteReport, "DCS Wise Summary Reports", "1.03.03.010", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptMultipleDeductionReport, "Manual Addition Deduction Report", "1.03.03.011", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptMCCShiftReportRouteWise, "MCC Shift Report (Route Wise)", "1.03.03.012", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptvillageslip, "Village Slip Report", "1.03.03.013", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkBillMCC, "MCC Milk Bill Details (MCC Wise) Report", "1.03.03.014", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkBillRouteWise, "MCC Milk Bill Details (Route Wise) Report", "1.03.03.015", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCMilkBillSummary, "MCC Milk Bill Summary", "1.03.03.016", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptLowProcurement, "Low Procurement", "1.03.03.017", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMonthlyVLCProcurement, "Monthly VLC Procurement", "1.03.03.018", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDailyDifferentReport, "Daily Different Report", "1.03.03.019", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCDA, "CDA Report", "1.03.03.020", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMillPurchaseBill, "Milk Puchase Bill", "1.03.03.021", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptDailyGainDay, "Milk Gain-Loss Report", "1.03.03.022", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPendingMilkSRN, "Pending Milk SRN Report", "1.03.03.023", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptShiftCodeWise, "Shift Report Code Wise", "1.03.03.024", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptShifReportZeroAmtSample, "Shift Report Zero Amount Sample", "1.03.03.025", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPendingProvisionReport, "Pending Provision Report", "1.03.03.026", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkStockLedgerSummary, "Milk Stock Ledger (Summary) Report", "1.03.03.027", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMilkRouteVehicleReport, "Milk Route Vehicle Report", "1.03.03.028", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMccSaleRegister, "MCC Sale Register Report", "1.03.03.029", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptDispatchOfMilkTransfer, "Dispatch Of Milk Transfer report", "1.03.03.030", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.MccSummaryReport, "Mcc Summary Report", "1.03.03.031", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCollectionLevelChart, "Collection Level Chart", "1.03.03.032", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCollectionCenterChart, "Collection Center Chart", "1.03.03.033", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMemberPaymentSlip, "Member Payment Slip", "1.03.03.034", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMPWiseMilkCollectionATPoolingPoint, "Farmer Milk Collection", "1.03.03.035", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMPWiseMilkCollection, "Milk Collection As Per Data Entry", "1.03.03.036", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVillageDiffReport, "Village Difference Report", "1.03.03.037", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPrimaryTransporter, "Primary Transporter Report", "1.03.03.038", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVSPOrVLCVarationRpt, "VSP Or VLC Varation Report", "1.03.03.039", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMPIDReport, "Farmer ID Report ", "1.03.03.040", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCollectionAnalysis, "Collection Analysis", "1.03.03.041", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmTDSReport, "TDS Report", "1.03.03.042", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCVLCVarationReport, "MCC Or VLC Varation Report ", "1.03.03.043", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVSPAssetIssue1, "VSP Asset Issue Report", "1.03.03.044", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMCCMilkStatus, "MCC Milk Status Report", "1.03.03.045", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.MCCProvisonReport, "MCC Provision Report", "1.03.03.046", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVillageDiffReportParas, "Village Difference Report(DEMO)", "1.03.03.047", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVSPIncentiveRegister, "VSP Incentive Register", "1.03.03.048", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.MonthlyProgressReport, "Monthly Progress Report", "1.03.03.049", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPDAccountSummariesReport, "PD Account Summaries Report", "1.03.03.050", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDCSDeductionReport, "DCS Deduction Report", "1.03.03.051", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerMilkQcVsERPEntry, "Tanker Milk Qc Vs ERP Entry Report", "1.03.03.052", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCVLCVariationReportPourersNo, "MCC/VLC Variation Report & Pourers No", "1.03.03.053", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMccSaleAdjustment, "MCC Sale Adjustment", "1.03.03.054", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMccBulkmilkRegister, "Day wise Milk Collection", "1.03.03.055", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDailyWiseMilkCost, "Daily wise Milk Cost", "1.03.03.056", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptDailyStanderdMilkQtyMCCWise, "Daily Standerd Milk Qty MCC Wise", "1.03.03.057", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptWeighmentRegister, "Weighment Register Report", "1.03.03.058", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptSecondaryTransporterReport, "Secondary Transport Format For ERP", "1.03.03.059", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVLCVehicleWeigmentRegister, "VLC Vehicle Weigment Register Report", "1.03.03.060", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMilkReceiptImproperWeight, "Milk Receipt Improper Weight Report", "1.03.03.061", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptRoutewiseTPTimeTable, "Routewise TP Table Report", "1.03.03.062", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCRouteTimeTable, "Milk Route Time Report", "1.03.03.063", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RpttankerReport, "Tanker Report", "1.03.03.064", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.MPIncentiveEntryReport, "MP Incentive Entry Register", "1.03.03.065", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.VLCProgressReport, "VLC Progress Report", "1.03.03.066", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptTankerInTransit, "Transit In Tanker", "1.03.03.067", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMccCollectionDetails, "MCC Collection Details", "1.03.03.068", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailMilkProcurementMCC, "Milk Procurement MCC Audit Trail Report", "1.03.03.069", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMccMasterDetail, "MCC Master Detail", "1.03.03.070", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptListofCowDCS, "List Of Cow DCS ", "1.03.03.071", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMccProcurementUploader, "MCC Procurement Uploader Report", "1.03.03.072", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCWiseAbstractReport, "MCC Wise Abstract Report", "1.03.03.073", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCollectionDataChangeReport, "Collection Data Change Report", "1.03.03.074", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCPaymentSummary, "MCC Payment Summary", "1.03.03.075", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCBillGenrationStatus, "MCC Bill Genration Status", "1.03.03.076", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMCCDataEntrySummaryReport, "MCC Data Entry Summary Report", "1.03.03.077", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerAllocationReport, "Tanker Allocation Report", "1.03.03.078", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkCollectionShiftWiseReport, "Milk Collection Shift Wise Report", "1.03.03.079", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMccDeductionReport, "Mcc Deduction Report", "1.03.03.080", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptTankerDispatchvsAckn, "Tanker Dispatch vs Acknowledgement", "1.03.03.081", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.MilkProcurementVisualReport, "Visual Report Procurement", "1.03.03.082", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessReport, "Payment Process Report", "1.03.03.083", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptLiabilityReport, "Liability Report", "1.03.03.084", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVSPMilkNotsold, "Closed DCS Outstanding Report", "1.03.03.085", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkBillProcurementSummary, "Milk Bill Procurement Summary", "1.03.03.086", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkAnalysis, "Milk Analysis", "1.03.03.087", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSocietyLedgerReport, "Society Ledger", "1.03.03.088", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkCostReport, "Milk Cost Report", "1.03.03.089", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPaymentProcessReportBMCSocietyWise, "Cattle Feed Adjustment Report", "1.03.03.090", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDCSFinancial, "DCS Financial Report", "1.03.03.091", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTruckSheetReport, "BMC Truck Sheet Report", "1.03.03.092", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDcsSecurity, "DCS Security Report", "1.03.03.093", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDayWiseCol, "Day Wise Collection Report", "1.03.03.094", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPaymentCycleWiseReport, "Payment Cycle Wise Report", "1.03.03.095", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptMobileAppMilkCollection, "Mobile App Milk Collection Report", "1.03.03.096", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDBTRecoVsIncentiveReport, "DBT Reco Vs Incentive Repor", "1.03.03.097", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmAutoAdditionDeductionReport, "Auto Addition/Deduction Report", "1.03.03.098", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptAutoMultipleAdditionDeduction, "Auto Multiple Addition Deduction", "1.03.03.099", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBmcCollection, "Bmc Collection Report", "1.03.03.100", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDBTSummaryMonthlyWise, "DBT Year Wise Summary", "1.03.03.101", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDBTMilkPayment, "DBT Milk Payment", "1.03.03.102", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBMCTankerTestingReport, "BMC Tanker Testing Report", "1.03.03.103", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkPaymentSummary, "Milk Payment Summary Report", "1.03.03.104", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptCattleFeedSaleReport, "Cattle Feed Sale Report", "1.03.03.105", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptTransferToSavingReport, "Transfer To Saving Report", "1.03.03.106", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDcsPaymentReport, "Dcs Payment Report", "1.03.03.107", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptYearlyMonthlyDcsCollectionReport, "Yearly/Monthly Dcs Collection", "1.03.03.108", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDCSSavingLedger, "DCS Saving Ledger", "1.03.03.109", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, 37)
        InsertDefaultValue(clsUserMgtCode.rptBMCCapacityUtilizationReport, "BMC Capacity Utilization Report", "1.03.03.110", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBMCPerodicalReport, "BMC Periodical Report", "1.03.03.111", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPerodicalDispatchReport, "Periodical Dispatch Report", "1.03.03.112", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDBTCappingReport, "DBT Capping Report", "1.03.03.112", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptHeadLoadReport, "Head Load  Report", "1.03.03.113", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDCSCredit, "DCS Credit Report", "1.03.03.114", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSuspanceDCSReport, "Suspense Milk Report", "1.03.03.115", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerGainLossReport, "Tanker Gain Loss Report", "1.03.03.116", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDcscollectionHistory, "Dcs collection History", "1.03.03.117", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmHeadLoadHistory, "Head Load History", "1.03.03.118", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmOneDayStop, "One Day Stop", "1.03.03.119", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmLocalMilkSaleRegister, "Local Milk Sale Register", "1.03.03.120", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptFarmerSaleRegister, "Farmer Sale Register", "1.03.03.121", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmMonthWiseBeneficieriesReport, "Month Wise Beneficieris Report", "1.03.03.122", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptBMCGazeCapacityUtilizationReport, "BMC Gaze Capacity Utilization Report", "1.03.03.123", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmTTSavingReport, "Saving Balance Report", "1.03.03.124", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBmcTankerProfitLossReport, "Bmc Tanker Profit Loss report", "1.03.03.125", clsUserMgtCode.SubModuleMCCMilkProcurementReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.BmcLabReport, "Bmc Lab Report", "1.03.03.126", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDayWiseMilkCollectionVsDispatch, "Day Wise Milk Collection Vs Dispatch", "1.03.03.127", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.CompareVoucherReport, "Compare Voucher Report", "1.03.03.128", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBankSavingAdvice, "Bank Saving Advice", "1.03.03.129", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDCSWiseHoldPayments, "DCS Wise Hold Payment & TDS", "1.03.03.130", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBMCMobileHistory, "BMC Mobile History Report", "1.03.03.131", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.YearlyBillReport, "DCS Income", "1.03.03.132", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmLastDCSCollectionItem, "Last DCS Collection/Item", "1.03.03.133", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptGheeAndCattleFeedDeductionStatementReport, "Ghee And Cattle Feed Deduction Statement Report", "1.03.03.134", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmMilkProcurementSocietySummary, "Milk Procurement Society Summary", "1.03.03.135", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.CanSummaryReport, "Can Summary Report", "1.03.03.136", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.YearlyDBTSummaryReport, "Yearly DBT Summary Report", "1.03.03.137", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmTotalDeductionReport, "Total DCS Deduction Report", "1.03.03.138", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBMCDailyQtyReport, "BMC Daily Qty Report", "1.03.03.139", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDBTMonthWiseFarmerDetail, "DBT Month Wise Farmer Detail", "1.03.03.140", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.DcsMonthWiseMilkCollectionReport, "Dcs Month Wise Milk Collection Report", "1.03.03.141", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.ZoneWiseReport, "Zone Wise Report", "1.03.03.142", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkProcurementReport, "Milk Procurement Report", "1.03.03.143", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDCSTruckSheetRegister, "DCS Truck Sheet Register", "1.03.03.144", clsUserMgtCode.SubModuleMCCMilkProcurementReport, 37)

        InsertDefaultValue(clsUserMgtCode.SubModuleMCCMilkProcurementBMCReport, "BMC Report", "1.03.04", clsUserMgtCode.ModuleMCCMilkProcurement, "SM", 30)
#End Region

#Region "Purchase"
        '--------------Purchase----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModulePurchase, "Purchase", "1.04", clsUserMgtCode.ExpertERP, "M", 23)

        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModulePurchaseSetup, "Setup", "1.04.01", clsUserMgtCode.ModulePurchase, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmPurchaseSetting, "Purchase Setting", "1.04.01.01", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VendorItemDetails, "Vendor Item Details", "1.04.01.02", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CostCenter, "Cost Center", "1.04.01.03", clsUserMgtCode.SubModulePurchaseSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmRequisitionApproval, "Requisition Approval", "1.04.01.04", clsUserMgtCode.SubModulePurchaseSetup, 27) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.RequisitSubTypeMaster, "Requisite SubType Master", "1.04.01.05", clsUserMgtCode.SubModulePurchaseSetup, 27) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.FrmFormIssueReceiptEntry, "Form Issue/Receive List Entry", "1.04.01.06", clsUserMgtCode.SubModulePurchaseSetup, 27)
        ''richa Ticket no BM00000003507
        InsertDefaultValue(clsUserMgtCode.FrmCostCentreGroupStores, "Cost Centre Group Stores", "1.04.01.07", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryTermsMaster, "Delivery Terms Master", "1.04.01.04", clsUserMgtCode.SubModulePurchaseSetup, 27)
        'STUTI
        InsertDefaultValue(clsUserMgtCode.capexmaster, "Capex Master", "1.04.01.09", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.capexbudget, "Sub Capex Master", "1.04.01.10", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmUnitMaster, "Cost Center Unit Master", "1.04.01.11", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCostCenterTypeMaster, "Cost Center Type Master", "1.04.01.12", clsUserMgtCode.SubModulePurchaseSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMaterialSalePriceChart, "Material Sale Price Chart", "1.04.01.13", clsUserMgtCode.SubModulePurchaseSetup, 27)
        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModulePurchaseTransaction, "Transaction", "1.04.02", clsUserMgtCode.ModulePurchase, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseRequistion, "Purchase Indent", "1.04.02.01", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.mbtnPendingApprovalOfReq, "Pending Indent For Approval", "1.04.02.02", clsUserMgtCode.SubModulePurchaseTransaction, 8) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.RMDemanApproval, "RM Demand Approval For RAL", "1.04.02.03", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.RFQ, "RFQ", "1.04.02.04", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, False, True, Nothing) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.VendorQuotation, "Vendor Quotation", "1.04.02.05", clsUserMgtCode.SubModulePurchaseTransaction, 8) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.VendorComparison, "Vendor Quotation Comparison", "1.04.02.06", clsUserMgtCode.SubModulePurchaseTransaction, 8) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.VendorComparisonApproval, "Vendor Quotation Approval", "1.04.02.07", clsUserMgtCode.SubModulePurchaseTransaction, 8) ''ExpertERP
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseOrder, "Purchase Order", "1.04.02.08", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, "erp@1234", Nothing, "Advance", "Internal", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.SetPOSchedule, "Set PO Schedule", "1.04.02.09", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, False, False, False, "", Nothing, "", "", "", "", "")
        InsertDefaultValue(clsUserMgtCode.mbtnGRN, "Gate Received Note", "1.04.02.10", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, "umang@123", Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.VisualRandomQC, "Visual Random QC", "1.04.02.11", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, "umang@123", Nothing, "", "Internal", "", "", "")

        InsertDefaultValue(clsUserMgtCode.POWeighment, "PO Weighment", "1.04.02.11", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.POUnloading, "PO Unloading", "1.04.02.12", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, Nothing, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.mbtnMRN, "Material Received Note", "1.04.02.13", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, "umang@321", Nothing, "", "Internal", "", "", "", False, "", 0, True, False)

        InsertDefaultValue(clsUserMgtCode.mbtnSRN, "Store Received Note", "1.04.02.15", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.PurchaseGateOut, "Gate Out", "1.04.02.16", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.SRNReturn, "Store Received Note Return", "1.04.02.17", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.RALNOC, "RAL NOC", "1.04.02.18", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, False, False, False, Nothing, Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.TenderShortPenalty, "Tender Short Penalty", "1.04.02.19", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseInvoice, "Purchase Invoice", "1.04.02.20", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseReturn, "Purchase Return", "1.04.02.21", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.mbtnNRGP, "NRGP Request", "1.04.02.22", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.mbtnGatePass, "RGP/NRGP", "1.04.02.23", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.mbtnIssueReturn, "Issue/Return/Transfer", "1.04.02.24", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, 0, "", "", "", False, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMaterialQuotation, "Material Quotation", "1.04.02.25", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMaterialQuotationOrder, "Material Quotation Order", "1.04.02.26", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMaterialQuotationComparison, "Material Quotation Comparison", "1.04.02.27", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.ScrapSale, "Material Sales", "1.04.02.28", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "")
        'InsertDefaultValue(clsUserMgtCode.FrmItemConversion, "Item Conversion", "1.04.02.26", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmStoreRequistion, "Store Requisition", "1.04.02.29", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmPurchaseSchedule, "Purchase Schedule", "1.04.02.30", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.FrmBulkCloser, "Bulk Closer", "1.04.02.29", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmScrapSaleGateOut, "Misc. Sale Gate Out", "1.04.02.31", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.JobWorkDispatch, "Job Work Dispatch", "1.04.02.31", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.mbtnPurchaseJobWork, "Purchase JobWork", "1.04.02.32", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.SRNReturnListCancellation, "SRN Return List For Cancellation", "1.04.02.33", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.ScrapSaleRetrun, "Material Sales Return", "1.04.02.33", clsUserMgtCode.SubModulePurchaseTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryMISSAle, "Sale Return Gate Entry Misc Sale", "1.04.02.34", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmTender, "Tender", "1.04.02.35", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCorrectionforWrongEntry, "Correction For Wrong Entry", "1.04.02.36", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDeletionForEntry, "Deletion For Entry", "1.04.02.37", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmShortSupplyPenalty, "Short Supply Penalty", "1.04.02.38", clsUserMgtCode.SubModulePurchaseTransaction, 8)

        'richa
        'InsertDefaultValue(clsUserMgtCode.frmItemQuantityInformation, "Item Quantity Information ", "1.04.03.17", clsUserMgtCode.SubModulePurchaseTransaction, 8)
        ' Reports 
        InsertDefaultValue(clsUserMgtCode.SubModulePurchaseReport, "Report", "1.04.03", clsUserMgtCode.ModulePurchase, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.FrmPendingRequisitionQty, "Pending Requisition", "1.04.03.01", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPendingGrn_Qty, "Pending GRN", "1.04.03.02", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPendingMrn_Qty, "Pending MRN", "1.04.03.03", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPo_action, "Purchase Order Tracking", "1.04.03.04", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderRegister, "Purchase Order Register", "1.04.03.05", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnGRNReport, "GRN Report", "1.04.03.06", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPendingSrn_Qty, "SRN Pending For Bill", "1.04.03.07", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPurchasebookReport1, "Purchase Book", "1.04.03.08", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPurchaseRegisterReport, "Purchase Register Report", "1.04.03.08", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRGP_Register_NRGP, "RGP/NRGP Register", "1.04.03.09", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmVendorWiseReturnableGoodBalance, "Vendor Wise Returnable Goods Balance", "1.04.03.10", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmItemWiseDispatchLedger3, "Scrap Sale Register", "1.04.03.11", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseHistory, "Purchase History ", "1.04.03.12", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptQualityStatus, "Quality Status Report", "1.04.03.13", clsUserMgtCode.SubModulePurchaseReport, 37)


        'InsertDefaultValue(clsUserMgtCode.mbtnStoreLedger, "Stores Ledger", "1.04.03.14", clsUserMgtCode.SubModulePurchaseReport, 37)
        'InsertDefaultValue(clsUserMgtCode.mbtnDailyRcptNoteSummary, "Daily Receipt Note Summary", "1.04.03.15", clsUserMgtCode.SubModulePurchaseReport, 37)
        'InsertDefaultValue(clsUserMgtCode.Store_Receipt_Note, "Daily Receipt Detail Work", "1.04.03.16", clsUserMgtCode.SubModulePurchaseReport, 37)


        InsertDefaultValue(clsUserMgtCode.FrmIssueOrReturnItemWiseSummary, "Net Iss-Ret Item Wise Summary", "1.04.03.14", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.StockStatement, "Stock Statement Report", "1.04.03.15", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmMorningReport, "Morning Report", "1.04.03.16", clsUserMgtCode.SubModulePurchaseReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RM_Consumption_Detail, "RM Consumption Breakup", "1.04.03.20", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.AddCharge, "Additional Charge Report", "1.04.03.17", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.Vendor_Rating_Rejection, "Vendor Rating Rejection", "1.04.03.18", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.Parti_VS_Rejected, "Partywise Receipts V/S Rejection Detail", "1.04.03.19", clsUserMgtCode.SubModulePurchaseReport, 37)



        InsertDefaultValue(clsUserMgtCode.frmPurchaseOrderList, "Purchase Order List", "1.04.03.20", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPendingSaleInvoiceforChilpPO, "Pending Sale Invoice To be Received", "1.04.03.21", clsUserMgtCode.SubModulePurchaseReport, 37)
        '=================Added By Preeti Gupta
        InsertDefaultValue(clsUserMgtCode.rptMaterialSendforJobWork, "Material send For job work", "1.04.03.22", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMaterialReceivedAfterJobWork, "Material received after job work", "1.04.03.23", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBalanceStockForJobWork, "Balance stock For job work", "1.04.03.24", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptRGPWiseJobWork, "RGP wise job work", "1.04.03.25", clsUserMgtCode.SubModulePurchaseReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptSparePartStatus, "Spare Part Status", "1.04.03.26", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSrnReport, "SRN Report", "1.04.03.27", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptUnpostedPO, "Unposted Purchase Order", "1.04.03.28", clsUserMgtCode.SubModulePurchaseReport, 37)

        'KUNAL > TICKET : BM00000010298 > CLIENT : UDL > DATE : 28-NOV-2016
        InsertDefaultValue(clsUserMgtCode.FrmRpt_OutStnd_Items_RGP, "OutStanding Items against RGP", "1.04.03.29", clsUserMgtCode.SubModulePurchaseReport, 37)

        '=======Sanjeet (11/01/2017)=================
        InsertDefaultValue(clsUserMgtCode.RptPendingPO, "PO Pending Report", "1.04.03.30", clsUserMgtCode.SubModulePurchaseReport, 37)
        '======================================

        '=======Sanjeet (12/01/2017)=================
        InsertDefaultValue(clsUserMgtCode.RptIssueReturnHirerachyWise, "Issue Return Hirerachy Wise", "1.04.03.31", clsUserMgtCode.SubModulePurchaseReport, 37)

        ' DATE : 14-FEB-2017 , CLIENT : UDL === Monthly Consumption Report =======
        InsertDefaultValue(clsUserMgtCode.frmMonthlyConsumptionReport, "Monthly / High-Value Consumptions ", "1.04.03.32", clsUserMgtCode.SubModulePurchaseReport, 37)

        '======================================
        InsertDefaultValue(clsUserMgtCode.frmPurchaseOrderAmd, "PO Amendment Report", "1.04.03.33", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCapexRegister, "Capex Purchase Register", "1.04.03.34", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptDayWisePurchasePriceReport, "Day Wise Purchase Price Report", "1.04.03.35", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPurchasePlanReport, "Purchase Plan Report", "1.04.03.36", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPurchaseMaterialRegister, "Purchase Material Register Report", "1.04.03.37", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptJobWorkDebitNoteReport, "Job Work Debit Note Report", "1.04.03.38", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmCostCentreConsumptionRpt, "Cost Centre Consumption Report", "1.04.03.39", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPOAgainstDocument, "Purchase Order Against Document List", "1.04.03.40", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailPurchase, "Purchase Audit Trail Report", "1.04.03.41", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBillChecklist, "Bill Checklist", "1.04.03.42", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCancelledTransactions_Purchase, "Purchase Cancelled Transactions Report", "1.04.03.43", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmTenderTrackingReport, "Tender Tracking Report", "1.04.03.44", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmERPStatusTrackingReport, "ERP Status Tracking Report", "1.04.03.45", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptrlPenaltyRegister, "Penalty Register Report", "1.04.03.46", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPerformanceReport, "Performance Report", "1.04.03.47", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.VehicleUnloadingReport, "Vehicle Unloading Report", "1.04.03.48", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptRMUnloading, "RM Unloading Report", "1.04.03.49", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTransactionWiseStock, "Transaction Wise Stock", "1.04.03.50", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.Rptralwisematerialreceivedandpayment, "Ral Wise Material Received And Payment", "1.04.03.51", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.RmSecurityDeduction, "Rm Security Deduction", "1.04.03.52", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.WeightmentAutoAndManualReport, "Weightment Auto And Manual Report", "1.04.03.53", clsUserMgtCode.SubModulePurchaseReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptRejectionReport, "Rejection Detail Report", "1.04.03.57", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.CancelPurchaseInvoiceReport, "Cancel Purchase Invoice Report", "1.04.03.58", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmRALQCReport, "RAL QC Report", "1.04.03.59", clsUserMgtCode.SubModulePurchaseReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPIPaymentStatusReport, "PI Payment Status Report", "1.04.03.60", clsUserMgtCode.SubModulePurchaseReport, 37)
        '=================End=====================

        ' Transaction Report 
        InsertDefaultValue(clsUserMgtCode.SubModulePurchaseTransactionReport, "Transaction Report", "1.04.04", clsUserMgtCode.ModulePurchase, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.mbtnRGP_NRGP_Rpt, "RGP/NRGP Report", "1.04.04.01", clsUserMgtCode.SubModulePurchaseTransactionReport, 37)
        InsertDefaultValue(clsUserMgtCode.MRDAReport, "MRDA Report", "1.04.04.02", clsUserMgtCode.SubModulePurchaseTransactionReport, 37)
        InsertDefaultValue(clsUserMgtCode.DebitAdviseReport, "Debit Advise Report", "1.04.04.03", clsUserMgtCode.SubModulePurchaseTransactionReport, 37)
        InsertDefaultValue(clsUserMgtCode.PJVReport, "PJV Report", "1.04.04.05", clsUserMgtCode.SubModulePurchaseTransactionReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCapxRevHis, "Capex Budget Revision History", "1.04.04.06", clsUserMgtCode.SubModulePurchaseTransactionReport, 37)

#End Region

#Region "Material Management"
        '--------------Material Management----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleMaterial, "Material Management", "1.05", clsUserMgtCode.ExpertERP, "M", 23)
        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleMaterialSetup, "Setup", "1.05.01", clsUserMgtCode.ModuleMaterial, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.itemMaster, "Item Master For Finished Goods", "1.05.01.01", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmItemMasterRMOther, "Item Master For RM Other", "1.05.01.02", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.inventorySetting, "Inventory Setting", "1.05.01.03", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.itemGroups, "Item Group", "1.05.01.04", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.itemStructure, "Item Structure", "1.05.01.05", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.itemPurchaseAccount, "Purchase Account Sets", "1.05.01.06", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.itemSaleAccount, "Sales Account Sets", "1.05.01.05", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.GSTunitMeasure, "Gst Unit Of Measure", "1.05.01.08", clsUserMgtCode.SubModuleMaterialSetup, 27)

        InsertDefaultValue(clsUserMgtCode.unitMaster, "Unit Of Measure", "1.05.01.09", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.chapterhead, "Item Sub Group Type", "1.05.01.10", clsUserMgtCode.SubModuleMaterialSetup, 27)
        ' InsertDefaultValue(clsUserMgtCode.mbtnItemCategory, "Item Category", "1.05.01.10", clsUserMgtCode.SubModuleMaterialSetup, 27)
        ' InsertDefaultValue(clsUserMgtCode.mbtnItemSubCategory, "Item Sub Category", "1.05.01.11", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.packType, "Pack Type", "1.05.01.11", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.PriceComponentMasters, "Price Component Master", "1.05.01.12", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.PriceComponentMapping, "Price Component Mapping", "1.05.01.13", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.PricePlan, "Price Plan", "1.05.01.14", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.PriceMaster, "Price Master", "1.05.01.15", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.SchemeMaster, "Scheme Master", "1.05.01.16", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnBreakageHead1, "Breakage Head", "1.05.01.17", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.locationMaster, "Location Master", "1.05.01.18", clsUserMgtCode.SubModuleMaterialSetup, 27)
        ' InsertDefaultValue(clsUserMgtCode.ItemLocationDetails, "Item Location Details", "1.05.01.19", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.ItemReorderLevel, "Item Reorder Level", "1.05.01.19", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.ItemExciseMapping, "Item Exice Mapping", "1.05.01.20", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.ItemBasicPrice, "Item Basic Price", "1.05.01.21", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmStandardscheme, "Standard Scheme", "1.05.01.22", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmStandardRateItem, "Standard Rate Of Item", "1.05.01.23", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmItemCategoryLevel, "Item Category Level", "1.05.01.24", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmItemCategoryStructure, "Item Category Structure", "1.05.01.25", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBarCodeGenerator, "Bar Code Generator", "1.05.01.26", clsUserMgtCode.SubModuleMaterialSetup, 27) ''Expert
        InsertDefaultValue(clsUserMgtCode.WarrantyMaster, "Warranty Master", "1.05.01.27", clsUserMgtCode.SubModuleMaterialSetup, 27) ''Expert
        InsertDefaultValue(clsUserMgtCode.frmSchemeMasterNew, "Scheme Master", "1.05.01.28", clsUserMgtCode.SubModuleMaterialSetup, 27) ''Expert
        InsertDefaultValue(clsUserMgtCode.frmWeightConversion, "Weight Conversion", "1.05.01.29", clsUserMgtCode.SubModuleMaterialSetup, 27) ''Expert
        InsertDefaultValue(clsUserMgtCode.frmPriceGroupMapping, "Price Group Mapping", "1.05.01.30", clsUserMgtCode.SubModuleMaterialSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.SublocationMaster, "Sub Location Master", "1.05.01.32", clsUserMgtCode.SubModuleMaterialSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.ItemLocationMapping, "Transfer Location Mapping", "1.05.01.33", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmLocationCategoryLevel, "Location Category Level", "1.05.01.31", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmLocationCategoryStructure, "Location Category Structure", "1.05.01.32", clsUserMgtCode.SubModuleMaterialSetup, 27)
        ''richa 21/08/2014
        InsertDefaultValue(clsUserMgtCode.FrmCatalogMaster, "Catalog Master", "1.05.01.33", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPartNoMaster, "Part No. Master", "1.05.01.34", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.InvetorySourceCode, "Inventory Source Code", "1.05.01.35", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmItemTypeMaster, "Item Type Master", "1.05.01.36", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmHSNMaster, "HSN Master", "1.05.01.37", clsUserMgtCode.SubModuleMaterialSetup, 27)

        InsertDefaultValue(clsUserMgtCode.frmOverheadCostMaster, "Overhead Cost Master", "1.05.01.38", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmOverheadCostGroup, "Overhead Cost Group", "1.05.01.39", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmItemCostMapping, "Item Cost Mapping", "1.05.01.40", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmWeightUomMaster, "Weight Uom Master", "1.05.01.41", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.TypeOfItemMaster, "Type Of Item", "1.05.01.42", clsUserMgtCode.SubModuleMaterialSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBatchManufacturingMaster, "Batch Manufacturing Master", "1.05.01.43", clsUserMgtCode.SubModuleMaterialSetup, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleMaterialTransaction, "Transaction", "1.05.02", clsUserMgtCode.ModuleMaterial, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.Indent, "Transfer Indent", "1.05.02.01", clsUserMgtCode.SubModuleMaterialTransaction, 8) ''FMCG
        InsertDefaultValue(clsUserMgtCode.Transfer, "Transfer", "1.05.02.02", clsUserMgtCode.SubModuleMaterialTransaction, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, False, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)  ''FMCG
        InsertDefaultValue(clsUserMgtCode.CreateTransfer, "Create Transfer", "1.05.02.03", clsUserMgtCode.SubModuleMaterialTransaction, 8) ''FMCG
        InsertDefaultValue(clsUserMgtCode.FrmItemMcMapping, "Item MC Mapping", "1.05.02.04", clsUserMgtCode.SubModuleMaterialTransaction, 8) ''FMCG
        InsertDefaultValue(clsUserMgtCode.mbtnEmptyTrans, "Empty Transactions", "1.05.02.05", clsUserMgtCode.SubModuleMaterialTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False) ''FMCG
        ' InsertDefaultValue(clsUserMgtCode.mbtnProductionEntry, "Production Entry", "1.05.02.06", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.mbtnStoreAdjustment, "Store Adjustment", "1.05.02.07", clsUserMgtCode.SubModuleMaterialTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmWarehouseBreakage, "Warehouse Breakage", "1.05.02.08", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        ' InsertDefaultValue(clsUserMgtCode.FrmExpiryDateEntry, "Expired Item Entry", "1.05.02.09", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmPhysicalStock, "Physical Stock", "1.05.02.10", clsUserMgtCode.SubModuleMaterialTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmPhysicalStockMultipleLocation, "Physical Stock Multiple Location", "1.05.02.11", clsUserMgtCode.SubModuleMaterialTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.ChangeItemSerialNumber, "Change Item Serial Number", "1.05.02.12", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        ' InsertDefaultValue(clsUserMgtCode.frmAssemblies, "Assemblies/Disassemblies", "1.05.02.12", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        ' InsertDefaultValue(clsUserMgtCode.ItemStockConversion, "Item Stock Conversion", "1.05.02.13", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.TransferReturn, "Transfer Return", "1.05.02.14", clsUserMgtCode.SubModuleMaterialTransaction, 8) ''FMCG
        ' InsertDefaultValue(clsUserMgtCode.GatePassTransfer, "Gate Pass Transfer", "1.05.02.15", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        ' InsertDefaultValue(clsUserMgtCode.TransferCrateReceived, "Transfer Crate Received", "1.05.02.16", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmTransferGateOut, "Transfer Gate Out", "1.05.02.17", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmGateEntryReturnTransfer, "Gate Return Transfer", "1.05.02.18", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        'If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Description from TSPL_FIXED_PARAMETER WHERE CODE='" & clsFixedParameterCode.ForUDLOnly & "' AND TYPE='" & clsFixedParameterType.ForUDLOnly & "'")), "1") = CompairStringResult.Equal Then
        '    InsertDefaultValue(clsUserMgtCode.FrmJobWorkInventory, "Job Work Inventory", "1.05.02.18", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        '    InsertDefaultValue(clsUserMgtCode.frmRawMilkConsumtion, "Raw Milk Consumtion", "1.05.02.19", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        'End If
        InsertDefaultValue(clsUserMgtCode.frmGeneralWeighment, "General Weighment", "1.05.02.20", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmRMProcessLoss, "RM Process Loss", "1.05.02.21", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAdjProductionEntry, "Production Entry", "1.05.02.22", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.frmAdjProductionEntryQC, "Production Entry QC", "1.0 7.02.22", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.frmAdjProductionStoreEntry, "Production Store Entry", "1.05.02.23", clsUserMgtCode.SubModuleMaterialTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmStartBatchEntry, "Start Batch Entry", "1.05.02.23", clsUserMgtCode.SubModuleMaterialTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)

        ' Reports 
        InsertDefaultValue(clsUserMgtCode.SubModuleMaterialReport, "Report", "1.05.03", clsUserMgtCode.ModuleMaterial, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.mbtnItemMovement, "Inventory Movement Report", "1.05.03.01", clsUserMgtCode.SubModuleMaterialReport, 37)
        'InsertDefaultValue(clsUserMgtCode.ItemLocationDetailsReport, "Item Location Details Report", "1.05.03.02", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.ItemPrice, "Item Price", "1.05.03.03", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.StockRecoReport, "Stock Reconciliation Report", "1.05.03.04", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmStockDispatchReport, "Stock Dispatch Report", "1.05.03.05", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnStockAdjustmentReport, "Stock Adjustment Report", "1.05.03.06", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmAdjustmentStatusReport1, "Adjustment Status Report", "1.05.03.07", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnBreakageReport, "Breakage Report", "1.05.03.08", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.BreakageReportSummary, "Breakage Report Summary", "1.05.03.09", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.RoutewiseBreakageReport, "Routewise Breakage Report", "1.05.03.10", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.SchemeReport, "Scheme Report", "1.05.03.11", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.StockReportForFinishedGoods, "Stock Report For Finished Goods", "1.05.03.12", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmAdjustmentReport, "Adjustment Report", "1.05.03.13", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVehicleWiseLoadout, "Vehicle Wise Loadout", "1.05.03.14", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPendingIndentTransferReport, "Pending Indent Transfer Report", "1.05.03.15", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmIndentReport, "Pending Indent Report", "1.05.03.16", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmExpiredItemDetails, "Expired Item Details Report", "1.05.03.17", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.stockRecoNew, "Stock Reco", "1.05.03.18", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.UnionStockReport, "Union Stock Report ", "1.05.03.18", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.stockRecoBatch, "Stock Reco Batch ", "1.05.03.19", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBarCodeGenerator1, "Serial No BarCode Generator", "1.05.03.20", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmItemSerialTrackingReport, "Item Serial Tracking Report", "1.05.03.21", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmInventoryAgeingReport, "Inventory Ageing Report", "1.05.03.22", clsUserMgtCode.SubModuleMaterialReport, 37)
        'InsertDefaultValue(clsUserMgtCode.FrmItemListRpt, "Item List Report", "1.05.03.23", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmTransferRegister, "Transfer Register", "1.05.03.23", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.MISStockLedgerReport, "Stock ledger Report", "1.05.03.24", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmItemReloadReport, "Item ReOrder Level Report", "1.05.03.25", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.MeterialstockReco, "Meterial Stock Reco", "1.05.03.26", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmFatSnfStockReport, "FAT/SNF Stock Report", "1.05.03.27", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmStockAgeingAnalysisReport, "Stock Ageing Analysis Report", "1.05.03.28", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptTransporterProvisionReport, "Transporter Provision Report", "1.05.03.37", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDatewiseQtyFatSnfStockReport, "Datewise Milk Stock", "1.05.03.31", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmMCCMilkLossGain, "Milk Loss Gain", "1.05.03.32", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptUnpostingTransItemQty, "Item Commited Quantity Report", "1.05.03.33", clsUserMgtCode.SubModuleMaterialReport, 37)
        'InsertDefaultValue(clsUserMgtCode.stockRecoNewJR, "Stock Reco Jobwork", "1.05.03.31", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBatchItemReport1, "Batch Item Report", "1.05.03.38", clsUserMgtCode.SubModuleMaterialReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptItemPurchaseSet, "Item Purchase Accounts Report", "1.05.03.39", clsUserMgtCode.SubModuleMaterialReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptSaleAccountSetList, "Sale Account Set List", "1.05.03.40", clsUserMgtCode.SubModuleMaterialReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptItemSalePurchaseSet, "Item Sale Purchase Account Report", "1.05.03.41", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmItemCostUOM, "Item Cost Report", "1.05.03.39", clsUserMgtCode.SubModuleMaterialReport, 37) '
        InsertDefaultValue(clsUserMgtCode.AuditTrailMaterialManagement, "Material Management Audit Trail Report", "1.05.03.40", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.FATSNFGainLoss, "FAT/SNF Gain Loss", "1.05.03.41", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptProductReceivingReport, "Product Receiving Report", "1.05.03.42", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCustomerWiseStockReco, "Customer Wise Stock Reco", "1.05.03.43", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.ItemStockReport, "Item Stock Report", "1.05.03.44", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.ItemRateStatusReport, "Item Rate Status Report", "1.05.03.45", clsUserMgtCode.SubModuleMaterialReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDailySTAFGSStatement, "Daily STA FGS Statement Report", "1.05.03.46", clsUserMgtCode.SubModuleMaterialReport, 37)


        'InsertDefaultValue(clsUserMgtCode.rptStockReport, " Stocks Report", "1.05.03.45", clsUserMgtCode.SubModuleMaterialReport, 37)
        '
#End Region

#Region "Dairy Production"
        '--------------Dairy Production----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleProductionDairy, "Dairy Production", "1.06", clsUserMgtCode.ExpertERP, "M", 23)

        ' Master    
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionSetupDairy, "Setup", "1.06.01", clsUserMgtCode.ModuleProductionDairy, "SM", 24)

        ' InsertDefaultValue(clsUserMgtCode.ACCSETMFGDairy, "Account set", "1.06.01.01", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSectionMaster, "Section Master", "1.06.01.02", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.FrmStageMaster, "Stage Master", "1.06.01.03", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmPPLogSheetMaster, "QC Log Sheet Master", "1.06.01.04", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionLogSheet, "QC Log Sheet", "1.06.01.05", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSectionStageMapping, "Section Stage Mapping", "1.06.01.06", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialDairy, "Bill Of Material", "1.06.01.07", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmProfitCenter, "Profit Center", "1.06.01.08", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.FrmLineMaster, "Line Master", "1.06.01.09", clsUserMgtCode.SubModuleProductionSetupDairy, 27)


        InsertDefaultValue(clsUserMgtCode.frmParameterMaster, "Parameter Master", "1.06.01.12", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmParameterRangeMaster, "Parameter Range Master", "1.06.01.13", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmParameterRangeMasterForQC, "Parameter Range Master For QC", "1.06.01.14", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.ParameterValueMaster, "Parameter Value Master", "1.06.01.15", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmPriceChartBulkProc, "Price Chart Master (Bulk Procurement)", "1.06.01.16", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmVendorPriceChartMapping, "Vendor Price Chart Mapping", "1.06.01.17", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.FrmContractTanker, "Contractor Tanker", "1.06.01.18", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSupplierMaster, "Supplier Master", "1.06.01.19", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmDivertedContractor, "Diverted Contractor", "1.06.01.20", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkTypeMast, "Milk Type Master", "1.06.01.21", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmMilkGradeMaster, "Milk Grade Master", "1.06.01.22", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.TankerCleaningItem, "Tanker Cleaning Item", "1.06.01.24", clsUserMgtCode.SubModuleProductionSetupDairy, 27)
        InsertDefaultValue(clsUserMgtCode.frmShiftMasterBulk, "Shift Master", "1.06.01.25", clsUserMgtCode.SubModuleProductionSetupDairy, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionTransactionDairy, "Transaction", "1.06.02", clsUserMgtCode.ModuleProductionDairy, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmProductionPlanningDairy, "Production Planning", "1.06.02.01", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderDairy, "Batch Order", "1.06.02.02", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionIssueEntry, "Production Issue Entry ", "1.06.02.03", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStandardization, "Production Standardization ", "1.06.02.04", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.ProcessProductionStandardizationFinalQC, "Production Standardization Final QC", "1.06.02.05", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProcessProductionStageProcess, "Stage Process", "1.06.02.06", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntry, "Production Entry", "1.06.02.07", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)

        InsertDefaultValue(clsUserMgtCode.frmProductionEntryFinalQC, "Production Entry Final QC", "1.06.02.08", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.DariyProductionUploader, "Production Uploader", "1.06.02.09", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)

        InsertDefaultValue(clsUserMgtCode.frmWreckageBooking, "Wreckage Booking", "1.06.02.10", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmAssembDis, "Assemblies/Disassemblies", "1.06.02.11", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.JobWorkDispatchProduction, "Job Work Dispatch", "1.06.02.12", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProductionEntryWithoutBatch, "Production Entry(Without Batch)", "1.06.02.13", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmProcessProdReturn, "Production Return", "1.06.02.14", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransfer, "Silo Milk Transfer", "1.06.02.15", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransfer_JOBWORK, "Job Work Silo Milk Transfer", "1.06.02.16", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmMRPForProduction, "MRP", "1.06.02.17", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSiloMilkTransferUploader, "Silo Milk Transfer Uploader", "1.06.02.18", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionStoreRequest, "Production Store Request", "1.06.02.19", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)


        InsertDefaultValue(clsUserMgtCode.frmPOBulkProc, "PO Bulk Procurement", "1.06.02.20", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", True, True, False, "udl@321", Nothing)
        InsertDefaultValue(clsUserMgtCode.frmIntimation, "Intimation", "1.06.02.21", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry, "Gate Entry", "1.06.02.22", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", True, True, False, "udl@321", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmWeighment, "Weighment", "1.06.02.23", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, Nothing, Nothing, Nothing, Nothing, False, False, True, "", Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheck, "Quality Check", "1.06.02.24", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, Nothing, Nothing, Nothing, False, False, False, "erp@1234", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmUnloading, "Unloading", "1.06.02.25", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCleaning, "Cleaning", "1.06.02.26", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, Nothing, Nothing, Nothing, Nothing, False, False, True, "", Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmGateOut, "Gate Out", "1.06.02.27", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRN, "Bulk Milk SRN", "1.06.02.28", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.BulkMilkPurchaseInvoiceMultiple, "Bulk Milk Purchase Invoice Multiple", "1.06.02.29", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkPurchaseInvoice, "Bulk Milk Purchase Invoice", "1.06.02.30", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmBulkPurchaseUploader, "Bulk Purchase\Sale Uploader", "1.06.02.31", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        'InsertDefaultValue(clsUserMgtCode.SecondarySettingForQC, "Secondary Setting for QC", "1.06.02.13", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferIn, "Milk Transfer In", "1.06.02.32", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkSuppliers, "Milk Suppliers", "1.06.02.33", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCRouteMapping, "MCC Route Mapping", "1.06.02.34", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCSuperwiserMapping, "MCC Superviser Mapping", "1.06.02.35", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCSupplierMapping, "MCC Supplier Mapping", "1.06.02.36", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkCollection, "Milk Collection", "1.06.02.37", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkQualityCheck, "Milk Quality Check", "1.06.02.38", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkRateProcessingScheme, "Milk Rate Processing Scheme", "1.06.02.39", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmVehicleMovement, "Vehicle Movement", "1.06.02.40", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkBillGeneration, "Bill Generation", "1.06.02.41", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmProvisionEntry, "Provision Entry", "1.06.02.42", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmBulkMilkSRNReturn, "Bulk Milk SRN Return", "1.06.02.43", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.MccMilkTransferPrice, "MCC Milk Transfer Price", "1.06.02.44", clsUserMgtCode.SubModuleMCCMilkProcurementTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmMilkPurchaseReturn, "Bulk Milk Purchase Return", "1.06.02.45", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        'InsertDefaultValue(clsUserMgtCode.frmjobWorkDebitNote, "Job Work Debit Note", "1.06.02.28", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmMilkTransferInReturn, "Milk Transfer In Return", "1.06.02.46", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmQCSeparation, "QC Separation", "1.06.02.47", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.frmTankerProvision, "Tanker Provision", "1.06.02.48", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionGenerate, "Generate Milk Collection", "1.06.02.49", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.MilkCollectionDCSMultipleDaysMerge, "DCS Milk Collection Multiple Days Merge", "1.06.02.52", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.BulkProcurementUploader, "Bulk Procurement Uploader", "1.06.02.56", clsUserMgtCode.SubModuleProductionTransactionDairy, 8)
        InsertDefaultValue(clsUserMgtCode.ProductionShiftMgmt, "Shift Management", "1.06.02.57", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, 0, "", "", "", False, False, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmGateEntryTransaction, "Gate Entry Transaction", "1.06.02.58", clsUserMgtCode.SubModuleProductionTransactionDairy, "", 8, Nothing, Nothing, Nothing, Nothing, False, False, True, "", Nothing, "", "Internal", "", "", "", False, "", 0, True, False)

        '----------------Preeti Gupta------------
        'Report-----------
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionReportDairy, "Report", "1.06.03", clsUserMgtCode.ModuleProductionDairy, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptproductionEntryReport, "Daily Production Report", "1.06.03.01", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptItemConsumptionReport, "Item Consumption Report", "1.06.03.02", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptProductionIssueStatusReport, "Production Issues Status Report", "1.06.03.03", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        'InsertDefaultValue(clsUserMgtCode.rptSectionWiseStockReport, "Section Wise Stock Report", "1.06.03.02", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptBatchStatus, "Batch Status Report", "1.06.03.04", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptWTPReport, " WIP Report", "1.06.03.05", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptDairyProductionWreckageReport, "Wreckage Report", "1.06.03.06", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        'InsertDefaultValue(clsUserMgtCode.rptJobWorkProduction, "Job Work Production Report", "1.06.03.07", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptAvailableQtyForProduction, "Available Quantity For Production Report", "1.06.03.08", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionCancelledTransation, "Production Cancelled Transactions Report", "1.06.03.09", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptStandardQCReport, "Standardization QC Report", "1.06.03.10", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptStandardActualConsumption, "Standard Vs Actual Consumption Report", "1.06.03.11", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptProductionQCReport, "Production QC Report", "1.06.03.12", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptProductionWiseStockReco, "Production Wise Stock Reco", "1.06.03.13", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptIssueWIPConsumptionReport, "Issue WIP Consumption Report", "1.06.03.14", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptProcessGainLossReport, "Process Gain Loss Report", "1.06.03.15", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailDairyProduction, "Dairy Production Audit Trail Report", "1.06.03.16", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptProductionStatusReport, "Production Status Report", "1.06.03.17", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionUtilityCost, "Production Utility Cost Report", "1.06.03.18", clsUserMgtCode.SubModuleProductionReportDairy, 37)

        'InsertDefaultValue(clsUserMgtCode.rptSecondaryQuality, "Secondary Quality", "1.06.03.01", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptWeighment, "Weighment Report", "1.06.03.19", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptTankerSummaryReport, "Tanker Summary Report", "1.06.03.20", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptMilkWeigmentRegister, "Milk Weigment Register Report", "1.06.03.21", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkStockLedgerSummary, "Milk Stock Ledger (Summary) Report", "1.06.03.22", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPendingProvisionReport, "Pending Provision Report", "1.06.03.23", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptBulkMilkRegister, "Bulk Milk Register", "1.06.03.24", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptBulkMilkMultiplePurchaseInvoice, " Bulk Milk Multiple Purchase Invoice", "1.06.03.25", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptVLCwiseTPTimeTable, "VLC Wise TP Table Report", "1.06.03.26", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        'InsertDefaultValue(clsUserMgtCode.rptMDConversion, "MD Conversion Report", "1.06.03..23", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        'InsertDefaultValue(clsUserMgtCode.rptTankerStatusReport, "Tanker Status Report", "1.06.03..20", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerStatusReport, "Tanker Status Report", "1.06.03.27", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailMilkProcurementBulk, "Milk Procurement Bulk Audit Trail Report", "1.06.03.28", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.RptTotalMilkProcurement, "Total Milk Procurement Report", "1.06.03.29", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptProductionStandardization, "Production Standardization Report", "1.06.03.30", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.ProductionEntryReport, "Production Entry Report", "1.06.03.31", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerProfitLossReport, "Tanker Profit Loss Report", "1.06.03.32", clsUserMgtCode.SubModuleProductionReportDairy, 37)
        InsertDefaultValue(clsUserMgtCode.rptBmcTankerInOutReport, "BMC Tanker In-Out Report", "1.06.03.33", clsUserMgtCode.SubModuleProductionReportDairy, 37)


#End Region

#Region "Standard Production"
        '--------------Standard Production----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleProductionSTD, "Standard Production", "1.07", clsUserMgtCode.ExpertERP, "M", 23)

        ' Master    
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionSetupSTD, "Setup", "1.07.01", clsUserMgtCode.ModuleProductionSTD, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.ACCSETMFGSTD, "Account set", "1.07.01.01", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.COSTMAINTAIN, "Cost Maintainance", "1.07.01.02", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.SETTSTD, "Settings", "1.07.01.03", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.EXPENSE, "Expense Head", "1.07.01.04", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.TOOLTYPE, "Tool Type", "1.07.01.05", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.frmWorkCenterMaster, "Work Center Master", "1.07.01.06", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.frmResourceMaster, "Resource Master", "1.07.01.07", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.TOOL, "Tool Master", "1.07.01.08", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.frmOperationMaster, "Operation Master", "1.07.01.09", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.ALTER, "Alternate Item", "1.07.01.10", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.FrmProcessMaster1, "Process Master", "1.07.01.11", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.FrmProductionLines, "Production Line", "1.07.01.07", clsUserMgtCode.SubModuleProductionSetupSTD, 27)
        InsertDefaultValue(clsUserMgtCode.frmBreakDownMaster, "Break Down Master", "1.07.01.13", clsUserMgtCode.SubModuleProductionSetupSTD, 27)


        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionTransactionSTD, "Transaction", "1.07.02", clsUserMgtCode.ModuleProductionSTD, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialCosting, "Bill Of Material", "1.07.02.01", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionPlanningSTD, "Production Planning", "1.07.02.02", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmManufacturingOrder, "Manufacturing Order", "1.07.02.03", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderSTD, "Batch Order", "1.07.02.04", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionRequisition, "Production Requisition", "1.07.02.05", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmStoreIssueSTD, "Store Issue", "1.07.02.06", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProductionReturnSTD, "Production Return", "1.07.02.07", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionReceiptSTD, "Production Receipt", "1.07.02.08", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.frmMRP, "MRP", "1.07.02.09", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmLabourWorkingSheet, "Labour Working Sheet", "1.07.02.10", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmDemoProductionPlanning, "Demo Production Planning ", "1.07.02.11", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionReceiptDemo, "Production Serialized Mapping ", "1.07.02.12", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmProductionItemSerialReplace, "Replace Serializing Entry", "1.07.02.13", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmRiceBOM, "Bill of Material(Rice)", "1.07.02.14", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmRiceMixingEntry, "Mixing Entry(Rice)", "1.07.02.15", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmRiceProcessingEntry, "Processing Entry(Rice)", "1.07.02.16", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.frmMRPAutoMobile, "MRP(STD)", "1.07.02.17", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmStanderdProductionEntry, "Std Production Entry", "1.07.02.18", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.frmBreakDownEntry, "Break Down Entry", "1.07.02.19", clsUserMgtCode.SubModuleProductionTransactionSTD, 8)
        InsertDefaultValue(clsUserMgtCode.FrmBagReceipt, "BAG Receipt Entry", "1.07.02.23", clsUserMgtCode.SubModuleProductionTransactionSTD, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)

        'Reports(Setup)
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionSetupReportSTD, "Setup Report", "1.07.03", clsUserMgtCode.ModuleProductionSTD, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.LALT, "AlternateItem", "1.07.03.01", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.LACCt, "AccountSet", "1.07.03.02", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.frmListOfBOM, "BOM", "1.07.03.03", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.LOIC, "ItemCost", "1.07.03.04", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.LOPER, "Operations", "1.07.03.05", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        'InsertDefaultValue(clsUserMgtCode.PRODREPORT, "ProductionLines", "1.07.03.06", clsUserMgtCode.SubModuleProductionSetupReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmListofRequisition, "Requisition", "1.07.03.06", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.Resource, "Resource", "1.07.03.07", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.LToolT, "Tool", "1.07.03.08", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.LTOOL, "ToolType", "1.07.03.09", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.LWC, "WorkCenter", "1.07.03.10", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionSerializedReport, "Serialized Report", "1.07.03.11", clsUserMgtCode.SubModuleProductionSetupReportSTD, 37)

        '' Reports(transactional)
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionReportSTD, "Report", "1.07.04", clsUserMgtCode.ModuleProductionSTD, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderReportSTD, "Batch Order Report", "1.07.04.01", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.FrmBatchReceiptSTD, "Batch Receipt", "1.07.04.02", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        'InsertDefaultValue(clsUserMgtCode.frmDatewiseProduction, "Datewise Production Report", "1.07.04.03", clsUserMgtCode.SubModuleProductionReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmIssueReturnItemWiseReportSTD, "Issue Return Item Wise Report", "1.07.04.04", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        'InsertDefaultValue(clsUserMgtCode.frmLineProductivity, "Line Productivity", "1.07.04.05", clsUserMgtCode.SubModuleProductionReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmOperaterEfficiencyReport, "Operator Efficiency Report", "1.07.04.06", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionPlanReportSTD, "Production Paln Report", "1.07.04.07", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionVarianceSTD, "Production Variance", "1.07.04.08", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailStandardProduction, "Standard Production Audit Trail Report", "1.07.04.09", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.rptSPItemConsumptionReport, "Material Consumption Report", "1.07.04.10", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.ProductionReport, "Production Report", "1.07.04.11", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        'InsertDefaultValue(clsUserMgtCode.QualitySummaryReport, "Quality Summary Report", "1.07.04.12", clsUserMgtCode.SubModuleProductionReportSTD, 37)

        InsertDefaultValue(clsUserMgtCode.FrmProductionAndSaleReport, "Production And Sale Report", "1.07.04.12", clsUserMgtCode.SubModuleProductionReportSTD, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)

#End Region

#Region "Quality Control"
        '=============Quality Control===================================================
        InsertDefaultValue(clsUserMgtCode.ModuleQualityControl, "Quality Control", "1.08", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleQCSetup, "Setup", "1.08.01", clsUserMgtCode.ModuleQualityControl, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmPPLogSheetMaster_QC, "QC Parameter Master", "1.08.01.01", clsUserMgtCode.SubModuleQCSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmQualityModuleParameterRangeMaster, "QC Parameter Range Master", "1.08.01.02", clsUserMgtCode.SubModuleQCSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVendorItemQCMapping, "QC Vendor-Item Mapping", "1.08.01.03", clsUserMgtCode.SubModuleQCSetup, 27)

        ''Transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleQCTransaction, "Transaction", "1.08.02", clsUserMgtCode.ModuleQualityControl, "SM", 26)
        'Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal trans As SqlTransaction) As Boolean
        InsertDefaultValue(clsUserMgtCode.VisualRandomQC, "Visual Random QC", "1.08.02.01", clsUserMgtCode.SubModuleQCTransaction, "", 27, 0, "", "", "", True, True, True, Nothing, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.NIRQC, "NIR QC", "1.08.02.02", clsUserMgtCode.SubModuleQCTransaction, "", 27, 0, "", "", "", True, True, True, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmQualityCheckForSRN, "Incoming Quality Check Entry", "1.08.02.03", clsUserMgtCode.SubModuleQCTransaction, "", 27, 0, "", "", "", True, True, True, Nothing, Nothing, "", "Internal", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmOutgoingQC, "Outgoing Quality Check Entry", "1.08.02.04", clsUserMgtCode.SubModuleQCTransaction, "", 27, 0, "", "", "", True, True, True, Nothing, Nothing, "", "Internal", "", "", "")
        InsertDefaultValue(clsUserMgtCode.frmQualityCheckApprovalForSRN, "Incoming Quality Approval Entry", "1.08.02.05", clsUserMgtCode.SubModuleQCTransaction, 27)

        InsertDefaultValue(clsUserMgtCode.SubModuleQCReport, "Report", "1.08.03", clsUserMgtCode.ModuleQualityControl, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptPendingQCReport, "QC Status Report", "1.08.03.01", clsUserMgtCode.SubModuleQCReport, 37)
        InsertDefaultValue(clsUserMgtCode.QualitySummaryReport, "Quality Summary Report", "1.08.03.02", clsUserMgtCode.SubModuleQCReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptQCAnalysisReport, "QC Analysis Report", "1.08.03.03", clsUserMgtCode.SubModuleQCReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptNIRQC, "NIR QC Report", "1.08.03.07", clsUserMgtCode.SubModuleQCReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmTruckWiseRejectionReport, "Truck Wise Rejection Report", "1.08.03.08", clsUserMgtCode.SubModuleQCReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmNIRQCvsWetQc, "NIRQC Vs WetQc", "1.08.03.08", clsUserMgtCode.SubModuleQCReport, 37)


        ''===========================end QC===========================================================
#End Region

#Region "Dairy Sales"
        '----------------------------------Dairy Sale---------------------------------------'

        InsertDefaultValue(clsUserMgtCode.ModuleSaleDairy, "Dairy Sales", "1.09", clsUserMgtCode.ExpertERP, "M", 23)

        '' setup
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleDairySetUp, "Setup", "1.09.01", clsUserMgtCode.ModuleSaleDairy, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFreshDS, "Sales Setting", "1.09.01.01", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.FrmZoneMasterDS, "Zone Master", "1.09.01.02", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmRouteMasterDS, "Route Master", "1.09.01.03", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.FrmFreshTransactionApprovalDS, "Fresh Transaction Approval", "1.09.01.04", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSchemeMasterDairyDS, "Scheme Master Dairy", "1.09.01.05", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmRouteFreightDetailsDS, "Route Freight Details", "1.09.01.09", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmLocationItemMapping, "Location Item Mapping", "1.09.01.07", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.SaleIncentiveMaster, "Customer Incentive Master", "1.09.01.09", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerDeduction, "Customer Deduction", "1.09.01.10", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.VehicleMasterForDairySale, "Vehicle Master", "1.09.01.11", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.SaleIncentiveMaster, "Customer Incentive Master", "1.09.01.08", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerDeduction, "Customer Deduction", "1.09.01.09", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.VehicleMasterForDairySale, "Vehicle Master", "1.09.01.10", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmDistributorCommission, "Distributor Commission", "1.09.01.11", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmDistributorRouteTagging, "Distributor Route Tagging", "1.09.01.12", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.FrmItemCapacityLimt, "Item Capacity Limit", "1.09.01.13", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBoothRouteMapping, "Booth Route Mapping", "1.09.01.14", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBoothCommissionMaster, "Booth Commission Master", "1.09.01.15", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmSubRouteMaster, "Sub Route Master", "1.09.01.16", clsUserMgtCode.SubModuleSaleDairySetUp, 27)

        '' transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleDairyTransaction, "Transaction", "1.09.02", clsUserMgtCode.ModuleSaleDairy, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmbookingdairy, "Dairy Booking", "1.09.02.01", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryOrderDairy, "Dairy Delivery Order", "1.09.02.02", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSaleDispatchDairy, "Dairy Dispatch", "1.09.02.03", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoicedairy, "Dairy Sale Invoice", "1.09.02.04", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturndairy, "Dairy Sale Return", "1.09.02.05", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmGatePassDairy, "Loadin Slip", "1.09.02.06", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedDairySale, "Crate Received Dairy ", "1.09.02.07", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmBookingDairyMultipleDistributor, "Dairy Booking Distributor ", "1.09.02.08", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmPOSBookingDairyMultipleDistributor, "POS Dairy Booking Distributor ", "1.09.02.09", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDairyGatePass, "Dairy GatePass", "1.09.02.10", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingCustomer, "Dairy Booking Customer", "1.09.02.11", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, True, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmPerformaInvoiceDairy, "Dairy Proforma Invoice", "1.09.02.12", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDairyFreshDispatchMultiple, "Dairy Fresh Dispatch Multiple", "1.09.02.13", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.CustomerIncentiveEntry, "Customer Incentive Entry", "1.09.02.15", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmGenerateBookingFreshSale, "Generate Booking", "1.09.02.17", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDairyBookingUploader, "Dairy Booking Uploader", "1.09.02.20", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCashIndentBookingMobApp, "Cash Indent Mob Booking", "1.09.02.21", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.POSDairySale, "POS Dairy Sale", "1.09.02.22", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmTranspoterDeduction, "Transporter Deduction Entry", "1.09.02.23", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmNotepadFileMatching, "Notepad Mismatch", "1.09.02.24", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDemandBooking, "Demand", "1.09.02.25", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDemandAdjustment, "Demand Adjustment", "1.09.02.26", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDemandApproval, "Demand Approval", "1.09.02.27", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDemand_Sheet, "Demand Sheet", "1.09.02.28", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDCSDEmandBooking, "DCS Demand", "1.09.02.29", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBoothDispatch, "Booth Dispatch", "1.09.02.29", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDemandUploader, "Demand Uploader", "1.09.02.30", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmProductDemandBooking, "Product Demand Booking", "1.09.02.31", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmProductDispatch, "Product Dispatch", "1.09.02.32", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmQuickDemand, "Quick Demand", "1.09.02.33", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLeakedSaleReturn, "Leaked Sale Return", "1.09.02.34", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMultipleInvoice, "Multiple Invoice", "1.09.02.35", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCommissionTPTCalculation, "Commission/TPT Calculation", "1.09.02.36", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductQuickDemand, "Product Quick Demand", "1.09.02.37", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCustomerPenalty, "Customer Penalty", "1.09.02.38", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmRouteWiseSaleTarget, "Route Wise Sale Target", "1.09.02.39", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmEwaybill, "Eway Bill Update", "1.09.02.40", clsUserMgtCode.SubModuleSaleDairyTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDispatchReturn, "Dispatch Return", "1.09.02.41", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmBoothCommission, "Booth Commission", "1.09.02.42", clsUserMgtCode.SubModuleSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        '' end of transaction

        '--Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleDairyReport, "Report", "1.09.03", clsUserMgtCode.ModuleSaleDairy, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.RptEffectiveRateReport1, "Effective Rate Report", "1.09.03.01", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPlantCustomerDemand, "Plant Customer Demand Report", "1.09.03.02", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptDairyBookingDistributorReport, "Dairy Booking Distributor Report", "1.09.03.03", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMatrixFreshSalesReportSaleDairy, "Matrix Fresh Sales Report", "1.09.03.04", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDairyTruckSheetReport, "Truck Sheet Report", "1.09.03.05", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleRegisterDetail, "Sale Register Detail", "1.09.03.06", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.FrmSaleVsReceipReport, "Sale vs Receipt", "1.09.03.07", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCustomerIncentiveEntry, "Turnover Discount Report", "1.09.03.08", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptInvoiceDetailReport, "Invoice Detail Report", "1.09.03.09", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptProvisionChart, "Provision Chart", "1.09.03.10", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesVehicleReport, "Sales Vehicle Report", "1.09.03.11", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBankDetailsWithVendorMargin, "Bank Details with Vendor Margin", "1.09.03.12", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptAbsentBooth, "Absent Booth", "1.09.03.13", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCustomerWiseSalesReport, "Customer Wise Sales Report", "1.09.03.14", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBookingReport, "Booking Report", "1.09.03.15", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBookingQtyAmtReport, "Demand Booking Report", "1.09.03.17", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesLedgerReport, "Sales Ledger Report", "1.09.03.18", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesStock, "Sales Stock Report", "1.09.03.18", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptAvgSaleDetailReport, "Average Sale Detail Report", "1.09.03.19", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptdairydispatch, "Dairy Dispatch Report", "1.09.03.20", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptRouteSaleRegister, "Route Sale Register", "1.09.03.21", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptSalesSummaryReport, "Sales Summary Report", "1.09.03.22", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptGatePassReport, "Gate Pass Report", "1.09.03.23", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptNewSalesReport, "New Sales Report", "1.09.03.24", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDemandReturnReport, "Demand Return Report", "1.09.03.25", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDailyStatementReport, "Daily Statement Report", "1.09.03.25", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptCreditCustomerReport, "Credit Customer Report", "1.09.03.26", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptD1D2Report, "D1D2 Report", "1.09.03.27", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptZoneWiseReport1, "Zone Wise Report", "1.09.03.28", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMonthWiseDMRReport, "Month Wise DMR Report", "1.09.03.29", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleRegisterUOMReport, "Sale Register UOM Wise Report", "1.09.03.30", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCommissionTPTReport, "Rpt Commission/TPT Report", "1.09.03.31", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmGatepassDetailReport, "Gatepass Detail Report", "1.09.03.32", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesComparisonReport, "Sales Comparison Report", "1.09.03.33", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmBoothDemandReport, "Booth Demand Report", "1.09.03.34", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmBoothCountReport, "Booth Count Report", "1.09.03.35", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDemandNill, "Booth Nill Demand Report", "1.09.03.36", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMonthlyBillSummaryReport, "Monthly Bill Summary Report", "1.09.03.37", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDayWiseBoothDemand, "Day Wise Booth Demand Report", "1.09.03.38", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSendBilltoDistributor, "Send Bill To Distributor", "1.09.03.39", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmRouteWiseSaleTargetReport, "Route Wise Sale Target Report", "1.09.03.40", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSchemeSaleReport, "Scheme Sale Report", "1.09.03.42", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RouteWiseDemandReport, "Route Wise Demand Report", "1.09.03.43", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCarteJaliRpt, "Crate Jali Report", "1.09.03.29", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptCustomerEffective_ItemRate, "Item Effective Price List", "1.09.03.30", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptDeliveryOrderReport1, "Delivery Order Report", "1.09.03.31", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDemandForsingleBranch, "Demand for the single Branch", "1.09.03.32", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCrateJalliReportForTransfer, "Crate Jalli Report For Transfer", "1.09.03.33", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmPrintDistributerInvoiceStatement, "Print Distributer Invoice", "1.09.03.34", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", True, True, False, "", Nothing, "Invoice Wise", "Party Wise", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDistributerLedgerReport, "Distributer Ledger", "1.09.03.35", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPendingBooking, "Pending Booking Report", "1.09.03.36", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCustomerZone, "Customer Zone Wise Report", "1.09.03.37", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCrateCanreceiptReport, "Crate/Can Receipt Report", "1.09.03.38", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDairySaleRegisterReport, "Dairy Sale Register Report", "1.09.03.39", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.InvoicePriceWiseReport, "Invoice Price Wise Report", "1.09.03.40", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptRouteWiseSaleRegister, "Route Wise Sale Register", "1.09.03.41", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptRouteWiseSaleRegister, "Route Wise Sale Register", "1.09.03.41", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.SaleEinvoiceReport, "Sale Einvoice Report", "1.09.03.42", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailDairySale, "Dairy Sale Audit Trail Report", "1.09.03.43", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.CancelledTransactionReportDS, "Dairy Sale Cancelled Transactions Report", "1.09.03.44", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDailyReceiptReport, "Daily Receipt Report", "1.09.03.45", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTransporterWiseReport, "Transporter Wise Report", "1.09.03.46", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptDetailedCardReport, "Detailed Card Report", "1.09.03.46", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDailyLeakageReplacementReport, "Daily Leakage Replacement Report", "1.09.03.47", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptUnpostedDocumentReport, "Unposted Document Report", "1.09.03.48", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmZoneWiseSKUReport, "Zone Wise SKU Report", "1.09.03.49", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDairySaleSchemeReport, "Dairy Sale Scheme Report", "1.09.03.50", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBookingWiseRegister, "Booking Wise Register", "1.09.03.51", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDemandHistory, "Demand History", "1.09.03.52", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.FrmPrintMultipleGatepass, "Print Multiple Gatepass", "1.09.03.52", clsUserMgtCode.SubModuleSaleDairyReport, "", 37, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptBoothTruckSheet, "Booth Truck Sheet", "1.09.03.53", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDemandReport, "Demand Report", "1.09.03.54", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDailyDemandReport, "Daily Demand Report", "1.09.03.55", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptHSNWiseSaleReport, "HSN Wise Sale Report", "1.09.03.56", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmperdayDetailRpt, "PER Day Detail Report", "1.09.03.59", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBookingVsDispatchReport, "Booking Pending Report", "1.09.03.60", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBoothWiseBillReport, "Booth Wise Bill Print", "1.09.03.61", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDayBook, "Day Book", "1.09.03.62", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMilkSupplySaleReport, "Milk Supply Sale Report", "1.09.03.62", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCustItemWiseSaleReport, "Customer Item Wise Sale Report", "1.09.03.63", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptItemAndShiftWiseSaleSummaryReport, "Item And Shift Wise Sale Summary Report", "1.09.03.64", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmQuickDemandReport, "Quick Demand Report", "1.09.03.65", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmIBPReport, "I B P Report", "1.09.03.66", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptNewTenderPartyListReport, "New Tender Party List Report", "1.09.03.67", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCrateRegister, "Crate Balance Register", "1.09.03.69", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleInvoiceStatusReport, "Sale Invoice Status Report", "1.09.03.71", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmEwaybill, "Eway Bill API", "1.09.03.72", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptItemWiseBillReport, "Customer Item Wise Bill Report", "1.09.03.73", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDistributorRouteMarginStatusReport, "Distributor Route Margin Status Report", "1.09.03.74", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPartWiseItemReport, "Day Wise Party Sale Report", "1.09.03.76", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAccountSalesReport, "Account Sales Report", "1.09.03.77", clsUserMgtCode.SubModuleSaleDairyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptProductCreditSaleReport, "Product Credit Sale Report", "1.09.03.77", clsUserMgtCode.SubModuleSaleDairyReport, 37)

        InsertDefaultValue(clsUserMgtCode.SubModuleVisualSaleReport, "Visual Report", "1.09.04", clsUserMgtCode.ModuleSaleDairy, "SM", 36)
        InsertDefaultValue(clsUserMgtCode.rptZoneWiseMSVisual, "Zone Wise Milk Sale Report", "1.09.04.01", clsUserMgtCode.SubModuleVisualSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.VisualReportSale, "Visual Report Sale", "1.09.04.02", clsUserMgtCode.SubModuleVisualSaleReport, 37)
        '-------------------------SERVICE MODULE END----------------------------------------
#End Region

#Region "Sales And Distribution 2"
        '--------------Sales And Distribution New----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleSalesNew, "Sales And Distribution", "1.10", clsUserMgtCode.ExpertERP, "M", 23)

        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleNewSetup, "Setup", "1.10.01", clsUserMgtCode.ModuleSalesNew, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.FrmSaleSetting, "Sales Setting", "1.10.01.01", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.customerItemDetails, "Item Price List", "1.10.01.02", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerItemDetailApproval, "Item Price List Approval", "1.10.01.03", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmItemPriceListLevel3, "Item Price List-Level3", "1.10.01.04", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frm_User_Customer_Rate_Settings, "User/Customer Rate Settings", "1.10.01.05", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSalesmanTarget, "Salesman Target", "1.10.01.10", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmApprovalSetting, "Approval Setting", "1.10.01.07", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.routetypemaster, "Route Type Master", "1.10.01.08", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.routeMaster, "Route Master", "1.10.01.09", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.vhicleMaster, "Vehicle Master", "1.10.01.10", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.transportMaster, "Transport Master", "1.10.01.11", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.customerItemMapping, "Buyer Item Mapping", "1.10.01.12", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSNShipmentImportExport, "Shipment Import Export", "1.10.01.13", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmRemarMaster, "Remark Master", "1.10.01.14", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDiscountMaster, "Discount Master", "1.10.01.15", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDCSforSale, "DCS Sale", "1.10.01.16", clsUserMgtCode.SubModuleSaleNewSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmFrieghtRateMaster, "Frieght Rate Master", "1.10.01.17", clsUserMgtCode.SubModuleSaleNewSetup, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleNewTransaction, "Transaction", "1.10.02", clsUserMgtCode.ModuleSalesNew, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmSaleQuotation, "Sales Quotation", "1.10.02.01", clsUserMgtCode.SubModuleSaleNewTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmSNSalesOrder, "Sales Order", "1.10.02.02", clsUserMgtCode.SubModuleSaleNewTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmSNShipment, "Shipment", "1.10.02.03", clsUserMgtCode.SubModuleSaleNewTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSNSaleInvoice, "Sale Invoice", "1.10.02.04", clsUserMgtCode.SubModuleSaleNewTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSNServiceInvoice, "Service Invoice", "1.10.02.05", clsUserMgtCode.SubModuleSaleNewTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmSNSaleReturn, "Sale Return", "1.10.02.06", clsUserMgtCode.SubModuleSaleNewTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSNPOS, "POS", "1.10.02.07", clsUserMgtCode.SubModuleSaleNewTransaction, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmPendingQuotationApproval, "Pending Quotation Approval", "1.10.02.08", clsUserMgtCode.SubModuleSaleNewTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmSNReceiptChallan, "Daily Receipt", "1.10.02.09", clsUserMgtCode.SubModuleSaleNewTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmAutoSTN, "Auto STN", "1.10.02.10", clsUserMgtCode.SubModuleSaleNewTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmProspect, "Prospect  Detail", "1.10.02.11", clsUserMgtCode.SubModuleSaleNewTransaction, 8)


        ' Report
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleNewReport, "Report", "1.10.03", clsUserMgtCode.ModuleSalesNew, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.rptSalesmanTarget, "Salesman Target", "1.10.03.01", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSaleOrderSummary, "Sale Order Summary", "1.10.03.02", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSaleOrderDetail, "Sale Order Detail", "1.10.03.03", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSaleInvoiceSummary, "Sale Invoice Summary", "1.10.03.04", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSaleInvoiceDetail, "Sale invoice Detail", "1.10.03.05", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmShipmentSummary, "Shipment Summary", "1.10.03.06", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmShipmentDetail, "Shipment Detail", "1.10.03.07", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSaleRegisterDemo, "Sales Register", "1.10.03.08", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.ORDNEW, "Order Tracking", "1.10.03.09", clsUserMgtCode.SubModuleSaleNewReport, 37)
        ' InsertDefaultValue(clsUserMgtCode.mbtnItemMovement, "Inventory Movement Report", "1.10.03.10", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesAnalysis, "Sales Analysis Report", "1.10.03.11", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCustomerBillWiseDetail, "Customer Bill Wise Detail", "1.10.03.12", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmRptFormOfGuarntee, "Form of Guarntee", "1.10.03.13", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCustomerWiseMonthlySalesAnalysis, "Customer Wise Monthly Sales Analysis", "1.10.03.14", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCustomerBillWiseDuesSummary, "Customer Bill Wise Dues Summary", "1.10.03.15", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmVendorGroupWiseSaleReport, "Vendor Group Wise Sale Report", "1.10.03.16", clsUserMgtCode.SubModuleSaleNewReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmPaySlipReport, "Pay Slip Report", "1.10.03.17", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDVAT32, "DVAT-32 Report", "1.10.03.18", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmVehiclePendingStatusRpt, "Vehicle Pending Status Report", "1.10.03.19", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSaleSummaryAgainstPO, "Sale Summary Against PO", "1.10.03.20", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmYearMonthWiseSaleComparison, "Year - Month Wise Sale Comparision", "1.10.03.21", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCompanyMonthWiseSaleComparison, "Customer Wise - Month Wise Sale Comparision", "1.10.03.22", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDistributerPerformance, "Distributor Performance", "1.10.03.23", clsUserMgtCode.SubModuleSaleNewReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptCustomerAgeingDrillDown, "Customer Aging (Drilldown)", "1.10.03.24", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmReceiptChallanReport, "Daily Receipt Report", "1.10.03.25", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSaleHistory, "Inquiry", "1.10.03.25", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmClaimReport, "Claim Report", "1.10.03.26", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDealerManagementReport, "Sales Dealer Report", "1.10.03.27", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmProspectDetailReport, "Prospect Detail Report", "1.10.03.28", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleRegisterForAdv, "Sales Register Report", "1.10.03.29", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesCancelledRegister2, "Sales Cancelled Register", "1.10.03.30", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailSaleAndDistribution, "Sale And Distribution Audit Trail Report", "1.10.03.31", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesReport, "Sales Report", "1.10.03.32", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDCSSaleRegister, "DCS Sale Register", "1.10.03.33", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.saleconsignee, "Sale Consignee", "1.10.03.34", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.RoutePartyWisePriceList, "Route and PartyWise Price List", "1.10.03.38", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleReportCustomerWise, "Sales Report Customer Wise", "1.10.03.41", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDealerSalesReport, "Dealer Sales Report", "1.10.03.42", clsUserMgtCode.SubModuleSaleNewReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPlantWiseAvailableStockQtyReport, "Plant Wise Available Stock Qty Report", "1.10.03.43", clsUserMgtCode.SubModuleSaleNewReport, 37)

#End Region

#Region "Receivable"
        '--------------Receivable----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleReceivable, "Receivable", "1.11", clsUserMgtCode.ExpertERP, "M", 23)

        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleReceivableSetup, "Setup", "1.11.01", clsUserMgtCode.ModuleReceivable, "SM", 24)

        ''richa Ticket no.BM00000003438 on 20/08/2014

        InsertDefaultValue(clsUserMgtCode.FrmReceivableSettings, "Receivable Settings", "1.11.01.01", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerType, "Customer Type", "1.11.01.02", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerAccountSet, "Customer Account Set", "1.11.01.03", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerGroup, "Customer Group", "1.11.01.04", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomeCategory, "Customer Category", "1.11.01.05", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerMaster, "Customer Master", "1.11.01.06", clsUserMgtCode.SubModuleReceivableSetup, "", 27, False, "", "", "", True, False, False, "", Nothing, "", "", "", "", "", False, "", 0)

        InsertDefaultValue(clsUserMgtCode.ShiptoLocation, "Ship To Location", "1.11.01.07", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmShipToLocationDetails, "Ship To Location Details", "1.11.01.07", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCustomerCategoryLevel, "Customer Category Level", "1.11.01.08", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCustomerCategoryStructure, "Customer Category Structure", "1.11.01.09", clsUserMgtCode.SubModuleReceivableSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmReceivablePaymentTerms, "Payment Terms", "1.11.01.10", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerLocationMapping, "Customer Location Mapping", "1.11.01.11", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.SecondaryCustomerMaster, "Retailer Master", "1.11.01.12", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmPOSGRoupMaster, "POS Group Master", "1.11.01.13", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmConsumerDetailsForm, "Consumer Master", "1.11.01.14", clsUserMgtCode.SubModuleReceivableSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmCompetitorMaster, "Competitor Master", "1.11.01.15", clsUserMgtCode.SubModuleReceivableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCustomerRouteShiftMaster, "Customer Route Shift Master", "1.11.01.16", clsUserMgtCode.SubModuleReceivableSetup, 27)

        ''Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleReceivableTransaction, "Transaction", "1.11.02", clsUserMgtCode.ModuleReceivable, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.ReceiptEntry, "Receipt Entry", "1.11.02.01", clsUserMgtCode.SubModuleReceivableTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        'InsertDefaultValue(clsUserMgtCode.FinanceAdjustment, "Adjustment Entry", "1.11.02.02", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.ReceiptAdjustmentEntry, "Receipt Adjustment Entry", "1.11.02.02", clsUserMgtCode.SubModuleReceivableTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.mbtnARInvoiceEntry, "AR Invoice Entry", "1.11.02.03", clsUserMgtCode.SubModuleReceivableTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.frmQuickBook, "Quick Book Entry", "1.11.02.04", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCustomerInquiry, "Customer Inquiry", "1.11.02.05", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersSetOff, "Customer Set Off", "1.11.02.06", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.FrmReceiptInvoiceMapping, "Receipt Invoice Mapping", "1.11.02.07", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCustomersOutstanding, "Customer Outstanding", "1.11.02.08", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmBankUpdateUploader, "Bank Update Uploader", "1.11.02.09", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmMakePayment, "Make Payment", "1.11.02.10", clsUserMgtCode.SubModuleReceivableTransaction, 8)
        ''Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleReceivableReport, "Report", "1.11.03", clsUserMgtCode.ModuleReceivable, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.CustomerGroupReport, "Customer Group Report", "1.11.03.01", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.CustomerDetails, "Outlet Details", "1.11.03.02", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerLedger, "Customer Ledger", "1.11.03.03", clsUserMgtCode.SubModuleReceivableReport, 37)
        'InsertDefaultValue(clsUserMgtCode.FrmCustomerOutstanding, "Customer Outstanding", "1.11.03.04", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmRoute_CustomerOutstanding, "Route And Customer Outstandings", "1.11.03.05", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.receiptreport, "Receipt Report", "1.11.03.06", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.CustomersListReport, "Customer List Report", "1.11.03.07", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCustomerAgeing, "Customer Ageing", "1.11.03.08", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.CustomerRouteHistoryReport, "Customer Route History", "1.11.03.09", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.cash_Register_Details4, "Cash Register Detail", "1.11.03.10", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerEmptyTrial, "Customer Empty Trial", "1.11.03.11", clsUserMgtCode.SubModuleReceivableReport, 37)
        ' InsertDefaultValue(clsUserMgtCode.FrmSecurityDeposit1, "Security Deposit", "1.11.03.12", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSecurityLevel, "Customer Security", "1.11.03.13", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCustomersSetOff, "Customer Set Off Report", "1.11.03.14", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCustomerAdvanceReg, "Customer Advance Register", "1.11.03.15", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRptCustomerTransList, "Customer Ageing Details", "1.11.03.16", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRptCustomerTransHistory, "Customer Transaction Report", "1.11.03.17", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptARReport, "AR Report", "1.11.03.18", clsUserMgtCode.SubModuleReceivableReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptShippmentClearing, "Shipment Clearing", "1.11.03.19", clsUserMgtCode.SubModuleReceivableReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptPayableClearing, "Payable Clearing", "1.11.03.20", clsUserMgtCode.SubModuleReceivableReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptCustomerAccountSet, "Customer Account Set Report", "1.11.03.19", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerLedgerZoneAreaWise, "Customer Ledger with Zone/Area Wise", "1.11.03.19", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailReceivable, "Receivable Audit Trail Report", "1.11.03.20", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.CustomerLedgerVsAgeing, "Customer Ledger Vs Ageing", "1.11.03.21", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.SundryDebtorsReport, "Sundry Debtors Report", "1.11.01.19", clsUserMgtCode.SubModuleReceivableReport, 37)
        InsertDefaultValue(clsUserMgtCode.AccountStatementReport, "Account Statement Report", "1.11.01.20", clsUserMgtCode.SubModuleReceivableReport, 37)

#End Region

#Region "Payables"
        '--------------Payables----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModulePayable, "Payables", "1.12", clsUserMgtCode.ExpertERP, "M", 23)

        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModulePayableSetup, "Setup", "1.12.01", clsUserMgtCode.ModulePayable, "SM", 24)


        ''Anubhooti 02-Sep-2014 Against Ticket No BM00000003425
        InsertDefaultValue(clsUserMgtCode.FrmPayableSettings, "Payable Settings", "1.12.01.01", clsUserMgtCode.SubModulePayableSetup, 27)

        InsertDefaultValue(clsUserMgtCode.vendoraccountset, "Vendor Account Set", "1.12.01.02", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.vendorgroup, "Vendor Group", "1.12.01.12", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.vendormaster, "Vendor Master", "1.12.01.04", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVendorCategoryLevel, "Vendor Category Level", "1.12.01.05", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVendorCategoryStructure, "Vendor Category Structure", "1.12.01.06", clsUserMgtCode.SubModulePayableSetup, 27)
        ''richa Ticket no.BM00000003438 on 20/08/2014
        'InsertDefaultValue(clsUserMgtCode.paymentTerms, "Payment Terms", "1.12.01.07", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VendorBankMaster, "Vendor Bank Master", "1.12.01.08", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.transportMasterVendor, "Transport Master", "1.12.01.09", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerVendorMappingVendor, "Customer Vendor Mapping", "1.12.01.10", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmHirerachyLevelMaster, "Hierarchy Level Master", "1.12.01.11", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.vendorSubgroup, "Vendor Sub Group", "1.12.01.12", clsUserMgtCode.SubModulePayableSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCustomerEmployeeMapping, "Customer Employee Mapping", "1.12.01.13", clsUserMgtCode.SubModulePayableSetup, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModulePayableTransaction, "Transaction", "1.12.02", clsUserMgtCode.ModulePayable, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.PaymentEntryNew, "Payment Entry", "1.12.02.01", clsUserMgtCode.SubModulePayableTransaction, "", 8, 0, "", "", "", False, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntry, "AP Invoice Entry", "1.12.02.02", clsUserMgtCode.SubModulePayableTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.FrmVendorInquiry, "Vendor Inquiry", "1.12.02.03", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmVendorService, "Vendor Service Charge", "1.12.02.04", clsUserMgtCode.SubModulePayableTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmVendorSetOff, "Vendor Set Off", "1.12.02.05", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmVSPCSASetOff, "VSP/CSA Set Off", "1.12.02.06", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmEmployeeSetOff, "Employee Set Off", "1.12.02.07", clsUserMgtCode.SubModulePayableTransaction, 8)

        '' Anubhooti 05-Sep-2014 BM00000003755
        InsertDefaultValue(clsUserMgtCode.FrmPaymentUploader, "Payment Uploader", "1.12.02.05", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.PaymentAdjustmentEntry, "Payment Adjustment Entry", "1.12.02.06", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmSupplierReg, "Supplier Registration", "1.12.02.07", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmApprovedSuppliers, "Approved Suppliers", "1.12.02.08", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCheckPrinting, "Check Printing", "1.12.02.09", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploader, "NEFT Uploader", "1.12.02.10", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.VendorRegistration, "Vendor Registration", "1.12.02.11", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmNEFTUploaderFarmer, "NEFT Uploader Farmer", "1.12.02.12", clsUserMgtCode.SubModulePayableTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmQuickPaymentBySingleCheque, "Quick Payment By Single Cheque", "1.12.02.13", clsUserMgtCode.SubModulePayableTransaction, 8)
        'Reports
        InsertDefaultValue(clsUserMgtCode.SubModulePayableReport, "Report", "1.12.03", clsUserMgtCode.ModulePayable, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.frmPaymentEntry, "Payment Entry Report", "1.12.03.01", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceReport, "AP Invoice Report", "1.12.03.02", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.VendorLedgerReport, "Vendor Ledger Report", "1.12.03.03", clsUserMgtCode.SubModulePayableReport, 37)
        '  InsertDefaultValue(clsUserMgtCode.FrmVendorsOutstandings, "Vendor Outstanding", "1.12.03.04", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAgingPayble, "AP Aging Report", "1.12.03.05", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmRptVendorList, "Vendor Lsit Report", "1.12.03.06", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmVendorListRPT, "Vendor List Report", "1.12.03.07", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMultipleRTGS, "Multiple RTGS Payment Report", "1.12.03.08", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptMultiplePaymentAdvice1, "Multiple Payment Advice Report", "1.12.03.09", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.VendorCustomerLedgerReport, "Customer Vendor Ledger Report", "1.12.03.10", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRptVendorTransList, "Vendor Transaction Report", "1.12.03.11", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVendorSecurity, "Vendor Security", "1.12.03.12", clsUserMgtCode.SubModulePayableReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmAgingDrillDown, "AP Aging(Drill Down)", "1.12.03.07", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRptAPInvoiceDetailsReport, "AP Invoice Details Report", "1.12.03.13", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptHierarchyWiseReport, "Hierarchy Wise Report", "1.12.03.14", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAdvancePaymentRegister, "Vendor Advance Register", "1.12.03.15", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRptVendorAgeingDetails, "Vendor Ageing Details", "1.12.03.16", clsUserMgtCode.SubModulePayableReport, 37)

        InsertDefaultValue(clsUserMgtCode.rptAPReport, "AP Report", "1.12.03.17", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptGSTRReport, "GSTR Report", "1.12.03.18", clsUserMgtCode.SubModulePayableReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptVendorAccountSetReport, "Vendor Account Set Report", "1.12.03.19", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailPayables, "Payables Audit Trail Report", "1.12.03.19", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.VendorPaymentDetails, "Vendor Payment Details Report", "1.12.03.20", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.VendorLedgerVsAgeing, "Vendor Ledger Vs Ageing", "1.12.03.03", clsUserMgtCode.SubModulePayableReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPaymentPayable, "Payment Payable Report", "1.12.03.21", clsUserMgtCode.SubModulePayableReport, 37)

#End Region

#Region "Tax Deducted At Source"
        '--------------Tax Deducted At Source----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleTDS, "Tax Deducted At Source", "1.13", clsUserMgtCode.ExpertERP, "M", 23)

        'Setup
        InsertDefaultValue(clsUserMgtCode.SubModuleTDSSetup, "Setup", "1.13.01", clsUserMgtCode.ModuleTDS, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.FinancialYear, "Financial Year", "1.13.01.01", clsUserMgtCode.SubModuleTDSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.ResponsiblePerson, "Responsible Person", "1.13.01.02", clsUserMgtCode.SubModuleTDSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.BranchDetails, "Branch Details", "1.13.01.03", clsUserMgtCode.SubModuleTDSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.NatureOfDeduction, "Nature Of Deduction", "1.13.01.04", clsUserMgtCode.SubModuleTDSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.PartyDetails, "Party Details", "1.13.01.05", clsUserMgtCode.SubModuleTDSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.TDSSection, "TDS Section", "1.13.01.06", clsUserMgtCode.SubModuleTDSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.StateCode, "State Code", "1.13.01.07", clsUserMgtCode.SubModuleTDSSetup, 27)


        'Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleTDSTransaction, "Transaction", "1.13.02", clsUserMgtCode.ModuleTDS, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.mbtnCreateRemittance, "Create Remittance", "1.13.02.01", clsUserMgtCode.SubModuleTDSTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.remittanceentry, "Remittance Entry", "1.13.02.02", clsUserMgtCode.SubModuleTDSTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.mbtnAPInvoiceEntryTDS, "AP Invoice TDS", "1.13.02.03", clsUserMgtCode.SubModuleTDSTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.TDSPAYMENT, "TDS Payment", "1.13.02.04", clsUserMgtCode.SubModuleTDSTransaction, 8)
        'Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleTDSReport, "Report", "1.13.03", clsUserMgtCode.ModuleTDS, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.frmrptTDSLedger, "TDS Ledger Report", "1.13.03.01", clsUserMgtCode.SubModuleTDSReport, 37)
        InsertDefaultValue(clsUserMgtCode.TDSForm26Q, "Form26/27 Report", "1.13.03.02", clsUserMgtCode.SubModuleTDSReport, 37)
        InsertDefaultValue(clsUserMgtCode.TDSSectionSummaryReport, "TDS Section Summary", "1.13.03.03", clsUserMgtCode.SubModuleTDSReport, 37)
        InsertDefaultValue(clsUserMgtCode.Form16AReport, "Form16 Report", "1.13.03.04", clsUserMgtCode.SubModuleTDSReport, 37) '
        InsertDefaultValue(clsUserMgtCode.AuditTrailTaxDeductedAtSource, "Tax Deducted At Source Audit Trail Report", "1.13.03.05", clsUserMgtCode.SubModuleTDSReport, 37)
#End Region

#Region "General Ledger"
        '--------------General Ledger----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleGL, "General Ledger", "1.14", clsUserMgtCode.ExpertERP, "M", 23)

        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleGLSetup, "Setup", "1.14.01", clsUserMgtCode.ModuleGL, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.glOptions, "GL Option", "1.14.01.01", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.segmentCode, "Segment Code Master", "1.14.01.02", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.accountStructure, "Account Structure Master", "1.14.01.03", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.AccountMainGroup, "Account Main Group", "1.14.01.14", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.accountGroup, "Account Group", "1.14.01.05", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.AccountSubGroup, "Account Sub Group", "1.14.01.06", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.AccountGLMainAccount, "GL Main Account", "1.14.01.07", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.createAccounts, "Create Account", "1.14.01.08", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.glAccount, "General Ledger Account", "1.14.01.09", clsUserMgtCode.SubModuleGLSetup, 27)

        InsertDefaultValue(clsUserMgtCode.sourceCode, "SOURCE CODE", "1.14.01.10", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.glsecurity, "GL Security", "1.14.01.11", clsUserMgtCode.SubModuleGLSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmGL_account_excluded, "GL Account Excluded For trial Balance", "1.14.01.12", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frm_Account_Mapping, "Account Mapping", "1.14.01.13", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmMapLedgerAccToTally, "Map Ledger Account To Tally", "1.14.01.14", clsUserMgtCode.SubModuleGLSetup, 27) 'Tally
        InsertDefaultValue(clsUserMgtCode.frmPostAllGLToTally, "Post All GL To Tally", "1.14.01.15", clsUserMgtCode.SubModuleGLSetup, 27) 'Tally
        InsertDefaultValue(clsUserMgtCode.frmProfitAndLossPerforma, "Profit And Loss Performa", "1.14.01.16", clsUserMgtCode.SubModuleGLSetup, 27) ''Demo
        InsertDefaultValue(clsUserMgtCode.frmBalanceSheetPerforma, "Balance Sheet Performa", "1.14.01.17", clsUserMgtCode.SubModuleGLSetup, 27) ''Demo
        'InsertDefaultValue(clsUserMgtCode.CashFlowPerforma, "Cash Flow Performa", "1.14.01.18", clsUserMgtCode.SubModuleGLSetup, 27) ''Demo
        InsertDefaultValue(clsUserMgtCode.FiscalYear, "Financial Year Master", "1.14.01.19", clsUserMgtCode.SubModuleGLSetup, 27) ''Demo
        InsertDefaultValue(clsUserMgtCode.CostCentreFinancial, "Cost Centre Master Financial", "1.14.01.20", clsUserMgtCode.SubModuleGLSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmGLControlAccountMapping, "Control Account Mapping", "1.14.01.21", clsUserMgtCode.SubModuleGLSetup, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleGLTransaction, "Transaction", "1.14.02", clsUserMgtCode.ModuleGL, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.journalEntry, "Journal Entry", "1.14.02.01", clsUserMgtCode.SubModuleGLTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)
        InsertDefaultValue(clsUserMgtCode.mbtnVCGLEntry, "Vendor/Customer/GL Entry", "1.14.02.02", clsUserMgtCode.SubModuleGLTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, True)

        InsertDefaultValue(clsUserMgtCode.frmJEReverse, "Reverse Journal Entry", "1.14.02.03", clsUserMgtCode.SubModuleGLTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.ReversejournalEntry, "Reverse Journal Entry", "1.14.02.04", clsUserMgtCode.SubModuleGLTransaction, 8)

        '' Reports 
        InsertDefaultValue(clsUserMgtCode.SubModuleGLReport, "Report", "1.14.03", clsUserMgtCode.ModuleGL, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.FrmCostCenterAnalysisRpt, "Cost Center Analysis", "1.14.03.01", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmGLTransReport, "GL Transaction", "1.14.03.02", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmJrnlVoucher, "Journal Voucher", "1.14.03.03", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTrialBalance, "Trial Balance", "1.14.03.04", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnJournalBook, "Journal Book", "1.14.03.05", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmUnpostedJV, "Pending Journal Voucher", "1.14.03.06", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptProfitAndLoss, "Profit And Loss", "1.14.03.07", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBalanceSheet, "Balance Sheet", "1.14.03.08", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.CashFlow, "Cash Flow", "1.14.03.09", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.JECheckSystem, "JE Check System", "1.14.03.10", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptChartOfAccount, "Chart of Accounts", "1.14.03.11", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTrialBalanceCV, "Trial Balance(Customer/Vendor)", "1.14.03.12", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailGeneralLedger, "General Ledger Audit Trail Report", "1.14.03.13", clsUserMgtCode.SubModuleGLReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCostCenterReport, "Cost Center Report", "1.14.03.14", clsUserMgtCode.SubModuleGLReport, 37)

#End Region

#Region "HR and Payroll"
        '--------------HR and Payroll----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleHR, "HR and Payroll", "1.15", clsUserMgtCode.ExpertERP, "M", 23)

        'Setup
        InsertDefaultValue(clsUserMgtCode.SubModuleHRSetup, "Setup", "1.15.01", clsUserMgtCode.ModuleHR, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmSalaryAccountSetting, "Salary Account Setting", "1.15.01.01", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCountryMaster, "Country Master", "1.15.01.02", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmStateMaster, "State Master", "1.15.01.03", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBranchMaster, "Branch Master", "1.15.01.04", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmOTMaster, "OT Master", "1.15.01.05", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmAttendanceMaster, "Attendance Master", "1.15.01.06", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBonusMaster, "Bonus Master", "1.15.01.07", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPayPeriodMaster, "Pay Period Master", "1.15.01.08", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPFRulesMaster, "PF Rules Master", "1.15.01.09", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmESIRulesMaster, "ESI Rules Master", "1.15.01.10", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmLeaveMaster, "Leave Master", "1.15.01.15", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmShiftMaster, "Shift Master", "1.15.01.12", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDepartmentMaster, "Department Master", "1.15.01.13", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDevisionMaster, "Division Master", "1.15.01.14", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCastCategoryMaster, "Cast Category Master", "1.15.01.15", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDocumentMaster, "Document Master", "1.15.01.16", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCourseMaster, "Course Master", "1.15.01.17", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmGradeMaster, "Grade Master", "1.15.01.18", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmLanguageMaster, "Language Master", "1.15.01.19", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmOccupationMaster, "Occupation Master", "1.15.01.20", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmReligionMaster, "Religion Master", "1.15.01.21", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSkillMaster, "Skill Master", "1.15.01.22", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmOTSlab, "OT Slab", "1.15.01.23", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmConveyanceRateMaster, "Conveyance Rate Master", "1.15.01.24", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmODMaster, "OD Master", "1.15.01.25", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPayrollSetting, "Payroll Setting", "1.15.01.26", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSubDepartmentMaster, "Sub Department Master", "1.15.01.14", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPayrollDesignationMaster, "Designation Master", "1.15.01.27", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPaymentMode, "Payment Mode Master", "1.15.01.28", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPTSlab, "PT Slab", "1.15.01.29", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmHRRSlab, "HRR Slab", "1.15.01.30", clsUserMgtCode.SubModuleHRSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmSavingsMaster, "Savings Master", "1.15.01.30", clsUserMgtCode.SubModuleHRSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmSectionAllowanceMaster, "Section/Allowance Master", "1.15.01.31", clsUserMgtCode.SubModuleHRSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmEmployeeSavingsMapping, "Employee Savings Mapping", "1.15.01.32", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeDeductionMaster, "Employee Deduction Master", "1.15.01.35", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmLICPolicyMaster, "LIC Policy Master", "1.15.01.37", clsUserMgtCode.SubModuleHRSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeOTEntry, "Employee OT Entry", "1.15.01.39", clsUserMgtCode.SubModuleHRSetup, 27)

        '-----------------Transaction---------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleHRTransaction, "Transaction", "1.15.02", clsUserMgtCode.ModuleHR, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmEmployee_Master, "Employee Master", "1.15.02.01", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeStatus, "Employee Status", "1.15.02.02", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmPayHeadDefinitions, "Pay Head Definitions", "1.15.02.03", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmSalaryStructure, "Salary Structure", "1.15.02.04", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMapPayHeadsToSalaStructure, "Map Pay Heads To Sala Structure", "1.15.02.05", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmEmpSalary, "Employee Salary", "1.15.02.06", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLeaveSetting, "Leave Setting", "1.15.02.07", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLeaveOpeningBalance, "Leave Opening Balance", "1.15.02.08", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLeaveStartingDateSetting, "Leave Starting Date Setting", "1.15.02.09", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmGeneralHolidays, "General Holidays", "1.15.02.10", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmWeeklyHolidays, "Weekly Holidays", "1.15.02.11", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmLeaveAllotment, "Leave Allotment", "1.15.02.12", clsUserMgtCode.SubModuleHRTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.frmLeaveApplication, "Leave Application", "1.15.02.13", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLeaveApplication, "Leave Application", "1.15.02.13", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmLeaveAdjustment, "Leave Adjustment", "1.15.02.14", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeIncrement, "Employee Increment", "1.15.02.15", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAllowanceDetails, "Allowance Details", "1.15.02.16", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDeductionDetails, "Deduction Details", "1.15.02.17", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmReimbursementDetails, "Employee Reimbursement", "1.15.02.18", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmApplyLoan, "Apply Loan", "1.15.02.19", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmArrear, "DA Arrear", "1.15.02.19", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmGenerateBonus, "Generate Bonus", "1.15.02.20", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLoanGeneration, "Loan Generation", "1.15.02.21", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmLoanAdjustment, "Loan Adjustment", "1.15.02.22", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDailyAttendance, "Daily Attendance", "1.15.02.23", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmHourlyAttendance, "Hourly Attendance", "1.15.02.24", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMonthlyAttendance, "Monthly Attendance", "1.15.02.25", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmOTSheet, "OT Sheet", "1.15.02.26", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAdjustmentVoucher, "Employee Adjustment Voucher", "1.15.02.27", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGeneration, "Salary Generation", "1.15.02.28", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeGratuity, "Employee Gratuity", "1.15.02.29", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmLTAClaim, "LTA Claim", "1.15.02.30", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMediclaimEntry, "Employee Mediclaim Entry", "1.15.02.31", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmFullAndFinalSettlement, "Full And Final Settlement", "1.15.02.32", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeShiftChange, "Employee Shift Change", "1.15.02.33", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmODSheet, "OD Sheet", "1.15.02.34", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmConveyanceClaim, "Conveyance Claim", "1.15.02.35", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmEmployeeTransfer, "Employee Transfer", "1.15.02.36", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmEPFund, "EPF", "1.15.02.37", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmEmpIncrement, "Employee Increment", "1.15.02.38", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmSentSalarySlip, "Sent Salary Slip", "1.15.02.39", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.FrmAllotmentOfLeaves, "ALlotment Of Leaves", "1.15.02.40", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.EmployeeBandMaster, "Employee Band Master", "1.15.02.41", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmDA_Arrear, "DA Arrear", "1.15.02.42", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmLeaveEncashment, "Leave Encashment", "1.15.02.43", clsUserMgtCode.SubModuleHRTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmLICEmployeeTagging, "Employee LIC Tagging", "1.15.02.45", clsUserMgtCode.SubModuleHRTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeOTCalculation, "Employee OT Calculation", "1.15.02.47", clsUserMgtCode.SubModuleHRTransaction, 8)


        'Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleHRReport, "Report", "1.15.03", clsUserMgtCode.ModuleHR, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.frmAdjustmentRegister, "Adjustment Register", "1.15.03.01", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAllownceRegister, "Allowance Register", "1.15.03.02", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAttendanceRegister, "Attendance Register", "1.15.03.03", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDeductionRegister, "Deduction Register", "1.15.03.04", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmEmp_Id, "Employee Id Card", "1.15.03.05", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeRegister, "Employee Register", "1.15.03.06", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmLabelPrinting, "Label Printing", "1.15.03.07", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmLeaveRegisterReport, "Leave Register", "1.15.03.08", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPF_Covering_Letter, "PF Covering Letter", "1.15.03.09", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmReimbursementRegister, "Reimbursement Register", "1.15.03.10", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGenerationRegister, "Salary Generation Register", "1.15.03.11", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmVoucherPaymentsRegister, "VoucherPayments", "1.15.03.12", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptEmployeeBday6, "Employee Bday/Anniversary", "1.15.03.13", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.EmployeeWiseReport, "Employee Wise Report", "1.15.03.14", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptEmployeeAdvanceLedger, "Employee Advance Ledger", "1.15.03.15", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptLoanStatement, "Loan Statement", "1.15.03.16", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryGenerationRegisterArrear, "Arrear Register", "1.15.03.17", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailHRAndPayroll, "HR And Payroll Audit Trail Report", "1.15.03.18", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmActiveInactiveEmpRegister, "Employee Active/Inactive Report", "1.15.03.19", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptEmployeeDeductionMasterReport, "Employee Deduction Master Report", "1.15.03.20", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTentativeReport, "Tentative Salary", "1.15.03.21", clsUserMgtCode.SubModuleHRReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptEmployeeOTReport, "Employee OT Report", "1.15.03.23", clsUserMgtCode.SubModuleHRReport, 37)

        'Monthly Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleHRMonthlyReport, "Monthly Report", "1.15.04", clsUserMgtCode.ModuleHR, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.frmAditionalEarning_DeductionReport, "Additional Earning / Deduction Report", "1.15.04.01", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAttendedDaysReport, "Attended Days Report", "1.15.04.02", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankStatement_Reports, "Bank Statement Reports", "1.15.04.03", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankSummary_Report, "Bank Summary Report", "1.15.04.04", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDeductionDetailsReport, "Deduction Details Report", "1.15.04.05", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmOT_Reports, "OT Reports", "1.15.04.06", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPaySlip_Reports, "Pay Slip Reports", "1.15.04.07", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryAbstractReport, "Salary Abstract Report", "1.15.04.08", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryCertificate, "Salary Certificate", "1.15.04.09", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryComponentDetails, "Salary Component Details", "1.15.04.10", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryIncrementReport, "Salary Increment Report", "1.15.04.11", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalarySheet_Reports, "Salary Sheet", "1.15.04.12", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSalaryVoucher_Reports, "Salary Voucher Report", "1.15.04.13", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmVarianceReport, "Variance Report", "1.15.04.14", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSalarySlipRpt, "Salary Slip Report", "1.15.04.15", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSalarySummaryRpt, "Salary Summary Report", "1.15.04.16", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmEmployeePFRpt, "Employee PF Organization", "1.15.04.17", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmNewSalCertificate, " New Salary Certificate", "1.15.04.18", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmESICRpt, " ESIC Report", "1.15.04.18", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptESIHalfYearly, " ESIC Half Year Report", "1.15.04.19", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptESICStatement, " ESIC Statement Report", "1.15.04.20", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptESICForm6, " ESIC Form 6 Report", "1.15.04.21", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptESICChallan, " ESIC Challan Report", "1.15.04.22", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptESICDeclarationForm, " ESIC Declaration Form Report", "1.15.04.23", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFForm2, " PF Form 2", "1.15.04.24", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFForm3A, " PF Form 3A", "1.15.04.25", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFForm5, " PF Form 5", "1.15.04.26", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFForm6, " PF Form 6A", "1.15.04.27", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFStatement, " PF Statement Report", "1.15.04.28", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFForm10, " PF Form 10", "1.15.04.29", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RPtPFForm11_Revised_, " PF Form11", "1.15.04.30", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFForm12A_revised_, " PF Form12A", "1.15.04.31", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPFChallanStatement, " PF Challan Statement", "1.15.04.32", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptKDILSalarySlip, " Salary Slip Report(New)", "1.15.04.33", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBOILetterReport, " BOI Letter Report", "1.15.04.34", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptForm22, " Form 22 Report ", "1.15.04.35", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBonusStatement, " Bonus Statement Report ", "1.15.04.36", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptForm34, " Form34 Report ", "1.15.04.37", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPerformaForContributiondetail, "Performa for Contribution Details ", "1.15.04.38", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBankTransferDetail, "Bank Transfer Detail ", "1.15.04.39", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDepartmentwiseSalarySheetRpt, "Departmentwise Salary Sheet", "1.15.04.40", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDetailOfWelfareFundAmount, "Detail Of Welfare Fund Amount Report", "1.15.04.41", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptFromNO21, "Form No 21", "1.15.04.42", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptActurialValuation, "Acturial Valuation", "1.15.04.43", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptAttendanceReport, "Attendance Report", "1.15.04.44", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptBiometricAttendanceReport, "Biometric Attendance Report", "1.15.04.45", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptEmployeeStatusReport, "Employee Status Report", "1.15.04.46", clsUserMgtCode.SubModuleHRMonthlyReport, 37)
        'Statutory Reports
        'InsertDefaultValue(clsUserMgtCode.SubModuleHRStatutoryReport, "Statutory Reports", "1.15.05", clsUserMgtCode.ModuleHR, "SM", 37)

        'InsertDefaultValue(clsUserMgtCode.frmAnnualCensusReport, "Annual Census Report", "1.15.05.01", clsUserMgtCode.SubModuleHRStatutoryReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmForm_T, "Form - T", "1.15.05.02", clsUserMgtCode.SubModuleHRStatutoryReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmPF_ESI_Reports, "PF / ESI Reports", "1.15.05,03", clsUserMgtCode.SubModuleHRStatutoryReport, 37)

#End Region


#Region "Fixed Asset"
        '--------------Fixed Assets----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleFixedAsset, "Fixed Assets", "1.16", clsUserMgtCode.ExpertERP, "M", 23)

        'Master
        InsertDefaultValue(clsUserMgtCode.SubModuleFixedAssetSetup, "Setup", "1.16.01", clsUserMgtCode.ModuleFixedAsset, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.AssetSegment, "Asset Segment", "1.16.01.01", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DepAccSets, "Account Sets", "1.16.01.02", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.Categories, "" & If(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, Nothing)) = "1", "Asset Group", "Asset Category") & "", "1.16.01.03", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CostFACenter, "CostCenter", "1.16.01.04", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmDepreciationField, "Depreciation Field", "1.16.01.05", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmAssetGroups, "" & If(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, Nothing)) = "1", "Asset Sub Group", "Asset Group") & "", "1.16.01.06", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDepreciationMethod, "Depreciation Method", "1.16.01.07", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DepPeriod, "Depreciation", "1.16.01.08", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.Template, "" & If(clsCommon.myCstr(clsFixedParameter.GetData(clsFixedParameterCode.ReadOnlyTemplateFieldsOnAcqusition, clsFixedParameterType.ReadOnlyTemplateFieldsOnAcqusition, Nothing)) = "1", "Asset Category Master", "Asset Template Master") & "", "1.16.01.13", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmAMAcquisitionCode, "Acquisition Codes", "1.16.01.16", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSecondaryCustomer, "Secondary Customer", "1.16.01.11", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.fixedsetting, "Fixed Setting", "1.16.01.12", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmAssetBookMaster, "Asset Book", "1.16.01.13", clsUserMgtCode.SubModuleFixedAssetSetup, 27)
        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleFixedAssetTransaction, "Transaction", "1.16.02", clsUserMgtCode.ModuleFixedAsset, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.FAAcquisitionEntry, "Acquisition Entry", "1.16.02.01", clsUserMgtCode.SubModuleFixedAssetTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FAAssetDepreciation, "Asset Depreciation", "1.16.02.02", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FADisposalEntry, "Disposal Entry", "1.16.02.03", clsUserMgtCode.SubModuleFixedAssetTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmVisi_Install_Pullout, "Market Equipment", "1.16.02.04", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAsset_Issue_Return, "Asset Issue/Return", "1.16.02.05", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetRequisition, "Asset Requisition", "1.16.02.06", clsUserMgtCode.SubModuleFixedAssetTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmAssetStoreRequistion, "Asset Store Requisition", "1.16.02.07", clsUserMgtCode.SubModuleFixedAssetTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmSecondaryCustomerSale, "Secondary Customer Sale", "1.16.02.08", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmIssueItemsToAsset, "Issue Items to Assemble Asset", "1.16.02.09", clsUserMgtCode.SubModuleFixedAssetTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FAAssetWork, "Asset Work Expanses", "1.16.02.10", clsUserMgtCode.SubModuleFixedAssetTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FAMergeAcquisitionEntry, "Asset Merging Entry", "1.16.02.11", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetAccountChange, "Asset Account Change", "1.16.02.12", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetDispatchRetailer, "Asset Dispatch Retailer", "1.16.02.13", clsUserMgtCode.SubModuleFixedAssetTransaction, 8)
        'Report
        InsertDefaultValue(clsUserMgtCode.SubModuleFixedAssetReport, "Report", "1.16.03", clsUserMgtCode.ModuleFixedAsset, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.FrmAssetRegister, "Asset Register", "1.16.03.01", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmAssetDetail, "Asset Detail", "1.16.03.02", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDisposalDetail, "Disposal Detail", "1.16.03.03", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmVisi_Install_Pullout_Report, "Market Equipment Report", "1.16.03.04", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmAsset_Issue_Return_Report, "Asset Issue/Return Report", "1.16.03.05", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDistributor_VS_SecondaryCustomer_Sale, "Distributor/Secondary Customer Sale", "1.16.03.06", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptFARReport, "FAR Report", "1.16.03.07", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCWIPReport, "CWIP Report", "1.16.03.08", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptALCRReport, "Capex Purchase Life Cycle Report", "1.16.03.09", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCapexConsumptionRpt, "Capex Consumption Report", "1.16.03.10", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailFixedAssets, "FixedAssets Audit Trail Report", "1.16.03.11", clsUserMgtCode.SubModuleFixedAssetReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptAssetDispatchRetailer, "Asset Dispatch Retailer Report", "1.16.03.12", clsUserMgtCode.SubModuleFixedAssetReport, 37)

#End Region

#Region "Sales And Distribution 1"
        '--------------Sales And Distribution----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleSales, "Sales And Distribution", "1.17", clsUserMgtCode.ExpertERP, "M", 23)
        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleSetup, "Setup", "1.17.01", clsUserMgtCode.ModuleSales, "SM", 24)


        InsertDefaultValue(clsUserMgtCode.routetypemaster, "Route Type Master", "1.17.01.02", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.vhicleMaster, "Vehicle Master", "1.17.01.03", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.transportType, "Transport Type", "1.17.01.04", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.channelCategory, "Channel Category", "1.17.01.17", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.channelMaster, "Channel Master", "1.17.01.06", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSalesManHierarchy, "SalesMan Hierarchy", "1.17.01.07", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCommissionMaster, "Commission Master", "1.17.01.08", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.routeMaster, "Route Master", "1.17.01.09", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSettlementMaster, "Settlement Master", "1.17.01.10", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.groupMasterRoute, "Group Master(Route)", "1.17.01.11", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.visiMaster, "Market Equipment Master", "1.17.01.12", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDiscountCategoryMaster, "Discount Category Master", "1.17.01.13", clsUserMgtCode.SubModuleSaleSetup, 27)
        '        InsertDefaultValue(clsUserMgtCode.frmDiscountMaster, "Discount Master", "1.17.01.14", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.Sampling_Master, "Sampling Master", "1.17.01.15", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnTmplateCreation, "Template Creation", "1.17.01.16", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmViewPunchingInvoice, "View Punching Invoice", "1.17.01.17", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnTargetMaster, "Target Master", "1.17.01.18", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRouteShifting, "Outlet Shifting", "1.17.01.19", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.CustomerVendorMapping, "Customer Vendor Mapping", "1.17.01.20", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.TDMTARGET, "TDM Wise Target", "1.17.01.21", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCustomerTargetFixing, "Customer Target Fixing", "1.17.01.22", clsUserMgtCode.SubModuleSaleSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmClaimMaster, "Claim Master", "1.17.01.23", clsUserMgtCode.SubModuleSaleSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmCompalintEntry, "Compalint Entry", "1.17.01.24", clsUserMgtCode.SubModuleSaleSetup, 27)

        '''''''''''''''''''' Zone Master
        'InsertDefaultValue(clsUserMgtCode.FrmZoneMaster, "Zone Master", "1.17.01.24", clsUserMgtCode.SubModuleSaleSetup, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleTransaction, "Transaction", "1.17.02", clsUserMgtCode.ModuleSales, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.saleOrders, "Sales Order", "1.17.02.01", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.LoadOut, "Load Out", "1.17.02.02", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.saleReturn, "Sales Return", "1.17.02.03", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.SaleReturnInterCompany, "Sales Return (Inter Company)", "1.17.02.04", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmQuickSettlement, "Quick Settlement", "1.17.02.05", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.LoadOutStatus, "Loadout Status", "1.17.02.06", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmSettlementEntry, "Settlement Entry", "1.17.02.07", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCompleteTransfer, "LoadOut-Sale Invoice Reconciliation", "1.17.02.08", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmTransferIncompleteRemarks1, "Transfer Incomplete Remarks", "1.17.02.09", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.ReverseEntry, "Reverse Quick Settlement", "1.17.02.10", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.ChangeInvoiceSalesman, "Change Invoice Salesman", "1.17.02.11", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmGatePassENtry1, "Gate Pass Entry", "1.17.02.12", clsUserMgtCode.SubModuleSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCheckSlipEntry, "Check Slip Entry", "1.17.02.13", clsUserMgtCode.SubModuleSaleTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.FrmBookingEntry, "Booking Entry", "1.17.02.14", clsUserMgtCode.SubModuleSaleTransaction, 8)


        ' Report
        InsertDefaultValue(clsUserMgtCode.SubModuleSaleReport, "Report", "1.17.03", clsUserMgtCode.ModuleSales, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.ItemDiscountReport, "Item Discount Report", "1.17.03.01", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.SaleAccountBreakDetail, "Trade Discount Report", "1.17.03.02", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.CashDiscount, "Cash Discount Report", "1.17.03.03", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.SaleAccountBreakage, "Sale Account Breakup Report", "1.17.03.04", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmMCDiscountReport, "Marginal Contribution Report", "1.17.03.05", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDiscountAnalysis, "Discount Anaylsis Report", "1.17.03.06", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSaleDiscount1, "Sale Reco chart Report", "1.17.03.07", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.crptLoadOut, "Sales Invoice", "1.17.03.08", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.nrptSales, "Sales Register", "1.17.03.09", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.SaleReport, "Sales Volume Report", "1.17.03.10", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCreditSaleReport, "Credit Sales Report", "1.17.03.11", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.NoSaleReport, "No Sale Report", "1.17.03.12", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.OtherPartySale, "Other Party Sale", "1.17.03.13", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnNetSaleReport, "Net Sale Report", "1.17.03.14", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptNetSaleDetailReport, "Net Sale Detail Wise", "1.17.03.15", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDistrbutorSaleTarget, "DistributorWiseSaleTarget", "1.17.03.16", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmDayReportDirectSale, "Day Report Direct Sale", "1.17.03.17", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.ItemCommissionSummary, "Item Commission Report", "1.17.03.18", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.DailySettlement, "Salesman Shortage", "1.17.03.19", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.ProvisionalSaleReport, "Provision Sale", "1.17.03.20", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.ProvSaleDetail, "Provision Sale Detail", "1.17.03.21", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnCustomerRanking, "Customer Ranking Report", "1.17.03.22", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPenetration, "Penetration Report", "1.17.03.23", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.PrimarySales, "Primary Sales Report", "1.17.03.24", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.SecondarySales, "Secondary Sales", "1.17.03.25", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnRouteSale, "Route Sale Report", "1.17.03.26", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.receiptFillreport, "Receipts Against Sales (FILLED)", "1.17.03.27", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.receiptWOTreport, "Invoice without receipt", "1.17.03.28", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmKPIReport, "KPI Report", "1.17.03.29", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmMismatchSettlement, "Mismatch Settlement", "1.17.03.30", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSaleOrderSummary, "Pending Sale Order", "1.17.03.31", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSettlementSheetReconcilationeport, "Settlement Sheet Reconcilation Report", "1.17.03.32", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.TransferRegister, "Transfer Register", "1.17.03.33", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPendingSettlement, "Pending Settlement", "1.17.03.34", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmLoadReport, "Combine Sale Report", "1.17.03.35", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmLoadOutInvoiceRecoReport, "LoadOut Invoice Reconcilation Report", "1.17.03.36", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmTargetReport1, "Target Report", "1.17.03.37", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmSettlement_CashMemoStatus, "Settlement/Cash Memo Status", "1.17.03.38", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmReverseSettlementDetail, "Reverse Settlement Detail", "1.17.03.39", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmMismatchCashMemo, "Mismatched Cash Memo", "1.17.03.40", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCanceledSaleInvoice, "Canceled Sale Invoice", "1.17.03.41", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmEmptyTransactionReport, "Empty Reco. Chart", "1.17.03.42", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmpendingLoadin, "Pending LoadIn", "1.17.03.43", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmRptSalesReturn, "Sale Return Register", "1.17.03.44", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmQuickSettlementHead, "Settlement Head Report", "1.17.03.45", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.LO_vs_Vechile, "Loadout V/S Vehicle Report", "1.17.03.46", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.KeyValue, "KeySale Report", "1.17.03.47", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmLeakageBreakage, "LeakageBreakage Report", "1.17.03.48", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.GatePass_Vs_actual, "GatePass V/S Actual Report", "1.17.03.49", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptInvoiceAgainstInward, "Invoice Against Inward Report", "1.17.03.50", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCustomerTargetReport, "Customer Target Report", "1.17.03.51", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptClaimMaster, "Claim Report", "1.17.03.52", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmTDMReport, "TDM Report", "1.17.03.53", clsUserMgtCode.SubModuleSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.SalesmanSalesOrderReport, "Salesman Sales Order", "1.17.03.54", clsUserMgtCode.SubModuleSaleReport, 37)

        ' shipping Report

        InsertDefaultValue(clsUserMgtCode.SubModuleSaleShippingReport, "Shippeing Report", "1.17.04", clsUserMgtCode.ModuleSales, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.FilloutwardRegisterReport1, "Filloutward Register Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.vehicle_Details_Report1, "Vehicle Details Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.visiDetail1, "Market Equipment Detail Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.OutletEmpty1, "Outlet Empty Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.VehiclewiseSale1, "Vehicle Wise Sale", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.mbtnVisiVPO1, "Visi VPO Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.Channelwisecustomer1, "Channel Wise Customer Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.RouteListReport11, "Route List Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.LoadOutStatusreport1, "Loadout Status Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmTransitBreakageReport1, "Transit Breakage Report", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmFilledOutWard, "Filled Transaction(Salesman)", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)
        InsertDefaultValue(clsUserMgtCode.EmptyReportDetail, "Empty Report Detail", "1.17.04.01", clsUserMgtCode.SubModuleSaleShippingReport, 37)


#End Region

#Region "Pepsi Production"
        '--------------Pepsi Production---------------------------------- 
        InsertDefaultValue(clsUserMgtCode.ModuleProductionPepsi, "Pepsi Production", "1.18", clsUserMgtCode.ExpertERP, "M", 23)

        ' Master    
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionSetupPepsi, "Setup", "1.18.01", clsUserMgtCode.ModuleProductionPepsi, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.ACCSETMFGPepsi, "Account set", "1.18.01.01", clsUserMgtCode.SubModuleProductionSetupPepsi, 27)
        InsertDefaultValue(clsUserMgtCode.ITEMCATEGORY, "Production Item Category", "1.18.01.02", clsUserMgtCode.SubModuleProductionSetupPepsi, 27)
        InsertDefaultValue(clsUserMgtCode.PRO, "Production Line", "1.18.01.03", clsUserMgtCode.SubModuleProductionSetupPepsi, 27)
        InsertDefaultValue(clsUserMgtCode.SETTPep, "Settings", "1.18.01.04", clsUserMgtCode.SubModuleProductionSetupPepsi, 27)
        InsertDefaultValue(clsUserMgtCode.frmBillOfMaterialPepsi, "Bill Of Material", "1.18.01.18", clsUserMgtCode.SubModuleProductionSetupPepsi, 27)
        InsertDefaultValue(clsUserMgtCode.frmBOMImport, "BOM Import", "1.18.01.06", clsUserMgtCode.SubModuleProductionSetupPepsi, 27)

        ' Transaction 
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionTransactionPepsi, "Transaction", "1.18.02", clsUserMgtCode.ModuleProductionPepsi, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmProductionPlanningPepsi, "Production Planning", "1.18.02.01", clsUserMgtCode.SubModuleProductionTransactionPepsi, 8)
        InsertDefaultValue(clsUserMgtCode.frmBatchOrderPepsi, "Batch Order", "1.18.02.02", clsUserMgtCode.SubModuleProductionTransactionPepsi, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionRequisition, "Production Requisition", "1.18.02.03", clsUserMgtCode.SubModuleProductionTransactionPepsi, 8)
        InsertDefaultValue(clsUserMgtCode.frmStoreIssuePepsi, "Store Issue", "1.18.02.04", clsUserMgtCode.SubModuleProductionTransactionPepsi, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionReturnPep, "Production Return", "1.18.02.05", clsUserMgtCode.SubModuleProductionTransactionPepsi, 8)
        InsertDefaultValue(clsUserMgtCode.frmProductionReceiptPepsi, "Production Receipt", "1.18.02.06", clsUserMgtCode.SubModuleProductionTransactionPepsi, 8)

        '' Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleProductionReportPepsi, "Report", "1.18.03", clsUserMgtCode.ModuleProductionPepsi, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.frmBatchOrderReportPepsi, "Batch Order Report", "1.18.03.01", clsUserMgtCode.SubModuleProductionReportPepsi, 37)
        InsertDefaultValue(clsUserMgtCode.FrmBatchReceiptPepsi, "Batch Receipt", "1.18.03.02", clsUserMgtCode.SubModuleProductionReportPepsi, 37)
        InsertDefaultValue(clsUserMgtCode.frmDatewiseProduction, "Datewise Production Report", "1.18.03.03", clsUserMgtCode.SubModuleProductionReportPepsi, 37)
        InsertDefaultValue(clsUserMgtCode.frmIssueReturnItemWiseReportPepsi, "Issue Return Item Wise Report", "1.18.03.04", clsUserMgtCode.SubModuleProductionReportPepsi, 37)
        InsertDefaultValue(clsUserMgtCode.frmLineProductivity, "Line Productivity", "1.18.03.05", clsUserMgtCode.SubModuleProductionReportPepsi, 37)
        'InsertDefaultValue(clsUserMgtCode.frmOperaterEfficiencyReport, "Operator Efficiency Report", "1.12.04.06", clsUserMgtCode.SubModuleProductionReportSTD, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionPlanReportPepsi, "Production Paln Report", "1.18.03.06", clsUserMgtCode.SubModuleProductionReportPepsi, 37)
        InsertDefaultValue(clsUserMgtCode.frmProductionVariancePepsi, "Production Variance", "1.18.03.07", clsUserMgtCode.SubModuleProductionReportPepsi, 37)

#End Region

#Region "MIS Reports"
        '----------------------------MIS REPORTS-------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleMISReports, "MIS Reports", "1.19", clsUserMgtCode.ExpertERP, "M", 23)
        InsertDefaultValue(clsUserMgtCode.SubModuleMISReports, "Report", "1.19.01", clsUserMgtCode.ModuleMISReports, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.MISDebtorReport, "MIS Debtor Analysis", "1.19.01.01", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISCreditorReport, "MIS Creditor Analysis", "1.19.01.02", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISStockReco, "MIS Stock Reco", "1.19.01.03", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegister, "MIS Sale Register", "1.19.01.04", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegisterWithCSASalePatti, "MIS Sale Register With Sale Patti", "1.19.01.05", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegisterWithCSASalePattiProductLocationWise, "Product And Location Wise Sale Report", "1.19.01.06", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISSaleRegisterWithCSASalePattiProductPackWise, "Product And Pack Wise Sale Report", "1.19.01.07", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.FrmDailySaleReport, "Sales Analysis Report", "1.19.01.08", clsUserMgtCode.SubModuleMISReports, "", 37)
        'InsertDefaultValue(clsUserMgtCode.rptSKUWiseSale, "SKU Wise Sale Report", "1.19.01.09", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.FrmMonthlySaleReport, "Monthly Sale Report", "1.19.01.09", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.FrmCustomerGroupOutstanding, "Customer Group Outstanding", "1.19.01.10", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.FrmCategoryAnalysisReport, "Category Analysis Report", "1.19.01.11", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.rptPromptMsgPendindDoc, "Declared Documents Report", "1.19.01.12", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.RptFlavouredMilk, "Flavoured Milk Report", "1.19.01.13", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.RptMonthWiseSaleAnalysis, "Month Wise Analysis Report", "1.19.01.14", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MISMassBalanceReport, "Mass Balance Report", "1.19.01.15", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.rptTCSLedger, "TCS Ledger", "1.19.01.16", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.RMStockConsumption, "Stock Consumption Report", "1.19.01.17", clsUserMgtCode.SubModuleMISReports, "", 37)
        InsertDefaultValue(clsUserMgtCode.MSIProductionSaleReport, "MIS ProductionSale Report", "1.19.01.18", clsUserMgtCode.SubModuleMISReports, "", 37)
        ' InsertDefaultValue(clsUserMgtCode.rptDCSSaleRegister, "DCS Sale Register", "1.19.01.17", clsUserMgtCode.SubModuleMISReports, "", 37)
        '----------------------------MIS REPORTS----------------------------------------ENDS HERE

#End Region

#Region "Business Intelligence"
        '--------------Business Intelligence----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleBI, "Business Intelligence", "1.20", clsUserMgtCode.ExpertERP, "M", 23)

        InsertDefaultValue(clsUserMgtCode.SubModuleBISetup, "Setup", "1.20.01", clsUserMgtCode.ModuleBI, "SM", 37)
        InsertDefaultValue(clsUserMgtCode.BICreateFilter, "Create BI Filter", "1.20.01.01", clsUserMgtCode.SubModuleBISetup, 37)
        InsertDefaultValue(clsUserMgtCode.BICreateReport, "Create BI Reports", "1.20.01.02", clsUserMgtCode.SubModuleBISetup, 37)
        InsertDefaultValue(clsUserMgtCode.BICreateDashBoard, "Create Dashboard", "1.20.01.03", clsUserMgtCode.SubModuleBISetup, 37)

        'Sale
        InsertDefaultValue(clsUserMgtCode.SubModuleBISale, "Sale", "1.20.02", clsUserMgtCode.ModuleBI, "SM", 37)

        InsertDefaultValue(clsUserMgtCode.BIMonthWiseSale, "Month Wise Sale", "1.20.02.01", clsUserMgtCode.SubModuleBISale, 37)
        InsertDefaultValue(clsUserMgtCode.BITopCustomer, "Top Customers", "1.20.02.02", clsUserMgtCode.SubModuleBISale, 37)
        InsertDefaultValue(clsUserMgtCode.BITopItemCategory, "Top Item Category", "1.20.02.03", clsUserMgtCode.SubModuleBISale, 37)
        InsertDefaultValue(clsUserMgtCode.BIDashBoadr, "Dashboard", "1.20.02.04", clsUserMgtCode.SubModuleBISale, 37)


        'Purchase
        InsertDefaultValue(clsUserMgtCode.SubModuleBIPurchase, "Purchase", "1.20.03", clsUserMgtCode.ModuleBI, "SM", 37)

        InsertDefaultValue(clsUserMgtCode.BIMonthWisePurchase, "Month Wise Purchase", "1.20.03.01", clsUserMgtCode.SubModuleBIPurchase, 37)
        InsertDefaultValue(clsUserMgtCode.BITopVendor, "Top Vendor", "1.20.03.02", clsUserMgtCode.SubModuleBIPurchase, 37)
        InsertDefaultValue(clsUserMgtCode.BITopItemCategoryPurchase, "Top Item Category Purchase", "1.20.03.03", clsUserMgtCode.SubModuleBIPurchase, 37)

        'Finance
        InsertDefaultValue(clsUserMgtCode.SubModuleBIFinance, "Fincance", "1.20.04", clsUserMgtCode.ModuleBI, "SM", 37)

        InsertDefaultValue(clsUserMgtCode.frmBankBookChart, "Bank/Cash Summary", "1.20.04.01", clsUserMgtCode.SubModuleBIFinance, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankBookClosing, "Bank/Cash Closing", "1.20.04.02", clsUserMgtCode.SubModuleBIFinance, 37)
        InsertDefaultValue(clsUserMgtCode.FrmCustomerAgingSummary, "Customer Ageing Summary", "1.20.04.03", clsUserMgtCode.SubModuleBIFinance, 37)
        InsertDefaultValue(clsUserMgtCode.frmVendorAgingSummary, "Vendor Ageing Summary", "1.20.04.04", clsUserMgtCode.SubModuleBIFinance, 37)
        InsertDefaultValue(clsUserMgtCode.BITopExpence, "Top Expenses", "1.20.04.05", clsUserMgtCode.SubModuleBIFinance, 37)

        'Production
        InsertDefaultValue(clsUserMgtCode.SubModuleBIProduction, "Production", "1.20.05", clsUserMgtCode.ModuleBI, "SM", 37)

        InsertDefaultValue(clsUserMgtCode.frmGraphicalBatchOrder, "Batch Order Status", "1.20.05.01", clsUserMgtCode.SubModuleBIProduction, 37)
        InsertDefaultValue(clsUserMgtCode.frmGraphicalCategorywiseProduction, "Graphical Categorywise Production", "1.20.05.02", clsUserMgtCode.SubModuleBIProduction, 37)


        'Asset
        InsertDefaultValue(clsUserMgtCode.SubModuleBIAsseet, "Assets", "1.20.06", clsUserMgtCode.ModuleBI, "SM", 37)

        InsertDefaultValue(clsUserMgtCode.BIMonthWiseAssset, "Month Wise Asset", "1.20.06.01", clsUserMgtCode.SubModuleBIAsseet, 37)


        'UserDefined
        InsertDefaultValue(clsUserMgtCode.SubModuleBIUserDefined, "BI Created Reports", "1.20.07", clsUserMgtCode.ModuleBI, "SM", 37)
        'InsertDefaultValue(clsUserMgtCode.BIMonthWiseAssset, "Month Wise Asset", "1.20.06.01", clsUserMgtCode.SubModuleBIUserDefined, 37)

#End Region

#Region "Project Management"
        '--------------PROJECT MANAGEMENT----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleProjectManagement, "Project Management", "1.21", clsUserMgtCode.ExpertERP, "M", 23)

        '' setup
        InsertDefaultValue(clsUserMgtCode.SubModuleProjectManagementSetup, "Setup", "1.21.01", clsUserMgtCode.ModuleProjectManagement, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmPJCSettings, "Settings", "1.21.01.01", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCostTypes, "Cost Types", "1.21.01.02", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPJCAccountSets, "Account Sets", "1.21.01.03", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmJobMaster, "Job Master", "1.21.01.04", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmTaskMaster, "Task Master", "1.21.01.05", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPJCEmployeeMaster, "Employee Master", "1.21.01.06", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmUserApproval, "User Approval Master", "1.21.01.07", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBudgetMaintenance, "Budget Maintenance Master", "1.21.01.08", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmExpenseType, "Expense Type Master", "1.21.01.09", clsUserMgtCode.SubModuleProjectManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmProjectStatus, "Project Status Master", "1.21.01.10", clsUserMgtCode.SubModuleProjectManagementSetup, 27)

        '' end of setup
        '' transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleProjectManagementTransaction, "Transaction", "1.21.02", clsUserMgtCode.ModuleProjectManagement, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.ProjectMaster, "Project Master", "1.21.02.01", clsUserMgtCode.SubModuleProjectManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmTimeSheet, "Time Sheet", "1.21.02.02", clsUserMgtCode.SubModuleProjectManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmUserLog, "User Logs", "1.21.02.03", clsUserMgtCode.SubModuleProjectManagementTransaction, 8)
        'InsertDefaultValue(clsUserMgtCode.frmAssemblies, "Assemblies/Disassemblies", "1.21.02.04", clsUserMgtCode.SubModuleProjectManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmPJCExpense, "PJC Expense Detail", "1.21.02.04", clsUserMgtCode.SubModuleProjectManagementTransaction, 8)
        '' end of transaction

        '--Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleProjectManagementReport, "Report", "1.21.03", clsUserMgtCode.ModuleProjectManagement, "SM", 30)

        InsertDefaultValue(clsUserMgtCode.frmProjectListReport, "Project List Report", "1.21.03.01", clsUserMgtCode.SubModuleProjectManagementReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmProjectDetails, "Project Details", "1.21.03.02", clsUserMgtCode.SubModuleProjectManagementReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmProjectProfitReport, "Project Profit Report", "1.21.03.03", clsUserMgtCode.SubModuleProjectManagementReport, 37)


#End Region

#Region "Service"
        '----------------------------------SERVICE---------------------------------------'

        InsertDefaultValue(clsUserMgtCode.ModuleService, "Service", "1.22", clsUserMgtCode.ExpertERP, "M", 23)

        '' setup
        InsertDefaultValue(clsUserMgtCode.SubModuleServiceSetup, "Setup", "1.22.01", clsUserMgtCode.ModuleService, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmComplaintGroupMaster, "Complaint Group Master", "1.22.01.01", clsUserMgtCode.SubModuleServiceSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPrimaryReasonMaster, "Primary Reason Master", "1.22.01.02", clsUserMgtCode.SubModuleServiceSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmComplaintMaster, "Complaint Master", "1.22.01.03", clsUserMgtCode.SubModuleServiceSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPendingReasonMaster, "Pending Reason Master", "1.22.01.04", clsUserMgtCode.SubModuleServiceSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmItemChargeCategoryMaster, "Charge Category Master", "1.22.01.05", clsUserMgtCode.SubModuleServiceSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmItemChargeFranchiseMappingMaster, "Item Charge Franchise Mapping Master", "1.22.01.06", clsUserMgtCode.SubModuleServiceSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmAssetServiceMaster, "Asset Master", "1.22.01.07", clsUserMgtCode.SubModuleServiceSetup, 27)
        '' end of setup

        '' transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleServiceTransaction, "Transaction", "1.22.02", clsUserMgtCode.ModuleService, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmAssetAgreement, "Asset Agreement", "1.22.02.01", clsUserMgtCode.SubModuleServiceTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetInstallPullOut, "Asset Install / PullOut", "1.22.02.02", clsUserMgtCode.SubModuleServiceTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmComplaintDetailEntry, "Complaint Detail Entry", "1.22.02.03", clsUserMgtCode.SubModuleServiceTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetDistatch, "Dispatch", "1.22.02.04", clsUserMgtCode.SubModuleServiceTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCartMaintenanceEntry, "Cart Maintenance Entry", "1.22.02.05", clsUserMgtCode.SubModuleServiceTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmPendingComplaintDetail, "Pending Complaint Detail", "1.22.02.06", clsUserMgtCode.SubModuleServiceTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmQuickComplaintDetailEntry, "Quick Complaint Detail Entry", "1.22.02.07", clsUserMgtCode.SubModuleServiceTransaction, 8)
        '' end of transaction

        '--Reports
        InsertDefaultValue(clsUserMgtCode.SubModuleServiceReport, "Report", "1.22.03", clsUserMgtCode.ModuleService, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.frmFranchiseChargesReport, "Franchise Charges Report", "1.22.03.01", clsUserMgtCode.SubModuleServiceReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCustomersListReport, " List Of Customers Report", "1.22.03.02", clsUserMgtCode.SubModuleServiceReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmPullOutRedeployReport, " PullOut/Redeploy Report", "1.22.03.03", clsUserMgtCode.SubModuleServiceReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmVendorBillDetails, " Vendor Bill Details", "1.22.03.04", clsUserMgtCode.SubModuleServiceReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmDaywisePendingComplaint, " DayWise Pending Complaint Report", "1.22.03.05", clsUserMgtCode.SubModuleServiceReport, 37)

        InsertDefaultValue(clsUserMgtCode.frmAssetDetailReport, "Asset Detail Report", "1.22.03.07", clsUserMgtCode.SubModuleServiceReport, 37)
        'InsertDefaultValue(clsUserMgtCode.frmspareStockReport2, " Spare Stock As On", "1.22.03.04", clsUserMgtCode.SubModuleServiceReport, 37)
        '-------------------------SERVICE MODULE END----------------------------------------'

#End Region

#Region "Human Resource"
        '' Anubhooti 05-Aug-2014
        '=================================================Human Resource================================================================
        InsertDefaultValue(clsUserMgtCode.ModuleHumanResource, "Human Resource", "1.23", clsUserMgtCode.ExpertERP, "M", 23)

        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResSetup, "Setup", "1.23.01", clsUserMgtCode.ModuleHumanResource, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmchkList, "HR Document Check List Master", "1.23.01.01", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.JobTitle, "HR Job Title Master", "1.23.01.02", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmHRParameterMaster, "HR Parameter Master", "1.23.01.03", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRoundMaster, "HR Round Master", "1.23.01.04", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmProfileMaster, "HR Profile Master", "1.23.01.05", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.EmployeeTypeMaster, "HR Employee Type Master", "1.23.01.06", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSourceTypeMaster, "Source Type Master", "1.23.01.07", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSourceTypeDetail, "Source Type Detail", "1.23.01.08", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmQualificationMaster, "HR Qualification Master", "1.23.01.09", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRelationMaster, "Relation Master", "1.23.01.10", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.AgencyMaster, "HR Agency Master", "1.23.01.11", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmDamageMaster, "Damage/Fine Master", "1.23.01.12", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DesignationMasterHierarchy, "Designation Master Hierarchy", "1.23.01.13", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.HRIndustryType, "HR Industry Type Master", "1.23.01.14", clsUserMgtCode.SubModuleHumResSetup, 27)
        InsertDefaultValue(clsUserMgtCode.HRVerticalMaster, "HR Vertical Master", "1.23.01.15", clsUserMgtCode.SubModuleHumResSetup, 27)
        ' Ticket No : TEC/19/03/19-000453 - remove screen from menu
        'InsertDefaultValue(clsUserMgtCode.FrmHRSettings, "HR Settings", "1.23.01.16", clsUserMgtCode.SubModuleHumResSetup, 27)
        ''  SetUp End
        ' Transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResTrans, "Recruitment", "1.23.02", clsUserMgtCode.ModuleHumanResource, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.HRBudgeting, "HR Recruitment Budgeting", "1.23.02.01", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.RequesitionEntry, "HR Requisition Entry", "1.23.02.02", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.RequesitionApproval, "HR Requisition Approval", "1.23.02.03", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.RequesitionClose, "HR Requisition Close", "1.23.02.04", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmApplicantEntry, "HR Applicant Entry", "1.23.02.05", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmShortlist, "HR Short List", "1.23.02.06", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmInterviewSchedule, "HR Interview Schedule", "1.23.02.07", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmInterviewFeedback, "HR Interview Feedback", "1.23.02.08", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.OfferChkList, "HR Document Check List ", "1.23.02.09", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmReferenceCheck, "HR Reference Check", "1.23.02.10", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmSalaryFitment, "HR Salary Fitment", "1.23.02.11", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmOfferLetterHR, "HR Offer Letter", "1.23.02.12", clsUserMgtCode.SubModuleHumResTrans, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.JOININGCHECKLIST, "HR Joining Check List", "1.23.02.13", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmAppointmentLetterHR, "HR Appointment Letter", "1.23.02.14", clsUserMgtCode.SubModuleHumResTrans, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmHireEmployee, "HR Hire Employee", "1.23.02.15", clsUserMgtCode.SubModuleHumResTrans, 8)
        InsertDefaultValue(clsUserMgtCode.frmDamageCaused, "Damage/Fine Detail", "1.23.02.16", clsUserMgtCode.SubModuleHumResTrans, 8)

        ' Training 
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResTraining, "Training", "1.23.03", clsUserMgtCode.ModuleHumanResource, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.TrainingMaster, "Training Master", "1.23.03.01", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.TrainingResourceMaster, "Training Resource Master", "1.23.03.02", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.TrainingInstituteMaster, "Institute Master", "1.23.03.03", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.TrainingRequestMaster, "Request For Training", "1.23.03.04", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.TRAINER_MASTER, "Trainer Master", "1.23.03.05", clsUserMgtCode.SubModuleHumResTraining, 8)

        InsertDefaultValue(clsUserMgtCode.Schedule_Training, "Schedule For Training Master", "1.23.03.06", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.Training_Attendance, "Training Attendance", "1.23.03.07", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.frmHrTraineeFeedBack, "Trainee Feedback", "1.23.03.08", clsUserMgtCode.SubModuleHumResTraining, 8)
        InsertDefaultValue(clsUserMgtCode.frmHrTrainerFeedBack, "Trainer Feedback", "1.23.03.09", clsUserMgtCode.SubModuleHumResTraining, 8)

        ' Performance Evaluation 
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResPerformance, "Performance Evaluation", "1.23.04", clsUserMgtCode.ModuleHumanResource, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.frmHRPerformanceCategoryMaster, "Performance Category", "1.23.04.01", clsUserMgtCode.SubModuleHumResPerformance, 8)
        InsertDefaultValue(clsUserMgtCode.frmHRPerformanceMaster, "Performance Master", "1.23.04.02", clsUserMgtCode.SubModuleHumResPerformance, 8)
        InsertDefaultValue(clsUserMgtCode.frmHRPerformanceGroup, "Performance Group", "1.23.04.03", clsUserMgtCode.SubModuleHumResPerformance, 8)
        InsertDefaultValue(clsUserMgtCode.frmHRPerformanceGroupMapping, "Performance Group Mapping", "1.23.04.04", clsUserMgtCode.SubModuleHumResPerformance, 8)
        InsertDefaultValue(clsUserMgtCode.frmHRPerformanceRating, "Performance Rating", "1.23.04.05", clsUserMgtCode.SubModuleHumResPerformance, 8)
        ' Reimbursement 
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResReimburse, "Reimbursement", "1.23.05", clsUserMgtCode.ModuleHumanResource, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmHRReimbursementTypeMaster, "Reimbursement Type Master", "1.23.05.01", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelPurposeMaster, "Travel Purpose Master", "1.23.05.02", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelCategoryMaster, "Travel Category Master", "1.23.05.03", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelBookedForMaster, "Travel Booked For Master", "1.23.05.04", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelModeTypeMaster, "Travel Mode Type Master", "1.23.05.05", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelCityMaster, "Travel City Master", "1.23.05.06", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelClassTypeMaster, "Travel Class Type Master", "1.23.05.07", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRHotelRatingMaster, "Hotel Rating Master", "1.23.05.08", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelRoomTypeMaster, "Travel Room Type Master", "1.23.05.08", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelCarTypeMaster, "Travel Car Type Master", "1.23.05.09", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRRaiseTravelRequisition, "Raise Travel Requisition", "1.23.05.10", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelReqApproval, "Travel Requisition Approval", "1.23.05.11", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelReimbursementExpense, "Travel Reimbursement Expense", "1.23.05.12", clsUserMgtCode.SubModuleHumResReimburse, 8)
        InsertDefaultValue(clsUserMgtCode.FrmHRTravelExpenseApproval, "Travel Reimbursement Expense Approval", "1.23.05.13", clsUserMgtCode.SubModuleHumResReimburse, 8)

        '=========Greivance===================Added by Rohit on Apr 20,2015=====
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResGrievancce, "Grievance", "1.23.06", clsUserMgtCode.ModuleHumanResource, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmGrievanceTypeMaster, "Grievance Type Master", "1.23.06.01", clsUserMgtCode.SubModuleHumResGrievancce, 8)
        InsertDefaultValue(clsUserMgtCode.frmGrievanceLogging, "Grievance Logging", "1.23.06.02", clsUserMgtCode.SubModuleHumResGrievancce, 8)
        InsertDefaultValue(clsUserMgtCode.frmGrievanceAllocation, "Grievance Allocation", "1.23.06.03", clsUserMgtCode.SubModuleHumResGrievancce, 8)
        InsertDefaultValue(clsUserMgtCode.frmGrievanceResolution, "Grievance Resolution", "1.23.06.04", clsUserMgtCode.SubModuleHumResGrievancce, 8)
        '=============================================

        '===================Added by Preeti Gupta on Apr 24,2015=====
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResExitManagment, "Exit Management", "1.23.07", clsUserMgtCode.ModuleHumanResource, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmHREMInterviewQuestion, "Interview Question Master", "1.23.07.01", clsUserMgtCode.SubModuleHumResExitManagment, 8)
        InsertDefaultValue(clsUserMgtCode.frmResignationLetter, "Resignation Letter", "1.23.07.02", clsUserMgtCode.SubModuleHumResExitManagment, 8)
        InsertDefaultValue(clsUserMgtCode.frmTerminationLetter, "Termination Letter", "1.23.07.03", clsUserMgtCode.SubModuleHumResExitManagment, 8)
        InsertDefaultValue(clsUserMgtCode.frmResignationAcceptanceOrRejection, "Resignation Acceptance Or Rejection", "1.23.07.04", clsUserMgtCode.SubModuleHumResExitManagment, 8)
        InsertDefaultValue(clsUserMgtCode.frmHREMExitInterview, "Exit Interview", "1.23.07.05", clsUserMgtCode.SubModuleHumResExitManagment, 8)

        '----------------------------------SERVICE---------------------------------------'


        '=========Greivance===================Added by Rohit on Apr 20,2015=====
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResEMployeeEquipment, "Employee Equipment Tracking", "1.23.09", clsUserMgtCode.ModuleHumanResource, "SM", 26)
        ' InsertDefaultValue(clsUserMgtCode.frmAssetTypeMaster, "Asset Type Master", "1.23.09.01", clsUserMgtCode.SubModuleHumResEMployeeEquipment, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetCategoryMaster, "Asset Category Master", "1.23.09.01", clsUserMgtCode.SubModuleHumResEMployeeEquipment, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetSubCategoryMaster, "Asset Sub Category Master", "1.23.09.02", clsUserMgtCode.SubModuleHumResEMployeeEquipment, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetMaster, "Asset Master", "1.23.09.03", clsUserMgtCode.SubModuleHumResEMployeeEquipment, 8)
        InsertDefaultValue(clsUserMgtCode.frmAssetIssueReturn, "Asset Issue/Return", "1.23.09.04", clsUserMgtCode.SubModuleHumResEMployeeEquipment, 8)
        '=============================================

        '=============================================shivani=======================================
        InsertDefaultValue(clsUserMgtCode.SubModuleHumResReport, "Reports", "1.23.10", clsUserMgtCode.ModuleHumanResource, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.RptRegisterOfDeduction, "Register Of Deduction", "1.23.10.01", clsUserMgtCode.SubModuleHumResReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmHRPerformanceRatingRpt, "Performance Rating", "1.23.10.02", clsUserMgtCode.SubModuleHumResReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailHumanResource, "Human Resource Audit Trail Report", "1.23.10.03", clsUserMgtCode.SubModuleHumResReport, 37)

#End Region

#Region "Service And Warranty"
        '' Anubhooti 28-Aug-2015
        '=================================================Service And Warranty================================================================
        InsertDefaultValue(clsUserMgtCode.ModuleServiceAndWarranty, "Service And Warranty", "1.24", clsUserMgtCode.ExpertERP, "M", 23)
        ' Setup 
        InsertDefaultValue(clsUserMgtCode.SubModuleServiceAndWarranty, "Setup", "1.24.01", clsUserMgtCode.ModuleServiceAndWarranty, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmFaultCategory, "Fault Category", "1.24.01.01", clsUserMgtCode.SubModuleServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmFaultMaster, "Fault Master", "1.24.01.02", clsUserMgtCode.SubModuleServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmServiceChargeMaster, "Service Charge Master", "1.24.01.03", clsUserMgtCode.SubModuleServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmProblemType, "Problem Type Master", "1.24.01.04", clsUserMgtCode.SubModuleServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmCallType, "Call Type Master", "1.24.01.05", clsUserMgtCode.SubModuleServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmActivityType, "Activity Type Master", "1.24.01.06", clsUserMgtCode.SubModuleServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSolutionKnowledgeBase, "Solution Knowledge Base", "1.24.01.07", clsUserMgtCode.SubModuleServiceAndWarranty, 27)

        InsertDefaultValue(clsUserMgtCode.SubModuleTransServiceAndWarranty, "Transaction", "1.24.02", clsUserMgtCode.ModuleServiceAndWarranty, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmServiceCall, "Service Call", "1.24.02.01", clsUserMgtCode.SubModuleTransServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmServiceEnquiry, "Service Enquiry", "1.24.02.02", clsUserMgtCode.SubModuleTransServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmServiceAllocation, "Service Allocation", "1.24.02.03", clsUserMgtCode.SubModuleTransServiceAndWarranty, 27)
        InsertDefaultValue(clsUserMgtCode.FrmServiceVisitDetails, "Service Visit Details", "1.24.02.04", clsUserMgtCode.SubModuleTransServiceAndWarranty, 27)
        '' -------------------------------------------------------------------------------------------------------------------------------

#End Region



#Region "Bulk Sale"
        ''changes by richa agarwal 03/11/2014

        InsertDefaultValue(clsUserMgtCode.ModuleBulkSale, "Bulk Sale", "1.25", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleBSSetup, "Setup", "1.25.01", clsUserMgtCode.ModuleBulkSale, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingBulk, "Sales Setting", "1.25.01.01", clsUserMgtCode.SubModuleBSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSalePriceChart, "Bulk Sale Price Chart", "1.25.01.02", clsUserMgtCode.SubModuleBSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBulkCreditLimitApproval, "Bulk Credit Limit Approval", "1.25.01.03", clsUserMgtCode.SubModuleBSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBulkTransactionApproval, "Bulk Transaction Approval", "1.25.01.04", clsUserMgtCode.SubModuleBSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.TankerMasterSale, "Tanker Master Sale", "1.25.01.05", clsUserMgtCode.SubModuleBSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.BulkSaleFreightMaster, "Bulk Sale Freight Master", "1.25.01.06", clsUserMgtCode.SubModuleBSSetup, 27)

        ''transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleBulkSaleTransaction, "Transaction", "1.25.02", clsUserMgtCode.ModuleBulkSale, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmSalesOrderBS, "Bulk Sales Order", "1.25.02.01", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmGateEntrySale, "Bulk Gate Entry", "1.25.02.02", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmWeighmentEntry, "Bulk Weighment Entry", "1.25.02.03", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        If clsCommon.CompairString(clsCommon.myCstr(clsDBFuncationality.getSingleValue("Select Description from TSPL_FIXED_PARAMETER WHERE CODE='" & clsFixedParameterCode.BulkSaleSequence & "' AND TYPE='" & clsFixedParameterType.BulkSaleSequence & "'")), "0") = CompairStringResult.Equal Then
            InsertDefaultValue(clsUserMgtCode.FrmLoadingTanker, "Bulk Loading Tanker", "1.25.02.04", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
            InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, "Bulk Quality Check", "1.25.02.05", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        Else
            InsertDefaultValue(clsUserMgtCode.FrmQualityCheckBulkSale, "Bulk Quality Check", "1.25.02.04", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
            InsertDefaultValue(clsUserMgtCode.FrmLoadingTanker, "Bulk Loading Tanker", "1.25.02.05", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        End If
        InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSale, "Bulk Dispatch", "1.25.02.06", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceBulkSale, "Bulk Invoice", "1.25.02.07", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTrade, "Bulk Dispatch Trade", "1.25.02.08", clsUserMgtCode.SubModuleBulkSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmTankerOut, "Bulk Tanker Out", "1.25.02.09", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmCreateAutoInvoiceBS, "Bulk Auto Invoice", "1.25.02.10", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmBulkDispatchReturnSale, "Bulk Dispatch Return", "1.25.02.11", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleReturn, "Bulk Sale Return", "1.25.02.12", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmCanSaleUploader, "Can Sale Uploader", "1.25.02.13", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmCanSale, "Can Sale", "1.25.02.14", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmCanReceived, "Can Received", "1.25.02.15", clsUserMgtCode.SubModuleBulkSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.FrmDispatchBulkSaleTradeReturn, "Bulk Dispatch Trade Return", "1.25.02.16", clsUserMgtCode.SubModuleBulkSaleTransaction, 8)
        ' Ticket No : KDI/26/04/18-000277  By prabhakar ( Tester Remarks )
        'InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryBulkSAle, "Sale Return Gate Entry Bulk Sale", "1.25.02.16", clsUserMgtCode.SubModuleBulkSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBulkSaleAcknowledgement, "Bulk Sale Acknowledgement", "1.25.02.17", clsUserMgtCode.SubModuleBulkSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBulkSaleAcknowledgementUploader, "Bulk Sale Acknowledgement Uploader", "1.25.02.18", clsUserMgtCode.SubModuleBulkSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBulkSaleFreightCalculation, "Bulk Sale Freight Calculation", "1.25.02.19", clsUserMgtCode.SubModuleBulkSaleTransaction, 8)

        ''Report
        InsertDefaultValue(clsUserMgtCode.SubModuleBulkSaleReport, "Report", "1.25.03", clsUserMgtCode.ModuleBulkSale, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.FrmPrintBulkInvoiceStatement, "Bulk Invoice Print Statement", "1.25.03.01", clsUserMgtCode.SubModuleBulkSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBulkSaleRegister, "Bulk Sale Register", "1.25.03.02", clsUserMgtCode.SubModuleBulkSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptBulkMultipleDispatch, "Bulk Dispatch Print Statement", "1.25.03.03", clsUserMgtCode.SubModuleBulkSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailBulkSale, "Bulk Sale Audit Trail Report", "1.25.03.04", clsUserMgtCode.SubModuleBulkSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.BulkSaleAcknowledgementUploaderReport, "Bulk Sale Acknowledgement Uploader Report", "1.25.03.05", clsUserMgtCode.SubModuleBulkSaleReport, 37)

        ''========================================
#End Region

#Region "CSA Sale"
        InsertDefaultValue(clsUserMgtCode.ModuleCSASale, "CSA Sale", "1.26", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleCSASetup, "Setup", "1.26.01", clsUserMgtCode.ModuleCSASale, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingCSA, "Sales Setting", "1.26.01.01", clsUserMgtCode.SubModuleCSASetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCSAPriceMaster, "CSA Price Master", "1.26.01.02", clsUserMgtCode.SubModuleCSASetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCSAAccountSet, "CSA Account Set", "1.26.01.03", clsUserMgtCode.SubModuleCSASetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCSACommissionItemWise, "CSA Commission/Freight Itemwise Mapping", "1.26.01.04", clsUserMgtCode.SubModuleCSASetup, 27)

        'transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleCSASaleTransaction, "Transaction", "1.26.02", clsUserMgtCode.ModuleCSASale, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmCSABooking, "CSA Booking", "1.26.02.01", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCSARequest, "CSA Request", "1.26.02.02", clsUserMgtCode.SubModuleCSASaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCSADeliveryOrder, "CSA Delivery Order", "1.26.02.03", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCSATransfer, "CSA Transfer", "1.26.02.04", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCSASaleInvoice, "CSA Sale Invoice/Sale Patti", "1.26.02.05", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmCSATransferReturn, "CSA Transfer Return", "1.26.02.06", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmShortCloseDOCS, "Short Close DO CSA", "1.26.02.07", clsUserMgtCode.SubModuleCSASaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCSASalePattiReturn, "CSA Sale Patti Return", "1.26.02.08", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmCSATransferGateOut, "CSA Transfer Gate Out", "1.26.02.09", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmGateEntryReturnCS, "Gate Return CSA Sale", "1.26.02.10", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryCSATransfer, "Sale Return Gate Entry CSA Transfer", "1.26.02.11", clsUserMgtCode.SubModuleCSASaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'Report
        InsertDefaultValue(clsUserMgtCode.SubModuleCSAReport, "Report", "1.26.03", clsUserMgtCode.ModuleCSASale, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.frmCSADOReport, "CSA Delivery Order Report", "1.26.03.01", clsUserMgtCode.SubModuleCSAReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmCSATransferReport, "CSA Transfer Report", "1.26.03.02", clsUserMgtCode.SubModuleCSAReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCSASaleRegister, "CSA Sale Register Report", "1.26.03.03", clsUserMgtCode.SubModuleCSAReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCSACustomerLedger, "CSA Customer Ledger", "1.26.03.04", clsUserMgtCode.SubModuleCSAReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptCSAmonthlywisereport, "CSA Monthly Wise Analysis Report", "1.26.03.05", clsUserMgtCode.SubModuleCSAReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailCSASale, "CSA Sale Audit Trail Report", "1.26.03.06", clsUserMgtCode.SubModuleCSAReport, 37)
        ' 
        '-----------------------------------------
#End Region


#Region "Fresh Sale"
        InsertDefaultValue(clsUserMgtCode.ModuleFreshSale, "Fresh Sale", "1.27", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleFSSetup, "Setup", "1.27.01", clsUserMgtCode.ModuleFreshSale, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingFresh, "Sales Setting", "1.27.01.01", clsUserMgtCode.SubModuleFSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRouteMaster, "Route Master", "1.27.01.03", clsUserMgtCode.SubModuleFSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmFreshTransactionApproval, "Fresh Transaction Approval", "1.27.01.04", clsUserMgtCode.SubModuleFSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmFreshCreditLimitApproval, "Fresh Credit Limit Approval", "1.27.01.05", clsUserMgtCode.SubModuleFSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSalesLevelHierarchy, "Sales Level Hierarchy", "1.27.01.06", clsUserMgtCode.SubModuleFSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSalesHierarchy, "Sales Hierarchy", "1.27.01.07", clsUserMgtCode.SubModuleFSSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSalesHierarchyMapping, "Sales Hierarchy Mapping", "1.27.01.08", clsUserMgtCode.SubModuleFSSetup, 27)


        ' '' transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleFreshSaleTransaction, "Transaction", "1.27.02", clsUserMgtCode.ModuleFreshSale, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmBookingEntry, "Fresh Booking Entry", "1.27.02.01", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryNoteFreshSale, "Fresh Delivery", "1.27.02.02", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", False, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.FrmDispatchFreshSale, "Fresh Dispatch", "1.27.02.03", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmDispatchMultipleFreshSale, "Fresh Dispatch Multiple", "1.27.02.04", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmInvoiceFreshSale, "Fresh Invoice", "1.27.02.05", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmCreateReceived, "Fresh Crate Received", "1.27.02.06", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnFreshSale, "Fresh Sale Return", "1.27.02.07", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmGatePassFS, "Fresh GatePass Entry", "1.27.02.08", clsUserMgtCode.SubModuleFreshSaleTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmShortCloseDO, "Short Close DO", "1.27.02.09", clsUserMgtCode.SubModuleFreshSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCrateReceviedBox, "Crate Received Box", "1.27.02.10", clsUserMgtCode.SubModuleFreshSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmInvoiceCrateLinerDetail, "Invoice Crate Liner Detail", "1.27.02.11", clsUserMgtCode.SubModuleFreshSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryFS, "Sale Return Gate Entry Fresh Sale", "1.27.02.12", clsUserMgtCode.SubModuleFreshSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmSubsidyCreditNote, "Create Subsidy Credit Note", "1.27.02.13", clsUserMgtCode.SubModuleFreshSaleTransaction, 8)

        'report
        InsertDefaultValue(clsUserMgtCode.SubModuleFreshSaleReport, "Report", "1.27.03", clsUserMgtCode.ModuleFreshSale, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.FrmPrintFreshInvoice, "Fresh Invoice Print Statement ", "1.27.03.01", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptFreshSaleRegister1, "Fresh Sale Register", "1.27.03.02", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        'InsertDefaultValue(clsUserMgtCode.rptCrateAccounting, "Crate Accounting Report", "1.27.03.03", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCrateAccountingReport, "Crate Accounting ", "1.27.03.04", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVehicleCapacityFreshSaleReport, "Vehicle Capacity Fresh Sale Report ", "1.27.03.05", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSalesHierarchyReport, "Sales Hierarchy Report", "1.27.03.06", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVehicleWiseReport, "Vehicle Wise Report", "1.27.03.07", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptFreshBookingStatus, "Fresh Booking Status", "1.27.03.08", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        'prabhakar'
        InsertDefaultValue(clsUserMgtCode.rptZoneWiseFreshSaleReport, "Zone Wise Fresh Sale Report", "1.27.03.09", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptDispatchChallanReportFresh, "Dispatch Challan Report", "1.27.03.10", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptMatrixFreshSalesReport, "Matrix Fresh Sales Report", "1.27.03.11", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        ''Parteek
        InsertDefaultValue(clsUserMgtCode.rptPriceChartFreshSalesReport, "Fresh Price Chart Report", "1.27.03.12", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        ''End
        ''-------------------------------------------
        InsertDefaultValue(clsUserMgtCode.rptCrateLinerReport, "Crate Liner Report", "1.27.03.13", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleReturnGateEntryReport, "Sale Return Gate Entry Report", "1.27.03.14", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptSubsidyCredit, "Subsidy Credit Note", "1.27.03.15", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptSchemeDetail, "Scheme Details", "1.27.03.16", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptLeakageDetail, "Leakage Details", "1.27.03.17", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailFreshSale, "Fresh Sale Audit Trail Report", "1.27.03.18", clsUserMgtCode.SubModuleFreshSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmSaleAnalysisReport, "Sale Analysis Report", "1.27.03.19", clsUserMgtCode.SubModuleSaleReport, 37)
        ' 
#End Region

#Region "Product Sale"
        InsertDefaultValue(clsUserMgtCode.ModuleProductSale, "Product Sale", "1.28", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleProductSetup, "Setup", "1.28.01", clsUserMgtCode.ModuleProductSale, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingProduct, "Sales Setting", "1.28.01.01", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmCreditLimitApproval, "Credit Limit Approval", "1.28.01.02", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmExpiryDate, "Expiry Date", "1.28.01.03", clsUserMgtCode.SubModuleProductSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmTransactionApproval, "Transaction Approval", "1.28.01.04", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSchemeMasterDairy, "Scheme Master Dairy", "1.28.01.05", clsUserMgtCode.SubModuleProductSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmDistanceMappingMaster, "Distance Mapping", "1.28.01.06", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmTragetMaster, " Target Master Product Sale", "1.28.01.07", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmRouteFreightDetails, "Route Freight Details", "1.28.01.08", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.DispatchChecklist, "Dispatch Checklist", "1.28.01.09", clsUserMgtCode.SubModuleProductSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmPrintFreshInvoice, "Fresh Invoice Print Statement ", "1.28.01.08", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.VehicleMasterForProductSale, "Vehicle Master", "1.28.01.10", clsUserMgtCode.SubModuleProductSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmPriceMasterPS, "Product Price Master", "1.28.01.11", clsUserMgtCode.SubModuleProductSetup, 27)

        'transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleProductSaleTransaction, "Transaction", "1.28.02", clsUserMgtCode.ModuleProductSale, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmBookingProductSale, "Product Booking Entry", "1.28.02.01", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing)
        InsertDefaultValue(clsUserMgtCode.frmdispatchAdviceProductSale, "Production Advice", "1.28.02.02", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmDeliveryPrderProductSale, "Product Delivery Order", "1.28.02.03", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderProductSale, "Product Sale Order", "1.28.02.04", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmShipmentProductSale, "Product Dispatch", "1.28.02.05", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSaleInvoiceProductSale, "Product Invoice", "1.28.02.06", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmSaleReturnProductSale, "Product Sale Return", "1.28.02.07", clsUserMgtCode.SubModuleProductSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.FrmShortCloseDOPS, "Short Close DO PS", "1.28.02.08", clsUserMgtCode.SubModuleCSASaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmProductDispatchGateOut, "Product Dispatch Gate Out", "1.28.02.09", clsUserMgtCode.SubModuleProductSaleTransaction, 8)
        ''====Parteek Added 10-01-2017
        ' Commented by priti KDI/03/04/18-000179
        'InsertDefaultValue(clsUserMgtCode.frmProductBooking, "Product Booking", "1.28.02.10", clsUserMgtCode.SubModuleProductSaleTransaction, 8)
        ''========
        'InsertDefaultValue(clsUserMgtCode.frmMultCustBookingDisp, "Dispatch Booking Multiple Customer", "1.28.02.11", clsUserMgtCode.SubModuleProductSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmGateEntryReturnPS, "Gate Return Product Sale", "1.28.02.10", clsUserMgtCode.SubModuleProductSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryPS, "Sale Return Gate Entry Product Sale", "1.28.02.11", clsUserMgtCode.SubModuleProductSaleTransaction, 8)

        InsertDefaultValue(clsUserMgtCode.FrmGatePassPS, "Product GatePass Entry", "1.28.02.12", clsUserMgtCode.SubModuleProductSaleTransaction, 27)
        'InsertDefaultValue(clsUserMgtCode.frmRouteFreightDetails, "Route Freight Details", "1.28.02.08", clsUserMgtCode.SubModuleProductSaleTransaction, 27)

        'Report
        InsertDefaultValue(clsUserMgtCode.SubModuleProductSaleReport, "Report", "1.28.03", clsUserMgtCode.ModuleProductSale, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.FrmPrintProductInvoiceStatement, "Product Invoice Print Statement ", "1.28.03.01", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptProductSaleRegister1, "Product Sale Register ", "1.28.03.02", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptProductBookingStatus, "Product Sale Booking Status Report ", "1.28.03.03", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptProductDispatchStatus, "Product Sale Dispatch Status Report ", "1.28.03.04", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptProductDOStatus, "Product Sale Delivery Status Report ", "1.28.03.05", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptProductSaleOrderStatus, "Product Sale Order Status Report ", "1.28.03.06", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptsaleRegisterReport, "Sale Register Report ", "1.28.03.07", clsUserMgtCode.SubModuleProductSaleReport, 37)

        InsertDefaultValue(clsUserMgtCode.RptPartyWiseSale, "Party Wise Sale Report ", "1.28.03.08", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptGPDetail, "GP Details Report", "1.28.03.09", clsUserMgtCode.SubModuleProductSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailProductSale, "Product Sale Audit Trail Report", "1.28.03.10", clsUserMgtCode.SubModuleProductSaleReport, 37)

#End Region

#Region "Export Sale"
        ' 
        ''-------------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleExportSale, "Export Sale", "1.29", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleExportSetup, "Setup", "1.29.01", clsUserMgtCode.ModuleExportSale, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingExport, "Sales Setting", "1.29.01.01", clsUserMgtCode.SubModuleExportSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmEnquiryMaster, "Enquiry Master", "1.29.01.02", clsUserMgtCode.SubModuleExportSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmNotifiedPartyMaster, "Notified Party Master", "1.29.01.03", clsUserMgtCode.SubModuleExportSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCHAChargeMaster, "CHA Charge Master", "1.29.01.04", clsUserMgtCode.SubModuleExportSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmExIncentiveMaster, "Export Incentive Master", "1.29.01.05", clsUserMgtCode.SubModuleExportSetup, 27)

        'transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleExportSaleTransaction, "Transaction", "1.29.02", clsUserMgtCode.ModuleExportSale, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesQuotation, "Export Sale Quotation", "1.29.02.01", clsUserMgtCode.SubModuleExportSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesOrder, "Export Sales Order", "1.29.02.02", clsUserMgtCode.SubModuleExportSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmEXPorformaInvoice, "Export Proforma Invoice", "1.29.02.03", clsUserMgtCode.SubModuleExportSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmEXCommercialInvoice, "Export Commercial Invoice & Packing List", "1.29.02.04", clsUserMgtCode.SubModuleExportSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesInvoice, "Export Sales Invoice", "1.29.02.05", clsUserMgtCode.SubModuleExportSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesReturn, "Export Sale Return", "1.29.02.06", clsUserMgtCode.SubModuleExportSaleTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmsaleReturnGateEntryExportSAle, "Sale Return Gate Entry Export Sale", "1.29.02.07", clsUserMgtCode.SubModuleExportSaleTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesOrderR, "Export Sales Order Return", "1.29.02.08", clsUserMgtCode.SubModuleExportSaleTransaction, 8)

        'Report

        InsertDefaultValue(clsUserMgtCode.SubModuleExportSaleReport, "Report", "1.29.03", clsUserMgtCode.ModuleExportSale, "SM", 30)
        'InsertDefaultValue(clsUserMgtCode.RptExportSaleRegister, "Export Sale Register Report", "1.29.03.01", clsUserMgtCode.SubModuleExportSaleReport, 27)
        InsertDefaultValue(clsUserMgtCode.RptMTPIInHead, "PO In Hand Report ", "1.28.03.02", clsUserMgtCode.SubModuleExportSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptEXProductWiseDetail, "ProductWise Detail Report ", "1.28.03.03", clsUserMgtCode.SubModuleExportSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptdateOfArrivalofCons, "Date Of Arrival of Consignment", "1.28.03.04", clsUserMgtCode.SubModuleExportSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmEXSalesCancelledTransation, "Export Sale Cancelled Transactions Report", "1.28.03.05", clsUserMgtCode.SubModuleExportSaleReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailExportSale, "Export Sale Audit Trail Report", "1.28.03.06", clsUserMgtCode.SubModuleExportSaleReport, 37)
        'InsertDefaultValue(clsUserMgtCode.RptPaymentRelization, "Payment Relization", "1.28.03.05", clsUserMgtCode.SubModuleExportSaleReport, 27) '
        '---------------------------------------------
#End Region

#Region "Merchant Trade"
        InsertDefaultValue(clsUserMgtCode.ModuleMerchantTradeSale, "Merchant Trade", "1.30", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleMTSetup, "Setup", "1.30.01", clsUserMgtCode.ModuleMerchantTradeSale, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmSaleSettingMerchant, "Sales Setting", "1.30.01.01", clsUserMgtCode.SubModuleMTSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMerchantPaymentTerms, "Merchant Payment Terms", "1.30.01.02", clsUserMgtCode.SubModuleMTSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMerchantPaymentTermsGroup, "Merchant Payment Terms Group", "1.30.01.03", clsUserMgtCode.SubModuleMTSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMTReportContextFormat, "Merchant LC Report Context Declaration", "1.30.01.04", clsUserMgtCode.SubModuleMTSetup, 27)
        ''transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleMerchantTradeTransaction, "Transaction", "1.30.02", clsUserMgtCode.ModuleMerchantTradeSale, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseOrderMT, "Merchant Purchase Order", "1.30.02.01", clsUserMgtCode.SubModuleMerchantTradeTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmSalesOrderMT, "Merchant Sales Order", "1.30.02.02", clsUserMgtCode.SubModuleMerchantTradeTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmProformaInvoiceMT, "Merchant Proforma Invoice", "1.30.02.03", clsUserMgtCode.SubModuleMerchantTradeTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.FrmLCRequest, "LC Request", "1.30.02.04", clsUserMgtCode.SubModuleMerchantTradeTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmLCCreation, "LC Creation", "1.30.02.05", clsUserMgtCode.SubModuleMerchantTradeTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmDocumentAcceptance, "Document Acceptance", "1.30.02.06", clsUserMgtCode.SubModuleMerchantTradeTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.FrmFixedDeposit, "Fixed Deposit", "1.30.02.07", clsUserMgtCode.SubModuleMerchantTradeTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCommercialInvoiceMT, "Merchant Commercial Invoice & Packing List", "1.30.02.08", clsUserMgtCode.SubModuleMerchantTradeTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.FrmSRNMT, "Merchant SRN", "1.30.02.09", clsUserMgtCode.SubModuleMerchantTradeTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmSalesInvoiceMT, "Merchant Sales Invoice", "1.30.02.10", clsUserMgtCode.SubModuleMerchantTradeTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        InsertDefaultValue(clsUserMgtCode.frmSalesReturnMT, "Merchant Sale Return", "1.30.02.11", clsUserMgtCode.SubModuleMerchantTradeTransaction, "", 8, 0, "", "", "", True, True, Nothing)
        'Report
        InsertDefaultValue(clsUserMgtCode.SubModuleMerchantReport, "Report", "1.30.03", clsUserMgtCode.ModuleMerchantTradeSale, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.RptCashAgainstDocs, "Cash Against Docs. ", "1.30.03.01", clsUserMgtCode.SubModuleMerchantReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmMTSalesCancelledTransation, "Merchant Sale Cancelled Transactions Report", "1.30.03.02", clsUserMgtCode.SubModuleMerchantReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailMerchantTrade, "Merchant Trade Audit Trail Report", "1.30.03.03", clsUserMgtCode.SubModuleMerchantReport, 37)


#End Region

#Region "Product Dairy Sales"
        '' Dairy Product Sale
        InsertDefaultValue(clsUserMgtCode.ModuleProductSaleDairy, "Product Dairy Sales", "1.31", clsUserMgtCode.ExpertERP, "M", 23)

        '' Dairy Product Sale ==> Setup
        InsertDefaultValue(clsUserMgtCode.SubModuleProductSaleDairySetUp, "Setup", "1.31.01", clsUserMgtCode.ModuleProductSaleDairy, "SM", 24)

        ''Dairy Product Sale ==> Transcation
        InsertDefaultValue(clsUserMgtCode.SubModuleProductSaleDairyTransaction, "Transcation", "1.31.02", clsUserMgtCode.ModuleProductSaleDairy, "SM", 26)

        '' Dairy Product Sale ==> Report
        InsertDefaultValue(clsUserMgtCode.SubModuleProductSaleDairyReport, "Report", "1.31.03", clsUserMgtCode.ModuleProductSaleDairy, "SM", 30)

        '' End of Dairy Product Sale
#End Region

#Region "APS Dairy Sales"
        '' Dairy APS Sale
        InsertDefaultValue(clsUserMgtCode.ModuleAPSSaleDairy, "APS Dairy Sales", "1.32", clsUserMgtCode.ExpertERP, "M", 23)

        '' Dairy APS Sale ==> Setup
        InsertDefaultValue(clsUserMgtCode.SubModuleAPSSaleDairySetUp, "Setup", "1.32.01", clsUserMgtCode.ModuleAPSSaleDairy, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmCustomerTender, "Customer Tender", "1.32.01.01", clsUserMgtCode.SubModuleAPSSaleDairySetUp, 27)
        ''Dairy APS Sale ==> Transcation
        InsertDefaultValue(clsUserMgtCode.SubModuleAPSSaleDairyTransaction, "Transcation", "1.32.02", clsUserMgtCode.ModuleAPSSaleDairy, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmSalesOrder, "Sales Order", "1.32.02.01", clsUserMgtCode.SubModuleAPSSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmSalesOrderDispatch, "Sales Order Dispatch", "1.32.02.02", clsUserMgtCode.SubModuleAPSSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmSalesOrderGatePass, "Sales Order GatePass", "1.32.02.03", clsUserMgtCode.SubModuleAPSSaleDairyTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)

        '' Dairy APS Sale ==> Report
        InsertDefaultValue(clsUserMgtCode.SubModuleAPSSaleDairyReport, "Report", "1.32.03", clsUserMgtCode.ModuleAPSSaleDairy, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.FrmAPSSaleReport, "APS Sale Report", "1.32.03.01", clsUserMgtCode.SubModuleAPSSaleDairyReport, 37)
        '' End of Dairy Product Sale
#End Region

#Region "Sale Purchase Security"
        '-----------------------Sale Purchase security----------------
        InsertDefaultValue(clsUserMgtCode.ModulesalePurchaseSecurity, "Sale Purchase Security", "1.33", clsUserMgtCode.ExpertERP, "M", 23)

        InsertDefaultValue(clsUserMgtCode.SubSalePurchaseSecuritySetup, "Setup", "1.33.01", clsUserMgtCode.ModulesalePurchaseSecurity, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.FrmBankPermission, "Bank Permission", "1.33.01.01", clsUserMgtCode.SubSalePurchaseSecuritySetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCustomerPermission, "Customer Permission", "1.33.01.02", clsUserMgtCode.SubSalePurchaseSecuritySetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmVendorPermission, "Vendor Permission", "1.33.01.03", clsUserMgtCode.SubSalePurchaseSecuritySetup, 27)
        InsertDefaultValue(clsUserMgtCode.SubSalePurchaseSecurityReport, "Report", "1.33.02", clsUserMgtCode.ModulesalePurchaseSecurity, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.frmCustomerPermissionReport, "Customer Permission Report", "1.33.02.01", clsUserMgtCode.SubSalePurchaseSecurityReport, 37)
        '---------------------End Sale purchase security----------------
#End Region

#Region "Visual Process Flow"
        '=================================================Visual Process Flow================================================================
        InsertDefaultValue(clsUserMgtCode.ModuleVPFlow, "Visual Process Flow", "1.34", clsUserMgtCode.ExpertERP, "M", 23)

        ' Module Wise Cycles 
        InsertDefaultValue(clsUserMgtCode.FrmVPFSettings, "VPF Settings", "1.34.01.01", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmVPFActiveReport, "Active Section", "1.34.01.02", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmGLCycle, "General Ledger", "1.34.01.03", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmCommonCycle, "Common", "1.34.01.04", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmAPCycle, "Vendor Management", "1.34.01.05", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmARCycle, "Customer Management", "1.34.01.06", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBulkSaleCycle, "Bulk Sale", "1.34.01.07", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmFreshSaleCycle, "Fresh Sale", "1.34.01.08", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmProductSaleCycle, "Product Sale", "1.34.01.09", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmCSASaleCycle, "CSA Sale", "1.34.01.10", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMMCycle, "Material Management", "1.34.01.11", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmMCCProcurementCycle, "MCC Procurement", "1.34.01.12", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBulkProcurementCycle, "Bulk Procurement", "1.34.01.13", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmPurchaseCycle, "Purchase", "1.34.01.14", clsUserMgtCode.ModuleVPFlow, 27)
        InsertDefaultValue(clsUserMgtCode.FrmDProductionCycle, "Production", "1.34.01.15", clsUserMgtCode.ModuleVPFlow, 27)

        '' -----------------------------------------------------------------------------------------------------------------------------------

#End Region


#Region "Utility"
        '--------------Utility----------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleUtility, "Utility", "1.35", clsUserMgtCode.ExpertERP, "M", 23)

        InsertDefaultValue(clsUserMgtCode.SubModuleUtilitySetup, "Setup", "1.35.01", clsUserMgtCode.ModuleUtility, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.Calculator, "Calculator", "1.35.01.01", clsUserMgtCode.SubModuleUtilitySetup, 27)
        'InsertDefaultValue(clsUserMgtCode.mbtnTakeBackup, "Backup", "1.20.01.02", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnCreateReceiptAgainstInvoice, "Create Receipt Against Invoice", "1.35.01.03", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnCreateReceiptAgainstSale, "Create Empities Against Sales Invoice", "1.35.01.04", clsUserMgtCode.SubModuleUtilitySetup, 27)

        InsertDefaultValue(clsUserMgtCode.FrmDayWiseLoadOutEntered, "Data Entry Status", "1.35.01.06", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmUserPerformanceDetail, "User Performance Detail", "1.35.01.07", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmStockReport, "Stock Report", "1.35.01.08", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmUtilityForm, "Utility Form", "1.35.01.09", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmReconciliationSetting, "Reconciliation Setting", "1.35.01.10", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.mbtnRestoreDB, "Restore DataBase", "1.35.01.11", clsUserMgtCode.SubModuleUtilitySetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmCancelAfterPosting, "Cancel After Posting", "1.35.01.12", clsUserMgtCode.SubModuleUtilitySetup, 27)
        '----------------------------------Utility End---------------------------------------'

#End Region

#Region "Milk Job Work"
        '=============Milk Job Work===================================================
        InsertDefaultValue(clsUserMgtCode.MilkJobWork, "Milk Job Work", "1.36", clsUserMgtCode.ExpertERP, "M", 23)
        ''set up
        InsertDefaultValue(clsUserMgtCode.SubModuleMJWTransaction, "Transaction", "1.36.01", clsUserMgtCode.MilkJobWork, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmMilkJobWork, "Milk RGP", "1.36.01.01", clsUserMgtCode.SubModuleMJWTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmMilkGateEntry, "Milk Gate Entry", "1.36.01.02", clsUserMgtCode.SubModuleMJWTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmMilkWeighment, "Milk Weighment", "1.36.01.03", clsUserMgtCode.SubModuleMJWTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmJobMilkQualityCheck, "Milk Quality Check", "1.36.01.04", clsUserMgtCode.SubModuleMJWTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.FrmMilkUnloading, "Milk Unloading", "1.36.01.05", clsUserMgtCode.SubModuleMJWTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'InsertDefaultValue(clsUserMgtCode.FrmMilkCleaning, "Milk Cleaning", "1.36.01.06", clsUserMgtCode.SubModuleMJWTransaction, 27)
        InsertDefaultValue(clsUserMgtCode.FrmJobMilkSRN, "Milk Jobwork SRN", "1.36.01.06", clsUserMgtCode.SubModuleMJWTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        '=================report
        InsertDefaultValue(clsUserMgtCode.SubModuleMJWReport, "Report", "1.36.02", clsUserMgtCode.MilkJobWork, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.RptJobWorkStatus, "Job Work Status", "1.36.02.01", clsUserMgtCode.SubModuleMJWReport, 37) ' 
        InsertDefaultValue(clsUserMgtCode.AuditTrailMilkJobWork, "Milk Jobwork Audit Trail Report", "1.36.02.02", clsUserMgtCode.SubModuleMJWReport, 37)

        ''===========================end Job Work===========================================================
#End Region

#Region "TDS Payroll"
        '============================TDS Module=================Added by Preeti Gupta============================================================
        InsertDefaultValue(clsUserMgtCode.ModuleTDSPayroll, "TDS Payroll", "1.37", clsUserMgtCode.ExpertERP, "M", 23)
        '==========================================================TDS Module SetUp
        InsertDefaultValue(clsUserMgtCode.SubModuleTDSPayrollSetup, "Setup", "1.37.01", clsUserMgtCode.ModuleTDSPayroll, "SM", 24)
        'InsertDefaultValue(clsUserMgtCode.frmIncomeTaxSlab, "Income Tax Slab", "1.37.01.01", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmITSection, "IT Section Old", "1.37.01.02", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmInvestmentType, "Investment Type Old", "1.37.01.03", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmHouseRentDeclaration, "House Rent Declaration", "1.37.01.04", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmInvestmentDeclaration, "Investment Declaration Old", "1.37.01.05", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmHRAExemptionRule, "HRA Exemption Rule", "1.37.01.06", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)

        ' By Prabhakar
        InsertDefaultValue(clsUserMgtCode.frmSectionAllowanceMaster, "IT Section/Allowance Master", "1.37.01.01", clsUserMgtCode.SubModuleTDSPayrollSetup, 27) ' Section/Allowance Master
        InsertDefaultValue(clsUserMgtCode.frmSavingsMaster, "Investment Master", "1.37.01.02", clsUserMgtCode.SubModuleTDSPayrollSetup, 27) ' Savings Master
        InsertDefaultValue(clsUserMgtCode.frmEmployeeSavingsMapping, "Investment Declaration Master", "1.37.01.03", clsUserMgtCode.SubModuleTDSPayrollSetup, 27) ' Employee Savings Mapping
        InsertDefaultValue(clsUserMgtCode.frmIncomeTaxSlab, "Income Tax Slab", "1.37.01.04", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmHouseRentDeclaration, "House Rent Declaration", "1.37.01.05", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmHRAExemptionRule, "HRA Exemption Rule", "1.37.01.06", clsUserMgtCode.SubModuleTDSPayrollSetup, 27)
        '==============================================================END TDS Module Transaction=====================================================
        InsertDefaultValue(clsUserMgtCode.SubModuleTDSPayrollTrans, "Transaction", "1.37.02", clsUserMgtCode.ModuleTDSPayroll, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.IncomeTaxTDSCalculation, "Payroll TDS Calcuation", "1.33.02.01", clsUserMgtCode.SubModuleTDSPayrollTrans, 8) ''BHA/17/01/19-000783 by balwinder on 18/01/2019

        InsertDefaultValue(clsUserMgtCode.SubModuleTDSPayrollReport, "Report", "1.37.03", clsUserMgtCode.ModuleTDSPayroll, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.frmEmployeeTDSRpt, "Payroll TDS Register", "1.37.03.01", clsUserMgtCode.SubModuleTDSPayrollReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailTDSPayroll, "TDS Payroll Audit Trail Report", "1.37.03.02", clsUserMgtCode.SubModuleTDSPayrollReport, 37)
        '================================================================END TDS Module============================================================


#End Region

#Region "Silage Production"
        '= Added By Kunal :Silage Production Module
        InsertDefaultValue(clsUserMgtCode.ModuleSilageProduction, "Silage Production", "1.38", clsUserMgtCode.ExpertERP, "M", 23)
        '==========================================================TDS Module SetUp
        InsertDefaultValue(clsUserMgtCode.SubModuleSilagePrdSetup, "Setup", "1.38.01", clsUserMgtCode.ModuleSilageProduction, "SM", 24)
        '----------------------Screens
        InsertDefaultValue(clsUserMgtCode.frmSilageCriteriaMaster, "Silage Criteria Master", "1.38.01.01", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageAreaMaster, "Silage Area Master", "1.38.01.02", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageProductionApplication, "Silage Application Form ", "1.38.01.03", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageEnterPrenur, "Silage Entrepreneur Selection", "1.38.01.04", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageFormerProduction, "Silage Farmer Selection ", "1.38.01.05", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageTankerTransporterMaster, "Tanker Transporter Master", "1.38.01.06", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageTankerMaster, "Tanker Master", "1.38.01.07", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageParameterMaster, "Parameter Master", "1.38.01.08", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSialgeParameterRangeMaster, "Parameter Range Master", "1.38.01.09", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSialgeParameterRangeMasterForQC, "Parameter Range Master For QC", "1.38.01.10", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSialgeParameterValueMaster, "Parameter Value Master", "1.38.01.11", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilagePriceChartBulkProc, "Price Chart Master (Bulk Procurement)", "1.38.01.12", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageVendorPriceChartMapping, "Vendor Price Chart Mapping", "1.38.01.13", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmSilageVendorPriceChartMapping, "Contractor Tanker", "1.38.01.14", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.FrmSialgeSupplierMaster, "Supplier Master", "1.38.01.15", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageDivertedContractor, "Diverted Contractor", "1.38.01.16", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageTypeMast, "Crop Type Master", "1.38.01.17", clsUserMgtCode.SubModuleSilagePrdSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageGradeMaster, "Crop Grade Master", "1.38.01.18", clsUserMgtCode.SubModuleSilagePrdSetup, 27)

        '' """"""""""""""""""""""""
        '=======transaction =====================================================================================================================================================
        InsertDefaultValue(clsUserMgtCode.SubModuleSilagePrdTrans, "Transaction", "1.38.02", clsUserMgtCode.ModuleSilageProduction, "SM", 26)

        '' Sigage transaction                
        'InsertDefaultValue(clsUserMgtCode.frmSialgeIntimation, "Intimation", "1.38.02.01", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageGateEntry, "Gate Entry", "1.38.02.02", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageWeighment, "Weighment", "1.38.02.03", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageQualityCheck, "Quality Check", "1.38.02.04", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageUnloading, "Unloading", "1.38.02.05", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSialgeCleaning, "Cleaning", "1.38.02.06", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageGateOut, "Gate Out", "1.38.02.07", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageBulkSRN, "SRN", "1.38.02.08", clsUserMgtCode.SubModuleSilagePrdTrans, 27)
        InsertDefaultValue(clsUserMgtCode.frmSilageBulkSRNReturn, "SRN Return", "1.38.02.09", clsUserMgtCode.SubModuleSilagePrdTrans, 27)

        ' Module Paravet Services


#End Region


#Region "Paravet Services"
        InsertDefaultValue(clsUserMgtCode.ModuleParavet, "Paravet Services", "1.39", clsUserMgtCode.ExpertERP, "M", 23)
        ' Submodule Master
        InsertDefaultValue(clsUserMgtCode.SubModuleParavetMaster, "Setup", "1.39.01", clsUserMgtCode.ModuleParavet, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmCattleType, "Cattle Type Master", "1.39.01.02", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.frmBredType, "Bred Type Master", "1.39.01.03", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.frmCattleColor, "Cattle Color Master", "1.39.01.04", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.frmCattleMaster, "Cattle Master", "1.39.01.05", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.FrmServiceGroup, "Service Group Master", "1.39.01.06", clsUserMgtCode.SubModuleParavetMaster, 27)
        'InsertDefaultValue(clsUserMgtCode.FrmServiceName, "Service Name Master", "1.39.01.07", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.FrmServiceMaster, "Service Master", "1.39.01.08", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.FrmNDDBMaster, "NDDB Master", "1.39.01.09", clsUserMgtCode.SubModuleParavetMaster, 27)
        InsertDefaultValue(clsUserMgtCode.FrmBullMaster, "Bull Master", "1.39.01.10", clsUserMgtCode.SubModuleParavetMaster, 27)
        ' Submodule Transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleParavetTrans, "Transaction", "1.39.02", clsUserMgtCode.ModuleParavet, "SM", 26)

        InsertDefaultValue(clsUserMgtCode.FrmFarmerServiceOrderWithRate, "Farmer Service Order", "1.39.02.02", clsUserMgtCode.SubModuleParavetTrans, 27)
        ' Submodule Report
        InsertDefaultValue(clsUserMgtCode.SubModuleParavetReport, "Reports", "1.39.03", clsUserMgtCode.ModuleParavet, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptSMSDetails, "SMS Details Report", "1.39.03.02", clsUserMgtCode.SubModuleParavetReport, 37)

#End Region

#Region "Farmer Payment"
        ' Module Farmer Payment
        InsertDefaultValue(clsUserMgtCode.ModuleFarmerPayment, "Farmer Payment", "1.40", clsUserMgtCode.ExpertERP, "M", 23)
        '' ad setup, transaction and Reports in Farmer payment
        InsertDefaultValue(clsUserMgtCode.SubModuleFarmerPaymentMaster, "Setup", "1.40.01", clsUserMgtCode.ModuleFarmerPayment, "SM", 24)
        '' add setup
        InsertDefaultValue(clsUserMgtCode.frmMCCFarmerMappingFP, "MCC Farmer Mapping", "1.40.01.01", clsUserMgtCode.SubModuleFarmerPaymentMaster, 27)
        InsertDefaultValue(clsUserMgtCode.Farmeraccountset, "Farmer Account Set", "1.40.01.02", clsUserMgtCode.SubModuleFarmerPaymentMaster, 27)

        InsertDefaultValue(clsUserMgtCode.SubModuleFarmerPaymentTransaction, "Transaction", "1.40.02", clsUserMgtCode.ModuleFarmerPayment, "SM", 26)
        '' add transactions
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialFarmer, "MCC Material Sale Farmer", "1.40.02.01", clsUserMgtCode.SubModuleFarmerPaymentTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialSaleReturnFarmer, "MCC Material Sale Return Farmer ", "1.40.02.02", clsUserMgtCode.SubModuleFarmerPaymentTransaction, "", 8, 0, "", "", "", True, True, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.MPBillGeneration, "Generate MP Bill", "1.40.02.03", clsUserMgtCode.SubModuleFarmerPaymentTransaction, 8)

        InsertDefaultValue(clsUserMgtCode.frmPaymentProcessFarmer, "Payment Process Farmer", "1.40.02.04", clsUserMgtCode.SubModuleFarmerPaymentTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmFarmerPaymentAdjustment, "Farmer Payment Adjustment", "1.40.02.05", clsUserMgtCode.SubModuleFarmerPaymentTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmLockMPCollectionPC, "Lock MP Collection", "1.40.02.06", clsUserMgtCode.SubModuleFarmerPaymentTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmFarmerPaymentEntry, "Farmer Payment Entry", "1.40.02.07", clsUserMgtCode.SubModuleFarmerPaymentTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmFarmerMilkPurchaseInvoice, "Farmer Milk Purchase Invoice", "1.40.02.08", clsUserMgtCode.SubModuleFarmerPaymentTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmMCCMaterialFarmerUploader, "MCC Material Sale Farmer Uploader", "1.40.02.09", clsUserMgtCode.SubModuleFarmerPaymentTransaction, "", 8, 0, "", "", "", True, True, Nothing)

        InsertDefaultValue(clsUserMgtCode.SubModuleFarmerPaymentReport, "Report", "1.40.03", clsUserMgtCode.ModuleFarmerPayment, "SM", 30)

        ' Module Farmer Payment Report
        InsertDefaultValue(clsUserMgtCode.frmFarmerLedgerReport, "Farmer Ledger Report", "1.40.03.01", clsUserMgtCode.SubModuleFarmerPaymentReport, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailFarmerPayment, "Farmer Payment Audit Trail Report", "1.40.03.02", clsUserMgtCode.SubModuleFarmerPaymentReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptFarmerAbstractReport, "Abstract Report", "1.40.03.03", clsUserMgtCode.SubModuleFarmerPaymentReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVspFarmerPaymentDetail, "VSP Farmer Payment Report", "1.40.03.04", clsUserMgtCode.SubModuleFarmerPaymentReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptFarmerNotMappedWithVSP, "Farmer Not Mapped With VSP", "1.40.03.05", clsUserMgtCode.SubModuleFarmerPaymentReport, 37)

#End Region

#Region "Job Work Outward"
        '' Module JobWork OutWard
        InsertDefaultValue(clsUserMgtCode.ModuleJobWorkOutWard, "Job Work Outward", "1.41", clsUserMgtCode.ExpertERP, "M", 23)
        '' ad setup, transaction and Reports in Job Work OutWord
        InsertDefaultValue(clsUserMgtCode.SubModuleSetupJobWorkOurWard, "Setup", "1.41.01", clsUserMgtCode.ModuleJobWorkOutWard, "SM", 24)

        InsertDefaultValue(clsUserMgtCode.frmJobWorkoutwordMaster, "Job Price Master", "1.41.01.01", clsUserMgtCode.SubModuleSetupJobWorkOurWard, 27)
        InsertDefaultValue(clsUserMgtCode.frmVendorItemChargeMaster, "Vendor Item Job work Charges", "1.41.01.02", clsUserMgtCode.SubModuleSetupJobWorkOurWard, 27)

        InsertDefaultValue(clsUserMgtCode.JWOParameterMaster, "JWO Parameter Master", "1.41.01.03", clsUserMgtCode.SubModuleSetupJobWorkOurWard, 27)
        InsertDefaultValue(clsUserMgtCode.JWOParameterFormula, "JWO Parameter Formula", "1.41.01.04", clsUserMgtCode.SubModuleSetupJobWorkOurWard, 27)
        InsertDefaultValue(clsUserMgtCode.JWOFormulaVendorMapping, "JWO Formul Vendor Mapping", "1.41.01.05", clsUserMgtCode.SubModuleSetupJobWorkOurWard, 27)

        InsertDefaultValue(clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "Transaction", "1.41.02", clsUserMgtCode.ModuleJobWorkOutWard, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransfer, "Job Work Transfer - Milk", "1.41.02.01", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransferReturn, "Job Work Transfer Return - Milk", "1.41.02.02", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransferOther, "Job Work Transfer - Other", "1.41.02.03", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmMilkJobWorkTransferOtherReturn, "Job Work Transfer Return - Other", "1.41.02.04", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmGateEntry_JWO, "Job Work Gate Entry", "1.41.02.05", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, 8)
        InsertDefaultValue(clsUserMgtCode.frmWeighment_JWO, "Job Work Weighment", "1.41.02.06", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.JWO_Item_Acceptance, "Job Work Item Accepatance", "1.41.02.07", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, 8)
        InsertDefaultValue(clsUserMgtCode.frmQC_JWO, "Job Work Quality Check", "1.41.02.08", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, 8)
        InsertDefaultValue(clsUserMgtCode.frmUnloading_JWO, "Job Work Unloading", "1.41.02.09", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.JWO_Estimation, "JWO Estimation", "1.41.02.10", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, 8)
        InsertDefaultValue(clsUserMgtCode.JWO_SRN, "Job Work SRN", "1.41.02.11", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.JWO_SRN_Return, "Job Work SRN Return", "1.41.02.12", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkConsumption, "Job Work Consumption", "1.41.02.13", clsUserMgtCode.SubModuleTransactionJobWorkOurWard, 8)

        '=========Job Work Outward Report==========
        InsertDefaultValue(clsUserMgtCode.SubModuleReportJobWorkOurWard, "Reports", "1.41.03", clsUserMgtCode.ModuleJobWorkOutWard, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.RptJobWorkRegister, "Job Work Register Report", "1.41.03.01", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.RptJobWorOutwardPurchasekRegister, "Job Work Outward Purchase Register", "1.41.03.02", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.frmJobworkTransfer, "Job Work Transfer Report", "1.41.03.03", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.frmJobworkChargesReport, "Job Work Charges Report", "1.41.03.04", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.frmJobworkSRNReceiptReport, "Job Work SRN Receipt Report", "1.41.03.05", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.stockRecoNewJR, "Stock Reco Jobwork", "1.41.03.06", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailJobWorkOutward, "Jobwork Outward Audit Trail Report", "1.41.03.07", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)
        InsertDefaultValue(clsUserMgtCode.frmJobworkItemInStatusReport, "Job Work Item In Status Report", "1.41.03.08", clsUserMgtCode.SubModuleReportJobWorkOurWard, 37)

#End Region

#Region "Reconciliation"
        '=====================Reconciliation==================================
        InsertDefaultValue(clsUserMgtCode.ModuleReco, "Reconciliation", "1.42", clsUserMgtCode.ExpertERP, "M", 23)
        InsertDefaultValue(clsUserMgtCode.SubModuleRecoReport, "Reports", "1.42.03", clsUserMgtCode.ModuleReco, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptSaleReco, "Sale Reco", "1.42.03.01", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptPurReco, "Purchase Reco", "1.42.03.02", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVendorReco, "Vendor Reco", "1.42.03.03", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptCustomerReco, "Customer Reco", "1.42.03.04", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptInvReco, "Inventory Reco Report", "1.42.03.05", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmReconcilationRpt, "Bank GL Reco Report", "1.42.03.06", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.frmBankBookRecoReport, "Bank Book Reco Report", "1.42.03.07", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptShippmentClearing, "Shipment Clearing", "1.42.03.08", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptPayableClearing, "Payable Clearing", "1.42.03.09", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptVSPCustomerReco, "VSP Customer Reco", "1.42.03.10", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSRNReco, "Milk SRN Reco", "1.42.03.11", clsUserMgtCode.SubModuleRecoReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptTankerDispatchGainLossReco, "Tanker Dispatch Gain-Loss Reco", "1.42.03.12", clsUserMgtCode.SubModuleRecoReport, 37)
        '=====================Job Work Inward==================================

#End Region


#Region "Job Work Inward"
        InsertDefaultValue(clsUserMgtCode.ModuleJobWorkInWard, "Job Work Inward", "1.43", clsUserMgtCode.ExpertERP, "M", 23)
        'Startup
        InsertDefaultValue(clsUserMgtCode.SubModuleSetupJobWorkInWard, "Setup", "1.43.01", clsUserMgtCode.ModuleJobWorkInWard, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmJWParameterMaster, "JW Parameter Master", "1.43.01.01", clsUserMgtCode.SubModuleSetupJobWorkInWard, 27)
        InsertDefaultValue(clsUserMgtCode.frmJWFormulaMaster, "JW Formula", "1.43.01.02", clsUserMgtCode.SubModuleSetupJobWorkInWard, 27)
        InsertDefaultValue(clsUserMgtCode.frmJWVendorFormula, "Vendor Formula Mapping", "1.43.01.03", clsUserMgtCode.SubModuleSetupJobWorkInWard, 27)
        InsertDefaultValue(clsUserMgtCode.JWIItemPriceMaster, "JWI Item Price Master", "1.43.01.04", clsUserMgtCode.SubModuleSetupJobWorkInWard, 27)
        'Transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleTransactionJobWorkInWard, "Transaction", "1.43.02", clsUserMgtCode.ModuleJobWorkInWard, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.FrmSRNJobWorkEstimate, "JW Estimate", "1.43.02.01", clsUserMgtCode.SubModuleTransactionJobWorkInWard, 8)
        InsertDefaultValue(clsUserMgtCode.frmJobWorkBillig, "JobWork Billing", "1.43.02.02", clsUserMgtCode.SubModuleTransactionJobWorkInWard, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        'Report
        InsertDefaultValue(clsUserMgtCode.SubModuleReportJobWorkInWard, "Reports", "1.43.03", clsUserMgtCode.ModuleJobWorkInWard, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptJobworkMilkReceipt, "Jobwork Milk Receipt Report", "1.43.03.01", clsUserMgtCode.SubModuleReportJobWorkInWard, 37)
        InsertDefaultValue(clsUserMgtCode.rptJobworkBilling, "Jobwork Billing Report", "1.43.03.02", clsUserMgtCode.SubModuleReportJobWorkInWard, 37)
        InsertDefaultValue(clsUserMgtCode.rptJobworkProductionReport, "Jobwork Milk Production Report", "1.43.03.03", clsUserMgtCode.SubModuleReportJobWorkInWard, 37)
        InsertDefaultValue(clsUserMgtCode.rptJWTankerStatusReport, "Jobwork Tanker Status Report", "1.43.03.04", clsUserMgtCode.SubModuleReportJobWorkInWard, 37)
        InsertDefaultValue(clsUserMgtCode.rptJWDispatchReport, "Jobwork Dispatch Report", "1.43.03.05", clsUserMgtCode.SubModuleReportJobWorkInWard, 37) '
        InsertDefaultValue(clsUserMgtCode.AuditTrailJobWorkInward, "Jobwork Inward Audit Trail Report", "1.43.03.06", clsUserMgtCode.SubModuleReportJobWorkInWard, 37)

#End Region

#Region "Electrical"
        ' Electrical
        InsertDefaultValue(clsUserMgtCode.ModuleElectrical, "Electrical", "1.44", clsUserMgtCode.ExpertERP, "M", 23)
        'Startup
        InsertDefaultValue(clsUserMgtCode.SubModuleSetupElectrical, "Setup", "1.44.01", clsUserMgtCode.ModuleElectrical, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmSlotMaster, "Slot Master", "1.44.01.01", clsUserMgtCode.SubModuleSetupElectrical, 27)
        InsertDefaultValue(clsUserMgtCode.frmDGMaster, "DG Master", "1.44.01.02", clsUserMgtCode.SubModuleSetupElectrical, 27)
        'Transaction
        InsertDefaultValue(clsUserMgtCode.SubModuleTransactionElectrical, "Transaction", "1.44.02", clsUserMgtCode.ModuleElectrical, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmDailyElectricalEntry, "Daily Electrical Entry", "1.44.02.01", clsUserMgtCode.SubModuleTransactionElectrical, 8)
        'Report
        InsertDefaultValue(clsUserMgtCode.SubModuleReportElectrical, "Reports", "1.44.03", clsUserMgtCode.ModuleElectrical, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptDailyElectricalEntryReport, "Daily Electrical Entry Report", "1.44.03.01", clsUserMgtCode.SubModuleReportElectrical, 37)
        InsertDefaultValue(clsUserMgtCode.AuditTrailElectrical, "Electrical Audit Trail Report", "1.44.03.02", clsUserMgtCode.SubModuleReportElectrical, 37)

#End Region

#Region "Account Set Report"
        'sanjay Ticket No-TEC/06/06/19-000522
        InsertDefaultValue(clsUserMgtCode.ModuleAccSet, "Account Set Report", "1.45", clsUserMgtCode.ExpertERP, "M", 23)
        InsertDefaultValue(clsUserMgtCode.SubModuleAccSetReport, "Reports", "1.45.03", clsUserMgtCode.ModuleAccSet, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.RptCustomerAccountSet, "Customer Account Set Report", "1.45.03.01", clsUserMgtCode.SubModuleAccSetReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptVendorAccountSetReport, "Vendor Account Set Report", "1.45.03.02", clsUserMgtCode.SubModuleAccSetReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptItemPurchaseSet, "Item Purchase Accounts Report", "1.45.03.03", clsUserMgtCode.SubModuleAccSetReport, 37)
        InsertDefaultValue(clsUserMgtCode.rptSaleAccountSetList, "Sale Account Set List", "1.45.03.04", clsUserMgtCode.SubModuleAccSetReport, 37)
        InsertDefaultValue(clsUserMgtCode.RptItemSalePurchaseSet, "Item Sale Purchase Account Report", "1.45.03.05", clsUserMgtCode.SubModuleAccSetReport, 37)
        InsertDefaultValue(clsUserMgtCode.FrmItemListRpt, "Item List Report", "1.45.03.06", clsUserMgtCode.SubModuleAccSetReport, 37)

#End Region

#Region "Engineering And Plant Management"
        'Sanjay ,ADD new module XpertERPEngineeringAndPlantManagement,Ticket no-VIJ/07/11/19-000056
        InsertDefaultValue(clsUserMgtCode.ModuleEngPlantMgt, "Engineering And Plant Management", "1.46", clsUserMgtCode.ExpertERP, "M", 23)
        'Startup
        InsertDefaultValue(clsUserMgtCode.SubModuleSetupEngPlantMgt, "Setup", "1.46.01", clsUserMgtCode.ModuleEngPlantMgt, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmSectionMasterEng, "Section Master Eng", "1.46.01.01", clsUserMgtCode.SubModuleSetupEngPlantMgt, 27)
        InsertDefaultValue(clsUserMgtCode.frmConsumptionTypeMaster, "Consumption Type Master", "1.46.01.02", clsUserMgtCode.SubModuleSetupEngPlantMgt, 27)
        InsertDefaultValue(clsUserMgtCode.frmSectionConsumptionMapping, "Section Consumption Mapping", "1.46.01.03", clsUserMgtCode.SubModuleSetupEngPlantMgt, 27)
        InsertDefaultValue(clsUserMgtCode.frmParameterMasterEng, "Parameter Master Eng", "1.46.01.04", clsUserMgtCode.SubModuleSetupEngPlantMgt, 27)
        InsertDefaultValue(clsUserMgtCode.frmSectionParameterMapping, "Section Parameter Mapping", "1.46.01.05", clsUserMgtCode.SubModuleSetupEngPlantMgt, 27)

        InsertDefaultValue(clsUserMgtCode.SubModuleTransactionEngPlantMgt, "Transaction", "1.46.02", clsUserMgtCode.ModuleEngPlantMgt, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmLogSheetEng, "Log Sheet Eng", "1.46.02.01", clsUserMgtCode.SubModuleTransactionEngPlantMgt, 8)
        InsertDefaultValue(clsUserMgtCode.WorkRequisitionEng, "Work Requisition", "1.46.02.02", clsUserMgtCode.SubModuleTransactionEngPlantMgt, 8)
        InsertDefaultValue(clsUserMgtCode.frmWorkEstimationEng, "Work Estimation", "1.46.02.03", clsUserMgtCode.SubModuleTransactionEngPlantMgt, 8)
        InsertDefaultValue(clsUserMgtCode.WorkOrderEng, "Work Order", "1.46.02.04", clsUserMgtCode.SubModuleTransactionEngPlantMgt, 8)
        InsertDefaultValue(clsUserMgtCode.frmWorkOrderStatusEng, "Work Order Status", "1.46.02.05", clsUserMgtCode.SubModuleTransactionEngPlantMgt, 8)
        InsertDefaultValue(clsUserMgtCode.SubModuleReportEngPlantMgt, "Reports", "1.46.03", clsUserMgtCode.ModuleEngPlantMgt, "SM", 30)
        'InsertDefaultValue(clsUserMgtCode.rptDailyElectricalEntryReport, "Daily Electrical Entry Report", "1.42.03.01", clsUserMgtCode.SubModuleReportElectrical, 37)

#End Region

#Region "Complaint"
        'Complaint Module By Prabhakar
        InsertDefaultValue(clsUserMgtCode.ModuleComplaint, "Complaint", "1.47", clsUserMgtCode.ExpertERP, "M", 23)
        InsertDefaultValue(clsUserMgtCode.SubModuleSetupComplaint, "Setup", "1.47.01", clsUserMgtCode.ModuleComplaint, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmCustomerComplaintMaster, "Customer Complaint Master", "1.47.01.01", clsUserMgtCode.SubModuleSetupComplaint, 27)

        InsertDefaultValue(clsUserMgtCode.SubModuleTransactionComplaint, "Transaction", "1.47.02", clsUserMgtCode.ModuleComplaint, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmCustomerComplaint, "Customer Complaint", "1.47.02.01", clsUserMgtCode.SubModuleTransactionComplaint, 8)
        InsertDefaultValue(clsUserMgtCode.frmLeakageReplacementUploader, "Leakage Replacement Uploader", "1.47.02.02", clsUserMgtCode.SubModuleTransactionComplaint, 8)


#End Region

#Region "Share"
        ''--------------------------------ShareModule---------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.ModuleShare, "Share", "1.48", clsUserMgtCode.ExpertERP, "M", 23)

        ''----------------SETUP--------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleShareSetup, "Setup", "1.48.01", clsUserMgtCode.ModuleShare, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.ShareMaster, "Share Master", "1.48.01.01", clsUserMgtCode.SubModuleShareSetup, 27)
        ''------------------------------TRANSACTION--------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleShareTransaction, "Transaction", "1.48.02", clsUserMgtCode.ModuleShare, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmMultipleShareAllotment, "Multiple Share Allotment", "1.48.02.01", clsUserMgtCode.SubModuleShareTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmShareAllotment, "Share Allotment", "1.48.02.02", clsUserMgtCode.SubModuleShareTransaction, 8)
        ''----------------------------------------------REPORT---------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleShareReport, "Report", "1.48.03", clsUserMgtCode.ModuleShare, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptShareRegister, "Share Register", "1.48.03.01", clsUserMgtCode.SubModuleShareReport, 37)

#End Region


#Region "MIS"
        '----------------------------------MIS---------------------------------------'

        InsertDefaultValue(clsUserMgtCode.ModuleMIS, "MIS", "1.49", clsUserMgtCode.ExpertERP, "M", 23)

        '' setup
        '=========================================================









        InsertDefaultValue(clsUserMgtCode.SubModuleMISSetUp, "Setup", "1.49.01", clsUserMgtCode.ModuleMIS, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.MISitemGroups, "MIS Item Group", "1.49.01.01", clsUserMgtCode.SubModuleMISSetUp, 27)
        InsertDefaultValue(clsUserMgtCode.MISitemMaster, "MIS Item Master", "1.49.01.02", clsUserMgtCode.SubModuleMISSetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmProductionTransactionType, "Production Transaction Type", "1.49.01.03", clsUserMgtCode.SubModuleMISSetUp, 27)

        ''Transaction
        ''============================================
        InsertDefaultValue(clsUserMgtCode.SubModuleMISTransaction, "Transaction", "1.49.02", clsUserMgtCode.ModuleMIS, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmDailyMilkProducts, "Daily Milk Products", "1.49.02.01", clsUserMgtCode.SubModuleMISTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)
        InsertDefaultValue(clsUserMgtCode.frmDailySMPProduction, "Daily SMP Production", "1.49.02.02", clsUserMgtCode.SubModuleMISTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)

        'Try
        '    clsDBFuncationality.ExecuteNonQuery("insert into TSPL_MODULE_PERMISSION values ('1','" + clsUserMgtCode.ModuleReco + "')")
        'Catch ex As Exception

#End Region

#Region "Xpert API"
        InsertDefaultValue(clsUserMgtCode.ModuleXpertAPI, "Xpert API", "1.50", clsUserMgtCode.ExpertERP, "M", 23, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.SubModuleXpertAPISetup, "Setup", "1.50.01", clsUserMgtCode.ModuleXpertAPI, "SM", 26, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)

        InsertDefaultValue(clsUserMgtCode.SubModuleXpertAPITrans, "Transaction", "1.50.02", clsUserMgtCode.ModuleXpertAPI, "SM", 26, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIWeighment, "Weighment", "1.50.02.01", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPILoadinSlip, "Loadin slip", "1.50.02.02", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIMilkReceipt, "Milk Receipt", "1.50.02.03", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIMilkSample, "Milk Sample", "1.50.02.04", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIMilkReject, "Milk Reject", "1.50.02.05", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIMilkEmptySample, "Empty Sample", "1.50.02.06", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIPlantWeighment, "Weighment", "1.50.02.07", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIQualityCheck, "Quality Check", "1.50.02.08", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIParlourShiftOpen, "Parlour Shift Open", "1.50.02.09", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIParlourPOS, "POS", "1.50.02.10", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIParlourShiftClose, "Parlour Shift Close", "1.50.02.11", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIDairyGatePass, "Dairy GatePass", "1.50.02.12", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIBMCDCSSamplingReceiving, "BMC DCS Sampling Receiving", "1.50.02.13", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIGateEntry, "Gate Entry", "1.50.02.14", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIBMCMilkQC, "BMC Milk QC", "1.50.02.15", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIDEMANDBOOKING, "Demand Booking", "1.50.02.16", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPISaleDispatchDairy, "Dairy Dispatch", "1.50.02.17", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        Dim IsRCDFCFP As Boolean
        Try
            IsRCDFCFP = IIf(clsCommon.myCdbl(clsFixedParameter.GetData(clsFixedParameterType.RCDFCFP, clsFixedParameterCode.RCDFCFP, Nothing)) = 1, True, False)
            If IsRCDFCFP Then
                InsertDefaultValue(clsUserMgtCode.XpertAPIRCDFChemicalQC, "RCDF Chemical QC", "1.50.02.17", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
                InsertDefaultValue(clsUserMgtCode.XpertAPIRCDFUnionMilkQC, "RCDF Union Milk QC", "1.50.02.18", clsUserMgtCode.SubModuleXpertAPITrans, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
            End If
        Catch ex As Exception

        End Try

        InsertDefaultValue(clsUserMgtCode.SubModuleXpertAPIReport, "Report", "1.50.03", clsUserMgtCode.ModuleXpertAPI, "SM", 30, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIMilkMilkRegister, "Milk Register", "1.50.03.01", clsUserMgtCode.SubModuleXpertAPIReport, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPITruckSheet, "Truck Sheet Report", "1.50.03.02", clsUserMgtCode.SubModuleXpertAPIReport, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPIBMCDCSSampleReceivedRegister, "BMC DCS Sample Received Register", "1.50.03.03", clsUserMgtCode.SubModuleXpertAPIReport, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        InsertDefaultValue(clsUserMgtCode.XpertAPITankerCollectionReport, "Tanker Collection Report", "1.50.03.04", clsUserMgtCode.SubModuleXpertAPIReport, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        If IsRCDFCFP Then
            InsertDefaultValue(clsUserMgtCode.XpertAPIRCDFChemicalQCRegister, "RCDF Chemical QC Register", "1.50.03.05", clsUserMgtCode.SubModuleXpertAPIReport, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
            InsertDefaultValue(clsUserMgtCode.XpertAPIRCDFUnionMilkQCRegister, "RCDF Union Milk QC Register", "1.50.03.06", clsUserMgtCode.SubModuleXpertAPIReport, "", 8, Nothing, Nothing, Nothing, Nothing, True, True, True, True, Nothing, "", "", "", "", "", False, "", 1)
        End If
#End Region

#Region "Milk Union"
        InsertDefaultValue(clsUserMgtCode.ModuleMilkUnion, "Milk Union", "1.51", clsUserMgtCode.ExpertERP, "M", 23)
        InsertDefaultValue(clsUserMgtCode.SubModuleMilkUnion, "Report", "1.51.03", clsUserMgtCode.ModuleMilkUnion, "SM", 30)
        InsertDefaultValue(clsUserMgtCode.rptMilkUnion, "Milk Union Report", "1.51.03.01", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptUnionMilkPurchase, "Union Milk SRN Report", "1.51.03.01", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.latestPaymentProcess, "latest Payment Process", "1.51.03.02", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDBTNEFTUnionReport, "DBT Union Report", "1.51.03.03", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDBTStatusAndLastDPTStatus, "DBT Status And Last DPT Status", "1.51.03.04", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDBTApprovalStatus, "DBT Approval Status", "1.51.03.05", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDBTNEFTPaymentDetailReport, "DBT NEFT Payment Detail Report", "1.51.03.06", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptExeVersionReport, "Exe Version Report", "1.51.03.07", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptProductionReport, "Production Report", "1.51.03.08", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptProductionStoreReport, "Production and Store Report", "1.51.03.09", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.frmDailyDemand, "Daily Demand", "1.51.03.10", clsUserMgtCode.SubModuleSaleDairySetUp, 27)
        InsertDefaultValue(clsUserMgtCode.frmDBTUnionPayment, "DBT Union Payment Report", "1.51.03.11", clsUserMgtCode.SubModuleSaleDairySetUp, "", 27, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmDBTUnionPayment, "DBT Union Payment Report", "1.51.03.12", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmUnionLastGstInvoice, "Union wise Last GST invoice", "1.51.03.13", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptCattleFeedProduction, "Cattle Feed Production", "1.51.03.17", clsUserMgtCode.SubModuleMilkUnion, 8)
        'InsertDefaultValue(clsUserMgtCode.frmUpdateAppLocation, "Update App Location", "1.51.03.19", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptUnionBookingReport, "Milk Marketing Report", "1.51.03.20", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.DBTPDAccountReport, "DBT PD Account Report", "1.51.03.21", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmUnionWiseAppUserReport, "Union Wise App User Report", "1.51.03.22", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.FrmCmSambalYojanaSubsidy, "Cm Sambal Yojana Subsidy", "1.51.03.23", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDBTBankResponse, "DBT Bank Response", "1.51.03.24", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptDBTStatusReport, "DBT Status Report", "1.51.03.25", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptMachineSurveyRegister, "Machine Survey Register", "1.51.03.26", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptUnionWiseMilkTankerCollectionDetail, "Union Wise Milk Tanker Collection Detail", "1.51.03.27", clsUserMgtCode.SubModuleMilkUnion, 8)
        InsertDefaultValue(clsUserMgtCode.rptMilkUnion2, "New Milk Union Report", "1.51.03.28", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptDbtStatusRCDF, "Dbt Status RCDF Report", "1.51.03.30", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.frmUnionWiseDockEntryReport, "Union Wise Dock Entry Report ", "1.51.03.31", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.FrmInActiveDCSReport, "InActive DCS Report", "1.51.03.33", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)
        InsertDefaultValue(clsUserMgtCode.rptFarmerMilkCollectionReport, "Farmer Milk Collection Report", "1.51.03.35", clsUserMgtCode.SubModuleMilkUnion, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, False, False, True)

#End Region

#Region "Bull Management"
        InsertDefaultValue(clsUserMgtCode.ModuleBullManagement, "Bull Management", "1.52", clsUserMgtCode.ExpertERP, "M", 23)


        ''----------------SETUP--------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleBullManagementSetup, "Setup", "1.52.01", clsUserMgtCode.ModuleBullManagement, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmBullMovementType, "Bull movement Type", "1.52.01.01", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullBreedMaster, "Bull Breed Master", "1.52.01.02", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullTestParameter, "Bull Test Parameter", "1.52.01.03", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullParameterGroup, "Bull Parameter Group", "1.52.01.04", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullShedMaster, "Bull Shed Master", "1.52.01.05", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullShedParameterMaster, "Bull Shed Parameter Master", "1.52.01.06", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullShedParameterGroup, "Bull Shed Parameter Group", "1.52.01.07", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullMasters, "Bull Master", "1.52.01.08", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullStatus, "Bull Status", "1.52.01.09", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullSubStatus, "Bull Sub Master", "1.52.01.10", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullSourcePainting, "Bull Source Painting", "1.52.01.11", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullRating, "Bull Rating", "1.52.01.12", clsUserMgtCode.SubModuleBullManagementSetup, 27)



        InsertDefaultValue(clsUserMgtCode.frmBullSpeciesMaster, "Bull Species Master", "1.52.01.13", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullCategoryMaster, "Bull Category Master", "1.52.01.14", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullSubCategoryMaster, "Bull Sub Category Master", "1.52.01.15", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullSSCentreMaster, "Bull SS Centre Master", "1.52.01.16", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullPenIDMaster, "Bull Pen ID Master", "1.52.01.17", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullSourceName, "Bull Source Name Master", "1.52.01.18", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullPurchaseNo, "Bull Purchase Request Master", "1.52.01.19", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullInsuranceMaster, "Bull Insurance Master", "1.52.01.20", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullInsuranceType, "Bull Insurance Type", "1.52.01.21", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmBullSourceName, "Bull Source Name Master", "1.52.01.18", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        'InsertDefaultValue(clsUserMgtCode.frmBullPurchaseNo, "Bull Purchase Request No", "1.52.01.18", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmDiseaseMaster, "Disease Master", "1.52.01.22", clsUserMgtCode.SubModuleBullManagementSetup, 27)
        InsertDefaultValue(clsUserMgtCode.frmBullCMUGrouping, "Bull CMU Grouping", "1.52.01.25", clsUserMgtCode.SubModuleBullManagementSetup, 27)


        ''------------------------------TRANSACTION--------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleBullManagementTransaction, "Transaction", "1.52.02", clsUserMgtCode.ModuleBullManagement, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmBullMovement, "Bull Movement", "1.52.02.01", clsUserMgtCode.SubModuleBullManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBullVaccinationEntry, "Bull Vaccination Entry", "1.52.02.02", clsUserMgtCode.SubModuleBullManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBullCurlingEntry, "Bull Curling Entry", "1.52.02.03", clsUserMgtCode.SubModuleBullManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBullInsurance, "Bull Insurance", "1.52.02.04", clsUserMgtCode.SubModuleBullManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmBullTestParameterEntry, "Bull Test Parameter Entry", "1.52.02.05", clsUserMgtCode.SubModuleBullManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmInsuranceTagAllocation, "Insurance Tag Allocation", "1.52.02.06", clsUserMgtCode.SubModuleBullManagementTransaction, 8)
        InsertDefaultValue(clsUserMgtCode.frmCMUChecklistEntry, "Bull CMU CheckList Entry", "1.52.02.08", clsUserMgtCode.SubModuleBullManagementTransaction, 8)

        ''----------------------------------------------REPORT---------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleBullManagementReport, "Report", "1.52.03", clsUserMgtCode.ModuleBullManagement, "SM", 30)


#End Region

#Region "Seed"
        ''=======================================SHEED=============================================
        InsertDefaultValue(clsUserMgtCode.ModuleSheed, "Seed", "1.53", clsUserMgtCode.ExpertERP, "M", 23)

        ''---------------------------------------SETUP--------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleSheedSetUp, "Setup", "1.53.01", clsUserMgtCode.ModuleSheed, "SM", 24)
        InsertDefaultValue(clsUserMgtCode.frmSeedGrowerMaster, "Seed Grower Master", "1.53.01.01", clsUserMgtCode.SubModuleSheedSetUp, 27)

        ''--------------------------------------TRANSACTION--------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleSheedTransaction, "Transaction", "1.53.02", clsUserMgtCode.ModuleSheed, "SM", 26)
        InsertDefaultValue(clsUserMgtCode.frmSeedSelectionEntry, "Seed Selection Entry", "1.53.02.01", clsUserMgtCode.SubModuleSheedTransaction, "", 8, 0, "", "", "", False, False, False, "", Nothing, "", "", "", "", "", False, "", 0, True, False)


        ''---------------------------------------REPORT---------------------------------------------------------
        InsertDefaultValue(clsUserMgtCode.SubModuleSheedReport, "Report", "1.53.03", clsUserMgtCode.ModuleSheed, "SM", 30)

#End Region









        Try
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.BankCashBook, "Bank Cash Book")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.VehicleUtili, "Vehicle Utilization")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.ProcMilkPur, "Procurement Milk Purchase")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.MilkRecAtFac, "Milk Received at Factory")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.MilkSale, "Milk Sale")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.ProductSale, "Product Sale")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.Transportcos, "Transport cost")
            clsDashBoard.InsertDefaultValueDashBoard(clsDashBoard.StoreReport, "Store Report")
        Catch ex As Exception
        End Try
        'Dashboard
    End Sub

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal ImageNumber As Integer) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, "", ImageNumber)
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, Nothing)
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, ByVal trans As SqlTransaction) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, False, "", "", "", trans)
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal trans As SqlTransaction) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, False, False, trans)
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal trans As SqlTransaction) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, False, trans)
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal trans As SqlTransaction) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, Is_Notification_Applied, "", trans)
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal AmdPassword As String, ByVal trans As SqlTransaction) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, Is_Notification_Applied, AmdPassword, trans, "", "", "", "", "")
    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal AmdPassword As String, ByVal trans As SqlTransaction, ByVal strES_Trans_Type_1 As String, ByVal strES_Trans_Type_2 As String, ByVal strES_Trans_Type_3 As String, ByVal strES_Trans_Type_4 As String, ByVal strES_Trans_Type_5 As String) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, Is_Notification_Applied, AmdPassword, trans, strES_Trans_Type_1, strES_Trans_Type_2, strES_Trans_Type_3, strES_Trans_Type_4, strES_Trans_Type_5, False, "")
    End Function
    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal AmdPassword As String, ByVal trans As SqlTransaction, ByVal strES_Trans_Type_1 As String, ByVal strES_Trans_Type_2 As String, ByVal strES_Trans_Type_3 As String, ByVal strES_Trans_Type_4 As String, ByVal strES_Trans_Type_5 As String, ByVal AddAsNewMenu As Boolean, ByVal ProgramCodeOriginal As String) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, Is_Notification_Applied, AmdPassword, trans, strES_Trans_Type_1, strES_Trans_Type_2, strES_Trans_Type_3, strES_Trans_Type_4, strES_Trans_Type_5, AddAsNewMenu, ProgramCodeOriginal, 0)
    End Function
    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal AmdPassword As String, ByVal trans As SqlTransaction, ByVal strES_Trans_Type_1 As String, ByVal strES_Trans_Type_2 As String, ByVal strES_Trans_Type_3 As String, ByVal strES_Trans_Type_4 As String, ByVal strES_Trans_Type_5 As String, ByVal AddAsNewMenu As Boolean, ByVal ProgramCodeOriginal As String, ByVal intAPP_No As Integer) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, Is_Notification_Applied, AmdPassword, trans, strES_Trans_Type_1, strES_Trans_Type_2, strES_Trans_Type_3, strES_Trans_Type_4, strES_Trans_Type_5, AddAsNewMenu, ProgramCodeOriginal, intAPP_No, False, False)

    End Function

    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal AmdPassword As String, ByVal trans As SqlTransaction, ByVal strES_Trans_Type_1 As String, ByVal strES_Trans_Type_2 As String, ByVal strES_Trans_Type_3 As String, ByVal strES_Trans_Type_4 As String, ByVal strES_Trans_Type_5 As String, ByVal AddAsNewMenu As Boolean, ByVal ProgramCodeOriginal As String, ByVal intAPP_No As Integer, ByVal isLockLocation As Boolean, ByVal isLockLocationSegment As Boolean) As Boolean
        Return InsertDefaultValue(strProgramCode, strProgramName, strSNo, strParent_Code, strType, ImageNumber, IsRunFromOtherAsm, AddAs, AsmPath, FormName, Is_SMS_Applied, Is_EMAIL_Applied, Is_Notification_Applied, AmdPassword, trans, strES_Trans_Type_1, strES_Trans_Type_2, strES_Trans_Type_3, strES_Trans_Type_4, strES_Trans_Type_5, AddAsNewMenu, ProgramCodeOriginal, intAPP_No, isLockLocation, isLockLocationSegment, False)

    End Function
    Public Shared Function InsertDefaultValue(ByVal strProgramCode As String, ByVal strProgramName As String, ByVal strSNo As String, ByVal strParent_Code As String, ByVal strType As String, ByVal ImageNumber As Integer, IsRunFromOtherAsm As Integer, AddAs As String, AsmPath As String, FormName As String, ByVal Is_SMS_Applied As Boolean, ByVal Is_EMAIL_Applied As Boolean, ByVal Is_Notification_Applied As Boolean, ByVal AmdPassword As String, ByVal trans As SqlTransaction, ByVal strES_Trans_Type_1 As String, ByVal strES_Trans_Type_2 As String, ByVal strES_Trans_Type_3 As String, ByVal strES_Trans_Type_4 As String, ByVal strES_Trans_Type_5 As String, ByVal AddAsNewMenu As Boolean, ByVal ProgramCodeOriginal As String, ByVal intAPP_No As Integer, ByVal isLockLocation As Boolean, ByVal isLockLocationSegment As Boolean, ByVal Web_Implemented As Boolean) As Boolean
        Dim qry As String = "select PROGRAM_NAME from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "' "
        Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)

        Dim coll As New Hashtable()
        clsCommon.AddColumnsForChange(coll, "Program_Name", strProgramName)
        clsCommon.AddColumnsForChange(coll, "SNo", strSNo)
        clsCommon.AddColumnsForChange(coll, "Type", strType)
        clsCommon.AddColumnsForChange(coll, "Parent_Code", strParent_Code, True)
        clsCommon.AddColumnsForChange(coll, "Image_Number", ImageNumber)
        clsCommon.AddColumnsForChange(coll, "IsLoadFromOtherAssembly", clsCommon.myCdbl(IsRunFromOtherAsm))
        clsCommon.AddColumnsForChange(coll, "OtherAssemblyFilePathAndName", clsCommon.myCstr(AsmPath))
        clsCommon.AddColumnsForChange(coll, "FormName", clsCommon.myCstr(FormName))
        clsCommon.AddColumnsForChange(coll, "addAs", clsCommon.myCstr(AddAs))
        clsCommon.AddColumnsForChange(coll, "Modify_By", objCommonVar.CurrentUserCode)
        clsCommon.AddColumnsForChange(coll, "Modify_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))

        clsCommon.AddColumnsForChange(coll, "Is_SMS_Applied", IIf(Is_SMS_Applied, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Is_EMAIL_Applied", IIf(Is_EMAIL_Applied, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Is_Notification_Applied", IIf(Is_Notification_Applied, 1, 0))
        'done by stuti on 07/03/2017
        If clsCommon.myLen(AmdPassword) > 0 Then
            clsCommon.AddColumnsForChange(coll, "Amendment_Pwd", clsCommon.EncryptString(AmdPassword))
        End If

        clsCommon.AddColumnsForChange(coll, "ES_Trans_Type_1", strES_Trans_Type_1)
        clsCommon.AddColumnsForChange(coll, "ES_Trans_Type_2", strES_Trans_Type_2)
        clsCommon.AddColumnsForChange(coll, "ES_Trans_Type_3", strES_Trans_Type_3)
        clsCommon.AddColumnsForChange(coll, "ES_Trans_Type_4", strES_Trans_Type_4)
        clsCommon.AddColumnsForChange(coll, "ES_Trans_Type_5", strES_Trans_Type_5)
        clsCommon.AddColumnsForChange(coll, "APP_No", intAPP_No)
        'clsCommon.AddColumnsForChange(coll, "Customise_SNo", strSNo)

        clsCommon.AddColumnsForChange(coll, "Lock_Location_Segment", IIf(isLockLocationSegment, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Lock_Location", IIf(isLockLocation, 1, 0))
        clsCommon.AddColumnsForChange(coll, "Web_Implemented", IIf(Web_Implemented, 1, 0))
        If dt Is Nothing OrElse dt.Rows.Count <= 0 OrElse AddAsNewMenu Then
            clsCommon.AddColumnsForChange(coll, "Program_Code", strProgramCode)
            clsCommon.AddColumnsForChange(coll, "Comp_Code", objCommonVar.CurrentCompanyCode)
            clsCommon.AddColumnsForChange(coll, "Created_By", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "Created_Date", clsCommon.GetPrintDate(clsCommon.GETSERVERDATE(trans), "dd/MMM/yyyy"))
            clsCommon.AddColumnsForChange(coll, "Customise_SNo", strSNo) ' Ticket No : TEC/06/06/19-000524 By prabhakar
            clsCommon.AddColumnsForChange(coll, "Program_Code_Original", ProgramCodeOriginal, True)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PROGRAM_MASTER", OMInsertOrUpdate.Insert, "", trans)
        Else
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_PROGRAM_MASTER", OMInsertOrUpdate.Update, "Program_Code='" + strProgramCode + "'", trans)
        End If
        Return True
    End Function

    Public Shared Function AddCustomLinks(ByVal strProgramCode As String, ByVal strProgramDescription As String, ByVal strParent_Code As String, ByVal trans As SqlTransaction) As Boolean
        Dim strParentSNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select SNo from TSPL_PROGRAM_MASTER where Program_Code='" + strParent_Code + "'", trans))
        Dim qry As String = "select MAX(SNo) from TSPL_PROGRAM_MASTER where SNo like   '" + strParentSNo + "+.100%'"
        Dim strSNo As String = clsCommon.myCstr(clsDBFuncationality.getSingleValue(qry, trans))
        If clsCommon.myLen(strSNo) <= 0 Then
            strSNo = strParentSNo + ".1001"
        Else
            strSNo = clsCommon.incval(strSNo)
        End If
        ProgramCodeNew.InsertDefaultValue(strProgramCode, strProgramDescription, strSNo, strParent_Code, "", 30, trans)
        Return True
    End Function

    Public Shared Function GetProgramName(ByVal strProgramCode As String) As String
        Dim ProgramName As String = String.Empty
        ProgramName = clsCommon.myCstr(clsDBFuncationality.getSingleValue("select case when LEN(ISNULL(Re_Name,''))>0 then Re_Name else Program_Name end as Program_Name from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "'"))
        Return ProgramName
    End Function

End Class

Public Class clsFavouriteMenu
    Public Program_Code As String = Nothing
    Public User_Code As String = Nothing
    Public SNo As String = Nothing

    Public Shared Function SaveData(ByVal strProgramCode As String) As Boolean
        Dim trans As SqlTransaction = clsDBFuncationality.GetTransactin()
        Try
            Dim qry As String = "select 1 from TSPL_FAVOURITE_MENU where Program_Code='" + strProgramCode + "' and User_Code='" + objCommonVar.CurrentUserCode + "'"
            Dim dt As DataTable = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("Already added in favourite")
            End If
            qry = "select 1 from TSPL_PROGRAM_MASTER where Program_Code='" + strProgramCode + "' and Type in ('M','SM')"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Throw New Exception("Module or Submodule can not be entered into favourite")
            End If

            Dim IntSNo As Integer = 1
            qry = "select 1 from TSPL_FAVOURITE_MENU where User_Code='" + objCommonVar.CurrentUserCode + "'"
            dt = clsDBFuncationality.GetDataTable(qry, trans)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                IntSNo = dt.Rows.Count + 1
            End If
            Dim strSNo As String = ""
            For ii As Integer = 1 To 3 - clsCommon.myLen(IntSNo)
                strSNo += "0"
            Next
            strSNo += clsCommon.myCstr(IntSNo)

            Dim coll As New Hashtable()
            clsCommon.AddColumnsForChange(coll, "Program_Code", strProgramCode)
            clsCommon.AddColumnsForChange(coll, "User_Code", objCommonVar.CurrentUserCode)
            clsCommon.AddColumnsForChange(coll, "SNo", "0.1." + strSNo)
            clsCommonFunctionality.UpdateDataTable(coll, "TSPL_FAVOURITE_MENU", OMInsertOrUpdate.Insert, "", trans)
            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw New Exception(ex.Message)
        End Try

        Return True
    End Function

    Public Shared Function DeleteData(ByVal strProgramCode As String) As Boolean
        Dim qry As String = "delete from TSPL_FAVOURITE_MENU where Program_Code='" + strProgramCode + "' and User_Code='" + objCommonVar.CurrentUserCode + "'"
        clsDBFuncationality.ExecuteNonQuery(qry)
        Return True
    End Function
End Class
Public Class clsProgramIdFormNameMapping

    Public Shared Sub setAllProgramFormName()
        UpdateFormName(clsUserMgtCode.FrmCompanyMaster, "FrmCompanyMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CostCenter, "FrmCostCenter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CostFACenter, "FrmFACostCenter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ReverseEntry, "FrmReverseEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.cityMaster, "frmCityMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.taxAuthority, "frmTaxAuthority", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.taxRate, "FrmTaxRates", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.taxGroup, "FrmTaxGroups", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.paymentTerms, "frmPaymentTerms", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.paymentCodes, "FrmPaymentCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.bankMaster, "frmBankMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.bankBranchMaster, "FrmBankBrachMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmNotificationScreen, "FrmNotificationScreen", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ChangePwd, "FrmChangePassword", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAbateMentMaster, "FrmAbateMentMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PrefixGeneration, "FrmPrefixGenerationNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FormMaster, "FrmFormMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFormSerialNoMaster, "FrmFormSerialNoMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAdditionalCharges, "FrmAdditionalCharges", "ERP.EXE")
        UpdateFormName("FrmChangePassword", "FrmChangePassword", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmCurrencyConversion, "frmCurrencyConversion", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.CustomFieldMaster, "FrmCustomFieldMaster", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.CustomFieldMapping, "frmCustomFieldMapping", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmModuleCurrencyMapping, "frmModuleCurrencyMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CommonServicesSetting, "frmCommonServicesSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRegionMaster, "FrmRegionMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCountryMaster1, "frmCountryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStateMaster1, "frmStateMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DistrictMaster, "frmDistrictMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AreaMaster, "frmAreaMaster", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.FrmBranchAccountMapping, "FrmBranchAccountMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMCCDiscountMaster, "FrmDiscountMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.bankTransfer, "FrmBankTransfer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.reverseTransaction, "frmReverseTransaction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBankBook, "FrmBankBook", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBankBookChart, "FrmBankBookChart", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBankBookClosing, "FrmBankBookClosing", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCustomerAgingSummary, "FrmBICustomerAgeing", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorAgingSummary, "FrmBIVendorAgeing", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmFormCollection, "FrmBIVendorAgeing", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorAgingSummary, "FrmBIVendorAgeing", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmFormCollection, "FrmFormCollection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmLoadReport, "FrmLoadReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLoadOutInvoiceRecoReport, "FrmLoadOutInvoiceRecoReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCFormEntry, "FrmCFormEntry", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.FrmBankGuaranteeMaster1, "FrmBankGuaranteeMaster1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BankOpeningReco, "frmBankOpeningReco", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.RevaluationEntry, "frmRevaluationEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DVAT30, "FrmDVAT30", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DVAT31, "FrmDVAT31", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDetailsOfForm2B, "FrmDetailsOfForm2B", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCashVoucher, "FrmCashVoucher", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TaxTracking, "FrmTaxTracking", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FormIssue, "FrmFormIssueDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBankReconcilliation, "RptBankReconcilliation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmShipToLocationDetails, "FrmShipToLocationDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmReceivablePaymentTerms, "FrmReceivablePaymentTerms", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmReceivableSettings, "FrmReceivableSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerType, "frmCustomerType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomeCategory, "frmCustomerCategory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerMaster, "frmCustomer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerGroup, "frmCustomerGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerAccountSet, "frmCustomerAccountSet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ReceiptEntry, "FrmReceipttNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ReceiptAdjustmentEntry, "frmAdj", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnARInvoiceEntry, "FrmARInvoiceEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQuickBook, "FrmQuickEntry1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBankReco, "FrmBankReco", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCustomerInquiry, "FrmCustomerInquiry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerGroupReport, "FrmCustomerGroupReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerDetails, "FrmCustomerMasterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCustomerLedger, "FrmRptCustomerLedgerDemo", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRoute_CustomerOutstanding, "FrmRoute_CustomerOutStanding", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerAgeing, "rptCustomerAgeingDrillDown", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomersListReport, "frmCustomerListRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RouteListReport11, "frmRouteListReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCustomerEmptyTrial, "FrmCustomerEmptyTrial2", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSecurityLevel, "FrmSecurityLevel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptVendorSecurity, "FrmVendorSecurity", "ERP.EXE")
        If objCommonVar.IsDemoERP Then
            UpdateFormName(clsUserMgtCode.frmExciseChapterWise, "frmER1Demo", "ERP.EXE")
        Else
            UpdateFormName(clsUserMgtCode.frmExciseChapterWise, "frmExciseChapterWise", "ERP.EXE")
        End If
        UpdateFormName(clsUserMgtCode.FrmBankReverse, "FrmBankReverse", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptQualityStatus, "RptQualityStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.glOptions, "frmgloption", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.segmentCode, "Frmsegmentcode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.accountStructure, "frmGLStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.accountGroup, "frmAccountGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.glAccount, "frmGLAccount", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.createAccounts, "frmCreateAccountNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.sourceCode, "FrmSourceCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.glsecurity, "Frmglsecurity", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBalanceSheetPerforma, "FrmBalanceSheetPerforma1", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.FrmGL_account_excluded, "FrmGL_account_excluded", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmGLControlAccountMapping, "frmGLControlAccountMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frm_Account_Mapping, "Frm_Account_Mapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMapLedgerAccToTally, "frmMapLedgerAccToTally", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPostAllGLToTally, "frmPostAllGLToTally", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FiscalYear, "frmFinancialYearMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CostCentreFinancial, "FrmCostCentreFinancial", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.JECheckSystem, "rptJECheck", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AccountMainGroup, "FrmAccountMainGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AccountSubGroup, "FrmAccountSubGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AccountGLMainAccount, "frmAccountMainGLAccount", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.journalEntry, "frmJournalEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnVCGLEntry, "frmVCGLEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmJEReverse, "FrmJEReverse", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCostCenterAnalysisRpt, "FrmCostCenterAnalysisRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGLTransReport, "GLTransReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmJrnlVoucher, "JrnlVoucherReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptTrialBalance, "frmRptTrialBalanceNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptTrialBalanceCV, "frmRptTrialBalanceVC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnJournalBook, "frmRptDayWiseJournalBook", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptBalanceSheet, "frmRptBalanceSheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBankBookDayWise, "FrmBankBookDayWise", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmUnpostedJV, "FrmUnpostedJV", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.EmployeeMaster, "frmEmployeeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DesignationMaster, "frmDesignationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DesignationMasterHierarchy, "frmDesignationHierarchyMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.userMaster, "FrmUserMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.userGroupMaster, "FrmUserGroupMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.userGroupMapping, "FrmUserGroupMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.groupProgramMapping, "GroupProgramMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmScheduling, "FrmScheduling", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmApprovalLevelScreen, "frmApprovalScreen", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmApprovalAlertSumm, "FrmApprovalAlertSumm", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLocationSetting, "frmLocationLogin", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmSynchronization, "frmSynchronization", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FisaclYearEndProcess, "FrmFiscalYearEndProcess", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAppIntegrator, "frmAppIntegrator", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.vendormaster, "frmVendorMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.capexmaster, "frmCapexMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.capexbudget, "frmCapexBudet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DispatchChecklist, "frmDispatchChecklist", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.vendoraccountset, "frmvendoraccountset", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.vendorgroup, "frmVendorGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.vendortype, "frmVendorType", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmRequisitionApproval, "FrmRequisitionApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RequisitSubTypeMaster, "FrmRequisitSubTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHirerachyLevelMaster, "FrmHirerachyLevelMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPayableSettings, "FrmPayableSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPaymentUploader, "FrmPaymentUploader", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorBankMaster, "FrmVendorBankMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PaymentEntryNew, "FrmPaymentNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptMultipleRTGS, "RptMultipleRTGS", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptMultiplePaymentAdvice1, "RptMultiplePaymentAdvice1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PaymentAdjustmentEntry, "frmPaymentAdjEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnAPInvoiceEntry, "FrmAPInvoiceEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVendorService, "FrmVendorService", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVendorInquiry, "FrmVendorInquiry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPurchaseHistory, "FrmPurchaseHistory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSupplierReg, "FrmSupplierReg", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorRegistration, "frmVendorReg", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmApprovedSuppliers, "FrmApprovedSuppliers", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmScrapSaleInvoice, "FrmScrapSaleInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPaymentEntry, "FrmPaymentEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnAPInvoiceReport, "frmRptAPInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRptAPInvoiceDetailsReport, "FrmRptAPInvoiceDetailsReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorLedgerReport, "frmRptVendorLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorCustomerLedgerReport, "frmRptVendorCustomerLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAgingPayble, "rptAPAgeingDrillDown", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRptVendorList, "FrmRptVendorList", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRptVendorTransList, "FrmRptVendorTransList", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PricePlan, "frmPriceMasterPlan", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PriceMaster, "FrmPriceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PriceComponentMasters, "FrmPriceComponantMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PriceComponentMapping, "FrmPriceComponantMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SchemeMaster, "FrmSchmeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.inventorySetting, "frmInventorySetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.chapterhead, "frmChapterHead", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.itemStructure, "frmItemStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.itemGroups, "frmItemGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.itemMaster, "frmItemMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.locationMaster, "frmLocationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.itemPurchaseAccount, "frmPurcahseAccountSetCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.itemSaleAccount, "frmSaleAccountSetCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.unitMaster, "frmUnitOfCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.packType, "Frmpacktype", "ERP.EXE")
        UpdateFormName("frmExcisableLocationDetails", "FrmExcisableLocationDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmItemMasterRMOther, "FrmItemMasterRMOther", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TypeOfItemMaster, "TypeOfItemMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPartNoMaster, "FrmPartNoMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.InvetorySourceCode, "frmInventorySourceCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGrievanceTypeMaster, "frmGrievanceTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGrievanceLogging, "frmGrievanceLogging", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGrievanceAllocation, "FrmGrievanceAllocation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGrievanceResolution, "FrmGrievanceResolution", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetCategoryMaster, "FrmAssetTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetSubCategoryMaster, "FrmAssetSubCategoryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetMaster, "FrmAssetDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetIssueReturn, "frmAssetsIssueReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRSettings, "FrmHRSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmResignationLetter, "FrmResignationLetter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptRegisterOfDeduction, "RptRegisterOfDeduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTerminationLetter, "FrmHREXTerminationLetter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmResignationAcceptanceOrRejection, "FrmResignationAcceptanceORRejection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHREMInterviewQuestion, "FrmHREMInterviewQuestion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHREMExitInterview, "FrmExitInterview", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRPerformanceRatingRpt, "rptPerformanceRating", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmItemListRpt, "FrmItemListRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVendorListRPT, "FrmVendorListRPT", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemReorderLevel, "frmItemReorderLevel1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnItemCategory, "FrmItemCategory1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnItemSubCategory, "FrmItemSubCategory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemExciseMapping, "FrmItemExciseMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemBasicPrice, "FrmItemBasicPrice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStandardscheme, "frmStandardscheme", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStandardRateItem, "frmStandardRateItem", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmItemCategoryLevel, "frmItemCategoryLevel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorCategoryLevel, "frmItemCategoryLevel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerCategoryLevel, "frmItemCategoryLevel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLocationCategoryLevel, "frmItemCategoryLevel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmItemCategoryStructure, "frmItemCategoryStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorCategoryStructure, "frmItemCategoryStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerCategoryStructure, "frmItemCategoryStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLocationCategoryStructure, "frmItemCategoryStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCatalogMaster, "FrmCatalogMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBarCodeGenerator, "FrmBarCodeGenerator", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBarCodeGenerator1, "FrmBarCodeGenerator1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.WarrantyMaster, "frmWarrentyMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSchemeMasterNew, "FrmSchemeMasterNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmWeightConversion, "FrmWeightCoversion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmInventoryAgeingReport, "FrmInventoryAgeingReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPriceGroupMapping, "FrmPriceGroupMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmTragetMaster, "FrmTragetMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPrintProductInvoiceStatement, "FrmPrintProductInvoiceStatement", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.RptProductSaleRegister1, "RptProductSaleRegister1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptProductBookingStatus, "RptProductBookingStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptProductDispatchStatus, "RptProductDispatchStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptProductDOStatus, "RptProductDOStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptProductSaleOrderStatus, "RptProductSaleOrderStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnEmptyTrans, "frmAdjustmentEmpty", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnProductionEntry, "frmAdjustmentProduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnStoreAdjustment, "frmAdjustmentStore", "ERP.EXE")
        If clsCommon.CompairString(objCommonVar.CurrentCompanyCode, "KL") = CompairStringResult.Equal OrElse clsCommon.CompairString(objCommonVar.CurrentIndustryType, "D") = CompairStringResult.Equal Then
            UpdateFormName(clsUserMgtCode.Transfer, "FrmTransferKDIL", "ERP.EXE")
        Else
            UpdateFormName(clsUserMgtCode.Transfer, "frmTransferDCC", "ERP.EXE")
        End If
        UpdateFormName(clsUserMgtCode.FrmTransferGateOut, "FrmTransferGateOut", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TransferReturn, "frmTransferKDILReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CreateTransfer, "FrmTransfer3rdDoc", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Indent, "frmIndent", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmItemMcMapping, "FrmItemMcMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmWarehouseBreakage, "FrmWarehouseBreakage", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmExpiryDateEntry, "FrmExpiryDateEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPhysicalStock, "FrmPhysicalStock", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPhysicalStockMultipleLocation, "frmPhysicalStockMultipleLocation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ChangeItemSerialNumber, "frmChangeSerialNumber", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemStockConversion, "frmItemToItemStockConverion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnItemMovement, "frmRptInventoryMovement", "ERP.EXE")
        UpdateFormName("mtbnTransfer", "frmRptTransfer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.KeyValue, "FrmKeyvalueReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmLeakageBreakage, "FrmLeakageBreakage", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.GatePass_Vs_actual, "RptGatePassVSActual", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptInvoiceAgainstInward, "RptInvoiceAgainstInward", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemPrice, "frmItemPrice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.StockRecoReport, "FrmShippingStockreport1", "ERP.EXE")


        UpdateFormName(clsUserMgtCode.FrmStockDispatchReport, "FrmStockDispatchReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnStockAdjustmentReport, "FrmStockAdjustmentReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAdjustmentStatusReport1, "FrmAdjustmentStatusReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnBreakageReport, "FrmBreakageReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BreakageReportSummary, "FrmBreakageReportSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RoutewiseBreakageReport, "FrmRoutewiseBreakageSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ReportTransfer, "ReportTransfer", "ERP.EXE")
        UpdateFormName("SamplingReportSummary", "FrmSamplingReportSummary1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SchemeReport, "FrmSchemeReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.StockReportForFinishedGoods, "FrmStockReportFinishedGoods", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAdjustmentReport, "FrmAdjustmentReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptVehicleWiseLoadout, "frmVehicleWiseTransfe", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingIndentTransferReport, "FrmPendingIndentTransferReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.stockRecoNew, "FrmStockReco", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.stockRecoBatch, "FrmStockRecoBatch", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmItemSerialTrackingReport, "FrmItemSerialTrackingReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTransferRegister, "FrmTransferRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmFatSnfStockReport, "FrmFatSNFStock", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDatewiseQtyFatSnfStockReport, "frmMCCMilkLossGain", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCMilkLossGain, "FrmDatewiseMilkStock", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Sampling_Master, "Sampling_Master", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.cash_Register_Details4, "Cash_Register_Details4", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQuickSettlement, "FrmQuickSettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSettlementMaster, "FrmSettlementMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.groupMasterRoute, "frmRouteGroupMaster", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.routeMaster, "frmRouteMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.transportMaster, "frmTransportMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.transportMasterVendor, "frmTransportMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.vhicleMaster, "frmVehicleMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.channelCategory, "frmchannelCategory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.channelMaster, "frmChannelMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.visiMaster, "frmvisimaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCommissionMaster, "FrmCommissionMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAbateMentMaster, "FrmAbateMentMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.transportType, "FrmTransportType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.routetypemaster, "FrmRouteTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesManHierarchy, "FrmSalesManHierarchy", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnBreakageHead1, "FrmBreakagehead", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCashDiscountReport, "FrmCashDiscountReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerVendorMapping, "FrmCustomerVendorMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCheckPrinting, "frmPrintCheckMultiple", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmNEFTUploader, "FrmNEFTUploader", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerVendorMappingVendor, "FrmCustomerVendorMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TDMTARGET, "TDMwiseTarget", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRouteShifting, "FrmRoute_Shifting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDiscountMaster, "FrmDiscountMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDiscountCategoryMaster, "FrmDiscountCategoryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnTargetMaster, "FrmTargetMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnTmplateCreation, "FrmTemplateCreation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmViewPunchingInvoice, "FrmViewPunchingInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDayWiseLoadOutEntered, "FrmDayWiseLoadOutEntered", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerTargetFixing, "frmCustomerTargetFixing", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmClaimMaster, "frmClaimMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNShipmentImportExport, "frmShipmentImportExport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRemarMaster, "frmRemarkMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmZoneMaster, "FrmZoneMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.customerItemDetails, "FrmCustomerItemDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBookingEntry, "FrmBookingEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmShortCloseDO, "FrmShortCloseDO", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDispatchFreshSale, "frmDispatchNoteFreshSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDispatchMultipleFreshSale, "frmDispatchMultipleFreshSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCreditLimitApproval, "FrmCreditLimitApprovalMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFreshCreditLimitApproval, "FrmCreditLimitApprovalMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesLevelHierarchy, "FrmSalesLevelhierarchy", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesHierarchy, "FrmsalesHierarchy", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesHierarchyMapping, "FrmSalesHierarchyMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkCreditLimitApproval, "FrmCreditLimitApprovalMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmInvoiceFreshSale, "frmInvoiceFreshSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCreateReceived, "frmCreateReceived", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmGateEntrySale, "FrmGateEntrySale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmWeighmentEntry, "FrmWeighmentEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmLoadingTanker, "FrmLoadingTanker", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmQualityCheckBulkSale, "FrmQualityCheckBulkSale", "ERP.EXE")
        ' UpdateFormName(clsUserMgtCode.frmTranReverse, "frmTransactionReverse", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDispatchBulkSale, "FrmDispatchBulkSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmInvoiceBulkSale, "FrmInvoiceBulkSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCreateAutoInvoiceBS, "FrmCreateAutoInvoiceBS", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkSaleReturn, "FrmBulkSaleReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkDispatchReturnSale, "FrmBulkDispatchReturnSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmTankerOut, "FrmTankerOut", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.FrmDispatchBulkSaleTrade, "FrmDispatchBulkSaleTrade", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFixedDeposit, "FrmFixedDeposit", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProformaInvoiceMT, "frmEXPorformaInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPurchaseOrderMT, "frmPurchaseOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesOrderMT, "frmEXSalesOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCommercialInvoiceMT, "frmEXCommercialInvoice", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.rptUnpostedPO, "frmRptUnpostedPO", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSRNMT, "frmSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesInvoiceMT, "frmEXSalesInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesReturnMT, "frmEXSalesReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmLCRequest, "FrmLCRequest", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmLCCreation, "FrmLCCreation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDocumentAcceptance, "FrmDocumentAcceptance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptCashAgainstDocs, "RptCashAgainstDocs", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptMTPIInHead, "RptMTPIInHead", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptEXProductWiseDetail, "RptMTProductWiseDetailReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptdateOfArrivalofCons, "RptDateOfArrivalOfCon", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkSalePriceChart, "FrmBulkSalePriceChart", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPrintBulkInvoiceStatement, "FrmPrintBulkInvoiceStatement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBulkSaleRegister, "RptBulkSaleRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEnquiryMaster, "FrmEnquiryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCHAChargeMaster, "FrmCHAChargeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmExIncentiveMaster, "FrmExIncentiveMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmNotifiedPartyMaster, "FrmNotifiedPartyMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEXSalesQuotation, "FrmEXSalesQuotation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEXSalesOrder, "frmEXSalesOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEXPorformaInvoice, "frmEXPorformaInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEXCommercialInvoice, "frmEXCommercialInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEXSalesInvoice, "frmEXSalesInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEXSalesReturn, "frmEXSalesReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptExportSaleRegister, "RptExportSaleRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSAPriceMaster, "FrmCSAPriceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSAAccountSet, "FrmCSAAccountSet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSABooking, "frmCSABooking", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSADeliveryOrder, "FrmCSADeliveryOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSATransfer, "frmCSATransfer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCSATransferGateOut, "FrmCSATransferGateOut", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.frmCSASaleInvoice, "FrmCSASaleInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSASalePattiReturn, "frmCSASalePattiReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSATransferReturn, "FrmCSATransferReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSATransferReport, "FrmCSATransferReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCSADOReport, "FrmCSADOReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptCSASaleRegister, "RptCSASaleRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptCSACustomerLedger, "frmRptCSACustomerLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkSaleSettings, "FrmBulkSaleSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeliveryNoteFreshSale, "frmDeliveryNoteFreshSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBookingProductSale, "frmBookingProductSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSchemeMasterDairy, "FrmSchemeMasterDairy", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRouteMaster, "frmRouteMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesOrderProductSale, "frmDeliveryOrderProductSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeliveryPrderProductSale, "frmSaleOrderProductSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmShipmentProductSale, "frmShipmentProductSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProductDispatchGateOut, "FrmProductDispatchGateOut", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleInvoiceProductSale, "frmSaleInvoiceProductSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleReturnProductSale, "frmSaleReturnProductSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptsaleRegisterReport, "RptSaleRegisterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSalesHierarchyReport, "rptSalesHierarchyReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmGatePassFS, "FrmGatePassFS", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleReturnFreshSale, "frmSaleReturnFreshSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRouteFreightDetails, "FrmRouteFreightDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPendingQuotationApproval, "FrmPendingQuotationApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmApprovalSetting, "FrmApprovalSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalesmanTarget, "FrmSalesmanTarget", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.customerItemDetails, "FrmCustomerItemDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerItemDetailApproval, "FrmCustomerItemDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.customerItemMapping, "frmCustomerItemMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmItemPriceListLevel3, "FrmItemPriceListLevel3", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frm_User_Customer_Rate_Settings, "Frm_User_Customer_Rate_Settings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.routeMaster, "frmRouteMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.routetypemaster, "FrmRouteTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleQuotation, "frmSNSalesQuotation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNSalesOrder, "frmSNSalesOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNShipment, "frmSNShipment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNSaleInvoice, "frmSNSaleInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNServiceInvoice, "frmSNServiceInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNSaleReturn, "frmSNSaleReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAutoSTN, "FrmAutoSTN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProspect, "frmProspectDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSetting, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSettingFresh, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSettingBulk, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSettingMerchant, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMerchantPaymentTerms, "FrmMerchantPaymentTerms", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMerchantPaymentTermsGroup, "FrmMerchantPaymentTermsGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMTReportContextFormat, "FrmMTReportContextFormat", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSettingExport, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSettingCSA, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleSettingProduct, "FrmSaleSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNPOS, "frmSNPOS", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSNReceiptChallan, "frmReceiptChallan", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSalesmanTarget, "frmRptSalesmanTarge", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleOrderDetail, "frmSaleOrderDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleOrderSummary, "FrmSaleOrderSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleInvoiceSummary, "FrmSaleInvoiceSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleInvoiceDetail, "FrmSaleInvoiceDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmShipmentDetail, "FrmShipmentDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmShipmentSummary, "FrmShipmentSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleRegisterDemo, "RptSaleRegisterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPendingSaleInvoiceforChilpPO, "frmPendingSaleInvoiceforChilpPO", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ORDNEW, "FrmOrdertracking", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmReceiptChallanReport, "frmReceiptChallanReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProspectDetailReport, "frmProspectDetailReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.saleOrders, "frmSaleOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCheckSlipEntry, "FrmCheckSlipEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LoadOut, "frmShipmentInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.saleReturn, "frmSalesReturnNew", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.SaleReturnInterCompany, "frmSaleReturnInter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCompleteTransfer, "FrmCompleteTransfer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LoadOutStatus, "FrmCompleteLoadout", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ScrapSale, "frmScrapSale", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.JobWorkDispatch, "frmJobWorkDispatch", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmScrapSaleGateOut, "FrmScrapSaleGateOut", "ERP.EXE")

        UpdateFormName("ScrapInvoice", "frmScrapInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSettlementEntry, "FrmSettlementEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmTransferIncompleteRemarks1, "FrmTransferIncompleteRemarks1", "ERP.EXE")
        'Dim frmP As New FrmPWD(Nothing)
        'frmP.strCode = "TempProvisional"
        'frmP.strType = "TempProvisional"
        'frmP.ShowDialog()
        'If frmP.isPasswordCorrect Then
        '    Try
        '        UpdateFormName("FrmReverseEntry", "FrmReverseEntry", "ERP.EXE")
        '    Catch ex As Exception
        '        common.clsCommon.MyMessageBoxShow(ex.Message)
        '    End Try
        'End If
        UpdateFormName(clsUserMgtCode.ChangeInvoiceSalesman, "FrmChangeInvoiceSalesman", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleHistory, "FrmSaleHistory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmExpiryDate, "FrmExpiryDate", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmTransactionApproval, "FrmTransactionApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFreshTransactionApproval, "FrmTransactionApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkTransactionApproval, "FrmTransactionApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMCCFarmerMapping, "FrmMCCFarmerMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VLCMappingForMPAmount, "frmVLCMappingForMPAmount", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FreightChargesMaster, "frmFreightChargesMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMCCTransactionApproval, "FrmTransactionApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPrintFreshInvoice, "FrmPrintFreshInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptFreshSaleRegister1, "RptFreshSaleRegister1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptCrateAccounting, "rptCrateAccounting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptCrateAccountingReport, "RptCrateAccountingReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptVehicleCapacityFreshSaleReport, "RptVehicleCapacityFreshSale", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.vehicle_Details_Report1, "Vehicle_Details_Report", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptTransfer_IncompleteReport, "RptTransfer_IncompleteReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.reportQuickSettlement, "FrmReportForQuickSettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.visiDetail1, "VisiDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PendingSaleOrderReport, "PendingSaleOrderReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCDiscountReport, "FrmMCDiscReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.crptLoadOut, "FrmLoadOutRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.receiptFillreport, "FrmRECEIPTSAGAINSTSALES_FILLED_", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.receiptWOTreport, "FrmInvoiceswithoutreceipt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.receiptreport, "Frmreceiptvoucher2", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.nrptSales, "FrmRptSales", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnRptSalesManSalesReport, "frmRptSalesManReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Settlement, "FrmSettlementReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CustomerRouteHistoryReport, "FrmRptCustomerRouteHistory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemDiscountReport, "FrmDiscountReport", "ERP.EXE")
        UpdateFormName("ItemCommissionReport", "FrmItemCommissionReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LoadOutStatusreport1, "FrmLoadOutStatusreport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LoadOutReport1, "LoadOut", "ERP.EXE")
        UpdateFormName("EmptyInwardRegisterSummary", "FrmEmptyInwardRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnNetSaleReport, "FrmNetSaleReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDistrbutorSaleTarget, "FrmDistrbutorSaleTarget", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDayReportDirectSale, "FrmDayReportDirectSale", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.ExciseSummary1, "FrmExciseSummaryNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.OtherPartySale, "FrmOtherPartySale1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTransitBreakageReport1, "FrmTransitBreakageReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptCreditSaleReport, "frmRptCreditSales", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.NoSaleReport, "frmRptNoSales", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptPenetration, "frmRptPenetration", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.FilloutwardRegisterReport1, "FrmFilledOutwardRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SaleReport, "FrmTDMSaleReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PrimarySales, "FrmPrimarySalesReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SecondarySales, "FrmSecondarySales", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.EmptyInwardSaleRegister1, "frmInwardRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ProvSaleDetail, "FrmProvisionalSalesRoutewise", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmGatePassENtry1, "FrmGatePassENtry1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SaleAccountBreakDetail, "FrmSaleAccountBreakOrCashDisc", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SaleAccountBreakage, "FrmSaleAccountBreakageReport", "ERP.EXE")
        UpdateFormName("DailyStockAccount", "FrmDailyStockAccountRpt", "ERP.EXE")
        UpdateFormName("OverAllDisc", "FrmOverallDiscountReport", "ERP.EXE")
        If objCommonVar.IsDemoERP Then
            UpdateFormName(clsUserMgtCode.CrptRG1Detail1, "frmRG1Demo", "ERP.EXE")
        Else
            UpdateFormName(clsUserMgtCode.CrptRG1Detail1, "frmRG1", "ERP.EXE")
        End If
        If objCommonVar.IsDemoERP Then
            UpdateFormName(clsUserMgtCode.frmCFormReport, "FrmCFormReport", "ERP.EXE")
        End If
        UpdateFormName(clsUserMgtCode.frmOpeningBalance, "FrmOpenningBalance", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.RptServiceTaxDetail, "RptServiceTaxDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBankWiseChequeIssue, "RptBankWiseChequeIssue", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ProvisionalSaleReport, "FrmProvionalSalesReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ItemCommissionSummary, "FrmItemCommissionSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DistributedDiscountReport, "FrmDistribuorDiscount", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleDiscount1, "FrmSaleDiscount1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CEAllocationReport, "FrmCEAllocationRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SalesCollection, "FrmSalesCollection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DailySettlement, "frmDailySettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMismatchSettlement, "FrmMismatchSettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSettlementSheetReconcilationeport, "FrmSettlementSheetReconcilationReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.OutletEmpty1, "FrmOutletEmptyReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VehiclewiseSale1, "FrmVehiclewiseSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Channelwisecustomer1, "FrmChannelwiseCustomer1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCustomerRanking, "FrmCustomerRankingReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnVisiVPO1, "FrmVisiVPOReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnMismatchReport, "FrmMismatchRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnRouteSale, "RouteSaleReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerTargetReport, "frmCustomerTargetReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTDMReport, "frmTDMReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SalesmanSalesOrderReport, "FrmSalemanSaleOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDealerManagementReport, "FrmDealerManagementReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSalesAnalysis, "frmSalesAnalysisReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRptFormOfGuarntee, "frmRptFormOfGuarntee", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptCustomerWiseMonthlySalesAnalysis, "frmCustomerWiseMonthlySalesAnalysis", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPaySlipReport, "frmCheckDepositPaySlip", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSalarySlipRpt, "FrmSalarySlipRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSalarySummaryRpt, "FrmSalarySummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmEmployeePFRpt, "FrmEmployeePF", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmESICRpt, "FrmESICRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptForm34, "RptForm34", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBonusStatement, "RptBonusStatement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBOILetterReport, "RptBOILetterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptForm22, "RptForm22", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptESIHalfYearly, "RptESIHalfYearly", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFForm3A, "RptPFForm3A", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFForm5, "RptPFForm5", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFForm10, "RptPFForm10", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RPtPFForm11_Revised_, "RPtPFForm11_Revised_", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFForm12A_revised_, "RptPFForm12A_revised_", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFChallanStatement, "RptPFChallanStatement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptKDILSalarySlip, "RptKDILSalarySlip", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPerformaForContributiondetail, "RptPayrollPerformaforcontribution", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPaymentMode, "FrmPaymentMode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBankTransferDetail, "RptBankTransferDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDepartmentwiseSalarySheetRpt, "RptDepartmentWiseSalarySheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptDetailOfWelfareFundAmount, "RptDetailOfWelfareFundAmount", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptESICStatement, "RptESICStatement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFStatement, "RptPFStatement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFForm6, "RptPFForm6", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptESICForm6, "RptESICForm6", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptESICChallan, "RptESICChallan", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptESICDeclarationForm, "RptESICDeclarationForm", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptActurialValuation, "RptActurialValuation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPFForm2, "RptPFForm2", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDVAT32, "frmRptDVAT32", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerBillWiseDetail, "frmCustomerBillWiseDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleSummaryAgainstPO, "frmSaleSummaryAgainstPO", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmYearMonthWiseSaleComparison, "frmYearMonthWiseSaleComparison", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCompanyMonthWiseSaleComparison, "frmCompanyMonthWiseSaleComparison", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptDistributerPerformance, "rptDistributerPerformance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerBillWiseDuesSummary, "frmCustomerBillWiseDuesSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorGroupWiseSaleReport, "frmVendorGroupWiseSaleReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBatchReceiptSTD, "FrmBatchReceipt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBatchReceiptPepsi, "FrmBatchReceipt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFilledOutWard, "FrmFilledOutWard", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CashDiscount, "rptCashDiscountReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.CashDiscountReport, "frmCashDiscountNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TransferRegister, "TransferRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.EmptyReportDetail, "EmptyReportDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCarteJaliRpt, "FrmCrateJaliReport", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.RptPendingSettlement, "RptPendingSettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDetailOfForm2A, "FrmDetailOfForm2A", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmTargetReport1, "FrmTargetReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDailySettlementActualAndProvisionalReport, "FrmDailySettlementActualAndProvisionalReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSettlement_CashMemoStatus, "FrmSettlement_CashMemoStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmReverseSettlementDetail, "FrmReverseSettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMismatchCashMemo, "FrmMismatchCashMemo", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCanceledSaleInvoice, "FrmCanceledSaleInvoice1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSaleVolumeTracker, "FrmSaleVolumeTracker", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmEmptyTransactionReport, "FrmRptEmptyTransaction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSaleOrderSummary, "FrmSaleOrderSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDiscountAnalysis, "FrmDiscountAnalysis", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmpendingLoadin, "FrmPendingLoadIn_Transfer_Type", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRptSalesReturn, "FrmRptSalesReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmQuickSettlementHead, "FrmQuickSettlementHead", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LO_vs_Vechile, "FrmloadoutVSvechileCapacity2", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptClaimMaster, "rptClaimMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVehiclePendingStatusRpt, "FrmVehiclePendingStatusRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.NatureOfDeduction, "frmNatureOfDeduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PartyDetails, "frmPartyDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TDSSection, "frmTDSSection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BranchDetails, "frmBranchDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FinancialYear, "frmFinancialYear", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.StateCode, "frmStateCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ResponsiblePerson, "frmResponsiblePerson", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCreateRemittance, "FrmCreateRemittance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.remittanceentry, "Frmremittanceentry", "ERP.EXE")
        ' Pravet Service
        UpdateFormName(clsUserMgtCode.frmCattleType, "FrmCattleTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBredType, "FrmBredTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCattleColor, "FrmCattleColorMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCattleMaster, "FrmCattleMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceGroup, "FrmServiceGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceMaster, "FrmServiceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmNDDBMaster, "FrmNDDBMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBullMaster, "FrmBullMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFarmerServiceOrderWithRate, "FrmFarmerServiceOrderWithRate", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSMSDetails, "frmFarmerServiceOrder", "ERP.EXE")
        ' End Pravet Service
        UpdateFormName(clsUserMgtCode.mbtnAPInvoiceEntryTDS, "FrmAPInvoiceEntryTDS", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmrptTDSLedger, "FrmrptTDSLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TDSForm26Q, "form26Q27Q", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TDSSectionSummaryReport, "FrmTDSsectionSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Form16AReport, "FrmForm16A", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPurchaseSetting, "frmPurchaseSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptPurchaseRegisterReport, "RptPurchaseRegisterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorItemDetails, "frmVendorItemDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnPurchaseRequistion, "frmPurchaseRequistion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStoreRequistion, "frmStoreRequistion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnPendingApprovalOfReq, "FrmPendingReqForApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RequisitSubTypeMaster, "FrmRequisitSubTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RFQ, "FrmRFQ", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorQuotation, "frmVendorQuotation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorComparison, "FrmVendorComparison1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.VendorComparisonApproval, "frmVendorComparisonApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnPurchaseOrder, "frmPurchaseOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPurchaseSchedule, "FrmPurchaseSchedule", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnMRN, "frmMRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnSRN, "frmSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SRNReturn, "frmSRNReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnPurchaseInvoice, "frmPurchaseInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnPurchaseReturn, "frmPurchaseReturn", "ERP.EXE")
        'KUNAL > CLIENT : UDIL > TICKET : BM00000010226 
        UpdateFormName(clsUserMgtCode.mbtnNRGP, "frmNRGPBooking", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnGatePass, "frmRGP", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnIssueReturn, "frmIssueReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRoutewiseSaleReport, "FrmRoutewiseSaleReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmKPIReport, "FrmKPIReport", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.FrmProvSaleExcel, "FrmProvSaleExcel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RM_Consumption_Detail, "RM_Consumption_Detail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMaterialSendforJobWork, "RptMaterialSendForJobWork", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMaterialReceivedAfterJobWork, "RptMaterialReceivedAfterJobWork", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptBalanceStockForJobWork, "RptBalanceStockForJobWork", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptRGPWiseJobWork, "RptRGPWiseJobWork", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmIndentReport, "FrmIndentReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmExpiredItemDetails, "FrmExpiredItemDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmConsumptionReport1, "FrmConsumptionReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DebitAdviseReport, "FrmDebitAdviseReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Vendor_Rating_Rejection, "Vendor_Rating_Rejection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Store_Receipt_Note, "Store_Receipt_Note", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnDailyRcptNoteSummary, "FrmDailyReceipNoteSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Parti_VS_Rejected, "Parti_VS_Rejected", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPurchaseOrderReport, "FrmPurchaseOrderReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPo_action, "frmPo_action", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingRequisitionQty, "FrmPendingRequisitionQty", "ERP.EXE")
        UpdateFormName("FrmPendingPO_Qty", "FrmPendingPO_Qty", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingGrn_Qty, "FrmPendingGrn_Qty", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingMrn_Qty, "FrmPendingMrn_Qty", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingSrn_Qty, "FrmPendingSrn_Qty", "ERP.EXE")
        UpdateFormName("FrmPendingInvoice_Qty", "FrmPendingInvoice_Qty", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MRDAReport, "FrmMRDAReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmIssueOrReturnItemWiseSummary, "FrmIssueOrReturnItemWiseSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPurchasebookReport1, "FrmPurchasebookReport1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStockAnalysis, "FrmStockAnalysis1", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.frmMorningReport, "FrmMorningReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AddCharge, "FrmAdditionalCharge1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSrnReport, "FrmSrnReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PJVReport, "FrmPJVReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.StockStatement, "FrmStockStatementReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPurchaseOrderRegister, "FrmPurchaseOrderRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmItemWiseDispatchLedger3, "FrmItemWiseDispatchLedger3", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnRGP_NRGP_Rpt, "frmRGP_NRGP_Rpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DetailofWtdPriceofRawMaterial, "FrmWTDRpt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorWiseReturnableGoodBalance, "FrmVendorWiseReturnableGoodBalance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnStoreLedger, "FrmStoresLedgerNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnStoresLedger, "FrmStoresLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRGP_Register_NRGP, "FrmRGP_Register_NRGP", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFreightCosting, "FrmFreightCosting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRptPurchaseReturnBook, "FrmRptPurchaseReturnBook", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPurchaseOrderList, "frmPurchaseOrderList", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnGRN, "frmGRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.POWeighment, "frmPOWeighment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.POUnloading, "frmPOUnloading", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnGRNReport, "FrmGRNReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AssetSegment, "FrmAssetSegment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Template, "FrmTemplateMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetBookMaster, "frmAssetBookMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DepAccSets, "FrmDepAccountSet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Categories, "FrmCategories", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDepreciationField, "FrmDepreciationField", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetGroups, "FrmGroups", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDepreciationMethod, "frmDepreciationMethod", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DepPeriod, "FrmDepreciationPeriods", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAMAcquisitionCode, "FrmAMAcquisitionCode", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSecondaryCustomer, "FrmSecondaryCustomer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FAAcquisitionEntry, "frmAcquisionEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FAAssetDepreciation, "FrmAssetDepreciation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FADisposalEntry, "frmAssetScrapSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVisi_Install_Pullout, "FrmVisi_Install_Pullout", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAsset_Issue_Return, "FrmAsset_Issue_Return", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetRequisition, "frmAssetRequisition", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSecondaryCustomerSale, "FrmSecondaryCustomerSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FAMergeAcquisitionEntry, "frmFAMergeAsset", "ERP.EXE")

        'KUNAL > TICKET : BM00000010298 > CLIENT : UDL > DATE : 28-NOV-2016
        UpdateFormName(clsUserMgtCode.FrmRpt_OutStnd_Items_RGP, "FrmRpt_OutStnd_Items_RGP", "ERP.EXE")

        '======================================================
        UpdateFormName(clsUserMgtCode.frmIssueItemsToAsset, "frmItemIssueToAssembledAsset", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FAAssetWork, "frmAssetWork", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAssetRegister, "FrmAssetRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAssetDetail, "FrmAssetDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDisposalDetail, "FrmDisposalDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVisi_Install_Pullout_Report, "FrmVisi_Install_Pullout_Report", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAsset_Issue_Return_Report, "FrmAsset_Issue_Return_Report", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDistributor_VS_SecondaryCustomer_Sale, "FrmDistributor_VS_SecondaryCustomer_Sale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCreateReceiptAgainstSale, "FrmCreateReceiptAgainstSales", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnCreateReceiptAgainstInvoice, "FrmCreateReceiptAgainstInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnTakeBackup, "FrmBackup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.mbtnPendingApproval1, "FrmPendingAproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBulkPostingNew, "FrmBulkPostingNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.lockTransaction, "FrmLockTransaction1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmUserPerformanceDetail, "FrmUserPerformanceDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStockReport, "FrmStockReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmReconciliationSetting, "FrmReconciliationSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmUtilityForm, "FrmUtility", "ERP.EXE")
        ' Paravet Services
        UpdateFormName(clsUserMgtCode.frmCattleType, "FrmCattleTypeMaster", "ERP.EXE")
        'frm = New FrmPWD(Nothing)
        'frm.strCode = "TempProvisional"
        'frm.strType = "TempProvisional"
        'frm.ShowDialog()
        'If frm.isPasswordCorrect Then
        '    Dim frmNew As New FrmUtility()
        '    formShow(frmNew, strProgramName, True, "")
        '    'Try
        '    '    clsSaleHead.SetTempProvisionSale()
        '    'Catch ex As Exception
        '    '    common.clsCommon.MyMessageBoxShow(ex.Message, Me.Text)
        '    'End Try
        'End If


        UpdateFormName(clsUserMgtCode.frmSalaryAccountSetting, "frmSalaryAccountSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDepartmentMaster, "frmDepartmentMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSubDepartmentMaster, "frmSubDepartmentMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmEmployeeTransfer, "FrmEmployeeTransfer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSkillMaster, "frmSkillMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLanguageMaster, "frmLanguageMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCourseMaster, "frmCourseMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmShiftMaster, "frmShiftMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDocumentMaster, "frmDocumentMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPayPeriodMaster, "frmPayPeriodMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCountryMaster, "frmCountryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmReligionMaster, "frmReligionMaster", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmCurrencyMaster, "frmCurrencyMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCastCategoryMaster, "frmCastCategoryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGradeMaster, "frmGradeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStateMaster, "frmStateMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDevisionMaster, "frmDevisionMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOccupationMaster, "frmOccupationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPFRulesMaster, "frmPFRulesMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmESIRulesMaster, "frmESIRulesMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOTMaster, "frmOTMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAttendanceMaster, "frmAttendanceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBranchMaster, "frmBranchMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBonusMaster, "frmBonusMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOTSlab, "frmOTSlab", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPTSlab, "frmPTSlab", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmConveyanceRateMaster, "frmConveyanceRateMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmODMaster, "frmODMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.DesignationMaster, "frmDesignationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOTSheet, "frmOTSheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveMaster, "frmLeaveMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveSetting, "frmLeaveSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGeneralHolidays, "frmGeneralHolidays", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveOpeningBalance, "frmLeaveOpeningBalance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveStartingDateSetting, "frmLeaveStartingDateSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveAllotment, "frmLeaveAllotment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPayHeadDefinitions, "frmPayHeadDefinitions", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryStructure, "frmSalaryStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMapPayHeadsToSalaStructure, "frmMapPayHeadsToSalaStructure", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmWeeklyHolidays, "frmWeeklyHolidays", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveApplication, "frmLeaveApplication", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMonthlyAttendance, "frmMonthlyAttendance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveAdjustment, "frmLeaveAdjustment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGenerateBonus, "frmGenerateBonus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDailyAttendance, "frmDailyAttendance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHourlyAttendance, "frmHourlyAttendance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAdjustmentVoucher, "frmAdjustmentVoucher", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TOOLTYPE, "FrmToolType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmReimbursementDetails, "frmReimbursementDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAllowanceDetails, "frmAllowanceDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeductionDetails, "frmDeductionDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmApplyLoan, "frmApplyLoan", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmployee_Master, "frmEmployee_Master", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmpSalary, "frmEmployee_Salary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmployeeStatus, "frmEmployee_Status", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLoanAdjustment, "frmLoanAdjustment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLoanGeneration, "frmLoanGeneration", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryGeneration, "frmSalaryGeneration", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmployeeGratuity, "FrmEmployeeGratuity", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLTAClaim, "frmLTAClaim", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMediclaimEntry, "FrmMediclaimEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmFullAndFinalSettlement, "frmEmpFullAndFinalSettlement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmployeeShiftChange, "frmEmployeeShiftChange", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmODSheet, "frmODSheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmConveyanceClaim, "frmConveyanceClaim", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPayrollSetting, "frmPayrollSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmEmpIncrement, "FrmEmpIncrement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSentSalarySlip, "FrmSentSalarySlip", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAllotmentOfLeaves, "FrmAllotmentOfLeaves", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryGenerationRegister, "frmSalaryGenerationRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryGenerationRegisterArrear, "frmSalaryGenerationRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAllownceRegister, "frmAllownceRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeductionRegister, "frmDeductionRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmReimbursementRegister, "frmReimbursementRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAdjustmentRegister, "frmAdjustmentRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAttendanceRegister, "frmAttendanceRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLeaveRegisterReport, "frmLeaveRegisterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmployeeRegister, "frmEmployeeRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPF_ESI_Reports, "frmPF_ESI_Reports", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptEmployeeBday6, "RptEmployeeBday6", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptEmployeeAdvanceLedger, "rptEmployeeAdvanceLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPaySlip_Reports, "frmPaySlip_Reports", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalarySheet_Reports, "frmSalaryGenerationRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryAbstractReport, "frmSalaryAbstractReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAnnualCensusReport, "frmAnnualCensusReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAttendedDaysReport, "frmAttendedDaysReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryVoucher_Reports, "frmSalaryVoucher_Reports", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBankStatement_Reports, "frmBankStatement_Reports", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOT_Reports, "frmOT_Reports", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVarianceReport, "frmVarianceReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryComponentDetails, "frmSalaryComponentDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAditionalEarning_DeductionReport, "frmAditionalEarning_DeductionReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryCertificate, "frmAditionalEarning_DeductionReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPF_Covering_Letter, "frmPF_Covering_Letter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryIncrementReport, "frmSalaryIncrementReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBankSummary_Report, "frmBankSummary_Report", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeductionDetailsReport, "frmDeductionDetailsReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVoucherPaymentsRegister, "frmVoucherPaymentsRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmEmp_Id, "frmEmp_Id", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLabelPrinting, "frmLabelPrinting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmForm_T, "frmForm_T", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ACCSETMFGSTD, "FrmAccountSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ACCSETMFGDairy, "FrmAccountSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ACCSETMFGPepsi, "FrmAccountSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ITEMCATEGORY, "FrmItemProductionCategory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TOOL, "FrmToolMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SETTSTD, "FrmSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.SETTPep, "FrmSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.EXPENSE, "FrmExpenseHead", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmWorkCenterMaster, "frmWorkCenterMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmResourceMaster, "frmResourceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ALTER, "FrmAlternateItem", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBillOfMaterialPepsi, "frmBillOfMaterial", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBillOfMaterialDairy, "frmBOM", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProcessProductionLogSheet, "FrmProcessProductionLogSheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPPLogSheetMaster, "frmPPLogSheetMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPPLogSheetMaster_QC, "frmPPLogSheetMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorItemQCMapping, "FrmVendorItemQCMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQualityCheckForSRN, "FrmQualityCheckForSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQualityCheckApprovalForSRN, "FrmQualityCheckApprovalForSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOperationMaster, "frmOperationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBOMImport, "frmBOMImport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.COSTMAINTAIN, "FrmCostMaintainance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PRO, "FrmProductionLines", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProcessMaster1, "FrmProcessMaster1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSectionMaster, "FrmSectionMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmStageMaster, "frmStageMasters", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSectionStageMapping, "FrmSectionStageMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBillOfMaterialCosting, "frmBillOfMaterialCosting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionPlanningSTD, "frmProductionPlanning", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMRPAutoMobile, "frmMRPAutoMobile", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionPlanningPepsi, "FrmProcessProductionPlanning", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionPlanningDairy, "FrmProcessProductionPlanning", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProcessProductionIssueEntry, "FrmProcessProductionIssueEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProcessProductionStandardization, "frmProcessProductionStandardization", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ProcessProductionStandardizationFinalQC, "frmProcessProductionStandardizationFinalQC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProcessProductionStageProcess, "frmProcessProductionStageProcess", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionEntry, "frmProductionEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionEntryFinalQC, "frmProductionEntryFinalQC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionEntryWithoutBatch, "frmProductionEntryWithoutBatch", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProcessProdReturn, "frmProcessProductionReturn", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmMRP, "frmMRP", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDemoProductionPlanning, "frmProductionPlanningDemo", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionRequisition, "frmProductionRequisition", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRiceBOM, "frmRiceBOM", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRiceMixingEntry, "FrmRiceMixingEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRiceProcessingEntry, "FrmRiceProcessingEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionItemSerialReplace, "frmProductionItemSerialReplace", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionSerializedReport, "frmProductionSerializedReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionReceiptDemo, "FrmProductionReceiptDemo", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmManufacturingOrder, "frmManufacturingOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptproductionEntryReport, "RptproductionEntryReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptItemConsumptionReport, "RptItemConsumptionReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptProductionIssueStatusReport, "RptProductionIssueStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptWTPReport, "RptWIPReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptBatchStatus, "RptBatchStatusReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSectionWiseStockReport, "RptSectionWiseStockReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBatchOrderSTD, "frmBatchOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBatchOrderPepsi, "frmBatchOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBatchOrderDairy, "FrmProcessBatchOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStoreIssueSTD, "frmStoreIssue", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmStoreIssuePepsi, "frmStoreIssue", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionReturnSTD, "frmProductionReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionReturnPep, "frmProductionReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionReceiptSTD, "frmProductionReceipt", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLabourWorkingSheet, "FrmLabourWorkingSheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Resource, "RptListOf_Resource", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LTOOL, "RptListOfToolType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LACCt, "RptListOfAccountSet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LALT, "RptListOfAlternateItem", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LWC, "RptListOfWorkCenter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LOIC, "rptListOfItemCost", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LOPER, "RptListOfOperations", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.PRODREPORT, "FrmListOfProductionLines", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LToolT, "RptListOfTools", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmListOfBOM, "frmListOfBOM", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmIssueReturnItemWiseReportSTD, "frmIssueReturnItemWiseReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmIssueReturnItemWiseReportPepsi, "frmIssueReturnItemWiseReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDatewiseProduction, "frmDatewiseProduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionPlanReportSTD, "frmProductionPlanReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionPlanReportPepsi, "frmProductionPlanReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmLineProductivity, "frmLineProductivity", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBatchOrderReportSTD, "frmBatchOrderReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBatchOrderReportPepsi, "frmBatchOrderReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGraphicalBatchOrder, "frmGraphicalBatchOrder", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGraphicalCategorywiseProduction, "frmGraphicalCategorywiseProduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmListofRequisition, "frmListofRequisition", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOperaterEfficiencyReport, "FrmOperatorEfficiencyReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BICreateReport, "frmCreateBIReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BICreateFilter, "frmCreateBIFilter", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BICreateDashBoard, "FrmCreateDashBoard", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BIMonthWiseSale, "FrmBIMonthWiseSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BITopCustomer, "frmBITopCustomer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BITopItemCategory, "frmBITopItemCategory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionVarianceSTD, "frmProductionVariance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionVariancePepsi, "frmProductionVariance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BIMonthWisePurchase, "FrmBIMonthWisePurchase", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BITopVendor, "frmBITopVendor", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BITopItemCategoryPurchase, "frmBITopItemCategoryPurchase", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BIMonthWiseAssset, "frmBIMonthWiseAsset", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BITopExpence, "frmBITopExpence", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BIDashBoadr, "frmBIDashBoard", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPJCSettings, "frmPJCSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCostTypes, "frmCostType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPJCAccountSets, "frmPJCAccountSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmJobMaster, "frmJobMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTaskMaster, "frmTaskMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPJCEmployeeMaster, "frmPJCEmployeeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmUserApproval, "FrmUserApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBudgetMaintenance, "FrmBudgetMaintenance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ProjectMaster, "FrmProjectMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmExpenseType, "FrmExpenseType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTimeSheet, "frmTimesheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmUserLog, "FrmUserLog", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProjectStatus, "FrmProjectStatus", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssemblies, "frmAssemblies", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPJCExpense, "FrmPJCExpense", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProjectListReport, "FrmProjectListReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProjectDetails, "FrmProjectDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProjectProfitReport, "FrmProjectProfitReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProfitAndLossPerforma, "frmProfitAndLossPerforma", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptProfitAndLoss, "frmRptProfitAndLoss", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmComplaintGroupMaster, "frmComplaintGroupMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPrimaryReasonMaster, "FrmPrimaryReasonMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmComplaintMaster, "frmComplaintMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPendingReasonMaster, "frmPendingReasonMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmItemChargeCategoryMaster, "frmItemChargeCategoryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmItemChargeFranchiseMappingMaster, "FrmItemChargeFranchiseMappingMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetServiceMaster, "FrmAssetServiceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetAgreement, "FrmAssetAgreement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetInstallPullOut, "frmAssetInstallPullOut", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmComplaintDetailEntry, "FrmComplaintDetailEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQuickComplaintDetailEntry, "FrmQuickComplaintDetailEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetDistatch, "frmAssetDispatch", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCartMaintenanceEntry, "FrmCartMaintenanceEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingComplaintDetail, "FrmPendingComplaintDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmFranchiseChargesReport, "FrmFranchiseChargesReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomersListReport, "FrmCustomersListReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPullOutRedeployReport, "frmPullOutRedeployReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorBillDetails, "frmVendorBillDetails", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDaywisePendingComplaint, "FrmDaywisePendingComplaint", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmClaimReport, "FrmClaimReportNew", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAssetDetailReport, "FrmAssetDetailReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkCollectionLevels, "frmMilkCollectionLevelsMain", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkAdvanceMaster, "AdvanceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkVehicleTypeMaster, "VehicleTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkComponentMaster, "MilkComponentMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkRateTypeMaster, "MilkRateType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkShiftMaster, "ShiftMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSeasonMaster, "SeasonMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkRouteMaster, "FrmMilkRouteMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkVehicleMaster, "frmVehicleMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVillageMaster, "FrmVillageMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVSPMaster, "frmVSPMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPrimaryTransporterMaster, "FrmPrimaryTransporterMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTankerTransporterMaster, "frmTankerTransporterMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCMaster, "FrmMCCMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmTankerMaster, "FrmTankerMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVLCMaster, "FrmVLCMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVLCMasterTarget, "FrmVlcTargetMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeliveryTermsMaster, "frmDeliveryTermsMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPriceChartUploader, "FrmPriceChartUploader", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkPricePlanning, "frmPriceChartPlanMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPriceChartMaster, "FrmPriceChartMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPriceChartMaster_Bulk, "FrmPriceChartMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFormIssueReceiptEntry, "FrmFormIssueReceiptEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCostCentreGroupStores, "FrmCostCentreGroupStores", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVLCUploader, "FrmVLCUploader", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMPMaster, "FrmMPMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmParameterMaster, "FrmParameterMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmParameterRangeMaster, "FrmParameterRangeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVLCRouteShiftMaster, "FrmVLCRouteShiftMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPrimaryTransporterVehicalMaster, "FrmPrimaryTransporterVehicalMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOpenMCCShift, "FrmOpenMCCShift", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOpenMCCShiftManual, "FrmOpenMCCShift", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmParameterRangeMasterForQC, "frmParameterRangeMasterForQC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQualityModuleParameterRangeMaster, "frmParameterRangeMasterForQC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPortSettings, "FrmPortSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPriceChartBulkProc, "frmPriceChartBulkProc", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ParameterValueMaster, "FrmParameterValueMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LostDefectSealNo, "FrmLostDefectSealNo", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.LocationDistanceMapping, "FrmLocationDistanceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MccMilkTransferPrice, "FrmMccMilkTransferPrice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorPriceChartMapping, "frmVendorPriceChartMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmdeductionGroup, "FrmDeductionGroup1", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDeductionMaster, "FrmDeductionMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGroupOfDeduction, "FrmGroupOfDeduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkCollectionCenters, "FrmMilkCollectionCenters", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkReject, "frmMilkRejectEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkReceipt, "frmMilkReceiptMCC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCancelAfterPosting, "FrmCancelAfterPosting", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmMCCGateEntry, "frmMCCGateEntry", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmMCCWeighment, "frmMCCWeighment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMCCMilkTransPortorInvoice, "FrmRecurringPayableInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGateEntry, "FrmGateEntry", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.MCCDispatchReturn, "frmTankerDispatchReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCDispatch, "FrmMccDispatch", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmWeighment, "FrmWeighment", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.MilkGateEntryIn, "frmMilkGateEntryIn", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.MilkRetesting, "frmMilkRetesting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkGateEntryWeightment, "frmMilkGateEntryWeighment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkGateEntryOut, "frmMilkGateEntryOut", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQualityCheck, "FrmQualityCheck", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkSample, "frmMilkSampleMCC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCSMSSettiing, "FrmMccSMSSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCMaterial, "frmMCCMaterialSale", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCMaterialSaleReturn, "frmMccMaterialSaleReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMCCMaterialSalePriceChart, "FrmMCCMaterialSalePriceChart", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkSRN, "frmMilkSRNMCC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MCCSetting, "frmMCCProcurementSetting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVSPIncentiveTagging, "FrmVSPIncentiveTagging", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkPurchaseInvoice, "frmMilkPurchaseInvoiceMCC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkShiftUploader, "frmMilkShiftUploader", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkProcurementCorrection, "frmCorrection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkShiftEndMCC, "frmMilkShiftClosingMCC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkTransferIn, "FrmMilkTransferIn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBulkMilkSRN, "FrmBulkMilkSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBulkMilkSRNReturn, "FrmBulkMilkSRNReturn", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmUnloading, "FrmUnloading", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBulkMilkPurchaseInvoice, "FrmMilkPurchaseInvoice", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMilkPurchaseReturn, "FrmMilkPurchaseReturn", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.frmCleaning, "FrmCleaning", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmGateOut, "FrmGateOut", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmBulkPurchaseUploader, "frmBulkPurchaseUploader", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVlcdataUploadar, "FrmVlcDataUploadar", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVLCDataUploaderManual, "FrmVLCDataUploaderManual", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.BulkMilkPurchaseInvoiceMultiple, "frmBulkMilkPurchaseInvoiceMultiple", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMCCShiftReportRouteWise, "FrmMCCShiftReportRouteWise", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmDispatchTransfer, "FrmDispatchTransfer", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptCollectionLevelChart, "rptCollectionLevelChart", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptCollectionCenterChart, "rptCollectionCenterChart", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptvillageslip, "RptVillageSlip", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMilkBillMCC, "RptMilkBillMCC", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMilkBillRouteWise, "RptMilkBillRouteWise", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMCCMilkBillSummary, "RptMccMilkBillSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptTankerSummaryReport, "RptTankerSummaryReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptShifReportZeroAmtSample, "RptShiftReportZeroAmtSample", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMilkPaymentRegister, "RptPaymentRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptLowProcurement, "RptLowProcurement", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptMccSaleRegister, "RptMCCsaleRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBulkMilkRegister, "RptBulkMilkRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptMilkRouteVehicleReport, "RptMilkRouteVehicleReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptCDA, "RptCDA", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMonthlyVLCProcurement, "RptMonthlyVLCProcurement1", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.rptSecondaryQuality, "RptSecondaryQualityReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptDailyDifferentReport, "RptDailyDifferentRow_vb", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVSPAssetIssue, "frmVSPAssetIssue", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.frmVSPItemIssue, "frmVSPItemIssue", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMillPurchaseBill, "RptMilkPurchaseBill", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.RptGainSheetPeriod, "RptGainSheetPeriod", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptWeighment, "RptWeightment", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.RptTankerVariation, "RptTankerVariationReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptDailyGainDay, "RptDailyGainDay", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptPendingMilkSRN, "RptPendingMilkSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPendingProvisionReport, "FrmPendingProvisionReport", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.rptPTReport, "FrmPTReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MCCMilkRegister, "FrmMCCMilkRegister", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MccSummaryReport, "FrmMCCSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMilkStockLedgerSummary, "RptMilkStockLegderSummary", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMemberPaymentSlip, "Rptmemberpaymentslip3", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMPWiseMilkCollection, "RptMPWiseMilkCollection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptMPWiseMilkCollectionATPoolingPoint, "RptMPWiseMilkCollectionAtPoolingPoint3", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptVillageDiffReport, "RptVillageDiffReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptPrimaryTransporter, "RptPrimaryTransporter", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.rptMCCMilkRegisterDripSaver, "RptMCCMilkRegisterDripSaver", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptVSPOrVLCVarationRpt, "RptVSPOrVLCVarationReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptCollectionAnalysis, "RptCollectionAnalysis", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MCCProvisonReport, "frmRptMCCProvision", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.rptTankerDispatchWidthDeduction, "frmRptTankerDispatchWithDeduction", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptMPIDReport, "RptMPIDReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProvisionEntry, "FrmProvisionEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPaymentProcess, "FrmPaymentProcess", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.rptVSPItemIssue, "RptVSPItemIssue", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptDispatchOfMilkTransfer, "RptDispatchofmilkTransfer2", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.RptVLCTragetMasterReport, "RptVLVCTragetMasterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptBulkMilkMultiplePurchaseInvoice, "RptBulkMilkMultiplePurchaseInvoice", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.SublocationMaster, "frmSubLocationMaster", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.ItemLocationMapping, "frmTransferLocationMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBankPermission, "FrmBankPermission", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmCustomerPermission, "FrmCustomerPermission", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmVendorPermission, "FrmVendorPermission", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.JobTitle, "FrmJobTilte", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRParameterMaster, "FrmHRParameterMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.EmployeeTypeMaster, "frmHREmployeeTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmQualificationMaster, "FrmQualificationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProfileMaster, "FrmProfileMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRoundMaster, "FrmRoundMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmchkList, "FrmCheckListMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmRelationMaster, "FrmRelationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSourceTypeMaster, "FrmSourceTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSourceTypeDetail, "FrmSourceTypeDetail", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.HRBudgeting, "FrmHRBudgeting", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RequesitionEntry, "FrmRequesitionEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RequesitionApproval, "FrmRequesitionApprovel", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RequesitionClose, "FrmCloseRequestion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.HRIndustryType, "frmHRIndustryType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.HRVerticalMaster, "FrmHRVerticalMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmApplicantEntry, "FrmApplicantEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmShortlist, "FrmShortlist", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmInterviewSchedule, "FrmInterviewSchedule", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmInterviewFeedback, "FrmInterviewFeedback", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.OfferChkList, "FrmOfferCheckList", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.JOININGCHECKLIST, "FrmJoiningChecklist", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmReferenceCheck, "FrmReferenceCheck", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmSalaryFitment, "FrmSalaryFitment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmOfferLetterHR, "FrmOfferLetterHR", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmAppointmentLetterHR, "frmAppointmentLetterHR", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHireEmployee, "FrmHireEmployee", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHrTraineeFeedBack, "FrmTraineeFeedback", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHrTrainerFeedBack, "FrmHrTrainerFeedback", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDamageCaused, "FrmDamageCaused", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.AgencyMaster, "FrmAgencyMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDamageMaster, "FrmDamageMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TrainingMaster, "frmTrainingMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TrainingResourceMaster, "frmTrainingResourceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TrainingInstituteMaster, "frmInstituteMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TrainingRequestMaster, "frmRequestForTrainingMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TRAINER_MASTER, "frmTrainerMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Schedule_Training, "frmScheduleForTraining", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.Training_Attendance, "frmTrainingAttendance", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.TrainingMaster, "frmTrainingMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRPerformanceCategoryMaster, "frmHRPerformanceCategoryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRPerformanceMaster, "FrmHRPerformanceMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRPerformanceGroup, "FrmHRPerformanceGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRPerformanceGroupMapping, "FrmPerformanceGroupMapping", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRPerformanceRating, "FrmPerformanceRating", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRReimbursementTypeMaster, "frmHRReimbursementTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelPurposeMaster, "FrmHRTravelPurposeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelCategoryMaster, "FrmHRTravelCategoryMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelBookedForMaster, "FrmHRTravelBookedForMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelModeTypeMaster, "frmHRTravelModeTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelCityMaster, "FrmHRTravelCityMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelClassTypeMaster, "FrmHRTravelClassTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRHotelRatingMaster, "FrmHRHotelRatingMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelRoomTypeMaster, "FrmHRTravelRoomTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelCarTypeMaster, "FrmHRTravelCarTypeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRRaiseTravelRequisition, "FrmHRRaiseTravelRequisition", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelReqApproval, "frmHRTravelReqApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelReimbursementExpense, "FrmHRTravelReimbursementExpense", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRTravelExpenseApproval, "FrmHRTravelExpenseApproval", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmHRApprovarCreationMaster, "FrmApproverCreationMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVPFSettings, "FrmVPFSettings", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmVPFActiveReport, "FrmVPFActiveScreens", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmGLCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCommonCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmAPCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmARCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkSaleCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFreshSaleCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProductSaleCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCSASaleCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMMCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMCCProcurementCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmBulkProcurementCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmPurchaseCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmDProductionCycle, "FrmModuleCycle", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFaultCategory, "FrmFaultCategory", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmFaultMaster, "FrmFaultMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceChargeMaster, "FrmServiceChargeMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmProblemType, "FrmProblemType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmCallType, "FrmCallType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmActivityType, "FrmActivityType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmSolutionKnowledgeBase, "FrmSolutionKnowledgeBase", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceCall, "FrmServiceCall", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceEnquiry, "FrmServiceEnquiry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceAllocation, "FrmServiceAllocation", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmServiceVisitDetails, "FrmServiceVisitDetails", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.FrmItemConversion, "FrmItemConversion", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmMilkReasonMaster, "frmMilkReasonMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmPaymentCycleMaster, "frmPaymentCycleMaster", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.MilkMPPayment, "FrmMilkVSPPayment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmIncentiveMaster, "frmIncentiveMaster", "ERP.EXE")
        'UpdateFormName(clsUserMgtCode.MilkTruckSheet, "FrmMilkTruckSheet", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkVSPPayment, "FrmMilkVSPPayment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkVSPIssuePayment, "FrmMilkVSPPayment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MilkRecurringScheduler, "FrmMilkRecurringScheduler", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISDebtorReport, "FrmRptCustomerLedgerDemo", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISStockReco, "FrmStockReco", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISSaleRegister, "RptSaleRegisterReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISCreditorReport, "frmRptVendorLedger", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmRptCustomerTransList, "FrmRptCustomerTransList", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISStockLedgerReport, "FrmStockReco", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISStockLedgerReport, "FrmStockReco", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmItemReloadReport, "FrmItemReloadReport", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.rptSKUWiseSale, "rptSKUWiseSale", "ERP.EXE")

        UpdateFormName(clsUserMgtCode.FrmMilkJobWork, "frmMilkRGP", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMilkGateEntry, "FrmMilkGateEntry", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMilkWeighment, "FrmMilkWeighment", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmJobMilkQualityCheck, "FrmJobMilkQualityCheck", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmJobMilkSRN, "FrmJobMilkSRN", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.FrmMilkUnloading, "FrmMilkUnloading", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmIncomeTaxSlab, "frmIncomeTaxSlab", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmITSection, "FrmITSection", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmInvestmentType, "FrmInvestmentType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHouseRentDeclaration, "FrmHouseRentDeclaration", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmHRAExemptionRule, "FrmHRAExemptionRule", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmInvestmentDeclaration, "FrmInvestmentDeclaration", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.RptJobWorkStatus, "RptJobWorkStatus", "ERP.EXE")

        'UpdateFormName(clsUserMgtCode.TimeTable, "frmTimeTable", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.ReversejournalEntry, "frmJournalEntry", "ERP.EXE")
        '===========done by sanjeet on 22/11/2016
        'UpdateFormName(clsUserMgtCode.FrmTruckSheetRouteWiseRpt, "FrmTruckSheetRouteWiseRpt", "ERP.EXE")
        '23/112016
        'UpdateFormName(clsUserMgtCode.FrmMccWeightDifferenceRpt, "FrmMccWeightDifferenceRpt", "ERP.EXE")
        '(09/01/2017)
        'UpdateFormName(clsUserMgtCode.RptAClassMilkRate, "RptAClassMilkRate", "ERP.EXE")
        '=================
        'UpdateFormName(clsUserMgtCode.ApproveFailedSample, "Approve Failed Sample", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISitemGroups, "frmMISItemGroup", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.MISitemMaster, "frmMISItemMaster", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDailyMilkProducts, "frmDailyMilkProducts", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmProductionTransactionType, "frmProductionTransactionType", "ERP.EXE")
        UpdateFormName(clsUserMgtCode.frmDailySMPProduction, "frmDailySMPProduction", "ERP.EXE")

    End Sub

    Public Shared Sub UpdateFormName(ProgramId As String, FormName As String, AssamblyName As String)
        Dim qry As String = " update tspl_program_master set MainFormName='" & FormName & "',AsmName='" & AssamblyName & "' where Program_code='" & ProgramId & "'"
        clsDBFuncationality.ExecuteNonQuery(qry)
    End Sub

End Class

Public Class clsStandardMethods


    ' Defaine Constants for ClassName Here
    '           for Example
    Public Const XpertERPEngineclsCreateAllTable As String = "XpertERPEngine.clsCreateAllTables"


    ' Define Constants for MethodName Here
    '           for Example
    Public Const testMethod As String = "testMethod"
    Public Const testMethodArg As String = "arg1,arg2"

    Public Shared Sub AddStandardFunction()
        Try
            ' Insert Method for Recognization
            '           for Example
            insertDefaultValue("MethodForTest", XpertERPEngineclsCreateAllTable, testMethod, testMethodArg)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Public Shared Sub insertDefaultValue(MethodUniqueCode As String, ClassName As String, MethodName As String, Optional MethodArg As String = Nothing)
        Try
            MethodArg = clsCommon.myCstr(testMethodArg)
            If clsCommon.myLen(MethodUniqueCode) <= 0 Then
                Throw New Exception("Please Specify Mehod Unique Code")
            End If
            If clsCommon.myLen(ClassName) <= 0 Then
                Throw New Exception("Please Specify Class Name")
            End If
            If clsCommon.myLen(MethodName) <= 0 Then
                Throw New Exception("Please Specify Method Name")
            End If
            Dim qry As String = ""
            Dim cnt As Integer = clsCommon.myCdbl(clsDBFuncationality.getSingleValue("select count(*) from tspl_Standard_Method where MethodCode='" & MethodUniqueCode & "'"))
            If cnt > 0 Then
                qry = " update tspl_Standard_Method set ClassName='" & ClassName & "',MethodName='" & MethodName & "',MethodArg='" & MethodArg & "' where MethodCode='" & MethodUniqueCode & "'"
            Else
                qry = " Insert into tspl_Standard_Method(MethodCode,ClassName,MethodName,MethodArg) values('" & MethodUniqueCode & "','" & ClassName & "','" & MethodName & "','" & MethodArg & "')"

            End If
            clsDBFuncationality.ExecuteNonQuery(qry)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
End Class